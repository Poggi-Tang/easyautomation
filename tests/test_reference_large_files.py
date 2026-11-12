"""Large media documents use disk-backed sources without rejecting their images."""
import hashlib
import io
from email.message import Message
from pathlib import Path
from zipfile import ZipFile

import pytest

from easy_uiauto import reference_library as refs
from test_knowledge_library import library
from test_reference_library import SETTINGS, understand


DOCUMENT = b'''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:r><w:t>Save the project to disk.</w:t></w:r></w:p></w:body></w:document>'''


def digest(path):
    checksum = hashlib.sha256()
    with path.open("rb") as stream:
        while block := stream.read(1024*1024):checksum.update(block)
    return checksum.hexdigest()


def test_import_over_40mb_preserves_source_and_deduplicates_without_read_bytes(library, tmp_path, understand, monkeypatch):
    source = tmp_path / "large-manual.docx"
    with ZipFile(source, "w") as archive:
        archive.writestr("word/document.xml", DOCUMENT)
        with archive.open("word/media/large-image.png", "w") as image:
            for _ in range(41):image.write(b"x" * (1024*1024))
    original_hash = digest(source)
    assert source.stat().st_size > 40*1024*1024
    staged, messages = [], []
    original_extract = refs.extract
    def extract(path, *args):
        assert isinstance(path, Path) and path != source
        staged.append(path)
        return original_extract(path, *args)
    monkeypatch.setattr(refs, "extract", extract)
    monkeypatch.setattr(Path, "read_bytes", lambda *_: pytest.fail("whole file read"))
    result = library.import_reference("demo", source, settings=SETTINGS, progress=messages.append)
    archived = library.directory("demo") / result["source_asset"]
    assert digest(archived) == original_hash == result["source_sha256"]
    assert "Save the project" in result["body"] and result["name"] == source.name
    assert staged and not staged[0].exists()
    assert any("MB" in message for message in messages)
    assert library.import_reference("demo", source, settings=SETTINGS)["id"] == result["id"]
    assert len(understand) == 1 and len(staged) == 1


def test_source_snapshot_survives_original_edit_and_is_cleaned_on_failure(tmp_path):
    source = tmp_path / "manual.md"
    source.write_text("Original content", encoding="utf-8")
    staged = None
    with pytest.raises(RuntimeError, match="cancelled"):
        with refs._read_source(str(source)) as (staged, suffix, title, origin, checksum):
            source.write_text("Changed content", encoding="utf-8")
            assert staged.read_text(encoding="utf-8") == "Original content"
            assert checksum == hashlib.sha256(b"Original content").hexdigest()
            assert (suffix, title, origin) == (".md", source.name, str(source.resolve()))
            raise RuntimeError("cancelled")
    assert not staged.exists() and source.read_text(encoding="utf-8") == "Changed content"


def test_large_embedded_media_is_not_decompressed_or_counted_as_markup(tmp_path, monkeypatch):
    source = tmp_path / "manual.docx"
    with ZipFile(source, "w") as archive:
        archive.writestr("word/document.xml", DOCUMENT)
        archive.writestr("word/media/image.png", b"x" * 5000)
    monkeypatch.setattr(refs, "MAX_MARKUP_BYTES", len(DOCUMENT)+1)
    original_open = ZipFile.open
    def open_part(archive, name, *args, **kwargs):
        assert name == "word/document.xml"
        return original_open(archive, name, *args, **kwargs)
    monkeypatch.setattr(ZipFile, "open", open_part)
    assert refs.extract(source, ".docx")[0][0]["text"] == "Save the project to disk."


def test_oversized_document_structure_fails_before_decompression(tmp_path, monkeypatch):
    source = tmp_path / "manual.docx"
    with ZipFile(source, "w") as archive:archive.writestr("word/document.xml", DOCUMENT)
    monkeypatch.setattr(refs, "MAX_MARKUP_BYTES", len(DOCUMENT)-1)
    monkeypatch.setattr(ZipFile, "open", lambda *_: pytest.fail("oversized markup decompressed"))
    with pytest.raises(ValueError, match="正文结构"):refs.extract(source, ".docx")


def test_pdf_reader_receives_a_seekable_file_and_stream_is_closed(tmp_path, monkeypatch):
    import pypdf
    from pypdf.generic import DictionaryObject, NameObject, DecodedStreamObject
    source = tmp_path / "manual.pdf"
    writer = pypdf.PdfWriter();page = writer.add_blank_page(width=400, height=300)
    font = DictionaryObject({NameObject("/Type"):NameObject("/Font"), NameObject("/Subtype"):NameObject("/Type1"),
                             NameObject("/BaseFont"):NameObject("/Helvetica")})
    page[NameObject("/Resources")] = DictionaryObject({NameObject("/Font"):DictionaryObject({NameObject("/F1"):writer._add_object(font)})})
    content = DecodedStreamObject();content.set_data(b"BT /F1 12 Tf 10 250 Td (Save project to disk) Tj ET")
    page[NameObject("/Contents")] = writer._add_object(content)
    writer.add_attachment("large-diagram.bin", b"x" * (41*1024*1024))
    writer.write(source)
    assert source.stat().st_size > 40*1024*1024
    reader = pypdf.PdfReader;streams = []
    def open_pdf(stream):
        assert not isinstance(stream, (bytes, str, Path, io.BytesIO))
        streams.append(stream)
        return reader(stream)
    monkeypatch.setattr(pypdf, "PdfReader", open_pdf)
    assert "Save project" in refs.extract(source, ".pdf")[0][0]["text"]
    assert streams[0].closed


def test_disk_backed_scanned_pdf_uses_path_for_ocr(tmp_path, monkeypatch):
    import pypdfium2
    from PIL import Image
    from easy_uiauto import local_image_worker
    source = tmp_path / "scanned.pdf"
    Image.new("RGB", (100, 100), "white").save(source)
    original = pypdfium2.PdfDocument;paths = []
    def render(path):paths.append(path);return original(path)
    monkeypatch.setattr(pypdfium2, "PdfDocument", render)
    monkeypatch.setattr(local_image_worker, "read_text", lambda *a, **kw: "Save the project to disk.")
    assert refs.extract(source, ".pdf")[0][0]["origin"] == "PDF 第 1 页"
    assert paths == [str(source)]


@pytest.mark.parametrize("encoding", ["utf-8-sig", "utf-16", "gb18030"])
def test_text_file_encodings_still_work(encoding, tmp_path):
    source = tmp_path / "manual.txt"
    source.write_bytes("保存当前工程。".encode(encoding))
    assert refs.extract(source, ".txt")[0][0]["text"] == "保存当前工程。"


def test_text_limit_is_enforced_before_reading_remainder(tmp_path, monkeypatch):
    source = tmp_path / "manual.txt";source.write_text("x"*50000, encoding="utf-8")
    monkeypatch.setattr(refs, "MAX_CHARS", 100)
    with source.open("rb") as stream:
        with pytest.raises(ValueError, match="文字"):
            refs._text_source(stream)
        assert stream.tell() < source.stat().st_size


@pytest.mark.parametrize("over_limit", [False, True])
def test_web_download_uses_bounded_reads_and_cleans_up(library, tmp_path, understand, monkeypatch, over_limit):
    import tempfile
    monkeypatch.setattr(refs, "MAX_DOWNLOAD_BYTES", 200)
    monkeypatch.setattr(refs, "COPY_BYTES", 16)
    folders = []
    def temporary(**kwargs):
        folder = tempfile.TemporaryDirectory(dir=tmp_path, **kwargs);folders.append(Path(folder.name));return folder
    monkeypatch.setattr(refs, "TemporaryDirectory", temporary)
    class Response(io.BytesIO):
        headers = Message()
        def geturl(self):return "https://manual.example.test/guide"
        def read(self, size=-1):
            assert 0 < size <= 16
            return super().read(size)
    Response.headers["Content-Type"] = "text/plain"
    monkeypatch.setattr(refs.request, "urlopen", lambda *a, **kw: Response(b"Save to disk. "*(20 if over_limit else 2)))
    if over_limit:
        with pytest.raises(ValueError, match="下载到本地"):
            library.import_reference("demo", "https://manual.example.test/guide", settings=SETTINGS)
        assert not understand and not library.reference_documents("demo")
    else:
        result = library.import_reference("demo", "https://manual.example.test/guide", settings=SETTINGS)
        assert "Save to disk" in result["body"]
    assert folders and all(not folder.exists() for folder in folders)
