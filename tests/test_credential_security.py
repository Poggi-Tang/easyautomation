"""Credentials must never fall back to plaintext storage or provider error logs."""

import io
import json
import os
import sys
import traceback
from contextlib import nullcontext
from types import SimpleNamespace
from urllib.error import HTTPError, URLError
from uuid import uuid4

import pytest
from PIL import Image

from easy_uiauto.mcp import configuration, credentials, diagnostics, scanner, server
from easy_uiauto.mcp import interaction_learning
from easy_uiauto import reference_library, region_learning


@pytest.fixture
def settings_store(monkeypatch):
    """Never read or modify the developer's registry or real app credential."""
    registry, secure, writes = {}, {}, []

    def query(_key, name):
        if name not in registry:
            raise FileNotFoundError()
        return registry[name]

    def write(_key, name, _reserved, kind, value):
        writes.append((name, value))
        registry[name] = (value, kind)

    def delete(_key, name):
        if name not in registry:
            raise FileNotFoundError()
        del registry[name]

    fake = SimpleNamespace(HKEY_CURRENT_USER=1, REG_SZ=1,
        CreateKey=lambda *_: nullcontext(1), QueryValueEx=query, SetValueEx=write,
        DeleteValue=delete)
    monkeypatch.setitem(sys.modules, "winreg", fake)
    monkeypatch.setattr(os, "environ", dict(os.environ))
    monkeypatch.setattr(credentials, "write_api_key", lambda key: secure.update(key=key))
    monkeypatch.setattr(configuration, "_broadcast_environment_change", lambda: None)
    return registry, secure, writes


def test_new_save_migrates_plaintext_and_does_not_export_secret(settings_store, tmp_path):
    registry, secure, writes = settings_store
    registry[configuration.VISION_API_KEY] = ("legacy-key", 1)
    os.environ[configuration.VISION_API_KEY] = "inherited-key"
    result = configuration.save_gui_configuration(str(tmp_path), "https://example.test/v1",
                                                   "new-secret", "vision-model")
    assert secure == {"key": "new-secret"}
    assert configuration.VISION_API_KEY not in registry
    assert configuration.VISION_API_KEY not in os.environ
    assert "new-secret" not in repr(writes) + repr(result)
    assert registry[configuration.VISION_MODEL][0] == "vision-model"


@pytest.mark.parametrize("legacy", [False, True])
def test_credential_failure_restores_old_settings_without_plaintext_fallback(
        settings_store, monkeypatch, tmp_path, legacy):
    registry, secure, writes = settings_store
    registry[configuration.VISION_API_URL] = ("https://old.test", 2)
    if legacy:
        registry[configuration.VISION_API_KEY] = ("old-secret", 1)
    secure["key"] = "existing-credential"
    before_registry, before_env = registry.copy(), os.environ.copy()

    def fail(_value):
        raise OSError("credential store unavailable")

    monkeypatch.setattr(credentials, "write_api_key", fail)
    with pytest.raises(OSError, match="unavailable"):
        configuration.save_gui_configuration(str(tmp_path), "https://new.test", "new-secret", "m")
    assert registry == before_registry
    assert os.environ == before_env
    assert secure == {"key": "existing-credential"}
    assert "new-secret" not in repr(writes)


def test_registry_failure_keeps_existing_credential(settings_store, monkeypatch, tmp_path):
    registry, secure, _writes = settings_store
    secure["key"] = "existing"
    registry[configuration.KNOWLEDGE_DIR] = ("old-vault", 1)
    before = registry.copy()
    fake = sys.modules["winreg"]
    original = fake.SetValueEx

    def fail(key, name, reserved, kind, value):
        if name == configuration.VISION_API_URL:
            raise OSError("denied")
        original(key, name, reserved, kind, value)

    monkeypatch.setattr(fake, "SetValueEx", fail)
    with pytest.raises(OSError, match="denied"):
        configuration.save_gui_configuration(str(tmp_path), "https://new.test", "new-secret", "m")
    assert registry == before and secure == {"key": "existing"}


def test_cli_key_save_uses_same_secure_store(settings_store):
    registry, secure, writes = settings_store
    configuration._write_user_environment(configuration.VISION_API_KEY, "cli-secret")
    assert secure == {"key": "cli-secret"}
    assert not registry and not writes


@pytest.mark.parametrize("source", ["credential", "registry", "process", "codex"])
def test_key_lookup_priority_and_legacy_compatibility(monkeypatch, source):
    sources = ["credential", "registry", "process", "codex"]
    enabled = set(sources[sources.index(source):])
    monkeypatch.setattr(credentials, "read_api_key", lambda: "credential" if "credential" in enabled else "")
    monkeypatch.setattr(configuration, "_read_user_environment", lambda _: "registry" if "registry" in enabled else "")
    monkeypatch.setenv(configuration.VISION_API_KEY, "process" if "process" in enabled else "")
    monkeypatch.setattr(configuration, "_read_existing_codex_environment", lambda _: "codex")
    assert configuration._existing_vision_value(configuration.VISION_API_KEY) == source


@pytest.mark.skipif(os.name != "nt", reason="Windows Credential Manager")
def test_windows_credential_roundtrip_uses_isolated_target(monkeypatch):
    # Exercise the native ABI with a disposable target, never the real user setting.
    monkeypatch.setattr(credentials, "TARGET_NAME", "easy-uiauto-test/" + uuid4().hex)
    try:
        assert credentials.read_api_key() == ""
        credentials.write_api_key("test-密钥-🔑")
        assert credentials.read_api_key() == "test-密钥-🔑"
        credentials.write_api_key("replacement")
        assert credentials.read_api_key() == "replacement"
    finally:
        credentials.delete_api_key()
    assert credentials.read_api_key() == ""


@pytest.mark.parametrize("consumer", ["models", "scan", "semantics", "locate", "diagnostics", "reference", "deep", "effects"])
@pytest.mark.parametrize("failure", ["http", "connection"])
def test_provider_errors_never_reach_results_or_tracebacks(monkeypatch, consumer, failure):
    secret = "sk-test-secret-should-never-be-logged"
    api_url = "https://example.test/v1/chat/completions"
    settings = dict(api_url=api_url, api_key=secret, api_model="test-model")
    picture = Image.new("RGB", (32, 32))

    def fail(*_args, **_kwargs):
        if failure == "http":
            raise HTTPError(api_url, 401, secret, {}, io.BytesIO(json.dumps(
                {"error": {"message": "Invalid API key: " + secret}}).encode()))
        raise URLError("Authorization: Bearer " + secret)

    monkeypatch.setattr(configuration.urlrequest, "urlopen", fail)
    monkeypatch.setattr(server, "_vision_api_settings", lambda _: (api_url, secret, "test-model"))
    monkeypatch.setattr(server.pyautogui, "screenshot", lambda **_: picture)
    calls = {
        "models": lambda: configuration.fetch_models(api_url, secret),
        "scan": lambda: scanner._segment_image(picture, api_url, secret, "test-model", "1"),
        "semantics": lambda: scanner._request_control_semantics(picture, [], {}, [], api_url, secret, "m", "1"),
        "locate": lambda: server._find_control_by_vision("button", "", None),
        "diagnostics": lambda: diagnostics._test_remote_vision(api_url, secret, "m", "1"),
        "reference": lambda: reference_library.complete("reference", {}, settings=settings),
        "deep": lambda: region_learning._request_plan(settings, {}, picture),
        "effects": lambda: interaction_learning._interpret_effects(picture, picture, [], {}, [], [], [], [],
                                                                    "click", api_url, secret, "m", "1"),
    }
    if consumer == "effects":
        result = calls[consumer]()
        assert result["confidence"] == 0
        assert secret not in json.dumps(result)
    else:
        with pytest.raises(RuntimeError) as caught:
            calls[consumer]()
        rendered = "".join(traceback.format_exception(caught.value))
        assert secret not in rendered
        assert caught.value.__cause__ is None and caught.value.__suppress_context__
