from __future__ import annotations

import pytest

from easy_uiauto import cli


def test_cli_version(capsys):
    with pytest.raises(SystemExit) as exc:
        cli.main(["--version"])
    assert exc.value.code == 0
    assert "0.1.8" in capsys.readouterr().out


def test_cli_help(capsys):
    with pytest.raises(SystemExit) as exc:
        cli.main(["--help"])
    assert exc.value.code == 0
    assert "record" in capsys.readouterr().out
