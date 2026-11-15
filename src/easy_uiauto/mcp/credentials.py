"""Current-user Windows Credential Manager storage, also usable by the config EXE."""

from __future__ import annotations

import ctypes
import os
from ctypes import wintypes

TARGET_NAME = "easy-uiauto/vision-api-key"
_GENERIC = 1
_LOCAL_MACHINE = 2  # Persist for this user across logons; not shared with other users.
_NOT_FOUND = 1168


class _Credential(ctypes.Structure):
    _fields_ = [
        ("Flags", wintypes.DWORD), ("Type", wintypes.DWORD),
        ("TargetName", wintypes.LPWSTR), ("Comment", wintypes.LPWSTR),
        ("LastWritten", wintypes.FILETIME), ("CredentialBlobSize", wintypes.DWORD),
        ("CredentialBlob", ctypes.POINTER(ctypes.c_ubyte)),
        ("Persist", wintypes.DWORD), ("AttributeCount", wintypes.DWORD),
        ("Attributes", ctypes.c_void_p), ("TargetAlias", wintypes.LPWSTR),
        ("UserName", wintypes.LPWSTR),
    ]


def _api():
    if os.name != "nt":
        raise OSError("Windows Credential Manager is available only on Windows")
    api = ctypes.WinDLL("advapi32", use_last_error=True)
    api.CredReadW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                             ctypes.POINTER(ctypes.POINTER(_Credential))]
    api.CredReadW.restype = wintypes.BOOL
    api.CredWriteW.argtypes = [ctypes.POINTER(_Credential), wintypes.DWORD]
    api.CredWriteW.restype = wintypes.BOOL
    api.CredDeleteW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD]
    api.CredDeleteW.restype = wintypes.BOOL
    api.CredFree.argtypes = [ctypes.c_void_p]
    api.CredFree.restype = None
    return api


def _failure(operation: str) -> OSError:
    # Never interpolate input, OS exception details, or the credential blob.
    return OSError(f"Windows Credential Manager could not {operation} the API key "
                   f"(Windows error {ctypes.get_last_error()})")


def read_api_key() -> str:
    if os.name != "nt":
        return ""
    api = _api()
    pointer = ctypes.POINTER(_Credential)()
    if not api.CredReadW(TARGET_NAME, _GENERIC, 0, ctypes.byref(pointer)):
        if ctypes.get_last_error() == _NOT_FOUND:
            return ""
        raise _failure("read")
    try:
        credential = pointer.contents
        blob = ctypes.string_at(credential.CredentialBlob, credential.CredentialBlobSize)
        try:
            return blob.decode("utf-16-le")
        except UnicodeError:
            raise OSError("Saved API key is invalid; save it again in Settings") from None
    finally:
        api.CredFree(pointer)


def write_api_key(value: str) -> None:
    if not value.strip() or any(char in value for char in ("\r", "\n", "\x00")):
        raise ValueError("请填写有效的 API Key")
    blob = value.encode("utf-16-le")
    if len(blob) > 2560:
        raise ValueError("API Key 超出 Windows 凭据管理器的长度限制")
    buffer = (ctypes.c_ubyte * len(blob)).from_buffer_copy(blob)
    credential = _Credential(
        Type=_GENERIC, TargetName=TARGET_NAME, CredentialBlobSize=len(blob),
        CredentialBlob=buffer, Persist=_LOCAL_MACHINE, UserName="easy-uiauto",
    )
    api = _api()
    try:
        if not api.CredWriteW(ctypes.byref(credential), 0):
            raise _failure("save")
    finally:
        ctypes.memset(buffer, 0, len(blob))


def delete_api_key() -> None:
    """Remove this application's credential, without touching any other target."""
    api = _api()
    if not api.CredDeleteW(TARGET_NAME, _GENERIC, 0):
        if ctypes.get_last_error() != _NOT_FOUND:
            raise _failure("delete")
