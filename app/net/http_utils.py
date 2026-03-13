from __future__ import annotations

import ssl
import urllib.request
from typing import Mapping


def _enable_windows_cert_patch() -> bool:
    """Enable certifi patch for Windows certificate store, if available."""
    try:
        import certifi_win32  # type: ignore # noqa: F401

        return True
    except Exception:
        return False


def _build_ssl_context() -> ssl.SSLContext | None:
    """Build SSL context that respects enterprise proxy certificates."""
    if not _enable_windows_cert_patch():
        return None

    try:
        import certifi  # type: ignore

        ca_file = certifi.where()
        if not ca_file:
            return None
        return ssl.create_default_context(cafile=ca_file)
    except Exception:
        return None


_SSL_CONTEXT = _build_ssl_context()


def _build_request(
    target: str | urllib.request.Request,
    headers: Mapping[str, str] | None,
) -> urllib.request.Request:
    if isinstance(target, urllib.request.Request):
        if headers:
            for key, value in headers.items():
                target.add_header(key, value)
        return target
    return urllib.request.Request(target, headers=dict(headers or {}))


def urlopen(
    target: str | urllib.request.Request,
    *,
    timeout: float = 10.0,
    headers: Mapping[str, str] | None = None,
):
    request = _build_request(target, headers)
    if _SSL_CONTEXT is not None:
        return urllib.request.urlopen(request, timeout=timeout, context=_SSL_CONTEXT)
    return urllib.request.urlopen(request, timeout=timeout)
