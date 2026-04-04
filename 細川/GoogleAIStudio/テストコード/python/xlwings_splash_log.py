# -*- coding: utf-8 -*-
"""
段階1/2: logging 出力を xlwings 経由で Excel の UserForm（txtExecutionLog）へ追記する。

有効条件:
  - 環境変数 PM_AI_SPLASH_XLWINGS=1（VBA の .cmd で set）
  - TASK_INPUT_WORKBOOK にマクロブックのフルパス（既に Excel で開いていること）
  - 標準モジュールに Public Sub SplashLog_AppendChunk(ByVal chunk As String)（名前変更時は PM_AI_XLWINGS_SPLASH_MACRO=モジュール名.マクロ名）

execution_log.txt へのファイルログは従来どおり（LOG シート取り込み用）。
"""
from __future__ import annotations

import os
import threading
import time

FLUSH_INTERVAL_SEC = 0.12
MAX_MACRO_ARG_CHARS = 2800

_buf: list[str] = []
_lock = threading.Lock()
_last_flush = 0.0
_book_ref = None


def enabled() -> bool:
    v = (os.environ.get("PM_AI_SPLASH_XLWINGS") or "").strip().lower()
    return v in ("1", "true", "yes", "on")


def _macro_qualified_name() -> str:
    return (os.environ.get("PM_AI_XLWINGS_SPLASH_MACRO") or "SplashLog_AppendChunk").strip()


def _task_workbook_abs() -> str:
    p = (os.environ.get("TASK_INPUT_WORKBOOK") or "").strip()
    if not p:
        return ""
    return os.path.normcase(os.path.normpath(os.path.abspath(p)))


def _get_book():
    global _book_ref
    import xlwings as xw

    target = _task_workbook_abs()
    if not target:
        return None
    if _book_ref is not None:
        try:
            fn = _book_ref.full_name
            if fn and os.path.normcase(os.path.normpath(os.path.abspath(fn))) == target:
                return _book_ref
        except Exception:
            pass
        _book_ref = None
    for app in xw.apps:
        try:
            for book in app.books:
                try:
                    fn = book.full_name
                    if fn and os.path.normcase(os.path.normpath(os.path.abspath(fn))) == target:
                        _book_ref = book
                        return book
                except Exception:
                    continue
        except Exception:
            continue
    return None


def _invoke_macro(chunk: str) -> None:
    global _book_ref

    name = _macro_qualified_name()
    for _ in range(2):
        book = _get_book()
        if book is None:
            return
        try:
            book.macro(name)(chunk)
            return
        except Exception:
            _book_ref = None


def _dispatch_payload(payload: str) -> None:
    while payload:
        part = payload[:MAX_MACRO_ARG_CHARS]
        payload = payload[MAX_MACRO_ARG_CHARS:]
        _invoke_macro(part)


def flush(force: bool = False) -> None:
    global _last_flush, _buf
    payload = ""
    with _lock:
        now = time.monotonic()
        total = sum(len(x) for x in _buf)
        if not _buf:
            return
        if not force and (now - _last_flush) < FLUSH_INTERVAL_SEC and total < 900:
            return
        payload = "".join(_buf)
        _buf = []
        _last_flush = now
    if payload:
        _dispatch_payload(payload)


def append_formatted_line(text: str) -> None:
    if not enabled() or not text:
        return
    line = text if text.endswith("\n") else text + "\n"
    with _lock:
        _buf.append(line)
    flush(force=False)


def shutdown() -> None:
    if enabled():
        flush(force=True)
