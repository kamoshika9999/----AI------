"""
平文 Gemini 認証 JSON を暗号化する CLI。
復号側（planning_core）はソース内の定数のみ使用するため、暗号化時のパスフレーズは社内手順どおりに指定すること。

    py encrypt_gemini_credentials.py plain.json encrypted.json
    py encrypt_gemini_credentials.py plain.json out.json --passphrase-file pass.txt
"""
from __future__ import annotations

import argparse
import base64
import json
import os
import secrets
import sys

# 相対パス（--passphrase-file 等）は呼び出し元のカレントに依存しないよう、先にスクリプト所在へ
os.chdir(os.path.dirname(os.path.abspath(__file__)))

_DEFAULT_ITERATIONS = 480_000
_DEFAULT_PASSPHRASE = "nagaoka1234"


def _derive_fernet_key(passphrase: str, salt: bytes, iterations: int) -> bytes:
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.backends import default_backend
    from cryptography.hazmat.primitives import hashes

    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=iterations,
        backend=default_backend(),
    )
    return base64.urlsafe_b64encode(kdf.derive(passphrase.encode("utf-8")))


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Gemini 認証 JSON を Fernet + PBKDF2 で暗号化（--passphrase / --passphrase-file でパスフレーズ指定。復号は planning_core の定数のみ）"
    )
    parser.add_argument("plain_json", help="平文 JSON（gemini_api_key を含む）")
    parser.add_argument("output_json", help="出力する暗号化 JSON のパス")
    parser.add_argument(
        "--iterations",
        type=int,
        default=_DEFAULT_ITERATIONS,
        help=f"PBKDF2 繰り返し回数（既定: {_DEFAULT_ITERATIONS}）",
    )
    parser.add_argument(
        "--passphrase",
        default=_DEFAULT_PASSPHRASE,
        help="暗号化パスフレーズ（--passphrase-file 未指定時。社内手順の値と planning_core の復号用定数を一致させること）",
    )
    parser.add_argument(
        "--passphrase-file",
        default="",
        help="パスフレーズを UTF-8 の1ファイルから読む（VBA マクロから渡すとき用。改行・前後空白は除去）",
    )
    args = parser.parse_args()

    try:
        from cryptography.fernet import Fernet
    except ImportError:
        print(
            "cryptography が未インストールです。次を実行してください:\n"
            "  py -3 -m pip install cryptography\n"
            "または:\n"
            "  py -3 -m pip install -r python\\requirements.txt",
            file=sys.stderr,
        )
        return 1

    try:
        # VBA（ADODB.Stream UTF-8）が先頭に BOM を付けるため utf-8-sig
        with open(args.plain_json, encoding="utf-8-sig") as f:
            inner = json.load(f)
    except OSError as ex:
        print(f"平文 JSON を開けません: {args.plain_json} ({ex})", file=sys.stderr)
        return 1
    except json.JSONDecodeError as ex:
        print(f"平文 JSON の構文エラー: {args.plain_json} ({ex})", file=sys.stderr)
        return 1
    if not isinstance(inner, dict):
        print("平文 JSON はオブジェクト形式である必要があります。", file=sys.stderr)
        return 1
    key = inner.get("gemini_api_key") or inner.get("GEMINI_API_KEY")
    if not key or not str(key).strip():
        print("平文 JSON に gemini_api_key（または GEMINI_API_KEY）がありません。", file=sys.stderr)
        return 1

    inner_min = {"gemini_api_key": str(key).strip()}
    inner_bytes = json.dumps(inner_min, ensure_ascii=False).encode("utf-8")

    phrase = (args.passphrase or "").strip()
    pfile = (args.passphrase_file or "").strip()
    if pfile:
        try:
            with open(pfile, encoding="utf-8-sig") as pf:
                phrase = pf.read().strip()
        except OSError as ex:
            print(f"パスフレーズファイルを読めません: {pfile} ({ex})", file=sys.stderr)
            return 1

    if not phrase:
        print("パスフレーズが空です。", file=sys.stderr)
        return 1

    salt = secrets.token_bytes(16)
    fkey = _derive_fernet_key(phrase, salt, args.iterations)
    token = Fernet(fkey).encrypt(inner_bytes)
    out_obj = {
        "format_version": 2,
        "kdf": "pbkdf2_sha256",
        "iterations": args.iterations,
        "salt_b64": base64.standard_b64encode(salt).decode("ascii"),
        "fernet_ciphertext": token.decode("ascii"),
        "description": "encrypt_gemini_credentials.py で生成。復号は planning_core の定数のみ（パスフレーズは社内手順で管理）。",
    }

    try:
        with open(args.output_json, "w", encoding="utf-8") as f:
            json.dump(out_obj, f, ensure_ascii=False, indent=2)
            f.write("\n")
    except OSError as ex:
        print(f"暗号化 JSON を書き込めません: {args.output_json} ({ex})", file=sys.stderr)
        return 1

    print(f"書き出しました: {args.output_json}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as ex:
        print(f"予期しないエラー: {ex}", file=sys.stderr)
        raise SystemExit(1)
