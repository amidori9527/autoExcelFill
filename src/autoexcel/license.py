from __future__ import annotations

import base64
import binascii
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
import json
import os
from pathlib import Path
import sys
import tempfile
from typing import Any

from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey, Ed25519PublicKey

from autoexcel.runtime_paths import application_directory


FORMAT_VERSION = "AX1"
PRODUCT = "autoexcel"
FEATURE_ORDER_DIFF = "order_diff"
FEATURE_FETCH_ORDERS = "fetch_orders"
ALLOWED_FEATURES = frozenset({FEATURE_ORDER_DIFF, FEATURE_FETCH_ORDERS})
PUBLIC_KEY_BASE64 = "fFG48x7sCqxD8w_Q9K5wiNN37GXRpkK4HHHZhYKk30s"
PROJECT_ROOT = Path(__file__).resolve().parents[2]


@dataclass(frozen=True)
class LicenseInfo:
    valid: bool
    message: str
    features: frozenset[str] = frozenset()
    license_id: str = ""
    issued_at: datetime | None = None
    expires_at: datetime | None = None

    def allows(self, feature: str) -> bool:
        return self.valid and feature in self.features


def license_file_path() -> Path:
    if getattr(sys, "frozen", False):
        return application_directory() / "license.key"
    return PROJECT_ROOT / "license.key"


def load_license(path: Path | None = None) -> LicenseInfo:
    target = path or license_file_path()
    if not target.exists():
        return LicenseInfo(False, "未配置密钥")
    try:
        token = target.read_text(encoding="utf-8").strip()
    except (OSError, UnicodeError):
        return LicenseInfo(False, "密钥文件无法读取")
    return validate_license(token)


def install_license(token: str, path: Path | None = None) -> LicenseInfo:
    normalized = token.strip()
    info = validate_license(normalized)
    if not info.valid:
        return info
    target = path or license_file_path()
    target.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary_name = tempfile.mkstemp(prefix=".license.", dir=target.parent)
    os.close(fd)
    temporary_path = Path(temporary_name)
    try:
        temporary_path.write_text(normalized + "\n", encoding="utf-8")
        os.replace(temporary_path, target)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()
    return info


def remove_license(path: Path | None = None) -> None:
    target = path or license_file_path()
    if target.exists():
        target.unlink()


def validate_license(token: str, now: datetime | None = None) -> LicenseInfo:
    if not token:
        return LicenseInfo(False, "未配置密钥")
    if len(token) > 8192:
        return LicenseInfo(False, "密钥格式不正确")
    parts = token.split(".")
    if len(parts) != 3 or parts[0] != FORMAT_VERSION:
        return LicenseInfo(False, "密钥格式不正确")
    try:
        payload_bytes = _decode_base64(parts[1])
        signature = _decode_base64(parts[2])
        public_key = Ed25519PublicKey.from_public_bytes(_decode_base64(PUBLIC_KEY_BASE64))
        public_key.verify(signature, payload_bytes)
    except (ValueError, binascii.Error, InvalidSignature):
        return LicenseInfo(False, "密钥签名无效")

    try:
        payload = json.loads(payload_bytes.decode("utf-8"))
        if not isinstance(payload, dict):
            raise ValueError
        if payload.get("version") != 1 or payload.get("product") != PRODUCT:
            return LicenseInfo(False, "密钥不适用于当前产品")
        license_id = str(payload.get("license_id") or "").strip()
        if not license_id:
            return LicenseInfo(False, "密钥缺少许可证编号")
        raw_features = payload.get("features")
        if not isinstance(raw_features, list) or not all(
            isinstance(feature, str) for feature in raw_features
        ):
            return LicenseInfo(False, "密钥功能列表无效")
        features = frozenset(raw_features)
        if not features <= ALLOWED_FEATURES:
            return LicenseInfo(False, "密钥包含未知功能")
        issued_at = _parse_utc_datetime(payload.get("issued_at"))
        expires_at = _parse_utc_datetime(payload.get("expires_at"), optional=True)
    except (TypeError, ValueError, UnicodeDecodeError, json.JSONDecodeError):
        return LicenseInfo(False, "密钥内容无效")

    current_time = now or datetime.now(timezone.utc)
    if current_time.tzinfo is None:
        current_time = current_time.replace(tzinfo=timezone.utc)
    current_time = current_time.astimezone(timezone.utc)
    if issued_at > current_time + timedelta(minutes=5):
        return LicenseInfo(False, "密钥签发时间无效")
    if expires_at is not None and expires_at <= issued_at:
        return LicenseInfo(False, "密钥有效期无效")
    if expires_at is not None and expires_at <= current_time:
        return LicenseInfo(False, "密钥已过期", license_id=license_id, expires_at=expires_at)
    return LicenseInfo(
        True,
        "密钥有效",
        features=features,
        license_id=license_id,
        issued_at=issued_at,
        expires_at=expires_at,
    )


def sign_license(payload: dict[str, Any], private_key: Ed25519PrivateKey) -> str:
    payload_bytes = json.dumps(
        payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")
    ).encode("utf-8")
    signature = private_key.sign(payload_bytes)
    return f"{FORMAT_VERSION}.{_encode_base64(payload_bytes)}.{_encode_base64(signature)}"


def _parse_utc_datetime(value: Any, optional: bool = False) -> datetime | None:
    if optional and value in {None, ""}:
        return None
    if not isinstance(value, str):
        raise ValueError
    parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    if parsed.tzinfo is None:
        raise ValueError
    return parsed.astimezone(timezone.utc)


def _encode_base64(value: bytes) -> str:
    return base64.urlsafe_b64encode(value).decode("ascii").rstrip("=")


def _decode_base64(value: str) -> bytes:
    padding = "=" * (-len(value) % 4)
    return base64.b64decode(value + padding, altchars=b"-_", validate=True)
