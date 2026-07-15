from __future__ import annotations

import base64
from datetime import datetime, timedelta, timezone
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey

from autoexcel import license as license_module


class LicenseTest(unittest.TestCase):
    def setUp(self) -> None:
        self.private_key = Ed25519PrivateKey.generate()
        public_bytes = self.private_key.public_key().public_bytes(
            serialization.Encoding.Raw, serialization.PublicFormat.Raw
        )
        self.public_key_base64 = base64.urlsafe_b64encode(public_bytes).decode().rstrip("=")
        self.now = datetime(2026, 7, 14, 10, 0, tzinfo=timezone.utc)

    def payload(self, **overrides: object) -> dict[str, object]:
        payload: dict[str, object] = {
            "version": 1,
            "product": "autoexcel",
            "license_id": "license-001",
            "issued_at": (self.now - timedelta(minutes=1)).isoformat(),
            "expires_at": None,
            "features": ["order_diff", "fetch_orders"],
        }
        payload.update(overrides)
        return payload

    def validate(self, token: str):  # type: ignore[no-untyped-def]
        with patch.object(license_module, "PUBLIC_KEY_BASE64", self.public_key_base64):
            return license_module.validate_license(token, now=self.now)

    def test_valid_permanent_license_opens_both_features(self) -> None:
        token = license_module.sign_license(self.payload(), self.private_key)

        info = self.validate(token)

        self.assertTrue(info.valid)
        self.assertTrue(info.allows("order_diff"))
        self.assertTrue(info.allows("fetch_orders"))
        self.assertIsNone(info.expires_at)

    def test_tampered_payload_is_rejected(self) -> None:
        token = license_module.sign_license(self.payload(), self.private_key)
        prefix, payload, signature = token.split(".")
        decoded = json.loads(base64.urlsafe_b64decode(payload + "=="))
        decoded["features"] = ["order_diff"]
        tampered_payload = base64.urlsafe_b64encode(
            json.dumps(decoded, sort_keys=True, separators=(",", ":")).encode()
        ).decode().rstrip("=")

        info = self.validate(f"{prefix}.{tampered_payload}.{signature}")

        self.assertFalse(info.valid)
        self.assertEqual(info.message, "密钥签名无效")

    def test_expired_license_is_rejected(self) -> None:
        token = license_module.sign_license(
            self.payload(expires_at=(self.now - timedelta(seconds=1)).isoformat()),
            self.private_key,
        )

        info = self.validate(token)

        self.assertFalse(info.valid)
        self.assertEqual(info.message, "密钥已过期")

    def test_wrong_product_is_rejected(self) -> None:
        token = license_module.sign_license(
            self.payload(product="another-product"), self.private_key
        )

        info = self.validate(token)

        self.assertFalse(info.valid)
        self.assertEqual(info.message, "密钥不适用于当前产品")

    def test_install_and_load_round_trip(self) -> None:
        token = license_module.sign_license(self.payload(), self.private_key)
        with TemporaryDirectory() as temporary_directory, patch.object(
            license_module, "PUBLIC_KEY_BASE64", self.public_key_base64
        ):
            path = Path(temporary_directory) / "license.key"

            installed = license_module.install_license(token, path)
            loaded = license_module.load_license(path)

        self.assertTrue(installed.valid)
        self.assertTrue(loaded.valid)
        self.assertEqual(loaded.license_id, "license-001")


if __name__ == "__main__":
    unittest.main()
