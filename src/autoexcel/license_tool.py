from __future__ import annotations

import argparse
from datetime import datetime, time, timezone
from pathlib import Path
import uuid

from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey

from autoexcel.license import (
    FEATURE_ADD_B2B,
    FEATURE_ADD_CARDS,
    FEATURE_FETCH_ORDERS,
    FEATURE_ORDER_DIFF,
    PRODUCT,
    PROJECT_ROOT,
    sign_license,
)


DEFAULT_PRIVATE_KEY = PROJECT_ROOT / ".license" / "autoexcel-ed25519-private.pem"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate a signed SmartSheet Desk license key.")
    parser.add_argument("--private-key", type=Path, default=DEFAULT_PRIVATE_KEY)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--expires", help="Optional expiration date in YYYY-MM-DD (UTC).")
    parser.add_argument("--customer", default="", help="Optional customer label.")
    parser.add_argument(
        "--features",
        nargs="+",
        choices=(
            FEATURE_ORDER_DIFF,
            FEATURE_FETCH_ORDERS,
            FEATURE_ADD_CARDS,
            FEATURE_ADD_B2B,
        ),
        default=(
            FEATURE_ORDER_DIFF,
            FEATURE_FETCH_ORDERS,
            FEATURE_ADD_CARDS,
            FEATURE_ADD_B2B,
        ),
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    private_key = serialization.load_pem_private_key(
        args.private_key.read_bytes(), password=None
    )
    if not isinstance(private_key, Ed25519PrivateKey):
        raise ValueError("Private key must be an Ed25519 key")
    now = datetime.now(timezone.utc)
    expires_at = None
    if args.expires:
        expiration_date = datetime.strptime(args.expires, "%Y-%m-%d").date()
        expires_at = datetime.combine(expiration_date, time(23, 59, 59), timezone.utc)
    payload = {
        "version": 1,
        "product": PRODUCT,
        "license_id": str(uuid.uuid4()),
        "issued_at": now.isoformat().replace("+00:00", "Z"),
        "expires_at": expires_at.isoformat().replace("+00:00", "Z") if expires_at else None,
        "features": sorted(set(args.features)),
    }
    if args.customer.strip():
        payload["customer"] = args.customer.strip()
    token = sign_license(payload, private_key)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(token + "\n", encoding="utf-8")
    print(f"License written to: {args.output}")
    print(f"License ID: {payload['license_id']}")
    print(f"Expires: {payload['expires_at'] or 'never'}")
    print(f"Features: {', '.join(payload['features'])}")


if __name__ == "__main__":
    main()
