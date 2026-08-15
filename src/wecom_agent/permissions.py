from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
import hashlib
import hmac
import secrets


class PermissionDenied(RuntimeError):
    """Raised when execution is requested without approval for the exact plan."""


@dataclass(frozen=True)
class Approval:
    task_id: str
    plan_hash: str
    token: str
    issued_at: str
    expires_at: str

    @property
    def token_hash(self) -> str:
        return hashlib.sha256(self.token.encode("utf-8")).hexdigest()


class ApprovalManager:
    def __init__(self, secret: str | None = None, default_ttl_seconds: int = 600) -> None:
        self._secret = (secret or secrets.token_hex(32)).encode("utf-8")
        self.default_ttl_seconds = default_ttl_seconds

    def issue(
        self,
        task_id: str,
        plan_hash: str,
        *,
        now: datetime | None = None,
        ttl_seconds: int | None = None,
    ) -> Approval:
        now = _utc(now)
        expires = now + timedelta(seconds=ttl_seconds or self.default_ttl_seconds)
        nonce = secrets.token_urlsafe(18)
        payload = "|".join([task_id, plan_hash, str(int(expires.timestamp())), nonce])
        signature = hmac.new(self._secret, payload.encode("utf-8"), hashlib.sha256).hexdigest()
        return Approval(
            task_id=task_id,
            plan_hash=plan_hash,
            token=f"{payload}|{signature}",
            issued_at=now.isoformat(),
            expires_at=expires.isoformat(),
        )

    def verify(
        self,
        task_id: str,
        plan_hash: str,
        token: str,
        *,
        now: datetime | None = None,
    ) -> Approval:
        parts = token.split("|")
        if len(parts) != 5:
            raise PermissionDenied("Approval token is malformed")
        token_task, token_plan, expires_text, _nonce, signature = parts
        payload = "|".join(parts[:-1])
        expected = hmac.new(self._secret, payload.encode("utf-8"), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(signature, expected):
            raise PermissionDenied("Approval token is invalid")
        if token_task != task_id or token_plan != plan_hash:
            raise PermissionDenied("Approval does not match this exact task and plan")
        try:
            expires = datetime.fromtimestamp(int(expires_text), tz=timezone.utc)
        except (ValueError, OverflowError) as exc:
            raise PermissionDenied("Approval expiry is invalid") from exc
        now = _utc(now)
        if now >= expires:
            raise PermissionDenied("Approval has expired")
        return Approval(
            task_id=task_id,
            plan_hash=plan_hash,
            token=token,
            issued_at="",
            expires_at=expires.isoformat(),
        )


def _utc(value: datetime | None = None) -> datetime:
    value = value or datetime.now(timezone.utc)
    return value if value.tzinfo else value.replace(tzinfo=timezone.utc)
