from __future__ import annotations

import json

from sqlalchemy.orm import Session

from ..database import SessionLocal
from ..models import AuditLog


def log_audit(
    session: Session,
    actor_username: str,
    action: str,
    target_type: str = "",
    target_id: str = "",
    message: str = "",
    metadata: dict[str, object] | None = None,
) -> None:
    session.add(
        AuditLog(
            actor_username=actor_username or "sistema",
            action=action,
            target_type=target_type,
            target_id=str(target_id or ""),
            message=message,
            metadata_json=json.dumps(metadata or {}, ensure_ascii=False, default=str),
        )
    )


def log_audit_with_new_session(
    actor_username: str,
    action: str,
    target_type: str = "",
    target_id: str = "",
    message: str = "",
    metadata: dict[str, object] | None = None,
) -> None:
    session = SessionLocal()
    try:
        log_audit(
            session,
            actor_username,
            action,
            target_type=target_type,
            target_id=target_id,
            message=message,
            metadata=metadata,
        )
        session.commit()
    finally:
        session.close()
