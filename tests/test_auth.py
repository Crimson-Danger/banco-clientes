from __future__ import annotations

from sqlalchemy import create_engine, select
from sqlalchemy.orm import sessionmaker

from inss_db_app.auth import hash_password, verify_password
from inss_db_app.database import Base
from inss_db_app.models import AppUser


def test_hash_and_verify_password() -> None:
    password_hash, password_salt = hash_password("senha123")
    assert verify_password("senha123", password_hash, password_salt) is True
    assert verify_password("senha-errada", password_hash, password_salt) is False


def test_create_default_admin_like_user() -> None:
    engine = create_engine("sqlite:///:memory:", future=True)
    Base.metadata.create_all(engine)
    session = sessionmaker(bind=engine, autoflush=False, autocommit=False, future=True)()
    try:
        password_hash, password_salt = hash_password("admin123")
        session.add(
            AppUser(
                username="admin",
                full_name="Administrador",
                password_hash=password_hash,
                password_salt=password_salt,
                is_active=True,
                is_admin=True,
                can_view_dashboard=True,
                can_import=True,
                can_search=True,
                can_export=True,
                can_view_history=True,
                can_delete_batches=True,
                can_manage_users=True,
            )
        )
        session.commit()
        admin = session.scalar(select(AppUser).where(AppUser.username == "admin"))
        assert admin is not None
        assert admin.is_admin is True
        assert admin.can_manage_users is True
        assert verify_password("admin123", admin.password_hash, admin.password_salt) is True
    finally:
        session.close()
