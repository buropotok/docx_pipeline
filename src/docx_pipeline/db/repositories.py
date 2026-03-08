from __future__ import annotations

from typing import Optional

from docx_pipeline.db.db import get_connection


def get_document_by_uid(uid: str) -> Optional[dict]:
    conn = get_connection()
    try:
        row = conn.execute(
            """
            SELECT *
            FROM document
            WHERE uid = ?
            """,
            (uid,),
        ).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def get_document_by_mail_uid_and_attachment_index(
    mail_uid: str,
    attachment_index: int,
) -> Optional[dict]:
    conn = get_connection()
    try:
        row = conn.execute(
            """
            SELECT *
            FROM document
            WHERE mail_uid = ? AND attachment_index = ?
            """,
            (mail_uid, attachment_index),
        ).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def insert_document(
    uid: str,
    mail_uid: str,
    attachment_index: int,
    source_filename: str,
    source_ext: str,
    source_abs_path: str,
    artifacts_abs_path: str,
    processing_status: str = "discovered",
) -> int:
    conn = get_connection()
    try:
        cur = conn.execute(
            """
            INSERT INTO document (
                uid,
                mail_uid,
                attachment_index,
                source_filename,
                source_ext,
                source_abs_path,
                artifacts_abs_path,
                processing_status
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                uid,
                mail_uid,
                attachment_index,
                source_filename,
                source_ext,
                source_abs_path,
                artifacts_abs_path,
                processing_status,
            ),
        )
        conn.commit()
        return int(cur.lastrowid)
    finally:
        conn.close()


def update_document_status(uid: str, processing_status: str) -> None:
    conn = get_connection()
    try:
        conn.execute(
            """
            UPDATE document
            SET processing_status = ?,
                update_date = datetime('now')
            WHERE uid = ?
            """,
            (processing_status, uid),
        )
        conn.commit()
    finally:
        conn.close()


def list_documents() -> list[dict]:
    conn = get_connection()
    try:
        rows = conn.execute(
            """
            SELECT *
            FROM document
            ORDER BY id
            """
        ).fetchall()
        return [dict(row) for row in rows]
    finally:
        conn.close()
