# -*- coding: utf-8 -*-
"""사용자 피드백 도메인 — 게시글 목록/작성/첨부 다운로드.

`대조 서버.py` 에서 `register(app, ...)` 로 등록. 저장 경로 전역만 주입받고
나머지 헬퍼(_clean_feedback_text 등)는 이 도메인 전용이라 register 안에 동봉해
주입받은 경로를 클로저로 참조한다. 라우트 본문·엔드포인트명은 원본 그대로.
"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from flask import jsonify, request, send_file
from werkzeug.utils import secure_filename


def register(app, *, FEEDBACK_POSTS_PATH, FEEDBACK_UPLOAD_DIR):

    def _load_feedback_posts() -> list[dict]:
        if not FEEDBACK_POSTS_PATH.exists():
            return []
        try:
            with FEEDBACK_POSTS_PATH.open("r", encoding="utf-8") as fp:
                payload = json.load(fp)
        except (json.JSONDecodeError, OSError):
            return []
        if isinstance(payload, dict):
            posts = payload.get("posts", [])
        else:
            posts = payload
        if not isinstance(posts, list):
            return []
        return sorted(posts, key=lambda item: str(item.get("created_at", "")), reverse=True)

    def _save_feedback_posts(posts: list[dict]) -> None:
        FEEDBACK_POSTS_PATH.parent.mkdir(parents=True, exist_ok=True)
        ordered = sorted(posts, key=lambda item: str(item.get("created_at", "")), reverse=True)
        with FEEDBACK_POSTS_PATH.open("w", encoding="utf-8") as fp:
            json.dump({"posts": ordered}, fp, ensure_ascii=False, indent=2)

    def _clean_feedback_text(value: object, limit: int) -> str:
        text = str(value or "").strip()
        text = " ".join(text.split()) if limit <= 80 else text
        return text[:limit]

    def _save_feedback_attachment(post_id: str) -> dict | None:
        uploaded = request.files.get("attachment")
        if uploaded is None or not uploaded.filename:
            return None
        original_name = Path(uploaded.filename).name
        safe_name = secure_filename(original_name)
        if not safe_name:
            safe_name = f"attachment_{post_id}"
        saved_name = f"{post_id}_{safe_name}"
        saved_path = FEEDBACK_UPLOAD_DIR / saved_name
        uploaded.save(saved_path)
        return {
            "original_name": original_name,
            "stored_name": saved_name,
            "size": saved_path.stat().st_size if saved_path.exists() else 0,
            "download_url": f"/api/feedback-attachments/{saved_name}",
        }

    @app.get("/api/feedback-posts")
    def feedback_posts():
        return jsonify({"ok": True, "posts": _load_feedback_posts()})

    @app.post("/api/feedback-posts")
    def create_feedback_post():
        if request.content_type and request.content_type.startswith("multipart/form-data"):
            source = request.form
        else:
            source = request.get_json(silent=True) or {}
        author = _clean_feedback_text(source.get("author") or "익명", 40) or "익명"
        title = _clean_feedback_text(source.get("title"), 80)
        body = str(source.get("body") or "").strip()[:3000]

        if not title:
            return jsonify({"ok": False, "message": "제목을 입력해주세요."}), 400
        if not body:
            return jsonify({"ok": False, "message": "개선의견 내용을 입력해주세요."}), 400

        posts = _load_feedback_posts()
        created_at = datetime.now().strftime("%Y-%m-%d %H:%M")
        post_id = datetime.now().strftime("%Y%m%d%H%M%S%f")
        attachment = _save_feedback_attachment(post_id)
        post = {
            "id": post_id,
            "author": author,
            "title": title,
            "body": body,
            "created_at": created_at,
            "attachment": attachment,
        }
        posts.insert(0, post)
        _save_feedback_posts(posts[:300])
        return jsonify({"ok": True, "message": "개선의견이 등록되었습니다.", "post": post})

    @app.get("/api/feedback-attachments/<path:filename>")
    def download_feedback_attachment(filename: str):
        safe_name = Path(filename).name
        target = FEEDBACK_UPLOAD_DIR / safe_name
        if not target.exists() or not target.is_file():
            return jsonify({"ok": False, "message": "첨부파일을 찾을 수 없습니다."}), 404
        return send_file(target, as_attachment=True)
