from unittest.mock import patch

import httpx
import pytest
from anthropic import APIStatusError
from fastapi.testclient import TestClient

from app.main import app
from app.schemas import MessageResponse

client = TestClient(app)


def _api_status_error(message: str) -> APIStatusError:
    request = httpx.Request("POST", "https://api.anthropic.com/v1/messages")
    response = httpx.Response(500, request=request)
    return APIStatusError(message, response=response, body=None)


def test_chat_returns_mocked_step() -> None:
    with patch(
        "app.api.chat.step",
        return_value=MessageResponse(type="message", content="done"),
    ):
        response = client.post(
            "/api/chat", json={"messages": [{"role": "user", "content": "hi"}]}
        )
    assert response.status_code == 200
    assert response.json() == {"type": "message", "content": "done"}


def test_chat_provider_error_is_502_without_traceback() -> None:
    with patch(
        "app.api.chat.step",
        side_effect=_api_status_error(
            "Traceback (most recent call last):\n  File secret.py"
        ),
    ):
        response = client.post(
            "/api/chat", json={"messages": [{"role": "user", "content": "hi"}]}
        )
    assert response.status_code == 502
    body = response.text
    assert "Traceback" not in body
    assert "secret.py" not in body
    assert "APIStatusError" in body


def test_chat_unexpected_error_is_500_without_traceback() -> None:
    with patch(
        "app.api.chat.step",
        side_effect=RuntimeError(
            "Traceback (most recent call last):\n  File secret.py"
        ),
    ):
        response = client.post(
            "/api/chat", json={"messages": [{"role": "user", "content": "hi"}]}
        )
    assert response.status_code == 500
    body = response.text
    assert "Traceback" not in body
    assert "secret.py" not in body
    assert "RuntimeError" in body


def test_chat_rejects_user_tool_calls() -> None:
    response = client.post(
        "/api/chat",
        json={
            "messages": [
                {
                    "role": "user",
                    "content": "hi",
                    "tool_calls": [{"id": "1", "name": "get_selection", "args": {}}],
                }
            ]
        },
    )
    assert response.status_code == 400


def test_health_ok_when_key_set(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("ANTHROPIC_API_KEY", "test-key")
    response = client.get("/api/health")
    assert response.status_code == 200
    assert response.json() == {"status": "ok"}


def test_health_503_when_key_missing(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)
    response = client.get("/api/health")
    assert response.status_code == 503
    assert "ANTHROPIC_API_KEY" in response.json()["detail"]
