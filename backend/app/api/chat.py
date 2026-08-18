import os

from anthropic import APIError, AuthenticationError, RateLimitError
from fastapi import APIRouter, HTTPException

from app.agent.graph import step
from app.schemas import StepRequest, StepResponse

router = APIRouter()


@router.get("/health")
def health() -> dict[str, str]:
    if not os.getenv("ANTHROPIC_API_KEY"):
        raise HTTPException(status_code=503, detail="ANTHROPIC_API_KEY is missing")
    return {"status": "ok"}


@router.post("/chat", response_model=StepResponse)
def chat(request: StepRequest) -> StepResponse:
    try:
        return step(request.messages)
    except AuthenticationError as exc:
        raise HTTPException(
            status_code=503, detail=f"Invalid API key ({type(exc).__name__})"
        ) from exc
    except RateLimitError as exc:
        raise HTTPException(
            status_code=429, detail=f"Rate limited ({type(exc).__name__})"
        ) from exc
    except APIError as exc:
        raise HTTPException(status_code=502, detail=type(exc).__name__) from exc
    except Exception as exc:  # noqa: BLE001 — class name only; no traceback in the pane
        if "api key" in str(exc).lower():
            raise HTTPException(
                status_code=503, detail=f"Invalid API key ({type(exc).__name__})"
            ) from exc
        raise HTTPException(status_code=500, detail=type(exc).__name__) from exc
