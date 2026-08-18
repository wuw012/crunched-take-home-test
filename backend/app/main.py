from fastapi import FastAPI
from fastapi.exceptions import RequestValidationError
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from app.api.chat import router as chat_router

app = FastAPI(title="Crunched take-home agent")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://localhost:3000", "https://127.0.0.1:3000"],
    allow_methods=["*"],
    allow_headers=["*"],
)
app.include_router(chat_router, prefix="/api")


@app.exception_handler(RequestValidationError)
async def request_validation_handler(
    _request: object, exc: RequestValidationError
) -> JSONResponse:
    return JSONResponse(status_code=400, content={"detail": exc.errors()})
