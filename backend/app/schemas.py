from __future__ import annotations

from typing import Annotated, Any, Literal

from pydantic import BaseModel, ConfigDict, Field

from app.limits import MAX_REQUEST_MESSAGES


class ToolCall(BaseModel):
    id: str
    name: str
    args: dict[str, Any] = Field(default_factory=dict)


class UserMessage(BaseModel):
    model_config = ConfigDict(extra="forbid")

    role: Literal["user"]
    content: str


class AssistantMessage(BaseModel):
    model_config = ConfigDict(extra="forbid")

    role: Literal["assistant"]
    content: str | None = None
    tool_calls: list[ToolCall] | None = None


class ToolResultMessage(BaseModel):
    model_config = ConfigDict(extra="forbid")

    role: Literal["tool"]
    tool_call_id: str
    content: str


ChatMessage = Annotated[
    UserMessage | AssistantMessage | ToolResultMessage,
    Field(discriminator="role"),
]


class StepRequest(BaseModel):
    messages: list[ChatMessage] = Field(max_length=MAX_REQUEST_MESSAGES)


class MessageResponse(BaseModel):
    type: Literal["message"]
    content: str


class ToolCallsResponse(BaseModel):
    type: Literal["tool_calls"]
    tool_calls: list[ToolCall]


StepResponse = Annotated[
    MessageResponse | ToolCallsResponse,
    Field(discriminator="type"),
]
