import os
from typing import Any

from langchain_anthropic import ChatAnthropic
from langchain_core.messages import (
    AIMessage,
    BaseMessage,
    HumanMessage,
    SystemMessage,
    ToolMessage,
)
from langgraph.graph import END, START, StateGraph
from langgraph.graph.message import add_messages
from pydantic import ValidationError
from typing_extensions import Annotated, TypedDict

from app.agent.prompt import SYSTEM_PROMPT
from app.agent.tools import TOOLS, TOOLS_BY_NAME
from app.schemas import ChatMessage, MessageResponse, StepResponse, ToolCall, ToolCallsResponse


class AgentState(TypedDict):
    messages: Annotated[list[BaseMessage], add_messages]


MODEL = ChatAnthropic(
    model=os.getenv("ANTHROPIC_MODEL", "claude-sonnet-4-5-20250929"),
    temperature=0,
    timeout=60,
).bind_tools(TOOLS)


def _content_text(content: Any) -> str:
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts: list[str] = []
        for block in content:
            if isinstance(block, str):
                parts.append(block)
            elif isinstance(block, dict) and block.get("type") == "text":
                parts.append(str(block.get("text", "")))
            elif hasattr(block, "text"):
                parts.append(str(block.text))
        return "".join(parts)
    return str(content or "")


def _normalize_args(name: str, args: dict[str, Any]) -> dict[str, Any]:
    schema = TOOLS_BY_NAME.get(name)
    if schema is None:
        return args
    try:
        return schema.model_validate(args).model_dump()
    except ValidationError:
        # Do not 400 the turn. Invalid args still go to the WebView so the
        # model gets a tool result (range_too_large, start_cell_not_a1).
        return args


def to_langchain(messages: list[ChatMessage]) -> list[BaseMessage]:
    converted: list[BaseMessage] = [SystemMessage(content=SYSTEM_PROMPT)]
    for message in messages:
        if message.role == "user":
            converted.append(HumanMessage(content=message.content))
        elif message.role == "assistant":
            tool_calls = [
                {
                    "id": call.id,
                    "name": call.name,
                    "args": call.args,
                    "type": "tool_call",
                }
                for call in (message.tool_calls or [])
            ]
            converted.append(
                AIMessage(content=message.content or "", tool_calls=tool_calls)
            )
        elif message.role == "tool":
            converted.append(
                ToolMessage(
                    content=message.content,
                    tool_call_id=message.tool_call_id,
                )
            )
    return converted


def reason(state: AgentState) -> dict[str, list[BaseMessage]]:
    response = MODEL.invoke(state["messages"])
    return {"messages": [response]}


def build_graph():
    graph = StateGraph(AgentState)
    graph.add_node("reason", reason)
    graph.add_edge(START, "reason")
    graph.add_edge("reason", END)
    return graph.compile()


GRAPH = build_graph()


def step(messages: list[ChatMessage]) -> StepResponse:
    result = GRAPH.invoke({"messages": to_langchain(messages)})
    last = result["messages"][-1]
    tool_calls = getattr(last, "tool_calls", None) or []
    if tool_calls:
        return ToolCallsResponse(
            type="tool_calls",
            tool_calls=[
                ToolCall(
                    id=call["id"],
                    name=call["name"],
                    args=_normalize_args(call["name"], call.get("args") or {}),
                )
                for call in tool_calls
            ],
        )
    return MessageResponse(type="message", content=_content_text(last.content))
