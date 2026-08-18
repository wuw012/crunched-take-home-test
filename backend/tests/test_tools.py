from unittest.mock import patch

from langchain_core.messages import AIMessage
from pydantic import TypeAdapter, ValidationError
import pytest

from app.agent.graph import step, to_langchain
from app.agent.tools import create_chart, write_range
from app.limits import MAX_CELLS
from app.schemas import AssistantMessage, ChatMessage, ToolResultMessage, UserMessage

ChatMessageAdapter = TypeAdapter(ChatMessage)


def test_write_range_rejects_oversize() -> None:
    values = [[0] * 50 for _ in range(50)]
    assert 50 * 50 > MAX_CELLS
    with pytest.raises(ValidationError):
        write_range(sheet="P&L", start_cell="A1", values=values)


def test_write_range_rejects_range_start_cell() -> None:
    with pytest.raises(ValidationError):
        write_range(sheet="P&L", start_cell="A1:B2", values=[["=B2-B3"]])


def test_write_range_accepts_formulas() -> None:
    payload = write_range(
        sheet="P&L",
        start_cell="A5",
        values=[["Gross Profit", "=B2-B3"]],
    )
    assert payload.values[0][1].startswith("=")


def test_create_chart_rejects_unknown_type() -> None:
    with pytest.raises(ValidationError):
        create_chart(sheet="P&L", source_a1="A1:B6", chart_type="radar")


def test_user_message_rejects_tool_calls() -> None:
    with pytest.raises(ValidationError):
        ChatMessageAdapter.validate_python(
            {
                "role": "user",
                "content": "fix the total",
                "tool_calls": [{"id": "1", "name": "get_selection", "args": {}}],
            }
        )


def test_tool_message_requires_tool_call_id() -> None:
    with pytest.raises(ValidationError):
        ChatMessageAdapter.validate_python({"role": "tool", "content": "{}"})


def test_to_langchain_preserves_tool_loop() -> None:
    messages = [
        UserMessage(role="user", content="fix the total"),
        AssistantMessage(
            role="assistant",
            content=None,
            tool_calls=[{"id": "1", "name": "get_selection", "args": {}}],
        ),
        ToolResultMessage(role="tool", tool_call_id="1", content='{"sheet":"P&L"}'),
    ]
    converted = to_langchain(messages)
    assert converted[1].type == "human"
    assert converted[2].tool_calls[0]["name"] == "get_selection"
    assert converted[3].type == "tool"
    assert converted[3].tool_call_id == "1"


def test_step_strips_extra_write_range_keys() -> None:
    ai = AIMessage(
        content="",
        tool_calls=[
            {
                "id": "1",
                "name": "write_range",
                "args": {
                    "sheet": "P&L",
                    "start_cell": "A5",
                    "values": [["=B2-B3"]],
                    "extra": "drop-me",
                },
            }
        ],
    )
    with patch("app.agent.graph.GRAPH.invoke", return_value={"messages": [ai]}):
        result = step([UserMessage(role="user", content="fix GP")])
    assert result.type == "tool_calls"
    assert result.tool_calls[0].args["values"] == [["=B2-B3"]]
    assert "extra" not in result.tool_calls[0].args


def test_step_invalid_start_cell_still_returns_tool_calls() -> None:
    args = {"sheet": "P&L", "start_cell": "A1:B2", "values": [["=B2-B3"]]}
    ai = AIMessage(
        content="",
        tool_calls=[{"id": "1", "name": "write_range", "args": args}],
    )
    with patch("app.agent.graph.GRAPH.invoke", return_value={"messages": [ai]}):
        result = step([UserMessage(role="user", content="fix GP")])
    assert result.type == "tool_calls"
    assert result.tool_calls[0].args["start_cell"] == "A1:B2"


def test_step_oversize_write_range_still_returns_tool_calls() -> None:
    """Invalid tool args must not 400 the step. The WebView refuses and the model sees the error."""
    values = [[0] * 50 for _ in range(50)]
    args = {"sheet": "P&L", "start_cell": "A1", "values": values}
    ai = AIMessage(
        content="",
        tool_calls=[{"id": "1", "name": "write_range", "args": args}],
    )
    with patch("app.agent.graph.GRAPH.invoke", return_value={"messages": [ai]}):
        result = step([UserMessage(role="user", content="fill the sheet")])
    assert result.type == "tool_calls"
    assert result.tool_calls[0].args == args
