# Keep in sync with crunched-take-home/src/shared/limits.ts
MAX_CELLS = 2000
# Client loop only (loop.ts). The server is one graph step per POST /api/chat.
MAX_STEPS = 16
# Client truncation in tools.ts; do not slice tool JSON on the server.
MAX_TOOL_RESULT_CHARS = 8000
# Client trim by complete user turn. Not an API max_length: trim never splits
# the current turn, which can exceed 12 messages during a MAX_STEPS loop.
MAX_MESSAGES = 12
# Generous POST /api/chat bound: one MAX_STEPS turn with all six tools in
# parallel is ~113 messages; 256 leaves room without using MAX_MESSAGES.
MAX_REQUEST_MESSAGES = 256
