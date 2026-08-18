# Enforced on write_range tool schemas. Client copy: crunched-take-home/src/shared/limits.ts
MAX_CELLS = 2000
# POST /api/chat bound. Client trim (MAX_MESSAGES=12 in loop.ts) is separate —
# a live turn can exceed 12 messages during the 16-step tool loop.
MAX_REQUEST_MESSAGES = 256
