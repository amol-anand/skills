---
name: slicc-handoff
description: >-
  Hand off tasks to SLICC agent. Use when the user says things like
  "handoff to slicc", "move this to slicc", "move to the browser",
  "test in the browser", "handoff to browser", or asks you to continue the work
  in the SLICC agent.
allowed-tools: bash
---

# slicc-handoff

Hand off tasks to the SLICC browser agent by generating a handoff URL with a base64url-encoded JSON payload.

## What to do

1. Build a compact JSON payload with:
   - `instruction` (required)
   - optional `title`, `urls`, `context`, `acceptanceCriteria`, `notes`
2. Run the script with `--open` to deliver the handoff:
   - `slicc-handoff --open payload.json`
   - or pipe JSON directly: `cat payload.json | slicc-handoff --open`
3. Tell the user that SLICC should show an `Accept` / `Dismiss` prompt in the Chat tab.

## Payload guidance

- Keep `instruction` direct and action-oriented.
- Use `urls` only when specific pages matter.
- Keep `acceptanceCriteria` short and concrete.
- Put supporting detail in `context` or `notes`; do not dump full logs unless they are necessary.

## Examples

```bash
slicc-handoff --open payload.json
```

```bash
cat payload.json | slicc-handoff --open
```
