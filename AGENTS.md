# Instructions for AI Agents

This repository uses `.ai/` as shared operational memory for all AI agents.

Claude should start from `CLAUDE.md`.
Other agents should start from this file and then read `.ai/INDEX.md`.

## Required first read

1. `.ai/INDEX.md`
2. `.ai/CURRENT_STATE.md`
3. `.ai/TASK_HANDOFF.md`

## General rules

- Verify repository facts before making changes.
- Prefer minimal, safe, reviewable changes.
- Do not perform unrelated refactors.
- Do not add dependencies without justification.
- Do not expose secrets.
- Update `.ai/` files after important changes.
