# Installed Skills / Agents / Commands — Sources & Attribution

This directory contains third-party Claude Code **skills**, **subagents**, and **slash
commands** vendored (copied) into this repository. They were installed on request and
committed so they persist across (ephemeral) Claude Code web sessions.

All source projects are MIT licensed. Original copyright remains with their authors.

| Source repo | Version | Installed | Notes |
|-------------|---------|-----------|-------|
| [coreyhaines31/marketingskills](https://github.com/coreyhaines31/marketingskills) | 2.5.1 | 45 skills | Pure markdown skills. |
| [AgriciDaniel/claude-seo](https://github.com/AgriciDaniel/claude-seo) | 2.2.0 | 25 skills + 18 agents | Skills + agents only — see "Intentionally excluded". |
| [obra/superpowers](https://github.com/obra/superpowers) | latest | 14 skills | Skills only — `SessionStart` hook **not** wired up. |
| [luongnv89/claude-howto](https://github.com/luongnv89/claude-howto) | 2.1.160 | 6 skills + 9 agents + 8 commands | Reusable assets only (tutorial modules not copied). |

**Totals:** 90 skills · 27 agents · 8 slash commands.

## Naming collision resolved

`seo-audit` existed in **both** marketingskills and claude-seo. To keep both:

- `claude-seo`'s version keeps the canonical name **`seo-audit`** (full technical audit
  with parallel subagents).
- `marketingskills`' version was renamed to **`seo-audit-marketing`** (folder + the
  `name:` field in its `SKILL.md`). Cross-references in other marketingskills skills that
  point to `seo-audit` now resolve to claude-seo's version.

## Intentionally excluded (no auto-running / code-executing parts)

To keep the install safe and dependency-free, the following were **not** installed:

- **claude-seo**: Python `scripts/`, `extensions/` (DataForSEO, Firecrawl, etc.), MCP
  integrations, and the `PostToolUse` hook that runs a Python schema validator after every
  Edit/Write. Skills/agents that call those scripts will have limited functionality until
  the scripts + Python deps are added.
- **superpowers**: the `SessionStart` hook (`hooks/hooks.json`). The skills remain usable
  on demand.
- **claude-howto**: `06-hooks/`, `07-plugins/`, and the numbered tutorial modules
  (`02-memory`, `05-mcp`, etc.). Only `03-skills`, `04-subagents`, and `01-slash-commands`
  were copied.

## Updating

These are vendored copies, not submodules. To update, re-copy from the upstream repo and
re-apply the `seo-audit` rename.
