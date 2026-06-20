# Installed Skills, Agents & Commands — Sources

This directory contains third-party Claude Code skills, subagents, and slash
commands installed into the project. They are version-controlled here so they
persist across the remote/ephemeral Claude Code sessions used for this repo.

All source projects are **MIT licensed**. See each upstream repo for full text.

| Source repo | Version | What was installed | Excluded |
|-------------|---------|--------------------|----------|
| [coreyhaines31/marketingskills](https://github.com/coreyhaines31/marketingskills) | 2.5.1 | 45 skills | — |
| [AgriciDaniel/claude-seo](https://github.com/AgriciDaniel/claude-seo) | 2.2.0 | 25 skills + 18 agents | Python scripts, MCP extensions, and the auto-running `PostToolUse` validation hook |
| [obra/superpowers](https://github.com/obra/superpowers) | — | 14 skills | `SessionStart` hook (not auto-enabled) |
| [luongnv89/claude-howto](https://github.com/luongnv89/claude-howto) | 2.1.x | 6 skills + 9 subagents + 8 slash commands | tutorial modules, `06-hooks/`, `07-plugins/` |

**Totals:** 90 skills · 27 agents · 8 commands

## Notes

- **`seo-audit` name collision:** both `marketingskills` and `claude-seo` ship a
  skill named `seo-audit`. The dedicated SEO suite keeps the canonical name
  (`seo-audit`); the marketing one was renamed to **`seo-audit-marketing`**
  (folder + frontmatter `name`) so both remain usable.
- **No auto-running hooks were installed.** Hooks from `claude-seo`,
  `superpowers`, and `claude-howto` execute code automatically (on edits or
  session start) and require extra Python/Node dependencies, so they were left
  out. The skills/agents themselves are invoked on demand.
- **`claude-seo` script-dependent steps are limited** because the helper Python
  scripts were not installed. Re-run with the scripts if you need full
  functionality.

## Updating

Re-clone the upstream repo and re-copy the relevant directories into
`.claude/skills`, `.claude/agents`, or `.claude/commands`.
