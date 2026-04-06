# Copilot Migration Guide

This document explains how to adapt the slide-deck skills for GitHub Copilot or other AI coding assistants.

## What Stays the Same

The following components are agent-agnostic and require no changes:

- **`.claude/scripts/`** — All Python and Node.js scripts work identically regardless of agent
- **`.claude/prompts/`** — All prompt rule files are plain markdown, usable by any LLM
- **`.claude/docs/`** — API reference and documentation are agent-agnostic
- **`slide-workspace/`** — Runtime workspace directory convention

## What Changes Per Agent

### Claude Code

Skills are placed in `.claude/skills/{name}/SKILL.md` and are automatically picked up via the YAML `triggers` field.

### GitHub Copilot

1. **Instruction file location**: Copy skill content to `.github/copilot-instructions.md`

   Copilot uses a single instruction file for workspace context. You can append all three skills to it, separated by headers:

   ```markdown
   ## Skill: Analyze PPTX Template

   When asked to analyze a PPTX template or extract template style...
   [paste content of slide-analyze/SKILL.md without the YAML header]

   ## Skill: Generate Presentation

   When asked to generate a presentation...
   [paste content of slide-generate/SKILL.md without the YAML header]

   ## Skill: Edit Presentation

   When asked to edit a presentation...
   [paste content of slide-edit/SKILL.md without the YAML header]
   ```

2. **Trigger mapping**: The YAML `triggers` array maps to keywords you can put in your chat prompt to Copilot:
   - `analyze template` → "Please analyze this template"
   - `generate presentation` → "Generate a presentation about..."
   - `edit presentation` → "Edit this presentation to..."

3. **YAML header**: Remove the YAML frontmatter block (between `---` markers) — Copilot does not use it.

4. **Directory references**: All path references in the skill workflows use relative paths from the project root. These work identically in Copilot.

### Other Agents (Cursor, Aider, etc.)

- **Cursor**: Place skill content in `.cursorrules` or `.cursor/rules/`
- **Aider**: Use `--system-prompt` argument or `.aider.conf.yml`
- **Generic**: Any agent that accepts a system prompt can use the `.claude/prompts/` files directly

## Step-by-Step Migration to Copilot

1. Create `.github/copilot-instructions.md` if it doesn't exist

2. For each skill in `.claude/skills/*/SKILL.md`:
   - Strip the YAML frontmatter (everything between the `---` markers)
   - Add a markdown header section: `## When to use this skill`
   - List the trigger phrases from the `triggers:` field

3. Append the three adapted skill sections to `.github/copilot-instructions.md`

4. The scripts, prompts, and docs directories require no changes

5. Test by opening Copilot Chat and typing one of the trigger phrases

## Trigger Phrase Reference

| Skill | Triggers (use in chat) |
|-------|------------------------|
| slide-analyze | "analyze template", "analyze pptx", "extract template style" |
| slide-generate | "generate presentation", "create presentation from template" |
| slide-edit | "edit presentation", "modify presentation", "update slides" |

## Dependency Installation

All dependency scripts work cross-agent:

```bash
# Unix/macOS
bash .claude/scripts/setup_deps.sh

# Windows
powershell -ExecutionPolicy Bypass -File .claude/scripts/setup_deps.ps1
```

These scripts install python-pptx, Pillow, and pptxgenjs — no Docker required.
