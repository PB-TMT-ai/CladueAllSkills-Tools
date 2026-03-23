# Everything Claude Code — Plugin Setup Guide

## Quick Connect (3 ways)

### Option 1: Plugin Marketplace (Recommended)
```bash
# In Claude Code, run:
/plugin marketplace add PB-TMT-ai/CladueAllSkills-Tools
/plugin install everything-claude-code@everything-claude-code
```

### Option 2: Direct settings.json
Add to your `~/.claude/settings.json`:
```json
{
  "extraKnownMarketplaces": {
    "everything-claude-code": {
      "source": {
        "source": "github",
        "repo": "PB-TMT-ai/CladueAllSkills-Tools"
      }
    }
  },
  "enabledPlugins": {
    "everything-claude-code@everything-claude-code": true
  }
}
```

### Option 3: Local Folder (No Internet)
Add to your `~/.claude/settings.json`:
```json
{
  "enabledPlugins": {
    "everything-claude-code@D:/": true
  }
}
```

## After Connecting — Install Rules Manually
Claude Code plugins cannot distribute rules automatically. Copy them:
```powershell
# Windows PowerShell
mkdir -Force $env:USERPROFILE\.claude\rules
Copy-Item -Recurse D:\rules\common\* $env:USERPROFILE\.claude\rules\
Copy-Item -Recurse D:\rules\typescript\* $env:USERPROFILE\.claude\rules\  # pick your stack
Copy-Item -Recurse D:\rules\python\* $env:USERPROFILE\.claude\rules\
```

## What You Get
| Component | Count | Examples |
|-----------|-------|---------|
| Skills | 119 | python-patterns, tdd-workflow, django-tdd, security-review |
| Agents | 28 | planner, code-reviewer, security-reviewer, tdd-guide |
| Commands | 60 | /plan, /tdd, /code-review, /e2e, /build-fix, /learn |
| Hook Events | 7 | PreToolUse, PostToolUse, Stop, SessionStart, etc. |
| Language Rules | 12 | TypeScript, Python, Go, Rust, Kotlin, Java, Swift, C++, PHP, Perl, C# |
| MCP Servers | 20+ | GitHub, Supabase, Vercel, Railway, ClickHouse, Exa |

## Verify Installation
```bash
/plugin list everything-claude-code@everything-claude-code
```
