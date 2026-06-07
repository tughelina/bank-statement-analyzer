# Installing the Bank Statement Analyzer Skill (Claude Code)

This is the **power-user version** for people who use **Claude Code** (the terminal/desktop developer version of Claude). If you just use Claude or ChatGPT in the browser, use the **Universal Prompt** instead — it needs no installation.

## Requirements
- **Claude Code** installed and working.
- **Python 3** with `openpyxl`:  `pip3 install openpyxl`

## Install (1 minute)
1. Copy the whole **`bank-statement-analyzer`** folder into your Claude Code skills directory:
   - For all your projects:  `~/.claude/skills/bank-statement-analyzer`
   - For one project only:   `<your-project>/.claude/skills/bank-statement-analyzer`
2. Restart Claude Code (or start a new session).

## Use
Just ask, in plain language:
> "Analyze my bank statements in the folder `~/Statements 2025` and build the Excel overview."

Claude will recognise the skill, look at your files, ask you a few quick questions about your accounts and categories, and build a clean 6-sheet Excel workbook right next to your statements.

You can also run the included **`template_analysis.py`** directly once you've set the CONFIG block at its top:
```
python3 template_analysis.py "/path/to/your/statements/folder"
```

## Tip
Export your statements as **CSV** from your bank for the cleanest result. The skill also handles Excel and PDF, but CSV is the smoothest.

---
*© 2026 SoulStrategy by Chantal Perrinjaquet — House Of Coaching. Shared as a gift; please keep this credit.*
