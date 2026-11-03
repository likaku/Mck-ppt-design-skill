#!/usr/bin/env bash
# One-time GitHub repo setup for discoverability. Requires: gh auth login (repo admin).
set -euo pipefail
REPO="likaku/Mck-ppt-design-skill"

gh repo edit "$REPO" \
  --description "Make your AI agent produce consulting-grade, fully editable PowerPoint. 62 layouts, native shapes, code-checked QA gates. Claude Code / Cursor / WorkBuddy. 麦肯锡风格 PPT 设计系统" \
  --homepage "https://github.com/$REPO#try-it-in-60-seconds" \
  --enable-discussions \
  --add-topic powerpoint --add-topic pptx --add-topic python-pptx \
  --add-topic ai-agent --add-topic agent-skills --add-topic claude-skills --add-topic claude-code \
  --add-topic cursor --add-topic workbuddy --add-topic codebuddy \
  --add-topic presentation --add-topic slides --add-topic mckinsey --add-topic consulting --add-topic llm

# Social preview cannot be set via API. Upload manually:
#   Settings > General > Social preview > docs/social-preview.png (1280x640)
echo "Done. Now upload docs/social-preview.png in Settings > Social preview."
