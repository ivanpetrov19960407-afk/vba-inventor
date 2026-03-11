# AGENTS.md

## Purpose
This repository contains source-only VBA modules for Autodesk Inventor drawing automation.

## Working rules for agents
- Be honest about uncertainty; do not fabricate Autodesk Inventor API names, methods, enums, or behaviors.
- Prefer official Autodesk Inventor API documentation when verifying behavior.
- User-facing documentation must be written in Russian.
- Keep VBA module comments and identifiers ASCII/English where practical to reduce `.bas` import/encoding issues.
- If Cyrillic text is used inside `.bas` and may be encoding-fragile, document the risk and fallback approach in project docs.
- Keep implementation modular and import-friendly for Inventor VBA.
- Do not merge all logic into a single module.
- Current scope is only border frame + prompted title block; do not add top statement tables or lower-left technical paragraph block.
