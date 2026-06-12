# Agent Instructions

This repository uses Sayat's shared agent-rule library as the primary source of
maintainer preferences.

Before making changes, read:

- `~/.gemini/GEMINI.md`
- `~/.gemini/rules/md_to_docx_local.md`

If those files are not available on the current machine, follow the public
maintainer rules from:

- https://github.com/Sayyat/antigravity-rules

Repository-specific summary:

- keep this project standalone;
- use `pnpm` for development, checks, builds, and releases;
- keep `markdown-docx` as the default engine and `pandoc` as optional;
- preserve deterministic academic DOCX styling;
- keep release and install documentation understandable for non-developers;
- do not move this converter into downstream document repositories.
