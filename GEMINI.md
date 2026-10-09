# Tool Discipline
- NEVER use terminal commands (Select-String, Get-Content, grep, cat, sed) to read, search, or edit files inside this workspace.
- ALWAYS use your native internal workspace tools (file read, edit, search, glob) for all in-folder inspection and modifications.
- Reserve terminal execution strictly for project runtime tasks (e.g., executing scripts, running dev servers).