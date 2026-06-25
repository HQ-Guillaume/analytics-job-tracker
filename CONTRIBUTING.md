# Contributing

Thanks for improving Custom Job Tracker.

This repository is intended to stay safe for public release while keeping local
job-search data private. Contributions should improve crawler reliability,
profile configuration, workbook output, filtering, tests, or documentation.

## Guidelines

- Do not commit personal job-search data, credentials, cookies, local profile
  files, generated workbooks, caches, backups, or diagnostics.
- Keep the public release profile-neutral. User-specific search profiles should
  remain local and ignored by Git.
- Preserve Windows support for the GUI launchers and PowerShell workflows.
- Update documentation when changing setup, profile behavior, output columns, or
  filtering rules.
- Add or update tests for crawler, scoring, filtering, or workbook behavior when
  practical.

## Pull Requests

Before opening a pull request:

- Run relevant PowerShell checks or tests.
- Explain which job-tracking workflow the change improves.
- Call out any new network source, dependency, permissions need, or data-retention
  implication.
