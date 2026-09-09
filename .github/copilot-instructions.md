<!--
Version: 1.0.0
Last updated: 2026-09-09
Owner: GitHub Copilot
-->

# Copilot Instructions – Markdown Versioning

These instructions apply whenever a user asks you to create, edit, or update any Markdown (`.md`) file in this repository (guides, instructions, documentation, README files).

## Rules to follow automatically

1. **Never commit directly to `main`.**
   Always create or use a branch named `feature/<short-description>` for the change.
2. **Every versioned Markdown file must have a version header** at the very top:

```
<!--
Version: X.Y.Z
Last updated: YYYY-MM-DD
Owner: <name of the person making the change>
-->
```

If the file doesn't have one yet, add it starting at `1.0.0`.
3. **Use Semantic Versioning (SemVer)** — `MAJOR.MINOR.PATCH`:

- **MAJOR**: the change alters the meaning of existing content or removes something users rely on.
- **MINOR**: new section or new content added, nothing existing removed or changed.
- **PATCH**: typo fix, wording clarification, formatting only.
4. **Never decide the version bump silently.**
   After making the edit, always:

- state what changed,
- propose whether it's a MAJOR, MINOR, or PATCH change and why,
- propose the new version number,
- **wait for the user to confirm** before updating the header, updating the changelog, or committing anything.
5. **Update `CHANGELOG.md`** for every confirmed change, using this format:

```
## [X.Y.Z] - YYYY-MM-DD
### Added / Changed / Fixed
- Short description of the change
```

If `CHANGELOG.md` doesn't exist yet, create it.
6. **Commit message format**: `type(scope): description`

- `docs` for Markdown/documentation changes
- `feat` for new functionality elsewhere in the repo
- `fix` for bug fixes
- `ci` for automation/workflow changes

Example: `docs(readme): clarified setup steps for new team members`
7. **When asked to tag a release**, use:

```
git tag -a vX.Y.Z -m "short description"
git push origin vX.Y.Z
```

Tags must be pushed separately — they are not sent automatically with a branch.
8. **When asked for history**, summarize commits/changelog entries in plain, non-technical language, grouped by version number. Do not just paste raw `git log` output unless explicitly asked for it.

## Golden rule

**One change = one version bump = one CHANGELOG entry.**
Always let the human make the final call on MAJOR vs. MINOR vs. PATCH — propose it, explain your reasoning, but never assume.
