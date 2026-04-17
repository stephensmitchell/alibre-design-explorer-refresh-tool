# Alibre Design Explorer Refresh Tool

> Note: This repository is undergoing significant changes and is currently a work in progress.

| Item | Value |
| --- | --- |
| Type | Desktop utility / tool |
| Primary stack | VB.NET, Alibre automation |

## Overview
This utility connects to a running Alibre Design session, walks the current part feature list, regenerates the model, and saves the file.

## Repository Layout
- source/: project source, solution or project files, and runtime assets.
- submodules/: external git submodules used by the repository when required.
- documentation/: supplementary notes, changelogs, and non-GitHub documentation.
- .github/: repository README, templates, and GitHub-specific community files.
- `source/ADDesignExplorerRefresh.sln`: key source or build entry point.
- `source/ADDesignExplorerRefresh.vbproj`: key source or build entry point.
- `LICENSE`: repository license file kept at the root.

## Requirements
- Windows development environment.
- A .NET build environment compatible with the projects under source/.
- Alibre Design installed if you need to run, debug, or validate the Alibre integration.

## Build and Use
1. Open `source/ADDesignExplorerRefresh.sln` in your preferred IDE.
2. Restore dependencies and build from the source/ layout.
3. Use the notes in documentation/ and .github/README.md as the primary repository guide.

## Current Limitations
- The repository has been normalized for layout consistency; any path-sensitive tooling should be revalidated against the new folder structure.
- Existing runtime behavior and project-specific limitations remain unchanged.

## License
See [LICENSE](../LICENSE).

