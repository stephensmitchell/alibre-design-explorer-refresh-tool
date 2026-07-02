# Alibre Design Explorer Refresh Tool

A command-line utility that attaches to a running Alibre Design session and refreshes the active part's Design Explorer by reactivating every feature, regenerating the model, and saving the file.

The tool is a standalone VB.NET console executable. It connects to Alibre Design through the `AlibreX` COM automation interface, operates on the topmost (active) session, confirms that session is a part, sets each feature in the part's feature list to active, then calls `RegenerateAll()` and `Save()`. It targets .NET Framework 4.8.1 (Windows, x64) and was developed against Alibre Design 29.0.0.29060.

## Table Of Contents

- [What Is Here](#what-is-here)
- [Official Alibre Resources](#official-alibre-resources)
- [Requirements](#requirements)
- [Quick Start](#quick-start)
- [Build](#build)
- [Usage](#usage)
- [Key Files](#key-files)
- [Key Folders](#key-folders)
- [Notes](#notes)
- [License](#license)

## What Is Here

- A VB.NET console application (`source/Init.vb`) that drives a running Alibre Design session over COM automation.
- Logic to attach to the `AlibreX.AutomationHook` object, read the topmost session, and validate that it is a part session.
- A pass over the part's feature list that sets each `IADPartFeature.IsActive` to `True`, followed by `RegenerateAll()` and `Save()`.
- Error handling that reports when Alibre Design is not running, when the active session is not a part, and when COM or argument errors occur.
- COM object cleanup through `Marshal.ReleaseComObject` in a `Finally` block.

## Official Alibre Resources

Alibre's official resources for API development and AI/LLM/agent workflows: <https://www.alibre.com/api/>

## Requirements

- Alibre Design installed and running, with the bundled `AlibreX.dll` automation library. The project references `AlibreX.dll` from Alibre Design 29.0.0.29060.
- .NET Framework 4.8.1 runtime (Windows).
- x64 platform target.
- Visual Studio or MSBuild to build from source.

## Quick Start

1. Start Alibre Design and open a part.
2. Make that part the active (topmost) session.
3. Run `ADDesignExplorerRefresh.exe`.

The tool runs the whole pass without further input and saves the part before it exits.

## Build

1. Open `source/ADDesignExplorerRefresh.sln` in Visual Studio, or build `source/ADDesignExplorerRefresh.vbproj` with MSBuild.
2. Build the `Release` configuration. The executable is produced at `source/bin/Release/net481/ADDesignExplorerRefresh.exe`.
3. If your Alibre Design version differs from 29.0.0.29060, update the `AlibreX` reference `HintPath` in `ADDesignExplorerRefresh.vbproj` to point at your local `AlibreX.dll`.

This is a standalone console executable. There is no Alibre add-on (`.adc`) manifest to register.

## Usage

Run `ADDesignExplorerRefresh.exe` with Alibre Design open and a part active. The program prints one of:

- A confirmation that features were activated, the session regenerated, and changes saved.
- `No features found in the current session.` when the active part has no features.
- An `Operation failed`, `COM error`, `Argument error`, or `Unexpected error` message describing what went wrong, for example when Alibre Design is not running or the active session is not a part.

## Key Files

| File | Purpose |
| --- | --- |
| `source/Init.vb` | Console entry point (`Module Init`, `Sub Main`); attaches to Alibre Design, activates features, regenerates, and saves the part. |
| `source/ADDesignExplorerRefresh.sln` | Visual Studio solution. |
| `source/ADDesignExplorerRefresh.vbproj` | VB.NET project targeting `net481`, `OutputType` `Exe`, x64, with the `AlibreX` reference. |
| `source/App.config` | Application configuration. |
| `source/alibre.disclaimer.txt` | Alibre disclaimer text. |
| `LICENSE` | Repository license. |
| `.gitignore` | Ignored build output and editor files. |

## Key Folders

| Folder | Purpose |
| --- | --- |
| `source/` | VB.NET project source, solution, and build configuration. |
| `source/My Project/` | Auto-generated VB project metadata (assembly info, resources, settings). |
| `documentation/` | Placeholder for documentation (currently empty). |
| `submodules/` | Placeholder for git submodules (currently empty). |
| `reviews/` | Code review notes. |
| `.github/` | Repository metadata: this README, issue templates, contributing guide, code of conduct. |

## Notes

- The active session must be a part. Assembly and other session types are rejected with `The current session is not a valid part session.`
- The feature loop calls `Thread.Sleep` for 500 ms after activating each feature, so parts with many features take longer to process.
- The tool saves the part during the run. Confirm the file is in the state you want before running it.
- The `AlibreX` reference resolves to a machine-specific path (`C:\Program Files\Alibre Design 29.0.0.29060\Program\AlibreX.dll`); adjust it for your installation.

## License

See [LICENSE](../LICENSE).
