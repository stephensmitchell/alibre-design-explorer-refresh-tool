# Alibre Design Explorer Refresh Tool

A command-line utility that connects to a running Alibre Design session and forces the active part's Design Explorer to refresh by reactivating every feature, regenerating the model, and saving the file.

## Features
- Attaches to an already-running Alibre Design instance through its `AlibreX` COM automation interface.
- Operates on the topmost (active) part session and validates that it is a part before acting.
- Iterates the part's feature list and sets each feature to active.
- Regenerates the full model and saves the part once all features are processed.
- Reports clear status and error messages (for example, when Alibre Design is not running or the active session is not a part).

## Requirements
- Alibre Design installed and running, with the bundled `AlibreX.dll` automation library (developed against Alibre Design 29.0.0.29060).
- .NET Framework 4.8.1 runtime (Windows, x64).

## Installation
1. Open `source/ADDesignExplorerRefresh.sln` in Visual Studio (or build `source/ADDesignExplorerRefresh.vbproj` with MSBuild).
2. Build the `Release` configuration. The output is produced at `source/bin/Release/net481/ADDesignExplorerRefresh.exe`.
3. Ensure the `AlibreX` reference resolves to your local Alibre Design installation if your version differs from the one in the project file.

This is a standalone console executable; there is no Alibre add-on (`.adc`) manifest to register.

## Usage
1. Start Alibre Design and open the part you want to refresh.
2. Make that part the active (topmost) session.
3. Run `ADDesignExplorerRefresh.exe`.

The tool reactivates each feature, regenerates the model, and saves the part automatically. On success it prints a confirmation; if Alibre Design is not running or the active session is not a part, it reports the reason and exits.

## License
See [LICENSE](../LICENSE).
