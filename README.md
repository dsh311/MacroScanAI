# MacroScanAI
**AI-Powered Macro Security Analyzer**  
MacroScanAI is a cybersecurity tool for identifying malicious Office VBA macros. It offers both a dark-themed, interactive UI for code visualization and forensic triage, and a command-line interface for automated analysis and JSON reporting.

**Note:** AI models can make mistakes, and new malware techniques may not be recognized immediately.

---

## User Interface (UI)

MacroScanAI provides a clean, dark-themed interface for exploring Office documents containing VBA macros.

### Screenshots

<p align="center">
  <img src="./docs/MacroNavigatorUI.png" width="45%" />
  <img src="./docs/ScanWithAI.png" width="45%" />
</p>

### Features

- **Interactive Macro Navigator**: Browse modules and procedures in your Office document.  
- **AI-Powered Scan Results**: Final verdicts, confidence scores, and module-by-module analysis.  
- **Dark-Themed Layout**: Easy on the eyes for long analysis sessions.  
- **Supports Office Formats**: Legacy OLE (.doc, .xls, .ppt) and modern OOXML (.docm, .xlsm, .pptm) documents.

### How it Works

1. Open an Office document containing VBA macros.  
2. MacroScanAI parses modules and procedures.  
3. Click "Scan with AI" to analyze each module’s behavior.  
4. View results in the UI with a final verdict and list of behavioral indicators.

---

## Command-Line Interface

For automation, MacroScanAI provides a CLI tool: **MacroScanAI.Cli.exe**.

### Basic Syntax

```
MacroScanAI.Cli.exe <SourceFilePath> <OutputDirectory> [--apikey <YourOpenAIKey>]
```

- `<SourceFilePath>`: Path to the Office file to scan (`.xls`, `.doc`, etc.)  
- `<OutputDirectory>`: Directory where the JSON report will be saved  
- `--apikey <YourOpenAIKey>` (optional): Your OpenAI API key. If omitted, uses `OPENAI_API_KEY` environment variable.

### Passing the API Key

1. **Environment variable**

```cmd
setx OPENAI_API_KEY "sk-fakeapikey1234567890"
MacroScanAI.Cli.exe "C:\myClientSocket.xls" "C:\TestResults"
```

2. **Directly on the command line**

```cmd
MacroScanAI.Cli.exe "C:\myClientSocket.xls" "C:\TestResults" --apikey sk-fakeapikey1234567890
```

### Output

MacroScanAI generates a JSON report named after the input file with the suffix `.report.json`.

Example:

```
Input: myClientSocket.xls
Output: myClientSocket.xls.report.json
```

### Example JSON Report

```json
{
  "SourceFilePath": "C:\\myClientSocket.xls",
  "ReportGeneratedAt": "2025-11-21T21:31:02.0193615Z",
  "AllCodeVerdict": "Suspicious",
  "Modules": [
    {
      "moduleName": "Module1.bas",
      "verdict": "Benign",
      "confidence": 80,
      "summary": "The VBA module contains functions related to socket communication and data processing. It does not exhibit typical indicators of malicious behavior.",
      "indicators": []
    },
    {
      "moduleName": "UserForm1.frm",
      "verdict": "Suspicious",
      "confidence": 85,
      "summary": "Potential communication over a network with unclear intentions.",
      "indicators": [
        "Communication with remote host",
        "Socket operations"
      ]
    }
  ]
}
```

**Field Descriptions**

- `SourceFilePath`: Path to the scanned file  
- `ReportGeneratedAt`: Timestamp when the report was generated  
- `AllCodeVerdict`: Overall verdict (`Benign`, `Suspicious`, `Malicious`)  
- `Modules`: Per-module results with summary, confidence, and behavioral indicators

---

## Disclaimer

MacroScanAI uses AI-powered analysis to classify VBA macros as **Benign**, **Suspicious**, or **Malicious**.  

- Results are **advisory only** and should **not** be considered a definitive guarantee of safety or threat.  
- Users are responsible for performing further analysis and exercising judgment based on the context of the scanned document.  
- AI models can make mistakes, and new malware techniques may not be recognized immediately.

---

## License

This project is licensed under **GNU GPL v3.0**.  
It also includes code under the **Mozilla Public License 2.0**, redistributed under GPL-3.0 for combined distribution.

- GPL-3.0: [https://www.gnu.org/licenses/gpl-3.0.html](https://www.gnu.org/licenses/gpl-3.0.html)  
- MPL-2.0: [https://www.mozilla.org/MPL/2.0/](https://www.mozilla.org/MPL/2.0/)

![GPLv3 License](https://img.shields.io/badge/License-GPLv3-blue.svg)
