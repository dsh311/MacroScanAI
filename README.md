# MacroScanAI
**AI-Powered Macro Security Analyzer**  
An AI-powered cybersecurity tool for scanning and classifying Office VBA macros. It detects malicious or suspicious behaviors such as obfuscation, encryption, and other threat-actor techniques, and includes an interactive code visualization interface with full syntax highlighting.

<p align="center">
  <img src="./docs/MacroNavigatorUI.png" width="30%" />
  <img src="./docs/ScanWithAI.png" width="30%" />
</p>

---

## 🚀 Overview

MacroScanAI is designed for security analysts, researchers, IT administrators, and developers who need a fast and accurate way to assess the risk posed by VBA macros embedded inside Office documents.

Combining static analysis with AI-enhanced pattern recognition, MacroScanAI helps identify:

- 🔍 **Malicious, Suspicious, or Benign VBA code**
- 🧬 **Obfuscation techniques** (string manipulation, hidden payloads, hex blobs, base64, etc.)
- 🔐 **Decryption or encryption routines**
- 🌐 **Network calls, shell execution, and COM automation abuse**
- 📦 **Suspicious API imports and unusual Office object usage**
- ⚠️ **Patterns commonly seen in malware-laced documents**

The tool also includes a **Code Visualization Navigator**, making it easy to explore modules, procedures, and scan results.

---

## 🖼️ Screenshots

### **Macro Navigator UI**
A clean, dark-themed interface for browsing document modules and visualizing macro code.

![Macro Navigator UI](./docs/MacroNavigatorUI.png)

---

### **AI-Powered Scan Results**
Summary verdict, indicators, and detailed module-by-module analysis.

![Scan With AI](./docs/ScanWithAI.png)

---

## ✨ Key Features

- 🤖 **AI-Assisted Macro Analysis**  
  Automatically classifies macros as *Benign*, *Suspicious*, or *Malicious* with confidence scoring.

- 🧩 **Threat Pattern Detection**  
  Detects obfuscation, payload construction, shell commands, API calls, and more.

- 🗂️ **Interactive Code Navigator**  
  Visualize modules, procedures, and scan results side-by-side.

- 📄 **Supports Office Document Formats**  
  Scans VBA included in legacy OLE (.doc, .xls, .ppt) and modern OOXML (.docm, .xlsm, .pptm) documents.

- ⚡ **Fast Local Analysis**  
  Extracts and analyzes macro content quickly.

---

## 🛠️ How It Works

1. **Load an Office document** containing VBA macros.  
2. **MacroScanAI extracts** modules and procedures using a built-in parser.  
3. **AI-powered analysis** evaluates each module’s behavior.  
4. **Results are displayed** in a structured, hyperlinked UI.  
5. You receive a **final verdict** and a list of **behavioral indicators**.

---
