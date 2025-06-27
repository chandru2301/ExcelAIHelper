# Excel AI Helper
.
<div align="center">
  <img src="assets/logo-filled.png" alt="Excel AI Helper Logo" width="200">
  <br>
  <h3>AI-powered assistance for Microsoft Excel</h3>
  
  ![License](https://img.shields.io/badge/license-MIT-blue)
  ![Platform](https://img.shields.io/badge/platform-Office%20Add--in-green)
  ![Status](https://img.shields.io/badge/status-active-brightgreen)
</div>

## 🚀 Features

- **AI Analysis**: Get intelligent insights from your Excel data
- **Dual-mode Interface**: Choose between chat or cell-based interactions
- **Smart Column Creation**: Add relevant columns with AI-generated data
- **Hovering Chat**: Access AI assistance without disrupting your workflow
- **Conversation History**: Continue complex analyses over multiple interactions

## 📋 Overview

Excel AI Helper is an Office Add-in that brings the power of AI to your Excel spreadsheets. It uses natural language processing to understand your requests and helps you analyze data, create new columns, and gain insights without complex formulas or manual data manipulation.

<div align="center">
  <img src="assets/screenshot.png" alt="Excel AI Helper Screenshot" width="600">
</div>

## 🛠️ Technology Stack

- **Frontend**: TypeScript, Office.js
- **AI**: OpenAI API (GPT-4o-mini)
- **Build Tools**: Webpack, Babel
- **Office Integration**: Office Add-in Framework

## 🔧 Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/chandru2301/ExcelAIHelper.git
   cd ExcelAIHelper
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Add your OpenAI API key:
   - Open `src/taskpane/taskpane.ts`
   - Replace `YOUR_OPENAI_API_KEY` with your actual API key

4. Start the development server:
   ```bash
   npm start
   ```

5. Sideload the add-in in Excel:
   - Follow the [official Microsoft documentation](https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing)

## 💡 Usage

1. **Select Data**: Highlight the cells you want to analyze
2. **Choose Mode**: Toggle between Cell mode (writes to spreadsheet) or Chat mode (interactive conversation)
3. **Run Analysis**: Click "Run AI Analysis" to process your data
4. **Follow-up Questions**: In chat mode, ask follow-up questions about your data
5. **Add Columns**: Ask the AI to add relevant columns to your data

## 🌟 Key Components

| Component | Role |
|-----------|------|
| Excel Add-in | Frontend chat UI + handler for Office.js commands |
| Office.js | API to programmatically control Excel (read/write cells) |
| OpenAI API | Converts user's natural language prompt into an Excel action |
| Excel Runtime | Reflects the changes (cells updated, formatted, etc.) |

## 📝 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🤝 Contributing

Contributions, issues, and feature requests are welcome! Feel free to check [issues page](https://github.com/chandru2301/ExcelAIHelper/issues).

---

<div align="center">
  Made with ❤️ for Excel users everywhere
</div> 
