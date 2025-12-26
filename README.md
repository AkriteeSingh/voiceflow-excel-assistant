# VoiceFlow – Voice Controlled Excel Assistant

VoiceFlow is a voice-controlled Microsoft Excel add-in that enables users to perform spreadsheet operations using natural language voice commands. The system integrates a FastAPI backend with Google Gemini AI to interpret spoken instructions and execute corresponding Excel actions through Office.js.

This project demonstrates the practical application of AI-driven natural language understanding combined with real-time Office automation.

---

## Key Features

- Voice-based interaction inside Microsoft Excel
- AI-powered interpretation of natural language commands
- Supports Excel operations such as:
  - Writing values to cells
  - Inserting new rows or columns
  - Calculating sum, average, and standard deviation
  - Sorting and filtering data
  - Creating chart
- Real-time execution of commands within Excel
- Modular frontend and backend architecture

---

## System Architecture

The system follows a client-server architecture:

User Voice Input
↓
Browser Speech Recognition API
↓
Excel Taskpane (Office.js)
↓
FastAPI Backend
↓
Google Gemini AI
↓
Structured JSON Command
↓
Excel Action Execution

---

## Technology Stack

### Frontend (Excel Add-in)
- Office.js
- TypeScript
- HTML and CSS
- Web Speech API

### Backend
- Python
- FastAPI
- Google Gemini AI
- Uvicorn

---

##Security Considerations

-API keys and sensitive credentials are excluded from version control
-Configuration files containing secrets are ignored using .gitignore
-No sensitive data is committed to the reposit

---

##Future Enhancements

-Support for multiple worksheets
-Undo and command history
-Enhanced natural language feedback
-Offline fallback for basic commands

