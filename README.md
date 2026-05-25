# Spreadsheet Calculator

A desktop spreadsheet application built with C# and Windows Forms, featuring a custom formula evaluation engine and cloud storage integration. This project demonstrates advanced concepts in lexical analysis, parsing, and data persistence.

## Features
* **Custom Expression Engine:** Utilizes ANTLR4 to build a lexer and parser for mathematical expressions, constructing Abstract Syntax Trees (AST) for accurate formula evaluation.
* **Dynamic Cell Referencing:** Supports complex cell references (e.g., `A1 + B2`) with real-time recalculation of dependencies.
* **Cloud & Local Storage:** Save and load workbooks locally in JSON format or directly to cloud storage using the Google Drive API.
* **Classic UI:** A familiar, easy-to-use grid interface built with Windows Forms.

## Tech Stack
* **Language:** C# (.NET)
* **Framework:** Windows Forms
* **Parsing & Lexing:** ANTLR4
* **Data Persistence:** JSON, Google Drive API

## Getting Started
1. Clone the repository:
   `git clone https://github.com/oleksandratupchiy/oop-lab1.git`
2. Open the `.sln` solution file in Visual Studio.
3. Ensure you have the necessary ANTLR4 dependencies installed via NuGet.
4. If testing Google Drive integration, ensure your `client_secrets.json` or equivalent auth files are properly configured in the project directory.
5. Build and run the project.

## Project Structure Highlights
* `/SpreadsheetApp/Parser/` - Contains ANTLR grammars and generated lexer/parser classes.
* `/SpreadsheetApp/Expressions/` - Logic for AST construction and expression evaluation.
* `/SpreadsheetApp/Persistence/` - Implementation of local JSON and Google Drive storage mechanisms.
