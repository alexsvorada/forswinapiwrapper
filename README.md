# 🤖 ForsWinAPIWrapper

A VBA-based Windows API wrapper specifically designed to automate interactions with Fors windows. This tool provides functionality to send keyboard commands, handle shortcuts, and manage transactions in a Fors environment.

## 🚀 Features

-   Windows API integration for window handling and keyboard input simulation
-   Support for multiple transaction types (APAB, APAG, APAR, APAS, APAZ, APSG, MAGD)
-   Command processing with various input types:
    -   Direct text input
    -   Keyboard shortcuts
    -   Address-based commands
    -   Repeated commands
-   Automatic logging of all operations
-   Multi-language support (English/Slovak comments)

## 🎯 Purpose

The wrapper serves as a bridge between Excel and Fors windows, enabling users to:

-   Automate repetitive data entry tasks
-   Standardize transaction processes
-   Reduce human error in data input
-   Save time on routine operations
-   Maintain consistent workflow patterns

## How It Works ⚙️

The tool utilizes Windows API calls to simulate keyboard inputs and window interactions, allowing for:

-   Direct text input to Fors windows
-   Execution of keyboard shortcuts
-   Navigation through transaction screens
-   Automated data entry from Excel cells
-   Command sequencing and repetition

## 📋 Requirements

-   Microsoft Excel with VBA support
-   Windows operating system
-   Access to Fors system

## 🔧 Setup

1. Import the following modules into your Excel VBA project:

    - `App.bas`
    - `ForsWinApiWrapper.bas`
    - `Transactions.cls`

2. Ensure your Excel workbook contains the following worksheets:

    - "Main"
    - "Data"
    - "Logger"

3. Configure the "Main" worksheet with:
    - Server name in cell B3
    - Run count tracking in cell A13

## 📝 Logging

All operations are automatically logged in the "Logger" worksheet with:

-   Timestamp
-   User and server information
-   Command details

## 📖 Usage

The application uses the following rules to determine the input type:

-   If input starts with $ → Position/Transaction command
-   If input starts with & → Address command
-   If input contains \* → Repeated command
-   If input matches known shortcuts → Shortcut command
-   Otherwise → Plain text input

## 💡 For example

-   "Hello" → Plain text
-   "$APAB.transaction" → Position command
-   "&A1" → Address command
-   "TEXT\*3" → Repeated text
-   "F1" → Shortcut command
-   "CTRL\*2" → Repeated shortcut
