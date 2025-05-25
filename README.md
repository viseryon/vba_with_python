# VBA scripts (and Python)

## 🚀 Features

- Run Python scripts from Excel macros
- Capture and display script standard out from terminal in Excel
- Log changes in your workbook automatically

## 📦 Installation

1. **Clone this repository:**

   ```sh
   git clone https://github.com/yourusername/vba_with_python.git
   cd vba_with_python
   ```

2. **Install dependencies with [uv](https://github.com/astral-sh/uv):**

   ```sh
   uv sync
   ```

## 💡 Tips

- Make sure Python 3.11+ is installed and available in your PATH.
- Work on vba scripts with xlwings vba command
- Use `uv run xlwings vba edit --file workbook.xlsm` to edit scripts from your favourite editor
- Use `pythonw` to run scripts without opening a console window.
- Customize the VBA and Python scripts as needed for your workflow.

## 🛠️ Requirements

- Python 3.11+
- Excel (with macro support)
- [uv](https://github.com/astral-sh/uv)
