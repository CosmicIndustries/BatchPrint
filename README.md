# 🖨️ Windows Batch Printer Script

A lightweight Python utility for batch printing `.docx` documents and images on **Windows OS**. It leverages the default system applications and Windows printing APIs to send print jobs without needing third-party libraries.

---

## 📦 Features

- ✅ Automatically prints `.docx` files using their default associated application
- ✅ Prints `.png`, `.jpg`, `.jpeg`, `.bmp`, `.gif` images using the Windows printing API (`win32print`, `win32ui`)
- ⚠️ Warns on `.pdf` files (requires user configuration or manual extension)
- 📂 Recursively prints files in a specified folder
- 🪟 Fully compatible with **Windows OS**

---

## 🛠 Requirements

- Python 3.7+
- [Pillow (PIL)](https://pypi.org/project/Pillow/)
- `pywin32` package (`win32print`, `win32ui`)

Install dependencies:

```bash
pip install pywin32 Pillow
