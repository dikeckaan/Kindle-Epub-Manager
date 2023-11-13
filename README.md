# EPUB Manager and Sender

## Overview
This Python application provides a user-friendly interface to manage and send EPUB files. Built with Tkinter and leveraging the `win32com` library for Windows integrations, it simplifies the process of handling EPUB files stored in a specific directory.

## Features

- **List EPUB Files**: Automatically lists all `.epub` files in a predefined directory.
- **Search Functionality**: Allows users to search through the listed EPUB files based on text input.
- **Open Random EPUB**: Selects and opens a random EPUB file from the list.
- **Open Selected EPUB**: Opens the EPUB file selected by the user.
- **Send to Kindle**: Sends the selected EPUB file to a predefined Kindle email address for easy reading on your device.

## How to Use

1. **Set Up Folder Path**: Change the `FOLDER_PATH` variable to the directory where your EPUB files are stored.
2. **Running the Application**: Execute the script to start the application. The GUI window will open.
3. **Searching**: Enter text in the search bar to filter the EPUB files.
4. **Open or Send Files**: Use the buttons provided to open an EPUB file or send it to your Kindle.

## Requirements

- Python
- Tkinter (usually comes with Python)
- `win32com` library for Python

## Installation

Ensure you have Python installed on your system. You can download and install it from [Python's official website](https://www.python.org/).

To install the required `win32com` library, run:

```bash
pip install pywin32
```

## Acknowledgements

Special thanks to everyone who contributed to the development and testing of this application. Your feedback and support have been invaluable.

---

*Generated with the assistance of ChatGPT.*

Feel free to modify or extend this template as needed for your GitHub repository!
