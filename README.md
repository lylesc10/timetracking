# TimeTrackingScript

## Overview

`script.py` is a Python application for processing and analyzing employee time tracking Excel files. It provides a graphical user interface (GUI) for selecting and processing one or more Excel files, and generates summary statistics and charts. The processed data and visualizations are saved back to new Excel files.

### Features

- **GUI for file selection:** Easily select one or more Excel files to process.
- **Data cleaning:** Filters out interns and removes summary rows.
- **Summary generation:** Calculates total hours per employee and per supervisor.
- **Comment analysis:** Aggregates hours by unique comments and generates a summary sheet.
- **Excel charting:** Automatically creates bar charts in the output Excel file.
- **Visual feedback:** Displays summary information in the GUI after processing.

## How to Use

### Requirements

- Python 3.9+
- The following Python packages:
  - pandas
  - openpyxl
  - tkinter (usually included with Python)
  - matplotlib
  - pillow

You can install the required packages with:

```bash
pip install -r requirements.txt
```

### Running the Script

#### GUI Mode

Simply run:

```bash
python script.py
```

A window will appear. Click the "Select Files" button and choose one or more Excel files (`.xlsx` or `.xls`). The script will process each file, generate new Excel files with summaries and charts, and display summary information in the GUI.

#### Command-Line Mode

You can also process a single file from the command line:

```bash
python script.py path/to/your/file.xlsx
```

### Output

For each input file, a new file named `<original_filename>_sorted.xlsx` will be created in the same directory. This file will contain:

- A cleaned and sorted data sheet
- A "Comment Summary" sheet with total hours per comment
- Bar charts visualizing the comment summary

## Docker

You can run the script in a Docker container. See the included `Dockerfile` for details. Build and run with:

```bash
docker build -t excel-processor .
docker run -it --rm -e DISPLAY=$DISPLAY -v /tmp/.X11-unix:/tmp/.X11-unix excel-processor
```

> **Note:** Running the GUI in Docker requires X11 forwarding and an X server on your host machine.

## License

MIT License

---

For questions or issues, please open an issue on the GitHub repository.
