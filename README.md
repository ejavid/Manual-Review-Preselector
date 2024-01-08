# Manual-Review-Preselector
A simple and efficient Python tool designed to aid analysts in locating and flagging documents that contain specific keywords for manual review. Supports PDF, Word, and Visio files.
# Review-Document-Finder

The `Review-Document-Finder` is a Python-based tool that streamlines the process for analysts who need to perform manual reviews of documents. It automatically searches through a directory of files, identifies documents containing specified keywords, and copies them to a designated location for further examination. This tool supports PDF, Word, and Visio file formats, simplifying the preliminary phase of document analysis.

## Features

- **Keyword Searching**: Locate documents containing specified keywords within their text.
- **File Type Support**: Handles PDF and Word documents for text search and flags Visio files based on their names.
- **Organized Output**: Copies flagged documents into keyword-specific folders for organized review.
- **Case-Insensitive Search**: The search process is not case-sensitive, increasing the tool's robustness.

## Prerequisites

Before running the script, ensure you have the following dependencies installed:

- Python 3.6 or higher
- PyMuPDF (for PDF handling)
- python-docx (for Word document handling)

These can be installed via pip:

```bash
pip install PyMuPDF python-docx


Usage
Clone the repository to your local machine.
Modify the src_folder, dst_folder, and search_words variables in the script to match your requirements.
Run the script using Python.
