# Document vs Webpage Comparison Tool

A GUI application that compares Microsoft Word documents (.docx) with live webpages to identify content differences.

## Features
- Compare Word documents with live webpages
- Generate detailed HTML and Markdown reports
- Visualize differences with color coding
- Batch processing of multiple documents
- Interactive URL matching interface

## Requirements
- Python 3.x
- Dependencies listed in requirements.txt

## Installation
1. Clone this repository
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage
1. Run the application:
   ```
   python main.py
   ```
2. Click "Run Manual Match & Compare"
3. Select a folder containing your Word documents
4. Match each document with its corresponding webpage URL
5. View the generated reports

## Output
- HTML reports for each comparison (saved alongside the Word documents)
- A summary Markdown report in the selected folder
- Console output with similarity scores

## License
MIT License