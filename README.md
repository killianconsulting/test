# Document Webpage Comparer

A Python-based tool for comparing draft Word documents against live webpages, with intelligent content alignment and visual comparison features.

## Features

- **Manual URL Matching**: Allows users to manually match DOCX files with their corresponding webpage URLs
- **Intelligent Content Comparison**: 
  - Advanced content alignment using similarity matching
  - Special handling of H1 headings to ensure proper content structure alignment
  - Partial content matching for similar but not identical content
- **Visual Comparison Interface**:
  - Side-by-side comparison view
  - Color-coded content blocks:
    - Green: Matched content
    - Red: Content present in draft but missing from live
    - Blue: Content present on live but not in draft
  - Perfect visual alignment with invisible spacers
- **Batch Processing**:
  - Process multiple documents in one go
  - Progress tracking with visual progress bar
  - Summary report generation
- **Multiple Output Formats**:
  - HTML reports with interactive side-by-side view
  - Markdown summary reports
  - Individual HTML reports for each comparison
  - Batch summary in both formats

## Technical Features

- **Content Processing**:
  - Intelligent heading detection and preservation
  - Smart whitespace normalization
  - HTML tag handling and cleaning
  - Link text preservation
  - Meta information extraction (title, description)
- **Comparison Algorithm**:
  - Primary similarity threshold of 90%
  - Secondary partial matching for similar content
  - Special handling of document structure elements
  - Intelligent block alignment based on H1 headings

## Requirements

- Python 3.x
- Required packages:
  - python-docx
  - beautifulsoup4
  - requests
  - tkinter (usually comes with Python)

## Usage

1. Run the program using `python main.py`
2. Click "Run Manual Match & Compare"
3. Select the folder containing your DOCX files
4. Match each DOCX file with its corresponding webpage URL
5. Wait for the comparison to complete
6. Review the generated reports:
   - Individual HTML reports for each comparison
   - Combined markdown report for the batch
   - Summary in the application window

## Output

The tool generates several types of output:
- Individual HTML reports named `report_X_filename.html`
- A combined markdown report named `comparison_report.md`
- A summary displayed in the application window

## Recent Updates

- Improved content alignment with H1 heading detection
- Enhanced visual spacing in the comparison view
- Perfect vertical alignment between draft and live content
- Better handling of missing content with invisible spacers
- Improved readability with consistent padding and margins

## License
MIT License