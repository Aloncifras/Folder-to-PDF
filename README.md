# Folder2PDF: Take the root of all your docs and make a single PDF for easier scraping

Easily merge and convert multiple Word documents from a directory (including subfolders) into a single PDF file while preserving formatting, images, and content. This Python script simplifies bulk conversion and document organization.
*(For now only DOC and DOCX, you can request more file formats if needed)

## Features

- **Root Folder Selection**: Prompt the user to select a folder containing `.doc` and `.docx` files on its branches.
- **File Merging**: Combines all Word documents in the folder into a single Word file, ensuring page breaks between documents.
- **PDF Conversion**: Converts the merged Word file into a single PDF file.
- **User-Friendly Interface**: Utilizes `Tkinter` and `ctypes` to provide prompts and notifications.
- **Error Handling**: Logs errors for inaccessible files and handles file conflicts.

## Requirements

- **Python**: 3.8 or higher.
- **Libraries**:
  - `docx2pdf`: For converting Word documents to PDF.
  - `pywin32`: For interacting with Microsoft Word.
  - `tkinter`: For GUI prompts.

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/your-username/convert-docs-to-pdf.git
   cd convert-docs-to-pdf
   ```

2. Install dependencies:
   ```bash
   pip install docx2pdf pywin32
   ```

3. Ensure Microsoft Word is installed on your system.

## Usage

1. Run the script:
   ```bash
   python convert_docs_to_pdf.py
   ```

2. Follow the prompts to select a folder containing `.doc` and `.docx` files.

3. The script will merge all documents in the folder and save the resulting PDF in the same folder with the folder name as the PDF name.

## How It Works

1. **Folder Selection**:
   - The script prompts the user to select a folder.

2. **File Processing**:
   - It iterates through all `.doc` and `.docx` files in the folder.
   - Each document's content is appended to a single Word document with page breaks in between.

3. **PDF Conversion**:
   - The merged Word document is converted to a PDF using `docx2pdf`.

4. **Notifications**:
   - The user is notified when the process is complete, and the location of the resulting PDF is displayed.

## Debugging and Logs

- The script provides logs for each processed file, indicating whether it was successfully merged or if errors occurred.
- The temporary merged document is retained for debugging (can be deleted by uncommenting the cleanup section in the script).

## Contributions

Contributions are welcome! Feel free to open issues or submit pull requests to improve this project.

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

