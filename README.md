# Document Search Tool
A Python-based tool for searching text in various document formats (e.g., `.doc`, `.docx`, `.pdf`, `.txt`, `.xls`, `.xlsx`, `.ppt`, `.pptx`). This tool provides a user-friendly GUI for performing text searches with fuzzy matching capabilities.

## Features

- **Supported File Types**: `.doc`, `.docx`, `.pdf`, `.txt`, `.xls`, `.xlsx`, `.ppt`, `.pptx`.
- **Fuzzy Matching**: Search for text with approximate matches.
- **GUI**: Built with Tkinter for ease of use.
- **Progress Tracking**: Displays the number of files processed in real-time.
- **Stop Search**: Allows users to stop the search process at any time.
- **Configurable**: Saves and loads search configurations automatically.

## Usage
# 1. Run the Application
```
python modified_search.py
```
# 2. Steps:
- Select a file containing search strings.
- Choose a target file or directory to search.
- Select the file types to include in the search.
- Click "Perform Search" to start.
# 3. Stopping the Search: 
- Use the "Stop Search" button to halt the process.
# 4. Results:
- The results will be displayed in the text area, showing matches and their context.
## Configuration
- The application saves its configuration (e.g., last used search strings file and target path) in config.ini. This allows you to resume your work seamlessly.
## Contributing
- Contributions are welcome! Feel free to fork the repository and submit pull requests.
## License
- This project is licensed under the MIT License. See the LICENSE file for details.
## Acknowledgments
- Built with Python and Tkinter.
- Uses libraries like PyPDF2, python-docx, openpyxl, and python-pptx for document processing.
