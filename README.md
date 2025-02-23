# Excel to XML Batch Converter

A web-based application that converts Excel files to XML format, supporting batch processing of multiple rows. The application automatically converts Excel files upon upload and generates individual XML files for each row, named according to the Wi_Number field in the data.

## Features

- Single-page web application with Bootstrap styling
- Automatic conversion of Excel files to XML on upload
- Batch processing of multiple rows
- Individual XML preview for each row
- Download options:
  - Download individual XML files
  - Download all files as ZIP
- Clear All functionality to reset the interface
- Responsive design that works on desktop and mobile devices
- Support for .xlsx and .xls file formats

## Technical Requirements

- Modern web browser with JavaScript enabled
- No server-side dependencies required
- Internet connection (for loading CDN resources)

## Dependencies

The application uses the following external libraries loaded via CDN:

- Bootstrap 5.3.2 - For styling and layout
- SheetJS 0.18.5 - For Excel file processing
- JSZip 3.10.1 - For creating ZIP archives of multiple XML files

## Installation

1. Clone this repository or download the files
2. No build process or package installation required
3. Can be served from any web server or opened directly in a browser

## Usage

1. Open `index.html` in a web browser
2. Click "Choose Excel File" or drag and drop an Excel file
3. The file will be automatically processed and each row will be converted to XML
4. For each row you can:
   - Preview the XML content in the table
   - Download individual XML files using row-specific download buttons
   - Download all XML files at once as a ZIP
5. Use the "Clear All" button to reset the interface and start fresh

## Expected Excel Format

The input Excel file should have:
- Headers in the first row
- Data in subsequent rows
- Required field: Wi_Number (used for file naming)
- Common fields include:
  - Customer_ID
  - Customer_Name
  - Title
  - Page_Extent
  - Production_Route
  - Admin_Status
  (and other relevant fields)

## Generated XML Structure

Each row generates an XML file following this structure:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<csv>
  <Wi_Number>319702</Wi_Number>
  <Customer_ID>THEHIS</Customer_ID>
  <Customer_Name>The History Press</Customer_Name>
  ...
</csv>
```

## Interface Elements

- File Upload: Select Excel file for conversion
- Preview Table:
  - Wi Number column
  - XML preview column
  - Download button for each row
- Action Buttons:
  - Download All XMLs as ZIP (green)
  - Clear All (grey)

## Error Handling

The application includes handling for:
- Invalid file types
- Empty rows (automatically skipped)
- Missing Wi_Number (falls back to 'unknown')
- File reading errors
- ZIP creation errors

## Browser Compatibility

Tested and working in:
- Chrome (latest)
- Firefox (latest)
- Safari (latest)
- Edge (latest)

## Security Considerations

- All processing is done client-side
- No data is sent to any server
- Files are processed in memory
- No data persistence between sessions

## Development Notes

### Project Structure

```
excel-to-xml-converter/
├── index.html     # Main application file
└── README.md      # Documentation
```

### Key JavaScript Functions

- `convertExcelToXml(file)`: Handles the Excel file processing
- `createXmlFromRow(headers, values)`: Generates XML for a single row
- `downloadSingleXml(xml, wiNumber)`: Downloads individual XML file
- `downloadAllXmls()`: Creates and downloads ZIP of all XML files
- `clearAll()`: Resets the interface state
- `updatePreviewTable()`: Updates the preview table with current data

## Customization

To modify the application:

1. XML Structure:
   - Modify the `createXmlFromRow` function to change XML format
   - Adjust header processing in `convertExcelToXml`

2. Styling:
   - Update Bootstrap classes in HTML
   - Modify table layout and preview sizing
   - Adjust button colors and positioning

3. Functionality:
   - Add validation rules in `convertExcelToXml`
   - Modify file naming convention
   - Add additional preview information
   - Customize ZIP file structure

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License

This project is available under the MIT License. Feel free to modify and distribute as needed.

## Support

For issues or questions:
1. Check the error handling section above
2. Verify Excel file format matches requirements
3. Ensure all dependencies are loading correctly
4. Check browser console for errors

## Future Enhancements

Potential improvements that could be added:
- Custom XML template selection
- Validation rules for input data
- Support for more complex Excel structures
- Export to different formats
- Progress indicators for large files
- Custom filename patterns
- Data filtering options
- Bulk editing capabilities