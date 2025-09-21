# Excel to XML Batch Converter for DBS

A streamlined web application that converts Excel files to XML format with embedded barcode strings for book production workflows. Features bidirectional processing - import Excel data or re-import previously exported XML files for editing and re-export.

## Overview

This client-side application processes Excel files and generates XML job tickets with embedded 37-digit barcode strings. The tool automatically handles production route adjustments, provides preview and edit modes, and supports round-trip XML processing for data revision workflows.

## Features

### **Core Functionality**
- **Multi-Format Support**: Processes Excel (.xlsx, .xls) and XML (.xml) files
- **Bidirectional Workflow**: Import Excel → Export XML → Re-import XML → Edit → Re-export
- **Batch XML Import**: Upload multiple previously exported XML files for combined editing
- **Intelligent ISBN Detection**: Automatically selects from Limp_ISBN or Cased_ISBN columns
- **Barcode Integration**: Generates 37-digit barcode strings included as XML elements
- **Interactive Barcodes**: Click barcode strings to copy to clipboard
- **Row Management**: Preview and edit modes for comprehensive data validation

### **Production Route Intelligence**
- **Automatic Trim Width Adjustment**: Detects "Limp P/Bound 8pp Cover" routes and increases Trim_Width by 10mm
- **Visual Indicators**: Adjusted rows highlighted in yellow throughout the interface
- **Audit Trail**: Summary notifications of all automatic adjustments made

### **Barcode Generation**
- **37-Digit Format**: Precisely formatted according to book production specifications
- **ISBN Priority**: Limp_ISBN checked first, then Cased_ISBN as fallback
- **Transfer Station Logic**: Position 37 set to 1 (Limp) or 2 (Cased)
- **Fixed Standards**: Trim Off Head standardized at 3mm (0030)
- **Live Preview**: Barcode strings displayed in data tables with click-to-copy functionality

## Installation

1. Download the application files:
   - `index.html`
   - `script.js`

2. Place files in a web directory or open `index.html` directly in a web browser

3. No server-side technology required - this is a completely client-side application

## Usage Workflows

### **Excel Import Workflow**
1. **Upload Excel File**: Select single Excel file containing production data
2. **Automatic Processing**: 
   - File parsing and validation
   - Production route logic applied automatically
   - Barcode strings generated for each row
3. **Review & Edit**: Use Preview/Edit modes to validate and modify data
4. **Export**: Download all XMLs as ZIP file

### **XML Re-Import Workflow**
1. **Upload XML Files**: Select one or more previously exported XML files
2. **Data Reconstruction**: 
   - Parses XML files back to editable format
   - Combines multiple files into single dataset
   - Regenerates barcode strings from imported data
3. **Edit & Update**: Modify data using Edit mode interface
4. **Re-Export**: Generate updated XML files

### **Round-Trip Processing**
- **Excel → XML**: Initial batch processing for production
- **XML → XML**: Revision and update workflows
- **Combined Processing**: Merge multiple XML sources for consolidated editing

## Interface Modes

### **Preview Mode** (Default)
- Read-only view of processed data
- Shows Wi Number, ISBN, Title, Production Route, Trim Width, and Barcode String
- Click barcode strings to copy to clipboard
- Individual XML download buttons for each row
- Visual indicators for adjusted trim widths

### **Edit Mode**
- Editable interface for 12 key production fields
- Sticky column headers for large datasets
- Real-time barcode string updates with click-to-copy
- Save changes to return to Preview Mode

## File Format Requirements

### **Excel Input (.xlsx, .xls)**

#### **Required Columns**:
- **Limp_ISBN** OR **Cased_ISBN**: 13-digit ISBN identifier
- **Wi_Number**: Work item identifier for file naming
- **Trim_Height**: Book height in mm
- **Trim_Width**: Book width in mm
- **Spine_Size**: Spine thickness in mm
- **Cut_Off**: Book block height before trimming in mm

#### **Optional Columns**:
- **Title**: Book title (commas automatically removed)
- **Production_Route**: Used for automatic trim width adjustments
- **Page_Extent**, **Reel_Width**, **Imposition**, **Paper_Code**: Additional production data

### **XML Input (.xml)**
- Must be previously exported from this application
- Supports multiple file selection for batch import
- Automatically reconstructs original data structure
- Regenerates barcode strings (not imported from XML)

## Barcode String Format

37-digit barcode string structure:

| Position | Length | Description | Format/Rules |
|----------|--------|-------------|--------------|
| 1-13     | 13     | ISBN | From Limp_ISBN or Cased_ISBN |
| 14-17    | 4      | Endsheet Height | Fixed "0000" |
| 18-20    | 3      | Spine Size | Formatted with trailing zero |
| 21-24    | 4      | Book Block Height | Cut-Off + trailing zero |
| 25-28    | 4      | Trim Off Head | **Fixed "0030" (3mm)** |
| 29-32    | 4      | Trim Height | Height + trailing zero |
| 33-36    | 4      | Trim Width | Width + trailing zero |
| 37       | 1      | Transfer Station | **1 = Limp, 2 = Cased** |

### **Key Business Rules**:
- **ISBN Selection**: Limp_ISBN takes priority over Cased_ISBN
- **Transfer Station**: Automatically set based on ISBN source
- **Spine Formatting**: 2-digit values get trailing zero, 1-digit values get leading and trailing zeros
- **Standardized Trim**: 3mm (0030) trim off head for consistent production

### **Clipboard Functionality**:
- **Click to Copy**: Click any barcode string to copy to clipboard
- **Visual Feedback**: Text briefly changes to "Copied!" in green
- **Cross-Browser**: Works with modern clipboard API and fallback methods

## Generated XML Structure

Each row generates an XML file with this format:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<csv>
  <Wi_Number>324623</Wi_Number>
  <Limp_ISBN>9781739665333</Limp_ISBN>
  <Cased_ISBN></Cased_ISBN>
  <Title>Sample Book Title</Title>
  <Trim_Height>198</Trim_Height>
  <Trim_Width>129</Trim_Width>
  <Spine_Size>24</Spine_Size>
  <Cut_Off>209</Cut_Off>
  <!-- Additional mapped fields -->
  <Barcode>9781739665333000024020900301980129901</Barcode>
</csv>
```

### **XML Features**:
- **UTF-8 Encoding**: Standard character encoding
- **Clean Formatting**: Proper indentation and line breaks
- **Barcode Element**: Always appears as the last element
- **Field Mapping**: Excel column names converted to XML element names
- **Data Cleaning**: Title field automatically stripped of commas
- **Round-Trip Compatible**: Can be re-imported for further editing

## Download Features

### **Batch ZIP Download**
- **Single Archive**: All XML files in one convenient download
- **Smart Naming**: 
  - Single record: `[Wi_Number]_xml_files.zip`
  - Multiple records: `multi_xml_files.zip`
- **Organized Structure**: Clean file naming for easy processing

### **Individual Downloads**
- **Per-Row Export**: Download button for each record in preview table
- **Instant Access**: Direct XML file download without ZIP packaging
- **Filename Format**: `[Wi_Number].xml`

## File Upload Capabilities

### **Single Excel File**
- **Batch Processing**: Processes entire spreadsheet at once
- **Production Logic**: Applies trim width adjustments automatically
- **Data Validation**: Ensures required columns are present

### **Multiple XML Files**
- **Batch Import**: Select and upload multiple XML files simultaneously
- **Data Consolidation**: Combines all files into single editable dataset
- **Flexible Selection**: Choose any combination of previously exported XML files

### **Smart File Detection**
- **Type Validation**: Automatically detects Excel vs XML files
- **Mixed Type Prevention**: Prevents uploading Excel and XML files together
- **Error Handling**: Clear messages for invali# Excel to XML Batch Converter for DBS

A streamlined web application that converts Excel files to XML format with embedded barcode strings for book production workflows. Designed specifically for DBS (Digital Book Services) operations with intelligent ISBN handling and automated production route logic.

## Overview

This client-side application processes Excel files and generates XML job tickets with embedded 37-digit barcode strings. The tool automatically handles production route adjustments and provides both preview and edit modes for data validation before export.

## Features

### **Core Functionality**
- **Excel Processing**: Supports .xlsx and .xls file formats
- **XML Generation**: Creates individual XML files with embedded barcode data
- **Intelligent ISBN Detection**: Automatically selects from Limp_ISBN or Cased_ISBN columns
- **Barcode Integration**: Generates 37-digit barcode strings included as XML elements
- **Batch Download**: Single ZIP file containing all XML files
- **Row Management**: Preview and edit modes for data validation

### **Production Route Intelligence**
- **Automatic Trim Width Adjustment**: Detects "Limp P/Bound 8pp Cover" routes and increases Trim_Width by 10mm
- **Visual Indicators**: Adjusted rows highlighted in yellow throughout the interface
- **Audit Trail**: Summary notifications of all automatic adjustments made

### **Barcode Generation**
- **37-Digit Format**: Precisely formatted according to book production specifications
- **ISBN Priority**: Limp_ISBN checked first, then Cased_ISBN as fallback
- **Transfer Station Logic**: Position 37 set to 1 (Limp) or 2 (Cased)
- **Fixed Standards**: Trim Off Head standardized at 3mm (0030)
- **Live Preview**: Barcode strings displayed in data tables with hover tooltips

## Installation

1. Download the application files:
   - `index.html`
   - `script.js`

2. Place files in a web directory or open `index.html` directly in a web browser

3. No server-side technology required - this is a completely client-side application

## Usage

### **Primary Workflow**

1. **Upload Excel File**: Click "Choose Excel File" and select your data file
2. **Automatic Processing**: 
   - File parsing and validation
   - Production route logic applied automatically
   - Barcode strings generated for each row
   - XML data prepared for download
3. **Review Data**: 
   - Preview table shows key information including barcode strings
   - Switch to Edit Mode for detailed field modifications
4. **Download**: Click "Download All XMLs as ZIP" for batch export

### **Interface Modes**

#### **Preview Mode** (Default)
- Read-only view of processed data
- Shows Wi Number, ISBN, Title, Production Route, Trim Width, and Barcode String
- Individual XML download buttons for each row
- Visual indicators for adjusted trim widths

#### **Edit Mode**
- Editable interface for 12 key production fields
- Sticky column headers for large datasets
- Real-time barcode string updates
- Save changes to return to Preview Mode

## File Format Requirements

### **Excel Input (.xlsx, .xls)**

#### **Required Columns**:
- **Limp_ISBN** OR **Cased_ISBN**: 13-digit ISBN identifier
- **Wi_Number**: Work item identifier for file naming
- **Trim_Height**: Book height in mm
- **Trim_Width**: Book width in mm
- **Spine_Size**: Spine thickness in mm
- **Cut_Off**: Book block height before trimming in mm

#### **Optional Columns**:
- **Title**: Book title (commas automatically removed)
- **Production_Route**: Used for automatic trim width adjustments
- **Page_Extent**, **Reel_Width**, **Imposition**, **Paper_Code**: Additional production data

## Barcode String Format

37-digit barcode string structure:

| Position | Length | Description | Format/Rules |
|----------|--------|-------------|--------------|
| 1-13     | 13     | ISBN | From Limp_ISBN or Cased_ISBN |
| 14-17    | 4      | Endsheet Height | Fixed "0000" |
| 18-20    | 3      | Spine Size | Formatted with trailing zero |
| 21-24    | 4      | Book Block Height | Cut-Off + trailing zero |
| 25-28    | 4      | Trim Off Head | **Fixed "0030" (3mm)** |
| 29-32    | 4      | Trim Height | Height + trailing zero |
| 33-36    | 4      | Trim Width | Width + trailing zero |
| 37       | 1      | Transfer Station | **1 = Limp, 2 = Cased** |

### **Key Business Rules**:
- **ISBN Selection**: Limp_ISBN takes priority over Cased_ISBN
- **Transfer Station**: Automatically set based on ISBN source
- **Spine Formatting**: 2-digit values get trailing zero, 1-digit values get leading and trailing zeros
- **Standardized Trim**: 3mm (0030) trim off head for consistent production

## Generated XML Structure

Each row generates an XML file with this format:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<csv>
  <Wi_Number>324623</Wi_Number>
  <Limp_ISBN>9781739665333</Limp_ISBN>
  <Cased_ISBN></Cased_ISBN>
  <Title>Sample Book Title</Title>
  <Trim_Height>198</Trim_Height>
  <Trim_Width>129</Trim_Width>
  <Spine_Size>24</Spine_Size>
  <Cut_Off>209</Cut_Off>
  <!-- Additional mapped fields -->
  <Barcode>9781739665333000024020900301980129901</Barcode>
</csv>
```

### **XML Features**:
- **UTF-8 Encoding**: Standard character encoding
- **Clean Formatting**: Proper indentation and line breaks
- **Barcode Element**: Always appears as the last element
- **Field Mapping**: Excel column names converted to XML element names
- **Data Cleaning**: Title field automatically stripped of commas

## Download Features

### **Batch ZIP Download**
- **Single Archive**: All XML files in one convenient download
- **Smart Naming**: 
  - Single record: `[Wi_Number]_xml_files.zip`
  - Multiple records: `multi_xml_files.zip`
- **Organized Structure**: Clean file naming for easy processing

### **Individual Downloads**
- **Per-Row Export**: Download button for each record in preview table
- **Instant Access**: Direct XML file download without ZIP packaging
- **Filename Format**: `[Wi_Number].xml`

## Technical Implementation

### **Client-Side Processing**
- **Browser-Based**: No server dependencies or data transmission
- **JavaScript Libraries**:
  - SheetJS (XLSX) for Excel file parsing
  - JSZip for archive creation
  - Bootstrap 5 for responsive UI design
- **Modern Browser Support**: Works with current Chrome, Firefox, Safari, and Edge

### **Data Processing Pipeline**
1. **File Reading**: Excel files parsed in-memory
2. **Rule Application**: Production route logic and trim adjustments
3. **Barcode Generation**: 37-digit strings calculated for each row
4. **XML Creation**: Individual XML documents generated
5. **Archive Creation**: ZIP file assembly for batch download

### **Security & Privacy**
- **Local Processing**: All data remains in browser memory
- **No External Calls**: No data transmitted to external servers
- **Session-Based**: No persistent data storage
- **Safe Downloads**: Generated files contain only processed Excel data

## Error Handling

### **File Upload Issues**
- **Format Validation**: Ensures .xlsx or .xls file types
- **Structure Checking**: Validates presence of required columns
- **Empty File Detection**: Handles files with no data rows

### **Data Processing Issues**
- **Missing ISBN**: Graceful handling when both ISBN columns are empty
- **Invalid Numbers**: Default values for malformed numeric fields
- **Barcode Validation**: Ensures 37-digit format before inclusion

### **User Feedback**
- **Processing Indicators**: Real-time status updates during file processing
- **Error Messages**: Clear notifications for validation failures
- **Adjustment Alerts**: Notifications when production route logic is applied

## Business Logic

### **Production Route Processing**
1. **Pattern Detection**: Scans for "Limp P/Bound 8pp Cover" in Production_Route column
2. **Automatic Adjustment**: Adds 10mm to Trim_Width for matching rows
3. **Visual Feedback**: Highlights adjusted rows in yellow throughout interface
4. **Summary Report**: Displays count of adjustments made

### **ISBN Handling**
1. **Priority System**: Limp_ISBN checked before Cased_ISBN
2. **Format Standardization**: Ensures 13-digit format with zero padding
3. **Transfer Station Assignment**: Automatically sets position 37 based on ISBN source
4. **Validation**: Confirms ISBN availability before barcode generation

## Browser Requirements

- **Modern JavaScript Support**: ES6+ features required
- **File API**: For Excel file processing
- **Canvas Support**: For ZIP file generation
- **Download API**: For file downloads
- **Local Storage**: Not required - all processing in memory

## Troubleshooting

### **Common Issues**
- **File Not Processing**: Ensure Excel file contains data beyond header row
- **Missing Barcodes**: Check that either Limp_ISBN or Cased_ISBN contains valid data
- **Download Problems**: Verify browser allows downloads from local files
- **Performance**: Large files (1000+ rows) may take longer to process

### **Data Quality Tips**
- **Clean ISBN Data**: Ensure 13-digit format in ISBN columns
- **Numeric Fields**: Verify measurements are in millimeters
- **Column Headers**: Use standard naming or check for variations
- **Row Validation**: Remove empty rows before upload

## Version History

### **Current Version**
- **Streamlined Interface**: Single-column layout matching DBS standards
- **Embedded Barcodes**: 37-digit strings included as XML elements
- **Fixed Standards**: Standardized 3mm trim off head
- **Enhanced UX**: Improved processing feedback and error handling
- **Batch Processing**: Efficient handling of large datasets

## License

This project is available under the MIT License. Free to modify and distribute for production workflows.

## Support

For technical issues:
1. **File Format**: Verify Excel file structure and required columns
2. **Browser Console**: Check for JavaScript errors during processing
3. **Data Validation**: Ensure numeric fields contain valid measurements
4. **Download Settings**: Check browser popup and download permissions

The application is designed for reliability in production environments with comprehensive error handling and user feedback systems.