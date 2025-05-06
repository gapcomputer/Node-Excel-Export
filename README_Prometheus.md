# Node-Excel-Export: A Lightweight and Efficient Node.js Excel Export Library

## Project Overview

Node-Excel-Export is a lightweight and efficient Node.js library designed to simplify the process of exporting data sets to Excel (xlsx) files. The library provides a straightforward solution for developers who need to generate Excel spreadsheets programmatically with minimal configuration.

### Key Features

- **Simple Data Export**: Easily convert data collections into Excel spreadsheets
- **Multiple Sheet Support**: Generate Excel files with multiple worksheets
- **Lightweight Implementation**: Minimal dependencies with a small footprint
- **Node.js Compatibility**: Designed for seamless integration with Node.js applications

### Purpose

The primary goal of this library is to streamline the process of creating Excel files from data sets. It addresses common challenges developers face when needing to export structured data to a spreadsheet format, providing a clean and intuitive API for Excel file generation.

### Problem Solved

Many developers struggle with complex Excel export libraries that are heavy, difficult to use, or require extensive configuration. Node-Excel-Export solves this by offering:

- A simple, straightforward method to convert data to Excel
- Minimal overhead and quick performance
- Flexible sheet generation with custom naming
- Support for various data types and structures

## Getting Started, Installation, and Setup

## Quick Start

### Installation

Install the package using npm:

```bash
npm install excel-export
```

### Basic Usage

```javascript
const ExcelExport = require('excel-export');

// Example configuration for an Excel sheet
const config = {
  cols: [
    { caption: 'Name', type: 'string' },
    { caption: 'Age', type: 'number' }
  ],
  rows: [
    ['John Doe', 28],
    ['Jane Smith', 32]
  ]
};

// Generate Excel file
const result = ExcelExport.execute(config);
```

### Prerequisites

- Node.js (version compatible with collections and node-zip)
- npm package manager

### Compatibility

This library supports:
- Node.js environments
- Excel xlsx file generation
- Single and multiple sheet export

### Advanced Configuration

#### Multiple Sheets

```javascript
const configs = [
  {
    name: 'Sheet1',
    cols: [...],
    rows: [...]
  },
  {
    name: 'Sheet2',
    cols: [...],
    rows: [...]
  }
];

const result = ExcelExport.execute(configs);
```

### Performance Notes

- Use `executeAsync()` for non-blocking file generation
- Large datasets may require additional memory management

### Dependencies

- `collections`: ^3.0.0
- `node-zip`: 1.x

### Platform Support

- Windows
- macOS
- Linux

## API Reference

### Main Exports

#### `execute(config)`
Generate an Excel (XLSX) file based on provided configuration.

- **Parameters**:
  - `config` (Object | Array): Configuration for Excel sheet generation
    - For single sheet: Configuration object
    - For multiple sheets: Array of configuration objects

- **Returns**: Raw Buffer containing the XLSX file

- **Example**:
```javascript
const xlsxGen = require('./index');

// Single sheet
const singleSheetConfig = {
  name: 'MySheet',
  cols: [
    { caption: 'Name', type: 'string', width: 20 },
    { caption: 'Age', type: 'number' }
  ],
  rows: [
    ['John Doe', 30],
    ['Jane Smith', 25]
  ]
};

// Multiple sheets
const multiSheetConfig = [
  {
    name: 'Sheet1',
    cols: [ ... ],
    rows: [ ... ]
  },
  {
    name: 'Sheet2',
    cols: [ ... ],
    rows: [ ... ]
  }
];

const xlsxBuffer = xlsxGen.execute(config);
```

#### `executeAsync(config, callback)`
Asynchronous version of `execute()` method.

- **Parameters**:
  - `config` (Object | Array): Same as `execute()`
  - `callback` (Function): Callback function receiving generated XLSX buffer

- **Example**:
```javascript
xlsxGen.executeAsync(config, (buffer) => {
  // Handle generated Excel file buffer
});
```

### Sheet Configuration Properties

Each sheet configuration supports the following properties:

#### Sheet Configuration Object
- `name` (string, optional): Name of the worksheet
- `cols` (Array): Column definitions
  - `caption` (string): Column header text
  - `type` (string): Cell type ('string', 'number', 'date', 'bool')
  - `width` (number, optional): Column width
  - `captionStyleIndex` (number, optional): Style index for column header
  - `beforeCellWrite` (function, optional): Pre-processing function for cell data
- `rows` (Array): 2D array of cell values
- `stylesXmlFile` (string, optional): Path to custom styles XML file

### Utility Methods

#### Date Prototype Extensions
- `Date.prototype.getJulian()`: Calculate Julian date
- `Date.prototype.oaDate()`: Convert date to Excel's serial date format

### Supported Cell Types
- Strings
- Numbers
- Dates
- Booleans

### Internal Components
- Uses `node-zip` for file compression
- Uses `collections/sorted-map` for efficient string tracking
- Generates XML files conforming to Office Open XML (OOXML) spreadsheet specification

### Limitations
- Maximum sheet configurations depends on system memory
- Large datasets may require performance optimization

## Project Structure

The project follows a standard Node.js module structure with the following key directories and files:

#### Root Directory
- `index.js`: Primary entry point for the module
- `package.json`: Project configuration, dependencies, and metadata
- `sheet.js`: Likely contains sheet-related functionality
- `.gitignore`: Specifies intentionally untracked files to ignore

#### Example Directory
- `example/app.js`: Demonstrates example usage of the module
- `example/package.json`: Dependencies specific to the example
- `example/styles.xml`: Potentially contains styling information for Excel exports

#### Test Directory
- `test/main.js`: Contains test suite for the module

The project is organized to separate core functionality, examples, and tests, making it easy to understand and extend the Excel export module.

## Technologies Used

### Programming Languages
- JavaScript (Node.js)

### Core Libraries and Dependencies
- `collections`: Data structure and utility library
- `node-zip`: ZIP file manipulation library

### Development and Testing Tools
- Mocha: Testing framework
- Should.js: Assertion library for testing

### Runtime Environment
- Node.js

### File Formats Supported
- Excel XLSX
- XML

### Package Management
- npm (Node Package Manager)

## Additional Notes

### Performance and Memory Considerations

This library is designed for converting data sets to Excel xlsx files with efficient memory management. Key characteristics include:

- Uses shared string optimization to reduce memory footprint for repeated string values
- Supports multiple sheet generation in a single Excel workbook
- Lightweight implementation with minimal external dependencies

### Date and Number Handling

The library provides specialized handling for different cell types:

- **Number Cells**: Stored with numeric precision
- **Date Cells**: Converted using Excel's date serialization format
- **Boolean Cells**: Converted to Excel's boolean representation (0 or 1)
- **String Cells**: XML-escaped and stored in shared strings to optimize file size

### Customization Options

Developers can customize export behavior through configuration options:

- Column-specific styling
- Custom cell value transformations via `beforeCellWrite` callback
- Dynamic column width configuration
- Custom styles through external XML stylesheet

### Column Type Support

Supports multiple column types:
- `string` (default)
- `number`
- `date`
- `bool`

### Limitations

- Large datasets may require careful memory management
- Complex formatting beyond basic styling might require custom XML manipulation
- Timezone considerations with date cells

### Compatibility

- Compatible with Node.js environments
- Generates Excel 2007+ (.xlsx) format spreadsheets
- Uses native Node.js modules for processing

## Contributing

We welcome contributions to this project! To ensure a smooth contribution process, please follow these guidelines:

### Contribution Process

1. Fork the repository and create your branch from `main`.
2. Ensure any new code is well-documented and follows the existing code style.
3. Write tests for your changes using Mocha.

### Development Setup

- Install dependencies using `npm install`
- Run tests using `npm test`

### Submitting Contributions

1. Ensure your code passes all existing tests.
2. Add new tests for any new functionality.
3. Update documentation as needed.
4. Submit a pull request with a clear description of your changes.

### Testing

The project uses Mocha for testing. All tests are located in the `test/` directory.
- Run tests with `npm test`
- Ensure 100% test coverage for new features

### Code Style

- Follow existing code conventions in the project
- Use clear, descriptive variable and function names
- Add comments to explain complex logic

### Reporting Issues

- Use GitHub Issues to report bugs or suggest improvements
- Provide a clear and detailed description
- Include steps to reproduce the issue if applicable

### Dependencies

The project currently uses:
- `collections`: Data structure utilities
- `node-zip`: For Excel file processing
- `mocha` and `should`: For testing

**Note:** By contributing, you agree that your contributions will be licensed under the project's BSD license.

## License

This project is licensed under the BSD License. 

#### License Details
The BSD License is a permissive free software license that allows users to:
- Use the software for any purpose
- Modify the software
- Distribute the software
- Incorporate the software into other works

#### Conditions
- Redistributions of source code must retain the above copyright notice
- Redistributions in binary form must reproduce the copyright notice in the documentation
- Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission

For the full license text, please refer to the standard BSD License.