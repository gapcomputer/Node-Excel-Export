# Node Excel Export: Lightweight Excel File Generation for Node.js

## Project Overview

Node Excel Export is a lightweight Node.js library designed to simplify the process of exporting data sets to Excel (.xlsx) files. Built to provide a straightforward and efficient solution for generating Excel spreadsheets programmatically.

### Key Features

- **Simple Data Export**: Easily convert data collections into Excel spreadsheets
- **Flexible Configuration**: Supports multiple sheet generation and custom column definitions
- **Lightweight Implementation**: Minimal dependencies with efficient performance
- **Multiple Data Type Support**: Handles numbers, strings, dates, and boolean values
- **Customizable Styling**: Supports custom column widths and cell styles

### Problem Domain

Many Node.js applications require the ability to generate Excel reports or export data for further analysis. This library addresses the common challenges of Excel file generation by providing a simple, intuitive interface for converting data sets into Excel-compatible formats.

### Core Benefits

- Eliminates the complexity of manually creating Excel files
- Reduces boilerplate code for data export functionality
- Provides a consistent and reliable method for generating spreadsheets
- Compatible with various data sources and structures
- Minimizes external dependencies compared to more complex Excel generation libraries

## Getting Started, Installation, and Setup

### Prerequisites

- Node.js (version compatible with Node.js)
- npm package manager

### Installation

Install the package using npm:

```bash
npm install excel-export
```

### Quick Start

#### Basic Usage

```javascript
const ExcelExport = require('excel-export');

// Configure your Excel sheet
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
const excelBuffer = ExcelExport.execute(config);
```

### Development and Production

#### Running Tests

To run the project tests:

```bash
npm test
```

### Platform Compatibility

The library supports:
- Windows
- macOS
- Linux

### Dependencies

Key dependencies will be automatically installed:
- `collections`: ^3.0.0
- `node-zip`: 1.x

### Performance Considerations

- For large datasets, use `executeAsync()` method
- Recommended to manage memory for extensive data exports

## Features / Capabilities

The Node Excel Export library provides a simple and flexible solution for exporting data sets to Excel (.xlsx) files with the following core features:

### Data Export Capabilities
- Export multiple data sets to separate worksheets
- Support for various data types:
  - Strings
  - Numbers
  - Dates
  - Boolean values
- Customizable column widths
- Ability to apply custom styles to cells and columns

### Flexible Configuration
- Dynamic column definition
- Optional cell preprocessing
- Custom caption styling
- Configurable sheet names

### Advanced Features
- Shared string optimization for efficient Excel file generation
- Support for custom XML styling
- Cross-platform compatibility via Node.js

### Supported Data Transformations
- Automatic XML escaping for special characters
- Cell value transformation through `beforeCellWrite` callback
- Flexible type handling for different data formats

### Examples
```javascript
// Basic export configuration
var config = {
  name: 'Sales Report',
  cols: [
    { caption: 'Name', type: 'string' },
    { caption: 'Amount', type: 'number' },
    { caption: 'Date', type: 'date' }
  ],
  rows: [
    ['John Doe', 1000, new Date()],
    ['Jane Smith', 1500, new Date()]
  ]
};
```

### Compatibility
- Works with Node.js environments
- Generates standard .xlsx files compatible with Microsoft Excel, Google Sheets, and other spreadsheet software

## Usage Examples

### Basic Export

Create a simple Excel export with predefined column configurations and data:

```javascript
const excel = require('excel-export');

// Define column configuration
const conf = {
  cols: [
    { caption: 'Name', type: 'string', width: 20 },
    { caption: 'Age', type: 'number', width: 10 },
    { caption: 'Date', type: 'date', width: 15 }
  ],
  rows: [
    ['John Doe', 28, new Date()],
    ['Jane Smith', 35, new Date()]
  ]
};

// Generate Excel file
const result = excel.execute(conf);
```

### Multiple Sheet Export

Export data to multiple sheets in a single workbook:

```javascript
const excel = require('excel-export');

// Configuration for multiple sheets
const multiSheetConfig = [
  {
    name: 'Employees',
    cols: [
      { caption: 'Name', type: 'string' },
      { caption: 'Department', type: 'string' }
    ],
    rows: [
      ['John Doe', 'Sales'],
      ['Jane Smith', 'Marketing']
    ]
  },
  {
    name: 'Departments',
    cols: [
      { caption: 'Department Name', type: 'string' },
      { caption: 'Head Count', type: 'number' }
    ],
    rows: [
      ['Sales', 10],
      ['Marketing', 8]
    ]
  }
];

// Generate multi-sheet Excel file
const result = excel.execute(multiSheetConfig);
```

### Advanced Column Configuration

Customize cell rendering and styling:

```javascript
const excel = require('excel-export');

const config = {
  cols: [
    { 
      caption: 'Employee', 
      type: 'string', 
      beforeCellWrite: (row, cellData, event) => {
        // Custom cell data transformation
        return cellData.toUpperCase();
      }
    },
    { 
      caption: 'Salary', 
      type: 'number', 
      beforeCellWrite: (row, cellData, event) => {
        // Custom styling based on value
        if (cellData > 50000) {
          event.styleIndex = 2; // High salary style
        }
        return cellData;
      }
    }
  ],
  rows: [
    ['John Doe', 45000],
    ['Jane Smith', 55000]
  ]
};

const result = excel.execute(config);
```

### Custom Styles

Apply custom XML styles to your Excel sheet:

```javascript
const excel = require('excel-export');

const config = {
  cols: [
    { caption: 'Name', type: 'string' },
    { caption: 'Score', type: 'number' }
  ],
  rows: [
    ['John', 85],
    ['Jane', 92]
  ],
  stylesXmlFile: '/path/to/custom/styles.xml'
};

const result = excel.execute(config);
```

## Project Structure

The project is organized with the following key directories and files:

#### Main Project Files
- `index.js`: The primary entry point of the library
- `sheet.js`: Likely contains core functionality for Excel sheet manipulation
- `package.json`: Defines project metadata, dependencies, and scripts

#### Example Directory
- `example/`: Contains example implementation and demonstration files
  - `app.js`: Sample application showcasing library usage
  - `styles.xml`: Potentially defines styling for Excel exports
  - `package.json`: Example project dependencies

#### Testing
- `test/`: Contains project test suite
  - `main.js`: Main test file for the library

#### Configuration
- `.gitignore`: Specifies intentionally untracked files to ignore

### Project Dependencies
The project uses minimal dependencies:
- `collections`: Data structure utilities
- `node-zip`: Excel file (xlsx) generation support
- Development dependencies include `mocha` for testing

## Technologies Used

#### Languages
- JavaScript (Node.js)

#### Core Dependencies
- `collections`: Data structure library for advanced collection manipulation
- `node-zip`: ZIP file creation and manipulation library

#### Development and Testing
- Mocha: Testing framework for JavaScript
- Should.js: Assertion library for test cases

#### Runtime Environment
- Node.js

#### File Formats
- Excel XLSX
- XML (via `styles.xml`)

#### Package Management
- npm (Node Package Manager)

## Additional Notes

### Performance Considerations

The library is designed with memory efficiency in mind for Excel file generation. Key considerations include:

- Optimized shared string handling to reduce memory usage
- Support for generating multiple sheets in a single workbook
- Lightweight implementation with minimal external dependencies

### Data Type Handling

The export mechanism provides specialized handling for different cell types:

- **Numbers**: Preserved with numeric precision
- **Dates**: Converted using Excel's native date serialization format
- **Booleans**: Translated to Excel's boolean representation
- **Strings**: XML-escaped and stored in shared strings to minimize file size

### Customization Capabilities

Developers can enhance export functionality through various configuration options:

- Dynamic column width configuration
- Custom cell value transformations
- Column-specific styling via XML stylesheet
- Flexible sheet and column type definitions

#### Supported Column Types
- `string` (default)
- `number`
- `date`
- `bool`

### Known Limitations

- Memory management may be required for large datasets
- Complex formatting beyond basic styling might need custom XML manipulation
- Potential timezone considerations with date cell conversions

### Compatibility Notes

- Generates Excel 2007+ (.xlsx) format spreadsheets
- Works across Windows, macOS, and Linux environments
- Requires Node.js runtime

### External Dependencies

- `collections` (^3.0.0): Data structure utilities
- `node-zip` (1.x): ZIP file processing

## Contributing

We welcome and appreciate contributions to this project! To ensure a smooth and collaborative contribution process, please follow these guidelines:

### How to Contribute

1. **Fork the Repository**
   - Create a fork of the main repository
   - Clone your forked repository to your local machine

2. **Create a Branch**
   - Create a new branch for your feature or bugfix
   - Use a clear and descriptive branch name
   - Example: `feature/add-new-export-option` or `bugfix/resolve-date-formatting-issue`

### Development Setup

1. **Prerequisites**
   - Node.js (compatible with the project's current version)
   - npm package manager

2. **Installation**
   ```bash
   git clone <repository-url>
   cd excel-export
   npm install
   ```

### Contribution Guidelines

#### Code Style
- Follow the existing code conventions in the project
- Use clear and descriptive variable and function names
- Add comments to explain complex logic
- Ensure consistent formatting and indentation

#### Testing
- Write unit tests for new features or bug fixes
- Ensure all tests pass before submitting a pull request
- Use Mocha for writing tests
- Aim for high test coverage

#### Submitting Changes
1. Commit your changes with a clear and descriptive commit message
2. Push your branch to your fork
3. Open a pull request against the main repository
4. Provide a detailed description of your changes
   - What problem does this solve?
   - What is the proposed solution?
   - Are there any side effects or potential issues?

### Pull Request Process
- Ensure your code passes all existing tests
- Add new tests for any new functionality
- Update documentation to reflect your changes
- Your pull request will be reviewed by the maintainers
- Be prepared to make revisions based on feedback

### Reporting Issues
- Use GitHub Issues to report bugs or suggest improvements
- Provide a clear and detailed description
- Include steps to reproduce the issue
- If possible, include code snippets or screenshots

### Code of Conduct
- Be respectful and considerate of others
- Collaborate constructively
- Help create an inclusive and welcoming environment

### Licensing
By contributing, you agree that your contributions will be licensed under the project's BSD License.

### Questions?
If you have any questions about contributing, please open an issue for discussion.

## License

This project is licensed under the BSD License. 

For the full license text, please refer to the license details specified in the `package.json` file. The BSD License is a permissive free software license that allows for reuse within both free and proprietary software.

### Key Provisions
- Redistribution and use in source and binary forms are permitted
- Modifications and derivative works are allowed
- Attribution to the original author is typically required

#### License Identifier
- SPDX License Identifier: BSD