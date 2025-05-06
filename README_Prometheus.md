# Node-Excel-Export: A Lightweight, Flexible Excel Export Library for Node.js

## Project Overview

Node-Excel-Export is a lightweight Node.js library designed to simplify the process of generating Excel (xlsx) files programmatically. It provides a straightforward and flexible solution for exporting data sets into Excel spreadsheets with minimal configuration.

### Key Features

- **Simple Data Export**: Easily convert JavaScript data arrays into Excel spreadsheets
- **Multiple Sheet Support**: Create Excel files with multiple worksheets
- **Flexible Column Configuration**: Define column types, captions, and custom styles
- **Lightweight Implementation**: Minimal dependencies with a focus on core functionality
- **Custom Cell Formatting**: Supports preprocessing and styling of cell data

### Primary Use Cases

- Generating reports from application data
- Bulk data exports
- Automated spreadsheet creation in Node.js applications
- Converting structured data to Excel format

### Benefits

- No external Excel dependencies required
- Pure JavaScript implementation
- Supports various data types (strings, numbers, dates, booleans)
- Efficient memory usage with shared string optimization

## Getting Started, Installation, and Setup

### Prerequisites

- Node.js (version compatible with the project's dependencies)
- npm (Node Package Manager)

### Installation

Install the package using npm:

```bash
npm install excel-export
```

### Quick Start

#### Basic Usage

```javascript
const nodeExcel = require('excel-export');

// Define configuration for Excel export
const conf = {
  cols: [
    { caption: 'Column 1', type: 'string' },
    { caption: 'Column 2', type: 'date' }
  ],
  rows: [
    ['Data 1', new Date()],
    ['Data 2', new Date()]
  ]
};

// Generate Excel file
const result = nodeExcel.execute(conf);
```

### Configuration Options

#### Columns Configuration
- `caption`: Column header text
- `type`: Data type (`string`, `date`, `bool`, `number`)
- `width`: Column width
- `beforeCellWrite`: Optional function for data transformation

#### Advanced Features
- Custom styling via XML file
- Date and formatting transformations
- Supports large datasets

### Development

#### Running Tests
```bash
npm test
```

### Compatibility
- Works with Node.js environments
- Generates Excel (.xlsx) files
- Compatible with various data types and large datasets

## Features / Capabilities

### Core Features

#### XLSX Export Functionality
- Generate Excel (.xlsx) files programmatically
- Support for multiple sheet creation
- Flexible data export with configurable column types and styles

#### Data Type Support
- Export various data types:
  - Text
  - Numbers
  - Dates
  - Boolean values

#### Advanced Configuration Options
- Custom column captions
- Column width customization
- Per-column style configuration
- Pre-cell write data transformation
- Shared string optimization for text cells

#### Technical Capabilities
- Generates standard Office Open XML (.xlsx) compatible spreadsheets
- Supports Node.js environments
- Lightweight and dependency-minimal implementation
- Async and synchronous export methods

### Supported Cell Types
| Type | Description | Example |
|------|-------------|---------|
| `string` | Text content | `"Hello World"` |
| `number` | Numeric values | `42`, `3.14` |
| `date` | Date objects | `new Date()` |
| `bool` | Boolean values | `true`, `false` |

### Configuration Flexibility
Supports comprehensive configuration for each sheet, including:
- Column definitions
- Data rows
- Custom styles
- Cell transformations

## Usage Examples

### Basic Export with String and Numeric Data

```javascript
const nodeExcel = require('excel-export');

// Define column configuration
const conf = {
  cols: [
    { 
      caption: 'Name', 
      type: 'string', 
      width: 20 
    },
    { 
      caption: 'Age', 
      type: 'number', 
      width: 10 
    }
  ],
  rows: [
    ['John Doe', 30],
    ['Jane Smith', 25]
  ]
};

// Generate Excel file
const excelBuffer = nodeExcel.execute(conf);

// If in Express/HTTP context
res.setHeader('Content-Type', 'application/vnd.openxmlformats');
res.setHeader('Content-Disposition', 'attachment; filename=report.xlsx');
res.end(excelBuffer, 'binary');
```

### Advanced Export with Multiple Data Types

```javascript
const conf = {
  cols: [
    {
      caption: 'Description',
      type: 'string',
      beforeCellWrite: (row, cellData) => cellData.toUpperCase(),
      width: 20
    },
    {
      caption: 'Date',
      type: 'date',
      beforeCellWrite: (row, cellData, options) => {
        if (cellData === null) {
          options.cellType = 'string';
          return 'N/A';
        }
        return cellData;
      },
      width: 15
    },
    {
      caption: 'Boolean',
      type: 'bool'
    },
    {
      caption: 'Numeric Value',
      type: 'number',
      width: 15
    }
  ],
  rows: [
    ['First Entry', new Date(), true, 3.14159],
    ['Second Entry', new Date(), false, 2.7182]
  ]
};

const result = nodeExcel.execute(conf);
```

### Generating Large Datasets

```javascript
const conf = {
  cols: Array.from({length: 10}, (_, i) => ({
    caption: `Column ${i+1}`,
    type: 'string'
  })),
  rows: Array.from({length: 1000}, () => 
    Array.from({length: 10}, () => Math.random().toString())
  )
};

const result = nodeExcel.execute(conf);
```

### Notes
- The library supports different cell types: `string`, `number`, `date`, `bool`
- Use `beforeCellWrite` for custom cell preprocessing
- Specify column width and captions
- Complex transformations can be applied to cell data

## Project Structure

The project is organized with the following structure:

```
.
├── example/
│   ├── app.js             # Example application demonstrating usage
│   ├── package.json       # Example project dependencies
│   └── styles.xml         # Custom styling configuration for Excel export
├── test/
│   └── main.js            # Test suite for the library
├── index.js               # Main entry point for the Excel export library
├── sheet.js               # Core implementation for generating Excel worksheets
└── package.json           # Project metadata and dependencies
```

### Key Directories and Files

#### Source Files
- `index.js`: The primary module for Excel export functionality, handling the core generation process
- `sheet.js`: Contains the `Sheet` class responsible for generating individual Excel worksheets

#### Example and Testing
- `example/app.js`: Demonstrates how to use the Excel export library
- `test/main.js`: Contains tests to verify library functionality

#### Configuration
- `package.json`: Defines project metadata, dependencies, and scripts
- `example/styles.xml`: Optional custom styling configuration for Excel sheets

## Additional Notes

### Performance Insights

The Excel export library is engineered with memory efficiency and performance in mind. It provides lightweight Excel generation with optimized handling of various data types and export scenarios.

#### Memory Management
- Minimizes memory overhead through shared string optimization
- Supports incremental data processing for large datasets
- Designed for efficient memory utilization during Excel file generation

### Data Processing Capabilities

#### Type Conversion Strategies
- Automatic type inference and conversion for Excel cell values
- Specialized handling for different data types:
  - Numbers: Precise numeric representation
  - Dates: Standardized Excel date serialization
  - Booleans: Native Excel boolean mapping
  - Strings: XML-escaped with shared string optimization

### Technical Considerations

#### Extensibility
- Flexible configuration options for custom export requirements
- Support for dynamic column definitions
- Preprocessing hooks for cell value transformations

#### Compatibility Constraints
- Generates Excel 2007+ (.xlsx) format files
- Requires Node.js runtime environment
- Cross-platform support (Windows, macOS, Linux)

### Potential Integration Scenarios
- Reporting systems
- Data analysis tools
- Administrative dashboards
- Export functionality for web applications

### Security and Reliability
- Minimal external dependencies reduce potential vulnerability surface
- XML-based processing with built-in escaping mechanisms
- Consistent and predictable export behavior across different data sources

## Contributing

We welcome contributions to the Excel Export library! To ensure a smooth and collaborative contribution process, please follow these guidelines:

### Contribution Process

1. **Reporting Issues**
   - Use GitHub Issues to report bugs or suggest improvements
   - Provide a clear and detailed description
   - Include steps to reproduce the issue
   - If possible, include code snippets or example configurations

2. **Development Setup**
   - Fork the repository
   - Clone your forked repository
   - Install dependencies: `npm install`
   - Create a new branch for your feature or bugfix

### Code Guidelines

#### Coding Standards
- Follow existing code conventions in the project
- Use clear and descriptive variable and function names
- Add comments to explain complex logic
- Maintain consistent code formatting

#### Testing
- Write unit tests for new features or bug fixes
- Use Mocha for testing
- Run tests using: `npm test`
- Ensure all existing tests pass before submitting a pull request

### Pull Request Workflow
1. Commit your changes with a clear, descriptive message
2. Push your branch to your fork
3. Open a pull request against the main repository
4. Provide a detailed description of your changes

### Code of Conduct
- Be respectful and collaborative
- Help create an inclusive development environment

### Development Dependencies
- Mocha: Testing framework
- Should.js: Assertion library

### Notes
- Contributions are accepted under the project's BSD License
- Aim to maintain the library's lightweight and efficient design

If you have any questions, please open an issue for discussion.

## License

This project is licensed under the BSD License.

### License Details
- **License Type**: BSD License
- **Full License Text**: Available in the project's `LICENSE` file
- **Key Permissions**:
  - Commercial use allowed
  - Modification permitted
  - Distribution allowed
  - Private use allowed

### Conditions
- Redistribution must retain the original copyright notice
- Attribution to the original author is required

For the complete and official license terms, please refer to the project's license file.