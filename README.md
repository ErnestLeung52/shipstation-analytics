# ShipStation Rates Calculator

A command-line tool to analyze ShipStation CSV data and calculate metrics.

## Features

-   Reads ShipStation CSV export files
-   Calculates order counts, total rates, and average rates by store
-   Calculates metrics for orders with specific tags
-   Efficient processing for files with 1500-2000+ rows

## Installation

1. Clone this repository
2. Install dependencies:

```bash
npm install
```

## Usage

Run the application with a CSV filename as an argument:

```bash
node src/index.js filename.csv
```

The application will look for files in the following locations:

1. The current directory
2. The "ShipStation Orders" folder in the current directory

You only need to provide the filename, not the full path.

### Options

-   `-s, --store-only`: Only calculate store metrics
-   `-t, --tag-only`: Only calculate tag metrics
-   `-h, --help`: Display help information
-   `-V, --version`: Display version information

## Example

```bash
# If your file is in the "ShipStation Orders" folder:
node src/index.js "Feb-March 2025.csv"
```

## File Organization

Place your ShipStation CSV files in a folder named "ShipStation Orders" in the project root:

```
ShipStation-Rates-Calculator/
├── ShipStation Orders/
│   ├── Feb-March 2025.csv
│   └── Other-Export.csv
├── src/
│   ├── index.js
│   └── ...
└── ...
```

## Project Structure

-   `src/index.js`: Main entry point
-   `src/utils/fileReader.js`: CSV file reading and parsing
-   `src/metrics/calculator.js`: Metrics calculation logic
-   `src/display/reporter.js`: Display and formatting of results

## Requirements

-   Node.js 14.x or higher
-   CSV files exported from ShipStation with fields like Order #, Store, Rate, Tags, etc.

## License

ISC
