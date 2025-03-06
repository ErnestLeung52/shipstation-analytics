# ShipStation Rates Calculator

A command-line tool to analyze ShipStation CSV data and calculate metrics.

## Features

-   Reads ShipStation CSV export files
-   Calculates order counts, total rates, and average rates by store
-   Calculates order total, average order value (AOV), and shipping paid by customers
-   Computes shipping profit/loss and net revenue after shipping costs
-   Calculates shipping metrics for orders with specific tags
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

## Metrics Calculated

### Store Metrics

-   **Order Metrics**
    -   Order count
    -   Total order value
    -   Average order value (AOV)
-   **Shipping Metrics**
    -   Total shipping cost (what you paid)
    -   Average shipping cost
    -   Total shipping paid by customers
    -   Average shipping paid
-   **Profit Metrics**
    -   Shipping profit/loss (shipping paid minus shipping cost)
    -   Shipping profit margin
    -   Net revenue (order value minus shipping cost)
    -   Net revenue margin

Note: Net revenue represents the revenue available after shipping expenses, but before accounting for cost of goods sold (COGS) and other expenses.

### Tag Metrics

Tags are used for labeling special orders such as giveaways, influencer promotions, lost packages, etc. For tags, we calculate:

-   Order count
-   Total shipping cost
-   Average shipping cost

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
