#!/usr/bin/env node

/**
 * ShipStation Rates Calculator
 *
 * A command-line tool to analyze ShipStation CSV data and calculate metrics.
 * This is the main entry point that handles command-line arguments and
 * orchestrates the application flow.
 */

import { Command } from 'commander';
import chalk from 'chalk';
import path from 'path';
import { fileURLToPath } from 'url';
import { readCSVFile } from './utils/fileReader.js';
import { calculateStoreMetrics, calculateTagMetrics } from './metrics/calculator.js';
import { displayStoreMetrics, displayTagMetrics } from './display/reporter.js';
import { selectCSVFile } from './utils/fileSelector.js';
import { saveReportToCSV } from './utils/reportExporter.js';
import { saveReportToExcel } from './utils/excelExporter.js';
import { promptDateRange, filterDataByDateRange } from './utils/dateFilter.js';

// Get the directory name in ESM
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Determines the date range from the data
 * @param {Array<Object>} data - Array of order data
 * @returns {Object} - Object with startDate, endDate, and periodName
 */
function determineDateRange(data) {
	// Find the earliest and latest dates in the data
	let earliestDate = new Date('2100-01-01'); // Future date as initial value
	let latestDate = new Date('1900-01-01'); // Past date as initial value
	let dateFields = ['Order Date', 'OrderDate', 'Date', 'Ship Date', 'ShipDate'];

	// Parse date strings in MM/DD/YYYY format
	const parseDate = (dateStr) => {
		// Handle MM/DD/YYYY format
		const parts = dateStr.split('/');
		if (parts.length === 3) {
			const month = parseInt(parts[0], 10);
			const day = parseInt(parts[1], 10);
			const year = parseInt(parts[2], 10);
			return new Date(year, month - 1, day);
		}
		// Fallback to standard Date parsing
		return new Date(dateStr);
	};

	// Iterate through the data to find the date range
	for (const order of data) {
		for (const field of dateFields) {
			if (order[field]) {
				const orderDate = parseDate(order[field]);

				// Skip invalid dates
				if (isNaN(orderDate.getTime())) continue;

				// Update earliest and latest dates
				if (orderDate < earliestDate) earliestDate = new Date(orderDate);
				if (orderDate > latestDate) latestDate = new Date(orderDate);

				// Once we find a valid date field, we can break the inner loop
				break;
			}
		}
	}

	// Format the period name
	const startMonthName = earliestDate.toLocaleString('en-US', { month: 'short' });
	const startDay = earliestDate.getDate();
	const endMonthName = latestDate.toLocaleString('en-US', { month: 'short' });
	const endDay = latestDate.getDate();
	const year = latestDate.getFullYear();

	// Create a period name that includes the exact dates
	const periodName = `${startMonthName} ${startDay}-${endMonthName} ${endDay}, ${year}`;

	// Format date strings for display
	const startDateStr = `${(earliestDate.getMonth() + 1).toString().padStart(2, '0')}/${earliestDate
		.getDate()
		.toString()
		.padStart(2, '0')}/${(earliestDate.getFullYear() % 100).toString().padStart(2, '0')}`;
	const endDateStr = `${(latestDate.getMonth() + 1).toString().padStart(2, '0')}/${latestDate
		.getDate()
		.toString()
		.padStart(2, '0')}/${(latestDate.getFullYear() % 100).toString().padStart(2, '0')}`;

	return {
		shouldFilter: false, // We're not filtering, just determining the range
		startDate: earliestDate,
		endDate: latestDate,
		periodName,
		dateRangeStr: `${startDateStr}-${endDateStr}`,
	};
}

// Set up command-line interface
const program = new Command();

program
	.name('shipstation-calculator')
	.description('Calculate metrics from ShipStation CSV data')
	.argument('[filename]', 'CSV file to analyze (optional - will show file selector if not provided)')
	.option('-s, --store-only', 'Only calculate store metrics')
	.option('-t, --tag-only', 'Only calculate tag metrics')
	.option('-c, --compact', 'Display metrics in compact table format')
	.option('--save', 'Save the report to an Excel file')
	.option('--csv', 'Save the report as CSV instead of Excel (when used with --save)')
	.option('-d, --date-range <range>', 'Filter by date range in MM/DD/YY-MM/DD/YY format')
	.option('--no-prompt', 'Skip interactive prompts and analyze all data')
	.action(async (filename, options) => {
		try {
			console.log(chalk.blue('ShipStation Rates Calculator'));

			// If no filename is provided, show the file selector
			const fileToAnalyze = filename || (await selectCSVFile());

			console.log(chalk.gray(`Analyzing file: ${fileToAnalyze}\n`));

			// Read and parse the CSV file
			console.log(chalk.yellow('Reading CSV file...'));
			let data = await readCSVFile(fileToAnalyze);
			console.log(chalk.green(`Successfully read ${data.length} records\n`));

			// Handle date filtering
			let dateFilter;

			// If date range is provided as a command line option, use it
			if (options.dateRange) {
				// Parse the date range from command line
				const [startDateStr, endDateStr] = options.dateRange.split('-');

				// Validate format
				if (!startDateStr || !endDateStr || !startDateStr.includes('/') || !endDateStr.includes('/')) {
					throw new Error('Invalid date range format. Use MM/DD/YY-MM/DD/YY');
				}

				// Parse start date
				const startParts = startDateStr.split('/');
				const startMonth = parseInt(startParts[0], 10);
				const startDay = parseInt(startParts[1], 10);
				const startYear = parseInt(startParts[2], 10);
				const fullStartYear = 2000 + startYear;

				// Parse end date
				const endParts = endDateStr.split('/');
				const endMonth = parseInt(endParts[0], 10);
				const endDay = parseInt(endParts[1], 10);
				const endYear = parseInt(endParts[2], 10);
				const fullEndYear = 2000 + endYear;

				const startDate = new Date(fullStartYear, startMonth - 1, startDay);
				const endDate = new Date(fullEndYear, endMonth - 1, endDay);

				// Create period name with exact dates
				const startMonthName = startDate.toLocaleString('en-US', { month: 'short' });
				const endMonthName = endDate.toLocaleString('en-US', { month: 'short' });
				const periodName = `${startMonthName} ${startDay}-${endMonthName} ${endDay}, ${endDate.getFullYear()}`;

				dateFilter = {
					shouldFilter: true,
					startDate,
					endDate,
					periodName,
					dateRangeStr: options.dateRange,
				};

				console.log(chalk.yellow(`Using date range from command line: ${startDateStr} to ${endDateStr}`));
			} else if (options.prompt === false) {
				// Skip prompt if --no-prompt option is provided
				dateFilter = determineDateRange(data);
				console.log(
					chalk.yellow(
						`Analyzing all data from ${dateFilter.startDate.toLocaleDateString()} to ${dateFilter.endDate.toLocaleDateString()}`
					)
				);
			} else {
				// Otherwise prompt for date range
				dateFilter = await promptDateRange();

				// If no date range was provided in the prompt, determine it from the data
				if (!dateFilter.shouldFilter) {
					dateFilter = determineDateRange(data);
					console.log(
						chalk.yellow(
							`Analyzing all data from ${dateFilter.startDate.toLocaleDateString()} to ${dateFilter.endDate.toLocaleDateString()}`
						)
					);
				}
			}

			// Apply date filter if requested
			if (dateFilter.shouldFilter) {
				const originalCount = data.length;
				data = filterDataByDateRange(data, dateFilter.startDate, dateFilter.endDate);
				console.log(
					chalk.yellow(
						`Filtered data by date range: ${dateFilter.startDate.toLocaleDateString()} to ${dateFilter.endDate.toLocaleDateString()}`
					)
				);
				console.log(
					chalk.green(
						`Filtered from ${originalCount} to ${data.length} records (${Math.round(
							(data.length / originalCount) * 100
						)}% of original data)\n`
					)
				);

				// Check if we have data after filtering
				if (data.length === 0) {
					console.log(
						chalk.red(
							'No data matches the specified date range. Please check your date format and try again.'
						)
					);
					process.exit(1);
				}
			}

			// Calculate metrics
			const storeMetrics = calculateStoreMetrics(data);
			const tagMetrics = calculateTagMetrics(data);

			// Check if we have store metrics
			if (Object.keys(storeMetrics).length === 0) {
				console.log(chalk.red('No store data found. Please check your date range or CSV file.'));
				process.exit(1);
			}

			// Calculate total orders across all stores
			let totalAllStoresOrders = 0;
			if (!options.tagOnly) {
				// Count total orders from store metrics
				const stores = Object.keys(storeMetrics);
				for (const store of stores) {
					totalAllStoresOrders += storeMetrics[store].count;
				}

				// Make total orders available globally
				global.totalAllStoresOrders = totalAllStoresOrders;

				// Display store metrics with date range in the title
				displayStoreMetrics(storeMetrics, dateFilter.periodName);
			}

			if (!options.storeOnly) {
				// Display tag metrics with total orders count
				displayTagMetrics(tagMetrics, dateFilter.periodName);
			}

			// Save report if --save option is provided
			if (options.save) {
				if (options.csv) {
					// Save as CSV if --csv option is provided
					console.log(chalk.yellow('\nSaving report to CSV file...'));
					const savedFilePath = await saveReportToCSV(storeMetrics, tagMetrics, dateFilter.periodName);
					console.log(chalk.green(`Report saved to: ${savedFilePath}`));
				} else {
					// Save as Excel by default
					console.log(chalk.yellow('\nSaving report to Excel file...'));
					const savedFilePath = await saveReportToExcel(storeMetrics, tagMetrics, dateFilter.periodName);
					console.log(chalk.green(`Report saved to: ${savedFilePath}`));
				}
			}
		} catch (error) {
			console.error(chalk.red(`Error: ${error.message}`));
			process.exit(1);
		}
	});

program.parse();
