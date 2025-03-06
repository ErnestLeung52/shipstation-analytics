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

// Get the directory name in ESM
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Set up command-line interface
const program = new Command();

program
	.name('shipstation-calculator')
	.description('Calculate metrics from ShipStation CSV data')
	.argument('[filename]', 'CSV file to analyze (optional - will show file selector if not provided)')
	.option('-s, --store-only', 'Only calculate store metrics')
	.option('-t, --tag-only', 'Only calculate tag metrics')
	.option('-c, --compact', 'Display metrics in compact table format')
	.action(async (filename, options) => {
		try {
			console.log(chalk.blue('ShipStation Rates Calculator'));

			// If no filename is provided, show the file selector
			const fileToAnalyze = filename || (await selectCSVFile());

			console.log(chalk.gray(`Analyzing file: ${fileToAnalyze}\n`));

			// Read and parse the CSV file
			console.log(chalk.yellow('Reading CSV file...'));
			const data = await readCSVFile(fileToAnalyze);
			console.log(chalk.green(`Successfully read ${data.length} records\n`));

			// Calculate and display metrics based on options
			if (!options.tagOnly) {
				const storeMetrics = calculateStoreMetrics(data);
				displayStoreMetrics(storeMetrics);
			}

			if (!options.storeOnly) {
				const tagMetrics = calculateTagMetrics(data);
				displayTagMetrics(tagMetrics);
			}
		} catch (error) {
			console.error(chalk.red(`Error: ${error.message}`));
			process.exit(1);
		}
	});

program.parse();
