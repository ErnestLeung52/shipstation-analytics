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

// Get the directory name in ESM
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Set up command-line interface
const program = new Command();

program
	.name('shipstation-calculator')
	.description('Calculate metrics from ShipStation CSV data')
	.version('1.0.0')
	.argument('<filename>', 'CSV file to analyze (will look in "ShipStation Orders" folder by default)')
	.option('-s, --store-only', 'Only calculate store metrics')
	.option('-t, --tag-only', 'Only calculate tag metrics')
	.action(async (filename, options) => {
		try {
			console.log(chalk.blue('ShipStation Rates Calculator'));
			console.log(chalk.gray(`Analyzing file: ${filename}\n`));

			// Read and parse the CSV file
			// The file reader will check both the current directory and the ShipStation Orders folder
			console.log(chalk.yellow('Reading CSV file...'));
			const data = await readCSVFile(filename);
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
