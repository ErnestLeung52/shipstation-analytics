/**
 * File Reader Utility
 *
 * This module provides functions to read and parse CSV files.
 * It handles file validation and transforms the data into a usable format.
 */

import fs from 'fs/promises';
import { createReadStream, existsSync, mkdirSync } from 'fs';
import csv from 'csv-parser';
import path from 'path';

// Default directory for ShipStation files
const DEFAULT_SHIPSTATION_DIR = 'ShipStation Orders';

/**
 * Ensures the ShipStation Orders directory exists
 */
function ensureShipStationDirExists() {
	const dirPath = path.join(process.cwd(), DEFAULT_SHIPSTATION_DIR);
	if (!existsSync(dirPath)) {
		try {
			mkdirSync(dirPath, { recursive: true });
			console.log(`Created directory: ${DEFAULT_SHIPSTATION_DIR}`);
		} catch (error) {
			console.warn(`Warning: Could not create ${DEFAULT_SHIPSTATION_DIR} directory: ${error.message}`);
		}
	}
}

/**
 * Resolves the file path, checking in the default ShipStation directory if needed
 * @param {string} filePath - Original file path provided by user
 * @returns {string} - Resolved file path
 * @throws {Error} - If the file doesn't exist
 */
async function resolveFilePath(filePath) {
	// Ensure the ShipStation directory exists
	ensureShipStationDirExists();

	// First check if the file exists as provided
	if (existsSync(filePath)) {
		return filePath;
	}

	// If not, check in the default ShipStation directory
	const shipStationPath = path.join(process.cwd(), DEFAULT_SHIPSTATION_DIR, filePath);
	if (existsSync(shipStationPath)) {
		return shipStationPath;
	}

	// If still not found, throw an error
	throw new Error(`File not found: ${filePath}. Also checked in ${DEFAULT_SHIPSTATION_DIR} directory.`);
}

/**
 * Reads and parses a CSV file into an array of objects
 * @param {string} filePath - Path to the CSV file
 * @returns {Promise<Array>} - Array of objects representing the CSV data
 * @throws {Error} - If the file doesn't exist or isn't a valid CSV
 */
export async function readCSVFile(filePath) {
	try {
		// Resolve the file path, checking in the default directory if needed
		const resolvedPath = await resolveFilePath(filePath);

		// Validate file extension
		const fileExtension = path.extname(resolvedPath).toLowerCase();
		if (fileExtension !== '.csv' && fileExtension !== '.xlsx' && fileExtension !== '.xls') {
			throw new Error('File must be a CSV or Excel file');
		}

		// For Excel files, we would need additional handling
		if (fileExtension === '.xlsx' || fileExtension === '.xls') {
			throw new Error('Excel files are not yet supported. Please export as CSV from ShipStation.');
		}

		// Parse CSV file
		return new Promise((resolve, reject) => {
			const results = [];
			let headers = [];
			let firstRow = true;
			let sampleData = null;

			createReadStream(resolvedPath)
				.pipe(csv())
				.on('data', (data) => {
					// Store headers from the first row
					if (firstRow) {
						headers = Object.keys(data);
						firstRow = false;
						sampleData = data;

						// Log available fields to help with debugging
						console.log('Available fields in CSV:', headers.join(', '));

						// Log potential field mappings for important metrics
						identifyPotentialFields(headers, sampleData);
					}

					// Clean and transform data
					const cleanedData = cleanData(data, headers);
					results.push(cleanedData);
				})
				.on('end', () => {
					if (results.length > 0) {
						// Log a sample of the first row after cleaning to help with debugging
						console.log('\nSample of processed data (first row):');
						const sampleKeys = ['Store', 'Rate', 'Order Total', 'Shipping Paid', 'Tags'];
						for (const key of sampleKeys) {
							if (results[0][key] !== undefined) {
								console.log(`  ${key}: ${results[0][key]}`);
							} else {
								console.log(`  ${key}: <not found>`);
							}
						}
						console.log('');
					}

					resolve(results);
				})
				.on('error', (error) => {
					reject(new Error(`Failed to parse CSV: ${error.message}`));
				});
		});
	} catch (error) {
		if (error.code === 'ENOENT') {
			throw new Error(`File not found: ${filePath}`);
		}
		throw error;
	}
}

/**
 * Identifies potential fields for important metrics based on headers and sample data
 * @param {Array<string>} headers - CSV headers
 * @param {Object} sampleData - Sample data from the first row
 */
function identifyPotentialFields(headers, sampleData) {
	// Define categories of fields we're looking for
	const fieldCategories = {
		'Rate/Cost Fields': ['rate', 'cost', 'shipping cost', 'shipping rate'],
		'Order Total Fields': ['order total', 'total', 'order amount', 'amount', 'price'],
		'Shipping Paid Fields': ['shipping', 'shipping paid', 'customer shipping', 'shipping charge'],
		'Store Fields': ['store', 'marketplace', 'channel', 'source'],
		'Tag Fields': ['tag', 'tags', 'label', 'category'],
	};

	console.log('\nPotential field mappings:');

	// For each category, find potential matching fields
	for (const [category, keywords] of Object.entries(fieldCategories)) {
		const matches = headers.filter((header) => keywords.some((keyword) => header.toLowerCase().includes(keyword)));

		if (matches.length > 0) {
			console.log(`  ${category}:`);
			for (const match of matches) {
				const value = sampleData[match];
				console.log(`    - ${match}: ${value}`);
			}
		} else {
			console.log(`  ${category}: No potential matches found`);
		}
	}
	console.log('');
}

/**
 * Cleans and transforms raw CSV data
 * @param {Object} data - Raw data object from CSV parser
 * @param {Array<string>} headers - CSV headers
 * @returns {Object} - Cleaned data object
 */
function cleanData(data, headers) {
	const cleanedData = {};

	// Common field name variations in ShipStation exports
	const fieldMappings = {
		// Store field variations
		Store: ['Store', 'Marketplace', 'Channel', 'Source', 'Platform'],

		// Rate field variations
		Rate: ['Rate', 'ShippingRate', 'Shipping Rate', 'Cost', 'Shipping Cost', 'Postage Cost', 'Postage'],

		// Order total field variations
		'Order Total': [
			'Order Total',
			'OrderTotal',
			'Total',
			'Order Amount',
			'OrderAmount',
			'Order Value',
			'OrderValue',
		],

		// Shipping paid field variations
		'Shipping Paid': [
			'Shipping',
			'Shipping Paid',
			'ShippingPaid',
			'Customer Shipping',
			'CustomerShipping',
			'Shipping Charge',
		],

		// Tags field variations
		Tags: ['Tags', 'Tag', 'Labels', 'Label', 'Categories', 'Category'],
	};

	// Copy all properties
	for (const [key, value] of Object.entries(data)) {
		// Clean up key names (remove whitespace, etc.)
		const cleanKey = key.trim();

		// Handle numeric values
		if (isLikelyNumeric(cleanKey, value)) {
			// Remove currency symbols and convert to number
			cleanedData[cleanKey] = parseFloat(value.replace(/[^0-9.-]+/g, '')) || 0;
		} else {
			cleanedData[cleanKey] = value;
		}
	}

	// Ensure we have standardized field names for important metrics
	for (const [standardField, variations] of Object.entries(fieldMappings)) {
		if (!cleanedData[standardField]) {
			// Look for variations of this field
			for (const variation of variations) {
				if (cleanedData[variation] !== undefined) {
					cleanedData[standardField] = cleanedData[variation];
					break;
				}
			}
		}
	}

	// If we still don't have the required fields, try case-insensitive matching
	for (const [standardField, variations] of Object.entries(fieldMappings)) {
		if (!cleanedData[standardField]) {
			for (const header of headers) {
				if (variations.some((v) => header.toLowerCase() === v.toLowerCase())) {
					cleanedData[standardField] = cleanedData[header];
					break;
				}
			}
		}
	}

	return cleanedData;
}

/**
 * Determines if a field is likely to contain numeric data based on field name and value
 * @param {string} fieldName - The name of the field
 * @param {string} value - The value to check
 * @returns {boolean} - True if the field is likely numeric
 */
function isLikelyNumeric(fieldName, value) {
	// List of field names that typically contain currency or numeric values
	const numericFields = [
		'rate',
		'cost',
		'price',
		'total',
		'amount',
		'shipping',
		'paid',
		'value',
		'weight',
		'quantity',
		'qty',
		'postage',
	];

	// Check if the field name contains any of the numeric indicators
	const isNumericField = numericFields.some((term) => fieldName.toLowerCase().includes(term));

	// Check if the value looks like a currency (starts with $ or contains a decimal point)
	const looksLikeCurrency = /^\s*\$?\s*\d+(\.\d+)?\s*$/.test(value);

	return isNumericField || looksLikeCurrency;
}
