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

			createReadStream(resolvedPath)
				.pipe(csv())
				.on('data', (data) => {
					// Clean and transform data
					const cleanedData = cleanData(data);
					results.push(cleanedData);
				})
				.on('end', () => {
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
 * Cleans and transforms raw CSV data
 * @param {Object} data - Raw data object from CSV parser
 * @returns {Object} - Cleaned data object
 */
function cleanData(data) {
	const cleanedData = {};

	// Copy all properties
	for (const [key, value] of Object.entries(data)) {
		// Clean up key names (remove whitespace, etc.)
		const cleanKey = key.trim();

		// Convert numeric values to numbers
		if (cleanKey === 'Rate' && value) {
			// Remove currency symbols and convert to number
			cleanedData[cleanKey] = parseFloat(value.replace(/[^0-9.-]+/g, '')) || 0;
		} else {
			cleanedData[cleanKey] = value;
		}
	}

	return cleanedData;
}
