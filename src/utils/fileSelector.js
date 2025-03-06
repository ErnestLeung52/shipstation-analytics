import inquirer from 'inquirer';
import fs from 'fs';
import path from 'path';

export async function selectCSVFile() {
	// Define the directory where CSV files are stored
	const ordersDir = path.join(process.cwd(), 'ShipStation Orders');

	// Get all CSV files from the directory
	const files = fs
		.readdirSync(ordersDir)
		.filter((file) => file.endsWith('.csv'))
		.map((file) => ({
			name: file,
			value: path.join('ShipStation Orders', file),
		}));

	if (files.length === 0) {
		throw new Error('No CSV files found in the ShipStation Orders directory');
	}

	const { selectedFile } = await inquirer.prompt([
		{
			type: 'list',
			name: 'selectedFile',
			message: 'Select a CSV file to analyze:',
			choices: files,
		},
	]);

	return selectedFile;
}
