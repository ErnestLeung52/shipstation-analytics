/**
 * Date Filter Utility
 *
 * This module provides functions to filter data by date range.
 */

import inquirer from 'inquirer';

/**
 * Prompts the user to enter a date range for filtering
 * @returns {Promise<Object>} - Object with startDate, endDate, and periodName
 */
export async function promptDateRange() {
	const { dateRange } = await inquirer.prompt([
		{
			type: 'input',
			name: 'dateRange',
			message: 'Enter date range (MM/DD/YY-MM/DD/YY) or press Enter to analyze all data:',
			default: '',
			validate: (input) => {
				// Empty input is valid (analyze all data)
				if (!input) return true;

				// Check format: MM/DD/YY-MM/DD/YY
				const regex = /^(\d{1,2})\/(\d{1,2})\/(\d{2})-(\d{1,2})\/(\d{1,2})\/(\d{2})$/;
				const match = input.match(regex);

				if (!match) {
					return 'Please enter a valid date range in MM/DD/YY-MM/DD/YY format or leave empty';
				}

				// Validate start date
				const startMonth = parseInt(match[1], 10);
				const startDay = parseInt(match[2], 10);
				const startYear = parseInt(match[3], 10);
				// Handle 2-digit years
				const fullStartYear = 2000 + startYear;

				// Validate end date
				const endMonth = parseInt(match[4], 10);
				const endDay = parseInt(match[5], 10);
				const endYear = parseInt(match[6], 10);
				// Handle 2-digit years
				const fullEndYear = 2000 + endYear;

				const startDate = new Date(fullStartYear, startMonth - 1, startDay);
				const endDate = new Date(fullEndYear, endMonth - 1, endDay);

				if (isNaN(startDate.getTime())) {
					return 'Invalid start date';
				}

				if (isNaN(endDate.getTime())) {
					return 'Invalid end date';
				}

				if (startDate > endDate) {
					return 'Start date must be before end date';
				}

				return true;
			},
		},
	]);

	// If no date range provided, return shouldFilter: false
	if (!dateRange) {
		return { shouldFilter: false };
	}

	console.log(`Debug: User entered date range: ${dateRange}`);

	// Parse the date range
	const [startDateStr, endDateStr] = dateRange.split('-');

	// Parse start date (MM/DD/YY)
	const startParts = startDateStr.split('/');
	const startMonth = parseInt(startParts[0], 10);
	const startDay = parseInt(startParts[1], 10);
	const startYear = parseInt(startParts[2], 10);
	const fullStartYear = 2000 + startYear;

	// Parse end date (MM/DD/YY)
	const endParts = endDateStr.split('/');
	const endMonth = parseInt(endParts[0], 10);
	const endDay = parseInt(endParts[1], 10);
	const endYear = parseInt(endParts[2], 10);
	const fullEndYear = 2000 + endYear;

	const startDate = new Date(fullStartYear, startMonth - 1, startDay);
	const endDate = new Date(fullEndYear, endMonth - 1, endDay);

	console.log(`Debug: Parsed dates - Start: ${startDate.toLocaleDateString()}, End: ${endDate.toLocaleDateString()}`);

	// Create a period name from the date range that includes the exact dates
	const startMonthName = startDate.toLocaleString('en-US', { month: 'short' });
	const endMonthName = endDate.toLocaleString('en-US', { month: 'short' });
	const periodName = `${startMonthName} ${startDay}-${endMonthName} ${endDay}, ${endDate.getFullYear()}`;

	return {
		shouldFilter: true,
		startDate,
		endDate,
		periodName,
		dateRangeStr: `${startDateStr}-${endDateStr}`, // Store the original date range string
	};
}

/**
 * Filters data by date range
 * @param {Array<Object>} data - Array of order data
 * @param {Date} startDate - Start date for filtering
 * @param {Date} endDate - End date for filtering
 * @returns {Array<Object>} - Filtered data
 */
export function filterDataByDateRange(data, startDate, endDate) {
	// Set start date to beginning of day
	startDate.setHours(0, 0, 0, 0);

	// Set end date to end of day
	const adjustedEndDate = new Date(endDate);
	adjustedEndDate.setHours(23, 59, 59, 999);

	console.log(
		`Debug: Filtering dates between ${startDate.toLocaleDateString()} and ${adjustedEndDate.toLocaleDateString()}`
	);

	// Check the first few records to debug date parsing
	if (data.length > 0) {
		console.log('Debug: Sample data record:');
		console.log(JSON.stringify(data[0], null, 2));
	}

	// Function to parse date strings in MM/DD/YYYY format
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

	const filteredData = data.filter((order) => {
		// Try different date field variations
		const dateFields = ['Order Date', 'OrderDate', 'Date', 'Ship Date', 'ShipDate'];

		for (const field of dateFields) {
			if (order[field]) {
				const dateStr = order[field];
				const orderDate = parseDate(dateStr);

				// Skip invalid dates
				if (isNaN(orderDate.getTime())) continue;

				// Check if date is within range
				const isInRange = orderDate >= startDate && orderDate <= adjustedEndDate;

				// For debugging the first few records
				if (data.indexOf(order) < 3) {
					console.log(
						`Debug: Record #${data.indexOf(
							order
						)}, ${field}="${dateStr}" => ${orderDate.toLocaleDateString()}, in range: ${isInRange}`
					);
				}

				return isInRange;
			}
		}

		// If no valid date field found, include the order by default
		return true;
	});

	console.log(`Debug: Filtered from ${data.length} to ${filteredData.length} records`);

	if (filteredData.length === 0) {
		console.log(
			'Warning: No records match the date filter. Check if the date format in your CSV matches MM/DD/YYYY.'
		);
	}

	return filteredData;
}
