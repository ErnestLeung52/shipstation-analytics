#!/usr/bin/env node

import { filterDataByDateRange } from './src/utils/dateFilter.js';

// Create sample data with dates in MM/DD/YYYY format
const sampleData = [
	{ 'Order Date': '02/01/2025', Store: 'TikTok Shop US Store', Rate: 4.33 },
	{ 'Order Date': '02/05/2025', Store: 'TikTok Shop US Store', Rate: 3.91 },
	{ 'Order Date': '02/10/2025', Store: 'Shopify Store', Rate: 5.22 },
	{ 'Order Date': '02/15/2025', Store: 'Shopify Store', Rate: 4.87 },
	{ 'Order Date': '02/20/2025', Store: 'Walmart Store', Rate: 6.12 },
	{ 'Order Date': '02/25/2025', Store: 'Walmart Store', Rate: 5.78 },
	{ 'Order Date': '03/01/2025', Store: 'TikTok Shop US Store', Rate: 4.45 },
	{ 'Order Date': '03/05/2025', Store: 'TikTok Shop US Store', Rate: 4.12 },
	{ 'Order Date': '03/10/2025', Store: 'Shopify Store', Rate: 5.33 },
	{ 'Order Date': '03/15/2025', Store: 'Shopify Store', Rate: 4.99 },
];

// Test different date ranges
function testDateFilter() {
	console.log('Testing date filter functionality...\n');

	// Test 1: Filter for early February
	const startDate1 = new Date(2025, 1, 1); // Feb 1, 2025
	const endDate1 = new Date(2025, 1, 10); // Feb 10, 2025
	console.log(`\nTest 1: Filter from ${startDate1.toLocaleDateString()} to ${endDate1.toLocaleDateString()}`);
	const filtered1 = filterDataByDateRange([...sampleData], startDate1, endDate1);
	console.log(`Filtered data (${filtered1.length} records):`);
	filtered1.forEach((record) => console.log(`  ${record['Order Date']} - ${record.Store}`));

	// Test 2: Filter for March only
	const startDate2 = new Date(2025, 2, 1); // Mar 1, 2025
	const endDate2 = new Date(2025, 2, 31); // Mar 31, 2025
	console.log(`\nTest 2: Filter from ${startDate2.toLocaleDateString()} to ${endDate2.toLocaleDateString()}`);
	const filtered2 = filterDataByDateRange([...sampleData], startDate2, endDate2);
	console.log(`Filtered data (${filtered2.length} records):`);
	filtered2.forEach((record) => console.log(`  ${record['Order Date']} - ${record.Store}`));

	// Test 3: Filter for a specific store in February
	const startDate3 = new Date(2025, 1, 1); // Feb 1, 2025
	const endDate3 = new Date(2025, 1, 28); // Feb 28, 2025
	console.log(
		`\nTest 3: Filter from ${startDate3.toLocaleDateString()} to ${endDate3.toLocaleDateString()} for Shopify Store`
	);
	const filtered3 = filterDataByDateRange([...sampleData], startDate3, endDate3).filter(
		(record) => record.Store === 'Shopify Store'
	);
	console.log(`Filtered data (${filtered3.length} records):`);
	filtered3.forEach((record) => console.log(`  ${record['Order Date']} - ${record.Store}`));
}

testDateFilter();
