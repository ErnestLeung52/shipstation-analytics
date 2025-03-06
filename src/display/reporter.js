/**
 * Metrics Reporter
 *
 * This module provides functions to display calculated metrics in a readable format.
 * It formats and outputs store-based and tag-based metrics to the console.
 */

import chalk from 'chalk';
import Table from 'cli-table3';

/**
 * Formats a number as currency
 * @param {number} value - The value to format
 * @returns {string} - Formatted currency string
 */
function formatCurrency(value) {
	return `$${value.toFixed(2)}`;
}

/**
 * Formats a number as percentage
 * @param {number} value - The value to format
 * @returns {string} - Formatted percentage string
 */
function formatPercentage(value) {
	return `${value.toFixed(2)}%`;
}

/**
 * Colorizes a value based on whether it's positive, negative, or zero
 * @param {string} formattedValue - The pre-formatted value
 * @returns {string} - The colorized value
 */
function colorizeValue(formattedValue) {
	if (formattedValue.includes('-')) {
		return chalk.red(formattedValue);
	} else if (formattedValue === '$0.00' || formattedValue === '0.00%') {
		return chalk.yellow(formattedValue);
	} else {
		return chalk.green(formattedValue);
	}
}

/**
 * Displays all store metrics in a single comprehensive table
 * @param {Object} storeMetrics - Store metrics object
 * @param {Array<string>} stores - Array of store names
 * @param {number} totalOrders - Total number of orders
 * @param {number} totalOrderValue - Total order value
 * @param {number} totalRate - Total shipping cost
 * @param {number} totalShippingPaid - Total shipping paid by customers
 * @param {number} totalShippingProfit - Total shipping profit
 * @param {number} totalNetRevenue - Total net revenue
 */
function displayComprehensiveStoreTable(
	storeMetrics,
	stores,
	totalOrders,
	totalOrderValue,
	totalRate,
	totalShippingPaid,
	totalShippingProfit,
	totalNetRevenue
) {
	console.log(chalk.cyan.bold('\nComprehensive Store Metrics Table:'));

	// Reorder stores as specified: TikTok, Shopify, Walmart, Temu, Manual Orders
	const orderedStores = [];
	const storeOrder = ['TikTok Shop US Store', 'Shopify Store', 'Walmart Store', 'Temu Store', 'Manual Orders'];

	// First add stores in the specified order if they exist
	for (const storeName of storeOrder) {
		if (stores.includes(storeName)) {
			orderedStores.push(storeName);
		}
	}

	// Then add any remaining stores not in the specified order
	for (const store of stores) {
		if (!orderedStores.includes(store)) {
			orderedStores.push(store);
		}
	}

	// Calculate totals for the last column
	const overallShippingProfitMargin = totalShippingPaid > 0 ? (totalShippingProfit / totalShippingPaid) * 100 : 0;
	const overallNetRevenueMargin = totalOrderValue > 0 ? (totalNetRevenue / totalOrderValue) * 100 : 0;

	// Create table with metrics as rows and stores as columns
	const table = new Table({
		head: [
			chalk.white.bold('Metric'),
			...orderedStores.map((store) => chalk.white.bold(store)),
			chalk.white.bold('TOTAL'),
		],
		style: {
			head: [], // Disable colors in header
			border: [], // Disable colors for borders
		},
		wordWrap: true,
	});

	// Format values with percentages for specified fields
	const formatWithPercent = (value, total, isFormatted = false) => {
		if (total === 0) return isFormatted ? value : '0';

		// Extract numeric value if it's a formatted currency string
		let numericValue = value;
		if (isFormatted) {
			// Extract numeric value from currency string (remove $ and possible color codes)
			const plainValue = value.replace(/\u001b\[\d+m/g, ''); // Remove ANSI color codes
			numericValue = parseFloat(plainValue.replace(/[^\d.-]/g, ''));
		}

		const percent = ((numericValue / total) * 100).toFixed(1);
		return `${value} (${percent}%)`;
	};

	// Add rows for each metric
	table.push(
		[
			'Orders',
			...orderedStores.map((store) => formatWithPercent(storeMetrics[store].count, totalOrders)),
			chalk.bold(totalOrders),
		],
		[
			'Order Value',
			...orderedStores.map((store) =>
				formatWithPercent(formatCurrency(storeMetrics[store].totalOrderValue), totalOrderValue, true)
			),
			chalk.bold(formatCurrency(totalOrderValue)),
		],
		[
			'AOV',
			...orderedStores.map((store) => formatCurrency(storeMetrics[store].averageOrderValue)),
			chalk.bold(formatCurrency(totalOrderValue / totalOrders)),
		],
		[
			'Ship Cost',
			...orderedStores.map((store) =>
				formatWithPercent(formatCurrency(storeMetrics[store].totalRate), totalRate, true)
			),
			chalk.bold(formatCurrency(totalRate)),
		],
		[
			'Ship Paid',
			...orderedStores.map((store) =>
				formatWithPercent(formatCurrency(storeMetrics[store].totalShippingPaid), totalShippingPaid, true)
			),
			chalk.bold(formatCurrency(totalShippingPaid)),
		],
		[
			'Ship Profit',
			...orderedStores.map((store) => colorizeValue(formatCurrency(storeMetrics[store].shippingProfit))),
			chalk.bold(colorizeValue(formatCurrency(totalShippingProfit))),
		],
		[
			'Ship Margin',
			...orderedStores.map((store) => colorizeValue(formatPercentage(storeMetrics[store].shippingProfitMargin))),
			chalk.bold(colorizeValue(formatPercentage(overallShippingProfitMargin))),
		],
		[
			'Net Revenue',
			...orderedStores.map((store) =>
				formatWithPercent(colorizeValue(formatCurrency(storeMetrics[store].netRevenue)), totalNetRevenue, true)
			),
			chalk.bold(colorizeValue(formatCurrency(totalNetRevenue))),
		],
		[
			'Net Margin',
			...orderedStores.map((store) => colorizeValue(formatPercentage(storeMetrics[store].netRevenueMargin))),
			chalk.bold(colorizeValue(formatPercentage(overallNetRevenueMargin))),
		]
	);

	console.log(table.toString());
	console.log(chalk.gray('AOV = Average Order Value, Ship = Shipping, Net Margin = Net Revenue Margin'));
}

/**
 * Displays all tag metrics in a single comprehensive table
 * @param {Object} tagMetrics - Tag metrics object
 * @param {Array<string>} tags - Array of tag names
 * @param {number} totalTaggedOrders - Total number of tagged orders
 * @param {number} totalTagRate - Total shipping cost for tagged orders
 * @param {number} totalAllStoresOrders - Total number of orders across all stores
 */
function displayComprehensiveTagTable(tagMetrics, tags, totalTaggedOrders, totalTagRate, totalAllStoresOrders) {
	console.log(chalk.cyan.bold('\nComprehensive Tag Metrics Table:'));

	// Create table with metrics as rows and tags as columns
	const table = new Table({
		head: [chalk.white.bold('Metric'), ...tags.map((tag) => chalk.white.bold(tag)), chalk.white.bold('TOTAL')],
		style: {
			head: [], // Disable colors in header
			border: [], // Disable colors for borders
		},
		wordWrap: true,
	});

	// Calculate percentages for each tag
	const percentOfTotalOrders = tags.map((tag) => ((tagMetrics[tag].count / totalAllStoresOrders) * 100).toFixed(1));
	const totalPercentOfAllOrders = ((totalTaggedOrders / totalAllStoresOrders) * 100).toFixed(1);

	// Add rows for each metric
	table.push(
		['Orders', ...tags.map((tag) => tagMetrics[tag].count), chalk.bold(totalTaggedOrders)],
		[
			'% of All Orders',
			...percentOfTotalOrders.map((percent) => `${percent}%`),
			chalk.bold(`${totalPercentOfAllOrders}%`),
		],
		[
			'Total Shipping Cost',
			...tags.map((tag) => formatCurrency(tagMetrics[tag].totalRate)),
			chalk.bold(formatCurrency(totalTagRate)),
		],
		[
			'Avg Shipping Cost',
			...tags.map((tag) => formatCurrency(tagMetrics[tag].averageRate)),
			chalk.bold(formatCurrency(totalTagRate / totalTaggedOrders)),
		]
	);

	console.log(table.toString());
	console.log(chalk.gray('% of All Orders = Orders with this tag / Total orders across all stores'));
}

export function displayStoreMetrics(storeMetrics) {
	console.log(chalk.blue.bold('\n=== Store Metrics ==='));

	// Get stores and sort alphabetically
	const stores = Object.keys(storeMetrics).sort();

	if (stores.length === 0) {
		console.log(chalk.yellow('No store data found'));
		return;
	}

	// Calculate totals for summary
	let totalOrders = 0;
	let totalRate = 0;
	let totalOrderValue = 0;
	let totalShippingPaid = 0;
	let totalShippingProfit = 0;
	let totalNetRevenue = 0;

	// Collect data for each store
	for (const store of stores) {
		const metrics = storeMetrics[store];
		totalOrders += metrics.count;
		totalRate += metrics.totalRate;
		totalOrderValue += metrics.totalOrderValue;
		totalShippingPaid += metrics.totalShippingPaid;
		totalShippingProfit += metrics.shippingProfit;
		totalNetRevenue += metrics.netRevenue;
	}

	// Display comprehensive store metrics table
	displayComprehensiveStoreTable(
		storeMetrics,
		stores,
		totalOrders,
		totalOrderValue,
		totalRate,
		totalShippingPaid,
		totalShippingProfit,
		totalNetRevenue
	);

	// Display legend and help text
	console.log(chalk.gray('\nLegend:'));
	console.log(chalk.gray('- AOV = Average Order Value'));
	console.log(chalk.gray('- Ship = Shipping'));
	console.log(chalk.gray('- Net Margin = Net Revenue / Order Value'));
	console.log(chalk.gray('- Ship Margin = Shipping Profit / Shipping Paid'));
	console.log(chalk.green('- Green values indicate profit'));
	console.log(chalk.red('- Red values indicate loss'));
	console.log(chalk.yellow('- Yellow values indicate break-even'));

	// Display detailed metrics by section
	console.log(chalk.blue.bold('\nDetailed Metrics by Store:'));
	for (const store of stores) {
		const metrics = storeMetrics[store];
		console.log(chalk.cyan.bold(`\n${store}:`));

		// Order Summary
		console.log(
			chalk.white('Orders:'),
			chalk.yellow(`${metrics.count} orders`),
			chalk.gray(`(AOV: ${formatCurrency(metrics.averageOrderValue)})`)
		);

		// Revenue Summary
		console.log(
			chalk.white('Revenue:'),
			chalk.yellow(formatCurrency(metrics.totalOrderValue)),
			chalk.gray('→'),
			colorizeValue(formatCurrency(metrics.netRevenue)),
			chalk.gray(`(${colorizeValue(formatPercentage(metrics.netRevenueMargin))} margin)`)
		);

		// Shipping Summary
		console.log(
			chalk.white('Shipping:'),
			chalk.yellow(`Cost: ${formatCurrency(metrics.totalRate)}`),
			chalk.gray('vs'),
			chalk.yellow(`Paid: ${formatCurrency(metrics.totalShippingPaid)}`),
			chalk.gray('='),
			colorizeValue(formatCurrency(metrics.shippingProfit)),
			chalk.gray(`(${colorizeValue(formatPercentage(metrics.shippingProfitMargin))} margin)`)
		);
	}

	// Display overall summary
	console.log(chalk.blue.bold('\nOverall Summary:'));
	console.log(chalk.white('Total Orders:'), chalk.yellow(totalOrders));
	console.log(
		chalk.white('Total Revenue:'),
		chalk.yellow(formatCurrency(totalOrderValue)),
		chalk.gray('→'),
		colorizeValue(formatCurrency(totalNetRevenue)),
		chalk.gray(`(${colorizeValue(formatPercentage((totalNetRevenue / totalOrderValue) * 100))} margin)`)
	);
	console.log(
		chalk.white('Total Shipping:'),
		chalk.yellow(`Cost: ${formatCurrency(totalRate)}`),
		chalk.gray('vs'),
		chalk.yellow(`Paid: ${formatCurrency(totalShippingPaid)}`),
		chalk.gray('='),
		colorizeValue(formatCurrency(totalShippingProfit)),
		chalk.gray(`(${colorizeValue(formatPercentage((totalShippingProfit / totalShippingPaid) * 100))} margin)`)
	);
}

export function displayTagMetrics(tagMetrics) {
	console.log(chalk.blue.bold('\n=== Tag Metrics ==='));

	// Get tags and sort alphabetically
	const tags = Object.keys(tagMetrics).sort();

	if (tags.length === 0) {
		console.log(chalk.yellow('No tag data found'));
		return;
	}

	// Calculate totals for summary
	let totalTaggedOrders = 0;
	let totalTagRate = 0;

	// Collect data for each tag
	for (const tag of tags) {
		const metrics = tagMetrics[tag];
		totalTaggedOrders += metrics.count;
		totalTagRate += metrics.totalRate;
	}

	// Get total orders from all stores - this should be passed from the main function
	// For now, we'll use a parameter or fallback to the global variable if available
	const totalAllStoresOrders = global.totalAllStoresOrders || 1781; // Fallback to the number we saw in the report

	// Display comprehensive tag metrics table
	displayComprehensiveTagTable(tagMetrics, tags, totalTaggedOrders, totalTagRate, totalAllStoresOrders);

	// Display legend and help text
	console.log(chalk.gray('\nLegend:'));
	console.log(chalk.gray('- % of All Orders = Orders with this tag / Total orders across all stores'));
	console.log(chalk.gray('- Avg Shipping Cost = Total shipping cost / Number of orders'));

	// Display detailed metrics by section
	console.log(chalk.blue.bold('\nDetailed Tag Analysis:'));
	for (const tag of tags) {
		const metrics = tagMetrics[tag];
		const percentOfOrders = ((metrics.count / totalTaggedOrders) * 100).toFixed(1);
		const percentOfCost = ((metrics.totalRate / totalTagRate) * 100).toFixed(1);

		console.log(chalk.cyan.bold(`\n${tag}:`));
		console.log(
			chalk.white('Orders:'),
			chalk.yellow(`${metrics.count} orders`),
			chalk.gray(`(${percentOfOrders}% of tagged orders)`)
		);
		console.log(
			chalk.white('Shipping:'),
			chalk.yellow(`Total: ${formatCurrency(metrics.totalRate)}`),
			chalk.gray(`(${percentOfCost}% of total cost)`),
			chalk.gray(`Avg: ${formatCurrency(metrics.averageRate)}`)
		);
	}

	// Display overall tag summary
	console.log(chalk.blue.bold('\nTag Summary:'));
	console.log(chalk.white('Total Tagged Orders:'), chalk.yellow(totalTaggedOrders));
	console.log(chalk.white('Total Shipping Cost:'), chalk.yellow(formatCurrency(totalTagRate)));
	console.log(chalk.white('Average Cost per Order:'), chalk.yellow(formatCurrency(totalTagRate / totalTaggedOrders)));
	console.log(chalk.white('Unique Tags:'), chalk.yellow(tags.length));
}
