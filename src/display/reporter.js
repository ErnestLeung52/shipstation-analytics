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
 * Displays metrics grouped by store
 * @param {Object} storeMetrics - Object with store metrics
 */
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

	// Display order metrics table
	displayOrderMetricsTable(storeMetrics, stores, totalOrders, totalOrderValue);

	// Display shipping metrics table
	displayShippingMetricsTable(storeMetrics, stores, totalRate, totalShippingPaid);

	// Display profit metrics table
	displayProfitMetricsTable(
		storeMetrics,
		stores,
		totalShippingProfit,
		totalNetRevenue,
		totalOrderValue,
		totalShippingPaid
	);

	// Display detailed metrics for each store (original format)
	if (stores.length > 0) {
		console.log(chalk.blue.bold('\nDetailed Store Metrics:'));
		console.log(chalk.gray('(Scroll up to see the comparison tables)'));

		for (const store of stores) {
			const metrics = storeMetrics[store];

			console.log(chalk.green(`\nStore: ${chalk.bold(store)}`));

			// Order metrics
			console.log(chalk.cyan('  Order Metrics:'));
			console.log(`    Orders: ${chalk.yellow(metrics.count)}`);
			console.log(`    Total Order Value: ${chalk.yellow(formatCurrency(metrics.totalOrderValue))}`);
			console.log(`    Average Order Value (AOV): ${chalk.yellow(formatCurrency(metrics.averageOrderValue))}`);

			// Shipping metrics
			console.log(chalk.cyan('  Shipping Metrics:'));
			console.log(`    Total Shipping Cost (Rate): ${chalk.yellow(formatCurrency(metrics.totalRate))}`);
			console.log(`    Average Shipping Cost: ${chalk.yellow(formatCurrency(metrics.averageRate))}`);
			console.log(
				`    Total Shipping Paid by Customers: ${chalk.yellow(formatCurrency(metrics.totalShippingPaid))}`
			);
			console.log(`    Average Shipping Paid: ${chalk.yellow(formatCurrency(metrics.averageShippingPaid))}`);

			// Profit metrics
			console.log(chalk.cyan('  Profit Metrics:'));
			console.log(`    Shipping Profit/Loss: ${formatProfitLoss(metrics.shippingProfit)}`);
			console.log(`    Shipping Profit Margin: ${formatProfitLossPercentage(metrics.shippingProfitMargin)}`);
			console.log(`    Net Revenue (Revenue - Shipping Cost): ${formatProfitLoss(metrics.netRevenue)}`);
			console.log(`    Net Revenue Margin: ${formatProfitLossPercentage(metrics.netRevenueMargin)}`);
		}
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

	const table = new Table({
		head: [
			chalk.white.bold('Store'),
			chalk.white.bold('Orders'),
			chalk.white.bold('Order Value'),
			chalk.white.bold('AOV'),
			chalk.white.bold('Ship Cost'),
			chalk.white.bold('Ship Paid'),
			chalk.white.bold('Ship Profit'),
			chalk.white.bold('Ship Margin'),
			chalk.white.bold('Net Revenue'),
			chalk.white.bold('Net Margin'),
		],
		style: {
			head: [], // Disable colors in header
			border: [], // Disable colors for borders
		},
		wordWrap: true,
	});

	// Add rows for each store
	for (const store of stores) {
		const metrics = storeMetrics[store];

		table.push([
			store,
			metrics.count,
			formatCurrency(metrics.totalOrderValue),
			formatCurrency(metrics.averageOrderValue),
			formatCurrency(metrics.totalRate),
			formatCurrency(metrics.totalShippingPaid),
			colorizeValue(formatCurrency(metrics.shippingProfit)),
			colorizeValue(formatPercentage(metrics.shippingProfitMargin)),
			colorizeValue(formatCurrency(metrics.netRevenue)),
			colorizeValue(formatPercentage(metrics.netRevenueMargin)),
		]);
	}

	// Add total row
	const overallShippingProfitMargin = totalShippingPaid > 0 ? (totalShippingProfit / totalShippingPaid) * 100 : 0;
	const overallNetRevenueMargin = totalOrderValue > 0 ? (totalNetRevenue / totalOrderValue) * 100 : 0;

	table.push([
		chalk.bold('TOTAL'),
		chalk.bold(totalOrders),
		chalk.bold(formatCurrency(totalOrderValue)),
		chalk.bold(formatCurrency(totalOrderValue / totalOrders)),
		chalk.bold(formatCurrency(totalRate)),
		chalk.bold(formatCurrency(totalShippingPaid)),
		chalk.bold(colorizeValue(formatCurrency(totalShippingProfit))),
		chalk.bold(colorizeValue(formatPercentage(overallShippingProfitMargin))),
		chalk.bold(colorizeValue(formatCurrency(totalNetRevenue))),
		chalk.bold(colorizeValue(formatPercentage(overallNetRevenueMargin))),
	]);

	console.log(table.toString());
	console.log(chalk.gray('AOV = Average Order Value, Ship = Shipping, Net Margin = Net Revenue Margin'));
}

/**
 * Displays order metrics in a table format
 * @param {Object} storeMetrics - Store metrics object
 * @param {Array<string>} stores - Array of store names
 * @param {number} totalOrders - Total number of orders
 * @param {number} totalOrderValue - Total order value
 */
function displayOrderMetricsTable(storeMetrics, stores, totalOrders, totalOrderValue) {
	console.log(chalk.cyan.bold('\nOrder Metrics by Store:'));

	const table = new Table({
		head: [
			chalk.white.bold('Store'),
			chalk.white.bold('Orders'),
			chalk.white.bold('% of Orders'),
			chalk.white.bold('Total Order Value'),
			chalk.white.bold('% of Value'),
			chalk.white.bold('AOV'),
		],
		style: {
			head: [], // Disable colors in header
			border: [], // Disable colors for borders
		},
	});

	// Add rows for each store
	for (const store of stores) {
		const metrics = storeMetrics[store];
		const percentOfOrders = ((metrics.count / totalOrders) * 100).toFixed(1);
		const percentOfValue = ((metrics.totalOrderValue / totalOrderValue) * 100).toFixed(1);

		table.push([
			store,
			metrics.count,
			`${percentOfOrders}%`,
			formatCurrency(metrics.totalOrderValue),
			`${percentOfValue}%`,
			formatCurrency(metrics.averageOrderValue),
		]);
	}

	// Add total row
	table.push([
		chalk.bold('TOTAL'),
		chalk.bold(totalOrders),
		chalk.bold('100.0%'),
		chalk.bold(formatCurrency(totalOrderValue)),
		chalk.bold('100.0%'),
		chalk.bold(formatCurrency(totalOrderValue / totalOrders)),
	]);

	console.log(table.toString());
}

/**
 * Displays shipping metrics in a table format
 * @param {Object} storeMetrics - Store metrics object
 * @param {Array<string>} stores - Array of store names
 * @param {number} totalRate - Total shipping cost
 * @param {number} totalShippingPaid - Total shipping paid by customers
 */
function displayShippingMetricsTable(storeMetrics, stores, totalRate, totalShippingPaid) {
	console.log(chalk.cyan.bold('\nShipping Metrics by Store:'));

	const table = new Table({
		head: [
			chalk.white.bold('Store'),
			chalk.white.bold('Total Shipping Cost'),
			chalk.white.bold('Avg Shipping Cost'),
			chalk.white.bold('Total Shipping Paid'),
			chalk.white.bold('Avg Shipping Paid'),
		],
		style: {
			head: [], // Disable colors in header
			border: [], // Disable colors for borders
		},
	});

	// Add rows for each store
	for (const store of stores) {
		const metrics = storeMetrics[store];

		table.push([
			store,
			formatCurrency(metrics.totalRate),
			formatCurrency(metrics.averageRate),
			formatCurrency(metrics.totalShippingPaid),
			formatCurrency(metrics.averageShippingPaid),
		]);
	}

	// Add total row
	const avgTotalRate = totalRate / stores.reduce((sum, store) => sum + storeMetrics[store].count, 0);
	const avgTotalShippingPaid = totalShippingPaid / stores.reduce((sum, store) => sum + storeMetrics[store].count, 0);

	table.push([
		chalk.bold('TOTAL'),
		chalk.bold(formatCurrency(totalRate)),
		chalk.bold(formatCurrency(avgTotalRate)),
		chalk.bold(formatCurrency(totalShippingPaid)),
		chalk.bold(formatCurrency(avgTotalShippingPaid)),
	]);

	console.log(table.toString());
}

/**
 * Displays profit metrics in a table format
 * @param {Object} storeMetrics - Store metrics object
 * @param {Array<string>} stores - Array of store names
 * @param {number} totalShippingProfit - Total shipping profit
 * @param {number} totalNetRevenue - Total net revenue
 * @param {number} totalOrderValue - Total order value
 * @param {number} totalShippingPaid - Total shipping paid
 */
function displayProfitMetricsTable(
	storeMetrics,
	stores,
	totalShippingProfit,
	totalNetRevenue,
	totalOrderValue,
	totalShippingPaid
) {
	console.log(chalk.cyan.bold('\nProfit Metrics by Store:'));

	const table = new Table({
		head: [
			chalk.white.bold('Store'),
			chalk.white.bold('Shipping Profit/Loss'),
			chalk.white.bold('Shipping Margin'),
			chalk.white.bold('Net Revenue'),
			chalk.white.bold('Net Revenue Margin'),
		],
		style: {
			head: [], // Disable colors in header
			border: [], // Disable colors for borders
		},
		colWidths: [20, 20, 15, 20, 18],
	});

	// Add rows for each store
	for (const store of stores) {
		const metrics = storeMetrics[store];

		table.push([
			store,
			colorizeValue(formatCurrency(metrics.shippingProfit)),
			colorizeValue(formatPercentage(metrics.shippingProfitMargin)),
			colorizeValue(formatCurrency(metrics.netRevenue)),
			colorizeValue(formatPercentage(metrics.netRevenueMargin)),
		]);
	}

	// Add total row
	const overallShippingProfitMargin = totalShippingPaid > 0 ? (totalShippingProfit / totalShippingPaid) * 100 : 0;
	const overallNetRevenueMargin = totalOrderValue > 0 ? (totalNetRevenue / totalOrderValue) * 100 : 0;

	table.push([
		chalk.bold('TOTAL'),
		chalk.bold(colorizeValue(formatCurrency(totalShippingProfit))),
		chalk.bold(colorizeValue(formatPercentage(overallShippingProfitMargin))),
		chalk.bold(colorizeValue(formatCurrency(totalNetRevenue))),
		chalk.bold(colorizeValue(formatPercentage(overallNetRevenueMargin))),
	]);

	console.log(table.toString());
}

/**
 * Displays metrics grouped by tags
 * @param {Object} tagMetrics - Object with tag metrics
 */
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

	// Display comprehensive tag metrics table
	displayComprehensiveTagTable(tagMetrics, tags, totalTaggedOrders, totalTagRate);

	// Display tag metrics in a table
	displayTagMetricsTable(tagMetrics, tags, totalTaggedOrders, totalTagRate);

	// Display detailed metrics for each tag (original format)
	if (tags.length > 0) {
		console.log(chalk.blue.bold('\nDetailed Tag Metrics:'));
		console.log(chalk.gray('(Scroll up to see the comparison table)'));

		for (const tag of tags) {
			const metrics = tagMetrics[tag];

			console.log(chalk.green(`\nTag: ${chalk.bold(tag)}`));
			console.log(`  Orders: ${chalk.yellow(metrics.count)}`);
			console.log(`  Total Shipping Cost: ${chalk.yellow(formatCurrency(metrics.totalRate))}`);
			console.log(`  Average Shipping Cost: ${chalk.yellow(formatCurrency(metrics.averageRate))}`);
		}

		// Display summary
		console.log(chalk.blue.bold('\nTag Summary:'));
		console.log(`  Total Unique Tags: ${chalk.yellow(tags.length)}`);
		console.log(`  Total Tagged Orders: ${chalk.yellow(totalTaggedOrders)}`);
		console.log(`  Total Shipping Cost: ${chalk.yellow(formatCurrency(totalTagRate))}`);
		console.log(
			`  Overall Average Shipping Cost: ${chalk.yellow(formatCurrency(totalTagRate / totalTaggedOrders))}`
		);
	}
}

/**
 * Displays all tag metrics in a single comprehensive table
 * @param {Object} tagMetrics - Tag metrics object
 * @param {Array<string>} tags - Array of tag names
 * @param {number} totalTaggedOrders - Total number of tagged orders
 * @param {number} totalTagRate - Total shipping cost for tagged orders
 */
function displayComprehensiveTagTable(tagMetrics, tags, totalTaggedOrders, totalTagRate) {
	console.log(chalk.cyan.bold('\nComprehensive Tag Metrics Table:'));

	const table = new Table({
		head: [
			chalk.white.bold('Tag'),
			chalk.white.bold('Orders'),
			chalk.white.bold('% of Tagged Orders'),
			chalk.white.bold('Total Shipping Cost'),
			chalk.white.bold('% of Cost'),
			chalk.white.bold('Avg Shipping Cost'),
		],
		style: {
			head: [], // Disable colors in header
			border: [], // Disable colors for borders
		},
		wordWrap: true,
	});

	// Add rows for each tag
	for (const tag of tags) {
		const metrics = tagMetrics[tag];
		const percentOfOrders = ((metrics.count / totalTaggedOrders) * 100).toFixed(1);
		const percentOfCost = ((metrics.totalRate / totalTagRate) * 100).toFixed(1);

		table.push([
			tag,
			metrics.count,
			`${percentOfOrders}%`,
			formatCurrency(metrics.totalRate),
			`${percentOfCost}%`,
			formatCurrency(metrics.averageRate),
		]);
	}

	// Add total row
	table.push([
		chalk.bold('TOTAL'),
		chalk.bold(totalTaggedOrders),
		chalk.bold('100.0%'),
		chalk.bold(formatCurrency(totalTagRate)),
		chalk.bold('100.0%'),
		chalk.bold(formatCurrency(totalTagRate / totalTaggedOrders)),
	]);

	console.log(table.toString());
}

/**
 * Displays tag metrics in a table format
 * @param {Object} tagMetrics - Tag metrics object
 * @param {Array<string>} tags - Array of tag names
 * @param {number} totalTaggedOrders - Total number of tagged orders
 * @param {number} totalTagRate - Total shipping cost for tagged orders
 */
function displayTagMetricsTable(tagMetrics, tags, totalTaggedOrders, totalTagRate) {
	console.log(chalk.cyan.bold('\nTag Metrics Comparison:'));

	const table = new Table({
		head: [
			chalk.white.bold('Tag'),
			chalk.white.bold('Orders'),
			chalk.white.bold('% of Tagged Orders'),
			chalk.white.bold('Total Shipping Cost'),
			chalk.white.bold('% of Cost'),
			chalk.white.bold('Avg Shipping Cost'),
		],
		style: {
			head: [], // Disable colors in header
			border: [], // Disable colors for borders
		},
	});

	// Add rows for each tag
	for (const tag of tags) {
		const metrics = tagMetrics[tag];
		const percentOfOrders = ((metrics.count / totalTaggedOrders) * 100).toFixed(1);
		const percentOfCost = ((metrics.totalRate / totalTagRate) * 100).toFixed(1);

		table.push([
			tag,
			metrics.count,
			`${percentOfOrders}%`,
			formatCurrency(metrics.totalRate),
			`${percentOfCost}%`,
			formatCurrency(metrics.averageRate),
		]);
	}

	// Add total row
	table.push([
		chalk.bold('TOTAL'),
		chalk.bold(totalTaggedOrders),
		chalk.bold('100.0%'),
		chalk.bold(formatCurrency(totalTagRate)),
		chalk.bold('100.0%'),
		chalk.bold(formatCurrency(totalTagRate / totalTaggedOrders)),
	]);

	console.log(table.toString());
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
 * Formats profit/loss values with appropriate color coding
 * @param {number} value - The profit/loss value
 * @returns {string} - Formatted and color-coded profit/loss string
 */
function formatProfitLoss(value) {
	if (value > 0) {
		return chalk.green(formatCurrency(value));
	} else if (value < 0) {
		return chalk.red(formatCurrency(value));
	} else {
		return chalk.yellow(formatCurrency(value));
	}
}

/**
 * Formats profit/loss percentage values with appropriate color coding
 * @param {number} value - The profit/loss percentage value
 * @returns {string} - Formatted and color-coded profit/loss percentage string
 */
function formatProfitLossPercentage(value) {
	if (value > 0) {
		return chalk.green(formatPercentage(value));
	} else if (value < 0) {
		return chalk.red(formatPercentage(value));
	} else {
		return chalk.yellow(formatPercentage(value));
	}
}
