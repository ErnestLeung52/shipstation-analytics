/**
 * Metrics Reporter
 *
 * This module provides functions to display calculated metrics in a readable format.
 * It formats and outputs store-based and tag-based metrics to the console.
 */

import chalk from 'chalk';

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

	// Display metrics for each store
	for (const store of stores) {
		const metrics = storeMetrics[store];
		totalOrders += metrics.count;
		totalRate += metrics.totalRate;
		totalOrderValue += metrics.totalOrderValue;
		totalShippingPaid += metrics.totalShippingPaid;
		totalShippingProfit += metrics.shippingProfit;
		totalNetRevenue += metrics.netRevenue;

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
		console.log(`    Total Shipping Paid by Customers: ${chalk.yellow(formatCurrency(metrics.totalShippingPaid))}`);
		console.log(`    Average Shipping Paid: ${chalk.yellow(formatCurrency(metrics.averageShippingPaid))}`);

		// Profit metrics
		console.log(chalk.cyan('  Profit Metrics:'));
		console.log(`    Shipping Profit/Loss: ${formatProfitLoss(metrics.shippingProfit)}`);
		console.log(`    Shipping Profit Margin: ${formatProfitLossPercentage(metrics.shippingProfitMargin)}`);
		console.log(`    Net Revenue (Revenue - Shipping Cost): ${formatProfitLoss(metrics.netRevenue)}`);
		console.log(`    Net Revenue Margin: ${formatProfitLossPercentage(metrics.netRevenueMargin)}`);
	}

	// Calculate overall metrics
	const overallShippingProfitMargin = totalShippingPaid > 0 ? (totalShippingProfit / totalShippingPaid) * 100 : 0;
	const overallNetRevenueMargin = totalOrderValue > 0 ? (totalNetRevenue / totalOrderValue) * 100 : 0;

	// Display summary
	console.log(chalk.blue.bold('\nStore Summary:'));
	console.log(`  Total Stores: ${chalk.yellow(stores.length)}`);
	console.log(`  Total Orders: ${chalk.yellow(totalOrders)}`);

	console.log(chalk.cyan('\n  Order Summary:'));
	console.log(`    Total Order Value: ${chalk.yellow(formatCurrency(totalOrderValue))}`);
	console.log(`    Overall Average Order Value: ${chalk.yellow(formatCurrency(totalOrderValue / totalOrders))}`);

	console.log(chalk.cyan('\n  Shipping Summary:'));
	console.log(`    Total Shipping Cost: ${chalk.yellow(formatCurrency(totalRate))}`);
	console.log(`    Total Shipping Paid by Customers: ${chalk.yellow(formatCurrency(totalShippingPaid))}`);

	console.log(chalk.cyan('\n  Profit Summary:'));
	console.log(`    Total Shipping Profit/Loss: ${formatProfitLoss(totalShippingProfit)}`);
	console.log(`    Overall Shipping Profit Margin: ${formatProfitLossPercentage(overallShippingProfitMargin)}`);
	console.log(`    Total Net Revenue (Revenue - Shipping Cost): ${formatProfitLoss(totalNetRevenue)}`);
	console.log(`    Overall Net Revenue Margin: ${formatProfitLossPercentage(overallNetRevenueMargin)}`);
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

	// Display metrics for each tag
	for (const tag of tags) {
		const metrics = tagMetrics[tag];
		totalTaggedOrders += metrics.count;
		totalTagRate += metrics.totalRate;

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
	console.log(`  Overall Average Shipping Cost: ${chalk.yellow(formatCurrency(totalTagRate / totalTaggedOrders))}`);
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
