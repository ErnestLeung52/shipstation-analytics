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

	// Display metrics for each store
	for (const store of stores) {
		const metrics = storeMetrics[store];
		totalOrders += metrics.count;
		totalRate += metrics.totalRate;

		console.log(chalk.green(`\nStore: ${chalk.bold(store)}`));
		console.log(`  Orders: ${chalk.yellow(metrics.count)}`);
		console.log(`  Total Rate: ${chalk.yellow(formatCurrency(metrics.totalRate))}`);
		console.log(`  Average Rate: ${chalk.yellow(formatCurrency(metrics.averageRate))}`);
	}

	// Display summary
	console.log(chalk.blue.bold('\nStore Summary:'));
	console.log(`  Total Stores: ${chalk.yellow(stores.length)}`);
	console.log(`  Total Orders: ${chalk.yellow(totalOrders)}`);
	console.log(`  Total Rate: ${chalk.yellow(formatCurrency(totalRate))}`);
	console.log(`  Overall Average Rate: ${chalk.yellow(formatCurrency(totalRate / totalOrders))}`);
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
		console.log(`  Total Rate: ${chalk.yellow(formatCurrency(metrics.totalRate))}`);
		console.log(`  Average Rate: ${chalk.yellow(formatCurrency(metrics.averageRate))}`);
	}

	// Display summary
	console.log(chalk.blue.bold('\nTag Summary:'));
	console.log(`  Total Unique Tags: ${chalk.yellow(tags.length)}`);
	console.log(`  Total Tagged Orders: ${chalk.yellow(totalTaggedOrders)}`);
	console.log(`  Total Rate for Tagged Orders: ${chalk.yellow(formatCurrency(totalTagRate))}`);
}
