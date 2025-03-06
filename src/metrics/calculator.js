/**
 * Metrics Calculator
 *
 * This module provides functions to calculate various metrics from ShipStation data.
 * It includes calculations for store-based metrics and tag-based metrics.
 */

/**
 * Calculates metrics grouped by store
 * @param {Array<Object>} data - Array of ShipStation order data
 * @returns {Object} - Object with store metrics
 */
export function calculateStoreMetrics(data) {
	// Initialize results object
	const storeMetrics = {};

	// Process each order
	for (const order of data) {
		const store = order.Store || 'Unknown';
		const rate = typeof order.Rate === 'number' ? order.Rate : parseFloat(order.Rate) || 0;

		// Initialize store data if it doesn't exist
		if (!storeMetrics[store]) {
			storeMetrics[store] = {
				count: 0,
				totalRate: 0,
				averageRate: 0,
			};
		}

		// Update metrics
		storeMetrics[store].count += 1;
		storeMetrics[store].totalRate += rate;
	}

	// Calculate average rates
	for (const store in storeMetrics) {
		const metrics = storeMetrics[store];
		metrics.averageRate = metrics.count > 0 ? metrics.totalRate / metrics.count : 0;

		// Round to 2 decimal places for currency
		metrics.totalRate = parseFloat(metrics.totalRate.toFixed(2));
		metrics.averageRate = parseFloat(metrics.averageRate.toFixed(2));
	}

	return storeMetrics;
}

/**
 * Calculates metrics grouped by tags
 * @param {Array<Object>} data - Array of ShipStation order data
 * @returns {Object} - Object with tag metrics
 */
export function calculateTagMetrics(data) {
	// Initialize results object
	const tagMetrics = {};

	// Process each order
	for (const order of data) {
		// Skip if no tags
		if (!order.Tags || order.Tags.trim() === '') {
			continue;
		}

		// Split tags (they might be comma-separated)
		const tags = order.Tags.split(',')
			.map((tag) => tag.trim())
			.filter((tag) => tag);
		const rate = typeof order.Rate === 'number' ? order.Rate : parseFloat(order.Rate) || 0;

		// Process each tag
		for (const tag of tags) {
			// Initialize tag data if it doesn't exist
			if (!tagMetrics[tag]) {
				tagMetrics[tag] = {
					count: 0,
					totalRate: 0,
					averageRate: 0,
				};
			}

			// Update metrics
			tagMetrics[tag].count += 1;
			tagMetrics[tag].totalRate += rate;
		}
	}

	// Calculate average rates
	for (const tag in tagMetrics) {
		const metrics = tagMetrics[tag];
		metrics.averageRate = metrics.count > 0 ? metrics.totalRate / metrics.count : 0;

		// Round to 2 decimal places for currency
		metrics.totalRate = parseFloat(metrics.totalRate.toFixed(2));
		metrics.averageRate = parseFloat(metrics.averageRate.toFixed(2));
	}

	return tagMetrics;
}
