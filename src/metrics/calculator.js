/**
 * Metrics Calculator
 *
 * This module provides functions to calculate various metrics from ShipStation data.
 * It includes calculations for store-based metrics and tag-based metrics.
 */

/**
 * Safely extracts a numeric value from a field that might be a string or number
 * @param {any} value - The value to extract a number from
 * @returns {number} - The extracted number or 0 if invalid
 */
function extractNumericValue(value) {
	if (value === undefined || value === null) {
		return 0;
	}

	if (typeof value === 'number') {
		return value;
	}

	if (typeof value === 'string') {
		// Remove currency symbols and convert to number
		return parseFloat(value.replace(/[^0-9.-]+/g, '')) || 0;
	}

	return 0;
}

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
		const rate = extractNumericValue(order.Rate);

		// Extract order total (assuming it's in a field called "Order Total" or similar)
		const orderTotal = extractNumericValue(order['Order Total']) || extractNumericValue(order['OrderTotal']) || 0;

		// Extract shipping paid by customer (assuming it's in a field called "Shipping" or similar)
		const shippingPaid = extractNumericValue(order['Shipping']) || extractNumericValue(order['Shipping Paid']) || 0;

		// Initialize store data if it doesn't exist
		if (!storeMetrics[store]) {
			storeMetrics[store] = {
				count: 0,
				totalRate: 0,
				averageRate: 0,
				totalOrderValue: 0,
				averageOrderValue: 0,
				totalShippingPaid: 0,
				averageShippingPaid: 0,
				shippingProfit: 0,
				shippingProfitMargin: 0,
				netRevenue: 0,
				netRevenueMargin: 0,
			};
		}

		// Update metrics
		storeMetrics[store].count += 1;
		storeMetrics[store].totalRate += rate;
		storeMetrics[store].totalOrderValue += orderTotal;
		storeMetrics[store].totalShippingPaid += shippingPaid;
	}

	// Calculate averages and profit metrics
	for (const store in storeMetrics) {
		const metrics = storeMetrics[store];

		// Calculate averages
		metrics.averageRate = metrics.count > 0 ? metrics.totalRate / metrics.count : 0;
		metrics.averageOrderValue = metrics.count > 0 ? metrics.totalOrderValue / metrics.count : 0;
		metrics.averageShippingPaid = metrics.count > 0 ? metrics.totalShippingPaid / metrics.count : 0;

		// Calculate shipping profit (what customer paid minus what we paid)
		metrics.shippingProfit = metrics.totalShippingPaid - metrics.totalRate;

		// Calculate shipping profit margin as a percentage
		metrics.shippingProfitMargin =
			metrics.totalShippingPaid > 0 ? (metrics.shippingProfit / metrics.totalShippingPaid) * 100 : 0;

		// Calculate net revenue (order value minus shipping cost)
		// This represents the revenue available after shipping expenses
		metrics.netRevenue = metrics.totalOrderValue - metrics.totalRate;

		// Calculate net revenue margin as a percentage
		metrics.netRevenueMargin =
			metrics.totalOrderValue > 0 ? (metrics.netRevenue / metrics.totalOrderValue) * 100 : 0;

		// Round to 2 decimal places for currency
		metrics.totalRate = parseFloat(metrics.totalRate.toFixed(2));
		metrics.averageRate = parseFloat(metrics.averageRate.toFixed(2));
		metrics.totalOrderValue = parseFloat(metrics.totalOrderValue.toFixed(2));
		metrics.averageOrderValue = parseFloat(metrics.averageOrderValue.toFixed(2));
		metrics.totalShippingPaid = parseFloat(metrics.totalShippingPaid.toFixed(2));
		metrics.averageShippingPaid = parseFloat(metrics.averageShippingPaid.toFixed(2));
		metrics.shippingProfit = parseFloat(metrics.shippingProfit.toFixed(2));
		metrics.shippingProfitMargin = parseFloat(metrics.shippingProfitMargin.toFixed(2));
		metrics.netRevenue = parseFloat(metrics.netRevenue.toFixed(2));
		metrics.netRevenueMargin = parseFloat(metrics.netRevenueMargin.toFixed(2));
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

		const rate = extractNumericValue(order.Rate);

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

		// Calculate average rate
		metrics.averageRate = metrics.count > 0 ? metrics.totalRate / metrics.count : 0;

		// Round to 2 decimal places for currency
		metrics.totalRate = parseFloat(metrics.totalRate.toFixed(2));
		metrics.averageRate = parseFloat(metrics.averageRate.toFixed(2));
	}

	return tagMetrics;
}
