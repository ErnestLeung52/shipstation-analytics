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
 * @param {string} periodName - Period name for the report (e.g., "Feb 1-Mar 15, 2025")
 */
function displayComprehensiveStoreTable(
	storeMetrics,
	stores,
	totalOrders,
	totalOrderValue,
	totalRate,
	totalShippingPaid,
	totalShippingProfit,
	totalNetRevenue,
	periodName
) {
	// Use the period name directly for the table title
	const period = periodName || 'Current Period';

	console.log(chalk.cyan.bold(`\n${period} Store Shipping Analytics | ${period} 店铺物流分析`));

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
			chalk.white.bold('Metric | 指标'),
			...orderedStores.map((store) => chalk.white.bold(store)),
			chalk.white.bold('TOTAL | 总计'),
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
		// Use chalk.gray for the percentage part to create visual contrast
		return `${value} ${chalk.gray(`(${percent}%)`)}`;
	};

	// Add rows for each metric
	table.push(
		[
			'Orders | 订单数',
			...orderedStores.map((store) => formatWithPercent(storeMetrics[store].count, totalOrders)),
			chalk.bold(totalOrders),
		],
		[
			'Order Value | 订单价值',
			...orderedStores.map((store) =>
				formatWithPercent(formatCurrency(storeMetrics[store].totalOrderValue), totalOrderValue, true)
			),
			chalk.bold(formatCurrency(totalOrderValue)),
		],
		[
			'AOV | 平均订单价值',
			...orderedStores.map((store) => formatCurrency(storeMetrics[store].averageOrderValue)),
			chalk.bold(formatCurrency(totalOrderValue / totalOrders)),
		],
		[
			'Ship Cost | 物流成本',
			...orderedStores.map((store) =>
				formatWithPercent(formatCurrency(storeMetrics[store].totalRate), totalRate, true)
			),
			chalk.bold(formatCurrency(totalRate)),
		],
		[
			'Ship Paid | 物流收入',
			...orderedStores.map((store) =>
				formatWithPercent(formatCurrency(storeMetrics[store].totalShippingPaid), totalShippingPaid, true)
			),
			chalk.bold(formatCurrency(totalShippingPaid)),
		],
		[
			'Ship Profit | 物流利润',
			...orderedStores.map((store) => colorizeValue(formatCurrency(storeMetrics[store].shippingProfit))),
			chalk.bold(colorizeValue(formatCurrency(totalShippingProfit))),
		],
		[
			'Ship Margin | 物流利润率',
			...orderedStores.map((store) => colorizeValue(formatPercentage(storeMetrics[store].shippingProfitMargin))),
			chalk.bold(colorizeValue(formatPercentage(overallShippingProfitMargin))),
		],
		[
			'Net Revenue | 净收入',
			...orderedStores.map((store) =>
				formatWithPercent(colorizeValue(formatCurrency(storeMetrics[store].netRevenue)), totalNetRevenue, true)
			),
			chalk.bold(colorizeValue(formatCurrency(totalNetRevenue))),
		],
		[
			'Net Margin | 净利润率',
			...orderedStores.map((store) => colorizeValue(formatPercentage(storeMetrics[store].netRevenueMargin))),
			chalk.bold(colorizeValue(formatPercentage(overallNetRevenueMargin))),
		]
	);

	console.log(table.toString());
	console.log(
		chalk.gray(
			'AOV = Average Order Value | 平均订单价值, Ship = Shipping | 物流, Net Margin = Net Revenue Margin | 净利润率'
		)
	);
}

/**
 * Displays all special order metrics in a single comprehensive table
 * @param {Object} tagMetrics - Tag metrics object
 * @param {Array<string>} tags - Array of tag names
 * @param {number} totalTaggedOrders - Total number of tagged orders
 * @param {number} totalTagRate - Total shipping cost for tagged orders
 * @param {number} totalAllStoresOrders - Total number of orders across all stores
 */
function displayComprehensiveTagTable(tagMetrics, tags, totalTaggedOrders, totalTagRate, totalAllStoresOrders) {
	console.log(chalk.cyan.bold('\nSpecial Orders Analysis | 特殊订单分析'));

	// Create a map of tag descriptions in Chinese
	const tagDescriptions = {
		'Fulfillment Error': '仓库错误',
		Giveaways: '免费赠品',
		Influencer: '网红推广',
		'Not Delivered': '未送达',
		Replacement: '替换订单',
	};

	// Create table with metrics as rows and tags as columns
	const table = new Table({
		head: [
			chalk.white.bold('Metric | 指标'),
			...tags.map((tag) => chalk.white.bold(`${tag}\n${tagDescriptions[tag] || ''}`)),
			chalk.white.bold('TOTAL | 总计'),
		],
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
		['Orders | 订单数', ...tags.map((tag) => tagMetrics[tag].count), chalk.bold(totalTaggedOrders)],
		[
			'% of All Orders | 占总订单百分比',
			...percentOfTotalOrders.map((percent) => chalk.gray(`${percent}%`)),
			chalk.bold(chalk.gray(`${totalPercentOfAllOrders}%`)),
		],
		[
			'Total Shipping Cost | 总物流成本',
			...tags.map((tag) => formatCurrency(tagMetrics[tag].totalRate)),
			chalk.bold(formatCurrency(totalTagRate)),
		],
		[
			'Avg Shipping Cost | 平均物流成本',
			...tags.map((tag) => formatCurrency(tagMetrics[tag].averageRate)),
			chalk.bold(formatCurrency(totalTagRate / totalTaggedOrders)),
		]
	);

	console.log(table.toString());
	console.log(chalk.gray('% of All Orders = Orders with this special category / Total orders across all stores'));
	console.log(chalk.gray('占总订单百分比 = 特殊类别订单数 / 所有店铺总订单数'));
}

export function displayStoreMetrics(storeMetrics, periodName) {
	console.log(chalk.blue.bold('\n=== Store Metrics | 店铺指标 ==='));

	// Get stores and sort by order count (descending)
	const stores = Object.keys(storeMetrics).sort((a, b) => storeMetrics[b].count - storeMetrics[a].count);

	if (stores.length === 0) {
		console.log(chalk.yellow('No store data found | 未找到店铺数据'));
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
		totalNetRevenue,
		periodName
	);

	// Display legend and help text
	console.log(chalk.gray('\nLegend | 图例:'));
	console.log(chalk.gray('- AOV = Average Order Value | 平均订单价值'));
	console.log(chalk.gray('- Ship = Shipping | 物流'));
	console.log(chalk.gray('- Net Margin = Net Revenue / Order Value | 净利润率 = 净收入 / 订单价值'));
	console.log(chalk.gray('- Ship Margin = Shipping Profit / Shipping Paid | 物流利润率 = 物流利润 / 物流收入'));
	console.log(chalk.green('- Green values indicate profit | 绿色表示盈利'));
	console.log(chalk.red('- Red values indicate loss | 红色表示亏损'));
	console.log(chalk.yellow('- Yellow values indicate break-even | 黄色表示收支平衡'));

	// Display detailed metrics by section
	console.log(chalk.blue.bold('\nStores Summary | 店铺摘要:'));
	for (const store of stores) {
		const metrics = storeMetrics[store];
		console.log(chalk.cyan.bold(`\n${store}:`));

		// Order Summary
		console.log(
			chalk.white('Orders | 订单:'),
			chalk.yellow(`${metrics.count} orders`),
			chalk.gray(`(AOV: ${formatCurrency(metrics.averageOrderValue)})`)
		);

		// Revenue Summary
		console.log(
			chalk.white('Revenue | 收入:'),
			chalk.yellow(formatCurrency(metrics.totalOrderValue)),
			chalk.gray('→'),
			colorizeValue(formatCurrency(metrics.netRevenue)),
			chalk.gray(`(${colorizeValue(formatPercentage(metrics.netRevenueMargin))} margin)`)
		);

		// Shipping Summary
		console.log(
			chalk.white('Shipping | 物流:'),
			chalk.yellow(`Cost: ${formatCurrency(metrics.totalRate)}`),
			chalk.gray('vs'),
			chalk.yellow(`Paid: ${formatCurrency(metrics.totalShippingPaid)}`),
			chalk.gray('='),
			colorizeValue(formatCurrency(metrics.shippingProfit)),
			chalk.gray(`(${colorizeValue(formatPercentage(metrics.shippingProfitMargin))} margin)`)
		);
	}

	// Display overall summary
	console.log(chalk.blue.bold('\nOverall Summary | 总体摘要:'));
	console.log(chalk.white('Total Orders | 总订单数:'), chalk.yellow(totalOrders));
	console.log(
		chalk.white('Total Revenue | 总收入:'),
		chalk.yellow(formatCurrency(totalOrderValue)),
		chalk.gray('→'),
		colorizeValue(formatCurrency(totalNetRevenue)),
		chalk.gray(`(${colorizeValue(formatPercentage((totalNetRevenue / totalOrderValue) * 100))} margin)`)
	);
	console.log(
		chalk.white('Total Shipping | 总物流:'),
		chalk.yellow(`Cost: ${formatCurrency(totalRate)}`),
		chalk.gray('vs'),
		chalk.yellow(`Paid: ${formatCurrency(totalShippingPaid)}`),
		chalk.gray('='),
		colorizeValue(formatCurrency(totalShippingProfit)),
		chalk.gray(`(${colorizeValue(formatPercentage((totalShippingProfit / totalShippingPaid) * 100))} margin)`)
	);
}

export function displayTagMetrics(tagMetrics) {
	console.log(chalk.blue.bold('\n=== Special Orders Analysis | 特殊订单分析 ==='));

	// Get tags and sort alphabetically
	const tags = Object.keys(tagMetrics).sort();

	if (tags.length === 0) {
		console.log(chalk.yellow('No special orders data found | 未找到特殊订单数据'));
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
	console.log(chalk.gray('\nLegend | 图例:'));
	console.log(chalk.gray('- % of All Orders = Orders with this special category / Total orders across all stores'));
	console.log(chalk.gray('- 占总订单百分比 = 特殊类别订单数 / 所有店铺总订单数'));
	console.log(chalk.gray('- Avg Shipping Cost = Total shipping cost / Number of orders'));
	console.log(chalk.gray('- 平均物流成本 = 总物流成本 / 订单数'));

	// Explain each special order category
	console.log(chalk.gray('\nSpecial Order Categories | 特殊订单类别:'));
	console.log(chalk.gray('- Fulfillment Error | 仓库错误: Orders with errors made by warehouse staff'));
	console.log(chalk.gray('- Giveaways | 免费赠品: Free products given for promotional purposes'));
	console.log(chalk.gray('- Influencer | 网红推广: Orders sent to influencers for promotion'));
	console.log(chalk.gray('- Not Delivered | 未送达: Orders that were not delivered to customers'));
	console.log(chalk.gray('- Replacement | 替换订单: Replacement orders for damaged products'));

	// Display detailed metrics by section
	console.log(chalk.blue.bold('\nDetailed Special Orders Analysis | 详细特殊订单分析:'));
	for (const tag of tags) {
		const metrics = tagMetrics[tag];
		const percentOfOrders = ((metrics.count / totalTaggedOrders) * 100).toFixed(1);
		const percentOfCost = ((metrics.totalRate / totalTagRate) * 100).toFixed(1);
		const percentOfAllOrders = ((metrics.count / totalAllStoresOrders) * 100).toFixed(1);

		console.log(chalk.cyan.bold(`\n${tag} | ${getChineseTagName(tag)}:`));
		console.log(
			chalk.white('Orders | 订单:'),
			chalk.yellow(`${metrics.count} orders`),
			chalk.gray(`(${percentOfOrders}% of special orders, ${percentOfAllOrders}% of all orders)`)
		);
		console.log(
			chalk.white('Shipping | 物流:'),
			chalk.yellow(`Total: ${formatCurrency(metrics.totalRate)}`),
			chalk.gray(`(${percentOfCost}% of special orders cost)`),
			chalk.gray(`Avg: ${formatCurrency(metrics.averageRate)}`)
		);
	}

	// Display overall tag summary
	console.log(chalk.blue.bold('\nSpecial Orders Summary | 特殊订单摘要:'));
	console.log(
		chalk.white('Total Special Orders | 总特殊订单:'),
		chalk.yellow(totalTaggedOrders),
		chalk.gray(`(${((totalTaggedOrders / totalAllStoresOrders) * 100).toFixed(1)}% of all orders)`)
	);
	console.log(chalk.white('Total Shipping Cost | 总物流成本:'), chalk.yellow(formatCurrency(totalTagRate)));
	console.log(
		chalk.white('Average Cost per Order | 每单平均成本:'),
		chalk.yellow(formatCurrency(totalTagRate / totalTaggedOrders))
	);
	console.log(chalk.white('Unique Categories | 独特类别:'), chalk.yellow(tags.length));
}

/**
 * Helper function to get Chinese tag names
 */
function getChineseTagName(tag) {
	const tagMap = {
		'Fulfillment Error': '仓库错误',
		Giveaways: '免费赠品',
		Influencer: '网红推广',
		'Not Delivered': '未送达',
		Replacement: '替换订单',
	};
	return tagMap[tag] || tag;
}
