/**
 * Report Exporter
 *
 * This module provides functions to export metrics reports to CSV files.
 */

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import os from 'os';

// Get the directory name in ESM
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Saves store metrics and tag metrics to a CSV file
 * @param {Object} storeMetrics - Store metrics object
 * @param {Object} tagMetrics - Tag metrics object
 * @param {string} inputFileName - Name of the input file that was analyzed
 * @param {string} outputPath - Path to save the CSV file (optional)
 * @returns {Promise<string>} - Path to the saved file
 */
export async function saveReportToCSV(storeMetrics, tagMetrics, inputFileName, outputPath = null) {
	// Extract period from filename (e.g., "Feb-March 2025" from "./ShipStation Orders/Feb-March 2025.csv")
	const periodMatch = inputFileName ? inputFileName.match(/([^\/]+)\.csv$/) : null;
	const period = periodMatch ? periodMatch[1] : 'Current_Period';

	// Generate output filename
	const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);

	// Get the Downloads folder path
	const homeDir = os.homedir();
	const downloadsDir = path.join(homeDir, 'Downloads');

	// Create the output file path
	const outputFileName = outputPath || path.join(downloadsDir, `${period}_ShipStation_Report_${timestamp}.csv`);

	// Prepare CSV content
	let csvContent = [];

	// Add BOM (Byte Order Mark) for Excel to correctly recognize UTF-8
	csvContent.push('\ufeff');

	// Add report header
	csvContent.push(`"ShipStation Analytics Report for ${period}"`);
	csvContent.push('');

	// Add store metrics section
	csvContent.push('"STORE METRICS | 店铺指标"');
	csvContent.push('');

	// Get stores and sort by order count
	const stores = Object.keys(storeMetrics).sort((a, b) => storeMetrics[b].count - storeMetrics[a].count);

	if (stores.length === 0) {
		csvContent.push('"No store data found | 未找到店铺数据"');
	} else {
		// Calculate totals
		let totalOrders = 0;
		let totalRate = 0;
		let totalOrderValue = 0;
		let totalShippingPaid = 0;
		let totalShippingProfit = 0;
		let totalNetRevenue = 0;

		for (const store of stores) {
			const metrics = storeMetrics[store];
			totalOrders += metrics.count;
			totalRate += metrics.totalRate;
			totalOrderValue += metrics.totalOrderValue;
			totalShippingPaid += metrics.totalShippingPaid;
			totalShippingProfit += metrics.shippingProfit;
			totalNetRevenue += metrics.netRevenue;
		}

		// Add store metrics table header
		csvContent.push('"Metric | 指标","' + stores.join('","') + '","TOTAL | 总计"');

		// Add rows for each metric
		csvContent.push(
			`"Orders | 订单数","${stores
				.map((store) => {
					const percent = ((storeMetrics[store].count / totalOrders) * 100).toFixed(1);
					return `${storeMetrics[store].count} (${percent}%)`;
				})
				.join('","')}","${totalOrders}"`
		);

		csvContent.push(
			`"Order Value | 订单价值","${stores
				.map((store) => {
					const percent = ((storeMetrics[store].totalOrderValue / totalOrderValue) * 100).toFixed(1);
					return `$${storeMetrics[store].totalOrderValue.toFixed(2)} (${percent}%)`;
				})
				.join('","')}","$${totalOrderValue.toFixed(2)}"`
		);

		csvContent.push(
			`"AOV | 平均订单价值","${stores
				.map((store) => {
					return `$${storeMetrics[store].averageOrderValue.toFixed(2)}`;
				})
				.join('","')}","$${(totalOrderValue / totalOrders).toFixed(2)}"`
		);

		csvContent.push(
			`"Ship Cost | 物流成本","${stores
				.map((store) => {
					const percent = ((storeMetrics[store].totalRate / totalRate) * 100).toFixed(1);
					return `$${storeMetrics[store].totalRate.toFixed(2)} (${percent}%)`;
				})
				.join('","')}","$${totalRate.toFixed(2)}"`
		);

		csvContent.push(
			`"Ship Paid | 物流收入","${stores
				.map((store) => {
					const percent =
						totalShippingPaid > 0
							? ((storeMetrics[store].totalShippingPaid / totalShippingPaid) * 100).toFixed(1)
							: '0.0';
					return `$${storeMetrics[store].totalShippingPaid.toFixed(2)} (${percent}%)`;
				})
				.join('","')}","$${totalShippingPaid.toFixed(2)}"`
		);

		csvContent.push(
			`"Ship Profit | 物流利润","${stores
				.map((store) => {
					return `$${storeMetrics[store].shippingProfit.toFixed(2)}`;
				})
				.join('","')}","$${totalShippingProfit.toFixed(2)}"`
		);

		const overallShippingProfitMargin = totalShippingPaid > 0 ? (totalShippingProfit / totalShippingPaid) * 100 : 0;
		csvContent.push(
			`"Ship Margin | 物流利润率","${stores
				.map((store) => {
					return `${storeMetrics[store].shippingProfitMargin.toFixed(2)}%`;
				})
				.join('","')}","${overallShippingProfitMargin.toFixed(2)}%"`
		);

		csvContent.push(
			`"Net Revenue | 净收入","${stores
				.map((store) => {
					const percent = ((storeMetrics[store].netRevenue / totalNetRevenue) * 100).toFixed(1);
					return `$${storeMetrics[store].netRevenue.toFixed(2)} (${percent}%)`;
				})
				.join('","')}","$${totalNetRevenue.toFixed(2)}"`
		);

		const overallNetRevenueMargin = totalOrderValue > 0 ? (totalNetRevenue / totalOrderValue) * 100 : 0;
		csvContent.push(
			`"Net Margin | 净利润率","${stores
				.map((store) => {
					return `${storeMetrics[store].netRevenueMargin.toFixed(2)}%`;
				})
				.join('","')}","${overallNetRevenueMargin.toFixed(2)}%"`
		);

		// Add legend
		csvContent.push('');
		csvContent.push('"Legend | 图例:"');
		csvContent.push(
			'"AOV = Average Order Value | 平均订单价值, Ship = Shipping | 物流, Net Margin = Net Revenue Margin | 净利润率"'
		);
		csvContent.push('"- Green values indicate profit | 绿色表示盈利"');
		csvContent.push('"- Red values indicate loss | 红色表示亏损"');
		csvContent.push('"- Yellow values indicate break-even | 黄色表示收支平衡"');

		// Add store summary
		csvContent.push('');
		csvContent.push('"Stores Summary | 店铺摘要:"');

		for (const store of stores) {
			const metrics = storeMetrics[store];
			csvContent.push(`"${store}:"`);
			csvContent.push(`"Orders | 订单: ${metrics.count} orders (AOV: $${metrics.averageOrderValue.toFixed(2)})"`);
			csvContent.push(
				`"Revenue | 收入: $${metrics.totalOrderValue.toFixed(2)} → $${metrics.netRevenue.toFixed(
					2
				)} (${metrics.netRevenueMargin.toFixed(2)}% margin)"`
			);
			csvContent.push(
				`"Shipping | 物流: Cost: $${metrics.totalRate.toFixed(2)} vs Paid: $${metrics.totalShippingPaid.toFixed(
					2
				)} = $${metrics.shippingProfit.toFixed(2)} (${metrics.shippingProfitMargin.toFixed(2)}% margin)"`
			);
			csvContent.push('');
		}

		// Add overall summary
		csvContent.push('"Overall Summary | 总体摘要:"');
		csvContent.push(`"Total Orders | 总订单数: ${totalOrders}"`);
		csvContent.push(
			`"Total Revenue | 总收入: $${totalOrderValue.toFixed(2)} → $${totalNetRevenue.toFixed(
				2
			)} (${overallNetRevenueMargin.toFixed(2)}% margin)"`
		);
		csvContent.push(
			`"Total Shipping | 总物流: Cost: $${totalRate.toFixed(2)} vs Paid: $${totalShippingPaid.toFixed(
				2
			)} = $${totalShippingProfit.toFixed(2)} (${overallShippingProfitMargin.toFixed(2)}% margin)"`
		);
	}

	// Add special orders section
	csvContent.push('');
	csvContent.push('"SPECIAL ORDERS ANALYSIS | 特殊订单分析"');
	csvContent.push('');

	// Get tags and sort alphabetically
	const tags = Object.keys(tagMetrics).sort();

	if (tags.length === 0) {
		csvContent.push('"No special orders data found | 未找到特殊订单数据"');
	} else {
		// Calculate totals
		let totalTaggedOrders = 0;
		let totalTagRate = 0;

		for (const tag of tags) {
			const metrics = tagMetrics[tag];
			totalTaggedOrders += metrics.count;
			totalTagRate += metrics.totalRate;
		}

		// Create a map of tag descriptions in Chinese
		const tagDescriptions = {
			'Fulfillment Error': '仓库错误',
			Giveaways: '免费赠品',
			Influencer: '网红推广',
			'Not Delivered': '未送达',
			Replacement: '替换订单',
		};

		// Add tag metrics table header
		csvContent.push(
			`"Metric | 指标","${tags
				.map((tag) => `${tag} | ${tagDescriptions[tag] || ''}`)
				.join('","')}","TOTAL | 总计"`
		);

		// Add rows for each metric
		csvContent.push(
			`"Orders | 订单数","${tags.map((tag) => tagMetrics[tag].count).join('","')}","${totalTaggedOrders}"`
		);

		// Get total orders from all stores
		const totalAllStoresOrders =
			global.totalAllStoresOrders ||
			Object.keys(storeMetrics).reduce((sum, store) => sum + storeMetrics[store].count, 0);
		const percentOfTotalOrders = tags.map((tag) =>
			((tagMetrics[tag].count / totalAllStoresOrders) * 100).toFixed(1)
		);
		const totalPercentOfAllOrders = ((totalTaggedOrders / totalAllStoresOrders) * 100).toFixed(1);

		csvContent.push(
			`"% of All Orders | 占总订单百分比","${percentOfTotalOrders
				.map((percent) => `${percent}%`)
				.join('","')}","${totalPercentOfAllOrders}%"`
		);

		csvContent.push(
			`"Total Shipping Cost | 总物流成本","${tags
				.map((tag) => `$${tagMetrics[tag].totalRate.toFixed(2)}`)
				.join('","')}","$${totalTagRate.toFixed(2)}"`
		);

		csvContent.push(
			`"Avg Shipping Cost | 平均物流成本","${tags
				.map((tag) => `$${tagMetrics[tag].averageRate.toFixed(2)}`)
				.join('","')}","$${(totalTagRate / totalTaggedOrders).toFixed(2)}"`
		);

		// Add legend
		csvContent.push('');
		csvContent.push('"Legend | 图例:"');
		csvContent.push('"% of All Orders = Orders with this special category / Total orders across all stores"');
		csvContent.push('"占总订单百分比 = 特殊类别订单数 / 所有店铺总订单数"');
		csvContent.push('"Avg Shipping Cost = Total shipping cost / Number of orders"');
		csvContent.push('"平均物流成本 = 总物流成本 / 订单数"');

		// Add special order categories explanation
		csvContent.push('');
		csvContent.push('"Special Order Categories | 特殊订单类别:"');
		csvContent.push('"- Fulfillment Error | 仓库错误: Orders with errors made by warehouse staff"');
		csvContent.push('"- Giveaways | 免费赠品: Free products given for promotional purposes"');
		csvContent.push('"- Influencer | 网红推广: Orders sent to influencers for promotion"');
		csvContent.push('"- Not Delivered | 未送达: Orders that were not delivered to customers"');
		csvContent.push('"- Replacement | 替换订单: Replacement orders for damaged products"');

		// Add detailed tag analysis
		csvContent.push('');
		csvContent.push('"Detailed Special Orders Analysis | 详细特殊订单分析:"');

		for (const tag of tags) {
			const metrics = tagMetrics[tag];
			const percentOfOrders = ((metrics.count / totalTaggedOrders) * 100).toFixed(1);
			const percentOfCost = ((metrics.totalRate / totalTagRate) * 100).toFixed(1);
			const percentOfAllOrders = ((metrics.count / totalAllStoresOrders) * 100).toFixed(1);

			csvContent.push(`"${tag} | ${tagDescriptions[tag] || ''}:"`);
			csvContent.push(
				`"Orders | 订单: ${metrics.count} orders (${percentOfOrders}% of special orders, ${percentOfAllOrders}% of all orders)"`
			);
			csvContent.push(
				`"Shipping | 物流: Total: $${metrics.totalRate.toFixed(
					2
				)} (${percentOfCost}% of special orders cost) Avg: $${metrics.averageRate.toFixed(2)}"`
			);
			csvContent.push('');
		}

		// Add special orders summary
		csvContent.push('"Special Orders Summary | 特殊订单摘要:"');
		csvContent.push(
			`"Total Special Orders | 总特殊订单: ${totalTaggedOrders} (${totalPercentOfAllOrders}% of all orders)"`
		);
		csvContent.push(`"Total Shipping Cost | 总物流成本: $${totalTagRate.toFixed(2)}"`);
		csvContent.push(`"Average Cost per Order | 每单平均成本: $${(totalTagRate / totalTaggedOrders).toFixed(2)}"`);
		csvContent.push(`"Unique Categories | 独特类别: ${tags.length}"`);
	}

	// Write to file with UTF-8 encoding
	await fs.promises.writeFile(outputFileName, csvContent.join('\n'), { encoding: 'utf8' });

	return outputFileName;
}
