/**
 * Excel Report Exporter
 *
 * This module provides functions to export metrics reports to Excel files with proper formatting.
 */

import * as XLSX from 'xlsx';
import path from 'path';
import { fileURLToPath } from 'url';
import os from 'os';

// Get the directory name in ESM
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Saves store metrics and tag metrics to an Excel file
 * @param {Object} storeMetrics - Store metrics object
 * @param {Object} tagMetrics - Tag metrics object
 * @param {string} inputFileName - Name of the input file that was analyzed
 * @param {string} outputPath - Path to save the Excel file (optional)
 * @returns {Promise<string>} - Path to the saved file
 */
export async function saveReportToExcel(storeMetrics, tagMetrics, inputFileName, outputPath = null) {
	// Extract period from filename (e.g., "Feb-March 2025" from "./ShipStation Orders/Feb-March 2025.csv")
	const periodMatch = inputFileName ? inputFileName.match(/([^\/]+)\.csv$/) : null;
	const period = periodMatch ? periodMatch[1] : 'Current_Period';

	// Generate output filename
	const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);

	// Get the Downloads folder path
	const homeDir = os.homedir();
	const downloadsDir = path.join(homeDir, 'Downloads');

	// Create the output file path
	const outputFileName = outputPath || path.join(downloadsDir, `${period}_ShipStation_Report_${timestamp}.xlsx`);

	// Create a new workbook
	const workbook = XLSX.utils.book_new();

	// Add title worksheet with overview and instructions
	createTitleWorksheet(workbook, period, inputFileName);

	// Add store metrics worksheet
	createStoreMetricsWorksheet(workbook, storeMetrics, period, inputFileName);

	// Add special orders worksheet
	createSpecialOrdersWorksheet(workbook, tagMetrics, storeMetrics);

	// Write to file
	XLSX.writeFile(workbook, outputFileName);

	return outputFileName;
}

/**
 * Creates a title worksheet with overview and instructions
 * @param {Object} workbook - XLSX workbook
 * @param {string} period - Period name from the input file
 * @param {string} fileName - Name of the file being analyzed
 */
function createTitleWorksheet(workbook, period, fileName) {
	const currentDate = new Date().toLocaleDateString('en-US', {
		year: 'numeric',
		month: 'long',
		day: 'numeric',
	});

	const data = [
		[`ShipStation Analytics Report - ${period} | ShipStation 分析报告 - ${period}`],
		[`Generated on: ${currentDate} | 生成日期: ${currentDate}`],
		[`Source file: ${fileName} | 源文件: ${fileName}`],
		[],
		['REPORT OVERVIEW | 报告概述'],
		[],
		[
			'This report contains detailed shipping and order analytics for your ShipStation data. | 此报告包含您的 ShipStation 数据的详细物流和订单分析。',
		],
		['The report is organized into the following worksheets: | 报告分为以下工作表:'],
		[],
		['1. Store Metrics | 店铺指标'],
		['   - Comprehensive table of metrics by store | 按店铺划分的综合指标表'],
		['   - Order counts, values, and shipping costs | 订单数量、价值和物流成本'],
		['   - Profit margins and revenue analysis | 利润率和收入分析'],
		['   - Detailed breakdown by store | 按店铺的详细分析'],
		[],
		['2. Special Orders | 特殊订单'],
		['   - Analysis of orders with special tags | 带有特殊标签的订单分析'],
		['   - Breakdown by tag category | 按标签类别的分析'],
		['   - Shipping costs for special order types | 特殊订单类型的物流成本'],
		['   - Percentage of total orders | 占总订单的百分比'],
		[],
		['INSTRUCTIONS | 使用说明'],
		[],
		['- Green values indicate profit | 绿色表示盈利'],
		['- Red values indicate loss | 红色表示亏损'],
		['- Yellow values indicate break-even | 黄色表示收支平衡'],
		[],
		['Key Metrics Explained | 关键指标解释:'],
		[
			'- AOV (Average Order Value): Total order value divided by number of orders | 平均订单价值: 总订单价值除以订单数量',
		],
		['- Ship Margin: Shipping profit as a percentage of shipping paid | 物流利润率: 物流利润占物流收入的百分比'],
		['- Net Margin: Net revenue as a percentage of order value | 净利润率: 净收入占订单价值的百分比'],
		['- Ship Cost: Total shipping cost paid to carrier | 物流成本: 支付给物流公司的总成本'],
		['- Ship Paid: Total shipping fees collected from customers | 物流收入: 从客户处收取的总物流费用'],
		['- Ship Profit: Difference between shipping paid and shipping cost | 物流利润: 物流收入与物流成本之间的差额'],
		['- Net Revenue: Order value minus shipping cost | 净收入: 订单价值减去物流成本'],
		[],
		['For questions or support, contact your analytics team. | 如有问题或需要支持，请联系您的分析团队。'],
	];

	// Create worksheet
	const ws = XLSX.utils.aoa_to_sheet(data);

	// Set column widths
	ws['!cols'] = [{ wch: 100 }];

	// Apply formatting
	applyTitleWorksheetFormatting(ws, data.length);

	// Add worksheet to workbook
	XLSX.utils.book_append_sheet(workbook, ws, 'Overview | 概述');
}

/**
 * Applies formatting to the title worksheet
 * @param {Object} ws - XLSX worksheet
 * @param {number} rowCount - Number of rows
 */
function applyTitleWorksheetFormatting(ws, rowCount) {
	// Format title
	const titleRef = XLSX.utils.encode_cell({ r: 0, c: 0 });
	ws[titleRef].s = {
		font: { bold: true, sz: 16, color: { rgb: '0000FF' } },
		alignment: { horizontal: 'center', vertical: 'center' },
		border: {
			top: { style: 'thin', color: { rgb: '000000' } },
			bottom: { style: 'thin', color: { rgb: '000000' } },
			left: { style: 'thin', color: { rgb: '000000' } },
			right: { style: 'thin', color: { rgb: '000000' } },
		},
		fill: { patternType: 'solid', fgColor: { rgb: 'E6F0FF' } },
	};

	// Format date and source file
	for (let i = 1; i < 3; i++) {
		const cellRef = XLSX.utils.encode_cell({ r: i, c: 0 });
		ws[cellRef].s = {
			font: { italic: true, color: { rgb: '666666' } },
			alignment: { horizontal: 'center', vertical: 'center' },
			border: {
				top: { style: 'thin', color: { rgb: '000000' } },
				bottom: { style: 'thin', color: { rgb: '000000' } },
				left: { style: 'thin', color: { rgb: '000000' } },
				right: { style: 'thin', color: { rgb: '000000' } },
			},
		};
	}

	// Format section headers
	const sectionHeaders = [
		'REPORT OVERVIEW | 报告概述',
		'INSTRUCTIONS | 使用说明',
		'Key Metrics Explained | 关键指标解释:',
	];
	for (let r = 0; r < rowCount; r++) {
		const cellRef = XLSX.utils.encode_cell({ r, c: 0 });
		if (ws[cellRef] && sectionHeaders.includes(ws[cellRef].v)) {
			ws[cellRef].s = {
				font: { bold: true, sz: 12, color: { rgb: 'FFFFFF' } },
				fill: { patternType: 'solid', fgColor: { rgb: '4472C4' } },
				alignment: { horizontal: 'center', vertical: 'center' },
				border: {
					top: { style: 'thin', color: { rgb: '000000' } },
					bottom: { style: 'thin', color: { rgb: '000000' } },
					left: { style: 'thin', color: { rgb: '000000' } },
					right: { style: 'thin', color: { rgb: '000000' } },
				},
			};
		}
	}

	// Format worksheet titles
	const worksheetTitles = ['1. Store Metrics | 店铺指标', '2. Special Orders | 特殊订单'];
	for (let r = 0; r < rowCount; r++) {
		const cellRef = XLSX.utils.encode_cell({ r, c: 0 });
		if (ws[cellRef] && worksheetTitles.some((title) => ws[cellRef].v.includes(title))) {
			ws[cellRef].s = {
				font: { bold: true, sz: 12, color: { rgb: '0000FF' } },
				border: {
					top: { style: 'thin', color: { rgb: '000000' } },
					bottom: { style: 'thin', color: { rgb: '000000' } },
					left: { style: 'thin', color: { rgb: '000000' } },
					right: { style: 'thin', color: { rgb: '000000' } },
				},
				fill: { patternType: 'solid', fgColor: { rgb: 'E6F0FF' } },
			};
		}
	}

	// Format color indicators
	const colorIndicators = [
		'- Green values indicate profit | 绿色表示盈利',
		'- Red values indicate loss | 红色表示亏损',
		'- Yellow values indicate break-even | 黄色表示收支平衡',
	];

	for (let r = 0; r < rowCount; r++) {
		const cellRef = XLSX.utils.encode_cell({ r, c: 0 });
		if (ws[cellRef]) {
			if (ws[cellRef].v === colorIndicators[0]) {
				ws[cellRef].s = {
					font: { color: { rgb: '008000' }, sz: 11 },
					border: {
						top: { style: 'thin', color: { rgb: '000000' } },
						bottom: { style: 'thin', color: { rgb: '000000' } },
						left: { style: 'thin', color: { rgb: '000000' } },
						right: { style: 'thin', color: { rgb: '000000' } },
					},
				}; // Green
			} else if (ws[cellRef].v === colorIndicators[1]) {
				ws[cellRef].s = {
					font: { color: { rgb: 'FF0000' }, sz: 11 },
					border: {
						top: { style: 'thin', color: { rgb: '000000' } },
						bottom: { style: 'thin', color: { rgb: '000000' } },
						left: { style: 'thin', color: { rgb: '000000' } },
						right: { style: 'thin', color: { rgb: '000000' } },
					},
				}; // Red
			} else if (ws[cellRef].v === colorIndicators[2]) {
				ws[cellRef].s = {
					font: { color: { rgb: 'FFC000' }, sz: 11 },
					border: {
						top: { style: 'thin', color: { rgb: '000000' } },
						bottom: { style: 'thin', color: { rgb: '000000' } },
						left: { style: 'thin', color: { rgb: '000000' } },
						right: { style: 'thin', color: { rgb: '000000' } },
					},
				}; // Yellow
			}
		}
	}
}

/**
 * Creates a worksheet for store metrics
 * @param {Object} workbook - XLSX workbook
 * @param {Object} storeMetrics - Store metrics object
 * @param {string} period - Period name from the input file
 * @param {string} fileName - Name of the file being analyzed
 */
function createStoreMetricsWorksheet(workbook, storeMetrics, period, fileName) {
	// Get stores and sort by order count
	const stores = Object.keys(storeMetrics).sort((a, b) => storeMetrics[b].count - storeMetrics[a].count);

	if (stores.length === 0) {
		const ws = XLSX.utils.aoa_to_sheet([
			[`ShipStation Analytics Report for ${period} | ShipStation ${period} 分析报告`],
			[],
			['STORE METRICS | 店铺指标'],
			[],
			['No store data found | 未找到店铺数据'],
		]);
		XLSX.utils.book_append_sheet(workbook, ws, 'Store Metrics | 店铺指标');
		return;
	}

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

	// Create header rows
	const data = [
		[`ShipStation Analytics Report for ${period} | ShipStation ${period} 分析报告`],
		[],
		['STORE METRICS | 店铺指标'],
		[],
	];

	// Create main header row with store names
	const mainHeaderRow = ['Metric | 指标'];
	stores.forEach((store) => {
		mainHeaderRow.push(store); // Store name
		mainHeaderRow.push(''); // Empty cell for the percentage column
	});
	mainHeaderRow.push('TOTAL | 总计');
	mainHeaderRow.push(''); // Empty cell for the percentage column
	data.push(mainHeaderRow);

	// Create subheader row with Value/% labels
	const subHeaderRow = [''];
	stores.forEach(() => {
		subHeaderRow.push('Value');
		subHeaderRow.push('%');
	});
	subHeaderRow.push('Value');
	subHeaderRow.push('%');
	data.push(subHeaderRow);

	// Add rows for each metric
	// Orders row
	const ordersRow = ['Orders | 订单数'];
	stores.forEach((store) => {
		ordersRow.push(storeMetrics[store].count);
		ordersRow.push(((storeMetrics[store].count / totalOrders) * 100).toFixed(1) + '%');
	});
	ordersRow.push(totalOrders);
	ordersRow.push('100.0%');
	data.push(ordersRow);

	// Order Value row
	const orderValueRow = ['Order Value | 订单价值'];
	stores.forEach((store) => {
		orderValueRow.push({ v: storeMetrics[store].totalOrderValue, t: 'n', z: '$#,##0.00' });
		orderValueRow.push(((storeMetrics[store].totalOrderValue / totalOrderValue) * 100).toFixed(1) + '%');
	});
	orderValueRow.push({ v: totalOrderValue, t: 'n', z: '$#,##0.00' });
	orderValueRow.push('100.0%');
	data.push(orderValueRow);

	// AOV row
	const aovRow = ['AOV | 平均订单价值'];
	stores.forEach((store) => {
		aovRow.push({ v: storeMetrics[store].averageOrderValue, t: 'n', z: '$#,##0.00' });
		aovRow.push(''); // No percentage for AOV
	});
	aovRow.push({ v: totalOrderValue / totalOrders, t: 'n', z: '$#,##0.00' });
	aovRow.push(''); // No percentage for AOV
	data.push(aovRow);

	// Ship Cost row
	const shipCostRow = ['Ship Cost | 物流成本'];
	stores.forEach((store) => {
		shipCostRow.push({ v: storeMetrics[store].totalRate, t: 'n', z: '$#,##0.00' });
		shipCostRow.push(((storeMetrics[store].totalRate / totalRate) * 100).toFixed(1) + '%');
	});
	shipCostRow.push({ v: totalRate, t: 'n', z: '$#,##0.00' });
	shipCostRow.push('100.0%');
	data.push(shipCostRow);

	// Ship Paid row
	const shipPaidRow = ['Ship Paid | 物流收入'];
	stores.forEach((store) => {
		shipPaidRow.push({ v: storeMetrics[store].totalShippingPaid, t: 'n', z: '$#,##0.00' });
		const percent =
			totalShippingPaid > 0
				? ((storeMetrics[store].totalShippingPaid / totalShippingPaid) * 100).toFixed(1)
				: '0.0';
		shipPaidRow.push(percent + '%');
	});
	shipPaidRow.push({ v: totalShippingPaid, t: 'n', z: '$#,##0.00' });
	shipPaidRow.push('100.0%');
	data.push(shipPaidRow);

	// Ship Profit row
	const shipProfitRow = ['Ship Profit | 物流利润'];
	stores.forEach((store) => {
		shipProfitRow.push({ v: storeMetrics[store].shippingProfit, t: 'n', z: '$#,##0.00' });
		shipProfitRow.push(''); // No percentage for Ship Profit
	});
	shipProfitRow.push({ v: totalShippingProfit, t: 'n', z: '$#,##0.00' });
	shipProfitRow.push(''); // No percentage for Ship Profit
	data.push(shipProfitRow);

	// Ship Margin row
	const shipMarginRow = ['Ship Margin | 物流利润率'];
	const overallShippingProfitMargin = totalShippingPaid > 0 ? (totalShippingProfit / totalShippingPaid) * 100 : 0;
	stores.forEach((store) => {
		shipMarginRow.push({ v: storeMetrics[store].shippingProfitMargin, t: 'n', z: '0.00"%"' });
		shipMarginRow.push(''); // No percentage for Ship Margin
	});
	shipMarginRow.push({ v: overallShippingProfitMargin, t: 'n', z: '0.00"%"' });
	shipMarginRow.push(''); // No percentage for Ship Margin
	data.push(shipMarginRow);

	// Net Revenue row
	const netRevenueRow = ['Net Revenue | 净收入'];
	stores.forEach((store) => {
		netRevenueRow.push({ v: storeMetrics[store].netRevenue, t: 'n', z: '$#,##0.00' });
		netRevenueRow.push(((storeMetrics[store].netRevenue / totalNetRevenue) * 100).toFixed(1) + '%');
	});
	netRevenueRow.push({ v: totalNetRevenue, t: 'n', z: '$#,##0.00' });
	netRevenueRow.push('100.0%');
	data.push(netRevenueRow);

	// Net Margin row
	const netMarginRow = ['Net Margin | 净利润率'];
	const overallNetRevenueMargin = totalOrderValue > 0 ? (totalNetRevenue / totalOrderValue) * 100 : 0;
	stores.forEach((store) => {
		netMarginRow.push({ v: storeMetrics[store].netRevenueMargin, t: 'n', z: '0.00"%"' });
		netMarginRow.push(''); // No percentage for Net Margin
	});
	netMarginRow.push({ v: overallNetRevenueMargin, t: 'n', z: '0.00"%"' });
	netMarginRow.push(''); // No percentage for Net Margin
	data.push(netMarginRow);

	// Add legend
	data.push(
		[],
		['Legend | 图例:'],
		[
			'AOV = Average Order Value | 平均订单价值, Ship = Shipping | 物流, Net Margin = Net Revenue Margin | 净利润率',
		],
		['- Green values indicate profit | 绿色表示盈利'],
		['- Red values indicate loss | 红色表示亏损'],
		['- Yellow values indicate break-even | 黄色表示收支平衡'],
		['- Ship Cost = Total shipping cost paid to carrier | 物流成本 = 支付给物流公司的总成本'],
		['- Ship Paid = Total shipping fees collected from customers | 物流收入 = 从客户处收取的总物流费用'],
		[
			'- Ship Profit = Difference between shipping paid and shipping cost | 物流利润 = 物流收入与物流成本之间的差额',
		],
		['- Ship Margin = Shipping Profit / Shipping Paid | 物流利润率 = 物流利润 / 物流收入'],
		['- Net Revenue = Order value minus shipping cost | 净收入 = 订单价值减去物流成本'],
		['- Net Margin = Net Revenue / Order Value | 净利润率 = 净收入 / 订单价值']
	);

	// Add store summary
	data.push([], ['Stores Summary | 店铺摘要:']);

	for (const store of stores) {
		const metrics = storeMetrics[store];
		data.push(
			[`${store}:`],
			[`Orders | 订单: ${metrics.count} orders (AOV: $${metrics.averageOrderValue.toFixed(2)})`],
			[
				`Revenue | 收入: $${metrics.totalOrderValue.toFixed(2)} → $${metrics.netRevenue.toFixed(
					2
				)} (${metrics.netRevenueMargin.toFixed(2)}% margin)`,
			],
			[
				`Shipping | 物流: Cost: $${metrics.totalRate.toFixed(2)} vs Paid: $${metrics.totalShippingPaid.toFixed(
					2
				)} = $${metrics.shippingProfit.toFixed(2)} (${metrics.shippingProfitMargin.toFixed(2)}%)`,
			],
			[]
		);
	}

	// Add overall summary
	data.push(
		['Overall Summary | 总体摘要:'],
		[`Total Orders | 总订单数: ${totalOrders}`],
		[
			`Total Revenue | 总收入: $${totalOrderValue.toFixed(2)} → $${totalNetRevenue.toFixed(
				2
			)} (${overallNetRevenueMargin.toFixed(2)}%)`,
		],
		[
			`Total Shipping | 总物流: Cost: $${totalRate.toFixed(2)} vs Paid: $${totalShippingPaid.toFixed(
				2
			)} = $${totalShippingProfit.toFixed(2)} (${overallShippingProfitMargin.toFixed(2)}%)`,
		]
	);

	// Create worksheet
	const ws = XLSX.utils.aoa_to_sheet(data);

	// Set column widths
	const colWidths = [
		{ wch: 25 }, // Metric column
	];

	// Add column widths for each store (value and percentage columns)
	stores.forEach(() => {
		colWidths.push({ wch: 15 }); // Value column
		colWidths.push({ wch: 8 }); // Percentage column
	});

	// Add column widths for total columns
	colWidths.push({ wch: 15 }); // Total value column
	colWidths.push({ wch: 8 }); // Total percentage column

	ws['!cols'] = colWidths;

	// Create merged cells for store headers
	const merges = [];

	// Calculate the starting column for stores (after the metric column)
	let startCol = 1;

	// Merge cells for each store header
	stores.forEach((_, index) => {
		const colIndex = startCol + index * 2;
		merges.push({
			s: { r: 4, c: colIndex },
			e: { r: 4, c: colIndex + 1 },
		});
	});

	// Merge cells for the total header
	merges.push({
		s: { r: 4, c: startCol + stores.length * 2 },
		e: { r: 4, c: startCol + stores.length * 2 + 1 },
	});

	// Add merges to worksheet
	ws['!merges'] = merges;

	// Apply formatting
	applyWorksheetFormatting(ws, data.length, stores.length * 2 + 3);

	// Add worksheet to workbook
	XLSX.utils.book_append_sheet(workbook, ws, 'Store Metrics | 店铺指标');
}

/**
 * Creates a worksheet for special orders
 * @param {Object} workbook - XLSX workbook
 * @param {Object} tagMetrics - Tag metrics object
 * @param {Object} storeMetrics - Store metrics object (for total orders)
 */
function createSpecialOrdersWorksheet(workbook, tagMetrics, storeMetrics) {
	// Get tags and sort alphabetically
	const tags = Object.keys(tagMetrics).sort();

	if (tags.length === 0) {
		const ws = XLSX.utils.aoa_to_sheet([
			['SPECIAL ORDERS ANALYSIS | 特殊订单分析'],
			[],
			['No special orders data found | 未找到特殊订单数据'],
		]);
		XLSX.utils.book_append_sheet(workbook, ws, 'Special Orders | 特殊订单');
		return;
	}

	// Calculate totals
	let totalTaggedOrders = 0;
	let totalTagRate = 0;

	for (const tag of tags) {
		const metrics = tagMetrics[tag];
		totalTaggedOrders += metrics.count;
		totalTagRate += metrics.totalRate;
	}

	// Calculate total orders from all stores
	const totalAllStoresOrders = Object.keys(storeMetrics).reduce((sum, store) => sum + storeMetrics[store].count, 0);

	// Create a map of tag descriptions in Chinese
	const tagDescriptions = {
		'Fulfillment Error': '仓库错误',
		Giveaways: '免费赠品',
		Influencer: '网红推广',
		'Not Delivered': '未送达',
		Replacement: '替换订单',
	};

	// Create header rows
	const data = [['SPECIAL ORDERS ANALYSIS | 特殊订单分析'], []];

	// Create main header row with tag names
	const mainHeaderRow = ['Metric | 指标'];
	tags.forEach((tag) => {
		mainHeaderRow.push(`${tag} | ${tagDescriptions[tag] || ''}`);
	});
	mainHeaderRow.push('TOTAL | 总计');
	data.push(mainHeaderRow);

	// Add rows for each metric
	// Orders row
	const ordersRow = ['Orders | 订单数'];
	tags.forEach((tag) => {
		ordersRow.push(tagMetrics[tag].count);
	});
	ordersRow.push(totalTaggedOrders);
	data.push(ordersRow);

	// % of All Orders row
	const percentOfTotalOrders = tags.map((tag) => ((tagMetrics[tag].count / totalAllStoresOrders) * 100).toFixed(1));
	const totalPercentOfAllOrders = ((totalTaggedOrders / totalAllStoresOrders) * 100).toFixed(1);
	const percentRow = ['% of All Orders | 占总订单百分比'];
	tags.forEach((_, index) => {
		percentRow.push({ v: parseFloat(percentOfTotalOrders[index]), t: 'n', z: '0.0"%"' });
	});
	percentRow.push({ v: parseFloat(totalPercentOfAllOrders), t: 'n', z: '0.0"%"' });
	data.push(percentRow);

	// Total Shipping Cost row
	const costRow = ['Total Shipping Cost | 总物流成本'];
	tags.forEach((tag) => {
		costRow.push({ v: tagMetrics[tag].totalRate, t: 'n', z: '$#,##0.00' });
	});
	costRow.push({ v: totalTagRate, t: 'n', z: '$#,##0.00' });
	data.push(costRow);

	// Avg Shipping Cost row
	const avgRow = ['Avg Shipping Cost | 平均物流成本'];
	tags.forEach((tag) => {
		avgRow.push({ v: tagMetrics[tag].averageRate, t: 'n', z: '$#,##0.00' });
	});
	avgRow.push({ v: totalTagRate / totalTaggedOrders, t: 'n', z: '$#,##0.00' });
	data.push(avgRow);

	// Format legend as a single merged cell with rich text
	const legendText = [
		'Legend | 图例:',
		'% of All Orders = Orders with this special category / Total orders across all stores | 占总订单百分比 = 特殊类别订单数 / 所有店铺总订单数',
		'Avg Shipping Cost = Total shipping cost / Number of orders | 平均物流成本 = 总物流成本 / 订单数',
	].join('\n');
	data.push([]);
	data.push([legendText]);

	// Add special order categories explanation
	const categoriesText = [
		'Special Order Categories | 特殊订单类别:',
		'- Fulfillment Error | 仓库错误: Orders with errors made by warehouse staff',
		'- Giveaways | 免费赠品: Free products given for promotional purposes',
		'- Influencer | 网红推广: Orders sent to influencers for promotion',
		'- Not Delivered | 未送达: Orders that were not delivered to customers',
		'- Replacement | 替换订单: Replacement orders for damaged products',
	].join('\n');
	data.push([]);
	data.push([categoriesText]);

	// Add detailed tag analysis
	data.push([]);
	data.push(['Detailed Special Orders Analysis | 详细特殊订单分析:']);

	for (const tag of tags) {
		const metrics = tagMetrics[tag];
		const percentOfOrders = ((metrics.count / totalTaggedOrders) * 100).toFixed(1);
		const percentOfCost = ((metrics.totalRate / totalTagRate) * 100).toFixed(1);
		const percentOfAllOrders = ((metrics.count / totalAllStoresOrders) * 100).toFixed(1);

		data.push([`${tag} | ${tagDescriptions[tag] || ''}:`]);
		data.push([
			`Orders | 订单: ${metrics.count} orders (${percentOfOrders}% of special orders, ${percentOfAllOrders}% of all orders)`,
		]);
		data.push([
			`Shipping | 物流: Total: $${metrics.totalRate.toFixed(
				2
			)} (${percentOfCost}% of special orders cost) Avg: $${metrics.averageRate.toFixed(2)}`,
		]);
		data.push([]);
	}

	// Add special orders summary
	data.push(['Special Orders Summary | 特殊订单摘要:']);
	data.push([`Total Special Orders | 总特殊订单: ${totalTaggedOrders} (${totalPercentOfAllOrders}% of all orders)`]);
	data.push([`Total Shipping Cost | 总物流成本: $${totalTagRate.toFixed(2)}`]);
	data.push([`Average Cost per Order | 每单平均成本: $${(totalTagRate / totalTaggedOrders).toFixed(2)}`]);
	data.push([`Unique Categories | 独特类别: ${tags.length}`]);

	// Create worksheet
	const ws = XLSX.utils.aoa_to_sheet(data);

	// Set column widths
	const colWidths = [
		{ wch: 30 }, // Metric column
		...tags.map(() => ({ wch: 20 })), // Tag columns
		{ wch: 20 }, // Total column
	];
	ws['!cols'] = colWidths;

	// Apply special orders table formatting
	applySpecialOrdersFormatting(ws, data.length, tags.length + 2);

	// Add worksheet to workbook
	XLSX.utils.book_append_sheet(workbook, ws, 'Special Orders | 特殊订单');
}

/**
 * Applies formatting to a worksheet
 * @param {Object} ws - XLSX worksheet
 * @param {number} rowCount - Number of rows
 * @param {number} colCount - Number of columns
 */
function applyWorksheetFormatting(ws, rowCount, colCount) {
	// Set header styles
	for (let i = 0; i < 5; i++) {
		const cellRef = XLSX.utils.encode_cell({ r: i, c: 0 });
		if (ws[cellRef]) {
			ws[cellRef].s = {
				font: { bold: true, sz: 12, color: { rgb: '0000FF' } },
				border: {
					top: { style: 'thin', color: { rgb: '000000' } },
					bottom: { style: 'thin', color: { rgb: '000000' } },
					left: { style: 'thin', color: { rgb: '000000' } },
					right: { style: 'thin', color: { rgb: '000000' } },
				},
			};
		}
	}

	// Set main header row styles (store names)
	for (let c = 0; c < colCount; c++) {
		const cellRef = XLSX.utils.encode_cell({ r: 4, c });
		if (ws[cellRef]) {
			ws[cellRef].s = {
				font: { bold: true, sz: 12, color: { rgb: 'FFFFFF' } },
				fill: { patternType: 'solid', fgColor: { rgb: '4472C4' } },
				alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
				border: {
					top: { style: 'thin', color: { rgb: '000000' } },
					bottom: { style: 'thin', color: { rgb: '000000' } },
					left: { style: 'thin', color: { rgb: '000000' } },
					right: { style: 'thin', color: { rgb: '000000' } },
				},
			};
		}
	}

	// Set subheader row styles (Value/%)
	for (let c = 1; c < colCount; c++) {
		const cellRef = XLSX.utils.encode_cell({ r: 5, c });
		if (ws[cellRef]) {
			ws[cellRef].s = {
				font: { bold: true, italic: true, sz: 11, color: { rgb: 'FFFFFF' } },
				fill: { patternType: 'solid', fgColor: { rgb: '5B9BD5' } },
				alignment: { horizontal: 'center', vertical: 'center' },
				border: {
					top: { style: 'thin', color: { rgb: '000000' } },
					bottom: { style: 'thin', color: { rgb: '000000' } },
					left: { style: 'thin', color: { rgb: '000000' } },
					right: { style: 'thin', color: { rgb: '000000' } },
				},
			};
		}
	}

	// Set metric column styles
	for (let r = 6; r < 15; r++) {
		const cellRef = XLSX.utils.encode_cell({ r, c: 0 });
		if (ws[cellRef]) {
			ws[cellRef].s = {
				font: { bold: true, sz: 11 },
				fill: { patternType: 'solid', fgColor: { rgb: 'D9E1F2' } },
				border: {
					top: { style: 'thin', color: { rgb: '000000' } },
					bottom: { style: 'thin', color: { rgb: '000000' } },
					left: { style: 'thin', color: { rgb: '000000' } },
					right: { style: 'thin', color: { rgb: '000000' } },
				},
			};
		}
	}

	// Apply alternating row colors and preserve colored text
	for (let r = 6; r < 15; r++) {
		// Apply alternating row background
		const rowColor = r % 2 === 0 ? 'F5F5F5' : 'FFFFFF';

		for (let c = 1; c < colCount; c++) {
			const cellRef = XLSX.utils.encode_cell({ r, c });
			if (ws[cellRef]) {
				// Initialize style if not exists
				if (!ws[cellRef].s) ws[cellRef].s = {};

				// Apply row background and border
				ws[cellRef].s.fill = { patternType: 'solid', fgColor: { rgb: rowColor } };
				ws[cellRef].s.border = {
					top: { style: 'thin', color: { rgb: '000000' } },
					bottom: { style: 'thin', color: { rgb: '000000' } },
					left: { style: 'thin', color: { rgb: '000000' } },
					right: { style: 'thin', color: { rgb: '000000' } },
				};

				// Apply special formatting for profit/loss values
				if (c % 2 === 1) {
					// Value columns
					const cellValue = ws[cellRef].v;
					if (typeof cellValue === 'number') {
						// Apply color based on value
						if (cellValue < 0) {
							ws[cellRef].s.font = { sz: 11, color: { rgb: 'FF0000' } }; // Red for negative
						} else if (cellValue > 0) {
							ws[cellRef].s.font = { sz: 11, color: { rgb: '008000' } }; // Green for positive
						} else {
							ws[cellRef].s.font = { sz: 11, color: { rgb: 'FFC000' } }; // Yellow for zero
						}
					}
				}

				// Center percentage columns
				if (c % 2 === 0 && ws[cellRef].v) {
					// Percentage columns
					ws[cellRef].s.font = { italic: true, sz: 11, color: { rgb: '666666' } };
					ws[cellRef].s.alignment = { horizontal: 'center', vertical: 'center' };
				}
			}
		}
	}

	// Set total column styles
	const totalValueCol = colCount - 2;
	const totalPercentCol = colCount - 1;
	for (let r = 6; r < 15; r++) {
		// Total value column
		const valueRef = XLSX.utils.encode_cell({ r, c: totalValueCol });
		if (ws[valueRef]) {
			ws[valueRef].s = {
				font: { bold: true, sz: 11 },
				fill: { patternType: 'solid', fgColor: { rgb: 'D9E1F2' } },
				border: {
					top: { style: 'thin', color: { rgb: '000000' } },
					bottom: { style: 'thin', color: { rgb: '000000' } },
					left: { style: 'thin', color: { rgb: '000000' } },
					right: { style: 'thin', color: { rgb: '000000' } },
				},
			};

			// Apply color based on value
			const cellValue = ws[valueRef].v;
			if (typeof cellValue === 'number') {
				if (cellValue < 0) {
					ws[valueRef].s.font.color = { rgb: 'FF0000' }; // Red for negative
				} else if (cellValue > 0) {
					ws[valueRef].s.font.color = { rgb: '008000' }; // Green for positive
				} else {
					ws[valueRef].s.font.color = { rgb: 'FFC000' }; // Yellow for zero
				}
			}
		}

		// Total percentage column
		const percentRef = XLSX.utils.encode_cell({ r, c: totalPercentCol });
		if (ws[percentRef] && ws[percentRef].v) {
			ws[percentRef].s = {
				font: { bold: true, italic: true, sz: 11, color: { rgb: '666666' } },
				alignment: { horizontal: 'center', vertical: 'center' },
				fill: { patternType: 'solid', fgColor: { rgb: 'D9E1F2' } },
				border: {
					top: { style: 'thin', color: { rgb: '000000' } },
					bottom: { style: 'thin', color: { rgb: '000000' } },
					left: { style: 'thin', color: { rgb: '000000' } },
					right: { style: 'thin', color: { rgb: '000000' } },
				},
			};
		}
	}

	// Format legend as a single merged cell with rich text
	const legendStartRow = 15;
	const legendText = [
		'Legend | 图例:',
		'AOV = Average Order Value | 平均订单价值, Ship = Shipping | 物流, Net Margin = Net Revenue Margin | 净利润率',
		'Green values indicate profit | 绿色表示盈利, Red values indicate loss | 红色表示亏损, Yellow values indicate break-even | 黄色表示收支平衡',
	].join('\n');

	// Create legend cell
	const legendCellRef = XLSX.utils.encode_cell({ r: legendStartRow, c: 0 });
	ws[legendCellRef] = {
		v: legendText,
		t: 's',
		s: {
			font: { bold: true, sz: 11, color: { rgb: '0000FF' } },
			alignment: { wrapText: true, vertical: 'top' },
			border: {
				top: { style: 'thin', color: { rgb: '000000' } },
				bottom: { style: 'thin', color: { rgb: '000000' } },
				left: { style: 'thin', color: { rgb: '000000' } },
				right: { style: 'thin', color: { rgb: '000000' } },
			},
		},
	};

	// Merge legend cells across all columns
	if (!ws['!merges']) ws['!merges'] = [];
	ws['!merges'].push({
		s: { r: legendStartRow, c: 0 },
		e: { r: legendStartRow + 2, c: colCount - 1 },
	});

	// Set section header styles for store summary and overall summary
	const sectionHeaders = ['Stores Summary | 店铺摘要:', 'Overall Summary | 总体摘要:'];
	for (let r = legendStartRow + 3; r < rowCount; r++) {
		const cellRef = XLSX.utils.encode_cell({ r, c: 0 });
		if (ws[cellRef] && sectionHeaders.includes(ws[cellRef].v)) {
			ws[cellRef].s = {
				font: { bold: true, sz: 12, color: { rgb: 'FFFFFF' } },
				fill: { patternType: 'solid', fgColor: { rgb: '4472C4' } },
				border: {
					top: { style: 'thin', color: { rgb: '000000' } },
					bottom: { style: 'thin', color: { rgb: '000000' } },
					left: { style: 'thin', color: { rgb: '000000' } },
					right: { style: 'thin', color: { rgb: '000000' } },
				},
			};

			// Merge section header across all columns
			if (!ws['!merges']) ws['!merges'] = [];
			ws['!merges'].push({
				s: { r, c: 0 },
				e: { r, c: colCount - 1 },
			});
		}
	}
}

/**
 * Applies formatting to the special orders worksheet
 * @param {Object} ws - XLSX worksheet
 * @param {number} rowCount - Number of rows
 * @param {number} colCount - Number of columns
 */
function applySpecialOrdersFormatting(ws, rowCount, colCount) {
	// Set header styles
	for (let i = 0; i < 3; i++) {
		const cellRef = XLSX.utils.encode_cell({ r: i, c: 0 });
		if (ws[cellRef]) {
			ws[cellRef].s = {
				font: { bold: true, sz: 12, color: { rgb: '0000FF' } },
				border: {
					top: { style: 'thin', color: { rgb: '000000' } },
					bottom: { style: 'thin', color: { rgb: '000000' } },
					left: { style: 'thin', color: { rgb: '000000' } },
					right: { style: 'thin', color: { rgb: '000000' } },
				},
			};
		}
	}

	// Set main header row styles (tag names)
	for (let c = 0; c < colCount; c++) {
		const cellRef = XLSX.utils.encode_cell({ r: 2, c });
		if (ws[cellRef]) {
			ws[cellRef].s = {
				font: { bold: true, sz: 12, color: { rgb: 'FFFFFF' } },
				fill: { patternType: 'solid', fgColor: { rgb: '4472C4' } },
				alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
				border: {
					top: { style: 'thin', color: { rgb: '000000' } },
					bottom: { style: 'thin', color: { rgb: '000000' } },
					left: { style: 'thin', color: { rgb: '000000' } },
					right: { style: 'thin', color: { rgb: '000000' } },
				},
			};
		}
	}

	// Set metric column styles
	for (let r = 3; r < 7; r++) {
		const cellRef = XLSX.utils.encode_cell({ r, c: 0 });
		if (ws[cellRef]) {
			ws[cellRef].s = {
				font: { bold: true, sz: 11 },
				fill: { patternType: 'solid', fgColor: { rgb: 'D9E1F2' } },
				border: {
					top: { style: 'thin', color: { rgb: '000000' } },
					bottom: { style: 'thin', color: { rgb: '000000' } },
					left: { style: 'thin', color: { rgb: '000000' } },
					right: { style: 'thin', color: { rgb: '000000' } },
				},
			};
		}
	}

	// Apply alternating row colors
	for (let r = 3; r < 7; r++) {
		// Apply alternating row background
		const rowColor = r % 2 === 0 ? 'F5F5F5' : 'FFFFFF';

		for (let c = 1; c < colCount; c++) {
			const cellRef = XLSX.utils.encode_cell({ r, c });
			if (ws[cellRef]) {
				// Initialize style if not exists
				if (!ws[cellRef].s) ws[cellRef].s = {};

				// Apply row background and border
				ws[cellRef].s.fill = { patternType: 'solid', fgColor: { rgb: rowColor } };
				ws[cellRef].s.border = {
					top: { style: 'thin', color: { rgb: '000000' } },
					bottom: { style: 'thin', color: { rgb: '000000' } },
					left: { style: 'thin', color: { rgb: '000000' } },
					right: { style: 'thin', color: { rgb: '000000' } },
				};
				ws[cellRef].s.alignment = { horizontal: 'center', vertical: 'center' };
				ws[cellRef].s.font = { sz: 11 };
			}
		}
	}

	// Set total column styles
	const totalCol = colCount - 1;
	for (let r = 3; r < 7; r++) {
		const cellRef = XLSX.utils.encode_cell({ r, c: totalCol });
		if (ws[cellRef]) {
			ws[cellRef].s = {
				font: { bold: true, sz: 11 },
				fill: { patternType: 'solid', fgColor: { rgb: 'D9E1F2' } },
				alignment: { horizontal: 'center', vertical: 'center' },
				border: {
					top: { style: 'thin', color: { rgb: '000000' } },
					bottom: { style: 'thin', color: { rgb: '000000' } },
					left: { style: 'thin', color: { rgb: '000000' } },
					right: { style: 'thin', color: { rgb: '000000' } },
				},
			};
		}
	}

	// Format legend as a single merged cell
	const legendRow = 8;
	const legendCellRef = XLSX.utils.encode_cell({ r: legendRow, c: 0 });
	if (ws[legendCellRef]) {
		ws[legendCellRef].s = {
			font: { bold: true, sz: 11, color: { rgb: '0000FF' } },
			alignment: { wrapText: true, vertical: 'top' },
			border: {
				top: { style: 'thin', color: { rgb: '000000' } },
				bottom: { style: 'thin', color: { rgb: '000000' } },
				left: { style: 'thin', color: { rgb: '000000' } },
				right: { style: 'thin', color: { rgb: '000000' } },
			},
			fill: { patternType: 'solid', fgColor: { rgb: 'E6F0FF' } },
		};

		// Merge legend cells across all columns
		if (!ws['!merges']) ws['!merges'] = [];
		ws['!merges'].push({
			s: { r: legendRow, c: 0 },
			e: { r: legendRow, c: colCount - 1 },
		});
	}

	// Format categories as a single merged cell
	const categoriesRow = 10;
	const categoriesCellRef = XLSX.utils.encode_cell({ r: categoriesRow, c: 0 });
	if (ws[categoriesCellRef]) {
		ws[categoriesCellRef].s = {
			font: { bold: true, sz: 11, color: { rgb: '0000FF' } },
			alignment: { wrapText: true, vertical: 'top' },
			border: {
				top: { style: 'thin', color: { rgb: '000000' } },
				bottom: { style: 'thin', color: { rgb: '000000' } },
				left: { style: 'thin', color: { rgb: '000000' } },
				right: { style: 'thin', color: { rgb: '000000' } },
			},
			fill: { patternType: 'solid', fgColor: { rgb: 'E6F0FF' } },
		};

		// Merge categories cells across all columns
		if (!ws['!merges']) ws['!merges'] = [];
		ws['!merges'].push({
			s: { r: categoriesRow, c: 0 },
			e: { r: categoriesRow, c: colCount - 1 },
		});
	}

	// Set section header styles
	const sectionHeaders = [
		'Detailed Special Orders Analysis | 详细特殊订单分析:',
		'Special Orders Summary | 特殊订单摘要:',
	];
	for (let r = 12; r < rowCount; r++) {
		const cellRef = XLSX.utils.encode_cell({ r, c: 0 });
		if (ws[cellRef] && sectionHeaders.includes(ws[cellRef].v)) {
			ws[cellRef].s = {
				font: { bold: true, sz: 12, color: { rgb: 'FFFFFF' } },
				fill: { patternType: 'solid', fgColor: { rgb: '4472C4' } },
				border: {
					top: { style: 'thin', color: { rgb: '000000' } },
					bottom: { style: 'thin', color: { rgb: '000000' } },
					left: { style: 'thin', color: { rgb: '000000' } },
					right: { style: 'thin', color: { rgb: '000000' } },
				},
			};

			// Merge section header across all columns
			if (!ws['!merges']) ws['!merges'] = [];
			ws['!merges'].push({
				s: { r, c: 0 },
				e: { r, c: colCount - 1 },
			});
		}
	}

	// Format tag headers
	for (let r = 13; r < rowCount; r++) {
		const cellRef = XLSX.utils.encode_cell({ r, c: 0 });
		if (ws[cellRef] && ws[cellRef].v && ws[cellRef].v.includes(' | ') && ws[cellRef].v.includes(':')) {
			ws[cellRef].s = {
				font: { bold: true, sz: 11, color: { rgb: '000000' } },
				fill: { patternType: 'solid', fgColor: { rgb: 'D9E1F2' } },
				border: {
					top: { style: 'thin', color: { rgb: '000000' } },
					bottom: { style: 'thin', color: { rgb: '000000' } },
					left: { style: 'thin', color: { rgb: '000000' } },
					right: { style: 'thin', color: { rgb: '000000' } },
				},
			};
		}
	}
}
