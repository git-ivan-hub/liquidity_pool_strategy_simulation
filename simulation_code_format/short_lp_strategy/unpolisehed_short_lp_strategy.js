const XLSX = require('xlsx');
const fs = require('fs');

// Constants
const monthlyInvestment = 150; // In stablecoin (e.g., USDC)
const annualInterestRate = 0.02;
const monthlyInterestRate = annualInterestRate / 12;
const liquidationThreshold = 0.825;
const months = 12;
const gasFeeLend = 0.02;
const gasFeeBorrow = 0.03;
const gasFeeProvideLiquidity = 0.05;
const totalGasFeePerMonth = gasFeeLend + gasFeeBorrow + gasFeeProvideLiquidity;

// ETH Price Scenarios
const ethPriceScenarios = {
  "Scenario 1 - Gradual Increase": {
    1: 2200, 2: 2200, 3: 2200,
    4: 2800, 5: 2800, 6: 2800,
    7: 3300, 8: 3300, 9: 3300,
    10: 4000, 11: 4000, 12: 4000
  },
  "Scenario 2 - Gradual Decrease": {
    1: 4000, 2: 4000, 3: 4000,
    4: 3300, 5: 3300, 6: 3300,
    7: 2800, 8: 2800, 9: 2800,
    10: 2200, 11: 2200, 12: 2200
  }
};

// Column Descriptions for README sheet
const columnDescriptions = [
  { Column: "Month", Description: "Represents the month number in the simulation (1-12)." },
  { Column: "ETH Price at Borrow ($/ETH)", Description: "The ETH price at the end of the month used for borrowing." },
  { Column: "Stablecoin Deposit ($)", Description: "The amount of money invested monthly in stablecoin (e.g., USDC)." },
  { Column: "Total Stablecoin Lent ($)", Description: "Total accumulated stablecoins with monthly compounding interest." },
  { Column: "Borrowed ETH (This Month)", Description: "The amount of ETH borrowed this month." },
  { Column: "Total ETH Borrowed (Cumulative)", Description: "Total ETH borrowed, accumulated with interest." },
  { Column: "LP Growth ($)", Description: "Cumulative value of ETH borrowed and deposited to LP." },
  { Column: "Impermanent Loss (%)", Description: "Impermanent loss percentage due to ETH price divergence from original price." },
  { Column: "LP Value After IL ($)", Description: "Liquidity Pool value after applying impermanent loss." },
  { Column: "ETH Debt Value ($)", Description: "Total ETH debt converted to USD at current ETH price." },
  { Column: "Gas Fee (Base Chain) ($)", Description: "Fixed gas fees paid monthly (lend, borrow, LP) on the Base chain." },
  { Column: "Net Capital ($)", Description: "Final capital after adding lent stablecoin, LP, deducting gas and ETH debt." },
  { Column: "LTV (%)", Description: "Loan-to-Value ratio based on current capital." },
  { Column: "ETH Liquidation Price ($)", Description: "Price at which ETH would trigger liquidation based on stablecoin collateral." },
  { Column: "Liquidated?", Description: "Whether your position would be liquidated this month ('Yes' or 'No')." },
  { Column: "Profit ($)", Description: "Net capital - Total Deposits, representing final profit/loss." }
];

// Utility
const format = (val, digits = 2) => typeof val === "number" ? val.toFixed(digits) : val;

function calculateIL(p0, p1) {
  const r = p1 / p0;
  return Math.max(0, 1 - (2 * Math.sqrt(r)) / (1 + r));
}

// 👇 Short Strategy Sheet Generation
function createShortScenarioSheet(ethPricesAtEnd) {
  const data = [];
  let lpGrowth = 0;
  let totalETHBorrowed = 0;
  let totalStablecoinLent = 0;
  let totalDeposited = 0;

  for (let month = 1; month <= months; month++) {
    const ethPrice = ethPricesAtEnd[month];

    // 1. Add stablecoin to lending pool and accrue interest
    totalStablecoinLent += monthlyInvestment;
    totalStablecoinLent *= (1 + monthlyInterestRate);
    totalDeposited += monthlyInvestment;

    // 2. Calculate how much ETH can be borrowed this month (target 45% LTV)
    const borrowableUSD = totalStablecoinLent * 0.45;
    const currentDebtUSD = totalETHBorrowed * ethPrice;
    const additionalBorrowUSD = borrowableUSD - currentDebtUSD;
    const borrowedETH = additionalBorrowUSD > 0 ? additionalBorrowUSD / ethPrice : 0;

    totalETHBorrowed = (totalETHBorrowed + borrowedETH) * (1 + monthlyInterestRate);

    // 3. Use borrowed ETH in LP
    lpGrowth += borrowedETH * ethPrice;

    const il = calculateIL(2200, ethPrice); // Base price = 2200
    const lpFinal = lpGrowth * (1 - il);

    // 4. Calculate debt and net capital
    const ethDebtValue = totalETHBorrowed * ethPrice;
    const netCapital = totalStablecoinLent + lpFinal - ethDebtValue - totalGasFeePerMonth;

    const totalAssets = totalStablecoinLent + lpFinal;
    const ltv = ethDebtValue / totalAssets;
    const isLiquidated = ltv > (1 / liquidationThreshold) ? "Yes" : "No";

    const liquidationPrice = totalETHBorrowed > 0
      ? (totalStablecoinLent * liquidationThreshold) / totalETHBorrowed
      : "N/A";

    const profit = netCapital - totalDeposited;

    data.push({
      "Month": month,
      "ETH Price at Borrow ($/ETH)": ethPrice,
      "Stablecoin Deposit ($)": format(monthlyInvestment),
      "Total Stablecoin Lent ($)": format(totalStablecoinLent),
      "Borrowed ETH (This Month)": format(borrowedETH, 6),
      "Total ETH Borrowed (Cumulative)": format(totalETHBorrowed, 6),
      "LP Growth ($)": format(lpGrowth),
      "Impermanent Loss (%)": format(il * 100),
      "LP Value After IL ($)": format(lpFinal),
      "ETH Debt Value ($)": format(ethDebtValue),
      "Gas Fee (Base Chain) ($)": format(totalGasFeePerMonth),
      "Net Capital ($)": format(netCapital),
      "LTV (%)": format(ltv * 100),
      "ETH Liquidation Price ($)": typeof liquidationPrice === "number" ? format(liquidationPrice) : "N/A",
      "Liquidated?": isLiquidated,
      "Profit ($)": format(profit)
    });
  }

  return data;
}

// Create workbook
const workbook = XLSX.utils.book_new();

// Create README
const readmeSheet = XLSX.utils.json_to_sheet(columnDescriptions);
XLSX.utils.book_append_sheet(workbook, readmeSheet, 'README');

// Add each short scenario as a new sheet
for (const [scenarioName, priceData] of Object.entries(ethPriceScenarios)) {
  const shortSheetData = createShortScenarioSheet(priceData);
  const shortSheet = XLSX.utils.json_to_sheet(shortSheetData);
  XLSX.utils.book_append_sheet(workbook, shortSheet, `${scenarioName} (Short)`);
}

// Save the workbook
const filename = 'ETH_Short_LP_Strategy.xlsx';
XLSX.writeFile(workbook, filename);
console.log(`✅ File generated: ${filename}`);
