const XLSX = require('xlsx');
const fs = require('fs');

// Constants
const asset = "Etherium";
const monthlyInvestment = 150;
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

// Column Descriptions
const columnDescriptions = [
  { Column: "Scenario", Description: "ETH price scenario name", Formula: "One of the defined scenarios" },
  { Column: "Month", Description: "Represents the month number in the simulation (1-12).", Formula: "Incremental: 1 through 12" },
  { Column: "ETH Price at Deposit ($/ETH)", Description: "The ETH price at the end of the month used for deposit.", Formula: "From ETH price scenario object" },
  { Column: "Deposit Amount ($)", Description: "The amount of money invested monthly in USD.", Formula: `Constant: $${monthlyInvestment}` },
  { Column: "Total ETH Held Without Interest (ETH)", Description: "Total accumulated ETH from deposits, without compounding interest.", Formula: "Σ (Deposit Amount / ETH Price)" },
  { Column: "Total ETH Held With Interest (ETH)", Description: "Total ETH including monthly compounding interest.", Formula: "(Previous ETH + new ETH) * (1 + monthlyInterestRate)" },
  { Column: "Total ETH Value (USD)", Description: "Value of total ETH held (with interest) at current ETH price.", Formula: "Total ETH With Interest * ETH Price" },
  { Column: "Borrowed Amount (This Month, $)", Description: "The new borrowed amount to maintain a 45% LTV in this month.", Formula: "(Total Capital * 0.45) - Total Borrowed" },
  { Column: "Total Borrowed with Interest ($)", Description: "Total borrowed amount accumulated with interest.", Formula: "(Previous Total Borrowed + This Month Borrowed) * (1 + monthlyInterestRate)" },
  { Column: "Borrow Rate (%)", Description: "Borrowed amount for this month as a percentage of capital.", Formula: "(This Month Borrowed / Total Capital) * 100" },
  { Column: "LTV (%)", Description: "Loan-to-Value ratio based on current capital.", Formula: "(Total Borrowed / Total Capital) * 100" },
  { Column: "LP Growth ($)", Description: "Cumulative growth of funds added to LP using borrowed capital.", Formula: "Σ (This Month Borrowed)" },
  { Column: "Impermanent Loss (%)", Description: "Impermanent loss percentage due to ETH price divergence from original price ($2200).", Formula: "1 - (2 * √(p1/p0)) / (1 + p1/p0)" },
  { Column: "LP Value After IL ($)", Description: "Liquidity Pool value after applying impermanent loss.", Formula: "LP Growth * (1 - IL)" },
  { Column: "Gas Fee (Base Chain) ($)", Description: "Fixed gas fees paid monthly (lend, borrow, LP) on the Base chain.", Formula: `${gasFeeLend} + ${gasFeeBorrow} + ${gasFeeProvideLiquidity}` },
  { Column: "Net Capital ($)", Description: "Final capital after adding ETH value, LP, deducting gas and debt.", Formula: "ETH Value + LP Value - Total Borrowed - Gas Fee" },
  { Column: "ETH Liquidation Price ($)", Description: "Price at which ETH would trigger liquidation based on current LTV.", Formula: "(Total Borrowed * 0.825) / Total ETH Held" },
  { Column: "Liquidated?", Description: "Whether your position would be liquidated this month ('Yes' or 'No').", Formula: "Yes if LTV > 82.5%, else No" },
  { Column: "Profit ($)", Description: "Net capital + LP - Total Deposits, representing final profit/loss.", Formula: "Net Capital - Total Deposits + LP Value" }
];

// Utility functions
const format = (val, digits = 2) => typeof val === "number" ? val.toFixed(digits) : val;

function calculateIL(p0, p1) {
  const r = p1 / p0;
  return Math.max(0, 1 - (2 * Math.sqrt(r)) / (1 + r));
}

// Main calculator
function createScenarioSheet(ethPricesAtEnd, scenarioName) {
  const data = [];
  let lpGrowth = 0;
  let totalBorrowedAmount = 0;
  let totalETHHeld = 0;
  let totalETHHeldWithoutInterest = 0;
  let totalDepositedAmount = 0;

  for (let month = 1; month <= months; month++) {
    const p0 = ethPricesAtEnd[month];
    const purchasedEth = monthlyInvestment / p0;
    totalETHHeldWithoutInterest += purchasedEth;
    totalETHHeld += purchasedEth;
    totalETHHeld *= (1 + monthlyInterestRate);

    const totalETHValueInUSD = totalETHHeld * p0;
    const totalCapital = totalETHValueInUSD + lpGrowth;

    const targetBorrowAmount = totalCapital * 0.45;
    const borrowedAmountInUSD = targetBorrowAmount - totalBorrowedAmount;
    totalBorrowedAmount = (totalBorrowedAmount + borrowedAmountInUSD) * (1 + monthlyInterestRate);

    lpGrowth += borrowedAmountInUSD;

    const il = calculateIL(p0, 2200);
    const lpFinal = lpGrowth * (1 - il);

    const netCapitalInUSD = totalETHValueInUSD + lpFinal - totalBorrowedAmount - totalGasFeePerMonth;

    let liquidationPrice = totalBorrowedAmount * liquidationThreshold / totalETHHeld;
    liquidationPrice = liquidationPrice > 0 ? format(liquidationPrice) : "N/A";

    const isLiquidated = totalBorrowedAmount > totalCapital * liquidationThreshold ? "Yes" : "No";

    totalDepositedAmount += monthlyInvestment;
    const profit = lpFinal + (netCapitalInUSD - totalDepositedAmount);
    const borrowRatePercentage = (borrowedAmountInUSD / totalCapital) * 100;
    const LTVPercentage = (totalBorrowedAmount / totalCapital) * 100;

    data.push({
      "Scenario": scenarioName,
      "Month": month,
      "ETH Price at Deposit ($/ETH)": p0,
      "Deposit Amount ($)": format(monthlyInvestment),
      "Total ETH Held Without Interest (ETH)": format(totalETHHeldWithoutInterest, 6),
      "Total ETH Held With Interest (ETH)": format(totalETHHeld, 6),
      "Total ETH Value (USD)": format(totalETHValueInUSD),
      "Borrowed Amount (This Month, $)": format(borrowedAmountInUSD),
      "Total Borrowed with Interest ($)": format(totalBorrowedAmount),
      "Borrow Rate (%)": format(borrowRatePercentage),
      "LTV (%)": format(LTVPercentage),
      "LP Growth ($)": format(lpGrowth),
      "Impermanent Loss (%)": format(il * 100),
      "LP Value After IL ($)": format(lpFinal),
      "Gas Fee (Base Chain) ($)": format(totalGasFeePerMonth),
      "Net Capital ($)": format(netCapitalInUSD),
      "ETH Liquidation Price ($)": liquidationPrice,
      "Liquidated?": isLiquidated,
      "Profit ($)": format(profit)
    });
  }

  return data;
}

// Write Excel Workbook
const workbook = XLSX.utils.book_new();

// Add README sheet
const readmeSheet = XLSX.utils.json_to_sheet(columnDescriptions);
XLSX.utils.book_append_sheet(workbook, readmeSheet, 'README');

// Generate unified sheet
let allCombinedData = [];
const includePerScenarioSheets = true;

for (const [scenarioName, priceData] of Object.entries(ethPriceScenarios)) {
  const scenarioData = createScenarioSheet(priceData, scenarioName);
  allCombinedData = allCombinedData.concat(scenarioData);

  if (includePerScenarioSheets) {
    const sheet = XLSX.utils.json_to_sheet(scenarioData);
    XLSX.utils.book_append_sheet(workbook, sheet, scenarioName);
  }
}

// Add combined sheet
const combinedSheet = XLSX.utils.json_to_sheet(allCombinedData);
XLSX.utils.book_append_sheet(workbook, combinedSheet, 'Combined Scenarios');

// Save file
const filename = `${asset} Long LP Strategy Simulation.xlsx`;
XLSX.writeFile(workbook, filename);
console.log(`✅ File generated: ${filename}`);
