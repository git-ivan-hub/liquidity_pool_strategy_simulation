const XLSX = require('xlsx');
const fs = require('fs');

// === Parameters ===
const monthlyInvestment = 150;
const annualInterestRate = 0.02;
const monthlyInterestRate = annualInterestRate / 12;
const liquidationThreshold = 0.825;
const months = 12;

// Gas Fees per transaction on Base Chain
const gasFeeLend = 0.02;
const gasFeeBorrow = 0.03;
const gasFeeProvideLiquidity = 0.05;
const totalGasFeePerMonth = gasFeeLend + gasFeeBorrow + gasFeeProvideLiquidity;

// === Scenarios ===
const ethPriceScenarios = {
  "Scenario 1 (ETH Rises)": {
    1: 2200, 2: 2200, 3: 2200,
    4: 2800, 5: 2800, 6: 2800,
    7: 3300, 8: 3300, 9: 3300,
    10: 4000, 11: 4000, 12: 4000
  },
  "Scenario 2 (ETH Falls)": {
    1: 4000, 2: 4000, 3: 4000,
    4: 3300, 5: 3300, 6: 3300,
    7: 2800, 8: 2800, 9: 2800,
    10: 2200, 11: 2200, 12: 2200
  }
};

// === Impermanent Loss Function ===
function calculateIL(p0, p1) {
  const r = p1 / p0;
  return Math.max(0, 1 - (2 * Math.sqrt(r)) / (1 + r));
}

// === Create Sheet per Scenario ===
const createScenarioSheet = (ethPricesAtEnd) => {
  const data = [];

  let lpGrowth = 0;
  let totalBorrowedAmount = 0;
  let totalETHHeld = 0;
  let totalETHHeldWithoutInterest = 0;
  let totalCapital = 0;

  for (let month = 1; month <= months; month++) {
    const p0 = ethPricesAtEnd[month];

    const purchasedEth = monthlyInvestment / p0;
    totalETHHeldWithoutInterest += purchasedEth;
    totalETHHeld += purchasedEth;
    totalETHHeld *= (1 + monthlyInterestRate);

    const totalETHValueInUSD = totalETHHeld * p0;
    totalCapital = totalETHValueInUSD + lpGrowth;

    const targetBorrowAmount = totalCapital * 0.45;
    const borrowedAmountInUSD = targetBorrowAmount - totalBorrowedAmount;
    totalBorrowedAmount = (totalBorrowedAmount + borrowedAmountInUSD) * (1 + monthlyInterestRate);

    lpGrowth += borrowedAmountInUSD;

    const il = calculateIL(p0, ethPricesAtEnd[1]);
    const lpFinal = lpGrowth * (1 - il);
    const netCapitalInUSD = totalETHValueInUSD + lpFinal - totalBorrowedAmount - totalGasFeePerMonth;

    let liquidationPrice = totalBorrowedAmount * liquidationThreshold / totalETHHeld;
    liquidationPrice = liquidationPrice > 0 ? liquidationPrice : "N/A";

    const isLiquidated = totalBorrowedAmount > totalCapital * liquidationThreshold ? "Yes" : "No";
    const totalDepositedAmount = monthlyInvestment * month;

    const profit = lpFinal + (netCapitalInUSD - totalDepositedAmount);
    const borrowRatePercentage = (borrowedAmountInUSD / totalCapital) * 100;
    const LTVPercentage = (totalBorrowedAmount / totalCapital) * 100;

    data.push({
      Month: month,
      "ETH Price at Deposit ($/ETH)": p0,
      "Deposit Amount ($)": monthlyInvestment.toFixed(2),
      "Total ETH Held Without Interest (ETH)": totalETHHeldWithoutInterest.toFixed(6),
      "Total ETH Held With Interest (ETH)": totalETHHeld.toFixed(6),
      "Total ETH Value (USD)": totalETHValueInUSD.toFixed(2),
      "Borrowed This Month ($)": borrowedAmountInUSD.toFixed(2),
      "Total Borrowed with Interest ($)": totalBorrowedAmount.toFixed(2),
      "Borrow Rate (%)": borrowRatePercentage.toFixed(2),
      "LTV (%)": LTVPercentage.toFixed(2),
      "LP Growth ($)": lpGrowth.toFixed(2),
      "Impermanent Loss (%)": (il * 100).toFixed(2),
      "LP Value After IL ($)": lpFinal.toFixed(2),
      "Gas Fee (Base Chain) ($)": totalGasFeePerMonth.toFixed(2),
      "Net Capital ($)": netCapitalInUSD.toFixed(2),
      "ETH Price for Liquidation ($)": liquidationPrice === "N/A" ? "N/A" : liquidationPrice.toFixed(2),
      "Liquidated?": isLiquidated,
      "Profit ($)": profit.toFixed(2)
    });
  }

  return data;
};

// === Create Workbook and Append Sheets ===
const workbook = XLSX.utils.book_new();

// 1. README Sheet
const readme = [
  ["Column Name", "Description"],
  ["Month", "Investment month (1–12)"],
  ["ETH Price at Deposit ($/ETH)", "Price of ETH at end of the month"],
  ["Deposit Amount ($)", "Monthly investment in USD"],
  ["Total ETH Held Without Interest (ETH)", "Cumulative ETH acquired (no interest)"],
  ["Total ETH Held With Interest (ETH)", "Total ETH held with monthly compound interest"],
  ["Total ETH Value (USD)", "Value of ETH held in USD at month-end price"],
  ["Borrowed This Month ($)", "Amount borrowed to maintain 45% LTV"],
  ["Total Borrowed with Interest ($)", "Cumulative borrowed amount including interest"],
  ["Borrow Rate (%)", "Monthly borrow rate as a % of total capital"],
  ["LTV (%)", "Loan-to-Value ratio"],
  ["LP Growth ($)", "Growth of LP from borrowed funds"],
  ["Impermanent Loss (%)", "Impermanent loss vs initial ETH price"],
  ["LP Value After IL ($)", "LP value after applying IL"],
  ["Gas Fee (Base Chain) ($)", "Total gas fees per month"],
  ["Net Capital ($)", "ETH + LP - borrow - gas"],
  ["ETH Price for Liquidation ($)", "ETH price that triggers liquidation"],
  ["Liquidated?", "Was portfolio liquidated this month?"],
  ["Profit ($)", "Net Capital + LP - Total Deposits"]
];
const readmeSheet = XLSX.utils.aoa_to_sheet(readme);
XLSX.utils.book_append_sheet(workbook, readmeSheet, "README");

// 2. Scenario Sheets
for (const [scenarioName, priceData] of Object.entries(ethPriceScenarios)) {
  const data = createScenarioSheet(priceData);
  const sheet = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(workbook, sheet, scenarioName);
}

// === Save File ===
const filename = 'ETH_LP_Strategy_Simulation.xlsx';
XLSX.writeFile(workbook, filename);
console.log(`✅ File generated: ${filename}`);
