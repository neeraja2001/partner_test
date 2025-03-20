const xlsx = require('xlsx');
const path = require('path');
const { test, expect, chromium } = require('@playwright/test');
const ExcelJS = require('exceljs'); // Install: npm install exceljs
const AsyncLock = require('async-lock');

const lock = new AsyncLock();
// File paths
const inputFilePath = path.join("C:\\Users\\018073\\Desktop\\Playwright\\partnerlistaccounts.xls");
const outputFilePath = path.join("C:\\Users\\018073\\Desktop\\Playwright\\partnerlistaccounts_updated.xlsx");

let workbook;
let worksheet;
let isWorksheetCleared = false;

// Function to set up the workbook and worksheet
async function setupWorkbook() {
  try {
    // Try to read the existing workbook
    workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(outputFilePath);
    worksheet = workbook.getWorksheet('AccountCounts');
    if (!worksheet) {
      worksheet = workbook.addWorksheet('AccountCounts');
      worksheet.addRow(['Username', 'Count', 'Status']); // Add header row if missing
    } else if (!isWorksheetCleared) {
      // Clear the worksheet only once
      worksheet.removeRows(1, worksheet.rowCount); // Clear all rows at once
      isWorksheetCleared = true; // Set the flag to true after clearing
    }
    await workbook.xlsx.writeFile(outputFilePath);
  } catch (error) {
    // If file doesn't exist, create a new workbook and sheet
    workbook = new ExcelJS.Workbook();
    worksheet = workbook.addWorksheet('AccountCounts');
    worksheet.addRow(['Username', 'Count', 'Status']); // Add header row
  }
}

// Function to reload the page until an element is visible
async function reloadUntilVisiblePolling(page, selector) {
  while (true) {
    try {
      await page.reload({ waitUntil: 'networkidle' });
      const isVisible = await page.locator(selector).isVisible();
      if (isVisible) {
        console.log(`Element "${selector}" is now visible after reload.`);
        break;
      }
      console.log(`Element "${selector}" not visible, reloading...`);
      await page.waitForTimeout(1000); // Optional delay
    } catch (error) {
      console.error("Error during reload or visibility check:", error);
      break; // Or handle the error as needed
    }
  }
}

// Run workbook setup before all tests
test.beforeAll(async () => {
  await setupWorkbook();
});

// Read data from the input file
const workbook_read = xlsx.readFile(inputFilePath);
const sheet = workbook_read.Sheets[workbook_read.SheetNames[0]];
const data = xlsx.utils.sheet_to_json(sheet);

test.describe('Partner circuit list', () => {
  // Loop through the accounts (process the first 5 rows for testing)
  data.slice(0, 5).forEach((account, index) => {
    const { Username: username, Password: password } = account;

    if (!username || !password) {
      console.error(`Missing username or password at row ${index + 1}`);
      return;
    }

    test(`Circuit list count for ${username} - Row ${index + 1}`, async () => {
      const browser = await chromium.launch({ headless: false });
      const context = await browser.newContext({
        ignoreHTTPSErrors: true,
        userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
      });
      test.setTimeout(400000);
      const page = await context.newPage();

      let countValue = 0; // Initialize count
      let status = "Pending"; // Initialize status

      try {
        // Login and navigation code
        await page.reload({ waitUntil: 'networkidle' });
        await page.goto("https://sifyaakaash.net/#/login");
        await reloadUntilVisiblePolling(page, "//button[@class='border-btn dropdown-toggle']");

        const delayBeforeLogin = Math.floor(Math.random() * (10000 - 2000 + 1)) + 2000;
        console.log(`Waiting for ${delayBeforeLogin} ms before login...`);
        await page.waitForTimeout(delayBeforeLogin);

        await page.click("//button[@class='border-btn dropdown-toggle']");
        await page.fill('id=username', username);
        await page.fill('id=password', password);
        console.log(`Login attempt for ${username}`);

        await page.click('id=kc-login');
        while (true) {
          await page.reload({ waitUntil: 'networkidle' });
          try {
            await page.waitForURL('https://sifyaakaash.net/#/partner-circuit-list', { timeout: 20000 });
            console.log('Successfully navigated to Partner Circuit List page');
            break;
          } catch (error) {
            console.log('Still not on the right page, reloading...');
          }
        }

        // Count retrieval
        const countLocator = page.locator('//my-app/partner/section/section/div[2]/data-grid/div/div[2]/div/div/div[3]/a');
        await page.waitForTimeout(2000);
        const text = await countLocator.textContent();
        console.log(`Total count for ${username}:`, text);
        countValue = parseInt(text, 10);
        expect(countValue).toBeGreaterThan(0);
        status = "Success";

        await page.click("//img[@src='../../../assets/images/logout-vector.png']");
        await context.clearCookies();
        await context.clearPermissions();
        await context.storageState({ path: null });
        await page.reload({ waitUntil: 'networkidle' });
        await page.waitForURL("https://sifyaakaash.net/#/login");
      } catch (error) {
        console.error(`Test failed for ${username} at row ${index + 1}:`, error.message);
        status = "Failed"; // Update status on failure
      } finally {
        await browser.close();
        await lock.acquire("excel", async () => {
          worksheet.addRow([username, countValue, status]);
          await workbook.xlsx.writeFile(outputFilePath);
        }); // Save after each test
      }
    });
  });
});

test.afterAll(async () => {
  console.log(`Excel file updated: ${outputFilePath}`);
});