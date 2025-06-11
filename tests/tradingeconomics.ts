import { test } from '@playwright/test';
import { writeFile } from 'fs';
import { chromium } from 'playwright-extra';

test('test', async () => {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({
    userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
  });
  const page = await context.newPage();
  await page.goto('https://tradingeconomics.com/calendar');
  await page.getByRole('button', { name: 'Countries' }).click();
  await page.getByText('Clear').click();
  await page.waitForTimeout(1000);
  const USA = page.locator('#te-c-all li').filter({ hasText: 'United States' });
  await USA.scrollIntoViewIfNeeded();
  await USA.click();
  await page.getByText('Save').click();
  let scrollsRemaining = 50;
  let isNotYetBottom = true;
  const cardHeader = page.locator('div.card-header');
  while (scrollsRemaining > 0) {
    await page.evaluate(() => window.scrollBy(0, 1000));
    await page.waitForTimeout(500);
    isNotYetBottom = await cardHeader.isHidden();
    if (!isNotYetBottom) {
      break;
    }
    scrollsRemaining--;
  }
  const content = await page.content();
  writeFile('./docs/index.html', content, (err) => {
    if (err) {
      console.error('Error writing to file', err);
    } else {
      console.log('File has been written');
    }
  });
  await browser.close();
});
