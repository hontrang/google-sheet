import { test } from '@playwright/test';
import { writeFile } from 'fs';
import { chromium } from 'playwright-extra'

test('test', async () => {
    const browser = await chromium.launch({ headless: true })
    const context = await browser.newContext({
        userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
    })
    const page = await context.newPage()
    await page.goto('https://tradingeconomics.com/calendar');
    await page.getByRole('button', { name: '  Countries' }).click();
    await page.getByText('Clear').click();
    await page.locator('#te-c-all li').filter({ hasText: 'United States' }).click();
    await page.getByRole('cell', { name: '  Recent    Impact  ' }).click();
    await page.getByRole('button', { name: '   Category' }).click();
    await page.getByText('Save').click();
    const content = await page.content();
    writeFile('./docs/index.html', content, (err) => {
        if (err) {
            console.error('Error writing to file', err);
        } else {
            console.log('File has been written');
        }
    });
});
