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
    const USA = page.locator('#te-c-all li').filter({ hasText: 'United States' });
    await USA.scrollIntoViewIfNeeded();
    await USA.click();
    await page.getByRole('cell', { name: '  Recent    Impact  ' }).click();
    await page.getByRole('button', { name: '   Category' }).click();
    await page.getByText('Save').click();
    const content = await page.content();
    await page.waitForTimeout(5000);
    let scrollsRemaining = 20;
    while (scrollsRemaining > 0) {
        await page.evaluate(() => window.scrollBy(0, 10000));
        await page.waitForTimeout(1000);
        scrollsRemaining--;
    }
    writeFile('./docs/index.html', content, (err) => {
        if (err) {
            console.error('Error writing to file', err);
        } else {
            console.log('File has been written');
        }
    });
    await browser.close();
});
