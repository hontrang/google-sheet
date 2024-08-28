import { test } from '@playwright/test';
import { writeFile } from 'fs';

test('test', async ({ page }) => {
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