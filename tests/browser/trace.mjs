import { readFile } from 'node:fs/promises';
import { once } from 'node:events';
import assert from 'node:assert/strict';
import { chromium, webkit } from 'playwright';
import issuer from '../../portable/certificate.cjs';
import { createPaneServer } from '../../scripts/pane-server.mjs';

const server = createPaneServer(issuer.createIdentity());
server.listen(0,'localhost');
await once(server,'listening');
let browser;
try {
  browser = process.env.CHROMIUM_EXECUTABLE
    ? await chromium.launch({executablePath:process.env.CHROMIUM_EXECUTABLE})
    : await webkit.launch();
  const page = await browser.newPage({ignoreHTTPSErrors:true,viewport:{width:350,height:625}});
  const errors = [];
  page.on('pageerror', e=>errors.push(e.message));
  const host = await readFile(new URL('./office-double.js', import.meta.url),'utf8');
  await page.route('https://appsforoffice.microsoft.com/**', route=>route.fulfill({contentType:'text/javascript',body:host}));
  await page.goto('https://localhost:'+server.address().port);
  await page.waitForFunction(()=>window.commands?.TraceInputs && document.querySelectorAll('#rows tr').length===3);
  async function step(key,address) {
    await page.keyboard.press(key);
    await page.waitForFunction(target=>window.active===target && document.activeElement?.matches('#rows tr[aria-current=true]') &&
      document.querySelector('#step').textContent.includes(target),address);
  }
  await page.evaluate(()=>commands.TraceInputs());
  await step('ArrowDown','Cash Flow!B4');
  await step('ArrowDown','Debt!B4');
  await step('ArrowUp','Cash Flow!B4');
  const leavesCommandQToHost = await page.evaluate(() =>
    document.activeElement.dispatchEvent(new KeyboardEvent('keydown', {
      key: 'q', metaKey: true, bubbles: true, cancelable: true
    }))
  );
  assert.equal(leavesCommandQToHost, true);
  await page.keyboard.press('ArrowRight');
  await page.waitForFunction(()=>document.querySelector('#formula').textContent==='=B2-B3');
  await step('ArrowDown','Cash Flow!B2');
  await step('ArrowLeft','Cash Flow!B4');
  await step('Home','Summary!B2');
  await page.evaluate(()=>commands.TraceUsers());
  await step('ArrowDown','Cash Flow!B7');
  await step('ArrowDown','Debt!B7');
  // The host simulates a successful selection and a subsequent loss of pane focus.
  // Each step above requires the application to regain focus and display the right row.
  assert.equal(await page.locator('#target-formula').textContent(),'=IF(Summary!B2>=1,"Pass","Fail")');
  await page.keyboard.press('Enter');
  await page.waitForFunction(()=>!window.visible && window.active==='Debt!B7');
  await page.evaluate(()=>commands.ReturnCell());
  assert.equal(await page.evaluate(()=>window.active),'Summary!B2');
  assert.deepEqual(errors,[]);
  console.log('Browser trace passed with simulated Office selection and real pane keyboard events.');
} finally {
  await browser?.close();
  server.close();
}
