const puppeteer = require('puppeteer');
(async () => {
  try {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    page.on('console', msg => console.log('PAGE LOG:', msg.text()));
    page.on('pageerror', error => console.error('PAGE ERROR:', error.message));
    page.on('response', response => {
      if (!response.ok()) console.log('PAGE RES ERR:', response.url(), response.status());
    });
    console.log('Navigating to index.html...');
    await page.goto('http://localhost:8081/workspace/index.html', {waitUntil: 'networkidle0'});
    console.log('Done.');
    await browser.close();
  } catch(e) {
    console.error('Pupt error', e);
  }
})();
