import puppeteer from 'puppeteer';

(async () => {
  try {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    
    page.on('console', msg => console.log('PAGE LOG:', msg.type(), msg.text()));
    page.on('pageerror', error => console.error('PAGE ERROR:', error.message));
    page.on('requestfailed', request => {
      console.log('REQ FAIL:', request.url(), request.failure().errorText);
    });

    console.log('Navigating to index.html...');
    await page.goto('http://localhost:8081/workspace/index.html', {waitUntil: 'networkidle0'});
    
    console.log('Clicking 3D viewer button...');
    await page.click('[data-open-3d="true"]');
    
    await new Promise(r => setTimeout(r, 2000));
    
    console.log('Done.');
    await browser.close();
  } catch(e) {
    console.error('Puppeteer error:', e);
  }
})();
