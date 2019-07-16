const puppeteer = require('puppeteer');
(async () => {
  const browser = await puppeteer.launch();
  const page = await browser.newPage();
  await page.goto('https://frs.gov.cz/cs/ioff/application-status');
  await page.$("#edit-ioff-application-number").then(res=> {return res});
  await page.type('#edit-ioff-application-number', '<request number>', { delay: 50 })
  await page.select('#edit-ioff-application-code', '<type>')
  await page.select('#edit-ioff-application-year', '<year>')
  await page.click('#edit-submit-button', { delay: 100 });
  await page.waitForSelector('body > div.main-container.container > div > section > div.alert.alert-block.alert-success.messages.status > ul > li:nth-child(1) > p');
  await page.screenshot({path: 'result.png'});
  await browser.close();
})();
