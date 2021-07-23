const puppeteer = require('puppeteer');
const assert = require('assert');
require('dotenv').config();
const email = process.env.EMAIL;
const password = process.env.PASSWORD;
(async () => {
  const browser = await puppeteer.launch({headless: false});
  const page = await browser.newPage();
  await page.goto('https://outlook.live.com/owa/');


  // Click on login button
  const firstLoginButtonSelector = 'body > header > div > aside > div > nav > ul > li:nth-child(2) > a'
  await page.waitForSelector(firstLoginButtonSelector);
  await page.click(firstLoginButtonSelector);

  // Input email address
  const emailAddressSelector = '#lightbox > div:nth-child(3) > div > div > div > div.row > div.form-group.col-md-24 > div > input.form-control.ltr_override.input.ext-input.text-box.ext-text-box'
  await page.waitForSelector(emailAddressSelector);
  await page.type(emailAddressSelector, email);

  // Click on Next button
  const nextButtonSelector = "#lightbox > div:nth-child(3) > div > div > div > div.win-button-pin-bottom > div > div > div > div > input"
  await page.waitForSelector(nextButtonSelector);
  await page.click(nextButtonSelector)

  // Input password
  const passwordSelector = "#lightbox > div:nth-child(3) > div > div.pagination-view.animate.has-identity-banner.slide-in-next > div > div:nth-child(11) > div > div.placeholderContainer > input"
  await page.waitForSelector(passwordSelector);
  await page.type(passwordSelector, password);

  // Click on login button
  const secondLoginSelector = "#lightbox > div:nth-child(3) > div > div.pagination-view.animate.has-identity-banner.slide-in-next > div > div.position-buttons > div.win-button-pin-bottom > div > div > div > div > input"
  await page.waitForSelector(secondLoginSelector);
  await page.click(secondLoginSelector);

  // Click on Write email button
  const writeEmailSelector = '[role="region"] > div > div:nth-child(2) > div > div  > button'
  await page.waitForSelector(writeEmailSelector);
  await page.click(writeEmailSelector);

  // Type in email address of Recipient
  const recipientSelector = '[role="complementary"]:nth-child(2) > div > div > div > div > div > div > div > div > div > div > div > div > div > div > div > input';
  await page.waitForSelector(recipientSelector);
  await page.type(recipientSelector, 'echorus78@gmail.com');
  await page.keyboard.press("Enter");

  // Type in subject
  var possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  var subject = ""
  for (var i=0; i < 5;i++){
      subject += possible.charAt(Math.floor(Math.random() * possible.length));
  }
  await page.waitForSelector(".ms-TextField-field");
  await page.type(".ms-TextField-field",subject);
  
  // Press ctrl+enter to send email.
  await page.keyboard.down("Control");
  await page.keyboard.press("Enter");
  await page.keyboard.up("Control");

  // Click on Sent messages tab.
  const sentMessagesSelector = '[role="tree"]:nth-child(3) > div:nth-child(5)'
  await page.waitForSelector(sentMessagesSelector);
  // Delay of 5000ms should solve the problem of clicking on sent mails before outlook actually sends the mail.
  await page.click(sentMessagesSelector ,{delay: 5000});

  // Check if there are any email in the list and if so then check if its the last sent one and then delete all.
  try{
    await page.waitForSelector('[data-convid]',{timeout:5000});
    const sentEmailsSelector = '[data-convid] > div > div > div > div > div:nth-child(2) > div > div'
    await page.waitForSelector(sentEmailsSelector);
    const potentialSubject = await page.$eval((sentEmailsSelector), div => div.textContent);
    
    assert.strictEqual(potentialSubject,subject);

    // Click on the dump folder button and confirm it
    await page.click('[role="menuitem"]', {delay: 1000});
    await page.waitForSelector('#ok-1');
    await page.click('#ok-1', {delay: 1000});
  } catch(err){
    console.log(err)
  }
})();
