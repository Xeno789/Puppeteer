const puppeteer = require('puppeteer');
const assert = require('assert');
require('dotenv').config();
const email = process.env.EMAIL;
const password = process.env.PASSWORD;
const targetEmail = process.env.TARGET_EMAIL;
(async () => {
  const browser = await puppeteer.launch({
    headless: false
  });
  const page = await browser.newPage();
  await page.goto('https://outlook.live.com/owa/');


  // Click on login button
  const firstLoginButtonXpath = "/html/body/header/div/aside/div/nav/ul/li[2]/a"
  const firstLoginButton = await page.waitForXPath(firstLoginButtonXpath)
  await firstLoginButton.click();

  // Input email address
  const emailAddressInputXpath = '//*[@id="lightbox"]/div[3]/div/div/div/div[2]/div[2]/div/input[1]'
  const emailAddressInput = await page.waitForXPath(emailAddressInputXpath);
  await emailAddressInput.type(email)

  // Click on Next button
  const nextButtonXpath = '//*[@id="lightbox"]/div[3]/div/div/div/div[4]/div/div/div/div/input'
  const nextButton = await page.waitForXPath(nextButtonXpath);
  await nextButton.click();

  // Input password
  const passwordInputXpath = '//*[@id="lightbox"]/div[3]/div/div[2]/div/div[2]/div/div[2]/input'
  const passwordInput = await page.waitForXPath(passwordInputXpath);
  await passwordInput.type(password);

  // Click on login button
  const secondLoginButtonXpath = '//*[@id="lightbox"]/div[3]/div/div[2]/div/div[3]/div[2]/div/div/div/div/input'
  const secondLoginButton = await page.waitForXPath(secondLoginButtonXpath);
  await secondLoginButton.click();

  // Time by Time Outlook asks if you want to stay signed in, if thats the case, this clicks no.
  try {
    const doNotStaySignedInButtonXpath = '/html/body/div/form/div/div/div[1]/div[2]/div/div[2]/div/div[3]/div[2]/div/div/div[1]/input'
    const doNotStaySignedInButton = await page.waitForXPath(doNotStaySignedInButtonXpath);
    await doNotStaySignedInButton.click();
  } catch {
    console.log("Outlook didnt ask if I want to stay signed in.")
  }

  // Click on write email button
  const writeEmailButtonXpath = '//*[@id="app"]/div/div[2]/div[2]/div[1]/div/div/div[1]/div[1]/div[2]/div/div/button'
  const writeEmailButton = await page.waitForXPath(writeEmailButtonXpath);
  await writeEmailButton.click();

  // Type in email address of Recipient
  const recipientInputXpath = '//*[@id="ReadingPaneContainerId"]/div/div/div/div[1]/div[1]/div[1]/div/div[1]/div/div[1]/div/div/div[2]/div[2]/input';
  const recipientInput = await page.waitForXPath(recipientInputXpath);
  await recipientInput.type(targetEmail);
  await page.keyboard.press("Enter")

  // Type in subject
  var possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  var subject = ""
  for (var i = 0; i < 5; i++) {
    subject += possible.charAt(Math.floor(Math.random() * possible.length));
  }
  const subjectInputXpath = '//*[@id="ReadingPaneContainerId"]/div/div/div/div[1]/div[1]/div[3]/div[2]/div/div/div/input';
  const subjectInput = await page.waitForXPath(subjectInputXpath);
  await subjectInput.type(subject);

  // Press send email button
  const sendMailButtonXpath = '//*[@id="ReadingPaneContainerId"]/div/div/div/div[1]/div[3]/div[2]/div[1]/div/span/button[1]'
  const sendMailButton = await page.waitForXPath(sendMailButtonXpath);
  await sendMailButton.click();

  // Click on Sent messages tab.
  const sentMessagesButtonXpath = '//*[@id="app"]/div/div[2]/div[2]/div[1]/div/div/div[1]/div[2]/div/div[1]/div/div[2]/div[5]/div/span[1]'
  const sentMessagesButton = await page.waitForXPath(sentMessagesButtonXpath);
  // Delay of 5000ms should solve the problem of clicking on sent mails before outlook actually sends the mail.
  await sentMessagesButton.click({
    delay: 5000
  });

  // Check if there are any email in the list and if so then check if its the last sent one and then delete all.

  const lastSentEmailSubjectXpath = '//*[@id="app"]/div/div[2]/div[2]/div[1]/div/div/div[3]/div[2]/div/div[1]/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div[2]/div/div[1]/span'
  const lastSentEmailSubject = await page.waitForXPath(lastSentEmailSubjectXpath, {
    timeout: 5000
  });
  const potentialSubject = await page.evaluate(async subject => await subject.innerText, lastSentEmailSubject);

  assert.strictEqual(potentialSubject, subject);

  // Click on the dump folder button and confirm it
  const dumpFolderButtonXpath = '//*[@id="app"]/div/div[2]/div[2]/div[1]/div/div/div[3]/div[1]/div/div/div/div/div/div[1]/div[1]/button';
  const dumpFolderButton = await page.waitForXPath(dumpFolderButtonXpath)
  await dumpFolderButton.click({
    delay: 1000
  });

  const confirmButtonXpath = '/html/body/div[12]/div/div/div/div[2]/div[2]/div/div[2]/div[2]/div/span[1]/button';
  const confirmButton = await page.waitForXPath(confirmButtonXpath)
  await confirmButton.click({
    delay: 1000
  });
})();