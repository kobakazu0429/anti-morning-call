import * as playwright from "playwright";
import { ElementHandleForTag } from "playwright/types/structs";

const DEBUG = false;

export class OutlookLite {
  private email: string;
  private password: string;
  public page!: playwright.Page;
  public browser!: playwright.Browser;

  constructor(email: string, password: string) {
    this.email = email;
    this.password = password;
  }

  public async init() {
    this.browser = await playwright["chromium"].launch({ headless: !DEBUG });
    const context = await this.browser.newContext({
      recordVideo: {
        dir: ".",
      },
    });
    this.page = await context.newPage();

    console.log("Outlook.init: initialized");
  }

  public async login() {
    const url = `https://outlook.office365.com/owa/?exsvurl=1&layout=light&wa=wsignin1.0.`;
    await this.page.goto(url);

    await this.page.fill('input[type="email"]', this.email);
    await this.page.click('input[type="submit"]');
    await this.page.fill('input[type="password"]', this.password);
    await this.page.click('input[type="submit"]');

    await this.page.waitForSelector(".inner");

    if (await this.page.isVisible(`text=サインインの状態を維持しますか?`)) {
      await this.click(`text=はい`);
    }
    console.log("Outlook.login: logined");
  }

  public async gotoOutlookPage() {
    const url = "https://outlook.office365.com/owa/";
    await this.page.goto(url);
    console.log("Outlook.gotoOutlookPage: moved");
  }

  public async searchMorningCall() {
    const err = await this.page.$("#errMsg");
    if (err) {
      throw new Error(await err.innerText());
    }

    await this.page.click("text=朝の点呼");
    const urlPrefix = `https://390390.jp/parent/enq/?ENV_CODE=webasp6&access_key=`;
    await this.page.click(`text=${urlPrefix}`);

    await this.page.waitForTimeout(2000);
    const pages = this.page.context().pages();
    this.page = pages[1];

    console.log("Outlook.searchMorningCall: done");
  }

  public async answerForm() {
    // 寮
    await this.page.check(`//*[@id="frmRegForm_EnqAnswer"]/div[1]/label/input`);
    // 元気です
    await this.page.check(`//*[@id="frmRegForm_EnqAnswer"]/div[5]/label/input`);
    console.log("Outlook.answerForm: done");
  }

  public async submitForm() {
    await this.page.click('input[type="submit"]');
    console.log("Outlook.submitForm: done");
  }

  public async close() {
    this.page.close();
    const pages = this.page.context().pages();
    this.page = pages[0];

    await this.page.click("text=サインアウト");
    await this.browser.close();
    console.log("Outlook.close: closed");
  }

  public async screenshotElement(element: string, filename: string) {
    await this.page.waitForSelector(element, { state: "attached" });
    const detectedElement = await this.page.$(element);
    if (detectedElement) {
      await detectedElement.screenshot({ path: `${filename}.png` });
      console.log(`Outlook.screenshotElement: screenshoted, ${element}`);
    } else {
      console.log(`Outlook.screenshotElement: screenshot failed, ${element}`);
    }
  }

  public async click(element: string) {
    const handle = await this.page.$(element);
    if (handle) {
      await handle.click();
      console.log(`Outlook.click: ${element}`);
    } else {
      console.log(`Outlook.click: failed, ${element}`);
    }
  }
}
