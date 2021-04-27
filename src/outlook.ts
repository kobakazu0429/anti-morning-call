import * as playwright from "playwright";

export class Outlook {
  private email: string;
  private password: string;
  public page!: playwright.Page;
  public browser!: playwright.Browser;

  constructor(email: string, password: string) {
    this.email = email;
    this.password = password;
  }

  public async init() {
    this.browser = await playwright["chromium"].launch({ headless: false });
    const context = await this.browser.newContext();
    this.page = await context.newPage();

    console.log("Outlook.init: initialized");
  }

  public async login() {
    const url = "https://login.microsoftonline.com/";
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
    const url = "https://outlook.office.com/mail";
    await this.page.goto(url);
    console.log("Outlook.gotoOutlookPage: moved");
  }

  public async searchMorningCall() {
    await this.page.waitForSelector("id=owaSearchBox", { state: "attached" });
    await this.page.fill("input", "朝の点呼");
    await this.page.click('button[aria-label="検索"]');
    console.log("Outlook.searchMorningCall: done");
    await this.page.waitForSelector(`text=上位の結果`, { state: "attached" });

    const mails = await this.page.$eval(
      'div[aria-label="メッセージ一覧"]',
      (e) => {
        Array.from(e.children[0].children[0].children).forEach((s) =>
          console.log(s.getAttribute("aria-label"))
        );
      }
    );
    console.log(mails);
  }

  public async close() {
    await this.browser.close();
    console.log("Outlook.close: closed");
  }

  public async wait() {
    await this.page.waitForSelector("#app", { state: "attached" });
  }

  public async screenshot(filename: string) {
    await this.wait();
    await this.page.screenshot({ path: `${filename}.png`, fullPage: true });
    console.log("Outlook.screenshot: screenshoted");
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

// export class AutoBodyTempMesument extends Outlook {
//   constructor(email: string, password: string) {
//     super(email, password);
//   }

//   public async gotoMesumentPage() {
//     const url = "https://k-mdl01.kure-nct.ac.jp/course/view.php?id=37";
//     await this.page.goto(url);
//     console.log("AutoBodyTempMesument.gotoMesumentPage: moved");
//   }

//   public async gotoDailyPage(month: number, date: number) {
//     await this.wait();
//     await this.click(`text=${month}/${date}`);
//   }

//   public async gotoAnswerFormPage() {
//     await this.wait();
//     await this.click(`text=質問に回答する`);
//     console.log("AutoBodyTempMesument.gotoAnswerFormPage: moved");
//   }

//   public async fillForm() {
//     await this.wait();
//     const selects = await this.page.$$("select");
//     await selects[0].selectOption("1");
//     await selects[1].selectOption("1");
//     console.log("AutoBodyTempMesument.fillForm: done");
//   }

//   public async answerForm() {
//     await this.click(`input[name=savevalues]`);
//     console.log("AutoBodyTempMesument.answerForm: done");
//   }
// }
