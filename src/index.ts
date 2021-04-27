import * as dotenv from "dotenv";
import moment from "moment-timezone";
// import { AutoBodyTempMesument } from "./outlook";
import { Outlook } from "./outlook";
import { uploadGyazo } from "./gyazo";
import { sendToSlack } from "./slack";

dotenv.config();

const { EMAIL, PASSWORD } = process.env;

if (!EMAIL || !PASSWORD)
  throw new Error(`EMAIL: ${EMAIL}, PASSWORD: ${PASSWORD}`);

const today = moment().tz("Asia/Tokyo");
const month = today.month() + 1;
const date = today.date();
const day = today.day();

console.log(today);
console.log(month, date, day);

if ([0, 6].includes(day)) {
  console.log("weekend.");
  process.exit();
}

const sleep = (ms: number) => {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
};

(async () => {
  const outlook = new Outlook(EMAIL, PASSWORD);
  await outlook.init();
  await outlook.login();
  await outlook.gotoOutlookPage();
  await outlook.searchMorningCall();

  // const filename = `${month}-${date}`;

  // await outlook.fillForm();
  // await outlook.screenshotElement("#region-main", `${filename}-form`);

  // await outlook.page.waitForTimeout(500);
  // await outlook.answerForm();
  // await outlook.page.waitForTimeout(2000);
  // await outlook.screenshotElement("#region-main", `${filename}-answer`);
  // await outlook.close();

  // const formImgUrl = await uploadGyazo(`${filename}-form.png`);
  // const answerImgUrl = await uploadGyazo(`${filename}-answer.png`);

  // await sendToSlack({ url: formImgUrl, title: `${filename}-form` });
  // await sendToSlack({ url: answerImgUrl, title: `${filename}-answer` });
  await sleep(10000);
})();
