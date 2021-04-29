import * as dotenv from "dotenv";
import moment from "moment-timezone";
import { OutlookLite } from "./outlook-lite";
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

// const sleep = (ms: number) => {
//   return new Promise((resolve) => {
//     setTimeout(resolve, ms);
//   });
// };

(async () => {
  const outlook = new OutlookLite(EMAIL, PASSWORD);
  await outlook.init();
  await outlook.login();
  await outlook.gotoOutlookPage();
  await outlook.searchMorningCall();

  const filename = `Anti Morning Call-${month}-${date}`;

  await outlook.answerForm();
  // await outlook.screenshotElement(".container", `${filename}-form`);

  // await outlook.submitForm();
  // await outlook.screenshotElement(".container", `${filename}-submit`);

  // const formImgUrl = await uploadGyazo(`${filename}-form.png`);
  // const submitImgUrl = await uploadGyazo(`${filename}-submit.png`);

  // await sendToSlack({ url: formImgUrl, title: `${filename}-form` });
  // await sendToSlack({ url: submitImgUrl, title: `${filename}-submit` });

  await outlook.close();
  // await sleep(10000);
})();
