import axios from "axios";

export async function sendToSlack(img: { url: string; title: string }) {
  const { SLACK_INCOMING_WEB_HOOK } = process.env;

  if (!SLACK_INCOMING_WEB_HOOK)
    throw new Error(`SLACK_INCOMING_WEB_HOOK: ${SLACK_INCOMING_WEB_HOOK}`);

  const res = await axios.post(SLACK_INCOMING_WEB_HOOK, {
    blocks: [
      {
        type: "image",
        title: {
          type: "plain_text",
          text: img.title,
          emoji: true,
        },
        image_url: img.url,
        alt_text: "marg",
      },
    ],
  });
  if (res.data === "ok") {
    console.log("sendToSlack: successed");
  } else {
    console.log("sendToSlack: failed");
    console.log(res);
  }
}
