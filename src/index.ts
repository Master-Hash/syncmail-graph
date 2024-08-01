/**
 * Welcome to Cloudflare Workers!
 *
 * This is a template for a Scheduled Worker: a Worker that can run on a
 * configurable interval:
 * https://developers.cloudflare.com/workers/platform/triggers/cron-triggers/
 *
 * - Run `npm run dev` in your terminal to start a development server
 * - Run `curl "http://localhost:8787/__scheduled?cron=*+*+*+*+*"` to see your worker in action
 * - Run `npm run deploy` to publish your worker
 *
 * Bind resources to your worker in `wrangler.toml`. After adding bindings, a type definition for the
 * `Env` object can be regenerated with `npm run cf-typegen`.
 *
 * Learn more at https://developers.cloudflare.com/workers/
 */

import { sha256 } from "hono/utils/crypto";

export default {
  // The scheduled handler is invoked at the interval set in our wrangler.toml's
  // [[triggers]] configuration.
  async scheduled(event, env, ctx): Promise<void> {
    // A Cron Trigger can make requests to other endpoints on the Internet,
    // publish to a Queue, query a D1 Database, and much more.
    //
    // We'll keep it simple and make an API call to a Cloudflare API:
    // let resp = await fetch('https://api.cloudflare.com/client/v4/ips');
    // let wasSuccessful = resp.ok ? 'success' : 'fail';
    //
    // You could store this result in KV, write to a D1 Database, or publish to a Queue.
    // In this template, we'll just log the result:
    // console.log(`trigger fired at ${event.cron}: ${wasSuccessful}`);

    const Modify1 = env.DB.prepare("INSERT OR IGNORE INTO GlobalMessages (Folder, MessageID, MessageIDHash, Epoch, InReplyTo, SubjectLine, Author, Recipients, RAWMessage, FolderSerial) VALUES (?,?,?,?,?,?,?,?,?,?)");
    const Modify2 = env.DB.prepare("UPDATE GlobalMessages SET FolderSerial = ? WHERE MessageIDHash = ?");
    const Modify3 = env.DB.prepare("DELETE FROM GlobalMessages WHERE Folder = ? AND FolderSerial > ?");
    const Modify4 = env.DB.prepare("DELETE FROM GlobalMessages WHERE Folder = ? AND MessageID <> ? AND FolderSerial = ?");


    const accessToken = await env.AZ_TOKENS.get("access_token");
    if (accessToken === null) {
      const refreshToken = await env.AZ_TOKENS.get("refresh_token");
      if (refreshToken === null || !env.CLIENT_ID || !env.CLIENT_SECRET) {
        throw new Error("No refresh token found");
      }
      const d = new URLSearchParams();
      d.append("client_id", env.CLIENT_ID);
      d.append("scope", "User.Read Mail.ReadBasic offline_access");
      d.append("refresh_token", refreshToken);
      d.append("grant_type", "refresh_token");
      d.append("client_secret", env.CLIENT_SECRET);
      console.log("Refreshing token", d.toString().replaceAll("+", "%20"));
      const resp = await fetch("https://login.microsoftonline.com/consumers/oauth2/v2.0/token",
        {
          method: "POST",
          headers: {
            "Content-Type": "application/x-www-form-urlencoded"
          },
          body: d.toString().replaceAll("+", "%20"),
        }
      );
      if (!resp.ok) {
        console.log(resp);
        console.log(await resp.text());
        throw new Error("Failed to refresh token");
      }
      const data = await resp.json() as {
        token_type: "Bearer";
        scope: string;
        expires_in: number;
        ext_expires_in: number;
        access_token: string;
        refresh_token: string;
      };
      await env.AZ_TOKENS.put("access_token", data.access_token, { expirationTtl: data.expires_in - 60 });
      await env.AZ_TOKENS.put("refresh_token", data.refresh_token);
    }
    const _folders = await fetch("https://graph.microsoft.com/v1.0/me/mailFolders?$filter=displayName%20eq%20'收件箱'%20or%20displayName%20eq%20'已发送邮件'&$select=displayName,totalItemCount", {
      headers: {
        "Authorization": `Bearer ${accessToken}`
      }
    });
    const folders = await _folders.json() as {
      value: Array<{
        displayName: string;
        totalItemCount: number;
      }>;
    };
    console.log(folders);
    // const f = Object.groupBy(folders.value, ({ displayName }) => displayName);
    const inbox = folders.value.find(({ displayName }) => displayName === "收件箱");
    const _i = await fetch("https://graph.microsoft.com/v1.0/me/mailFolders('Inbox')/messages?$select=receivedDateTime,subject,internetMessageId,internetMessageHeaders,from,toRecipients,ccRecipients&$top=16", {
      headers: {
        "Authorization": `Bearer ${accessToken}`
      }
    });
    const i = await _i.json() as { value: Message[]; };
    console.log(i);
    await modifyDatabaseStatements("Inbox", i.value, inbox!.totalItemCount);

    const sentItems = folders.value.find(({ displayName }) => displayName === "已发送邮件");
    const _s = await fetch("https://graph.microsoft.com/v1.0/me/mailFolders('SentItems')/messages?$select=receivedDateTime,subject,internetMessageId,internetMessageHeaders,from,toRecipients,ccRecipients&$top=16", {
      headers: {
        "Authorization": `Bearer ${accessToken}`
      }
    });
    const s = await _s.json() as { value: Message[]; };
    console.log(s);
    await modifyDatabaseStatements("Sent", s.value, sentItems!.totalItemCount);

    async function modifyDatabaseStatements(box: string, messages: Message[], totalItemCount: number): Promise<void> {
      const batch = [] as Array<D1PreparedStatement>;
      for (const [index, message] of messages.entries()) {
        const serial = totalItemCount - index;
        const s = await sha256(message.internetMessageId);
        batch.push(Modify1
          .bind(
            box,
            message.internetMessageId,
            s,
            new Date(message.receivedDateTime).valueOf() / 1000,
            // message.internetMessageHeaders.find(({ name }) => name.toLowerCase() === "in-reply-to")?.value ?? null,
            null,
            message.subject,
            JSON.stringify(message.from.emailAddress),
            JSON.stringify(message.toRecipients.map(({ emailAddress }) => emailAddress)),
            0,
            serial,
          ));
        batch.push(Modify2
          .bind(
            serial,
            s,
          ));
      }
      batch.push(Modify3
        .bind(
          box,
          totalItemCount,
        ));
      for (const [index, message] of messages.entries()) {
        const serial = totalItemCount - index;
        batch.push(
          Modify4
            .bind(
              box,
              message.internetMessageId,
              serial,
            )
        );
      }
      const e = await env.DB.batch(batch);
      console.log(e);
    }


  },
} satisfies ExportedHandler<Env>;



type Message = {
  receivedDateTime: string;
  internetMessageId: string;
  subject: string;
  internetMessageHeaders: Array<{
    name: string;
    value: string;
  }>;
  from: EmailAddress;
  toRecipients: Array<EmailAddress>;
  ccRecipients: Array<EmailAddress>;
};

type EmailAddress = {
  emailAddress: {
    name: string;
    address: string;
  };
}

