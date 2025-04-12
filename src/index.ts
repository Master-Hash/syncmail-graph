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

    const session = env.DB.withSession("first-primary");
    const Modify5 = session.prepare("INSERT INTO GlobalMessages (Folder, MessageID, MessageIDHash, Epoch, InReplyTo, SubjectLine, Author, Recipients, RAWMessage, FolderSerial) VALUES(?,?,?,?,?,?,?,?,?,?),(?,?,?,?,?,?,?,?,?,?),(?,?,?,?,?,?,?,?,?,?),(?,?,?,?,?,?,?,?,?,?),(?,?,?,?,?,?,?,?,?,?),(?,?,?,?,?,?,?,?,?,?),(?,?,?,?,?,?,?,?,?,?),(?,?,?,?,?,?,?,?,?,?) ON conflict (MessageID) do UPDATE SET FolderSerial = excluded.FolderSerial WHERE GlobalMessages.FolderSerial IS DISTINCT FROM excluded.FolderSerial");
    const Modify6 = session.prepare("WITH Cleaner (MessageID) AS (VALUES (?),(?),(?),(?),(?),(?),(?),(?)) DELETE FROM GlobalMessages WHERE Folder = ? AND ((MessageID NOT IN (SELECT MessageID FROM Cleaner) AND FolderSerial > ? - 8) OR FolderSerial > ?)");

    let accessToken = await env.AZ_TOKENS.get("access_token");
    if (accessToken === null) {
      const refreshToken = await env.AZ_TOKENS.get("refresh_token");
      if (refreshToken === null || !env.CLIENT_ID || !env.CLIENT_SECRET) {
        throw new Error("No refresh token found");
      }
      const d = new URLSearchParams();
      d.append("client_id", env.CLIENT_ID);
      d.append("scope", "User.Read Mail.Read offline_access");
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
      await env.AZ_TOKENS.put("access_token", data.access_token, { expirationTtl: data.expires_in - 90 });
      await env.AZ_TOKENS.put("refresh_token", data.refresh_token);
      accessToken = data.access_token;
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
    const sentItems = folders.value.find(({ displayName }) => displayName === "已发送邮件");
    const [_i, _s] = await Promise.all([fetch("https://graph.microsoft.com/v1.0/me/mailFolders('Inbox')/messages?$select=receivedDateTime,subject,internetMessageId,from,toRecipients,ccRecipients&$top=8&$expand=singleValueExtendedProperties($filter=id%20eq%20'String%200x1042')", {
      headers: {
        "Authorization": `Bearer ${accessToken}`
      }
    }), fetch("https://graph.microsoft.com/v1.0/me/mailFolders('SentItems')/messages?$select=receivedDateTime,subject,internetMessageId,from,toRecipients,ccRecipients&$top=8&$expand=singleValueExtendedProperties($filter=id%20eq%20'String%200x1042')", {
      headers: {
        "Authorization": `Bearer ${accessToken}`
      }
    })]);
    const [i, s] = await Promise.all([
      _i.json(),
      _s.json()
    ]) as [
        { value: OutlookMessage[]; },
        { value: OutlookMessage[]; }
      ];

    // const i = await _i.json() as { value: OutlookMessage[]; };
    // console.log(i.value.find(({ singleValueExtendedProperties }) => singleValueExtendedProperties)?.singleValueExtendedProperties);
    // console.log(i);
    const s1 = await modifyDatabaseStatements2("Inbox", i.value, inbox!.totalItemCount);

    // const s = await _s.json() as { value: OutlookMessage[]; };
    // console.log(s.value.find(({ singleValueExtendedProperties }) => singleValueExtendedProperties)?.singleValueExtendedProperties);
    const s2 = await modifyDatabaseStatements2("Sent", s.value, sentItems!.totalItemCount);

    const e = await env.DB.batch([
      ...s1,
      ...s2,
    ]);

    console.log(e);

    async function modifyDatabaseStatements2(box: "Inbox" | "Sent", messages: OutlookMessage[], totalItemCount: number): Promise<Array<D1PreparedStatement>> {
      const batch = [] as Array<D1PreparedStatement>;
      const argsUpsert = [] as Array<[
        "Inbox" | "Sent", // Folder
        string, // MessageID
        string, // MessageIDHash
        number, // Epoch
        string | null, // InReplyTo
        string, // SubjectLine
        string, // Author
        string, // Recipients
        0, // RAWMessage
        number, // FolderSerial
      ]>;
      for (const [index, message] of messages.entries()) {
        const serial = totalItemCount - index;
        const s = await sha256(message.internetMessageId);
        argsUpsert.push([
          box,
          message.internetMessageId,
          s!,
          new Date(message.receivedDateTime).valueOf() / 1000,
          message.singleValueExtendedProperties?.find(({ id }) => id === "String 0x1042")?.value ?? null,
          message.subject,
          JSON.stringify(message.from.emailAddress),
          JSON.stringify([...message.toRecipients, ...message.ccRecipients].map(({ emailAddress }) => emailAddress)),
          0,
          serial,
        ]);
      }
      console.log(argsUpsert.flat());
      batch.push(Modify5.bind(...argsUpsert.flat()));

      batch.push(Modify6.bind(
        ...argsUpsert.map(i => i[1]),
        box,
        totalItemCount,
        totalItemCount,
      ));
      return batch;
    }
  },
} satisfies ExportedHandler<Env>;



type OutlookMessage = {
  receivedDateTime: string;
  internetMessageId: string;
  subject: string;
  // internetMessageHeaders: Array<{
  //   name: string;
  //   value: string;
  // }>;
  from: EmailAddress;
  toRecipients: Array<EmailAddress>;
  ccRecipients: Array<EmailAddress>;
  singleValueExtendedProperties?: Array<{
    id: string;
    value: string;
  }>;
};

type EmailAddress = {
  emailAddress: {
    name: string;
    address: string;
  };
}

