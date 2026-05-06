/// <reference types="office-js" />

interface ThreadMessage {
  from?: string;
  date?: string;
  subject?: string;
  body: string;
}

interface DraftRequest {
  mode: "read" | "compose";
  subject: string;
  from?: { name: string; address: string };
  body: string;
  thread?: ThreadMessage[];
  tone: string;
  language: "auto" | "es" | "en";
  length: "auto" | "short" | "medium" | "long";
  instructions: string;
}

const PREF_KEYS = {
  tone: "rra.tone",
  language: "rra.language",
  length: "rra.length",
  instructions: "rra.instructions",
} as const;

function hydratePrefs() {
  for (const id of ["tone", "language", "length"] as const) {
    const stored = localStorage.getItem(PREF_KEYS[id]);
    if (stored) (document.getElementById(id) as HTMLSelectElement).value = stored;
  }
  const instr = localStorage.getItem(PREF_KEYS.instructions);
  if (instr) (document.getElementById("instructions") as HTMLTextAreaElement).value = instr;
}

function bindPrefPersistence() {
  for (const id of ["tone", "language", "length"] as const) {
    const el = document.getElementById(id) as HTMLSelectElement;
    el.addEventListener("change", () => localStorage.setItem(PREF_KEYS[id], el.value));
  }
  const instr = document.getElementById("instructions") as HTMLTextAreaElement;
  instr.addEventListener("input", () => localStorage.setItem(PREF_KEYS.instructions, instr.value));
}

const $ = <T extends HTMLElement = HTMLElement>(id: string) =>
  document.getElementById(id) as T;

let mode: "read" | "compose" = "read";
let currentAbort: AbortController | null = null;

Office.onReady(({ host }) => {
  if (host !== Office.HostType.Outlook) {
    setStatus("This add-in only runs in Outlook.");
    return;
  }

  const item = Office.context.mailbox.item;
  if (!item) {
    setStatus("No email selected.");
    return;
  }

  // Read-mode items expose displayReplyForm; compose-mode items don't.
  mode = "displayReplyForm" in item ? "read" : "compose";

  setStatus(mode === "read" ? "Ready. Incoming email loaded." : "Ready. Composing.");

  $("generate").addEventListener("click", () => generate());
  $("insert").addEventListener("click", () => insertDraft());
  $("copy").addEventListener("click", () => copyDraft());
  $("regen").addEventListener("click", () => generate());

  // Quick-action chips: clicking fills (or appends to) the instructions textarea
  document.querySelectorAll<HTMLButtonElement>(".chip").forEach((chip) => {
    chip.addEventListener("click", () => {
      const instr = chip.dataset.instr ?? "";
      const ta = $("instructions") as HTMLTextAreaElement;
      // Toggle: clicking the same active chip clears it
      if (chip.classList.contains("active")) {
        chip.classList.remove("active");
        if (ta.value === instr) ta.value = "";
      } else {
        document.querySelectorAll(".chip.active").forEach((c) => c.classList.remove("active"));
        chip.classList.add("active");
        ta.value = instr;
      }
    });
  });

  $("learn").addEventListener("click", () => learnFromSentItems());
  $("add-sample").addEventListener("click", () => addPastedSample());
  $("clear-learned").addEventListener("click", () => clearLearned());

  hydratePrefs();
  bindPrefPersistence();
  refreshSamplesCount();
});

async function addPastedSample() {
  const ta = $("paste-sample") as HTMLTextAreaElement;
  const raw = ta.value.trim();
  if (raw.length < 20) {
    $("samples-status").textContent = "Sample too short (need at least 20 chars).";
    return;
  }
  const cleaned = cleanSentBody(raw);
  if (cleaned.length < 20) {
    $("samples-status").textContent = "After cleanup the sample was too short. Paste more content.";
    return;
  }

  const btn = $("add-sample") as HTMLButtonElement;
  btn.disabled = true;
  try {
    const r = await fetch("/api/samples/add", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ reply: cleaned, language: guessLanguage(cleaned) }),
    });
    if (!r.ok) throw new Error(`API ${r.status}`);
    const data = (await r.json()) as { total: number; duplicate?: boolean };
    ta.value = "";
    $("samples-status").textContent = data.duplicate
      ? `Already saved (skipped duplicate) · ${data.total} total in use.`
      : `Sample added · ${data.total} total now in use.`;
  } catch (e) {
    $("samples-status").textContent = `Error: ${(e as Error).message}`;
  } finally {
    btn.disabled = false;
  }
}

async function clearLearned() {
  if (!confirm("Clear all learned samples? Seed samples in style-samples.json are kept.")) return;
  const r = await fetch("/api/samples/clear", { method: "POST" });
  const data = (await r.json()) as { total: number };
  $("samples-status").textContent = `Learned samples cleared · ${data.total} total now in use.`;
}

async function refreshSamplesCount() {
  try {
    const r = await fetch("/api/samples/count");
    const data = (await r.json()) as { total: number; seed: number; learned: number };
    $("samples-status").textContent = `${data.total} samples loaded (${data.seed} seed + ${data.learned} learned).`;
  } catch {
    $("samples-status").textContent = "Could not reach backend for sample count.";
  }
}

const SENT_ITEMS_SOAP_FULL = (max: number) => `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013"/>
  </soap:Header>
  <soap:Body>
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:BodyType>Text</t:BodyType>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject"/>
          <t:FieldURI FieldURI="item:Body"/>
          <t:FieldURI FieldURI="item:DateTimeSent"/>
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:IndexedPageItemView MaxEntriesReturned="${max}" Offset="0" BasePoint="Beginning"/>
      <m:SortOrder>
        <t:FieldOrder Order="Descending">
          <t:FieldURI FieldURI="item:DateTimeSent"/>
        </t:FieldOrder>
      </m:SortOrder>
      <m:ParentFolderIds>
        <t:DistinguishedFolderId Id="sentitems"/>
      </m:ParentFolderIds>
    </m:FindItem>
  </soap:Body>
</soap:Envelope>`;

const SENT_ITEMS_SOAP_MINIMAL = (max: number) => `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013"/>
  </soap:Header>
  <soap:Body>
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>AllProperties</t:BaseShape>
      </m:ItemShape>
      <m:IndexedPageItemView MaxEntriesReturned="${max}" Offset="0" BasePoint="Beginning"/>
      <m:ParentFolderIds>
        <t:DistinguishedFolderId Id="sentitems"/>
      </m:ParentFolderIds>
    </m:FindItem>
  </soap:Body>
</soap:Envelope>`;

async function tryEws(soap: string): Promise<{ ok: boolean; xml?: string; error?: string }> {
  const result = await new Promise<Office.AsyncResult<string>>((resolve) => {
    Office.context.mailbox.makeEwsRequestAsync(soap, resolve);
  });
  if (result.status !== Office.AsyncResultStatus.Succeeded) {
    return { ok: false, error: result.error?.message ?? "unknown" };
  }
  const xml = result.value;
  if (xml.includes("<faultcode>")) {
    const m = /<faultstring[^>]*>([^<]*)</.exec(xml);
    return { ok: false, error: `SOAP fault: ${m?.[1] ?? "unknown"}` };
  }
  // Check for ResponseClass="Error"
  if (/ResponseClass="Error"/.test(xml)) {
    const m = /<m:ResponseCode>([^<]+)<\/m:ResponseCode>/.exec(xml);
    return { ok: false, error: `EWS ${m?.[1] ?? "error"}` };
  }
  return { ok: true, xml };
}

/** Strip quoted thread, signature, and disclaimers; keep only the user's own writing. */
function cleanSentBody(body: string): string {
  if (!body) return "";
  // Cut at the first line that looks like a quoted-thread boundary.
  const lines = body.split(/\r?\n/);
  const cutPatterns = [
    /^\s*From:\s/i,
    /^\s*De:\s/i,
    /^\s*On\s.+\swrote:\s*$/i,
    /^\s*El\s.+\sescribió:\s*$/i,
    /^\s*Enviado:\s/i,
    /^\s*Sent:\s/i,
    /^\s*>+/,
    /^_{3,}\s*$/,
    /^-{3,}\s*Original Message/i,
    /^-{3,}\s*Mensaje original/i,
  ];
  let cut = lines.length;
  for (let i = 0; i < lines.length; i++) {
    if (cutPatterns.some((p) => p.test(lines[i]))) {
      cut = i;
      break;
    }
  }
  let out = lines.slice(0, cut).join("\n");

  // Strip standard email signature delimiter (line containing only "-- ").
  const sigIdx = out.search(/\n--\s*\n/);
  if (sigIdx !== -1) out = out.slice(0, sigIdx);

  return out.trim();
}

/** Heuristic language guess from a body. */
function guessLanguage(text: string): "es" | "en" {
  const sample = text.slice(0, 2000).toLowerCase();
  if (/[ñ¿¡áéíóú]/.test(sample)) return "es";
  const es = (sample.match(/\b(que|para|por|con|pero|hola|saludos|gracias|cuando)\b/g) ?? []).length;
  const en = (sample.match(/\b(the|and|for|with|that|but|hello|thanks|when)\b/g) ?? []).length;
  return es > en ? "es" : "en";
}

async function learnFromSentItems() {
  const btn = $("learn") as HTMLButtonElement;
  btn.disabled = true;
  $("samples-status").textContent = "Pulling sent items via EWS (full request)…";

  try {
    // Try the rich request first (sorted, body-only). If tenant is finicky, fall
    // back to the minimal request that asks for default shape only.
    let attempt = await tryEws(SENT_ITEMS_SOAP_FULL(25));
    if (!attempt.ok) {
      console.warn("[learn] full EWS request failed:", attempt.error, "(retrying minimal)");
      $("samples-status").textContent = "Full request failed. Retrying with minimal EWS shape…";
      attempt = await tryEws(SENT_ITEMS_SOAP_MINIMAL(25));
    }
    if (!attempt.ok) {
      $("samples-status").textContent = `EWS unavailable: ${attempt.error}. Use the manual paste below.`;
      return;
    }

    const doc = new DOMParser().parseFromString(attempt.xml!, "text/xml");
    const T = "http://schemas.microsoft.com/exchange/services/2006/types";
    const messages = doc.getElementsByTagNameNS(T, "Message");

    const learned: { reply: string; language: "es" | "en" }[] = [];
    const seen = new Set<string>();
    for (let i = 0; i < messages.length; i++) {
      const m = messages.item(i)!;
      const bodyText = m.getElementsByTagNameNS(T, "Body")[0]?.textContent ?? "";
      const cleaned = cleanSentBody(bodyText);
      // Filter: skip if too short, too long, or duplicate
      if (cleaned.length < 80) continue;
      if (cleaned.length > 4000) continue;
      const sig = cleaned.slice(0, 200);
      if (seen.has(sig)) continue;
      seen.add(sig);
      learned.push({ reply: cleaned, language: guessLanguage(cleaned) });
    }

    if (learned.length === 0) {
      $("samples-status").textContent = `Found ${messages.length} sent items, but none had usable cleaned bodies.`;
      return;
    }

    $("samples-status").textContent = `Saving ${learned.length} samples…`;
    const r = await fetch("/api/samples/learn", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(learned),
    });
    const data = (await r.json()) as { learned: number; total: number };
    $("samples-status").textContent = `Learned ${data.learned} samples · ${data.total} total now in use.`;
  } catch (e) {
    $("samples-status").textContent = `Error: ${(e as Error).message}`;
  } finally {
    btn.disabled = false;
  }
}

function setStatus(text: string) {
  $("status").textContent = text;
}

async function readEmail(): Promise<{ subject: string; from?: { name: string; address: string }; body: string }> {
  const item = Office.context.mailbox.item!;

  const subject = await new Promise<string>((resolve) => {
    if (typeof item.subject === "string") resolve(item.subject);
    else (item.subject as Office.Subject).getAsync((r) => resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : ""));
  });

  let from: { name: string; address: string } | undefined;
  if (mode === "read" && (item as Office.MessageRead).from) {
    const f = (item as Office.MessageRead).from;
    from = { name: f.displayName, address: f.emailAddress };
  }

  const body = await new Promise<string>((resolve) => {
    item.body.getAsync(Office.CoercionType.Text, (r) =>
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : "")
    );
  });

  return { subject, from, body };
}

/**
 * Pull the rest of the conversation via EWS. Returns [] if EWS is disabled on
 * the tenant, no conversationId is available, or the SOAP response can't be
 * parsed. Caller should treat the empty case as "use body only".
 *
 * Logs diagnostics to the browser console so failures are visible.
 */
async function fetchThread(): Promise<{ messages: ThreadMessage[]; reason: string }> {
  const item = Office.context.mailbox.item;
  const convId = (item as Office.MessageRead | undefined)?.conversationId;
  if (!convId) {
    console.warn("[thread] no conversationId on item (compose mode or unavailable)");
    return { messages: [], reason: "no conversationId" };
  }

  const soap = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:xsd="http://www.w3.org/2001/XMLSchema"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013"/>
  </soap:Header>
  <soap:Body>
    <m:GetConversationItems xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:BodyType>Text</t:BodyType>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject"/>
          <t:FieldURI FieldURI="item:Body"/>
          <t:FieldURI FieldURI="item:DateTimeReceived"/>
          <t:FieldURI FieldURI="message:From"/>
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:Conversations>
        <t:Conversation>
          <t:ConversationId Id="${escapeXml(convId)}"/>
        </t:Conversation>
      </m:Conversations>
    </m:GetConversationItems>
  </soap:Body>
</soap:Envelope>`;

  const ewsResult = await new Promise<Office.AsyncResult<string>>((resolve) => {
    Office.context.mailbox.makeEwsRequestAsync(soap, resolve);
  });

  if (ewsResult.status !== Office.AsyncResultStatus.Succeeded) {
    console.warn("[thread] makeEwsRequestAsync failed:", ewsResult.error);
    return { messages: [], reason: `EWS error: ${ewsResult.error?.message ?? "unknown"}` };
  }

  const xml = ewsResult.value;
  console.log("[thread] EWS response length:", xml?.length ?? 0);

  // Detect SOAP fault before trying to parse messages
  if (xml && /<(soap|faultcode|m:ResponseClass="Error")/i.test(xml.slice(0, 500)) === false && xml.includes("ResponseClass=\"Error\"")) {
    console.warn("[thread] EWS response contains an error class:", xml.slice(0, 1000));
  }
  if (xml && xml.includes("<faultcode>")) {
    const m = /<faultstring[^>]*>([^<]*)</.exec(xml);
    console.warn("[thread] SOAP fault:", m?.[1] ?? "(unknown)");
    return { messages: [], reason: `SOAP fault: ${m?.[1] ?? "unknown"}` };
  }

  try {
    const doc = new DOMParser().parseFromString(xml, "text/xml");
    const T = "http://schemas.microsoft.com/exchange/services/2006/types";
    const messages = doc.getElementsByTagNameNS(T, "Message");
    console.log(`[thread] parsed ${messages.length} <t:Message> elements`);

    // EWS sometimes returns an MessageResponseShape error per conversation
    const responseMessages = doc.getElementsByTagNameNS(
      "http://schemas.microsoft.com/exchange/services/2006/messages",
      "GetConversationItemsResponseMessage",
    );
    if (responseMessages.length > 0) {
      const cls = responseMessages[0].getAttribute("ResponseClass");
      const code = responseMessages[0].getElementsByTagNameNS(
        "http://schemas.microsoft.com/exchange/services/2006/messages",
        "ResponseCode",
      )[0]?.textContent;
      console.log(`[thread] ResponseClass=${cls} ResponseCode=${code}`);
      if (cls === "Error") {
        return { messages: [], reason: `EWS ${code ?? "error"}` };
      }
    }

    const out: ThreadMessage[] = [];
    for (let i = 0; i < messages.length; i++) {
      const m = messages.item(i)!;
      const subject = m.getElementsByTagNameNS(T, "Subject")[0]?.textContent ?? undefined;
      const date = m.getElementsByTagNameNS(T, "DateTimeReceived")[0]?.textContent ?? undefined;
      const body = m.getElementsByTagNameNS(T, "Body")[0]?.textContent ?? "";
      const fromMailbox = m.getElementsByTagNameNS(T, "From")[0]?.getElementsByTagNameNS(T, "Mailbox")[0];
      const fromName = fromMailbox?.getElementsByTagNameNS(T, "Name")[0]?.textContent ?? "";
      const fromAddr = fromMailbox?.getElementsByTagNameNS(T, "EmailAddress")[0]?.textContent ?? "";
      const from = [fromName, fromAddr ? `<${fromAddr}>` : ""].filter(Boolean).join(" ");
      if (body.trim()) out.push({ from, date, subject, body: body.trim() });
    }
    out.sort((a, b) => (a.date ?? "").localeCompare(b.date ?? ""));
    return { messages: out, reason: out.length ? "ok" : "no messages in response" };
  } catch (e) {
    console.warn("[thread] parse error:", e);
    return { messages: [], reason: "parse error" };
  }
}

function escapeXml(s: string): string {
  return s.replace(/[<>&"']/g, (c) => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;", '"': "&quot;", "'": "&apos;" })[c]!);
}

async function generate() {
  const btn = $("generate") as HTMLButtonElement;
  btn.disabled = true;
  setStatus("Reading email…");

  try {
    const { subject, from, body } = await readEmail();

    setStatus("Loading thread…");
    const { messages: thread, reason: threadReason } = await fetchThread();
    if (thread.length > 1) setStatus(`Drafting reply (${thread.length} messages in thread)…`);
    else setStatus("Drafting reply…");

    const payload: DraftRequest = {
      mode,
      subject,
      from,
      body,
      thread: thread.length > 0 ? thread : undefined,
      tone: ($("tone") as HTMLSelectElement).value,
      language: ($("language") as HTMLSelectElement).value as DraftRequest["language"],
      length: ($("length") as HTMLSelectElement).value as DraftRequest["length"],
      instructions: ($("instructions") as HTMLTextAreaElement).value,
    };

    // Cancel any prior in-flight stream (e.g. user clicked Regenerate)
    currentAbort?.abort();
    const controller = new AbortController();
    currentAbort = controller;

    const draftEl = $("draft") as HTMLTextAreaElement;
    draftEl.value = "";
    document.querySelector<HTMLElement>(".output")?.removeAttribute("hidden");

    const res = await fetch("/api/draft", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
      signal: controller.signal,
    });

    if (!res.ok) {
      const err = await res.text();
      throw new Error(`API ${res.status}: ${err}`);
    }

    const detectedLang = res.headers.get("X-Detected-Language");
    const reader = res.body!.getReader();
    const decoder = new TextDecoder();

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      draftEl.value += decoder.decode(value, { stream: true });
      draftEl.scrollTop = draftEl.scrollHeight;
    }
    draftEl.value += decoder.decode();

    const parts: string[] = ["Drafted"];
    if (detectedLang) parts.push(`(${detectedLang})`);
    parts.push(thread.length > 0 ? `· thread: ${thread.length} msgs` : `· no thread (${threadReason})`);
    setStatus(parts.join(" "));
  } catch (e) {
    if ((e as Error).name === "AbortError") {
      setStatus("Cancelled.");
    } else {
      setStatus(`Error: ${(e as Error).message}`);
    }
  } finally {
    btn.disabled = false;
    if (currentAbort?.signal.aborted === false) currentAbort = null;
  }
}

function insertDraft() {
  const draft = ($("draft") as HTMLTextAreaElement).value;
  if (!draft.trim()) return;

  const item = Office.context.mailbox.item!;

  if (mode === "read") {
    (item as Office.MessageRead).displayReplyForm({ htmlBody: draftToHtml(draft) });
    setStatus("Reply form opened with draft.");
    return;
  }

  item.body.setSelectedDataAsync(
    draftToHtml(draft),
    { coercionType: Office.CoercionType.Html },
    (r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) setStatus("Inserted.");
      else setStatus(`Insert failed: ${r.error.message}`);
    }
  );
}

async function copyDraft() {
  const draft = ($("draft") as HTMLTextAreaElement).value;
  await navigator.clipboard.writeText(draft);
  setStatus("Copied to clipboard.");
}

function draftToHtml(text: string): string {
  return text
    .split(/\n\n+/)
    .map((p) => `<p>${escapeHtml(p).replace(/\n/g, "<br>")}</p>`)
    .join("");
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);
}
