import "dotenv/config";
import express, { Request, Response } from "express";
import cors from "cors";
import Anthropic from "@anthropic-ai/sdk";
import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(__dirname, "..");
const SEED_PATH = resolve(ROOT, "style-samples.json");
const LEARNED_PATH = resolve(ROOT, "learned-samples.json");

interface StyleSample {
  context?: string;       // the email being replied to (optional; learned samples have no context)
  reply: string;
  language?: "es" | "en";
}

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

const LEARN_CAP = 25;

const LENGTH_INSTRUCTIONS: Record<DraftRequest["length"], string | null> = {
  auto: null,
  short: "Length: keep it to 1 to 3 sentences. No greeting beyond a single word if any.",
  medium: "Length: one short paragraph (3 to 6 sentences).",
  long: "Length: multiple paragraphs are fine if the situation calls for it.",
};

function formatThread(thread: ThreadMessage[]): string {
  return thread
    .map((m, i) => {
      const header = [
        `Message ${i + 1}/${thread.length}`,
        m.from ? `From: ${m.from}` : null,
        m.date ? `Date: ${m.date}` : null,
        m.subject ? `Subject: ${m.subject}` : null,
      ]
        .filter(Boolean)
        .join(" | ");
      return `[${header}]\n${m.body.trim()}`;
    })
    .join("\n\n---\n\n");
}

function detectLanguage(text: string): "es" | "en" {
  const sample = text.slice(0, 4000).toLowerCase();
  if (/[ñ¿¡]|á|é|í|ó|ú/.test(sample)) return "es";
  const esHits = (sample.match(/\b(que|para|por|con|pero|porque|hola|saludos|gracias|favor|cuando|donde|este|esta|estos|estas|cualquier|hacia|sobre|entre)\b/g) ?? []).length;
  const enHits = (sample.match(/\b(the|and|for|with|that|but|because|hello|thanks|when|where|this|these|any|toward|about|between)\b/g) ?? []).length;
  return esHits > enHits ? "es" : "en";
}

function loadSamples(): StyleSample[] {
  const seed: StyleSample[] = JSON.parse(readFileSync(SEED_PATH, "utf-8"));
  const learned: StyleSample[] = existsSync(LEARNED_PATH)
    ? JSON.parse(readFileSync(LEARNED_PATH, "utf-8"))
    : [];
  return [...seed, ...learned];
}

const SYSTEM_INSTRUCTIONS = `You draft email replies in the user's voice. Match their cadence, register, vocabulary, and sign-off style. Do not invent a generic professional tone.

Hard rules:
- Output only the reply body. No subject line, no quoted thread, no commentary about your own draft.
- Mirror the language of the incoming email (Spanish if the email is in Spanish, English if in English). If the user's tone instruction is "match incoming", also mirror its formality.
- Preserve concrete commitments and dates from the incoming email. Do not invent facts. If something needs a fact you don't have, leave a clear "[fill in: X]" placeholder.
- If the user gave extra instructions, follow them. They override defaults.
- Never use em dashes (—) or en dashes (–). Replace with commas, periods, colons, parentheses, or two separate sentences. This rule overrides any tendency toward "literary" punctuation.
- Be concise. Replies should be as short as the situation allows.

Below are real reply samples written by the user. Treat them as the canonical source of voice, tone, structure, opening style, and sign-off. Do not copy phrases; absorb the style.`;

function buildStyleBlock(samples: StyleSample[]): string {
  return samples
    .map((s, i) => {
      const lang = s.language ? ` lang="${s.language}"` : "";
      if (s.context && s.context.trim()) {
        return (
          `<sample n="${i + 1}"${lang}>\n` +
          `<incoming>\n${s.context.trim()}\n</incoming>\n` +
          `<reply>\n${s.reply.trim()}\n</reply>\n` +
          `</sample>`
        );
      }
      return (
        `<sample n="${i + 1}"${lang}>\n` +
        `<email_i_wrote>\n${s.reply.trim()}\n</email_i_wrote>\n` +
        `</sample>`
      );
    })
    .join("\n\n");
}

const app = express();
app.use(cors());
app.use(express.json({ limit: "1mb" }));

const client = new Anthropic();
const MODEL = process.env.ANTHROPIC_MODEL ?? "claude-sonnet-4-6";

function loadLearned(): StyleSample[] {
  return existsSync(LEARNED_PATH) ? JSON.parse(readFileSync(LEARNED_PATH, "utf-8")) : [];
}

app.post("/api/samples/learn", (req: Request, res: Response) => {
  const samples = req.body as StyleSample[];
  if (!Array.isArray(samples)) {
    return res.status(400).json({ error: "expected an array of {reply, language?}" });
  }
  // Cap to keep cached system prompt bounded; client sends newest first.
  const capped = samples.slice(0, LEARN_CAP);
  writeFileSync(LEARNED_PATH, JSON.stringify(capped, null, 2));
  const all = loadSamples();
  console.log(`[learn] saved ${capped.length} learned samples (cap ${LEARN_CAP}, ${all.length} total)`);
  res.json({ learned: capped.length, total: all.length, cap: LEARN_CAP });
});

app.post("/api/samples/add", (req: Request, res: Response) => {
  const sample = req.body as StyleSample;
  if (!sample || typeof sample.reply !== "string" || sample.reply.trim().length < 20) {
    return res.status(400).json({ error: "reply must be a string of at least 20 chars" });
  }
  const reply = sample.reply.trim();
  const learned = loadLearned();

  // Dedup: skip if this reply (or near-identical) is already saved
  const fingerprint = (s: string) => s.replace(/\s+/g, " ").slice(0, 200).toLowerCase();
  const fp = fingerprint(reply);
  const isDup = learned.some((s) => fingerprint(s.reply) === fp);
  if (isDup) {
    const all = loadSamples();
    console.log(`[add] skipped duplicate (${all.length} total)`);
    return res.json({ added: 0, duplicate: true, total: all.length });
  }

  learned.push({ reply, language: sample.language });
  // Evict oldest if over the cap.
  while (learned.length > LEARN_CAP) learned.shift();
  writeFileSync(LEARNED_PATH, JSON.stringify(learned, null, 2));
  const all = loadSamples();
  console.log(`[add] appended sample (${all.length} total, learned at ${learned.length}/${LEARN_CAP})`);
  res.json({ added: 1, duplicate: false, total: all.length });
});

app.post("/api/samples/clear", (_req: Request, res: Response) => {
  writeFileSync(LEARNED_PATH, "[]\n");
  const all = loadSamples();
  res.json({ cleared: true, total: all.length });
});

app.get("/api/samples/count", (_req, res) => {
  const all = loadSamples();
  const seedCount = JSON.parse(readFileSync(SEED_PATH, "utf-8")).length;
  res.json({ total: all.length, seed: seedCount, learned: all.length - seedCount });
});

app.post("/api/draft", async (req: Request, res: Response) => {
  const body = req.body as DraftRequest;

  if (!body || typeof body.body !== "string") {
    return res.status(400).json({ error: "missing email body" });
  }

  const detected = detectLanguage(body.body);
  const targetLanguage = body.language === "auto" ? detected : body.language;
  const langInstruction =
    body.language === "auto"
      ? `Reply in ${targetLanguage === "es" ? "Spanish" : "English"} (mirroring the incoming email).`
      : `Reply in ${targetLanguage === "es" ? "Spanish" : "English"} regardless of the incoming email's language.`;

  const hasThread = !!body.thread && body.thread.length > 1;

  const userMsg = [
    hasThread ? "--- conversation history (oldest first) ---" : null,
    hasThread ? formatThread(body.thread!) : null,
    hasThread ? "--- end of history ---" : null,
    hasThread ? "" : null,
    "--- the email I am replying to ---",
    body.from ? `From: ${body.from.name} <${body.from.address}>` : null,
    body.subject ? `Subject: ${body.subject}` : null,
    "",
    body.body.trim(),
    "--- end ---",
    "",
    `Language: ${langInstruction}`,
    `Tone: ${body.tone}`,
    LENGTH_INSTRUCTIONS[body.length ?? "auto"],
    body.instructions.trim() ? `Extra instructions from me: ${body.instructions.trim()}` : null,
    "",
    body.mode === "compose"
      ? "I have already started composing. Draft a complete reply that I can paste into the compose window."
      : hasThread
        ? "Draft a reply. Use the conversation history above for context, but only respond to the latest email."
        : "Draft a reply. The email body may contain quoted text from earlier messages in the thread (often prefixed with '>' or 'On <date>, X wrote:'). Use that as context if present, but respond only to the most recent message.",
  ]
    .filter((x) => x !== null)
    .join("\n");

  console.log(
    `\n[draft] mode=${body.mode} lang=${targetLanguage} tone=${body.tone} length=${body.length ?? "auto"}` +
      (body.from ? ` from="${body.from.name} <${body.from.address}>"` : "") +
      ` subject="${body.subject}"` +
      ` body_len=${body.body.length}` +
      ` thread=${body.thread?.length ?? 0}`,
  );
  if (process.env.LOG_BODY === "1") console.log(`[draft] body:\n${body.body}\n---`);

  res.setHeader("Content-Type", "text/plain; charset=utf-8");
  res.setHeader("X-Detected-Language", targetLanguage);
  res.setHeader("X-Thread-Count", String(body.thread?.length ?? 0));
  res.flushHeaders();

  try {
    const samples = loadSamples();
    const stream = client.messages.stream({
      model: MODEL,
      max_tokens: 2048,
      thinking: { type: "adaptive" },
      output_config: { effort: "high" },
      system: [
        { type: "text", text: SYSTEM_INSTRUCTIONS },
        {
          type: "text",
          text: buildStyleBlock(samples),
          cache_control: { type: "ephemeral" },
        },
      ],
      messages: [{ role: "user", content: userMsg }],
    });

    let carry = "";
    for await (const event of stream) {
      if (event.type === "content_block_delta" && event.delta.type === "text_delta") {
        const chunk = carry + event.delta.text;
        // Hold the last char in case a dash is split across chunks (paranoid, but cheap).
        carry = chunk.slice(-1);
        const emit = chunk.slice(0, -1).replace(/\s*[—–]\s*/g, ", ");
        if (emit) res.write(emit);
      }
    }
    if (carry) res.write(carry.replace(/[—–]/g, ","));
    res.end();

    const final = await stream.finalMessage();
    const u = final.usage;
    const cacheRead = u.cache_read_input_tokens ?? 0;
    const cacheWrite = u.cache_creation_input_tokens ?? 0;
    const hitPct = cacheRead + cacheWrite > 0 ? Math.round((cacheRead / (cacheRead + cacheWrite)) * 100) : 0;
    console.log(
      `[draft] usage in=${u.input_tokens} out=${u.output_tokens}` +
        ` cache_read=${cacheRead} cache_write=${cacheWrite} cache_hit=${hitPct}%` +
        ` stop=${final.stop_reason}`,
    );
  } catch (e) {
    const err = e as Error;
    console.error("[draft] error:", err);
    if (!res.headersSent) {
      res.status(500).json({ error: err.message });
    } else {
      res.write(`\n\n[error: ${err.message}]`);
      res.end();
    }
  }
});

app.get("/api/health", (_req, res) => res.json({ ok: true, model: MODEL, samples: loadSamples().length }));

if (!process.env.ANTHROPIC_API_KEY) {
  console.error("error: ANTHROPIC_API_KEY is not set. Export it in your shell (or put it in .env).");
  process.exit(1);
}

const port = Number(process.env.PORT ?? 3001);
app.listen(port, () => {
  console.log(`api listening on http://localhost:${port}`);
  console.log(`model: ${MODEL}, style samples: ${loadSamples().length}`);
});
