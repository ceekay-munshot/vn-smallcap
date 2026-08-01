#!/usr/bin/env node
// Governance flags scraper — v2.
//
// PROBLEM with v1 (May 2026): scraped 4 landing pages (BSE corporate
// announcements, BSE press releases, NSE corporate filings, SEBI orders
// listing) but none of them actually surface SEBI orders against named
// listed entities. The BSE/NSE pages list routine LODR Reg-30 filings
// (board meetings, financial results); SEBI's orders.html lists order
// titles like "WTM/SK/CFD/CFD-SEC-2/29183/2024" — no entity names. Net
// result: 0 flagged companies in most weekly runs, even though high-
// profile SEBI activity against Nifty 500 names is well-documented.
//
// v2 STRATEGY: financial-news SEBI topic pages. Mint / Business Standard
// / Economic Times each maintain a "SEBI" topic page where every story
// of the last few weeks is listed with a headline. Headlines naturally
// name the listed entity ("SEBI fines XYZ Ltd ₹2 cr for…", "Adani gets
// SEBI notice on…"). LLM extraction works cleanly on this surface.
//
// LAYERS:
//   1) Curated baseline (screener-test/static/known-sebi-cases.json) —
//      hand-maintained high-profile cases so we're never at zero coverage.
//   2) News topic pages (Mint + BS + ET) — caught automatically.
//   3) SEBI orders page (kept as backup; rarely fires but no extra cost).
//   4) Combined, deduped, fuzzy-matched to Nifty 500 tickers.
//
// Cost: ~10-20 Firecrawl credits per weekly run (one credit per source).

import { readFileSync, writeFileSync, mkdirSync, existsSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const SCREENER_PATH = resolve(__dirname, "../public/data/screener-companies.json");
const CURATED_PATH  = resolve(__dirname, "static/known-sebi-cases.json");
const OUT_PATH      = resolve(__dirname, "../public/data/governance-flags.json");

const FIRECRAWL_URL = "https://api.firecrawl.dev/v1/scrape";
const API_KEY = process.env.FIRECRAWL_API_KEY;

// ---------- Source list ----------
// News pages name listed entities directly in headlines — the highest-
// signal surface we can scrape. SEBI's own orders.html stays as a backup
// (occasionally catches direct mentions when an order title includes a
// company name).
const SOURCES = [
  { url: "https://www.livemint.com/topic/sebi",                label: "Mint — SEBI topic (news)",                    mode: "news" },
  { url: "https://www.business-standard.com/topic/sebi",       label: "Business Standard — SEBI topic (news)",       mode: "news" },
  { url: "https://economictimes.indiatimes.com/topic/sebi",    label: "Economic Times — SEBI topic (news)",          mode: "news" },
  { url: "https://www.sebi.gov.in/enforcement/orders.html",    label: "SEBI — orders listing (backup)",              mode: "regulator" },
];

// ---------- LLM extract schemas — per-source mode ----------

// News-page schema: titles like "SEBI bans XYZ Ltd promoters for…"
// or "Adani gets SEBI show-cause on…". The LLM must extract the
// entity NAMED IN THE HEADLINE (not third-party advisers).
const NEWS_SCHEMA = {
  type: "object",
  properties: {
    flagged_entities: {
      type: "array",
      description: "Indian listed companies named in recent NEWS HEADLINES on this page as subjects of SEBI enforcement, investigation, or disciplinary action. INCLUDE: stories where the headline contains a company name AND any of these SEBI-action keywords: ban, bar, prohibit, fine, penalty, show-cause, adjudication, investigation, enforcement, order, settlement, consent, noticee, insider trading, fraud, misrepresent, debarment, restrain. STRICTLY EXCLUDE: stories where the company is mentioned in passing (e.g. as a market beneficiary of a SEBI policy change), opinion pieces, analyst commentary, generic 'SEBI updates rules' stories, foreign companies, and individuals named only in personal capacity. EMPTY OK: if no headline names a listed entity as the SUBJECT of an enforcement matter, return [] — do not invent.",
      items: {
        type: "object",
        properties: {
          name:             { type: "string", description: "Exact company name as it appears in the news headline." },
          order_type:       { type: "string", description: "Best fit: sebi-investigation | sebi-ban | sebi-fine | show-cause | adjudication | settlement | enforcement-order | other." },
          is_listed_entity: { type: "boolean", description: "true if the entity looks like a publicly listed Indian company." },
          headline_snippet: { type: "string", description: "VERBATIM ~15-30 word excerpt from the news headline or summary that names the company AND the SEBI action. Required." },
        },
        required: ["name", "is_listed_entity", "headline_snippet"],
      },
    },
  },
  required: ["flagged_entities"],
};

// Regulator-page schema: SEBI's orders.html lists order titles +
// occasionally entity names. Slightly different prompt to handle the
// formal language ("In the matter of XYZ Ltd…").
const REGULATOR_SCHEMA = {
  type: "object",
  properties: {
    flagged_entities: {
      type: "array",
      description: "Indian listed companies named as NOTICEES, RESPONDENTS, or subjects in SEBI enforcement order titles on this page. Order titles often follow the pattern 'In the matter of [Company]' or 'WTM Order against [Company]'. INCLUDE only if a company name is clearly part of the order title. STRICTLY EXCLUDE non-listed shell entities, foreign companies, and individuals named in personal capacity. EMPTY OK: return [] if no order title names a listed entity.",
      items: {
        type: "object",
        properties: {
          name:             { type: "string", description: "Exact entity name as it appears in the SEBI order title." },
          order_type:       { type: "string", description: "Best fit: wtm-order | ao-order | settlement | consent-order | interim-order | enforcement-order | other." },
          is_listed_entity: { type: "boolean", description: "true if the entity looks like a publicly listed Indian company." },
          headline_snippet: { type: "string", description: "VERBATIM ~15-30 word excerpt from the order title/summary that names the entity. Required." },
        },
        required: ["name", "is_listed_entity", "headline_snippet"],
      },
    },
  },
  required: ["flagged_entities"],
};

const SCHEMA_BY_MODE = { news: NEWS_SCHEMA, regulator: REGULATOR_SCHEMA };

// Hallucination guard: classic placeholder patterns.
const PLACEHOLDER_RE = /\b(XYZ|ABC|LMN|PQR|Foo|Bar|Sample|Example|Placeholder)\b/i;

// Signal verification: the SNIPPET (verbatim from the source) must
// contain a SEBI-action keyword. order_type alone is the LLM's "best
// guess" — accepting it lets through entries where the LLM tagged a
// random buyback story as "sebi-fine" with no evidence in the snippet.
const SEBI_SIGNAL_RE = /\b(sebi|show[- ]?cause|adjudicat\w+|investigat\w+|enforcement|violation|penalty|fine|ban|bar|prohibit|debarment|restrain|insider[- ]?trading|fraud|misrepresent\w+|consent\s+order|noticee|respondent|wtm\s+order|ao\s+order|interim\s+order)\b/i;

// Resolved / favourable-to-company signals. If any of these appear in
// the snippet, the SEBI proceeding has gone IN FAVOUR of the company
// (or has been resolved) — it's no longer an "active SEBI issue" and
// we should NOT hard-fail the stock.
//
// Expanded after the first ad-hoc run still flagged RELIANCE with a
// snippet that used wording my first regex didn't cover:
//   "Reliance Industries gets partial relief from Supreme Court as it
//    sets aside Rs 447 crore Sebi recovery order"
// Additions: sets/setting aside (verb conjugations); "partial relief
// from Supreme Court / High Court / SAT"; "provides relief to";
// "favourable verdict"; "appeal upheld" (already there); plus
// "in (the) favour of" without the "ruled" prefix.
const RESOLVED_SIGNAL_RE = new RegExp(
  "\\b(" +
    "quash\\w*"                                                          + "|" + // quashes, quashed, quashing
    "(?:set|sets|setting)\\s+aside"                                      + "|" + // set/sets/setting aside
    "dismiss\\w*"                                                        + "|" + // dismisses, dismissed
    "exonerat\\w*"                                                       + "|" + // exonerate(s/d)
    "acquit\\w*"                                                         + "|" + // acquit(s/ted)
    "absolv\\w*"                                                         + "|" + // absolves, absolved
    "(?:partial\\s+)?relief\\s+(?:from|by|to)\\s+(?:the\\s+)?(?:supreme\\s+court|high\\s+court|sat|sebi)" + "|" +
    "(?:provides|grants|granting)\\s+relief\\s+to"                       + "|" + // SC provides/grants relief to X
    "ruled\\s+in\\s+(?:favou?r|favor)\\s+of"                             + "|" +
    "in\\s+(?:the\\s+)?favou?r\\s+of\\s+(?:the\\s+)?(?:company|noticee|appellant|respondent)" + "|" +
    "favou?rable\\s+(?:verdict|ruling|order|judgement|judgment)"         + "|" +
    "no\\s+contravention"                                                + "|" +
    "cleared\\s+by"                                                      + "|" +
    "withdraw\\w*\\s+(?:by\\s+)?sebi"                                    + "|" +
    "appeal\\s+(?:allowed|upheld|succeeded)"                             + "|" +
    // Appellate-court language that signals SEBI/SAT got OVERTURNED.
    // E.g. RIL snippet: "the SAT, in its majority verdict, committed
    // an 'egregious error' in passing the judgement"
    "majority\\s+verdict"                                                + "|" +
    "(?:egregious|gross|fundamental|grave|patent)\\s+error"              + "|" +
    "committed\\s+an?\\s+\\w+\\s+error"                                  + "|" +
    "(?:apex\\s+court|supreme\\s+court|sat)\\s+(?:said|held|noted|observed|found)\\s+(?:the\\s+)?(?:sat|sebi)" +
  ")\\b",
  "i"
);

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const DEBUG_PER_SOURCE = [];

// ---------- Firecrawl call ----------

async function firecrawlExtract(src) {
  const { url, label, mode } = src;
  const schema = SCHEMA_BY_MODE[mode] || NEWS_SCHEMA;
  console.log(`  Firecrawl LLM extract on: ${url}   (${mode})`);
  const entry = { url, label, mode, at: new Date().toISOString() };
  try {
    const r = await fetch(FIRECRAWL_URL, {
      method: "POST",
      headers: { Authorization: `Bearer ${API_KEY}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        url,
        formats: ["json"],
        jsonOptions: { schema },
        onlyMainContent: false,
      }),
    });
    entry.fc_status = r.status;
    const text = await r.text();
    if (!r.ok) {
      entry.outcome = "fc_http_error";
      entry.error_body = text.slice(0, 300);
      console.log(`    Firecrawl HTTP ${r.status}: ${text.slice(0, 200)}`);
      DEBUG_PER_SOURCE.push(entry);
      return [];
    }
    let j;
    try { j = JSON.parse(text); }
    catch {
      entry.outcome = "non_json_response";
      entry.error_body = text.slice(0, 300);
      DEBUG_PER_SOURCE.push(entry);
      return [];
    }
    const data = j?.data || null;
    const extract = data?.json || data?.llm_extraction || data?.extract || {};
    let items = Array.isArray(extract.flagged_entities) ? extract.flagged_entities : [];
    entry.outcome = "ok";
    entry.upstream_status = data?.metadata?.statusCode || null;
    entry.total_extracted_raw = items.length;
    // Hallucination guard #1: upstream not 200 → distrust the LLM.
    if (entry.upstream_status && entry.upstream_status !== 200) {
      entry.dropped_reason = `upstream_status=${entry.upstream_status} — refusing LLM output to avoid hallucination`;
      items = [];
    }
    // Hallucination guard #2: placeholder names.
    const flagged_placeholders = items.filter((it) => PLACEHOLDER_RE.test(String(it.name || "")));
    if (flagged_placeholders.length) {
      entry.dropped_placeholders = flagged_placeholders.map((it) => it.name);
      items = items.filter((it) => !PLACEHOLDER_RE.test(String(it.name || "")));
    }
    entry.total_extracted = items.length;
    entry.prefilter_sample = items.slice(0, 15).map((it) => ({
      name: it.name,
      order_type: it.order_type,
      is_listed_entity: it.is_listed_entity,
      snippet: (it.headline_snippet || "").slice(0, 150),
    }));
    // Filter chain:
    //   (a) must have a name + must be listed entity
    //   (b) the SNIPPET (verbatim, not the LLM's order_type guess) must
    //       contain a SEBI-action keyword
    //   (c) the snippet must NOT contain a "resolved in company's
    //       favour" signal (quashed / set aside / dismissed / etc.)
    const dropped = { not_listed: 0, no_sebi_signal_in_snippet: 0, resolved_in_favour: 0, no_name: 0 };
    const filtered = items.filter((it) => {
      if (!it.name || String(it.name).trim().length < 3) { dropped.no_name++; return false; }
      if (it.is_listed_entity === false) { dropped.not_listed++; return false; }
      const snippet = String(it.headline_snippet || "");
      // (b) — REQUIRE SEBI signal in the snippet itself. order_type is
      // unreliable on its own (LLM has been seen guessing "sebi-fine"
      // for unrelated buyback / earnings stories).
      if (!SEBI_SIGNAL_RE.test(snippet)) {
        dropped.no_sebi_signal_in_snippet++; return false;
      }
      // (c) — DROP when the snippet describes a resolution in the
      // company's favour. The matter is no longer active.
      if (RESOLVED_SIGNAL_RE.test(snippet)) {
        dropped.resolved_in_favour++; return false;
      }
      return true;
    });
    entry.after_filter_count = filtered.length;
    entry.dropped_counts = dropped;
    entry.sample = filtered.slice(0, 10).map((it) => ({ name: it.name, order_type: it.order_type, snippet: (it.headline_snippet || "").slice(0, 120) }));
    console.log(`    Extracted ${items.length} entities, ${filtered.length} after filter (dropped: ${JSON.stringify(dropped)})`);
    DEBUG_PER_SOURCE.push(entry);
    return filtered;
  } catch (err) {
    entry.outcome = "fetch_error";
    entry.error = err.message;
    console.log(`    Error: ${err.message}`);
    DEBUG_PER_SOURCE.push(entry);
    return [];
  }
}

// ---------- Fuzzy name → ticker matcher ----------

function normaliseName(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/\b(ltd|limited|pvt|private|corporation|corp|inc|company|co|the|of|and|&)\b/g, " ")
    .replace(/[^\w\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function generateNameVariants(name) {
  const out = new Set();
  const n = normaliseName(name);
  if (!n) return out;
  out.add(n);                       // full normalised
  out.add(n.replace(/\s+/g, ""));   // no spaces
  const parts = n.split(/\s+/);
  if (parts.length >= 2) {
    out.add(parts.slice(0, 2).join(" "));
    out.add(parts.slice(0, 2).join(""));
  }
  if (parts.length >= 3) {
    out.add(parts.slice(0, 3).join(" "));
  }
  // First-word-only prefix for groups like "Adani", "Reliance", "Tata"
  // when news headlines abbreviate to just the group name. We add this
  // ONLY when the first word is reasonably distinctive (>= 4 chars).
  if (parts[0] && parts[0].length >= 4) {
    out.add(parts[0]);
  }
  return out;
}

function buildNameMap(screener) {
  const map = new Map();
  // Track which variants were added per name so we can deduplicate
  // sensibly (don't overwrite a precise match with a loose prefix).
  for (const c of screener) {
    const fullName = String(c.Company || "");
    const m = String(c["Screener URL"] || "").match(/\/company\/([^/]+)/);
    if (!m) continue;
    const ticker = m[1].toUpperCase();
    const variants = generateNameVariants(fullName);
    for (const v of variants) {
      // Don't overwrite an existing exact match with a vague one — e.g.,
      // "adani" alone would otherwise map to whichever Adani entity
      // ended up last in the loop. Skip if the variant is already mapped.
      if (!map.has(v)) map.set(v, ticker);
    }
  }
  return map;
}

function resolveTicker(rawName, nameMap) {
  const variants = generateNameVariants(rawName);
  for (const v of variants) {
    if (nameMap.has(v)) return nameMap.get(v);
  }
  return null;
}

// ---------- Curated baseline merger ----------

function loadCurated() {
  if (!existsSync(CURATED_PATH)) {
    console.log("  No curated baseline file at", CURATED_PATH);
    return [];
  }
  try {
    const j = JSON.parse(readFileSync(CURATED_PATH, "utf8"));
    const cases = Array.isArray(j.active_cases) ? j.active_cases : [];
    console.log(`  Loaded ${cases.length} curated active cases.`);
    return cases;
  } catch (e) {
    console.log("  Failed to parse curated file:", e.message);
    return [];
  }
}

// ---------- Output ----------

function writeStub(reason) {
  mkdirSync(dirname(OUT_PATH), { recursive: true });
  writeFileSync(OUT_PATH, JSON.stringify({
    generated_at: new Date().toISOString(),
    flagged_companies: {},
    error: reason,
  }, null, 2) + "\n");
}

// ---------- main ----------

async function main() {
  console.log("Loading Nifty 500 universe...");
  const screener = JSON.parse(readFileSync(SCREENER_PATH, "utf8"));
  const nameMap = buildNameMap(screener);
  console.log(`Built ticker map with ${nameMap.size} name variants for ${screener.length} companies.`);

  console.log("\nLoading curated baseline...");
  const curated = loadCurated();

  // ---- Pull from live sources (if API key present) ----
  const allEntities = [];
  if (!API_KEY) {
    console.log("\nFIRECRAWL_API_KEY not set — skipping live sources, using curated baseline only.");
  } else {
    console.log("\nScraping live sources...");
    for (const src of SOURCES) {
      const items = await firecrawlExtract(src);
      for (const it of items) allEntities.push({
        ...it,
        _source: src.label,
        _url: src.url,
        _origin: "scraper",
      });
      await sleep(400);
    }
    console.log(`\nTotal live entities collected: ${allEntities.length}`);
  }

  // ---- Merge curated cases into the entity stream ----
  for (const c of curated) {
    allEntities.push({
      name: c.name,
      order_type: c.order_type,
      headline_snippet: c.snippet,
      is_listed_entity: true,
      _source: c.source_label || "Curated baseline (static/known-sebi-cases.json)",
      _url: null,
      _origin: "curated",
      _explicit_ticker: c.ticker, // honour the curator's ticker mapping
    });
  }

  // ---- Match to Nifty 500 tickers ----
  const flagged = {};
  const unmatched = [];
  for (const e of allEntities) {
    // Curated entries carry an explicit ticker — trust the curator.
    let ticker = e._explicit_ticker || null;
    if (!ticker) ticker = resolveTicker(e.name, nameMap);
    if (!ticker) {
      unmatched.push({ name: e.name, type: e.order_type, source: e._source, origin: e._origin });
      continue;
    }
    if (!flagged[ticker]) {
      flagged[ticker] = {
        primary_name: e.name,
        order_type: e.order_type || null,
        snippet: e.headline_snippet || null,
        source_url: e._url,
        source_label: e._source,
        origins: new Set(),
        mentions: [],
        first_seen_at: new Date().toISOString(),
      };
    }
    flagged[ticker].origins.add(e._origin);
    flagged[ticker].mentions.push({
      source: e._source,
      origin: e._origin,
      order_type: e.order_type || null,
      snippet: (e.headline_snippet || "").slice(0, 200),
    });
  }
  // Convert origins Set → sorted array for JSON
  for (const t in flagged) {
    flagged[t].origins = [...flagged[t].origins].sort();
  }

  console.log(`\nMatched ${Object.keys(flagged).length} flagged tickers from ${allEntities.length} entities (${unmatched.length} unmatched).`);
  if (Object.keys(flagged).length) {
    console.log("Flagged tickers:");
    Object.entries(flagged).forEach(([t, v]) => {
      console.log(`  ${t.padEnd(15)} — ${v.primary_name}  [${v.origins.join("+")}]  (${v.order_type || "unknown type"})`);
    });
  }

  const payload = {
    generated_at: new Date().toISOString(),
    source_count: SOURCES.length,
    curated_baseline_count: curated.length,
    total_entities_seen: allEntities.length,
    flagged_companies: flagged,
    unmatched_sample: unmatched.slice(0, 30),
    sources: DEBUG_PER_SOURCE,
  };

  mkdirSync(dirname(OUT_PATH), { recursive: true });
  writeFileSync(OUT_PATH, JSON.stringify(payload, null, 2) + "\n");
  console.log(`\nWrote ${OUT_PATH}`);
}

main().catch((err) => {
  console.error("Fatal:", err.stack || err.message);
  writeStub(err.message);
  process.exit(0);
});
