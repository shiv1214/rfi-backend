import express from "express";
import fetch from "node-fetch";
import cors from "cors";
import dotenv from "dotenv";
import multer from "multer";
import XLSX from "xlsx";
import fs from "fs";
import { MongoClient } from "mongodb";

dotenv.config();
const app = express();

// ── MONGODB CONNECTION ────────────────────────────────────────────
let db = null;

async function connectDB() {
  if (!process.env.MONGODB_URI) {
    console.log("⚠ No MONGODB_URI — using in-memory storage only");
    return;
  }
  try {
    const client = new MongoClient(process.env.MONGODB_URI);
    await client.connect();
    db = client.db("kavach");
    console.log("✅ MongoDB connected — kavach database");
  } catch(e) {
    console.error("❌ MongoDB connection failed:", e.message);
    console.log("⚠ Falling back to in-memory storage");
  }
}
connectDB();
app.use(cors({
  origin: process.env.FRONTEND_URL || "*",
  methods: ["GET", "POST"],
  allowedHeaders: ["Content-Type", "Authorization"]
}));
app.use(express.json());
const upload = multer({ dest: "uploads/" });

let RFI_DATA = null;
let SUPPLIER_DATA = null;
let SUPPLIERS_REGISTRY = {};  // in-memory fallback

// ── PARSERS ──────────────────────────────────────────────────────

function parseRFI(wb) {
  const r = {
    project: {},
    must_have: [],
    good_to_have: [],
    subjective: []
  };
  if (wb.SheetNames.includes("Project Details")) {
    XLSX.utils.sheet_to_json(wb.Sheets["Project Details"], { defval: "" }).forEach(row => {
      const q = String(row["Question"] || row[Object.keys(row)[1]] || "").trim();
      const a = String(row["Answer"]   || row[Object.keys(row)[2]] || "").trim();
      if (q && a) r.project[q] = a;
    });
  }
  if (wb.SheetNames.includes("RFI Requirements")) {
    XLSX.utils.sheet_to_json(wb.Sheets["RFI Requirements"], { defval: "" }).forEach(row => {
      const param = String(row["Parameter"] || "").trim();
      const cat   = String(row["Category of Reqirement"] || row["Category"] || "").toLowerCase();
      const req   = String(row["Project"] || row["Requirement"] || row["Value"] || "").trim();
      if (!param) return;
      const entry = { parameter: param, requirement: req };
      if      (cat.includes("must"))       r.must_have.push(entry);
      else if (cat.includes("good"))       r.good_to_have.push(entry);
      else if (cat.includes("subjective")) r.subjective.push(entry);
    });
  }
  return r;
}

function parseSupplier(wb) {
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows  = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  if (!rows.length) return { variants: [] };
  const keys     = Object.keys(rows[0]);
  const paramCol = keys.find(k => String(k).toLowerCase().includes("param") || k === "Parameter") || keys[1];
  const vCols    = keys.slice(keys.indexOf(paramCol) + 1).filter(k => !k.startsWith("__") || rows[1]?.[k]);
  const vNames   = vCols.map(c => c.startsWith("__") ? String(rows[0]?.[c] || c) : c).filter(n => n && n !== "Value");
  const variants = vNames.map(name => ({ name, specs: {} }));
  rows.forEach(row => {
    const param = String(row[paramCol] || "").trim();
    if (!param || param === "Parameter") return;
    vNames.forEach((name, i) => {
      const val = row[vCols[i]];
      if (val !== undefined && String(val).trim()) variants[i].specs[param] = String(val).trim();
    });
  });
  return { variants: variants.filter(v => Object.keys(v.specs).length > 0) };
}

// ── UPLOAD ────────────────────────────────────────────────────────

app.post("/upload/rfi", upload.single("file"), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No file" });
    RFI_DATA = parseRFI(XLSX.readFile(req.file.path));
    fs.unlinkSync(req.file.path);
    console.log("RFI loaded:", RFI_DATA.must_have.length, "MH,", RFI_DATA.good_to_have.length, "GTH");
    res.json({ must_have_count: RFI_DATA.must_have.length, good_to_have_count: RFI_DATA.good_to_have.length, subjective_count: RFI_DATA.subjective.length });
  } catch(e) { console.error(e); res.status(500).json({ error: e.message }); }
});

app.post("/upload/supplier", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No file" });
    const supplierName = req.body?.supplierName || `Supplier_${Date.now()}`;
    const parsed = parseSupplier(XLSX.readFile(req.file.path));
    fs.unlinkSync(req.file.path);

    // Set as active supplier in memory
    SUPPLIER_DATA = parsed;
    SUPPLIERS_REGISTRY[supplierName] = {
      data: parsed,
      variants: parsed.variants.map(v => v.name),
      uploadedAt: new Date().toISOString()
    };

    // Save to MongoDB if connected
    if (db) {
      try {
        await db.collection("suppliers").updateOne(
          { name: supplierName },
          { $set: {
              name: supplierName,
              variants: parsed.variants,
              variantNames: parsed.variants.map(v => v.name),
              uploadedAt: new Date()
          }},
          { upsert: true }
        );
        console.log(`✅ Supplier "${supplierName}" saved to MongoDB`);
      } catch(e) {
        console.error("MongoDB save error:", e.message);
      }
    }

    console.log(`Supplier "${supplierName}" loaded:`, parsed.variants.map(v => v.name).join(", "));
    res.json({ variants_found: parsed.variants.map(v => v.name), supplierName });
  } catch(e) { console.error(e); res.status(500).json({ error: e.message }); }
});

// Get all suppliers — from MongoDB if available, else in-memory
app.get("/suppliers", async (_, res) => {
  try {
    if (db) {
      const suppliers = await db.collection("suppliers")
        .find({}, { projection: { name:1, variantNames:1, uploadedAt:1, _id:0 } })
        .sort({ uploadedAt: -1 })
        .toArray();
      return res.json({ suppliers: suppliers.map(s => ({
        name: s.name,
        variants: s.variantNames || [],
        uploadedAt: s.uploadedAt
      }))});
    }
    // Fallback to in-memory
    const suppliers = Object.entries(SUPPLIERS_REGISTRY).map(([name, s]) => ({
      name, variants: s.variants, uploadedAt: s.uploadedAt
    }));
    res.json({ suppliers });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// OEM selects a supplier — load from MongoDB or memory
app.post("/select/supplier", async (req, res) => {
  const { supplierName } = req.body || {};
  if (!supplierName) return res.status(400).json({ error: "supplierName required" });

  try {
    // Try MongoDB first
    if (db) {
      const sup = await db.collection("suppliers").findOne({ name: supplierName });
      if (sup) {
        SUPPLIER_DATA = { variants: sup.variants };
        console.log(`Active supplier set to: ${supplierName} (from MongoDB)`);
        return res.json({ ok: true, supplierName, variants: sup.variants.map(v => v.name) });
      }
    }

    // Fallback to in-memory registry
    if (SUPPLIERS_REGISTRY[supplierName]) {
      SUPPLIER_DATA = SUPPLIERS_REGISTRY[supplierName].data;
      console.log(`Active supplier set to: ${supplierName} (from memory)`);
      return res.json({ ok: true, supplierName, variants: SUPPLIER_DATA.variants.map(v => v.name) });
    }

    return res.status(404).json({ error: "Supplier not found" });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("/data/rfi",      (_, res) => res.json(RFI_DATA || {}));
app.get("/data/supplier", (_, res) => res.json(SUPPLIER_DATA || {}));

// ── DETERMINISTIC EVALUATOR ──────────────────────────────────────
// Compares catalogue spec against RFI requirement purely from data.
// No hardcoded parameter names, no hardcoded values.

function nums(s) {
  return (String(s).match(/\d+\.?\d*/g) || []).map(Number).filter(n => !isNaN(n) && n > 0);
}

function isRange(s) {
  return /\d+\s*(to|-|–)\s*\d+/i.test(s) || nums(s).length >= 2;
}

function findSpec(variant, reqParam) {
  const clean = s => String(s).toLowerCase().replace(/[^a-z0-9]/g, "");
  const rp = clean(reqParam);
  // 1. exact normalised match
  let key = Object.keys(variant.specs).find(k => clean(k) === rp);
  if (key) return { key, val: variant.specs[key] };
  // 2. one contains the other
  key = Object.keys(variant.specs).find(k => clean(k).includes(rp) || rp.includes(clean(k)));
  if (key) return { key, val: variant.specs[key] };
  // 3. token overlap — at least one 3-char token matches
  const reqTok = rp.match(/[a-z0-9]{3,}/g) || [];
  key = Object.keys(variant.specs).find(k => {
    const kt = clean(k).match(/[a-z0-9]{3,}/g) || [];
    return reqTok.some(t => kt.some(kt2 => kt2.includes(t) || t.includes(kt2)));
  });
  if (key) return { key, val: variant.specs[key] };
  return null;
}

// ── NUMERIC evaluator (deterministic — math only) ─────────────────
// Used for range, min, max, single-number requirements
function evaluateNumeric(variantName, offered, reqVal) {
  const offStr  = String(offered).toLowerCase().trim();
  const reqStr  = String(reqVal).toLowerCase().trim();
  const offNums = nums(offStr);
  const reqNums = nums(reqStr);

  // RANGE: subset or superset both pass
  if (isRange(reqStr) && reqNums.length >= 2) {
    const [reqLo, reqHi] = [Math.min(...reqNums), Math.max(...reqNums)];
    if (offNums.length >= 2) {
      const [offLo, offHi] = [Math.min(...offNums), Math.max(...offNums)];
      const pass = (offLo >= reqLo && offHi <= reqHi) || (offLo <= reqLo && offHi >= reqHi);
      return {
        numeric: true, pass, offered,
        reason: pass
          ? `${variantName} range ${offered} is within/covers required range ${reqVal}`
          : `${variantName} range ${offered} (${offLo}–${offHi}) is outside acceptable window ${reqVal} (${reqLo}–${reqHi})`
      };
    }
    if (offNums.length === 1) {
      const [reqLo, reqHi] = [Math.min(...reqNums), Math.max(...reqNums)];
      const pass = offNums[0] >= reqLo && offNums[0] <= reqHi;
      return { numeric: true, pass, offered, reason: pass ? `${variantName} ${offered} within range ${reqVal}` : `${variantName} ${offered} outside range ${reqVal}` };
    }
  }

  // MINIMUM
  if (/\bmin(imum)?\b|>=|at least/i.test(reqStr) && reqNums.length === 1 && offNums.length >= 1) {
    const pass = Math.max(...offNums) >= reqNums[0];
    return { numeric: true, pass, offered, reason: pass ? `${variantName} ${offered} meets minimum ${reqNums[0]}` : `${variantName} ${offered} below minimum ${reqNums[0]}` };
  }

  // MAXIMUM
  if (/\bmax(imum)?\b|<=|not more|up to/i.test(reqStr) && reqNums.length === 1 && offNums.length >= 1) {
    const pass = Math.min(...offNums) <= reqNums[0];
    return { numeric: true, pass, offered, reason: pass ? `${variantName} ${offered} within maximum ${reqNums[0]}` : `${variantName} ${offered} exceeds maximum ${reqNums[0]}` };
  }

  // SINGLE NUMBER (5% tolerance)
  if (reqNums.length === 1 && offNums.length >= 1) {
    const tol  = reqNums[0] * 0.05;
    const pass = Math.abs(offNums[0] - reqNums[0]) <= Math.max(tol, 0.5);
    return { numeric: true, pass, offered, reason: pass ? `${variantName} ${offered} matches ${reqVal}` : `${variantName} ${offered} does not match ${reqVal}` };
  }

  return null; // Not a numeric comparison — fall through to LLM
}

// ── Main evaluate — numeric first, LLM for text/semantic ──────────
// Returns a promise for text params (needs LLM), sync result for numeric
async function evaluate(variantName, reqParam, reqVal) {
  const variant = SUPPLIER_DATA.variants.find(v => v.name.toLowerCase() === variantName.toLowerCase());
  if (!variant) return { pass: false, offered: "N/A", reason: "Variant not found", method: "error" };

  const found = findSpec(variant, reqParam);
  if (!found) return { pass: false, offered: "Not in catalogue", reason: `"${reqParam}" not found in ${variantName} catalogue`, method: "not_found" };

  const offered = found.val;
  const reqStr  = String(reqVal).toLowerCase().trim();
  const reqNums = nums(reqStr);

  // If requirement has numbers → use deterministic numeric evaluator
  if (reqNums.length > 0) {
    const numResult = evaluateNumeric(variantName, offered, reqVal);
    if (numResult) return { ...numResult, method: "numeric" };
  }

  // No numbers, or numeric evaluator couldn't handle it → use LLM for semantic reasoning
  // This handles: "eAxle" vs "Rear Wheel Central Drive", "Water/Liquid" vs "Water-Glycol", "PMSM" vs "Induction", etc.
  return await evaluateWithLLM(variantName, reqParam, reqVal, offered);
}

// ── LLM semantic evaluator ─────────────────────────────────────────
// Uses GPT-4o-mini to reason about equivalence, not string matching
async function evaluateWithLLM(variantName, reqParam, reqVal, offered) {
  const key   = process.env.OPENAI_API_KEY;
  const model = process.env.AI_MODEL || "gpt-4o-mini";

  const prompt = `You are a technical procurement engineer evaluating supplier compliance.

PARAMETER BEING EVALUATED: ${reqParam}
REQUIRED BY TML: ${reqVal}
OFFERED BY SUPPLIER (${variantName}): ${offered}

Your job: Does the offered value satisfy the requirement? Use engineering knowledge and common sense.

Examples of correct reasoning:
- Required "Water/Liquid cooling", offered "Water-Glycol (50-50)" → PASS (glycol is a liquid, water-glycol IS liquid cooling)
- Required "eAxle", offered "Rear Wheel Central Drive" → FAIL (completely different architecture)
- Required "PMSM", offered "Permanent Magnet Synchronous Motor" → PASS (same thing, different abbreviation)
- Required "CAN 2.0B", offered "CAN FD" → PASS (CAN FD is backward compatible with CAN 2.0B)
- Required "IP6K9K", offered "IP67" → FAIL (IP67 is not the same as IP6K9K — automotive standard differs)
- Required "ASIL-C", offered "ASIL-B" → FAIL (lower safety level)

Respond with ONLY a JSON object, nothing else:
{"pass": true or false, "reason": "one sentence explanation"}`;

  try {
    const r = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: { "Authorization": `Bearer ${key}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        model,
        messages: [{ role: "user", content: prompt }],
        temperature: 0,
        max_tokens: 100
      })
    });
    const data = await r.json();
    const text = data.choices?.[0]?.message?.content?.trim() || "{}";
    const parsed = JSON.parse(text.replace(/```json|```/g, "").trim());
    return {
      pass: !!parsed.pass,
      offered,
      reason: parsed.reason || `LLM evaluated: ${offered} vs ${reqVal}`,
      method: "llm"
    };
  } catch(e) {
    console.error("LLM evaluator error:", e.message);
    // Fallback: basic token match if LLM fails
    const offTok = offered.toLowerCase().replace(/[^a-z0-9 ]/g," ").split(/\s+/).filter(w=>w.length>=3);
    const reqTok = reqVal.toLowerCase().replace(/[^a-z0-9 ]/g," ").split(/\s+/).filter(w=>w.length>=3);
    const matched = reqTok.filter(t => offTok.some(ot => ot.includes(t) || t.includes(ot)));
    const pass = matched.length / Math.max(reqTok.length, 1) >= 0.5;
    return { pass, offered, reason: `Token match fallback: ${matched.length}/${reqTok.length} terms matched`, method: "fallback" };
  }
}

// Evaluate ALL variants against ALL must-have requirements at once
// Uses async LLM evaluation for text params, deterministic for numeric
app.post("/evaluate/all_must_have", async (_, res) => {
  try {
    if (!RFI_DATA || !SUPPLIER_DATA) return res.status(400).json({ error: "Upload both files first" });
    const results = {};

    // Run all evaluations in parallel per variant
    await Promise.all(SUPPLIER_DATA.variants.map(async v => {
      results[v.name] = await Promise.all(RFI_DATA.must_have.map(async req => {
        const evalResult = await evaluate(v.name, req.parameter, req.requirement);
        return { parameter: req.parameter, requirement: req.requirement, ...evalResult };
      }));
    }));

    console.log("Hybrid evaluation complete:");
    Object.entries(results).forEach(([vn, reqs]) => {
      reqs.forEach(r => {
        const icon = r.pass ? "✓" : "✗";
        console.log(`  ${icon} ${vn} — ${r.parameter}: ${r.offered} [${r.method}]`);
      });
    });
    res.json(results);
  } catch(e) { console.error(e); res.status(500).json({ error: e.message }); }
});

// ── PROMPTS ───────────────────────────────────────────────────────

function buildTMLPrompt(phase, reqParam, reqVal) {
  const proj = Object.entries(RFI_DATA.project).map(([k,v]) => `${k}: ${v}`).join("\n");
  const mh   = RFI_DATA.must_have.map((r,i) => `${i+1}. ${r.parameter}: ${r.requirement}`).join("\n");
  const gth  = RFI_DATA.good_to_have.map((r,i) => `${i+1}. ${r.parameter}: ${r.requirement}`).join("\n");
  const subj = RFI_DATA.subjective?.map((r,i) => `${i+1}. ${r.parameter}: ${r.requirement}`).join("\n") || "";

  let task = "";

  if (phase === "general") {
    task = `PHASE 1 — You are ANSWERING the supplier's question, not asking one.
Read their last message carefully. Answer it directly and naturally using the project data below.
Sound like a TML engineer in a meeting — confident, precise, factual.
1-2 sentences max. Do NOT ask anything back.`;
  }
  else if (phase === "must_have") {
    task = `PHASE 2 — MUST HAVE VERIFICATION
Ask about this ONE specific requirement only:
  Parameter: ${reqParam}
  Required value: ${reqVal}

Ask naturally and professionally. Example:
"Regarding ${reqParam} — our specification requires ${reqVal}. Can you confirm the exact value your variant offers?"
One question only. Do not mention any other requirement.`;
  }
  else if (phase === "good_to_have") {
    task = `PHASE 3 — GOOD TO HAVE PREFERENCE
Ask about this preference conversationally:
  Parameter: ${reqParam}
  Preference: ${reqVal}
Example: "On ${reqParam}, we prefer ${reqVal}. What does your variant offer here?"
One question, friendly but professional.`;
  }
  else if (phase === "negotiation") {
    task = `PHASE 4 — NEGOTIATION
Review the conversation. Identify the key deviation or gap found.
Ask the supplier specifically: can they close this gap through engineering?
Ask for: feasibility, timeline, and cost impact. Be direct and solution-focused. 2 sentences.`;
  }
  else if (phase === "recommendation") {
    task = `PHASE 5 — FINAL SUMMARY REQUEST
Professionally close the evaluation. Ask the supplier to provide:
1. All modifications they can commit to
2. Timeline for each
3. Technical risks TML should consider
Sound like a senior engineer wrapping up a formal technical review. One clear request.`;
  }

  return `You are a Senior Design Engineer at Tata Motors Limited evaluating suppliers for the Columbus eAxle project.
You have ONLY the RFI document below. You cannot see the supplier catalogue.

PERSONALITY: Professional, direct, technically sharp. You ask precise questions. You acknowledge good answers briefly. You never repeat yourself.

PROJECT DETAILS:
${proj}

MUST HAVE REQUIREMENTS:
${mh}

GOOD TO HAVE PREFERENCES:
${gth}

SUBJECTIVE SPECIFICATIONS:
${subj}

OUTPUT RULES:
- Output ONLY your message. No labels, no preamble like "Sure" or "Certainly".
- Never say "as per RFI" or "my document says".
- Sound like a real engineer in a professional meeting.
- Stay strictly on the current task.

CURRENT TASK:
${task}`;
}

function buildSupplierPrompt(variantName, phase, reqParam, reqVal) {
  const v = SUPPLIER_DATA.variants.find(x => x.name.toLowerCase() === variantName.toLowerCase()) || SUPPLIER_DATA.variants[0];
  if (!v) return "You are a Supplier Engineer. No catalogue data available.";
  const specs = Object.entries(v.specs).map(([k, val]) => `  ${k}: ${val}`).join("\n");
  const allVariants = SUPPLIER_DATA.variants.map(x => x.name).join(", ");

  let task = "";

  if (phase === "general") {
    task = `PHASE 1 — PROJECT DISCOVERY (You are ASKING a question)
Ask TML ONE specific professional question to understand their project requirements.
Good topics: vehicle type, SOP date, production volumes, target market, warranty expectations, quality standards.
Do NOT reveal your specs yet. Do NOT repeat questions already asked in conversation.
One concise, professional sentence only.`;
  }
  else if (phase === "must_have" || phase === "must_have_eval") {
    task = `PHASE 2 — TECHNICAL COMPLIANCE
TML asked about your ${v.name} value for: "${reqParam}" (they require: ${reqVal})

Steps:
1. Find "${reqParam}" in your ${v.name} specs below
2. State your exact value
3. Say whether it meets the requirement

Response format: "Our ${v.name} [meets/does not meet] this requirement — our ${reqParam} is [exact value from specs]."

Rules:
- Use ONLY values from your specs. Never guess.
- If spec not listed: "That parameter is not in our ${v.name} catalogue."
- If your value is within an acceptable range, state it confidently.
- Be honest if it doesn't meet requirements.`;
  }
  else if (phase === "good_to_have") {
    task = `PHASE 3 — PREFERENCE CHECK
TML asked about ${reqParam} (their preference: ${reqVal}).
Find this in your ${v.name} specs and respond naturally.
"For ${reqParam}, our ${v.name} offers [your value], which [aligns with/differs from] your preference of ${reqVal}."
If not in catalogue: "That parameter is not listed in our ${v.name} catalogue."
One sentence, honest and professional.`;
  }
  else if (phase === "negotiation") {
    task = `PHASE 4 — ENGINEERING NEGOTIATION
TML is asking about closing a technical gap through product modification.
Respond as a professional application engineer with engineering authority.
If feasible: state specifically what can be changed, realistic timeline (e.g. "10-14 weeks"), constraints.
If not feasible: explain why clearly, suggest nearest alternative.
2-3 sentences. Realistic, professional. Do not over-promise.`;
  }
  else if (phase === "recommendation") {
    task = `PHASE 5 — FINAL ENGINEERING SUMMARY
Provide a structured response:
1. Specific modifications you can commit to
2. Timeline for each
3. Technical risks or constraints TML should factor in
Be specific and realistic. Sound like a senior engineer closing a formal proposal. 4-6 sentences.`;
  }

  return `You are the Application Engineer for a supplier presenting the ${v.name} eAxle variant to Tata Motors.
You have ONLY your product catalogue. You do NOT have TML's RFI document.
Available variants: ${allVariants}

PERSONALITY: Technically knowledgeable, honest, professional. You know your product well. You admit gaps honestly rather than bluffing.

${v.name.toUpperCase()} — COMPLETE SPECIFICATIONS:
${specs}

OUTPUT RULES:
- Output ONLY your message. No labels, no preamble like "Certainly" or "Of course".
- NEVER say "I do not have data" for specs that ARE listed above — look carefully first.
- If genuinely not listed: "That parameter is not in our ${v.name} catalogue."
- Sound like a real engineer in a professional meeting.

CURRENT TASK:
${task}`;
}

// ── SCRIPTED TML ANSWERS FOR PHASE 1 (no AI, no hallucination) ───

function scriptedAnswer(supplierQuestion) {
  const q = supplierQuestion.toLowerCase();
  const p = RFI_DATA.project;

  // Helper: find value from project details by keyword in key
  const get = (...keys) => {
    for (const k of keys) {
      const found = Object.entries(p).find(([key]) => key.toLowerCase().includes(k));
      if (found) return found[1];
    }
    return null;
  };

  // All project details as a readable string for generic fallback
  const allDetails = Object.entries(p).map(([k,v]) => `${k}: ${v}`).join("\n");

  // ── Specific keyword matches — most specific first ──────────────
  if (q.includes("start of production") || q.includes(" sop ") || q.match(/\bsop\b/))
    return `The Start of Production (SOP) for Project Columbus is planned for ${get("sop","start of production","production date") || "October 2026"}.`;

  if (q.includes("annual volume") || q.includes("production volume") || q.includes("annual production") || q.includes("quantity per year") || (q.includes("volume") && !q.includes("environment")))
    return `First year volume requirement: ${get("first year","volume y1","year 1 volume") || "15,000 units"}. Second year: ${get("second year","volume y2","year 2 volume") || "25,000 units"}.`;

  if (q.includes("warranty") || q.includes("durability standard") || q.includes("durability requirement"))
    return `Part level warranty expectation: ${get("warranty","part level warranty","durability") || "10 years or 5 lac KM whichever is earlier"}.`;

  if (q.includes("target market") || q.includes("geographical") || q.includes("geography") || q.includes("which market") || q.includes("which country") || q.includes("region"))
    return `The target market for Project Columbus is ${get("target market","market","geography") || "India and SARC Countries"}.`;

  if (q.includes("vehicle type") || q.includes("vehicle category") || q.includes("gvw") || q.includes("gross vehicle") || q.includes("what type of vehicle") || q.includes("which vehicle"))
    return `The vehicle type is ${get("vehicle type","vehicle") || "N3 Heavy Commercial Vehicle with 46T GVW"}.`;

  if (q.includes("localisation") || q.includes("localization") || q.includes("pli") || q.includes("local content") || q.includes("make in india"))
    return `${get("localisation","pli","local") || "Yes, localisation compliance is required as per PLI/PM-E drive guidelines"}.`;

  if (q.includes("manufacturing") || q.includes("plant location") || q.includes("production site") || q.includes("where will") || q.includes("facility"))
    return `Manufacturing location: ${get("manufacturing","location","plant","facility") || "Jamshedpur, India"}.`;

  if (q.includes("eaxle") || q.includes("e-axle") || q.includes("component") || q.includes("what are you sourcing") || q.includes("what is being sourced"))
    return `We are sourcing a complete eAxle Drive System for integration into a heavy commercial vehicle platform.`;

  if (q.includes("alpha build") || q.includes("beta build") || q.includes("prototype") || q.includes("build plan") || q.includes("sample"))
    return `Alpha build: ${get("alpha build","alpha") || "June 2025"} (${get("alpha qty","alpha parts") || "6 units"}). Beta build: ${get("beta build","beta") || "December 2025"} (${get("beta qty","beta parts") || "15 units"}).`;

  if (q.includes("project name") || q.includes("name of the project"))
    return `The project name is ${get("project name","project") || "Columbus"}.`;

  // ── Generic fallback: send relevant project details ─────────────
  // For questions about "performance", "application", "environment", "challenges", etc.
  // Return a summary of the key project facts rather than one repeated answer
  return `Here are the key project details: ${allDetails}. Please let me know if you need any specific clarification.`;
}

// ── AI ROUTE ──────────────────────────────────────────────────────

app.post("/api/ai", async (req, res) => {
  try {
    const { messages, agent, phase, variant, reqParam, reqVal } = req.body;
    if (!messages || !agent) return res.status(400).json({ error: "Missing messages or agent" });
    if (!RFI_DATA || !SUPPLIER_DATA) return res.status(400).json({ error: "Upload both files first" });

    // Phase 1 TML: scripted, zero hallucination
    if (agent === "tml" && phase === "general") {
      const lastSup = [...messages].reverse().find(m => m.role === "assistant")?.content || "";
      return res.json({ choices: [{ message: { content: scriptedAnswer(lastSup) } }] });
    }

    const prompt = agent === "tml"
      ? buildTMLPrompt(phase, reqParam, reqVal)
      : buildSupplierPrompt(variant, phase, reqParam, reqVal);

    const key   = process.env.OPENAI_API_KEY;
    const model = process.env.AI_MODEL || "gpt-4o-mini";

    const r = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: { "Authorization": `Bearer ${key}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        model,
        messages: [{ role: "system", content: prompt }, ...messages],
        temperature: 0.1,
        max_tokens: 250
      })
    });

    const data = await r.json();
    if (data.error) { console.error("OpenAI error:", data.error); return res.status(500).json({ error: data.error.message }); }
    res.json(data);
  } catch(e) { console.error(e); res.status(500).json({ error: e.message }); }
});

// ── REPORT ────────────────────────────────────────────────────────

app.post("/api/report", async (req, res) => {
  try {
    const { variantResults, conversation } = req.body;
    if (!RFI_DATA || !SUPPLIER_DATA) return res.status(400).json({ error: "No data" });

    const mh  = RFI_DATA.must_have.map(r => `${r.parameter}: ${r.requirement}`).join("\n");
    const gth = RFI_DATA.good_to_have.map(r => `${r.parameter}: ${r.requirement}`).join("\n");
    const varSum = (variantResults||[]).map(v =>
      `${v.name}: ${v.status}, MH ${v.mhPassed}/${v.mhTotal}, deviations: ${(v.deviations||[]).join("; ")||"none"}, eliminated: ${v.eliminationReason||"N/A"}`
    ).join("\n");

    const prompt = `TML procurement analyst. Generate RFI compliance report for Project Columbus eAxle evaluation as JSON only (no markdown).\n\nMUST HAVE:\n${mh}\n\nGOOD TO HAVE:\n${gth}\n\nVARIANT RESULTS:\n${varSum}\n\nReturn:\n{"executive_summary":"...","verdict":"RECOMMENDED or CONDITIONALLY RECOMMENDED or NOT RECOMMENDED","recommended_variant":"name or null","recommendation_reason":"...","variant_results":[{"variant":"","status":"PASSED or ELIMINATED","elimination_reason":"","overall_fit":"","deviations":[]}],"modifications_required":[{"item":"","timeline":"","feasibility":""}],"risks":[{"risk":"","severity":"","mitigation":""}]}`;

    const key   = process.env.OPENAI_API_KEY;
    const model = process.env.AI_MODEL || "gpt-4o-mini";
    const r = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: { "Authorization": `Bearer ${key}`, "Content-Type": "application/json" },
      body: JSON.stringify({ model, messages: [{ role: "user", content: prompt }], temperature: 0.1, max_tokens: 1500 })
    });
    const data = await r.json();
    const txt  = data.choices?.[0]?.message?.content || "{}";
    try { res.json({ report: JSON.parse(txt.replace(/```json|```/g, "").trim()) }); }
    catch { res.json({ report: null, raw: txt }); }
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── HEALTH ────────────────────────────────────────────────────────

app.get("/health", (_, res) => res.json({
  ok: true,
  rfi: !!RFI_DATA,
  supplier: !!SUPPLIER_DATA,
  model: process.env.AI_MODEL || "gpt-4o-mini",
  variants: SUPPLIER_DATA?.variants?.map(v => v.name),
  routes: ["/upload/rfi", "/upload/supplier", "/evaluate/all_must_have", "/api/ai", "/api/report"]
}));

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log(`Server running at http://localhost:${PORT}\nHealth: http://localhost:${PORT}/health`));