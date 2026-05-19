/**
 * ============================================================================
 *  Prizzz Scorecard — free backend (Google Apps Script)
 *  Stores every submission in a Google Sheet (your V1 "CRM") and emails
 *  the lead + a notification to Banj Media. $0, no Supabase / Airtable / Vercel.
 * ============================================================================
 *
 *  ONE-TIME SETUP (≈5 min)
 *  -----------------------
 *  1. Create a new Google Sheet (e.g. "Prizzz Leads"). Leave it empty —
 *     the header row is created automatically on first submission.
 *  2. In that Sheet: Extensions ▸ Apps Script. Delete the sample code,
 *     paste THIS ENTIRE FILE, and save.
 *  3. Edit NOTIFY_EMAIL below to the inbox that should get hot-lead alerts.
 *  4. Deploy ▸ New deployment ▸ type "Web app".
 *       - Description: prizzz-scorecard
 *       - Execute as:  Me
 *       - Who has access: Anyone
 *     Click Deploy, authorize when prompted, COPY the Web app URL.
 *  5. Paste that URL into score/index.html →  CONFIG.ENDPOINT_URL.
 *  6. (After any code change here) Deploy ▸ Manage deployments ▸ edit ▸
 *     "New version" so the live URL runs the latest code.
 *
 *  Rows are UPSERTED by session_id: the ungated completion snapshot creates
 *  the row; if the user later leaves their email, the same row is enriched
 *  (no duplicates). Emails are sent only when an email is present.
 * ============================================================================
 */

const NOTIFY_EMAIL = "contact@banjmedia.com"; // sales inbox (lead alerts + reply-to)
const SHEET_NAME    = "Leads";

const HEADERS = [
  "session_id","submitted_at","lead_category","lead_quality_score",
  "name","company","role","email","phone",
  "instagram","facebook","x","tiktok","linkedin","website",
  "sector","location","platforms","moments","moment_date","moment_importance",
  "production_need","budget","commitment","decision","start_timing","lead_source",
  "utm_source","utm_medium","utm_campaign","utm_content","referrer",
  "overall","strategy","production","platform","packaging","distribution","moment",
  "result_category","recommended_plan","activation","budget_constrained",
  "strongest","weakest","answers_json","status","next_action"
];

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents || "{}");
    const sheet = getSheet_();
    const row = toRow_(p);

    // Upsert by session_id (column A)
    const ids = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 1).getValues();
    let target = -1;
    for (let i = 0; i < ids.length; i++) if (ids[i][0] === p.session_id) { target = i + 2; break; }
    if (target === -1) sheet.appendRow(row);
    else sheet.getRange(target, 1, 1, row.length).setValues([row]);

    const email = (p.contact && p.contact.email) ? String(p.contact.email).trim() : "";
    if (email) {
      sendUserEmail_(p, email);
      notifyTeam_(p);
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet() { return json_({ ok: true, service: "prizzz-scorecard" }); }

/* ---------- helpers ---------- */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
  }
  return sh;
}
function toRow_(p) {
  const c = p.contact || {}, q = p.qualification || {}, u = p.utm || {},
        s = p.scores || {}, d = (s.dimensions || {}), r = p.recommendation || {},
        l = p.lead || {}, cl = c.links || {};
  return [
    p.session_id || "", p.submitted_at || new Date().toISOString(),
    l.category || "", l.quality_score == null ? "" : l.quality_score,
    c.name || "", c.company || "", c.role || q.role || "", c.email || "", c.phone || "",
    cl.instagram || "", cl.facebook || "", cl.x || "", cl.tiktok || "", cl.linkedin || "", cl.website || "",
    q.sector || "", q.loc || "", (q.platforms || []).join(", "), (q.moments || []).join(", "),
    q.mdate || "", q.mimp || "", q.prod || "", q.budget || "", q.commit || "",
    q.decide || "", q.timing || "", q.source || "",
    u.utm_source || "", u.utm_medium || "", u.utm_campaign || "", u.utm_content || "", u.referrer || "",
    s.overall == null ? "" : s.overall,
    d.strategy ?? "", d.production ?? "", d.platform ?? "", d.packaging ?? "", d.distribution ?? "", d.moment ?? "",
    s.category || "", r.plan || "", r.activation ? "Yes" : "No", r.budget_constrained ? "Yes" : "No",
    (r.strongest || []).join(", "), (r.weakest || []).join(", "),
    JSON.stringify(p.answers || {}), "New submission", ""
  ];
}
function sendUserEmail_(p, email) {
  const s = p.scores || {}, r = p.recommendation || {}, c = p.contact || {};
  const name = c.name || "";
  const weak = (r.weakest || []).map(labelFr_).join(", ");
  const subject = "Votre Score Prizzz est prêt";
  const body =
`Bonjour ${name},

Votre Score Prizzz est de ${s.overall}/100.
Votre marque est actuellement au niveau : ${s.category}.

Vos principales zones de progression : ${weak}.

Cela signifie que votre communication peut gagner en structure, en régularité et en impact si elle est organisée comme un système mensuel.

Selon vos réponses, le chemin recommandé est : Prizzz ${r.plan}${r.activation ? " + Activation" : ""}.

Vous pouvez planifier un diagnostic avec Banj Media :
https://form.typeform.com/to/fcw8U53n

— Banj Media`;
  MailApp.sendEmail({ to: email, subject: subject, body: body, name: "Banj Media — Prizzz", replyTo: NOTIFY_EMAIL });
}
function notifyTeam_(p) {
  const s = p.scores || {}, r = p.recommendation || {}, c = p.contact || {}, l = p.lead || {};
  const cl = c.links || {};
  const linksTxt = ["instagram","facebook","x","tiktok","linkedin","website"]
    .filter(k => cl[k]).map(k => "  " + k + ": " + cl[k]).join("\n") || "  (aucun)";
  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: `[Prizzz] ${l.category || "Lead"} — ${c.name || c.email} (${s.overall}/100)`,
    body:
`Nouveau lead Prizzz Scorecard

Nom: ${c.name}
Entreprise: ${c.company}
Email: ${c.email}
WhatsApp: ${c.phone}
Liens:
${linksTxt}

Score: ${s.overall}/100 — ${s.category}
Plan recommandé: ${r.plan}${r.activation ? " + Activation" : ""}
Budget-contraint: ${r.budget_constrained ? "Oui" : "Non"}
Lead category: ${l.category} (quality ${l.quality_score})
Faiblesses: ${(r.weakest || []).join(", ")}

Voir la feuille pour le détail complet.`
  });
}
function labelFr_(k){return ({strategy:"Stratégie",production:"Production",platform:"Plateformes",packaging:"Packaging",distribution:"Distribution",moment:"Moments"})[k]||k;}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
