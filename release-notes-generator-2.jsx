import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";

// ── Infoblox Mountain Meadow brand tokens ─────────────────────────────────────
const C = {
  bg:"#090E0C", card:"#0F1612", active:"#132019", hover:"#172618",
  border:"#1E3028", borderHi:"#2E5040",
  green:"#1AB87C", gdark:"#138F5E", gdim:"#072818", gtext:"#3DD68C", gfaint:"#0F2E1E",
  red:"#E05050", rdim:"#2E1010",
  amber:"#E0A030", adim:"#2E2010",
  blue:"#4090D0", bdim:"#0A1E2E",
  teal:"#20C0A0", tdim:"#0A2820",
  text:"#E8F5F0", muted:"#6A9E88", faint:"#2A5040",
  mono:"'JetBrains Mono','Fira Code',monospace",
  sans:"'DM Sans',system-ui,sans-serif",
  disp:"'Barlow Condensed','DM Sans',system-ui,sans-serif",
};

const GS = `
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600&family=Barlow+Condensed:wght@600;700&family=JetBrains+Mono:wght@400&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{background:${C.bg};color:${C.text};font-family:${C.sans};min-height:100vh}
::-webkit-scrollbar{width:5px}::-webkit-scrollbar-track{background:transparent}
::-webkit-scrollbar-thumb{background:${C.border};border-radius:3px}
input[type=text],input[type=date],textarea{background:${C.card};color:${C.text};border:1px solid ${C.border};border-radius:6px;font-family:${C.sans};font-size:14px;padding:9px 12px;outline:none;width:100%;transition:border-color .15s,box-shadow .15s}
input[type=text]:focus,input[type=date]:focus,textarea:focus{border-color:${C.green};box-shadow:0 0 0 3px ${C.gdim}}
input::placeholder,textarea::placeholder{color:${C.faint}}
input[type=date]{color-scheme:dark}
input[type=checkbox]{width:auto;accent-color:${C.green};cursor:pointer;flex-shrink:0}
@keyframes fadeUp{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
@keyframes spin{to{transform:rotate(360deg)}}
@keyframes pulse{0%,100%{opacity:.6}50%{opacity:1}}
`;

// ── API (no key — runs inside Claude.ai artifact sandbox) ─────────────────────
async function callClaude(system, userMsg, withMcp = false) {
  const body = {
    model: "claude-sonnet-4-20250514",
    max_tokens: 2000,
    system,
    messages: [{ role: "user", content: userMsg }],
  };
  if (withMcp) body.mcp_servers = [{ type: "url", url: "https://mcp.atlassian.com/v1/mcp", name: "atlassian" }];
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`API error ${res.status}`);
  const data = await res.json();
  if (data.error) throw new Error(data.error.message);
  return (data.content || []).map(b => b.type === "text" ? b.text : (b.content?.[0]?.text || "")).join("");
}

function parseJsonArray(raw) {
  const m = raw.match(/\[[\s\S]*?\]/);
  try { return m ? JSON.parse(m[0]) : null; } catch { return null; }
}

// ── Column map for Jira exports ───────────────────────────────────────────────
const COL_MAP = {
  key:         ["issue key","key","id","ticket","issue id"],
  summary:     ["summary","title","name","subject"],
  type:        ["issue type","issuetype","type","kind"],
  status:      ["status","state","resolution status"],
  priority:    ["priority","severity"],
  assignee:    ["assignee","assigned to","owner"],
  fixVersion:  ["fix version/s","fix version","fix versions","version","release","custom field (release)"],
  description: ["description","details","body"],
  project:     ["project key","project","project name"],
};

function findCol(headers, candidates) {
  const lc = headers.map(h => (h || "").toLowerCase().trim());
  for (const c of candidates) { const i = lc.indexOf(c); if (i !== -1) return headers[i]; }
  return null;
}

function rowsToIssues(rows, headers) {
  const col = {};
  for (const [f, cs] of Object.entries(COL_MAP)) col[f] = findCol(headers, cs);
  return rows
    .filter(r => col.key ? r[col.key] : true)
    .map(r => ({
      key:         col.key         ? String(r[col.key]         || "").trim() : "",
      summary:     col.summary     ? String(r[col.summary]     || "").trim() : "",
      type:        col.type        ? String(r[col.type]        || "").trim() : "Task",
      status:      col.status      ? String(r[col.status]      || "").trim() : "",
      priority:    col.priority    ? String(r[col.priority]    || "").trim() : "",
      assignee:    col.assignee    ? String(r[col.assignee]    || "").trim() : "",
      fixVersion:  col.fixVersion  ? String(r[col.fixVersion]  || "").trim() : "",
      description: col.description ? String(r[col.description] || "").slice(0, 200).trim() : "",
      project:     col.project     ? String(r[col.project]     || "").trim() : "",
    }))
    .filter(r => r.key || r.summary);
}

function parseFile(file) {
  return new Promise((resolve, reject) => {
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "csv") {
      Papa.parse(file, {
        header: true, skipEmptyLines: true,
        complete: r => resolve({ rows: r.data, headers: r.meta.fields || [] }),
        error: e => reject(new Error(e.message)),
      });
    } else if (ext === "xlsx" || ext === "xls") {
      const reader = new FileReader();
      reader.onload = e => {
        try {
          const wb = XLSX.read(e.target.result, { type: "array" });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
          resolve({ rows, headers: rows.length ? Object.keys(rows[0]) : [] });
        } catch (err) { reject(err); }
      };
      reader.onerror = () => reject(new Error("File read failed"));
      reader.readAsArrayBuffer(file);
    } else {
      reject(new Error("Unsupported file type. Please upload .csv, .xlsx, or .xls"));
    }
  });
}

// ── Infoblox style system prompt ──────────────────────────────────────────────
function buildStyleSystem(fmt, extra) {
  const fmtGuide = {
    infoblox:   "Use the Infoblox house style exactly as described.",
    markdown:   "Output as Markdown with Infoblox prose style. Use ## for any section headings.",
    confluence: "Output as Confluence wiki markup. h2. for section headings. * for bullets.",
    html:       "Output as clean semantic HTML: <h2>, <ul><li>, <strong> for titles. No inline styles.",
    json:       `Output ONLY valid JSON: {"version":"...","date":"...","entries":[{"title":"Category: Feature Title","summary":"one declarative sentence","detail":"prose paragraph ending with For more information, see [Topic]."}]}`,
  };
  return `You are a senior technical writer at Infoblox. Write release notes in the exact Infoblox house style.

INFOBLOX RELEASE NOTE STYLE — follow every rule precisely:

1. ENTRY FORMAT — each feature:
   * [Category]: [Feature Title]:

   [One-sentence summary — declarative, present tense, NO "now"]
   [Full prose detail paragraph ending with "For more information, see [Topic Name]."]

2. TITLE: "* Category: Feature Title:" — title case, colons required, starts with "* "
   Category = product/area name: "NIOS-X", "BloxOne", "Infoblox Portal", "API"

3. SUMMARY LINE: Short declarative sentence. Present tense. NO "now". Just states the capability.
   Example: "NIOS-X supports automatic server right-sizing."

4. DETAIL PARAGRAPH: Opens with product name + "now supports/provides/includes/allows".
   Cover: WHAT it does · WHERE (exact UI page names) · HOW briefly · WHY (user benefit).
   MUST end: "For more information, see [Topic Name]."
   Keep to 2–4 sentences max.

5. LANGUAGE: Professional, active voice, present tense. Expand acronyms first use (e.g. OCI).
   Include version numbers where relevant. NEVER include Jira ticket keys. No markdown symbols in plain text.

6. SPACING: Blank line between title and body. Blank line between entries.

FORMAT: ${fmtGuide[fmt]}
${extra ? `EXTRA INSTRUCTIONS: ${extra}` : ""}

Output ONLY the entries. No preamble, no headings, no explanation.

EXAMPLE — match exactly:

* NIOS-X: Oracle Cloud Support:

NIOS-X supports deployment on Oracle Cloud Infrastructure (OCI).
You can now deploy NIOS-X servers on Oracle Cloud Infrastructure (OCI) using the Infoblox-provided OCI package, which you can download from the Infoblox Portal. For more information, see Oracle Cloud Infrastructure (OCI) Deployment.

* NIOS-X: Automatic Server Right-Sizing:

NIOS-X supports automatic server right-sizing.
NIOS-X now supports automatic server right-sizing, scaling servers to the next Infoblox-recommended size when capacity thresholds are exceeded multiple times within a monthly cycle, ensuring servers remain within supported performance limits. For more information, see Automatic Server Right-Sizing.

* Add Support for Apply Upgrades Now:

Infoblox now provides the ability to apply software updates immediately from the Servers page.
The Servers page includes a new Apply software updates now option that allows you to install the latest software updates for applications running on a selected NIOS-X server immediately, providing faster and more flexible update management. For more information, see Scheduling Software Updates for Servers.`;
}

// ── Shared tiny UI components ─────────────────────────────────────────────────
const Spin = ({ s = 16 }) => (
  <div style={{ width: s, height: s, border: `2px solid ${C.border}`, borderTopColor: C.green, borderRadius: "50%", animation: "spin .7s linear infinite", flexShrink: 0 }} />
);

const Badge = ({ children, color = "muted" }) => {
  const map = { green:[C.gdim,C.gtext], red:[C.rdim,C.red], amber:[C.adim,C.amber], blue:[C.bdim,C.blue], teal:[C.tdim,C.teal], muted:["#1A1E14",C.muted] };
  const [bg, fg] = map[color] || map.muted;
  return <span style={{ background:bg, color:fg, fontSize:11, fontWeight:600, padding:"2px 8px", borderRadius:4, whiteSpace:"nowrap", textTransform:"uppercase", letterSpacing:".03em" }}>{children}</span>;
};

const Btn = ({ onClick, disabled, children, variant = "primary", full, sx = {} }) => {
  const base = { fontFamily:C.sans, fontSize:13, fontWeight:500, padding:"8px 16px", borderRadius:6, border:"none", cursor:disabled?"not-allowed":"pointer", opacity:disabled?.45:1, transition:"all .15s", whiteSpace:"nowrap", display:"inline-flex", alignItems:"center", gap:6, ...(full?{width:"100%",justifyContent:"center",padding:"13px",fontSize:15}:{}) };
  const vs = { primary:{ background:C.green, color:"#090E0C", fontWeight:700 }, secondary:{ background:"transparent", color:C.text, border:`1px solid ${C.border}` }, ghost:{ background:"transparent", color:C.muted } };
  return <button onClick={onClick} disabled={disabled} style={{ ...base, ...(vs[variant]||vs.primary), ...sx }}>{children}</button>;
};

const Card = ({ children, glow, sx = {}, fwdRef }) => (
  <div ref={fwdRef} style={{ background:C.card, border:`1px solid ${glow?C.green+"55":C.border}`, borderRadius:8, padding:"1.25rem", marginBottom:12, animation:"fadeUp .35s ease", ...(glow?{boxShadow:`0 0 0 1px ${C.gdim}, 0 0 20px ${C.gdim}`}:{}), ...sx }}>
    {children}
  </div>
);

const StepHeader = ({ n, label, done }) => (
  <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:16 }}>
    <div style={{ width:24, height:24, borderRadius:4, background:done?C.green:C.gdim, color:done?"#090E0C":C.gtext, fontSize:12, fontWeight:700, display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0, fontFamily:C.disp, transition:"all .2s" }}>
      {done ? "✓" : n}
    </div>
    <span style={{ fontSize:11, fontWeight:700, color:C.muted, textTransform:"uppercase", letterSpacing:".1em", fontFamily:C.disp }}>{label}</span>
    {done && <div style={{ flex:1, height:1, background:C.gdim, marginLeft:4 }} />}
  </div>
);

const Field = ({ label, children }) => (
  <div style={{ display:"flex", flexDirection:"column", gap:6, flex:1, minWidth:0 }}>
    <label style={{ fontSize:11, color:C.muted, fontWeight:600, textTransform:"uppercase", letterSpacing:".06em" }}>{label}</label>
    {children}
  </div>
);

const Sel = ({ value, onChange, children, loading, placeholder, disabled }) => (
  <div style={{ position:"relative", flex:1, minWidth:0 }}>
    <select value={value} onChange={onChange} disabled={disabled || loading}
      style={{ width:"100%", background:C.card, color:value?C.text:C.faint, border:`1px solid ${C.border}`, borderRadius:6, fontFamily:C.sans, fontSize:14, padding:"9px 34px 9px 12px", outline:"none", cursor:loading?"wait":(disabled?"not-allowed":"pointer"), appearance:"none", WebkitAppearance:"none", opacity:disabled?.5:1, transition:"border-color .15s" }}
      onFocus={e=>{ e.target.style.borderColor=C.green; e.target.style.boxShadow=`0 0 0 3px ${C.gdim}`; }}
      onBlur={e=>{ e.target.style.borderColor=C.border; e.target.style.boxShadow="none"; }}>
      {placeholder && <option value="" disabled>{loading ? "Loading…" : placeholder}</option>}
      {children}
    </select>
    <div style={{ position:"absolute", right:10, top:"50%", transform:"translateY(-50%)", pointerEvents:"none" }}>
      {loading ? <Spin s={12} /> : <svg width="12" height="12" viewBox="0 0 12 12"><path d="M2 4l4 4 4-4" stroke={C.faint} strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" fill="none"/></svg>}
    </div>
  </div>
);

const MetricTile = ({ val, lbl, col }) => (
  <div style={{ background:C.active, border:`1px solid ${C.border}`, borderRadius:6, padding:"10px 14px" }}>
    <div style={{ fontSize:24, fontWeight:700, color:col||C.green, lineHeight:1, fontFamily:C.disp }}>{val}</div>
    <div style={{ fontSize:12, color:C.muted, marginTop:4, textTransform:"uppercase", letterSpacing:".05em" }}>{lbl}</div>
  </div>
);

const ErrBox = ({ msg }) => msg ? (
  <div style={{ marginTop:10, background:C.rdim, color:C.red, borderRadius:6, padding:"10px 14px", fontSize:13, border:`1px solid rgba(224,80,80,.2)`, lineHeight:1.5 }}>⚠ {msg}</div>
) : null;

const HintBox = ({ msg }) => msg ? (
  <div style={{ background:C.gdim, color:C.muted, borderRadius:6, padding:"8px 13px", fontSize:12, lineHeight:1.5, marginBottom:10 }}>{msg}</div>
) : null;

// ── Issue helpers ─────────────────────────────────────────────────────────────
const typeMeta = (t = "") => {
  const l = t.toLowerCase();
  if (l.includes("bug"))    return { color:"red",   lbl:"Bug" };
  if (l.includes("epic"))   return { color:"teal",  lbl:"Epic" };
  if (l.includes("story"))  return { color:"blue",  lbl:"Story" };
  if (l.includes("feature") || l.includes("improvement")) return { color:"green", lbl: l.includes("feature") ? "Feature" : "Improvement" };
  return { color:"muted", lbl: t || "Task" };
};
const statusColor = (s = "") => {
  const l = s.toLowerCase();
  if (l.includes("done") || l.includes("resolv") || l.includes("close")) return "green";
  if (l.includes("progress") || l.includes("review")) return "amber";
  return "muted";
};
const priorityColor = (p = "") => {
  const l = p.toLowerCase();
  if (l === "highest" || l === "critical") return "red";
  if (l === "high")   return "amber";
  if (l === "low" || l === "lowest") return "blue";
  return null;
};

// ════════════════════════════════════════════════════════════════════════════════
// MAIN APP
// ════════════════════════════════════════════════════════════════════════════════
export default function App() {
  const [mode, setMode]           = useState("import"); // "import" | "jira"

  // Jira state
  const [projects, setProjects]       = useState([]);
  const [projLoading, setProjLoading] = useState(false);
  const [projError, setProjError]     = useState("");
  const [selProject, setSelProject]   = useState("");
  const [sprints, setSprints]         = useState([]);
  const [sprintLoading, setSprintLoading] = useState(false);
  const [selSprint, setSelSprint]     = useState("");
  const [versions, setVersions]       = useState([]);
  const [verLoading, setVerLoading]   = useState(false);
  const [selVersion, setSelVersion]   = useState("");
  const [useJql, setUseJql]           = useState(false);
  const [customJql, setCustomJql]     = useState("");
  const [jiraStatus, setJiraStatus]   = useState("all");
  const [jiraType, setJiraType]       = useState("all");
  const [jiraMax, setJiraMax]         = useState("50");
  const [jiraFrom, setJiraFrom]       = useState("");
  const [jiraTo, setJiraTo]           = useState("");
  const [fetchLoading, setFetchLoading] = useState(false);
  const [fetchError, setFetchError]   = useState("");

  // Import state
  const [allRaw, setAllRaw]               = useState([]);
  const [importLoading, setImportLoading] = useState(false);
  const [importError, setImportError]     = useState("");
  const [importStats, setImportStats]     = useState(null);
  const [importProjects, setImportProjects] = useState([]);
  const [importVersions, setImportVersions] = useState([]);
  const [selImpProject, setSelImpProject] = useState("");
  const [selImpVersion, setSelImpVersion] = useState("");
  const [impStatus, setImpStatus]         = useState("all");
  const [impType, setImpType]             = useState("all");
  const fileRef = useRef(null);

  // Shared
  const [issues, setIssues]         = useState([]);
  const [selected, setSelected]     = useState(new Set());
  const [search, setSearch]         = useState("");

  // Generate
  const [verName, setVerName]       = useState("");
  const [fmt, setFmt]               = useState("infoblox");
  const [extra, setExtra]           = useState("");
  const [generating, setGenerating] = useState(false);
  const [genError, setGenError]     = useState("");
  const [output, setOutput]         = useState("");
  const [copied, setCopied]         = useState(false);
  const outRef = useRef(null);

  // ── Load Jira projects ──────────────────────────────────────────────────────
  const loadProjects = useCallback(async () => {
    setProjLoading(true); setProjError("");
    try {
      const raw = await callClaude(
        `Use the Jira MCP to list all accessible projects. Return ONLY a JSON array: [{"key":"PROJ","name":"Name","id":"1"},...]. No markdown.`,
        "List all Jira projects I have access to.", true
      );
      const list = parseJsonArray(raw);
      if (!list || !list.length) throw new Error("No projects returned. Check your Atlassian connection in Settings → Integrations.");
      setProjects(list);
    } catch (e) { setProjError(e.message); }
    finally { setProjLoading(false); }
  }, []);

  useEffect(() => {
    if (mode === "jira" && projects.length === 0 && !projLoading) loadProjects();
  }, [mode]);

  // ── Load sprints + versions on project change ───────────────────────────────
  useEffect(() => {
    if (!selProject) return;
    setSprints([]); setVersions([]); setSelSprint(""); setSelVersion("");
    const p = projects.find(x => x.key === selProject);
    if (p) setVerName(`${p.name} — Release`);

    setSprintLoading(true);
    callClaude(`Return ONLY JSON array of sprints: [{"id":"1","name":"Sprint 12","state":"active"},...] state:active|closed|future. Most recent first.`, `List all sprints for Jira project "${selProject}".`, true)
      .then(r => setSprints(parseJsonArray(r) || []))
      .catch(() => setSprints([]))
      .finally(() => setSprintLoading(false));

    setVerLoading(true);
    callClaude(`Return ONLY JSON array of fix versions: [{"id":"1","name":"v2.0","released":true,"releaseDate":"2026-01-01"},...]`, `List all fix versions for Jira project "${selProject}".`, true)
      .then(r => setVersions(parseJsonArray(r) || []))
      .catch(() => setVersions([]))
      .finally(() => setVerLoading(false));
  }, [selProject]);

  // ── Fetch Jira issues ───────────────────────────────────────────────────────
  const fetchJira = useCallback(async () => {
    if (!selProject && !useJql) { setFetchError("Select a project first."); return; }
    setFetchLoading(true); setFetchError(""); setIssues([]); setOutput(""); setSelected(new Set());

    let jql;
    if (useJql) { jql = customJql.trim(); }
    else {
      const pts = [`project = "${selProject}"`];
      if (selSprint)  pts.push(`sprint = "${selSprint}"`);
      if (selVersion) pts.push(`fixVersion = "${selVersion}"`);
      if (jiraStatus === "done")     pts.push(`status in ("Done","Resolved","Closed")`);
      if (jiraStatus === "released") pts.push(`status = "Released"`);
      if (jiraType === "bug")        pts.push(`issuetype = Bug`);
      if (jiraType === "feature")    pts.push(`issuetype in (Story,Feature,"New Feature",Improvement)`);
      if (jiraType === "no-epic")    pts.push(`issuetype != Epic`);
      if (jiraFrom) pts.push(`updated >= "${jiraFrom}"`);
      if (jiraTo)   pts.push(`updated <= "${jiraTo}"`);
      jql = pts.join(" AND ") + " ORDER BY created DESC";
    }
    try {
      const raw = await callClaude(
        `Use Jira MCP searchJiraIssuesUsingJql. Return ONLY a JSON array. Each: {"key":"PROJ-1","summary":"...","type":"Bug|Story|Task|Feature|Epic|Improvement","status":"...","assignee":"name or null","priority":"High|Medium|Low|null","description":"first 150 chars","fixVersion":"v or null"}. Limit ${jiraMax}.`,
        `Search Jira with JQL: ${jql}`, true
      );
      const list = parseJsonArray(raw);
      if (!list || !list.length) throw new Error("No issues found. Try setting status to \"All statuses\" or check your project key.");
      setIssues(list); setSelected(new Set(list.map(i => i.key)));
    } catch (e) { setFetchError(e.message); }
    finally { setFetchLoading(false); }
  }, [selProject, selSprint, selVersion, jiraStatus, jiraType, jiraMax, jiraFrom, jiraTo, useJql, customJql]);

  // ── Handle file import ──────────────────────────────────────────────────────
  const handleImport = useCallback(async (file) => {
    setImportLoading(true); setImportError(""); setAllRaw([]); setIssues([]); setOutput(""); setSelected(new Set()); setImportStats(null);
    try {
      const { rows, headers } = await parseFile(file);
      if (!rows.length) throw new Error("File is empty.");
      const parsed = rowsToIssues(rows, headers);
      if (!parsed.length) throw new Error("No valid issues found. Make sure the file has 'Issue Key' and 'Summary' columns.");

      setAllRaw(parsed);
      const projs = [...new Set(parsed.map(i => i.project || (i.key ? i.key.split("-")[0] : "")).filter(Boolean))];
      const vers  = [...new Set(parsed.map(i => i.fixVersion).filter(Boolean))];
      const stats = [...new Set(parsed.map(i => i.status).filter(Boolean))];

      setImportProjects(projs);
      setImportVersions(vers);
      if (projs.length === 1) { setSelImpProject(projs[0]); setVerName(`${projs[0]} — Release`); }
      if (vers.length  === 1) setSelImpVersion(vers[0]);
      setImpStatus("all");
      setImportStats({ filename: file.name, total: parsed.length, projects: projs.length, versions: vers.length, statuses: stats });
    } catch (e) { setImportError(e.message); }
    finally { setImportLoading(false); }
  }, []);

  // ── Apply import filters ────────────────────────────────────────────────────
  useEffect(() => {
    if (!allRaw.length) return;
    let f = allRaw;
    if (selImpProject) f = f.filter(i => (i.project || (i.key ? i.key.split("-")[0] : "")) === selImpProject);
    if (selImpVersion) f = f.filter(i => i.fixVersion === selImpVersion);
    if (impStatus === "done") f = f.filter(i => { const l = (i.status||"").toLowerCase(); return l.includes("done")||l.includes("resolv")||l.includes("close"); });
    if (impType === "bug")     f = f.filter(i => (i.type||"").toLowerCase().includes("bug"));
    if (impType === "feature") f = f.filter(i => { const l=(i.type||"").toLowerCase(); return l.includes("story")||l.includes("feature")||l.includes("improvement"); });
    if (impType === "no-epic") f = f.filter(i => !(i.type||"").toLowerCase().includes("epic"));
    setIssues(f); setSelected(new Set(f.map(i => i.key)));
  }, [allRaw, selImpProject, selImpVersion, impStatus, impType]);

  // ── Generate ────────────────────────────────────────────────────────────────
  const generate = useCallback(async () => {
    const sel = issues.filter(i => selected.has(i.key));
    if (!sel.length) { setGenError("Select at least one issue."); return; }
    setGenerating(true); setGenError(""); setOutput("");

    const lines = sel.map(i =>
      `- Key: ${i.key}\n  Summary: ${i.summary}` +
      (i.type        ? `\n  Type: ${i.type}` : "") +
      (i.status      ? `\n  Status: ${i.status}` : "") +
      (i.description ? `\n  Description: ${i.description.slice(0, 150)}` : "") +
      (i.fixVersion  ? `\n  Fix Version: ${i.fixVersion}` : "")
    ).join("\n");

    try {
      const raw = await callClaude(
        buildStyleSystem(fmt, extra),
        `Write Infoblox release notes for: "${verName || "this release"}"\n\nIssues:\n${lines}`,
        false
      );
      if (!raw.trim()) throw new Error("Empty response. Try again.");
      setOutput(raw.trim());
      setTimeout(() => outRef.current?.scrollIntoView({ behavior: "smooth", block: "start" }), 150);
    } catch (e) { setGenError(e.message); }
    finally { setGenerating(false); }
  }, [issues, selected, fmt, verName, extra]);

  // ── Issue list helpers ──────────────────────────────────────────────────────
  const toggleIssue = key => { const s = new Set(selected); s.has(key) ? s.delete(key) : s.add(key); setSelected(s); };
  const toggleAll   = v   => setSelected(v ? new Set(issues.map(i => i.key)) : new Set());
  const filtered    = issues.filter(i => !search || i.key.toLowerCase().includes(search.toLowerCase()) || i.summary.toLowerCase().includes(search.toLowerCase()));
  const mx = { total: issues.length, bugs: issues.filter(i=>(i.type||"").toLowerCase().includes("bug")).length, feat: issues.filter(i=>{const l=(i.type||"").toLowerCase();return l.includes("story")||l.includes("feature")||l.includes("improvement");}).length, hi: issues.filter(i=>{const p=(i.priority||"").toLowerCase();return p==="high"||p==="highest"||p==="critical";}).length };
  const activeSprint = sprints.find(s => s.state === "active");
  const stepN = n => mode === "jira" ? n : n - 1;

  // ── Render ──────────────────────────────────────────────────────────────────
  return (
    <>
      <style>{GS}</style>
      <div style={{ maxWidth:840, margin:"0 auto", padding:"1.5rem" }}>

        {/* Header */}
        <div style={{ marginBottom:"1.5rem", paddingBottom:"1.25rem", borderBottom:`1px solid ${C.border}` }}>
          <div style={{ display:"flex", alignItems:"center", gap:14, marginBottom:8 }}>
            <svg width="34" height="34" viewBox="0 0 36 36" fill="none">
              <polygon points="18,2 34,18 18,34 2,18" fill={C.green}/>
              <polygon points="18,8 28,18 18,28 8,18" fill={C.bg}/>
              <polygon points="18,13 23,18 18,23 13,18" fill={C.green}/>
            </svg>
            <div>
              <div style={{ display:"flex", alignItems:"baseline", gap:8 }}>
                <span style={{ fontFamily:C.disp, fontSize:26, fontWeight:700, color:C.text, letterSpacing:".02em" }}>RELEASE NOTES</span>
                <span style={{ fontFamily:C.disp, fontSize:26, fontWeight:700, color:C.green, letterSpacing:".02em" }}>GENERATOR</span>
              </div>
              <div style={{ fontSize:13, color:C.muted, marginTop:2 }}>Infoblox · Import CSV / Excel or connect Jira · Infoblox house style AI</div>
            </div>
          </div>
          <div style={{ height:2, background:`linear-gradient(90deg,${C.green},${C.gdim})`, borderRadius:1 }} />
        </div>

        {/* Mode tabs */}
        <div style={{ display:"flex", gap:6, background:C.card, border:`1px solid ${C.border}`, borderRadius:8, padding:4, marginBottom:14 }}>
          {[
            { id:"import", icon:"↑", label:"Import CSV / Excel" },
            { id:"jira",   icon:"⬡", label:"Connect to Jira" },
          ].map(t => (
            <button key={t.id} onClick={() => { setMode(t.id); setIssues([]); setOutput(""); setSelected(new Set()); }}
              style={{ flex:1, display:"flex", alignItems:"center", justifyContent:"center", gap:7, padding:"10px 16px", border:"none", borderRadius:6, cursor:"pointer", fontFamily:C.sans, fontSize:14, fontWeight:600, transition:"all .15s", background:mode===t.id?C.green:"transparent", color:mode===t.id?"#090E0C":C.muted }}>
              <span style={{ fontSize:15 }}>{t.icon}</span>{t.label}
            </button>
          ))}
        </div>

        {/* ── IMPORT MODE ──────────────────────────────────────────────────── */}
        {mode === "import" && (
          <Card>
            <StepHeader n="1" label="Import Jira export file" done={!!importStats} />

            {/* Drop zone */}
            {!importStats && !importLoading && (
              <div onDrop={e=>{e.preventDefault();e.currentTarget.classList.remove("dz-over");const f=e.dataTransfer.files[0];if(f)handleImport(f);}}
                   onDragOver={e=>{e.preventDefault();e.currentTarget.classList.add("dz-over");}}
                   onDragLeave={e=>e.currentTarget.classList.remove("dz-over")}
                   onClick={() => fileRef.current?.click()}
                   style={{ border:`2px dashed ${C.borderHi}`, borderRadius:8, padding:"2.25rem 1.5rem", textAlign:"center", cursor:"pointer", background:C.active, transition:"all .15s" }}
                   onMouseEnter={e=>{e.currentTarget.style.borderColor=C.green;e.currentTarget.style.background=C.gfaint;}}
                   onMouseLeave={e=>{e.currentTarget.style.borderColor=C.borderHi;e.currentTarget.style.background=C.active;}}>
                <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls" style={{ display:"none" }} onChange={e=>{if(e.target.files[0])handleImport(e.target.files[0]);}} />
                <svg width="40" height="40" viewBox="0 0 40 40" fill="none" style={{ margin:"0 auto 12px", display:"block" }}>
                  <rect width="40" height="40" rx="8" fill={C.gdim}/>
                  <path d="M20 10v14M13 17l7-7 7 7M10 28h20" stroke={C.green} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                </svg>
                <div style={{ fontSize:15, fontWeight:600, color:C.text, marginBottom:6 }}>Drop your Jira export here</div>
                <div style={{ fontSize:13, color:C.muted }}>or click to browse · <strong style={{color:C.gtext}}>.csv</strong> · <strong style={{color:C.gtext}}>.xlsx</strong> · <strong style={{color:C.gtext}}>.xls</strong></div>
                <div style={{ fontSize:11, color:C.faint, marginTop:10, lineHeight:1.7 }}>
                  Jira → Issues → Export → Export Excel CSV (all fields)<br/>
                  Auto-maps: Issue Key, Summary, Type, Status, Priority, Fix Version, Assignee
                </div>
              </div>
            )}

            {importLoading && <div style={{ display:"flex", alignItems:"center", gap:10, color:C.muted, fontSize:13, padding:"12px 0" }}><Spin />Parsing file and extracting issues…</div>}
            <ErrBox msg={importError} />

            {/* Imported summary */}
            {importStats && !importLoading && (
              <>
                {/* File pill */}
                <div style={{ display:"flex", alignItems:"center", gap:10, background:C.gfaint, border:`1px solid ${C.green}33`, borderRadius:6, padding:"10px 14px", marginBottom:14 }}>
                  <svg width="16" height="16" viewBox="0 0 16 16" fill="none"><rect width="16" height="16" rx="3" fill={C.gdim}/><path d="M4 5h8M4 8h8M4 11h5" stroke={C.green} strokeWidth="1.5" strokeLinecap="round"/></svg>
                  <span style={{ fontSize:13, color:C.gtext, fontWeight:600, flex:1, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{importStats.filename}</span>
                  <span style={{ fontSize:12, color:C.muted, whiteSpace:"nowrap" }}>{importStats.total} issues · {importStats.projects} project{importStats.projects!==1?"s":""} · {importStats.versions} version{importStats.versions!==1?"s":""}</span>
                  <button onClick={()=>{setImportStats(null);setAllRaw([]);setIssues([]);setOutput("");setSelected(new Set());if(fileRef.current)fileRef.current.value="";}} style={{ background:"none", border:"none", color:C.faint, cursor:"pointer", fontSize:18, lineHeight:1, padding:"0 3px" }}>✕</button>
                </div>

                <HintBox msg={`Detected statuses: ${importStats.statuses.join(", ") || "(none)"}. Status filter defaults to "All statuses" so nothing is hidden.`} />

                <div style={{ display:"flex", gap:12, marginBottom:12, flexWrap:"wrap" }}>
                  <Field label={`Project — ${importProjects.length} detected`}>
                    <Sel value={selImpProject} onChange={e=>setSelImpProject(e.target.value)} placeholder={importProjects.length?"Filter by project":"All projects"}>
                      <option value="">All projects</option>
                      {importProjects.map(p => <option key={p} value={p}>{p}</option>)}
                    </Sel>
                  </Field>
                  <Field label={`Fix version — ${importVersions.length} detected`}>
                    <Sel value={selImpVersion} onChange={e=>setSelImpVersion(e.target.value)} placeholder={importVersions.length?"Filter by version":"No versions found"}>
                      <option value="">All versions</option>
                      {importVersions.map(v => <option key={v} value={v}>{v}</option>)}
                    </Sel>
                  </Field>
                </div>
                <div style={{ display:"flex", gap:12, marginBottom:12, flexWrap:"wrap" }}>
                  <Field label="Status filter">
                    <Sel value={impStatus} onChange={e=>setImpStatus(e.target.value)}>
                      <option value="all">All statuses — show everything</option>
                      <option value="done">Done / Resolved / Closed only</option>
                    </Sel>
                  </Field>
                  <Field label="Issue type">
                    <Sel value={impType} onChange={e=>setImpType(e.target.value)}>
                      <option value="all">All types</option>
                      <option value="bug">Bugs only</option>
                      <option value="feature">Features, Stories &amp; Improvements</option>
                      <option value="no-epic">Exclude Epics</option>
                    </Sel>
                  </Field>
                </div>
                <div style={{ display:"flex", alignItems:"center" }}>
                  <span style={{ fontSize:13, color:C.muted }}>{issues.length} issue{issues.length!==1?"s":""} ready</span>
                  <button onClick={()=>{setImportStats(null);setAllRaw([]);setIssues([]);setOutput("");setSelected(new Set());if(fileRef.current)fileRef.current.value="";}} style={{ background:"none", border:"none", color:C.muted, cursor:"pointer", fontSize:13, textDecoration:"underline", padding:0, marginLeft:"auto" }}>Upload different file</button>
                </div>
              </>
            )}
          </Card>
        )}

        {/* ── JIRA MODE ────────────────────────────────────────────────────── */}
        {mode === "jira" && (
          <>
            <Card>
              <StepHeader n="1" label="Select project & sprint" done={!!selProject} />

              {projLoading && <div style={{ display:"flex", alignItems:"center", gap:10, color:C.muted, fontSize:13, padding:"8px 0" }}><Spin />Connecting to Jira…</div>}
              <ErrBox msg={projError} />
              {projError && (
                <div style={{ marginTop:8, fontSize:13, color:C.muted }}>
                  Can't connect? <button onClick={()=>setMode("import")} style={{ background:"none", border:"none", color:C.green, cursor:"pointer", fontSize:13, fontWeight:600, textDecoration:"underline", padding:0 }}>Import a CSV / Excel file instead →</button>
                </div>
              )}
              {!projLoading && !projError && projects.length === 0 && (
                <div style={{ background:C.bdim, color:C.blue, borderRadius:6, padding:"9px 13px", fontSize:12, marginTop:8 }}>No projects loaded. Check your Atlassian connection in Settings → Integrations.</div>
              )}
              {projects.length > 0 && (
                <>
                  <div style={{ display:"flex", gap:12, marginBottom:12, flexWrap:"wrap" }}>
                    <Field label={`Jira project — ${projects.length} available`}>
                      <Sel value={selProject} onChange={e=>setSelProject(e.target.value)} placeholder="Choose a project">
                        {projects.map(p => <option key={p.key} value={p.key}>{p.key} — {p.name}</option>)}
                      </Sel>
                    </Field>
                    <Field label={<span>Sprint {activeSprint && <span style={{ background:C.gdim, color:C.green, fontSize:10, fontWeight:700, padding:"1px 7px", borderRadius:3, marginLeft:5 }}>● Active: {activeSprint.name}</span>}</span>}>
                      <Sel value={selSprint} onChange={e=>setSelSprint(e.target.value)} loading={sprintLoading} placeholder={!selProject?"Select project first":"All sprints"} disabled={!selProject}>
                        <option value="">All sprints</option>
                        {sprints.map(s => <option key={s.id} value={s.name}>{s.state==="active"?"★ ":s.state==="future"?"○ ":""}{s.name}</option>)}
                      </Sel>
                    </Field>
                    <Field label="Fix version">
                      <Sel value={selVersion} onChange={e=>setSelVersion(e.target.value)} loading={verLoading} placeholder={!selProject?"Select project first":"Any version"} disabled={!selProject}>
                        <option value="">Any version</option>
                        {versions.map(v => <option key={v.id} value={v.name}>{v.name}{v.released?" ✓":""}{v.releaseDate?` · ${v.releaseDate}`:""}</option>)}
                      </Sel>
                    </Field>
                  </div>
                  {sprints.length > 0 && <div style={{ fontSize:11, color:C.faint, marginBottom:12 }}>★ active · ○ future · ✓ released</div>}
                </>
              )}
              <label style={{ display:"flex", alignItems:"center", gap:8, cursor:"pointer", fontSize:13, color:C.muted }}>
                <input type="checkbox" checked={useJql} onChange={e=>setUseJql(e.target.checked)} />
                Override with custom JQL
              </label>
              {useJql && (
                <textarea value={customJql} onChange={e=>setCustomJql(e.target.value)} rows={2} placeholder='e.g. project = "PROJ" AND sprint in openSprints() AND status = Done' style={{ fontFamily:C.mono, fontSize:13, marginTop:10 }} />
              )}
            </Card>

            <Card>
              <StepHeader n="2" label="Filters" />
              <div style={{ display:"flex", gap:12, marginBottom:12, flexWrap:"wrap" }}>
                <Field label="Status">
                  <Sel value={jiraStatus} onChange={e=>setJiraStatus(e.target.value)}>
                    <option value="all">All statuses</option>
                    <option value="done">Done / Resolved / Closed</option>
                    <option value="released">Released</option>
                  </Sel>
                </Field>
                <Field label="Issue type">
                  <Sel value={jiraType} onChange={e=>setJiraType(e.target.value)}>
                    <option value="all">All types</option>
                    <option value="bug">Bugs only</option>
                    <option value="feature">Features, Stories &amp; Improvements</option>
                    <option value="no-epic">Exclude Epics</option>
                  </Sel>
                </Field>
                <Field label="Max issues">
                  <Sel value={jiraMax} onChange={e=>setJiraMax(e.target.value)}>
                    {["20","50","100","200"].map(v=><option key={v} value={v}>{v}</option>)}
                  </Sel>
                </Field>
              </div>
              <div style={{ display:"flex", gap:12, flexWrap:"wrap", alignItems:"flex-end" }}>
                <Field label="Updated from"><input type="date" value={jiraFrom} onChange={e=>setJiraFrom(e.target.value)} /></Field>
                <Field label="Updated to"><input type="date" value={jiraTo} onChange={e=>setJiraTo(e.target.value)} /></Field>
                <Btn onClick={fetchJira} disabled={fetchLoading||(!selProject&&!useJql)} sx={{ padding:"9px 24px", fontSize:14, flexShrink:0 }}>
                  {fetchLoading ? <><Spin />Fetching…</> : "Fetch issues →"}
                </Btn>
              </div>
              <ErrBox msg={fetchError} />
            </Card>
          </>
        )}

        {/* ── SELECT ISSUES (shared) ────────────────────────────────────────── */}
        {issues.length > 0 && (
          <Card>
            <StepHeader n={stepN(3)} label="Select issues" done={selected.size > 0} />

            <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:8, marginBottom:14 }}>
              <MetricTile val={mx.total} lbl="Total" />
              <MetricTile val={mx.bugs}  lbl="Bugs"     col={C.red} />
              <MetricTile val={mx.feat}  lbl="Features" col={C.green} />
              <MetricTile val={mx.hi}    lbl="High pri" col={C.amber} />
            </div>

            <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:10, flexWrap:"wrap" }}>
              <input type="text" value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search issues…" style={{ maxWidth:220, padding:"7px 12px", fontSize:13 }} />
              <span style={{ fontSize:13, color:C.muted, marginLeft:"auto" }}>{selected.size} / {issues.length} selected</span>
              <Btn variant="secondary" onClick={()=>toggleAll(true)}  sx={{ fontSize:12, padding:"6px 12px" }}>All</Btn>
              <Btn variant="ghost"     onClick={()=>toggleAll(false)} sx={{ fontSize:12, padding:"6px 12px" }}>Clear</Btn>
            </div>

            <div style={{ display:"flex", flexDirection:"column", gap:5, maxHeight:340, overflowY:"auto", paddingRight:3 }}>
              {filtered.length === 0 && (
                <div style={{ padding:"2rem", textAlign:"center", color:C.faint, fontSize:13 }}>No issues match. Try changing the status filter to "All statuses".</div>
              )}
              {filtered.map(issue => {
                const tm = typeMeta(issue.type), sc = statusColor(issue.status), pc = priorityColor(issue.priority), sel = selected.has(issue.key);
                return (
                  <div key={issue.key} onClick={()=>toggleIssue(issue.key)}
                    style={{ display:"flex", alignItems:"center", gap:10, padding:"9px 12px", border:`1px solid ${sel?C.green+"55":C.border}`, borderLeft:`3px solid ${sel?C.green:C.border}`, borderRadius:6, cursor:"pointer", background:sel?C.gfaint:C.card, transition:"all .12s" }}>
                    <input type="checkbox" checked={sel} onChange={()=>{}} onClick={e=>e.stopPropagation()} />
                    <span style={{ fontSize:12, color:C.gtext, fontFamily:C.mono, minWidth:80, flexShrink:0 }}>{issue.key}</span>
                    <span style={{ fontSize:13, color:C.text, flex:1, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{issue.summary}</span>
                    <div style={{ display:"flex", gap:5, flexShrink:0, alignItems:"center" }}>
                      <Badge color={tm.color}>{tm.lbl}</Badge>
                      {pc && <Badge color={pc}>{issue.priority}</Badge>}
                      <Badge color={sc}>{issue.status || "—"}</Badge>
                    </div>
                  </div>
                );
              })}
            </div>
          </Card>
        )}

        {/* ── GENERATE (shared) ────────────────────────────────────────────── */}
        {issues.length > 0 && (
          <Card>
            <StepHeader n={stepN(4)} label="Generate release notes" />

            {/* Style badge */}
            <div style={{ display:"flex", alignItems:"center", gap:8, background:C.gfaint, border:`1px solid ${C.green}33`, borderRadius:6, padding:"8px 13px", marginBottom:14 }}>
              <svg width="13" height="13" viewBox="0 0 36 36" fill="none"><polygon points="18,2 34,18 18,34 2,18" fill={C.green}/><polygon points="18,9 27,18 18,27 9,18" fill={C.card}/><polygon points="18,14 22,18 18,22 14,18" fill={C.green}/></svg>
              <span style={{ fontSize:12, color:C.gtext, fontWeight:600 }}>Infoblox house style applied</span>
              <span style={{ fontSize:12, color:C.muted, marginLeft:3 }}>· Category: Title · Summary · Prose · "For more information, see…"</span>
            </div>

            <div style={{ display:"flex", gap:12, marginBottom:12, flexWrap:"wrap" }}>
              <Field label="Version / release name">
                <input type="text" value={verName} onChange={e=>setVerName(e.target.value)} placeholder="e.g. NIOS-X v4.1 — May 2026" />
              </Field>
              <Field label="Output format">
                <Sel value={fmt} onChange={e=>setFmt(e.target.value)}>
                  <option value="infoblox">Infoblox plain text (default)</option>
                  <option value="markdown">Markdown</option>
                  <option value="confluence">Confluence wiki markup</option>
                  <option value="html">HTML</option>
                  <option value="json">JSON (structured)</option>
                </Sel>
              </Field>
            </div>
            <div style={{ marginBottom:16 }}>
              <Field label="Extra instructions (optional)">
                <input type="text" value={extra} onChange={e=>setExtra(e.target.value)} placeholder="e.g. Group under NIOS-X, highlight breaking changes, include OCI references…" />
              </Field>
            </div>

            <Btn onClick={generate} disabled={generating || selected.size === 0} full>
              {generating ? <><Spin />Writing release notes…</> : `Generate notes for ${selected.size} issue${selected.size!==1?"s":""} →`}
            </Btn>
            <ErrBox msg={genError} />
          </Card>
        )}

        {/* ── OUTPUT ───────────────────────────────────────────────────────── */}
        {output && (
          <Card glow fwdRef={outRef}>
            <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:14 }}>
              <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                <div style={{ width:8, height:8, borderRadius:2, background:C.green, animation:"pulse 2s ease-in-out infinite" }} />
                <span style={{ fontSize:11, fontWeight:700, color:C.gtext, textTransform:"uppercase", letterSpacing:".1em", fontFamily:C.disp }}>Output · {fmt}</span>
              </div>
              <div style={{ display:"flex", gap:8 }}>
                <Btn variant="secondary" onClick={()=>navigator.clipboard.writeText(output).then(()=>{setCopied(true);setTimeout(()=>setCopied(false),2000);})} sx={{ fontSize:12, padding:"6px 12px" }}>
                  {copied ? "✓ Copied" : "Copy"}
                </Btn>
                <Btn variant="secondary" onClick={generate} sx={{ fontSize:12, padding:"6px 12px" }}>Regenerate</Btn>
              </div>
            </div>
            <pre style={{ fontFamily:C.mono, fontSize:13, lineHeight:1.8, color:C.text, background:C.active, borderRadius:6, padding:"1rem", overflowX:"auto", whiteSpace:"pre-wrap", wordBreak:"break-word", maxHeight:520, overflowY:"auto" }}>
              {output}
            </pre>
          </Card>
        )}

        {/* Footer */}
        <div style={{ marginTop:"1.5rem", paddingTop:"1rem", borderTop:`1px solid ${C.border}`, display:"flex", alignItems:"center", gap:8 }}>
          <svg width="13" height="13" viewBox="0 0 36 36" fill="none"><polygon points="18,2 34,18 18,34 2,18" fill={C.green} opacity=".6"/><polygon points="18,9 27,18 18,27 9,18" fill={C.bg}/><polygon points="18,14 22,18 18,22 14,18" fill={C.green} opacity=".6"/></svg>
          <span style={{ fontSize:12, color:C.faint }}>Infoblox Release Notes Generator · Powered by Claude AI · No API key required</span>
        </div>
      </div>
    </>
  );
}
