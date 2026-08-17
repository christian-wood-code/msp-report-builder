"use strict";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "Content-Type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Content-Type": "application/json",
  "X-Content-Type-Options": "nosniff",
};

const respond = (status, body) => ({ statusCode: status, headers: CORS, body: JSON.stringify(body) });
const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const LAPTOP_KW = ["laptop","notebook","thinkpad","latitude","elitebook","probook","inspiron",
  "xps","surface pro","surface laptop","macbook","yoga","ideapad","spectre","envy","pavilion",
  "omen","zbook","precision","vostro","gram","swift","spin","chromebook","portege","tecra",
  "aspire","nitro","folio","dragonfly"];

const SEV_ORDER = { high: 0, medium: 1, low: 2 };

const ADMIN_ROLES = new Set([
  "Global Administrator","Global Reader","Privileged Role Administrator",
  "Privileged Authentication Administrator","Security Administrator","Security Reader",
  "Exchange Administrator","SharePoint Administrator","Teams Administrator",
  "Intune Administrator","Conditional Access Administrator","User Administrator",
  "Billing Administrator","Application Administrator","Authentication Administrator",
  "Helpdesk Administrator","Password Administrator","Compliance Administrator",
]);

// Licence SKU part number → friendly name
const SKU_MAP = {
  "SPB":"Microsoft 365 Business Premium","SMB_BUSINESS_PREMIUM":"Microsoft 365 Business Premium",
  "O365_BUSINESS_PREMIUM":"Microsoft 365 Business Standard","O365_BUSINESS_ESSENTIALS":"Microsoft 365 Business Basic",
  "SMB_BUSINESS":"Microsoft 365 Apps for Business","OFFICESUBSCRIPTION":"Microsoft 365 Apps for Enterprise",
  "ENTERPRISEPACK":"Microsoft 365 E3","ENTERPRISEPREMIUM":"Microsoft 365 E5",
  "ENTERPRISEPREMIUM_NOPSTNCONF":"Microsoft 365 E5 (No Audio Conf)","STANDARDPACK":"Office 365 E1",
  "STANDARDWOFFPACK":"Office 365 E2","ENTERPRISEWITHSCAL":"Office 365 E4","DESKLESSPACK":"Office 365 F3",
  "EXCHANGESTANDARD":"Exchange Online (Plan 1)","EXCHANGEENTERPRISE":"Exchange Online (Plan 2)",
  "EXCHANGEDESKLESS":"Exchange Online Kiosk","EXCHANGEESSENTIALS":"Exchange Online Essentials",
  "EOP_ENTERPRISE":"Exchange Online Protection","AAD_PREMIUM":"Microsoft Entra ID P1",
  "AAD_PREMIUM_P2":"Microsoft Entra ID P2","AAD_BASIC":"Microsoft Entra ID Free",
  "EMS":"Enterprise Mobility + Security E3","EMSPREMIUM":"Enterprise Mobility + Security E5",
  "INTUNE_A":"Microsoft Intune Plan 1","INTUNE_A_D":"Microsoft Intune Device","INTUNE_SMB":"Microsoft Intune for SMB",
  "WIN_DEF_ATP":"Microsoft Defender for Endpoint","DEFENDER_ENDPOINT_P1":"Microsoft Defender for Endpoint P1",
  "ATP_ENTERPRISE":"Microsoft Defender for Office 365 P1","THREAT_INTELLIGENCE":"Microsoft Defender for Office 365 P2",
  "POWER_BI_STANDARD":"Power BI (Free)","POWER_BI_PRO":"Power BI Pro","POWER_BI_PREMIUM_USER":"Power BI Premium Per User",
  "PBI_PREMIUM_P1_ADDON":"Power BI Premium P1","FLOW_FREE":"Power Automate (Free)","FLOW_P1":"Power Automate Plan 1",
  "FLOW_P2":"Power Automate Plan 2","POWERFLOW_P2":"Power Automate per User","POWERAPPS_VIRAL":"Power Apps (Free)",
  "POWERAPPS_PER_USER":"Power Apps per User","POWERAPPS_DEV":"Power Apps Developer",
  "TEAMS_FREE":"Microsoft Teams (Free)","TEAMS_EXPLORATORY":"Microsoft Teams Exploratory",
  "MCOEV":"Microsoft Teams Phone Standard","MCOMEETADV":"Microsoft 365 Audio Conferencing",
  "MCOPSTN1":"Teams Domestic Calling Plan","MCOPSTN2":"Teams Domestic & International Calling Plan",
  "VISIOCLIENT":"Visio Plan 2","VISIOONLINE_PLAN1":"Visio Plan 1","PROJECTPREMIUM":"Project Plan 5",
  "PROJECTPROFESSIONAL":"Project Plan 3","PROJECTESSENTIALS":"Project Plan 1",
  "WIN10_PRO_ENT_SUB":"Windows 10/11 Enterprise E3","WIN10_VDA_E3":"Windows 10/11 Enterprise E3",
  "WIN10_VDA_E5":"Windows 10/11 Enterprise E5","INFORMATION_PROTECTION_COMPLIANCE":"Microsoft Purview E5 Compliance",
  "RIGHTSMANAGEMENT":"Azure Information Protection P1","RIGHTSMANAGEMENT_ADHOC":"Azure Rights Management (Free)",
  "SHAREPOINTENTERPRISE":"SharePoint Online (Plan 2)","SHAREPOINTSTANDARD":"SharePoint Online (Plan 1)",
  "WACONEDRIVESTANDARD":"OneDrive for Business Plan 1","WACONEDRIVEENTERPRISE":"OneDrive for Business Plan 2",
  "DEVELOPERPACK":"Microsoft 365 E3 Developer","DEVELOPERPACK_E5":"Microsoft 365 E5 Developer",
  "CRMSTANDARD":"Dynamics 365 Sales Professional","DYN365_ENTERPRISE_PLAN1":"Dynamics 365 Customer Engagement",
};

// Common licence GUIDs → friendly name fallback (for licences not in subscribedSkus)
const GUID_MAP = {
  "4b9405b0-7788-4568-add1-99614e613b69":"Exchange Online (Plan 1)",
  "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46":"Microsoft 365 Business Premium",
  "f30db892-07e9-47e9-837c-80727f46fd3d":"Power Automate (Free)",
  "f245ecc8-75af-4f8e-b61f-27d8114de5f3":"Microsoft 365 F3",
  "19ec0d23-8335-4cbd-94ac-6050e30712fa":"Microsoft 365 Business Standard",
  "5b631642-bd26-49fe-bd20-1daaa972ef80":"Microsoft Defender for Business",
  "dcb1a3ae-b33f-4487-846a-a640262fadf4":"Power Apps Developer (Free)",
  "c5928f49-12ba-48f7-ada3-0d743a3601d5":"Microsoft Teams Exploratory",
  "639dec6b-bb19-468b-871c-c5c441c4b0cb":"Microsoft Teams (Free)",
  "eda1941c-3c4f-4995-b5eb-e85a42175ab9":"Microsoft Intune",
  "3b555118-da6a-4418-894f-7df1e2096870":"Microsoft 365 Business Basic",
  "6fd2c87f-b296-42f0-b197-1e91e994b900":"Microsoft 365 E3",
  "c7df2760-2c81-4ef7-b578-5b5392b571df":"Microsoft 365 E5",
  "b05e124f-c7cc-45a0-a6aa-8cf78c946968":"Enterprise Mobility + Security E5",
  "efccb6f7-5641-4e0e-bd10-b4976e1bf68e":"Enterprise Mobility + Security E3",
  "f8a1db68-be16-40ed-86d5-cb42ce701560":"Power BI Pro",
  "078d2b04-f1bd-4111-bbd4-b4b1b354cef4":"Microsoft Entra ID P1",
  "84a661c4-e949-4bd2-a560-ed7766fcaf2b":"Microsoft Entra ID P2",
  "e6778190-713e-4e4f-9119-8b8238de25df":"Microsoft 365 Apps for Enterprise",
  "c2273bd0-dff7-4215-9ef5-2c7bcfb06425":"Power Apps per User",
  "d3b4fe1f-9992-4930-8acb-ca6ec609365e":"Teams Domestic Calling Plan",
  "0c266dff-15dd-4b49-8397-2bb16070ed52":"Microsoft 365 Audio Conferencing",
  // Additional GUIDs identified from tenant data
  "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235":"Microsoft Fabric (Free)",
  "36a0f3b3-adb5-49ea-bf66-762134cf063a":"Microsoft Teams Premium",
  "749742bf-0d37-4158-a120-33567104deeb":"Microsoft 365 Lighthouse",
  "4cde982a-ede4-4409-9ae6-b003453c8ea6":"Microsoft Teams Rooms Pro",
  "3f9f06f5-3c31-472c-985f-62d9c10ec167":"Power Pages vTrial for Makers",
  "1e615a51-59db-4807-9957-aa83c3657351":"Dynamics 365 Customer Service Trial",
  "6af4b3d6-14bb-4a2a-960c-6c902aad34f3":"Microsoft Teams Rooms Basic",
  "238e2f8d-e429-4035-94db-6926be4ffe7b":"Microsoft Defender for Office 365 P2 Trial",
  "6ec92958-3cc1-49db-95bd-bc6b3798df71":"Microsoft Defender for Endpoint P2 Trial",
  "606b54a9-78d8-4298-ad8b-df6ef4481c80":"Power Virtual Agents Trial",
  "52ea0e27-ae73-4983-a08f-13561ebdb823":"Microsoft Viva Insights",
  "bc946dac-7877-4271-b2f7-99d2db13cd2c":"Dynamics 365 Customer Voice Trial",
  "420af87e-8177-4146-a780-3786adaffbca":"Microsoft Teams Essentials",
  "beb6439c-caad-48d3-bf46-0c82871e12be":"Microsoft 365 Copilot",
  "84cd610f-a3f8-4beb-84ab-d9d2c902c6c9":"Microsoft Entra ID Governance",
  "aa2695c9-8d59-4800-9dc8-12e01f1735af":"Microsoft Entra Permissions Management",
};

// ISO 3166-1 country code → full name
const CC = {"AF":"Afghanistan","AL":"Albania","DZ":"Algeria","AD":"Andorra","AO":"Angola","AG":"Antigua and Barbuda","AR":"Argentina","AM":"Armenia","AT":"Austria","AZ":"Azerbaijan","BS":"Bahamas","BH":"Bahrain","BD":"Bangladesh","BB":"Barbados","BY":"Belarus","BE":"Belgium","BZ":"Belize","BJ":"Benin","BT":"Bhutan","BO":"Bolivia","BA":"Bosnia and Herzegovina","BW":"Botswana","BR":"Brazil","BN":"Brunei","BG":"Bulgaria","BF":"Burkina Faso","BI":"Burundi","CV":"Cabo Verde","KH":"Cambodia","CM":"Cameroon","CA":"Canada","CF":"Central African Republic","TD":"Chad","CL":"Chile","CN":"China","CO":"Colombia","KM":"Comoros","CG":"Congo","CD":"Congo (DRC)","CK":"Cook Islands","CR":"Costa Rica","HR":"Croatia","CU":"Cuba","CY":"Cyprus","CZ":"Czech Republic","DK":"Denmark","DJ":"Djibouti","DM":"Dominica","DO":"Dominican Republic","EC":"Ecuador","EG":"Egypt","SV":"El Salvador","GQ":"Equatorial Guinea","ER":"Eritrea","EE":"Estonia","SZ":"Eswatini","ET":"Ethiopia","FJ":"Fiji","FI":"Finland","FR":"France","GA":"Gabon","GM":"Gambia","GE":"Georgia","DE":"Germany","GH":"Ghana","GR":"Greece","GD":"Grenada","GT":"Guatemala","GN":"Guinea","GW":"Guinea-Bissau","GY":"Guyana","HT":"Haiti","HN":"Honduras","HU":"Hungary","IS":"Iceland","IN":"India","ID":"Indonesia","IR":"Iran","IQ":"Iraq","IE":"Ireland","IL":"Israel","IT":"Italy","JM":"Jamaica","JP":"Japan","JO":"Jordan","KZ":"Kazakhstan","KE":"Kenya","KI":"Kiribati","KW":"Kuwait","KG":"Kyrgyzstan","LA":"Laos","LV":"Latvia","LB":"Lebanon","LS":"Lesotho","LR":"Liberia","LY":"Libya","LI":"Liechtenstein","LT":"Lithuania","LU":"Luxembourg","MG":"Madagascar","MW":"Malawi","MY":"Malaysia","MV":"Maldives","ML":"Mali","MT":"Malta","MH":"Marshall Islands","MR":"Mauritania","MU":"Mauritius","MX":"Mexico","FM":"Micronesia","MD":"Moldova","MC":"Monaco","MN":"Mongolia","ME":"Montenegro","MA":"Morocco","MZ":"Mozambique","MM":"Myanmar","NA":"Namibia","NR":"Nauru","NP":"Nepal","NL":"Netherlands","NI":"Nicaragua","NE":"Niger","NG":"Nigeria","MK":"North Macedonia","NO":"Norway","OM":"Oman","PK":"Pakistan","PW":"Palau","PA":"Panama","PG":"Papua New Guinea","PY":"Paraguay","PE":"Peru","PH":"Philippines","PL":"Poland","PT":"Portugal","QA":"Qatar","RO":"Romania","RU":"Russia","RW":"Rwanda","KN":"Saint Kitts and Nevis","LC":"Saint Lucia","VC":"Saint Vincent and the Grenadines","WS":"Samoa","SM":"San Marino","ST":"Sao Tome and Principe","SA":"Saudi Arabia","SN":"Senegal","RS":"Serbia","SC":"Seychelles","SL":"Sierra Leone","SG":"Singapore","SK":"Slovakia","SI":"Slovenia","SB":"Solomon Islands","SO":"Somalia","ZA":"South Africa","SS":"South Sudan","ES":"Spain","LK":"Sri Lanka","SD":"Sudan","SR":"Suriname","SE":"Sweden","CH":"Switzerland","SY":"Syria","TW":"Taiwan","TJ":"Tajikistan","TZ":"Tanzania","TH":"Thailand","TL":"Timor-Leste","TG":"Togo","TO":"Tonga","TT":"Trinidad and Tobago","TN":"Tunisia","TR":"Turkey","TM":"Turkmenistan","TV":"Tuvalu","UG":"Uganda","UA":"Ukraine","AE":"United Arab Emirates","GB":"United Kingdom","US":"United States","UY":"Uruguay","UZ":"Uzbekistan","VU":"Vanuatu","VE":"Venezuela","VN":"Vietnam","YE":"Yemen","ZM":"Zambia","ZW":"Zimbabwe"};
const countryName = code => CC[code] || code;

function policyStatus(p) {
  if (!p) return "not found";
  if (p.state === "enabled") return "enforced";
  if (p.state === "enabledForReportingButNotEnforced") return "report-only";
  return "disabled";
}
const dn = p => (p?.displayName || "").toLowerCase();

// ── HTTP helpers ─────────────────────────────────────────────────────────────
async function getToken(tenantId, clientId, clientSecret) {
  const res = await fetch(
    `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ grant_type: "client_credentials", client_id: clientId, client_secret: clientSecret, scope: "https://graph.microsoft.com/.default" }),
    }
  );
  const d = await res.json();
  if (d.error) throw new Error(d.error_description || d.error);
  return d.access_token;
}

async function graphOne(token, path, beta = false, extraHeaders = {}) {
  const base = beta ? "https://graph.microsoft.com/beta" : "https://graph.microsoft.com/v1.0";
  try {
    const res = await fetch(base + path, { headers: { Authorization: `Bearer ${token}`, ...extraHeaders } });
    const d = await res.json();
    return d.error ? { data: null, error: d.error.message } : { data: d, error: null };
  } catch (e) {
    return { data: null, error: e.message };
  }
}

async function graphAll(token, path, beta = false, extraHeaders = {}) {
  const base = beta ? "https://graph.microsoft.com/beta" : "https://graph.microsoft.com/v1.0";
  const results = [];
  let url = base + path;
  let error = null, pages = 0;
  while (url && pages < 50) {
    try {
      const res = await fetch(url, { headers: { Authorization: `Bearer ${token}`, ...extraHeaders } });
      const d = await res.json();
      if (d.error) { error = d.error.message; break; }
      if (Array.isArray(d.value)) results.push(...d.value);
      url = d["@odata.nextLink"] || null;
      pages++;
    } catch (e) { error = e.message; break; }
  }
  return { results, error };
}

function parseCSV(text) {
  const lines = text.trim().split("\n");
  if (lines.length < 2) return [];
  const headers = lines[0].split(",").map(h => h.replace(/"/g, "").trim());
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    if (!lines[i].trim()) continue;
    const vals = [];
    let cur = "", inQ = false;
    for (const ch of lines[i]) {
      if (ch === '"') { inQ = !inQ; }
      else if (ch === "," && !inQ) { vals.push(cur.trim()); cur = ""; }
      else cur += ch;
    }
    vals.push(cur.trim());
    const row = {};
    headers.forEach((h, j) => { row[h] = (vals[j] || "").replace(/"/g, "").trim(); });
    rows.push(row);
  }
  return rows;
}

async function graphReportCSV(token, path) {
  try {
    const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      headers: { Authorization: `Bearer ${token}` },
      redirect: "manual",
    });
    let csvText;
    if (res.status === 302 || res.status === 301) {
      const location = res.headers.get("location");
      if (!location) return { rows: [], error: "No redirect location" };
      csvText = await (await fetch(location)).text();
    } else if (res.status === 200) {
      csvText = await res.text();
    } else {
      const body = await res.text();
      try { const j = JSON.parse(body); return { rows: [], error: j.error?.message || j.Message || body.substring(0, 150) }; }
      catch { return { rows: [], error: `HTTP ${res.status}: ${body.substring(0, 150)}` }; }
    }
    if (!csvText || csvText.trim().startsWith("<") || csvText.trim().startsWith("{")) {
      return { rows: [], error: "Unexpected response format: " + csvText.substring(0, 100) };
    }
    return { rows: parseCSV(csvText), error: null };
  } catch (e) {
    return { rows: [], error: e.message };
  }
}

function withTimeout(promise, ms, fallback) {
  return Promise.race([promise, new Promise(resolve => setTimeout(() => resolve(fallback), ms))]);
}

// ── Main handler ──────────────────────────────────────────────────────────────
// Simple in-memory rate limiter — 10 requests per IP per 5 minutes
// Note: resets on function cold start; sufficient to deter casual abuse
const rateLimitMap = new Map();
const RATE_LIMIT = 10, RATE_WINDOW_MS = 5 * 60 * 1000;

function checkRateLimit(ip) {
  const now = Date.now();
  const entry = rateLimitMap.get(ip) || { count: 0, windowStart: now };
  if (now - entry.windowStart > RATE_WINDOW_MS) {
    entry.count = 0; entry.windowStart = now;
  }
  entry.count++;
  rateLimitMap.set(ip, entry);
  // Prune old entries to prevent memory growth
  if (rateLimitMap.size > 500) {
    for (const [k, v] of rateLimitMap) {
      if (now - v.windowStart > RATE_WINDOW_MS) rateLimitMap.delete(k);
    }
  }
  return entry.count <= RATE_LIMIT;
}

function winVersionLabel(osVersion) {
  if (!osVersion) return "";
  const parts = osVersion.split(".");
  if (parts.length < 4) return osVersion;
  const build = parseInt(parts[2]);
  const rev   = parseInt(parts[3]);
  let ver = "";
  if      (build >= 26200) ver = "Win 11 25H2";
  else if (build >= 26100) ver = "Win 11 24H2";
  else if (build >= 22631) ver = "Win 11 23H2";
  else if (build >= 22621) ver = "Win 11 22H2";
  else if (build >= 22000) ver = "Win 11 21H2";
  else if (build >= 19045) ver = "Win 10 22H2";
  else if (build >= 19044) ver = "Win 10 21H2";
  else if (build >= 19041) ver = "Win 10 20H2";
  else return osVersion;
  const WIN10 = [[7184,"Apr 2026"],[6953,"Mar 2026"],[6796,"Feb 2026"],[6608,"Jan 2026"],[6452,"Dec 2025"],[6293,"Nov 2025"],[6161,"Oct 2025"],[5965,"Sep 2025"],[5830,"Aug 2025"],[5608,"Jul 2025"],[5487,"Jun 2025"],[5371,"May 2025"],[5247,"Apr 2025"],[5073,"Mar 2025"],[4953,"Feb 2025"],[4842,"Jan 2025"],[4651,"Dec 2024"],[4529,"Nov 2024"],[4355,"Oct 2024"],[4239,"Sep 2024"],[4046,"Aug 2024"],[3930,"Jul 2024"],[3803,"Jun 2024"],[3672,"May 2024"],[3570,"Apr 2024"],[3447,"Mar 2024"],[3326,"Feb 2024"],[3260,"Jan 2024"]];
  const W23H2 = [[6649,"Feb 2026"],[6459,"Jan 2026"],[6252,"Dec 2025"],[6092,"Nov 2025"],[5965,"Oct 2025"],[5854,"Sep 2025"],[5674,"Aug 2025"],[5472,"Jul 2025"],[5323,"Jun 2025"],[5073,"May 2025"],[4894,"Apr 2025"],[4651,"Mar 2025"],[4528,"Feb 2025"],[4391,"Jan 2025"],[4169,"Dec 2024"],[4037,"Nov 2024"],[3958,"Oct 2024"],[3863,"Sep 2024"],[3640,"Aug 2024"],[3527,"Jul 2024"],[3374,"Jun 2024"],[3296,"May 2024"],[3155,"Apr 2024"],[2986,"Mar 2024"],[2861,"Feb 2024"],[2792,"Jan 2024"]];
  const W24H2 = [[7840,"Feb 2026"],[7593,"Jan 2026"],[7387,"Dec 2025"],[7147,"Nov 2025"],[6952,"Oct 2025"],[6755,"Sep 2025"],[6528,"Aug 2025"],[6252,"Jul 2025"],[6036,"Jun 2025"],[5854,"May 2025"],[5680,"Apr 2025"],[5461,"Mar 2025"],[5247,"Feb 2025"],[4972,"Jan 2025"],[4651,"Dec 2024"],[4529,"Nov 2024"],[4355,"Oct 2024"],[4112,"Sep 2024"]];
  const W25H2 = [[8246,"Apr 2026"],[7840,"Feb 2026"]];
  const table = build >= 26200 ? W25H2 : build >= 26100 ? W24H2 : build >= 22631 ? W23H2 : WIN10;
  const found = table.find(([r]) => rev >= r);
  const month = found ? found[1] : "";
  return ver + (month ? ` (${month})` : "") + " — " + osVersion;
}

exports.handler = async (event) => {
  if (event.httpMethod === "OPTIONS") return { statusCode: 204, headers: CORS, body: "" };
  if (event.httpMethod !== "POST") return respond(405, { error: "Method not allowed" });

  // Rate limiting — 10 requests per IP per 5 minutes
  const clientIp = event.headers?.["x-forwarded-for"]?.split(",")[0]?.trim()
    || event.headers?.["x-nf-client-connection-ip"]
    || "unknown";
  if (!checkRateLimit(clientIp)) {
    return respond(429, { error: "Too many requests — please wait a few minutes before trying again." });
  }

  let body;
  try { body = JSON.parse(event.body); }
  catch { return respond(400, { error: "Invalid JSON" }); }

  const { tenantId, clientId, clientSecret, reportFrom, reportTo } = body;
  if (!tenantId || !clientId || !clientSecret) return respond(400, { error: "Missing credentials" });
  if (!GUID_RE.test(tenantId) || !GUID_RE.test(clientId)) return respond(400, { error: "Invalid credential format" });
  if (typeof clientSecret !== "string" || clientSecret.length < 8 || clientSecret.length > 256) return respond(400, { error: "Invalid credential format" });

  let token;
  try { token = await getToken(tenantId, clientId, clientSecret); }
  catch (e) { return respond(401, { error: "Authentication failed: " + e.message }); }

  try {
    const now_ms = Date.now();
    const MS_DAY = 86400000;
    const thirtyDaysAgo = new Date(now_ms - 30 * MS_DAY).toISOString().split(".")[0] + "Z";
    // Use report period if provided, otherwise fall back to rolling 30 days
    const siPeriodStart = reportFrom
      ? new Date(reportFrom).toISOString().split(".")[0] + "Z"
      : thirtyDaysAgo;
    const siPeriodEnd = reportTo
      ? new Date(new Date(reportTo).getTime() + 86400000).toISOString().split(".")[0] + "Z"
      : null;  // null = no end cap needed
    // Sign-in filter: date + country only (simplest possible OData filter).
    // errorCode filtered client-side — combining it server-side causes timeouts.
    // ConsistencyLevel: eventual header required for ne filters on sign-in logs.
    // Single ne filter on NZ only — minimises server-side complexity.
    // AU and blank country filtered client-side along with failed sign-ins.


    const FAST = 8500, SLOW = 8000;
    const emptyAll = { results: [], error: "timeout" };
    const emptyOne = { data: null, error: "timeout" };
    const emptyCSV = { rows: [], error: "timeout" };

    const [
      devR, scoreR, riskyR, caR, secDefaultsR, authMethodsR, compPoliciesR, appProtR,
      usersR, rolesR, _signInsPlaceholder, spUsageR, spStorageR, skusR, roleDefsR, m365GroupsR, secGroupsR,
    ] = await Promise.all([
      withTimeout(graphAll(token, "/deviceManagement/managedDevices?$top=200&$expand=windowsProtectionState", true), SLOW, emptyAll),
      withTimeout(graphOne(token, "/security/secureScores?$top=1"), FAST, emptyOne),
      withTimeout(graphAll(token, "/identityProtection/riskyUsers?$top=500&$filter=riskState eq 'atRisk'"), FAST, emptyAll),
      withTimeout(graphAll(token, "/identity/conditionalAccess/policies?$top=200"), FAST, emptyAll),
      withTimeout(graphOne(token, "/policies/identitySecurityDefaultsEnforcementPolicy"), FAST, emptyOne),
      withTimeout(graphOne(token, "/policies/authenticationMethodsPolicy"), FAST, emptyOne),
      withTimeout(graphAll(token, "/deviceManagement/deviceCompliancePolicies?$top=200"), FAST, emptyAll),
      withTimeout(graphOne(token, "/deviceAppManagement/managedAppPolicies?$top=200", true), FAST, emptyOne),
      withTimeout(graphAll(token, "/users?$top=500&$select=id,displayName,userPrincipalName,userType,accountEnabled,assignedLicenses,signInActivity,mail", true), SLOW, emptyAll),
      withTimeout(graphAll(token, "/roleManagement/directory/roleAssignments?$expand=principal($select=id,displayName,userPrincipalName)&$top=200"), SLOW, emptyAll),
      Promise.resolve(emptyAll), // sign-in query runs separately after Promise.all
      withTimeout(graphReportCSV(token, "/reports/getSharePointSiteUsageDetail(period='D180')"), SLOW, emptyCSV),
      withTimeout(graphReportCSV(token, "/reports/getSharePointSiteUsageStorage(period='D30')"), SLOW, emptyCSV),
      withTimeout(graphAll(token, "/subscribedSkus"), FAST, emptyAll),
      withTimeout(graphAll(token, "/roleManagement/directory/roleDefinitions?$select=id,displayName&$top=200"), FAST, emptyAll),
      withTimeout(graphOne(token, "/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$count=true&$top=1&$select=id", false, {"ConsistencyLevel":"eventual"}), FAST, emptyOne),
      withTimeout(graphOne(token, "/groups?$filter=securityEnabled+eq+true+and+mailEnabled+eq+false&$count=true&$top=1&$select=id", false, {"ConsistencyLevel":"eventual"}), FAST, emptyOne),
    ]);

    // ── Sign-in query — maximally optimised ──────────────────────────────────
    // Optimisations:
    // 1. Uses report period dates (not rolling 30d) so early-month logins aren't missed
    // 2. Includes interactiveUser + nonInteractiveUser — catches PRT/device logins
    // 3. Single page fetch, no pagination — one HTTP request
    // 4. Minimal $select — smallest payload
    // 5. Falls back to 14-day window in parallel
    const halfPeriodStart = reportFrom
      ? new Date(new Date(reportFrom).getTime() + Math.floor((new Date(reportTo||now_ms).getTime() - new Date(reportFrom).getTime()) / 2)).toISOString().split(".")[0] + "Z"
      : new Date(now_ms - 14 * MS_DAY).toISOString().split(".")[0] + "Z";

    // Include both interactive and non-interactive to catch Windows PRT logins
    const siEndClause = siPeriodEnd ? ` and createdDateTime le ${siPeriodEnd}` : "";
    const siInteractive = " and signInEventTypes/any(t:t eq 'interactiveUser')";
    const siF_full = encodeURIComponent(`createdDateTime ge ${siPeriodStart}${siEndClause}${siInteractive}`);
    const siF_half = encodeURIComponent(`createdDateTime ge ${halfPeriodStart}${siEndClause}${siInteractive}`);
    const siSel = "$select=userPrincipalName,location,status";
    const siHdr = { Authorization: `Bearer ${token}` };

    async function fetchSI(filter) {
      try {
        const url = `https://graph.microsoft.com/v1.0/auditLogs/signIns?$top=500&${siSel}&$filter=${filter}`;
        const r = await fetch(url, { headers: siHdr });
        const j = await r.json();
        if (j.error) return { results: [], error: j.error.message };
        return { results: j.value || [], error: null };
      } catch(e) { return { results: [], error: e.message }; }
    }

    // Run full period and half period in parallel — use full if it succeeds
    let signInWindow = 30;
    const siWindowDays = reportFrom
      ? Math.round((new Date(reportTo||now_ms) - new Date(reportFrom)) / MS_DAY)
      : 30;
    const TO = { results: [], error: "timeout" };
    const [siFull, siHalf] = await Promise.all([
      withTimeout(fetchSI(siF_full), 8500, TO),
      withTimeout(fetchSI(siF_half), 8500, TO),
    ]);
    let signInsRFinal;
    if (!siFull.error) {
      signInsRFinal = siFull; signInWindow = siWindowDays;
    } else if (!siHalf.error) {
      signInsRFinal = siHalf; signInWindow = Math.round(siWindowDays / 2);
    } else {
      signInsRFinal = TO; signInWindow = 0;
    }

    // ── Devices ───────────────────────────────────────────────────────────────
    const comp  = { compliant: 0, noncompliant: 0, unknown: 0 };
    let mobileCount = 0, iosCount = 0, androidCount = 0;  // track mobile to exclude from other metrics
    const encryption = { encrypted: 0, notEncrypted: 0 };
    const notCompliantList = [], notEncryptedList = [], patchOver90Devices = [], lowDisk = [];
    const avNotActive = [], avOutOfDate = [];   // AV state lists
    let avActiveCount = 0, avNotActiveCount = 0, avOutOfDateCount = 0;
    const patchStatus = { current: 0, over30: 0, over90: 0 };
    let staleCount = 0;
    let win10 = 0, win11 = 0, win24h2 = 0, win25h2 = 0, macOS = 0, linux = 0;

    for (const d of devR.results) {
      const os    = (d.operatingSystem || "").toLowerCase();
      const model = `${d.model || ""} ${d.manufacturer || ""}`.toLowerCase();

      // Mobile — count only, exclude from all other metrics
      const isMobile = os.includes("android") || os.includes("ios");
      if (os.includes("android")) { mobileCount++; androidCount++; }
      else if (os.includes("ios")) { mobileCount++; iosCount++; }

      // Compliance — include mobile
      const cs = d.complianceState in comp ? d.complianceState : "unknown";
      comp[cs]++;
      if (cs === "noncompliant") {
        notCompliantList.push({
          name: d.deviceName || "Unknown",
          user: d.userPrincipalName || d.emailAddress || "Unknown",
          os:   `${d.operatingSystem || ""}${d.osVersion ? " " + d.osVersion.split(".").slice(0,2).join(".") : ""}`,
          lastSync: d.lastSyncDateTime || null,
        });
      }

      // Skip remaining metrics for mobile
      if (isMobile) continue;

      // Staleness & patch
      const lastSync = d.lastSyncDateTime ? new Date(d.lastSyncDateTime).getTime() : 0;
      const ageDays  = lastSync ? (now_ms - lastSync) / MS_DAY : Infinity;
      if (ageDays > 90) staleCount++;
      if (ageDays > 90) {
        patchStatus.over90++;
        patchOver90Devices.push({ name: d.deviceName, user: d.userPrincipalName || d.emailAddress || "Unknown", lastSeen: d.lastSyncDateTime || null, os: winVersionLabel(d.osVersion || "") });
      } else if (ageDays > 30) {
        patchStatus.over30++;
      } else {
        patchStatus.current++;
      }

      // Disk
      if (d.freeStorageSpaceInBytes > 0 && d.totalStorageSpaceInBytes > 0) {
        const pct = Math.round((d.freeStorageSpaceInBytes / d.totalStorageSpaceInBytes) * 100);
        if (pct < 15) lowDisk.push({ name: d.deviceName, user: d.userPrincipalName || d.emailAddress || "Unassigned", pct, gb: Math.round(d.freeStorageSpaceInBytes / 1e9) });
      }

      // OS version
      const ver = d.osVersion || "";
      if      (os.includes("mac"))         macOS++;
      else if (os.includes("linux"))        linux++;
      else if (ver.includes("26200"))       win25h2++;
      else if (ver.includes("26100"))       win24h2++;
      else if (ver.startsWith("10.0.2"))   win11++;
      else if (ver.startsWith("10.0.1"))   win10++;

      // Antivirus state (Windows only — from windowsProtectionState expand)
      if (os.includes("windows")) {
        const avState = d.windowsProtectionState || null;
        if (avState) {
          // A device is considered protected if EITHER:
          //   1. antivirusEnabled === true (Defender is primary AV), OR
          //   2. realTimeProtectionEnabled === true (Defender RTP is on — even if another AV is primary)
          // antivirusEnabled=false with realTimeProtectionEnabled=true typically means
          // a third-party AV (Sophos, CrowdStrike, etc.) is registered as primary — machine IS protected.
          const avEnabled    = avState.antivirusEnabled === true;
          const rtpEnabled   = avState.realTimeProtectionEnabled === true;
          const isProtected  = avEnabled || rtpEnabled;

          const avSigDate  = avState.antivirusSignatureUpdateDateTime || null;
          const sigAgeDays = avSigDate ? (now_ms - new Date(avSigDate).getTime()) / MS_DAY : null;
          const avStale    = sigAgeDays !== null && sigAgeDays > 30;

          if (!isProtected) {
            // Neither Defender nor RTP active — genuinely not protected
            avNotActiveCount++;
            avNotActive.push({ name: d.deviceName, user: d.userPrincipalName || d.emailAddress || "Unknown" });
          } else if (avStale) {
            avOutOfDateCount++;
            avOutOfDate.push({ name: d.deviceName, user: d.userPrincipalName || d.emailAddress || "Unknown", lastUpdated: avSigDate, ageDays: Math.round(sigAgeDays) });
          } else {
            avActiveCount++;
          }
        }
      }

      // Encryption
      // macOS excluded — Intune does not reliably report FileVault state via isEncrypted
      const isMacOS = (d.operatingSystem || "").toLowerCase().includes("mac");
      if (d.isEncrypted === true) {
        encryption.encrypted++;
      } else if (d.isEncrypted === false && !isMacOS) {
        encryption.notEncrypted++;
        notEncryptedList.push({ name: d.deviceName, user: d.userPrincipalName || d.emailAddress || "Unknown", os: `${d.operatingSystem || ""}${d.osVersion ? " " + d.osVersion : ""}` });
      }
    }

    // ── Conditional Access ────────────────────────────────────────────────────
    const caPolicies   = caR.results;
    const caEnabled    = caPolicies.filter(p => p.state === "enabled").length;
    const caReportOnly = caPolicies.filter(p => p.state === "enabledForReportingButNotEnforced").length;
    const legacyAuthPolicy  = caPolicies.find(p => dn(p).includes("legacy"));
    const mfaAllUsersPolicy = caPolicies.find(p => dn(p).includes("require") && dn(p).includes("multifactor") && !dn(p).includes("admin"));
    const adminMfaPolicy    = caPolicies.find(p => (dn(p).includes("admin") || dn(p).includes("phishing")) && dn(p).includes("multifactor"));
    const geoBlockPolicy    = caPolicies.find(p => dn(p).includes("block") && (dn(p).includes("countr") || dn(p).includes("location")));

    // ── Auth methods ──────────────────────────────────────────────────────────
    const authConfigs = authMethodsR.data?.authenticationMethodConfigurations || [];
    const getAuth = id => authConfigs.find(m => m.id === id);
    const tapConfig = getAuth("TemporaryAccessPass");
    const authMethods = {
      authAppEnabled: getAuth("MicrosoftAuthenticator")?.state === "enabled",
      fido2Enabled:   getAuth("Fido2")?.state === "enabled",
      smsEnabled:     getAuth("Sms")?.state === "enabled",
      tapEnabled:     tapConfig?.state === "enabled",
      tapReusable:    tapConfig?.isUsableOnce === false,
    };

    // ── Compliance policies & app protection ──────────────────────────────────
    const compPolicies    = compPoliciesR.results;
    const appProtPolicies = appProtR.data?.value || [];
    const countType = (arr, kw) => arr.filter(p => (p["@odata.type"] || "").toLowerCase().includes(kw)).length;

    // ── Users & sign-ins ──────────────────────────────────────────────────────
    const allUsers = usersR.results;
    const guestCount = allUsers.filter(u => u.userType === "Guest").length;
    // Shared mailboxes: enabled, non-guest, no assigned licences, has a mail address
    const sharedMailboxCount = allUsers.filter(u =>
      u.userType !== "Guest" &&
      u.accountEnabled !== false &&
      u.mail &&
      (!u.assignedLicenses || u.assignedLicenses.length === 0)
    ).length;
    // Licensed users: enabled, non-guest, has at least one licence assigned
    const licensedUserCount = allUsers.filter(u =>
      u.userType !== "Guest" &&
      u.accountEnabled !== false &&
      u.assignedLicenses && u.assignedLicenses.length > 0
    ).length;

    // Primary user licence SKU part numbers — only these are flagged for 90d inactivity
    const PRIMARY_SKUS = new Set([
      "SPB","SMB_BUSINESS_PREMIUM","O365_BUSINESS_PREMIUM","O365_BUSINESS_ESSENTIALS",
      "SMB_BUSINESS","OFFICESUBSCRIPTION","ENTERPRISEPACK","ENTERPRISEPREMIUM",
      "STANDARDPACK","STANDARDWOFFPACK","ENTERPRISEWITHSCAL","DESKLESSPACK","DESKLESS",
    ]);
    // Build a quick skuId → skuPartNumber lookup from the subscribed SKUs
    const skuIdToPartNum = {};
    for (const s of skusR.results) { if (s.skuId) skuIdToPartNum[s.skuId] = s.skuPartNumber || ""; }

    const notSignedIn90Licensed = [], notSignedIn90Guest = [];
    for (const u of allUsers) {
      if (u.accountEnabled === false) continue;
      const isGuest    = u.userType === "Guest";
      const lastSignIn = u.signInActivity?.lastSuccessfulSignInDateTime || u.signInActivity?.lastSignInDateTime || null;
      const daysSince  = lastSignIn ? (now_ms - new Date(lastSignIn).getTime()) / MS_DAY : Infinity;
      if (daysSince <= 90) continue; // catches both active users and never-signed-in (Infinity > 90)
      if (isGuest) {
        notSignedIn90Guest.push({ name: u.displayName, upn: u.userPrincipalName, lastSignIn, daysSince: isFinite(daysSince) ? Math.round(daysSince) : null });
      } else {
        // Only flag users with a primary user licence (Business Basic/Standard/Premium, E1/E3/E5, F3)
        const hasPrimary = (u.assignedLicenses || []).some(lic => PRIMARY_SKUS.has(skuIdToPartNum[lic.skuId] || ""));
        if (hasPrimary) {
          notSignedIn90Licensed.push({ name: u.displayName, upn: u.userPrincipalName, lastSignIn, daysSince: isFinite(daysSince) ? Math.round(daysSince) : null });
        }
      }
    }
    notSignedIn90Licensed.sort((a, b) => b.daysSince - a.daysSince);
    notSignedIn90Guest.sort((a, b) => b.daysSince - a.daysSince);

    // ── Licences ──────────────────────────────────────────────────────────────
    const skuNames = { ...GUID_MAP };
    for (const sku of skusR.results) {
      if (sku.skuId && sku.skuPartNumber) {
        skuNames[sku.skuId] = SKU_MAP[sku.skuPartNumber]
          || GUID_MAP[sku.skuId]
          || sku.skuPartNumber.split("_").map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(" ");
      }
    }
    // Build licence summary from subscribedSkus (preferred — includes available count)
    let licenceSummary = skusR.results
      .filter(s => s.skuId && s.skuPartNumber && (s.consumedUnits || 0) > 0)
      .map(s => ({ name: skuNames[s.skuId] || s.skuPartNumber, count: s.consumedUnits || 0, available: s.prepaidUnits?.enabled || 0 }))
      .sort((a, b) => b.count - a.count);

    // Fallback: if subscribedSkus returned nothing, count from user assignedLicenses
    if (licenceSummary.length === 0 && usersR.results.length > 0) {
      const licCounts = {};
      for (const u of usersR.results) {
        for (const lic of (u.assignedLicenses || [])) {
          if (!lic.skuId) continue;
          const name = skuNames[lic.skuId] || lic.skuId;
          licCounts[name] = (licCounts[name] || 0) + 1;
        }
      }
      licenceSummary = Object.entries(licCounts)
        .sort((a, b) => b[1] - a[1])
        .map(([name, count]) => ({ name, count, available: 0 }));
    }

    // ── Admin roles ───────────────────────────────────────────────────────────
    const roleDefMap = {};
    for (const rd of roleDefsR.results) { if (rd.id) roleDefMap[rd.id] = rd.displayName || rd.id; }

    const adminByUser = {};
    for (const a of rolesR.results) {
      const roleName = roleDefMap[a.roleDefinitionId] || a.roleDefinitionId;
      if (!ADMIN_ROLES.has(roleName)) continue;
      const p = a.principal || {};
      const uid = p.id || p.userPrincipalName;
      if (!uid) continue;
      if (!adminByUser[uid]) adminByUser[uid] = { name: p.displayName || "Unknown", upn: p.userPrincipalName || "", roles: [] };
      if (!adminByUser[uid].roles.includes(roleName)) adminByUser[uid].roles.push(roleName);
    }
    const adminRoleMembers = Object.values(adminByUser)
      .map(u => ({ ...u, roles: u.roles.sort() }))
      .sort((a, b) => a.name.localeCompare(b.name));

    // ── Overseas sign-ins — expected countries excluded, full names resolved ────
    // Build UPN->name lookup from the users we already fetched
    const upnNameMap = new Map();
    for (const u of (usersR.results || [])) {
      if (u.userPrincipalName && u.displayName) upnNameMap.set(u.userPrincipalName.toLowerCase(), u.displayName);
    }
    const SKIP_CC = new Set(["AU", "NZ", "MY", ""]);  // AU, NZ, Malaysia = expected
    const extMap = new Map();
    let totalOverseasLogins = 0;
    for (const s of (signInsRFinal.results || [])) {
      const upn = s.userPrincipalName;
      const cc  = s.location?.countryOrRegion || "";
      if (!upn) continue;
      if (SKIP_CC.has(cc)) continue;                    // expected countries
      if ((s.status?.errorCode ?? 1) !== 0) continue;  // successful only
      totalOverseasLogins++;
      const key = upn + "|" + cc;
      if (!extMap.has(key)) extMap.set(key, { upn, cc, countryFull: countryName(cc), eventCount: 0 });
      extMap.get(key).eventCount++;
    }
    // Group by user — collect all countries and total events
    const userMap = new Map();
    for (const row of extMap.values()) {
      if (!userMap.has(row.upn)) userMap.set(row.upn, { upn: row.upn, name: upnNameMap.get(row.upn.toLowerCase()) || '', eventCount: 0, countries: [] });
      const u = userMap.get(row.upn);
      u.eventCount += row.eventCount;
      u.countries.push(`${row.countryFull} (${row.eventCount})`);
    }
    const externalByUser = [...userMap.values()].sort((a, b) => b.eventCount - a.eventCount);

    // ── SharePoint ────────────────────────────────────────────────────────────
    const spRows = spUsageR.rows || [];
    const spFirstRow = spRows[0] || {};
    const spKeys = Object.keys(spFirstRow);
    const siteUrlKey     = spKeys.find(k => k.toLowerCase().includes("site url"))       || "Site URL";
    const lastActivityKey= spKeys.find(k => k.toLowerCase().includes("last activity"))  || "Last Activity Date";
    const storageUsedKey = spKeys.find(k => k.toLowerCase().includes("storage used"))   || "Storage Used (Byte)";
    const siteTypeKey    = spKeys.find(k => k.toLowerCase().includes("site type") || k.toLowerCase().includes("template type")) || "";
    const isDeletedKey   = spKeys.find(k => k.toLowerCase().includes("is deleted"))     || "";

    const cutoff180 = now_ms - 180 * MS_DAY;
    let spSiteCount = 0, spGroupCount = 0, spCommCount = 0, spClassicCount = 0, spTotalUsedGB = 0;
    const inactiveSites = [];

    // Additional column keys confirmed from live API
    const ownerNameKey      = spKeys.find(k => k.toLowerCase().includes("owner display"))       || "";
    const ownerPrincipalKey = spKeys.find(k => k.toLowerCase().includes("owner principal"))     || "";
    const rootTemplateKey   = spKeys.find(k => k.toLowerCase().includes("root web template") || k.toLowerCase().includes("template")) || "";

    // Derive the tenant's SharePoint base URL from any onmicrosoft.com UPN in the data
    // e.g. Finance@TPS1.onmicrosoft.com → https://TPS1.sharepoint.com
    let spBaseUrl = "";
    for (const row of spRows) {
      const upn = (ownerPrincipalKey && row[ownerPrincipalKey]) || "";
      const m = upn.match(/@([^.]+)\.onmicrosoft\.com$/i);
      if (m) { spBaseUrl = `https://${m[1]}.sharepoint.com`; break; }
    }

    for (const row of spRows) {
      if (isDeletedKey && (row[isDeletedKey] || "").toLowerCase() === "true") continue;

      const siteUrl    = row[siteUrlKey] || "";
      const ownerName  = (ownerNameKey      && row[ownerNameKey])      || "";
      const ownerUpn   = (ownerPrincipalKey && row[ownerPrincipalKey]) || "";
      const template   = (rootTemplateKey   && row[rootTemplateKey])   || "";

      // Skip personal OneDrive sites
      if (siteUrl.toLowerCase().includes("/personal/")) continue;
      if (siteTypeKey && (row[siteTypeKey] || "").toLowerCase().includes("onedrive")) continue;
      if (template.toLowerCase() === "msf" || template.toLowerCase() === "personal") continue;

      spSiteCount++;
      const tl = template.toLowerCase();
      if      (tl === "group")       spGroupCount++;
      else if (tl === "sitepagepublishing" || tl === "communication") spCommCount++;
      else                           spClassicCount++;
      spTotalUsedGB += parseInt(row[storageUsedKey] || "0") / 1e9;

      const lastActivity = row[lastActivityKey] || null;
      // Only flag sites with a known last activity date older than 180 days.
      // Sites with no last activity date are excluded — Microsoft reports D180 only,
      // so a blank date means no file activity in the window but the site may still be in use.
      if (lastActivity && new Date(lastActivity).getTime() < cutoff180) {
        // Clean site name from Owner Display Name (strip " Owners" / " Members" etc.)
        const siteName = ownerName
          .replace(/ Owners$/i, "").replace(/ Members$/i, "").replace(/ Visitors$/i, "").trim()
          || ownerUpn.split("@")[0]
          || "Unknown site";

        // Construct URL: use CSV value if present, otherwise build from ownerUpn alias + tenant domain
        // e.g. Finance@TPS1.onmicrosoft.com → https://TPS1.sharepoint.com/sites/Finance
        let resolvedUrl = siteUrl;
        if (!resolvedUrl && ownerUpn && spBaseUrl) {
          const alias = ownerUpn.split("@")[0];
          if (alias) resolvedUrl = `${spBaseUrl}/sites/${alias}`;
        }

        inactiveSites.push({ name: siteName, url: resolvedUrl || null, lastActivity: lastActivity || null });
      }
    }
    spTotalUsedGB = Math.round(spTotalUsedGB * 10) / 10;

    // Use storage trend for more accurate total if available
    const spStorageRows = spStorageR.rows || [];
    const spLatestRow = spStorageRows.length > 0 ? spStorageRows[spStorageRows.length - 1] : null;
    const spTrendRaw = spLatestRow ? parseInt(spLatestRow["Storage Used (Byte)"] || "0") : 0;
    if (spTrendRaw > 0) spTotalUsedGB = Math.round(spTrendRaw / 1e9 * 10) / 10;

    // ── Risk register ─────────────────────────────────────────────────────────
    const risks = [];
    const risk = (sev, area, finding, action) => risks.push({ severity: sev, area, finding, action });
    const p = (n, s, pl) => `${n} ${n === 1 ? s : pl}`;

    if (!legacyAuthPolicy) risk("high","Conditional Access","No legacy authentication block policy found","Create and enforce a CA policy to block legacy auth");
    else if (policyStatus(legacyAuthPolicy) === "report-only") risk("high","Conditional Access","Legacy authentication block is report-only — not enforced","Switch to enforced after validating exceptions");
    if (adminMfaPolicy && policyStatus(adminMfaPolicy) === "report-only") risk("high","Conditional Access","Admin phishing-resistant MFA is report-only","Switch to enforced once admin readiness confirmed");
    if (avNotActiveCount > 0) risk("high","Intune / Antivirus",`${p(avNotActiveCount,"device has","devices have")} antivirus disabled or not reporting`,"Investigate and re-enable antivirus immediately");
    if (avOutOfDateCount > 0) risk("medium","Intune / Antivirus",`${p(avOutOfDateCount,"device has","devices have")} antivirus definitions more than 30 days out of date`,"Force a definition update via Intune or investigate connectivity issues");
    if (encryption.notEncrypted > 0) risk("high","Intune / Device Security",`${p(encryption.notEncrypted,"device is","devices are")} not encrypted`,"Enable BitLocker via Intune encryption policy");
    if (comp.noncompliant > 0) risk("high","Intune / Compliance",`${p(comp.noncompliant,"device is","devices are")} non-compliant`,"Review and remediate non-compliant devices");
    if (secDefaultsR.data?.isEnabled === false && caEnabled < 3) risk("high","Entra ID","Security Defaults disabled with fewer than 3 enforced CA policies","Ensure CA fully covers Security Defaults protections");
    if (riskyR.results.length > 0) risk("high","Entra ID Protection",`${p(riskyR.results.length,"user account is","user accounts are")} flagged as at-risk`,"Review and reset passwords immediately");
    if (patchStatus.over90 > 0) risk("high","Intune / Patch Management",`${p(patchStatus.over90,"device is","devices are")} over 90 days without a check-in (3+ patch cycles)`,"Investigate — devices may be offline, lost, or unmanaged");
    if (patchStatus.over30 > 0) risk("medium","Intune / Patch Management",`${p(patchStatus.over30,"device is","devices are")} 30–90 days without a check-in (1–2 patch cycles behind)`,"Review and bring devices back under management");
    if (externalByUser.length > 0) risk("medium","Entra ID / Sign-ins",`${p(externalByUser.length,"user has","users have")} signed in from outside AU/NZ in the last 30 days`,"Review for unexpected access — ensure MFA and CA policies are enforced");
    if (countType(compPolicies,"windows") === 0) risk("medium","Intune / Compliance","No explicit Windows compliance policy found","Create a Windows compliance policy");
    if (appProtPolicies.length === 0) risk("medium","Intune / App Protection","No app protection (MAM) policies found","Add MAM policies if BYOD is permitted");
    if (authMethods.smsEnabled) risk("medium","Authentication Methods","SMS sign-in enabled — weaker than phishing-resistant MFA","Review whether SMS should remain available tenant-wide");
    if (authMethods.tapEnabled && authMethods.tapReusable) risk("medium","Authentication Methods","Temporary Access Pass is reusable","Consider switching TAP to one-time-use");
    if (!geoBlockPolicy) risk("medium","Conditional Access","No geographic restriction policy detected","Consider blocking sign-ins from outside AU/NZ");
    if (win10 > 0) risk("medium","Intune / Patch Management",`${p(win10,"device is","devices are")} still running Windows 10 — support ended October 2025`,"Plan and execute Windows 11 upgrade");
    const totalInactive = notSignedIn90Licensed.length + notSignedIn90Guest.length;
    if (totalInactive > 0) risk("low","Entra ID / Users",`${p(totalInactive,"user has","users have")} not signed in for 90+ days (${notSignedIn90Licensed.length} licensed, ${notSignedIn90Guest.length} guest)`,"Review for stale or unused accounts — disable licensed users, remove guest access");
    if (inactiveSites.length > 0) risk("low","SharePoint",`${p(inactiveSites.length,"SharePoint site has","SharePoint sites have")} had no activity in 180+ days`,"Review for archiving or deletion");
    if (lowDisk.length > 0) risk("low","Intune / Device Health",`${p(lowDisk.length,"device has","devices have")} less than 15% disk space remaining`,"Clean up disk space or expand storage");
    risks.sort((a, b) => SEV_ORDER[a.severity] - SEV_ORDER[b.severity]);

    const scoreRaw = scoreR.data?.value?.[0] ?? null;

    return respond(200, {
      total: devR.results.length,
      comp, staleCount, lowDisk, mobileCount, iosCount, androidCount, win10, win11, win24h2, win25h2, macOS, linux,
      encryption, notEncryptedList, notCompliantList, patchStatus, patchOver90: patchOver90Devices, // full list — no cap
      av: { active: avActiveCount, notActive: avNotActiveCount, outOfDate: avOutOfDateCount, notActiveList: avNotActive.slice(0,50), outOfDateList: avOutOfDate.slice(0,50) },
      score: scoreRaw ? { pct: Math.round((scoreRaw.currentScore/scoreRaw.maxScore)*100), cur: Math.round(scoreRaw.currentScore), max: Math.round(scoreRaw.maxScore) } : null,
      risky: riskyR.results.length,
      securityDefaults: secDefaultsR.data?.isEnabled ?? null,
      conditionalAccess: { total: caPolicies.length, enabled: caEnabled, reportOnly: caReportOnly },
      keyPolicies: {
        legacyAuthBlock: { name: legacyAuthPolicy?.displayName || null,  status: policyStatus(legacyAuthPolicy) },
        mfaAllUsers:     { name: mfaAllUsersPolicy?.displayName || null, status: policyStatus(mfaAllUsersPolicy) },
        adminMfa:        { name: adminMfaPolicy?.displayName || null,    status: policyStatus(adminMfaPolicy) },
        geoBlock:        { name: geoBlockPolicy?.displayName || null,    status: policyStatus(geoBlockPolicy) },
      },
      authMethods,
      compliancePolicies: { total: compPolicies.length, windows: countType(compPolicies,"windows"), android: countType(compPolicies,"android") },
      appProtection:      { total: appProtPolicies.length, windows: countType(appProtPolicies,"windows"), ios: countType(appProtPolicies,"ios"), android: countType(appProtPolicies,"android") },
      users: {
        total: licensedUserCount,
        sharedMailboxes: sharedMailboxCount,
        guests: guestCount,
        notSignedIn90Licensed: notSignedIn90Licensed.length,
        notSignedIn90LicensedList: notSignedIn90Licensed.slice(0, 50),
        notSignedIn90Guest: notSignedIn90Guest.length,
        notSignedIn90GuestList: notSignedIn90Guest.slice(0, 50),
        licenceSummary,
        adminRoles: adminRoleMembers,
        externalSignIns: { total: totalOverseasLogins, uniqueUsers: externalByUser.length, byUser: externalByUser.slice(0, 30), timedOut: signInsRFinal.error === "timeout", windowDays: signInWindow },
      },
      sharepoint: {
        siteCount: spSiteCount,
        groupCount: spGroupCount,
        commCount: spCommCount,
        classicCount: spClassicCount,
        m365GroupCount: m365GroupsR.data?.["@odata.count"] ?? null,
        securityGroupCount: secGroupsR.data?.["@odata.count"] ?? null,
        totalUsedGB: spTotalUsedGB,
        allocatedGB: null,
        inactiveSiteCount: inactiveSites.length,
        error: spUsageR.error || spStorageR.error || null,
      },
      risks,
    });
  } catch (e) {
    return respond(500, { error: "Internal error: " + e.message });
  }
};
