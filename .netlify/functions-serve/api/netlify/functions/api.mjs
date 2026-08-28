
import {createRequire as ___nfyCreateRequire} from "module";
import {fileURLToPath as ___nfyFileURLToPath} from "url";
import {dirname as ___nfyPathDirname} from "path";
let __filename=___nfyFileURLToPath(import.meta.url);
let __dirname=___nfyPathDirname(___nfyFileURLToPath(import.meta.url));
let require=___nfyCreateRequire(import.meta.url);


// netlify/functions/api.mjs
import { getStore } from "@netlify/blobs";
import XLSX from "xlsx";
import crypto from "node:crypto";
var USERS_KEY = "ledgerly-users.xlsx";
var EXPENSES_KEY = "ledgerly-expenses.xlsx";
var headers = { "content-type": "application/json", "cache-control": "no-store" };
var usersStore = () => getStore({ name: "ledgerly-users", consistency: "strong" });
var expensesStore = () => getStore({ name: "ledgerly-expenses", consistency: "strong" });
var json = (statusCode, body) => new Response(JSON.stringify(body), { status: statusCode, headers });
var readWorkbook = async (store, key) => {
  const data = await store.get(key, { type: "arrayBuffer" });
  return data ? XLSX.read(data, { type: "array" }) : XLSX.utils.book_new();
};
var writeWorkbook = async (store, key, workbook) => store.set(key, XLSX.write(workbook, { bookType: "xlsx", type: "array" }));
var safeSheet = (profile) => profile.replace(/[\\/?*\[\]:]/g, " ").trim().slice(0, 31) || "User";
var rowsFor = (workbook, profile) => XLSX.utils.sheet_to_json(workbook.Sheets[safeSheet(profile)] || {}).map((row) => ({ ...row, Profile: profile }));
var profilesFrom = (workbook) => XLSX.utils.sheet_to_json(workbook.Sheets.Profiles || {}).map((row) => String(row["Profile Name"] || "")).filter(Boolean);
var authRows = (workbook) => XLSX.utils.sheet_to_json(workbook.Sheets.Profiles || {});
var passwordHash = (password, salt = crypto.randomBytes(16).toString("hex")) => `${salt}:${crypto.scryptSync(password, salt, 64).toString("hex")}`;
var passwordMatches = (password, stored) => {
  const [salt, digest] = String(stored || "").split(":");
  if (!salt || !digest) return false;
  const actual = crypto.scryptSync(password, salt, 64);
  return crypto.timingSafeEqual(actual, Buffer.from(digest, "hex"));
};
var api_default = async (request) => {
  try {
    const url = new URL(request.url);
    const method = request.method || request.httpMethod;
    if (url.pathname.endsWith("/health")) return json(200, { ok: true });
    if (method === "GET" && url.pathname.endsWith("/profiles")) {
      const wb = await readWorkbook(usersStore(), USERS_KEY);
      return json(200, { profiles: profilesFrom(wb) });
    }
    if (method === "POST" && url.pathname.endsWith("/auth/register")) {
      const body = await request.json();
      const username = String(body.username || "").trim();
      const password = String(body.password || "");
      if (!/^[a-zA-Z0-9_.-]{3,40}$/.test(username) || password.length < 8) return json(400, { error: "Use a username with 3-40 letters/numbers and a password of at least 8 characters." });
      const store = usersStore();
      const wb = await readWorkbook(store, USERS_KEY);
      const rows = authRows(wb);
      if (rows.some((row) => String(row.Username || "").toLowerCase() === username.toLowerCase())) return json(409, { error: "That username is already registered." });
      rows.push({ Username: username, "Password Hash": passwordHash(password), "Profile Name": username, "Created Date": (/* @__PURE__ */ new Date()).toISOString() });
      wb.Sheets.Profiles = XLSX.utils.json_to_sheet(rows);
      if (!wb.SheetNames.includes("Profiles")) XLSX.utils.book_append_sheet(wb, wb.Sheets.Profiles, "Profiles");
      await writeWorkbook(store, USERS_KEY, wb);
      return json(201, { profile: username, profiles: profilesFrom(wb) });
    }
    if (method === "POST" && url.pathname.endsWith("/auth/login")) {
      const body = await request.json();
      const username = String(body.username || "").trim();
      const password = String(body.password || "");
      const wb = await readWorkbook(usersStore(), USERS_KEY);
      const user = authRows(wb).find((row) => String(row.Username || "").toLowerCase() === username.toLowerCase());
      if (!user || !passwordMatches(password, user["Password Hash"])) return json(401, { error: "Invalid username or password." });
      return json(200, { profile: String(user["Profile Name"] || user.Username), profiles: profilesFrom(wb) });
    }
    if (method === "POST" && url.pathname.endsWith("/profiles")) {
      const profile2 = String((await request.json()).profile || "").trim();
      if (!profile2) return json(400, { error: "A profile is required." });
      const store = usersStore();
      const wb = await readWorkbook(store, USERS_KEY);
      const profiles = profilesFrom(wb);
      if (!profiles.includes(profile2)) {
        const rows = [...profiles, profile2].map((name) => ({ "Profile Name": name, "Created Date": (/* @__PURE__ */ new Date()).toISOString() }));
        wb.Sheets.Profiles = XLSX.utils.json_to_sheet(rows);
        if (!wb.SheetNames.includes("Profiles")) XLSX.utils.book_append_sheet(wb, wb.Sheets.Profiles, "Profiles");
        await writeWorkbook(store, USERS_KEY, wb);
      }
      return json(200, { profile: profile2, profiles: [.../* @__PURE__ */ new Set([...profiles, profile2])] });
    }
    const payload = method === "GET" ? Object.fromEntries(url.searchParams) : await request.json();
    const profile = String(payload.profile || "").trim();
    if (method === "PUT" && url.pathname.endsWith("/settings")) {
      const store = usersStore();
      const wb = await readWorkbook(store, USERS_KEY);
      const rows = XLSX.utils.sheet_to_json(wb.Sheets.Settings || {}).filter((row) => row.Profile !== profile);
      rows.push({ Profile: profile, Categories: payload.categories || "", "Default Currency": payload.defaultCurrency || "LKR", "Theme Preference": payload.theme || "dark" });
      wb.Sheets.Settings = XLSX.utils.json_to_sheet(rows);
      if (!wb.SheetNames.includes("Settings")) XLSX.utils.book_append_sheet(wb, wb.Sheets.Settings, "Settings");
      await writeWorkbook(store, USERS_KEY, wb);
      return json(200, { settings: payload });
    }
    if (!profile) return json(400, { error: "A profile is required." });
    if (method === "GET") {
      const wb = await readWorkbook(expensesStore(), EXPENSES_KEY);
      return json(200, { expenses: rowsFor(wb, profile) });
    }
    if (method === "POST" || method === "PUT" || method === "DELETE") {
      const store = expensesStore();
      const wb = await readWorkbook(store, EXPENSES_KEY);
      const sheet = safeSheet(profile);
      const rows = rowsFor(wb, profile);
      let next = rows;
      if (method === "POST") next = [...rows, payload.expense];
      if (method === "PUT") next = rows.map((row) => row.ID === payload.expense.ID ? payload.expense : row);
      if (method === "DELETE") next = rows.filter((row) => row.ID !== payload.id);
      const clean = next.map((row) => {
        const copy = { ...row };
        delete copy.Profile;
        return copy;
      });
      wb.Sheets[sheet] = XLSX.utils.json_to_sheet(clean);
      if (!wb.SheetNames.includes(sheet)) XLSX.utils.book_append_sheet(wb, wb.Sheets[sheet], sheet);
      await writeWorkbook(store, EXPENSES_KEY, wb);
      return json(200, { expenses: next });
    }
    return json(405, { error: "Method not allowed." });
  } catch (error) {
    return json(500, { error: error.message || "Storage request failed." });
  }
};
export {
  api_default as default
};
//# sourceMappingURL=data:application/json;base64,ewogICJ2ZXJzaW9uIjogMywKICAic291cmNlcyI6IFsibmV0bGlmeS9mdW5jdGlvbnMvYXBpLm1qcyJdLAogICJzb3VyY2VzQ29udGVudCI6IFsiaW1wb3J0IHsgZ2V0U3RvcmUgfSBmcm9tICdAbmV0bGlmeS9ibG9icyc7XG5pbXBvcnQgWExTWCBmcm9tICd4bHN4JztcbmltcG9ydCBjcnlwdG8gZnJvbSAnbm9kZTpjcnlwdG8nO1xuXG5jb25zdCBVU0VSU19LRVkgPSAnbGVkZ2VybHktdXNlcnMueGxzeCc7XG5jb25zdCBFWFBFTlNFU19LRVkgPSAnbGVkZ2VybHktZXhwZW5zZXMueGxzeCc7XG5jb25zdCBoZWFkZXJzID0geyAnY29udGVudC10eXBlJzogJ2FwcGxpY2F0aW9uL2pzb24nLCAnY2FjaGUtY29udHJvbCc6ICduby1zdG9yZScgfTtcbmNvbnN0IHVzZXJzU3RvcmUgPSAoKSA9PiBnZXRTdG9yZSh7IG5hbWU6ICdsZWRnZXJseS11c2VycycsIGNvbnNpc3RlbmN5OiAnc3Ryb25nJyB9KTtcbmNvbnN0IGV4cGVuc2VzU3RvcmUgPSAoKSA9PiBnZXRTdG9yZSh7IG5hbWU6ICdsZWRnZXJseS1leHBlbnNlcycsIGNvbnNpc3RlbmN5OiAnc3Ryb25nJyB9KTtcbmNvbnN0IGpzb24gPSAoc3RhdHVzQ29kZSwgYm9keSkgPT4gbmV3IFJlc3BvbnNlKEpTT04uc3RyaW5naWZ5KGJvZHkpLCB7IHN0YXR1czogc3RhdHVzQ29kZSwgaGVhZGVycyB9KTtcbmNvbnN0IHJlYWRXb3JrYm9vayA9IGFzeW5jIChzdG9yZSwga2V5KSA9PiB7IGNvbnN0IGRhdGEgPSBhd2FpdCBzdG9yZS5nZXQoa2V5LCB7IHR5cGU6ICdhcnJheUJ1ZmZlcicgfSk7IHJldHVybiBkYXRhID8gWExTWC5yZWFkKGRhdGEsIHsgdHlwZTogJ2FycmF5JyB9KSA6IFhMU1gudXRpbHMuYm9va19uZXcoKTsgfTtcbmNvbnN0IHdyaXRlV29ya2Jvb2sgPSBhc3luYyAoc3RvcmUsIGtleSwgd29ya2Jvb2spID0+IHN0b3JlLnNldChrZXksIFhMU1gud3JpdGUod29ya2Jvb2ssIHsgYm9va1R5cGU6ICd4bHN4JywgdHlwZTogJ2FycmF5JyB9KSk7XG5jb25zdCBzYWZlU2hlZXQgPSBwcm9maWxlID0+IHByb2ZpbGUucmVwbGFjZSgvW1xcXFwvPypcXFtcXF06XS9nLCAnICcpLnRyaW0oKS5zbGljZSgwLCAzMSkgfHwgJ1VzZXInO1xuY29uc3Qgcm93c0ZvciA9ICh3b3JrYm9vaywgcHJvZmlsZSkgPT4gWExTWC51dGlscy5zaGVldF90b19qc29uKHdvcmtib29rLlNoZWV0c1tzYWZlU2hlZXQocHJvZmlsZSldIHx8IHt9KS5tYXAocm93ID0+ICh7IC4uLnJvdywgUHJvZmlsZTogcHJvZmlsZSB9KSk7XG5jb25zdCBwcm9maWxlc0Zyb20gPSB3b3JrYm9vayA9PiBYTFNYLnV0aWxzLnNoZWV0X3RvX2pzb24od29ya2Jvb2suU2hlZXRzLlByb2ZpbGVzIHx8IHt9KS5tYXAocm93ID0+IFN0cmluZyhyb3dbJ1Byb2ZpbGUgTmFtZSddIHx8ICcnKSkuZmlsdGVyKEJvb2xlYW4pO1xuY29uc3QgYXV0aFJvd3MgPSB3b3JrYm9vayA9PiBYTFNYLnV0aWxzLnNoZWV0X3RvX2pzb24od29ya2Jvb2suU2hlZXRzLlByb2ZpbGVzIHx8IHt9KTtcbmNvbnN0IHBhc3N3b3JkSGFzaCA9IChwYXNzd29yZCwgc2FsdCA9IGNyeXB0by5yYW5kb21CeXRlcygxNikudG9TdHJpbmcoJ2hleCcpKSA9PiBgJHtzYWx0fToke2NyeXB0by5zY3J5cHRTeW5jKHBhc3N3b3JkLCBzYWx0LCA2NCkudG9TdHJpbmcoJ2hleCcpfWA7XG5jb25zdCBwYXNzd29yZE1hdGNoZXMgPSAocGFzc3dvcmQsIHN0b3JlZCkgPT4geyBjb25zdCBbc2FsdCwgZGlnZXN0XSA9IFN0cmluZyhzdG9yZWQgfHwgJycpLnNwbGl0KCc6Jyk7IGlmICghc2FsdCB8fCAhZGlnZXN0KSByZXR1cm4gZmFsc2U7IGNvbnN0IGFjdHVhbCA9IGNyeXB0by5zY3J5cHRTeW5jKHBhc3N3b3JkLCBzYWx0LCA2NCk7IHJldHVybiBjcnlwdG8udGltaW5nU2FmZUVxdWFsKGFjdHVhbCwgQnVmZmVyLmZyb20oZGlnZXN0LCAnaGV4JykpOyB9O1xuXG5leHBvcnQgZGVmYXVsdCBhc3luYyByZXF1ZXN0ID0+IHtcbiAgdHJ5IHtcbiAgICBjb25zdCB1cmwgPSBuZXcgVVJMKHJlcXVlc3QudXJsKTtcbiAgICBjb25zdCBtZXRob2QgPSByZXF1ZXN0Lm1ldGhvZCB8fCByZXF1ZXN0Lmh0dHBNZXRob2Q7XG4gICAgaWYgKHVybC5wYXRobmFtZS5lbmRzV2l0aCgnL2hlYWx0aCcpKSByZXR1cm4ganNvbigyMDAsIHsgb2s6IHRydWUgfSk7XG4gICAgaWYgKG1ldGhvZCA9PT0gJ0dFVCcgJiYgdXJsLnBhdGhuYW1lLmVuZHNXaXRoKCcvcHJvZmlsZXMnKSkge1xuICAgICAgY29uc3Qgd2IgPSBhd2FpdCByZWFkV29ya2Jvb2sodXNlcnNTdG9yZSgpLCBVU0VSU19LRVkpO1xuICAgICAgcmV0dXJuIGpzb24oMjAwLCB7IHByb2ZpbGVzOiBwcm9maWxlc0Zyb20od2IpIH0pO1xuICAgIH1cbiAgICBpZiAobWV0aG9kID09PSAnUE9TVCcgJiYgdXJsLnBhdGhuYW1lLmVuZHNXaXRoKCcvYXV0aC9yZWdpc3RlcicpKSB7XG4gICAgICBjb25zdCBib2R5ID0gYXdhaXQgcmVxdWVzdC5qc29uKCk7XG4gICAgICBjb25zdCB1c2VybmFtZSA9IFN0cmluZyhib2R5LnVzZXJuYW1lIHx8ICcnKS50cmltKCk7XG4gICAgICBjb25zdCBwYXNzd29yZCA9IFN0cmluZyhib2R5LnBhc3N3b3JkIHx8ICcnKTtcbiAgICAgIGlmICghL15bYS16QS1aMC05Xy4tXXszLDQwfSQvLnRlc3QodXNlcm5hbWUpIHx8IHBhc3N3b3JkLmxlbmd0aCA8IDgpIHJldHVybiBqc29uKDQwMCwgeyBlcnJvcjogJ1VzZSBhIHVzZXJuYW1lIHdpdGggMy00MCBsZXR0ZXJzL251bWJlcnMgYW5kIGEgcGFzc3dvcmQgb2YgYXQgbGVhc3QgOCBjaGFyYWN0ZXJzLicgfSk7XG4gICAgICBjb25zdCBzdG9yZSA9IHVzZXJzU3RvcmUoKTsgY29uc3Qgd2IgPSBhd2FpdCByZWFkV29ya2Jvb2soc3RvcmUsIFVTRVJTX0tFWSk7IGNvbnN0IHJvd3MgPSBhdXRoUm93cyh3Yik7XG4gICAgICBpZiAocm93cy5zb21lKHJvdyA9PiBTdHJpbmcocm93LlVzZXJuYW1lIHx8ICcnKS50b0xvd2VyQ2FzZSgpID09PSB1c2VybmFtZS50b0xvd2VyQ2FzZSgpKSkgcmV0dXJuIGpzb24oNDA5LCB7IGVycm9yOiAnVGhhdCB1c2VybmFtZSBpcyBhbHJlYWR5IHJlZ2lzdGVyZWQuJyB9KTtcbiAgICAgIHJvd3MucHVzaCh7IFVzZXJuYW1lOiB1c2VybmFtZSwgJ1Bhc3N3b3JkIEhhc2gnOiBwYXNzd29yZEhhc2gocGFzc3dvcmQpLCAnUHJvZmlsZSBOYW1lJzogdXNlcm5hbWUsICdDcmVhdGVkIERhdGUnOiBuZXcgRGF0ZSgpLnRvSVNPU3RyaW5nKCkgfSk7XG4gICAgICB3Yi5TaGVldHMuUHJvZmlsZXMgPSBYTFNYLnV0aWxzLmpzb25fdG9fc2hlZXQocm93cyk7IGlmICghd2IuU2hlZXROYW1lcy5pbmNsdWRlcygnUHJvZmlsZXMnKSkgWExTWC51dGlscy5ib29rX2FwcGVuZF9zaGVldCh3Yiwgd2IuU2hlZXRzLlByb2ZpbGVzLCAnUHJvZmlsZXMnKTsgYXdhaXQgd3JpdGVXb3JrYm9vayhzdG9yZSwgVVNFUlNfS0VZLCB3Yik7XG4gICAgICByZXR1cm4ganNvbigyMDEsIHsgcHJvZmlsZTogdXNlcm5hbWUsIHByb2ZpbGVzOiBwcm9maWxlc0Zyb20od2IpIH0pO1xuICAgIH1cbiAgICBpZiAobWV0aG9kID09PSAnUE9TVCcgJiYgdXJsLnBhdGhuYW1lLmVuZHNXaXRoKCcvYXV0aC9sb2dpbicpKSB7XG4gICAgICBjb25zdCBib2R5ID0gYXdhaXQgcmVxdWVzdC5qc29uKCk7IGNvbnN0IHVzZXJuYW1lID0gU3RyaW5nKGJvZHkudXNlcm5hbWUgfHwgJycpLnRyaW0oKTsgY29uc3QgcGFzc3dvcmQgPSBTdHJpbmcoYm9keS5wYXNzd29yZCB8fCAnJyk7XG4gICAgICBjb25zdCB3YiA9IGF3YWl0IHJlYWRXb3JrYm9vayh1c2Vyc1N0b3JlKCksIFVTRVJTX0tFWSk7IGNvbnN0IHVzZXIgPSBhdXRoUm93cyh3YikuZmluZChyb3cgPT4gU3RyaW5nKHJvdy5Vc2VybmFtZSB8fCAnJykudG9Mb3dlckNhc2UoKSA9PT0gdXNlcm5hbWUudG9Mb3dlckNhc2UoKSk7XG4gICAgICBpZiAoIXVzZXIgfHwgIXBhc3N3b3JkTWF0Y2hlcyhwYXNzd29yZCwgdXNlclsnUGFzc3dvcmQgSGFzaCddKSkgcmV0dXJuIGpzb24oNDAxLCB7IGVycm9yOiAnSW52YWxpZCB1c2VybmFtZSBvciBwYXNzd29yZC4nIH0pO1xuICAgICAgcmV0dXJuIGpzb24oMjAwLCB7IHByb2ZpbGU6IFN0cmluZyh1c2VyWydQcm9maWxlIE5hbWUnXSB8fCB1c2VyLlVzZXJuYW1lKSwgcHJvZmlsZXM6IHByb2ZpbGVzRnJvbSh3YikgfSk7XG4gICAgfVxuICAgIGlmIChtZXRob2QgPT09ICdQT1NUJyAmJiB1cmwucGF0aG5hbWUuZW5kc1dpdGgoJy9wcm9maWxlcycpKSB7XG4gICAgICBjb25zdCBwcm9maWxlID0gU3RyaW5nKChhd2FpdCByZXF1ZXN0Lmpzb24oKSkucHJvZmlsZSB8fCAnJykudHJpbSgpO1xuICAgICAgaWYgKCFwcm9maWxlKSByZXR1cm4ganNvbig0MDAsIHsgZXJyb3I6ICdBIHByb2ZpbGUgaXMgcmVxdWlyZWQuJyB9KTtcbiAgICAgIGNvbnN0IHN0b3JlID0gdXNlcnNTdG9yZSgpO1xuICAgICAgY29uc3Qgd2IgPSBhd2FpdCByZWFkV29ya2Jvb2soc3RvcmUsIFVTRVJTX0tFWSk7XG4gICAgICBjb25zdCBwcm9maWxlcyA9IHByb2ZpbGVzRnJvbSh3Yik7XG4gICAgICBpZiAoIXByb2ZpbGVzLmluY2x1ZGVzKHByb2ZpbGUpKSB7XG4gICAgICAgIGNvbnN0IHJvd3MgPSBbLi4ucHJvZmlsZXMsIHByb2ZpbGVdLm1hcChuYW1lID0+ICh7ICdQcm9maWxlIE5hbWUnOiBuYW1lLCAnQ3JlYXRlZCBEYXRlJzogbmV3IERhdGUoKS50b0lTT1N0cmluZygpIH0pKTtcbiAgICAgICAgd2IuU2hlZXRzLlByb2ZpbGVzID0gWExTWC51dGlscy5qc29uX3RvX3NoZWV0KHJvd3MpO1xuICAgICAgICBpZiAoIXdiLlNoZWV0TmFtZXMuaW5jbHVkZXMoJ1Byb2ZpbGVzJykpIFhMU1gudXRpbHMuYm9va19hcHBlbmRfc2hlZXQod2IsIHdiLlNoZWV0cy5Qcm9maWxlcywgJ1Byb2ZpbGVzJyk7XG4gICAgICAgIGF3YWl0IHdyaXRlV29ya2Jvb2soc3RvcmUsIFVTRVJTX0tFWSwgd2IpO1xuICAgICAgfVxuICAgICAgcmV0dXJuIGpzb24oMjAwLCB7IHByb2ZpbGUsIHByb2ZpbGVzOiBbLi4ubmV3IFNldChbLi4ucHJvZmlsZXMsIHByb2ZpbGVdKV0gfSk7XG4gICAgfVxuICAgIGNvbnN0IHBheWxvYWQgPSBtZXRob2QgPT09ICdHRVQnID8gT2JqZWN0LmZyb21FbnRyaWVzKHVybC5zZWFyY2hQYXJhbXMpIDogYXdhaXQgcmVxdWVzdC5qc29uKCk7XG4gICAgY29uc3QgcHJvZmlsZSA9IFN0cmluZyhwYXlsb2FkLnByb2ZpbGUgfHwgJycpLnRyaW0oKTtcbiAgICBpZiAobWV0aG9kID09PSAnUFVUJyAmJiB1cmwucGF0aG5hbWUuZW5kc1dpdGgoJy9zZXR0aW5ncycpKSB7XG4gICAgICBjb25zdCBzdG9yZSA9IHVzZXJzU3RvcmUoKTtcbiAgICAgIGNvbnN0IHdiID0gYXdhaXQgcmVhZFdvcmtib29rKHN0b3JlLCBVU0VSU19LRVkpO1xuICAgICAgY29uc3Qgcm93cyA9IFhMU1gudXRpbHMuc2hlZXRfdG9fanNvbih3Yi5TaGVldHMuU2V0dGluZ3MgfHwge30pLmZpbHRlcihyb3cgPT4gcm93LlByb2ZpbGUgIT09IHByb2ZpbGUpO1xuICAgICAgcm93cy5wdXNoKHsgUHJvZmlsZTogcHJvZmlsZSwgQ2F0ZWdvcmllczogcGF5bG9hZC5jYXRlZ29yaWVzIHx8ICcnLCAnRGVmYXVsdCBDdXJyZW5jeSc6IHBheWxvYWQuZGVmYXVsdEN1cnJlbmN5IHx8ICdMS1InLCAnVGhlbWUgUHJlZmVyZW5jZSc6IHBheWxvYWQudGhlbWUgfHwgJ2RhcmsnIH0pO1xuICAgICAgd2IuU2hlZXRzLlNldHRpbmdzID0gWExTWC51dGlscy5qc29uX3RvX3NoZWV0KHJvd3MpO1xuICAgICAgaWYgKCF3Yi5TaGVldE5hbWVzLmluY2x1ZGVzKCdTZXR0aW5ncycpKSBYTFNYLnV0aWxzLmJvb2tfYXBwZW5kX3NoZWV0KHdiLCB3Yi5TaGVldHMuU2V0dGluZ3MsICdTZXR0aW5ncycpO1xuICAgICAgYXdhaXQgd3JpdGVXb3JrYm9vayhzdG9yZSwgVVNFUlNfS0VZLCB3Yik7XG4gICAgICByZXR1cm4ganNvbigyMDAsIHsgc2V0dGluZ3M6IHBheWxvYWQgfSk7XG4gICAgfVxuICAgIGlmICghcHJvZmlsZSkgcmV0dXJuIGpzb24oNDAwLCB7IGVycm9yOiAnQSBwcm9maWxlIGlzIHJlcXVpcmVkLicgfSk7XG4gICAgaWYgKG1ldGhvZCA9PT0gJ0dFVCcpIHtcbiAgICAgIGNvbnN0IHdiID0gYXdhaXQgcmVhZFdvcmtib29rKGV4cGVuc2VzU3RvcmUoKSwgRVhQRU5TRVNfS0VZKTtcbiAgICAgIHJldHVybiBqc29uKDIwMCwgeyBleHBlbnNlczogcm93c0Zvcih3YiwgcHJvZmlsZSkgfSk7XG4gICAgfVxuICAgIGlmIChtZXRob2QgPT09ICdQT1NUJyB8fCBtZXRob2QgPT09ICdQVVQnIHx8IG1ldGhvZCA9PT0gJ0RFTEVURScpIHtcbiAgICAgIGNvbnN0IHN0b3JlID0gZXhwZW5zZXNTdG9yZSgpO1xuICAgICAgY29uc3Qgd2IgPSBhd2FpdCByZWFkV29ya2Jvb2soc3RvcmUsIEVYUEVOU0VTX0tFWSk7XG4gICAgICBjb25zdCBzaGVldCA9IHNhZmVTaGVldChwcm9maWxlKTtcbiAgICAgIGNvbnN0IHJvd3MgPSByb3dzRm9yKHdiLCBwcm9maWxlKTtcbiAgICAgIGxldCBuZXh0ID0gcm93cztcbiAgICAgIGlmIChtZXRob2QgPT09ICdQT1NUJykgbmV4dCA9IFsuLi5yb3dzLCBwYXlsb2FkLmV4cGVuc2VdO1xuICAgICAgaWYgKG1ldGhvZCA9PT0gJ1BVVCcpIG5leHQgPSByb3dzLm1hcChyb3cgPT4gcm93LklEID09PSBwYXlsb2FkLmV4cGVuc2UuSUQgPyBwYXlsb2FkLmV4cGVuc2UgOiByb3cpO1xuICAgICAgaWYgKG1ldGhvZCA9PT0gJ0RFTEVURScpIG5leHQgPSByb3dzLmZpbHRlcihyb3cgPT4gcm93LklEICE9PSBwYXlsb2FkLmlkKTtcbiAgICAgIGNvbnN0IGNsZWFuID0gbmV4dC5tYXAocm93ID0+IHsgY29uc3QgY29weSA9IHsgLi4ucm93IH07IGRlbGV0ZSBjb3B5LlByb2ZpbGU7IHJldHVybiBjb3B5OyB9KTtcbiAgICAgIHdiLlNoZWV0c1tzaGVldF0gPSBYTFNYLnV0aWxzLmpzb25fdG9fc2hlZXQoY2xlYW4pO1xuICAgICAgaWYgKCF3Yi5TaGVldE5hbWVzLmluY2x1ZGVzKHNoZWV0KSkgWExTWC51dGlscy5ib29rX2FwcGVuZF9zaGVldCh3Yiwgd2IuU2hlZXRzW3NoZWV0XSwgc2hlZXQpO1xuICAgICAgYXdhaXQgd3JpdGVXb3JrYm9vayhzdG9yZSwgRVhQRU5TRVNfS0VZLCB3Yik7XG4gICAgICByZXR1cm4ganNvbigyMDAsIHsgZXhwZW5zZXM6IG5leHQgfSk7XG4gICAgfVxuICAgIHJldHVybiBqc29uKDQwNSwgeyBlcnJvcjogJ01ldGhvZCBub3QgYWxsb3dlZC4nIH0pO1xuICB9IGNhdGNoIChlcnJvcikgeyByZXR1cm4ganNvbig1MDAsIHsgZXJyb3I6IGVycm9yLm1lc3NhZ2UgfHwgJ1N0b3JhZ2UgcmVxdWVzdCBmYWlsZWQuJyB9KTsgfVxufVxuIl0sCiAgIm1hcHBpbmdzIjogIjs7Ozs7Ozs7OztBQUFBLFNBQVMsZ0JBQWdCO0FBQ3pCLE9BQU8sVUFBVTtBQUNqQixPQUFPLFlBQVk7QUFFbkIsSUFBTSxZQUFZO0FBQ2xC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
