import { getStore } from '@netlify/blobs';
import XLSX from 'xlsx';
import crypto from 'node:crypto';

const USERS_KEY = 'ledgerly-users.xlsx';
const EXPENSES_KEY = 'ledgerly-expenses.xlsx';
const headers = { 'content-type': 'application/json', 'cache-control': 'no-store' };
const usersStore = () => getStore({ name: 'ledgerly-users', consistency: 'strong' });
const expensesStore = () => getStore({ name: 'ledgerly-expenses', consistency: 'strong' });
const json = (statusCode, body) => new Response(JSON.stringify(body), { status: statusCode, headers });
const readWorkbook = async (store, key) => { const data = await store.get(key, { type: 'arrayBuffer' }); return data ? XLSX.read(data, { type: 'array' }) : XLSX.utils.book_new(); };
const writeWorkbook = async (store, key, workbook) => store.set(key, XLSX.write(workbook, { bookType: 'xlsx', type: 'array' }));
const safeSheet = profile => profile.replace(/[\\/?*\[\]:]/g, ' ').trim().slice(0, 31) || 'User';
const rowsFor = (workbook, profile) => XLSX.utils.sheet_to_json(workbook.Sheets[safeSheet(profile)] || {}).map(row => ({ ...row, Profile: profile }));
const profilesFrom = workbook => XLSX.utils.sheet_to_json(workbook.Sheets.Profiles || {}).map(row => String(row['Profile Name'] || '')).filter(Boolean);
const authRows = workbook => XLSX.utils.sheet_to_json(workbook.Sheets.Profiles || {});
const passwordHash = (password, salt = crypto.randomBytes(16).toString('hex')) => `${salt}:${crypto.scryptSync(password, salt, 64).toString('hex')}`;
const passwordMatches = (password, stored) => { const [salt, digest] = String(stored || '').split(':'); if (!salt || !/^[a-f0-9]{128}$/i.test(digest)) return false; try { const actual = crypto.scryptSync(password, salt, 64); return crypto.timingSafeEqual(actual, Buffer.from(digest, 'hex')); } catch { return false; } };

export default async request => {
  try {
    const url = new URL(request.url);
    const method = request.method || request.httpMethod;
    if (url.pathname.endsWith('/health')) return json(200, { ok: true });
    if (method === 'GET' && url.pathname.endsWith('/profiles')) {
      const wb = await readWorkbook(usersStore(), USERS_KEY);
      return json(200, { profiles: profilesFrom(wb) });
    }
    if (method === 'POST' && url.pathname.endsWith('/auth/register')) {
      const body = await request.json();
      const username = String(body.username || '').trim();
      const password = String(body.password || '');
      if (!/^[a-zA-Z0-9_.-]{3,40}$/.test(username) || password.length < 8) return json(400, { error: 'Use a username with 3-40 letters/numbers and a password of at least 8 characters.' });
      const store = usersStore(); const wb = await readWorkbook(store, USERS_KEY); const rows = authRows(wb);
      if (rows.some(row => String(row.Username || '').toLowerCase() === username.toLowerCase())) return json(409, { error: 'That username is already registered.' });
      rows.push({ Username: username, 'Password Hash': passwordHash(password), 'Profile Name': username, 'Created Date': new Date().toISOString() });
      wb.Sheets.Profiles = XLSX.utils.json_to_sheet(rows); if (!wb.SheetNames.includes('Profiles')) XLSX.utils.book_append_sheet(wb, wb.Sheets.Profiles, 'Profiles'); await writeWorkbook(store, USERS_KEY, wb);
      return json(201, { profile: username, profiles: profilesFrom(wb) });
    }
    if (method === 'POST' && url.pathname.endsWith('/auth/login')) {
      const body = await request.json(); const username = String(body.username || '').trim(); const password = String(body.password || '');
      const wb = await readWorkbook(usersStore(), USERS_KEY); const user = authRows(wb).find(row => String(row.Username || '').toLowerCase() === username.toLowerCase());
      if (!user || !passwordMatches(password, user['Password Hash'])) return json(401, { error: 'Invalid username or password.' });
      return json(200, { profile: String(user['Profile Name'] || user.Username), profiles: profilesFrom(wb) });
    }
    if (method === 'POST' && url.pathname.endsWith('/profiles')) {
      const profile = String((await request.json()).profile || '').trim();
      if (!profile) return json(400, { error: 'A profile is required.' });
      const store = usersStore();
      const wb = await readWorkbook(store, USERS_KEY);
      const profiles = profilesFrom(wb);
      if (!profiles.includes(profile)) {
        const rows = [...profiles, profile].map(name => ({ 'Profile Name': name, 'Created Date': new Date().toISOString() }));
        wb.Sheets.Profiles = XLSX.utils.json_to_sheet(rows);
        if (!wb.SheetNames.includes('Profiles')) XLSX.utils.book_append_sheet(wb, wb.Sheets.Profiles, 'Profiles');
        await writeWorkbook(store, USERS_KEY, wb);
      }
      return json(200, { profile, profiles: [...new Set([...profiles, profile])] });
    }
    const payload = method === 'GET' ? Object.fromEntries(url.searchParams) : await request.json();
    const profile = String(payload.profile || '').trim();
    if (method === 'PUT' && url.pathname.endsWith('/settings')) {
      const store = usersStore();
      const wb = await readWorkbook(store, USERS_KEY);
      const rows = XLSX.utils.sheet_to_json(wb.Sheets.Settings || {}).filter(row => row.Profile !== profile);
      rows.push({ Profile: profile, Categories: payload.categories || '', 'Default Currency': payload.defaultCurrency || 'LKR', 'Theme Preference': payload.theme || 'dark' });
      wb.Sheets.Settings = XLSX.utils.json_to_sheet(rows);
      if (!wb.SheetNames.includes('Settings')) XLSX.utils.book_append_sheet(wb, wb.Sheets.Settings, 'Settings');
      await writeWorkbook(store, USERS_KEY, wb);
      return json(200, { settings: payload });
    }
    if (!profile) return json(400, { error: 'A profile is required.' });
    if (method === 'GET') {
      const wb = await readWorkbook(expensesStore(), EXPENSES_KEY);
      return json(200, { expenses: rowsFor(wb, profile) });
    }
    if (method === 'POST' || method === 'PUT' || method === 'DELETE') {
      const store = expensesStore();
      const wb = await readWorkbook(store, EXPENSES_KEY);
      const sheet = safeSheet(profile);
      const rows = rowsFor(wb, profile);
      let next = rows;
      if (method === 'POST') next = [...rows, payload.expense];
      if (method === 'PUT') next = rows.map(row => row.ID === payload.expense.ID ? payload.expense : row);
      if (method === 'DELETE') next = rows.filter(row => row.ID !== payload.id);
      const clean = next.map(row => { const copy = { ...row }; delete copy.Profile; return copy; });
      wb.Sheets[sheet] = XLSX.utils.json_to_sheet(clean);
      if (!wb.SheetNames.includes(sheet)) XLSX.utils.book_append_sheet(wb, wb.Sheets[sheet], sheet);
      await writeWorkbook(store, EXPENSES_KEY, wb);
      return json(200, { expenses: next });
    }
    return json(405, { error: 'Method not allowed.' });
  } catch (error) { return json(500, { error: error.message || 'Storage request failed.' }); }
}
