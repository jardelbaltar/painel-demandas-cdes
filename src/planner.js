const tenantId = import.meta.env.VITE_MICROSOFT_TENANT_ID || 'de23d5f0-ccac-4c84-81d6-2892a8c055aa';
const clientId = import.meta.env.VITE_MICROSOFT_CLIENT_ID;
const planId = import.meta.env.VITE_PLANNER_PLAN_ID || '_IFjpmPlW02Q7eVsII-VQmQADmgL';
const scopes = ['openid', 'profile', 'offline_access', 'User.Read', 'Tasks.Read', 'Group.Read.All'];
const cacheKey = 'cdes-planner-auth';

export const isPlannerConfigured = () => Boolean(clientId);

const encode = bytes => btoa(String.fromCharCode(...new Uint8Array(bytes))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
const redirectUri = () => window.location.origin + window.location.pathname;
const readCache = () => JSON.parse(localStorage.getItem(cacheKey) || 'null');
const writeCache = value => localStorage.setItem(cacheKey, JSON.stringify(value));

const exchangeToken = async params => {
  const response = await fetch(tokenEndpoint, {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: clientId, redirect_uri: redirectUri(), ...params }),
  });
  const result = await response.json();
  if (!response.ok) throw new Error(result.error_description || 'Não foi possível concluir a autenticação Microsoft');
  const cached = { ...result, expires_at: Date.now() + result.expires_in * 1000 };
  writeCache(cached);
  return cached;
};

const completeRedirect = async () => {
  const params = new URLSearchParams(window.location.search);
  const code = params.get('code');
  if (!code) return null;
  const verifier = sessionStorage.getItem(`${cacheKey}-verifier`);
  const expectedState = sessionStorage.getItem(`${cacheKey}-state`);
  if (!verifier || params.get('state') !== expectedState) throw new Error('Resposta de autenticação inválida');
  const token = await exchangeToken({ grant_type: 'authorization_code', code, code_verifier: verifier, scope: scopes.join(' ') });
  sessionStorage.removeItem(`${cacheKey}-verifier`);
  sessionStorage.removeItem(`${cacheKey}-state`);
  history.replaceState({}, document.title, redirectUri());
  return token;
};

const beginLogin = async () => {
  const verifier = encode(crypto.getRandomValues(new Uint8Array(64)));
  const challenge = encode(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
  const state = encode(crypto.getRandomValues(new Uint8Array(24)));
  sessionStorage.setItem(`${cacheKey}-verifier`, verifier);
  sessionStorage.setItem(`${cacheKey}-state`, state);
  const authorize = new URL(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`);
  authorize.search = new URLSearchParams({ client_id: clientId, response_type: 'code', redirect_uri: redirectUri(), response_mode: 'query', scope: scopes.join(' '), state, code_challenge: challenge, code_challenge_method: 'S256', prompt: 'select_account' });
  window.location.assign(authorize);
  return new Promise(() => {});
};

const getAccessToken = async interactive => {
  if (!clientId) throw new Error('Integração com o Planner não configurada');
  const redirected = await completeRedirect();
  if (redirected) return redirected.access_token;
  const cached = readCache();
  if (cached?.access_token && cached.expires_at > Date.now() + 60000) return cached.access_token;
  if (cached?.refresh_token) {
    try { return (await exchangeToken({ grant_type: 'refresh_token', refresh_token: cached.refresh_token, scope: scopes.join(' ') })).access_token; }
    catch { localStorage.removeItem(cacheKey); }
  }
  if (interactive) return beginLogin();
  throw new Error('login_required');
};

const getAccount = token => {
  try { return JSON.parse(new TextDecoder().decode(Uint8Array.from(atob(token.split('.')[1].replace(/-/g, '+').replace(/_/g, '/')), character => character.charCodeAt(0)))); }
  catch { return {}; }
};

const fetchGraphPage = async (url, token) => {
  const response = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!response.ok) {
    const body = await response.json().catch(() => ({}));
    const error = new Error(body.error?.message || `Microsoft Graph retornou HTTP ${response.status}`);
    error.status = response.status;
    throw error;
  }
  return response.json();
};

const graphCollection = async (path, token) => {
  const items = [];
  let next = `https://graph.microsoft.com/v1.0${path}`;
  while (next) {
    const page = await fetchGraphPage(next, token);
    items.push(...(page.value || []));
    next = page['@odata.nextLink'] || '';
  }
  return items;
};

export const parseBucketName = value => {
  const originalName = String(value || '').trim();
  const match = originalName.match(/^(.*?)\s*-\s*(\d+)\s*$/);
  return match ? { name: match[1].trim(), developers: Number(match[2]), originalName } : { name: originalName, developers: 0, originalName };
};

const priorityLabel = priority => priority === 0 ? 'Urgente' : priority <= 3 ? 'Importante' : priority <= 7 ? 'Média' : 'Baixa';
const isoDate = value => value ? value.slice(0, 10) : '';

export async function loadPlanner({ interactive = false } = {}) {
  const token = await getAccessToken(interactive);
  const [buckets, plannerTasks] = await Promise.all([
    graphCollection(`/planner/plans/${encodeURIComponent(planId)}/buckets`, token),
    graphCollection(`/planner/plans/${encodeURIComponent(planId)}/tasks`, token),
  ]);
  const taskDetails = await Promise.allSettled(plannerTasks.map(task => fetchGraphPage(`/planner/tasks/${encodeURIComponent(task.id)}/details`, token)));
  const bucketById = new Map(buckets.map(bucket => [bucket.id, parseBucketName(bucket.name)]));
  const teams = buckets.map(bucket => ({ id: bucket.id, ...parseBucketName(bucket.name) }));
  const tasks = plannerTasks.map((task, index) => {
    const bucket = bucketById.get(task.bucketId) || parseBucketName('Sem bucket');
    const progress = Number(task.percentComplete) || 0;
    const priority = priorityLabel(Number(task.priority));
    const plannerDetail = taskDetails[index].status === 'fulfilled' ? taskDetails[index].value : {};
    const checklist = Object.values(plannerDetail.checklist || {}).map(item => ({ title: item.title, completed: Boolean(item.isChecked) }));
    return {
      id: task.id, title: task.title, team: bucket.name,
      start: isoDate(task.startDateTime), end: isoDate(task.dueDateTime),
      status: progress >= 100 ? 'Concluída' : progress > 0 ? 'Em execução' : 'Planejada',
      progress, priority, highlightedLabel: '', checklist,
      details: [{ label: 'Bucket no Planner', value: bucket.originalName }, { label: 'Progresso', value: `${progress}%` }, { label: 'Prioridade', value: priority }],
    };
  }).filter(task => task.title);
  return { teams, tasks, account: getAccount(token), syncedAt: new Date() };
}
