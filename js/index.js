/* ============================================================
   TASKFLOW — index.js  (optimised build)
   ============================================================
   QUICK-START FOR NEW FEATURES
   ─────────────────────────────────────────────────────────────
   1. DOM access         →  $(id)  or  $$(selector)
   2. Cached elements    →  EL.xxx  (populated in initializeApp)
                            Add new ones to the EL block there.
   3. Debouncing         →  debounce(fn, ms)
   4. Magic numbers      →  CONFIG.XXX  (defined below)
   5. Calendar refresh   →  renderCalendarDebounced()
                            (use this instead of renderCalendar()
                             after data mutations)
   6. Sections (Ctrl+F)  →  "SUPABASE CLIENT"
                             "API SHIM"
                             "SUPER-ADMIN"
                             "STATE"
                             "NAVIGATION / VIEWS"
                             "TASK BOARD"
                             "CALENDAR"
                             "ADD / EDIT TASK MODAL"
                             "NOTIFICATIONS"
                             "IMPORT USERS"
   ─────────────────────────────────────────────────────────────
*/

/* ============================================================
   PERFORMANCE UTILITIES
   Fast DOM helpers, debounce, and shared config.
   ============================================================ */

/** Fast getElementById alias */
function $(id) {
  return document.getElementById(id);
}

/** Fast querySelectorAll alias — returns Array */
function $$(sel, root) {
  return Array.from((root || document).querySelectorAll(sel));
}

/**
 * Debounce: collapses rapid repeated calls into one.
 * Usage: const fn = debounce(() => doWork(), 80);
 */
function debounce(fn, ms) {
  var t;
  return function () {
    var args = arguments;
    var ctx = this;
    clearTimeout(t);
    t = setTimeout(function () {
      fn.apply(ctx, args);
    }, ms);
  };
}

/** App-wide configuration. Centralises all magic numbers. */
const CONFIG = {
  POLL_INTERVAL_MS: 30000, // how often we poll for notifications
  DEBOUNCE_RENDER_MS: 60, // calendar / board re-render debounce
  DEBOUNCE_SEARCH_MS: 120, // search input debounce
  SCROLL_TO_TASK_MS: 100, // delay before scrolling to a focused task
  TOAST_DURATION_MS: 4000, // how long toasts stay visible
  CAL_BLOCK_H: 22, // px — height of one calendar task block
};

/* ============================================================
   SUPABASE CLIENT — multi-device, real-time database
   URL: https://xtlaqgititvfjorxdbgi.supabase.co
   ============================================================ */
const SUPA_URL = 'https://xtlaqgititvfjorxdbgi.supabase.co';
const SUPA_KEY =
  'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inh0bGFxZ2l0aXR2ZmpvcnhkYmdpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzE3NjkzMDQsImV4cCI6MjA4NzM0NTMwNH0.Lrqo4504z-4r9SZWn1FagPHHk8ogFqpjqnnsqxJascg';

// ---------- Core Supabase REST helpers ----------
async function supa(method, table, body, filters, options) {
  filters = filters || '';
  options = options || {};
  const url = SUPA_URL + '/rest/v1/' + table + (filters ? '?' + filters : '');
  const headers = {
    apikey: SUPA_KEY,
    Authorization: 'Bearer ' + SUPA_KEY,
    'Content-Type': 'application/json',
    Accept: 'application/json',
  };
  // Always ask Supabase to return the full row(s) after mutating
  if (method === 'POST' || method === 'PATCH' || method === 'PUT') {
    headers['Prefer'] = 'return=representation';
  }
  const opts = { method: method, headers: headers };
  if (body && method !== 'GET' && method !== 'DELETE') {
    opts.body = JSON.stringify(body);
  }
  // Offline: queue mutations and return optimistic null
  if (!navigator.onLine && method !== 'GET') {
    if (typeof OfflineQueue !== 'undefined') {
      const path = '/direct/' + table + (filters ? '?' + filters : '');
      OfflineQueue.enqueue({ method, path, body: body || null }).catch(
        () => { },
      );
    }
    return null; // Callers handle null gracefully
  }
  const res = await fetch(url, opts);
  // 204 = No Content (DELETE / PATCH with no return)
  if (res.status === 204) return null;
  const text = await res.text();
  if (!text) return null;
  let data;
  try {
    data = JSON.parse(text);
  } catch (e) {
    throw new Error('Bad JSON from Supabase: ' + text.slice(0, 120));
  }
  if (!res.ok) {
    const msg = Array.isArray(data) ? data[0] : data;
    throw new Error(
      msg?.message ||
      msg?.error ||
      msg?.hint ||
      'Supabase error ' + res.status + ': ' + text.slice(0, 200),
    );
  }
  return data;
}

const SB = {
  async select(table, filters) {
    return supa('GET', table, null, filters);
  },
  async insert(table, body) {
    const r = await supa('POST', table, body);
    // Supabase returns an array for inserts — grab first row
    return Array.isArray(r) ? r[0] : r;
  },
  async update(table, id, body) {
    const r = await supa('PATCH', table, body, 'id=eq.' + id);
    return Array.isArray(r) ? r[0] : r;
  },
  async updateWhere(table, filters, body) {
    const r = await supa('PATCH', table, body, filters);
    return r;
  },
  async delete(table, id) {
    return supa('DELETE', table, null, 'id=eq.' + id);
  },
  async deleteWhere(table, filters) {
    return supa('DELETE', table, null, filters);
  },
};

// ---------- API shim — multi-tenant, company-scoped ----------
// Every query is automatically filtered to the current user's company_id.
// Super-admin (role='superadmin') bypasses company scoping.

function _cid() {
  return localStorage.getItem('tf_cid') || '';
}

const API = {
  getToken() {
    return localStorage.getItem('tf_token');
  },
  setToken(t) {
    localStorage.setItem('tf_token', t);
  },
  clearToken() {
    ['tf_token', 'tf_uid', 'tf_cid', 'tf_role'].forEach((k) =>
      localStorage.removeItem(k),
    );
  },

  async request(method, path, body) {
    const parts = path.split('?')[0].split('/').filter(Boolean);
    const [resource, id, action] = parts;
    const cid = _cid();
    const role = localStorage.getItem('tf_role') || '';
    const isSA = role === 'superadmin';

    // ── AUTH ────────────────────────────────────────────────────
    if (resource === 'auth') {
      if (id === 'login') {
        const rows = await SB.select(
          'tf_users',
          `username=eq.${body.username}&select=*&limit=1`,
        );
        if (
          !rows ||
          rows.length === 0 ||
          rows[0].password_hash !== body.password
        )
          throw new Error('Invalid credentials');
        const u = rows[0];
        localStorage.setItem('tf_uid', u.id);
        localStorage.setItem('tf_cid', u.company_id || '');
        localStorage.setItem('tf_role', u.role);
        return {
          token: 'app-session-' + u.id,
          user: {
            id: u.id,
            name: u.name,
            username: u.username,
            role: u.role,
            companyId: u.company_id,
          },
        };
      }
      if (id === 'logout') {
        this.clearToken();
        return null;
      }
    }

    // ── COMPANIES (super-admin only) ─────────────────────────────
    if (resource === 'companies') {
      if (!isSA) throw new Error('Not authorised');
      if (method === 'GET')
        return SB.select('tf_companies', 'select=*&order=created_at.asc');
      if (method === 'POST')
        return SB.insert('tf_companies', { name: body.name });
      if (method === 'DELETE') {
        await SB.delete('tf_companies', id);
        return null;
      }
    }

    // ── USERS ────────────────────────────────────────────────────
    if (resource === 'users') {
      const cidFilter = isSA ? '' : `company_id=eq.${cid}&`;
      if (method === 'GET' && !id) {
        const rows = await SB.select(
          'tf_users',
          cidFilter +
          'select=id,name,username,role,company_id,email,email_notif,created_at',
        );
        return (rows || []).map((u) => ({
          id: u.id,
          name: u.name,
          username: u.username,
          role: u.role,
          companyId: u.company_id,
          email: u.email,
          emailNotif: u.email_notif,
          createdAt: u.created_at,
        }));
      }
      if (method === 'POST') {
        const exists = await SB.select(
          'tf_users',
          `username=eq.${body.username}&select=id&limit=1`,
        );
        if (exists && exists.length > 0)
          throw new Error('Username already taken');
        const ins = {
          name: body.name,
          username: body.username,
          password_hash: body.password,
          role: body.role || 'user',
          email: body.email || null,
          email_notif: body.emailNotif !== false,
        };
        if (!isSA) ins.company_id = cid;
        else if (body.companyId) ins.company_id = body.companyId;
        const nu = await SB.insert('tf_users', ins);
        return {
          id: nu.id,
          name: nu.name,
          username: nu.username,
          role: nu.role,
          companyId: nu.company_id,
          email: nu.email,
          emailNotif: nu.email_notif,
          createdAt: nu.created_at,
        };
      }
      if (method === 'PUT' && id) {
        const upd = {};
        if (body.password !== undefined) upd.password_hash = body.password;
        if (body.name !== undefined) upd.name = body.name;
        if (body.email !== undefined) upd.email = body.email;
        if (body.emailNotif !== undefined) upd.email_notif = body.emailNotif;
        let u = await SB.update('tf_users', id, upd);
        if (!u) {
          const rows = await SB.select('tf_users', `id=eq.${id}&limit=1`);
          if (rows && rows.length > 0) u = rows[0];
        }
        return u
          ? {
            id: u.id,
            name: u.name,
            username: u.username,
            role: u.role,
            email: u.email,
            emailNotif: u.email_notif,
          }
          : null;
      }
      if (method === 'DELETE' && id) {
        await SB.deleteWhere('tf_tasks', `user_id=eq.${id}`);
        await SB.delete('tf_users', id);
        return null;
      }
    }

    // ── TASKS ────────────────────────────────────────────────────
    if (resource === 'tasks') {
      const cidFilter = isSA ? '' : `company_id=eq.${cid}&`;
      if (method === 'GET') {
        const rows = await SB.select(
          'tf_tasks',
          cidFilter + 'select=*&order=created_at.desc',
        );
        return (rows || []).map(_mapTask);
      }
      if (method === 'POST') {
        const ins = _taskToDb(body);
        if (!isSA) ins.company_id = cid;
        return _mapTask(await SB.insert('tf_tasks', ins));
      }
      if (method === 'PUT' && id) {
        const r = await SB.update('tf_tasks', id, _taskToDb(body));
        return r ? _mapTask(r) : null;
      }
      if (method === 'DELETE' && id) {
        await SB.delete('tf_tasks', id);
        return null;
      }
    }

    // ── LOGS ─────────────────────────────────────────────────────
    if (resource === 'logs') {
      const cidFilter = isSA ? '' : `company_id=eq.${cid}&`;
      if (method === 'GET') {
        const rows = await SB.select(
          'tf_logs',
          cidFilter + 'select=*&order=created_at.desc&limit=500',
        );
        return (rows || []).map((r) => ({
          id: r.id,
          taskId: r.task_id,
          taskTitle: r.task_title,
          action: r.action,
          actorName: r.actor_name,
          userId: r.user_id,
          timestamp: r.created_at,
        }));
      }
      if (method === 'POST') {
        const ins = {
          task_id: body.taskId,
          task_title: body.taskTitle,
          action: body.action,
          actor_name: body.actorName,
          user_id: body.userId,
        };
        if (!isSA) ins.company_id = cid;
        const r = await SB.insert('tf_logs', ins);
        return {
          id: r.id,
          taskId: r.task_id,
          taskTitle: r.task_title,
          action: r.action,
          actorName: r.actor_name,
          userId: r.user_id,
          timestamp: r.created_at,
        };
      }
    }

    // ── NOTIFICATIONS ─────────────────────────────────────────────
    if (resource === 'notifications') {
      const uid = localStorage.getItem('tf_uid');
      if (method === 'GET') {
        const rows = await SB.select(
          'tf_notifications',
          `user_id=eq.${uid}&select=*&order=created_at.desc&limit=200`,
        );
        return (rows || []).map(_mapNotif);
      }
      if (method === 'POST') {
        const ins = {
          user_id: body.userId,
          title: body.title,
          body: body.body,
          task_id: body.taskId || null,
          metadata: body.metadata || null,
          is_read: false,
        };
        if (!isSA) ins.company_id = cid;
        return _mapNotif(await SB.insert('tf_notifications', ins));
      }
      if (method === 'PUT' && id && action === 'read') {
        await SB.update('tf_notifications', id, { is_read: true });
        return null;
      }
      if (method === 'PUT' && id === 'read-all') {
        await SB.updateWhere('tf_notifications', `user_id=eq.${uid}`, {
          is_read: true,
        });
        return null;
      }
    }

    // ── LEAVES ───────────────────────────────────────────────────
    if (resource === 'leaves') {
      const cidFilter = isSA ? '' : `company_id=eq.${cid}&`;
      if (method === 'GET') {
        const rows = await SB.select(
          'tf_leaves',
          cidFilter + 'select=*&order=start_date.asc',
        );
        return (rows || []).map(_mapLeave);
      }
      if (method === 'POST') {
        const ins = {
          user_id: body.userId,
          type: body.type,
          start_date: body.startDate,
          end_date: body.endDate,
          reason: body.reason || null,
          status: body.status || 'approved',
          requested_by: body.requestedBy || null,
        };
        if (!isSA) ins.company_id = cid;
        return _mapLeave(await SB.insert('tf_leaves', ins));
      }
      if (method === 'PATCH' && id) {
        const upd = {};
        if (body.status !== undefined) upd.status = body.status;
        if (body.requestedBy !== undefined) upd.requested_by = body.requestedBy;
        return _mapLeave(await SB.update('tf_leaves', id, upd));
      }
      if (method === 'PUT' && id) {
        const upd = {};
        if (body.userId !== undefined) upd.user_id = body.userId;
        if (body.type !== undefined) upd.type = body.type;
        if (body.startDate !== undefined) upd.start_date = body.startDate;
        if (body.endDate !== undefined) upd.end_date = body.endDate;
        if (body.reason !== undefined) upd.reason = body.reason;
        return _mapLeave(await SB.update('tf_leaves', id, upd));
      }
      if (method === 'DELETE' && id) {
        await SB.delete('tf_leaves', id);
        return null;
      }
    }

    // ── SCHEDULES ────────────────────────────────────────────────
    if (resource === 'schedules') {
      const cidFilter = isSA ? '' : `company_id=eq.${cid}&`;
      if (method === 'GET') {
        const rows = await SB.select(
          'tf_schedules',
          cidFilter + 'select=*&order=created_at.asc',
        );
        if (!rows || rows.length === 0)
          return [
            {
              id: 'ws-default',
              name: 'Standard Workday',
              start: 9,
              end: 18,
              active: true,
            },
          ];
        return rows.map(_mapSchedule);
      }
      if (method === 'POST') {
        const ins = {
          name: body.name,
          start_hour: body.startHour,
          end_hour: body.endHour,
          is_active: false,
        };
        if (!isSA) ins.company_id = cid;
        return _mapSchedule(await SB.insert('tf_schedules', ins));
      }
      if (method === 'PUT' && id && action === 'activate') {
        const f = isSA
          ? 'id=neq.00000000-0000-0000-0000-000000000000'
          : `company_id=eq.${cid}`;
        await SB.updateWhere('tf_schedules', f, { is_active: false });
        await SB.update('tf_schedules', id, { is_active: true });
        return null;
      }
      if (method === 'DELETE' && id) {
        await SB.delete('tf_schedules', id);
        return null;
      }
    }

    // ── TEAMS ────────────────────────────────────────────────────
    if (resource === 'teams') {
      const cidFilter = isSA ? '' : `company_id=eq.${cid}&`;
      if (method === 'GET') {
        const rows = await SB.select(
          'tf_teams',
          cidFilter + 'select=*&order=created_at.asc',
        );
        return (rows || []).map(_mapTeam);
      }
      if (method === 'POST') {
        const ins = {
          name: body.name,
          color: body.color || '#F59E0B',
          member_ids: body.memberIds || [],
        };
        if (!isSA) ins.company_id = cid;
        return _mapTeam(await SB.insert('tf_teams', ins));
      }
      if (method === 'PUT' && id) {
        const upd = {};
        if (body.name !== undefined) upd.name = body.name;
        if (body.color !== undefined) upd.color = body.color;
        if (body.memberIds !== undefined) upd.member_ids = body.memberIds;
        return _mapTeam(await SB.update('tf_teams', id, upd));
      }
      if (method === 'DELETE' && id) {
        await SB.delete('tf_teams', id);
        return null;
      }
    }

    return null;
  },

  get(path) {
    return this.request('GET', path);
  },
  post(path, body) {
    return this.request('POST', path, body);
  },
  put(path, body) {
    return this.request('PUT', path, body);
  },
  patch(path, body) {
    return this.request('PATCH', path, body);
  },
  del(path) {
    return this.request('DELETE', path);
  },
};

// ---------- Field mappers ------------------------------------------------
function _mapTask(r) {
  if (!r) return null;
  return {
    id: r.id,
    userId: r.user_id,
    title: r.title,
    requestor: r.requestor,
    priority: r.priority,
    start: r.start_at,
    deadline: r.deadline_at,
    description: r.description || [],
    done: r.is_done || false,
    doneAt: r.done_at,
    cancelled: r.is_cancelled || false,
    cancelledAt: r.cancelled_at,
    cancelReason: r.cancel_reason,
    createdAt: r.created_at,
    isMultiPersonnel: r.is_multi || false,
    multiGroupId: r.group_id,
    isTeamTask: r.is_team_task || false,
    teamId: r.team_id,
    teamName: r.team_name,
    createdBy: r.created_by || null,
  };
}
function _taskToDb(b) {
  const d = {};
  if (b.userId !== undefined) d.user_id = b.userId;
  if (b.title !== undefined) d.title = b.title;
  if (b.requestor !== undefined) d.requestor = b.requestor;
  if (b.priority !== undefined) d.priority = b.priority;
  if (b.start !== undefined) d.start_at = b.start;
  if (b.deadline !== undefined) d.deadline_at = b.deadline;
  if (b.description !== undefined) d.description = b.description;
  if (b.done !== undefined) d.is_done = b.done;
  if (b.doneAt !== undefined) d.done_at = b.doneAt;
  if (b.cancelled !== undefined) d.is_cancelled = b.cancelled;
  if (b.cancelledAt !== undefined) d.cancelled_at = b.cancelledAt;
  if (b.cancelReason !== undefined) d.cancel_reason = b.cancelReason;
  if (b.isTeamTask !== undefined) d.is_team_task = b.isTeamTask;
  if (b.teamId !== undefined) d.team_id = b.teamId;
  if (b.teamName !== undefined) d.team_name = b.teamName;
  if (b.multiGroupId !== undefined) d.group_id = b.multiGroupId;
  if (b.createdBy !== undefined) d.created_by = b.createdBy;
  return d;
}

// ── Shared task helpers ─────────────────────────────────────────────────────
// Group metadata is embedded as a hidden sentinel item in the description array.
// It is filtered out everywhere in the UI so users never see it.
// Format: { __meta: true, groupId: 'tg-...' | 'mg-...', memberIds: [...] }

function _makeGroupMeta(groupId, memberIds) {
  return { __meta: true, groupId, memberIds: memberIds.slice() };
}
function _getGroupMeta(description) {
  if (!Array.isArray(description)) return null;
  return description.find((d) => d.__meta) || null;
}
// Returns the visible (non-meta) checklist items
function _visibleDesc(description) {
  if (!Array.isArray(description)) return [];
  return description.filter((d) => !d.__meta);
}
// Build a full description array from visible items + existing meta
function _buildDesc(visibleItems, meta) {
  if (!meta) return visibleItems;
  return [...visibleItems, meta];
}
// Propagate a checklist change to every task in the same group
async function _syncGroupCheck(changedTask, newDescription) {
  const meta = _getGroupMeta(changedTask.description);
  if (!meta || !meta.groupId) return; // not a group task — nothing to sync
  const groupId = meta.groupId;
  const allTasks = getTasks();
  const groupTasks = allTasks.filter((t) => {
    const m = _getGroupMeta(t.description);
    return m && m.groupId === groupId && t.id !== changedTask.id;
  });
  // Build the synced description: visible items from newDescription + each task's own meta
  const newVisible = _visibleDesc(newDescription);
  for (const t of groupTasks) {
    const tMeta = _getGroupMeta(t.description);
    const synced = _buildDesc(newVisible, tMeta);
    t.description = synced;
    // Persist silently
    API.put(`/tasks/${t.id}`, { description: synced }).catch(() => { });
  }
}

// Returns all sibling tasks that share the same team group (same group_id in __meta).
// Used to propagate done/cancel/reopen/delete/edit to every member's copy of a team task.
function _getTeamSiblings(task) {
  if (!task || !task.isTeamTask) return [];
  const meta = _getGroupMeta(task.description);
  if (!meta || !meta.groupId) return [];
  const groupId = meta.groupId;
  return getTasks().filter(
    (t) =>
      t.id !== task.id &&
      t.isTeamTask &&
      (() => {
        const m = _getGroupMeta(t.description);
        return m && m.groupId === groupId;
      })(),
  );
}

function _mapNotif(r) {
  if (!r) return null;
  return {
    id: r.id,
    userId: r.user_id,
    title: r.title,
    body: r.body,
    taskId: r.task_id,
    metadata: r.metadata,
    read: r.is_read || false,
    timestamp: r.created_at,
  };
}
function _mapLeave(r) {
  if (!r) return null;
  return {
    id: r.id,
    userId: r.user_id,
    type: r.type,
    startDate: r.start_date,
    endDate: r.end_date,
    reason: r.reason,
    status: r.status || 'approved',
    requestedBy: r.requested_by || null,
    createdAt: r.created_at,
  };
}
function _mapSchedule(r) {
  if (!r) return null;
  return {
    id: r.id,
    name: r.name,
    start: r.start_hour,
    end: r.end_hour,
    active: r.is_active || false,
  };
}
function _mapTeam(r) {
  if (!r) return null;
  return {
    id: r.id,
    name: r.name,
    color: r.color || '#F59E0B',
    memberIds: r.member_ids || [],
    createdAt: r.created_at,
  };
}

// ---------- In-memory cache ---------------------------------------------
const cache = {
  users: null,
  tasks: null,
  logs: null,
  notifications: null,
  leaves: null,
  workSchedules: null,
  teams: null,
  companies: null,
};

// Poll for new notifications every 30s when logged in
let _pollInterval = null;
async function startPolling() {
  clearInterval(_pollInterval);
  _pollInterval = setInterval(async () => {
    try {
      cache.notifications = (await API.get('/notifications')) || [];
      updateNotifBadge();
    } catch { }
  }, 30000);
}
function stopPolling() {
  clearInterval(_pollInterval);
}

// Legacy synchronous-style getters (now use cache)
function getUsers() {
  return cache.users || [];
}
function getTasks() {
  return cache.tasks || [];
}
function getLogs() {
  return cache.logs || [];
}
function getNotifs() {
  return cache.notifications || [];
}
function getLeaves() {
  return cache.leaves || [];
}
function getWorkSchedules() {
  return (
    cache.workSchedules || [
      {
        id: 'ws-default',
        name: 'Standard Workday',
        start: 9,
        end: 18,
        active: true,
      },
    ]
  );
}
function getTeams() {
  return cache.teams || [];
}

// Async save helpers
async function saveTasks(t) {
  // Handled inline in each action via API calls; cache stays in sync
  cache.tasks = t;
}
async function saveUsers(u) {
  cache.users = u;
}
async function saveLogs(l) {
  cache.logs = l;
}
async function saveNotifs(n) {
  cache.notifications = n;
}
async function saveLeaves(l) {
  cache.leaves = l;
}
async function saveWorkSchedules(arr) {
  cache.workSchedules = arr;
}

function getWorkHours() {
  const scheds = getWorkSchedules();
  return (
    scheds.find((s) => s.active) ||
    scheds[0] || { name: 'Standard Workday', start: 9, end: 18 }
  );
}

async function loadAll() {
  if (localStorage.getItem('tf_role') === 'superadmin') return;
  try {
    const [users, tasks, logs, notifs, leaves, schedules, teams] =
      await Promise.all([
        API.get('/users'),
        API.get('/tasks'),
        API.get('/logs'),
        API.get('/notifications'),
        API.get('/leaves'),
        API.get('/schedules'),
        API.get('/teams'),
      ]);
    cache.users = users || [];
    cache.tasks = tasks || [];
    cache.logs = logs || [];
    cache.notifications = notifs || [];
    cache.leaves = leaves || [];
    cache.workSchedules = schedules || [];
    cache.teams = teams || [];
  } catch (e) {
    console.error('loadAll failed:', e);
  }
  _initialLoadDone = true;
}

/* ============================================================
   SUPER-ADMIN FUNCTIONS
   ============================================================ */
async function loadSACompanies() {
  const list = document.getElementById('sa-companies-list');
  list.innerHTML =
    '<div style="text-align:center;padding:40px;color:var(--text3);">Loading...</div>';
  try {
    const companies = await API.get('/companies');
    cache.companies = companies || [];
    if (!companies || companies.length === 0) {
      list.innerHTML = `<div class="sa-empty"><div class="sa-empty-icon">🏢</div><div>No companies yet. Add your first one above.</div></div>`;
      return;
    }
    // Load user counts per company
    const allUsers = await SB.select('tf_users', 'select=id,company_id,role');
    list.innerHTML = companies
      .map((co) => {
        const coUsers = (allUsers || []).filter(
          (u) => u.company_id === co.id && u.role !== 'admin',
        );
        const admins = (allUsers || []).filter(
          (u) => u.company_id === co.id && u.role === 'admin',
        );
        return `<div class="sa-company-card">
        <div class="sa-company-icon">🏢</div>
        <div class="sa-company-info">
          <div class="sa-company-name">${escHtml(co.name)}</div>
          <div class="sa-company-meta">${admins.length} admin · ${coUsers.length} users · Created ${new Date(co.created_at).toLocaleDateString()}</div>
        </div>
        <div class="sa-company-actions">
          <button class="btn-secondary" style="font-size:11px;padding:6px 12px;" onclick="viewCompanyUsers('${co.id}','${escHtml(co.name)}')">👥 Users</button>
          <button class="btn-secondary" style="font-size:11px;padding:6px 10px;" onclick="openSAChangeAdminPw('${co.id}','${escHtml(co.name)}')">🔑 Admin PW</button>
          <button class="btn-danger" style="font-size:11px;padding:6px 10px;" onclick="deleteCompany('${co.id}','${escHtml(co.name)}')">✕</button>
        </div>
      </div>`;
      })
      .join('');
  } catch (e) {
    list.innerHTML = `<div style="color:var(--danger);padding:20px;">${e.message}</div>`;
  }
}

function openCreateCompany() {
  document.getElementById('sa-co-name').value = '';
  document.getElementById('sa-admin-name').value = '';
  document.getElementById('sa-admin-username').value = '';
  document.getElementById('sa-admin-password').value = '';
  openModal('sa-company-modal');
  setTimeout(() => document.getElementById('sa-co-name').focus(), 100);
}

async function saveCompany() {
  const name = document.getElementById('sa-co-name').value.trim();
  const adminName = document.getElementById('sa-admin-name').value.trim();
  const adminUser = document
    .getElementById('sa-admin-username')
    .value.trim()
    .toLowerCase()
    .replace(/\s/g, '');
  const adminPass = document.getElementById('sa-admin-password').value.trim();
  if (!name || !adminName || !adminUser || !adminPass) {
    toast('All fields are required.', 'error');
    return;
  }
  try {
    // 1. Create company
    const co = await API.post('/companies', { name });
    // 2. Create admin user for that company
    await API.post('/users', {
      name: adminName,
      username: adminUser,
      password: adminPass,
      role: 'admin',
      companyId: co.id,
    });
    closeModal('sa-company-modal');
    toast(`Company "${name}" created with admin account! ✓`, 'success');
    await loadSACompanies();
  } catch (e) {
    toast(e.message || 'Failed to create company.', 'error');
  }
}

async function deleteCompany(id, name) {
  if (
    !confirm(
      `Delete company "${name}" and ALL their data (users, tasks, teams)? This cannot be undone.`,
    )
  )
    return;
  try {
    // Delete all company data in order
    await SB.deleteWhere('tf_tasks', `company_id=eq.${id}`);
    await SB.deleteWhere('tf_teams', `company_id=eq.${id}`);
    await SB.deleteWhere('tf_leaves', `company_id=eq.${id}`);
    await SB.deleteWhere('tf_logs', `company_id=eq.${id}`);
    await SB.deleteWhere('tf_notifications', `company_id=eq.${id}`);
    await SB.deleteWhere('tf_schedules', `company_id=eq.${id}`);
    await SB.deleteWhere('tf_users', `company_id=eq.${id}`);
    await API.del('/companies/' + id);
    toast(`Company "${name}" deleted.`, 'info');
    await loadSACompanies();
  } catch (e) {
    toast(e.message || 'Failed to delete company.', 'error');
  }
}

async function viewCompanyUsers(companyId, companyName) {
  document.getElementById('sa-users-modal-title').textContent =
    '👥 ' + companyName + ' — Users';
  const list = document.getElementById('sa-users-list');
  list.innerHTML =
    '<div style="text-align:center;padding:20px;color:var(--text3);">Loading...</div>';
  openModal('sa-users-modal');
  try {
    const users = await SB.select(
      'tf_users',
      `company_id=eq.${companyId}&select=id,name,username,role,created_at`,
    );
    const roleColors = {
      admin: '#A78BFA',
      manager: '#38BDF8',
      user: 'var(--p4)',
    };
    const roleLabels = {
      admin: '🛡 Admin',
      manager: '🏢 Manager',
      user: '👤 User',
    };
    list.innerHTML =
      (users || [])
        .map((u) => {
          const rc = roleColors[u.role] || 'var(--p4)';
          return `<div style="display:flex;align-items:center;gap:12px;padding:10px 12px;background:var(--bg3);border:1px solid var(--border);border-radius:8px;">
        <div style="width:34px;height:34px;border-radius:50%;background:${rc}18;border:1px solid ${rc}44;display:flex;align-items:center;justify-content:center;font-size:14px;font-weight:700;color:${rc};flex-shrink:0;">${escHtml((u.name || '?')[0].toUpperCase())}</div>
        <div style="flex:1;min-width:0;">
          <div style="font-weight:600;font-size:13px;">${escHtml(u.name)}</div>
          <div style="font-size:11px;color:var(--text3);">@${escHtml(u.username)}</div>
        </div>
        <span style="font-size:10px;font-weight:700;padding:3px 8px;border-radius:10px;background:${rc}18;color:${rc};border:1px solid ${rc}30;">${roleLabels[u.role] || u.role}</span>
      </div>`;
        })
        .join('') ||
      '<div style="text-align:center;padding:20px;color:var(--text3);">No users yet</div>';
  } catch (e) {
    list.innerHTML = `<div style="color:var(--danger);padding:12px;">${e.message}</div>`;
  }
}

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2);
}

/* ============================================================
   SUPER-ADMIN: Change Admin Password
   ============================================================ */
async function openSAChangeAdminPw(companyId, companyName) {
  document.getElementById('sa-chpw-info').textContent =
    'Set a new password for an admin account of "' + companyName + '".';
  document.getElementById('sa-chpw-new').value = '';
  document.getElementById('sa-chpw-confirm').value = '';
  const sel = document.getElementById('sa-chpw-admin-select');
  sel.innerHTML = '<option value="">Loading...</option>';
  openModal('sa-chpw-modal');
  try {
    const users = await SB.select(
      'tf_users',
      `company_id=eq.${companyId}&role=eq.admin&select=id,name,username`,
    );
    if (!users || users.length === 0) {
      sel.innerHTML = '<option value="">No admin found</option>';
    } else {
      sel.innerHTML = users
        .map(
          (u) =>
            `<option value="${u.id}">${escHtml(u.name)} (@${escHtml(u.username)})</option>`,
        )
        .join('');
    }
  } catch (e) {
    sel.innerHTML = '<option value="">Error loading admins</option>';
  }
}

async function saveSAAdminPw() {
  const adminId = document.getElementById('sa-chpw-admin-select').value;
  const newPw = document.getElementById('sa-chpw-new').value.trim();
  const confPw = document.getElementById('sa-chpw-confirm').value.trim();
  if (!adminId) {
    toast('Please select an admin account.', 'error');
    return;
  }
  if (!newPw) {
    toast('Please enter a new password.', 'error');
    return;
  }
  if (newPw !== confPw) {
    toast('Passwords do not match.', 'error');
    return;
  }
  if (newPw.length < 4) {
    toast('Password must be at least 4 characters.', 'error');
    return;
  }
  try {
    await SB.update('tf_users', adminId, { password_hash: newPw });
    closeModal('sa-chpw-modal');
    toast('Admin password updated successfully. ✓', 'success');
  } catch (e) {
    toast(e.message || 'Failed to update password.', 'error');
  }
}

function initDB() {
  /* no-op: handled by backend */
}

/* ============================================================
   STATE
   ============================================================ */
let state = {
  currentUser: null,
  view: 'login',
  previousView: null,
  targetUserId: null,
  editingTaskId: null,
  editingLeaveId: null,
  contextTaskId: null,
  deleteTaskId: null,
  cancelTaskId: null,
  waParsed: null,
  descItems: [],
  timelineScale: 'week',
  currentViewMode: 'board', // 'board' or 'timeline'
  calendarMode: 'personal', // 'personal' or 'team'
  _returnToCalendar: false,
};

var gcalState = {
  mode: 'week', // 'day' | 'week' | 'month'
  anchorDate: new Date(), // reference date for current view
  hourHeight: 60, // px per hour — basis for all positioning
  startHour: 7, // first visible hour
  endHour: 23, // last visible hour
};

/* ============================================================
   GLOBAL OPERATION LOCK — prevents double-submission on all
   async action buttons. Each operation has a unique key.
   A locked key means the operation is already in-flight.
   ============================================================ */
const _opLocks = new Set();

function _lockOp(key, btn, busyText) {
  if (_opLocks.has(key)) return false;
  _opLocks.add(key);
  if (btn) {
    btn.disabled = true;
    btn._origText = btn.textContent;
    if (busyText) btn.textContent = busyText;
  }
  // Show global overlay with process name centred on screen
  const ov = document.getElementById('global-overlay');
  const tx = document.getElementById('global-overlay-text');
  if (ov) {
    if (tx) tx.textContent = busyText || 'Processing…';
    ov.style.display = 'flex';
  }
  return true;
}

function _unlockOp(key, btn) {
  _opLocks.delete(key);
  if (btn) {
    btn.disabled = false;
    if (btn._origText !== undefined) btn.textContent = btn._origText;
  }
  // Hide overlay only when ALL locks released
  if (_opLocks.size === 0) {
    const ov = document.getElementById('global-overlay');
    if (ov) ov.style.display = 'none';
  }
}

/* ── Nav history for Home / Back / Next header buttons ── */
const _navHist = [];
let _navIdx = -1;

function _navRecord(key) {
  if (_navHist[_navIdx] === key) return; // same page, skip
  _navHist.splice(_navIdx + 1); // drop forward stack
  _navHist.push(key);
  _navIdx = _navHist.length - 1;
  _navRefresh();
}

function _navRefresh() {
  const back = document.getElementById('nav-back-btn');
  const next = document.getElementById('nav-next-btn');
  const canBack = _navIdx > 0;
  const canNext = _navIdx < _navHist.length - 1;
  if (back) {
    back.style.opacity = canBack ? '1' : '0.35';
    back.style.pointerEvents = canBack ? 'auto' : 'none';
  }
  if (next) {
    next.style.opacity = canNext ? '1' : '0.35';
    next.style.pointerEvents = canNext ? 'auto' : 'none';
  }
}

function _navJump(key) {
  if (!key) return;
  if (key.startsWith('user-tasks:')) {
    const uid = key.slice('user-tasks:'.length);
    state.targetUserId = uid;
    state.view = 'user-tasks';
    showView('user-tasks');
  } else {
    showView(key);
  }
}

function headerNavHome() {
  const r = state.currentUser?.role;
  if (r === 'admin' || r === 'manager') goAdminHome();
  else showView('worker-home');
}

function headerNavBack() {
  if (_navIdx <= 0) return;
  _navIdx--;
  _navRefresh();
  _navJump(_navHist[_navIdx]);
}

function headerNavNext() {
  if (_navIdx >= _navHist.length - 1) return;
  _navIdx++;
  _navRefresh();
  _navJump(_navHist[_navIdx]);
}
async function signIn() {
  const loginBtn = document.getElementById('login-btn');
  if (!_lockOp('login', loginBtn, '⏳ Signing in...')) return;

  const username = document.getElementById('login-username').value.trim();
  const password = document.getElementById('login-password').value.trim();
  const errEl = document.getElementById('login-error');
  errEl.classList.add('hidden');

  try {
    const data = await API.post('/auth/login', { username, password });
    if (!data) {
      _unlockOp('login', loginBtn);
      return;
    }
    API.setToken(data.token);
    state.currentUser = data.user;

    (EL.loginScreen || document.getElementById('login-screen')).classList.add(
      'hidden',
    );

    if (data.user.role === 'superadmin') {
      // Super-admin goes to their own panel, not the main app
      document.getElementById('sa-panel').classList.remove('hidden');
      await loadSACompanies();
      _unlockOp('login', loginBtn);
    } else {
      document.getElementById('app').classList.remove('hidden');
      setupHeader();
      requestNotifPermission();
      registerSW();
      startPolling();
      startRealtimeNotifs();
      await loadAll();
      _unlockOp('login', loginBtn);
    }

    if (data.user.role !== 'superadmin') {
      if (data.user.role === 'admin' || data.user.role === 'manager') {
        showView('admin-home');
      } else {
        // Regular users get calendar button in header
        const calBtn = document.getElementById('my-cal-btn');
        if (calBtn) calBtn.classList.remove('hidden');
        showView('worker-home');
      }
    }
  } catch (err) {
    errEl.textContent = err.message || 'Invalid credentials';
    errEl.classList.remove('hidden');
    _unlockOp('login', loginBtn);
  }
}

/**
 * Sign out the current user and return to the login screen.
 * Clears all cached data, stops polling, and resets the UI.
 * @returns {Promise<void>}
 */
async function signOut() {
  try {
    await API.post('/auth/logout');
  } catch { }
  API.clearToken();
  stopPolling();
  stopRealtimeNotifs();
  state.currentUser = null;
  state.view = 'login';
  Object.keys(cache).forEach((k) => (cache[k] = null));
  document.getElementById('app').classList.add('hidden');
  document.getElementById('sa-panel').classList.add('hidden');
  (EL.loginScreen || document.getElementById('login-screen')).classList.remove(
    'hidden',
  );
  document.getElementById('login-username').value = '';
  document.getElementById('login-password').value = '';
  document.getElementById('login-error').classList.add('hidden');
  clearInterval(deadlineInterval);
  // Clear saved view state
  localStorage.removeItem('tf_view');
  localStorage.removeItem('tf_target_uid');
  history.replaceState(null, '', location.pathname);
}

/* ============================================================
   HEADER
   ============================================================ */
function setupHeader() {
  const u = state.currentUser;
  document.getElementById('header-username').textContent = u.name;
  const roleEl = document.getElementById('header-role');
  const roleLabels = { admin: 'Admin', manager: 'Manager', user: 'User' };
  roleEl.textContent = roleLabels[u.role] || u.role.toUpperCase();
  const roleClass = {
    admin: 'role-admin',
    manager: 'role-manager',
    user: 'role-user',
  };
  roleEl.className = 'header-role ' + (roleClass[u.role] || 'role-user');
  // Settings (change password) button — admin only
  const chpwBtn = document.getElementById('header-chpw-btn');
  if (chpwBtn) chpwBtn.classList.toggle('hidden', u.role !== 'admin');
  updateNotifBadge();
}

/* ============================================================
   VIEW ROUTING — with hash-based URLs and persistence
   ============================================================ */
const views = [
  'admin-home',
  'worker-home',
  'task-board',
  'timeline-view',
  'leave-calendar-view',
  'user-list-view',
  'calendar-view',
  'teams-view',
];

// Map view names to URL-friendly hashes
var _viewToHash = {
  'admin-home': 'home',
  'worker-home': 'home',
  'my-tasks': 'tasks',
  'user-list': 'users',
  'user-tasks': 'user-tasks',
  calendar: 'calendar',
  teams: 'teams',
  'leave-calendar': 'leaves',
};
var _hashToView = {
  home: null, // resolved by role
  tasks: 'my-tasks',
  users: 'user-list',
  'user-tasks': 'user-tasks',
  calendar: 'calendar',
  teams: 'teams',
  leaves: 'leave-calendar',
};

function _saveViewState(v) {
  localStorage.setItem('tf_view', v);
  if (state.targetUserId)
    localStorage.setItem('tf_target_uid', state.targetUserId);
  // Update URL hash without triggering hashchange handler
  var hash = _viewToHash[v] || v;
  if (location.hash !== '#' + hash) {
    history.replaceState(null, '', '#' + hash);
  }
}

function _getSavedView() {
  // Check URL hash first, then localStorage
  var hash = location.hash.replace('#', '');
  if (hash && _hashToView[hash] !== undefined) {
    var mapped = _hashToView[hash];
    if (mapped === null) return null; // role-dependent home
    return mapped;
  }
  return localStorage.getItem('tf_view') || null;
}

// Listen for browser back/forward
window.addEventListener('hashchange', function () {
  if (!state.currentUser) return;
  var hash = location.hash.replace('#', '');
  var mapped = _hashToView[hash];
  if (mapped === null) {
    // home — resolve by role
    mapped =
      state.currentUser.role === 'admin' || state.currentUser.role === 'manager'
        ? 'admin-home'
        : 'worker-home';
  }
  if (mapped && mapped !== state.view) {
    if (mapped === 'user-tasks') {
      var savedUid = localStorage.getItem('tf_target_uid');
      if (savedUid) {
        state.targetUserId = savedUid;
      }
    }
    showView(mapped);
  }
});

/**
 * Switch the visible view in the application.
 * Handles hiding/showing view containers, updating breadcrumbs,
 * mobile nav state, URL hash, and triggering necessary renders.
 * @param {string} v - View identifier (e.g. 'admin-home', 'my-tasks', 'calendar', 'teams')
 */
function showView(v) {
  state.previousView = state.view;
  state.view = v;
  _saveViewState(v);
  views.forEach((id) => document.getElementById(id)?.classList.add('hidden'));
  (
    EL.boardFilterBar || document.getElementById('board-filter-bar')
  )?.classList.remove('visible');
  (EL.addTaskFab || document.getElementById('add-task-fab')).classList.add(
    'hidden',
  );
  document.getElementById('breadcrumb').classList.add('hidden');
  // Restore mobile "New Task" button when leaving calendar
  if (v !== 'calendar') {
    var mobAddBtn = document.getElementById('mob-nav-add');
    if (mobAddBtn) mobAddBtn.style.display = '';
  }
  const isElevated =
    state.currentUser.role === 'admin' || state.currentUser.role === 'manager';
  // Track nav for Back/Next buttons (skip 'user-tasks' here — viewUserTasks records with uid)
  if (v !== 'user-tasks') _navRecord(v);

  if (v === 'worker-home') {
    document.getElementById('worker-home').classList.remove('hidden');
    document.getElementById('worker-name-display').textContent =
      state.currentUser.name.split(' ')[0];
    updateMobileNavActive('mob-nav-home');
    // Render groups strip
    const myTeams = getTeams().filter((t) =>
      (t.memberIds || []).includes(state.currentUser.id),
    );
    const strip = document.getElementById('worker-groups-strip');
    if (strip) {
      if (myTeams.length === 0) {
        strip.style.display = 'none';
      } else {
        strip.style.display = 'flex';
        // Keep the header label, replace rest
        const header =
          '<div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.08em;color:var(--text3);text-align:center;margin-bottom:4px;">My Teams</div>';
        const cards = myTeams
          .map((team) => {
            const members = (team.memberIds || [])
              .map((id) => getUsers().find((u) => u.id === id))
              .filter(Boolean);
            const avatars = members
              .slice(0, 6)
              .map(
                (m) =>
                  `<div title="${escHtml(m.name)}" style="width:26px;height:26px;border-radius:50%;background:${team.color || '#F59E0B'}33;border:1.5px solid ${team.color || '#F59E0B'};display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;color:${team.color || '#F59E0B'};flex-shrink:0;">${escHtml((m.name || '?')[0].toUpperCase())}</div>`,
              )
              .join('');
            const extra =
              members.length > 6
                ? `<div style="font-size:10px;color:var(--text3);margin-left:2px;">+${members.length - 6}</div>`
                : '';
            return `<div style="display:flex;align-items:center;gap:10px;padding:10px 14px;background:var(--bg2);border:1px solid ${team.color || '#F59E0B'}33;border-radius:12px;">
            <div style="width:10px;height:10px;border-radius:50%;background:${team.color || '#F59E0B'};flex-shrink:0;"></div>
            <div style="font-size:12px;font-weight:700;color:var(--text);white-space:nowrap;min-width:60px;max-width:120px;overflow:hidden;text-overflow:ellipsis;">${escHtml(team.name)}</div>
            <div style="display:flex;gap:3px;align-items:center;flex-wrap:wrap;flex:1;">${avatars}${extra}</div>
          </div>`;
          })
          .join('');
        strip.innerHTML = header + cards;
      }
    }
  } else if (v === 'admin-home') {
    document.getElementById('admin-home').classList.remove('hidden');
    document.getElementById('admin-name-display').textContent =
      state.currentUser.name.split(' ')[0];
    updateMobileNavActive('mob-nav-home');
    // Hide Register User and Leave Calendar cards for non-admins
    const regCard = document.querySelector(
      '.admin-nav-card[onclick="openRegisterUser()"]',
    );
    const leaveCard = document.querySelector(
      '.admin-nav-card[onclick="goToLeaveCalendar()"]',
    );
    const chpwCard = document.getElementById('admin-chpw-card');
    const settingsCard = document.getElementById('admin-settings-card');
    const mgrMultiCard = document.getElementById('manager-multi-card');
    const multiTeamCard = document.getElementById('multi-team-task-card');
    const importCard = document.getElementById('admin-import-card');
    const deleteAllCard = document.getElementById('admin-deleteall-card');
    if (regCard)
      regCard.style.display = state.currentUser.role === 'admin' ? '' : 'none';
    if (leaveCard) leaveCard.style.display = isElevated ? '' : 'none';
    if (chpwCard)
      chpwCard.style.display = state.currentUser.role === 'admin' ? '' : 'none';
    if (settingsCard)
      settingsCard.style.display =
        state.currentUser.role === 'admin' ? '' : 'none';
    if (importCard)
      importCard.style.display =
        state.currentUser.role === 'admin' ? '' : 'none';
    if (deleteAllCard)
      deleteAllCard.style.display =
        state.currentUser.role === 'admin' ? '' : 'none';
    if (mgrMultiCard) mgrMultiCard.style.display = 'none'; // removed
    if (multiTeamCard) multiTeamCard.style.display = isElevated ? '' : 'none';
  } else if (v === 'my-tasks') {
    if (state.currentViewMode === 'timeline') {
      showTimelineView(state.currentUser.id);
    } else {
      showBoardView(state.currentUser.id);
    }
  } else if (v === 'user-list') {
    document.getElementById('user-list-view').classList.remove('hidden');
    // Hide calendar when showing user list
    document.getElementById('calendar-view').classList.add('hidden');
    setBreadcrumb([
      { label: 'Home', fn: 'goAdminHome' },
      { label: 'User Tasks' },
    ]);
    renderUserList();
  } else if (v === 'user-tasks') {
    if (state.currentViewMode === 'timeline') {
      showTimelineView(state.targetUserId);
    } else {
      showBoardView(state.targetUserId);
    }
  } else if (v === 'leave-calendar') {
    document.getElementById('leave-calendar-view').classList.remove('hidden');
    setBreadcrumb([
      { label: 'Home', fn: 'goAdminHome' },
      { label: 'Calendar' },
    ]);
    renderLeaveCalendar();
  } else if (v === 'teams') {
    document.getElementById('teams-view').classList.remove('hidden');
    // Hide calendar when showing teams view
    document.getElementById('calendar-view').classList.add('hidden');
    setBreadcrumb([{ label: 'Home', fn: 'goAdminHome' }, { label: 'Teams' }]);
    const teamAssignBtn = document.getElementById('team-assign-btn');
    if (teamAssignBtn) teamAssignBtn.style.display = isElevated ? '' : 'none';
    // Highlight Teams tab in mobile nav
    updateMobileNavActive('mob-nav-teams');
    renderTeamsView();
  } else if (v === 'calendar') {
    document.getElementById('calendar-view').classList.remove('hidden');
    const isPersonal = state.calendarMode === 'personal';
    const label = isPersonal ? 'My Calendar' : 'Calendar';
    var homeFn =
      state.currentUser.role === 'user' ? 'goWorkerHome' : 'goAdminHome';
    setBreadcrumb([
      { label: 'Home', fn: isPersonal ? homeFn : homeFn },
      { label },
    ]);

    // Show breadcrumbs on desktop; hide on mobile (browser back handles navigation)
    var breadcrumb = document.getElementById('breadcrumb');
    var isMobileView = window.innerWidth <= 768 || 'ontouchstart' in window;
    if (breadcrumb) {
      if (isMobileView) {
        breadcrumb.classList.add('hidden');
      } else {
        breadcrumb.classList.remove('hidden');
      }
    }

    // Activate Calendar tab in mobile nav
    updateMobileNavActive('mob-nav-cal');
    // Hide mobile "New Task" button in calendar view
    var mobAddBtn = document.getElementById('mob-nav-add');
    if (mobAddBtn) mobAddBtn.style.display = 'none';

    // Add back button to calendar header (desktop only)
    var calHeader = document.querySelector('.calendar-header');
    if (
      calHeader &&
      !document.getElementById('cal-back-btn') &&
      !isMobileView
    ) {
      var backBtn = document.createElement('button');
      backBtn.id = 'cal-back-btn';
      backBtn.className = 'btn-secondary';
      backBtn.innerHTML = '← Back';
      backBtn.style.cssText =
        'padding:8px 12px;font-size:12px;margin-right:8px;';
      backBtn.onclick = function () {
        if (state.previousView && state.previousView !== 'calendar') {
          showView(state.previousView);
        } else if (
          state.currentUser.role === 'admin' ||
          state.currentUser.role === 'manager'
        ) {
          showView('admin-home');
        } else {
          showView('my-tasks');
        }
      };
      calHeader.insertBefore(backBtn, calHeader.firstChild);
    }
    // Reset to today's month on every open (unless jumping to a specific task date)
    if (!state._calendarJumpDate) {
      var now = new Date();
      calState.year = now.getFullYear();
      calState.month = now.getMonth();
      calState.zoom = 'day';
    }

    // Force all calendar sub-containers to a known good state before rendering
    var _sw = document.getElementById('cal-split-wrapper');
    var _gcal = document.getElementById('gcal-view');
    var isGcal = localStorage.getItem('tf_gcal_active') === '1';

    if (_gcal) _gcal.classList.toggle('hidden', !isGcal);
    if (_sw) _sw.style.display = isGcal ? 'none' : 'flex';

    var _yg = document.getElementById('cal-year-grid');
    if (_yg) _yg.remove();
    var _lt = document.getElementById('cal-table');
    if (_lt) _lt.style.display = 'none';
    var _ws = document.getElementById('cal-week-strip');
    if (_ws) _ws.classList.toggle('hidden', isGcal);

    // Update toggle button text/state
    const gcalBtn = document.getElementById('cal-gcal-toggle-btn');
    if (gcalBtn) {
      gcalBtn.textContent = isGcal ? '⬅ Grid View' : '📅 Task View';
      gcalBtn.classList.toggle('active', isGcal);
    }

    renderCalendar();
    // Scroll to today's column after rendering (when not jumping to a task)
    if (!state._calendarJumpDate) {
      var todayDateStr = new Date().toISOString().slice(0, 10);
      setTimeout(function () {
        var todayCol = document.querySelector(
          '#cal-dates-panel [data-date="' + todayDateStr + '"]',
        );
        if (todayCol)
          todayCol.scrollIntoView({
            behavior: 'smooth',
            block: 'nearest',
            inline: 'center',
          });
      }, 80);
    }

    // If we jumped from a task, scroll to the task's date column (and worker's row)
    if (state._calendarJumpDate) {
      var jumpDate = state._calendarJumpDate;
      var highlightUserId = state._calendarHighlightUserId || null;
      var focusTaskId = state._calendarFocusTaskId || null;
      calState.year = jumpDate.getFullYear();
      calState.month = jumpDate.getMonth();
      state._calendarJumpDate = null;
      state._calendarHighlightUserId = null;
      state._calendarFocusTaskId = null;
      renderCalendar();
      setTimeout(function () {
        // Scroll to and highlight the specific task block if available
        var focused = false;
        if (focusTaskId) {
          var taskBlock = document.querySelector(
            '.cal-block[data-taskid="' + focusTaskId + '"]',
          );
          if (taskBlock) {
            taskBlock.scrollIntoView({
              behavior: 'smooth',
              block: 'center',
              inline: 'center',
            });
            taskBlock.style.outline = '2px solid var(--amber)';
            taskBlock.style.outlineOffset = '2px';
            taskBlock.style.zIndex = '100';
            setTimeout(function () {
              taskBlock.style.outline = '';
              taskBlock.style.outlineOffset = '';
              taskBlock.style.zIndex = '';
            }, 3000);
            focused = true;
          }
        }
        // Fallback: scroll to the date column header
        if (!focused) {
          var taskDateStr = jumpDate.toISOString().slice(0, 10);
          var cols = document.querySelectorAll('#cal-dates-panel [data-date]');
          for (var i = 0; i < cols.length; i++) {
            if (cols[i].dataset.date === taskDateStr) {
              cols[i].scrollIntoView({
                behavior: 'smooth',
                block: 'nearest',
                inline: 'center',
              });
              cols[i].style.outline = '2px solid var(--amber)';
              setTimeout(
                function (el) {
                  el.style.outline = '';
                }.bind(null, cols[i]),
                3000,
              );
              break;
            }
          }
        }
        // If viewing a worker's row, scroll to and highlight that row
        if (highlightUserId) {
          var rows = document.querySelectorAll(
            '#cal-names-table tr[data-user-id]',
          );
          for (var r = 0; r < rows.length; r++) {
            if (rows[r].dataset.userId === highlightUserId) {
              rows[r].scrollIntoView({ behavior: 'smooth', block: 'center' });
              rows[r].style.outline = '2px solid var(--amber)';
              setTimeout(
                function (el) {
                  el.style.outline = '';
                }.bind(null, rows[r]),
                3000,
              );
              break;
            }
          }
        }
      }, 100);
    }
  }
}

function showBoardView(userId) {
  document.getElementById('task-board').classList.remove('hidden');
  // Hide calendar when showing task board
  document.getElementById('calendar-view').classList.add('hidden');
  (EL.addTaskFab || document.getElementById('add-task-fab')).classList.remove(
    'hidden',
  );
  (
    EL.boardFilterBar || document.getElementById('board-filter-bar')
  ).classList.add('visible');
  _boardSearchVal = '';
  _boardFilter = 'all';
  const inp = document.getElementById('board-search');
  if (inp) inp.value = '';
  const clr = document.getElementById('board-search-clear');
  if (clr) clr.classList.add('hidden');
  document
    .querySelectorAll('.board-filter-pill')
    .forEach((p) => p.classList.toggle('active', p.dataset.filter === 'all'));
  const isElevated =
    state.currentUser.role === 'admin' || state.currentUser.role === 'manager';
  if (isElevated) {
    if (state.view === 'user-tasks') {
      const targetUser = getUsers().find((u) => u.id === state.targetUserId);
      setBreadcrumb([
        { label: 'Home', fn: 'goAdminHome' },
        { label: 'User Tasks', fn: 'goToUserList' },
        { label: targetUser?.name || 'User' },
      ]);
    } else {
      setBreadcrumb([
        { label: 'Home', fn: 'goAdminHome' },
        { label: 'My Tasks' },
      ]);
    }
  } else {
    setBreadcrumb([
      { label: 'Home', fn: 'goWorkerHome' },
      { label: 'My Tasks' },
    ]);
  }
  // Highlight Tasks tab in mobile nav
  updateMobileNavActive('mob-nav-tasks');
  renderTasks(userId);
  startDeadlineChecker();
}

function showTimelineView(userId) {
  document.getElementById('timeline-view').classList.remove('hidden');
  // Hide calendar when showing timeline view
  document.getElementById('calendar-view').classList.add('hidden');
  (EL.addTaskFab || document.getElementById('add-task-fab')).classList.remove(
    'hidden',
  );
  // Hide add-task FAB on mobile for timeline view
  if (window.innerWidth <= 768 || 'ontouchstart' in window) {
    (EL.addTaskFab || document.getElementById('add-task-fab')).classList.add(
      'hidden',
    );
  }
  const isElevated =
    state.currentUser.role === 'admin' || state.currentUser.role === 'manager';
  const targetUser = getUsers().find((u) => u.id === userId);
  document.getElementById('timeline-title').textContent =
    isElevated && state.view === 'user-tasks'
      ? `Timeline: ${targetUser?.name}`
      : 'My Timeline';
  if (isElevated) {
    if (state.view === 'user-tasks') {
      setBreadcrumb([
        { label: 'Home', fn: 'goAdminHome' },
        { label: 'User Tasks', fn: 'goToUserList' },
        { label: targetUser?.name || 'User' },
      ]);
    } else {
      setBreadcrumb([
        { label: 'Home', fn: 'goAdminHome' },
        { label: 'My Timeline' },
      ]);
    }
  } else {
    setBreadcrumb([
      { label: 'Home', fn: 'goWorkerHome' },
      { label: 'My Timeline' },
    ]);
  }
  renderTimeline(userId);
  startDeadlineChecker();
}

var calSearchFilter = '';

function onCalSearch(val) {
  calSearchFilter = (val || '').trim().toLowerCase();
  const clearBtn = document.getElementById('cal-search-clear');
  if (clearBtn) clearBtn.classList.toggle('hidden', !calSearchFilter);
  if (calState.zoom === 'day') renderCalendar();
}

function clearCalSearch() {
  calSearchFilter = '';
  const inp = document.getElementById('cal-search-input');
  if (inp) inp.value = '';
  const clearBtn = document.getElementById('cal-search-clear');
  if (clearBtn) clearBtn.classList.add('hidden');
  renderCalendar();
}

function calToggleTeamView() {
  const isUser = state.currentUser.role === 'user';
  if (isUser) {
    // Users can toggle between personal and team (read-only)
    state.calendarMode =
      state.calendarMode === 'personal' ? 'team' : 'personal';
  } else {
    state.calendarMode = state.calendarMode === 'team' ? 'personal' : 'team';
  }

  if (state.calendarMode === 'personal') {
    const gcalWrap = document.getElementById('gcal-view');
    if (gcalWrap && !gcalWrap.classList.contains('hidden')) {
      calToggleGcalView();
    }
  }

  updateCalHeaderUI();
  renderCalendar();
}

function updateCalHeaderUI() {
  const isPersonal = state.calendarMode === 'personal';
  const isUser = state.currentUser?.role === 'user';
  const toggleBtn = document.getElementById('cal-team-toggle-btn');
  const gcalBtn = document.getElementById('cal-gcal-toggle-btn');
  const searchBar = document.getElementById('cal-search-bar');
  const workerApplyBtn = $('worker-apply-leave-btn');
  const addLeaveBtn = $('add-leave-btn');
  const isElevated =
    state.currentUser?.role === 'admin' ||
    state.currentUser?.role === 'manager';

  // Toggle button text
  if (toggleBtn) {
    toggleBtn.style.display = '';
    toggleBtn.textContent = isPersonal ? '👥 Team View' : '👤 My View';
  }

  // GCal toggle button visibility
  if (gcalBtn) {
    gcalBtn.style.display = (state.view === 'calendar' && !isPersonal) ? '' : 'none';
  }

  // Search only visible in day view + team mode
  if (searchBar)
    searchBar.classList.toggle(
      'hidden',
      isPersonal || calState.zoom === 'year',
    );

  // Manager/Admin button - only in team view for elevated roles
  if (addLeaveBtn)
    addLeaveBtn.style.display = isElevated && !isPersonal ? '' : 'none';

  // Worker button - only for regular users
  if (workerApplyBtn)
    workerApplyBtn.style.display = isUser ? '' : 'none';
}
function goToUserList() {
  showView('user-list');
}
function goAdminHome() {
  showView('admin-home');
}
/** Navigate the current user to their personal task board. */
function goToMyTasks() {
  state.targetUserId = null;
  state.view = 'my-tasks';
  showView('my-tasks');
}
/** Navigate to the team leave calendar view. */
function goToLeaveCalendar() {
  state.calendarMode = 'team';
  showView('calendar');
}
function goToMyCalendar() {
  state.calendarMode = 'personal';
  const gcalWrap = document.getElementById('gcal-view');
  if (gcalWrap && !gcalWrap.classList.contains('hidden')) {
    calToggleGcalView();
  }
  showView('calendar');
}
function goWorkerHome() {
  showView('worker-home');
}

function setBreadcrumb(items) {
  const bc = document.getElementById('breadcrumb');
  bc.classList.remove('hidden');
  bc.innerHTML = items
    .map((item, i) => {
      if (item.fn)
        return `<span class="breadcrumb-item" onclick="${item.fn}()">${item.label}</span>`;
      return `<span class="breadcrumb-current">${item.label}</span>`;
    })
    .join('<span class="breadcrumb-sep"> / </span>');
}

function switchToBoard() {
  state.currentViewMode = 'board';
  const userId =
    state.view === 'user-tasks' ? state.targetUserId : state.currentUser.id;
  showBoardView(userId);
}

function switchToTimeline() {
  state.currentViewMode = 'timeline';
  const userId =
    state.view === 'user-tasks' ? state.targetUserId : state.currentUser.id;
  showTimelineView(userId);
}

function setTimelineScale(scale) {
  state.timelineScale = scale;
  document.querySelectorAll('.timeline-scale-btn').forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.scale === scale);
  });
  const userId =
    state.view === 'user-tasks' ? state.targetUserId : state.currentUser.id;
  renderTimeline(userId);
}

/* ============================================================
   LEAVE MANAGEMENT
   ============================================================ */
/* ============================================================
   WORKING DAYS & LEAVE BALANCE ENGINE
   ============================================================ */

// Get company-level settings (weekend working, leave quota)
function getCompanySettings() {
  try {
    return JSON.parse(
      localStorage.getItem('tf_co_settings_' + (_cid() || 'default')) || '{}',
    );
  } catch {
    return {};
  }
}
function saveCompanySettings(obj) {
  const key = 'tf_co_settings_' + (_cid() || 'default');
  const cur = getCompanySettings();
  localStorage.setItem(key, JSON.stringify(Object.assign(cur, obj)));
}

// Called from worksettings modal checkboxes
function saveWeekendSettings() {
  const sat = document.getElementById('ws-sat-working')?.checked || false;
  const sun = document.getElementById('ws-sun-working')?.checked || false;
  saveCompanySettings({ satWorking: sat, sunWorking: sun });
  toast('Weekend settings saved.', 'success');
}
function loadWeekendSettings() {
  const s = getCompanySettings();
  const satEl = document.getElementById('ws-sat-working');
  const sunEl = document.getElementById('ws-sun-working');
  if (satEl) satEl.checked = !!s.satWorking;
  if (sunEl) sunEl.checked = !!s.sunWorking;
}

// Annual leave quota (working days, default 30)
const ANNUAL_LEAVE_DAYS = 30;

// Is a date a working day (Mon-Fri + optional Sat/Sun)?
function isWorkingDay(date) {
  const d = new Date(date);
  d.setHours(12, 0, 0, 0);
  const dow = d.getDay(); // 0=Sun,1=Mon...6=Sat
  const s = getCompanySettings();
  if (dow >= 1 && dow <= 5) return true; // Mon-Fri always
  if (dow === 6 && s.satWorking) return true; // Sat if enabled
  if (dow === 0 && s.sunWorking) return true; // Sun if enabled
  return false;
}

// Count working days between two date strings (inclusive)
function countWorkingDays(startStr, endStr) {
  const s = new Date(startStr + 'T00:00:00');
  const e = new Date(endStr + 'T23:59:59');
  let count = 0;
  const cur = new Date(s);
  while (cur <= e) {
    if (isWorkingDay(cur)) count++;
    cur.setDate(cur.getDate() + 1);
  }
  return count;
}

// Get all working-day dates for a user in the current year
function getWorkedDays(userId) {
  const year = new Date().getFullYear();
  const leaves = getLeaves().filter((l) => l.userId === userId);
  let worked = 0;
  const start = new Date(year, 0, 1);
  const end = new Date(year, 11, 31);
  const cur = new Date(start);
  while (cur <= end) {
    if (cur > new Date()) break; // only past days
    if (isWorkingDay(cur)) {
      const dayStr = cur.toISOString().slice(0, 10);
      // AWOL subtracts from working days
      const awol = leaves.find(
        (l) =>
          l.type === 'awol' && l.startDate <= dayStr && l.endDate >= dayStr,
      );
      if (!awol) worked++;
    }
    cur.setDate(cur.getDate() + 1);
  }
  return worked;
}

// Count accrued lieu days for a user (weekend days they worked)
// Lieu days EARNED = number of weekends the user worked (stored in userWorkdays)
function countLieuDaysEarned(userId) {
  const workdays = getUserCustomWorkdays(userId).length;
  const s = getCompanySettings();
  const bonus = s.lieuBonus && s.lieuBonus[userId] ? s.lieuBonus[userId] : 0;
  return workdays + bonus;
}
// Lieu days USED = lieu leave records that are NOT auto-created weekend markers
function countLieuDaysUsed(userId) {
  const leaves = getLeaves().filter(
    (l) =>
      l.userId === userId &&
      l.type === 'lieu' &&
      (l.status === 'approved' || !l.status) && // backwards compat
      !(l.reason || '').includes('auto lieu day'),
  );
  let total = 0;
  leaves.forEach((l) => {
    total += countWorkingDays(l.startDate, l.endDate);
  });
  return total;
}
// Available lieu balance
function countLieuDays(userId) {
  return Math.max(0, countLieuDaysEarned(userId) - countLieuDaysUsed(userId));
}

// Count used LOA days for a user this year
function countUsedLeaveDays(userId) {
  const year = new Date().getFullYear();
  const leaves = getLeaves().filter(
    (l) =>
      l.userId === userId &&
      l.type === 'loa' &&
      l.startDate &&
      l.startDate.startsWith(year.toString()) &&
      (l.status === 'approved' || !l.status),
  );
  let total = 0;
  leaves.forEach((l) => {
    total += countWorkingDays(l.startDate, l.endDate);
  });
  return total;
}

// Remaining LOA days
function remainingLeaveDays(userId) {
  return Math.max(0, ANNUAL_LEAVE_DAYS - countUsedLeaveDays(userId));
}

// Count working days this month for a user (excluding AWOL)
function countMonthlyWorkingDays(userId, year, month) {
  const leaves = getLeaves().filter((l) => l.userId === userId);
  let worked = 0;
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  for (let d = 1; d <= daysInMonth; d++) {
    const cur = new Date(year, month, d);
    if (cur > today) break;
    if (!isWorkingDay(cur)) continue;
    const dayStr = cur.toISOString().slice(0, 10);
    const awol = leaves.find(
      (l) => l.type === 'awol' && l.startDate <= dayStr && l.endDate >= dayStr,
    );
    if (!awol) worked++;
  }
  return worked;
}

/* ============================================================
   LEAVE MODAL — enhanced with balance checks
   ============================================================ */

function filterLeaveUsers(query) {
  const select = document.getElementById('leave-user');
  const options = Array.from(select.options);
  const lowerQuery = query.toLowerCase();

  options.forEach((opt) => {
    if (opt.value === '') {
      opt.style.display = ''; // Always show the placeholder
      return;
    }
    const text = opt.textContent.toLowerCase();
    opt.style.display = text.includes(lowerQuery) ? '' : 'none';
  });
}

function openAddLeave() {
  var role = state.currentUser ? state.currentUser.role : '';
  if (role !== 'manager' && role !== 'admin') {
    toast('Only managers can manage leave.', 'error');
    return;
  }
  state.editingLeaveId = null;
  document.getElementById('leave-modal-title').textContent = 'Add Leave';
  const _delBtn = document.getElementById('delete-leave-btn');
  _delBtn.classList.add('hidden');
  _delBtn.onclick = null;

  const userSelect = document.getElementById('leave-user');
  userSelect.innerHTML = '<option value="">Select employee...</option>';
  const users = getUsers().filter((u) => u.role !== 'admin');
  users.forEach((u) => {
    const option = document.createElement('option');
    option.value = u.id;
    option.textContent = u.name + ' (@' + u.username + ')';
    userSelect.appendChild(option);
  });

  document.getElementById('leave-type').value = 'lieu';
  const today = new Date();
  const todayStr = today.toISOString().slice(0, 10);
  document.getElementById('leave-start').value = todayStr;
  document.getElementById('leave-end').value = todayStr;
  document.getElementById('leave-reason').value = '';

  updateLeaveTypeOptions();
  openModal('leave-modal');
}

// Update leave type dropdown based on available balances
function updateLeaveTypeOptions() {
  const userId = document.getElementById('leave-user')?.value;
  const lieuSel = document.querySelector('#leave-type option[value="lieu"]');
  const loaSel = document.querySelector('#leave-type option[value="loa"]');
  if (!lieuSel || !loaSel) return;
  if (!userId) {
    lieuSel.disabled = false;
    loaSel.disabled = false;
    lieuSel.textContent = '🟢 Lieu Day — Compensatory time off';
    loaSel.textContent = '🔵 Leave of Absence — Authorized extended leave';
    return;
  }
  const lieuDays = countLieuDays(userId);
  const loaDays = remainingLeaveDays(userId);
  const s = getCompanySettings();
  const canLieu = s.satWorking || s.sunWorking || lieuDays > 0; // allow if balance exists (e.g. manually added)
  if (!canLieu || lieuDays <= 0) {
    lieuSel.disabled = true;
    lieuSel.textContent = '🟢 Lieu Day — None available';
    if (document.getElementById('leave-type').value === 'lieu')
      document.getElementById('leave-type').value = 'loa';
  } else {
    lieuSel.disabled = false;
    lieuSel.textContent = `🟢 Lieu Day — ${lieuDays} available`;
  }
  if (loaDays <= 0) {
    loaSel.disabled = true;
    loaSel.textContent = '🔵 Leave of Absence — None remaining';
    if (document.getElementById('leave-type').value === 'loa')
      document.getElementById('leave-type').value = 'awol';
  } else {
    loaSel.disabled = false;
    loaSel.textContent = `🔵 Leave of Absence — ${loaDays} days remaining`;
  }
}

// Validate and enforce leave limits before saving
function validateLeaveRequest(userId, type, startDate, endDate) {
  const requestedDays = countWorkingDays(startDate, endDate);
  const s = getCompanySettings();

  if (type === 'lieu') {
    const available = countLieuDays(userId);
    if (available <= 0)
      return {
        ok: false,
        msg: 'No lieu days available. Earn lieu days by working overtime or on assigned rest days.',
      };
    if (requestedDays > 10)
      return {
        ok: false,
        msg: 'Lieu day leave cannot exceed 10 consecutive working days.',
      };
    if (requestedDays > available) {
      // Partial — trim end date to available days
      let remaining = available;
      let cur = new Date(startDate + 'T00:00:00');
      let lastValid = startDate;
      while (remaining > 0) {
        if (isWorkingDay(cur)) {
          lastValid = cur.toISOString().slice(0, 10);
          remaining--;
        }
        cur.setDate(cur.getDate() + 1);
      }
      return {
        ok: 'partial',
        trimmedEnd: lastValid,
        used: available,
        msg: `Only ${available} lieu day(s) available. Leave will be set from ${startDate} to ${lastValid}.`,
      };
    }
    return { ok: true, days: requestedDays };
  }

  if (type === 'loa') {
    const remaining = remainingLeaveDays(userId);
    if (remaining <= 0)
      return {
        ok: false,
        msg: `No leave days remaining for this year (${ANNUAL_LEAVE_DAYS}-day annual entitlement fully used).`,
      };
    if (requestedDays > remaining) {
      let rem = remaining;
      let cur = new Date(startDate + 'T00:00:00');
      let lastValid = startDate;
      while (rem > 0) {
        if (isWorkingDay(cur)) {
          lastValid = cur.toISOString().slice(0, 10);
          rem--;
        }
        cur.setDate(cur.getDate() + 1);
      }
      return {
        ok: 'partial',
        trimmedEnd: lastValid,
        used: remaining,
        msg: `Only ${remaining} leave day(s) remaining. Leave will be set from ${startDate} to ${lastValid}.`,
      };
    }
    return { ok: true, days: requestedDays };
  }

  return { ok: true, days: requestedDays };
}

function openEditLeave(leaveId) {
  const leave = getLeaves().find((l) => l.id === leaveId);
  if (!leave) return;

  state.editingLeaveId = leaveId;
  document.getElementById('leave-modal-title').textContent = 'Edit Leave';
  const delBtn = document.getElementById('delete-leave-btn');
  delBtn.classList.remove('hidden');
  delBtn.onclick = function () {
    deleteLeaveById(leaveId);
  };

  const userSelect = document.getElementById('leave-user');
  userSelect.innerHTML = '<option value="">Select employee...</option>';
  const users = getUsers().filter((u) => u.role !== 'admin');
  users.forEach((u) => {
    const option = document.createElement('option');
    option.value = u.id;
    option.textContent = u.name + ' (@' + u.username + ')';
    if (u.id === leave.userId) option.selected = true;
    userSelect.appendChild(option);
  });

  document.getElementById('leave-type').value = leave.type;
  document.getElementById('leave-start').value =
    leave.startDate || (leave.start ? leave.start.slice(0, 10) : '');
  document.getElementById('leave-end').value =
    leave.endDate || (leave.end ? leave.end.slice(0, 10) : '');
  document.getElementById('leave-reason').value = leave.reason || '';

  updateLeaveTypeOptions();
  openModal('leave-modal');
}

async function saveLeave() {
  const _btn = document.querySelector(
    '#leave-modal .btn-primary[onclick="saveLeave()"]',
  );
  if (!_lockOp('saveLeave', _btn, 'Saving…')) return;
  const userId = document.getElementById('leave-user').value;
  let type = document.getElementById('leave-type').value;
  const startDate = document.getElementById('leave-start').value;
  let endDate = document.getElementById('leave-end').value;
  const reason = document.getElementById('leave-reason').value.trim();
  if (!userId || !startDate || !endDate) {
    toast('Please fill in all required fields.', 'error');
    _unlockOp('saveLeave', _btn);
    return;
  }
  const payload = {
    userId,
    type,
    startDate,
    endDate,
    reason,
    status: 'approved',
    requestedBy: state.currentUser.id,
  };
  if (startDate > endDate) {
    toast('End date must be on or after start date.', 'error');
    _unlockOp('saveLeave', _btn);
    return;
  }

  // Weekend check
  if (countWorkingDays(startDate, endDate) === 0) {
    toast('The selected date range falls entirely on weekends.', 'error');
    _unlockOp('saveLeave', _btn);
    return;
  }

  // Block if any day in the range already has a leave (for new leaves only)
  if (!state.editingLeaveId) {
    const existingLeaves = getLeaves().filter(
      (l) => l.userId === userId && l.status !== 'denied',
    );
    const cur = new Date(startDate + 'T12:00:00');
    const end = new Date(endDate + 'T12:00:00');
    while (cur <= end) {
      if (isWorkingDay(cur)) {
        const dStr = cur.toISOString().slice(0, 10);
        const clash = existingLeaves.find((l) => {
          const ls = l.startDate || (l.start ? l.start.slice(0, 10) : '');
          const le = l.endDate || (l.end ? l.end.slice(0, 10) : '');
          return ls <= dStr && le >= dStr;
        });
        if (clash) {
          toast(
            'Leave already exists on ' +
            dStr +
            '. Remove it first before re-adding.',
            'error',
          );
          _unlockOp('saveLeave', _btn);
          return;
        }
      }
      cur.setDate(cur.getDate() + 1);
    }
  }

  const user = getUsers().find((u) => u.id === userId);

  // Validate leave balance (only for new leaves)
  if (!state.editingLeaveId) {
    const v = validateLeaveRequest(userId, type, startDate, endDate);
    if (v.ok === false) {
      toast(v.msg, 'error');
      _unlockOp('saveLeave', _btn);
      return;
    }
    if (v.ok === 'partial') {
      endDate = v.trimmedEnd;
      toast(v.msg, 'warning');
    }
  }

  const fmtDate = (d) =>
    new Date(d + 'T12:00:00').toLocaleDateString('en-GB', {
      day: '2-digit',
      month: 'short',
      year: 'numeric',
    });
  const typeLabel =
    type === 'awol' ? 'AWOL' : type === 'loa' ? 'Leave of Absence' : 'Lieu Day';
  const days = countWorkingDays(startDate, endDate);

  try {
    if (state.editingLeaveId) {
      const updated = await API.put(`/leaves/${state.editingLeaveId}`, {
        userId,
        type,
        startDate,
        endDate,
        reason,
      });
      const idx = cache.leaves.findIndex((l) => l.id === state.editingLeaveId);
      if (idx !== -1) cache.leaves[idx] = updated;
      toast('Leave updated!', 'success');
    } else {
      const newLeave = await API.post('/leaves', {
        userId,
        type,
        startDate,
        endDate,
        reason,
      });
      cache.leaves.push(newLeave);

      const icon = type === 'awol' ? '🔴' : type === 'loa' ? '🔵' : '🟢';
      const notifBody =
        `${typeLabel} — ${fmtDate(startDate)}${endDate !== startDate ? ' to ' + fmtDate(endDate) : ''} (${days} working day${days !== 1 ? 's' : ''})` +
        (reason ? '. Note: ' + reason : '');

      // In-app notification
      pushNotification(
        userId,
        `${icon} ${typeLabel} Recorded`,
        notifBody,
        newLeave.id,
        { type: 'leave', leaveType: type },
      );

      // Email notification (fire & forget)
      const emailUser = user || {};
      if (emailUser.emailNotif !== false && emailUser.email) {
        sendLeaveEmail(emailUser, {
          type,
          startDate,
          endDate,
          reason,
          days,
          typeLabel,
          fmtDate,
          icon,
        });
      }

      // Notify all managers/admins too (for AWOL)
      if (type === 'awol') {
        getUsers()
          .filter(
            (u) =>
              (u.role === 'manager' || u.role === 'admin') &&
              u.id !== state.currentUser.id,
          )
          .forEach((mgr) => {
            pushNotification(
              mgr.id,
              '🔴 AWOL Recorded — ' + (user?.name || 'User'),
              (user?.name || 'Employee') +
              ' is recorded AWOL from ' +
              fmtDate(startDate) +
              (endDate !== startDate ? ' to ' + fmtDate(endDate) : ''),
              newLeave.id,
              { type: 'leave', leaveType: 'awol' },
            );
          });
      }
      toast('Leave added!', 'success');
    }
  } catch (err) {
    toast(err.message || 'Failed to save leave.', 'error');
    _unlockOp('saveLeave', _btn);
    return;
  }
  _unlockOp('saveLeave', _btn);
  closeModal('leave-modal');
  if (state.view === 'calendar') renderCalendarDebounced();
  else renderLeaveCalendar();
}

async function deleteLeaveById(leaveId) {
  if (!leaveId) {
    toast('No leave ID provided.', 'error');
    return;
  }
  if (!confirm('Delete this leave record?')) return;
  try {
    await API.del(`/leaves/${leaveId}`);
    cache.leaves = cache.leaves.filter((l) => l.id !== leaveId);
    state.editingLeaveId = null;
    toast('Leave removed.', 'info');
    closeModal('leave-modal');
    if (state.view === 'calendar') renderCalendarDebounced();
    else renderLeaveCalendar();
  } catch (err) {
    toast(err.message || 'Failed to delete leave.', 'error');
  }
}

/* Remove a single day from a multi-day leave without a confirmation dialog.
   Trims the start/end, or splits the record if it's a middle day. */
async function removeLeaveDay(leaveId, dayStr) {
  const leave = getLeaves().find((l) => l.id === leaveId);
  if (!leave) {
    toast('Leave record not found.', 'error');
    return;
  }
  const sd = leave.startDate || (leave.start ? leave.start.slice(0, 10) : '');
  const ed = leave.endDate || (leave.end ? leave.end.slice(0, 10) : '');

  if (!confirm('Remove ' + dayStr + ' from this leave?')) return;

  const prevDay = (d) => {
    const dt = new Date(d + 'T12:00:00');
    dt.setDate(dt.getDate() - 1);
    return dt.toISOString().slice(0, 10);
  };
  const nextDay = (d) => {
    const dt = new Date(d + 'T12:00:00');
    dt.setDate(dt.getDate() + 1);
    return dt.toISOString().slice(0, 10);
  };

  try {
    if (sd === dayStr && ed === dayStr) {
      // Single-day leave — delete entirely
      await API.del(`/leaves/${leaveId}`);
      cache.leaves = cache.leaves.filter((l) => l.id !== leaveId);
    } else if (sd === dayStr) {
      // Remove first day — shift start forward
      const newStart = nextDay(dayStr);
      const updated = await API.put(`/leaves/${leaveId}`, {
        startDate: newStart,
        endDate: ed,
      });
      const idx = cache.leaves.findIndex((l) => l.id === leaveId);
      if (idx !== -1)
        cache.leaves[idx] = updated || {
          ...cache.leaves[idx],
          startDate: newStart,
        };
    } else if (ed === dayStr) {
      // Remove last day — shift end backward
      const newEnd = prevDay(dayStr);
      const updated = await API.put(`/leaves/${leaveId}`, {
        startDate: sd,
        endDate: newEnd,
      });
      const idx = cache.leaves.findIndex((l) => l.id === leaveId);
      if (idx !== -1)
        cache.leaves[idx] = updated || {
          ...cache.leaves[idx],
          endDate: newEnd,
        };
    } else {
      // Middle day — split: shorten original, create new for remainder
      const endPart1 = prevDay(dayStr);
      const startPart2 = nextDay(dayStr);
      const updated = await API.put(`/leaves/${leaveId}`, {
        startDate: sd,
        endDate: endPart1,
      });
      const idx = cache.leaves.findIndex((l) => l.id === leaveId);
      if (idx !== -1)
        cache.leaves[idx] = updated || {
          ...cache.leaves[idx],
          endDate: endPart1,
        };
      const newLeave = await API.post('/leaves', {
        userId: leave.userId,
        type: leave.type,
        startDate: startPart2,
        endDate: ed,
        reason: leave.reason || '',
      });
      if (newLeave) cache.leaves.push(newLeave);
    }
    toast('Day removed from leave.', 'info');
    if (state.view === 'calendar') renderCalendarDebounced();
    else renderLeaveCalendar();
  } catch (err) {
    toast(err.message || 'Failed to remove day.', 'error');
  }
}

function formatDateShort(date) {
  return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
}

/* ============================================================
   LEAVE/TASK CONFLICT CHECKING
   ============================================================ */
function checkLeaveConflict(userId, start, end) {
  const leaves = getLeaves().filter((l) => l.userId === userId);
  const taskStart = new Date(start);
  const taskEnd = new Date(end);

  return leaves.filter((leave) => {
    // Support both old ISO and new date-string format
    var leaveStart = leave.startDate
      ? new Date(leave.startDate + 'T00:00:00')
      : new Date(leave.start);
    var leaveEnd = leave.endDate
      ? new Date(leave.endDate + 'T23:59:59')
      : new Date(leave.end);
    return taskStart < leaveEnd && taskEnd > leaveStart;
  });
}

function notifyManagerOfLeaveConflict(task, conflictingLeaves) {
  const managers = getUsers().filter(
    (u) => u.role === 'manager' || u.role === 'admin',
  );
  const user = getUsers().find((u) => u.id === task.userId);

  const leaveTypes = conflictingLeaves
    .map((l) => {
      const types = { lieu: 'Lieu Day', loa: 'Leave of Absence', awol: 'AWOL' };
      return types[l.type];
    })
    .join(', ');

  managers.forEach((manager) => {
    if (manager.id === state.currentUser.id) return; // Don't notify self

    pushNotification(
      manager.id,
      '🚫 Task Conflicts with Leave',
      `Task "${task.title}" for ${user?.name} overlaps with ${leaveTypes}. Task cannot be created during this period.`,
      task.id,
      {
        type: 'leave-conflict',
        leaveIds: conflictingLeaves.map((l) => l.id),
        taskId: task.id,
      },
    );
  });
}

/* ============================================================
   TIMELINE RENDERING (Monday.com Style)
   ============================================================ */
function renderTimeline(userId) {
  const container = document.getElementById('timeline-container');
  const grid = document.getElementById('timeline-grid');
  const header = document.getElementById('timeline-header');
  const body = document.getElementById('timeline-body');

  let tasks = getTasks().filter(
    (t) => t.userId === userId && !t.cancelled && !t.done,
  );

  // Sort by priority then start date
  tasks.sort((a, b) => {
    if (a.priority !== b.priority) return a.priority - b.priority;
    return new Date(a.start) - new Date(b.start);
  });

  // Calculate date range
  const now = new Date();
  let startDate, endDate;

  if (tasks.length === 0) {
    startDate = new Date(now);
    endDate = new Date(now);
    endDate.setDate(endDate.getDate() + 7);
  } else {
    const dates = tasks.flatMap((t) => [
      new Date(t.start),
      new Date(t.deadline),
    ]);
    startDate = new Date(Math.min(...dates));
    endDate = new Date(Math.max(...dates));
    // Add buffer
    startDate.setDate(startDate.getDate() - 2);
    endDate.setDate(endDate.getDate() + 3);
  }

  // Generate header cells based on scale
  header.innerHTML = '';
  body.innerHTML = '';

  const scale = state.timelineScale;
  let cellWidth, cellCount;
  let current = new Date(startDate);

  if (scale === 'day') {
    cellWidth = 60;
    const diffTime = endDate - startDate;
    cellCount = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
    for (let i = 0; i < cellCount; i++) {
      const cell = document.createElement('div');
      cell.className = 'timeline-header-cell';
      cell.style.width = cellWidth + 'px';
      const d = new Date(startDate);
      d.setDate(d.getDate() + i);
      const isToday = d.toDateString() === now.toDateString();
      if (isToday) cell.classList.add('today');
      cell.textContent = d.toLocaleDateString('en-US', {
        weekday: 'short',
        day: 'numeric',
      });
      header.appendChild(cell);
    }
  } else if (scale === 'week') {
    cellWidth = 120;
    const diffTime = endDate - startDate;
    cellCount = Math.ceil(diffTime / (1000 * 60 * 60 * 24 * 7)) + 1;
    for (let i = 0; i < cellCount; i++) {
      const cell = document.createElement('div');
      cell.className = 'timeline-header-cell';
      cell.style.width = cellWidth + 'px';
      const d = new Date(startDate);
      d.setDate(d.getDate() + i * 7);
      const weekEnd = new Date(d);
      weekEnd.setDate(weekEnd.getDate() + 6);
      cell.textContent = `${d.getMonth() + 1}/${d.getDate()} - ${weekEnd.getMonth() + 1}/${weekEnd.getDate()}`;
      header.appendChild(cell);
    }
  } else if (scale === 'month') {
    cellWidth = 150;
    let months = [];
    let d = new Date(startDate);
    while (d <= endDate) {
      months.push(new Date(d));
      d.setMonth(d.getMonth() + 1);
    }
    cellCount = months.length;
    months.forEach((m) => {
      const cell = document.createElement('div');
      cell.className = 'timeline-header-cell';
      cell.style.width = cellWidth + 'px';
      cell.textContent = m.toLocaleDateString('en-US', {
        month: 'short',
        year: 'numeric',
      });
      header.appendChild(cell);
    });
  }

  // Add grid lines
  const gridLines = document.createElement('div');
  gridLines.className = 'timeline-grid-lines';
  for (let i = 0; i < cellCount; i++) {
    const line = document.createElement('div');
    line.className = 'timeline-grid-line';
    line.style.width =
      (scale === 'day' ? 60 : scale === 'week' ? 120 : 150) + 'px';
    const d = new Date(startDate);
    if (scale === 'day') d.setDate(d.getDate() + i);
    else if (scale === 'week') d.setDate(d.getDate() + i * 7);
    else d.setMonth(d.getMonth() + i);
    if (d.getDay() === 0 || d.getDay() === 6) line.classList.add('weekend');
    gridLines.appendChild(line);
  }
  body.appendChild(gridLines);

  // Add today line
  if (now >= startDate && now <= endDate) {
    const todayLine = document.createElement('div');
    todayLine.className = 'timeline-today-line';
    const totalMs = endDate - startDate;
    const elapsed = now - startDate;
    const pct = (elapsed / totalMs) * 100;
    todayLine.style.left = `calc(200px + ${pct}%)`;
    body.appendChild(todayLine);
  }

  // Render task rows
  if (tasks.length === 0) {
    const empty = document.createElement('div');
    empty.style.cssText =
      'text-align:center;padding:60px;color:var(--text3);font-size:13px;';
    empty.innerHTML =
      '📭 No active tasks to display<br><span style="font-size:11px;opacity:0.7;">Add tasks to see them on the timeline</span>';
    body.appendChild(empty);
  } else {
    // ── Hierarchy-aware rendering ─────────────────────────────
    // Build a set of child task IDs so we can skip them in the main loop
    // and render them nested beneath their parent row instead.
    const childIds = new Set(
      tasks.filter((t) => !!t.parentId).map((t) => t.id),
    );

    function _buildTaskRow(task, isChild, totalMs) {
      const row = document.createElement('div');
      row.className = 'timeline-row';
      if (isChild)
        row.style.cssText =
          'background:rgba(245,158,11,0.04);border-left:3px solid var(--amber);margin-left:24px;';

      const label = document.createElement('div');
      label.className = 'timeline-row-label';
      if (isChild) {
        label.style.cssText = 'padding-left:10px;';
        label.innerHTML = `<span style="display:flex;align-items:center;overflow:hidden;min-width:0;gap:4px;">
          <span style="font-size:9px;color:var(--amber);flex-shrink:0;">↳</span>
          <span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${escHtml(task.title)}</span>
        </span>
        <span style="margin-left:auto;font-size:10px;color:var(--text3);">P${task.priority}</span>`;
      } else {
        label.innerHTML = `<span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${escHtml(task.title)}</span>
          <span style="margin-left:auto;font-size:10px;color:var(--text3);">P${task.priority}</span>`;
      }
      row.appendChild(label);

      const content = document.createElement('div');
      content.className = 'timeline-row-content';

      const taskStart = new Date(task.start);
      const taskEnd = new Date(task.deadline);
      const startOffset = ((taskStart - startDate) / totalMs) * 100;
      const duration = ((taskEnd - taskStart) / totalMs) * 100;

      const bar = document.createElement('div');
      bar.className = `timeline-bar p${task.priority}`;
      bar.style.left = `${startOffset}%`;
      bar.style.width = `${Math.max(duration, 2)}%`;
      bar.innerHTML = `<span class="bar-title">${escHtml(task.title)}</span>
        <span class="bar-time">${formatTimeShort(taskStart)} - ${formatTimeShort(taskEnd)}</span>`;

      const overlaps = checkTaskOverlap(task, tasks);
      if (overlaps.length > 0) {
        const hasCritical = overlaps.some((o) => o.priority <= task.priority);
        bar.classList.add(hasCritical ? 'overlap-critical' : 'overlap-warning');
      }

      bar.onclick = (e) => {
        e.stopPropagation();
        openEditTask(task.id);
      };
      content.appendChild(bar);
      row.appendChild(content);
      return row;
    }

    const totalMs = endDate - startDate;

    tasks.forEach((task) => {
      // Skip children — they will be rendered nested under their parent
      if (childIds.has(task.id)) return;

      // Render the parent (or standalone) row
      body.appendChild(_buildTaskRow(task, false, totalMs));

      // If this task has children, render them immediately after, indented
      const children = tasks.filter((t) => t.parentId === task.id);
      children.forEach((child) => {
        body.appendChild(_buildTaskRow(child, true, totalMs));
      });
    });
  }

  // Update view toggle buttons
  document
    .getElementById('view-board-btn')
    .classList.toggle('active', state.currentViewMode === 'board');
  document
    .getElementById('view-timeline-btn')
    .classList.toggle('active', state.currentViewMode === 'timeline');
}

function formatTimeShort(date) {
  return date.toLocaleTimeString('en-US', {
    hour: '2-digit',
    minute: '2-digit',
  });
}

/* ============================================================
   OVERLAP DETECTION & RESOLUTION
   ============================================================ */
/* ============================================================
   WORK-HOUR AWARE SCHEDULING
   ============================================================ */

// Returns next working datetime at or after `from` for a given userId
function nextWorkStart(from, userId) {
  const wh = getWorkHours();
  const d = new Date(from);
  for (let i = 0; i < 60; i++) {
    // max 60 days look-ahead
    const dow = d.getDay();
    const isWknd = dow === 0 || dow === 6;
    const dayStr = d.toISOString().slice(0, 10);
    const userWds = getUserCustomWorkdays(userId);
    const isWorking = !isWknd || userWds.includes(dayStr);
    if (isWorking) {
      if (
        d.getHours() < wh.start ||
        (d.getHours() === wh.start && d.getMinutes() === 0)
      ) {
        d.setHours(wh.start, 0, 0, 0);
        return new Date(d);
      }
      if (d.getHours() < wh.end) return new Date(d); // mid-day: ok
    }
    // Advance to next day work start
    d.setDate(d.getDate() + 1);
    d.setHours(wh.start, 0, 0, 0);
  }
  return new Date(d);
}

// Calculate task end given start + durationHours, respecting work hours and user weekends
function calculateScheduledEnd(startISO, durationHours, userId) {
  const wh = getWorkHours();
  const workDayHours = wh.end - wh.start;
  let remaining = durationHours;
  const cursor = new Date(startISO);

  // Snap cursor to work start if before it
  if (cursor.getHours() < wh.start) cursor.setHours(wh.start, 0, 0, 0);

  for (let i = 0; i < 365 && remaining > 0; i++) {
    const dow = cursor.getDay();
    const isWknd = dow === 0 || dow === 6;
    const dayStr = cursor.toISOString().slice(0, 10);
    const userWds = getUserCustomWorkdays(userId);
    if (isWknd && !userWds.includes(dayStr)) {
      cursor.setDate(cursor.getDate() + 1);
      cursor.setHours(wh.start, 0, 0, 0);
      continue;
    }
    const hoursLeftToday =
      wh.end - (cursor.getHours() + cursor.getMinutes() / 60);
    if (hoursLeftToday <= 0) {
      cursor.setDate(cursor.getDate() + 1);
      cursor.setHours(wh.start, 0, 0, 0);
      continue;
    }
    if (remaining <= hoursLeftToday) {
      cursor.setTime(cursor.getTime() + remaining * 3600000);
      remaining = 0;
    } else {
      remaining -= hoursLeftToday;
      cursor.setDate(cursor.getDate() + 1);
      cursor.setHours(wh.start, 0, 0, 0);
    }
  }
  return new Date(cursor);
}

// Check overlaps between a task and existing tasks
function checkTaskOverlap(task, allTasks) {
  const taskStart = new Date(task.start);
  const taskEnd = new Date(task.deadline);
  return allTasks.filter((t) => {
    if (t.id === task.id) return false;
    return taskStart < new Date(t.deadline) && taskEnd > new Date(t.start);
  });
}

// New overlap resolver with priority stacking rules:
// - High/Urgent (P1/P2) can overlap with Mid/Low (P3-P5) — high goes on top
// - Two High/Urgent tasks cannot share the same time — first created has priority
// - Mid/Low tasks are always pushed after any High/Urgent that blocks them
function resolveOverlapsNew(newTask, existingTasks, userId) {
  const wh = getWorkHours();
  const priority = parseInt(newTask.priority);
  const durationMs = new Date(newTask.deadline) - new Date(newTask.start);
  const durationHours = durationMs / 3600000;

  // Determine which tasks actually block this one
  const blockers = existingTasks
    .filter((t) => {
      const op = parseInt(t.priority);
      const np = priority;
      // High/urgent (P1/P2) blocks mid/low (P3-P5)
      if (op <= 2 && np >= 3) return true;
      // Two high/urgent tasks block each other (first one wins)
      if (op <= 2 && np <= 2) return true;
      // Mid/low (P3-P5) also blocks each other — no same-time scheduling
      if (op >= 3 && np >= 3) return true;
      // Mid/low does NOT block high/urgent
      return false;
    })
    .sort((a, b) => new Date(a.start) - new Date(b.start));

  let candidateStart = new Date(newTask.start);
  candidateStart = nextWorkStart(candidateStart, userId || newTask.userId);

  let moved = false;
  let overlappingWith = [];

  for (let iter = 0; iter < 200; iter++) {
    const candidateEnd = calculateScheduledEnd(
      candidateStart.toISOString(),
      durationHours,
      userId || newTask.userId,
    );
    const blocking = blockers.filter((t) => {
      return (
        candidateStart < new Date(t.deadline) &&
        candidateEnd > new Date(t.start)
      );
    });
    if (blocking.length === 0) break;
    const latestEnd = blocking.reduce((m, t) => {
      const e = new Date(t.deadline);
      return e > m ? e : m;
    }, new Date(0));
    overlappingWith = blocking;
    moved = true;
    candidateStart = nextWorkStart(
      new Date(latestEnd.getTime() + 60000),
      userId || newTask.userId,
    );
  }

  const candidateEnd = calculateScheduledEnd(
    candidateStart.toISOString(),
    durationHours,
    userId || newTask.userId,
  );

  if (!moved) {
    // Only P1/P2 can visually overlap P3-P5 (high priority tasks show on top)
    // P3-P5 tasks are always moved to avoid overlap (handled by blockers above)
    const warns = checkTaskOverlap(
      {
        ...newTask,
        start: candidateStart.toISOString(),
        deadline: candidateEnd.toISOString(),
      },
      existingTasks,
    ).filter((t) => parseInt(t.priority) > priority && priority <= 2); // only warn if new task is high and overlaps low
    return {
      resolved: true,
      moved: false,
      overlaps: warns,
      newStart: candidateStart.toISOString(),
      newDeadline: candidateEnd.toISOString(),
    };
  }

  return {
    resolved: true,
    moved: true,
    overlaps: overlappingWith,
    newStart: candidateStart.toISOString(),
    newDeadline: candidateEnd.toISOString(),
  };
}

// Live overlap check while filling in the task modal
function onTaskFieldChange() {
  const startVal = document.getElementById('f-start')?.value;
  const durVal = getDurationHours();
  const priority = parseInt(
    document.getElementById('f-priority')?.value || '3',
  );
  const previewEl = document.getElementById('f-deadline-preview');
  const alertEl = document.getElementById('overlap-alert');
  const alertText = document.getElementById('overlap-alert-text');
  const leaveAlertEl = document.getElementById('leave-blocked-alert');
  const leaveAlertText = document.getElementById('leave-blocked-alert-text');

  if (!startVal || !durVal || durVal <= 0) {
    if (previewEl) previewEl.textContent = '';
    return;
  }

  const targetUserId = state.editingTaskId
    ? getTasks().find((t) => t.id === state.editingTaskId)?.userId ||
    state.currentUser.id
    : state.view === 'user-tasks'
      ? state.targetUserId
      : state.currentUser.id;

  // Preview: simple wall-clock end (start + duration). calculateScheduledEnd is for work-hours scheduling only.
  const wallClockEnd = new Date(
    new Date(startVal).getTime() + durVal * 3600000,
  );
  if (previewEl)
    previewEl.textContent =
      '⏱ Ends: ' +
      wallClockEnd.toLocaleString('en-US', {
        weekday: 'short',
        month: 'short',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
      });

  const scheduledEnd = calculateScheduledEnd(
    new Date(startVal).toISOString(),
    durVal,
    targetUserId,
  );

  // Weekend restriction check
  const startDate = new Date(startVal);
  const dow = startDate.getDay();
  const isWeekendStart = dow === 0 || dow === 6;
  if (isWeekendStart && priority >= 3) {
    if (alertEl) {
      alertEl.className = 'overlap-alert critical';
      if (alertText)
        alertText.textContent =
          '⚠️ Only P1 (Critical) and P2 (High) tasks can be scheduled on weekends. Task will be rescheduled to next Monday.';
      alertEl.classList.remove('hidden');
    }
    return;
  }

  // Overlap check
  const existingTasks = getTasks().filter(
    (t) =>
      t.userId === targetUserId &&
      !t.cancelled &&
      !t.done &&
      (state.editingTaskId ? t.id !== state.editingTaskId : true),
  );
  const proposed = {
    start: new Date(startVal).toISOString(),
    deadline: scheduledEnd.toISOString(),
    priority,
    id: state.editingTaskId || '__new__',
  };
  const overlaps = checkTaskOverlap(proposed, existingTasks);

  if (overlaps.length > 0) {
    const blockers = overlaps.filter(
      (o) =>
        parseInt(o.priority) <= priority &&
        (priority >= 3 || parseInt(o.priority) <= priority),
    );
    const warns = overlaps.filter((o) => parseInt(o.priority) > priority);
    // High/urgent (P1/P2) can overlap mid/low but same priority cannot
    const realBlockers = blockers.filter((o) => {
      const op = parseInt(o.priority);
      const np = priority;
      // Block if: new task is mid/low (>=3) AND existing is high/urgent (<=2)
      // OR both are high/urgent (<=2) — they cannot share same time
      if (op <= 2 && np >= 3) return true; // existing is high, new is mid/low → blocked
      if (op <= 2 && np <= 2) return true; // both high/urgent → blocked
      return false;
    });
    if (realBlockers.length > 0) {
      alertEl.className = 'overlap-alert critical';
      alertText.textContent = `⚠️ Blocked by higher-priority task(s): ${realBlockers.map((o) => `"${o.title}" (P${o.priority})`).join(', ')}. Task will be auto-rescheduled after them.`;
      alertEl.classList.remove('hidden');
    } else if (warns.length > 0) {
      alertEl.className = 'overlap-alert';
      alertText.textContent = `ℹ️ Overlaps with lower-priority task(s): ${warns.map((o) => `"${o.title}" (P${o.priority})`).join(', ')}. They can share time slot (lower priority will show below).`;
      alertEl.classList.remove('hidden');
    } else {
      alertEl.classList.add('hidden');
    }
  } else {
    alertEl.classList.add('hidden');
  }

  // Leave conflict check
  const leaveConflicts = checkLeaveConflict(
    targetUserId,
    new Date(startVal).toISOString(),
    scheduledEnd.toISOString(),
  );
  if (leaveConflicts.length > 0) {
    const types = { lieu: 'Lieu Day', loa: 'Leave of Absence', awol: 'AWOL' };
    leaveAlertEl.classList.remove('hidden');
    leaveAlertText.textContent = `Leave conflict: ${leaveConflicts.map((l) => types[l.type]).join(', ')}`;
  } else {
    leaveAlertEl.classList.add('hidden');
  }
}

function notifyManagerOfOverlap(task, overlaps, wasMoved) {
  const managers = getUsers().filter(
    (u) => u.role === 'manager' || u.role === 'admin',
  );
  const user = getUsers().find((u) => u.id === task.userId);

  const overlapNames = overlaps.map((o) => `"${o.title}"`).join(', ');
  const severity = overlaps.some((o) => o.priority <= task.priority)
    ? 'critical'
    : 'warning';

  managers.forEach((manager) => {
    pushNotification(
      manager.id,
      `⚠️ Task Overlap ${severity === 'critical' ? 'Resolved' : 'Detected'}`,
      `${user?.name || 'A user'}'s task "${task.title}" overlaps with ${overlapNames}.${wasMoved ? ' Task was auto-rescheduled.' : ''}`,
      task.id,
      {
        type: 'overlap',
        severity,
        taskId: task.id,
        overlapIds: overlaps.map((o) => o.id),
      },
    );
  });
}

/* ============================================================
   TASK RENDERING
   ============================================================ */
/**
 * Render the task board for a specific user.
 * Builds task cards grouped by status (active → done → cancelled),
 * sorted by priority and deadline. Supports search filtering.
 * @param {string} userId - The user ID whose tasks to display
 */
function renderTasks(userId) {
  const board = document.getElementById('task-board');
  board.innerHTML = '';
  let tasks = getTasks().filter((t) => t.userId === userId);
  const isElevated =
    state.currentUser.role === 'admin' || state.currentUser.role === 'manager';

  // Sort: active by priority → cancelled → done
  tasks.sort((a, b) => {
    const statusOrder = (t) => (t.done ? 2 : t.cancelled ? 1 : 0);
    if (statusOrder(a) !== statusOrder(b))
      return statusOrder(a) - statusOrder(b);
    if (!a.done && !a.cancelled && !b.done && !b.cancelled) {
      if (a.priority !== b.priority) return a.priority - b.priority;
      return (
        new Date(a.start || a.createdAt) - new Date(b.start || b.createdAt)
      );
    }
    return 0;
  });

  if (tasks.length === 0) {
    board.innerHTML = `<div class="board-empty"><div class="board-empty-icon">📋</div><div class="board-empty-text">No tasks yet</div><div class="board-empty-sub">Your task board is clear.</div></div>`;
    const fab = EL.addTaskFab || document.getElementById('add-task-fab');
    if (fab) fab.classList.remove('show-fab');
    return;
  }

  const fab = EL.addTaskFab || document.getElementById('add-task-fab');
  if (fab && window.innerWidth <= 768) fab.classList.add('show-fab');
  const now = new Date();
  let lastPriority = null;
  let shownCancelledHeader = false;
  let shownDoneHeader = false;

  // For team tasks, only render the first copy per group; collect members for that group
  const allTasks = getTasks();
  const renderedGroups = new Set();

  tasks.forEach((task) => {
    // ── Team task deduplication ──
    // If this task belongs to a team group, check if we already rendered that group
    const meta = task.isTeamTask ? _getGroupMeta(task.description) : null;
    const groupId = meta ? meta.groupId : null;
    if (groupId) {
      if (renderedGroups.has(groupId)) return; // already rendered, skip
      renderedGroups.add(groupId);
    }

    const deadline = new Date(task.deadline);
    const hoursLeft = (deadline - now) / 1000 / 60 / 60;
    const isActive = !task.done && !task.cancelled;
    const isDue = isActive && hoursLeft <= 24 && hoursLeft > 0;
    const isOverdue = isActive && hoursLeft <= 0;
    const isWarning = isActive && hoursLeft <= 72 && hoursLeft > 24;

    const checkedCount = _visibleDesc(task.description).filter(
      (d) => d.checked,
    ).length;
    const totalItems = _visibleDesc(task.description).length;
    const pct =
      totalItems > 0 ? Math.round((checkedCount / totalItems) * 100) : 100;
    const allChecked = totalItems === 0 || checkedCount === totalItems;

    // Section dividers
    if (isActive && lastPriority !== task.priority) {
      const pLabels = {
        1: 'Critical',
        2: 'High',
        3: 'Medium',
        4: 'Low',
        5: 'Minimal',
      };
      board.insertAdjacentHTML(
        'beforeend',
        `
        <div class="board-section">
          <span class="section-label">Priority ${task.priority} — ${pLabels[task.priority]}</span>
          <span class="section-line"></span>
        </div>
      `,
      );
      lastPriority = task.priority;
    }
    if (task.cancelled && !shownCancelledHeader) {
      board.insertAdjacentHTML(
        'beforeend',
        `
        <div class="board-section">
          <span class="section-label" style="color:var(--danger)">🚫 Cancelled Tasks</span>
          <span class="section-line" style="background:rgba(239,68,68,0.2)"></span>
        </div>
      `,
      );
      shownCancelledHeader = true;
    }
    if (task.done && !shownDoneHeader) {
      board.insertAdjacentHTML(
        'beforeend',
        `
        <div class="board-section">
          <span class="section-label" style="color:var(--success)">✅ Completed Tasks</span>
          <span class="section-line" style="background:rgba(34,197,94,0.2)"></span>
        </div>
      `,
      );
      shownDoneHeader = true;
    }

    let flickerClass = '';
    if (isOverdue || isDue) flickerClass = 'flicker-critical';
    else if (isWarning) flickerClass = 'flicker-warning';

    let deadlineStr = formatDeadline(deadline);
    let deadlineClass = isOverdue
      ? 'urgent'
      : isDue
        ? 'urgent'
        : isWarning
          ? 'warning'
          : '';

    // Determine footer content
    let footerContent = '';
    if (task.cancelled) {
      footerContent = `<span class="cancelled-tag">🚫 Cancelled</span>
        ${isElevated ? `<button class="btn-secondary" style="font-size:10px;padding:5px 10px;" onclick="reopenTask(event,'${task.id}')">🔄 Reopen</button>` : ''}`;
    } else if (task.done) {
      footerContent = `<button class="done-btn is-done" disabled>✓ Completed</button>
        ${isElevated ? `<button class="btn-secondary" style="font-size:10px;padding:5px 10px;margin-left:6px;" onclick="reopenTask(event,'${task.id}')">🔄 Reopen</button>` : ''}
        <button class="btn-danger" style="font-size:10px;padding:5px 10px;margin-left:6px;" onclick="deleteCompletedTask(event,'${task.id}')" data-tip="Delete this completed task">🗑 Delete</button>`;
    } else {
      footerContent = `<button class="done-btn" onclick="markDone(event,'${task.id}')" ${!allChecked ? 'disabled data-tip="Tick all checklist items above before marking done"' : 'data-tip="Mark this task as completed"'}>
        ${!allChecked ? `⬜ ${checkedCount}/${totalItems} Done` : '✓ Mark Done'}
      </button>`;
    }

    const checklistDisabled = task.cancelled || task.done;

    // ── Build team member list for grouped team tasks ──
    let memberListHtml = '';
    if (groupId) {
      const siblings = allTasks.filter((t) => {
        if (!t.isTeamTask) return false;
        const m = _getGroupMeta(t.description);
        return m && m.groupId === groupId;
      });
      const allMembers = siblings.concat([task]);
      const users = getUsers();
      memberListHtml = `
        <div style="margin-top:10px;padding-top:10px;border-top:1px solid var(--border);">
          <div style="font-size:10px;font-weight:700;color:var(--text3);letter-spacing:0.08em;text-transform:uppercase;margin-bottom:6px;">👥 Assigned to (${allMembers.length})</div>
          <div style="display:flex;flex-direction:column;gap:4px;">
            ${allMembers
          .map((m) => {
            const u = users.find((u) => u.id === m.userId);
            const mDone = m.done;
            const mCancelled = m.cancelled;
            const statusIcon = mDone ? '✅' : mCancelled ? '🚫' : '🔄';
            const statusColor = mDone
              ? 'var(--success)'
              : mCancelled
                ? 'var(--danger)'
                : 'var(--text3)';
            return `<div style="display:flex;align-items:center;gap:8px;font-size:11px;padding:3px 0;">
                <span style="color:${statusColor};font-size:12px;">${statusIcon}</span>
                <span style="font-weight:600;color:var(--text2);">${escHtml(u ? u.name : 'Unknown')}</span>
                <span style="color:var(--text3);font-size:10px;">${mDone ? 'Done' : mCancelled ? 'Cancelled' : 'In progress'}</span>
              </div>`;
          })
          .join('')}
          </div>
        </div>`;
    }

    const card = document.createElement('div');
    card.className = `task-card p${task.priority} ${flickerClass} ${task.done ? 'done' : ''} ${task.cancelled ? 'cancelled' : ''}`;
    card.dataset.id = task.id;
    card.innerHTML = `
      <div class="card-priority-bar"></div>
      <div class="card-body">
        <div class="card-header">
          <div class="card-title">${escHtml(task.title)}</div>
          <div class="priority-badge">${task.priority}</div>
        </div>
        ${task.teamName ? `<div style="font-size:10px;color:#F59E0B;margin-bottom:4px;letter-spacing:0.04em;">🏷️ ${escHtml(task.teamName)}</div>` : ''}
        ${task.isMultiPersonnel ? `<div style="font-size:10px;color:var(--info);margin-bottom:4px;">👥 Multi-personnel task</div>` : ''}
        ${task.cancelled && task.cancelReason ? `<div class="cancel-reason">Reason: ${escHtml(task.cancelReason)}</div>` : ''}
        <div class="card-meta">
          <span class="meta-item"><span class="icon">👤</span>${escHtml(task.requestor)}</span>
          <span class="meta-item meta-deadline ${deadlineClass}"><span class="icon">⏰</span>${deadlineStr}${isOverdue ? ' (OVERDUE)' : isDue ? ' (TODAY)' : ''}</span>
          ${task.cancelledAt ? `<span class="meta-item"><span class="icon">🚫</span>Cancelled ${formatTime(new Date(task.cancelledAt))}</span>` : ''}
          ${task.doneAt ? `<span class="meta-item"><span class="icon">✅</span>Done ${formatTime(new Date(task.doneAt))}</span>` : ''}
        </div>
        ${totalItems > 0
        ? `
        <div class="checklist-preview">
          <div class="checklist-bar"><div class="checklist-fill" style="width:${pct}%"></div></div>
          <span>${checkedCount}/${totalItems}</span>
        </div>`
        : ''
      }
      </div>
      <div class="card-expand-section" id="expand-${task.id}">
        ${totalItems > 0
        ? `<div>
          <div class="expand-label">Checklist${_getGroupMeta(task.description) ? ' <span style="font-size:10px;color:var(--amber);font-weight:600;">· shared</span>' : ''}</div>
          <div class="checklist-items">
            ${_visibleDesc(task.description)
          .map(
            (d, idx) => `
              <div class="checklist-item ${d.checked ? 'checked' : ''}" id="cli-${task.id}-${idx}">
                <input type="checkbox" id="chk-${task.id}-${idx}" ${d.checked ? 'checked' : ''} ${checklistDisabled ? 'disabled' : ''}
                  onchange="toggleCheck('${task.id}', ${idx}, this.checked)">
                <label for="chk-${task.id}-${idx}">${escHtml(d.text)}</label>
              </div>
            `,
          )
          .join('')}
          </div>
        </div>`
        : `<div style="font-size:12px;color:var(--text3);">${task.cancelled ? 'Task was cancelled.' : 'No checklist items — task can be marked done directly.'}</div>`
      }
        ${memberListHtml}
      </div>
      <div class="card-footer">
        ${footerContent}
        <span class="card-right-hint">click to expand</span>
      </div>
    `;

    card
      .querySelector('.card-body')
      .addEventListener('click', () => toggleExpand(task.id));
    card.addEventListener('click', (e) => {
      if (
        e.target.closest('.done-btn') ||
        e.target.closest('.btn-secondary') ||
        e.target.closest('.btn-danger') ||
        e.target.closest('input[type="checkbox"]') ||
        e.target.closest('label')
      )
        return;
      if (e.target.closest('.card-body')) return;
      toggleExpand(task.id);
    });
    card.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      showContextMenu(e, task.id);
    });

    let _lpTimer = null;
    card.addEventListener(
      'touchstart',
      (e) => {
        _lpTimer = setTimeout(() => {
          _lpTimer = null;
          const touch = e.touches[0];
          const fakeEvent = {
            clientX: touch.clientX,
            clientY: touch.clientY,
            preventDefault: () => { },
          };
          showContextMenu(fakeEvent, task.id);
        }, 600);
      },
      { passive: true },
    );
    card.addEventListener('touchend', () => {
      clearTimeout(_lpTimer);
      _lpTimer = null;
    });
    card.addEventListener('touchmove', () => {
      clearTimeout(_lpTimer);
      _lpTimer = null;
    });

    board.appendChild(card);
  });
}

/* ============================================================
   BOARD SEARCH / FILTER
   ============================================================ */
let _boardSearchVal = '';
let _boardFilter = 'all';

function onBoardSearch(val) {
  _boardSearchVal = (val || '').trim().toLowerCase();
  const clearBtn = EL.boardSearch
    ? EL.boardSearch.parentElement.querySelector('#board-search-clear')
    : $('board-search-clear');
  if (clearBtn) clearBtn.classList.toggle('hidden', !_boardSearchVal);
  applyBoardFilter();
}

function clearBoardSearch() {
  _boardSearchVal = '';
  if (EL.boardSearch) EL.boardSearch.value = '';
  const clearBtn = $('board-search-clear');
  if (clearBtn) clearBtn.classList.add('hidden');
  applyBoardFilter();
}

function setBoardFilter(filter) {
  _boardFilter = filter;
  document.querySelectorAll('.board-filter-pill').forEach((pill) => {
    pill.classList.toggle('active', pill.dataset.filter === filter);
  });
  applyBoardFilter();
}

function applyBoardFilter() {
  const cards = document.querySelectorAll('.task-card');
  const sections = document.querySelectorAll('.board-section');
  const now = new Date();
  let anyVisible = false;
  cards.forEach((card) => {
    const title = (
      card.querySelector('.card-title')?.textContent || ''
    ).toLowerCase();
    const meta = (
      card.querySelector('.card-meta')?.textContent || ''
    ).toLowerCase();
    const matchSearch =
      !_boardSearchVal ||
      title.includes(_boardSearchVal) ||
      meta.includes(_boardSearchVal);
    let matchFilter = true;
    if (_boardFilter === 'active')
      matchFilter =
        !card.classList.contains('done') &&
        !card.classList.contains('cancelled');
    else if (_boardFilter === 'overdue') {
      const deadlineEl = card.querySelector('.meta-deadline');
      matchFilter =
        deadlineEl &&
        (deadlineEl.classList.contains('urgent') ||
          deadlineEl.textContent.includes('OVERDUE'));
    } else if (_boardFilter === 'done')
      matchFilter = card.classList.contains('done');
    const show = matchSearch && matchFilter;
    card.style.display = show ? '' : 'none';
    if (show) anyVisible = true;
  });
  // Show/hide section headers based on visible cards
  sections.forEach((sec) => {
    sec.style.display = '';
  });
  // Show a "no results" hint if nothing visible
  let noResults = document.getElementById('board-no-results');
  if (!anyVisible) {
    if (!noResults) {
      noResults = document.createElement('div');
      noResults.id = 'board-no-results';
      noResults.style.cssText =
        'grid-column:1/-1;text-align:center;padding:40px;color:var(--text3);font-size:13px;';
      noResults.textContent = 'No tasks match your filter.';
      document.getElementById('task-board').appendChild(noResults);
    }
    noResults.style.display = '';
  } else if (noResults) {
    noResults.style.display = 'none';
  }
}

/* ============================================================
   OFFLINE INDICATOR
   ============================================================ */

function toggleExpand(taskId) {
  const el = document.getElementById('expand-' + taskId);
  if (!el) return;
  el.classList.toggle('open');
  const card = document.querySelector(`.task-card[data-id="${taskId}"]`);
  if (el.classList.contains('open')) {
    card?.querySelector('.card-right-hint')?.textContent &&
      (card.querySelector('.card-right-hint').textContent =
        'click to collapse');
  } else {
    card?.querySelector('.card-right-hint')?.textContent &&
      (card.querySelector('.card-right-hint').textContent = 'click to expand');
  }
}

async function toggleCheck(taskId, idx, checked) {
  const tasks = getTasks();
  const task = tasks.find((t) => t.id === taskId);
  if (!task) return;
  // idx here is the index in the *visible* description (excludes __meta items)
  const visible = _visibleDesc(task.description);
  if (!visible[idx]) return;
  visible[idx].checked = checked;
  // Rebuild full description preserving meta
  const meta = _getGroupMeta(task.description);
  task.description = _buildDesc(visible, meta);

  // Update DOM for this task
  const item = document.getElementById(`cli-${taskId}-${idx}`);
  if (item) item.classList.toggle('checked', checked);
  const allChecked = visible.every((d) => d.checked);
  const doneBtn = document.querySelector(
    `.task-card[data-id="${taskId}"] .done-btn`,
  );
  if (doneBtn && !task.done) {
    doneBtn.disabled = !allChecked;
    const total = visible.length;
    const checkedCount = visible.filter((d) => d.checked).length;
    doneBtn.textContent = allChecked
      ? '✓ Mark Done'
      : `⬜ ${checkedCount}/${total} Done`;
  }
  const total = visible.length;
  const checkedCount = visible.filter((d) => d.checked).length;
  const pct = total > 0 ? Math.round((checkedCount / total) * 100) : 100;
  const fill = document.querySelector(
    `.task-card[data-id="${taskId}"] .checklist-fill`,
  );
  if (fill) fill.style.width = pct + '%';
  const preview = document.querySelector(
    `.task-card[data-id="${taskId}"] .checklist-preview span`,
  );
  if (preview) preview.textContent = `${checkedCount}/${total}`;

  // Persist this task
  API.put(`/tasks/${taskId}`, { description: task.description }).catch(
    () => { },
  );
  // Propagate change to all other tasks in the same group
  await _syncGroupCheck(task, task.description);
}

async function markDone(e, taskId) {
  e.stopPropagation();
  const tasks = getTasks();
  const task = tasks.find((t) => t.id === taskId);
  if (!task) return;
  const allChecked =
    _visibleDesc(task.description).length === 0 ||
    _visibleDesc(task.description).every((d) => d.checked);
  if (!allChecked) {
    toast('Complete all checklist items first!', 'error');
    return;
  }
  const doneAt = new Date().toISOString();
  task.done = true;
  task.doneAt = doneAt;
  try {
    await API.put(`/tasks/${taskId}`, { done: true, doneAt });
    addLog({
      taskId: task.id,
      taskTitle: task.title,
      action: 'done',
      actorName: state.currentUser.name,
      userId: task.userId,
    });
    // Sync to all sibling team task copies
    const siblings = _getTeamSiblings(task);
    for (const s of siblings) {
      s.done = true;
      s.doneAt = doneAt;
      API.put(`/tasks/${s.id}`, { done: true, doneAt }).catch(() => { });
    }
    const count = siblings.length + 1;
    toast(
      count > 1
        ? `Task marked as done! ✓ (${count} members updated)`
        : 'Task marked as done! ✓',
      'success',
    );
    // Notify managers and admins that this task was completed
    if (state.currentUser.role === 'user') {
      var managers = getUsers().filter(
        (u) => u.role === 'admin' || u.role === 'manager',
      );
      managers.forEach((m) => {
        pushNotification(
          m.id,
          `✅ Task Completed`,
          `"${task.title}" has been marked as done by ${state.currentUser.name}.`,
          task.id,
          { type: 'task', action: 'completed' },
        );
      });
    }
    if (state.view === 'calendar') {
      renderCalendarDebounced();
      return;
    }
    const userId =
      state.view === 'user-tasks' ? state.targetUserId : state.currentUser.id;
    if (state.currentViewMode === 'timeline') renderTimeline(userId);
    else renderTasks(userId);
  } catch (err) {
    task.done = false;
    task.doneAt = null;
    toast('Failed to update task.', 'error');
  }
}

async function confirmDelete() {
  if (!state.deleteTaskId) return;
  if (!_lockOp('confirmDelete')) return;
  const tasks = getTasks();
  const task = tasks.find((t) => t.id === state.deleteTaskId);
  const id = state.deleteTaskId;
  state.deleteTaskId = null;
  closeModal('confirm-modal');
  // Collect siblings before deletion
  const siblings = task ? _getTeamSiblings(task) : [];
  try {
    await API.del(`/tasks/${id}`);
    if (task)
      addLog({
        taskId: task.id,
        taskTitle: task.title,
        action: 'deleted',
        actorName: state.currentUser.name,
        userId: task.userId,
      });
    if (task) {
      const deletedUser = getUsers().find((u) => u.id === task.userId);
      if (deletedUser)
        sendTaskEmail(deletedUser, task, 'cancelled').catch(() => { });
    }
    cache.tasks = cache.tasks.filter((t) => t.id !== id);
    // Delete all sibling copies
    for (const s of siblings) {
      cache.tasks = cache.tasks.filter((t) => t.id !== s.id);
      API.del(`/tasks/${s.id}`).catch(() => { });
    }
    const count = siblings.length + 1;
    toast(
      count > 1 ? `Task deleted (${count} members updated).` : 'Task deleted.',
      'info',
    );
  } catch (err) {
    toast('Failed to delete task.', 'error');
    _unlockOp('confirmDelete');
    return;
  }
  _unlockOp('confirmDelete');
  if (state.view === 'calendar') {
    renderCalendarDebounced();
    return;
  }
  const userId =
    state.view === 'user-tasks' ? state.targetUserId : state.currentUser.id;
  if (state.currentViewMode === 'timeline') renderTimeline(userId);
  else renderTasks(userId);
}

async function confirmCancel() {
  if (!state.cancelTaskId) return;
  if (!_lockOp('confirmCancel')) return;
  const reason = document.getElementById('cancel-reason-input').value.trim();
  const tasks = getTasks();
  const task = tasks.find((t) => t.id === state.cancelTaskId);
  state.cancelTaskId = null;
  closeModal('cancel-modal');
  if (!task) return;
  const siblings = _getTeamSiblings(task);
  try {
    const updated = await API.put(`/tasks/${task.id}`, {
      cancelled: true,
      cancelReason: reason || '',
    });
    if (updated) {
      Object.assign(task, updated);
    } else {
      task.cancelled = true;
      task.cancelReason = reason || '';
    }
    addLog({
      taskId: task.id,
      taskTitle: task.title,
      action: 'cancelled',
      actorName: state.currentUser.name,
      userId: task.userId,
    });
    if (task.userId !== state.currentUser.id) {
      pushNotification(
        task.userId,
        '🚫 Task Cancelled',
        `"${task.title}" was cancelled by ${state.currentUser.name}${reason ? ': ' + reason : '.'}`,
        task.id,
      );
    }
    const cancelledUser = getUsers().find((u) => u.id === task.userId);
    if (cancelledUser)
      sendTaskEmail(cancelledUser, task, 'cancelled').catch(() => { });
    // Cancel all sibling copies
    for (const s of siblings) {
      s.cancelled = true;
      s.cancelReason = reason || '';
      API.put(`/tasks/${s.id}`, {
        cancelled: true,
        cancelReason: reason || '',
      }).catch(() => { });
    }
    const count = siblings.length + 1;
    toast(
      count > 1
        ? `Task cancelled (${count} members updated).`
        : 'Task cancelled.',
      'info',
    );
  } catch (err) {
    toast('Failed to cancel task.', 'error');
    _unlockOp('confirmCancel');
    return;
  }
  _unlockOp('confirmCancel');
  if (state.view === 'calendar') {
    renderCalendarDebounced();
    return;
  }
  const userId =
    state.view === 'user-tasks' ? state.targetUserId : state.currentUser.id;
  if (state.currentViewMode === 'timeline') renderTimeline(userId);
  else renderTasks(userId);
}

async function reopenTask(e, taskId) {
  if (e) e.stopPropagation();
  if (!_lockOp('reopen_' + taskId, null, 'Reopening task…')) return;
  const tasks = getTasks();
  const task = tasks.find((t) => t.id === taskId);
  if (!task) {
    _unlockOp('reopen_' + taskId);
    return;
  }
  const siblings = _getTeamSiblings(task);
  try {
    const updated = await API.put(`/tasks/${taskId}`, {
      cancelled: false,
      cancelReason: '',
      done: false,
      doneAt: null,
    });
    if (updated) {
      Object.assign(task, updated);
    } else {
      task.cancelled = false;
      task.cancelReason = '';
      task.done = false;
      task.doneAt = null;
    }
    addLog({
      taskId: task.id,
      taskTitle: task.title,
      action: 'reopened',
      actorName: state.currentUser.name,
      userId: task.userId,
    });
    if (task.userId !== state.currentUser.id) {
      pushNotification(
        task.userId,
        '🔄 Task Reopened',
        `"${task.title}" has been reopened by ${state.currentUser.name}.`,
        task.id,
      );
    }
    const reopenedUser = getUsers().find((u) => u.id === task.userId);
    if (reopenedUser)
      sendTaskEmail(reopenedUser, task, 'reopened').catch(() => { });
    for (const s of siblings) {
      s.cancelled = false;
      s.cancelReason = '';
      s.done = false;
      s.doneAt = null;
      API.put(`/tasks/${s.id}`, {
        cancelled: false,
        cancelReason: '',
        done: false,
        doneAt: null,
      }).catch(() => { });
    }
    const count = siblings.length + 1;
    toast(
      count > 1
        ? `Task reopened! ✓ (${count} members updated)`
        : 'Task reopened! ✓',
      'success',
    );
  } catch (err) {
    toast('Failed to reopen task.', 'error');
    _unlockOp('reopen_' + taskId);
    return;
  }
  _unlockOp('reopen_' + taskId);
  if (state.view === 'calendar') {
    renderCalendarDebounced();
    return;
  }
  const userId =
    state.view === 'user-tasks' ? state.targetUserId : state.currentUser.id;
  if (state.currentViewMode === 'timeline') renderTimeline(userId);
  else renderTasks(userId);
}

/* ============================================================
   CONTEXT MENU
   ============================================================ */
function showContextMenu(e, taskId) {
  state.contextTaskId = taskId;
  const task = getTasks().find((t) => t.id === taskId);
  const isElevated =
    state.currentUser.role === 'admin' || state.currentUser.role === 'manager';
  const menu = document.getElementById('context-menu');

  // Show/hide items based on role and task state
  const editBtn = document.getElementById('ctx-edit');
  const viewCalBtn = document.getElementById('ctx-view-cal');
  const cancelBtn = document.getElementById('ctx-cancel');
  const reopenBtn = document.getElementById('ctx-reopen');
  const deleteBtn = document.getElementById('ctx-delete');

  // Edit: elevated roles always; user can edit their own active tasks (only if they created it)
  const isOwner = task.userId === state.currentUser.id;
  const isCreator = !task.createdBy || task.createdBy === state.currentUser.id;
  const userCanEdit = isOwner && isCreator;
  editBtn.style.display =
    !task.cancelled && !task.done && (isElevated || userCanEdit) ? '' : 'none';

  // View in Calendar: hide when already in calendar view
  if (viewCalBtn)
    viewCalBtn.style.display = state.view === 'calendar' ? 'none' : '';

  // Cancel: only elevated, only on active tasks
  cancelBtn.style.display =
    isElevated && !task.cancelled && !task.done ? '' : 'none';

  // Reopen: only elevated, on cancelled OR completed tasks
  reopenBtn.style.display =
    isElevated && (task.cancelled || task.done) ? '' : 'none';

  // Delete: elevated always; regular user only when task is done
  deleteBtn.style.display = isElevated || (isOwner && task.done) ? '' : 'none';

  menu.classList.remove('hidden');
  let x = e.clientX,
    y = e.clientY;
  menu.style.left = x + 'px';
  menu.style.top = y + 'px';
  requestAnimationFrame(() => {
    const rect = menu.getBoundingClientRect();
    if (rect.right > window.innerWidth) menu.style.left = x - rect.width + 'px';
    if (rect.bottom > window.innerHeight)
      menu.style.top = y - rect.height + 'px';
  });
}
function hideContextMenu() {
  document.getElementById('context-menu').classList.add('hidden');
}
document.addEventListener('click', hideContextMenu);
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    hideContextMenu();
    closeAllModals();
  }
});

function ctxEdit() {
  hideContextMenu();
  openEditTask(state.contextTaskId);
}
function ctxViewInCalendar() {
  hideContextMenu();
  const task = getTasks().find((t) => t.id === state.contextTaskId);
  if (!task) return;
  const taskStart = new Date(task.start || task.createdAt);
  state._calendarJumpDate = taskStart;
  state._calendarFocusTaskId = task.id;

  const isElevated =
    state.currentUser.role === 'admin' || state.currentUser.role === 'manager';
  const isViewingWorker =
    isElevated && state.view === 'user-tasks' && state.targetUserId;

  if (isViewingWorker) {
    // Manager viewing a worker's tasks → go to team calendar and highlight that worker's row
    state.calendarMode = 'team';
    state._calendarHighlightUserId = task.userId;
  } else {
    // Worker or manager viewing their own tasks → go to personal calendar
    state.calendarMode = 'personal';
    state._calendarHighlightUserId = null;
  }

  showView('calendar');
  toast('📅 Jumped to task in calendar.', 'info');
}
function ctxCancel() {
  hideContextMenu();
  const task = getTasks().find((t) => t.id === state.contextTaskId);
  if (!task) return;
  const isElevated =
    state.currentUser.role === 'admin' || state.currentUser.role === 'manager';
  if (!isElevated) {
    toast('Only managers can cancel tasks.', 'error');
    return;
  }
  if (task.cancelled) {
    toast('Task is already cancelled.', 'info');
    return;
  }
  if (task.done) {
    toast('Cannot cancel a completed task.', 'error');
    return;
  }
  state.cancelTaskId = task.id;
  document.getElementById('cancel-task-name').textContent =
    '"' + task.title + '"';
  document.getElementById('cancel-reason-input').value = '';
  openModal('cancel-modal');
}
function ctxDelete() {
  hideContextMenu();
  const task = getTasks().find((t) => t.id === state.contextTaskId);
  if (!task) return;
  const isElevated =
    state.currentUser.role === 'admin' || state.currentUser.role === 'manager';
  if (!isElevated) {
    // Regular users: only allow delete on completed tasks
    if (!task.done) {
      toast('You can only delete completed tasks.', 'error');
      return;
    }
    deleteCompletedTask(null, task.id);
    return;
  }
  state.deleteTaskId = task.id;
  document.getElementById('confirm-task-name').textContent = `"${task.title}"`;
  openModal('confirm-modal');
}
function ctxReopen() {
  hideContextMenu();
  const isElevated =
    state.currentUser.role === 'admin' || state.currentUser.role === 'manager';
  if (!isElevated) {
    toast('Only managers can reopen tasks.', 'error');
    return;
  }
  reopenTask(null, state.contextTaskId);
}

/* ============================================================
   ADD / EDIT TASK MODAL
   ============================================================ */
/* Schedule mode: 'duration' or 'enddate' */
var _taskScheduleMode = 'duration';
function setTaskScheduleMode(mode) {
  _taskScheduleMode = mode;
  const durBtn = document.getElementById('f-mode-duration');
  const endBtn = document.getElementById('f-mode-enddate');
  const durGrp = document.getElementById('f-duration-group');
  const endGrp = document.getElementById('f-enddate-group');
  const activeStyle =
    'flex:1;padding:7px;border-radius:6px;font-size:11px;font-weight:700;border:1.5px solid var(--amber);background:var(--amber);color:var(--bg);cursor:pointer;transition:all 0.15s;';
  const inactiveStyle =
    'flex:1;padding:7px;border-radius:6px;font-size:11px;font-weight:700;border:1.5px solid var(--border);background:var(--bg3);color:var(--text2);cursor:pointer;transition:all 0.15s;';
  if (mode === 'duration') {
    durBtn.style.cssText = activeStyle;
    endBtn.style.cssText = inactiveStyle;
    durGrp.style.display = '';
    endGrp.style.display = 'none';
  } else {
    endBtn.style.cssText = activeStyle;
    durBtn.style.cssText = inactiveStyle;
    endGrp.style.display = '';
    durGrp.style.display = 'none';
  }
  onTaskFieldChange();
}
function getDurationHours() {
  if (_taskScheduleMode === 'enddate') {
    const startVal = document.getElementById('f-start').value;
    const endVal = document.getElementById('f-enddate').value;
    if (!startVal || !endVal) return 0;
    const diff = (new Date(endVal) - new Date(startVal)) / 3600000;
    return Math.max(0.25, diff);
  }
  const raw = parseFloat(document.getElementById('f-duration').value);
  const unit = document.getElementById('f-duration-unit').value;
  if (!raw || raw <= 0) return 0;
  return unit === 'days' ? raw * 8 : raw; // 1 day = 8 work hours
}

/**
 * Open the Add Task modal with empty fields.
 * Resets all form inputs and description items for a new task.
 */
function openAddTask() {
  state.editingTaskId = null;
  state.descItems = [];
  _taskScheduleMode = 'duration';
  document.getElementById('task-modal-title').textContent = 'Add New Task';
  document.getElementById('f-title').value = '';
  document.getElementById('f-requestor').value = state.currentUser?.name || '';
  document.getElementById('f-priority').value = '3';
  const now = new Date();
  document.getElementById('f-start').value = toLocalISO(now);
  document.getElementById('f-duration').value = '1';
  document.getElementById('f-duration-unit').value = 'hours';
  document.getElementById('f-enddate').value = '';
  setTaskScheduleMode('duration');
  document.getElementById('f-deadline-preview').textContent = '';
  document.getElementById('f-desc-input').value = '';
  document.getElementById('overlap-alert').classList.add('hidden');
  document.getElementById('leave-blocked-alert').classList.add('hidden');
  renderDescItems();
  onTaskFieldChange();
  // If called from calendar, stay on calendar after save
  if (state.view === 'calendar') state._returnToCalendar = true;
  openModal('task-modal');
}

function openEditTask(taskId) {
  const task = getTasks().find((t) => t.id === taskId);
  if (!task) return;
  state.editingTaskId = taskId;
  state.descItems = _visibleDesc(task.description).map((d) => ({
    text: d.text,
    checked: d.checked,
  }));
  document.getElementById('task-modal-title').textContent = 'Edit Task';
  document.getElementById('f-title').value = task.title;
  document.getElementById('f-requestor').value = task.requestor;
  document.getElementById('f-priority').value = task.priority;
  const startDate = task.start
    ? new Date(task.start)
    : new Date(task.createdAt);
  document.getElementById('f-start').value = toLocalISO(startDate);
  const deadline = new Date(task.deadline);
  const durationHours = Math.round(((deadline - startDate) / 36e5) * 4) / 4;
  document.getElementById('f-duration').value = Math.max(0.25, durationHours);
  document.getElementById('f-duration-unit').value = 'hours';
  document.getElementById('f-enddate').value = '';
  _taskScheduleMode = 'duration';
  setTaskScheduleMode('duration');
  renderDescItems();
  onTaskFieldChange();
  openModal('task-modal');
}

function toLocalISO(date) {
  const offset = date.getTimezoneOffset() * 60000;
  return new Date(date.getTime() - offset).toISOString().slice(0, 16);
}
// Local date string — avoids UTC offset shifting the date
function toDateStr(date) {
  return (
    date.getFullYear() +
    '-' +
    String(date.getMonth() + 1).padStart(2, '0') +
    '-' +
    String(date.getDate()).padStart(2, '0')
  );
}

function addDescItem() {
  const input = document.getElementById('f-desc-input');
  const text = input.value.trim();
  if (!text) return;
  state.descItems.push({ text, checked: false });
  input.value = '';
  renderDescItems();
  input.focus();
}

function removeDescItem(idx) {
  state.descItems.splice(idx, 1);
  renderDescItems();
}

function renderDescItems() {
  const container = document.getElementById('f-desc-items');
  container.innerHTML = state.descItems
    .map(
      (item, i) => `
    <div class="desc-item">
      <span>${escHtml(item.text)}</span>
      <button onclick="removeDescItem(${i})">✕</button>
    </div>
  `,
    )
    .join('');
}

/**
 * Validate and save a new or edited task.
 * Handles priority overlap resolution, deadline calculation,
 * multi-personnel syncing, notification dispatch, and email alerts.
 * @returns {Promise<void>}
 */
async function saveTask() {
  const _btn = document.querySelector(
    '#task-modal .btn-primary[onclick="saveTask()"]',
  );
  if (!_lockOp('saveTask', _btn, 'Saving…')) return;
  try {
    const title = document.getElementById('f-title').value.trim();
    const requestor = document.getElementById('f-requestor').value.trim();
    const priority = parseInt(document.getElementById('f-priority').value);
    const startVal = document.getElementById('f-start').value;
    const durVal = getDurationHours();

    if (!title || !requestor || !startVal || !durVal || durVal <= 0) {
      toast('Please fill in all required fields.', 'error');
      _unlockOp('saveTask', _btn);
      return;
    }

    const targetUserId =
      state.view === 'user-tasks' ? state.targetUserId : state.currentUser.id;
    const userId = state.editingTaskId
      ? getTasks().find((t) => t.id === state.editingTaskId)?.userId ||
      targetUserId
      : targetUserId;

    // Weekend restriction: only P1/P2 allowed on weekends
    let scheduledStart = nextWorkStart(new Date(startVal), userId);
    const startDow = scheduledStart.getDay();
    const isWeekend = startDow === 0 || startDow === 6;
    if (isWeekend && priority >= 3) {
      // Reschedule to next Monday
      while (scheduledStart.getDay() === 0 || scheduledStart.getDay() === 6) {
        scheduledStart.setDate(scheduledStart.getDate() + 1);
      }
      const wh = getWorkHours();
      scheduledStart.setHours(wh.start, 0, 0, 0);
      toast(
        'Task rescheduled to next weekday (weekends are for P1/P2 only).',
        'warning',
      );
    }

    // Calculate deadline respecting work hours / user weekends
    const scheduledEnd = calculateScheduledEnd(
      scheduledStart.toISOString(),
      durVal,
      userId,
    );

    const leaveConflicts = checkLeaveConflict(
      userId,
      scheduledStart.toISOString(),
      scheduledEnd.toISOString(),
    );
    if (leaveConflicts.length > 0) {
      const types = { lieu: 'Lieu Day', loa: 'Leave of Absence', awol: 'AWOL' };
      toast(
        `Cannot create task: overlaps with ${leaveConflicts.map((l) => types[l.type]).join(', ')}`,
        'error',
      );
      notifyManagerOfLeaveConflict(
        {
          title,
          requestor,
          priority,
          start: scheduledStart.toISOString(),
          deadline: scheduledEnd.toISOString(),
          userId,
        },
        leaveConflicts,
      );
      _unlockOp('saveTask', _btn);
      return;
    }

    const tasks = getTasks();
    const description = state.descItems.map((d) => ({
      text: d.text,
      checked: false,
    }));

    if (state.editingTaskId) {
      const task = tasks.find((t) => t.id === state.editingTaskId);
      if (!task) {
        _unlockOp('saveTask', _btn);
        return;
      }
      const otherTasks = tasks.filter(
        (t) =>
          t.id !== state.editingTaskId &&
          t.userId === task.userId &&
          !t.cancelled &&
          !t.done,
      );
      const overlapCheck = resolveOverlapsNew(
        {
          ...task,
          title,
          requestor,
          priority,
          start: scheduledStart.toISOString(),
          deadline: scheduledEnd.toISOString(),
        },
        otherTasks,
        task.userId,
      );
      const finalStart = overlapCheck.newStart || scheduledStart.toISOString();
      const finalDeadline =
        overlapCheck.newDeadline || scheduledEnd.toISOString();
      const oldDesc = task.description || [];
      const oldMeta = _getGroupMeta(oldDesc);
      const finalDesc = _buildDesc(
        state.descItems.map((d) => {
          const ex = _visibleDesc(oldDesc).find((od) => od.text === d.text);
          return { text: d.text, checked: ex ? ex.checked : d.checked };
        }),
        oldMeta,
      );
      if (overlapCheck.moved) {
        toast('Task auto-rescheduled due to overlap.', 'warning');
        notifyManagerOfOverlap(task, overlapCheck.overlaps, true);
      }
      const updated = await API.put(`/tasks/${state.editingTaskId}`, {
        title,
        requestor,
        priority,
        start: finalStart,
        deadline: finalDeadline,
        description: finalDesc,
      });
      // Merge returned fields into cached task; fall back to sent values if Supabase returns null
      if (updated) {
        Object.assign(task, updated);
      } else {
        Object.assign(task, {
          title,
          requestor,
          priority,
          start: finalStart,
          deadline: finalDeadline,
          description: finalDesc,
        });
      }
      addLog({
        taskId: task.id,
        taskTitle: task.title,
        action: 'edited',
        actorName: state.currentUser.name,
        userId: task.userId,
      });
      const editedUser = getUsers().find((u) => u.id === task.userId);
      if (editedUser)
        sendTaskEmail(editedUser, task, 'updated').catch(() => { });
      // Propagate edits to all sibling team task copies (keep each sibling's own meta)
      const siblings = _getTeamSiblings(task);
      for (const s of siblings) {
        const sMeta = _getGroupMeta(s.description);
        const sDesc = _buildDesc(_visibleDesc(finalDesc), sMeta);
        Object.assign(s, { title, requestor, priority, description: sDesc });
        API.put(`/tasks/${s.id}`, {
          title,
          requestor,
          priority,
          description: sDesc,
        }).catch(() => { });
      }
      toast(
        siblings.length > 0
          ? `Task updated! ✓ (${siblings.length + 1} members synced)`
          : 'Task updated!',
        'success',
      );
    } else {
      const existingTasks = tasks.filter(
        (t) => t.userId === userId && !t.cancelled && !t.done,
      );
      const taskData = {
        title,
        requestor,
        priority,
        start: scheduledStart.toISOString(),
        deadline: scheduledEnd.toISOString(),
        description,
        createdBy: state.currentUser.id,
      };
      const overlapCheck = resolveOverlapsNew(taskData, existingTasks, userId);
      if (overlapCheck.moved) {
        taskData.start = overlapCheck.newStart;
        taskData.deadline = overlapCheck.newDeadline;
        toast(
          'Task auto-rescheduled — placed after higher-priority tasks.',
          'warning',
        );
      }
      const newTask = await API.post('/tasks', { ...taskData, userId });
      if (!newTask) {
        // Offline — task queued; create optimistic local entry
        const optimisticTask = {
          id: 'offline_' + uid(),
          ...taskData,
          userId,
          done: false,
          cancelled: false,
          createdAt: new Date().toISOString(),
        };
        if (!cache.tasks) cache.tasks = [];
        cache.tasks.push(optimisticTask);
        toast(
          '📵 Offline — task saved locally and will sync when reconnected.',
          'warning',
        );
        _unlockOp('saveTask', _btn);
        closeModal('task-modal');
        const renderUserId =
          state.view === 'user-tasks'
            ? state.targetUserId
            : state.currentUser.id;
        renderTasks(renderUserId);
        return;
      }
      if (!cache.tasks) cache.tasks = [];
      cache.tasks.push(newTask);
      if (
        (state.currentUser.role === 'admin' ||
          state.currentUser.role === 'manager') &&
        userId !== state.currentUser.id
      ) {
        var actorRole =
          state.currentUser.role === 'admin' ? 'Admin' : 'Manager';
        var assignedName =
          getUsers().find((u) => u.id === userId)?.name || 'User';
        // Notify the assignee
        pushNotification(
          userId,
          `📌 New Task Assigned`,
          `"${title}" has been added by ${actorRole}.`,
          newTask.id,
          { type: 'task', action: 'assigned' },
        );
        // Notify the setter (confirmation)
        pushNotification(
          state.currentUser.id,
          `✅ Task Assigned to ${assignedName}`,
          `"${title}" has been assigned successfully.`,
          newTask.id,
          { type: 'task', action: 'assigned-confirm', sendEmail: false },
        );
      }
      const assignedUser = getUsers().find((u) => u.id === userId);
      if (assignedUser)
        sendTaskEmail(assignedUser, newTask, 'assigned').catch(() => { });
      if (overlapCheck.overlaps.length > 0)
        notifyManagerOfOverlap(
          newTask,
          overlapCheck.overlaps,
          overlapCheck.moved,
        );
      addLog({
        taskId: newTask.id,
        taskTitle: newTask.title,
        action: 'added',
        actorName: state.currentUser.name,
        userId: newTask.userId,
      });
      toast('Task added! ✓', 'success');
    }
  } catch (err) {
    // Error handling
    toast(err.message || 'Failed to save task.', 'error');
    _unlockOp('saveTask', _btn);
    return;
  }
  _unlockOp('saveTask', _btn);
  closeModal('task-modal');
  // Stay on calendar if that's where we are (or if flagged to return)
  if (state.view === 'calendar' || state._returnToCalendar) {
    state._returnToCalendar = false;
    state.view = 'calendar';
    renderCalendar();
    return;
  }
  const renderUserId =
    state.view === 'user-tasks' ? state.targetUserId : state.currentUser.id;
  if (state.currentViewMode === 'timeline') renderTimeline(renderUserId);
  else renderTasks(renderUserId);
}

function viewUserTasks(userId) {
  state.targetUserId = userId;
  state.view = 'user-tasks';
  _navRecord('user-tasks:' + userId);
  showView('user-tasks');
}

async function deleteUser(userId) {
  if (!confirm('Delete this user and all their tasks?')) return;
  try {
    await API.del(`/users/${userId}`);
    cache.users = cache.users.filter((u) => u.id !== userId);
    cache.tasks = cache.tasks.filter((t) => t.userId !== userId);
    toast('User deleted.', 'info');
    renderUserList();
  } catch (err) {
    toast(err.message || 'Failed to delete user.', 'error');
  }
}

/* ============================================================
   REGISTER USER
   ============================================================ */
function openRegisterUser() {
  // Populate team dropdown
  const teamSel = document.getElementById('reg-team');
  if (teamSel) {
    const teams = getTeams();
    teamSel.innerHTML =
      '<option value="">No team</option>' +
      teams
        .map(
          (t) => `<option value="${escHtml(t.id)}">${escHtml(t.name)}</option>`,
        )
        .join('');
  }
  openModal('register-modal');
  setTimeout(() => document.getElementById('reg-name').focus(), 100);
}

async function registerUser() {
  const _btn = document.querySelector(
    '#register-modal .btn-primary[onclick="registerUser()"]',
  );
  if (!_lockOp('registerUser', _btn, 'Saving…')) return;
  const name = document.getElementById('reg-name').value.trim();
  const username = document
    .getElementById('reg-username')
    .value.trim()
    .toLowerCase()
    .replace(/\s/g, '');
  const password = document.getElementById('reg-password').value.trim();
  const role = document.getElementById('reg-role').value;
  const teamId = document.getElementById('reg-team')?.value || '';
  const email = document.getElementById('reg-email')?.value.trim() || '';
  const emailNotif =
    document.getElementById('reg-email-notif')?.checked !== false;
  if (!name || !username || !password) {
    toast('All fields are required.', 'error');
    return;
  }
  try {
    const newUser = await API.post('/users', {
      name,
      username,
      password,
      role,
      email,
      emailNotif,
    });
    cache.users.push(newUser);
    // Auto-add to selected team
    if (teamId) {
      const team = cache.teams
        ? cache.teams.find((t) => t.id === teamId)
        : null;
      if (team) {
        if (!team.memberIds) team.memberIds = [];
        if (!team.memberIds.includes(newUser.id)) {
          team.memberIds.push(newUser.id);
          await API.put('/teams/' + teamId, { memberIds: team.memberIds });
        }
      }
    }
    document.getElementById('reg-name').value = '';
    document.getElementById('reg-username').value = '';
    document.getElementById('reg-password').value = '';
    document.getElementById('reg-role').value = 'user';
    if (document.getElementById('reg-team'))
      document.getElementById('reg-team').value = '';
    closeModal('register-modal');
    const roleLabel =
      role === 'admin' ? 'Admin' : role === 'manager' ? 'Manager' : 'User';
    toast(`${roleLabel} "${name}" registered! ✓`, 'success');
    if (state.view === 'user-list') renderUserList();
    if (state.view === 'teams') renderTeamsView();
  } catch (err) {
    toast(err.message || 'Failed to register user.', 'error');
  } finally {
    _unlockOp('registerUser', _btn);
  }
}

/* ============================================================
   WHATSAPP PARSER (Claude API)
   ============================================================ */
function openWhatsApp() {
  state.waParsed = null;
  document.getElementById('wa-result').classList.add('hidden');
  document.getElementById('wa-use-btn').classList.add('hidden');
  document.getElementById('wa-loading').classList.add('hidden');
  document.getElementById('wa-text').value = '';
  openModal('wa-modal');
}

async function parseWhatsApp() {
  const text = document.getElementById('wa-text').value.trim();
  if (!text) {
    toast('Paste a message first.', 'error');
    return;
  }
  document.getElementById('wa-loading').classList.remove('hidden');
  document.getElementById('wa-result').classList.add('hidden');
  document.getElementById('wa-parse-btn').disabled = true;

  try {
    const prompt = `You are a task extraction assistant. Extract task information from this WhatsApp message and return ONLY a JSON object with these fields:
- title: string (task name, short and clear)
- requestor: string (who sent/requested the task, use "Unknown" if not clear)
- priority: number 1-5 (1=critical/urgent, 3=medium, 5=low)
- start: string in ISO format YYYY-MM-DDTHH:mm (use current time if not specified, today is ${new Date().toISOString().slice(0, 10)})
- deadline: string in ISO format YYYY-MM-DDTHH:mm (estimate from context, use today+3days if not specified)
- description: array of strings (checklist steps, max 5 items, extract from message or create logical steps)

Message: "${text}"

Return ONLY valid JSON, no explanation.`;

    const res = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 1000,
        messages: [{ role: 'user', content: prompt }],
      }),
    });
    const data = await res.json();
    const raw = data.content?.[0]?.text || '';
    const cleaned = raw.replace(/```json|```/g, '').trim();
    const parsed = JSON.parse(cleaned);
    state.waParsed = parsed;

    const pLabels = {
      1: 'Critical',
      2: 'High',
      3: 'Medium',
      4: 'Low',
      5: 'Minimal',
    };
    document.getElementById('wa-parsed-display').innerHTML = `
      <strong>Task:</strong> ${escHtml(parsed.title)}<br>
      <strong>Requestor:</strong> ${escHtml(parsed.requestor)}<br>
      <strong>Priority:</strong> ${parsed.priority} — ${pLabels[parsed.priority]}<br>
      <strong>Start:</strong> ${formatDeadline(new Date(parsed.start))}<br>
      <strong>Deadline:</strong> ${formatDeadline(new Date(parsed.deadline))}<br>
      <strong>Checklist:</strong><br>
      ${parsed.description.map((d, i) => `&nbsp;&nbsp;${i + 1}. ${escHtml(d)}`).join('<br>')}
    `;
    document.getElementById('wa-result').classList.remove('hidden');
    document.getElementById('wa-use-btn').classList.remove('hidden');
  } catch (err) {
    toast(
      "Couldn't read the message. Try adding more detail like a deadline or task name.",
      'error',
    );
  } finally {
    document.getElementById('wa-loading').classList.add('hidden');
    document.getElementById('wa-parse-btn').disabled = false;
  }
}

function useWaParsed() {
  if (!state.waParsed) return;
  const p = state.waParsed;
  closeModal('wa-modal');
  state.editingTaskId = null;
  state.descItems = (p.description || []).map((d) => ({
    text: d,
    checked: false,
  }));
  document.getElementById('task-modal-title').textContent =
    'Task from WhatsApp';
  document.getElementById('f-title').value = p.title || '';
  document.getElementById('f-requestor').value = p.requestor || '';
  document.getElementById('f-priority').value = p.priority || 3;
  try {
    document.getElementById('f-start').value = p.start.slice(0, 16);
    document.getElementById('f-enddate').value = p.deadline.slice(0, 16);
  } catch { }
  renderDescItems();
  openModal('task-modal');
}

/* ============================================================
   LOGS
   ============================================================ */
function addLog({ taskId, taskTitle, action, actorName, userId }) {
  const entry = {
    id: uid(),
    taskId,
    taskTitle,
    action,
    actorName,
    userId,
    timestamp: new Date().toISOString(),
  };
  if (!cache.logs) cache.logs = [];
  cache.logs.unshift(entry);
  if (cache.logs.length > 500) cache.logs.splice(500);
  // Fire-and-forget to API
  API.post('/logs', { taskId, taskTitle, action, actorName, userId }).catch(
    () => { },
  );
}

function openLogs() {
  const logs = getLogs();
  const list = document.getElementById('logs-list');
  const isAdmin = state.currentUser.role === 'admin';
  const filtered = isAdmin
    ? logs
    : logs.filter((l) => l.userId === state.currentUser.id);

  if (filtered.length === 0) {
    list.innerHTML = '<div class="logs-empty">📭 No activity yet</div>';
  } else {
    const icons = {
      added: '➕',
      done: '✅',
      edited: '✏️',
      deleted: '🗑️',
      cancelled: '🚫',
      reopened: '🔄',
    };
    const colors = {
      added: 'var(--amber)',
      done: 'var(--success)',
      edited: 'var(--p4)',
      deleted: 'var(--danger)',
      cancelled: 'var(--danger)',
      reopened: '#A78BFA',
    };
    list.innerHTML = filtered
      .map((l) => {
        const user = getUsers().find((u) => u.id === l.userId);
        return `
        <div class="log-entry">
          <div class="log-icon">${icons[l.action] || '📝'}</div>
          <div class="log-content">
            <div class="log-title" style="color:${colors[l.action]}">${l.action.toUpperCase()} — ${escHtml(l.taskTitle)}</div>
            <div class="log-meta">by ${escHtml(l.actorName)}${isAdmin && user ? ` (for ${escHtml(user.name)})` : ''}</div>
          </div>
          <div class="log-time">${formatTime(new Date(l.timestamp))}</div>
        </div>
      `;
      })
      .join('');
  }
  openModal('logs-modal');
}

/* ============================================================
   NOTIFICATIONS
   ============================================================ */
function pushNotification(userId, title, body, taskId, metadata = null) {
  const notif = {
    id: uid(),
    userId,
    title,
    body,
    taskId,
    read: false,
    timestamp: new Date().toISOString(),
    metadata,
  };
  if (!cache.notifications) cache.notifications = [];
  cache.notifications.unshift(notif);
  updateNotifBadge();
  // Persist to Supabase
  API.post('/notifications', { userId, title, body, taskId, metadata }).catch(
    () => { },
  );
  // Browser notification
  if ('serviceWorker' in navigator && navigator.serviceWorker.controller) {
    navigator.serviceWorker.controller.postMessage({
      type: 'SHOW_NOTIFICATION',
      title,
      body,
      tag: 'task-' + (taskId || 'notif'),
    });
  } else if (Notification.permission === 'granted') {
    new Notification(title, { body, tag: 'tf-' + (taskId || uid()) });
  }
  // Email notification for task assignments
  if (
    metadata?.sendEmail !== false &&
    (metadata?.type === 'task' || !metadata?.type)
  ) {
    const user = getUsers().find((u) => u.id === userId);
    if (user?.emailNotif !== false && user?.email) {
      const task = taskId ? getTasks().find((t) => t.id === taskId) : null;
      if (task)
        sendTaskEmail(user, task, metadata?.action || 'assigned').catch(
          () => { },
        );
    }
  }
}

function updateNotifBadge() {
  const notifs = getNotifs().filter(
    (n) => n.userId === state.currentUser?.id && !n.read,
  );
  const btn = document.getElementById('notif-btn');
  if (notifs.length > 0) {
    btn.classList.add('notification-dot');
    btn.setAttribute('data-count', notifs.length);
    btn.title =
      notifs.length + ' unread notification' + (notifs.length !== 1 ? 's' : '');
  } else {
    btn.classList.remove('notification-dot');
    btn.removeAttribute('data-count');
    btn.title = 'Notifications';
  }
  // Also update mobile bottom nav badge
  updateMobileNavNotifBadge();
}

/* ============================================================
   EMAIL NOTIFICATIONS via mailto / Supabase Edge Function stub
   ============================================================ */
// Professional email composer — creates a mailto: link as fallback
// In production, replace sendEmailViaAPI with a real Supabase Edge Function

async function sendLeaveEmail(user, leaveInfo) {
  if (!user.email) return;
  const { type, startDate, endDate, days, typeLabel, fmtDate, icon, reason } =
    leaveInfo;
  const companyName =
    cache.companies?.find?.((c) => c.id === _cid())?.name || 'Your Company';
  const subject = `[TaskFlow] ${typeLabel} Recorded — ${user.name}`;
  const body = composeLeaveEmail(user, leaveInfo, companyName);
  await sendEmailViaAPI(user.email, subject, body);
}

async function sendTaskEmail(user, task, action) {
  if (!user || user.emailNotif === false || !user.email) {
    return;
  }
  const companyName =
    cache.companies?.find?.((c) => c.id === _cid())?.name || 'Your Company';
  const subject = `[TaskFlow] Task ${action} — ${task.title}`;
  const body = composeTaskEmail(user, task, action, companyName);
  await sendEmailViaAPI(user.email, subject, body);
}

// Email via Supabase Edge Function (deploy separately) or fallback
async function sendEmailViaAPI(to, subject, htmlBody) {
  try {
    const result = await SB.insert('tf_email_queue', {
      company_id: _cid() || null,
      to_email: to,
      subject: subject,
      html_body: htmlBody,
      sent: false,
    });
  } catch (e) {
    console.error('[EMAIL] ❌ Failed to queue:', e.message, e);
  }
}

// Debug: call testEmailQueue() from browser console to test email insert

function composeLeaveEmail(user, info, companyName) {
  const startDate = info.startDate || info.start_date || info.start || '';
  const endDate = info.endDate || info.end_date || info.end || '';
  const days = Number(info.days) || 1;
  const typeLabel = info.typeLabel || info.type || 'Leave';
  const reason = info.reason || info.notes || '';
  const icon = info.icon || '📅';
  const type = info.type || 'lieu';

  const fmt = (d) =>
    d
      ? new Date(d + 'T12:00:00').toLocaleDateString('en-GB', {
        weekday: 'long',
        day: 'numeric',
        month: 'long',
        year: 'numeric',
      })
      : '—';
  const colors = { lieu: '#10B981', loa: '#3B82F6', awol: '#EF4444' };
  const color = colors[type] || '#F59E0B';
  return `
  <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#f8f9fa;padding:20px;">
    <div style="background:#08090D;padding:24px 32px;border-radius:12px 12px 0 0;">
      <div style="font-size:22px;font-weight:800;color:#F59E0B;letter-spacing:-0.02em;">TaskFlow</div>
      <div style="color:#9CA3AF;font-size:12px;margin-top:4px;">${companyName}</div>
    </div>
    <div style="background:#fff;padding:32px;border-radius:0 0 12px 12px;border:1px solid #e5e7eb;border-top:none;">
      <div style="display:inline-block;padding:6px 14px;background:${color}18;color:${color};border:1px solid ${color}33;border-radius:20px;font-size:12px;font-weight:700;margin-bottom:20px;">${icon} ${typeLabel.toUpperCase()}</div>
      <h2 style="margin:0 0 8px;font-size:20px;color:#111;">Leave Period Recorded</h2>
      <p style="color:#6B7280;font-size:14px;margin:0 0 24px;">Dear <strong>${user.name}</strong>, the following leave has been recorded on your account.</p>
      <table style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:24px;">
        <tr style="border-bottom:1px solid #f0f0f0;"><td style="padding:10px 0;color:#6B7280;width:40%;">Leave Type</td><td style="padding:10px 0;font-weight:600;color:#111;">${typeLabel}</td></tr>
        <tr style="border-bottom:1px solid #f0f0f0;"><td style="padding:10px 0;color:#6B7280;">From</td><td style="padding:10px 0;font-weight:600;color:#111;">${fmt(startDate)}</td></tr>
        <tr style="border-bottom:1px solid #f0f0f0;"><td style="padding:10px 0;color:#6B7280;">To</td><td style="padding:10px 0;font-weight:600;color:#111;">${fmt(endDate)}</td></tr>
        <tr style="border-bottom:1px solid #f0f0f0;"><td style="padding:10px 0;color:#6B7280;">Working Days</td><td style="padding:10px 0;font-weight:600;color:${color};">${days} day${days !== 1 ? 's' : ''}</td></tr>
        ${reason ? `<tr><td style="padding:10px 0;color:#6B7280;vertical-align:top;">Notes</td><td style="padding:10px 0;font-weight:600;color:#111;">${reason}</td></tr>` : ''}
      </table>
      ${type === 'awol' ? '<div style="padding:14px 16px;background:#FEF2F2;border:1px solid #FECACA;border-radius:8px;color:#DC2626;font-size:13px;margin-bottom:20px;"><strong>⚠️ Important:</strong> AWOL is recorded as an unauthorized absence and will impact your attendance record.</div>' : ''}
      <p style="font-size:12px;color:#9CA3AF;border-top:1px solid #f0f0f0;padding-top:16px;margin:0;">This is an automated notification from TaskFlow. Please contact your manager if you believe this is an error.</p>
    </div>
  </div>`;
}

function composeTaskEmail(user, task, action, companyName) {
  const title = task.title || task.name || '—';
  const requestor = task.requestor || task.requested_by || '—';
  const priority = Number(task.priority) || 3;
  const start = task.start || task.start_time || task.scheduled_start || null;
  const deadline = task.deadline || task.end_time || task.scheduled_end || null;
  const description = Array.isArray(task.description)
    ? task.description
    : task.checklist || task.items || [];

  const pLabels = [
    '',
    '🔴 Critical',
    '🟠 High',
    '🟡 Medium',
    '🔵 Low',
    '⚪ Minimal',
  ];
  const pColors = ['', '#FF4040', '#FF8C00', '#F5C518', '#38BDF8', '#64748B'];
  const p = priority;
  const fmt = (d) =>
    d
      ? new Date(d).toLocaleDateString('en-GB', {
        weekday: 'short',
        day: 'numeric',
        month: 'short',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
      })
      : '—';
  const actionLabels = {
    assigned: 'assigned to you',
    updated: 'updated',
    cancelled: 'cancelled',
    done: 'marked as complete',
    reopened: 'reopened',
  };
  const actionLabel = actionLabels[action] || action;
  return `
  <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#f8f9fa;padding:20px;">
    <div style="background:#08090D;padding:24px 32px;border-radius:12px 12px 0 0;">
      <div style="font-size:22px;font-weight:800;color:#F59E0B;">TaskFlow</div>
      <div style="color:#9CA3AF;font-size:12px;margin-top:4px;">${companyName}</div>
    </div>
    <div style="background:#fff;padding:32px;border-radius:0 0 12px 12px;border:1px solid #e5e7eb;border-top:none;">
      <div style="display:inline-block;padding:6px 14px;background:${pColors[p]}18;color:${pColors[p]};border:1px solid ${pColors[p]}33;border-radius:20px;font-size:12px;font-weight:700;margin-bottom:20px;">P${p} ${pLabels[p]}</div>
      <h2 style="margin:0 0 8px;font-size:20px;color:#111;">Task ${actionLabel.charAt(0).toUpperCase() + actionLabel.slice(1)}</h2>
      <p style="color:#6B7280;font-size:14px;margin:0 0 24px;">Dear <strong>${user.name}</strong>, a task has been <strong>${actionLabel}</strong>.</p>
      <div style="background:#f8f9fa;border:1px solid #e5e7eb;border-left:4px solid ${pColors[p]};border-radius:4px;padding:16px 20px;margin-bottom:24px;">
        <div style="font-size:18px;font-weight:700;color:#111;margin-bottom:8px;">${title}</div>
        <div style="font-size:13px;color:#6B7280;">Requested by: <strong>${requestor}</strong></div>
      </div>
      <table style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:24px;">
        <tr style="border-bottom:1px solid #f0f0f0;"><td style="padding:10px 0;color:#6B7280;width:40%;">Priority</td><td style="padding:10px 0;font-weight:700;color:${pColors[p]};">${pLabels[p]}</td></tr>
        <tr style="border-bottom:1px solid #f0f0f0;"><td style="padding:10px 0;color:#6B7280;">Start</td><td style="padding:10px 0;font-weight:600;color:#111;">${fmt(start)}</td></tr>
        <tr style="border-bottom:1px solid #f0f0f0;"><td style="padding:10px 0;color:#6B7280;">Deadline</td><td style="padding:10px 0;font-weight:600;color:#111;">${fmt(deadline)}</td></tr>
        ${description.length ? `<tr><td style="padding:10px 0;color:#6B7280;vertical-align:top;">Checklist</td><td style="padding:10px 0;">${description.map((d) => `<div style="margin-bottom:4px;font-size:13px;">${d.checked ? '✅' : '☐'} ${d.text}</div>`).join('')}</td></tr>` : ''}
      </table>
      <p style="font-size:12px;color:#9CA3AF;border-top:1px solid #f0f0f0;padding-top:16px;margin:0;">This is an automated notification from TaskFlow.</p>
    </div>
  </div>`;
}

/* ============================================================
   NOTIFICATION POLLING
   Polls for new notifications every 8 seconds.
   ============================================================ */
let _notifPollInterval = null;
let _lastNotifCount = 0;
let _initialLoadDone = false;

function startRealtimeNotifs() {
  clearInterval(_notifPollInterval);

  // Core poll function — checks for new notifications
  async function _pollNotifs() {
    if (!state.currentUser) return;
    try {
      const fresh = await API.get('/notifications');
      if (!fresh) return;
      cache.notifications = fresh;
      const unread = fresh.filter(
        (n) => n.userId === state.currentUser.id && !n.read,
      ).length;
      if (unread > _lastNotifCount && _lastNotifCount >= 0) {
        const newest = fresh.filter(
          (n) => n.userId === state.currentUser.id && !n.read,
        )[0];
        if (newest) {
          _showBrowserNotif(newest.title, newest.body, 'tf-' + newest.id);
        }
        const btn = document.getElementById('notif-btn');
        if (btn) {
          btn.style.animation = 'none';
          setTimeout(() => (btn.style.animation = ''), 10);
        }
      }
      _lastNotifCount = unread;
      updateNotifBadge();
    } catch { }
  }

  _notifPollInterval = setInterval(_pollNotifs, 8000);

  // When tab comes back from background, immediately check for notifications
  document.addEventListener('visibilitychange', function () {
    if (!document.hidden && state.currentUser) {
      _pollNotifs();
    }
  });
}

function stopRealtimeNotifs() {
  clearInterval(_notifPollInterval);
}

// Alias: mobile nav calls toggleNotifications(), desktop uses openNotifications()
function toggleNotifications() {
  openNotifications();
}

/**
 * Open the notifications modal showing all notifications
 * for the current user, grouped by read status.
 */
function openNotifications() {
  const notifs = getNotifs().filter((n) => n.userId === state.currentUser.id);
  const list = document.getElementById('notif-list');
  if (notifs.length === 0) {
    list.innerHTML = '<div class="notif-empty">🔔 No notifications</div>';
  } else {
    list.innerHTML = notifs
      .map((n) => {
        const isOverlap = n.metadata?.type === 'overlap';
        const isLeave = n.metadata?.type === 'leave';
        const isLeaveConflict = n.metadata?.type === 'leave-conflict';
        const severity = n.metadata?.severity || 'warning';
        return `
        <div class="notif-item ${n.read ? '' : 'unread'} ${isOverlap ? 'overlap-' + severity : ''} ${isLeave ? 'leave-blocked' : ''}" onclick="handleNotifClick('${n.id}')">
          <div class="notif-title">
            ${isOverlap ? '⚠️' : isLeave ? '📅' : isLeaveConflict ? '🚫' : ''} ${escHtml(n.title)}
            ${isOverlap ? `<span class="notif-badge">${severity.toUpperCase()}</span>` : ''}
            ${isLeave ? `<span class="notif-badge leave">${n.metadata?.leaveType?.toUpperCase()}</span>` : ''}
            ${isLeaveConflict ? `<span class="notif-badge awol">BLOCKED</span>` : ''}
          </div>
          <div class="notif-body">${escHtml(n.body)}</div>
          <div class="notif-time">${formatTime(new Date(n.timestamp))}</div>
        </div>
      `;
      })
      .join('');
  }
  openModal('notif-modal');
}

function handleNotifClick(notifId) {
  const notif = getNotifs().find((n) => n.id === notifId);
  if (!notif) return;
  notif.read = true;
  updateNotifBadge();
  API.put(`/notifications/${notifId}/read`).catch(() => { });
  if (notif.metadata?.type === 'overlap' && notif.taskId) {
    closeModal('notif-modal');
    openEditTask(notif.taskId);
    setTimeout(() => {
      const alertEl = document.getElementById('overlap-alert');
      const alertText = document.getElementById('overlap-alert-text');
      alertEl.className = `overlap-alert ${notif.metadata.severity === 'critical' ? 'critical' : ''}`;
      alertText.textContent = `This task has scheduling conflicts. Please review and reschedule if necessary.`;
      alertEl.classList.remove('hidden');
    }, 100);
  }
  if (notif.metadata?.type === 'leave-conflict' && notif.taskId) {
    closeModal('notif-modal');
    openEditTask(notif.taskId);
  }
}

function markNotifsRead() {
  (cache.notifications || []).forEach((n) => {
    if (n.userId === state.currentUser.id) n.read = true;
  });
  updateNotifBadge();
  API.put('/notifications/read-all').catch(() => { });
}

function clearNotifs() {
  cache.notifications = (cache.notifications || []).filter(
    (n) => n.userId !== state.currentUser.id,
  );
  closeModal('notif-modal');
  updateNotifBadge();
}

async function requestNotifPermission() {
  if (!('Notification' in window)) return;
  if (Notification.permission === 'default') {
    // Prompt user with a toast first, then request
    toast(
      '📱 Enable notifications to get task alerts outside the app.',
      'info',
    );
    setTimeout(async () => {
      const result = await Notification.requestPermission();
      if (result === 'granted') {
        toast('✅ Notifications enabled!', 'success');
      }
    }, 1500);
  }
}

// Show a browser/OS notification that appears OUTSIDE the website
function _showBrowserNotif(title, body, tag) {
  if (Notification.permission !== 'granted') return;
  // Only show when tab is NOT focused (so it appears outside the site)
  // Or always show via service worker (which shows even when focused)
  if ('serviceWorker' in navigator && navigator.serviceWorker.controller) {
    navigator.serviceWorker.controller.postMessage({
      type: 'SHOW_NOTIFICATION',
      title: title,
      body: body,
      tag: tag || 'taskflow-notif',
    });
  } else {
    // Fallback: direct Notification API (only shows if tab is not focused in some browsers)
    try {
      new Notification(title, {
        body: body,
        icon: '/favicon.ico',
        tag: tag || 'tf-notif',
      });
    } catch (e) {
      /* Mobile browsers may block Notification constructor */
    }
  }
}

/* ============================================================
   SERVICE WORKER
   ============================================================ */
async function registerSW() {
  if (!('serviceWorker' in navigator)) return;
  try {
    var reg = await navigator.serviceWorker.register('./sw.js');
    // Wait for SW to activate and claim this page so controller is available
    if (!navigator.serviceWorker.controller) {
      await new Promise(function (resolve) {
        navigator.serviceWorker.addEventListener('controllerchange', resolve, {
          once: true,
        });
        // Timeout after 3s in case SW doesn't claim
        setTimeout(resolve, 3000);
      });
    }
    console.log('[SW] Registered and controlling page');
  } catch (e) {
    console.warn('SW registration failed:', e);
  }
}

/* ============================================================
   DEADLINE CHECKER (interval)
   ============================================================ */
async function deleteCompletedTask(e, taskId) {
  if (e) e.stopPropagation();
  if (!confirm('Delete this completed task? This cannot be undone.')) return;
  const tasks = getTasks();
  const task = tasks.find((t) => t.id === taskId);
  if (!task || !task.done) {
    toast('Task is not completed.', 'error');
    return;
  }
  try {
    await API.del(`/tasks/${taskId}`);
    cache.tasks = cache.tasks.filter((t) => t.id !== taskId);
    addLog({
      taskId,
      taskTitle: task.title,
      action: 'deleted',
      actorName: state.currentUser.name,
      userId: task.userId,
    });
    toast('Completed task deleted.', 'success');
    if (state.view === 'calendar') {
      renderCalendarDebounced();
      return;
    }
    const uid =
      state.view === 'user-tasks' ? state.targetUserId : state.currentUser.id;
    if (state.currentViewMode === 'timeline') renderTimeline(uid);
    else renderTasks(uid);
  } catch (err) {
    toast(err.message || 'Failed to delete task.', 'error');
  }
}

async function autoDeleteEndOfMonthTasks() {
  const now = new Date();
  // Only run on the last day of the month at any time, or if it's past the last day
  const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  if (now.getDate() !== lastDay) return;
  const doneTasks = getTasks().filter(
    (t) => t.done && t.userId === state.currentUser?.id,
  );
  if (doneTasks.length === 0) return;
  for (const task of doneTasks) {
    try {
      await API.del(`/tasks/${task.id}`);
      cache.tasks = cache.tasks.filter((t) => t.id !== task.id);
    } catch (e) { }
  }
  if (doneTasks.length > 0) {
    toast(
      `🗑 Auto-deleted ${doneTasks.length} completed task${doneTasks.length > 1 ? 's' : ''} at end of month.`,
      'info',
    );
    const uid =
      state.view === 'user-tasks' ? state.targetUserId : state.currentUser?.id;
    if (uid) {
      if (state.currentViewMode === 'timeline') renderTimeline(uid);
      else renderTasks(uid);
    }
  }
}

let deadlineInterval = null;
function startDeadlineChecker() {
  clearInterval(deadlineInterval);
  deadlineInterval = setInterval(() => {
    const userId =
      state.view === 'user-tasks' ? state.targetUserId : state.currentUser?.id;
    if (!userId) return;
    const tasks = getTasks().filter((t) => t.userId === userId && !t.done);
    if ('serviceWorker' in navigator && navigator.serviceWorker.controller) {
      navigator.serviceWorker.controller.postMessage({
        type: 'CHECK_DEADLINES',
        tasks,
        userName: state.currentUser?.name,
      });
    }
    // Re-render to update flicker classes
    const now = new Date();
    tasks.forEach((task) => {
      const card = document.querySelector(`.task-card[data-id="${task.id}"]`);
      if (!card) return;
      const deadline = new Date(task.deadline);
      const hoursLeft = (deadline - now) / 1000 / 60 / 60;
      card.classList.remove('flicker-critical', 'flicker-warning');
      if (hoursLeft <= 24 && !task.done) card.classList.add('flicker-critical');
      else if (hoursLeft <= 72 && !task.done)
        card.classList.add('flicker-warning');
    });
    // Auto-delete completed tasks at end of month
    autoDeleteEndOfMonthTasks();
  }, 60000); // every minute
}

/* ============================================================
   MODAL HELPERS
   ============================================================ */
let _openModalCount = 0;
function openModal(id) {
  document.getElementById(id).classList.remove('hidden');
  _openModalCount++;
  document.body.classList.add('modal-open');
}
function closeModal(id) {
  document.getElementById(id).classList.add('hidden');
  _openModalCount = Math.max(0, _openModalCount - 1);
  if (_openModalCount === 0) document.body.classList.remove('modal-open');
}
function closeAllModals() {
  [
    'task-modal',
    'leave-modal',
    'wa-modal',
    'logs-modal',
    'notif-modal',
    'register-modal',
    'confirm-modal',
    'cancel-modal',
    'chpw-modal',
    'multi-task-modal',
    'worksettings-modal',
    'team-modal',
    'team-task-modal',
    'team-tasks-modal',
    'sa-company-modal',
    'sa-users-modal',
    'lieu-day-modal',
    'import-users-modal',
    'delete-all-modal',
  ].forEach((id) => document.getElementById(id).classList.add('hidden'));
  _openModalCount = 0;
  document.body.classList.remove('modal-open');
}
// Note: modal overlay click handlers are wired in initializeApp() after body.html is injected

/* ============================================================
   TOAST
   ============================================================ */
function toast(msg, type = 'info') {
  // Delegate to app-style push notification
  const titles = {
    success: 'Success',
    error: 'Error',
    info: 'TaskFlow',
    warning: 'Warning',
  };
  showPushNotification(titles[type] || 'TaskFlow', msg, type);
}

/* ============================================================
   APP-STYLE PUSH NOTIFICATION SYSTEM
   ============================================================ */
const PUSH_MAX = 4;
const PUSH_DURATION = 4500;

function showPushNotification(title, body, type = 'info') {
  const icons = { success: '✅', error: '❌', info: '💡', warning: '⚠️' };
  const container = document.getElementById('push-notif-container');
  if (!container) return;

  // Remove oldest if too many
  const existing = container.querySelectorAll('.push-notif');
  if (existing.length >= PUSH_MAX) dismissPushNotif(existing[0]);

  const el = document.createElement('div');
  el.className = 'push-notif';
  el.setAttribute('role', 'alert');
  el.setAttribute('aria-live', 'polite');

  const now = new Date();
  const timeStr = now.toLocaleTimeString('en-US', {
    hour: '2-digit',
    minute: '2-digit',
  });

  el.innerHTML = `
    <div class="push-notif-body">
      <div class="push-notif-icon ${type}"><span>${icons[type] || '💡'}</span></div>
      <div class="push-notif-text">
        <div class="push-notif-app">TaskFlow</div>
        <div class="push-notif-title">${escHtmlSimple(title)}</div>
        <div class="push-notif-message">${escHtmlSimple(body)}</div>
        <div class="push-notif-time">${timeStr}</div>
      </div>
    </div>
    <div class="push-notif-progress">
      <div class="push-notif-progress-bar ${type}" style="width:100%;"></div>
    </div>
  `;

  // Swipe to dismiss (touch)
  let startX = 0;
  el.addEventListener(
    'touchstart',
    (e) => {
      startX = e.touches[0].clientX;
    },
    { passive: true },
  );
  el.addEventListener('touchend', (e) => {
    const dx = e.changedTouches[0].clientX - startX;
    if (Math.abs(dx) > 60) dismissPushNotif(el);
  });
  el.addEventListener('click', () => dismissPushNotif(el));

  container.appendChild(el);

  // Animate progress bar
  const bar = el.querySelector('.push-notif-progress-bar');
  if (bar) {
    bar.style.transition = `width ${PUSH_DURATION}ms linear`;
    requestAnimationFrame(() => {
      bar.style.width = '0%';
    });
  }

  // Auto dismiss
  const timer = setTimeout(() => dismissPushNotif(el), PUSH_DURATION);
  el._dismissTimer = timer;
}

function dismissPushNotif(el) {
  if (!el || !el.parentNode) return;
  if (el._dismissTimer) clearTimeout(el._dismissTimer);
  el.classList.add('dismissing');
  setTimeout(() => {
    if (el.parentNode) el.parentNode.removeChild(el);
  }, 320);
}

function escHtmlSimple(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/* ============================================================
   MOBILE NAV HELPERS
   ============================================================ */
function mobileNavHome() {
  // ── Hide features based on role ──────────────────────────
  var mgCard = document.getElementById('manager-add-card');
  if (mgCard) {
    if (
      state.currentUser.role === 'admin' ||
      state.currentUser.role === 'manager'
    ) {
      mgCard.style.display = '';
    } else {
      mgCard.style.display = 'none';
    }
  }

  // Show Teams mobile nav icon for all roles
  var mobTeamsBtn = document.getElementById('mob-nav-teams');
  if (mobTeamsBtn) {
    mobTeamsBtn.style.display = '';
    if (typeof updateMobileNavTabCount === 'function')
      updateMobileNavTabCount();
  }
  updateMobileNavActive('mob-nav-home');
  if (!state.currentUser) return;
  if (
    state.currentUser.role === 'admin' ||
    state.currentUser.role === 'manager'
  ) {
    showView('admin-home');
  } else {
    showView('worker-home');
  }
}

function mobileNavTasks() {
  updateMobileNavActive('mob-nav-tasks');
  if (!state.currentUser) return;
  state.view = 'my-tasks';
  showView('my-tasks');
}

function mobileNavTeams() {
  if (!state.currentUser) return;
  if (state.currentUser.role === 'user') return;
  updateMobileNavActive('mob-nav-teams');
  showView('teams');
}

function mobileNavCal() {
  if (!state.currentUser) return;
  updateMobileNavActive('mob-nav-cal');
  // Always reset jump state so calendar reloads fresh
  state._calendarJumpDate = null;
  state._calendarHighlightUserId = null;
  state._calendarFocusTaskId = null;
  const isElevated =
    state.currentUser.role === 'admin' || state.currentUser.role === 'manager';
  state.calendarMode = isElevated ? 'team' : 'personal';

  // If cache is not yet populated, wait for it then render
  if (!cache.tasks || !cache.users) {
    showView('calendar');
    loadAll().then(function () {
      renderCalendar();
    });
  } else {
    showView('calendar');
  }
}

function updateMobileNavActive(activeId) {
  document
    .querySelectorAll('.mob-nav-item')
    .forEach((btn) => btn.classList.remove('active'));
  const el = document.getElementById(activeId);
  if (el) el.classList.add('active');
  // Hide "New Task" button on calendar view; show on all other views including tasks
  var addBtn = document.getElementById('mob-nav-add');
  if (addBtn) {
    if (activeId === 'mob-nav-cal') {
      addBtn.style.display = 'none';
    } else {
      addBtn.style.display = '';
    }
  }

  // Permanently hide teams icon for regular users
  var mobTeamsBtn = document.getElementById('mob-nav-teams');
  if (mobTeamsBtn && state.currentUser && state.currentUser.role === 'user') {
    mobTeamsBtn.style.display = 'none';
  }

  // Remove / Hide Back Button on Mobile Calendar
  var backBtn = document.getElementById('nav-back-btn');
  if (backBtn) backBtn.style.display = activeId === 'mob-nav-cal' ? 'none' : '';
}

function updateMobileNavNotifBadge() {
  const badge = document.getElementById('mob-notif-badge');
  if (!badge || !state.currentUser) return;
  const unread = getNotifs().filter(
    (n) => n.userId === state.currentUser.id && !n.read,
  ).length;
  if (unread > 0) {
    badge.textContent = unread > 9 ? '9+' : unread;
    badge.classList.remove('hidden');
  } else {
    badge.classList.add('hidden');
  }
}

/* ============================================================
   HELPERS
   ============================================================ */
function escHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formatDeadline(date) {
  const now = new Date();
  const diff = date - now;
  const hours = Math.round(diff / 1000 / 60 / 60);
  const days = Math.round(diff / 1000 / 60 / 60 / 24);
  if (diff < 0)
    return date.toLocaleDateString('en-US', {
      month: 'short',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    });
  if (hours < 1) return 'Due in <1 hour';
  if (hours < 24) return `Due in ${hours}h`;
  if (days === 1) return 'Due tomorrow';
  if (days < 7) return `Due in ${days} days`;
  return date.toLocaleDateString('en-US', {
    month: 'short',
    day: 'numeric',
    year: date.getFullYear() !== now.getFullYear() ? 'numeric' : undefined,
    hour: '2-digit',
    minute: '2-digit',
  });
}

function formatTime(date) {
  const now = new Date();
  const diff = now - date;
  if (diff < 60000) return 'just now';
  if (diff < 3600000) return Math.round(diff / 60000) + 'm ago';
  if (diff < 86400000) return Math.round(diff / 3600000) + 'h ago';
  return date.toLocaleDateString('en-US', {
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });
}

/* ============================================================
   TOOLTIP ENGINE
   ============================================================ */
document.addEventListener('DOMContentLoaded', function () {
  (function () {
    const tip = document.createElement('div');
    tip.className = 'tf-tooltip';
    document.body.appendChild(tip);

    let showTimer = null;
    let hideTimer = null;
    let currentTarget = null;

    function showTip(el, text) {
      clearTimeout(hideTimer);
      tip.textContent = text;
      tip.classList.remove('visible', 'tip-above');

      // Position
      const rect = el.getBoundingClientRect();
      const tipW = 220;
      let left = rect.left + rect.width / 2;
      let top = rect.bottom + 10;

      // Flip above if it would go off-screen bottom
      const approxH = 40;
      if (top + approxH > window.innerHeight - 20) {
        top = rect.top - approxH - 10;
        tip.classList.add('tip-above');
      }

      // Clamp left
      left = Math.max(10, Math.min(left, window.innerWidth - tipW / 2 - 10));

      tip.style.left = left + 'px';
      tip.style.top = top + 'px';
      tip.style.transform = 'translateX(-50%)';

      // Force reflow then show
      tip.getBoundingClientRect();
      tip.classList.add('visible');
    }

    function hideTip() {
      clearTimeout(showTimer);
      tip.classList.remove('visible');
      currentTarget = null;
    }

    document.addEventListener('mouseover', (e) => {
      const el = e.target.closest('[data-tip]');
      if (!el) return;
      if (el === currentTarget) return;
      currentTarget = el;
      clearTimeout(showTimer);
      clearTimeout(hideTimer);
      showTimer = setTimeout(() => {
        if (currentTarget === el) showTip(el, el.dataset.tip);
      }, 700); // 700ms delay
    });

    document.addEventListener('mouseout', (e) => {
      const el = e.target.closest('[data-tip]');
      if (!el) return;
      clearTimeout(showTimer);
      hideTimer = setTimeout(hideTip, 100);
    });

    // Hide on scroll, click or key
    document.addEventListener('scroll', hideTip, true);
    document.addEventListener('click', hideTip, true);
    document.addEventListener('keydown', hideTip, true);
  })();
}); // end DOMContentLoaded (tooltip)

/* ============================================================
   TEAMS — JS
   ============================================================ */
let _editingTeamId = null;
let ttDescItems = [];

// ---- Render the Teams grid view ----
function renderTeamsView() {
  const grid = document.getElementById('teams-grid');
  const teams = getTeams();
  const users = getUsers();
  const isElevated =
    state.currentUser.role === 'admin' || state.currentUser.role === 'manager';

  if (teams.length === 0) {
    grid.innerHTML = `<div style="grid-column:1/-1;text-align:center;padding:80px 20px;color:var(--text3);">
      <div style="font-size:48px;margin-bottom:16px;opacity:0.3;">🏷️</div>
      <div style="font-family:var(--display);font-size:18px;color:var(--text2);margin-bottom:8px;">No Teams Yet</div>
      <div style="font-size:13px;">Create a team to group people and assign tasks together.</div>
    </div>`;
    return;
  }

  grid.innerHTML = teams
    .map((team) => {
      const members = (team.memberIds || [])
        .map((id) => users.find((u) => u.id === id))
        .filter(Boolean);
      const color = team.color || '#F59E0B';
      const memberChips =
        members.length === 0
          ? '<div style="font-size:11px;color:var(--text3);padding:8px 0;">No members assigned</div>'
          : members
            .map(
              (m) => `
        <div class="team-member-chip">
          <div class="team-member-avatar" style="background:${color}18;border-color:${color}55;color:${color};">${escHtml((m.name || '?')[0].toUpperCase())}</div>
          <div class="team-member-name">${escHtml(m.name)}</div>
          <div class="team-member-role">${m.role}</div>
        </div>`,
            )
            .join('');

      const editBtn = isElevated
        ? `<button class="btn-secondary" style="font-size:11px;padding:6px 12px;" onclick="openEditTeam('${team.id}')">✏️ Edit</button>`
        : '';
      const assignBtn = isElevated
        ? `<button class="btn-primary" style="width:auto;font-size:11px;padding:6px 14px;" onclick="openTeamTaskModal('${team.id}')">📋 Assign Multi Team Task</button>`
        : '';
      const viewTasksBtn = `<button class="btn-secondary" style="font-size:11px;padding:6px 12px;" onclick="showTeamTasks('${team.id}')">📌 View Tasks</button>`;

      return `<div class="team-card">
      <div class="team-card-header">
        <div style="display:flex;align-items:center;gap:8px;">
          <div class="team-color-dot" style="width:12px;height:12px;border-radius:50%;background:${color};flex-shrink:0;"></div>
          <div class="team-card-name">${escHtml(team.name)}</div>
        </div>
        <div class="team-card-badge">${members.length} member${members.length !== 1 ? 's' : ''}</div>
      </div>
      <div class="team-members-list">${memberChips}</div>
      <div class="team-card-actions">${viewTasksBtn}${editBtn}${assignBtn}</div>
    </div>`;
    })
    .join('');
}

// ---- Show tasks per team ----
function showTeamTasks(teamId) {
  const team = getTeams().find((t) => t.id === teamId);
  if (!team) return;

  const allTasks = getTasks();
  const teamTasks = allTasks.filter(
    (t) => t.teamId === teamId || (t.isTeamTask && t.teamId === teamId),
  );

  const active = teamTasks.filter((t) => !t.done && !t.cancelled);
  const done = teamTasks.filter((t) => t.done);
  const cancelled = teamTasks.filter((t) => t.cancelled);

  const color = team.color || '#F59E0B';
  const users = getUsers();

  function _taskRow(task) {
    const u = users.find((u) => u.id === task.userId);
    const assignee = u ? u.name : 'Unknown';
    const start = task.start
      ? new Date(task.start).toLocaleDateString('en-GB', {
        day: '2-digit',
        month: 'short',
      })
      : '—';
    const deadline = task.deadline
      ? new Date(task.deadline).toLocaleDateString('en-GB', {
        day: '2-digit',
        month: 'short',
      })
      : '—';
    const pColors = {
      1: '#EF4444',
      2: '#F97316',
      3: '#F59E0B',
      4: '#3B82F6',
      5: '#6B7280',
    };
    const pLabel = ['', 'P1', 'P2', 'P3', 'P4', 'P5'][task.priority] || '';
    const pColor = pColors[task.priority] || '#6B7280';
    const statusBadge = task.done
      ? '<span style="font-size:10px;padding:2px 8px;border-radius:20px;background:rgba(52,211,153,0.15);color:#10B981;">✓ Done</span>'
      : task.cancelled
        ? '<span style="font-size:10px;padding:2px 8px;border-radius:20px;background:rgba(239,68,68,0.12);color:#EF4444;">✗ Cancelled</span>'
        : '<span style="font-size:10px;padding:2px 8px;border-radius:20px;background:rgba(245,158,11,0.15);color:var(--amber);">Active</span>';
    const isElevated =
      state.currentUser.role === 'admin' ||
      state.currentUser.role === 'manager';
    const editBtn =
      isElevated && !task.done && !task.cancelled
        ? `<button onclick="closeModal('team-tasks-modal');openEditTask('${task.id}')" style="font-size:10px;padding:3px 10px;border-radius:6px;border:1px solid var(--border);background:var(--bg3);color:var(--text2);cursor:pointer;">✏️ Edit</button>`
        : '';
    return `<div style="display:flex;align-items:flex-start;gap:10px;padding:10px 12px;background:var(--bg3);border-radius:10px;border:1px solid var(--border);">
      <div style="width:28px;height:28px;border-radius:50%;background:${pColor}22;border:1.5px solid ${pColor};display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:800;color:${pColor};flex-shrink:0;">${pLabel}</div>
      <div style="flex:1;min-width:0;">
        <div style="font-size:13px;font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${escHtml(task.title)}</div>
        <div style="font-size:11px;color:var(--text3);margin-top:2px;">👤 ${escHtml(assignee)} &nbsp;·&nbsp; 📅 ${start} → ${deadline}</div>
      </div>
      <div style="display:flex;align-items:center;gap:6px;flex-shrink:0;">${statusBadge}${editBtn}</div>
    </div>`;
  }

  function _section(label, tasks, emptyMsg) {
    if (tasks.length === 0)
      return `<div style="font-size:12px;color:var(--text3);padding:8px 4px;">${emptyMsg}</div>`;
    return `<div style="font-size:10px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:0.08em;margin-bottom:6px;">${label} (${tasks.length})</div>
      <div style="display:flex;flex-direction:column;gap:6px;margin-bottom:16px;">${tasks.map(_taskRow).join('')}</div>`;
  }

  const body = document.getElementById('team-tasks-modal-body');
  body.innerHTML = `
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:16px;">
      <div style="width:14px;height:14px;border-radius:50%;background:${color};flex-shrink:0;"></div>
      <div style="font-family:var(--display);font-size:17px;font-weight:800;">${escHtml(team.name)}</div>
      <div style="font-size:11px;color:var(--text3);background:var(--bg3);padding:3px 10px;border-radius:20px;border:1px solid var(--border);">${teamTasks.length} task${teamTasks.length !== 1 ? 's' : ''} total</div>
    </div>
    ${_section('Active', active, 'No active tasks for this team.')}
    ${done.length > 0 ? _section('Completed', done, '') : ''}
    ${cancelled.length > 0 ? _section('Cancelled', cancelled, '') : ''}
  `;

  openModal('team-tasks-modal');
}

// ---- Create / Edit Team modal ----
function openCreateTeam() {
  _editingTeamId = null;
  document.getElementById('team-modal-title').textContent = '🏷️ Create Team';
  document.getElementById('tm-name').value = '';
  document.getElementById('tm-search').value = '';
  document.getElementById('tm-delete-btn').classList.add('hidden');
  // Reset color picker
  const radios = document.querySelectorAll('input[name="tm-color"]');
  if (radios.length) radios[0].checked = true;
  renderTeamMemberPicker('');
  openModal('team-modal');
  setTimeout(() => document.getElementById('tm-name').focus(), 100);
}

function openEditTeam(teamId) {
  const team = getTeams().find((t) => t.id === teamId);
  if (!team) return;
  _editingTeamId = teamId;
  document.getElementById('team-modal-title').textContent = '✏️ Edit Team';
  document.getElementById('tm-name').value = team.name || '';
  document.getElementById('tm-search').value = '';
  document.getElementById('tm-delete-btn').classList.remove('hidden');
  // Set color
  const radios = document.querySelectorAll('input[name="tm-color"]');
  radios.forEach((r) => {
    r.checked = r.value === team.color;
  });
  renderTeamMemberPicker('');
  openModal('team-modal');
}

function renderTeamMemberPicker(filter) {
  const container = document.getElementById('tm-member-list');
  const currentTeam = _editingTeamId
    ? getTeams().find((t) => t.id === _editingTeamId)
    : null;
  const currentMembers = currentTeam ? currentTeam.memberIds || [] : [];
  // Preserve checked state from DOM
  const checked = {};
  container.querySelectorAll('input[type="checkbox"]').forEach((cb) => {
    checked[cb.value] = cb.checked;
  });
  // Merge with saved members for new opens (before first interaction)
  currentMembers.forEach((id) => {
    if (!(id in checked)) checked[id] = true;
  });

  const users = getUsers().filter((u) => u.role !== 'admin');
  const q = (filter || '').toLowerCase();
  const filtered = q
    ? users.filter(
      (u) =>
        u.name.toLowerCase().includes(q) ||
        u.username.toLowerCase().includes(q),
    )
    : users;

  const roleIcon = { manager: '🏢', user: '👤' };
  container.innerHTML =
    filtered.length === 0
      ? '<div style="padding:12px;text-align:center;color:var(--text3);font-size:12px;">No users found</div>'
      : filtered
        .map(
          (u) => `
      <label class="team-modal-member-item">
        <input type="checkbox" value="${escHtml(u.id)}" ${checked[u.id] ? 'checked' : ''} onchange="updateTmSelectedCount()">
        <div style="width:28px;height:28px;border-radius:50%;background:var(--bg4);border:1px solid var(--border);display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;color:var(--text2);flex-shrink:0;">${escHtml((u.name || '?')[0].toUpperCase())}</div>
        <div style="flex:1;min-width:0;">
          <div style="font-size:13px;font-weight:600;">${escHtml(u.name)}</div>
          <div style="font-size:10px;color:var(--text3);">${roleIcon[u.role] || '👤'} ${u.role} · @${escHtml(u.username)}</div>
        </div>
      </label>`,
        )
        .join('');

  updateTmSelectedCount();
}

function updateTmSelectedCount() {
  const n = document.querySelectorAll('#tm-member-list input:checked').length;
  const el = document.getElementById('tm-selected-count');
  if (el) el.textContent = n + ' selected';
}

function tmSelectAll() {
  document
    .querySelectorAll('#tm-member-list input[type="checkbox"]')
    .forEach((cb) => (cb.checked = true));
  updateTmSelectedCount();
}
function tmClearAll() {
  document
    .querySelectorAll('#tm-member-list input[type="checkbox"]')
    .forEach((cb) => (cb.checked = false));
  updateTmSelectedCount();
}

async function saveTeam() {
  const _btn = document.querySelector(
    '#team-modal .btn-primary[onclick="saveTeam()"]',
  );
  if (!_lockOp('saveTeam', _btn, 'Saving…')) return;
  const name = document.getElementById('tm-name').value.trim();
  if (!name) {
    toast('Team name is required.', 'error');
    return;
  }
  const color =
    document.querySelector('input[name="tm-color"]:checked')?.value ||
    '#F59E0B';
  const memberIds = Array.from(
    document.querySelectorAll('#tm-member-list input:checked'),
  ).map((cb) => cb.value);

  try {
    if (_editingTeamId) {
      const updated = await API.put('/teams/' + _editingTeamId, {
        name,
        color,
        memberIds,
      });
      const idx = cache.teams.findIndex((t) => t.id === _editingTeamId);
      if (idx !== -1) cache.teams[idx] = updated;
      toast('Team updated!', 'success');
    } else {
      const newTeam = await API.post('/teams', { name, color, memberIds });
      cache.teams.push(newTeam);
      toast('Team "' + name + '" created! ✓', 'success');
    }
    closeModal('team-modal');
    renderTeamsView();
  } catch (err) {
    toast(err.message || 'Failed to save team.', 'error');
  } finally {
    _unlockOp('saveTeam', _btn);
  }
}

async function deleteTeam() {
  if (!_editingTeamId) return;
  const _dBtn = document.getElementById('tm-delete-btn');
  if (!_lockOp('deleteTeam', _dBtn, 'Deleting…')) return;
  const team = getTeams().find((t) => t.id === _editingTeamId);
  if (
    !confirm(
      'Delete team "' +
      (team ? team.name : '') +
      '"? This does not delete the users or their tasks.',
    )
  )
    return;
  try {
    await API.del('/teams/' + _editingTeamId);
    cache.teams = cache.teams.filter((t) => t.id !== _editingTeamId);
    closeModal('team-modal');
    toast('Team deleted.', 'info');
    renderTeamsView();
  } catch (err) {
    toast(err.message || 'Failed to delete team.', 'error');
  } finally {
    _unlockOp('deleteTeam', _dBtn);
  }
}

// ---- Team Task Duration/EndDate toggle ----
var _ttScheduleMode = 'duration';
function setTtScheduleMode(mode) {
  _ttScheduleMode = mode;
  const activeStyle =
    'flex:1;padding:7px;border-radius:6px;font-size:11px;font-weight:700;border:1.5px solid var(--amber);background:var(--amber);color:var(--bg);cursor:pointer;transition:all 0.15s;';
  const inactiveStyle =
    'flex:1;padding:7px;border-radius:6px;font-size:11px;font-weight:700;border:1.5px solid var(--border);background:var(--bg3);color:var(--text2);cursor:pointer;transition:all 0.15s;';
  const durBtn = document.getElementById('tt-mode-duration');
  const endBtn = document.getElementById('tt-mode-enddate');
  const durGrp = document.getElementById('tt-duration-group');
  const endGrp = document.getElementById('tt-enddate-group');
  if (mode === 'duration') {
    durBtn.style.cssText = activeStyle;
    endBtn.style.cssText = inactiveStyle;
    durGrp.style.display = '';
    endGrp.style.display = 'none';
  } else {
    endBtn.style.cssText = activeStyle;
    durBtn.style.cssText = inactiveStyle;
    endGrp.style.display = '';
    durGrp.style.display = 'none';
  }
  updateTtDeadlinePreview();
  checkTeamTaskConflicts();
}
function getTtDeadline() {
  const start = document.getElementById('tt-start').value;
  if (!start) return null;
  if (_ttScheduleMode === 'enddate') {
    const ed = document.getElementById('tt-deadline')
      ? document.getElementById('tt-deadline').value
      : '';
    return ed || null;
  }
  const raw = parseFloat(document.getElementById('tt-duration').value);
  const unit = document.getElementById('tt-duration-unit').value;
  if (!raw || raw <= 0) return null;
  const hours = unit === 'days' ? raw * 8 : raw;
  // Simple wall-clock: start + hours, converted back to local datetime string
  const end = new Date(new Date(start).getTime() + hours * 3600000);
  return toLocalISO(end);
}
function updateTtDeadlinePreview() {
  const prev = document.getElementById('tt-deadline-preview');
  if (!prev) return;
  const dl = getTtDeadline();
  if (dl) {
    const d = new Date(dl);
    prev.textContent =
      '⏱ Ends: ' +
      d.toLocaleString('en-US', {
        weekday: 'short',
        month: 'short',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
      });
  } else {
    prev.textContent = '';
  }
}

// ---- Team Task Assignment ----
function openTeamTaskModal(preselectedTeamId) {
  ttDescItems = [];
  _ttScheduleMode = 'duration';
  document.getElementById('tt-title').value = '';
  document.getElementById('tt-requestor').value = state.currentUser.name;
  document.getElementById('tt-priority').value = '3';
  document.getElementById('tt-overlap-alert').classList.add('hidden');
  document.getElementById('tt-leave-alert').classList.add('hidden');
  document.getElementById('tt-desc-items').innerHTML = '';
  document.getElementById('tt-desc-input').value = '';

  const now = new Date();
  document.getElementById('tt-start').value = toLocalISO(now);
  // Reset duration fields
  if (document.getElementById('tt-duration'))
    document.getElementById('tt-duration').value = '1';
  if (document.getElementById('tt-duration-unit'))
    document.getElementById('tt-duration-unit').value = 'hours';
  setTtScheduleMode('duration');

  // Populate team selector — if opened from a specific team card, lock to that team only
  const teams = getTeams();
  const teamList = document.getElementById('tt-team-list');
  if (preselectedTeamId) {
    // Locked to one team — just show its name, no other choices
    const selTeam = teams.find((t) => t.id === preselectedTeamId);
    teamList.innerHTML = selTeam
      ? `<div style="display:flex;align-items:center;gap:8px;padding:6px 8px;background:${selTeam.color || '#F59E0B'}15;border:1px solid ${selTeam.color || '#F59E0B'}44;border-radius:8px;">
           <div style="width:10px;height:10px;border-radius:50%;background:${selTeam.color || '#F59E0B'};flex-shrink:0;"></div>
           <span style="font-size:13px;font-weight:700;color:var(--text);">${escHtml(selTeam.name)}</span>
           <span style="font-size:11px;color:var(--text3);margin-left:auto;">${(selTeam.memberIds || []).length} members</span>
           <input type="checkbox" value="${escHtml(selTeam.id)}" checked style="display:none;">
         </div>`
      : '<div style="color:var(--danger);font-size:12px;">Team not found.</div>';
  } else {
    // Generic open — show all teams as checkboxes
    teamList.innerHTML =
      teams.length === 0
        ? '<div style="color:var(--text3);font-size:12px;padding:4px;">No teams yet. Create a team first.</div>'
        : teams
          .map(
            (t) => `
        <label style="display:flex;align-items:center;gap:8px;cursor:pointer;padding:4px 6px;border-radius:6px;transition:background .15s;"
          onmouseenter="this.style.background='var(--bg4)'" onmouseleave="this.style.background=''">
          <input type="checkbox" value="${escHtml(t.id)}"
            onchange="onTeamTaskTeamChange()" style="width:14px;height:14px;accent-color:var(--amber);cursor:pointer;">
          <span style="font-size:12px;font-weight:600;color:var(--text);">${escHtml(t.name)}</span>
          <span style="font-size:11px;color:var(--text3);margin-left:auto;">${(t.memberIds || []).length} members</span>
        </label>`,
          )
          .join('');
  }

  document.getElementById('tt-members-preview').style.display = 'none';
  document.getElementById('tt-members-list').innerHTML = '';
  document.getElementById('tt-team-count').textContent = '(none selected)';

  if (preselectedTeamId) onTeamTaskTeamChange();

  openModal('team-task-modal');
}

function onTeamTaskTeamChange() {
  const checkedIds = Array.from(
    document.querySelectorAll('#tt-team-list input[type=checkbox]:checked'),
  ).map((c) => c.value);
  const preview = document.getElementById('tt-members-preview');
  const list = document.getElementById('tt-members-list');
  const countEl = document.getElementById('tt-team-count');

  if (checkedIds.length === 0) {
    preview.style.display = 'none';
    countEl.textContent = '(none selected)';
    return;
  }

  countEl.textContent =
    '(' +
    checkedIds.length +
    ' team' +
    (checkedIds.length > 1 ? 's' : '') +
    ' selected)';

  // Collect unique members across all selected teams
  const teams = getTeams();
  const users = getUsers();
  const seenMemberIds = new Set();
  const memberEntries = [];

  checkedIds.forEach((teamId) => {
    const team = teams.find((t) => t.id === teamId);
    if (!team) return;
    (team.memberIds || []).forEach((id) => {
      if (!seenMemberIds.has(id)) {
        seenMemberIds.add(id);
        memberEntries.push({
          id,
          teamColor: team.color || '#F59E0B',
          teamName: team.name,
        });
      }
    });
  });

  if (memberEntries.length === 0) {
    preview.style.display = 'none';
    toast('Selected teams have no members yet.', 'warning');
    return;
  }

  list.innerHTML = memberEntries
    .map(({ id, teamColor, teamName }) => {
      const u = users.find((u) => u.id === id);
      if (!u) return '';
      return `<div title="${escHtml(teamName)}" style="display:flex;align-items:center;gap:6px;padding:4px 8px;background:${teamColor}18;border:1px solid ${teamColor}33;border-radius:8px;font-size:12px;font-weight:600;">
      <div style="width:20px;height:20px;border-radius:50%;background:${teamColor}33;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;color:${teamColor};">${escHtml((u.name || '?')[0].toUpperCase())}</div>
      ${escHtml(u.name)}
    </div>`;
    })
    .join('');
  preview.style.display = '';
  checkTeamTaskConflicts();
}

function checkTeamTaskConflicts() {
  const checkedIds = Array.from(
    document.querySelectorAll('#tt-team-list input[type=checkbox]:checked'),
  ).map((c) => c.value);
  const start = document.getElementById('tt-start').value;
  const deadline = getTtDeadline
    ? getTtDeadline()
    : document.getElementById('tt-deadline')
      ? document.getElementById('tt-deadline').value
      : null;
  if (checkedIds.length === 0 || !start || !deadline) return;

  // Gather unique member IDs across all selected teams
  const teams = getTeams();
  const memberIds = new Set();
  checkedIds.forEach((tid) => {
    const t = teams.find((t) => t.id === tid);
    if (t) (t.memberIds || []).forEach((id) => memberIds.add(id));
  });
  if (memberIds.size === 0) return;

  const overlapMsgs = [],
    leaveMsgs = [];
  const allTasks = getTasks();
  const typeLabels = {
    lieu: 'Lieu Day',
    loa: 'Leave of Absence',
    awol: 'AWOL',
  };

  memberIds.forEach((uid) => {
    const user = getUsers().find((u) => u.id === uid);
    const name = user ? user.name : uid;
    const userTasks = allTasks.filter(
      (t) => t.userId === uid && !t.cancelled && !t.done,
    );
    const tempTask = {
      id: 'temp',
      start: new Date(start).toISOString(),
      deadline: new Date(deadline).toISOString(),
    };
    const overlaps = checkTaskOverlap(tempTask, userTasks);
    if (overlaps.length)
      overlapMsgs.push(
        name +
        ': conflicts with "' +
        overlaps.map((o) => o.title).join('", "') +
        '"',
      );
    const lc = checkLeaveConflict(uid, start, deadline);
    if (lc.length)
      leaveMsgs.push(
        name + ' is on ' + lc.map((l) => typeLabels[l.type]).join(', '),
      );
  });

  const overlapAlert = document.getElementById('tt-overlap-alert');
  const leaveAlert = document.getElementById('tt-leave-alert');
  if (overlapMsgs.length) {
    document.getElementById('tt-overlap-text').textContent =
      overlapMsgs.join(' | ');
    overlapAlert.classList.remove('hidden');
  } else overlapAlert.classList.add('hidden');
  if (leaveMsgs.length) {
    document.getElementById('tt-leave-text').textContent =
      leaveMsgs.join(' | ');
    leaveAlert.classList.remove('hidden');
  } else leaveAlert.classList.add('hidden');
}

function addTtDescItem() {
  const input = document.getElementById('tt-desc-input');
  const text = input.value.trim();
  if (!text) return;
  ttDescItems.push({ text, checked: false });
  input.value = '';
  renderTtDescItems();
  input.focus();
}
function removeTtDescItem(idx) {
  ttDescItems.splice(idx, 1);
  renderTtDescItems();
}
function renderTtDescItems() {
  document.getElementById('tt-desc-items').innerHTML = ttDescItems
    .map(
      (item, i) =>
        `<div class="desc-item"><span>${escHtml(item.text)}</span><button onclick="removeTtDescItem(${i})">✕</button></div>`,
    )
    .join('');
}

async function saveTeamTask() {
  const _btn = document.querySelector(
    '#team-task-modal .btn-primary[onclick="saveTeamTask()"]',
  );
  if (!_lockOp('saveTeamTask', _btn, 'Assigning…')) return;
  const checkedTeamIds = Array.from(
    document.querySelectorAll('#tt-team-list input[type=checkbox]:checked'),
  ).map((c) => c.value);
  const title = document.getElementById('tt-title').value.trim();
  const requestor = document.getElementById('tt-requestor').value.trim();
  const priority = parseInt(document.getElementById('tt-priority').value);
  const start = document.getElementById('tt-start').value;
  const deadlineRaw = getTtDeadline();
  const deadline = deadlineRaw || '';

  if (checkedTeamIds.length === 0) {
    toast('Please select at least one team.', 'error');
    _unlockOp('saveTeamTask', _btn);
    return;
  }
  if (!title || !requestor || !start || !deadline) {
    toast('Please fill in all required fields.', 'error');
    _unlockOp('saveTeamTask', _btn);
    return;
  }
  if (new Date(start) >= new Date(deadline)) {
    toast('End time must be after start time.', 'error');
    _unlockOp('saveTeamTask', _btn);
    return;
  }

  const _wh = getWorkHours();
  const _startHr =
    new Date(start).getHours() + new Date(start).getMinutes() / 60;
  const _endHr =
    new Date(deadline).getHours() + new Date(deadline).getMinutes() / 60;
  const _pad = (h) => String(h).padStart(2, '0');
  if (
    _startHr < _wh.start ||
    _endHr > _wh.end ||
    _startHr >= _wh.end ||
    _endHr <= _wh.start
  ) {
    toast(
      'Bookings must be within work hours: ' +
      _pad(_wh.start) +
      ':00 – ' +
      _pad(_wh.end) +
      ':00',
      'error',
    );
    return;
  }

  // Collect unique member IDs and their team name (first team they belong to)
  const allTeams = getTeams();
  const memberTeamMap = new Map(); // memberId → first team they're in
  checkedTeamIds.forEach((teamId) => {
    const team = allTeams.find((t) => t.id === teamId);
    if (!team) return;
    (team.memberIds || []).forEach((memberId) => {
      if (!memberTeamMap.has(memberId)) memberTeamMap.set(memberId, team);
    });
  });

  if (memberTeamMap.size === 0) {
    toast('Selected teams have no members.', 'error');
    return;
  }

  const description = ttDescItems.map((d) => ({
    text: d.text,
    checked: false,
  }));
  const groupId = 'tg-' + uid();
  const allTasks = getTasks();
  let assignedCount = 0;
  const skippedLeave = [];

  // Collect all member IDs upfront for the group metadata
  const allMemberIds = Array.from(memberTeamMap.keys());
  const groupMeta = _makeGroupMeta(groupId, allMemberIds);
  const descWithMeta = _buildDesc(description, groupMeta);

  for (const [memberId, team] of memberTeamMap) {
    const user = getUsers().find((u) => u.id === memberId);
    const name = user ? user.name : memberId;
    const lc = checkLeaveConflict(memberId, start, deadline);
    if (lc.length) {
      skippedLeave.push(name);
      continue;
    }
    const existingTasks = allTasks.filter(
      (t) => t.userId === memberId && !t.cancelled && !t.done,
    );
    const taskData = {
      title,
      requestor,
      priority,
      start: new Date(start).toISOString(),
      deadline: new Date(deadline).toISOString(),
      description: descWithMeta,
      isTeamTask: true,
      teamId: team.id,
      teamName: team.name,
      multiGroupId: groupId,
    };
    const overlapCheck = resolveOverlapsNew(taskData, existingTasks, memberId);
    if (overlapCheck.moved) {
      taskData.start = overlapCheck.newStart;
      taskData.deadline = overlapCheck.newDeadline;
    }
    try {
      const newTask = await API.post('/tasks', {
        ...taskData,
        userId: memberId,
      });
      if (!cache.tasks) cache.tasks = [];
      if (newTask) cache.tasks.push(newTask);
      const teamNames = checkedTeamIds
        .map((tid) => {
          const t = allTeams.find((t) => t.id === tid);
          return t ? t.name : '';
        })
        .filter(Boolean)
        .join(', ');
      pushNotification(
        memberId,
        '🏷️ Team Task Assigned',
        '"' +
        title +
        '" has been assigned to you via team(s) "' +
        teamNames +
        '" by ' +
        state.currentUser.name +
        '.',
        newTask?.id,
      );
      if (newTask)
        addLog({
          taskId: newTask.id,
          taskTitle: newTask.title,
          action: 'added',
          actorName: state.currentUser.name,
          userId: memberId,
        });
      assignedCount++;
    } catch (e) {
      console.error('Failed to assign to', memberId, e);
    }
  }

  _unlockOp('saveTeamTask', _btn);
  closeModal('team-task-modal');
  if (skippedLeave.length > 0) {
    toast(
      'Task assigned to ' +
      assignedCount +
      ' member(s). Skipped: ' +
      skippedLeave.join(', ') +
      ' (on leave).',
      'warning',
    );
  } else {
    toast(
      'Team task assigned to ' + assignedCount + ' member(s)! ✓',
      'success',
    );
  }
  renderTeamsView();
}

/* ============================================================
   CALENDAR TOAST — quick navigation shortcut
   ============================================================ */
var _calToastShown = false;
function showCalendarToast() {
  if (_calToastShown) return;
  _calToastShown = true;
  // Use the existing push notification system
  const container = document.getElementById('push-notif-container');
  if (!container) return;
  const el = document.createElement('div');
  el.className = 'push-notif';
  el.style.cursor = 'pointer';
  el.innerHTML = `
    <div class="push-notif-body">
      <div class="push-notif-icon info" style="font-size:20px;">📅</div>
      <div class="push-notif-text">
        <div class="push-notif-app">TaskFlow</div>
        <div class="push-notif-title">View Your Calendar</div>
        <div class="push-notif-message">Tap to open your schedule and leave calendar</div>
      </div>
      <button onclick="event.stopPropagation();this.closest('.push-notif').remove();" style="background:none;border:none;color:rgba(255,255,255,0.4);font-size:16px;cursor:pointer;padding:0 4px;align-self:flex-start;margin-top:-2px;">×</button>
    </div>
    <div class="push-notif-progress"><div class="push-notif-progress-bar info" style="width:100%;transition:width 6s linear;"></div></div>
  `;
  el.onclick = function () {
    el.classList.add('dismissing');
    setTimeout(() => el.remove(), 300);
    goToMyCalendar();
  };
  container.appendChild(el);
  // Animate progress bar
  requestAnimationFrame(() => {
    const bar = el.querySelector('.push-notif-progress-bar');
    if (bar) {
      requestAnimationFrame(() => {
        bar.style.width = '0%';
      });
    }
  });
  // Auto-dismiss after 6s
  setTimeout(() => {
    if (el.parentNode) {
      el.classList.add('dismissing');
      setTimeout(() => el.remove(), 300);
    }
  }, 6000);
}

/* ============================================================
   WORK SETTINGS (Admin) — multi-schedule with history
   ============================================================ */
function openWorkSettings() {
  renderWorkScheduleList();
  document.getElementById('ws-new-name').value = '';
  document.getElementById('ws-new-start').value = '9';
  document.getElementById('ws-new-end').value = '18';
  document.getElementById('ws-add-form').classList.add('hidden');
  updateWsNewPreview();
  loadWeekendSettings();
  openModal('worksettings-modal');
}

function renderWorkScheduleList() {
  var scheds = getWorkSchedules();
  var active = getWorkHours();
  var list = document.getElementById('ws-list');
  if (!scheds || scheds.length === 0) {
    list.innerHTML =
      '<div style="padding:12px;color:var(--text3);font-size:12px;text-align:center;">No schedules saved yet.</div>';
    return;
  }
  list.innerHTML = scheds
    .map(function (s) {
      var fmt = function (h) {
        return (h < 10 ? '0' : '') + h + ':00';
      };
      var isActive = s.id === active.id || (s.active && s.id === active.id);
      return (
        '<div class="ws-item' +
        (isActive ? ' ws-item-active' : '') +
        '">' +
        '<div class="ws-item-info">' +
        '<div class="ws-item-name">' +
        escHtml(s.name) +
        (isActive ? ' <span class="ws-badge">Active</span>' : '') +
        '</div>' +
        '<div class="ws-item-hours">' +
        fmt(s.start) +
        ' – ' +
        fmt(s.end) +
        ' (' +
        (s.end - s.start) +
        ' hrs)</div>' +
        '</div>' +
        '<div class="ws-item-actions">' +
        (!isActive
          ? '<button class="btn-secondary ws-btn" onclick="activateSchedule(\'' +
          s.id +
          '\')">Set Active</button>'
          : '') +
        '<button class="btn-secondary ws-btn" style="color:var(--danger);border-color:var(--danger);" onclick="deleteSchedule(\'' +
        s.id +
        '\')">Delete</button>' +
        '</div>' +
        '</div>'
      );
    })
    .join('');
}

async function activateSchedule(id) {
  try {
    await API.put(`/schedules/${id}/activate`);
    cache.workSchedules.forEach((s) => (s.active = s.id === id));
    renderWorkScheduleList();
    toast('Schedule activated.', 'success');
  } catch (err) {
    toast(err.message || 'Failed.', 'error');
  }
}

async function deleteSchedule(id) {
  const scheds = getWorkSchedules();
  if (scheds.length <= 1) {
    toast('Cannot delete the only schedule.', 'error');
    return;
  }
  if (getWorkHours().id === id) {
    toast('Cannot delete the active schedule.', 'error');
    return;
  }
  try {
    await API.del(`/schedules/${id}`);
    cache.workSchedules = cache.workSchedules.filter((s) => s.id !== id);
    renderWorkScheduleList();
    toast('Schedule deleted.', 'info');
  } catch (err) {
    toast(err.message || 'Failed.', 'error');
  }
}

function updateWsNewPreview() {
  var s = parseInt(document.getElementById('ws-new-start').value);
  var e = parseInt(document.getElementById('ws-new-end').value);
  var name = document.getElementById('ws-new-name').value || 'New Schedule';
  var fmt = function (h) {
    return (h < 10 ? '0' : '') + h + ':00';
  };
  var p = document.getElementById('ws-new-preview');
  if (p)
    p.textContent =
      name + ': ' + fmt(s) + ' – ' + fmt(e) + ' (' + (e - s) + ' hrs)';
}

async function saveNewSchedule() {
  var _btn = document.querySelector(
    '#worksettings-modal .btn-primary[onclick="saveNewSchedule()"]',
  );
  if (!_lockOp('saveSchedule', _btn, 'Saving…')) return;
  var name = document.getElementById('ws-new-name').value.trim();
  var startHour = parseInt(document.getElementById('ws-new-start').value);
  var endHour = parseInt(document.getElementById('ws-new-end').value);
  if (!name) {
    toast('Please enter a schedule name.', 'error');
    _unlockOp('saveSchedule', _btn);
    return;
  }
  if (endHour <= startHour) {
    toast('End hour must be after start hour.', 'error');
    _unlockOp('saveSchedule', _btn);
    return;
  }
  try {
    const newSched = await API.post('/schedules', { name, startHour, endHour });
    cache.workSchedules.push(newSched);
    document.getElementById('ws-new-name').value = '';
    document.getElementById('ws-add-form').classList.add('hidden');
    renderWorkScheduleList();
    toast('Schedule "' + name + '" saved!', 'success');
  } catch (err) {
    toast(err.message || 'Failed to save schedule.', 'error');
  } finally {
    _unlockOp('saveSchedule', _btn);
  }
}

// Legacy compat
function updateWsPreview() {
  updateWsNewPreview();
}
function saveWorkSettings() {
  saveNewSchedule();
}

/* ============================================================
   CHANGE PASSWORD (Admin)
   ============================================================ */
function openChangePassword() {
  const userSelect = document.getElementById('chpw-user');
  userSelect.innerHTML = '<option value="">Select user...</option>';
  getUsers()
    .filter((u) => u.id !== state.currentUser.id)
    .forEach((u) => {
      const opt = document.createElement('option');
      opt.value = u.id;
      opt.textContent = u.name + ' (@' + u.username + ') — ' + u.role;
      userSelect.appendChild(opt);
    });
  document.getElementById('chpw-new').value = '';
  document.getElementById('chpw-confirm').value = '';
  openModal('chpw-modal');
}

async function changeUserPassword() {
  const _btn = document.querySelector(
    '#chpw-modal .btn-primary[onclick="changeUserPassword()"]',
  );
  if (!_lockOp('chpwUser', _btn, 'Updating…')) return;
  const userId = document.getElementById('chpw-user').value;
  const newPw = document.getElementById('chpw-new').value.trim();
  const confirmPw = document.getElementById('chpw-confirm').value.trim();
  if (!userId) {
    toast('Please select a user.', 'error');
    _unlockOp('chpwUser', _btn);
    return;
  }
  if (!newPw || newPw.length < 4) {
    toast('Password must be at least 4 characters.', 'error');
    _unlockOp('chpwUser', _btn);
    return;
  }
  if (newPw !== confirmPw) {
    toast('Passwords do not match.', 'error');
    _unlockOp('chpwUser', _btn);
    return;
  }
  const user = getUsers().find((u) => u.id === userId);
  try {
    const result = await API.put(`/users/${userId}`, { password: newPw });
    if (!result)
      throw new Error('Password update failed - no response from server');
    addLog({
      taskId: null,
      taskTitle: 'Password changed for ' + (user ? user.name : userId),
      action: 'edited',
      actorName: state.currentUser.name,
      userId,
    });
    pushNotification(
      userId,
      '🔑 Password Changed',
      'Your account password was changed by ' + state.currentUser.name + '.',
      null,
    );
    closeModal('chpw-modal');
    toast(
      'Password updated' + (user ? ' for ' + user.name : '') + '!',
      'success',
    );
  } catch (err) {
    console.error('Password change error:', err);
    toast(
      err.message || 'Failed to update password. Please check your connection.',
      'error',
    );
  } finally {
    _unlockOp('chpwUser', _btn);
  }
}

function openChangePasswordFor(userId) {
  openChangePassword();
  setTimeout(function () {
    document.getElementById('chpw-user').value = userId;
  }, 50);
}

/* ============================================================
   MULTI-PERSONNEL TASK
   ============================================================ */
var mtDescItems = [];
var _mtScheduleMode = 'duration';

function setMtScheduleMode(mode) {
  _mtScheduleMode = mode;
  var activeStyle =
    'flex:1;padding:7px;border-radius:6px;font-size:11px;font-weight:700;border:1.5px solid var(--amber);background:var(--amber);color:var(--bg);cursor:pointer;transition:all 0.15s;';
  var inactiveStyle =
    'flex:1;padding:7px;border-radius:6px;font-size:11px;font-weight:700;border:1.5px solid var(--border);background:var(--bg3);color:var(--text2);cursor:pointer;transition:all 0.15s;';
  var durBtn = document.getElementById('mt-mode-duration');
  var endBtn = document.getElementById('mt-mode-enddate');
  var durGrp = document.getElementById('mt-duration-group');
  var endGrp = document.getElementById('mt-enddate-group');
  if (mode === 'duration') {
    durBtn.style.cssText = activeStyle;
    endBtn.style.cssText = inactiveStyle;
    durGrp.style.display = '';
    endGrp.style.display = 'none';
  } else {
    endBtn.style.cssText = activeStyle;
    durBtn.style.cssText = inactiveStyle;
    endGrp.style.display = '';
    durGrp.style.display = 'none';
  }
  updateMtDeadlinePreview();
}

function getMtDeadline() {
  var start = document.getElementById('mt-start').value;
  if (!start) return null;
  if (_mtScheduleMode === 'enddate') {
    var ed = document.getElementById('mt-enddate').value;
    return ed || null;
  }
  var raw = parseFloat(document.getElementById('mt-duration').value);
  var unit = document.getElementById('mt-duration-unit').value;
  if (!raw || raw <= 0) return null;
  var hours = unit === 'days' ? raw * 8 : raw;
  // Simple wall-clock: start + hours, converted to local datetime string
  var end = new Date(new Date(start).getTime() + hours * 3600000);
  return toLocalISO(end);
}

function updateMtDeadlinePreview() {
  var prev = document.getElementById('mt-deadline-preview');
  if (!prev) return;
  var dl = getMtDeadline();
  if (dl) {
    var d = new Date(dl);
    prev.textContent =
      '⏱ Ends: ' +
      d.toLocaleString('en-US', {
        weekday: 'short',
        month: 'short',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
      });
  } else {
    prev.textContent = '';
  }
}

function openMultiTask() {
  var role = state.currentUser ? state.currentUser.role : '';
  if (role !== 'manager' && role !== 'admin') {
    toast('Only managers can create multi-personnel tasks.', 'error');
    return;
  }
  mtDescItems = [];
  _mtScheduleMode = 'duration';
  document.getElementById('mt-title').value = '';
  document.getElementById('mt-requestor').value = state.currentUser?.name || '';
  document.getElementById('mt-priority').value = '3';
  var now = new Date();
  document.getElementById('mt-start').value = toLocalISO(now);
  document.getElementById('mt-duration').value = '1';
  document.getElementById('mt-duration-unit').value = 'hours';
  document.getElementById('mt-enddate').value = '';
  document.getElementById('multi-overlap-alert').classList.add('hidden');
  document.getElementById('multi-leave-alert').classList.add('hidden');
  if (document.getElementById('mt-search'))
    document.getElementById('mt-search').value = '';
  setMtScheduleMode('duration');
  renderMtDescItems();
  renderMtUserList('');
  updateMtSelectedCount();
  openModal('multi-task-modal');
}

// Store all users for filtering
function renderMtUserList(filter) {
  var container = document.getElementById('mt-users');
  var users = getUsers().filter(function (u) {
    return u.role === 'user' || u.role === 'manager';
  });
  var q = (filter || '').toLowerCase();
  var filtered = q
    ? users.filter(function (u) {
      return (
        u.name.toLowerCase().includes(q) ||
        u.username.toLowerCase().includes(q)
      );
    })
    : users;
  // Preserve checked state
  var checked = {};
  container.querySelectorAll('input[type="checkbox"]').forEach(function (cb) {
    if (cb.checked) checked[cb.value] = true;
  });
  container.innerHTML = filtered
    .map(function (u) {
      var isChecked = checked[u.id] ? 'checked' : '';
      return (
        '<label class="multi-user-option"><input type="checkbox" value="' +
        escHtml(u.id) +
        '" ' +
        isChecked +
        ' onchange="checkMultiTaskConflicts();updateMtSelectedCount()"><div><div class="multi-user-label">' +
        escHtml(u.name) +
        '</div><div class="multi-user-sub">@' +
        escHtml(u.username) +
        ' — ' +
        u.role +
        '</div></div></label>'
      );
    })
    .join('');
  if (filtered.length === 0)
    container.innerHTML =
      '<div style="padding:12px;text-align:center;color:var(--text3);font-size:12px;">No team members found</div>';
}

function filterMtUsers(val) {
  renderMtUserList(val);
  updateMtSelectedCount();
}

function selectAllMtUsers() {
  document
    .querySelectorAll('#mt-users input[type="checkbox"]')
    .forEach(function (cb) {
      cb.checked = true;
    });
  updateMtSelectedCount();
  checkMultiTaskConflicts();
}

function clearAllMtUsers() {
  document
    .querySelectorAll('#mt-users input[type="checkbox"]')
    .forEach(function (cb) {
      cb.checked = false;
    });
  updateMtSelectedCount();
  document.getElementById('multi-overlap-alert').classList.add('hidden');
  document.getElementById('multi-leave-alert').classList.add('hidden');
}

function updateMtSelectedCount() {
  var n = document.querySelectorAll(
    '#mt-users input[type="checkbox"]:checked',
  ).length;
  var el = document.getElementById('mt-selected-count');
  if (el) el.textContent = n + ' selected';
}

function openMultiTaskFor(userId) {
  openMultiTask();
  setTimeout(function () {
    var cb = document.querySelector('#mt-users input[value="' + userId + '"]');
    if (cb) {
      cb.checked = true;
    }
  }, 100);
}

function addMtDescItem() {
  var input = document.getElementById('mt-desc-input');
  var text = input.value.trim();
  if (!text) return;
  mtDescItems.push({ text: text, checked: false });
  input.value = '';
  renderMtDescItems();
}

function removeMtDescItem(idx) {
  mtDescItems.splice(idx, 1);
  renderMtDescItems();
}

function renderMtDescItems() {
  var container = document.getElementById('mt-desc-items');
  container.innerHTML = mtDescItems
    .map(function (item, i) {
      return (
        '<div class="desc-item"><span>' +
        escHtml(item.text) +
        '</span><button onclick="removeMtDescItem(' +
        i +
        ')">✕</button></div>'
      );
    })
    .join('');
}

function checkMultiTaskConflicts() {
  updateMtDeadlinePreview();
  var start = document.getElementById('mt-start').value;
  var deadline = getMtDeadline();
  if (!start || !deadline) return;
  var selectedIds = Array.from(
    document.querySelectorAll('#mt-users input:checked'),
  ).map(function (c) {
    return c.value;
  });
  if (selectedIds.length === 0) return;
  var overlapMsgs = [];
  var leaveMsgs = [];
  var allTasks = getTasks();
  var typeLabelsLeave = {
    lieu: 'Lieu Day',
    loa: 'Leave of Absence',
    awol: 'AWOL',
  };
  selectedIds.forEach(function (uid) {
    var user = getUsers().find(function (u) {
      return u.id === uid;
    });
    var userTasks = allTasks.filter(function (t) {
      return t.userId === uid && !t.cancelled && !t.done;
    });
    var tempTask = {
      id: 'temp',
      start: new Date(start).toISOString(),
      deadline: new Date(deadline).toISOString(),
    };
    var overlaps = checkTaskOverlap(tempTask, userTasks);
    if (overlaps.length > 0)
      overlapMsgs.push(
        (user ? user.name : uid) +
        ': conflicts with "' +
        overlaps
          .map(function (o) {
            return o.title;
          })
          .join('", "') +
        '"',
      );
    var leaveConflicts = checkLeaveConflict(uid, start, deadline);
    if (leaveConflicts.length > 0)
      leaveMsgs.push(
        (user ? user.name : uid) +
        ' is on ' +
        leaveConflicts
          .map(function (l) {
            return typeLabelsLeave[l.type];
          })
          .join(', '),
      );
  });
  var overlapAlert = document.getElementById('multi-overlap-alert');
  var leaveAlert = document.getElementById('multi-leave-alert');
  if (overlapMsgs.length > 0) {
    document.getElementById('multi-overlap-text').textContent =
      overlapMsgs.join(' | ');
    overlapAlert.classList.remove('hidden');
  } else overlapAlert.classList.add('hidden');
  if (leaveMsgs.length > 0) {
    document.getElementById('multi-leave-text').textContent =
      leaveMsgs.join(' | ');
    leaveAlert.classList.remove('hidden');
  } else leaveAlert.classList.add('hidden');
}

async function saveMultiTask() {
  var _btn = document.querySelector(
    '#multi-task-modal .btn-primary[onclick="saveMultiTask()"]',
  );
  if (!_lockOp('saveMultiTask', _btn, 'Assigning…')) return;
  var title = document.getElementById('mt-title').value.trim();
  var requestor = document.getElementById('mt-requestor').value.trim();
  var priority = parseInt(document.getElementById('mt-priority').value);
  var start = document.getElementById('mt-start').value;
  var deadline = getMtDeadline();
  if (!title || !requestor || !start || !deadline) {
    toast('Please fill all required fields.', 'error');
    _unlockOp('saveMultiTask', _btn);
    return;
  }
  if (new Date(start) >= new Date(deadline)) {
    toast('End time must be after start time.', 'error');
    _unlockOp('saveMultiTask', _btn);
    return;
  }

  var _mwh = getWorkHours();
  var _mStartHr =
    new Date(start).getHours() + new Date(start).getMinutes() / 60;
  var _mEndHr =
    new Date(deadline).getHours() + new Date(deadline).getMinutes() / 60;
  var _mPad = function (h) {
    return String(h).padStart(2, '0');
  };
  if (
    _mStartHr < _mwh.start ||
    _mEndHr > _mwh.end ||
    _mStartHr >= _mwh.end ||
    _mEndHr <= _mwh.start
  ) {
    toast(
      'Bookings must be within work hours: ' +
      _mPad(_mwh.start) +
      ':00 – ' +
      _mPad(_mwh.end) +
      ':00',
      'error',
    );
    _unlockOp('saveMultiTask', _btn);
    return;
  }

  var selectedIds = Array.from(
    document.querySelectorAll('#mt-users input:checked'),
  ).map(function (c) {
    return c.value;
  });
  if (selectedIds.length === 0) {
    toast('Select at least one user.', 'error');
    _unlockOp('saveMultiTask', _btn);
    return;
  }

  var tasks = getTasks();
  var assignedCount = 0;
  var skippedLeave = [];
  var groupId = 'mg-' + uid();
  var visibleDesc = mtDescItems.map(function (d) {
    return { text: d.text, checked: false };
  });
  var groupMeta = _makeGroupMeta(groupId, selectedIds.slice());
  var description = _buildDesc(visibleDesc, groupMeta);

  for (var i = 0; i < selectedIds.length; i++) {
    var uid_val = selectedIds[i];
    var user = getUsers().find(function (u) {
      return u.id === uid_val;
    });
    var leaveConflicts = checkLeaveConflict(uid_val, start, deadline);
    if (leaveConflicts.length > 0) {
      skippedLeave.push(user ? user.name : uid_val);
      continue;
    }
    var existingTasks = tasks.filter(function (t) {
      return t.userId === uid_val && !t.cancelled && !t.done;
    });
    var taskData = {
      title: title,
      requestor: requestor,
      priority: priority,
      start: new Date(start).toISOString(),
      deadline: new Date(deadline).toISOString(),
      description: description,
      isMultiPersonnel: true,
      multiGroupId: groupId,
    };
    var overlapCheck = resolveOverlapsNew(taskData, existingTasks, uid_val);
    if (overlapCheck.moved) {
      taskData.start = overlapCheck.newStart;
      taskData.deadline = overlapCheck.newDeadline;
      notifyManagerOfOverlap(
        Object.assign({}, taskData, { userId: uid_val }),
        overlapCheck.overlaps,
        true,
      );
    } else if (overlapCheck.overlaps && overlapCheck.overlaps.length > 0) {
      notifyManagerOfOverlap(
        Object.assign({}, taskData, { userId: uid_val }),
        overlapCheck.overlaps,
        false,
      );
    }
    try {
      var newTask = await API.post(
        '/tasks',
        Object.assign({}, taskData, { userId: uid_val }),
      );
      if (!cache.tasks) cache.tasks = [];
      if (newTask) cache.tasks.push(newTask);
      pushNotification(
        uid_val,
        '📌 Multi-Personnel Task Assigned',
        '"' +
        title +
        '" has been assigned to you by ' +
        state.currentUser.name +
        '.',
        newTask?.id,
      );
      if (newTask)
        addLog({
          taskId: newTask.id,
          taskTitle: newTask.title,
          action: 'added',
          actorName: state.currentUser.name,
          userId: uid_val,
        });
      assignedCount++;
    } catch (e) {
      console.error('Failed to assign to', uid_val, e);
    }
  }

  _unlockOp('saveMultiTask', _btn);
  closeModal('multi-task-modal');
  if (skippedLeave.length > 0)
    toast(
      'Task assigned to ' +
      assignedCount +
      ' user(s). Skipped: ' +
      skippedLeave.join(', ') +
      ' (on leave).',
      'warning',
    );
  else toast('Task assigned to ' + assignedCount + ' user(s)! ✓', 'success');
  if (state.view === 'calendar') renderCalendarDebounced();
  else if (state.view === 'user-list') renderUserList();
  else {
    var userId =
      state.view === 'user-tasks' ? state.targetUserId : state.currentUser.id;
    if (state.currentViewMode === 'timeline') renderTimeline(userId);
    else renderTasks(userId);
  }
}

/* ============================================================
   TABLE-BASED CALENDAR
   ============================================================ */
var calState = {
  year: new Date().getFullYear(),
  month: new Date().getMonth(),
  zoom: 'day',
};

/**
 * Debounced version of renderCalendar.
 * Use this wherever the calendar needs to refresh in response to
 * data changes (task save, cancel, reopen, leave update etc.)
 * so rapid successive calls are collapsed into one paint.
 */
var renderCalendarDebounced = debounce(function () {
  renderCalendar();
  if ($('gcal-view') && !$('gcal-view').classList.contains('hidden'))
    renderGcal();
}, CONFIG.DEBOUNCE_RENDER_MS);
var _calColWidth = 80; // px — default column width for zoom

function calZoom(dir) {
  // dir: +1 = zoom in (wider), -1 = zoom out (narrower)
  var steps = [40, 56, 80, 110, 150, 200];
  var idx = steps.indexOf(_calColWidth);
  if (idx === -1) idx = 2; // default to 80
  idx = Math.max(0, Math.min(steps.length - 1, idx + dir));
  _calColWidth = steps[idx];
  document.documentElement.style.setProperty(
    '--cal-col-w',
    _calColWidth + 'px',
  );
}

function calToggleZoom() {
  calState.zoom = calState.zoom === 'day' ? 'year' : 'day';
  var btn = document.getElementById('cal-zoom-btn');
  if (btn) btn.textContent = calState.zoom === 'year' ? '🔍 Year' : '🔍 Days';
  updateCalHeaderUI();
  renderCalendar();
}

function calNav(dir) {
  if (calState.zoom === 'year') {
    calState.year += dir;
  } else {
    calState.month += dir;
    if (calState.month > 11) {
      calState.month = 0;
      calState.year++;
    }
    if (calState.month < 0) {
      calState.month = 11;
      calState.year--;
    }
  }
  renderCalendar();
}

/**
 * calNavigateToTask — clicking a calendar task block navigates to that
 * user's task board and highlights / expands the specific task card.
 * • Admin/manager viewing team calendar → goes to user-tasks for that user.
 * • Personal calendar or worker → stays on own board.
 */
function calNavigateToTask(task) {
  const role = state.currentUser?.role;
  const isElevated = role === 'admin' || role === 'manager';
  const isOtherUser = isElevated && task.userId !== state.currentUser.id;

  function _scrollAndExpand(taskId) {
    // Wait a tick for renderTasks to finish, then scroll + expand
    setTimeout(function () {
      var card = document.querySelector('.task-card[data-id="' + taskId + '"]');
      if (card) {
        card.scrollIntoView({ behavior: 'smooth', block: 'center' });
        card.style.outline = '2px solid var(--amber)';
        setTimeout(function () {
          card.style.outline = '';
        }, 3000);
        // Auto-expand the card
        var expandSection = document.getElementById('expand-' + taskId);
        if (expandSection && !expandSection.classList.contains('open')) {
          expandSection.classList.add('open');
        }
      }
    }, 120);
  }

  if (isOtherUser) {
    // Navigate to that user's task board
    state.targetUserId = task.userId;
    state.view = 'user-tasks';
    _navRecord('user-tasks:' + task.userId);
    showView('user-tasks');
    _scrollAndExpand(task.id);
  } else {
    // Own tasks — just navigate to my task board
    state.view = 'my-tasks';
    showBoardView(state.currentUser.id);
    _scrollAndExpand(task.id);
  }
}

function calGoToday() {
  var now = new Date();
  calState.year = now.getFullYear();
  calState.month = now.getMonth();
  renderCalendar();
  // Scroll to today column after render
  setTimeout(function () {
    var todayCol = document.querySelector('#cal-dates-panel .today-col');
    if (todayCol) {
      todayCol.scrollIntoView({
        behavior: 'smooth',
        block: 'nearest',
        inline: 'center',
      });
    }
  }, 80);
}

/* ============================================================
   SWIPE GESTURES ON CALENDAR (mobile) — improved: only triggers
   when swipe starts on a non-scrollable area, not on the grid
   ============================================================ */
document.addEventListener('DOMContentLoaded', function () {
  (function initCalSwipe() {
    var _swipeStartX = 0;
    var _swipeStartY = 0;
    var _swipeOnStrip = false; // swipe started on week-strip (chip area)

    document.addEventListener(
      'touchstart',
      function (e) {
        var cv = document.getElementById('calendar-view');
        if (!cv || cv.classList.contains('hidden')) return;
        _swipeStartX = e.touches[0].clientX;
        _swipeStartY = e.touches[0].clientY;
        // Only allow month-swipe when touching the calendar header area or week strip
        var strip = document.getElementById('cal-week-strip');
        var header = document.querySelector('#calendar-view .calendar-header');
        _swipeOnStrip =
          (strip && strip.contains(e.target)) ||
          (header && header.contains(e.target));
      },
      { passive: true },
    );

    document.addEventListener(
      'touchend',
      function (e) {
        var cv = document.getElementById('calendar-view');
        if (!cv || cv.classList.contains('hidden')) return;
        if (!_swipeOnStrip) return; // only navigate from header / strip area
        var dx = e.changedTouches[0].clientX - _swipeStartX;
        var dy = e.changedTouches[0].clientY - _swipeStartY;
        if (Math.abs(dx) > 55 && Math.abs(dx) > Math.abs(dy) * 1.8) {
          if (dx < 0) calNav(1);
          else calNav(-1);
        }
      },
      { passive: true },
    );
  })();
});

/* ============================================================
   RICH TASK TOOLTIP SYSTEM (desktop hover + mobile tap hint)
   ============================================================ */
const TaskTooltip = (function () {
  const PCOLORS = {
    1: '#EF4444',
    2: '#FB923C',
    3: '#F59E0B',
    4: '#38BDF8',
    5: '#64748B',
  };
  const PLABELS = { 1: 'P1', 2: 'P2', 3: 'P3', 4: 'P4', 5: 'P5' };

  let _tip = null;
  let _hideTimer = null;
  let _currentTaskId = null;
  const isMobile = () => window.innerWidth <= 768 || 'ontouchstart' in window;

  function _buildHTML(task) {
    const desc = _visibleDesc(task.description || []);
    const total = desc.length;
    const checked = desc.filter((d) => d.checked).length;
    const pct = task.done
      ? 100
      : total > 0
        ? Math.round((checked / total) * 100)
        : 0;
    const pColor = PCOLORS[task.priority] || '#F59E0B';
    const pLabel = PLABELS[task.priority] || 'P3';

    const ts = new Date(task.start || task.createdAt);
    const te = new Date(task.deadline);
    const now = new Date();
    const isOverdue = !task.done && te < now;
    const isToday = te.toDateString() === now.toDateString();
    const daysLeft = Math.ceil((te - now) / 86400000);

    function fmt(d) {
      return (
        d.toLocaleDateString(undefined, { month: 'short', day: 'numeric' }) +
        ' ' +
        d.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' })
      );
    }

    const deadlineClass = isOverdue ? 'danger' : daysLeft <= 1 ? 'warn' : 'ok';
    const deadlineText = isOverdue
      ? 'OVERDUE'
      : isToday
        ? 'Today'
        : daysLeft + 'd left';

    // Checklist preview (up to 4 items)
    const previewItems = desc.slice(0, 4);
    const moreCount = Math.max(0, desc.length - 4);
    const checklistHtml =
      previewItems.length > 0
        ? `
      <div class="tt-divider"></div>
      <div class="tt-checklist">
        ${previewItems
          .map(
            (item) => `
          <div class="tt-check-item${item.checked ? ' done' : ''}">
            <div class="tt-check-dot${item.checked ? ' done' : ''}"></div>
            <span>${escHtml(item.text || item)}</span>
          </div>`,
          )
          .join('')}
        ${moreCount > 0 ? `<div class="tt-more-items">+${moreCount} more item${moreCount > 1 ? 's' : ''}</div>` : ''}
      </div>`
        : '';

    return `
      <style>#tf-task-tooltip { --tt-accent: ${pColor}; }</style>
      <div class="tt-header">
        <div class="tt-priority-badge" style="background:${pColor}">${pLabel}</div>
        <div class="tt-title">${escHtml(task.title)}</div>
        ${task.done ? '<div class="tt-done-badge">✓ Done</div>' : ''}
      </div>
      <div class="tt-divider"></div>
      <div class="tt-row">
        <span class="tt-label">Start</span>
        <span class="tt-value">${fmt(ts)}</span>
      </div>
      <div class="tt-row">
        <span class="tt-label">Deadline</span>
        <span class="tt-value ${deadlineClass}">${fmt(te)}</span>
      </div>
      <div class="tt-row">
        <span class="tt-label">Status</span>
        <span class="tt-value ${deadlineClass}">${task.done ? '✓ Completed' : task.cancelled ? '✕ Cancelled' : deadlineText}</span>
      </div>
      ${task.requestor ? `<div class="tt-row"><span class="tt-label">From</span><span class="tt-value">${escHtml(task.requestor)}</span></div>` : ''}
      ${total > 0
        ? `
      <div class="tt-divider"></div>
      <div class="tt-row">
        <span class="tt-label">Progress</span>
        <span class="tt-value">${checked}/${total} (${pct}%)</span>
      </div>
      <div class="tt-progress-wrap">
        <div class="tt-progress-bar" style="width:${pct}%;background:${task.done ? '#22C55E' : pColor}"></div>
      </div>`
        : ''
      }
      ${checklistHtml}
      ${task.isMultiPersonnel ? '<div class="tt-requestor">👥 Multi-personnel task</div>' : ''}`;
  }

  function show(task, anchorEl) {
    if (isMobile()) {
      _showMobile(task);
      return;
    }
    clearTimeout(_hideTimer);
    _currentTaskId = task.id;

    let tip = document.getElementById('tf-task-tooltip');
    if (!tip) return;
    _tip = tip;
    tip.innerHTML = _buildHTML(task);
    tip.classList.add('visible');

    // Position tooltip near the anchor element
    const rect = anchorEl.getBoundingClientRect();
    const TW = 310,
      TH = 220;
    let left = rect.right + 10;
    let top = rect.top - 10;

    if (left + TW > window.innerWidth - 12) left = rect.left - TW - 10;
    if (left < 12) left = 12;
    if (top + TH > window.innerHeight - 12) top = window.innerHeight - TH - 12;
    if (top < 12) top = 12;

    tip.style.left = left + 'px';
    tip.style.top = top + 'px';
  }

  function hide(delay) {
    if (isMobile()) {
      _hideMobile();
      return;
    }
    delay = delay || 0;
    clearTimeout(_hideTimer);
    _hideTimer = setTimeout(function () {
      const tip = document.getElementById('tf-task-tooltip');
      if (tip) tip.classList.remove('visible');
      _currentTaskId = null;
    }, delay);
  }

  function _showMobile(task) {
    const mob = document.getElementById('tf-task-tooltip-mobile');
    if (!mob) return;
    const pColor = PCOLORS[task.priority] || '#F59E0B';
    const desc = _visibleDesc(task.description || []);
    const total = desc.length;
    const checked = desc.filter((d) => d.checked).length;
    const pct = task.done
      ? 100
      : total > 0
        ? Math.round((checked / total) * 100)
        : 0;
    const te = new Date(task.deadline);
    const now = new Date();
    const isOverdue = !task.done && te < now;

    document.getElementById('ttm-title') &&
      (document.getElementById('ttm-title').textContent = task.title);
    document.getElementById('ttm-meta').textContent =
      'P' +
      task.priority +
      ' · ' +
      (task.done
        ? 'Done'
        : isOverdue
          ? 'OVERDUE'
          : te.toLocaleDateString(undefined, {
            month: 'short',
            day: 'numeric',
          }));
    const bar = document.getElementById('ttm-bar');
    bar.style.width = pct + '%';
    bar.style.background = task.done ? '#22C55E' : pColor;

    mob.style.display = '';
    mob.style.borderTopColor = pColor;
    requestAnimationFrame(() => mob.classList.add('visible'));
    clearTimeout(_hideTimer);
    _hideTimer = setTimeout(() => _hideMobile(), 3500);
  }

  function _hideMobile() {
    const mob = document.getElementById('tf-task-tooltip-mobile');
    if (!mob) return;
    mob.classList.remove('visible');
    setTimeout(() => {
      if (mob) mob.style.display = 'none';
    }, 200);
  }

  return { show, hide };
})();

/* Attach rich tooltip to all cal-block elements via event delegation */
document.addEventListener('DOMContentLoaded', function () {
  (function initCalBlockTooltips() {
    var _pendingShow = null;

    // Desktop: mouseenter/mouseleave
    document.addEventListener(
      'mouseenter',
      function (e) {
        var block = e.target.closest('.cal-block[data-taskid]');
        if (!block) return;
        var taskId = block.dataset.taskid;
        var task = (getTasks ? getTasks() : []).find(function (t) {
          return t.id === taskId;
        });
        if (!task) return;
        clearTimeout(_pendingShow);
        _pendingShow = setTimeout(function () {
          TaskTooltip.show(task, block);
        }, 180);
      },
      true,
    );

    document.addEventListener(
      'mouseleave',
      function (e) {
        var block = e.target.closest('.cal-block[data-taskid]');
        if (!block) return;
        clearTimeout(_pendingShow);
        TaskTooltip.hide(120);
      },
      true,
    );

    // Mobile: touchstart on block shows hint briefly, touchend opens task
    document.addEventListener(
      'touchstart',
      function (e) {
        var block = e.target.closest('.cal-block[data-taskid]');
        if (!block) return;
        var taskId = block.dataset.taskid;
        var task = (getTasks ? getTasks() : []).find(function (t) {
          return t.id === taskId;
        });
        if (!task) return;
        TaskTooltip.show(task, block);
      },
      { passive: true },
    );
  })();
});

/* ============================================================
   BUSYCAL DATE CHIP STRIP — build / rebuild on renderCalendar
   ============================================================ */
function buildCalWeekStrip(year, month, todayStr, users, allTasks, allLeaves) {
  var strip = document.getElementById('cal-week-strip');
  if (!strip) return;

  // Date chip strip removed — always hidden on all screen sizes
  strip.classList.add('hidden');
  strip.innerHTML = '';
  return;

  var daysInMonth = new Date(year, month + 1, 0).getDate();
  var dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  var pColors = {
    1: '#EF4444',
    2: '#FB923C',
    3: '#F59E0B',
    4: '#38BDF8',
    5: '#64748B',
  };

  strip.innerHTML = '';
  strip.classList.remove('hidden');

  // Month label chip at start
  var monthNames = [
    'Jan',
    'Feb',
    'Mar',
    'Apr',
    'May',
    'Jun',
    'Jul',
    'Aug',
    'Sep',
    'Oct',
    'Nov',
    'Dec',
  ];
  var monthLabel = document.createElement('div');
  monthLabel.style.cssText =
    'font-size:10px;font-weight:800;color:var(--amber);writing-mode:vertical-lr;transform:rotate(180deg);padding:0 6px 6px;letter-spacing:0.06em;text-transform:uppercase;flex-shrink:0;';
  monthLabel.textContent = monthNames[month];
  strip.appendChild(monthLabel);

  for (var d = 1; d <= daysInMonth; d++) {
    var date = new Date(year, month, d);
    var dateStr =
      year +
      '-' +
      String(month + 1).padStart(2, '0') +
      '-' +
      String(d).padStart(2, '0');
    var dow = date.getDay();
    var isToday = dateStr === todayStr;
    var isWeekend = dow === 0 || dow === 6;

    // Collect dots for this day
    var dots = [];
    if (allTasks) {
      allTasks.forEach(function (t) {
        var ts = new Date(t.start || t.createdAt);
        var te = new Date(t.deadline);
        var ds = new Date(date);
        ds.setHours(0, 0, 0, 0);
        var de = new Date(date);
        de.setHours(23, 59, 59, 999);
        if (ts <= de && te >= ds && dots.length < 3) {
          dots.push(pColors[t.priority] || pColors['3']);
        }
      });
    }
    if (allLeaves) {
      allLeaves.forEach(function (l) {
        var sd = l.startDate || '';
        var ed = l.endDate || '';
        if (sd <= dateStr && ed >= dateStr && dots.length < 3) {
          var c =
            l.type === 'lieu'
              ? '#10B981'
              : l.type === 'loa'
                ? '#3B82F6'
                : '#EF4444';
          dots.push(c);
        }
      });
    }

    var chip = document.createElement('div');
    chip.className =
      'cal-week-chip' +
      (isToday ? ' today-chip' : '') +
      (isWeekend ? ' weekend-chip' : '');
    chip.dataset.dateStr = dateStr;

    var dayEl = document.createElement('div');
    dayEl.className = 'cal-week-chip-day';
    dayEl.textContent = dayNames[dow].slice(0, 2);

    var numWrap = document.createElement('div');
    if (isToday) {
      var circle = document.createElement('div');
      circle.className = 'cal-week-chip-num-circle';
      circle.textContent = d;
      numWrap.appendChild(circle);
    } else {
      numWrap.className = 'cal-week-chip-num';
      numWrap.textContent = d;
    }

    var dotsEl = document.createElement('div');
    dotsEl.className = 'cal-week-chip-dots';
    dots.slice(0, 3).forEach(function (c) {
      var dot = document.createElement('div');
      dot.className = 'cal-week-chip-dot';
      dot.style.background = c;
      dotsEl.appendChild(dot);
    });

    chip.appendChild(dayEl);
    chip.appendChild(numWrap);
    chip.appendChild(dotsEl);

    // Tap: scroll that date column into view
    chip.addEventListener(
      'click',
      (function (ds) {
        return function () {
          var col = document.querySelector(
            '#cal-dates-panel th[data-date="' + ds + '"]',
          );
          if (col)
            col.scrollIntoView({
              behavior: 'smooth',
              block: 'nearest',
              inline: 'center',
            });
        };
      })(dateStr),
    );

    strip.appendChild(chip);
  }

  // Scroll today's chip into center
  requestAnimationFrame(function () {
    var todayChip = strip.querySelector('.today-chip');
    if (todayChip)
      todayChip.scrollIntoView({
        behavior: 'smooth',
        block: 'nearest',
        inline: 'center',
      });
  });
}

/* ============================================================
   CALENDAR USER TOOLTIP
   ============================================================ */
let _calTooltipEl = null;
function showCalUserTooltip(user, tdEl) {
  hideCalUserTooltip();
  const year = calState.year;
  const month = calState.month;
  const leaves = getLeaves().filter((l) => l.userId === user.id);

  // Count working days this month (Mon-Fri, minus AWOL)
  const workDays = countMonthlyWorkingDays(user.id, year, month);

  // Lieu days balance
  const lieuDays = countLieuDays(user.id);
  const lieuEarned = countLieuDaysEarned(user.id);
  const lieuUsed = countLieuDaysUsed(user.id);

  // LOA balance
  const loaUsed = countUsedLeaveDays(user.id);
  const loaRemaining = remainingLeaveDays(user.id);

  // AWOL count this month
  const monthStr = year + '-' + String(month + 1).padStart(2, '0');
  const awolDays = leaves
    .filter(
      (l) =>
        l.type === 'awol' && l.startDate && l.startDate.startsWith(monthStr),
    )
    .reduce((sum, l) => sum + countWorkingDays(l.startDate, l.endDate), 0);

  const s = getCompanySettings();
  const weekendNote =
    s.satWorking || s.sunWorking
      ? '<div style="margin-top:4px;font-size:10px;color:#F59E0B;">Weekend working: ' +
      [s.satWorking ? 'Sat' : '', s.sunWorking ? 'Sun' : '']
        .filter(Boolean)
        .join('+') +
      '</div>'
      : '';

  const html = `
    <div style="font-family:monospace;font-size:12px;">
      <div style="font-weight:800;font-size:13px;color:#F59E0B;margin-bottom:10px;border-bottom:1px solid #2A2D38;padding-bottom:6px;">${escHtml(user.name)}</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px 16px;">
        <div style="color:#9CA3AF;font-size:10px;text-transform:uppercase;letter-spacing:.05em;">Working Days</div>
        <div style="font-weight:700;color:#E8EAF0;">${workDays}</div>
        <div style="color:#9CA3AF;font-size:10px;text-transform:uppercase;letter-spacing:.05em;">Lieu Earned</div>
        <div style="font-weight:700;color:#10B981;">${lieuEarned}</div>
        <div style="color:#9CA3AF;font-size:10px;text-transform:uppercase;letter-spacing:.05em;">Lieu Used</div>
        <div style="font-weight:700;color:#6B7280;">${lieuUsed}</div>
        <div style="color:#9CA3AF;font-size:10px;text-transform:uppercase;letter-spacing:.05em;">Lieu Balance</div>
        <div style="font-weight:700;color:${lieuDays > 0 ? '#10B981' : '#EF4444'};">${lieuDays}</div>
        <div style="color:#9CA3AF;font-size:10px;text-transform:uppercase;letter-spacing:.05em;">Leave Used</div>
        <div style="font-weight:700;color:#3B82F6;">${loaUsed} / ${ANNUAL_LEAVE_DAYS}</div>
        <div style="color:#9CA3AF;font-size:10px;text-transform:uppercase;letter-spacing:.05em;">Leave Left</div>
        <div style="font-weight:700;color:${loaRemaining <= 5 ? '#EF4444' : loaRemaining <= 10 ? '#FB923C' : '#22C55E'};">${loaRemaining} days</div>
        ${awolDays > 0 ? `<div style="color:#9CA3AF;font-size:10px;text-transform:uppercase;letter-spacing:.05em;">AWOL This Mo.</div><div style="font-weight:700;color:#EF4444;">${awolDays} day${awolDays !== 1 ? 's' : ''}</div>` : ''}
      </div>
      ${weekendNote}
    </div>`;

  const tip = document.createElement('div');
  tip.id = 'cal-user-tooltip';
  tip.style.cssText = [
    'position:fixed',
    'z-index:9999',
    'background:#0F1117',
    'border:1px solid #2A2D38',
    'border-radius:10px',
    'padding:14px 16px',
    'pointer-events:none',
    'box-shadow:0 8px 32px rgba(0,0,0,0.6)',
    'min-width:200px',
    'max-width:260px',
    'transition:opacity .15s ease',
    'opacity:0',
  ].join(';');
  tip.innerHTML = html;
  document.body.appendChild(tip);
  _calTooltipEl = tip;

  // Position: to the right of the cell
  const rect = tdEl.getBoundingClientRect();
  let left = rect.right + 8;
  let top = rect.top;
  // Keep within viewport
  const tw = 260;
  const th = 180;
  if (left + tw > window.innerWidth - 10) left = rect.left - tw - 8;
  if (top + th > window.innerHeight - 10) top = window.innerHeight - th - 10;
  tip.style.left = left + 'px';
  tip.style.top = Math.max(10, top) + 'px';
  requestAnimationFrame(() => {
    tip.style.opacity = '1';
  });
}
function hideCalUserTooltip() {
  if (_calTooltipEl) {
    _calTooltipEl.remove();
    _calTooltipEl = null;
  }
}

/* ============================================================
   CALENDAR YEAR VIEW — iPhone-style mini month grids
   ============================================================ */
function renderCalendarYearView() {
  var year = calState.year;
  document.getElementById('cal-month-label').textContent = String(year);
  var table = document.getElementById('cal-table');
  table.style.display = 'none';
  var splitWrapper = document.getElementById('cal-split-wrapper');
  if (splitWrapper) splitWrapper.style.display = 'none';
  var container = document.getElementById('calendar-container');
  var old = document.getElementById('cal-year-grid');
  if (old) old.remove();

  var today = new Date();
  var allTasks = getTasks();
  var allLeaves = getLeaves();
  var isPersonal =
    state.calendarMode === 'personal' ||
    (state.currentUser.role === 'user' && state.calendarMode !== 'team');
  var viewUserId = isPersonal ? state.currentUser.id : null;
  var monthNames = [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December',
  ];
  var dayLetters = ['S', 'M', 'T', 'W', 'T', 'F', 'S'];
  var pColors = {
    1: '#EF4444',
    2: '#FB923C',
    3: '#FBBF24',
    4: '#60A5FA',
    5: '#9CA3AF',
  };
  var leaveColors = { lieu: '#10B981', loa: '#3B82F6', awol: '#EF4444' };

  var grid = document.createElement('div');
  grid.id = 'cal-year-grid';
  grid.style.cssText =
    'display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:16px;padding:16px;overflow-y:auto;';

  for (var m = 0; m < 12; m++) {
    var isCurrentMonth = today.getFullYear() === year && today.getMonth() === m;
    var daysInMonth = new Date(year, m + 1, 0).getDate();
    var firstDow = new Date(year, m, 1).getDay();

    var card = document.createElement('div');
    card.style.cssText =
      'background:var(--bg3);border:1px solid ' +
      (isCurrentMonth ? 'var(--amber)' : 'var(--border)') +
      ';border-radius:12px;padding:12px;cursor:pointer;transition:all .15s;';
    card.onmouseenter = function () {
      this.style.transform = 'translateY(-2px)';
      this.style.boxShadow = '0 6px 24px rgba(0,0,0,0.3)';
    };
    card.onmouseleave = function () {
      this.style.transform = '';
      this.style.boxShadow = '';
    };
    (function (mo) {
      card.onclick = function () {
        calState.month = mo;
        calState.zoom = 'day';
        var btn = document.getElementById('cal-zoom-btn');
        if (btn) btn.textContent = '🔍 Days';
        renderCalendar();
      };
    })(m);

    var title = document.createElement('div');
    title.style.cssText =
      'font-size:13px;font-weight:800;color:' +
      (isCurrentMonth ? 'var(--amber)' : 'var(--text)') +
      ';margin-bottom:8px;';
    title.textContent = monthNames[m];
    card.appendChild(title);

    var dowRow = document.createElement('div');
    dowRow.style.cssText =
      'display:grid;grid-template-columns:repeat(7,1fr);gap:1px;margin-bottom:4px;';
    dayLetters.forEach(function (l) {
      var h = document.createElement('div');
      h.style.cssText =
        'font-size:8px;color:var(--text3);text-align:center;font-weight:700;';
      h.textContent = l;
      dowRow.appendChild(h);
    });
    card.appendChild(dowRow);

    var daysGrid = document.createElement('div');
    daysGrid.style.cssText =
      'display:grid;grid-template-columns:repeat(7,1fr);gap:1px;';
    for (var e = 0; e < firstDow; e++)
      daysGrid.appendChild(document.createElement('div'));

    for (var d = 1; d <= daysInMonth; d++) {
      (function (day, mo2) {
        var dayDate = new Date(year, mo2, day);
        var dayStr = toDateStr(dayDate);
        var isToday = dayDate.toDateString() === today.toDateString();
        var isWknd = dayDate.getDay() === 0 || dayDate.getDay() === 6;
        var dots = [];

        if (viewUserId) {
          allLeaves
            .filter(function (l) {
              return (
                l.userId === viewUserId &&
                l.status !== 'denied' &&
                l.startDate <= dayStr &&
                l.endDate >= dayStr
              );
            })
            .forEach(function (l) {
              if (l.status === 'pending') {
                dots.push('var(--text3)');
              } else {
                dots.push(leaveColors[l.type] || '#888');
              }
            });
        } else {
          if (
            allLeaves.some(function (l) {
              return (
                l.status !== 'denied' &&
                l.startDate <= dayStr &&
                l.endDate >= dayStr
              );
            })
          ) {
            // If any are leaves, show a color. We'll stick to a mixed indicator
            // or just green if any is approved, grey if all are pending.
            const dayLeaves = allLeaves.filter(
              (l) =>
                l.status !== 'denied' &&
                l.startDate <= dayStr &&
                l.endDate >= dayStr,
            );
            const hasApproved = dayLeaves.some(
              (l) => l.status !== 'pending' && l.status !== 'denied',
            );
            dots.push(hasApproved ? '#10B981' : '#9CA3AF');
          }
        }

        var cell = document.createElement('div');
        cell.style.cssText =
          'display:flex;flex-direction:column;align-items:center;padding:1px 0;' +
          (isWknd && !isToday ? 'opacity:0.45;' : '');
        if (isToday) {
          cell.style.cssText += 'background:var(--amber);border-radius:50%;';
        }
        var num = document.createElement('div');
        num.style.cssText =
          'font-size:9px;font-weight:' +
          (isToday ? '800' : '400') +
          ';color:' +
          (isToday ? 'var(--bg)' : 'var(--text2)') +
          ';width:16px;height:14px;display:flex;align-items:center;justify-content:center;';
        num.textContent = day;
        cell.appendChild(num);
        if (dots.length > 0) {
          var dr = document.createElement('div');
          dr.style.cssText =
            'display:flex;gap:1px;justify-content:center;margin-top:1px;';
          dots.slice(0, 3).forEach(function (c) {
            var dot = document.createElement('div');
            dot.style.cssText =
              'width:3px;height:3px;border-radius:50%;background:' +
              c +
              ';flex-shrink:0;';
            dr.appendChild(dot);
          });
          cell.appendChild(dr);
        }
        daysGrid.appendChild(cell);
      })(d, m);
    }
    card.appendChild(daysGrid);
    grid.appendChild(card);
  }
  container.appendChild(grid);
}

/**
 * Render the calendar view (day/week/month/year).
 * Builds the full calendar grid with task blocks, leave indicators,
 * workday markers, sticky headers, and the BusyCal-style date chip strip.
 */
function renderCalendar() {
  var old = $('cal-year-grid');
  var legacyTable = $('cal-table');
  updateCalHeaderUI();
  if (calState.zoom === 'year') {
    var splitW = $('cal-split-wrapper');
    if (splitW) splitW.style.display = 'none';
    if (legacyTable) legacyTable.style.display = 'none';
    var yearStrip = $('cal-week-strip');
    if (yearStrip) yearStrip.classList.add('hidden');
    renderCalendarYearView();
    return;
  }
  if (old) old.remove();
  // Show split wrapper, hide legacy table
  var splitWrapper = $('cal-split-wrapper');
  if (splitWrapper) splitWrapper.style.display = 'flex';
  if (legacyTable) legacyTable.style.display = 'none';

  var isPersonal = state.calendarMode === 'personal';
  var isElevated =
    state.currentUser.role === 'admin' || state.currentUser.role === 'manager';

  var year = calState.year;
  var month = calState.month;
  var monthNames = [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December',
  ];
  if (EL.calMonthLabel)
    EL.calMonthLabel.textContent = monthNames[month] + ' ' + year;

  var namesTable = EL.calNamesTable || $('cal-names-table');
  var datesTable = EL.calDatesTable || $('cal-dates-table');
  namesTable.innerHTML = '';
  datesTable.innerHTML = '';

  var daysInMonth = new Date(year, month + 1, 0).getDate();
  var today = new Date();
  var days = [];
  for (var d = 1; d <= daysInMonth; d++) days.push(new Date(year, month, d));

  // ── Names panel header ──
  var namesThead = document.createElement('thead');
  var namesTh = document.createElement('tr');
  var memberTh = document.createElement('th');
  memberTh.textContent = 'Team Members';
  namesTh.appendChild(memberTh);
  namesThead.appendChild(namesTh);
  namesTable.appendChild(namesThead);

  // ── Dates panel header ──
  var datesThead = document.createElement('thead');
  var datesHeaderRow = document.createElement('tr');
  var customWorkdays = getCustomWorkdays();
  var todayStr = toDateStr(today);
  var dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  days.forEach(function (day) {
    var th = document.createElement('th');
    var isToday = day.toDateString() === today.toDateString();
    var isWeekend = day.getDay() === 0 || day.getDay() === 6;
    var dayStr = toDateStr(day);
    var isCustomWorkday = customWorkdays.includes(dayStr);
    var isMobile = window.innerWidth <= 768;
    th.innerHTML =
      '<span class="cal-date-day">' +
      dayNames[day.getDay()] +
      '</span><span class="cal-date-num">' +
      day.getDate() +
      '</span>';
    th.style.minWidth = isMobile ? 'var(--cal-col-w, 46px)' : '80px';
    th.dataset.date = dayStr; /* for strip chip tap-scroll */
    if (isToday) th.classList.add('today-col');
    else if (isWeekend) th.classList.add('weekend-col');
    if (isCustomWorkday) th.classList.add('custom-workday-col');
    datesHeaderRow.appendChild(th);
  });
  datesThead.appendChild(datesHeaderRow);
  datesTable.appendChild(datesThead);

  var allTasks = getTasks();
  var allLeaves = getLeaves();
  var allUsers = getUsers().filter(function (u) {
    return u.role !== 'admin';
  });
  var users = isPersonal
    ? allUsers.filter(function (u) {
      return u.id === state.currentUser.id;
    })
    : allUsers;
  if (!isPersonal && calSearchFilter) {
    users = users.filter(function (u) {
      return (
        u.name.toLowerCase().includes(calSearchFilter) ||
        u.username.toLowerCase().includes(calSearchFilter)
      );
    });
  }

  var namesTbody = document.createElement('tbody');
  var datesTbody = document.createElement('tbody');
  // Fragments for batched DOM writes — avoids per-row reflow
  var namesFrag = document.createDocumentFragment();
  var datesFrag = document.createDocumentFragment();
  var roleColors = { manager: '#A78BFA', user: 'var(--p4)' };
  var typeLabels = { lieu: '🟢 Lieu', loa: '🔵 LOA', awol: '🔴 AWOL' };

  if (users.length === 0) {
    var emptyTr = document.createElement('tr');
    var emptyTd = document.createElement('td');
    emptyTd.style.cssText =
      'text-align:center;padding:60px;color:var(--text3);font-size:13px;';
    emptyTd.textContent = 'No team members registered yet.';
    emptyTr.appendChild(emptyTd);
    namesTbody.appendChild(emptyTr);
  }

  users.forEach(function (user) {
    // ── Names panel row ──
    var namesTr = document.createElement('tr');
    namesTr.dataset.userId = user.id;
    var userTd = document.createElement('td');
    var rc = roleColors[user.role] || 'var(--p4)';
    userTd.innerHTML =
      '<div class="cal-user-label" style="cursor:pointer;" title="View ' +
      escHtml(user.name) +
      '\'s tasks"><div class="cal-user-name">' +
      escHtml(user.name) +
      '</div><div class="cal-user-role" style="color:' +
      rc +
      '">@' +
      escHtml(user.username) +
      '</div>' +
      (isElevated
        ? '<div class="cal-lieu-btns"><button class="cal-lieu-btn" onclick="event.stopPropagation();openAddLieuDay(\'' +
        user.id +
        "','" +
        escHtml(user.name) +
        '\')">+ Lieu</button><button class="cal-lieu-btn cal-lieu-btn-minus" onclick="event.stopPropagation();removeLieuDay(\'' +
        user.id +
        "','" +
        escHtml(user.name) +
        '\')">− Lieu</button></div>'
        : '') +
      '</div>';
    userTd.onclick = (function (uid) {
      return function () {
        viewUserTasks(uid);
      };
    })(user.id);
    userTd.style.position = 'relative';
    userTd.addEventListener(
      'mouseenter',
      (function (u) {
        return function (e) {
          showCalUserTooltip(u, e.currentTarget);
        };
      })(user),
    );
    userTd.addEventListener('mouseleave', hideCalUserTooltip);
    namesTr.appendChild(userTd);
    namesTbody.appendChild(namesTr);

    // ── Dates panel row ──
    var datesTr = document.createElement('tr');
    var userLeaves = allLeaves.filter(function (l) {
      return l.userId === user.id;
    });

    // ── Pre-compute leave slot assignments (each unique leave period = one row) ──
    // Merge consecutive/overlapping leaves of the same type into single bars
    var rawSlottedLeaves = userLeaves
      .filter(function (l) {
        var sd = l.startDate || (l.start ? l.start.slice(0, 10) : '');
        var ed = l.endDate || (l.end ? l.end.slice(0, 10) : '');
        return (
          sd <= toDateStr(days[days.length - 1]) && ed >= toDateStr(days[0])
        );
      })
      .sort(function (a, b) {
        var sa = a.startDate || (a.start ? a.start.slice(0, 10) : '');
        var sb = b.startDate || (b.start ? b.start.slice(0, 10) : '');
        return sa < sb ? -1 : sa > sb ? 1 : 0;
      });
    // Merge consecutive leaves of same type into virtual bars
    var slottedLeaves = (function (leaves) {
      var merged = [];
      leaves.forEach(function (l) {
        var sd = l.startDate || (l.start ? l.start.slice(0, 10) : '');
        var ed = l.endDate || (l.end ? l.end.slice(0, 10) : '');
        if (!merged.length) {
          merged.push({
            id: l.id,
            userId: l.userId,
            type: l.type,
            startDate: sd,
            endDate: ed,
            reason: l.reason,
            _ids: [l.id],
          });
          return;
        }
        var last = merged[merged.length - 1];
        // Check if same type and consecutive (adjacent or overlapping dates)
        if (last.type === l.type) {
          var nextDay = (function (d) {
            var dt = new Date(d + 'T12:00:00');
            dt.setDate(dt.getDate() + 1);
            return dt.toISOString().slice(0, 10);
          })(last.endDate);
          if (sd <= nextDay) {
            // Merge: extend end date
            if (ed > last.endDate) last.endDate = ed;
            last._ids.push(l.id);
            return;
          }
        }
        merged.push({
          id: l.id,
          userId: l.userId,
          type: l.type,
          startDate: sd,
          endDate: ed,
          reason: l.reason,
          _ids: [l.id],
        });
      });
      return merged;
    })(rawSlottedLeaves);

    // slotMap persists across days: tracks which slot index each spanning item owns.
    // Key: 'L_<id>' for leaves, 'T_<id>' for tasks. Value: slot index (integer).
    var slotMap = {};

    days.forEach(function (day) {
      var td = document.createElement('td');
      var isToday = day.toDateString() === today.toDateString();
      var isWeekend = day.getDay() === 0 || day.getDay() === 6;
      var dayStr = toDateStr(day);
      td.dataset.dateStr = dayStr;
      td.dataset.userId = user.id;
      td.dataset.userName = user.name;
      var isPast = dayStr < todayStr;
      var userWorkdays = getUserCustomWorkdays(user.id);
      var isUserWorkday = userWorkdays.includes(dayStr);
      var isAutoWorkday = isPast && !isWeekend;
      if (isToday) {
        td.classList.add('today-col');
      } else if (isWeekend) {
        if (isUserWorkday) {
          td.classList.add('day-off');
          td.classList.add('user-workday');
        } else td.classList.add('day-off');
      } else if (isUserWorkday || isAutoWorkday) {
        var hasLeave = userLeaves.some(function (leave) {
          var sd =
            leave.startDate || (leave.start ? leave.start.slice(0, 10) : '');
          var ed = leave.endDate || (leave.end ? leave.end.slice(0, 10) : '');
          return (
            sd <= dayStr &&
            ed >= dayStr &&
            (leave.type === 'lieu' || leave.type === 'loa')
          );
        });
        if (!hasLeave) {
          td.classList.add('past-workday');
        }
      }
      var dayStart = new Date(day);
      dayStart.setHours(0, 0, 0, 0);
      var dayEnd = new Date(day);
      dayEnd.setHours(23, 59, 59, 999);
      var cell = document.createElement('div');
      cell.className = 'cal-cell-content';

      // ── Unified slot rendering: leaves then tasks ──
      // Rules:
      //   1. Spanning items lock the same slot index on every day they cover (slotMap).
      //   2. Single-day / newly-starting items fill the topmost free slot each day.
      //   3. On collision days (a spanning item ENDS the same day another STARTS),
      //      the ending item keeps its locked slot and the starting item goes below it.

      var pColors = {
        1: 'var(--p1)',
        2: 'var(--p2)',
        3: 'var(--p3)',
        4: 'var(--p4)',
        5: 'var(--p5)',
      };

      // Build a flat ordered list of all items (leaves first, then tasks).
      // Each entry carries enough metadata so slot assignment works uniformly.
      var allItems = [];
      slottedLeaves.forEach(function (lv) {
        var sd = lv.startDate || (lv.start ? lv.start.slice(0, 10) : '');
        var ed = lv.endDate || (lv.end ? lv.end.slice(0, 10) : '');
        allItems.push({
          kind: 'leave',
          item: lv,
          key: 'L_' + lv.id,
          startStr: sd,
          endStr: ed,
          isSpanning: sd !== ed,
        });
      });

      // Detect collision: any item ending today that started before today,
      // AND any item starting today.
      var hasEndingToday = allItems.some(function (e) {
        return e.endStr === dayStr && e.startStr < dayStr;
      });
      var hasStartingToday = allItems.some(function (e) {
        return e.startStr === dayStr;
      });
      var collisionDay = hasEndingToday && hasStartingToday;

      // On collision days, sort so items that started before today come first
      // (they keep their locked slot), items starting today come after.
      if (collisionDay) {
        allItems.sort(function (a, b) {
          var aStartsToday = a.startStr === dayStr;
          var bStartsToday = b.startStr === dayStr;
          if (!aStartsToday && bStartsToday) return -1;
          if (aStartsToday && !bStartsToday) return 1;
          return 0;
        });
      }

      // ── Slot assignment for this day ──
      // Pass 1: mark slots already locked by spanning items.
      var usedSlots = {}; // slot# → key
      allItems.forEach(function (e) {
        if (slotMap[e.key] !== undefined) usedSlots[slotMap[e.key]] = e.key;
      });

      // Pass 2: assign new slots to items active for the first time today.
      allItems.forEach(function (e) {
        if (slotMap[e.key] !== undefined) return; // already locked
        // Is this item active today?
        var activeNow;
        if (e.kind === 'leave') {
          activeNow = e.startStr <= dayStr && e.endStr >= dayStr;
        } else {
          activeNow = e.ts <= dayEnd && e.te >= dayStart;
        }
        if (!activeNow) return;

        // Find the first free slot
        var s = 0;
        while (usedSlots[s] !== undefined) s++;
        usedSlots[s] = e.key;
        e._slot = s;

        // Lock into slotMap only if it spans beyond today
        if (e.isSpanning) {
          slotMap[e.key] = s;
        }
      });

      // Release slots for items that fully ended before today.
      Object.keys(slotMap).forEach(function (key) {
        var e = allItems.find(function (x) {
          return x.key === key;
        });
        if (e && e.endStr < dayStr) delete slotMap[key];
      });

      // Determine max slot used today.
      var maxSlot = -1;
      allItems.forEach(function (e) {
        var s = slotMap[e.key] !== undefined ? slotMap[e.key] : e._slot;
        if (s !== undefined && s > maxSlot) maxSlot = s;
      });

      // Build slot → entry map for rendering.
      var slotRender = {};
      allItems.forEach(function (e) {
        var s = slotMap[e.key] !== undefined ? slotMap[e.key] : e._slot;
        if (s !== undefined) slotRender[s] = e;
      });

      // ── Render slots 0..maxSlot ──
      for (var si = 0; si <= maxSlot; si++) {
        var entry = slotRender[si];

        // Gap spacer
        if (!entry) {
          var gsp = document.createElement('div');
          gsp.className = 'cal-slot-spacer';
          cell.appendChild(gsp);
          continue;
        }

        if (entry.kind === 'leave') {
          var leave = entry.item;
          var sd = entry.startStr;
          var ed = entry.endStr;
          var activeOnDay = sd <= dayStr && ed >= dayStr;

          if (
            !activeOnDay ||
            leave.status === 'denied' ||
            !isWorkingDay(new Date(dayStr + 'T12:00:00'))
          ) {
            var sp = document.createElement('div');
            sp.className = 'cal-slot-spacer';
            cell.appendChild(sp);
            continue;
          }

          var isLvStart = sd === dayStr;
          var isLvEnd = ed === dayStr;
          var spansMulti = sd !== ed;
          var spanCls = '';
          if (spansMulti) {
            if (isLvStart && !isLvEnd) spanCls = ' span-start';
            else if (!isLvStart && isLvEnd) spanCls = ' span-end';
            else if (!isLvStart && !isLvEnd) spanCls = ' span-middle';
          }

          var slotEl = document.createElement('div');
          slotEl.className = 'cal-leave-slot';
          var bar = document.createElement('div');
          bar.className = 'cal-leave-bar leave-' + leave.type + spanCls;
          bar.title =
            (typeLabels[leave.type] || leave.type) +
            (leave.reason ? ': ' + leave.reason : '');
          var lbl = document.createElement('span');
          lbl.className = 'cal-leave-day-label';
          lbl.textContent =
            isLvStart || dayStr === toDateStr(days[0])
              ? typeLabels[leave.type] || leave.type
              : typeLabels[leave.type] || leave.type;

          if (leave.status === 'pending') {
            const isElevated =
              state.currentUser.role === 'admin' ||
              state.currentUser.role === 'manager';

            bar.style.background = 'var(--bg4)';
            bar.style.border = '1px solid var(--border2)';
            bar.style.color = 'var(--text3)';
            bar.style.opacity = '0.7';
            bar.title = isElevated
              ? '⏳ Pending approval — click to review'
              : '⏳ Awaiting manager approval';
            if (leave.reason) bar.title += ': ' + leave.reason;
            lbl.textContent = '⏳ ' + lbl.textContent;

            if (isElevated) {
              bar.style.cursor = 'pointer';
              bar.onclick = function (e) {
                e.stopPropagation();
                openLeaveRequests(leave.id);
              };
            }
          }

          bar.appendChild(lbl);
          if (isElevated) {
            var delB = document.createElement('button');
            delB.className = 'cal-leave-del';
            delB.textContent = '×';
            delB.title = 'Remove ' + dayStr + ' from this leave';
            delB.onclick = (function (lids, dStr) {
              return function (e) {
                e.stopPropagation();
                var actualLeave = getLeaves().filter(function (l) {
                  var lsd =
                    l.startDate || (l.start ? l.start.slice(0, 10) : '');
                  var led = l.endDate || (l.end ? l.end.slice(0, 10) : '');
                  return (
                    l.userId === leave.userId &&
                    l.type === leave.type &&
                    lsd <= dStr &&
                    led >= dStr
                  );
                })[0];
                if (actualLeave) removeLeaveDay(actualLeave.id, dStr);
              };
            })(leave._ids, dayStr);
            bar.appendChild(delB);
          }
          slotEl.appendChild(bar);
          cell.appendChild(slotEl);
        }
      }

      td.appendChild(cell);
      // Badges
      if (!isToday) {
        var hasLeaveBadge = userLeaves.some(function (leave) {
          var sd =
            leave.startDate || (leave.start ? leave.start.slice(0, 10) : '');
          var ed = leave.endDate || (leave.end ? leave.end.slice(0, 10) : '');
          return sd <= dayStr && ed >= dayStr;
        });
        if (isWeekend && !isUserWorkday) {
          var badge = document.createElement('div');
          badge.className = 'cal-dayoff-label';
          badge.textContent = 'day off';
          cell.appendChild(badge);
        } else if (isWeekend && isUserWorkday && !hasLeaveBadge) {
          var badge = document.createElement('div');
          badge.className = 'cal-workday-badge';
          badge.textContent = '📅 workday';
          badge.title = 'Set as work day for ' + user.name;
          cell.appendChild(badge);
        } else if (isAutoWorkday && !hasLeaveBadge) {
          var badge = document.createElement('div');
          badge.className = 'cal-workday-badge';
          badge.textContent = '✓ workday';
          badge.title = 'Past working day';
          cell.appendChild(badge);
        }
      }
      if (isElevated) {
        td.style.cursor = 'pointer';
        td.title = 'Double-click to add leave';
        td.ondblclick = (function (uid_inner, day_inner) {
          return function (e) {
            e.stopPropagation();
            openAddLeave();
            setTimeout(function () {
              document.getElementById('leave-user').value = uid_inner;
              var dateStr = day_inner.toISOString().slice(0, 10);
              document.getElementById('leave-start').value = dateStr;
              document.getElementById('leave-end').value = dateStr;
            }, 60);
          };
        })(user.id, day);
      }
      datesTr.appendChild(td);
    });
    datesFrag.appendChild(datesTr);
  });

  // Flush fragments → single reflow per panel
  namesTbody.appendChild(namesFrag);
  datesTbody.appendChild(datesFrag);
  namesTable.appendChild(namesTbody);
  datesTable.appendChild(datesTbody);

  // ── Sync vertical scroll between panels ──
  var namesPanel = EL.calNamesPanel || $('cal-names-panel');
  var datesPanel = EL.calDatesPanel || $('cal-dates-panel');
  // Remove old listeners by cloning
  var newNamesPanel = namesPanel.cloneNode(false);
  var newDatesPanel = datesPanel.cloneNode(false);
  newNamesPanel.appendChild(namesTable);
  newDatesPanel.appendChild(datesTable);
  if (namesPanel.parentNode)
    namesPanel.parentNode.replaceChild(newNamesPanel, namesPanel);
  if (datesPanel.parentNode)
    datesPanel.parentNode.replaceChild(newDatesPanel, datesPanel);
  // Update cached references so subsequent renders use the live DOM nodes
  EL.calNamesPanel = newNamesPanel;
  EL.calDatesPanel = newDatesPanel;
  // passive:true lets the browser skip calling preventDefault — improves scroll performance
  // ── Sticky header via JS transform ──
  function stickyHeader(panel) {
    // Disabled JS translateY. Relying on native CSS position: sticky
    // This prevents the double-movement bug on mobile scroll
  }

  newDatesPanel.addEventListener(
    'scroll',
    function () {
      newNamesPanel.scrollTop = newDatesPanel.scrollTop;
      stickyHeader(newDatesPanel);
      stickyHeader(newNamesPanel);
    },
    { passive: true },
  );
  newNamesPanel.addEventListener(
    'scroll',
    function () {
      newDatesPanel.scrollTop = newNamesPanel.scrollTop;
      stickyHeader(newDatesPanel);
      stickyHeader(newNamesPanel);
    },
    { passive: true },
  );

  // ── Wheel: vertical scrolls the panel content, Shift+wheel scrolls calendar horizontally ──
  function calWheelHandler(e) {
    var isHorizontal = e.shiftKey || Math.abs(e.deltaX) > Math.abs(e.deltaY);
    if (isHorizontal) {
      // Consume and apply horizontal scroll to dates panel
      e.preventDefault();
      e.stopPropagation();
      newDatesPanel.scrollLeft += e.shiftKey ? e.deltaY : e.deltaX;
    } else {
      // Vertical scroll — apply to both panels directly
      var dy = e.deltaY;
      newDatesPanel.scrollTop += dy;
      newNamesPanel.scrollTop = newDatesPanel.scrollTop;
      stickyHeader(newDatesPanel);
      stickyHeader(newNamesPanel);
    }
  }
  newDatesPanel.addEventListener('wheel', calWheelHandler, { passive: false });
  newNamesPanel.addEventListener('wheel', calWheelHandler, { passive: false });

  // ── Sync row heights between panels ──
  // Separates DOM reads from DOM writes to avoid layout thrashing.
  function syncRowHeights() {
    var nameRows = namesTable.querySelectorAll('tr');
    var dateRows = datesTable.querySelectorAll('tr');
    var len = Math.min(nameRows.length, dateRows.length);
    // PHASE 1: reset all heights (single write pass)
    for (var i = 0; i < len; i++) {
      Array.from(nameRows[i].cells).forEach(function (c) {
        c.style.height = '';
      });
      Array.from(dateRows[i].cells).forEach(function (c) {
        c.style.height = '';
      });
    }
    // PHASE 2: read all heights in one pass (no interleaved writes)
    var heights = new Array(len);
    for (var i = 0; i < len; i++) {
      heights[i] = Math.max(
        nameRows[i].getBoundingClientRect().height,
        dateRows[i].getBoundingClientRect().height,
      );
    }
    // PHASE 3: write all heights in one pass (no interleaved reads)
    for (var i = 0; i < len; i++) {
      var hpx = heights[i] + 'px';
      Array.from(nameRows[i].cells).forEach(function (c) {
        c.style.height = hpx;
      });
      Array.from(dateRows[i].cells).forEach(function (c) {
        c.style.height = hpx;
      });
    }
  }
  // Double rAF ensures browser has fully laid out the DOM before measuring
  requestAnimationFrame(function () {
    requestAnimationFrame(syncRowHeights);
  });

  // Scroll today's column into view
  setTimeout(function () {
    var todayCol = newDatesPanel.querySelector('th.today-col');
    if (todayCol)
      todayCol.scrollIntoView({ inline: 'center', block: 'nearest' });
  }, 50);

  // BusyCal-style date chip strip (mobile)
  try {
    buildCalWeekStrip(year, month, todayStr, users, allTasks, allLeaves);
  } catch (e) {
    /* strip is optional */
  }
}

/* ============================================================
   CALENDAR DATE CONTEXT MENU & DRAG-SELECT
   (Commented out for future reference — drag-select needs refinement)
   ============================================================
(function () {
  var _dragStart = null;    // date string of drag start
  var _dragEnd = null;      // date string of drag end
  var _isDragging = false;
  var _selectedDates = [];  // array of date strings
  var _dragUserId = null;   // user ID of the row being dragged

  // ── Helpers ──────────────────────────────────────────────
  function getDateFromTh(th) {
    return th && th.dataset && th.dataset.date ? th.dataset.date : null;
  }

  function getAllDateThs() {
    var panel = document.getElementById('cal-dates-panel') ||
      (EL && EL.calDatesPanel);
    if (!panel) return [];
    return Array.from(panel.querySelectorAll('thead th[data-date]'));
  }

  function getDateRange(startStr, endStr) {
    var start = new Date(startStr + 'T00:00:00');
    var end = new Date(endStr + 'T00:00:00');
    if (start > end) { var tmp = start; start = end; end = tmp; }
    var dates = [];
    var d = new Date(start);
    while (d <= end) {
      dates.push(d.toISOString().slice(0, 10));
      d.setDate(d.getDate() + 1);
    }
    return dates;
  }

  function highlightDates(dateStrs, userId) {
    // Clear all highlights first
    document.querySelectorAll('.cal-drag-selected').forEach(function (el) {
      el.classList.remove('cal-drag-selected');
    });
    if (!dateStrs || dateStrs.length === 0) return;
    var panel = document.getElementById('cal-dates-panel') ||
      (EL && EL.calDatesPanel);
    if (!panel) return;
    dateStrs.forEach(function (ds) {
      // Only highlight body cells for the specific user's row
      var selector = 'tbody td[data-date-str="' + ds + '"]';
      if (userId) selector += '[data-user-id="' + userId + '"]';
      var tds = panel.querySelectorAll(selector);
      tds.forEach(function (td) { td.classList.add('cal-drag-selected'); });
    });
  }

  function clearSelection() {
    _dragStart = null;
    _dragEnd = null;
    _isDragging = false;
    _selectedDates = [];
    _dragUserId = null;
    highlightDates([]);
  }

  // ── Context menu ─────────────────────────────────────────
  function showCalContextMenu(e, dates) {
    e.preventDefault();
    e.stopPropagation();
    var menu = document.getElementById('cal-date-context-menu');
    if (!menu) return;

    _selectedDates = dates;
    var startDate = dates[0];
    var endDate = dates[dates.length - 1];

    // Update label
    var label = document.getElementById('cal-ctx-date-label');
    if (label) {
      if (dates.length === 1) {
        var d = new Date(startDate + 'T00:00:00');
        label.textContent = d.toLocaleDateString(undefined, { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' });
      } else {
        var d1 = new Date(startDate + 'T00:00:00');
        var d2 = new Date(endDate + 'T00:00:00');
        label.textContent = d1.toLocaleDateString(undefined, { month: 'short', day: 'numeric' }) +
          ' → ' + d2.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' }) +
          ' (' + dates.length + ' days)';
      }
    }

    // Position menu
    var x = e.clientX;
    var y = e.clientY;
    menu.style.left = x + 'px';
    menu.style.top = y + 'px';
    menu.classList.add('visible');

    // Adjust if off-screen
    requestAnimationFrame(function () {
      var rect = menu.getBoundingClientRect();
      if (rect.right > window.innerWidth - 8) {
        menu.style.left = (window.innerWidth - rect.width - 8) + 'px';
      }
      if (rect.bottom > window.innerHeight - 8) {
        menu.style.top = (window.innerHeight - rect.height - 8) + 'px';
      }
    });
  }

  function hideCalContextMenu() {
    var menu = document.getElementById('cal-date-context-menu');
    if (menu) menu.classList.remove('visible');
  }

  // ── Menu actions ─────────────────────────────────────────
  function onMultiTask() {
    hideCalContextMenu();
    if (!_selectedDates.length) return;
    var startDate = _selectedDates[0];
    var endDate = _selectedDates[_selectedDates.length - 1];

    // Open multi-personnel task modal
    if (typeof openTeamTaskModal === 'function') {
      openTeamTaskModal(null);
      // Pre-fill start date after modal opens
      setTimeout(function () {
        var startEl = document.getElementById('tt-start');
        if (startEl) startEl.value = startDate + 'T09:00';
        // Switch to fixed end-date mode and set end date
        if (typeof setTtScheduleMode === 'function') setTtScheduleMode('fixed');
        var endEl = document.getElementById('tt-end');
        if (endEl) endEl.value = endDate + 'T17:00';
      }, 100);
    }
    clearSelection();
  }

  function onAddTask() {
    hideCalContextMenu();
    if (!_selectedDates.length) return;
    var startDate = _selectedDates[0];
    var endDate = _selectedDates[_selectedDates.length - 1];

    if (typeof openAddTask === 'function') {
      openAddTask();
      setTimeout(function () {
        var startEl = document.getElementById('task-start');
        if (startEl) startEl.value = startDate + 'T09:00';
        var deadlineEl = document.getElementById('task-deadline');
        if (deadlineEl) deadlineEl.value = endDate + 'T17:00';
      }, 100);
    }
    clearSelection();
  }

  function onAddLeave() {
    hideCalContextMenu();
    if (!_selectedDates.length) return;
    var startDate = _selectedDates[0];
    var endDate = _selectedDates[_selectedDates.length - 1];

    if (typeof openAddLeave === 'function') {
      openAddLeave();
      setTimeout(function () {
        var startEl = document.getElementById('leave-start');
        if (startEl) startEl.value = startDate;
        var endEl = document.getElementById('leave-end');
        if (endEl) endEl.value = endDate;
      }, 100);
    }
    clearSelection();
  }

  // ── Wire up menu button clicks ───────────────────────────
  document.addEventListener('DOMContentLoaded', function () {
    var multiBtn = document.getElementById('cal-ctx-multi-task');
    if (multiBtn) multiBtn.addEventListener('click', onMultiTask);
  });

  // ── Hide menu on click outside ───────────────────────────
  document.addEventListener('mousedown', function (e) {
    var menu = document.getElementById('cal-date-context-menu');
    if (menu && menu.classList.contains('visible') && !menu.contains(e.target)) {
      hideCalContextMenu();
      clearSelection();
    }
  });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') {
      hideCalContextMenu();
      clearSelection();
    }
  });

  // ── Helper to find a body cell (td) with data-date-str under mouse ──
  function findDateCell(target) {
    var el = target;
    while (el && el !== document.body) {
      if (el.tagName === 'TD' && el.dataset && el.dataset.dateStr) {
        // Verify it's inside a dates panel
        var panel = el.closest('.cal-dates-panel, #cal-dates-panel, [id="cal-dates-panel"]');
        if (panel) return el;
      }
      el = el.parentElement;
    }
    return null;
  }

  function getDateFromCell(td) {
    return td && td.dataset && td.dataset.dateStr ? td.dataset.dateStr : null;
  }

  // ── Drag-select on schedule body ─ uses mousedown/move/up ──
  document.addEventListener('mousedown', function (e) {
    var td = findDateCell(e.target);
    if (!td || e.button !== 0) return; // left click only
    hideCalContextMenu();
    _isDragging = true;
    _dragStart = getDateFromCell(td);
    _dragEnd = _dragStart;
    _dragUserId = td.dataset.userId || null;
    _selectedDates = [_dragStart];
    highlightDates(_selectedDates, _dragUserId);
    // Prevent text selection during drag via CSS
    document.body.style.userSelect = 'none';
    document.body.style.webkitUserSelect = 'none';
  });

  document.addEventListener('mousemove', function (e) {
    if (!_isDragging || !_dragStart) return;
    var td = findDateCell(e.target);
    if (td) {
      _dragEnd = getDateFromCell(td);
      if (_dragEnd) {
        _selectedDates = getDateRange(_dragStart, _dragEnd);
        highlightDates(_selectedDates, _dragUserId);
      }
    }
  });

  document.addEventListener('mouseup', function () {
    if (_isDragging) {
      _isDragging = false;
      // Restore text selection
      document.body.style.userSelect = '';
      document.body.style.webkitUserSelect = '';
      // Selection stays highlighted until right-click or click-away
    }
  });

  // ── Right-click on date HEADER → show context menu ───────
  document.addEventListener('contextmenu', function (e) {
    // Only trigger on date header th, not body cells
    var el = e.target;
    var th = null;
    while (el && el !== document.body) {
      if (el.tagName === 'TH' && el.dataset && el.dataset.date) {
        var panel = el.closest('.cal-dates-panel, #cal-dates-panel, [id="cal-dates-panel"]');
        if (panel) { th = el; break; }
      }
      el = el.parentElement;
    }
    if (!th) return;

    var clickedDate = th.dataset.date;
    if (!clickedDate) return;

    _selectedDates = [clickedDate];
    showCalContextMenu(e, [clickedDate]);
  });

  // ── Touch support for mobile (long-press = right-click) ──
  var _touchTimer = null;
  var _touchStartTd = null;
  document.addEventListener('touchstart', function (e) {
    var el = e.target;
    var th = null;
    while (el && el !== document.body) {
      if (el.tagName === 'TH' && el.dataset && el.dataset.date) {
        var panel = el.closest('.cal-dates-panel, #cal-dates-panel, [id="cal-dates-panel"]');
        if (panel) { th = el; break; }
      }
      el = el.parentElement;
    }
    if (!th) {
      var td = findDateCell(e.target);
      if (!td) return;
      _touchStartTd = td;
      _touchTimer = setTimeout(function () {
        var date = getDateFromCell(_touchStartTd);
        if (date) {
          _selectedDates = [date];
          highlightDates(_selectedDates);
          var touch = e.touches[0];
          showCalContextMenu({
            preventDefault: function () { },
            stopPropagation: function () { },
            clientX: touch.clientX,
            clientY: touch.clientY
          }, [date]);
        }
      }, 600);
    } else {
      var clickedDate = th.dataset.date;
      _touchTimer = setTimeout(function () {
        if (clickedDate) {
          _selectedDates = [clickedDate];
          var touch = e.touches[0];
          showCalContextMenu({
            preventDefault: function () { },
            stopPropagation: function () { },
            clientX: touch.clientX,
            clientY: touch.clientY
          }, [clickedDate]);
        }
      }, 600);
    }
  }, { passive: true });
  document.addEventListener('touchend', function () {
    clearTimeout(_touchTimer);
    _touchStartTd = null;
  }, { passive: true });
  document.addEventListener('touchmove', function () {
    clearTimeout(_touchTimer);
  }, { passive: true });
})();
============================================================ */

// saveLeave calendar refresh is handled inside the function itself

var _lieuTargetUserId = null;
var _lieuTargetUserName = null;

function openAddLieuDay(userId, userName) {
  if (
    state.currentUser.role !== 'manager' &&
    state.currentUser.role !== 'admin'
  ) {
    toast('Only managers can add lieu days.', 'error');
    return;
  }
  _lieuTargetUserId = userId;
  _lieuTargetUserName = userName;
  const currentBalance = countLieuDays(userId);
  const infoEl = document.getElementById('lieu-day-info');
  if (infoEl)
    infoEl.textContent =
      'Adding lieu days for ' +
      userName +
      '. Current balance: ' +
      currentBalance +
      ' day(s).';
  const countEl = document.getElementById('lieu-day-count');
  if (countEl) countEl.value = '1';
  openModal('lieu-day-modal');
}

function removeLieuDay(userId, userName) {
  if (
    state.currentUser.role !== 'manager' &&
    state.currentUser.role !== 'admin'
  ) {
    toast('Only managers can remove lieu days.', 'error');
    return;
  }
  const currentBalance = countLieuDays(userId);
  if (currentBalance <= 0) {
    toast('No lieu days to remove for ' + userName + '.', 'error');
    return;
  }
  // Subtract from bonus balance
  const s = getCompanySettings();
  const lieuBonus = s.lieuBonus || {};
  lieuBonus[userId] = (lieuBonus[userId] || 0) - 1;
  saveCompanySettings({ lieuBonus: lieuBonus });
  toast(
    'Removed 1 lieu day from ' +
    userName +
    '. Balance: ' +
    countLieuDays(userId),
    'success',
  );
  if (state.view === 'calendar') renderCalendarDebounced();
}

// Alias: HTML calls saveLieuDays(), logic lives in confirmAddLieuDays()
function saveLieuDays() {
  confirmAddLieuDays();
}

function confirmAddLieuDays() {
  const countInput =
    document.getElementById('lieu-day-count') ||
    document.getElementById('lieu-days-count');
  const count = parseInt(countInput ? countInput.value : 0);
  if (!count || count < 1) {
    toast('Please enter a valid number of lieu days.', 'error');
    return;
  }
  if (!_lieuTargetUserId) return;

  // Add bonus lieu days directly to the balance (no fake weekend workdays)
  const s = getCompanySettings();
  const lieuBonus = s.lieuBonus || {};
  lieuBonus[_lieuTargetUserId] = (lieuBonus[_lieuTargetUserId] || 0) + count;
  saveCompanySettings({ lieuBonus: lieuBonus });

  closeModal('lieu-day-modal');
  toast(
    `Added ${count} lieu day${count !== 1 ? 's' : ''} to ${_lieuTargetUserName}'s balance.`,
    'success',
  );
  if (state.view === 'calendar') renderCalendarDebounced();
}

function addUserCustomWorkday(userId, dateStr) {
  saveUserCustomWorkday(userId, dateStr);
}

/* ============================================================
   USER LIST (admin/manager)
   ============================================================ */
function renderUserList(filterText) {
  var grid = document.getElementById('user-grid');
  var isAdmin = state.currentUser.role === 'admin';
  var isElevated = isAdmin || state.currentUser.role === 'manager';
  var chpwBtn = document.getElementById('chpw-list-btn');
  var multiBtn = document.getElementById('multi-task-list-btn');
  var regBtn = document.getElementById('reg-user-list-btn');
  if (chpwBtn) chpwBtn.style.display = isAdmin ? '' : 'none';
  if (multiBtn) multiBtn.style.display = isElevated ? '' : 'none';
  if (regBtn) regBtn.style.display = isAdmin ? '' : 'none';
  var importBtn = document.getElementById('import-user-list-btn');
  if (importBtn) importBtn.style.display = isAdmin ? '' : 'none';

  // Search bar — insert once
  var searchWrap = document.getElementById('user-list-search-wrap');
  if (!searchWrap) {
    searchWrap = document.createElement('div');
    searchWrap.id = 'user-list-search-wrap';
    searchWrap.style.cssText =
      'width:100%;max-width:400px;margin:0 auto 16px;position:relative;';
    searchWrap.innerHTML =
      '<span style="position:absolute;left:10px;top:50%;transform:translateY(-50%);font-size:14px;pointer-events:none;">🔍</span>' +
      '<input type="text" id="user-list-search" placeholder="Search users…" ' +
      'style="width:100%;padding:10px 14px 10px 34px;background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius);color:var(--text);font-size:13px;font-family:var(--mono);">';
    grid.parentNode.insertBefore(searchWrap, grid);
    document
      .getElementById('user-list-search')
      .addEventListener('input', function () {
        renderUserList(this.value.trim().toLowerCase());
      });
  }
  // Keep search value in sync
  var searchInput = document.getElementById('user-list-search');
  if (searchInput && filterText === undefined)
    filterText = searchInput.value.trim().toLowerCase();

  var users = getUsers().filter(function (u) {
    return u.role !== 'admin';
  });
  // Apply search filter
  if (filterText) {
    users = users.filter(function (u) {
      return (
        u.name.toLowerCase().includes(filterText) ||
        u.username.toLowerCase().includes(filterText)
      );
    });
  }
  var tasks = getTasks();
  if (users.length === 0) {
    grid.innerHTML =
      '<div style="grid-column:1/-1;text-align:center;padding:60px;color:var(--text3);">' +
      (filterText
        ? 'No users match your search.'
        : 'No users registered yet. <br>Click &quot;Register User&quot; to add one.') +
      '</div>';
    return;
  }
  var roleColors = { manager: '#A78BFA', user: 'var(--p4)' };
  var roleLabels = { manager: '🏢 Manager', user: '👤 User' };
  grid.innerHTML = users
    .map(function (u) {
      var userTasks = tasks.filter(function (t) {
        return t.userId === u.id;
      });
      var active = userTasks.filter(function (t) {
        return !t.done && !t.cancelled;
      }).length;
      var cancelled = userTasks.filter(function (t) {
        return t.cancelled;
      }).length;
      var done = userTasks.filter(function (t) {
        return t.done;
      }).length;
      var roleColor = roleColors[u.role] || 'var(--p4)';
      var roleLabel = roleLabels[u.role] || u.role;
      var userTeams = getTeams().filter(function (t) {
        return (t.memberIds || []).includes(u.id);
      });
      var teamBadges =
        userTeams.length > 0
          ? '<div style="display:flex;flex-wrap:wrap;gap:4px;margin:5px 0 2px;">' +
          userTeams
            .map(function (t) {
              var c = t.color || '#F59E0B';
              return (
                '<span style="font-size:10px;font-weight:600;padding:2px 8px;border-radius:20px;background:' +
                c +
                '18;border:1px solid ' +
                c +
                '44;color:' +
                c +
                ';">🏷 ' +
                escHtml(t.name) +
                '</span>'
              );
            })
            .join('') +
          '</div>'
          : '';
      return (
        '<div class="user-card" onclick="viewUserTasks(\'' +
        u.id +
        '\')">' +
        '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;">' +
        '<div class="user-card-name">' +
        escHtml(u.name) +
        '</div>' +
        '<span style="font-size:10px;font-weight:700;padding:3px 8px;border-radius:12px;background:' +
        roleColor +
        '18;color:' +
        roleColor +
        ';border:1px solid ' +
        roleColor +
        '30;">' +
        roleLabel +
        '</span>' +
        '</div>' +
        '<div class="user-card-username">@' +
        escHtml(u.username) +
        '</div>' +
        teamBadges +
        '<div class="user-card-stats">' +
        '<div class="stat"><div class="stat-num">' +
        userTasks.length +
        '</div><div class="stat-label">Total</div></div>' +
        '<div class="stat"><div class="stat-num" style="color:var(--p1)">' +
        active +
        '</div><div class="stat-label">Active</div></div>' +
        '<div class="stat"><div class="stat-num" style="color:var(--danger)">' +
        cancelled +
        '</div><div class="stat-label">Cancelled</div></div>' +
        '<div class="stat"><div class="stat-num" style="color:var(--success)">' +
        done +
        '</div><div class="stat-label">Done</div></div>' +
        '</div>' +
        '<div class="user-card-actions" onclick="event.stopPropagation()">' +
        '<button class="btn-secondary" style="flex:1;font-size:11px;padding:6px;" onclick="viewUserTasks(\'' +
        u.id +
        '\')">View Tasks</button>' +
        (isElevated
          ? '<button class="btn-secondary" style="font-size:11px;padding:6px 10px;" data-tip="Assign multi-task" onclick="openMultiTaskFor(\'' +
          u.id +
          '\')">👥</button>'
          : '') +
        (isAdmin
          ? '<button class="btn-secondary" style="font-size:11px;padding:6px 10px;" data-tip="Change password" onclick="openChangePasswordFor(\'' +
          u.id +
          '\')">🔑</button>'
          : '') +
        (isAdmin
          ? '<button class="btn-danger" style="font-size:11px;padding:6px 10px;" onclick="deleteUser(\'' +
          u.id +
          '\')">✕</button>'
          : '') +
        '</div>' +
        '</div>'
      );
    })
    .join('');
}

initDB();

// Prevent context menu outside cards / calendar cells
document.addEventListener('contextmenu', (e) => {
  const calTd = e.target.closest('#cal-dates-panel td');

  // ── Calendar day right-click: task blocks get their own menu ──
  const isCalBlock = e.target.closest('.cal-block[data-taskid]');

  if (calTd && calTd.dataset.dateStr && !isCalBlock) {
    const role = state.currentUser?.role;
    const isElevated = role === 'admin' || role === 'manager';
    const isWorker = role === 'user';
    const d = new Date(calTd.dataset.dateStr + 'T12:00:00');
    const isWeekend = d.getDay() === 0 || d.getDay() === 6;

    // Weekend workday marker — admin/manager (right-click on day off shows work question tooltip)
    if (isWeekend && isElevated) {
      e.preventDefault();
      showCalCtxMenu(
        e,
        calTd.dataset.dateStr,
        calTd.dataset.userId,
        calTd.dataset.userName,
      );
      return;
    }

    // Day action menu — elevated roles on any day
    if (isElevated) {
      e.preventDefault();
      showCalDayCtxMenu(
        e,
        calTd.dataset.dateStr,
        calTd.dataset.userId,
        calTd.dataset.userName,
      );
      return;
    }
  }

  // Allow right-click on task cards, timeline bars, leave blocks, AND calendar task blocks
  if (
    !e.target.closest('.task-card') &&
    !e.target.closest('.timeline-bar') &&
    !e.target.closest('.leave-block') &&
    !isCalBlock
  )
    e.preventDefault();
});

/* ============================================================
   CALENDAR DAY CONTEXT MENU — per-user weekend workday marking
   ============================================================ */
var _calCtxDate = null;
var _calCtxUserId = null;

function showCalCtxMenu(e, dateStr, userId, userName) {
  hideCalCtxMenu();
  _calCtxDate = dateStr;
  _calCtxUserId = userId;
  const menu = document.getElementById('cal-ctx-menu');
  const userDays = getUserCustomWorkdays(userId);
  const isMarked = userDays.includes(dateStr);
  const d = new Date(dateStr + 'T12:00:00');
  const dayNames = [
    'Sunday',
    'Monday',
    'Tuesday',
    'Wednesday',
    'Thursday',
    'Friday',
    'Saturday',
  ];
  const dayLabel =
    dayNames[d.getDay()] +
    ', ' +
    d.toLocaleDateString('en-US', {
      month: 'long',
      day: 'numeric',
      year: 'numeric',
    });
  document.getElementById('cal-ctx-date-label').textContent = dayLabel;
  document.getElementById('cal-ctx-user-label').textContent =
    '👤 ' + (userName || 'Employee');
  document.getElementById('cal-ctx-question').textContent = isMarked
    ? 'This day is marked as worked. Remove it?'
    : 'Did this employee work on this day?';
  document.getElementById('cal-ctx-set-btn').style.display = isMarked
    ? 'none'
    : '';
  document.getElementById('cal-ctx-unset-btn').style.display = isMarked
    ? ''
    : 'none';
  menu.classList.remove('hidden');
  menu.style.left = e.clientX + 'px';
  menu.style.top = e.clientY + 'px';
  requestAnimationFrame(() => {
    const rect = menu.getBoundingClientRect();
    if (rect.right > window.innerWidth)
      menu.style.left = e.clientX - rect.width + 'px';
    if (rect.bottom > window.innerHeight)
      menu.style.top = e.clientY - rect.height + 'px';
  });
}
function hideCalCtxMenu() {
  document.getElementById('cal-ctx-menu').classList.add('hidden');
  _calCtxDate = null;
  _calCtxUserId = null;
}
document.addEventListener('click', hideCalCtxMenu);

function getUserCustomWorkdays(userId) {
  const s = getCompanySettings();
  const uwd = s.userWorkdays || {};
  return Array.isArray(uwd[userId]) ? uwd[userId] : [];
}
function saveUserCustomWorkday(userId, dateStr) {
  const s = getCompanySettings();
  const uwd = s.userWorkdays || {};
  const days = Array.isArray(uwd[userId]) ? uwd[userId] : [];
  if (!days.includes(dateStr)) uwd[userId] = [...days, dateStr];
  saveCompanySettings({ userWorkdays: uwd });
}
function removeUserCustomWorkday(userId, dateStr) {
  const s = getCompanySettings();
  const uwd = s.userWorkdays || {};
  uwd[userId] = (Array.isArray(uwd[userId]) ? uwd[userId] : []).filter(
    (d) => d !== dateStr,
  );
  saveCompanySettings({ userWorkdays: uwd });
}
function getCustomWorkdays() {
  // kept for legacy — returns flat global list (unused now)
  const s = getCompanySettings();
  return Array.isArray(s.customWorkdays) ? s.customWorkdays : [];
}
function calCtxSetWorkday() {
  if (!_calCtxDate || !_calCtxUserId) return;
  saveUserCustomWorkday(_calCtxUserId, _calCtxDate);
  const user = getUsers().find((u) => u.id === _calCtxUserId);
  toast(
    '📅 Weekend workday recorded for ' +
    (user?.name || 'employee') +
    ' — lieu balance updated',
    'success',
  );
  renderCalendar();
  hideCalCtxMenu();
}
function calCtxUnsetWorkday() {
  if (!_calCtxDate || !_calCtxUserId) return;
  removeUserCustomWorkday(_calCtxUserId, _calCtxDate);
  toast('Weekend workday removed — lieu balance updated', 'info');
  renderCalendar();
  hideCalCtxMenu();
}

/* ============================================================
   CALENDAR DAY RIGHT-CLICK CONTEXT MENU — Create Task / Set Status
   ============================================================ */
var _calDayCtxDate = null;
var _calDayCtxUserId = null;
var _calDayCtxUserName = null;

function showCalDayCtxMenu(e, dateStr, userId, userName) {
  hideCalDayCtxMenu();
  _calDayCtxDate = dateStr;
  _calDayCtxUserId = userId;
  _calDayCtxUserName = userName;

  const d = new Date(dateStr + 'T12:00:00');
  const dayNames = [
    'Sunday',
    'Monday',
    'Tuesday',
    'Wednesday',
    'Thursday',
    'Friday',
    'Saturday',
  ];
  const monthNames = [
    'Jan',
    'Feb',
    'Mar',
    'Apr',
    'May',
    'Jun',
    'Jul',
    'Aug',
    'Sep',
    'Oct',
    'Nov',
    'Dec',
  ];
  const label =
    dayNames[d.getDay()] +
    ', ' +
    monthNames[d.getMonth()] +
    ' ' +
    d.getDate() +
    ' ' +
    d.getFullYear();

  document.getElementById('cal-day-ctx-date-label').textContent = label;
  document.getElementById('cal-day-ctx-user-label').textContent =
    '👤 ' + (userName || 'Employee');

  // Gray out Lieu Day button if user has no lieu days available
  const lieuBtnBalance = document.querySelector(
    '#cal-day-ctx-menu .ctx-item[onclick="calDayCtxSetStatus(\'lieu\')"]',
  );
  if (lieuBtnBalance && userId) {
    const lieuBalance = countLieuDays(userId);
    lieuBtnBalance.disabled = lieuBalance <= 0;
    lieuBtnBalance.style.opacity = lieuBalance <= 0 ? '0.4' : '';
    lieuBtnBalance.style.cursor = lieuBalance <= 0 ? 'not-allowed' : '';
    lieuBtnBalance.title =
      lieuBalance <= 0 ? 'No lieu days available for this employee' : '';
  }

  const menu = document.getElementById('cal-day-ctx-menu');
  const isWorker = state.currentUser.role === 'user';

  // Role-based visibility for context menu items
  const setStatusSection = menu.querySelector('.cal-day-ctx-section');
  const lieuBtn = menu.querySelector('.ctx-item[onclick*="lieu"]');
  const loaBtn = menu.querySelector('.ctx-item[onclick*="loa"]');
  const awolBtn = menu.querySelector('.ctx-item[onclick*="awol"]');
  const createTaskBtn = menu.querySelector('.ctx-item[onclick*="CreateTask"]');

  if (isWorker) {
    if (setStatusSection) setStatusSection.style.display = 'none';
    if (lieuBtn) lieuBtn.style.display = 'none';
    if (loaBtn) loaBtn.style.display = 'none';
    if (awolBtn) awolBtn.style.display = 'none';
    if (createTaskBtn) createTaskBtn.style.display = 'none';

    let applyBtn = document.getElementById('cal-ctx-apply-leave-btn');
    if (!applyBtn) {
      applyBtn = document.createElement('button');
      applyBtn.className = 'ctx-item';
      applyBtn.id = 'cal-ctx-apply-leave-btn';
      applyBtn.innerHTML = '🏖️ &nbsp;Apply for Leave';
      applyBtn.onclick = function () {
        openApplyLeave(_calDayCtxDate);
        hideCalDayCtxMenu();
      };
      menu.appendChild(applyBtn);
    }
    applyBtn.style.display = '';
  } else {
    if (setStatusSection) setStatusSection.style.display = '';
    if (lieuBtn) lieuBtn.style.display = '';
    if (loaBtn) loaBtn.style.display = '';
    if (awolBtn) awolBtn.style.display = '';
    if (createTaskBtn) createTaskBtn.style.display = '';

    const applyBtn = document.getElementById('cal-ctx-apply-leave-btn');
    if (applyBtn) applyBtn.style.display = 'none';
  }

  menu.classList.remove('hidden');
  menu.style.left = e.clientX + 'px';
  menu.style.top = e.clientY + 'px';
  requestAnimationFrame(() => {
    const rect = menu.getBoundingClientRect();
    if (rect.right > window.innerWidth)
      menu.style.left = e.clientX - rect.width + 'px';
    if (rect.bottom > window.innerHeight)
      menu.style.top = e.clientY - rect.height + 'px';
  });
}

function hideCalDayCtxMenu() {
  const menu = document.getElementById('cal-day-ctx-menu');
  if (menu) menu.classList.add('hidden');
  _calDayCtxDate = null;
  _calDayCtxUserId = null;
  _calDayCtxUserName = null;
}
document.addEventListener('click', hideCalDayCtxMenu);

function calDayCtxCreateTask() {
  if (!_calDayCtxDate) return;
  const dateStr = _calDayCtxDate;
  const userId = _calDayCtxUserId;
  hideCalDayCtxMenu();

  // If viewing another user's calendar, set target but stay in calendar context
  if (userId && userId !== state.currentUser.id) {
    state.targetUserId = userId;
    state.view = 'user-tasks'; // so saveTask assigns to the right user
  }

  openAddTask();

  // Pre-populate start date from the clicked calendar day
  setTimeout(function () {
    const _wh = getWorkHours();
    const startHour = _wh.start || 8;
    const d = new Date(
      dateStr + 'T' + String(startHour).padStart(2, '0') + ':00:00',
    );
    const startEl = document.getElementById('f-start');
    if (startEl) {
      startEl.value = toLocalISO(d);
      if (typeof onTaskFieldChange === 'function') onTaskFieldChange();
    }
    // Store that we were in calendar so saveTask returns to calendar
    state._returnToCalendar = true;
  }, 80);
}

function calDayCtxSetStatus(type) {
  if (!_calDayCtxDate) return;
  const dateStr = _calDayCtxDate;
  const userId = _calDayCtxUserId;

  if (state.currentUser.role === 'user') {
    openApplyLeave(dateStr);
    hideCalDayCtxMenu();
    return;
  }

  hideCalDayCtxMenu();

  const role = state.currentUser?.role;
  if (role !== 'admin' && role !== 'manager') {
    toast('Only managers and admins can set leave status.', 'error');
    return;
  }

  // Block lieu if no balance
  if (type === 'lieu' && userId) {
    const lieuBalance = countLieuDays(userId);
    if (lieuBalance <= 0) {
      toast('No lieu days available for this employee.', 'error');
      return;
    }
  }

  // Block if a leave already exists on this day for this user
  if (userId) {
    const existing = getLeaves()
      .filter((l) => l.userId === userId)
      .find((l) => {
        const ls = l.startDate || (l.start ? l.start.slice(0, 10) : '');
        const le = l.endDate || (l.end ? l.end.slice(0, 10) : '');
        return ls <= dateStr && le >= dateStr;
      });
    if (existing) {
      toast(
        'A leave already exists on ' +
        dateStr +
        ' for this person. Remove it first.',
        'error',
      );
      return;
    }
  }

  openAddLeave();

  // Pre-populate leave modal fields
  setTimeout(function () {
    const userSel = document.getElementById('leave-user');
    if (userSel && userId) userSel.value = userId;
    const typeEl = document.getElementById('leave-type');
    if (typeEl) {
      typeEl.value = type;
      if (typeof updateLeaveTypeOptions === 'function')
        updateLeaveTypeOptions();
    }
    const startEl = document.getElementById('leave-start');
    const endEl = document.getElementById('leave-end');
    if (startEl) startEl.value = dateStr;
    if (endEl) endEl.value = dateStr;
  }, 80);
}

// ── NEW LEAVE REQUEST FUNCTIONS ─────────────────────────────────────────────

function openApplyLeave(dateStr) {
  if (state.currentUser.role !== 'user') return;

  // If no date passed, default to today or first working day of viewed month
  if (!dateStr) {
    const today = new Date();
    const viewYear = calState.year;
    const viewMonth = calState.month;
    const isSameMonth =
      today.getFullYear() === viewYear && today.getMonth() === viewMonth;
    dateStr = isSameMonth
      ? toDateStr(today)
      : toDateStr(new Date(viewYear, viewMonth, 1));
  }

  const todayStr = toDateStr(new Date());
  $('apply-leave-start').value = dateStr || todayStr;
  $('apply-leave-end').value = dateStr || todayStr;
  $('apply-leave-reason').value = '';
  $('apply-leave-type').value = 'lieu';
  $('apply-leave-modal-title').textContent = '🏖️ Apply for Leave';
  openModal('apply-leave-modal');
}

async function submitLeaveRequest() {
  const userId = state.currentUser.id;
  const type = $('apply-leave-type').value;
  const startDate = $('apply-leave-start').value;
  const endDate = $('apply-leave-end').value;
  const reason = $('apply-leave-reason').value.trim();

  if (!startDate || !endDate) {
    toast('Please fill in all dates.', 'error');
    return;
  }
  if (startDate > endDate) {
    toast('End date must be after start date.', 'error');
    return;
  }

  // Weekend check
  if (countWorkingDays(startDate, endDate) === 0) {
    toast(
      'The selected date range falls entirely on weekends. Please choose working days.',
      'error',
    );
    return;
  }

  // Check for existing leave overlap
  const existingLeaves = getLeaves().filter(
    (l) => l.userId === userId && l.status !== 'denied',
  );
  const cur = new Date(startDate + 'T12:00:00');
  const end = new Date(endDate + 'T12:00:00');
  while (cur <= end) {
    if (isWorkingDay(cur)) {
      const dStr = cur.toISOString().slice(0, 10);
      const clash = existingLeaves.find((l) => {
        const ls = l.startDate || (l.start ? l.start.slice(0, 10) : '');
        const le = l.endDate || (l.end ? l.end.slice(0, 10) : '');
        return ls <= dStr && le >= dStr;
      });
      if (clash) {
        toast('Leave already exists on ' + dStr + '. Remove it first.', 'error');
        return;
      }
    }
    cur.setDate(cur.getDate() + 1);
  }

  try {
    const newLeave = await API.post('/leaves', {
      userId,
      type,
      startDate,
      endDate,
      reason,
      status: 'pending',
      requestedBy: userId,
    });
    cache.leaves.push(newLeave);

    // Notify all managers + admins
    const userName = state.currentUser.name;
    const typeLabel = type === 'lieu' ? 'Lieu Day' : 'Leave of Absence';
    getUsers()
      .filter((u) => u.role === 'manager' || u.role === 'admin')
      .forEach((mgr) => {
        pushNotification(
          mgr.id,
          `🕐 Leave Request — ${userName}`,
          `${userName} has requested ${typeLabel} from ${startDate} to ${endDate}.`,
          newLeave.id,
          { type: 'leaveRequest', leaveId: newLeave.id },
        );
      });

    toast('Leave request submitted! Awaiting manager approval.', 'success');
    closeModal('apply-leave-modal');
    updateLeaveRequestsBadge();
    renderCalendarDebounced();
  } catch (err) {
    toast(err.message || 'Failed to submit request.', 'error');
  }
}

function openLeaveRequests(highlightId) {
  const pending = getLeaves().filter((l) => l.status === 'pending');
  const list = $('leave-requests-list');

  if (pending.length === 0) {
    list.innerHTML =
      '<div style="padding:24px;text-align:center;color:var(--text3);font-size:13px;">No pending leave requests.</div>';
  } else {
    list.innerHTML = pending
      .map((l) => {
        const user = getUsers().find((u) => u.id === l.userId);
        const typeLabel =
          l.type === 'lieu' ? '🟢 Lieu Day' : '🔵 Leave of Absence';
        return `
        <div data-leave-id="${l.id}" style="padding:14px 16px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;transition:background 0.3s;">
          <div>
            <div style="font-weight:700;font-size:13px;">${user?.name || 'Unknown'
          }</div>
            <div style="font-size:12px;color:var(--text2);margin-top:2px;">${typeLabel} · ${l.startDate
          }${l.endDate !== l.startDate ? ' → ' + l.endDate : ''}</div>
            ${l.reason
            ? `<div style="font-size:11px;color:var(--text3);margin-top:2px;">${l.reason}</div>`
            : ''
          }
          </div>
          <div style="display:flex;gap:8px;">
            <button class="btn-primary" style="width:auto;padding:7px 16px;font-size:12px;" onclick="resolveLeaveRequest('${l.id
          }','approved')">✅ Approve</button>
            <button class="btn-danger" style="padding:7px 16px;font-size:12px;" onclick="resolveLeaveRequest('${l.id
          }','denied')">❌ Deny</button>
          </div>
        </div>`;
      })
      .join('');
  }
  openModal('leave-requests-modal');

  if (highlightId) {
    setTimeout(() => {
      const el = list.querySelector(`[data-leave-id="${highlightId}"]`);
      if (el) {
        el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        el.style.background = 'rgba(245,158,11,0.12)';
        setTimeout(() => (el.style.background = ''), 2000);
      }
    }, 300);
  }
}

async function resolveLeaveRequest(leaveId, newStatus) {
  try {
    const updated = await API.patch('/leaves/' + leaveId, { status: newStatus });
    const idx = cache.leaves.findIndex((l) => l.id === leaveId);
    if (idx !== -1)
      cache.leaves[idx] = { ...cache.leaves[idx], status: newStatus };

    const leave = cache.leaves[idx];
    const user = getUsers().find((u) => u.id === leave?.userId);
    const typeLabel = leave?.type === 'lieu' ? 'Lieu Day' : 'Leave of Absence';
    const icon = newStatus === 'approved' ? '✅' : '❌';
    const label = newStatus === 'approved' ? 'Approved' : 'Denied';

    pushNotification(
      leave.userId,
      `${icon} Leave ${label} — ${typeLabel}`,
      `Your ${typeLabel} request (${leave.startDate} to ${leave.endDate}) has been ${label.toLowerCase()} by a manager.`,
      leaveId,
      { type: 'leaveResolution', status: newStatus },
    );

    toast(`Leave ${label.toLowerCase()} successfully.`, 'success');
    updateLeaveRequestsBadge();
    openLeaveRequests(); // refresh the list
    renderCalendarDebounced();
  } catch (err) {
    toast(err.message || 'Failed to update leave.', 'error');
  }
}

function updateLeaveRequestsBadge() {
  const count = getLeaves().filter((l) => l.status === 'pending').length;
  const badge = $('leave-requests-badge');
  if (!badge) return;
  if (count > 0) {
    badge.textContent = count;
    badge.classList.remove('hidden');
    badge.style.display = 'flex';
  } else {
    badge.classList.add('hidden');
    badge.style.display = 'none';
  }
}

// Backspace key triggers back button
document.addEventListener('keydown', function (e) {
  if (
    e.key === 'Backspace' &&
    e.target.tagName !== 'INPUT' &&
    e.target.tagName !== 'TEXTAREA' &&
    e.target.tagName !== 'SELECT'
  ) {
    var backBtn = document.getElementById('cal-back-btn');
    if (backBtn && state.view === 'calendar') {
      e.preventDefault();
      backBtn.click();
    } else if (
      state.previousView &&
      state.previousView !== state.view &&
      state.previousView !== 'login'
    ) {
      e.preventDefault();
      showView(state.previousView);
    }
  }
});

/* ============================================================
   IMPORT USERS — Excel / CSV bulk import (multi-sheet)
   ============================================================ */
var _importRows = [];
var _importRunning = false;
var _importWb = null; // cached XLSX workbook
var _importFileName = '';

function openImportUsers() {
  if (!state.currentUser || state.currentUser.role !== 'admin') {
    toast('Only admins can import users.', 'error');
    return;
  }
  importReset();
  openModal('import-users-modal');
}

function closeImportModal() {
  if (_importRunning) return;
  closeModal('import-users-modal');
  importReset();
}

function importReset() {
  _importRows = [];
  _importRunning = false;
  _importWb = null;
  _importFileName = '';
  document.getElementById('import-step-upload').classList.remove('hidden');
  document.getElementById('import-step-sheets').classList.add('hidden');
  document.getElementById('import-step-preview').classList.add('hidden');
  document.getElementById('import-submit-btn').classList.add('hidden');
  document.getElementById('import-table-body').innerHTML = '';
  document.getElementById('import-sheet-grid').innerHTML = '';
  document.getElementById('import-progress-wrap').classList.add('hidden');
  document.getElementById('import-progress-fill').style.width = '0%';
  document.getElementById('import-file-input').value = '';
  var dz = document.getElementById('import-drop-zone');
  if (dz) dz.classList.remove('dragover');
}

function importOnDragOver(e) {
  e.preventDefault();
  document.getElementById('import-drop-zone').classList.add('dragover');
}
function importOnDragLeave(e) {
  document.getElementById('import-drop-zone').classList.remove('dragover');
}
function importOnDrop(e) {
  e.preventDefault();
  document.getElementById('import-drop-zone').classList.remove('dragover');
  var file = e.dataTransfer.files[0];
  if (file) importHandleFile(file);
}

function importHandleFile(file) {
  if (!file) return;
  if (!window.XLSX) {
    toast('Excel parser not loaded yet — please try again.', 'error');
    return;
  }
  var ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx', 'xls', 'csv'].includes(ext)) {
    toast('Unsupported file. Use .xlsx, .xls or .csv', 'error');
    return;
  }
  _importFileName = file.name;
  var reader = new FileReader();
  reader.onload = function (ev) {
    try {
      var data = new Uint8Array(ev.target.result);
      var wb = XLSX.read(data, { type: 'array' });
      _importWb = wb;

      if (wb.SheetNames.length === 1) {
        // Single sheet — skip picker, go straight to preview
        var ws = wb.Sheets[wb.SheetNames[0]];
        var json = XLSX.utils.sheet_to_json(ws, { defval: '' });
        if (!json || !json.length) {
          toast('Sheet is empty.', 'error');
          return;
        }
        importParseRows(json, file.name + ' › ' + wb.SheetNames[0]);
      } else {
        // Multiple sheets — show picker
        importShowSheetPicker(wb, file.name);
      }
    } catch (err) {
      toast('Could not read file: ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function importShowSheetPicker(wb, fileName) {
  document.getElementById('import-step-upload').classList.add('hidden');
  document.getElementById('import-sheet-file-label').textContent =
    fileName +
    ' — ' +
    wb.SheetNames.length +
    ' sheets found. Click a sheet to preview it.';

  var grid = document.getElementById('import-sheet-grid');
  grid.innerHTML = wb.SheetNames.map(function (name) {
    var ws = wb.Sheets[name];
    var json = XLSX.utils.sheet_to_json(ws, { defval: '' });
    var rowTxt = json.length + ' row' + (json.length !== 1 ? 's' : '');
    return (
      '<div class="sheet-tab" onclick="importSelectSheet(\'' +
      name.replace(/\\/g, '\\\\').replace(/'/g, "\\'") +
      '\')">' +
      '📋 ' +
      escHtml(name) +
      '<div class="sheet-tab-rows">' +
      rowTxt +
      '</div>' +
      '</div>'
    );
  }).join('');

  document.getElementById('import-step-sheets').classList.remove('hidden');
}

function importSelectSheet(sheetName) {
  if (!_importWb) return;
  var ws = _importWb.Sheets[sheetName];
  var json = XLSX.utils.sheet_to_json(ws, { defval: '' });
  if (!json || !json.length) {
    toast('That sheet is empty.', 'error');
    return;
  }
  document.getElementById('import-step-sheets').classList.add('hidden');
  importParseRows(json, _importFileName + ' › ' + sheetName);
}

function importBackToSheets() {
  if (!_importWb) {
    importReset();
    return;
  }
  _importRows = [];
  document.getElementById('import-step-preview').classList.add('hidden');
  document.getElementById('import-submit-btn').classList.add('hidden');
  document.getElementById('import-table-body').innerHTML = '';
  document.getElementById('import-progress-wrap').classList.add('hidden');
  document.getElementById('import-progress-fill').style.width = '0%';
  if (_importWb.SheetNames.length === 1) {
    importReset();
    return;
  }
  importShowSheetPicker(_importWb, _importFileName);
}

function importParseRows(json, label) {
  var ALIAS = {
    'full name': 'name',
    fullname: 'name',
    'display name': 'name',
    'user name': 'username',
    user_name: 'username',
    login: 'username',
    pass: 'password',
    pwd: 'password',
    'e-mail': 'email',
    mail: 'email',
    group: 'team',
    department: 'team',
    dept: 'team',
    type: 'role',
    'user type': 'role',
    'account type': 'role',
  };
  var existingUnames = getUsers().map(function (u) {
    return u.username.toLowerCase();
  });
  var rows = [];

  json.slice(0, 1000).forEach(function (raw, idx) {
    var norm = {};
    Object.keys(raw).forEach(function (k) {
      var key = k.trim().toLowerCase();
      norm[ALIAS[key] || key] = String(raw[k]).trim();
    });
    var name = norm['name'] || '';
    var username = (norm['username'] || '').toLowerCase().replace(/\s+/g, '');
    var password = norm['password'] || '';
    var email = norm['email'] || '';
    var team = norm['team'] || '';
    var role = (norm['role'] || 'user').toLowerCase();

    var errors = [];
    if (!name) errors.push('name required');
    if (!username) errors.push('username required');
    if (!password) errors.push('password required');
    if (password && password.length < 4) errors.push('password min 4 chars');
    if (role !== 'user' && role !== 'manager')
      errors.push('role must be "user" or "manager"');
    if (existingUnames.includes(username))
      errors.push('username already taken');

    rows.push({
      idx: idx + 1,
      name,
      username,
      password,
      email,
      team,
      role,
      errors,
      valid: errors.length === 0,
      status: errors.length === 0 ? '✓ Ready' : '⚠ ' + errors[0],
    });
  });

  _importRows = rows;
  var validCount = rows.filter(function (r) {
    return r.valid;
  }).length;
  var errCount = rows.length - validCount;

  document.getElementById('import-cnt-total').textContent = rows.length;
  document.getElementById('import-cnt-valid').textContent = validCount;
  document.getElementById('import-cnt-errors').textContent = errCount;
  document.getElementById('import-file-name').textContent = label;

  var tbody = document.getElementById('import-table-body');
  tbody.innerHTML = rows
    .map(function (r) {
      var rowCls = r.valid ? 'row-ok' : 'row-error';
      var roleCls =
        r.role === 'manager'
          ? 'import-role-manager'
          : r.role === 'user'
            ? 'import-role-user'
            : 'import-role-invalid';
      return (
        '<tr class="' +
        rowCls +
        '">' +
        '<td style="color:var(--text3)">' +
        r.idx +
        '</td>' +
        '<td>' +
        escHtml(r.name) +
        '</td>' +
        '<td style="color:var(--amber)">@' +
        escHtml(r.username) +
        '</td>' +
        '<td style="color:var(--text3)">••••••</td>' +
        '<td style="color:var(--text3)">' +
        escHtml(r.email || '—') +
        '</td>' +
        '<td>' +
        (r.team
          ? escHtml(r.team)
          : '<span style="color:var(--text3)">—</span>') +
        '</td>' +
        '<td><span class="import-role-badge ' +
        roleCls +
        '">' +
        escHtml(r.role) +
        '</span></td>' +
        '<td class="import-status">' +
        escHtml(r.status) +
        '</td>' +
        '</tr>'
      );
    })
    .join('');

  document.getElementById('import-step-preview').classList.remove('hidden');

  if (validCount > 0) {
    var btn = document.getElementById('import-submit-btn');
    btn.textContent =
      'Import ' +
      validCount +
      ' Valid User' +
      (validCount !== 1 ? 's' : '') +
      ' →';
    btn.classList.remove('hidden');
  }
}

async function importSubmit() {
  var validRows = _importRows.filter(function (r) {
    return r.valid;
  });
  if (!validRows.length) {
    toast('No valid rows to import.', 'error');
    return;
  }
  if (_importRunning) return;
  _importRunning = true;
  var _importBtn = document.getElementById('import-submit-btn');
  _lockOp('importUsers', _importBtn, 'Importing users…');
  document.getElementById('import-progress-wrap').classList.remove('hidden');

  var fill = document.getElementById('import-progress-fill');
  var label = document.getElementById('import-progress-label');
  var teams = getTeams();
  var done = 0,
    failed = 0;
  var total = validRows.length;

  for (var i = 0; i < validRows.length; i++) {
    var r = validRows[i];
    label.textContent =
      'Importing ' + (i + 1) + ' / ' + total + ' — ' + r.name + '…';
    fill.style.width = Math.round(((i + 1) / total) * 100) + '%';

    var trEl = null;
    var allTr = document
      .getElementById('import-table-body')
      .querySelectorAll('tr');
    for (var ti = 0; ti < allTr.length; ti++) {
      var cells = allTr[ti].querySelectorAll('td');
      if (cells.length && parseInt(cells[0].textContent) === r.idx) {
        trEl = allTr[ti];
        break;
      }
    }
    if (trEl) {
      trEl.querySelector('.import-status').textContent = '⏳';
      trEl.className = '';
    }

    try {
      var newUser = await API.post('/users', {
        name: r.name,
        username: r.username,
        password: r.password,
        role: r.role,
        email: r.email,
        emailNotif: !!r.email,
      });
      cache.users.push(newUser);

      if (r.team) {
        var matchTeam = teams.find(function (t) {
          return t.name.toLowerCase() === r.team.toLowerCase();
        });
        if (!matchTeam) {
          // Auto-create the team if it doesn't exist yet
          try {
            matchTeam = await API.post('/teams', {
              name: r.team,
              color: '#F59E0B',
              memberIds: [],
            });
            teams.push(matchTeam);
            cache.teams.push(matchTeam);
          } catch (teamErr) {
            matchTeam = null; // team creation failed, skip assignment
          }
        }
        if (matchTeam) {
          if (!matchTeam.memberIds) matchTeam.memberIds = [];
          if (!matchTeam.memberIds.includes(newUser.id)) {
            matchTeam.memberIds.push(newUser.id);
            await API.put('/teams/' + matchTeam.id, {
              memberIds: matchTeam.memberIds,
            });
          }
        }
      }
      done++;
      if (trEl) {
        trEl.querySelector('.import-status').textContent = '✓ Imported';
        trEl.className = 'row-done';
      }
    } catch (err) {
      failed++;
      var msg = err && err.message ? err.message : 'Failed';
      if (trEl) {
        trEl.querySelector('.import-status').textContent = '✕ ' + msg;
        trEl.className = 'row-fail';
      }
    }
  }

  fill.style.width = '100%';
  label.textContent =
    '✓ Done — ' +
    done +
    ' imported' +
    (failed ? ', ' + failed + ' failed' : '') +
    '.';
  _importRunning = false;
  _unlockOp('importUsers', _importBtn);
  document.getElementById('import-submit-btn').classList.add('hidden');

  try {
    cache.teams = (await API.get('/teams')) || cache.teams;
  } catch (e) { }

  var summary = done + ' user' + (done !== 1 ? 's' : '') + ' imported!';
  if (failed) summary += ' ' + failed + ' failed — see table for details.';
  toast(summary, done > 0 ? 'success' : 'error');

  if (state.view === 'user-list') renderUserList();
  if (state.view === 'teams') renderTeamsView();
}

function importDownloadTemplate() {
  var csv =
    'name,username,password,email,team,role\n' +
    'Jane Smith,jane.smith,password123,jane@company.com,Engineering,user\n' +
    'John Doe,john.doe,password123,john@company.com,Sales,manager\n';
  var blob = new Blob([csv], { type: 'text/csv' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = 'taskflow_import_template.csv';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/* ============================================================
   DELETE ALL — danger zone
   ============================================================ */
function openDeleteAll() {
  if (!state.currentUser || state.currentUser.role !== 'admin') {
    toast('Only admins can use this feature.', 'error');
    return;
  }
  // Reset state
  ['del-tasks', 'del-users', 'del-teams', 'del-all'].forEach(function (id) {
    document.getElementById(id).checked = false;
  });
  document.getElementById('del-confirm-input').value = '';
  document.getElementById('delete-all-btn').disabled = true;
  document.getElementById('del-progress-wrap').classList.add('hidden');
  document.getElementById('del-progress-fill').style.width = '0%';
  openModal('delete-all-modal');
}

function toggleDeleteAll(cb) {
  var checked = cb.checked;
  ['del-tasks', 'del-users', 'del-teams'].forEach(function (id) {
    document.getElementById(id).checked = checked;
  });
  updateDeleteBtn();
}

function updateDeleteBtn() {
  // Sync "Everything" checkbox state
  var allChecked = ['del-tasks', 'del-users', 'del-teams'].every(function (id) {
    return document.getElementById(id).checked;
  });
  document.getElementById('del-all').checked = allChecked;

  var anyChecked = ['del-tasks', 'del-users', 'del-teams', 'del-all'].some(
    function (id) {
      return document.getElementById(id).checked;
    },
  );
  var confirmed =
    document.getElementById('del-confirm-input').value.trim().toUpperCase() ===
    'DELETE';
  document.getElementById('delete-all-btn').disabled = !(
    anyChecked && confirmed
  );
}

async function executeDeleteAll() {
  var delTasks = document.getElementById('del-tasks').checked;
  var delUsers = document.getElementById('del-users').checked;
  var delTeams = document.getElementById('del-teams').checked;
  var confirmed =
    document.getElementById('del-confirm-input').value.trim().toUpperCase() ===
    'DELETE';

  if (!confirmed) {
    toast('Type DELETE to confirm.', 'error');
    return;
  }
  if (!delTasks && !delUsers && !delTeams) {
    toast('Select at least one option.', 'error');
    return;
  }

  var _delBtn = document.getElementById('delete-all-btn');
  _lockOp('deleteAll', _delBtn, 'Deleting data…');
  document.getElementById('del-progress-wrap').classList.remove('hidden');
  var fill = document.getElementById('del-progress-fill');
  var label = document.getElementById('del-progress-label');

  var steps = [];
  if (delTasks || delUsers) steps.push('tasks');
  if (delUsers) steps.push('users');
  if (delTeams) steps.push('teams');
  var total = steps.length;
  var step = 0;

  function progress(msg) {
    step++;
    label.textContent = msg;
    fill.style.width = Math.round((step / total) * 100) + '%';
  }

  try {
    if (delTasks || delUsers) {
      progress('Deleting all tasks…');
      var tasks = getTasks();
      for (var i = 0; i < tasks.length; i++) {
        try {
          await API.del('/tasks/' + tasks[i].id);
        } catch (e) { }
      }
      cache.tasks = [];
    }

    if (delUsers) {
      progress('Deleting all user accounts…');
      var users = getUsers().filter(function (u) {
        return u.role !== 'admin';
      });
      for (var j = 0; j < users.length; j++) {
        try {
          await API.del('/users/' + users[j].id);
        } catch (e) { }
      }
      cache.users = cache.users.filter(function (u) {
        return u.role === 'admin';
      });
    }

    if (delTeams) {
      progress('Deleting all teams…');
      var teams = getTeams();
      for (var k = 0; k < teams.length; k++) {
        try {
          await API.del('/teams/' + teams[k].id);
        } catch (e) { }
      }
      cache.teams = [];
    }

    fill.style.width = '100%';
    label.textContent = '✓ Done.';

    var parts = [];
    if (delTasks || delUsers) parts.push('tasks');
    if (delUsers) parts.push('accounts');
    if (delTeams) parts.push('teams');
    toast('Deleted all ' + parts.join(', ') + '.', 'success');

    setTimeout(function () {
      _unlockOp('deleteAll', _delBtn);
      closeModal('delete-all-modal');
      if (state.view === 'user-list') renderUserList();
      else if (state.view === 'task-board') renderTasks(state.currentUser.id);
      else if (state.view === 'teams') renderTeamsView();
    }, 800);
  } catch (err) {
    toast('Delete failed: ' + (err.message || 'Unknown error'), 'error');
    _unlockOp('deleteAll', _delBtn);
  }
}

/* ============================================================
   SWIPE DOWN TO DISMISS MODALS (mobile bottom-sheet gesture)
   ============================================================ */
document.addEventListener('DOMContentLoaded', function () {
  (function initModalSwipeDismiss() {
    if (!('ontouchstart' in window)) return;

    let _startY = 0,
      _modal = null,
      _isDragging = false;

    document.addEventListener(
      'touchstart',
      (e) => {
        const header = e.target.closest('.modal-header');
        if (!header) return;
        _modal = header.closest('.modal');
        if (!_modal) return;
        _startY = e.touches[0].clientY;
        _isDragging = true;
        _modal.style.transition = 'none';
      },
      { passive: true },
    );

    document.addEventListener(
      'touchmove',
      (e) => {
        if (!_isDragging || !_modal) return;
        const dy = e.touches[0].clientY - _startY;
        if (dy > 0) _modal.style.transform = 'translateY(' + dy + 'px)';
      },
      { passive: true },
    );

    document.addEventListener(
      'touchend',
      (e) => {
        if (!_isDragging || !_modal) return;
        const dy = e.changedTouches[0].clientY - _startY;
        _modal.style.transition = '';
        if (dy > 80) {
          const overlay = _modal.closest('.modal-overlay');
          if (overlay && overlay.id) closeModal(overlay.id);
          else _modal.style.transform = '';
        } else {
          _modal.style.transform = '';
        }
        _isDragging = false;
        _modal = null;
      },
      { passive: true },
    );
  })();
});

/* ============================================================
   MOBILE KEYBOARD AVOIDANCE FOR MODALS
   Uses visualViewport API to detect keyboard open/close.
   When keyboard opens  → modal shrinks to fit visible space.
   When keyboard closes → modal smoothly expands back to full
   height and stays anchored at the bottom (no position jump).
   ============================================================ */
document.addEventListener('DOMContentLoaded', function () {
  (function initModalKeyboardAwareness() {
    if (!('ontouchstart' in window)) return;
    const vv = window.visualViewport;
    if (!vv) return;

    // Natural full height the modal should return to after keyboard hides
    const FULL_MAX_H = '92dvh';

    const _origOpenModal = window.openModal;
    window.openModal = function (id) {
      _origOpenModal(id);
      const overlay = document.getElementById(id);
      if (!overlay) return;
      // Scroll modal body to top when freshly opened
      const body = overlay.querySelector('.modal-body');
      if (body) body.scrollTop = 0;
    };

    const _origCloseModal = window.closeModal;
    window.closeModal = function (id) {
      _origCloseModal(id);
      // Restore any keyboard-imposed sizing on close
      const overlay = document.getElementById(id);
      if (overlay) {
        const modal = overlay.querySelector('.modal');
        if (modal) {
          modal.style.removeProperty('max-height');
          modal.style.removeProperty('transition');
        }
      }
    };

    function _adjustForKeyboard() {
      const overlays = document.querySelectorAll('.modal-overlay:not(.hidden)');
      if (overlays.length === 0) return;

      const viewportH = vv.height;
      const viewportTop = vv.offsetTop;
      const windowH = window.innerHeight;
      const kbHeight = windowH - viewportTop - viewportH;
      const kbOpen = kbHeight > 100; // > 100px means keyboard is up

      overlays.forEach((overlay) => {
        const modal = overlay.querySelector('.modal');
        if (!modal) return;

        if (kbOpen) {
          // Keyboard visible — shrink modal to fit remaining space.
          // No overlay repositioning needed; the modal is position:fixed bottom:0
          // so it naturally sits above the keyboard as the visual viewport shrinks.
          const targetH = Math.floor(viewportH * 0.96);
          modal.style.transition = 'max-height 0.18s ease';
          modal.style.maxHeight = targetH + 'px';

          // Scroll the focused input into the visible portion of the modal body
          const focused = document.activeElement;
          if (focused && modal.contains(focused)) {
            setTimeout(
              () =>
                focused.scrollIntoView({
                  behavior: 'smooth',
                  block: 'nearest',
                }),
              120,
            );
          }
        } else {
          // Keyboard hidden — smoothly restore full height.
          // Use a transition so the sheet expands naturally rather than snapping.
          modal.style.transition = 'max-height 0.28s ease';
          modal.style.maxHeight = FULL_MAX_H;
          // After the transition finishes, remove the inline style so CSS takes over
          const clearTransition = () => {
            modal.style.removeProperty('max-height');
            modal.style.removeProperty('transition');
            modal.removeEventListener('transitionend', clearTransition);
          };
          modal.addEventListener('transitionend', clearTransition, {
            once: true,
          });
        }
      });
    }

    vv.addEventListener('resize', _adjustForKeyboard);

    // Scroll focused inputs into view (after keyboard animation completes)
    document.addEventListener(
      'focusin',
      (e) => {
        const target = e.target;
        if (!target.closest('.modal')) return;
        setTimeout(
          () => target.scrollIntoView({ behavior: 'smooth', block: 'nearest' }),
          380,
        );
      },
      true,
    );
  })();
});

// Enhance openModal to always focus first input field (improves UX on desktop too)
document.addEventListener('DOMContentLoaded', function () {
  (function enhanceOpenModal() {
    const _base = window.openModal;
    window.openModal = function (id) {
      _base(id);
      requestAnimationFrame(() => {
        const modal = document.getElementById(id);
        if (!modal) return;
        // Focus first visible text/number/date input in modal body
        const inp = modal.querySelector(
          '.modal-body input:not([type=checkbox]):not([type=radio]):not([hidden]), .modal-body textarea',
        );
        if (inp) {
          setTimeout(() => inp.focus(), 150);
        }
      });
    };
  })();
});

/* ============================================================
   OFFLINE QUEUE — IndexedDB-backed write queue
   When offline, mutations (POST/PATCH/DELETE) are queued here.
   On reconnect, the queue is replayed in order.
   ============================================================ */
const OfflineQueue = (() => {
  const DB_NAME = 'tf_offline_db';
  const DB_VERSION = 1;
  const STORE = 'queue';
  let _db = null;

  function _open() {
    return new Promise((resolve, reject) => {
      if (_db) {
        resolve(_db);
        return;
      }
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains(STORE)) {
          db.createObjectStore(STORE, { keyPath: 'id', autoIncrement: true });
        }
      };
      req.onsuccess = (e) => {
        _db = e.target.result;
        resolve(_db);
      };
      req.onerror = (e) => reject(e.target.error);
    });
  }

  async function enqueue(op) {
    // op: { method, path, body, localId, localType }
    const db = await _open();
    const tx = db.transaction(STORE, 'readwrite');
    const store = tx.objectStore(STORE);
    store.add({ ...op, ts: Date.now() });
    return new Promise((res, rej) => {
      tx.oncomplete = res;
      tx.onerror = rej;
    });
  }

  async function getAll() {
    const db = await _open();
    const tx = db.transaction(STORE, 'readonly');
    const store = tx.objectStore(STORE);
    return new Promise((res, rej) => {
      const req = store.getAll();
      req.onsuccess = () => res(req.result || []);
      req.onerror = () => rej(req.error);
    });
  }

  async function remove(id) {
    const db = await _open();
    const tx = db.transaction(STORE, 'readwrite');
    const store = tx.objectStore(STORE);
    store.delete(id);
    return new Promise((res, rej) => {
      tx.oncomplete = res;
      tx.onerror = rej;
    });
  }

  async function _pingServer() {
    // Lightweight connectivity check — hit the Supabase health endpoint
    try {
      const res = await fetch(SUPA_URL + '/rest/v1/', {
        method: 'GET',
        headers: { apikey: SUPA_KEY },
        signal: AbortSignal.timeout ? AbortSignal.timeout(5000) : undefined,
      });
      return res.ok || res.status === 400; // 400 = no table specified but server is alive
    } catch {
      return false;
    }
  }

  async function replay() {
    if (!navigator.onLine) return 0;
    // Verify actual server reachability before attempting to sync
    const serverUp = await _pingServer();
    if (!serverUp) {
      console.log(
        '[OfflineQueue] navigator.onLine=true but server unreachable — skipping sync',
      );
      return 0;
    }
    const items = await getAll();
    if (items.length === 0) return 0;
    let synced = 0;
    for (const item of items) {
      try {
        await API.request(item.method, item.path, item.body);
        await remove(item.id);
        synced++;
      } catch (e) {
        // If it's a network error, stop — don't process further
        if (!navigator.onLine) break;
        // If server went down mid-replay, stop
        const stillUp = await _pingServer();
        if (!stillUp) break;
        // If it's a data/server error, remove it (don't block the queue forever)
        await remove(item.id);
      }
    }
    if (synced > 0) {
      // Remove optimistic offline tasks from local cache so they get replaced
      // by the real server-synced tasks when loadAll() re-fetches
      if (cache.tasks) {
        cache.tasks = cache.tasks.filter(
          (t) => !t.id || !String(t.id).startsWith('offline_'),
        );
      }
      // Refresh data from server after sync
      try {
        await loadAll();
      } catch (e) { }
      const userId = state.currentUser?.id;
      if (userId) {
        if (state.currentViewMode === 'timeline') renderTimeline(userId);
        else if (
          document.getElementById('task-board') &&
          !document.getElementById('task-board').classList.contains('hidden')
        )
          renderTasks(userId);
        if (state.view === 'calendar') renderCalendar();
      }
      toast(
        `📶 ${synced} offline change${synced > 1 ? 's' : ''} synced to server.`,
        'success',
      );
    }
    return synced;
  }

  // Listen for reconnect — wait 2s for connection to stabilize before syncing
  if (typeof window !== 'undefined' && window.addEventListener) {
    window.addEventListener('online', () => {
      if (state.currentUser) setTimeout(() => replay(), 2000);
    });
  }

  async function count() {
    const items = await getAll();
    return items.length;
  }

  // Offline indicator in header
  function updateOnlineStatus() {
    let indicator = document.getElementById('tf-offline-indicator');
    if (!indicator) {
      indicator = document.createElement('div');
      indicator.id = 'tf-offline-indicator';
      indicator.style.cssText = [
        'position:fixed',
        'top:10px',
        'left:50%',
        'transform:translateX(-50%)',
        'background:#EF4444',
        'color:#fff',
        'font-size:11px',
        'font-weight:700',
        'padding:5px 14px',
        'border-radius:20px',
        'z-index:99999',
        'display:none',
        'align-items:center',
        'gap:6px',
        'box-shadow:0 4px 16px rgba(0,0,0,0.4)',
      ].join(';');
      indicator.innerHTML = '📵 Offline — changes will sync when reconnected';
      document.body.appendChild(indicator);
    }
    indicator.style.display = navigator.onLine ? 'none' : 'flex';
  }

  if (typeof window !== 'undefined' && window.addEventListener) {
    window.addEventListener('online', updateOnlineStatus);
    window.addEventListener('offline', updateOnlineStatus);
    // Initial check on load
    setTimeout(updateOnlineStatus, 1000);
  }

  return { enqueue, getAll, remove, replay, count };
})();

/* ============================================================
   SINGLE-SESSION CONTROL
   One active session per account at a time.
   Login from a new device shows a transfer prompt; the old
   session is invalidated. Works by writing a session token to
   the Supabase user row (session_token column).
   Since that column may not exist, we store it in localStorage
   with a device fingerprint approach instead.
   ============================================================ */
const Session = (() => {
  const DEVICE_KEY = 'tf_device_id';
  const SESSION_KEY = 'tf_session_id';

  function getDeviceId() {
    let id = localStorage.getItem(DEVICE_KEY);
    if (!id) {
      id =
        'dev_' +
        Date.now().toString(36) +
        '_' +
        Math.random().toString(36).slice(2);
      localStorage.setItem(DEVICE_KEY, id);
    }
    return id;
  }

  function newSessionId() {
    return (
      'sess_' +
      Date.now().toString(36) +
      '_' +
      Math.random().toString(36).slice(2)
    );
  }

  // Store active session in Supabase under tf_users.active_session
  // Falls back gracefully if the column doesn't exist
  async function claimSession(userId) {
    const sessId = newSessionId();
    const deviceId = getDeviceId();
    localStorage.setItem(SESSION_KEY, sessId);
    try {
      await SB.update('tf_users', userId, {
        active_session: JSON.stringify({ sessId, deviceId, ts: Date.now() }),
      });
    } catch (e) {
      // Column may not exist — that's fine, degrade gracefully
    }
    return sessId;
  }

  // Check if another device has taken over the session
  async function checkSession(userId) {
    const mySession = localStorage.getItem(SESSION_KEY);
    const myDevice = getDeviceId();
    if (!mySession) return 'ok'; // no session to check
    try {
      const rows = await SB.select(
        'tf_users',
        `id=eq.${userId}&select=active_session&limit=1`,
      );
      if (!rows || rows.length === 0) return 'ok';
      const raw = rows[0].active_session;
      if (!raw) return 'ok'; // column empty or doesn't exist
      const saved = JSON.parse(raw);
      if (saved.sessId !== mySession) {
        return 'stolen'; // another device claimed the session
      }
    } catch (e) {
      return 'ok'; // column doesn't exist or error — degrade gracefully
    }
    return 'ok';
  }

  // Show a non-dismissable banner that session was taken over
  function showSessionStolenBanner() {
    const b = document.createElement('div');
    b.style.cssText = `
      position:fixed; inset:0; z-index:999999; background:rgba(0,0,0,0.92);
      display:flex; flex-direction:column; align-items:center; justify-content:center;
      font-family:var(--mono); gap:20px; padding:24px; text-align:center;
    `;
    b.innerHTML = `
      <div style="font-size:32px;">📱➡️💻</div>
      <div style="font-size:18px;font-weight:700;color:var(--amber);">Session transferred</div>
      <div style="font-size:13px;color:var(--text2);max-width:340px;line-height:1.6;">
        Your account was opened on another device. TaskFlow allows only one active session at a time.
      </div>
      <button onclick="signOut()" style="background:var(--amber);color:#000;border:none;border-radius:10px;
        padding:12px 28px;font-size:14px;font-weight:700;cursor:pointer;font-family:var(--mono);">
        Sign in again
      </button>
    `;
    document.body.appendChild(b);
  }

  // Poll every 30s to check if session is still ours
  let _pollTimer = null;
  function startSessionPoll(userId) {
    clearInterval(_pollTimer);
    _pollTimer = setInterval(async () => {
      if (!state.currentUser) {
        clearInterval(_pollTimer);
        return;
      }
      const result = await checkSession(userId);
      if (result === 'stolen') {
        clearInterval(_pollTimer);
        showSessionStolenBanner();
      }
    }, 30000);
  }

  function stopSessionPoll() {
    clearInterval(_pollTimer);
  }

  return {
    claimSession,
    checkSession,
    startSessionPoll,
    stopSessionPoll,
    showSessionStolenBanner,
    getDeviceId,
  };
})();

// Hook into signIn to claim session

/* ============================================================
   GOOGLE CALENDAR-STYLE TASK VIEW (GCal View)
   ============================================================ */

function calToggleGcalView() {
  const gcalWrap = $('gcal-view');
  const splitWrap = $('cal-split-wrapper');
  const btn = $('cal-gcal-toggle-btn');
  const isGcal = !gcalWrap.classList.contains('hidden');

  if (isGcal) {
    // Switch back to grid view
    gcalWrap.classList.add('hidden');
    splitWrap.style.display = 'flex';
    btn.textContent = '📅 Task View';
    btn.classList.remove('active');
    $('cal-week-strip')?.classList.remove('hidden');
    localStorage.removeItem('tf_gcal_active');
  } else {
    // Switch to gcal view
    splitWrap.style.display = 'none';
    gcalWrap.classList.remove('hidden');
    btn.textContent = '⬅ Grid View';
    btn.classList.add('active');
    $('cal-week-strip')?.classList.add('hidden');
    // Align anchor date with standard calendar if switching for first time
    gcalState.anchorDate = new Date(calState.year, calState.month, 1);
    renderGcal();
    localStorage.setItem('tf_gcal_active', '1');
  }
}

function setGcalMode(mode) {
  gcalState.mode = mode;
  $$('.gcal-mode-btn').forEach((b) =>
    b.classList.toggle('active', b.dataset.mode === mode),
  );
  renderGcal();
}

function gcalNav(dir) {
  const d = new Date(gcalState.anchorDate);
  if (gcalState.mode === 'day') d.setDate(d.getDate() + dir);
  else if (gcalState.mode === 'week') d.setDate(d.getDate() + dir * 7);
  else d.setMonth(d.getMonth() + dir);
  gcalState.anchorDate = d;
  renderGcal();
}

function gcalGoToday() {
  gcalState.anchorDate = new Date();
  renderGcal();
}

function renderGcal() {
  updateGcalRangeLabel();
  if (gcalState.mode === 'month') {
    $('gcal-grid-wrap').style.display = 'none';
    $('gcal-month-grid').classList.remove('hidden');
    renderGcalMonth();
  } else {
    $('gcal-grid-wrap').style.display = 'flex';
    $('gcal-month-grid').classList.add('hidden');
    renderGcalDayWeek();
  }
}

function updateGcalRangeLabel() {
  const d = gcalState.anchorDate;
  const label = $('gcal-range-label');
  const months = [
    'Jan',
    'Feb',
    'Mar',
    'Apr',
    'May',
    'Jun',
    'Jul',
    'Aug',
    'Sep',
    'Oct',
    'Nov',
    'Dec',
  ];
  const days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  if (gcalState.mode === 'day') {
    label.textContent =
      days[d.getDay()] +
      ' ' +
      d.getDate() +
      ' ' +
      months[d.getMonth()] +
      ' ' +
      d.getFullYear();
  } else if (gcalState.mode === 'week') {
    const mon = getWeekStart(d);
    const sun = new Date(mon);
    sun.setDate(mon.getDate() + 6);
    label.textContent =
      mon.getDate() +
      ' – ' +
      sun.getDate() +
      ' ' +
      months[sun.getMonth()] +
      ' ' +
      sun.getFullYear();
  } else {
    label.textContent = months[d.getMonth()] + ' ' + d.getFullYear();
  }
}

function getWeekStart(date) {
  const d = new Date(date);
  const dow = d.getDay(); // 0=Sun
  const diff = dow === 0 ? -6 : 1 - dow; // shift to Monday
  d.setDate(d.getDate() + diff);
  d.setHours(0, 0, 0, 0);
  return d;
}

function getGcalDays() {
  const d = gcalState.anchorDate;
  if (gcalState.mode === 'day') return [new Date(d)];
  if (gcalState.mode === 'week') {
    const mon = getWeekStart(d);
    return Array.from({ length: 7 }, (_, i) => {
      const day = new Date(mon);
      day.setDate(mon.getDate() + i);
      return day;
    });
  }
  return [];
}

function getGcalTasks() {
  const isElevated =
    state.currentUser.role === 'admin' || state.currentUser.role === 'manager';
  return getTasks().filter((t) => {
    if (t.status === 'cancelled') return false;
    if (!isElevated && t.userId !== state.currentUser.id) return false;
    return true;
  });
}

function getGcalTasksDeduped() {
  const tasks = getGcalTasks();
  const seen = new Map();
  const solo = [];

  tasks.forEach(task => {
    const gid = task.multiGroupId || task.teamId || null;

    if (!gid) {
      solo.push(task);
      return;
    }

    if (seen.has(gid)) {
      const rep = seen.get(gid);
      const user = getUsers().find(u => u.id === task.userId);
      if (user && !rep._assigneeNames.includes(user.name)) {
        rep._assigneeNames.push(user.name);
      }
    } else {
      const user = getUsers().find(u => u.id === task.userId);
      seen.set(gid, {
        ...task,
        _isGroup: true,
        _groupId: gid,
        _assigneeNames: user ? [user.name] : [],
      });
    }
  });

  return [...solo, ...seen.values()];
}

function renderGcalDayWeek() {
  const days = getGcalDays();
  const tasks = getGcalTasksDeduped();
  const H = gcalState.hourHeight;
  const startH = gcalState.startHour;
  const endH = gcalState.endHour;
  const totalHours = endH - startH;
  const totalHeight = totalHours * H;
  const today = toDateStr(new Date());
  const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

  // ── Day headers ──
  const headersEl = $('gcal-day-headers');
  headersEl.innerHTML = '';
  days.forEach((day) => {
    const dStr = toDateStr(day);
    const isToday = dStr === today;
    const isWeekend = day.getDay() === 0 || day.getDay() === 6;
    const el = document.createElement('div');
    el.className =
      'gcal-day-header' +
      (isToday ? ' today-col' : '') +
      (isWeekend ? ' weekend-col' : '');
    el.style.flex = '1';
    el.style.minWidth = '120px';
    el.innerHTML = `<span>${dayNames[day.getDay()]
      }</span><span class="gcal-day-num">${day.getDate()}</span>`;
    headersEl.appendChild(el);
  });

  // ── Time gutter ──
  const gutter = $('gcal-time-gutter');
  gutter.innerHTML = '';
  gutter.style.position = 'relative';
  gutter.style.height = totalHeight + 'px';
  for (let h = startH; h <= endH; h++) {
    const label = document.createElement('div');
    label.className = 'gcal-hour-label';
    label.style.top = (h - startH) * H + 'px';
    label.textContent =
      h === 0
        ? '12 AM'
        : h < 12
          ? h + ' AM'
          : h === 12
            ? '12 PM'
            : h - 12 + ' PM';
    gutter.appendChild(label);
  }

  // ── Columns ──
  const colsEl = $('gcal-columns');
  colsEl.innerHTML = '';
  colsEl.style.height = totalHeight + 'px';
  colsEl.style.display = 'flex';
  colsEl.style.flex = '1';

  // Hour grid lines
  for (let h = startH; h < endH; h++) {
    const line = document.createElement('div');
    line.className = 'gcal-hour-row';
    line.style.top = (h - startH) * H + 'px';
    colsEl.appendChild(line);
  }

  days.forEach((day) => {
    const dStr = toDateStr(day);
    const isToday = dStr === today;
    const isWeekend = day.getDay() === 0 || day.getDay() === 6;

    const col = document.createElement('div');
    col.className =
      'gcal-day-col' +
      (isToday ? ' today-col' : '') +
      (isWeekend ? ' weekend-col' : '');
    col.style.height = totalHeight + 'px';

    const dayTasks = tasks.filter((t) => {
      const ts = new Date(t.start);
      const te = new Date(t.end || t.deadline);
      const dStart = new Date(dStr + 'T00:00:00');
      const dEnd = new Date(dStr + 'T23:59:59');
      return ts <= dEnd && te >= dStart;
    });

    const columns = resolveGcalOverlaps(dayTasks, dStr);

    columns.forEach((group, colIdx) => {
      const colCount = columns.length;
      group.forEach((task) => {
        const block = buildGcalTaskBlock(
          task,
          dStr,
          colIdx,
          colCount,
          H,
          startH,
          endH,
        );
        if (block) col.appendChild(block);
      });
    });

    colsEl.appendChild(col);
  });

  positionGcalNowLine(H, startH, endH);

  const scrollEl = $('gcal-body-scroll');
  const now = new Date();
  const scrollHour = Math.max(startH, now.getHours() - 1);
  scrollEl.scrollTop = (scrollHour - startH) * H;
}

function resolveGcalOverlaps(tasks, dayStr) {
  if (!tasks.length) return [];
  const sorted = tasks
    .slice()
    .sort((a, b) => new Date(a.start) - new Date(b.start));
  const columns = [];

  sorted.forEach((task) => {
    const tStart = new Date(task.start).getTime();
    let placed = false;
    for (const col of columns) {
      const last = col[col.length - 1];
      const lastEnd = new Date(last.end || last.deadline).getTime();
      if (tStart >= lastEnd) {
        col.push(task);
        placed = true;
        break;
      }
    }
    if (!placed) columns.push([task]);
  });

  return columns;
}

function buildGcalTaskBlock(task, dayStr, colIdx, colCount, H, startH, endH) {
  const tStart = new Date(task.start);
  const tEnd = new Date(task.end || task.deadline);

  const visStart = new Date(
    dayStr + 'T' + String(startH).padStart(2, '0') + ':00:00',
  );
  const visEnd = new Date(
    dayStr + 'T' + String(endH).padStart(2, '0') + ':00:00',
  );
  const clampedStart = tStart < visStart ? visStart : tStart;
  const clampedEnd = tEnd > visEnd ? visEnd : tEnd;

  const topFrac =
    clampedStart.getHours() + clampedStart.getMinutes() / 60 - startH;
  const heightFrac =
    clampedEnd.getHours() +
    clampedEnd.getMinutes() / 60 -
    (clampedStart.getHours() + clampedStart.getMinutes() / 60);

  if (heightFrac <= 0) return null;

  const topPx = Math.max(0, topFrac * H);
  const heightPx = Math.max(18, heightFrac * H - 2);

  const pct = 100 / colCount;
  const left = colIdx * pct;
  const width = pct - 1;

  const colours = {
    1: { bg: 'var(--p1)', text: '#fff' },
    2: { bg: 'var(--p2)', text: '#fff' },
    3: { bg: 'var(--p3)', text: '#1a1a2e' },
    4: { bg: 'var(--p4)', text: '#fff' },
    5: { bg: 'var(--p5)', text: '#fff' },
  };
  const c = colours[task.priority] || colours[3];

  const block = document.createElement('div');
  block.className = 'gcal-task-block';
  block.style.top = topPx + 'px';
  block.style.height = heightPx + 'px';
  block.style.left = left + '%';
  block.style.width = width + '%';
  block.style.background = c.bg;
  block.style.color = c.text;

  const startFmt = tStart.toLocaleTimeString('en-GB', {
    hour: '2-digit',
    minute: '2-digit',
  });
  const endFmt = tEnd.toLocaleTimeString('en-GB', {
    hour: '2-digit',
    minute: '2-digit',
  });
  block.innerHTML = `
    <span class="gcal-task-title">${escHtml(task.title)}</span>
    ${heightPx > 32
      ? `<span class="gcal-task-time">${startFmt} – ${endFmt}</span>`
      : ''
    }
  `;

  const isElevated =
    state.currentUser.role === 'admin' || state.currentUser.role === 'manager';
  if (isElevated) {
    block.addEventListener('click', (e) => {
      e.stopPropagation();
      if (task._isGroup && task._groupId) {
        openTeamTaskViewById(task._groupId);
      } else {
        openEditTask(task.id);
      }
    });
  } else {
    block.addEventListener('mouseenter', (e) => showGcalTooltip(task, e));
    block.addEventListener('mousemove', (e) => moveGcalTooltip(e));
    block.addEventListener('mouseleave', () => hideGcalTooltip());
    block.addEventListener('click', (e) => {
      e.stopPropagation();
      showGcalTooltip(task, e);
    });
  }

  if (task._isGroup && task._assigneeNames && task._assigneeNames.length > 0) {
    const names = task._assigneeNames;
    const label = names.length <= 2
      ? names.join(', ')
      : names.slice(0, 2).join(', ') + ' +' + (names.length - 2) + ' more';

    const assigneeLine = document.createElement('span');
    assigneeLine.className = 'gcal-task-time';
    assigneeLine.textContent = '👥 ' + label;
    block.appendChild(assigneeLine);
  }

  return block;
}

function positionGcalNowLine(H, startH, endH) {
  const line = $('gcal-now-line');
  if (!line) return;
  const now = new Date();
  const frac = now.getHours() + now.getMinutes() / 60;
  if (frac < startH || frac > endH) {
    line.style.display = 'none';
    return;
  }
  line.style.display = 'block';
  line.style.top = (frac - startH) * H + 'px';
}

setInterval(() => {
  if (
    !$('gcal-view').classList.contains('hidden') &&
    gcalState.mode !== 'month'
  ) {
    positionGcalNowLine(
      gcalState.hourHeight,
      gcalState.startHour,
      gcalState.endHour,
    );
  }
}, 60000);

function renderGcalMonth() {
  const d = gcalState.anchorDate;
  const year = d.getFullYear();
  const month = d.getMonth();
  const tasks = getGcalTasksDeduped();
  const today = toDateStr(new Date());
  const grid = $('gcal-month-grid');
  grid.innerHTML = '';

  ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'].forEach((name) => {
    const h = document.createElement('div');
    h.style.cssText =
      'padding:6px;font-size:10px;font-weight:700;color:var(--text3);text-align:center;border-bottom:1px solid var(--border);font-family:var(--mono);text-transform:uppercase;letter-spacing:0.06em;';
    h.textContent = name;
    grid.appendChild(h);
  });

  const firstDay = new Date(year, month, 1);
  const startOffset = firstDay.getDay() === 0 ? 6 : firstDay.getDay() - 1;
  for (let i = 0; i < startOffset; i++) {
    const blank = document.createElement('div');
    blank.className = 'gcal-month-day';
    blank.style.background = 'rgba(0,0,0,0.1)';
    grid.appendChild(blank);
  }

  const daysInMonth = new Date(year, month + 1, 0).getDate();
  for (let day = 1; day <= daysInMonth; day++) {
    const dStr = toDateStr(new Date(year, month, day));
    const dow = new Date(year, month, day).getDay();
    const isToday = dStr === today;
    const isWeekend = dow === 0 || dow === 6;

    const cell = document.createElement('div');
    cell.className =
      'gcal-month-day' +
      (isToday ? ' today-col' : '') +
      (isWeekend ? ' weekend-col' : '');

    const numEl = document.createElement('div');
    numEl.className = 'gcal-month-day-num';
    numEl.textContent = day;
    cell.appendChild(numEl);

    const dayTasks = tasks.filter((t) => {
      const ts = new Date(t.start);
      const te = new Date(t.end || t.deadline);
      const ds = new Date(dStr + 'T00:00:00');
      const de = new Date(dStr + 'T23:59:59');
      return ts <= de && te >= ds;
    });

    dayTasks.slice(0, 4).forEach((task) => {
      const colours = {
        1: 'var(--p1)',
        2: 'var(--p2)',
        3: 'var(--p3)',
        4: 'var(--p4)',
        5: 'var(--p5)',
      };
      const chip = document.createElement('div');
      chip.className = 'gcal-month-task-chip';
      chip.style.background = colours[task.priority] || colours[3];
      chip.style.color = task.priority === 3 ? '#1a1a2e' : '#fff';
      chip.textContent = task.title;

      const isElevated =
        state.currentUser.role === 'admin' ||
        state.currentUser.role === 'manager';
      if (isElevated) {
        chip.addEventListener('click', (e) => {
          e.stopPropagation();
          if (task._isGroup && task._groupId) {
            openTeamTaskViewById(task._groupId);
          } else {
            openEditTask(task.id);
          }
        });
      } else {
        chip.addEventListener('mouseenter', (e) => showGcalTooltip(task, e));
        chip.addEventListener('mousemove', (e) => moveGcalTooltip(e));
        chip.addEventListener('mouseleave', () => hideGcalTooltip());
      }
      cell.appendChild(chip);
    });

    if (dayTasks.length > 4) {
      const more = document.createElement('div');
      more.style.cssText = 'font-size:10px;color:var(--text3);padding:1px 4px;';
      more.textContent = '+' + (dayTasks.length - 4) + ' more';
      cell.appendChild(more);
    }

    grid.appendChild(cell);
  }
}

function showGcalTooltip(task, e) {
  let tip = $('gcal-tooltip');
  if (!tip) {
    tip = document.createElement('div');
    tip.id = 'gcal-tooltip';
    document.body.appendChild(tip);
  }
  const startFmt = new Date(task.start).toLocaleString('en-GB', {
    day: '2-digit',
    month: 'short',
    hour: '2-digit',
    minute: '2-digit',
  });
  const endFmt = new Date(task.end || task.deadline).toLocaleString('en-GB', {
    day: '2-digit',
    month: 'short',
    hour: '2-digit',
    minute: '2-digit',
  });
  const pLabels = {
    1: '🔴 P1',
    2: '🟠 P2',
    3: '🟡 P3',
    4: '🔵 P4',
    5: '⚪ P5',
  };
  tip.innerHTML = `
    <div style="font-weight:800;font-size:13px;margin-bottom:6px;">${escHtml(
    task.title,
  )}</div>
    <div style="color:var(--text3);font-size:11px;margin-bottom:2px;">🕐 ${startFmt}</div>
    <div style="color:var(--text3);font-size:11px;margin-bottom:6px;">⏱ ends ${endFmt}</div>
    <div style="font-size:11px;">${pLabels[task.priority] || ''} · ${escHtml(
    task.requestor || '',
  )}</div>
    ${task.status === 'done'
      ? '<div style="color:#10B981;font-size:11px;margin-top:4px;">✅ Done</div>'
      : ''
    }
  `;
  if (task._isGroup && task._assigneeNames && task._assigneeNames.length > 0) {
    const namesHtml = task._assigneeNames.map(n =>
      `<span style="display:inline-block;background:var(--bg3);border-radius:4px;padding:1px 6px;font-size:10px;margin:2px 2px 0 0;">${n}</span>`
    ).join('');
    tip.innerHTML += `
      <div style="margin-top:6px;border-top:1px solid var(--border);padding-top:6px;">
        <div style="font-size:10px;color:var(--text3);margin-bottom:4px;text-transform:uppercase;letter-spacing:0.06em;font-weight:700;">Assigned to</div>
        <div>${namesHtml}</div>
      </div>
    `;
  }
  tip.classList.add('visible');
  moveGcalTooltip(e);
}

function moveGcalTooltip(e) {
  const tip = $('gcal-tooltip');
  if (!tip) return;
  const x = e.clientX + 14;
  const y = e.clientY - 10;
  const overflowX = x + 270 > window.innerWidth;
  tip.style.left = (overflowX ? e.clientX - 274 : x) + 'px';
  tip.style.top = Math.max(0, y) + 'px';
}

function openTeamTaskViewById(groupId) {
  const groupTasks = getTasks().filter(t =>
    (t.multiGroupId || t.teamId) === groupId
  );
  if (!groupTasks.length) return;

  const rep = groupTasks[0];
  const users = groupTasks.map(t => getUsers().find(u => u.id === t.userId)).filter(Boolean);

  const pLabels = { 1: '🔴 P1', 2: '🟠 P2', 3: '🟡 P3', 4: '🔵 P4', 5: '⚪ P5' };
  const startFmt = new Date(rep.start).toLocaleString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
  const endFmt = new Date(rep.end || rep.deadline).toLocaleString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });

  const body = $('team-tasks-modal-body');
  body.innerHTML = `
    <div style="padding:16px;border-bottom:1px solid var(--border);">
      <div style="font-size:16px;font-weight:800;margin-bottom:8px;">${rep.title}</div>
      <div style="font-size:12px;color:var(--text2);margin-bottom:4px;">🕐 ${startFmt} → ${endFmt}</div>
      <div style="font-size:12px;color:var(--text2);margin-bottom:12px;">${pLabels[rep.priority] || ''} · Requested by ${rep.requestor || '—'}</div>
      <div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:0.06em;margin-bottom:8px;">Assigned to (${users.length})</div>
      <div style="display:flex;flex-wrap:wrap;gap:6px;">
        ${users.map(u => `
          <div style="background:var(--bg3);border:1px solid var(--border);border-radius:8px;padding:6px 12px;font-size:12px;font-weight:600;">
            ${u.name}
            <span style="font-size:10px;color:var(--text3);margin-left:4px;">${u.role}</span>
          </div>
        `).join('')}
      </div>
    </div>
    ${rep.description && rep.description.length ? `
      <div style="padding:14px 16px;">
        <div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:0.06em;margin-bottom:8px;">Checklist</div>
        ${rep.description.map(item => `
          <div style="display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:1px solid rgba(255,255,255,0.04);font-size:12px;">
            <span>${item.done ? '✅' : '⬜'}</span>
            <span style="color:${item.done ? 'var(--text3)' : 'var(--text)'};text-decoration:${item.done ? 'line-through' : 'none'}">${item.text}</span>
          </div>
        `).join('')}
      </div>
    ` : ''}
  `;

  openModal('team-tasks-modal');
}

function hideGcalTooltip() {
  const tip = $('gcal-tooltip');
  if (tip) tip.classList.remove('visible');
}

function refreshAllCalendarViews() {
  renderCalendarDebounced();
  if ($('gcal-view') && !$('gcal-view').classList.contains('hidden'))
    renderGcal();
}

async function initializeApp() {
  /* ── DOM element cache ──────────────────────────────────────
     Populated once so the rest of the app can read EL.xxx
     instead of calling getElementById() on every render.
     To add a new cached element: add a line here.
  ────────────────────────────────────────────────────────── */
  window.EL = {
    app: $('app'),
    loginScreen: $('login-screen'),
    loginBtn: $('login-btn'),
    loginUser: $('login-username'),
    loginPw: $('login-password'),
    loginError: $('login-error'),
    saPanel: $('sa-panel'),
    taskBoard: $('task-board'),
    taskModal: $('task-modal'),
    cancelModal: $('cancel-modal'),
    confirmModal: $('confirm-modal'),
    calView: $('calendar-view'),
    calNamesPanel: $('cal-names-panel'),
    calDatesPanel: $('cal-dates-panel'),
    calNamesTable: $('cal-names-table'),
    calDatesTable: $('cal-dates-table'),
    calMonthLabel: $('cal-month-label'),
    addTaskFab: $('add-task-fab'),
    breadcrumb: $('breadcrumb'),
    boardFilterBar: $('board-filter-bar'),
    boardSearch: $('board-search'),
    headerRole: $('header-role'),
    headerUsername: $('header-username'),
    notifBtn: $('notif-btn'),
    globalOverlay: $('global-overlay'),
    pushContainer: $('push-notif-container'),
    mobNavAdd: $('mob-nav-add'),
  };

  // Wire up login event listeners
  EL.loginBtn.addEventListener('click', signIn);
  EL.loginPw.addEventListener('keydown', function (e) {
    if (e.key === 'Enter') signIn();
  });
  EL.loginUser.addEventListener('keydown', function (e) {
    if (e.key === 'Enter') EL.loginPw.focus();
  });

  // Wire up modal overlay click-to-close
  document.querySelectorAll('.modal-overlay').forEach((overlay) => {
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) closeModal(overlay.id);
    });
  });

  // Mark body as ready so app panel becomes visible (prevents FOUC)
  document.body.classList.add('app-ready');

  /* ============================================================
     AUTO-RESTORE SESSION ON RELOAD
     ============================================================ */
  const token = localStorage.getItem('tf_token');
  const uid = localStorage.getItem('tf_uid');
  const role = localStorage.getItem('tf_role');
  const cid = localStorage.getItem('tf_cid');
  if (!token || !uid || !role) return; // no saved session

  try {
    // Re-fetch the user record to confirm session is still valid
    const rows = await SB.select(
      'tf_users',
      `id=eq.${uid}&select=id,name,username,role,company_id&limit=1`,
    );
    if (!rows || rows.length === 0) {
      API.clearToken();
      return;
    }
    const u = rows[0];
    state.currentUser = {
      id: u.id,
      name: u.name,
      username: u.username,
      role: u.role,
      companyId: u.company_id,
    };

    (EL.loginScreen || document.getElementById('login-screen')).classList.add(
      'hidden',
    );

    if (role === 'superadmin') {
      document.getElementById('sa-panel').classList.remove('hidden');
      await loadSACompanies();
    } else {
      document.getElementById('app').classList.remove('hidden');
      setupHeader();
      requestNotifPermission();
      registerSW();
      startPolling();
      startRealtimeNotifs();
      await loadAll();
      // Restore saved view or default to home
      var savedView = _getSavedView();
      // Always show calendar button for regular users
      if (role !== 'admin' && role !== 'manager') {
        const calBtn = document.getElementById('my-cal-btn');
        if (calBtn) calBtn.classList.remove('hidden');
      }
      if (
        savedView &&
        savedView !== 'admin-home' &&
        savedView !== 'worker-home'
      ) {
        if (savedView === 'user-tasks') {
          var savedUid = localStorage.getItem('tf_target_uid');
          if (savedUid) {
            state.targetUserId = savedUid;
            showView('user-tasks');
          } else {
            showView(
              role === 'admin' || role === 'manager' ? 'user-list' : 'my-tasks',
            );
          }
        } else {
          showView(savedView);
        }
      } else if (role === 'admin' || role === 'manager') {
        showView('admin-home');
      } else {
        showView('worker-home');
      }
      // Check & claim session after restore (single-session enforcement)
      if (typeof window._sessionClaimOnRestore === 'function') {
        setTimeout(window._sessionClaimOnRestore, 500);
      }
      updateLeaveRequestsBadge();
    }
  } catch (e) {
    // Session invalid or network error — stay on login screen
    API.clearToken();
  }
}
const _origSignIn = signIn;
window.signIn = async function () {
  await _origSignIn.apply(this, arguments);
  // After successful login, claim session
  if (state.currentUser && state.currentUser.role !== 'superadmin') {
    await Session.claimSession(state.currentUser.id);
    Session.startSessionPoll(state.currentUser.id);
  }
};

// Hook into session restore to claim + start polling
const _origTryRestore = window._tryRestoreSession;
// We patch the startup by adding a post-restore hook via a flag
window._sessionClaimOnRestore = async function () {
  if (state.currentUser && state.currentUser.role !== 'superadmin') {
    const result = await Session.checkSession(state.currentUser.id);
    if (result === 'stolen') {
      // Someone else is logged in — boot us to login
      API.clearToken();
      document.getElementById('app').classList.add('hidden');
      (
        EL.loginScreen || document.getElementById('login-screen')
      ).classList.remove('hidden');
      toast(
        'Your session was opened on another device. Please sign in.',
        'warning',
      );
      return;
    }
    await Session.claimSession(state.currentUser.id);
    Session.startSessionPoll(state.currentUser.id);
  }
};

// signOut should also stop session poll
const _origSignOut = signOut;
window.signOut = async function () {
  Session.stopSessionPoll();
  await _origSignOut.apply(this, arguments);
};

// Expose initializeApp globally so bootstrap script can always find it
if (typeof window !== 'undefined') window.initializeApp = initializeApp;

window.debugEmailHTML = function () {
  const user = { name: 'Alice', emailNotif: true, email: 'test@example.com' };
  const task = { title: 'Test Task', requestor: 'Bob', priority: 2, description: [{ text: 'Do it', checked: false }] };
  const action = 'assigned';
  const html = composeTaskEmail(user, task, action, 'My Company');
  console.log('--- HTML OUTPUT ---');
  console.log(html);

  // also test the insert
  sendEmailViaAPI(user.email, 'Test', html).then(() => {
    console.log('[DEBUG] Insert test fired. Check table.');
  });
};
