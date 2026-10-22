const DEADLINE = Date.parse("2026-10-22T12:00:00+03:00");
const STORE_KEY = "applications";
const TRAINING = { passed: "Прошёл", missed: "Не прошёл" };
const PROGRESS = {
  assignment_sent: "Отправлено задание",
  bitrix_entered: "Зашёл в Битрикс",
  project_started: "Начал делать проект",
  project_done: "Сделал проект",
};
const FIELDS = {
  name: [2, 120],
  position: [2, 160],
  process: [0, 500],
};
const ALLOWED_ORIGINS = [
  "https://nikisheva.github.io",
  "http://127.0.0.1:8765",
  "http://localhost:8765",
];

function corsHeaders(request) {
  const origin = request.headers.get("Origin") || "";
  const headers = { Vary: "Origin" };
  if (ALLOWED_ORIGINS.some((item) => origin === item || origin.startsWith(item + "/"))) {
    headers["Access-Control-Allow-Origin"] = origin;
    headers["Access-Control-Allow-Methods"] = "GET, POST, PATCH, DELETE, OPTIONS";
    headers["Access-Control-Allow-Headers"] = "Content-Type, X-Admin-Token, bypass-tunnel-reminder";
  }
  return headers;
}

function json(request, code, payload) {
  return new Response(JSON.stringify(payload), {
    status: code,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
      ...corsHeaders(request),
    },
  });
}

function cleanText(value, limit) {
  return String(value || "").split(/\s+/).join(" ").trim().slice(0, limit);
}

function nowMsk() {
  return new Date().toLocaleString("sv-SE", { timeZone: "Europe/Moscow" }).replace(" ", "T") + "+03:00";
}

function isAdmin(request, env) {
  const expected = String(env.ADMIN_TOKEN || "").trim();
  const given = request.headers.get("X-Admin-Token") || "";
  if (!expected || given.length !== expected.length) return false;
  let diff = 0;
  for (let i = 0; i < expected.length; i += 1) {
    diff |= expected.charCodeAt(i) ^ given.charCodeAt(i);
  }
  return diff === 0;
}

function validate(body) {
  if (Date.now() > DEADLINE) {
    return { error: "Приём заявок закрыт. Срок — 22 октября 2026, 12:00 мск." };
  }
  if (String(body.company_website || "").trim()) {
    return { honeypot: true };
  }
  const item = {};
  for (const [key, [low, high]] of Object.entries(FIELDS)) {
    const value = cleanText(body[key], high);
    if (value.length < low) return { error: "Заполните все обязательные поля." };
    item[key] = value;
  }
  const training = String(body.training || "");
  if (!TRAINING[training]) return { error: "Укажите, прошли обучение или нет." };
  item.training = training;
  return { item };
}

async function loadAll(env) {
  const raw = await env.APPS.get(STORE_KEY);
  if (!raw) return [];
  try {
    const data = JSON.parse(raw);
    return Array.isArray(data) ? data : [];
  } catch {
    return [];
  }
}

async function saveAll(env, rows) {
  await env.APPS.put(STORE_KEY, JSON.stringify(rows));
}

function upsert(rows, item) {
  const stamp = nowMsk();
  const found = rows.find((row) => (row.name || "").toLowerCase() === item.name.toLowerCase());
  if (found) {
    Object.assign(found, item, { updated_at: stamp });
    return { row: found, updated: true };
  }
  const row = {
    id: crypto.randomUUID().replace(/-/g, "").slice(0, 8),
    created_at: stamp,
    updated_at: stamp,
    ...item,
  };
  rows.push(row);
  return { row, updated: false };
}

function toCsv(rows) {
  const header = ["id", "создана", "обновлена", "имя", "должность", "обучение", "процесс", ...Object.values(PROGRESS)];
  const lines = [header.join(",")];
  for (const row of rows) {
    const values = [
      row.id,
      row.created_at,
      row.updated_at,
      row.name,
      row.position,
      TRAINING[row.training] || "",
      row.process,
      ...Object.keys(PROGRESS).map((key) => (row[key] ? "да" : "")),
    ].map((value) => `"${String(value || "").replace(/"/g, '""')}"`);
    lines.push(values.join(","));
  }
  return "\uFEFF" + lines.join("\n");
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const path = url.pathname;

    if (request.method === "OPTIONS") {
      if (path !== "/api/health" && path !== "/api/applications" && path !== "/api/applications.csv") {
        return json(request, 404, { error: "Страница не найдена." });
      }
      return new Response(null, { status: 204, headers: corsHeaders(request) });
    }

    if (request.method === "GET" && path === "/api/health") {
      return json(request, 200, {
        ok: true,
        open: Date.now() <= DEADLINE,
        deadline: "2026-10-22T12:00:00+03:00",
        auth: "none",
      });
    }

    if (request.method === "GET" && (path === "/api/applications" || path === "/api/applications.csv")) {
      if (!isAdmin(request, env)) {
        return json(request, 403, { error: "Список заявок доступен только организатору." });
      }
      const rows = (await loadAll(env)).slice().sort((a, b) => String(a.created_at).localeCompare(String(b.created_at)));
      if (path.endsWith(".csv")) {
        const raw = toCsv(rows);
        return new Response(raw, {
          status: 200,
          headers: {
            "Content-Type": "text/csv; charset=utf-8",
            "Content-Disposition": "attachment; filename=ai-challenge-applications.csv",
            "Cache-Control": "no-store",
            ...corsHeaders(request),
          },
        });
      }
      return json(request, 200, { items: rows, count: rows.length });
    }

    if ((request.method === "PATCH" || request.method === "DELETE") && path === "/api/applications") {
      if (!isAdmin(request, env)) {
        return json(request, 403, { error: "Список заявок доступен только организатору." });
      }
      let body;
      try {
        body = await request.json();
      } catch {
        return json(request, 400, { error: "Не удалось прочитать заявку." });
      }
      const appId = String((body && body.id) || "").trim();
      if (!appId) return json(request, 400, { error: "Не указана заявка." });
      const rows = await loadAll(env);
      const index = rows.findIndex((row) => String(row.id || "") === appId);
      if (index < 0) return json(request, 404, { error: "Заявка не найдена." });
      if (request.method === "DELETE") {
        const removed = rows.splice(index, 1)[0];
        await saveAll(env, rows);
        return json(request, 200, { ok: true, id: removed.id });
      }
      const flags = {};
      for (const key of Object.keys(PROGRESS)) {
        if (Object.prototype.hasOwnProperty.call(body, key)) flags[key] = Boolean(body[key]);
      }
      if (!Object.keys(flags).length) return json(request, 400, { error: "Нечего обновлять." });
      Object.assign(rows[index], flags, { updated_at: nowMsk() });
      await saveAll(env, rows);
      return json(request, 200, { ok: true, id: rows[index].id, item: rows[index] });
    }

    if (request.method === "POST" && path === "/api/applications") {
      let body;
      try {
        body = await request.json();
      } catch {
        return json(request, 400, { error: "Не удалось прочитать заявку." });
      }
      if (!body || typeof body !== "object") {
        return json(request, 400, { error: "Не удалось прочитать заявку." });
      }
      const checked = validate(body);
      if (checked.error) {
        return json(request, checked.error.includes("закрыт") ? 403 : 400, { error: checked.error });
      }
      if (checked.honeypot) {
        return json(request, 201, { ok: true, id: "00000000", updated: false });
      }
      const rows = await loadAll(env);
      const { row, updated } = upsert(rows, checked.item);
      await saveAll(env, rows);
      return json(request, updated ? 200 : 201, { ok: true, id: row.id, updated });
    }

    return json(request, 404, { error: "Страница не найдена." });
  },
};
