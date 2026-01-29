const PHASES = ["A embarcar", "Em Trânsito", "A Registrar", "A Desembaraçar", "A Carregar"];
const TEAMS_FILE_NAME = "Equipes.json";
const WEBHOOK_URL = "https://n8n.grupobgr.com.br/webhook/b52163be-ad8d-4264-9f2c-74dd8bf2cf2e";
const CHATBOT_WEBHOOK_URL = "https://n8n.grupobgr.com.br/webhook/aed320b0-92db-4922-992f-f4854d186ec5/chat";

// Fallback para ambientes que bloqueiam fetch (ex.: abrir via file://).
const DEFAULT_TEAMS = {
  "Operacional 1": ["TATIANI", "YAGOCORREA"],
  "Operacional 2": ["PRISCILA", "JULIABATISTA"],
  "Operacional 3": ["CHARLIANEBATISTA", "THAMIRYS", "MARCOSAZEVEDO"],
  "Operacional 4": ["JULIANAAPRIGIO"],
  "Operacional 5": ["BRUNOFREIRES", "EDUARDA"],
  "Operacional 6": ["INGRID"],
  "Operacional 7": ["POLIANA"],
  "Operacional 8": ["DIANEVITOR"],
};

// Os dados agora vêm do Excel (raiz do projeto) ou do upload manual do arquivo.
let processes = [];
let teamByMemberKey = new Map(); // memberKey -> teamName
let apiReqSeq = 0;
let apiAbort = null;
let isRefreshing = false;

// Registro de mensagens (Logs)
let messageLogs = [];
let recurrentAlertInterval = null;
let recurrentAlertTimeout = null;

const els = {
  list: document.getElementById("processList"),
  search: document.getElementById("searchInput"),
  filter: document.getElementById("phaseFilter"),
  teamFilter: document.getElementById("teamFilter"),
  openFrom: document.getElementById("openFrom"),
  openTo: document.getElementById("openTo"),
  refresh: document.getElementById("refreshBtn"),
  reset: document.getElementById("resetBtn"),
  totalCount: document.getElementById("totalCount"),
  resultsMeta: document.getElementById("resultsMeta"),
  notificationContainer: document.getElementById("notificationContainer"),
};

// Sistema de Notificações
function showAlert(message, type = "info", duration = 4000) {
  if (!els.notificationContainer) return;

  const notification = document.createElement("div");
  notification.className = `notification notification--${type}`;
  
  notification.innerHTML = `
    <div class="notification__content">${message}</div>
    <button class="notification__close" aria-label="Fechar">&times;</button>
  `;

  els.notificationContainer.appendChild(notification);

  const closeFn = () => {
    notification.classList.add("notification--closing");
    notification.addEventListener("animationend", () => {
      notification.remove();
      // Se for um alerta de aviso (recorrente), para a recorrência ao clicar no X
      if (type === "warning") {
        stopRecurrentAlert();
      }
    }, { once: true });
  };

  notification.querySelector(".notification__close").onclick = closeFn;

  if (duration > 0) {
    setTimeout(closeFn, duration);
  }
}

// estado de expansão por cliente
const groupExpanded = Object.create(null);

function norm(s) {
  return String(s ?? "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function memberKey(s) {
  // Equipes.json usa nomes tipo "JULIABATISTA" (sem espaços). Normaliza para comparar.
  return norm(s).replace(/[^a-z0-9]/g, "").toUpperCase();
}

function buildTeamIndex(teamsObj) {
  const idx = new Map();
  for (const [teamName, members] of Object.entries(teamsObj || {})) {
    if (!Array.isArray(members)) continue;
    for (const m of members) {
      const k = memberKey(m);
      if (!k) continue;
      if (!idx.has(k)) idx.set(k, String(teamName));
    }
  }
  return idx;
}

function applyTeamsToProcesses(list) {
  const arr = Array.isArray(list) ? list : processes;
  for (const p of arr) {
    const resp = String(p?.responsavel || "").trim();

    // tenta casar por partes (caso venha "NOME1 / NOME2", etc.)
    const parts = resp
      ? resp
          .split(/[\/,;|]+/g)
          .map((x) => x.trim())
          .filter(Boolean)
      : [];

    let equipe = "";
    for (const part of parts) {
      const k = memberKey(part);
      const hit = teamByMemberKey.get(k);
      if (hit) {
        equipe = hit;
        break;
      }
    }

    if (!equipe) {
      const kAll = memberKey(resp);
      equipe = teamByMemberKey.get(kAll) || "";

      // match “contém” (ex.: e-mail/strings maiores) e “contido” (ex.: abreviações)
      if (!equipe && kAll) {
        for (const [memberK, teamName] of teamByMemberKey.entries()) {
          if (!memberK) continue;
          if (kAll.includes(memberK) || memberK.includes(kAll)) {
            equipe = teamName;
            break;
          }
        }
      }
    }

    // Agrupamento deve seguir APENAS o Equipes.json
    p.equipe = equipe || "Sem equipe";
  }
}

function setHint() {}

function clamp(n, min, max) {
  return Math.max(min, Math.min(max, n));
}

function phaseIndex(phaseName) {
  const idx = PHASES.findIndex((p) => p === phaseName);
  return idx === -1 ? 0 : idx;
}

function getTransitoStatus(process) {
  // Verifica se está em trânsito: tem embarque, não tem chegada
  const hasEmbarque = Boolean(process.dtaEmbarqueBl && isValidIsoDateStr(process.dtaEmbarqueBl));
  const hasChegada = Boolean(process.dtaChegadaBl && isValidIsoDateStr(process.dtaChegadaBl));
  const hasRegistro = Boolean(process.registroDi && String(process.registroDi).trim().length > 0);

  if (hasEmbarque && !hasChegada) {
    return hasRegistro ? "registrado" : "nao_registrado";
  }
  return null;
}

function buildPhaseFilterOptions() {
  const opts = [
    { value: "all", label: "Todas" },
    ...PHASES.map((p) => ({ value: p, label: p })),
  ];

  els.filter.innerHTML = opts
    .map((o) => `<option value="${escapeHtml(o.value)}">${escapeHtml(o.label)}</option>`)
    .join("");
}

function escapeHtml(str) {
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function formatDateBr(iso) {
  if (!iso) return "";
  const d = new Date(`${iso}T00:00:00`);
  if (Number.isNaN(d.getTime())) return String(iso);
  return d.toLocaleDateString("pt-BR");
}

function titleCase(s) {
  const str = String(s ?? "").trim();
  if (!str) return "";
  return str.charAt(0).toUpperCase() + str.slice(1);
}

function isoFromDateParts(y, m, d) {
  if (!y || !m || !d) return "";
  const mm = String(m).padStart(2, "0");
  const dd = String(d).padStart(2, "0");
  return `${y}-${mm}-${dd}`;
}

function isoFromAnyDateValue(v) {
  if (!v) return "";
  // Date (quando cellDates = true)
  if (v instanceof Date && !Number.isNaN(v.getTime())) {
    return isoFromDateParts(v.getFullYear(), v.getMonth() + 1, v.getDate());
  }

  // Excel serial (número)
  if (typeof v === "number" && window.XLSX?.SSF?.parse_date_code) {
    const parsed = window.XLSX.SSF.parse_date_code(v);
    if (parsed?.y && parsed?.m && parsed?.d) {
      return isoFromDateParts(parsed.y, parsed.m, parsed.d);
    }
  }

  // String em formatos comuns (dd/mm/yyyy, yyyy-mm-dd)
  const s = String(v).trim();
  const m1 = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(s);
  if (m1) return isoFromDateParts(Number(m1[3]), Number(m1[2]), Number(m1[1]));
  const m2 = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (m2) return isoFromDateParts(Number(m2[1]), Number(m2[2]), Number(m2[3]));

  // fallback: tenta Date.parse
  const dt = new Date(s);
  if (!Number.isNaN(dt.getTime())) {
    return isoFromDateParts(dt.getFullYear(), dt.getMonth() + 1, dt.getDate());
  }
  return "";
}

function isValidIsoDateStr(s) {
  const str = String(s || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(str)) return false;
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(str);
  if (!m) return false;
  const y = Number(m[1]);
  const mo = Number(m[2]);
  const da = Number(m[3]);
  if (!y || mo < 1 || mo > 12 || da < 1 || da > 31) return false;
  // valida sem efeito de fuso horário
  const d = new Date(Date.UTC(y, mo - 1, da));
  if (Number.isNaN(d.getTime())) return false;
  return (
    d.getUTCFullYear() === y &&
    d.getUTCMonth() + 1 === mo &&
    d.getUTCDate() === da
  );
}

function sanitizeDateRange(from, to) {
  const f = isValidIsoDateStr(from) ? String(from).trim() : "";
  const t = isValidIsoDateStr(to) ? String(to).trim() : "";
  if (f && t && f > t) return { from: t, to: f };
  return { from: f, to: t };
}

function flashInvalidInput(el) {
  if (!el?.classList) return;
  el.classList.add("field__input--invalid");
  window.setTimeout(() => el.classList.remove("field__input--invalid"), 1200);
}

function syncDateInputs() {
  const rawFrom = els.openFrom?.value || "";
  const rawTo = els.openTo?.value || "";
  const r = sanitizeDateRange(rawFrom, rawTo);

  // se usuário tentou algo inválido, limpa e dá feedback visual
  if (rawFrom && !r.from && els.openFrom) {
    els.openFrom.value = "";
    flashInvalidInput(els.openFrom);
  }
  if (rawTo && !r.to && els.openTo) {
    els.openTo.value = "";
    flashInvalidInput(els.openTo);
  }

  // se vier invertido, troca no input para não permitir período inválido
  if (r.from && r.to) {
    if (els.openFrom && els.openFrom.value !== r.from) els.openFrom.value = r.from;
    if (els.openTo && els.openTo.value !== r.to) els.openTo.value = r.to;
  }

  // trava seleção (quando possível) com min/max
  if (els.openTo) els.openTo.min = r.from || "";
  if (els.openFrom) els.openFrom.max = r.to || "";

  // "De" e "Até" obrigatórios para requisitar na API
  if (els.refresh) {
    const ok = Boolean(r.from && r.to);
    const isLoading = els.refresh.dataset?.loading === "true";
    const operacional = els.teamFilter?.value || "all";
    els.refresh.disabled = isLoading ? true : !ok;
    els.refresh.title = ok ? buildWebhookUrl({ from: r.from, to: r.to, operacional }) : "Informe Abertura (De) e (Até) para atualizar.";
  }

  return r;
}

function pickFirstHeader(headersNormToOrig, aliases) {
  for (const a of aliases) {
    const k = norm(a);
    if (headersNormToOrig[k]) return headersNormToOrig[k];
  }
  return "";
}

function buildHeadersMap(headers) {
  const map = Object.create(null);
  for (const h of headers) {
    const key = norm(h).replace(/\s+/g, " ").trim();
    if (!key) continue;
    // mantém o primeiro
    if (!map[key]) map[key] = h;
  }
  return map;
}

function matchPhaseLabel(value) {
  const v = norm(value);
  const found = PHASES.find((p) => norm(p) === v);
  return found || "";
}

function statusToPhase(statusValue) {
  const raw = String(statusValue ?? "").trim();
  if (!raw) return "";

  // se já vier igual ao label, usa direto
  const exact = matchPhaseLabel(raw);
  if (exact) return exact;

  const v = norm(raw);
  if (v.includes("embar")) return "A embarcar";
  // Verifica "transito" ANTES de "registr" para capturar "Em transito já registrado" e "Em transito não registrado"
  if (v.includes("transit") || v.includes("transito")) return "Em Trânsito";
  if (v.includes("registr")) return "A Registrar";
  if (v.includes("desembar") || v.includes("desembara")) return "A Desembaraçar";
  if (v.includes("carreg")) return "A Carregar";

  return "";
}

function findPhaseDateHeader(headersNormToOrig, phaseLabel) {
  const ph = norm(phaseLabel);
  const candidates = Object.keys(headersNormToOrig).filter((h) => h.includes(ph));
  if (!candidates.length) return "";
  const preferred = candidates.find((h) => /\b(data|dt)\b/.test(h));
  return headersNormToOrig[preferred || candidates[0]];
}

function sheetToRowsWithUniqueHeaders(ws) {
  // Lê como matriz para não perder colunas com cabeçalho repetido (ex.: STATUS, STATUS)
  const aoa = window.XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (!aoa?.length) return [];

  const requiredKeys = ["processo", "ref_cliente", "cliente", "responsavel", "abertura", "modal", "origem", "status"];

  function scoreHeaderRow(row) {
    const cells = (row || []).map((c) => norm(c).trim());
    let score = 0;
    for (const rk of requiredKeys) {
      if (cells.some((x) => x === rk || x.includes(rk))) score++;
    }
    return score;
  }

  let headerRowIdx = 0;
  let bestScore = scoreHeaderRow(aoa[0]);
  const scanLimit = Math.min(aoa.length, 25);
  for (let i = 1; i < scanLimit; i++) {
    const s = scoreHeaderRow(aoa[i]);
    if (s > bestScore) {
      bestScore = s;
      headerRowIdx = i;
    }
  }

  const rawHeaders = (aoa[headerRowIdx] || []).map((h) => String(h ?? "").trim());
  const used = Object.create(null);
  const headers = rawHeaders.map((h, idx) => {
    const base = h || `COL_${idx + 1}`;
    const key = norm(base) || `col_${idx + 1}`;
    used[key] = (used[key] || 0) + 1;
    return used[key] === 1 ? base : `${base}_${used[key]}`;
  });

  const rows = [];
  for (let i = headerRowIdx + 1; i < aoa.length; i++) {
    const rowArr = aoa[i] || [];
    const rowObj = {};
    for (let c = 0; c < headers.length; c++) {
      rowObj[headers[c]] = rowArr[c] ?? "";
    }
    rows.push(rowObj);
  }
  return rows;
}

function mapRowsToProcesses(rows) {
  if (!rows?.length) return [];
  const headers = Object.keys(rows[0] || {});
  const headersNormToOrig = buildHeadersMap(headers);

  const idH = pickFirstHeader(headersNormToOrig, [
    "id",
    "processo",
    "processo ",
    "processo id",
    "po",
    "pedido",
    "pedido de compra",
    "numero",
    "número",
  ]);
  const refClienteH = pickFirstHeader(headersNormToOrig, [
    "ref_cliente",
    "ref cliente",
    "ref. cliente",
    "referencia cliente",
    "referência cliente",
    "referencia do cliente",
    "referência do cliente",
  ]);
  const clienteNomeH = pickFirstHeader(headersNormToOrig, ["cliente", "client"]);
  const fornecedorH = pickFirstHeader(headersNormToOrig, ["fornecedor", "supplier", "exportador", "vendedor"]);
  const respH = pickFirstHeader(headersNormToOrig, ["responsavel", "responsável", "responsavel(a)", "responsável(a)"]);
  const aberturaH = pickFirstHeader(headersNormToOrig, [
    "abertura",
    "data abertura",
    "dt abertura",
    "data_abertura",
    "dt_abertura",
    "data de abertura",
    "data_de_abertura",
    "data_abertura_",
    "dataabertura",
  ]);
  const modalH = pickFirstHeader(headersNormToOrig, ["modal", "transporte", "modalidade"]);
  const origemH = pickFirstHeader(headersNormToOrig, ["origem", "orig"]);
  const faseH = pickFirstHeader(headersNormToOrig, ["fase", "fase atual", "fase_atual", "status", "etapa", "fluxo"]);
  const tipoRegimeH = pickFirstHeader(headersNormToOrig, [
    "tipo_regime",
    "tipo regime",
    "tipo de regime",
    "regime",
    "regime aduaneiro",
  ]);
  const dtaEmbarqueBlH = pickFirstHeader(headersNormToOrig, [
    "dta_embarque_bl",
    "dta embarque bl",
    "data embarque bl",
    "dt embarque bl",
    "embarque_bl",
    "embarque bl",
    "data embarque",
    "dt embarque",
  ]);
  const dtaChegadaBlH = pickFirstHeader(headersNormToOrig, [
    "dta_chegada_bl",
    "dta chegada bl",
    "data chegada bl",
    "dt chegada bl",
    "chegada_bl",
    "chegada bl",
    "data chegada",
    "dt chegada",
  ]);
  const registroDiH = pickFirstHeader(headersNormToOrig, [
    "registro_di",
    "registro di",
    "di",
    "numero_di",
    "numero di",
    "num_di",
    "num di",
  ]);

  const phaseDateHeaders = Object.fromEntries(
    PHASES.map((ph) => [ph, findPhaseDateHeader(headersNormToOrig, ph)])
  );

  const out = [];
  for (const row of rows) {
    const id = String((idH && row[idH]) || "").trim();
    const refCliente = String((refClienteH && row[refClienteH]) || "").trim();
    const clienteNome = String((clienteNomeH && row[clienteNomeH]) || "").trim();
    const cliente = clienteNome || refCliente;
    const fornecedor =
      String((fornecedorH && row[fornecedorH]) || "").trim() ||
      String((respH && row[respH]) || "").trim();
    const responsavel = String((respH && row[respH]) || "").trim();
    const modal = String((modalH && row[modalH]) || "").trim();
    const origem = String((origemH && row[origemH]) || "").trim();
    const aberturaIso = aberturaH ? isoFromAnyDateValue(row[aberturaH]) : "";
    const tipoRegime = String((tipoRegimeH && row[tipoRegimeH]) || "").trim();
    const dtaEmbarqueBl = dtaEmbarqueBlH ? isoFromAnyDateValue(row[dtaEmbarqueBlH]) : "";
    const dtaChegadaBl = dtaChegadaBlH ? isoFromAnyDateValue(row[dtaChegadaBlH]) : "";
    const registroDi = registroDiH ? String((row[registroDiH] || "").trim()) : "";

    // ignora linhas vazias
    if (!id && !cliente && !refCliente && !clienteNome && !fornecedor && !modal && !origem) continue;

    const datas = {};
    for (const ph of PHASES) {
      const h = phaseDateHeaders[ph];
      const iso = h ? isoFromAnyDateValue(row[h]) : "";
      if (iso) datas[ph] = iso;
    }

    let faseAtual = "";
    if (faseH) faseAtual = statusToPhase(row[faseH]) || matchPhaseLabel(row[faseH]);
    if (!faseAtual) {
      // se não vier a fase atual, tenta inferir pela última data preenchida
      const lastWithDate = [...PHASES].reverse().find((ph) => datas[ph]);
      faseAtual = lastWithDate || PHASES[0];
    }

    // Se não houver datas por etapa no Excel, usa "ABERTURA" como referência (aparece no primeiro passo)
    if (aberturaIso && !datas[PHASES[0]]) datas[PHASES[0]] = aberturaIso;

    out.push({
      id: id || "(sem id)",
      cliente,
      refCliente,
      clienteNome,
      fornecedor,
      responsavel,
      aberturaIso,
      modal,
      origem,
      tipoRegime,
      faseAtual,
      datas,
      dtaEmbarqueBl,
      dtaChegadaBl,
      registroDi,
    });
  }

  // mantém ordem estável (se houver id)
  out.sort((a, b) => String(a.id).localeCompare(String(b.id), "pt-BR"));
  return out;
}

function buildTeamFilterOptions(teamsObj) {
  if (!els.teamFilter) return;
  
  const teams = Object.keys(teamsObj || {}).sort((a, b) => a.localeCompare(b, "pt-BR"));
  const options = [
    { value: "all", label: "Todas as equipes" },
    ...teams.map((t) => ({ value: t, label: t })),
    { value: "Sem equipe", label: "Sem equipe" }
  ];

  els.teamFilter.innerHTML = options
    .map((o) => `<option value="${escapeHtml(o.value)}">${escapeHtml(o.label)}</option>`)
    .join("");
}

async function tryLoadTeamsFromRoot() {
  try {
    const url = `./${encodeURI(TEAMS_FILE_NAME)}`;
    const res = await fetch(url, { cache: "no-store" });
    if (!res.ok) throw new Error(`Falha ao buscar Equipes (${res.status}).`);
    const obj = await res.json();
    teamByMemberKey = buildTeamIndex(obj);
    buildTeamFilterOptions(obj);
    applyTeamsToProcesses(processes);
    render();
  } catch {
    // fallback: usa o JSON embutido (útil em file://)
    teamByMemberKey = buildTeamIndex(DEFAULT_TEAMS);
    buildTeamFilterOptions(DEFAULT_TEAMS);
    applyTeamsToProcesses(processes);
    render();
  }
}

function normalizeOperacionalName(equipeName) {
  // Converte "Operacional 1" para "Operacional1" (remove espaço)
  if (!equipeName || equipeName === "all") return "OPERACIONAL%";
  if (equipeName === "Sem equipe") return "";
  return String(equipeName).replace(/\s+/g, "");
}

function buildWebhookUrl({ from, to, operacional } = {}) {
  const u = new URL(WEBHOOK_URL);

  // Sempre passar as datas informadas (quando preenchidas)
  const r = sanitizeDateRange(from, to);
  // O fluxo do n8n espera estes parâmetros na query — sempre envia as chaves
  u.searchParams.set("data_inicio", r.from || "");
  u.searchParams.set("data_fim", r.to || "");

  // Adiciona parâmetro operacional sempre (normalizado)
  const operacionalNormalized = normalizeOperacionalName(operacional || "all");
  u.searchParams.set("operacional", operacionalNormalized);

  // evita cache agressivo intermediário
  u.searchParams.set("_t", String(Date.now()));
  return u.toString();
}

function unwrapToArray(payload) {
  function walk(x, depth) {
    if (Array.isArray(x)) return x;
    if (!x || typeof x !== "object") return [];
    if (depth > 4) return [];

    // n8n: às vezes retorna um item único { json: {...} }
    if (x.json && typeof x.json === "object") return [x];

    const keys = ["data", "output", "items", "rows", "processes", "result", "body"];
    for (const k of keys) {
      if (!(k in x)) continue;
      const v = x[k];
      const got = walk(v, depth + 1);
      if (got.length) return got;
    }

    // último fallback: se parecer um registro, considera como 1 linha
    const hasSomeUsefulKey =
      "PROCESSO" in x ||
      "processo" in x ||
      "STATUS" in x ||
      "status" in x ||
      "CLIENTE" in x ||
      "cliente" in x ||
      "ABERTURA" in x ||
      "abertura" in x ||
      "aberturaIso" in x ||
      "faseAtual" in x;
    return hasSomeUsefulKey ? [x] : [];
  }

  return walk(payload, 0);
}

async function tryLoadFromWebhook({ from, to, operacional, signal } = {}) {
  const url = buildWebhookUrl({ from, to, operacional });
  // útil para validar no DevTools (sem mostrar mensagens na tela)
  try {
    window.__lastWebhookUrl = url;
  } catch {}
  const res = await fetch(url, { cache: "no-store", signal });
  if (!res.ok) throw new Error(`Falha ao buscar webhook (${res.status}).`);
  const payload = await res.json();
  const rows = unwrapToArray(payload);
  // n8n comumente retorna [{ json: {...} }]
  const flat = (rows || [])
    .map((it) => (it && typeof it === "object" && it.json && typeof it.json === "object" ? it.json : it))
    .filter((it) => it && typeof it === "object");

  // debug sem UI
  try {
    window.__lastWebhookPayload = payload;
    window.__lastWebhookRows = flat;
  } catch {}

  if (!flat.length) return [];

  function normalizeProcessFromApi(obj) {
    const o = obj || {};
    const id =
      String(
        o.id ??
          o.processo ??
          o.PROCESSO ??
          o["PROCESSO "] ??
          o.po ??
          o.PO ??
          ""
      ).trim();

    const refCliente = String(o.refCliente ?? o.ref_cliente ?? o.REF_CLIENTE ?? "").trim();
    const clienteNome = String(o.clienteNome ?? o.cliente_nome ?? o.CLIENTE ?? o.cliente ?? "").trim();
    const fornecedor = String(o.fornecedor ?? o.FORNECEDOR ?? "").trim();
    const responsavel = String(o.responsavel ?? o.responsável ?? o.RESPONSÁVEL ?? o.RESPONSAVEL ?? "").trim();
    const modal = String(o.modal ?? o.MODAL ?? "").trim();
    const origem = String(o.origem ?? o.ORIGEM ?? "").trim();
    const tipoRegime = String(o.tipoRegime ?? o.tipo_regime ?? o.TIPO_REGIME ?? "").trim();

    const aberturaIso =
      o.aberturaIso
        ? String(o.aberturaIso).trim()
        : isoFromAnyDateValue(o.abertura ?? o.ABERTURA ?? o["DATA ABERTURA"] ?? o["DT ABERTURA"]);

    const dtaEmbarqueBl =
      o.dtaEmbarqueBl
        ? String(o.dtaEmbarqueBl).trim()
        : isoFromAnyDateValue(o.dta_embarque_bl ?? o.DTA_EMBARQUE_BL ?? o["DTA EMBARQUE BL"] ?? o["DATA EMBARQUE BL"] ?? o.embarque ?? o.EMBARQUE ?? "");
    
    const dtaChegadaBl =
      o.dtaChegadaBl
        ? String(o.dtaChegadaBl).trim()
        : isoFromAnyDateValue(o.dta_chegada_bl ?? o.DTA_CHEGADA_BL ?? o["DTA CHEGADA BL"] ?? o["DATA CHEGADA BL"] ?? o.chegada ?? o.CHEGADA ?? "");
    
    const registroDi = String(
      o.registroDi ??
      o.registro_di ??
      o.REGISTRO_DI ??
      o["REGISTRO DI"] ??
      o.di ??
      o.DI ??
      o.numero_di ??
      o.NUMERO_DI ??
      ""
    ).trim();

    const faseOriginal =
      String(o.faseAtual ?? o.fase_atual ?? o["FASE ATUAL"] ?? o.fase ?? o.FASE ?? "").trim() ||
      String(o.status ?? o.STATUS ?? "").trim();
    let faseAtual = statusToPhase(faseOriginal) || matchPhaseLabel(faseOriginal) || "";

    const datas = (o.datas && typeof o.datas === "object" && !Array.isArray(o.datas)) ? o.datas : {};
    const cliente = String(o.cliente ?? clienteNome ?? refCliente ?? "").trim();

    return {
      id: id || "(sem id)",
      cliente,
      refCliente,
      clienteNome,
      fornecedor,
      responsavel,
      aberturaIso: String(aberturaIso || "").trim(),
      modal,
      origem,
      tipoRegime,
      faseAtual: faseAtual || PHASES[0],
      faseOriginal: faseOriginal || "",
      datas,
      dtaEmbarqueBl: String(dtaEmbarqueBl || "").trim(),
      dtaChegadaBl: String(dtaChegadaBl || "").trim(),
      registroDi: String(registroDi || "").trim(),
    };
  }

  const looksLikeProcess =
    flat.some((o) => Object.prototype.hasOwnProperty.call(o, "faseAtual")) ||
    flat.some((o) => Object.prototype.hasOwnProperty.call(o, "aberturaIso")) ||
    flat.some((o) => Object.prototype.hasOwnProperty.call(o, "datas"));

  // Reaproveita o mesmo mapeamento do Excel:
  // - Se vierem chaves tipo PROCESSO/CLIENTE/STATUS/TIPO_REGIME, funciona via aliases
  // - Se vierem chaves já "internas", também mapeia (por similaridade/aliases)
  const mapped = looksLikeProcess ? flat.map(normalizeProcessFromApi) : mapRowsToProcesses(flat);
  applyTeamsToProcesses(mapped);
  return mapped;
}

async function refreshFromApi({ showLoading } = { showLoading: true }) {
  if (isRefreshing) return;
  const { from, to } = syncDateInputs();
  // Datas obrigatórias
  if (!from || !to) {
    if (!from && els.openFrom) flashInvalidInput(els.openFrom);
    if (!to && els.openTo) flashInvalidInput(els.openTo);
    return;
  }

  // Obtém o valor do filtro de equipe operacional
  const operacional = els.teamFilter?.value || "all";

  const btn = els.refresh;
  const prevText = btn?.textContent;

  apiReqSeq += 1;
  const seq = apiReqSeq;
  try {
    if (apiAbort) apiAbort.abort();
  } catch {}
  apiAbort = new AbortController();

  isRefreshing = true;
  if (showLoading && btn) {
    btn.disabled = true;
    btn.textContent = "Atualizando...";
    btn.setAttribute("aria-busy", "true");
    btn.dataset.loading = "true";
    btn.title = buildWebhookUrl({ from, to, operacional });
  }

  try {
    const loaded = await tryLoadFromWebhook({ from, to, operacional, signal: apiAbort.signal });
    // Só aplica se for a última requisição (evita race)
    if (seq === apiReqSeq) {
      processes = Array.isArray(loaded) ? loaded : [];
      applyTeamsToProcesses(processes);
      render();
      showAlert("Dados atualizados com sucesso!", "success");
    }
  } catch (err) {
    // Se foi abortado por uma nova chamada, ignora
    if (String(err?.name) === "AbortError") return;
    showAlert("Erro ao atualizar dados. Tente novamente.", "danger");
    // Sem mensagens na tela; mantém dados atuais em caso de falha
  } finally {
    isRefreshing = false;
    if (showLoading && btn) {
      btn.disabled = false;
      btn.textContent = prevText || "Atualizar";
      btn.removeAttribute("aria-busy");
      delete btn.dataset.loading;
    }
  }
}

function groupByEquipe(items) {
  const groups = new Map();
  for (const p of items) {
    const key = p.equipe || "Sem equipe";
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(p);
  }
  // ordenação: equipe asc; processos por id asc
  const sorted = [...groups.entries()].sort(([a], [b]) => a.localeCompare(b, "pt-BR"));
  for (const [, arr] of sorted) {
    arr.sort((x, y) => String(x.id).localeCompare(String(y.id), "pt-BR"));
  }
  return sorted;
}

function getBaseFilteredProcesses() {
  const q = norm(els.search.value.trim());
  const { from, to } = sanitizeDateRange(els.openFrom?.value || "", els.openTo?.value || "");

  return processes.filter((p) => {
    const matchesText =
      !q ||
      norm(p.refCliente).includes(q) ||
      norm(p.cliente).includes(q);

    const abertura = String(p.aberturaIso || "");
    const matchesOpenFrom = !from ? true : (abertura && abertura >= from);
    const matchesOpenTo = !to ? true : (abertura && abertura <= to);

    return matchesText && matchesOpenFrom && matchesOpenTo;
  });
}

function getFilteredProcesses() {
  const phase = els.filter.value;
  const team = els.teamFilter?.value || "all";
  const base = getBaseFilteredProcesses();
  return base.filter((p) => {
    const matchesPhase = phase === "all" ? true : p.faseAtual === phase;
    const matchesTeam = team === "all" ? true : (p.equipe || "Sem equipe") === team;
    return matchesPhase && matchesTeam;
  });
}

function updateSummary() {
  const base = getBaseFilteredProcesses();
  const total = base.length;
  const fmt = (n) => Number(n || 0).toLocaleString("pt-BR");
  els.totalCount.textContent = fmt(total);

  const counts = Object.fromEntries(PHASES.map((p) => [p, 0]));
  for (const p of base) {
    if (counts[p.faseAtual] !== undefined) counts[p.faseAtual] += 1;
  }

  for (const card of document.querySelectorAll?.(".kpi-card[data-phase]") || []) {
    const phase = card.getAttribute("data-phase") || "";
    const n = counts[phase] ?? 0;
    const countEl = card.querySelector?.("[data-role='phase-count']");
    if (countEl) countEl.textContent = fmt(n);
  }

  // Segmentação por TIPO_REGIME dentro de "A Registrar"
  let entreposto = 0;
  let consumo = 0;
  for (const p of base) {
    if (p.faseAtual !== "A Registrar") continue;
    const v = norm(p.tipoRegime);
    if (!v) continue;
    if (v.includes("entre")) entreposto += 1;
    else if (v.includes("consum")) consumo += 1;
  }

  const regCard = document.querySelector?.(".kpi-card[data-phase='A Registrar']");
  if (regCard) {
    const elEnt = regCard.querySelector?.("[data-role='regime-entreposto']");
    const elCon = regCard.querySelector?.("[data-role='regime-consumo']");
    if (elEnt) elEnt.textContent = fmt(entreposto);
    if (elCon) elCon.textContent = fmt(consumo);
  }

  // Segmentação por status de registro dentro de "Em Trânsito"
  // Os dois status que somam "Em Trânsito" são: "Em transito já registrado" e "Em transito não registrado"
  let transitoRegistrado = 0;
  let transitoNaoRegistrado = 0;
  for (const p of base) {
    // IMPORTANTE: p.faseAtual deve ser exatamente "Em Trânsito"
    if (p.faseAtual !== "Em Trânsito") continue;
    
    const v = norm(p.faseOriginal || "");
    
    // Verifica se contém "ja registrado"
    if (v.includes("ja registrado")) {
      transitoRegistrado += 1;
    }
    // Verifica se contém "nao registrado"
    else if (v.includes("nao registrado")) {
      transitoNaoRegistrado += 1;
    }
    // Se não encontrou nenhum padrão mas está em "Em Trânsito", conta como não registrado (fallback)
    else {
      transitoNaoRegistrado += 1;
    }
    
    // Log para depuração no console do navegador
    if (p.faseAtual === "Em Trânsito") {
      console.log("Processo em Trânsito:", {
        ref: p.refCliente,
        faseOriginal: p.faseOriginal,
        norm: v,
        segmento: v.includes("ja registrado") ? "Registrado" : "A Registrar"
      });
    }
  }

  // Atualiza os elementos no DOM
  const transCard = document.querySelector('.kpi-card[data-phase="Em Trânsito"]');
  if (transCard) {
    const elReg = transCard.querySelector("[data-role='transito-registrado']");
    const elNaoReg = transCard.querySelector("[data-role='transito-nao-registrado']");
    if (elReg) elReg.textContent = fmt(transitoRegistrado);
    if (elNaoReg) elNaoReg.textContent = fmt(transitoNaoRegistrado);
  }
}

function render() {
  updateSummary();
  const items = getFilteredProcesses();
  els.resultsMeta.textContent = `${items.length} resultado${items.length === 1 ? "" : "s"}`;

  if (!items.length) {
    els.list.innerHTML = `<div class="empty">Nenhum processo encontrado com os filtros atuais.</div>`;
    return;
  }

  const groups = groupByEquipe(items);
  els.list.innerHTML = groups.map(([equipe, arr]) => renderGroup(equipe, arr)).join("");
}

function renderGroup(equipe, arr) {
  const total = arr.length;

  const expanded = groupExpanded[equipe] ?? true;
  const collapsedAttr = expanded ? "false" : "true";
  const bodyStyle = expanded ? "" : " style=\"display:none\"";

  const header = `
    <button class="group__toggle" type="button" data-action="toggle-group" data-team="${escapeHtml(equipe)}" aria-expanded="${expanded ? "true" : "false"}">
      <span class="group__chev" aria-hidden="true"></span>
      <span class="group__title">${escapeHtml(equipe)}</span>
      <span class="group__meta">
        <span class="group__pill">${total} processo${total === 1 ? "" : "s"}</span>
      </span>
    </button>
  `;

  const body = `
    <div class="group__body"${bodyStyle}>
      ${arr.map(renderCard).join("")}
    </div>
  `;

  return `<section class="group" data-collapsed="${collapsedAttr}">${header}${body}</section>`;
}

function renderCard(p) {
  const idx = phaseIndex(p.faseAtual);
  const pct = clamp(((idx + 1) / PHASES.length) * 100, 0, 100);
  const aberturaBr = formatDateBr(p.aberturaIso);
  const clienteLabel = p.clienteNome ? `Cliente: ${p.clienteNome}` : "Cliente:";
  const refLabel = p.refCliente ? `Ref: ${p.refCliente}` : "";

  const steps = PHASES.map((ph, i) => {
    let cls = "step step--pending";
    if (i < idx) cls = "step step--done";
    if (i === idx) cls = "step step--current";
    const dateIso = (p.datas || {})[ph];
    const dateBr = formatDateBr(dateIso);
    const title = ph;
    return `
      <div class="${cls}" title="${escapeHtml(title)}">
        <div class="step__label">${escapeHtml(ph)}</div>
        <div class="step__date">&nbsp;</div>
      </div>
    `;
  }).join("");

  return `
    <article class="card">
      <div class="card__top">
        <div class="card__title">
          <strong>${escapeHtml(p.refCliente || "-")}</strong>
          <div class="card__sub">
            <span class="chip"><span class="chip__dot"></span>${escapeHtml(clienteLabel)} ${escapeHtml(p.clienteNome ? "" : p.cliente)}</span>
            <span class="chip"><span class="chip__dot"></span>Responsável: ${escapeHtml(p.responsavel || p.fornecedor)}</span>
            <span class="chip"><span class="chip__dot"></span>Modal: ${escapeHtml(p.modal)}</span>
            <span class="chip"><span class="chip__dot"></span>Origem: ${escapeHtml(p.origem)}</span>
            <span class="chip"><span class="chip__dot"></span>Abertura: ${escapeHtml(aberturaBr || "-")}</span>
          </div>
        </div>
      </div>

      <div class="flow">
        <div class="progress" aria-label="Progresso do processo">
          <div class="progress__bar" style="width:${pct.toFixed(0)}%"></div>
        </div>
        <div class="steps">
          ${steps}
        </div>
      </div>
    </article>
  `;
}

function attachEvents() {
  els.search.addEventListener("input", () => render());
  els.filter.addEventListener("change", () => render());
  els.teamFilter?.addEventListener("change", () => {
    render();
    // Atualiza o título do botão de atualizar para mostrar a URL completa
    syncDateInputs();
  });

  els.openFrom?.addEventListener("change", () => {
    syncDateInputs();
    render();
  });
  els.openTo?.addEventListener("change", () => {
    syncDateInputs();
    render();
  });
  els.openFrom?.addEventListener("blur", () => {
    syncDateInputs();
    render();
  });
  els.openTo?.addEventListener("blur", () => {
    syncDateInputs();
    render();
  });

  els.refresh?.addEventListener("click", async () => {
    await refreshFromApi({ showLoading: true });
  });

  els.reset.addEventListener("click", () => {
    els.search.value = "";
    els.filter.value = "all";
    if (els.teamFilter) els.teamFilter.value = "all";
    if (els.openFrom) els.openFrom.value = "";
    if (els.openTo) els.openTo.value = "";
    render();
    showAlert("Filtros limpos.", "info", 2000);
  });

  // accordion (delegação)
  els.list.addEventListener("click", (e) => {
    const btn = e.target.closest?.("button[data-action='toggle-group']");
    if (!btn) return;
    const equipe = btn.getAttribute("data-team") || "Sem equipe";
    groupExpanded[equipe] = !(groupExpanded[equipe] ?? true);
    render();
  });

  // Chatbot events
  const chatBot = document.getElementById("chatbot");
  const openChatBtn = document.getElementById("openChat");
  const closeChatBtn = document.getElementById("closeChat");
  const sendChatBtn = document.getElementById("sendChat");
  const chatInput = document.getElementById("chatInput");
  const chatMessages = document.getElementById("chatMessages");
  const chatInvite = document.getElementById("chatInvite");

  const inviteMessages = [
    "Dúvidas sobre o operacional? Pergunte aqui!",
    "Posso te ajudar com o registro DUIMP?",
    "Precisa de informações sobre algum processo?",
    "Olá! Vamos agilizar seu operacional hoje?",
    "Clique aqui para tirar suas dúvidas!"
  ];

  if (chatInvite) {
    let msgIdx = 0;
    setInterval(() => {
      msgIdx = (msgIdx + 1) % inviteMessages.length;
      chatInvite.style.opacity = "0";
      setTimeout(() => {
        chatInvite.textContent = inviteMessages[msgIdx];
        chatInvite.style.opacity = "1";
      }, 500);
    }, 8000); // Troca a mensagem a cada 8 segundos
  }

  if (openChatBtn && chatBot) {
    openChatBtn.addEventListener("click", () => {
      chatBot.classList.add("chatbot--open");
      openChatBtn.style.display = "none";
    });
  }

  if (closeChatBtn && chatBot && openChatBtn) {
    closeChatBtn.addEventListener("click", () => {
      chatBot.classList.remove("chatbot--open");
      openChatBtn.style.display = "grid";
    });
  }

  function addChatMessage(text, side) {
    const msg = document.createElement("div");
    msg.className = `chat-msg chat-msg--${side}`;
    msg.textContent = text;
    chatMessages.appendChild(msg);
    chatMessages.scrollTop = chatMessages.scrollHeight;

    // Registrar no log
    messageLogs.push({
      timestamp: new Date().toISOString(),
      sender: side,
      message: text
    });
    console.log("Log de Mensagens Atualizado:", messageLogs);
  }

  function startRecurrentAlert() {
    if (recurrentAlertInterval) return;

    const sendAlert = () => {
      const alertMsg = "Lembrete: Processo X chegará em 15 dias. Podemos trabalhar no registro DUIMP?";
      showAlert(alertMsg, "warning", 0); // 0 para não fechar sozinho

      // Também adiciona ao chat como bot se o chat estiver aberto ou para registro
      messageLogs.push({
        timestamp: new Date().toISOString(),
        sender: "system-alert",
        message: alertMsg
      });
    };

    // Envia o primeiro imediatamente
    sendAlert();

    // Define o intervalo de 1 minuto
    recurrentAlertInterval = setInterval(sendAlert, 60000);
  }

  function stopRecurrentAlert() {
    if (recurrentAlertInterval) {
      clearInterval(recurrentAlertInterval);
      recurrentAlertInterval = null;
      showAlert("Lembretes pausados. Eles voltarão em 30 minutos.", "info", 5000);

      // Agenda o reinício após 30 minutos (1.800.000 ms)
      if (recurrentAlertTimeout) clearTimeout(recurrentAlertTimeout);
      recurrentAlertTimeout = setTimeout(() => {
        showAlert("Reiniciando lembretes operacionais...", "info", 4000);
        startRecurrentAlert();
      }, 1800000); 
    }
  }

  // Função para demonstrar todos os modelos de alertas
  function showDemoAlerts() {
    showAlert("Este é um alerta de SUCESSO! (Verde)", "success", 5000);
    setTimeout(() => showAlert("Este é um alerta de AVISO! (Amarelo/Laranja)", "warning", 6000), 1000);
    setTimeout(() => showAlert("Este é um alerta de ERRO! (Vermelho)", "danger", 7000), 2000);
    setTimeout(() => showAlert("Este é um alerta de INFORMAÇÃO! (Azul)", "info", 8000), 3000);
  }

  async function sendToChatbotAPI(text) {
    try {
      const response = await fetch(CHATBOT_WEBHOOK_URL, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          message: text,
          timestamp: new Date().toISOString(),
          context: "operacional_dashboard"
        }),
      });

      if (!response.ok) throw new Error("Erro na resposta da API do Chatbot");

      const data = await response.json();
      
      // Assume que o n8n retorna um campo 'output', 'response' ou 'message'
      const botResponse = data.output || data.response || data.message || data.text || "Recebi sua mensagem, mas não consegui processar a resposta.";
      
      addChatMessage(botResponse, "bot");

      // Verifica se a resposta contém gatilhos para alertas ou recorrência
      const lowerResp = botResponse.toLowerCase();
      if (lowerResp.includes("duimp") || lowerResp.includes("registro")) {
        startRecurrentAlert();
      }
    } catch (error) {
      console.error("Erro ao chamar API do Chatbot:", error);
      addChatMessage("Desculpe, estou com dificuldades para me conectar ao servidor agora.", "bot");
      showAlert("Erro de conexão com o Assistente", "danger");
    }
  }

  function handleChatSend() {
    const text = chatInput.value.trim();
    if (!text) return;

    addChatMessage(text, "user");
    chatInput.value = "";

    // Chama a API real do n8n
    sendToChatbotAPI(text);
  }

  if (sendChatBtn) {
    sendChatBtn.addEventListener("click", handleChatSend);
  }

  if (chatInput) {
    chatInput.addEventListener("keypress", (e) => {
      if (e.key === "Enter") handleChatSend();
    });
  }
}

function init() {
  buildPhaseFilterOptions();
  attachEvents();
  syncDateInputs();

  tryLoadTeamsFromRoot().catch(() => {
    // se não carregar equipes, ainda dá para ver os dados; cai em "Sem equipe"
    render();
  });
  // Não faz requisição automaticamente; só via botão "Atualizar".
  render();
}

init();

