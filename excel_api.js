/**
 * excel_api.js — Eduh Dev
 * Módulo central: todas as funções da Microsoft Graph API para Excel
 * Importe em qualquer script: const api = require("./excel_api")
 *
 * npm install @azure/msal-node axios dotenv
 */

require("dotenv").config();
const msal  = require("@azure/msal-node");
const axios = require("axios");
const fs    = require("fs");

// ─── Auth ──────────────────────────────────────────────────────────────────────
const SCOPES = ["https://graph.microsoft.com/Files.ReadWrite.All"];
let _token  = null;
let _msalApp = null;

function _cachePlugin() {
  const path = process.env.TOKEN_CACHE_PATH || ".token_cache.json";
  return {
    beforeCacheAccess: async (ctx) => {
      if (fs.existsSync(path)) ctx.tokenCache.deserialize(fs.readFileSync(path, "utf8"));
    },
    afterCacheAccess: async (ctx) => {
      if (ctx.cacheHasChanged) fs.writeFileSync(path, ctx.tokenCache.serialize());
    },
  };
}

function _buildApp() {
  const cfg = {
    auth: {
      clientId:  process.env.CLIENT_ID,
      authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
      ...(process.env.CLIENT_SECRET ? { clientSecret: process.env.CLIENT_SECRET } : {}),
    },
    cache: { cachePlugin: _cachePlugin() },
  };
  return process.env.AUTH_MODE === "client_credentials" && process.env.CLIENT_SECRET
    ? new msal.ConfidentialClientApplication(cfg)
    : new msal.PublicClientApplication(cfg);
}

/**
 * Autentica e retorna o access token.
 * deviceCodeCallback(info) → chamado no Device Code Flow com link + código.
 */
async function getToken(deviceCodeCallback) {
  if (_token) return _token;
  if (!_msalApp) _msalApp = _buildApp();

  if (process.env.AUTH_MODE === "client_credentials") {
    const r = await _msalApp.acquireTokenByClientCredential({
      scopes: ["https://graph.microsoft.com/.default"],
    });
    return (_token = r.accessToken);
  }

  const accounts = await _msalApp.getTokenCache().getAllAccounts();
  if (accounts.length) {
    try {
      const r = await _msalApp.acquireTokenSilent({ scopes: SCOPES, account: accounts[0] });
      return (_token = r.accessToken);
    } catch (_) {}
  }

  const r = await _msalApp.acquireTokenByDeviceCode({
    scopes: SCOPES,
    deviceCodeCallback: deviceCodeCallback || ((i) => {
      console.log(`\nAcesse: ${i.verificationUri}\nCódigo: ${i.userCode}\n`);
    }),
  });
  return (_token = r.accessToken);
}

/** Invalida o token em memória (força novo login) */
function resetToken() { _token = null; }

// ─── Cliente HTTP ──────────────────────────────────────────────────────────────
async function _graph(method, endpoint, body = null, params = {}, extraHeaders = {}) {
  const token = await getToken();
  try {
    const res = await axios({
      method,
      url: `https://graph.microsoft.com/v1.0${endpoint}`,
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...extraHeaders },
      data: body,
      params,
    });
    return res.data;
  } catch (e) {
    throw new Error(e.response?.data?.error?.message || e.message);
  }
}

function _driveBase(fileId) {
  return process.env.SHAREPOINT_SITE_ID
    ? `/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${fileId}`
    : `/me/drive/items/${fileId}`;
}

function _enc(name) { return encodeURIComponent(name); }

// ══════════════════════════════════════════════════════════════════════════════
//  ARQUIVOS
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Lista arquivos .xlsx do OneDrive.
 * @param {string} [pasta] - Caminho relativo no OneDrive (vazio = raiz)
 * @returns {Array<{id, name, sizeKB, modified, webUrl}>}
 */
async function listarArquivos(pasta = "") {
  const ep = pasta
    ? `/me/drive/root:/${pasta}:/children`
    : `/me/drive/root/children`;
  const data = await _graph("GET", ep, null, {
    $select: "id,name,size,lastModifiedDateTime,webUrl",
  });
  return (data.value || [])
    .filter((f) => f.name?.endsWith(".xlsx") || f.name?.endsWith(".xls"))
    .map((f) => ({
      id:       f.id,
      name:     f.name,
      sizeKB:   Math.round((f.size || 0) / 1024),
      modified: f.lastModifiedDateTime?.slice(0, 10),
      webUrl:   f.webUrl,
    }));
}

/**
 * Retorna metadados de um arquivo.
 * @param {string} fileId
 */
async function infoArquivo(fileId) {
  const [f, ws] = await Promise.all([
    _graph("GET", `/me/drive/items/${fileId}`),
    _graph("GET", `${_driveBase(fileId)}/workbook/worksheets`),
  ]);
  return {
    id:       f.id,
    name:     f.name,
    sizeKB:   Math.round((f.size || 0) / 1024),
    created:  f.createdDateTime?.slice(0, 10),
    modified: f.lastModifiedDateTime?.slice(0, 10),
    webUrl:   f.webUrl,
    sheets:   (ws.value || []).map((s) => s.name),
  };
}

/**
 * Deleta um arquivo pelo ID.
 * @param {string} fileId
 */
async function deletarArquivo(fileId) {
  await _graph("DELETE", `/me/drive/items/${fileId}`);
}

// ══════════════════════════════════════════════════════════════════════════════
//  ABAS (Worksheets)
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Lista todas as abas de um arquivo.
 * @param {string} fileId
 * @returns {Array<{name, position, visibility}>}
 */
async function listarAbas(fileId) {
  const data = await _graph("GET", `${_driveBase(fileId)}/workbook/worksheets`);
  return (data.value || []).map((s) => ({
    name:       s.name,
    position:   s.position,
    visibility: s.visibility,
  }));
}

/**
 * Cria uma nova aba.
 * @param {string} fileId
 * @param {string} nome
 * @returns {{ name, position }}
 */
async function criarAba(fileId, nome) {
  const res = await _graph("POST", `${_driveBase(fileId)}/workbook/worksheets/add`, { name: nome });
  return { name: res.name, position: res.position };
}

/**
 * Renomeia uma aba.
 * @param {string} fileId
 * @param {string} nomeAtual
 * @param {string} novoNome
 */
async function renomearAba(fileId, nomeAtual, novoNome) {
  await _graph("PATCH", `${_driveBase(fileId)}/workbook/worksheets/${_enc(nomeAtual)}`, { name: novoNome });
}

/**
 * Deleta uma aba pelo nome.
 * @param {string} fileId
 * @param {string} nome
 */
async function deletarAba(fileId, nome) {
  await _graph("DELETE", `${_driveBase(fileId)}/workbook/worksheets/${_enc(nome)}`);
}

// ══════════════════════════════════════════════════════════════════════════════
//  LEITURA DE DADOS
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Lê um range de uma aba. Sem range = usedRange.
 * @param {string} fileId
 * @param {string} aba
 * @param {string} [range] - ex: "A1:D10"
 * @returns {{ address, rowCount, columnCount, values: Array<Array> }}
 */
async function lerRange(fileId, aba, range = "") {
  const ep = range
    ? `${_driveBase(fileId)}/workbook/worksheets/${_enc(aba)}/range(address='${range}')`
    : `${_driveBase(fileId)}/workbook/worksheets/${_enc(aba)}/usedRange`;
  const data = await _graph("GET", ep);
  return {
    address:     data.address,
    rowCount:    data.rowCount,
    columnCount: data.columnCount,
    values:      data.values || [],
  };
}

/**
 * Lê um range e retorna como array de objetos usando a 1ª linha como header.
 * @param {string} fileId
 * @param {string} aba
 * @param {string} [range]
 * @returns {Array<Object>}
 */
async function lerRangeComoObjetos(fileId, aba, range = "") {
  const { values } = await lerRange(fileId, aba, range);
  if (!values.length) return [];
  const [header, ...rows] = values;
  return rows.map((r) => {
    const obj = {};
    header.forEach((h, i) => { obj[h || `Col${i + 1}`] = r[i] ?? ""; });
    return obj;
  });
}

/**
 * Lê uma única célula.
 * @param {string} fileId
 * @param {string} aba
 * @param {string} celula - ex: "B3"
 * @returns {{ address, value, formula, text }}
 */
async function lerCelula(fileId, aba, celula) {
  const data = await _graph(
    "GET",
    `${_driveBase(fileId)}/workbook/worksheets/${_enc(aba)}/range(address='${celula}')`,
    null,
    { $select: "address,values,formulas,text" }
  );
  return {
    address: data.address,
    value:   data.values?.[0]?.[0] ?? null,
    formula: data.formulas?.[0]?.[0] ?? null,
    text:    data.text?.[0]?.[0] ?? null,
  };
}

// ══════════════════════════════════════════════════════════════════════════════
//  ESCRITA DE DADOS
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Escreve valores em um range.
 * @param {string} fileId
 * @param {string} aba
 * @param {string} range - ex: "A1" ou "A1:C3"
 * @param {Array<Array>} values - matriz de valores
 */
async function escreverRange(fileId, aba, range, values) {
  await _graph(
    "PATCH",
    `${_driveBase(fileId)}/workbook/worksheets/${_enc(aba)}/range(address='${range}')`,
    { values }
  );
}

/**
 * Escreve uma fórmula em uma célula.
 * @param {string} fileId
 * @param {string} aba
 * @param {string} celula - ex: "D5"
 * @param {string} formula - ex: "=SUM(A1:A10)"
 */
async function escreverFormula(fileId, aba, celula, formula) {
  await _graph(
    "PATCH",
    `${_driveBase(fileId)}/workbook/worksheets/${_enc(aba)}/range(address='${celula}')`,
    { formulas: [[formula]] }
  );
}

/**
 * Limpa o conteúdo de um range.
 * @param {string} fileId
 * @param {string} aba
 * @param {string} range
 */
async function limparRange(fileId, aba, range) {
  await _graph(
    "POST",
    `${_driveBase(fileId)}/workbook/worksheets/${_enc(aba)}/range(address='${range}')/clear`,
    { applyTo: "Contents" }
  );
}

// ══════════════════════════════════════════════════════════════════════════════
//  TABELAS
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Lista tabelas de um arquivo.
 * @param {string} fileId
 * @returns {Array<{id, name}>}
 */
async function listarTabelas(fileId) {
  const data = await _graph("GET", `${_driveBase(fileId)}/workbook/tables`);
  return (data.value || []).map((t) => ({ id: t.id, name: t.name }));
}

/**
 * Lê os dados de uma tabela como array de objetos.
 * @param {string} fileId
 * @param {string} tabelaId - nome ou ID da tabela
 * @returns {Array<Object>}
 */
async function lerTabela(fileId, tabelaId) {
  const [rows, cols] = await Promise.all([
    _graph("GET", `${_driveBase(fileId)}/workbook/tables/${tabelaId}/rows`),
    _graph("GET", `${_driveBase(fileId)}/workbook/tables/${tabelaId}/columns`),
  ]);
  const headers = (cols.value || []).map((c) => c.name);
  return (rows.value || []).map((r) => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = r.values[0][i] ?? ""; });
    return obj;
  });
}

/**
 * Adiciona uma linha em uma tabela.
 * @param {string} fileId
 * @param {string} tabelaId
 * @param {Array} values - array simples com os valores de cada coluna
 */
async function adicionarLinhaTabela(fileId, tabelaId, values) {
  await _graph(
    "POST",
    `${_driveBase(fileId)}/workbook/tables/${tabelaId}/rows/add`,
    { values: [values] }
  );
}

// ══════════════════════════════════════════════════════════════════════════════
//  SESSÃO (edições em lote com melhor performance)
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Cria uma sessão persistente de edição (evita múltiplos round-trips).
 * @param {string} fileId
 * @returns {string} sessionId - use no header "workbook-session-id"
 */
async function criarSessao(fileId) {
  const res = await _graph("POST", `${_driveBase(fileId)}/workbook/createSession`, { persistChanges: true });
  return res.id;
}

/**
 * Fecha uma sessão.
 * @param {string} fileId
 * @param {string} sessionId
 */
async function fecharSessao(fileId, sessionId) {
  await _graph("POST", `${_driveBase(fileId)}/workbook/closeSession`, null, {}, {
    "workbook-session-id": sessionId,
  });
}

// ══════════════════════════════════════════════════════════════════════════════
//  EXPORTS
// ══════════════════════════════════════════════════════════════════════════════
module.exports = {
  // auth
  getToken,
  resetToken,
  // arquivos
  listarArquivos,
  infoArquivo,
  deletarArquivo,
  // abas
  listarAbas,
  criarAba,
  renomearAba,
  deletarAba,
  // leitura
  lerRange,
  lerRangeComoObjetos,
  lerCelula,
  // escrita
  escreverRange,
  escreverFormula,
  limparRange,
  // tabelas
  listarTabelas,
  lerTabela,
  adicionarLinhaTabela,
  // sessão
  criarSessao,
  fecharSessao,
};
