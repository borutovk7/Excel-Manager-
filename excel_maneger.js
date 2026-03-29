#!/usr/bin/env node
/**
 * excel_manager.js — Eduh Dev
 * CLI interativo que usa todas as funções do excel_api.js
 *
 * npm install @azure/msal-node axios dotenv
 * node excel_manager.js
 */

const api      = require("./excel_api");
const readline = require("readline");

// ─── Cores ─────────────────────────────────────────────────────────────────────
const CLR = { reset:"\x1b[0m", bold:"\x1b[1m", red:"\x1b[91m", green:"\x1b[92m", yellow:"\x1b[93m", blue:"\x1b[94m", cyan:"\x1b[96m" };
const P   = (t, ...s) => s.join("") + t + CLR.reset;
const ok  = (m) => console.log(P(`\n  ✔  ${m}`, CLR.green));
const err = (m) => console.log(P(`\n  ✘  ${m}`, CLR.red));

// ─── Input ─────────────────────────────────────────────────────────────────────
const rl  = readline.createInterface({ input: process.stdin, output: process.stdout });
const ask = (q) => new Promise((res) => rl.question(P(`  ${q} `, CLR.cyan), res));

// ─── Tabela bonita ─────────────────────────────────────────────────────────────
function printTable(rows, cols) {
  if (!rows.length) return console.log(P("  (vazio)", CLR.yellow));
  const keys = cols || Object.keys(rows[0]);
  const w = keys.map((k) => Math.max(k.length, ...rows.map((r) => String(r[k] ?? "").length)));
  const sep = (l, m, r) => l + w.map((n) => "─".repeat(n + 2)).join(m) + r;
  console.log("\n" + sep("┌","┬","┐"));
  console.log("│" + keys.map((k, i) => ` ${P(k.padEnd(w[i]), CLR.bold)} `).join("│") + "│");
  console.log(sep("├","┼","┤"));
  rows.forEach((r) =>
    console.log("│" + keys.map((k, i) => ` ${String(r[k] ?? "").padEnd(w[i])} `).join("│") + "│")
  );
  console.log(sep("└","┴","┘") + "\n");
}

// ─── Comandos (usam funções do excel_api.js) ───────────────────────────────────
const cmds = {

  // ── ARQUIVOS ────────────────────────────────────────────────────────────────
  async listarArquivos() {
    const pasta = await ask("Pasta no OneDrive (Enter = raiz):");
    const files = await api.listarArquivos(pasta.trim());
    if (!files.length) return console.log(P("  Nenhum Excel encontrado.", CLR.yellow));
    printTable(files, ["id", "name", "sizeKB", "modified"]);
  },

  async infoArquivo() {
    const id   = await ask("ID do arquivo:");
    const info = await api.infoArquivo(id.trim());
    console.log(`
${P("  Nome:", CLR.bold)}     ${info.name}
${P("  ID:", CLR.bold)}       ${info.id}
${P("  Tamanho:", CLR.bold)}  ${info.sizeKB} KB
${P("  Criado:", CLR.bold)}   ${info.created}
${P("  Editado:", CLR.bold)}  ${info.modified}
${P("  Abas:", CLR.bold)}     ${info.sheets.join(", ")}
${P("  URL:", CLR.bold)}      ${info.webUrl}
    `);
  },

  async deletarArquivo() {
    const id   = await ask("ID do arquivo:");
    const conf = await ask("Confirma deletar? (s/n):");
    if (conf.toLowerCase() !== "s") return console.log(P("  Cancelado.", CLR.yellow));
    await api.deletarArquivo(id.trim());
    ok("Arquivo deletado.");
  },

  // ── ABAS ────────────────────────────────────────────────────────────────────
  async listarAbas() {
    const id   = await ask("ID do arquivo:");
    const abas = await api.listarAbas(id.trim());
    printTable(abas, ["name", "position", "visibility"]);
  },

  async criarAba() {
    const id   = await ask("ID do arquivo:");
    const nome = await ask("Nome da nova aba:");
    const res  = await api.criarAba(id.trim(), nome.trim());
    ok(`Aba '${res.name}' criada na posição ${res.position}.`);
  },

  async renomearAba() {
    const id   = await ask("ID do arquivo:");
    const aba  = await ask("Nome atual:");
    const novo = await ask("Novo nome:");
    await api.renomearAba(id.trim(), aba.trim(), novo.trim());
    ok(`Aba renomeada para '${novo.trim()}'.`);
  },

  async deletarAba() {
    const id   = await ask("ID do arquivo:");
    const aba  = await ask("Nome da aba:");
    const conf = await ask("Confirma? (s/n):");
    if (conf.toLowerCase() !== "s") return console.log(P("  Cancelado.", CLR.yellow));
    await api.deletarAba(id.trim(), aba.trim());
    ok(`Aba '${aba.trim()}' deletada.`);
  },

  // ── LEITURA ─────────────────────────────────────────────────────────────────
  async lerRange() {
    const id    = await ask("ID do arquivo:");
    const aba   = await ask("Nome da aba:");
    const range = await ask("Range (ex: A1:D10, Enter = usedRange):");
    const res   = await api.lerRange(id.trim(), aba.trim(), range.trim());
    console.log(P(`\n  ${aba.trim()} → ${res.address}  (${res.rowCount}R × ${res.columnCount}C)`, CLR.bold));
    const [header, ...rows] = res.values;
    if (!header) return console.log(P("  Vazio.", CLR.yellow));
    printTable(
      rows.map((r) => {
        const o = {};
        header.forEach((h, i) => { o[h || `Col${i+1}`] = r[i] ?? ""; });
        return o;
      })
    );
  },

  async lerCelula() {
    const id     = await ask("ID do arquivo:");
    const aba    = await ask("Nome da aba:");
    const celula = await ask("Célula (ex: B3):");
    const res    = await api.lerCelula(id.trim(), aba.trim(), celula.trim().toUpperCase());
    console.log(`
${P("  Endereço:", CLR.bold)} ${res.address}
${P("  Valor:", CLR.bold)}    ${res.value ?? "(vazio)"}
${P("  Fórmula:", CLR.bold)}  ${res.formula || "-"}
${P("  Texto:", CLR.bold)}    ${res.text ?? "-"}
    `);
  },

  // ── ESCRITA ─────────────────────────────────────────────────────────────────
  async escreverRange() {
    const id    = await ask("ID do arquivo:");
    const aba   = await ask("Nome da aba:");
    const range = await ask("Range inicial (ex: A1):");
    console.log(P("  Digite cada linha — colunas separadas por vírgula. Linha vazia = fim.", CLR.yellow));
    const values = [];
    while (true) {
      const linha = await ask(`  Linha ${values.length + 1}:`);
      if (!linha.trim()) break;
      values.push(linha.split(",").map((v) => isNaN(v.trim()) ? v.trim() : Number(v.trim())));
    }
    if (!values.length) return console.log(P("  Nenhum dado.", CLR.yellow));
    await api.escreverRange(id.trim(), aba.trim(), range.trim(), values);
    ok(`${values.length} linha(s) escritas em '${range.trim()}'.`);
  },

  async escreverFormula() {
    const id      = await ask("ID do arquivo:");
    const aba     = await ask("Nome da aba:");
    const celula  = await ask("Célula (ex: D5):");
    const formula = await ask("Fórmula (ex: =SUM(A1:A10)):");
    await api.escreverFormula(id.trim(), aba.trim(), celula.trim(), formula.trim());
    ok(`Fórmula escrita em ${celula.trim()}.`);
  },

  async limparRange() {
    const id    = await ask("ID do arquivo:");
    const aba   = await ask("Nome da aba:");
    const range = await ask("Range (ex: B2:D10):");
    await api.limparRange(id.trim(), aba.trim(), range.trim());
    ok(`Range '${range.trim()}' limpo.`);
  },

  // ── TABELAS ─────────────────────────────────────────────────────────────────
  async listarTabelas() {
    const id     = await ask("ID do arquivo:");
    const tabelas = await api.listarTabelas(id.trim());
    if (!tabelas.length) return console.log(P("  Nenhuma tabela.", CLR.yellow));
    printTable(tabelas, ["id", "name"]);
  },

  async lerTabela() {
    const id    = await ask("ID do arquivo:");
    const tabId = await ask("Nome ou ID da tabela:");
    const rows  = await api.lerTabela(id.trim(), tabId.trim());
    printTable(rows);
  },

  async adicionarLinhaTabela() {
    const id    = await ask("ID do arquivo:");
    const tabId = await ask("Nome ou ID da tabela:");
    const vals  = await ask("Valores separados por vírgula:");
    const values = vals.split(",").map((v) => isNaN(v.trim()) ? v.trim() : Number(v.trim()));
    await api.adicionarLinhaTabela(id.trim(), tabId.trim(), values);
    ok("Linha adicionada.");
  },

  // ── SESSÃO ──────────────────────────────────────────────────────────────────
  async criarSessao() {
    const id  = await ask("ID do arquivo:");
    const sid = await api.criarSessao(id.trim());
    ok("Sessão criada!");
    console.log(P(`  workbook-session-id: ${sid}\n`, CLR.yellow));
  },

  async fecharSessao() {
    const id  = await ask("ID do arquivo:");
    const sid = await ask("Session ID:");
    await api.fecharSessao(id.trim(), sid.trim());
    ok("Sessão fechada.");
  },
};

// ─── Menu ──────────────────────────────────────────────────────────────────────
function menu() {
  console.log(`
${P("══════════════════════════════════════════════", CLR.blue)}
${P("    EXCEL MANAGER — Microsoft Graph API", CLR.bold + CLR.blue)}
${P("══════════════════════════════════════════════", CLR.blue)}

${P(" ARQUIVOS", CLR.yellow)}
  listarArquivos       Listar .xlsx no OneDrive
  infoArquivo          Detalhes de um arquivo
  deletarArquivo       Deletar arquivo por ID

${P(" ABAS", CLR.yellow)}
  listarAbas           Listar abas
  criarAba             Criar nova aba
  renomearAba          Renomear aba
  deletarAba           Deletar aba

${P(" LEITURA", CLR.yellow)}
  lerRange             Ler range ou usedRange
  lerCelula            Ler célula específica

${P(" ESCRITA", CLR.yellow)}
  escreverRange        Escrever dados num range
  escreverFormula      Inserir fórmula numa célula
  limparRange          Limpar conteúdo de um range

${P(" TABELAS", CLR.yellow)}
  listarTabelas        Listar tabelas
  lerTabela            Ver dados de uma tabela
  adicionarLinhaTabela Adicionar linha numa tabela

${P(" SESSÃO", CLR.yellow)}
  criarSessao          Criar sessão para edições em lote
  fecharSessao         Fechar sessão

  ${P("sair", CLR.red)}                Encerrar
`);
}

// ─── Main ──────────────────────────────────────────────────────────────────────
(async () => {
  if (!process.env.TENANT_ID || !process.env.CLIENT_ID) {
    err("TENANT_ID e CLIENT_ID ausentes no .env. Copie .env.example → .env.");
    process.exit(1);
  }

  console.log(P("\n  Conectando à Microsoft Graph API...", CLR.cyan));
  try {
    await api.getToken((info) => {
      console.log(P(`\n  Acesse: `, CLR.bold) + P(info.verificationUri, CLR.green));
      console.log(P(`  Código: `, CLR.bold) + P(info.userCode, CLR.yellow + CLR.bold) + "\n");
    });
    ok("Autenticado!");
  } catch (e) {
    err(`Falha na autenticação: ${e.message}`);
    rl.close();
    process.exit(1);
  }

  menu();

  while (true) {
    const input = (await ask("\n➜  Comando:")).trim();
    if (!input) continue;
    if (input === "sair") { console.log(P("\n  Até logo! 👋\n", CLR.green)); rl.close(); break; }
    if (["ajuda","menu","help"].includes(input)) { menu(); continue; }
    if (cmds[input]) {
      try { await cmds[input](); }
      catch (e) { err(e.message); }
    } else {
      err(`Comando desconhecido: '${input}'. Digite 'menu'.`);
    }
  }
})();
