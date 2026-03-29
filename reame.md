# 📊 Excel Manager — Eduh Dev

CLI em Node.js para gerenciar planilhas Excel via **Microsoft Graph API** (Microsoft 365 / OneDrive).

---

## 📁 Estrutura

```
excel-manager/
├── excel_api.js      # Módulo central — todas as funções da Graph API
├── excel_manager.js  # CLI interativo — consome o excel_api.js
├── .env.example      # Modelo de variáveis de ambiente
├── .env              # Suas credenciais (não commitar!)
├── .token_cache.json # Cache de autenticação gerado automaticamente
└── package.json
```

---

## ⚙️ Setup

### 1. Instalar dependências

```bash
npm install
```

### 2. Criar app no Azure AD

1. Acesse [portal.azure.com](https://portal.azure.com)
2. Vá em **Azure Active Directory → App registrations → New registration**
3. Dê um nome (ex: `excel-manager`) e clique em **Register**
4. Copie o **Application (client) ID** e o **Directory (tenant) ID**
5. Em **API permissions → Add a permission → Microsoft Graph → Delegated**:
   - Adicione `Files.ReadWrite.All`
   - Clique em **Grant admin consent**
6. *(Opcional — modo app-only)* Em **Certificates & secrets → New client secret**, gere um secret e copie o valor

### 3. Configurar o `.env`

```bash
cp .env.example .env
```

Abra o `.env` e preencha:

```env
TENANT_ID=seu_tenant_id
CLIENT_ID=seu_client_id
CLIENT_SECRET=          # deixe vazio para usar Device Code (login no browser)
AUTH_MODE=device_code   # ou client_credentials
```

### 4. Rodar

```bash
npm start
```

---

## 🔐 Modos de Autenticação

| `AUTH_MODE` | Descrição | Quando usar |
|---|---|---|
| `device_code` | Login via browser — acessa o OneDrive do usuário logado | Uso pessoal / dev |
| `client_credentials` | App-only — sem login de usuário, precisa de `CLIENT_SECRET` | Bots / servidores |

> No modo `device_code`, ao iniciar o script será exibido um link e um código para autenticar no browser. O token é cacheado em `.token_cache.json` para não precisar repetir.

---

## 🖥️ Comandos do CLI

Após rodar `npm start`, use os comandos abaixo no terminal:

### 📂 Arquivos

| Comando | Descrição |
|---|---|
| `listarArquivos` | Lista todos os `.xlsx` do OneDrive |
| `infoArquivo` | Exibe detalhes de um arquivo (nome, abas, URL...) |
| `deletarArquivo` | Deleta um arquivo pelo ID |

### 📑 Abas

| Comando | Descrição |
|---|---|
| `listarAbas` | Lista todas as abas de um arquivo |
| `criarAba` | Cria uma nova aba |
| `renomearAba` | Renomeia uma aba existente |
| `deletarAba` | Deleta uma aba |

### 📖 Leitura

| Comando | Descrição |
|---|---|
| `lerRange` | Lê um range de células (ou todo o `usedRange`) |
| `lerCelula` | Lê valor, fórmula e texto de uma célula específica |

### ✏️ Escrita

| Comando | Descrição |
|---|---|
| `escreverRange` | Escreve dados em um range (linha por linha) |
| `escreverFormula` | Insere uma fórmula em uma célula |
| `limparRange` | Limpa o conteúdo de um range |

### 📋 Tabelas

| Comando | Descrição |
|---|---|
| `listarTabelas` | Lista tabelas do arquivo |
| `lerTabela` | Exibe os dados de uma tabela |
| `adicionarLinhaTabela` | Adiciona uma linha no final de uma tabela |

### ⚡ Sessão

| Comando | Descrição |
|---|---|
| `criarSessao` | Cria sessão persistente para edições em lote (mais performance) |
| `fecharSessao` | Encerra a sessão |

Digite `menu` a qualquer momento para ver todos os comandos.

---

## 📦 Usando o `excel_api.js` em outros scripts

O `excel_api.js` é um módulo reutilizável. Importe em qualquer script do seu projeto:

```js
const api = require("./excel_api");

// Listar arquivos
const files = await api.listarArquivos();
console.log(files);

// Ler range como array de objetos
const rows = await api.lerRangeComoObjetos(fileId, "Sheet1", "A1:E20");
console.log(rows);

// Escrever dados
await api.escreverRange(fileId, "Sheet1", "A1", [
  ["Nome", "Idade", "Cidade"],
  ["Dudu", 22, "SP"],
]);

// Inserir fórmula
await api.escreverFormula(fileId, "Sheet1", "D2", "=SUM(B2:C2)");

// Adicionar linha em tabela
await api.adicionarLinhaTabela(fileId, "Tabela1", ["Dudu", 22, "SP"]);
```

### Funções exportadas

```
getToken(deviceCodeCallback?)   Autentica e retorna o access token
resetToken()                    Invalida o token em memória

listarArquivos(pasta?)          → [{id, name, sizeKB, modified, webUrl}]
infoArquivo(fileId)             → {id, name, sizeKB, created, modified, webUrl, sheets}
deletarArquivo(fileId)

listarAbas(fileId)              → [{name, position, visibility}]
criarAba(fileId, nome)          → {name, position}
renomearAba(fileId, atual, novo)
deletarAba(fileId, nome)

lerRange(fileId, aba, range?)   → {address, rowCount, columnCount, values}
lerRangeComoObjetos(fileId, aba, range?)  → [{...}]
lerCelula(fileId, aba, celula)  → {address, value, formula, text}

escreverRange(fileId, aba, range, values)
escreverFormula(fileId, aba, celula, formula)
limparRange(fileId, aba, range)

listarTabelas(fileId)           → [{id, name}]
lerTabela(fileId, tabelaId)     → [{...}]
adicionarLinhaTabela(fileId, tabelaId, values)

criarSessao(fileId)             → sessionId
fecharSessao(fileId, sessionId)
```

---

## 🛡️ .gitignore recomendado

```
.env
.token_cache.json
node_modules/
```

---

## 📄 Licença

MIT — Eduh Dev
