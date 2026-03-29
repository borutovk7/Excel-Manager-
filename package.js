{
  "name": "excel-manager",
  "version": "1.0.0",
  "description": "CLI para gerenciar planilhas Excel via Microsoft Graph API",
  "main": "excel_api.js",
  "bin": {
    "excel-manager": "./excel_manager.js"
  },
  "scripts": {
    "start": "node excel_manager.js",
    "dev": "node --watch excel_manager.js"
  },
  "keywords": ["excel", "microsoft-graph", "onedrive", "xlsx", "cli"],
  "author": "Eduh Dev",
  "license": "MIT",
  "dependencies": {
    "@azure/msal-node": "^2.16.2",
    "axios": "^1.7.9",
    "dotenv": "^16.4.7"
  },
  "engines": {
    "node": ">=18.0.0"
  }
}
