const path = require("path");
const express = require("express");

const app = express();
const PORT = process.env.PORT || 4173;
const ROOT = __dirname;

app.use(express.static(ROOT));

app.use((_req, res) => {
  res.sendFile(path.join(ROOT, "index.html"));
});

app.listen(PORT, () => {
  console.log(`Fort Lar CRM ativo em http://localhost:${PORT}`);
  console.log("Configure o Supabase em crm/config.js antes de usar o login.");
});
