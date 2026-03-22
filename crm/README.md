# Fort Lar CRM

CRM da Fort Lar preparado para GitHub + Supabase.

## O que esta pronto

- login com Supabase Auth
- banco PostgreSQL no Supabase
- storage de anexos por ordem de servico
- pipeline de leads
- agenda de visitas
- ordens de servico
- controle financeiro
- projeto local pronto para Git

## Arquivos principais

- `index.html`: interface principal
- `app.js`: logica do CRM conectada ao Supabase
- `styles.css`: layout responsivo
- `config.example.js`: modelo de configuracao
- `config.js`: configuracao local do Supabase
- `supabase-schema.sql`: schema, policies e bucket
- `server.js`: servidor estatico local

## Como configurar o Supabase

1. Crie um projeto no Supabase
2. Abra o SQL Editor
3. Rode o arquivo `supabase-schema.sql`
4. Se quiser iniciar com dados de exemplo, rode `supabase-seed.sql`
5. Em Authentication > Users, crie o primeiro usuario
6. Em Project Settings > API, copie:
   - Project URL
   - anon public key
7. Preencha `config.js` com esses dados

Exemplo:

```js
window.FORTLAR_SUPABASE_CONFIG = {
  supabaseUrl: "https://SEU-PROJETO.supabase.co",
  supabaseAnonKey: "SUA_ANON_KEY",
  storageBucket: "crm-anexos",
};
```

## Como rodar localmente

```bash
npm install
npm start
```

Abra:

```text
http://localhost:4173
```

## Como subir para o GitHub

No terminal da pasta raiz do projeto:

```bash
git add .
git commit -m "Preparar CRM Fort Lar para Supabase"
git remote add origin https://github.com/SEU-USUARIO/SEU-REPO.git
git push -u origin main
```

## Publicacao

Como o CRM agora fala direto com o Supabase, ele pode ser publicado como site estatico em:

- Vercel
- Netlify
- GitHub Pages

Observacao:

- para GitHub Pages, o `config.js` fica publico no cliente, entao use apenas a `anon key`
- nunca coloque service role key no front-end

## Vercel

No painel da Vercel:

1. Clique em `Add New Project`
2. Importe o repositório `Fortlar`
3. Em `Root Directory`, selecione `crm`
4. Mantenha o deploy como projeto estatico
5. Clique em `Deploy`

O arquivo `vercel.json` ja esta pronto em `crm/vercel.json`.
