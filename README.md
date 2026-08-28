# Verbo Shop — Acervo da biblioteca

App de controle de estoque de livros (React + Vite + Tailwind), extraído do
artifact para rodar como site de verdade — principalmente pra testar a
leitura de código de barras pela câmera, que fica restrita dentro do
sandbox de pré-visualização do Claude.

## O que muda em relação ao artifact

- **Persistência**: no artifact os dados ficam em `window.storage`. Aqui,
  `src/main.jsx` cria um polyfill que usa `localStorage` do navegador com a
  mesma API (`get`/`set`/`delete`/`list`), então o app funciona sem editar
  `App.jsx`. Isso significa que os dados ficam salvos **só nesse navegador**
  (não sincronizam entre dispositivos).
- **Câmera**: fora do sandbox do artifact, o navegador vai pedir permissão
  de câmera normalmente ao tocar em "Ler código de barras". Precisa de
  **HTTPS** (ou `localhost`) — todo serviço de deploy abaixo já entrega isso.
- **Busca de ISBN**: as chamadas para Google Books e Open Library devem
  funcionar normalmente aqui (sem as restrições de rede do sandbox). O
  fallback via IA (api.anthropic.com) só funciona dentro do ambiente Claude
  — fora dele, ele falha silenciosamente e cai direto no cadastro manual,
  o que é esperado.

## Rodar localmente

```bash
npm install
npm run dev
```

Abre em `http://localhost:5173`. Como é `localhost`, a câmera funciona
mesmo sem HTTPS.

## Domínio próprio

O site está configurado pra rodar em `verboshop.verbomontesclaros.com`
(arquivo `public/CNAME` + `base: "/"` no `vite.config.js`). Os passos de
DNS e GitHub Pages estão descritos na conversa com o Claude — resumindo:

1. No provedor de DNS do domínio, crie um registro CNAME: host `verboshop`
   apontando para `SEU-USUARIO.github.io.`
2. No GitHub → Settings → Pages → Custom domain, digite
   `verboshop.verbomontesclaros.com` e salve.
3. Publique normalmente (`npm run deploy`) — o arquivo `public/CNAME`
   garante que o domínio não se perde a cada publicação.

## Publicar (pra testar em qualquer celular)

### Opção A — GitHub Pages (recomendado se você já usa GitHub)

1. Crie um repositório novo no GitHub chamado **`verbo-shop`** (ou outro
   nome — se usar outro nome, edite `base` em `vite.config.js` e
   `homepage` em `package.json` pra bater com ele).
2. Suba o código:
   ```bash
   git init
   git add .
   git commit -m "primeira versão"
   git branch -M main
   git remote add origin https://github.com/SEU-USUARIO/verbo-shop.git
   git push -u origin main
   ```
3. Rode:
   ```bash
   npm install
   npm run deploy
   ```
   Isso builda o projeto e publica a pasta `dist` numa branch `gh-pages`
   automaticamente (via pacote `gh-pages`).
4. No GitHub, vá em **Settings → Pages** do repositório e confirme que a
   fonte está apontando pra branch `gh-pages` (geralmente já detecta
   sozinho). Em alguns minutos o site fica em:
   `https://SEU-USUARIO.github.io/verbo-shop/`
5. Pra atualizar depois de qualquer mudança, é só repetir o passo 3
   (`npm run deploy`).

### Opção B — Vercel (mais simples)

1. Crie um repositório novo no GitHub e suba esta pasta inteira.
2. Entre em [vercel.com](https://vercel.com), clique em "Add New Project",
   importe o repositório.
3. O Vercel detecta Vite automaticamente (build command `vite build`,
   output `dist`). Clique em Deploy.
4. Abra a URL gerada (`https://seu-projeto.vercel.app`) no celular.

### Opção C — Netlify (arrastar e soltar, sem GitHub)

1. Rode localmente:
   ```bash
   npm install
   npm run build
   ```
2. Isso gera a pasta `dist/`.
3. Entre em [app.netlify.com/drop](https://app.netlify.com/drop) e arraste
   a pasta `dist` pra lá. Ele te dá uma URL pública na hora.

### Opção D — CodeSandbox (testar rápido, sem instalar nada)

1. Entre em [codesandbox.io](https://codesandbox.io), crie um sandbox Vite +
   React novo.
2. Substitua os arquivos gerados pelos desta pasta (`src/App.jsx`,
   `src/main.jsx`, `src/index.css`, `tailwind.config.js`,
   `postcss.config.js`, `index.html`, `package.json`).
3. O CodeSandbox já te dá uma URL pública com HTTPS.

## Estrutura

```
verbo-shop-app/
├── index.html
├── package.json
├── postcss.config.js
├── tailwind.config.js
├── vite.config.js
└── src/
    ├── App.jsx      ← o app inteiro (mesmo conteúdo do artifact)
    ├── main.jsx     ← ponto de entrada + polyfill de window.storage
    └── index.css    ← diretivas do Tailwind
```
