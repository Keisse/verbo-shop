import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// IMPORTANTE para GitHub Pages: troque "verbo-shop" pelo nome exato
// do seu repositório no GitHub (ex.: se o repo se chama "acervo-verbo",
// use base: "/acervo-verbo/"). Se publicar em outro serviço (Netlify,
// Firebase), pode deixar base: "/" novamente.
export default defineConfig({
  plugins: [react()],
  base: "/verbo-shop/",
});
