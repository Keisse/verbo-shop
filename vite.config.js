import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Com domínio próprio (verboshop.verbomontesclaros.com), o site é servido
// na raiz — por isso base volta a ser "/". Se um dia parar de usar domínio
// próprio e voltar pro endereço padrão do GitHub Pages
// (usuario.github.io/nome-do-repo/), troque de volta pra "/nome-do-repo/".
export default defineConfig({
  plugins: [react()],
  base: "/",
});
