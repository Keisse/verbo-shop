import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";
import "./index.css";

// ---------------------------------------------------------------------------
// Fora do ambiente de artifacts da Claude, window.storage não existe.
// Este polyfill reproduz a mesma API (get/set/delete/list) usando
// localStorage, para que o app funcione sem nenhuma outra alteração.
// ---------------------------------------------------------------------------
if (typeof window !== "undefined" && !window.storage) {
  const PREFIX = "verbo-shop:";

  window.storage = {
    async get(key) {
      const raw = localStorage.getItem(PREFIX + key);
      if (raw === null) throw new Error(`Chave "${key}" não encontrada`);
      return { key, value: raw, shared: false };
    },
    async set(key, value) {
      localStorage.setItem(PREFIX + key, value);
      return { key, value, shared: false };
    },
    async delete(key) {
      localStorage.removeItem(PREFIX + key);
      return { key, deleted: true, shared: false };
    },
    async list(prefix = "") {
      const keys = Object.keys(localStorage)
        .filter((k) => k.startsWith(PREFIX + prefix))
        .map((k) => k.slice(PREFIX.length));
      return { keys, prefix, shared: false };
    },
  };
}

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
