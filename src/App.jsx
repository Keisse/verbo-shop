import { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  Plus, Search, X, Tag, PackagePlus, Loader2, Phone,
  Settings2, Trash2, ScanBarcode, CircleCheck, CircleAlert, ArrowLeft,
  Users, UserCircle2, UserPlus, ShoppingBag, Download, Tags, Menu, Receipt, Lightbulb, Pencil, Home, FileSpreadsheet, Eye, EyeOff,
} from "lucide-react";

// ---------------------------------------------------------------------------
// Verbo Shop — controle de estoque de livros em formato de fichário.
// Cada ficha representa um TÍTULO com quantidade em estoque. Vender abate do
// estoque e pede nome + telefone do comprador. Gestão de estoque concentra
// cadastro (com leitura de código de barras), edição de quantidade e exclusão.
// ---------------------------------------------------------------------------

const STATUS = {
  disponivel: { label: "Disponível", color: "#3F6B4F", bg: "#E8EFE7" },
  esgotado: { label: "Esgotado", color: "#B04A4A", bg: "#F3E1E1" },
};

const COR_ESTOQUE_BAIXO = "#B8752E";
const LIMITE_ESTOQUE_BAIXO = 2;

// Bolinha + texto, sem contorno — usado nas fichas e no detalhe rápido.
function statusVisual(livro) {
  if (livro.quantidade <= 0) {
    return { dot: STATUS.esgotado.color, texto: STATUS.esgotado.label, cor: STATUS.esgotado.color };
  }
  return { dot: STATUS.disponivel.color, texto: STATUS.disponivel.label, cor: STATUS.disponivel.color };
}

function estoqueBaixo(livro) {
  return livro.quantidade > 0 && livro.quantidade <= LIMITE_ESTOQUE_BAIXO;
}


// ---------------------------------------------------------------------------
// Supabase — banco compartilhado entre todos os usuários do app. Usamos a
// API REST direto via fetch (sem SDK), pra funcionar tanto no artifact de
// preview quanto no site publicado sem depender de um pacote npm extra.
// A chave "anon" é pública por design — a segurança real vem das policies
// de RLS configuradas nas tabelas.
// ---------------------------------------------------------------------------
const SUPABASE_URL = "https://viixukbodizfgzkiagim.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZpaXh1a2JvZGl6Zmd6a2lhZ2ltIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODc4NzczNjAsImV4cCI6MjEwMzQ1MzM2MH0.qHx4Fkvayc4TcKFit4en-8eh2PYsXLj27itValKeEDQ";

// Token de acesso da sessão atual — atualizado no login/logout/refresh.
// As funções sb* abaixo sempre leem o valor mais recente na hora da chamada.
let currentAccessToken = null;

function authHeaders() {
  return {
    apikey: SUPABASE_ANON_KEY,
    Authorization: `Bearer ${currentAccessToken || SUPABASE_ANON_KEY}`,
    "Content-Type": "application/json",
  };
}

async function sbSelect(tabela, filtro = "") {
  const url = `${SUPABASE_URL}/rest/v1/${tabela}?select=*${filtro ? `&${filtro}` : ""}`;
  const res = await fetch(url, { headers: authHeaders() });
  if (!res.ok) throw new Error(`Falha ao buscar "${tabela}" (${res.status})`);
  return res.json();
}

async function sbInsert(tabela, linhaOuLinhas) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${tabela}`, {
    method: "POST",
    headers: { ...authHeaders(), Prefer: "return=representation" },
    body: JSON.stringify(linhaOuLinhas),
  });
  if (!res.ok) throw new Error(`Falha ao inserir em "${tabela}" (${res.status})`);
  return res.json();
}

async function sbUpsert(tabela, linhas) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${tabela}?on_conflict=id`, {
    method: "POST",
    headers: { ...authHeaders(), Prefer: "resolution=merge-duplicates,return=representation" },
    body: JSON.stringify(linhas),
  });
  if (!res.ok) throw new Error(`Falha ao sincronizar "${tabela}" (${res.status})`);
  return res.json();
}

async function sbUpdate(tabela, id, campos) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${tabela}?id=eq.${encodeURIComponent(id)}`, {
    method: "PATCH",
    headers: { ...authHeaders(), Prefer: "return=representation" },
    body: JSON.stringify(campos),
  });
  if (!res.ok) throw new Error(`Falha ao atualizar "${tabela}" (${res.status})`);
  return res.json();
}

async function sbDelete(tabela, id) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${tabela}?id=eq.${encodeURIComponent(id)}`, {
    method: "DELETE",
    headers: authHeaders(),
  });
  if (!res.ok) throw new Error(`Falha ao excluir de "${tabela}" (${res.status})`);
}

// ---- autenticação (Supabase Auth via REST, sem SDK) ----
async function authLogin(email, senha) {
  const res = await fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=password`, {
    method: "POST",
    headers: { apikey: SUPABASE_ANON_KEY, "Content-Type": "application/json" },
    body: JSON.stringify({ email, password: senha }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error_description || data.msg || "E-mail ou senha incorretos.");
  return data; // { access_token, refresh_token, expires_in, user }
}

async function authRefresh(refreshToken) {
  const res = await fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=refresh_token`, {
    method: "POST",
    headers: { apikey: SUPABASE_ANON_KEY, "Content-Type": "application/json" },
    body: JSON.stringify({ refresh_token: refreshToken }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error("Sessão expirada.");
  return data;
}

async function authAlterarMinhaSenha(novaSenha) {
  const res = await fetch(`${SUPABASE_URL}/auth/v1/user`, {
    method: "PUT",
    headers: authHeaders(),
    body: JSON.stringify({ password: novaSenha }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.msg || "Não foi possível trocar a senha.");
  return data;
}

// Chama a Edge Function admin-usuarios (roda no servidor, com privilégios de
// admin) — usada pra criar pessoa, redefinir senha de outros e excluir pessoa.
async function chamarAdminUsuarios(acao, dados) {
  const res = await fetch(`${SUPABASE_URL}/functions/v1/admin-usuarios`, {
    method: "POST",
    headers: authHeaders(),
    body: JSON.stringify({ acao, ...dados }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Falha ao executar a operação.");
  return data;
}

// Conversões entre o formato usado no app (camelCase) e as colunas do banco
// (snake_case) — só necessário pra movimentações, o resto já bate 1:1.
function movParaLinha(mov) {
  return {
    id: mov.id,
    livro_id: mov.livroId,
    tipo: mov.tipo,
    data: mov.data,
    detalhe: mov.detalhe || null,
    comprador_nome: mov.compradorNome || null,
    comprador_telefone: mov.compradorTelefone || null,
    vendedor_nome: mov.vendedorNome || null,
    quantidade_vendida: mov.quantidadeVendida || null,
    desconto: mov.desconto || null,
    valor_total: mov.valorTotal != null ? mov.valorTotal : null,
  };
}

function linhaParaMov(row) {
  return {
    id: row.id,
    livroId: row.livro_id,
    tipo: row.tipo,
    data: row.data,
    detalhe: row.detalhe,
    compradorNome: row.comprador_nome,
    compradorTelefone: row.comprador_telefone,
    vendedorNome: row.vendedor_nome,
    quantidadeVendida: row.quantidade_vendida,
    desconto: row.desconto,
    valorTotal: row.valor_total,
  };
}

const PAPEIS = ["Liderado", "Líder"];

function maskTelefone(valor) {
  const digitos = valor.replace(/\D/g, "").slice(0, 11);
  if (digitos.length === 0) return "";
  if (digitos.length <= 2) return `(${digitos}`;
  if (digitos.length <= 6) return `(${digitos.slice(0, 2)}) ${digitos.slice(2)}`;
  if (digitos.length <= 10) return `(${digitos.slice(0, 2)}) ${digitos.slice(2, 6)}-${digitos.slice(6)}`;
  return `(${digitos.slice(0, 2)}) ${digitos.slice(2, 7)}-${digitos.slice(7)}`;
}

function uid(prefix) {
  return `${prefix}_${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
}

function formatDate(iso) {
  const d = new Date(iso);
  return d.toLocaleDateString("pt-BR", { day: "2-digit", month: "short", year: "numeric" });
}

function formatBRL(valor) {
  if (valor == null || isNaN(valor)) return null;
  return valor.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
}

// Vendas registradas antes deste recurso só têm o texto em "detalhe"
// (ex.: "Maria · (38) 99999-0000 · vendido por Fulano"). Extrai o que der.
function compradorDaMovimentacao(m) {
  if (m.compradorNome) return { nome: m.compradorNome, telefone: m.compradorTelefone || "" };
  if (!m.detalhe) return null;
  const partes = m.detalhe.split(" · ");
  if (partes.length < 2) return null;
  return { nome: partes[0].trim(), telefone: partes[1].trim() };
}

function statusDe(quantidade) {
  return quantidade > 0 ? "disponivel" : "esgotado";
}

// ---------------------------------------------------------------------------
// Leitura de ISBN — usada só para reconhecer um livro JÁ CADASTRADO (por
// comparação local do código) e somar exemplar automaticamente. Cadastro de
// título novo é sempre manual (título, autor, categoria, ISBN digitado).
// ---------------------------------------------------------------------------

function normalizeIsbn(rawCode) {
  const digits = rawCode.replace(/[^0-9]/g, "");
  if (digits.length === 13 && (digits.startsWith("978") || digits.startsWith("979"))) {
    return digits;
  }
  if (digits.length === 10) {
    return digits;
  }
  return null;
}

const CATEGORIAS_PADRAO = [
  "Cura",
  "Milagres",
  "Escatologia",
  "Motivacional",
  "Fé",
  "Oração",
  "Confissão",
  "História",
  "Apóstolos",
  "Autoridade do crente",
];

const SEED_BOOKS = [
  { titulo: "O Novo Nascimento", autor: "Kenneth E. Hagin", categoria: "Fé" },
  { titulo: "O Poder Criativo de Deus", autor: "Kenneth E. Hagin", categoria: "Confissão" },
  { titulo: "O Espírito dentro de nós e o Espírito sobre nós", autor: "Kenneth E. Hagin", categoria: "Autoridade do crente" },
  { titulo: "O Espírito Santo e você!", autor: "Rick Renner", categoria: "Autoridade do crente" },
  { titulo: "O Maravilhoso Nome de Jesus", autor: "E. W. Kenyon", categoria: "Autoridade do crente" },
  { titulo: "Lançando Suas Ansiedades Sobre O Senhor", autor: "Kenneth E. Hagin", categoria: "Oração" },
  { titulo: "Louvor, adoração e as coisas do coração", autor: "Manassés Guerra", categoria: "Oração" },
  { titulo: "Não Fale Negativo - Vol.2 (Coleção Legado)", autor: "Bud Wright", categoria: "Confissão" },
  { titulo: "Edifica seu fundamento", autor: "Rick Renner", categoria: "Fé" },
  { titulo: "Em busca de Timóteo", autor: "Tony Cooke", categoria: "Apóstolos" },
  { titulo: "Excelência", autor: "Andrew Wommack", categoria: "Motivacional" },
  { titulo: "Curai enfermos e expulsai demônios", autor: "T. L. Osborn", categoria: "Cura" },
  { titulo: "Curando os enfermos", autor: "T. L. Osborn", categoria: "Cura" },
  { titulo: "Curso de estudo da Cura Bíblica", autor: "Kenneth E. Hagin", categoria: "Cura" },
  { titulo: "Conhecendo o Pai", autor: "David Emigh", categoria: "Fé" },
  { titulo: "Conquistando almas", autor: "T. L. Osborn", categoria: "Apóstolos" },
  { titulo: "Crescendo espiritualmente", autor: "Kenneth E. Hagin", categoria: "Fé" },
  { titulo: "Amor caminho para vitória", autor: "Kenneth E. Hagin", categoria: "Motivacional" },
  { titulo: "Como Manter a Cabeça no Lugar Neste Mundo Louco", autor: "Rick Renner", categoria: "Motivacional" },
  { titulo: "Como ser dirigido pelo Espírito de Deus (Edição Legado)", autor: "Kenneth E. Hagin", categoria: "Autoridade do crente" },
  { titulo: "A Igreja Triunfante", autor: "Kenneth E. Hagin", categoria: "Autoridade do crente" },
  { titulo: "A palavra de Deus sobre Cura", autor: "Kenneth E. Hagin", categoria: "Cura" },
  { titulo: "A unção da Cura", autor: "Kenneth E. Hagin", categoria: "Cura" },
].map((l) => {
  const quantidade = 5;
  return { id: uid("bk"), titulo: l.titulo, autor: l.autor, categoria: l.categoria, quantidade, status: statusDe(quantidade) };
});

export default function Acervo() {
  const [livros, setLivros] = useState(null);
  const [movs, setMovs] = useState(null);
  const [busca, setBusca] = useState("");
  const [filtro, setFiltro] = useState("todos");
  const [categoriaFiltro, setCategoriaFiltro] = useState("todas");
  const [visualizacao, setVisualizacao] = useState("fichas"); // fichas | estante
  const [spineSelecionado, setSpineSelecionado] = useState(null); // livroId
  const [saving, setSaving] = useState(false);
  const [sbErro, setSbErro] = useState(null);
  const loadedOnce = useRef(false);

  // Venda
  const [vendaModal, setVendaModal] = useState(null);
  const [vQtd, setVQtd] = useState("1");
  const [vNome, setVNome] = useState("");
  const [vTelefone, setVTelefone] = useState("");
  const [vVendedorId, setVVendedorId] = useState("");
  const [vDesconto, setVDesconto] = useState("");
  const [vSugestoesAbertas, setVSugestoesAbertas] = useState(false);

  // Equipe / usuários
  const [usuarios, setUsuarios] = useState(null);
  const [equipeAba, setEquipeAba] = useState("equipe"); // equipe | cadastrar
  const [equipeBusca, setEquipeBusca] = useState("");
  const [compradoresBusca, setCompradoresBusca] = useState("");
  const [compradorSelecionado, setCompradorSelecionado] = useState(null);
  const [vendasBusca, setVendasBusca] = useState("");
  const [vendaSelecionada, setVendaSelecionada] = useState(null);
  const [vendasDataInicio, setVendasDataInicio] = useState("");
  const [vendasDataFim, setVendasDataFim] = useState("");
  const [vendasFiltroVendedor, setVendasFiltroVendedor] = useState("todos");

  // Categorias
  const [categoriasCadastradas, setCategoriasCadastradas] = useState(null);
  const [novaCategoria, setNovaCategoria] = useState("");
  const [confirmExcluirCategoria, setConfirmExcluirCategoria] = useState(null);
  const [confirmExcluirUsuario, setConfirmExcluirUsuario] = useState(null);
  const [novoUsuario, setNovoUsuario] = useState({ nome: "", email: "", telefone: "", papel: "Liderado", senha: "" });

  // Autenticação (Supabase Auth) — o "perfil" agora vem da pessoa logada de verdade
  const [sessao, setSessao] = useState(null); // { accessToken, refreshToken, userId }
  const [autenticando, setAutenticando] = useState(true); // true enquanto tenta restaurar sessão salva
  const [loginEmail, setLoginEmail] = useState("");
  const [loginSenha, setLoginSenha] = useState("");
  const [verLoginSenha, setVerLoginSenha] = useState(false);
  const [verSenhaPropria, setVerSenhaPropria] = useState(false);
  const [verSenhaNovoUsuario, setVerSenhaNovoUsuario] = useState(false);
  const [verSenhaRedefinir, setVerSenhaRedefinir] = useState(false);
  const [loginErro, setLoginErro] = useState("");
  const [loginCarregando, setLoginCarregando] = useState(false);
  const [perfil, setPerfil] = useState(null);
  const [novaSenhaPropria, setNovaSenhaPropria] = useState("");
  const [senhaAlterada, setSenhaAlterada] = useState(false);
  const [redefinirSenhaId, setRedefinirSenhaId] = useState(null); // usuarioId sendo redefinido pelo Líder
  const [redefinirSenhaValor, setRedefinirSenhaValor] = useState("");

  // Reposição rápida (a partir da ficha)
  const [repoModal, setRepoModal] = useState(null);
  const [rQtd, setRQtd] = useState("1");

  // Gestão de estoque
  const [pagina, setPagina] = useState("acervo"); // acervo | gestao | equipe | categorias | compradores | perfil
  const [menuAberto, setMenuAberto] = useState(false);
  const [gestaoAba, setGestaoAba] = useState("catalogo"); // catalogo | cadastrar
  const [gestaoBusca, setGestaoBusca] = useState("");
  const [gestaoFiltroStatus, setGestaoFiltroStatus] = useState("todos");
  const [gestaoFiltroCategoria, setGestaoFiltroCategoria] = useState("todas");
  const [confirmExcluir, setConfirmExcluir] = useState(null);
  const [novo, setNovo] = useState({ titulo: "", autor: "", categoria: "", quantidade: "1", isbn: "", preco: "" });
  const [livroEditandoId, setLivroEditandoId] = useState(null);
  const [edicaoForm, setEdicaoForm] = useState({ titulo: "", autor: "", categoria: "", isbn: "", preco: "" });
  const [importProcessando, setImportProcessando] = useState(false);
  const [arquivoSelecionado, setArquivoSelecionado] = useState(null);
  const [importResultado, setImportResultado] = useState(null); // { adicionados, atualizados, ignorados }
  const importInputRef = useRef(null);
  const [scannerAberto, setScannerAberto] = useState(false);
  const [scanStatus, setScanStatus] = useState("idle"); // idle | carregando | starting | scanning | error | lookup
  const [codigoManual, setCodigoManual] = useState("");
  const [scanMensagem, setScanMensagem] = useState(null); // { tipo: 'exemplar'|'nao_encontrado', texto }
  const [scanErroDetalhe, setScanErroDetalhe] = useState("");
  const scannerInstanceRef = useRef(null);
  const scannerRodandoRef = useRef(false);
  const SCANNER_DIV_ID = "verbo-shop-leitor-codigo";
  const [torchDisponivel, setTorchDisponivel] = useState(false);
  const [torchLigado, setTorchLigado] = useState(false);

  const SESSAO_KEY = "acervo:sessao"; // guarda só o refresh_token, localmente por aparelho

  const carregarCatalogo = async () => {
    try {
      let linhasLivros = await sbSelect("livros");
      if (linhasLivros.length === 0) {
        linhasLivros = await sbInsert("livros", SEED_BOOKS);
      }
      setLivros(linhasLivros);
    } catch (err) {
      console.error("Falha ao carregar livros do Supabase:", err);
      setSbErro("Não foi possível conectar ao banco de dados. Verifique sua internet e recarregue a página.");
      setLivros(SEED_BOOKS);
    }

    try {
      const linhasMovs = await sbSelect("movimentacoes");
      setMovs(linhasMovs.map(linhaParaMov).sort((a, b) => new Date(b.data) - new Date(a.data)));
    } catch (err) {
      console.error("Falha ao carregar movimentações do Supabase:", err);
      setMovs([]);
    }

    try {
      let linhasUsuarios = await sbSelect("usuarios");
      setUsuarios(linhasUsuarios);
    } catch (err) {
      console.error("Falha ao carregar equipe do Supabase:", err);
      setUsuarios([]);
    }

    try {
      let linhasCategorias = await sbSelect("categorias");
      if (linhasCategorias.length === 0) {
        linhasCategorias = await sbInsert(
          "categorias",
          CATEGORIAS_PADRAO.map((nome) => ({ id: nome, nome }))
        );
      }
      setCategoriasCadastradas(linhasCategorias.map((l) => l.nome));
    } catch (err) {
      console.error("Falha ao carregar categorias do Supabase:", err);
      setCategoriasCadastradas(CATEGORIAS_PADRAO);
    }

    loadedOnce.current = true;
  };

  // Depois de logar (ou restaurar sessão), busca a ficha da pessoa na equipe
  // pelo auth_id — é isso que vira o "perfil" (nome, papel, etc).
  const carregarPerfilDaSessao = async (userId) => {
    const linhas = await sbSelect("usuarios", `auth_id=eq.${userId}`);
    if (linhas.length === 0) {
      throw new Error("Sua conta não está vinculada a nenhuma pessoa da equipe. Fale com o Líder.");
    }
    setPerfil(linhas[0]);
  };

  const iniciarSessao = async (dadosAuth) => {
    currentAccessToken = dadosAuth.access_token;
    try {
      await carregarPerfilDaSessao(dadosAuth.user.id);
      await carregarCatalogo();
      setSessao({ accessToken: dadosAuth.access_token, refreshToken: dadosAuth.refresh_token, userId: dadosAuth.user.id });
      await window.storage.set(SESSAO_KEY, JSON.stringify({ refreshToken: dadosAuth.refresh_token }));
    } catch (err) {
      currentAccessToken = null;
      throw err;
    }
  };

  const encerrarSessao = async () => {
    currentAccessToken = null;
    setSessao(null);
    setPerfil(null);
    setLivros(null);
    setMovs(null);
    setUsuarios(null);
    setCategoriasCadastradas(null);
    loadedOnce.current = false;
    try {
      await window.storage.delete(SESSAO_KEY);
    } catch {
      // sem sessão salva, tudo bem
    }
  };

  const fazerLogin = async () => {
    if (!loginEmail.trim() || !loginSenha) return;
    setLoginCarregando(true);
    setLoginErro("");
    try {
      const dadosAuth = await authLogin(loginEmail.trim(), loginSenha);
      await iniciarSessao(dadosAuth);
      setLoginSenha("");
    } catch (err) {
      setLoginErro(err.message || "Não foi possível entrar. Confira e-mail e senha.");
    } finally {
      setLoginCarregando(false);
    }
  };

  // ---- restaurar sessão salva neste aparelho, ao abrir o app ----
  useEffect(() => {
    (async () => {
      try {
        const salvo = await window.storage.get(SESSAO_KEY);
        const { refreshToken } = JSON.parse(salvo.value);
        const dadosAuth = await authRefresh(refreshToken);
        await iniciarSessao(dadosAuth);
      } catch {
        // sem sessão salva (ou expirada/inválida) — mostra a tela de login normalmente
        currentAccessToken = null;
        try {
          await window.storage.delete(SESSAO_KEY);
        } catch {
          // não tinha sessão salva mesmo, tudo bem
        }
      } finally {
        setAutenticando(false);
      }
    })();
  }, []);

  // ---- limpeza da câmera ao desmontar ----
  useEffect(() => {
    return () => pararScanner();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Envolve uma chamada ao Supabase com indicador de "salvando" e captura de erro,
  // pra usar nas funções de cadastro/edição/exclusão sem repetir try/catch toda hora.
  const sincronizar = async (fn) => {
    setSaving(true);
    try {
      await fn();
      setSbErro(null);
    } catch (err) {
      console.error(err);
      setSbErro(err?.message || "Não foi possível salvar no banco compartilhado. Verifique sua internet e tente de novo.");
    } finally {
      setSaving(false);
    }
  };

  const registrarMovimento = (livroId, tipo, detalhe, extra = {}) => {
    const mov = { id: uid("mv"), livroId, tipo, data: new Date().toISOString(), detalhe: detalhe || "", ...extra };
    setMovs((prev) => [mov, ...prev]);
    sincronizar(() => sbInsert("movimentacoes", movParaLinha(mov)));
  };

  // ---- venda ----
  const abrirVendaModal = (livroId) => {
    setVQtd("1");
    setVNome("");
    setVTelefone("");
    setVDesconto("");
    setVSugestoesAbertas(false);
    const eu = perfil && (usuarios || []).find((u) => u.id === perfil.id && u.papel === "Liderado");
    setVVendedorId(eu ? eu.id : "");
    setVendaModal(livroId);
  };

  const confirmarVenda = () => {
    if (!vendaModal) return;
    const livro = livros.find((l) => l.id === vendaModal);
    const qtd = Math.max(1, Math.min(parseInt(vQtd, 10) || 1, livro.quantidade));
    const vendedor = (usuarios || []).find((u) => u.id === vVendedorId);
    if (!vNome.trim() || !vTelefone.trim() || !vendedor || livro.quantidade <= 0) return;

    const descontoNum = Math.max(0, parseFloat((vDesconto || "0").replace(",", ".")) || 0);
    const subtotal = livro.preco != null ? livro.preco * qtd : null;
    const valorTotal = subtotal != null ? Math.max(0, subtotal - descontoNum) : null;
    const novaQtd = livro.quantidade - qtd;

    setLivros((prev) => prev.map((l) => (l.id === vendaModal ? { ...l, quantidade: novaQtd, status: statusDe(novaQtd) } : l)));
    sincronizar(() => sbUpdate("livros", vendaModal, { quantidade: novaQtd, status: statusDe(novaQtd) }));
    registrarMovimento(
      vendaModal,
      `venda de ${qtd} ${qtd === 1 ? "exemplar" : "exemplares"}`,
      `${vNome.trim()} · ${vTelefone.trim()} · vendido por ${vendedor.nome}`,
      {
        compradorNome: vNome.trim(),
        compradorTelefone: vTelefone.trim(),
        vendedorNome: vendedor.nome,
        quantidadeVendida: qtd,
        desconto: descontoNum || null,
        valorTotal,
      }
    );
    setVendaModal(null);
  };

  // ---- reposição rápida (ficha) ----
  const abrirRepoModal = (livroId) => {
    setRQtd("1");
    setRepoModal(livroId);
  };

  const confirmarRepo = () => {
    if (!repoModal) return;
    const qtd = Math.max(1, parseInt(rQtd, 10) || 1);
    const livro = livros.find((l) => l.id === repoModal);
    const novaQtd = livro.quantidade + qtd;
    setLivros((prev) => prev.map((l) => (l.id === repoModal ? { ...l, quantidade: novaQtd, status: statusDe(novaQtd) } : l)));
    sincronizar(() => sbUpdate("livros", repoModal, { quantidade: novaQtd, status: statusDe(novaQtd) }));
    registrarMovimento(repoModal, `reposição de ${qtd} ${qtd === 1 ? "exemplar" : "exemplares"}`, "");
    setRepoModal(null);
  };

  // ---- gestão de estoque: editar quantidade / excluir ----
  const editarQuantidade = (livroId, novaQtd) => {
    const qtd = Math.max(0, parseInt(novaQtd, 10) || 0);
    setLivros((prev) => prev.map((l) => (l.id === livroId ? { ...l, quantidade: qtd, status: statusDe(qtd) } : l)));
    sincronizar(() => sbUpdate("livros", livroId, { quantidade: qtd, status: statusDe(qtd) }));
  };

  const excluirLivro = (livroId) => {
    setLivros((prev) => prev.filter((l) => l.id !== livroId));
    setConfirmExcluir(null);
    sincronizar(() => sbDelete("livros", livroId));
  };

  const abrirEdicaoLivro = (livro) => {
    setEdicaoForm({
      titulo: livro.titulo,
      autor: livro.autor,
      categoria: livro.categoria,
      isbn: livro.isbn || "",
      preco: livro.preco != null ? String(livro.preco) : "",
    });
    setLivroEditandoId(livro.id);
  };

  const salvarEdicaoLivro = () => {
    if (!livroEditandoId || !edicaoForm.titulo.trim()) return;
    const precoNum = parseFloat(edicaoForm.preco.replace(",", "."));
    const campos = {
      titulo: edicaoForm.titulo.trim(),
      autor: edicaoForm.autor.trim() || "Autor desconhecido",
      categoria: edicaoForm.categoria || "Sem categoria",
      isbn: edicaoForm.isbn.trim() || null,
      preco: !isNaN(precoNum) && precoNum >= 0 ? precoNum : null,
    };
    setLivros((prev) => prev.map((l) => (l.id === livroEditandoId ? { ...l, ...campos } : l)));
    sincronizar(() => sbUpdate("livros", livroEditandoId, campos));
    setLivroEditandoId(null);
  };

  // ---- gestão da equipe ----
  const souLider = !!(perfil && perfil.papel === "Líder");

  const adicionarUsuario = async () => {
    if (!souLider) return;
    if (!novoUsuario.nome.trim() || !novoUsuario.email.trim() || !novoUsuario.senha || novoUsuario.senha.length < 6) return;
    await sincronizar(async () => {
      const resultado = await chamarAdminUsuarios("criar_pessoa", {
        nome: novoUsuario.nome.trim(),
        email: novoUsuario.email.trim(),
        telefone: novoUsuario.telefone.trim(),
        papel: novoUsuario.papel,
        senha: novoUsuario.senha,
      });
      setUsuarios((prev) => [resultado.usuario, ...(prev || [])]);
      setNovoUsuario({ nome: "", email: "", telefone: "", papel: "Liderado", senha: "" });
    });
  };

  const editarPapelUsuario = (usuarioId, papel) => {
    if (!souLider) return;
    setUsuarios((prev) => prev.map((u) => (u.id === usuarioId ? { ...u, papel } : u)));
    sincronizar(() => sbUpdate("usuarios", usuarioId, { papel }));
  };

  const excluirUsuario = async (usuarioId) => {
    if (!souLider) return;
    await sincronizar(async () => {
      await chamarAdminUsuarios("excluir_pessoa", { usuarioId });
      setUsuarios((prev) => prev.filter((u) => u.id !== usuarioId));
      setConfirmExcluirUsuario(null);
    });
  };

  // ---- categorias ----
  const adicionarCategoria = () => {
    const nome = novaCategoria.trim();
    if (!nome) return;
    const jaExiste = (categoriasCadastradas || []).some((c) => c.toLowerCase() === nome.toLowerCase());
    if (jaExiste) return;
    setCategoriasCadastradas((prev) => [...(prev || []), nome]);
    sincronizar(() => sbInsert("categorias", { id: nome, nome }));
    setNovaCategoria("");
  };

  const excluirCategoria = (nome) => {
    setCategoriasCadastradas((prev) => (prev || []).filter((c) => c !== nome));
    setConfirmExcluirCategoria(null);
    sincronizar(() => sbDelete("categorias", nome));
  };

  // ---- perfil / senha ----
  const trocarMinhaSenha = async () => {
    if (!novaSenhaPropria || novaSenhaPropria.length < 6) return;
    await sincronizar(async () => {
      await authAlterarMinhaSenha(novaSenhaPropria);
      setNovaSenhaPropria("");
      setSenhaAlterada(true);
      setTimeout(() => setSenhaAlterada(false), 2500);
    });
  };

  const abrirRedefinirSenha = (usuarioId) => {
    setRedefinirSenhaValor("");
    setRedefinirSenhaId(usuarioId);
  };

  const confirmarRedefinirSenha = async () => {
    if (!redefinirSenhaId || !redefinirSenhaValor || redefinirSenhaValor.length < 6) return;
    await sincronizar(async () => {
      await chamarAdminUsuarios("redefinir_senha", { usuarioId: redefinirSenhaId, novaSenha: redefinirSenhaValor });
      setRedefinirSenhaId(null);
    });
  };

  // ---- cadastro de novo livro ----
  const resetNovo = () => setNovo({ titulo: "", autor: "", categoria: "", quantidade: "1", isbn: "", preco: "" });

  const adicionarLivro = () => {
    if (!novo.titulo.trim()) return;
    const quantidade = Math.max(0, parseInt(novo.quantidade, 10) || 0);
    const precoNum = parseFloat(novo.preco.replace(",", "."));
    const livro = {
      id: uid("bk"),
      titulo: novo.titulo.trim(),
      autor: novo.autor.trim() || "Autor desconhecido",
      categoria: novo.categoria || "Sem categoria",
      quantidade,
      status: statusDe(quantidade),
      isbn: novo.isbn.trim() || null,
      preco: !isNaN(precoNum) && precoNum >= 0 ? precoNum : null,
    };
    setLivros((prev) => [livro, ...prev]);
    sincronizar(() => sbInsert("livros", livro));
    resetNovo();
  };

  // ---- importação de livros por planilha ----
  const baixarModeloPlanilha = () => {
    const linhas = [
      { Título: "O Novo Nascimento", Autor: "Kenneth E. Hagin", Categoria: "Fé", ISBN: "9788578600000", Quantidade: 5, "Preço": 29.9 },
      { Título: "Curando os Enfermos", Autor: "T. L. Osborn", Categoria: "Cura", ISBN: "", Quantidade: 3, "Preço": 24.5 },
    ];
    const planilha = XLSX.utils.json_to_sheet(linhas);
    planilha["!cols"] = [{ wch: 30 }, { wch: 22 }, { wch: 16 }, { wch: 16 }, { wch: 12 }, { wch: 10 }];
    const livro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(livro, planilha, "Livros");
    XLSX.writeFile(livro, "modelo-importacao-livros.xlsx");
  };

  const normalizarChave = (str) =>
    str
      .toString()
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "");

  const pegarCampo = (linhaNormalizada, chaves) => {
    for (const chave of chaves) {
      if (linhaNormalizada[chave] !== undefined && linhaNormalizada[chave] !== "") {
        return linhaNormalizada[chave];
      }
    }
    return "";
  };

  const importarPlanilha = async (arquivo) => {
    setImportProcessando(true);
    setImportResultado(null);
    try {
      const buffer = await arquivo.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const primeiraAba = workbook.SheetNames[0];
      const linhas = XLSX.utils.sheet_to_json(workbook.Sheets[primeiraAba], { defval: "" });

      let adicionados = 0;
      let atualizados = 0;
      let ignorados = 0;
      const categoriasNovas = new Set();
      let livrosAtualizados = [...livros];

      linhas.forEach((linhaOriginal) => {
        const linha = {};
        Object.keys(linhaOriginal).forEach((k) => {
          linha[normalizarChave(k)] = linhaOriginal[k];
        });

        const titulo = String(pegarCampo(linha, ["titulo", "título", "nome"])).trim();
        if (!titulo) {
          ignorados++;
          return;
        }
        const autor = String(pegarCampo(linha, ["autor", "autores"])).trim();
        const categoria = String(pegarCampo(linha, ["categoria", "categorias"])).trim();
        const isbnBruto = String(pegarCampo(linha, ["isbn", "codigo", "código", "ean"])).trim();
        const isbn = isbnBruto ? normalizeIsbn(isbnBruto) || isbnBruto : "";
        const quantidadeBruta = pegarCampo(linha, ["quantidade", "qtd", "estoque"]);
        const quantidade = Math.max(0, parseInt(quantidadeBruta, 10) || 0);
        const precoBruto = String(pegarCampo(linha, ["preco", "preço", "valor"])).replace(",", ".");
        const precoNum = parseFloat(precoBruto);
        const preco = !isNaN(precoNum) && precoNum >= 0 ? precoNum : null;

        const existenteIdx = isbn
          ? livrosAtualizados.findIndex((l) => l.isbn && normalizeIsbn(l.isbn) === normalizeIsbn(isbn))
          : -1;

        if (existenteIdx >= 0) {
          const existente = livrosAtualizados[existenteIdx];
          const novaQtd = existente.quantidade + quantidade;
          livrosAtualizados[existenteIdx] = { ...existente, quantidade: novaQtd, status: statusDe(novaQtd) };
          atualizados++;
        } else {
          if (categoria) categoriasNovas.add(categoria);
          livrosAtualizados = [
            {
              id: uid("bk"),
              titulo,
              autor: autor || "Autor desconhecido",
              categoria: categoria || "Sem categoria",
              quantidade,
              status: statusDe(quantidade),
              isbn: isbn || null,
              preco,
            },
            ...livrosAtualizados,
          ];
          adicionados++;
        }
      });

      setLivros(livrosAtualizados);
      sincronizar(() => sbUpsert("livros", livrosAtualizados));

      if (categoriasNovas.size > 0) {
        let novasParaSalvar = [];
        setCategoriasCadastradas((prev) => {
          const atuais = new Set((prev || []).map((c) => c.toLowerCase()));
          const novas = Array.from(categoriasNovas).filter((c) => !atuais.has(c.toLowerCase()));
          novasParaSalvar = novas;
          return [...(prev || []), ...novas];
        });
        if (novasParaSalvar.length > 0) {
          sincronizar(() => sbInsert("categorias", novasParaSalvar.map((nome) => ({ id: nome, nome }))));
        }
      }

      setImportResultado({ adicionados, atualizados, ignorados, erro: null });
    } catch (err) {
      console.error("Falha ao importar planilha:", err);
      setImportResultado({ erro: "Não foi possível ler o arquivo. Confira se é um .xlsx, .xls ou .csv válido." });
    } finally {
      setImportProcessando(false);
      setArquivoSelecionado(null);
      if (importInputRef.current) importInputRef.current.value = "";
    }
  };

  // ---- leitor de código de barras ----
  // Carrega a lib html5-qrcode uma única vez (via cdnjs, mesma versão que você
  // já usava). Ela decodifica os frames por conta própria — muito mais
  // confiável entre navegadores do que a API nativa BarcodeDetector, que não
  // existe no Safari/iPhone e é inconsistente até em Chrome desktop.
  const carregarHtml5Qrcode = () => {
    if (window.Html5Qrcode) return Promise.resolve();
    if (window.__html5QrcodeLoading) return window.__html5QrcodeLoading;
    window.__html5QrcodeLoading = new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = "https://cdnjs.cloudflare.com/ajax/libs/html5-qrcode/2.3.8/html5-qrcode.min.js";
      script.async = true;
      script.onload = () => resolve();
      script.onerror = () => reject(new Error("Não foi possível carregar a biblioteca de leitura."));
      document.body.appendChild(script);
    });
    return window.__html5QrcodeLoading;
  };

  const abrirScanner = async () => {
    setScannerAberto(true);
    setScanStatus("carregando");
    setScanMensagem(null);
    setScanErroDetalhe("");
    setTorchDisponivel(false);
    setTorchLigado(false);

    try {
      await carregarHtml5Qrcode();
    } catch (err) {
      console.error(err);
      setScanErroDetalhe("Falha ao carregar a biblioteca de leitura (" + (err?.message || err) + ").");
      setScanStatus("error");
      return;
    }

    if (!window.Html5Qrcode) {
      setScanErroDetalhe("A biblioteca de leitura carregou, mas Html5Qrcode não ficou disponível.");
      setScanStatus("error");
      return;
    }

    setScanStatus("starting");
    // dá um instante pro <div id="verbo-shop-leitor-codigo"> montar no DOM
    await new Promise((r) => setTimeout(r, 50));

    try {
      const formatos = [
        window.Html5QrcodeSupportedFormats.EAN_13,
        window.Html5QrcodeSupportedFormats.EAN_8,
        window.Html5QrcodeSupportedFormats.UPC_A,
        window.Html5QrcodeSupportedFormats.UPC_E,
        window.Html5QrcodeSupportedFormats.CODE_128,
      ];
      const instancia = new window.Html5Qrcode(SCANNER_DIV_ID, {
        formatsToSupport: formatos,
        verbose: true, // ajuda a diagnosticar pelo console do navegador
      });
      scannerInstanceRef.current = instancia;

      await instancia.start(
        { facingMode: "environment" },
        { fps: 10, qrbox: { width: 300, height: 160 } },
        (decodedText) => {
          // achou um código — para o scanner e busca os dados
          console.log("[Scanner] código detectado pela câmera:", decodedText);
          pararScanner();
          setScannerAberto(false);
          buscarLivroPorCodigo(decodedText);
        },
        () => {} // ignora frames sem detecção
      );
      scannerRodandoRef.current = true;
      setScanStatus("scanning");

      // Lanterna, quando o aparelho suportar (ajuda muito em ambiente com pouca luz)
      try {
        const capacidades = instancia.getRunningTrackCapabilities?.();
        setTorchDisponivel(!!capacidades?.torch);
      } catch {
        setTorchDisponivel(false);
      }
    } catch (err) {
      console.error("Falha ao abrir câmera:", err);
      const nome = err?.name ? `${err.name}: ` : "";
      setScanErroDetalhe(nome + (err?.message || String(err)));
      setScanStatus("error");
    }
  };

  const alternarTorch = async () => {
    const instancia = scannerInstanceRef.current;
    if (!instancia) return;
    try {
      const novoEstado = !torchLigado;
      await instancia.applyVideoConstraints({ advanced: [{ torch: novoEstado }] });
      setTorchLigado(novoEstado);
    } catch (err) {
      console.warn("Lanterna não suportada neste aparelho:", err);
    }
  };

  const pararScanner = () => {
    const instancia = scannerInstanceRef.current;
    if (instancia && scannerRodandoRef.current) {
      scannerRodandoRef.current = false;
      instancia
        .stop()
        .then(() => instancia.clear())
        .catch(() => {});
    }
    scannerInstanceRef.current = null;
  };

  const fecharScanner = () => {
    pararScanner();
    setScannerAberto(false);
    setScanStatus("idle");
    setTorchDisponivel(false);
    setTorchLigado(false);
  };

  const buscarLivroPorCodigo = (codigo) => {
    console.log("[Scanner] código recebido:", codigo);
    setScanMensagem(null);

    const isbn = normalizeIsbn(codigo);
    const existente = isbn ? (livros || []).find((l) => l.isbn && normalizeIsbn(l.isbn) === isbn) : null;

    if (!existente) {
      setGestaoAba("catalogo");
      setScanMensagem({
        tipo: "nao_encontrado",
        texto: `Nenhum título com o código ${codigo} está cadastrado.`,
      });
      setTimeout(() => setScanMensagem(null), 6000);
      return;
    }

    if (existente.quantidade <= 0) {
      setGestaoAba("catalogo");
      setScanMensagem({
        tipo: "nao_encontrado",
        texto: `"${existente.titulo}" já está sem exemplares em estoque.`,
      });
      setTimeout(() => setScanMensagem(null), 6000);
      return;
    }

    // Encontrou e tem estoque — abre o modal de venda de verdade, pra
    // registrar quem vendeu e pra quem, igual ao botão "Vender" da ficha.
    setGestaoAba("catalogo");
    abrirVendaModal(existente.id);
  };

  const buscarCodigoManual = () => {
    if (!codigoManual.trim()) return;
    fecharScanner();
    buscarLivroPorCodigo(codigoManual.trim());
    setCodigoManual("");
  };

  // ---- derivados ----
  const totais = useMemo(() => {
    const l = livros || [];
    return {
      titulos: l.length,
      exemplares: l.reduce((acc, x) => acc + x.quantidade, 0),
      disponiveis: l.filter((x) => x.status === "disponivel").length,
      esgotados: l.filter((x) => x.status === "esgotado").length,
    };
  }, [livros]);

  const categorias = useMemo(() => {
    const set = new Set((livros || []).map((l) => l.categoria).filter(Boolean));
    return Array.from(set).sort((a, b) => a.localeCompare(b, "pt-BR"));
  }, [livros]);

  const filtrados = useMemo(() => {
    return (livros || [])
      .filter((l) => (filtro === "todos" ? true : l.status === filtro))
      .filter((l) => (categoriaFiltro === "todas" ? true : l.categoria === categoriaFiltro))
      .filter((l) => {
        const q = busca.trim().toLowerCase();
        if (!q) return true;
        return l.titulo.toLowerCase().includes(q) || l.autor.toLowerCase().includes(q);
      });
  }, [livros, busca, filtro, categoriaFiltro]);

  const movsRecentes = useMemo(() => (movs || []).slice(0, 8), [movs]);
  const livroPorId = (id) => (livros || []).find((l) => l.id === id);
  const livroVenda = vendaModal ? livroPorId(vendaModal) : null;

  const catalogoGestaoFiltrado = useMemo(() => {
    const q = gestaoBusca.trim().toLowerCase();
    return (livros || [])
      .filter((l) => (gestaoFiltroStatus === "todos" ? true : l.status === gestaoFiltroStatus))
      .filter((l) => (gestaoFiltroCategoria === "todas" ? true : l.categoria === gestaoFiltroCategoria))
      .filter((l) => (!q ? true : l.titulo.toLowerCase().includes(q) || l.autor.toLowerCase().includes(q)));
  }, [livros, gestaoBusca, gestaoFiltroStatus, gestaoFiltroCategoria]);

  const vendedores = useMemo(() => (usuarios || []).filter((u) => u.papel === "Liderado"), [usuarios]);

  const equipeFiltrada = useMemo(() => {
    const q = equipeBusca.trim().toLowerCase();
    if (!q) return usuarios || [];
    return (usuarios || []).filter((u) => u.nome.toLowerCase().includes(q));
  }, [usuarios, equipeBusca]);

  const compradores = useMemo(() => {
    const mapa = new Map();
    (movs || []).forEach((m) => {
      if (!m.tipo || !m.tipo.startsWith("venda")) return;
      const c = compradorDaMovimentacao(m);
      if (!c || !c.nome) return;
      const chave = `${c.nome.trim().toLowerCase()}|${c.telefone.trim()}`;
      const livro = livroPorId(m.livroId);
      const qtd = m.quantidadeVendida || 1;
      if (!mapa.has(chave)) {
        mapa.set(chave, {
          nome: c.nome,
          telefone: c.telefone,
          compras: 0,
          exemplares: 0,
          ultimaCompra: m.data,
          itens: [],
        });
      }
      const registro = mapa.get(chave);
      registro.compras += 1;
      registro.exemplares += qtd;
      if (new Date(m.data) > new Date(registro.ultimaCompra)) registro.ultimaCompra = m.data;
      registro.itens.push({ titulo: livro ? livro.titulo : "Livro removido", quantidade: qtd, data: m.data });
    });
    return Array.from(mapa.values()).sort((a, b) => new Date(b.ultimaCompra) - new Date(a.ultimaCompra));
  }, [movs, livros]);

  const sugestoesComprador = useMemo(() => {
    const q = vNome.trim().toLowerCase();
    if (!q) return [];
    return compradores.filter((c) => c.nome.toLowerCase().includes(q)).slice(0, 5);
  }, [compradores, vNome]);

  const vendasRegistradas = useMemo(() => {
    return (movs || [])
      .filter((m) => m.tipo && m.tipo.startsWith("venda"))
      .map((m) => {
        const livro = livroPorId(m.livroId);
        const c = compradorDaMovimentacao(m);
        return {
          id: m.id,
          data: m.data,
          quantidade: m.quantidadeVendida || 1,
          livroTitulo: livro ? livro.titulo : "Livro removido",
          compradorNome: c ? c.nome : "Não identificado",
          compradorTelefone: c ? c.telefone : "",
          vendedorNome: m.vendedorNome || "",
          desconto: m.desconto || null,
          valorTotal: m.valorTotal != null ? m.valorTotal : null,
        };
      })
      .sort((a, b) => new Date(b.data) - new Date(a.data));
  }, [movs, livros]);

  const totalExemplaresVendidos = useMemo(
    () => vendasRegistradas.reduce((acc, v) => acc + v.quantidade, 0),
    [vendasRegistradas]
  );

  const vendedoresComVenda = useMemo(() => {
    const set = new Set(vendasRegistradas.map((v) => v.vendedorNome).filter(Boolean));
    return Array.from(set).sort((a, b) => a.localeCompare(b, "pt-BR"));
  }, [vendasRegistradas]);

  const vendasFiltradas = useMemo(() => {
    const q = vendasBusca.trim().toLowerCase();
    const inicio = vendasDataInicio ? new Date(vendasDataInicio + "T00:00:00") : null;
    const fim = vendasDataFim ? new Date(vendasDataFim + "T23:59:59") : null;
    return vendasRegistradas.filter((v) => {
      if (q && !v.livroTitulo.toLowerCase().includes(q) && !v.compradorNome.toLowerCase().includes(q)) return false;
      if (vendasFiltroVendedor !== "todos" && v.vendedorNome !== vendasFiltroVendedor) return false;
      const data = new Date(v.data);
      if (inicio && data < inicio) return false;
      if (fim && data > fim) return false;
      return true;
    });
  }, [vendasRegistradas, vendasBusca, vendasDataInicio, vendasDataFim, vendasFiltroVendedor]);

  const resumoPeriodo = useMemo(() => {
    return vendasFiltradas.reduce(
      (acc, v) => ({
        vendas: acc.vendas + 1,
        exemplares: acc.exemplares + v.quantidade,
        total: acc.total + (v.valorTotal || 0),
        temValor: acc.temValor || v.valorTotal != null,
      }),
      { vendas: 0, exemplares: 0, total: 0, temValor: false }
    );
  }, [vendasFiltradas]);

  const compradoresFiltrados = useMemo(() => {
    const q = compradoresBusca.trim().toLowerCase();
    if (!q) return compradores;
    return compradores.filter((c) => c.nome.toLowerCase().includes(q) || c.telefone.includes(q));
  }, [compradores, compradoresBusca]);

  const contagemPorCategoria = useMemo(() => {
    const mapa = {};
    (livros || []).forEach((l) => {
      mapa[l.categoria] = (mapa[l.categoria] || 0) + 1;
    });
    return mapa;
  }, [livros]);

  const exportarCompradoresExcel = () => {
    const linhas = compradores.map((c) => ({
      Nome: c.nome,
      Telefone: c.telefone,
      "Total de compras": c.compras,
      "Exemplares comprados": c.exemplares,
      "Última compra": formatDate(c.ultimaCompra),
      "Livros comprados": c.itens.map((i) => `${i.titulo} (${i.quantidade})`).join("; "),
    }));
    const planilha = XLSX.utils.json_to_sheet(linhas);
    planilha["!cols"] = [{ wch: 28 }, { wch: 18 }, { wch: 14 }, { wch: 16 }, { wch: 14 }, { wch: 50 }];
    const livro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(livro, planilha, "Compradores");
    XLSX.writeFile(livro, "compradores-verbo-shop.xlsx");
  };

  const fontesCSS = `
    @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600;9..144,700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
    .stamp { transform: none; }
    .card-hole { box-shadow: inset 0 0 0 2px ${RULE}; }
    ::selection { background: ${ACCENT}33; }
  `;

  const itensMenu = [
    {
      key: "acervo",
      label: "Acervo",
      icon: <Home size={18} />,
      sub: `${totais.titulos} ${totais.titulos === 1 ? "título" : "títulos"}`,
    },
    { separador: true },
    {
      key: "scanner",
      label: "Ler código de barras",
      icon: <ScanBarcode size={18} />,
      sub: "Vender um título",
      acao: () => {
        setPagina("gestao");
        setGestaoAba("catalogo");
        abrirScanner();
      },
    },
    {
      key: "gestao",
      label: "Estoque",
      icon: <Settings2 size={18} />,
      sub: `${totais.exemplares} ${totais.exemplares === 1 ? "exemplar" : "exemplares"}`,
    },
    {
      key: "categorias",
      label: "Categorias",
      icon: <Tags size={18} />,
      sub: `${(categoriasCadastradas || []).length} ${(categoriasCadastradas || []).length === 1 ? "categoria" : "categorias"}`,
    },
    { separador: true },
    {
      key: "compradores",
      label: "Compradores",
      icon: <ShoppingBag size={18} />,
      sub: `${compradores.length} ${compradores.length === 1 ? "pessoa" : "pessoas"}`,
    },
    {
      key: "vendas",
      label: "Vendas",
      icon: <Receipt size={18} />,
      sub: `${totalExemplaresVendidos} ${totalExemplaresVendidos === 1 ? "exemplar vendido" : "exemplares vendidos"}`,
    },
    { separador: true },
    {
      key: "equipe",
      label: "Equipe",
      icon: <Users size={18} />,
      sub: `${(usuarios || []).length} ${(usuarios || []).length === 1 ? "pessoa" : "pessoas"}`,
    },
    { separador: true },
    {
      key: "perfil",
      label: "Meu perfil",
      icon: <UserCircle2 size={18} />,
      sub: perfil ? `${perfil.nome} · ${perfil.papel}` : "",
    },
  ];

  const sbErroBanner = sbErro && (
    <div className="fixed top-0 inset-x-0 z-50 flex justify-center px-4 pt-3 pointer-events-none">
      <div
        style={{ background: "#B04A4A", color: "#fff", fontFamily: MONO, fontSize: 12 }}
        className="pointer-events-auto rounded-sm px-4 py-2.5 shadow-md flex items-center gap-3 max-w-lg"
      >
        <CircleAlert size={16} className="shrink-0" />
        <span className="flex-1">{sbErro}</span>
        <button onClick={() => setSbErro(null)} aria-label="Fechar aviso" className="shrink-0">
          <X size={14} />
        </button>
      </div>
    </div>
  );

  const menuLateral = menuAberto && (
    <div className="fixed inset-0 z-50 flex">
      <div className="absolute inset-0 bg-black/40" onClick={() => setMenuAberto(false)} />
      <div
        style={{ background: PAPER, borderColor: RULE }}
        className="relative w-72 max-w-[85vw] h-full border-r flex flex-col p-5 overflow-y-auto"
      >
        <div className="flex items-center justify-between mb-5">
          <button
            onClick={() => {
              setPagina("acervo");
              setMenuAberto(false);
            }}
            className="flex items-center"
          >
            <img src={LOGO_URL} alt="Verbo Shop" style={{ height: 32 }} />
          </button>
          <button onClick={() => setMenuAberto(false)} style={{ color: INK_SOFT }} aria-label="Fechar menu">
            <X size={20} />
          </button>
        </div>

        <p style={{ fontFamily: MONO, fontSize: 10.5, color: INK_SOFT, letterSpacing: "0.08em" }} className="mb-2">
          GESTÃO
        </p>
        <div className="flex flex-col gap-1">
          {itensMenu.map((item, idx) =>
            item.separador ? (
              <hr key={`sep-${idx}`} style={{ borderColor: `${COR_MENU}33` }} className="my-2 border-t" />
            ) : (
              <button
                key={item.key}
                onClick={() => {
                  setMenuAberto(false);
                  if (item.acao) item.acao();
                  else setPagina(item.key);
                }}
                style={{
                  background: !item.acao && pagina === item.key ? `${COR_MENU}18` : "transparent",
                  color: !item.acao && pagina === item.key ? COR_MENU : INK,
                }}
                className="flex items-center gap-3 rounded-sm px-3 py-2.5 text-left hover:opacity-80 transition"
              >
                <span style={{ color: !item.acao && pagina === item.key ? COR_MENU : INK_SOFT }}>{item.icon}</span>
                <span className="flex-1 min-w-0">
                  <span style={{ fontFamily: SANS, fontSize: 14, fontWeight: 600 }} className="block">
                    {item.label}
                  </span>
                  <span style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT }} className="block truncate">
                    {item.sub}
                  </span>
                </span>
              </button>
            )
          )}
        </div>
      </div>
    </div>
  );

  const TopBar = () => (
    <div className="relative flex items-center justify-between mb-5 gap-2">
      <button
        onClick={() => setMenuAberto(true)}
        style={{ borderColor: COR_MENU, background: CARD, color: COR_MENU }}
        className="border rounded-sm w-9 h-9 flex items-center justify-center hover:opacity-80 transition shrink-0"
        aria-label="Abrir menu"
      >
        <Menu size={17} />
      </button>
      <button
        onClick={() => setPagina("acervo")}
        className="absolute left-1/2 top-1/2"
        style={{ transform: "translate(-50%, -50%)" }}
        aria-label="Ir para o acervo"
      >
        <img src={LOGO_URL} alt="Verbo Shop" style={{ height: 32 }} />
      </button>
      <button
        onClick={() => setPagina("perfil")}
        style={{
          borderColor: pagina === "perfil" ? COR_MENU : RULE,
          background: pagina === "perfil" ? `${COR_MENU}18` : CARD,
          color: COR_MENU,
        }}
        className="border rounded-full w-9 h-9 flex items-center justify-center hover:opacity-80 transition shrink-0"
        aria-label="Gestão do perfil"
        title={perfil ? `${perfil.nome} · ${perfil.papel}` : "Gestão do perfil"}
      >
        {perfil && perfil.nome ? (
          <span style={{ fontFamily: SERIF, fontWeight: 700, fontSize: 14 }}>
            {perfil.nome.trim().charAt(0).toUpperCase()}
          </span>
        ) : (
          <UserCircle2 size={18} />
        )}
      </button>
    </div>
  );

  const botaoVoltarAcervo = pagina !== "acervo" && (
    <div className="fixed bottom-5 inset-x-0 flex justify-center z-30 px-5 pointer-events-none">
      <button
        onClick={() => setPagina("acervo")}
        style={{ fontFamily: MONO, fontSize: 13, color: INK, borderColor: RULE, background: CARD }}
        className="pointer-events-auto flex items-center gap-1.5 border rounded-full px-4 py-2.5 shadow-md hover:opacity-90 transition"
      >
        <ArrowLeft size={15} />
        Voltar ao acervo
      </button>
    </div>
  );

  const vendaModalOverlay = vendaModal && livroVenda && (
    <div className="fixed inset-0 bg-black/30 flex items-center justify-center px-5 z-20">
      <div style={{ background: CARD, borderColor: RULE }} className="border rounded-md p-5 w-full max-w-sm">
        <h3 style={{ fontFamily: SERIF, fontSize: 18, fontWeight: 600 }} className="mb-1">Registrar venda</h3>
        <p style={{ color: INK_SOFT, fontSize: 13 }} className="mb-3">
          {livroVenda.titulo} · {livroVenda.quantidade} em estoque
        </p>
        <label style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT }}>Quantidade vendida</label>
        <div className="mt-1 mb-3">
          <QtyStepper value={vQtd} onChange={setVQtd} min={1} max={livroVenda.quantidade} color={ACCENT} />
        </div>

        {vendedores.length === 0 ? (
          <p style={{ fontFamily: MONO, fontSize: 11.5, color: "#B04A4A" }} className="mb-3">
            Nenhum liderado cadastrado ainda. Abra o menu (ícone ☰) → Gestão da equipe para cadastrar.
          </p>
        ) : (
          <>
            <label style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT }}>Vendido por *</label>
            <select
              value={vVendedorId}
              onChange={(e) => setVVendedorId(e.target.value)}
              style={{ borderColor: RULE, background: PAPER, fontFamily: SANS, color: vVendedorId ? INK : INK_SOFT }}
              className="border rounded-sm px-3 py-2 text-sm w-full mb-3 mt-1"
            >
              <option value="">Selecionar…</option>
              {vendedores.map((v) => (
                <option key={v.id} value={v.id}>{v.nome}</option>
              ))}
            </select>
          </>
        )}

        <div className="relative mb-2">
          <input
            placeholder="Nome do comprador *"
            value={vNome}
            onChange={(e) => {
              setVNome(e.target.value);
              setVSugestoesAbertas(true);
            }}
            onFocus={() => setVSugestoesAbertas(true)}
            onBlur={() => setTimeout(() => setVSugestoesAbertas(false), 150)}
            style={{ borderColor: RULE, background: PAPER, fontFamily: SANS }}
            className="border rounded-sm px-3 py-2 text-sm w-full"
          />
          {vSugestoesAbertas && sugestoesComprador.length > 0 && (
            <div
              style={{ background: CARD, borderColor: RULE }}
              className="absolute left-0 right-0 top-full mt-1 border rounded-sm shadow-md z-10 overflow-hidden"
            >
              {sugestoesComprador.map((c) => (
                <button
                  key={`${c.nome}|${c.telefone}`}
                  onClick={() => {
                    setVNome(c.nome);
                    setVTelefone(c.telefone);
                    setVSugestoesAbertas(false);
                  }}
                  style={{ fontFamily: SANS }}
                  className="w-full text-left px-3 py-2 text-sm hover:opacity-80 transition flex items-center justify-between gap-2"
                >
                  <span className="truncate">{c.nome}</span>
                  <span style={{ fontFamily: MONO, fontSize: 10.5, color: ACCENT }} className="shrink-0">
                    já comprou {c.compras}×
                  </span>
                </button>
              ))}
            </div>
          )}
        </div>
        <input
          type="tel"
          placeholder="Telefone *"
          value={vTelefone}
          onChange={(e) => setVTelefone(maskTelefone(e.target.value))}
          style={{ borderColor: RULE, background: PAPER, fontFamily: SANS }}
          className="border rounded-sm px-3 py-2 text-sm w-full mb-3"
        />

        <label style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT }}>Desconto (opcional)</label>
        <input
          type="text"
          inputMode="decimal"
          placeholder="R$ 0,00"
          value={vDesconto}
          onChange={(e) => setVDesconto(e.target.value.replace(/[^0-9,.]/g, ""))}
          onKeyDown={(e) => e.key === "Enter" && confirmarVenda()}
          style={{ borderColor: RULE, background: PAPER, fontFamily: MONO }}
          className="border rounded-sm px-3 py-2 text-sm w-full mb-3 mt-1"
        />

        {livroVenda.preco != null &&
          (() => {
            const qtdAtual = Math.max(1, Math.min(parseInt(vQtd, 10) || 1, livroVenda.quantidade));
            const subtotal = livroVenda.preco * qtdAtual;
            const descontoAtual = Math.max(0, parseFloat((vDesconto || "0").replace(",", ".")) || 0);
            const total = Math.max(0, subtotal - descontoAtual);
            return (
              <div style={{ background: PAPER, borderColor: RULE }} className="border rounded-sm px-3 py-2.5 mb-4 flex flex-col gap-1">
                <div className="flex justify-between" style={{ fontFamily: MONO, fontSize: 11.5, color: INK_SOFT }}>
                  <span>Subtotal ({qtdAtual} × {formatBRL(livroVenda.preco)})</span>
                  <span>{formatBRL(subtotal)}</span>
                </div>
                {descontoAtual > 0 && (
                  <div className="flex justify-between" style={{ fontFamily: MONO, fontSize: 11.5, color: COR_ESTOQUE_BAIXO }}>
                    <span>Desconto</span>
                    <span>− {formatBRL(descontoAtual)}</span>
                  </div>
                )}
                <div className="flex justify-between" style={{ fontFamily: SERIF, fontWeight: 700, fontSize: 16, color: ACCENT }}>
                  <span>Total</span>
                  <span>{formatBRL(total)}</span>
                </div>
              </div>
            );
          })()}

        <div className="flex gap-2 justify-end">
          <button onClick={() => setVendaModal(null)} style={{ fontFamily: MONO, fontSize: 12.5, color: INK_SOFT }} className="px-3 py-2">
            Cancelar
          </button>
          <button
            onClick={confirmarVenda}
            disabled={!vNome.trim() || !vTelefone.trim() || !vVendedorId}
            style={{
              background: ACCENT,
              color: "#fff",
              fontFamily: MONO,
              fontSize: 12.5,
              opacity: !vNome.trim() || !vTelefone.trim() || !vVendedorId ? 0.4 : 1,
            }}
            className="rounded-sm px-3 py-2"
          >
            Confirmar venda
          </button>
        </div>
      </div>
    </div>
  );

  const scannerOverlay = scannerAberto && (
    <div className="fixed inset-0 bg-black flex flex-col z-40">
      <div className="flex items-center justify-between gap-3 px-4 py-4">
        <div className="flex items-center gap-3">
          <button onClick={fecharScanner} style={{ color: "#fff" }}>
            <ArrowLeft size={20} />
          </button>
          <p style={{ color: "#fff", fontFamily: MONO, fontSize: 13 }}>Aponte para vender</p>
        </div>
        {torchDisponivel && (
          <button
            onClick={alternarTorch}
            style={{ color: torchLigado ? "#F5C77E" : "#fff", borderColor: "#444" }}
            className="border rounded-full w-9 h-9 flex items-center justify-center shrink-0"
            aria-label="Lanterna"
          >
            <Lightbulb size={16} />
          </button>
        )}
      </div>

      <div className="flex-1 relative flex items-center justify-center overflow-hidden">
        <div id={SCANNER_DIV_ID} className="w-full h-full" />
        {(scanStatus === "carregando" || scanStatus === "starting") && (
          <div className="absolute inset-0 flex items-center justify-center">
            <Loader2 size={28} className="animate-spin" style={{ color: "#F5C77E" }} />
          </div>
        )}
        <div
          className="absolute border-2 rounded-md pointer-events-none"
          style={{ borderColor: scanStatus === "scanning" ? "#8FE3A5" : "#F5C77E", width: "82%", height: 105 }}
        />
        {scanStatus === "scanning" && (
          <p
            style={{ color: "#fff", fontFamily: MONO, fontSize: 11 }}
            className="absolute bottom-3 left-0 right-0 text-center px-6"
          >
            Aproxime a ~10–15cm, boa luz, alinhe o código na caixa
          </p>
        )}
      </div>

      <div className="px-5 py-5" style={{ background: "#111" }}>
        {scanStatus === "carregando" && (
          <p style={{ color: "#F5C77E", fontFamily: MONO, fontSize: 12.5 }} className="flex items-center gap-2 mb-3">
            <Loader2 size={14} className="animate-spin" /> Carregando leitor…
          </p>
        )}
        {scanStatus === "starting" && (
          <p style={{ color: "#F5C77E", fontFamily: MONO, fontSize: 12.5 }} className="flex items-center gap-2 mb-3">
            <Loader2 size={14} className="animate-spin" /> Abrindo câmera…
          </p>
        )}
        {scanStatus === "scanning" && (
          <p style={{ color: "#8FE3A5", fontFamily: MONO, fontSize: 12.5 }} className="flex items-center gap-2 mb-3">
            <CircleCheck size={14} /> Procurando código…
          </p>
        )}
        {scanStatus === "error" && (
          <div className="mb-3">
            <p style={{ color: "#F0A0A0", fontFamily: MONO, fontSize: 12.5 }} className="flex items-center gap-2">
              <CircleAlert size={14} />
              Não foi possível acessar a câmera ou carregar o leitor. Digite o código abaixo ou confira as permissões.
            </p>
            {scanErroDetalhe && (
              <p style={{ color: "#F0A0A0", fontFamily: MONO, fontSize: 10.5, opacity: 0.8 }} className="mt-1 pl-5">
                Detalhe: {scanErroDetalhe}
              </p>
            )}
          </div>
        )}
        <div className="flex gap-2">
          <input
            placeholder="Digitar ISBN pra vender"
            value={codigoManual}
            onChange={(e) => setCodigoManual(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && buscarCodigoManual()}
            style={{ fontFamily: MONO, background: "#1E1E1E", color: "#fff", borderColor: "#333" }}
            className="border rounded-sm px-3 py-2 text-sm flex-1"
          />
          <button
            onClick={buscarCodigoManual}
            style={{ background: ACCENT, color: "#fff", fontFamily: MONO, fontSize: 12.5 }}
            className="rounded-sm px-3 py-2"
          >
            Vender
          </button>
        </div>
      </div>
    </div>
  );

  if (autenticando) {
    return (
      <div style={{ background: PAPER, minHeight: "100vh" }} className="flex items-center justify-center">
        <Loader2 className="animate-spin" size={28} style={{ color: INK }} />
      </div>
    );
  }

  if (!sessao) {
    return (
      <div style={{ background: PAPER, minHeight: "100vh", color: INK, fontFamily: SANS }} className="flex items-center justify-center px-5">
        <style>{fontesCSS}</style>
        <div className="w-full max-w-sm">
          <div className="flex justify-center mb-6">
            <img src={LOGO_URL} alt="Verbo Shop" style={{ height: 48 }} />
          </div>
          <div style={{ background: CARD, borderColor: RULE }} className="border rounded-md p-6">
            <h1 style={{ fontFamily: SERIF, fontWeight: 600, fontSize: 22 }} className="mb-1">Entrar</h1>
            <p style={{ color: INK_SOFT, fontSize: 13 }} className="mb-5">Use o e-mail e senha da sua conta da equipe.</p>

            <label style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT }}>E-mail</label>
            <input
              type="email"
              autoFocus
              value={loginEmail}
              onChange={(e) => setLoginEmail(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && fazerLogin()}
              style={{ fontFamily: SANS, borderColor: RULE, background: PAPER }}
              className="border rounded-sm px-3 py-2 text-sm w-full mt-1 mb-3"
            />

            <label style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT }}>Senha</label>
            <div className="relative mt-1 mb-4">
              <input
                type={verLoginSenha ? "text" : "password"}
                value={loginSenha}
                onChange={(e) => setLoginSenha(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && fazerLogin()}
                style={{ fontFamily: SANS, borderColor: RULE, background: PAPER }}
                className="border rounded-sm pl-3 pr-10 py-2 text-sm w-full"
              />
              <button
                type="button"
                onClick={() => setVerLoginSenha((v) => !v)}
                style={{ color: INK_SOFT }}
                className="absolute right-3 top-1/2 -translate-y-1/2"
                aria-label={verLoginSenha ? "Esconder senha" : "Mostrar senha"}
              >
                {verLoginSenha ? <EyeOff size={16} /> : <Eye size={16} />}
              </button>
            </div>

            {loginErro && (
              <p style={{ fontFamily: MONO, fontSize: 12, color: "#B04A4A" }} className="mb-3">{loginErro}</p>
            )}

            <button
              onClick={fazerLogin}
              disabled={loginCarregando || !loginEmail.trim() || !loginSenha}
              style={{
                background: ACCENT,
                color: "#fff",
                fontFamily: MONO,
                fontSize: 13,
                opacity: loginCarregando || !loginEmail.trim() || !loginSenha ? 0.5 : 1,
              }}
              className="rounded-sm px-3 py-2.5 w-full flex items-center justify-center gap-2 hover:opacity-90 transition"
            >
              {loginCarregando && <Loader2 size={14} className="animate-spin" />}
              {loginCarregando ? "Entrando…" : "Entrar"}
            </button>
          </div>
          <p style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT }} className="text-center mt-4">
            Não tem uma conta? Peça pro Líder da equipe te cadastrar.
          </p>
        </div>
      </div>
    );
  }

  if (livros === null) {
    return (
      <div style={{ background: PAPER, minHeight: "100vh" }} className="flex items-center justify-center">
        <Loader2 className="animate-spin" size={28} style={{ color: INK }} />
      </div>
    );
  }

  // ---- página: Gestão de estoque (cabeçalho no mesmo estilo do acervo) ----
  if (pagina === "gestao") {
    return (
      <div style={{ background: PAPER, minHeight: "100vh", color: INK, fontFamily: SANS }}>
        <style>{fontesCSS}</style>

        <header className="px-5 pt-8 pb-6 max-w-5xl mx-auto">
          <TopBar />

          <h1 style={{ fontFamily: SERIF, fontWeight: 600, fontSize: "clamp(24px, 7vw, 32px)", lineHeight: 1.1, marginTop: 6 }}>
            Estoque
          </h1>
          <p style={{ color: INK_SOFT, fontSize: 14, marginTop: 4 }}>
            {totais.titulos} {totais.titulos === 1 ? "título cadastrado" : "títulos cadastrados"} · {totais.exemplares}{" "}
            {totais.exemplares === 1 ? "exemplar em estoque" : "exemplares em estoque"}
            {saving && <span style={{ opacity: 0.5 }}> · salvando…</span>}
          </p>

          {/* Abas */}
          <div className="flex gap-2 mt-5">
            {[
              { key: "catalogo", label: "Catálogo" },
              { key: "cadastrar", label: "Cadastrar novo" },
              { key: "importar", label: "Importar lista" },
            ].map((aba) => (
              <button
                key={aba.key}
                onClick={() => setGestaoAba(aba.key)}
                style={{
                  fontFamily: MONO,
                  fontSize: 13,
                  background: gestaoAba === aba.key ? ACCENT : CARD,
                  color: gestaoAba === aba.key ? "#fff" : INK_SOFT,
                  borderColor: gestaoAba === aba.key ? ACCENT : RULE,
                }}
                className="border rounded-sm px-4 py-2.5 transition hover:opacity-90"
              >
                {aba.label}
              </button>
            ))}
          </div>

          {/* Filtros — só fazem sentido no catálogo */}
          {gestaoAba === "catalogo" && (
            <>
              <div className="flex flex-wrap gap-2 mt-5">
                {[
                  { key: "todos", label: "Todos", n: totais.titulos, color: INK },
                  { key: "disponivel", label: STATUS.disponivel.label, n: totais.disponiveis, color: STATUS.disponivel.color },
                  { key: "esgotado", label: STATUS.esgotado.label, n: totais.esgotados, color: STATUS.esgotado.color },
                ].map((s) => (
                  <button
                    key={s.key}
                    onClick={() => setGestaoFiltroStatus(s.key)}
                    style={{
                      fontFamily: MONO,
                      fontSize: 12.5,
                      borderColor: gestaoFiltroStatus === s.key ? s.color : RULE,
                      color: gestaoFiltroStatus === s.key ? s.color : INK_SOFT,
                      background: gestaoFiltroStatus === s.key ? "#FFFFFFAA" : "transparent",
                    }}
                    className="border rounded-full px-3 py-1.5 transition"
                  >
                    {s.label} <span style={{ opacity: 0.6 }}>· {s.n}</span>
                  </button>
                ))}
              </div>

              {categorias.length > 0 && (
                <div className="flex items-center gap-2 mt-3">
                  <span style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT, letterSpacing: "0.06em" }}>
                    CATEGORIA
                  </span>
                  <select
                    value={gestaoFiltroCategoria}
                    onChange={(e) => setGestaoFiltroCategoria(e.target.value)}
                    style={{ fontFamily: MONO, fontSize: 12.5, borderColor: RULE, background: CARD, color: INK }}
                    className="border rounded-full px-3 py-1.5"
                  >
                    <option value="todas">Todas</option>
                    {categorias.map((c) => (
                      <option key={c} value={c}>{c}</option>
                    ))}
                  </select>
                </div>
              )}
            </>
          )}

          {/* Busca — só faz sentido no catálogo */}
          {gestaoAba === "catalogo" && (
            <div className="flex items-center gap-2 mt-4 border-b pb-3" style={{ borderColor: RULE }}>
              <Search size={16} style={{ color: INK_SOFT }} />
              <input
                placeholder="Buscar por título ou autor…"
                value={gestaoBusca}
                onChange={(e) => setGestaoBusca(e.target.value)}
                style={{ fontFamily: SANS, background: "transparent" }}
                className="text-sm flex-1 outline-none py-1"
              />
            </div>
          )}
        </header>

        <main className="px-5 pb-16 max-w-5xl mx-auto">
          {gestaoAba === "catalogo" && scanMensagem && (
            <div
              style={{
                background:
                  scanMensagem.tipo === "exemplar" ? `${STATUS.disponivel.color}18` : `${COR_ESTOQUE_BAIXO}18`,
                color: scanMensagem.tipo === "exemplar" ? STATUS.disponivel.color : COR_ESTOQUE_BAIXO,
                fontFamily: MONO,
                fontSize: 12,
              }}
              className="rounded-sm px-3 py-2.5 mb-4"
            >
              {scanMensagem.texto}
            </div>
          )}
          {gestaoAba === "catalogo" && (
            <div className="flex flex-col gap-2">
              {catalogoGestaoFiltrado.length === 0 && (
                <p style={{ color: INK_SOFT, fontSize: 13 }} className="py-6 text-center">
                  Nenhum título encontrado.
                </p>
              )}
              {catalogoGestaoFiltrado.map((livro) => (
                <div
                  key={livro.id}
                  style={{ background: CARD, borderColor: RULE }}
                  className="border rounded-sm p-3 flex items-center gap-3 flex-wrap sm:flex-nowrap"
                >
                  <button
                    onClick={() => abrirEdicaoLivro(livro)}
                    className="flex-1 min-w-[140px] text-left hover:opacity-75 transition"
                  >
                    <p style={{ fontFamily: SANS, fontSize: 13.5, fontWeight: 600 }} className="truncate">
                      {livro.titulo}
                    </p>
                    <p style={{ color: INK_SOFT, fontSize: 12 }} className="truncate">
                      {livro.autor} · {livro.categoria}
                      {formatBRL(livro.preco) && ` · ${formatBRL(livro.preco)}`}
                    </p>
                  </button>
                  <QtyStepper
                    value={String(livro.quantidade)}
                    onChange={(v) => editarQuantidade(livro.id, v)}
                    min={0}
                    color={ACCENT}
                  />
                  <button
                    onClick={() => abrirEdicaoLivro(livro)}
                    style={{ color: ACCENT }}
                    className="p-1.5 shrink-0 hover:opacity-70 transition"
                    aria-label={`Editar ${livro.titulo}`}
                  >
                    <Pencil size={16} />
                  </button>
                  {confirmExcluir === livro.id ? (
                    <div className="flex gap-1 shrink-0">
                      <button
                        onClick={() => excluirLivro(livro.id)}
                        style={{ background: "#B04A4A", color: "#fff", fontFamily: MONO, fontSize: 11 }}
                        className="rounded-sm px-2 py-1.5"
                      >
                        Excluir
                      </button>
                      <button
                        onClick={() => setConfirmExcluir(null)}
                        style={{ color: INK_SOFT, fontFamily: MONO, fontSize: 11 }}
                        className="px-2 py-1.5"
                      >
                        Cancelar
                      </button>
                    </div>
                  ) : (
                    <button
                      onClick={() => setConfirmExcluir(livro.id)}
                      style={{ color: "#B04A4A" }}
                      className="p-1.5 shrink-0 hover:opacity-70 transition"
                      aria-label={`Excluir ${livro.titulo}`}
                    >
                      <Trash2 size={16} />
                    </button>
                  )}
                </div>
              ))}
            </div>
          )}

          {gestaoAba === "cadastrar" && (
            <div className="max-w-md">
              <div className="flex flex-col gap-3">
                <input
                  placeholder="Título *"
                  value={novo.titulo}
                  onChange={(e) => setNovo({ ...novo, titulo: e.target.value })}
                  style={{ fontFamily: SANS, borderColor: RULE, background: CARD }}
                  className="border rounded-sm px-3 py-2 text-sm"
                />
                <input
                  placeholder="Autor"
                  value={novo.autor}
                  onChange={(e) => setNovo({ ...novo, autor: e.target.value })}
                  style={{ fontFamily: SANS, borderColor: RULE, background: CARD }}
                  className="border rounded-sm px-3 py-2 text-sm"
                />
                <select
                  value={novo.categoria}
                  onChange={(e) => setNovo({ ...novo, categoria: e.target.value })}
                  style={{ fontFamily: SANS, borderColor: RULE, background: CARD, color: novo.categoria ? INK : INK_SOFT }}
                  className="border rounded-sm px-3 py-2 text-sm"
                >
                  <option value="">Categoria…</option>
                  {(categoriasCadastradas || []).map((c) => (
                    <option key={c} value={c}>{c}</option>
                  ))}
                </select>
                <input
                  placeholder="ISBN (opcional)"
                  value={novo.isbn}
                  onChange={(e) => setNovo({ ...novo, isbn: e.target.value })}
                  style={{ fontFamily: MONO, borderColor: RULE, background: CARD }}
                  className="border rounded-sm px-3 py-2 text-sm"
                />
                <div>
                  <label style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT }}>Preço (R$)</label>
                  <input
                    type="text"
                    inputMode="decimal"
                    placeholder="0,00"
                    value={novo.preco}
                    onChange={(e) => setNovo({ ...novo, preco: e.target.value.replace(/[^0-9,.]/g, "") })}
                    style={{ fontFamily: MONO, borderColor: RULE, background: CARD }}
                    className="border rounded-sm px-3 py-2 text-sm w-full mt-1"
                  />
                </div>
                <div>
                  <label style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT }}>Quantidade inicial</label>
                  <div className="mt-1">
                    <QtyStepper
                      value={novo.quantidade}
                      onChange={(v) => setNovo({ ...novo, quantidade: v })}
                      min={0}
                      color={ACCENT}
                    />
                  </div>
                </div>
                <button
                  onClick={adicionarLivro}
                  style={{ background: ACCENT, color: "#fff", fontFamily: MONO, fontSize: 13 }}
                  className="rounded-sm px-3 py-2.5 hover:opacity-90 transition"
                >
                  Catalogar livro
                </button>
              </div>
              <p style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT }} className="mt-3">
                Preencher o ISBN aqui permite que, no futuro, escanear o código de barras desse livro já abra o
                registro de venda (menu → Ler código de barras).
              </p>
            </div>
          )}

          {gestaoAba === "importar" && (
            <div className="max-w-md">
              <div style={{ background: CARD, borderColor: RULE }} className="border rounded-sm p-4">
                <p style={{ fontFamily: SANS, fontSize: 13, fontWeight: 600 }} className="mb-2">
                  Como formatar sua planilha
                </p>
                <p style={{ fontFamily: SANS, fontSize: 12.5, color: INK_SOFT }} className="mb-2">
                  Use uma linha de cabeçalho com estas colunas (a ordem não importa; só{" "}
                  <strong style={{ color: INK }}>Título</strong> é obrigatório):
                </p>
                <div className="overflow-x-auto mb-3">
                  <table style={{ fontFamily: MONO, fontSize: 11 }} className="w-full border-collapse">
                    <thead>
                      <tr>
                        {["Título", "Autor", "Categoria", "ISBN", "Quantidade", "Preço"].map((c) => (
                          <th
                            key={c}
                            style={{ borderColor: RULE, color: INK_SOFT }}
                            className="border px-2 py-1 text-left whitespace-nowrap"
                          >
                            {c}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      <tr>
                        {["O Novo Nascimento", "Kenneth E. Hagin", "Fé", "9788578600000", "5", "29,90"].map(
                          (v, i) => (
                            <td key={i} style={{ borderColor: RULE, color: INK }} className="border px-2 py-1 whitespace-nowrap">
                              {v}
                            </td>
                          )
                        )}
                      </tr>
                    </tbody>
                  </table>
                </div>
                <ul style={{ fontFamily: SANS, fontSize: 12, color: INK_SOFT }} className="list-disc pl-4 mb-3 flex flex-col gap-1">
                  <li>Categoria vazia vira "Sem categoria"; categorias novas são cadastradas automaticamente.</li>
                  <li>Se o ISBN já existir no seu catálogo, a Quantidade é somada ao estoque existente em vez de duplicar o título.</li>
                  <li>Formatos aceitos: .xlsx, .xls ou .csv.</li>
                </ul>
                <button
                  onClick={baixarModeloPlanilha}
                  style={{ fontFamily: MONO, fontSize: 12, color: ACCENT, borderColor: ACCENT }}
                  className="border rounded-full px-3 py-1.5 flex items-center gap-1.5 hover:opacity-80 transition mb-4"
                >
                  <Download size={13} />
                  Baixar planilha modelo (.xlsx)
                </button>

                <label style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT, letterSpacing: "0.06em" }} className="block mb-2">
                  ARQUIVO DA PLANILHA
                </label>

                <input
                  ref={importInputRef}
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  onChange={(e) => {
                    const arquivo = e.target.files?.[0];
                    if (arquivo) {
                      setImportResultado(null);
                      setArquivoSelecionado(arquivo);
                    }
                  }}
                  className="hidden"
                />

                {!arquivoSelecionado ? (
                  <button
                    onClick={() => importInputRef.current?.click()}
                    style={{ borderColor: RULE, background: PAPER, color: INK }}
                    className="border border-dashed rounded-sm w-full py-6 flex flex-col items-center justify-center gap-2 hover:opacity-80 transition"
                  >
                    <FileSpreadsheet size={22} style={{ color: ACCENT }} />
                    <span style={{ fontFamily: MONO, fontSize: 12.5 }}>Escolher arquivo</span>
                    <span style={{ fontFamily: SANS, fontSize: 11, color: INK_SOFT }}>.xlsx, .xls ou .csv</span>
                  </button>
                ) : (
                  <div style={{ borderColor: RULE, background: PAPER }} className="border rounded-sm p-3">
                    <div className="flex items-center gap-3">
                      <span style={{ background: `${ACCENT}18`, color: ACCENT }} className="rounded-sm p-2 shrink-0">
                        <FileSpreadsheet size={18} />
                      </span>
                      <div className="min-w-0 flex-1">
                        <p style={{ fontFamily: SANS, fontSize: 13, fontWeight: 600 }} className="truncate">
                          {arquivoSelecionado.name}
                        </p>
                        <p style={{ fontFamily: MONO, fontSize: 10.5, color: INK_SOFT }}>
                          {(arquivoSelecionado.size / 1024).toFixed(0)} KB
                        </p>
                      </div>
                      {!importProcessando && (
                        <button
                          onClick={() => {
                            setArquivoSelecionado(null);
                            if (importInputRef.current) importInputRef.current.value = "";
                          }}
                          style={{ color: INK_SOFT }}
                          className="p-1 shrink-0 hover:opacity-70 transition"
                          aria-label="Remover arquivo"
                        >
                          <X size={16} />
                        </button>
                      )}
                    </div>

                    {!importProcessando && (
                      <p style={{ fontFamily: SANS, fontSize: 12, color: INK_SOFT }} className="mt-3">
                        Confirma a importação? Os livros dessa planilha serão adicionados ao catálogo.
                      </p>
                    )}

                    <div className="flex gap-2 justify-end mt-3">
                      {!importProcessando && (
                        <button
                          onClick={() => importInputRef.current?.click()}
                          style={{ fontFamily: MONO, fontSize: 12, color: INK_SOFT }}
                          className="px-3 py-2"
                        >
                          Trocar arquivo
                        </button>
                      )}
                      <button
                        onClick={() => importarPlanilha(arquivoSelecionado)}
                        disabled={importProcessando}
                        style={{
                          background: ACCENT,
                          color: "#fff",
                          fontFamily: MONO,
                          fontSize: 12.5,
                          opacity: importProcessando ? 0.6 : 1,
                        }}
                        className="rounded-sm px-4 py-2 flex items-center gap-2"
                      >
                        {importProcessando && <Loader2 size={14} className="animate-spin" />}
                        {importProcessando ? "Importando…" : "Confirmar importação"}
                      </button>
                    </div>
                  </div>
                )}

                {importResultado && !importResultado.erro && (
                  <p style={{ fontFamily: MONO, fontSize: 12, color: STATUS.disponivel.color }} className="mt-3">
                    {importResultado.adicionados} {importResultado.adicionados === 1 ? "título novo" : "títulos novos"} ·{" "}
                    {importResultado.atualizados} {importResultado.atualizados === 1 ? "atualizado" : "atualizados"}
                    {importResultado.ignorados > 0 && ` · ${importResultado.ignorados} ignorado(s) sem título`}
                  </p>
                )}
                {importResultado?.erro && (
                  <p style={{ fontFamily: MONO, fontSize: 12, color: "#B04A4A" }} className="mt-3">
                    {importResultado.erro}
                  </p>
                )}
              </div>
            </div>
          )}
        </main>

        {/* Modal de edição de livro */}
        {livroEditandoId && (
          <div className="fixed inset-0 bg-black/30 flex items-center justify-center px-5 z-20">
            <div style={{ background: CARD, borderColor: RULE }} className="border rounded-md p-5 w-full max-w-sm">
              <h3 style={{ fontFamily: SERIF, fontSize: 18, fontWeight: 600 }} className="mb-3">Editar livro</h3>
              <div className="flex flex-col gap-3">
                <input
                  placeholder="Título *"
                  value={edicaoForm.titulo}
                  onChange={(e) => setEdicaoForm({ ...edicaoForm, titulo: e.target.value })}
                  style={{ fontFamily: SANS, borderColor: RULE, background: PAPER }}
                  className="border rounded-sm px-3 py-2 text-sm"
                />
                <input
                  placeholder="Autor"
                  value={edicaoForm.autor}
                  onChange={(e) => setEdicaoForm({ ...edicaoForm, autor: e.target.value })}
                  style={{ fontFamily: SANS, borderColor: RULE, background: PAPER }}
                  className="border rounded-sm px-3 py-2 text-sm"
                />
                <select
                  value={edicaoForm.categoria}
                  onChange={(e) => setEdicaoForm({ ...edicaoForm, categoria: e.target.value })}
                  style={{ fontFamily: SANS, borderColor: RULE, background: PAPER, color: edicaoForm.categoria ? INK : INK_SOFT }}
                  className="border rounded-sm px-3 py-2 text-sm"
                >
                  <option value="">Categoria…</option>
                  {(categoriasCadastradas || []).map((c) => (
                    <option key={c} value={c}>{c}</option>
                  ))}
                </select>
                <input
                  placeholder="ISBN (opcional)"
                  value={edicaoForm.isbn}
                  onChange={(e) => setEdicaoForm({ ...edicaoForm, isbn: e.target.value })}
                  style={{ fontFamily: MONO, borderColor: RULE, background: PAPER }}
                  className="border rounded-sm px-3 py-2 text-sm"
                />
                <div>
                  <label style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT }}>Preço (R$)</label>
                  <input
                    type="text"
                    inputMode="decimal"
                    placeholder="0,00"
                    value={edicaoForm.preco}
                    onChange={(e) => setEdicaoForm({ ...edicaoForm, preco: e.target.value.replace(/[^0-9,.]/g, "") })}
                    onKeyDown={(e) => e.key === "Enter" && salvarEdicaoLivro()}
                    style={{ fontFamily: MONO, borderColor: RULE, background: PAPER }}
                    className="border rounded-sm px-3 py-2 text-sm w-full mt-1"
                  />
                </div>
              </div>
              <div className="flex gap-2 justify-end mt-4">
                <button
                  onClick={() => setLivroEditandoId(null)}
                  style={{ fontFamily: MONO, fontSize: 12.5, color: INK_SOFT }}
                  className="px-3 py-2"
                >
                  Cancelar
                </button>
                <button
                  onClick={salvarEdicaoLivro}
                  disabled={!edicaoForm.titulo.trim()}
                  style={{
                    background: ACCENT,
                    color: "#fff",
                    fontFamily: MONO,
                    fontSize: 12.5,
                    opacity: !edicaoForm.titulo.trim() ? 0.4 : 1,
                  }}
                  className="rounded-sm px-3 py-2"
                >
                  Salvar alterações
                </button>
              </div>
            </div>
          </div>
        )}

        {vendaModalOverlay}
        {scannerOverlay}
        {menuLateral}
        {botaoVoltarAcervo}
        {sbErroBanner}
      </div>
    );
  }

  // ---- página: Compradores ----
  if (pagina === "compradores") {
    return (
      <div style={{ background: PAPER, minHeight: "100vh", color: INK, fontFamily: SANS }}>
        <style>{fontesCSS}</style>
        <header className="px-5 pt-8 pb-6 max-w-5xl mx-auto">
          <TopBar />

          <h1 style={{ fontFamily: SERIF, fontWeight: 600, fontSize: "clamp(24px, 7vw, 32px)", lineHeight: 1.1, marginTop: 6 }}>
            Compradores
          </h1>
          <p style={{ color: INK_SOFT, fontSize: 14, marginTop: 4 }}>
            {compradores.length} {compradores.length === 1 ? "pessoa" : "pessoas"} · a partir do histórico de vendas
          </p>

          <button
            onClick={exportarCompradoresExcel}
            disabled={compradores.length === 0}
            style={{
              background: ACCENT,
              color: "#fff",
              fontFamily: MONO,
              fontSize: 13,
              opacity: compradores.length === 0 ? 0.4 : 1,
            }}
            className="rounded-sm px-4 py-2.5 mt-5 flex items-center justify-center gap-2 hover:opacity-90 transition w-full sm:w-auto"
          >
            <Download size={16} />
            Exportar para Excel
          </button>

          {compradores.length > 0 && (
            <div className="flex items-center gap-2 mt-5 border-b pb-3" style={{ borderColor: RULE }}>
              <Search size={16} style={{ color: INK_SOFT }} />
              <input
                placeholder="Buscar por nome ou telefone…"
                value={compradoresBusca}
                onChange={(e) => setCompradoresBusca(e.target.value)}
                style={{ fontFamily: SANS, background: "transparent" }}
                className="text-sm flex-1 outline-none py-1"
              />
            </div>
          )}
        </header>

        <main className="px-5 pb-16 max-w-5xl mx-auto">
          {compradores.length === 0 ? (
            <p style={{ color: INK_SOFT, fontSize: 13 }} className="py-6 text-center">
              Nenhuma venda registrada ainda. Assim que a primeira venda for feita, o comprador aparece aqui.
            </p>
          ) : compradoresFiltrados.length === 0 ? (
            <p style={{ color: INK_SOFT, fontSize: 13 }} className="py-6 text-center">
              Nenhuma pessoa encontrada.
            </p>
          ) : (
            <div className="grid sm:grid-cols-2 gap-4">
              {compradoresFiltrados.map((c) => (
                <button
                  key={`${c.nome}|${c.telefone}`}
                  onClick={() => setCompradorSelecionado(c)}
                  style={{ background: CARD, borderColor: RULE }}
                  className="border rounded-sm p-4 text-left hover:opacity-90 transition"
                >
                  <h3 style={{ fontFamily: SERIF, fontWeight: 600, fontSize: 17 }}>{c.nome}</h3>
                  <p style={{ color: INK_SOFT, fontSize: 13, marginTop: 2, display: "flex", alignItems: "center", gap: 4 }}>
                    <Phone size={12} /> {c.telefone || "sem telefone"}
                  </p>
                  <div className="flex items-center gap-3 mt-3" style={{ fontFamily: MONO, fontSize: 11.5, color: INK_SOFT }}>
                    <span>
                      <strong style={{ color: ACCENT, fontFamily: SERIF, fontSize: 15 }}>{c.compras}</strong>{" "}
                      {c.compras === 1 ? "compra" : "compras"}
                    </span>
                    <span>
                      <strong style={{ color: ACCENT, fontFamily: SERIF, fontSize: 15 }}>{c.exemplares}</strong>{" "}
                      {c.exemplares === 1 ? "exemplar" : "exemplares"}
                    </span>
                  </div>
                  <p style={{ fontFamily: MONO, fontSize: 10.5, color: INK_SOFT, marginTop: 6 }}>
                    Última compra: {formatDate(c.ultimaCompra)}
                  </p>
                  <p
                    style={{ fontFamily: MONO, fontSize: 11, color: ACCENT, borderColor: RULE, marginTop: 12, paddingTop: 12 }}
                    className="border-t"
                  >
                    Ver livros comprados →
                  </p>
                </button>
              ))}
            </div>
          )}
        </main>
        {menuLateral}
        {botaoVoltarAcervo}
        {sbErroBanner}

        {/* Detalhe completo de um comprador */}
        {compradorSelecionado && (
          <div className="fixed inset-0 bg-black/30 flex items-center justify-center px-5 z-20">
            <div style={{ background: CARD, borderColor: RULE }} className="border rounded-md p-5 w-full max-w-sm max-h-[80vh] flex flex-col">
              <div className="flex items-start justify-between gap-3">
                <div className="min-w-0">
                  <h3 style={{ fontFamily: SERIF, fontWeight: 600, fontSize: 19 }} className="truncate">
                    {compradorSelecionado.nome}
                  </h3>
                  <p style={{ color: INK_SOFT, fontSize: 13, marginTop: 2, display: "flex", alignItems: "center", gap: 4 }}>
                    <Phone size={12} /> {compradorSelecionado.telefone || "sem telefone"}
                  </p>
                </div>
                <button onClick={() => setCompradorSelecionado(null)} style={{ color: INK_SOFT }} className="shrink-0">
                  <X size={20} />
                </button>
              </div>

              <div className="flex items-center gap-3 mt-3" style={{ fontFamily: MONO, fontSize: 11.5, color: INK_SOFT }}>
                <span>
                  <strong style={{ color: ACCENT, fontFamily: SERIF, fontSize: 16 }}>{compradorSelecionado.compras}</strong>{" "}
                  {compradorSelecionado.compras === 1 ? "compra" : "compras"}
                </span>
                <span>
                  <strong style={{ color: ACCENT, fontFamily: SERIF, fontSize: 16 }}>{compradorSelecionado.exemplares}</strong>{" "}
                  {compradorSelecionado.exemplares === 1 ? "exemplar" : "exemplares"}
                </span>
              </div>

              <p style={{ fontFamily: MONO, fontSize: 10.5, color: INK_SOFT, letterSpacing: "0.06em" }} className="mt-4 mb-2">
                TÍTULOS COMPRADOS
              </p>
              <div className="overflow-y-auto flex-1 flex flex-col gap-2">
                {[...compradorSelecionado.itens]
                  .sort((a, b) => new Date(b.data) - new Date(a.data))
                  .map((item, idx) => (
                    <div key={idx} style={{ background: PAPER, borderColor: RULE }} className="border rounded-sm px-3 py-2">
                      <p style={{ fontFamily: SANS, fontSize: 13 }}>{item.titulo}</p>
                      <p style={{ fontFamily: MONO, fontSize: 10.5, color: INK_SOFT, marginTop: 2 }}>
                        {item.quantidade} {item.quantidade === 1 ? "exemplar" : "exemplares"} · {formatDate(item.data)}
                      </p>
                    </div>
                  ))}
              </div>
            </div>
          </div>
        )}
      </div>
    );
  }

  // ---- página: Vendas ----
  if (pagina === "vendas") {
    return (
      <div style={{ background: PAPER, minHeight: "100vh", color: INK, fontFamily: SANS }}>
        <style>{fontesCSS}</style>
        <header className="px-5 pt-8 pb-6 max-w-5xl mx-auto">
          <TopBar />

          <h1 style={{ fontFamily: SERIF, fontWeight: 600, fontSize: "clamp(24px, 7vw, 32px)", lineHeight: 1.1, marginTop: 6 }}>
            Vendas
          </h1>
          <p style={{ color: INK_SOFT, fontSize: 14, marginTop: 4 }}>
            {vendasRegistradas.length} {vendasRegistradas.length === 1 ? "venda registrada" : "vendas registradas"} ·{" "}
            {totalExemplaresVendidos} {totalExemplaresVendidos === 1 ? "exemplar vendido" : "exemplares vendidos"}
          </p>

          {vendasRegistradas.length > 0 && (
            <>
              <div className="flex flex-wrap gap-3 mt-5">
                <div>
                  <label style={{ fontFamily: MONO, fontSize: 10.5, color: INK_SOFT, letterSpacing: "0.06em" }}>DE</label>
                  <input
                    type="date"
                    value={vendasDataInicio}
                    onChange={(e) => setVendasDataInicio(e.target.value)}
                    style={{ fontFamily: MONO, fontSize: 12.5, borderColor: RULE, background: CARD, color: INK }}
                    className="border rounded-sm px-2 py-1.5 block mt-1"
                  />
                </div>
                <div>
                  <label style={{ fontFamily: MONO, fontSize: 10.5, color: INK_SOFT, letterSpacing: "0.06em" }}>ATÉ</label>
                  <input
                    type="date"
                    value={vendasDataFim}
                    onChange={(e) => setVendasDataFim(e.target.value)}
                    style={{ fontFamily: MONO, fontSize: 12.5, borderColor: RULE, background: CARD, color: INK }}
                    className="border rounded-sm px-2 py-1.5 block mt-1"
                  />
                </div>
                {(vendasDataInicio || vendasDataFim) && (
                  <button
                    onClick={() => {
                      setVendasDataInicio("");
                      setVendasDataFim("");
                    }}
                    style={{ fontFamily: MONO, fontSize: 11.5, color: INK_SOFT }}
                    className="self-end pb-2 hover:opacity-70 transition"
                  >
                    Limpar datas
                  </button>
                )}
                {vendedoresComVenda.length > 0 && (
                  <div>
                    <label style={{ fontFamily: MONO, fontSize: 10.5, color: INK_SOFT, letterSpacing: "0.06em" }}>VENDEDOR</label>
                    <select
                      value={vendasFiltroVendedor}
                      onChange={(e) => setVendasFiltroVendedor(e.target.value)}
                      style={{ fontFamily: MONO, fontSize: 12.5, borderColor: RULE, background: CARD, color: INK }}
                      className="border rounded-sm px-2 py-1.5 block mt-1"
                    >
                      <option value="todos">Todos</option>
                      {vendedoresComVenda.map((v) => (
                        <option key={v} value={v}>{v}</option>
                      ))}
                    </select>
                  </div>
                )}
              </div>

              <div
                style={{ background: `${ACCENT}12`, borderColor: RULE }}
                className="border rounded-sm px-4 py-3 mt-4 flex flex-wrap items-center gap-x-6 gap-y-1"
              >
                <span style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT, letterSpacing: "0.06em" }}>
                  NO PERÍODO FILTRADO
                </span>
                <span style={{ fontFamily: SANS, fontSize: 13 }}>
                  <strong style={{ fontFamily: SERIF, color: ACCENT, fontSize: 16 }}>{resumoPeriodo.vendas}</strong>{" "}
                  {resumoPeriodo.vendas === 1 ? "venda" : "vendas"}
                </span>
                <span style={{ fontFamily: SANS, fontSize: 13 }}>
                  <strong style={{ fontFamily: SERIF, color: ACCENT, fontSize: 16 }}>{resumoPeriodo.exemplares}</strong>{" "}
                  {resumoPeriodo.exemplares === 1 ? "exemplar" : "exemplares"}
                </span>
                {resumoPeriodo.temValor && (
                  <span style={{ fontFamily: SANS, fontSize: 13 }}>
                    Total:{" "}
                    <strong style={{ fontFamily: SERIF, color: ACCENT, fontSize: 16 }}>
                      {formatBRL(resumoPeriodo.total)}
                    </strong>
                  </span>
                )}
              </div>

              <div className="flex items-center gap-2 mt-4 border-b pb-3" style={{ borderColor: RULE }}>
                <Search size={16} style={{ color: INK_SOFT }} />
                <input
                  placeholder="Buscar por livro ou comprador…"
                  value={vendasBusca}
                  onChange={(e) => setVendasBusca(e.target.value)}
                  style={{ fontFamily: SANS, background: "transparent" }}
                  className="text-sm flex-1 outline-none py-1"
                />
              </div>
            </>
          )}
        </header>

        <main className="px-5 pb-16 max-w-5xl mx-auto">
          {vendasRegistradas.length === 0 ? (
            <p style={{ color: INK_SOFT, fontSize: 13 }} className="py-6 text-center">
              Nenhuma venda registrada ainda. Assim que vender um livro, ela aparece aqui.
            </p>
          ) : vendasFiltradas.length === 0 ? (
            <p style={{ color: INK_SOFT, fontSize: 13 }} className="py-6 text-center">
              Nenhuma venda encontrada.
            </p>
          ) : (
            <div className="grid sm:grid-cols-2 gap-3">
              {vendasFiltradas.map((v) => (
                <button
                  key={v.id}
                  onClick={() => setVendaSelecionada(v)}
                  style={{ background: CARD, borderColor: RULE }}
                  className="border rounded-sm p-4 text-left hover:opacity-90 transition flex items-start gap-3"
                >
                  <div
                    style={{ background: `${ACCENT}18`, color: ACCENT }}
                    className="rounded-sm w-11 h-11 flex flex-col items-center justify-center shrink-0"
                  >
                    <span style={{ fontFamily: SERIF, fontWeight: 700, fontSize: 16, lineHeight: 1 }}>{v.quantidade}</span>
                    <span style={{ fontFamily: MONO, fontSize: 8 }}>UN.</span>
                  </div>
                  <div className="min-w-0 flex-1">
                    <div className="flex items-start justify-between gap-2">
                      <p style={{ fontFamily: SANS, fontSize: 14, fontWeight: 600 }} className="truncate">
                        {v.livroTitulo}
                      </p>
                      {v.valorTotal != null && (
                        <span style={{ fontFamily: MONO, fontSize: 12.5, color: ACCENT }} className="shrink-0">
                          {formatBRL(v.valorTotal)}
                        </span>
                      )}
                    </div>
                    <p style={{ color: INK_SOFT, fontSize: 12.5, marginTop: 1 }} className="truncate">
                      Comprador: {v.compradorNome}
                    </p>
                    <p style={{ fontFamily: MONO, fontSize: 10.5, color: INK_SOFT, marginTop: 4 }}>
                      {formatDate(v.data)}
                    </p>
                  </div>
                </button>
              ))}
            </div>
          )}
        </main>
        {menuLateral}
        {botaoVoltarAcervo}
        {sbErroBanner}

        {/* Detalhe de uma venda */}
        {vendaSelecionada && (
          <div className="fixed inset-0 bg-black/30 flex items-center justify-center px-5 z-20">
            <div style={{ background: CARD, borderColor: RULE }} className="border rounded-md p-5 w-full max-w-sm">
              <div className="flex items-start justify-between gap-3 mb-3">
                <h3 style={{ fontFamily: SERIF, fontWeight: 600, fontSize: 19 }} className="min-w-0 truncate">
                  {vendaSelecionada.livroTitulo}
                </h3>
                <button onClick={() => setVendaSelecionada(null)} style={{ color: INK_SOFT }} className="shrink-0">
                  <X size={20} />
                </button>
              </div>
              <div className="flex flex-col gap-2" style={{ fontFamily: SANS, fontSize: 13.5 }}>
                <p>
                  <span style={{ color: INK_SOFT }}>Quantidade vendida: </span>
                  <strong>{vendaSelecionada.quantidade}</strong>
                </p>
                <p>
                  <span style={{ color: INK_SOFT }}>Comprador: </span>
                  {vendaSelecionada.compradorNome}
                </p>
                {vendaSelecionada.compradorTelefone && (
                  <p style={{ display: "flex", alignItems: "center", gap: 4 }}>
                    <Phone size={12} style={{ color: INK_SOFT }} /> {vendaSelecionada.compradorTelefone}
                  </p>
                )}
                {vendaSelecionada.vendedorNome && (
                  <p>
                    <span style={{ color: INK_SOFT }}>Vendido por: </span>
                    {vendaSelecionada.vendedorNome}
                  </p>
                )}
                {vendaSelecionada.desconto && (
                  <p>
                    <span style={{ color: INK_SOFT }}>Desconto: </span>
                    {formatBRL(vendaSelecionada.desconto)}
                  </p>
                )}
                {vendaSelecionada.valorTotal != null && (
                  <p>
                    <span style={{ color: INK_SOFT }}>Total: </span>
                    <strong style={{ color: ACCENT }}>{formatBRL(vendaSelecionada.valorTotal)}</strong>
                  </p>
                )}
                <p style={{ fontFamily: MONO, fontSize: 11.5, color: INK_SOFT, marginTop: 4 }}>
                  {formatDate(vendaSelecionada.data)}
                </p>
              </div>
            </div>
          </div>
        )}
      </div>
    );
  }

  // ---- página: Categorias ----
  if (pagina === "categorias") {
    const lista = categoriasCadastradas || [];
    return (
      <div style={{ background: PAPER, minHeight: "100vh", color: INK, fontFamily: SANS }}>
        <style>{fontesCSS}</style>
        <header className="px-5 pt-8 pb-6 max-w-5xl mx-auto">
          <TopBar />

          <h1 style={{ fontFamily: SERIF, fontWeight: 600, fontSize: "clamp(24px, 7vw, 32px)", lineHeight: 1.1, marginTop: 6 }}>
            Categorias
          </h1>
          <p style={{ color: INK_SOFT, fontSize: 14, marginTop: 4 }}>
            {lista.length} {lista.length === 1 ? "categoria cadastrada" : "categorias cadastradas"}
          </p>
        </header>

        <main className="px-5 pb-16 max-w-5xl mx-auto">
          <div className="max-w-md flex gap-2 mb-6">
            <input
              placeholder="Nova categoria…"
              value={novaCategoria}
              onChange={(e) => setNovaCategoria(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && adicionarCategoria()}
              style={{ fontFamily: SANS, borderColor: RULE, background: CARD }}
              className="border rounded-sm px-3 py-2 text-sm flex-1"
            />
            <button
              onClick={adicionarCategoria}
              disabled={!novaCategoria.trim()}
              style={{ background: ACCENT, color: "#fff", fontFamily: MONO, fontSize: 13, opacity: !novaCategoria.trim() ? 0.4 : 1 }}
              className="rounded-sm px-4 py-2 flex items-center gap-1.5 hover:opacity-90 transition shrink-0"
            >
              <Plus size={15} />
              Adicionar
            </button>
          </div>

          {lista.length === 0 ? (
            <p style={{ color: INK_SOFT, fontSize: 13 }} className="py-6 text-center">
              Nenhuma categoria cadastrada ainda.
            </p>
          ) : (
            <div className="flex flex-col gap-2 max-w-md">
              {lista.map((c) => {
                const emUso = contagemPorCategoria[c] || 0;
                return (
                  <div
                    key={c}
                    style={{ background: CARD, borderColor: RULE }}
                    className="border rounded-sm p-3 flex items-center gap-3"
                  >
                    <div className="flex-1 min-w-0">
                      <p style={{ fontFamily: SANS, fontSize: 13.5, fontWeight: 600 }} className="truncate">{c}</p>
                      <p style={{ color: INK_SOFT, fontSize: 12 }}>
                        {emUso} {emUso === 1 ? "livro" : "livros"} nessa categoria
                      </p>
                    </div>
                    {confirmExcluirCategoria === c ? (
                      emUso > 0 ? (
                        <div className="flex items-center gap-2">
                          <span style={{ fontFamily: MONO, fontSize: 11, color: "#B04A4A" }} className="max-w-[140px]">
                            Em uso — mude a categoria desses livros antes.
                          </span>
                          <button
                            onClick={() => setConfirmExcluirCategoria(null)}
                            style={{ color: INK_SOFT, fontFamily: MONO, fontSize: 11 }}
                            className="px-2 py-1.5 shrink-0"
                          >
                            Ok
                          </button>
                        </div>
                      ) : (
                        <div className="flex gap-1 shrink-0">
                          <button
                            onClick={() => excluirCategoria(c)}
                            style={{ background: "#B04A4A", color: "#fff", fontFamily: MONO, fontSize: 11 }}
                            className="rounded-sm px-2 py-1.5"
                          >
                            Excluir
                          </button>
                          <button
                            onClick={() => setConfirmExcluirCategoria(null)}
                            style={{ color: INK_SOFT, fontFamily: MONO, fontSize: 11 }}
                            className="px-2 py-1.5"
                          >
                            Cancelar
                          </button>
                        </div>
                      )
                    ) : (
                      <button
                        onClick={() => setConfirmExcluirCategoria(c)}
                        style={{ color: "#B04A4A" }}
                        className="p-1.5 shrink-0 hover:opacity-70 transition"
                        aria-label={`Excluir ${c}`}
                      >
                        <Trash2 size={16} />
                      </button>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </main>
        {menuLateral}
        {botaoVoltarAcervo}
        {sbErroBanner}
      </div>
    );
  }

  // ---- página: Gestão da equipe ----
  if (pagina === "equipe") {
    return (
      <div style={{ background: PAPER, minHeight: "100vh", color: INK, fontFamily: SANS }}>
        <style>{fontesCSS}</style>
        <header className="px-5 pt-8 pb-6 max-w-5xl mx-auto">
          <TopBar />

          <h1 style={{ fontFamily: SERIF, fontWeight: 600, fontSize: "clamp(24px, 7vw, 32px)", lineHeight: 1.1, marginTop: 6 }}>
            Equipe
          </h1>
          <p style={{ color: INK_SOFT, fontSize: 14, marginTop: 4 }}>
            {(usuarios || []).length} {(usuarios || []).length === 1 ? "pessoa cadastrada" : "pessoas cadastradas"} ·{" "}
            {vendedores.length} {vendedores.length === 1 ? "liderado" : "liderados"}
          </p>

          {/* Abas */}
          <div className="flex gap-2 mt-5">
            {[
              { key: "equipe", label: "Equipe" },
              { key: "cadastrar", label: "Cadastrar novo" },
            ].map((aba) => (
              <button
                key={aba.key}
                onClick={() => setEquipeAba(aba.key)}
                style={{
                  fontFamily: MONO,
                  fontSize: 13,
                  background: equipeAba === aba.key ? ACCENT : CARD,
                  color: equipeAba === aba.key ? "#fff" : INK_SOFT,
                  borderColor: equipeAba === aba.key ? ACCENT : RULE,
                }}
                className="border rounded-sm px-4 py-2.5 transition hover:opacity-90"
              >
                {aba.label}
              </button>
            ))}
          </div>

          {equipeAba === "equipe" && (
            <div className="flex items-center gap-2 mt-4 border-b pb-3" style={{ borderColor: RULE }}>
              <Search size={16} style={{ color: INK_SOFT }} />
              <input
                placeholder="Buscar por nome…"
                value={equipeBusca}
                onChange={(e) => setEquipeBusca(e.target.value)}
                style={{ fontFamily: SANS, background: "transparent" }}
                className="text-sm flex-1 outline-none py-1"
              />
            </div>
          )}
        </header>

        <main className="px-5 pb-16 max-w-5xl mx-auto">
          {equipeAba === "equipe" && (
            <div className="flex flex-col gap-2">
              {equipeFiltrada.length === 0 && (
                <p style={{ color: INK_SOFT, fontSize: 13 }} className="py-6 text-center">
                  {usuarios && usuarios.length === 0
                    ? 'Ninguém cadastrado ainda. Use a aba "Cadastrar novo".'
                    : "Nenhuma pessoa encontrada."}
                </p>
              )}
              {equipeFiltrada.map((u) => (
                <div
                  key={u.id}
                  style={{ background: CARD, borderColor: RULE }}
                  className="border rounded-sm p-3 flex items-center gap-3 flex-wrap sm:flex-nowrap"
                >
                  <div className="flex-1 min-w-[140px]">
                    <p style={{ fontFamily: SANS, fontSize: 13.5, fontWeight: 600 }} className="truncate">
                      {u.nome}
                    </p>
                    <p style={{ color: INK_SOFT, fontSize: 12 }} className="truncate">
                      {u.email || "sem e-mail"} {u.telefone ? `· ${u.telefone}` : ""}
                    </p>
                  </div>
                  <select
                    value={u.papel}
                    disabled={!souLider}
                    onChange={(e) => editarPapelUsuario(u.id, e.target.value)}
                    style={{
                      borderColor: RULE,
                      background: PAPER,
                      fontFamily: MONO,
                      fontSize: 12,
                      opacity: souLider ? 1 : 0.6,
                    }}
                    className="border rounded-sm px-2 py-1.5"
                  >
                    {PAPEIS.map((p) => (
                      <option key={p} value={p}>{p}</option>
                    ))}
                  </select>
                  {souLider && (
                    <button
                      onClick={() => abrirRedefinirSenha(u.id)}
                      style={{ borderColor: RULE, color: INK_SOFT, fontFamily: MONO, fontSize: 11 }}
                      className="border rounded-full px-2.5 py-1.5 shrink-0 hover:opacity-80 transition"
                    >
                      Redefinir senha
                    </button>
                  )}
                  {souLider &&
                    (confirmExcluirUsuario === u.id ? (
                      <div className="flex gap-1 shrink-0">
                        <button
                          onClick={() => excluirUsuario(u.id)}
                          style={{ background: "#B04A4A", color: "#fff", fontFamily: MONO, fontSize: 11 }}
                          className="rounded-sm px-2 py-1.5"
                        >
                          Excluir
                        </button>
                        <button
                          onClick={() => setConfirmExcluirUsuario(null)}
                          style={{ color: INK_SOFT, fontFamily: MONO, fontSize: 11 }}
                          className="px-2 py-1.5"
                        >
                          Cancelar
                        </button>
                      </div>
                    ) : (
                      <button
                        onClick={() => setConfirmExcluirUsuario(u.id)}
                        style={{ color: "#B04A4A" }}
                        className="p-1.5 shrink-0 hover:opacity-70 transition"
                        aria-label={`Excluir ${u.nome}`}
                      >
                        <Trash2 size={16} />
                      </button>
                    ))}
                </div>
              ))}
            </div>
          )}

          {equipeAba === "cadastrar" && (
            <div className="max-w-md">
              {!souLider ? (
                <div style={{ background: CARD, borderColor: RULE }} className="border rounded-sm p-4">
                  <p style={{ fontFamily: SANS, fontSize: 13.5, color: INK_SOFT }}>
                    Só quem tem o papel <strong style={{ color: INK }}>Líder</strong> pode cadastrar novas pessoas na
                    equipe. Se você acha que deveria ser Líder, peça pra alguém que já é te promover aqui mesmo,
                    na lista da equipe.
                  </p>
                </div>
              ) : (
                <div className="flex flex-col gap-3">
                  <input
                    placeholder="Nome *"
                    value={novoUsuario.nome}
                    onChange={(e) => setNovoUsuario({ ...novoUsuario, nome: e.target.value })}
                    style={{ fontFamily: SANS, borderColor: RULE, background: CARD }}
                    className="border rounded-sm px-3 py-2 text-sm"
                  />
                  <input
                    type="email"
                    placeholder="E-mail *"
                    value={novoUsuario.email}
                    onChange={(e) => setNovoUsuario({ ...novoUsuario, email: e.target.value })}
                    style={{ fontFamily: SANS, borderColor: RULE, background: CARD }}
                    className="border rounded-sm px-3 py-2 text-sm"
                  />
                  <input
                    type="tel"
                    placeholder="Telefone"
                    value={novoUsuario.telefone}
                    onChange={(e) => setNovoUsuario({ ...novoUsuario, telefone: maskTelefone(e.target.value) })}
                    style={{ fontFamily: SANS, borderColor: RULE, background: CARD }}
                    className="border rounded-sm px-3 py-2 text-sm"
                  />
                  <select
                    value={novoUsuario.papel}
                    onChange={(e) => setNovoUsuario({ ...novoUsuario, papel: e.target.value })}
                    style={{ fontFamily: SANS, borderColor: RULE, background: CARD }}
                    className="border rounded-sm px-3 py-2 text-sm"
                  >
                    {PAPEIS.map((p) => (
                      <option key={p} value={p}>{p}</option>
                    ))}
                  </select>
                  <div className="relative">
                    <input
                      type={verSenhaNovoUsuario ? "text" : "password"}
                      placeholder="Senha inicial * (mín. 6 caracteres)"
                      value={novoUsuario.senha}
                      onChange={(e) => setNovoUsuario({ ...novoUsuario, senha: e.target.value })}
                      onKeyDown={(e) => e.key === "Enter" && adicionarUsuario()}
                      style={{ fontFamily: SANS, borderColor: RULE, background: CARD }}
                      className="border rounded-sm pl-3 pr-10 py-2 text-sm w-full"
                    />
                    <button
                      type="button"
                      onClick={() => setVerSenhaNovoUsuario((v) => !v)}
                      style={{ color: INK_SOFT }}
                      className="absolute right-3 top-1/2 -translate-y-1/2"
                      aria-label={verSenhaNovoUsuario ? "Esconder senha" : "Mostrar senha"}
                    >
                      {verSenhaNovoUsuario ? <EyeOff size={16} /> : <Eye size={16} />}
                    </button>
                  </div>
                  <p style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT }}>
                    Combine essa senha com a pessoa — ela pode trocar depois em "Meu perfil".
                  </p>
                  <button
                    onClick={adicionarUsuario}
                    disabled={
                      !novoUsuario.nome.trim() ||
                      !novoUsuario.email.trim() ||
                      !novoUsuario.senha ||
                      novoUsuario.senha.length < 6
                    }
                    style={{
                      background: ACCENT,
                      color: "#fff",
                      fontFamily: MONO,
                      fontSize: 13,
                      opacity:
                        !novoUsuario.nome.trim() || !novoUsuario.email.trim() || !novoUsuario.senha || novoUsuario.senha.length < 6
                          ? 0.4
                          : 1,
                    }}
                    className="rounded-sm px-3 py-2.5 hover:opacity-90 transition flex items-center justify-center gap-2"
                  >
                    <UserPlus size={16} />
                    Cadastrar pessoa
                  </button>
                </div>
              )}
            </div>
          )}
        </main>

        {/* Modal de redefinir senha (Líder redefinindo a senha de outra pessoa) */}
        {redefinirSenhaId && (
          <div className="fixed inset-0 bg-black/30 flex items-center justify-center px-5 z-20">
            <div style={{ background: CARD, borderColor: RULE }} className="border rounded-md p-5 w-full max-w-sm">
              <h3 style={{ fontFamily: SERIF, fontSize: 18, fontWeight: 600 }} className="mb-1">Redefinir senha</h3>
              <p style={{ color: INK_SOFT, fontSize: 13 }} className="mb-3">
                {(usuarios || []).find((u) => u.id === redefinirSenhaId)?.nome}
              </p>
              <div className="relative mb-4">
                <input
                  type={verSenhaRedefinir ? "text" : "password"}
                  placeholder="Nova senha (mín. 6 caracteres)"
                  value={redefinirSenhaValor}
                  onChange={(e) => setRedefinirSenhaValor(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && confirmarRedefinirSenha()}
                  style={{ borderColor: RULE, background: PAPER, fontFamily: SANS }}
                  className="border rounded-sm pl-3 pr-10 py-2 text-sm w-full"
                />
                <button
                  type="button"
                  onClick={() => setVerSenhaRedefinir((v) => !v)}
                  style={{ color: INK_SOFT }}
                  className="absolute right-3 top-1/2 -translate-y-1/2"
                  aria-label={verSenhaRedefinir ? "Esconder senha" : "Mostrar senha"}
                >
                  {verSenhaRedefinir ? <EyeOff size={16} /> : <Eye size={16} />}
                </button>
              </div>
              <div className="flex gap-2 justify-end">
                <button onClick={() => setRedefinirSenhaId(null)} style={{ fontFamily: MONO, fontSize: 12.5, color: INK_SOFT }} className="px-3 py-2">
                  Cancelar
                </button>
                <button
                  onClick={confirmarRedefinirSenha}
                  disabled={!redefinirSenhaValor || redefinirSenhaValor.length < 6}
                  style={{
                    background: ACCENT,
                    color: "#fff",
                    fontFamily: MONO,
                    fontSize: 12.5,
                    opacity: !redefinirSenhaValor || redefinirSenhaValor.length < 6 ? 0.4 : 1,
                  }}
                  className="rounded-sm px-3 py-2"
                >
                  Confirmar
                </button>
              </div>
            </div>
          </div>
        )}

        {menuLateral}
        {botaoVoltarAcervo}
        {sbErroBanner}
      </div>
    );
  }

  // ---- página: Gestão do perfil ----
  if (pagina === "perfil") {
    return (
      <div style={{ background: PAPER, minHeight: "100vh", color: INK, fontFamily: SANS }}>
        <style>{fontesCSS}</style>
        <header className="px-5 pt-8 pb-6 max-w-5xl mx-auto">
          <TopBar />

          <h1 style={{ fontFamily: SERIF, fontWeight: 600, fontSize: "clamp(24px, 7vw, 32px)", lineHeight: 1.1, marginTop: 6 }}>
            Meu perfil
          </h1>
          <p style={{ color: INK_SOFT, fontSize: 14, marginTop: 4 }}>Seus dados como pessoa da equipe.</p>
        </header>

        <main className="px-5 pb-16 max-w-5xl mx-auto">
          <div style={{ background: CARD, borderColor: RULE }} className="border rounded-sm p-4 max-w-md mb-5">
            <p style={{ fontFamily: SERIF, fontWeight: 600, fontSize: 17 }}>{perfil?.nome}</p>
            <p style={{ color: INK_SOFT, fontSize: 13, marginTop: 2 }}>{perfil?.email}</p>
            {perfil?.telefone && <p style={{ color: INK_SOFT, fontSize: 13 }}>{perfil.telefone}</p>}
            <span
              style={{ background: `${ACCENT}18`, color: ACCENT, fontFamily: MONO, fontSize: 11 }}
              className="inline-block rounded-full px-2.5 py-1 mt-2"
            >
              {perfil?.papel}
            </span>
            <p style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT }} className="mt-3">
              Nome, e-mail e papel são definidos pelo Líder na Gestão da equipe — fale com ele(a) pra corrigir algo.
            </p>
          </div>

          <div className="max-w-md">
            <label style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT }}>Trocar minha senha</label>
            <div className="relative mt-1 mb-3">
              <input
                type={verSenhaPropria ? "text" : "password"}
                placeholder="Nova senha (mín. 6 caracteres)"
                value={novaSenhaPropria}
                onChange={(e) => setNovaSenhaPropria(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && trocarMinhaSenha()}
                style={{ fontFamily: SANS, borderColor: RULE, background: CARD }}
                className="border rounded-sm pl-3 pr-10 py-2 text-sm w-full"
              />
              <button
                type="button"
                onClick={() => setVerSenhaPropria((v) => !v)}
                style={{ color: INK_SOFT }}
                className="absolute right-3 top-1/2 -translate-y-1/2"
                aria-label={verSenhaPropria ? "Esconder senha" : "Mostrar senha"}
              >
                {verSenhaPropria ? <EyeOff size={16} /> : <Eye size={16} />}
              </button>
            </div>
            <button
              onClick={trocarMinhaSenha}
              disabled={!novaSenhaPropria || novaSenhaPropria.length < 6}
              style={{
                background: ACCENT,
                color: "#fff",
                fontFamily: MONO,
                fontSize: 13,
                opacity: !novaSenhaPropria || novaSenhaPropria.length < 6 ? 0.4 : 1,
              }}
              className="rounded-sm px-3 py-2.5 hover:opacity-90 transition"
            >
              Trocar senha
            </button>
            {senhaAlterada && (
              <p style={{ fontFamily: MONO, fontSize: 12, color: STATUS.disponivel.color }} className="mt-2">
                Senha alterada ✓
              </p>
            )}
          </div>

          <button
            onClick={encerrarSessao}
            style={{ fontFamily: MONO, fontSize: 13, color: "#B04A4A", borderColor: "#B04A4A" }}
            className="border rounded-sm px-4 py-2.5 mt-8 hover:opacity-80 transition flex items-center gap-2"
          >
            <ArrowLeft size={15} />
            Sair da conta
          </button>
        </main>
        {menuLateral}
        {botaoVoltarAcervo}
        {sbErroBanner}
      </div>
    );
  }

  return (
    <div style={{ background: PAPER, minHeight: "100vh", color: INK, fontFamily: SANS }}>
      <style>{fontesCSS}</style>

      {/* Cabeçalho */}
      <header className="px-5 pt-8 pb-6 max-w-5xl mx-auto">
        <TopBar />
        <div className="flex items-start justify-between gap-4 flex-wrap">
          <div className="min-w-0">
            <h1 style={{ fontFamily: SERIF, fontWeight: 600, fontSize: "clamp(24px, 7vw, 32px)", lineHeight: 1.1, marginTop: 6 }}>
              Acervo da biblioteca
            </h1>
            <p style={{ color: INK_SOFT, fontSize: 14, marginTop: 4 }}>
              {totais.titulos} {totais.titulos === 1 ? "título catalogado" : "títulos catalogados"}
              {saving && <span style={{ opacity: 0.5 }}> · salvando…</span>}
            </p>
          </div>
        </div>

        <button
          onClick={() => {
            setPagina("gestao");
            setGestaoAba("catalogo");
            abrirScanner();
          }}
          style={{ background: ACCENT, color: "#fff", fontFamily: MONO, fontSize: 13.5 }}
          className="w-full rounded-sm py-3 mt-5 flex items-center justify-center gap-2 hover:opacity-90 transition"
        >
          <ScanBarcode size={18} />
          Ler código de barras
        </button>

        {/* Estoque total em destaque */}
        <div className="stamp inline-flex items-baseline gap-2 mt-5 px-1 py-1 w-full" style={{ color: ACCENT }}>
          <span style={{ fontFamily: SERIF, fontWeight: 700, fontSize: 24 }}>{totais.exemplares}</span>
          <span style={{ fontFamily: MONO, fontSize: 11.5, letterSpacing: "0.08em" }}>
            {totais.exemplares === 1 ? "EXEMPLAR EM ESTOQUE" : "EXEMPLARES EM ESTOQUE"}
          </span>
        </div>

        {/* Filtro por categoria */}
        {categorias.length > 0 && (
          <div className="flex items-center gap-2 mt-3">
            <span style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT, letterSpacing: "0.06em" }}>
              CATEGORIA
            </span>
            <select
              value={categoriaFiltro}
              onChange={(e) => setCategoriaFiltro(e.target.value)}
              style={{ fontFamily: MONO, fontSize: 12.5, borderColor: RULE, background: CARD, color: INK }}
              className="border rounded-full px-3 py-1.5"
            >
              <option value="todas">Todas</option>
              {categorias.map((c) => (
                <option key={c} value={c}>
                  {c}
                </option>
              ))}
            </select>
          </div>
        )}

        {/* Busca */}
        <div className="flex items-center gap-2 mt-4 border-b pb-3" style={{ borderColor: RULE }}>
          <Search size={16} style={{ color: INK_SOFT }} />
          <input
            placeholder="Buscar por título ou autor…"
            value={busca}
            onChange={(e) => setBusca(e.target.value)}
            style={{ fontFamily: SANS, background: "transparent" }}
            className="text-sm flex-1 outline-none py-1"
          />
        </div>
      </header>

      {/* Corpo: fichas ou estante + movimentações */}
      <main className="px-5 pb-16 max-w-5xl mx-auto grid lg:grid-cols-[1fr_280px] gap-8">
        <section className="min-w-0">
          <div className="flex items-center justify-between gap-3 mb-4 flex-wrap">
            <p style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT, letterSpacing: "0.08em" }}>
              {filtrados.length} {filtrados.length === 1 ? "TÍTULO" : "TÍTULOS"} NESTA VISÃO
            </p>
            <div className="flex gap-2">
              {[
                { key: "fichas", label: "Fichas" },
                { key: "estante", label: "Estante virtual" },
              ].map((v) => (
                <button
                  key={v.key}
                  onClick={() => setVisualizacao(v.key)}
                  style={{
                    fontFamily: MONO,
                    fontSize: 12,
                    borderColor: visualizacao === v.key ? ACCENT : RULE,
                    color: visualizacao === v.key ? ACCENT : INK_SOFT,
                    background: visualizacao === v.key ? "#FFFFFFAA" : "transparent",
                  }}
                  className="border rounded-full px-3 py-1.5 transition"
                >
                  {v.label}
                </button>
              ))}
            </div>
          </div>

          {filtrados.length === 0 ? (
            <p style={{ color: INK_SOFT, fontFamily: SANS }} className="text-sm py-10 text-center">
              Nenhuma ficha encontrada. Ajuste a busca ou use "Gestão de estoque" para catalogar um livro.
            </p>
          ) : visualizacao === "fichas" ? (
            <div className="grid sm:grid-cols-2 gap-4">
              {filtrados.map((livro) => (
                <FichaLivro
                  key={livro.id}
                  livro={livro}
                  onVender={() => abrirVendaModal(livro.id)}
                  onRepor={() => abrirRepoModal(livro.id)}
                />
              ))}
            </div>
          ) : (
            <EstanteVirtual livros={filtrados} onSelecionar={(id) => setSpineSelecionado(id)} />
          )}
        </section>

        <aside>
          <h2 style={{ fontFamily: MONO, letterSpacing: "0.1em", fontSize: 11, color: INK_SOFT }} className="mb-3">
            ÚLTIMAS MOVIMENTAÇÕES
          </h2>
          {movsRecentes.length === 0 ? (
            <p style={{ color: INK_SOFT }} className="text-xs">Nenhum registro ainda.</p>
          ) : (
            <ol className="relative border-l pl-4" style={{ borderColor: RULE }}>
              {movsRecentes.map((m) => {
                const livro = livroPorId(m.livroId);
                return (
                  <li key={m.id} className="mb-4 relative">
                    <span className="absolute -left-[21px] top-1 w-2 h-2 rounded-full" style={{ background: ACCENT }} />
                    <p style={{ fontSize: 13, fontFamily: SANS }}>
                      <strong>{livro ? livro.titulo : "Livro removido"}</strong> — {m.tipo}
                      {m.detalhe ? ` · ${m.detalhe}` : ""}
                    </p>
                    <p style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT }}>{formatDate(m.data)}</p>
                  </li>
                );
              })}
            </ol>
          )}
        </aside>
      </main>

      {/* Modal de venda */}
      {vendaModalOverlay}

      {/* Modal de reposição rápida */}
      {repoModal && (
        <div className="fixed inset-0 bg-black/30 flex items-center justify-center px-5 z-20">
          <div style={{ background: CARD, borderColor: RULE }} className="border rounded-md p-5 w-full max-w-sm">
            <h3 style={{ fontFamily: SERIF, fontSize: 18, fontWeight: 600 }} className="mb-1">Repor estoque</h3>
            <p style={{ color: INK_SOFT, fontSize: 13 }} className="mb-3">{livroPorId(repoModal)?.titulo}</p>
            <label style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT }}>Exemplares recebidos</label>
            <div className="mt-1 mb-4">
              <QtyStepper value={rQtd} onChange={setRQtd} min={1} color={STATUS.disponivel.color} />
            </div>
            <div className="flex gap-2 justify-end">
              <button onClick={() => setRepoModal(null)} style={{ fontFamily: MONO, fontSize: 12.5, color: INK_SOFT }} className="px-3 py-2">
                Cancelar
              </button>
              <button onClick={confirmarRepo} style={{ background: STATUS.disponivel.color, color: "#fff", fontFamily: MONO, fontSize: 12.5 }} className="rounded-sm px-3 py-2">
                Confirmar
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Detalhe rápido de um livro selecionado na estante virtual */}
      {spineSelecionado && livroPorId(spineSelecionado) && (
        <div className="fixed inset-0 bg-black/30 flex items-center justify-center px-5 z-20">
          <div style={{ background: CARD, borderColor: RULE }} className="border rounded-md p-5 w-full max-w-sm">
            {(() => {
              const livro = livroPorId(spineSelecionado);
              const sv = statusVisual(livro);
              return (
                <>
                  <p style={{ fontFamily: MONO, fontSize: 10.5, color: INK_SOFT, letterSpacing: "0.06em" }}>
                    {livro.categoria.toUpperCase()}
                  </p>
                  <h3 style={{ fontFamily: SERIF, fontWeight: 600, fontSize: 20 }} className="mt-1">{livro.titulo}</h3>
                  <p style={{ fontStyle: "italic", color: INK_SOFT, fontSize: 13.5, marginTop: 2 }}>{livro.autor}</p>
                  <div className="flex items-center gap-2 mt-3 flex-wrap">
                    <div
                      style={{ background: `${sv.cor}22`, color: sv.cor, fontFamily: MONO, fontSize: 9.5, letterSpacing: "0.06em" }}
                      className="rounded-full px-2 py-0.5 flex items-center gap-1"
                    >
                      <span className="w-1 h-1 rounded-full shrink-0" style={{ background: sv.dot }} />
                      {sv.texto.toUpperCase()}
                    </div>
                    <span style={{ fontFamily: MONO, fontSize: 12, color: INK_SOFT }}>
                      {livro.quantidade} {livro.quantidade === 1 ? "exemplar" : "exemplares"}
                    </span>
                    {formatBRL(livro.preco) && (
                      <span style={{ fontFamily: MONO, fontSize: 12, color: ACCENT, marginLeft: "auto" }}>
                        {formatBRL(livro.preco)}
                      </span>
                    )}
                    {estoqueBaixo(livro) && (
                      <span
                        style={{ background: `${COR_ESTOQUE_BAIXO}22`, color: COR_ESTOQUE_BAIXO, fontFamily: MONO, fontSize: 9.5, letterSpacing: "0.06em" }}
                        className="rounded-full px-2 py-0.5 flex items-center gap-1"
                      >
                        <span className="w-1 h-1 rounded-full shrink-0" style={{ background: COR_ESTOQUE_BAIXO }} />
                        ACABANDO
                      </span>
                    )}
                  </div>
                  <div className="flex gap-2 flex-wrap mt-4">
                    <AcaoBtn
                      onClick={() => {
                        setSpineSelecionado(null);
                        abrirVendaModal(livro.id);
                      }}
                      icon={<Tag size={13} />}
                      label="Vender"
                      color={STATUS.disponivel.color}
                      disabled={livro.quantidade <= 0}
                    />
                    <AcaoBtn
                      onClick={() => {
                        setSpineSelecionado(null);
                        abrirRepoModal(livro.id);
                      }}
                      icon={<PackagePlus size={13} />}
                      label="Repor estoque"
                      color={ACCENT}
                    />
                  </div>
                  <button
                    onClick={() => setSpineSelecionado(null)}
                    style={{ fontFamily: MONO, fontSize: 12, color: INK_SOFT }}
                    className="mt-4"
                  >
                    Fechar
                  </button>
                </>
              );
            })()}
          </div>
        </div>
      )}
      {menuLateral}
      {botaoVoltarAcervo}
      {sbErroBanner}
    </div>
  );
}

function hashStr(str) {
  let h = 0;
  for (let i = 0; i < str.length; i++) h = (h * 31 + str.charCodeAt(i)) >>> 0;
  return h;
}

const PALETA_LOMBADA = ["#790E0C", "#3F6B4F", "#2E5266", "#B8752E", "#6B4A3A", "#8A6D3B", "#4B5A6B", "#9C4A4A", "#5C6B3E", "#5A4A73"];

function corLombada(categoria) {
  const h = hashStr(categoria || "geral");
  return PALETA_LOMBADA[h % PALETA_LOMBADA.length];
}

function EstanteVirtual({ livros, onSelecionar }) {
  const grupos = useMemo(() => {
    const map = new Map();
    livros.forEach((l) => {
      const chave = l.categoria || "Sem categoria";
      if (!map.has(chave)) map.set(chave, []);
      map.get(chave).push(l);
    });
    return Array.from(map.entries()).sort((a, b) => a[0].localeCompare(b[0], "pt-BR"));
  }, [livros]);

  return (
    <div className="flex flex-col gap-8">
      {grupos.map(([categoria, itens]) => (
        <div key={categoria}>
          <p style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT, letterSpacing: "0.08em" }} className="mb-2">
            {categoria.toUpperCase()} <span style={{ opacity: 0.6 }}>· {itens.length}</span>
          </p>
          <div
            className="flex items-end gap-[3px] overflow-x-auto pb-1"
            style={{ borderBottom: `10px solid #8A6A4A`, boxShadow: "inset 0 -3px 4px rgba(0,0,0,0.18)" }}
          >
            {itens.map((livro) => {
              const h = hashStr(livro.id);
              const altura = 148 + (h % 38); // 148–186px
              const largura = 30 + (h % 13); // 30–42px
              const cor = corLombada(livro.categoria);
              const esgotado = livro.status === "esgotado";
              return (
                <button
                  key={livro.id}
                  onClick={() => onSelecionar(livro.id)}
                  title={`${livro.titulo} — ${livro.autor}`}
                  style={{
                    background: esgotado ? `${cor}66` : cor,
                    width: `clamp(24px, 7vw, ${largura}px)`,
                    height: altura,
                    borderTop: "3px solid rgba(255,255,255,0.25)",
                  }}
                  className="shrink-0 rounded-t-[2px] relative flex flex-col items-center justify-between py-2 hover:-translate-y-1 transition-transform"
                >
                  {esgotado && (
                    <span
                      className="absolute -top-1.5 right-0.5 w-2 h-2 rounded-full"
                      style={{ background: "#B04A4A", boxShadow: "0 0 0 1.5px #FBF8F1" }}
                    />
                  )}
                  <span
                    style={{
                      writingMode: "vertical-rl",
                      fontFamily: SERIF,
                      fontWeight: 600,
                      fontSize: 11.5,
                      color: "#FBF8F1",
                      letterSpacing: "0.01em",
                      maxHeight: altura - 34,
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                      whiteSpace: "nowrap",
                    }}
                  >
                    {livro.titulo}
                  </span>
                  <span
                    style={{
                      fontFamily: MONO,
                      fontSize: 9,
                      color: "#FBF8F1CC",
                      background: "rgba(0,0,0,0.22)",
                      borderRadius: 999,
                      padding: "1px 5px",
                    }}
                  >
                    {livro.quantidade}
                  </span>
                </button>
              );
            })}
          </div>
        </div>
      ))}
    </div>
  );
}

function QtyStepper({ value, onChange, min = 0, max, color = INK }) {
  const num = parseInt(value, 10) || 0;
  const dec = () => onChange(String(Math.max(min, num - 1)));
  const inc = () => onChange(String(max !== undefined ? Math.min(max, num + 1) : num + 1));
  return (
    <div className="flex items-center border rounded-sm overflow-hidden" style={{ borderColor: RULE }}>
      <button
        type="button"
        onClick={dec}
        disabled={num <= min}
        style={{ fontFamily: MONO, fontSize: 16, color, opacity: num <= min ? 0.3 : 1, background: CARD }}
        className="w-9 h-9 flex items-center justify-center shrink-0 hover:opacity-70 transition"
        aria-label="Diminuir"
      >
        −
      </button>
      <input
        type="number"
        value={value}
        min={min}
        max={max}
        onChange={(e) => onChange(e.target.value)}
        style={{ fontFamily: MONO, background: PAPER, color: INK, width: 48 }}
        className="text-sm text-center h-9 outline-none border-x"
      />
      <button
        type="button"
        onClick={inc}
        disabled={max !== undefined && num >= max}
        style={{ fontFamily: MONO, fontSize: 16, color, opacity: max !== undefined && num >= max ? 0.3 : 1, background: CARD }}
        className="w-9 h-9 flex items-center justify-center shrink-0 hover:opacity-70 transition"
        aria-label="Aumentar"
      >
        +
      </button>
    </div>
  );
}

function FichaLivro({ livro, onVender, onRepor }) {
  const sv = statusVisual(livro);
  return (
    <div style={{ background: CARD, borderColor: RULE }} className="card-hole border rounded-sm p-4 relative overflow-hidden flex flex-col">
      <div className="absolute top-3 right-3 flex flex-col items-end gap-1">
        <div
          style={{ background: `${sv.cor}22`, color: sv.cor, fontFamily: MONO, fontSize: 9.5, letterSpacing: "0.06em" }}
          className="rounded-full px-2 py-0.5 flex items-center gap-1"
        >
          <span className="w-1 h-1 rounded-full shrink-0" style={{ background: sv.dot }} />
          {sv.texto.toUpperCase()}
        </div>
        {estoqueBaixo(livro) && (
          <span
            style={{ background: `${COR_ESTOQUE_BAIXO}22`, color: COR_ESTOQUE_BAIXO, fontFamily: MONO, fontSize: 9.5, letterSpacing: "0.06em" }}
            className="rounded-full px-2 py-0.5 flex items-center gap-1"
          >
            <span className="w-1 h-1 rounded-full shrink-0" style={{ background: COR_ESTOQUE_BAIXO }} />
            ACABANDO
          </span>
        )}
      </div>

      <p style={{ fontFamily: MONO, fontSize: 10.5, color: INK_SOFT, letterSpacing: "0.06em" }}>
        {livro.categoria.toUpperCase()}
      </p>
      <h3 style={{ fontFamily: SERIF, fontWeight: 600, fontSize: 19, lineHeight: 1.25, marginTop: 3, paddingRight: 70 }}>
        {livro.titulo}
      </h3>
      <p style={{ fontStyle: "italic", color: INK_SOFT, fontSize: 13.5, marginTop: 2 }}>{livro.autor}</p>

      <div style={{ marginTop: 10, display: "flex", alignItems: "baseline", gap: 6, flexWrap: "wrap" }}>
        <span style={{ fontFamily: SERIF, fontWeight: 700, fontSize: 20, color: sv.cor }}>{livro.quantidade}</span>
        <span style={{ fontFamily: MONO, fontSize: 11, color: INK_SOFT }}>
          {livro.quantidade === 1 ? "exemplar em estoque" : "exemplares em estoque"}
        </span>
        {formatBRL(livro.preco) && (
          <span style={{ fontFamily: MONO, fontSize: 12, color: ACCENT, marginLeft: "auto" }}>
            {formatBRL(livro.preco)}
          </span>
        )}
      </div>
      
      <div className="mt-4 pt-3 flex gap-2 flex-wrap border-t" style={{ borderColor: RULE }}>
        <AcaoBtn onClick={onVender} icon={<Tag size={13} />} label="Vender" color={STATUS.disponivel.color} disabled={livro.quantidade <= 0} />
        <AcaoBtn onClick={onRepor} icon={<PackagePlus size={13} />} label="Repor estoque" color={ACCENT} />
      </div>
    </div>
  );
}

function AcaoBtn({ onClick, icon, label, color, disabled }) {
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      style={{ borderColor: color, color, fontFamily: MONO, fontSize: 11.5, opacity: disabled ? 0.35 : 1 }}
      className="border rounded-full px-2.5 py-1.5 flex items-center gap-1.5 hover:opacity-80 transition"
    >
      {icon}
      {label}
    </button>
  );
}

// ---- token de design ----
const PAPER = "#EDE6D6";
const CARD = "#FBF8F1";
const INK = "#2B2620";
const INK_SOFT = "#75695A";
const ACCENT = "#790E0C";
const COR_MENU = "#790E0C"; // cor de marca da Verbo Shop, usada só nos menus
const RULE = "#C9BFA8";
const SERIF = "'Fraunces', serif";
const SANS = "'Inter', sans-serif";
const MONO = "'IBM Plex Mono', monospace";
const LOGO_URL = "data:image/png;base64,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";
