/* MODO DEMONSTRAÇÃO — só é usado quando o config.js ainda não tem o link da planilha.
   Roda o mesmo motor da planilha (Code.gs) dentro do navegador, guardando os dados só neste aparelho. */
(function () {
  const CHAVE = 'demo_felipe_burguer_v1';
  let salvo = {}; try { salvo = JSON.parse(localStorage.getItem(CHAVE) || '{}'); } catch (e) {}
  const limpa = v => typeof v === 'string' && v.startsWith("'") ? v.slice(1) : (v === undefined || v === null ? '' : v);
  function Range(sh, r, c, nr, nc) { this.sh = sh; this.r = r; this.c = c; this.nr = nr || 1; this.nc = nc || 1; }
  Range.prototype.getValues = function () { const o = []; for (let i = 0; i < this.nr; i++) { const row = this.sh.d[this.r - 1 + i] || []; const x = []; for (let j = 0; j < this.nc; j++) x.push(row[this.c - 1 + j] ?? ''); o.push(x); } return o; };
  Range.prototype.setValues = function (v) { v.forEach((row, i) => row.forEach((val, j) => this.sh.set(this.r + i, this.c + j, val))); return this; };
  Range.prototype.setValue = function (v) { this.sh.set(this.r, this.c, v); return this; };
  Range.prototype.setNumberFormat = function () { return this; };
  Range.prototype.setFontWeight = function () { return this; };
  function Sheet(d) { this.d = d || []; }
  Sheet.prototype.set = function (r, c, v) { while (this.d.length < r) this.d.push([]); this.d[r - 1][c - 1] = limpa(v); };
  Sheet.prototype.getLastRow = function () { let n = this.d.length; while (n && !this.d[n - 1].some(x => x !== '' && x !== undefined && x !== null)) n--; return n; };
  Sheet.prototype.getMaxRows = function () { return 1000; };
  Sheet.prototype.getRange = function (r, c, nr, nc) { return new Range(this, r, c, nr, nc); };
  Sheet.prototype.appendRow = function (a) { const r = this.getLastRow() + 1; a.forEach((v, j) => this.set(r, j + 1, v)); };
  Sheet.prototype.deleteRow = function (r) { this.d.splice(r - 1, 1); };
  Sheet.prototype.setFrozenRows = function () {};
  const abas = {}; Object.keys(salvo.abas || {}).forEach(k => abas[k] = new Sheet(salvo.abas[k]));
  const props = Object.assign({}, salvo.props || {}), cache = {};
  const uuid = () => (crypto.randomUUID ? crypto.randomUUID() : 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => { const r = Math.random() * 16 | 0; return (c === 'x' ? r : (r & 3 | 8)).toString(16); }));
  const fmt = (d, tz, f) => {
    const p = {}; new Intl.DateTimeFormat('en-CA', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' }).formatToParts(d).forEach(x => p[x.type] = x.value);
    return f.replace("'T'", 'T').replace('yyyy', p.year).replace('MM', p.month).replace('dd', p.day).replace('HH', p.hour).replace('mm', p.minute).replace('ss', p.second);
  };
  const SpreadsheetApp = { getActiveSpreadsheet: () => ({ getSheetByName: n => abas[n] || null, insertSheet: n => (abas[n] = new Sheet()) }) };
  const CacheService = { getScriptCache: () => ({ get: k => k in cache ? cache[k] : null, put: (k, v) => { cache[k] = String(v); }, remove: k => { delete cache[k]; } }) };
  const PropertiesService = { getScriptProperties: () => ({ getProperty: k => k in props ? props[k] : null, setProperty: (k, v) => { props[k] = String(v); } }) };
  const LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) };
  const Utilities = { formatDate: fmt, getUuid: uuid, base64Decode: s => s, newBlob: (b, m, n) => ({ b, m, n }) };
  const DriveApp = { Access: {}, Permission: {}, getFoldersByName: () => ({ hasNext: () => false }), createFolder: () => ({ createFile: () => ({ getId: () => uuid(), setSharing() {} }) }) };
  const ContentService = { MimeType: { JSON: 'json' }, createTextOutput: s => ({ s, setMimeType() { return this; } }) };

  const motor = (function () {
/**
 * FELIPE BURGUER DRINKS — Cardápio digital (motor na Planilha Google + Drive)
 * Sales Soluções
 *
 * INSTALAÇÃO (uma vez):
 *  1) Crie uma planilha nova > Extensões > Apps Script > apague tudo e cole este código.
 *  2) Troque o PIN_INICIAL abaixo (senha do painel, 6 dígitos ou mais). Salve.
 *  3) Escolha a função "setup" e clique em Executar (autorize o acesso).
 *  4) Implantar > Nova implantação > App da Web
 *       Executar como: Eu  |  Quem pode acessar: Qualquer pessoa
 *  5) Copie o link (termina em /exec) e cole no config.js do site.
 *
 * ATUALIZAR depois: cole o código novo, salve e vá em
 *  Implantar > Gerenciar implantações > lápis > Versão: Nova versão > Implantar
 *  (NÃO use "Nova implantação", senão o link muda).
 *
 * Depois de instalado, a senha é trocada pelo painel (Ajustes). Mudar o PIN_INICIAL aqui não altera a senha.
 */

const PIN_INICIAL = '123456';
const PASTA_FOTOS = 'Felipe Burguer Drinks - Fotos';

const ABAS = {
  Pedidos:   ['id', 'dia', 'numero', 'hora', 'criado_em', 'atualizado_em', 'turno', 'cliente', 'telefone', 'tipo', 'endereco', 'agendado',
              'resumo', 'subtotal', 'taxa', 'desconto', 'total', 'pagamento', 'troco', 'cupom', 'fidelidade', 'obs', 'status', 'codigo', 'itens'],
  Produtos:  ['id', 'nome', 'descricao', 'preco', 'categoria', 'turnos', 'qtd_proteinas', 'sabores', 'conta_limite', 'foto', 'ativo', 'esgotado', 'destaque', 'ordem', 'opcao_titulo', 'sabores_off', 'com_acomp', 'preco_noite'],
  Proteinas: ['id', 'nome', 'almoco', 'noite', 'esgotado', 'ordem', 'turnos'],
  Despesas:  ['id', 'dia', 'descricao', 'categoria', 'valor'],
  Cupons:    ['id', 'codigo', 'tipo', 'valor', 'minimo', 'ativo', 'usos'],
  Clientes:  ['telefone', 'nome', 'pedidos', 'entregues', 'desde_premio', 'ultimo_pedido']
};
const COLS_TEXTO = { Pedidos: ['dia', 'criado_em', 'atualizado_em', 'telefone', 'agendado', 'codigo'], Despesas: ['dia'], Clientes: ['telefone', 'ultimo_pedido'], Cupons: ['codigo'] };
const STATUS = ['Aguardando', 'Novo', 'Preparando', 'Pronto', 'Saiu', 'Entregue', 'Cancelado'];

const CONFIG_PADRAO = {
  tz: 'America/Fortaleza',
  slogan: 'Eleito o melhor da Região',
  pedidos_ativos: true,
  modo: 'auto',   // abre e fecha sozinho pelo horário (no painel dá pra forçar Aberta/Fechada)
  turnos: {
    almoco: { nome: 'Cardápio', ativo: true, inicio: '18:00', fim: '23:59' },
    noite:  { nome: 'Madrugada', ativo: false, inicio: '00:00', fim: '01:00' }
  },
  entrega: { ativo: true, taxa: 0, bairros: [], acrescimo_noite: { valor: 0, inicio: '23:30', fim: '05:00' },
    gratis_combo: { ativo: false, palavra: 'entrega grátis', ate: 0 } }, retirada: true, local: true,
  escolher_acomp: false, tel_local: false,
  pix_antes: false, pix_minutos: 15, cidade: '',
  pagamentos: { pix: true, dinheiro: true, cartao: true },
  chave_pix: '', whatsapp: '',
  endereco: '',
  link: '',
  acompanhamentos: { almoco: '', noite: '' },
  msg_inicio: '', msg_fim: '',
  limite: { ativo: false, almoco: 50, noite: 50 },
  arte_abertura: false, pedir_de_novo: true, encomenda: false, cupons: false,
  fidelidade: { ativo: false, modo: 'pedidos', a_cada: 10, valor_ponto: 10, desconto: 10, tipo_premio: 'valor', minimo: 0,
                tipos: { entrega: true, retirada: true, local: true } },
  categorias: ['Hambúrgueres', 'Acompanhamentos', 'Adicionais', 'Pizzas', 'Pizzas Doces', 'Massas', 'Bebidas', 'Drinks']
};

/* =====================================================================
   INSTALAÇÃO
   ===================================================================== */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(ABAS).forEach(function (nome) {
    let sh = ss.getSheetByName(nome);
    if (!sh) sh = ss.insertSheet(nome);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, ABAS[nome].length).setValues([ABAS[nome]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    (COLS_TEXTO[nome] || []).forEach(function (c) {
      const col = ABAS[nome].indexOf(c) + 1;
      sh.getRange(1, col, sh.getMaxRows(), 1).setNumberFormat('@');
    });
  });
  const p = props_();
  if (!p.getProperty('config')) p.setProperty('config', JSON.stringify(CONFIG_PADRAO));
  if (!p.getProperty('nome')) p.setProperty('nome', 'Felipe Burguer Drinks');
  if (!p.getProperty('pin')) p.setProperty('pin', String(PIN_INICIAL));
  if (!p.getProperty('seq')) p.setProperty('seq', '0');

  if (aba_('Produtos').getLastRow() < 2) {
    const prod = [
      ["Burguer 1 (90g)", "Pão brioche, hambúrguer 90g, cheddar e molho da casa.", 25, "Hambúrgueres", "", 0, "", false, "", false, 1, ""],
      ["Cheeseburguer (160g)", "Pão brioche, hambúrguer 160g, cheddar, coalho e molho da casa.", 30, "Hambúrgueres", "", 0, "", false, "", false, 2, ""],
      ["Clássico da Casa (160g)", "Pão brioche, hambúrguer 160g, cheddar, coalho, alface, tomate, cebola caramelizada e molho da casa.", 35, "Hambúrgueres", "", 0, "", false, "", true, 3, ""],
      ["Bacon Burguer (160g)", "Pão brioche, hambúrguer 160g, cheddar, bacon em tiras e molho da casa.", 35, "Hambúrgueres", "", 0, "", false, "", false, 4, ""],
      ["Bacon Picante ou Barbecue (160g)", "Pão brioche, hambúrguer 160g, cheddar, bacon, cebola caramelizada e molho picante ou barbecue.", 35, "Hambúrgueres", "", 0, "Picante, Barbecue", false, "", false, 5, "Molho"],
      ["Burguer da Casa (160g)", "Pão brioche, hambúrguer 160g, cheddar, coalho, bacon, ovo, cebola caramelizada e molho da casa.", 40, "Hambúrgueres", "", 0, "", false, "img/p/burguer.jpg", true, 6, ""],
      ["Frango Desfiado (160g)", "Pão brioche, frango 160g, alface, tomate e molho da casa.", 28, "Hambúrgueres", "", 0, "", false, "", false, 7, ""],
      ["Bacon Burguer Doble (2x 160g)", "Pão brioche, duas carnes 160g, cheddar, bacon, cebola caramelizada e molho da casa.", 50, "Hambúrgueres", "", 0, "", false, "", true, 8, ""],
      ["Burguer Tropical (160g)", "Pão brioche, hambúrguer 160g, abacaxi, alface americana e tomate seco.", 40, "Hambúrgueres", "", 0, "", false, "", false, 9, ""],
      ["Burguer Calabresa (160g)", "Pão brioche, hambúrguer 160g, cheddar, calabresa em rodelas e molho da casa.", 35, "Hambúrgueres", "", 0, "", false, "", false, 10, ""],
      ["Batata 150g", "Porção de batata frita 150g.", 15, "Acompanhamentos", "", 0, "", false, "", false, 11, ""],
      ["Batata 300g", "Porção de batata frita 300g.", 30, "Acompanhamentos", "", 0, "", false, "", false, 12, ""],
      ["Batata Cheddar e Bacon 300g", "Batata frita 300g com cheddar e bacon.", 45, "Acompanhamentos", "", 0, "", false, "", false, 13, ""],
      ["Batata Carne do Sol e Cheddar 300g", "Batata frita 300g com carne do sol e cheddar.", 45, "Acompanhamentos", "", 0, "", false, "", false, 14, ""],
      ["Adicional de molho", "Maionese temperada ou ketchup.", 3, "Adicionais", "", 0, "Maionese temperada, Ketchup", false, "", false, 15, "Molho"],
      ["Adicional de salada, ovo, bacon ou calabresa", "Escolha o adicional para o seu burguer.", 5, "Adicionais", "", 0, "Salada, Ovo, Bacon, Calabresa", false, "", false, 16, "Adicional"],
      ["Adicional de carne", "Carne extra de 90g ou 150g.", 10, "Adicionais", "", 0, "90g – R$ 10, 150g – R$ 15", false, "", false, 17, "Tamanho"],
      ["Pizza Portuguesa", "Presunto, calabresa, azeitona, ovo, cebola e tomate.", 60, "Pizzas", "", 0, "Média – R$ 60, Grande – R$ 70", false, "", true, 18, "Tamanho"],
      ["Pizza Calabresa", "Mussarela, calabresa e cebola.", 55, "Pizzas", "", 0, "Média – R$ 55, Grande – R$ 65", false, "", false, 19, "Tamanho"],
      ["Pizza Marguerita", "Mussarela, tomate e manjericão.", 55, "Pizzas", "", 0, "Média – R$ 55, Grande – R$ 65", false, "", false, 20, "Tamanho"],
      ["Pizza Mussarela", "Mussarela e tomate.", 55, "Pizzas", "", 0, "Média – R$ 55, Grande – R$ 65", false, "", false, 21, "Tamanho"],
      ["Pizza Mista", "Mussarela e presunto.", 55, "Pizzas", "", 0, "Média – R$ 55, Grande – R$ 65", false, "", false, 22, "Tamanho"],
      ["Pizza Frango", "Mussarela, frango e catupiry.", 58, "Pizzas", "", 0, "Média – R$ 58, Grande – R$ 68", false, "", false, 23, "Tamanho"],
      ["Pizza Bacon", "Mussarela e bacon.", 58, "Pizzas", "", 0, "Média – R$ 58, Grande – R$ 68", false, "", false, 24, "Tamanho"],
      ["Pizza Calabresa Picante", "Pimenta calabresa, bastante cebola e mussarela.", 60, "Pizzas", "", 0, "Média – R$ 60, Grande – R$ 70", false, "", false, 25, "Tamanho"],
      ["Pizza Pepperoni", "Mussarela e pepperoni.", 65, "Pizzas", "", 0, "Média – R$ 65, Grande – R$ 75", false, "", false, 26, "Tamanho"],
      ["Pizza Quatro Queijos", "Mussarela, gorgonzola, provolone e catupiry.", 67, "Pizzas", "", 0, "Média – R$ 67, Grande – R$ 77", false, "", false, 27, "Tamanho"],
      ["Pizza Vegetariana", "", 67, "Pizzas", "", 0, "Média – R$ 67, Grande – R$ 77", false, "", false, 28, "Tamanho"],
      ["Pizza Atum", "Mussarela, atum e cebola.", 67, "Pizzas", "", 0, "Média – R$ 67, Grande – R$ 77", false, "", false, 29, "Tamanho"],
      ["Pizza Isca de Filé Mignon", "Mussarela, isca de filé, champignon e molho.", 70, "Pizzas", "", 0, "Média – R$ 70, Grande – R$ 80", false, "img/p/pizza-isca.jpg", true, 30, "Tamanho"],
      ["Pizza Camarão", "Mussarela, molho de tomate e manjericão.", 70, "Pizzas", "", 0, "Média – R$ 70, Grande – R$ 80", false, "", false, 31, "Tamanho"],
      ["Pizza Carne do Sol", "Mussarela, coalho e cebola.", 70, "Pizzas", "", 0, "Média – R$ 70, Grande – R$ 80", false, "", false, 32, "Tamanho"],
      ["Pizza Chocolate", "", 60, "Pizzas Doces", "", 0, "Média – R$ 60, Grande – R$ 70", false, "", false, 33, "Tamanho"],
      ["Pizza Banana com Doce de Leite", "", 60, "Pizzas Doces", "", 0, "Média – R$ 60, Grande – R$ 70", false, "", false, 34, "Tamanho"],
      ["Isca de Filé Mignon (Talharim Artesanal)", "Molho de tomate pelado, shoyo e gengibre. Serve 2 pessoas.", 120, "Massas", "", 0, "", false, "img/p/massa-isca.jpg", false, 35, ""],
      ["Camarão (Talharim Artesanal)", "Molho de tomate pelado, manjericão e camarões flambados. Serve 2 pessoas.", 120, "Massas", "", 0, "", false, "", false, 36, ""],
      ["Quatro Queijos com Iscas (Talharim Artesanal)", "Molho quatro queijos com iscas de mignon. Serve 2 pessoas.", 130, "Massas", "", 0, "", false, "", false, 37, ""]
    ];
    aba_('Produtos').getRange(2, 1, prod.length, ABAS.Produtos.length).setValues(prod.map(function (x) {
      return [novoId_(), x[0], x[1], x[2], x[3], x[4], x[5], x[6], x[7], x[8], true, false, x[9], x[10], x[11] || '', '', false, ''];
    }));
  }
  if (aba_('Proteinas').getLastRow() < 2) {
    const pr = [];
    if (pr.length) aba_('Proteinas').getRange(2, 1, pr.length, ABAS.Proteinas.length).setValues(pr.map(function (n, i) { return [novoId_(), n, true, true, false, i + 1, 'almoco,noite']; }));
  }
  pasta_();
  limparCache_();
}

/* =====================================================================
   ENTRADA (o site chama aqui)
   ===================================================================== */
function doGet(e) {
  const acao = e && e.parameter && e.parameter.acao;
  if (acao === 'cardapio') return json_(executar_({ acao: 'cardapio' }));
  return json_({ ok: true, app: 'Felipe Burguer Drinks' });
}

function doPost(e) {
  let b;
  try { b = JSON.parse((e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, erro: 'Requisição inválida' }); }
  return json_(executar_(b));
}

const PUBLICAS = {
  cardapio: function () { return cardapio_(); },
  acompanhar: function (b) { return acompanhar_(b.p_id, b.p_codigo); },
  checar_cupom: function (b) { return checarCupom_(b.p_codigo, b.p_subtotal); },
  fidelidade_status: function (b) { return fidelidadeStatus_(b.p_telefone); },
  criar_pedido: function (b) { return comTrava_(function () { return criarPedido_(b.p || {}, false); }); }
};
const ADMIN = {
  admin_login: function () { return { ok: true }; },
  admin_dados: function () { return adminDados_(); },
  admin_pedidos: function (b) { return { ok: true, versao: versao_(), agora: new Date().toISOString(), pedidos: pedidosDoDia_(String(b.p_dia || hoje_().dia)) }; },
  admin_mudancas: function (b) { cancelarVencidos_(); return adminMudancas_(b.p_versao); },
  admin_confirmar_pix: function (b) { return comTrava_(function () { return confirmarPix_(String(b.p_id)); }); },
  admin_status: function (b) { return comTrava_(function () { return adminStatus_(b.p_id, b.p_status); }); },
  admin_produto: function (b) { return comTrava_(function () { return salvarProduto_(b.p || {}); }); },
  admin_proteina: function (b) { return comTrava_(function () { return salvarProteina_(b.p || {}); }); },
  admin_cupom: function (b) { return comTrava_(function () { return salvarCupom_(b.p || {}); }); },
  admin_excluir: function (b) { return comTrava_(function () { return excluir_(b.p_tipo, b.p_id); }); },
  admin_flag: function (b) { return comTrava_(function () { return flag_(b.p_tipo, b.p_id, b.p_campo, b.p_valor); }); },
  admin_flags: function (b) { return comTrava_(function () { return flags_(b.p_lista || []); }); },
  admin_config: function (b) { return comTrava_(function () { return salvarConfig_(b.p, b.p_nome); }); },
  admin_foto: function (b) { return { ok: true, url: enviarFoto_(b) }; },
  admin_senha_dono_criar: function (b) {   // só funciona enquanto o dono ainda não criou a senha dele
    if (props_().getProperty('pin_dono')) return { ok: false, erro: 'A senha do dono já existe.' };
    const nova = String(b.nova || ''); if (nova.length < 6) return { ok: false, erro: 'A senha precisa ter pelo menos 6 dígitos.' };
    if (nova === String(props_().getProperty('pin') || PIN_INICIAL)) return { ok: false, erro: 'Use uma senha diferente da senha do painel.' };
    props_().setProperty('pin_dono', nova); return { ok: true };
  },
  admin_pedido_manual: function (b) { return comTrava_(function () { return criarPedido_(b.p || {}, true); }); }
};

function executar_(b) {
  try {
    const acao = String(b.acao || '');
    if (PUBLICAS[acao]) return PUBLICAS[acao](b);
    if (ADMIN[acao] || DONO[acao]) {
      const cc = CacheService.getScriptCache(), falhas = parseInt(cc.get('pin_falhas') || '0', 10);
      if (falhas >= 10) return { ok: false, erro: 'Muitas tentativas erradas. Aguarde 10 minutos.' };
      if (String(b.pin || '') !== String(props_().getProperty('pin') || PIN_INICIAL)) {
        cc.put('pin_falhas', String(falhas + 1), 600);   // trava por 10 min depois de 10 erros
        return { ok: false, erro: 'PIN inválido' };
      }
      if (DONO[acao] || (acao === 'admin_excluir' && b.p_tipo === 'despesa')) {
        const pd = props_().getProperty('pin_dono');
        if (!pd) return { ok: false, erro: 'SEM_SENHA_DONO' };
        if (String(b.pin_dono || '') !== pd) { cc.put('pin_falhas', String(falhas + 1), 600); return { ok: false, erro: 'Senha do dono inválida' }; }
        if (DONO[acao]) return DONO[acao](b);
      }
      return ADMIN[acao] ? ADMIN[acao](b) : { ok: false, erro: 'Ação desconhecida' };
    }
    return { ok: false, erro: 'Ação desconhecida' };
  } catch (err) {
    return { ok: false, erro: String(err && err.message || err) };
  }
}

/* =====================================================================
   CARDÁPIO (público)
   ===================================================================== */
function cardapio_() {
  const cache = CacheService.getScriptCache();
  let base = cache.get('menu');
  if (base) base = JSON.parse(base);
  else {
    const cfg = config_();
    base = {
      nome: nome_(), config: cfg,
      produtos: ler_('Produtos').map(prod_).filter(function (p) { return p.ativo; }).map(function (p) { delete p.ativo; return p; }),
      proteinas: ler_('Proteinas').map(prot_).filter(function (x) { return x.almoco || x.noite; }).sort(porOrdem_)
                   .map(function (x) { return { id: x.id, nome: x.nome, almoco: x.almoco, noite: x.noite, esgotado: x.esgotado }; })
    };
    base.produtos.sort(porOrdem_);
    guardarCache_('menu', base, 21600);
  }
  const h = hoje_(base.config);
  base.ok = true;
  base.agora = h.iso;
  base.turno_atual = turnoEm_(base.config, h.hm);
  base.vendidos = vendidos_(h.dia);
  return base;
}

/* =====================================================================
   PEDIDO NOVO
   ===================================================================== */
function listaAcomp_(txt) {
  return String(txt || '').split(/,|;|\s+e\s+/i).map(function (s) { return s.trim(); }).filter(Boolean);
}
function taxaBairro_(cfg, bairro) {
  const bs = (cfg.entrega && Array.isArray(cfg.entrega.bairros)) ? cfg.entrega.bairros : [];
  if (!bs.length) return { ok: true, taxa: num_(cfg.entrega && cfg.entrega.taxa) };
  const b = norm_(bairro);
  const x = bs.filter(function (y) { return norm_(y.nome) === b; })[0];
  return x ? { ok: true, taxa: num_(x.taxa), nome: x.nome } : { ok: false };
}
// Acréscimo na taxa de entrega de madrugada (ex.: +R$ 2 depois da meia-noite)
function acrescimoNoite_(cfg, hm) {
  const a = cfg.entrega && cfg.entrega.acrescimo_noite;
  if (!a || !num_(a.valor)) return 0;
  const m = min_(hm), i = min_(a.inicio || '00:00'), f = min_(a.fim || '05:00');
  const dentro = i <= f ? (m >= i && m < f) : (m >= i || m < f);
  return dentro ? num_(a.valor) : 0;
}
// Combo com "Entrega Grátis" no nome: zera a taxa nos bairros próximos (taxa até o limite)
function entregaGratis_(cfg, itens, bairro) {
  const g = cfg.entrega && cfg.entrega.gratis_combo;
  if (!g || !sim_(g.ativo, false)) return false;
  const tem = itens.some(function (x) { return norm_(x.nome).indexOf(norm_(g.palavra || 'entrega gratis')) >= 0; });
  if (!tem) return false;
  const tb = taxaBairro_(cfg, bairro);
  return tb.ok && tb.taxa <= num_(g.ate || 0);
}
// opção com preço próprio: termina em "R$ 70" (ex.: "Média – R$ 60")
function precoOpcao_(s) { const m = /R\$\s*(\d+(?:\.\d{1,2})?)\s*\)?\s*$/.exec(String(s || '')); return m ? Number(m[1]) : null; }
function norm_(s) { return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim(); }

function criarPedido_(p, manual) {
  const cfg = config_();
  const erro = function (m) { return { ok: false, erro: m }; };
  if (!manual && !sim_(cfg.pedidos_ativos, true)) return erro('Pedidos online desativados no momento.');
  const h = hoje_(cfg), modo = manual ? 'aberta' : (cfg.modo || 'auto');
  if (modo === 'fechada') return erro('Estamos fechados no momento.');
  const turno = String(p.turno || '');
  if (['almoco', 'noite'].indexOf(turno) < 0 || (!manual && !turnoAtivo_(cfg, turno))) return erro('Turno indisponível.');
  const ini = min_(cfg.turnos[turno].inicio), fim = min_(cfg.turnos[turno].fim), agoraMin = min_(h.hm);
  const atual = turnoEm_(cfg, h.hm);
  const agendado = String(p.agendado || '').trim();
  if (agendado) {
    if (!sim_(cfg.encomenda, false)) return erro('Encomendas com horário estão desativadas.');
    const a = min_(agendado);
    if (a === null) return erro('Horário inválido.');
    if (a < ini || a > fim) return erro('Escolha um horário entre ' + cfg.turnos[turno].inicio + ' e ' + cfg.turnos[turno].fim + '.');
    if (a <= agoraMin) return erro('Escolha um horário mais tarde.');
  } else if (modo !== 'aberta' && atual !== turno) {
    return erro(atual ? 'Esse cardápio não está disponível agora.' : 'Estamos fechados agora.');
  }

  const nome = String(p.cliente || '').trim().slice(0, 60) || (manual ? 'Balcão' : '');
  const tel = String(p.telefone || '').replace(/\D/g, '');
  if (nome.length < 2) return erro('Informe seu nome.');
  const telOpcional = manual || (String(p.tipo || '') === 'local' && !sim_(cfg.tel_local, false));   // comer no local: WhatsApp é opcional
  if (telOpcional ? (tel && (tel.length < 10 || tel.length > 13)) : (tel.length < 10 || tel.length > 13)) return erro('Informe o telefone com DDD.');

  const tipo = String(p.tipo || '');
  let endereco = String(p.endereco || '').trim().slice(0, 250), taxa = 0;
  if (tipo === 'entrega') {
    if (!manual && !sim_(cfg.entrega && cfg.entrega.ativo, false)) return erro('Entrega indisponível no momento.');
    if (endereco.length < 5) return erro('Informe o endereço de entrega.');
    const tb = taxaBairro_(cfg, p.bairro);
    if (!tb.ok) return erro('Escolha o bairro da entrega.');
    taxa = manual && p.taxa !== undefined && p.taxa !== '' ? Math.max(0, num_(p.taxa)) : tb.taxa + acrescimoNoite_(cfg, h.hm);
  } else if (tipo === 'retirada') {
    if (!manual && !sim_(cfg.retirada, true)) return erro('Retirada indisponível.'); endereco = '';
  } else if (tipo === 'local') {
    if (!manual && !sim_(cfg.local, true)) return erro('Consumo no local indisponível.'); endereco = '';
  } else return erro('Escolha como quer receber o pedido.');

  if (!Array.isArray(p.itens) || !p.itens.length) return erro('Seu carrinho está vazio.');
  if (p.itens.length > 40) return erro('Pedido grande demais.');
  const produtos = {}; ler_('Produtos').map(prod_).forEach(function (x) { produtos[x.id] = x; });
  const proteinas = {}; ler_('Proteinas').map(prot_).forEach(function (x) { proteinas[x.id] = x; });

  const itens = []; let sub = 0, lim = 0;
  for (let i = 0; i < p.itens.length; i++) {
    const it = p.itens[i] || {}, pr = produtos[String(it.id)];
    if (!pr || !pr.ativo) return erro('Um item do carrinho não está mais disponível. Atualize a página.');
    if (pr.esgotado) return erro(pr.nome + ' esgotou.');
    if (pr.turnos.length && pr.turnos.indexOf(turno) < 0) return erro(pr.nome + ' não está disponível neste turno.');
    const acompLista = listaAcomp_((cfg.acompanhamentos || {})[turno]);
    const qtd = Math.min(parseInt(it.qtd, 10) || 0, 50);
    if (qtd <= 0) continue;
    let det = '';
    if (pr.qtd_proteinas > 0) {
      const ids = Array.isArray(it.proteinas) ? it.proteinas : [];
      if (ids.length < 1 || ids.length > pr.qtd_proteinas) return erro(pr.qtd_proteinas === 1 ? 'Escolha a proteína de ' + pr.nome + '.' : 'Escolha de 1 a ' + pr.qtd_proteinas + ' proteínas para ' + pr.nome + '.');   // o cliente escolhe ATÉ o máximo
      for (let k = 0; k < ids.length; k++) {
        const x = proteinas[String(ids[k])];
        if (!x || x.esgotado || !x[turno]) return erro('Uma proteína escolhida acabou. Escolha outra, por favor.');
        det += (det ? ' + ' : '') + x.nome;
      }
    }
    if (pr.qtd_proteinas > 0 && pr.com_acomp && acompLista.length && Array.isArray(it.acomp)) {
      const esc = it.acomp.map(function (a) { return norm_(a); });
      const com = acompLista.filter(function (a) { return esc.indexOf(norm_(a)) >= 0; });
      det += (det ? ' · ' : '') + (com.length ? 'com ' + com.join(', ').toLowerCase() : 'sem acompanhamentos');
    }
    if (pr.sabores.length) {
      if (pr.sabores_off.indexOf(String(it.sabor || '')) >= 0) return erro('O sabor ' + it.sabor + ' de ' + pr.nome + ' acabou. Escolha outro, por favor.');
      if (pr.sabores.indexOf(String(it.sabor || '')) < 0) return erro('Escolha ' + (pr.opcao_titulo ? pr.opcao_titulo.toLowerCase() : 'o sabor') + ' de ' + pr.nome + '.');
      det += (det ? ' · ' : '') + (pr.opcao_titulo ? pr.opcao_titulo + ': ' : '') + it.sabor;
    }
    let preco = turno === 'noite' && num_(pr.preco_noite) > 0 ? num_(pr.preco_noite) : pr.preco;   // preço do turno
    if (pr.sabores.length) { const po = precoOpcao_(it.sabor); if (po !== null) preco = po; }   // opção com preço (ex.: "Grande – R$ 70")
    itens.push({ id: pr.id, nome: pr.nome, qtd: qtd, preco: preco, detalhes: det, obs: String(it.obs || '').slice(0, 120), lim: pr.conta_limite, cat: pr.categoria });
    sub += preco * qtd;
    if (pr.conta_limite) lim += qtd;
  }
  if (!itens.length) return erro('Seu carrinho está vazio.');
  if (tipo === 'entrega' && !manual && entregaGratis_(cfg, itens, p.bairro)) taxa = 0;

  const doDia = pedidosDoDia_(h.dia);
  if (cfg.limite && sim_(cfg.limite.ativo, false) && lim > 0) {
    const rest = (parseInt(cfg.limite[turno], 10) || 0) - contarVendidos_(doDia, turno);
    if (lim > rest) return erro(rest <= 0 ? 'As marmitas/pratinhos esgotaram por hoje.' : rest === 1 ? 'Só resta 1 marmita/pratinho hoje.' : 'Só restam ' + rest + ' marmitas/pratinhos hoje.');
  }

  let desc = 0, codCupom = '', cupomRow = null;
  const cod = String(p.cupom || '').trim().toUpperCase();
  if (cod && sim_(cfg.cupons, false)) {
    const c = ler_('Cupons').map(cupom_).filter(function (x) { return x.codigo === cod && x.ativo; })[0];
    if (!c) return erro('Cupom inválido ou expirado.');
    if (sub < c.minimo) return erro('Esse cupom vale para pedidos a partir de ' + brl_(c.minimo) + '.');
    desc = c.tipo === 'pct' ? Math.round(sub * c.valor) / 100 : c.valor;
    codCupom = c.codigo; cupomRow = c;
  }

  let fid = false;
  const cli = tel ? cliente_(tel) : null;
  if (tel && cfg.fidelidade && sim_(cfg.fidelidade.ativo, false) && cli && cli.desde_premio >= Math.max(parseInt(cfg.fidelidade.a_cada, 10) || 10, 1)) {
    fid = true; desc += cfg.fidelidade.tipo_premio === 'pct' ? Math.round(sub * num_(cfg.fidelidade.desconto)) / 100 : num_(cfg.fidelidade.desconto);
  }
  desc = Math.min(desc, sub);
  const total = Math.round((sub + taxa - desc) * 100) / 100;

  const pag = String(p.pagamento || '');
  const chavePag = { 'Pix': 'pix', 'Dinheiro': 'dinheiro', 'Cartão': 'cartao' }[pag];
  if (!chavePag || (!manual && !sim_(cfg.pagamentos && cfg.pagamentos[chavePag], true))) return erro('Escolha a forma de pagamento.');
  let troco = pag === 'Dinheiro' ? num_(p.troco) : 0;
  if (troco > 0 && troco < total) return erro('O troco precisa ser para um valor maior que o total.');

  const pixAntes = !manual && pag === 'Pix' && sim_(cfg.pix_antes, false) && String(cfg.chave_pix || '').trim() !== '';
  // grava
  if (cupomRow) atualizarLinha_('Cupons', cupomRow._row, { usos: cupomRow.usos + 1 });
  const id = proximoId_();
  const numero = doDia.reduce(function (m, x) { return Math.max(m, x.numero); }, 0) + 1;
  const agoraIso = new Date().toISOString();
  const codigo = (manual ? 'm' : '') + Utilities.getUuid().replace(/-/g, '').slice(0, manual ? 11 : 12);
  const resumo = itens.map(function (i) { return i.qtd + 'x ' + i.nome + (i.detalhes ? ' (' + i.detalhes + ')' : ''); }).join('; ');
  const ped = { id: id, dia: h.dia, numero: numero, hora: parseInt(h.hm, 10), criado_em: agoraIso, atualizado_em: agoraIso, turno: turno,
    cliente: seguro_(nome), telefone: tel, tipo: tipo, endereco: seguro_(endereco), agendado: agendado, resumo: seguro_(resumo),
    subtotal: sub, taxa: taxa, desconto: desc, total: total, pagamento: pag, troco: troco, cupom: codCupom, fidelidade: fid,
    obs: seguro_(String(p.obs || '').slice(0, 300)), status: manual ? 'Preparando' : (pixAntes ? 'Aguardando' : 'Novo'), codigo: codigo, itens: JSON.stringify(itens) };
  aba_('Pedidos').appendRow(ABAS.Pedidos.map(function (c) { return ped[c]; }));
  if (tel) registrarCliente_(cli, tel, nome, h.dia, fid);
  guardarCache_('p_' + id, { codigo: codigo, status: ped.status, numero: numero, tipo: tipo }, 21600);
  doDia.push(pedObj_(ABAS.Pedidos.map(function (c) { return ped[c]; })));
  guardarCache_('vend_' + h.dia, { almoco: contarVendidos_(doDia, 'almoco'), noite: contarVendidos_(doDia, 'noite') }, 21600);
  if (pixAntes) { const ag = pendentesPix_(); ag[id] = Date.now(); props_().setProperty('pix_pendentes', JSON.stringify(ag)); }
  subirVersao_();
  return { ok: true, id: id, numero: numero, codigo: codigo, total: total, desconto: desc, fidelidade: fid, status: ped.status, expira_em: pixAntes ? expiraPix_(Date.now(), cfg) : '' };
}

/* =====================================================================
   ACOMPANHAR / CUPOM / FIDELIDADE (público)
   ===================================================================== */
function acompanhar_(id, codigo) {
  let c = CacheService.getScriptCache().get('p_' + id);
  if (c) c = JSON.parse(c);
  else {
    const p = pedidoPorId_(id);
    if (p) { c = { codigo: p.codigo, status: p.status, numero: p.numero, tipo: p.tipo }; guardarCache_('p_' + id, c, 21600); }
  }
  if (!c || !codigo || String(c.codigo) !== String(codigo)) return { ok: false, erro: 'Pedido não encontrado' };
  if (c.status === 'Aguardando') {
    cancelarVencidos_();
    const ag = pendentesPix_();
    if (!ag[id]) { const p = pedidoPorId_(id); if (p) c.status = p.status; }
    else return { ok: true, status: 'Aguardando', numero: c.numero, tipo: c.tipo, expira_em: expiraPix_(ag[id], config_()) };
  }
  return { ok: true, status: c.status, numero: c.numero, tipo: c.tipo };
}

/* ---------- Pix antes de preparar ---------- */
function pendentesPix_() { try { return JSON.parse(props_().getProperty('pix_pendentes') || '{}'); } catch (e) { return {}; } }
function expiraPix_(ms, cfg) { return new Date(Number(ms) + Math.max(parseInt(cfg.pix_minutos, 10) || 15, 5) * 60000).toISOString(); }
// cancela sozinho quem não pagou no prazo (só olha a lista de pendentes, é rápido)
function cancelarVencidos_() {
  const ag = pendentesPix_(), ids = Object.keys(ag); if (!ids.length) return;
  const cfg = config_(), agora = Date.now(); let mudou = false;
  ids.forEach(function (id) {
    if (agora > new Date(expiraPix_(ag[id], cfg)).getTime()) {
      const p = pedidoPorId_(id);
      if (p && p.status === 'Aguardando') adminStatus_(id, 'Cancelado');
      delete ag[id]; mudou = true;
    }
  });
  if (mudou) props_().setProperty('pix_pendentes', JSON.stringify(ag));
}
function confirmarPix_(id) {
  const p = pedidoPorId_(id);
  if (!p) return { ok: false, erro: 'Pedido não encontrado' };
  if (p.status !== 'Aguardando') return { ok: false, erro: 'Esse pedido não está aguardando Pix.' };
  const ag = pendentesPix_(); delete ag[id]; props_().setProperty('pix_pendentes', JSON.stringify(ag));
  return adminStatus_(id, 'Novo');
}

function checarCupom_(codigo, subtotal) {
  const cfg = config_();
  if (!sim_(cfg.cupons, false)) return { ok: false, erro: 'Cupons indisponíveis.' };
  const cod = String(codigo || '').trim().toUpperCase();
  const c = ler_('Cupons').map(cupom_).filter(function (x) { return x.codigo === cod && x.ativo; })[0];
  if (!c) return { ok: false, erro: 'Cupom inválido ou expirado.' };
  if (num_(subtotal) < c.minimo) return { ok: false, erro: 'Vale para pedidos a partir de ' + brl_(c.minimo) + '.' };
  return { ok: true, codigo: c.codigo, tipo: c.tipo, valor: c.valor, minimo: c.minimo };
}

function fidelidadeStatus_(telefone) {
  const cfg = config_(), tel = String(telefone || '').replace(/\D/g, '');
  if (!cfg.fidelidade || !sim_(cfg.fidelidade.ativo, false) || tel.length < 10) return { ok: true, ativo: false };
  const c = cliente_(tel);
  const f = cfg.fidelidade;
  return { ok: true, ativo: true, a_cada: Math.max(parseInt(f.a_cada, 10) || 10, 1), modo: f.modo || 'pedidos', valor_ponto: num_(f.valor_ponto) || 10,
           desconto: num_(f.desconto), tipo_premio: f.tipo_premio || 'valor', minimo: num_(f.minimo), tipos: f.tipos || {}, feitos: c ? c.desde_premio : 0 };
}

/* =====================================================================
   PAINEL
   ===================================================================== */
function cabecalhos_() {   // garante o nome das colunas novas nas abas que já existiam
  Object.keys(ABAS).forEach(function (nome) {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome); if (!sh) return;
    const cab = ABAS[nome], at = sh.getRange(1, 1, 1, cab.length).getValues()[0];
    if (cab.some(function (c, i) { return at[i] !== c; })) sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
  });
}
function adminDados_() {
  if (props_().getProperty('cab_v') !== '5') { cabecalhos_(); props_().setProperty('cab_v', '5'); }
  const cfg = config_(), h = hoje_(cfg);
  return { ok: true, nome: nome_(), config: cfg, tem_senha_dono: !!props_().getProperty('pin_dono'), hoje: h.dia, agora: h.iso, turno_atual: turnoEm_(cfg, h.hm),
    vendidos: vendidos_(h.dia),
    produtos: ler_('Produtos').map(prod_).sort(porOrdem_),
    proteinas: ler_('Proteinas').map(prot_).sort(porOrdem_),
    cupons: ler_('Cupons').map(cupom_) };
}

// Só lê a planilha se algo mudou (deixa o painel rápido)
function adminMudancas_(v) {
  const atual = versao_();
  if (String(v) === String(atual)) return { ok: true, versao: atual, mudou: false };
  return { ok: true, versao: atual, mudou: true, hoje: hoje_().dia, pedidos: pedidosDoDia_(hoje_().dia) };
}

function adminStatus_(id, status) {
  if (STATUS.indexOf(status) < 0) return { ok: false, erro: 'Status inválido' };
  const p = pedidoPorId_(id);
  if (!p) return { ok: false, erro: 'Pedido não encontrado' };
  const antes = p.status;
  atualizarLinha_('Pedidos', p._row, { status: status, atualizado_em: new Date().toISOString() });
  p.status = status;
  if (antes !== status && (antes === 'Entregue' || status === 'Entregue')) ajustarEntregues_(p.telefone, status === 'Entregue' ? 1 : -1, pontosDoPedido_(config_(), p));
  guardarCache_('p_' + p.id, { codigo: p.codigo, status: status, numero: p.numero, tipo: p.tipo }, 21600);
  if (antes === 'Cancelado' || status === 'Cancelado') {
    const doDia = pedidosDoDia_(p.dia);
    guardarCache_('vend_' + p.dia, { almoco: contarVendidos_(doDia, 'almoco'), noite: contarVendidos_(doDia, 'noite') }, 21600);
  }
  subirVersao_();
  delete p._row; delete p.codigo;
  return { ok: true, pedido: p };
}

function salvarProduto_(p) {
  const nome = String(p.nome || '').trim();
  if (!nome) return { ok: false, erro: 'Informe o nome.' };
  const ex = p.id ? ler_('Produtos').filter(function (x) { return String(x.id) === String(p.id); })[0] : null;
  const o = {
    id: ex ? ex.id : novoId_(), nome: seguro_(nome.slice(0, 80)), descricao: seguro_(String(p.descricao || '').slice(0, 300)),
    preco: Math.max(num_(p.preco), 0), categoria: seguro_(String(p.categoria || '').trim() || 'Outros'),
    turnos: lista_(p.turnos).filter(function (t) { return t === 'almoco' || t === 'noite'; }).join(','),
    qtd_proteinas: Math.min(Math.max(parseInt(p.qtd_proteinas, 10) || 0, 0), 5), sabores: lista_(p.sabores).map(seguro_).join(', '),
    conta_limite: sim_(p.conta_limite, false), foto: (Array.isArray(p.fotos) ? p.fotos : [p.foto]).map(function (u) { return String(u || '').replace(/\|/g, '').trim(); }).filter(Boolean).slice(0, 5).join('|'),
    ativo: sim_(p.ativo, true), esgotado: sim_(p.esgotado, false),
    destaque: sim_(p.destaque, false), ordem: parseInt(p.ordem, 10) || 0, opcao_titulo: seguro_(String(p.opcao_titulo || '').trim().slice(0, 30)),
    sabores_off: lista_(p.sabores_off).filter(function (x) { return lista_(p.sabores).indexOf(x) >= 0; }).map(seguro_).join(', '),
    com_acomp: sim_(p.com_acomp, true),
    preco_noite: p.preco_noite === '' || p.preco_noite === null || p.preco_noite === undefined || !(num_(p.preco_noite) > 0) ? '' : num_(p.preco_noite)
  };
  if (ex) atualizarLinha_('Produtos', ex._row, o);
  else aba_('Produtos').appendRow(ABAS.Produtos.map(function (c) { return o[c]; }));
  limparCache_();
  return { ok: true, item: prod_(o) };
}

function salvarProteina_(p) {
  const nome = String(p.nome || '').trim();
  if (!nome) return { ok: false, erro: 'Informe o nome.' };
  const ex = p.id ? ler_('Proteinas').filter(function (x) { return String(x.id) === String(p.id); })[0] : null;
  let tur = lista_(p.turnos).filter(function (t) { return t === 'almoco' || t === 'noite'; });
  if (p.turnos === undefined || p.turnos === null) tur = ['almoco', 'noite'];
  if (!tur.length) return { ok: false, erro: 'Escolha pelo menos um turno (almoço ou noite).' };
  const o = { id: ex ? ex.id : novoId_(), nome: seguro_(nome.slice(0, 60)),
              almoco: tur.indexOf('almoco') >= 0 && sim_(p.almoco, true), noite: tur.indexOf('noite') >= 0 && sim_(p.noite, true),
              esgotado: sim_(p.esgotado, false), ordem: parseInt(p.ordem, 10) || 0, turnos: tur.join(',') };
  if (ex) atualizarLinha_('Proteinas', ex._row, o);
  else aba_('Proteinas').appendRow(ABAS.Proteinas.map(function (c) { return o[c]; }));
  limparCache_();
  return { ok: true, item: prot_(o) };
}

function salvarCupom_(p) {
  const cod = String(p.codigo || '').replace(/\s/g, '').toUpperCase().slice(0, 20);
  if (cod.length < 3) return { ok: false, erro: 'O código precisa ter pelo menos 3 letras.' };
  const lista = ler_('Cupons').map(cupom_);
  const ex = p.id ? lista.filter(function (x) { return x.id === String(p.id); })[0] : null;
  if (lista.some(function (x) { return x.codigo === cod && (!ex || x.id !== ex.id); })) return { ok: false, erro: 'Já existe um cupom com esse código.' };
  const o = { id: ex ? ex.id : novoId_(), codigo: cod, tipo: p.tipo === 'pct' ? 'pct' : 'valor', valor: Math.max(num_(p.valor), 0),
              minimo: Math.max(num_(p.minimo), 0), ativo: sim_(p.ativo, true), usos: ex ? ex.usos : 0 };
  if (ex) atualizarLinha_('Cupons', ex._row, o);
  else aba_('Cupons').appendRow(ABAS.Cupons.map(function (c) { return o[c]; }));
  return { ok: true, item: cupom_(o) };
}

function salvarDespesa_(p) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(p.dia || ''))) return { ok: false, erro: 'Informe a data.' };
  if (num_(p.valor) <= 0) return { ok: false, erro: 'Informe o valor.' };
  const ex = p.id ? ler_('Despesas').filter(function (x) { return String(x.id) === String(p.id); })[0] : null;
  const o = { id: ex ? ex.id : novoId_(), dia: String(p.dia), descricao: seguro_(String(p.descricao || '').slice(0, 120)),
              categoria: seguro_(String(p.categoria || 'Outros')), valor: num_(p.valor) };
  if (ex) atualizarLinha_('Despesas', ex._row, o);
  else aba_('Despesas').appendRow(ABAS.Despesas.map(function (c) { return o[c]; }));
  return { ok: true, item: despesa_(o) };
}

function excluir_(tipo, id) {
  const nome = { produto: 'Produtos', proteina: 'Proteinas', cupom: 'Cupons', despesa: 'Despesas' }[tipo];
  if (!nome) return { ok: false, erro: 'Tipo inválido' };
  const linha = ler_(nome).filter(function (x) { return String(x.id) === String(id); })[0];
  if (linha) aba_(nome).deleteRow(linha._row);
  if (tipo === 'produto' || tipo === 'proteina') limparCache_();
  return { ok: true };
}

function flag_(tipo, id, campo, valor) {
  const ok = { produto: ['ativo', 'esgotado', 'destaque'], proteina: ['almoco', 'noite', 'esgotado'] }[tipo];
  if (!ok || ok.indexOf(campo) < 0) return { ok: false, erro: 'Campo inválido' };
  const nome = tipo === 'produto' ? 'Produtos' : 'Proteinas';
  const linha = ler_(nome).filter(function (x) { return String(x.id) === String(id); })[0];
  if (!linha) return { ok: false, erro: 'Não encontrado' };
  const o = {}; o[campo] = sim_(valor, false);
  atualizarLinha_(nome, linha._row, o);
  limparCache_();
  return { ok: true };
}

// Vários liga/desliga de uma vez (ex.: "marcar todas", "novo dia")
function flags_(lista) {
  const ok = { produto: ['ativo', 'esgotado', 'destaque'], proteina: ['almoco', 'noite', 'esgotado'] };
  const tabelas = { produto: ler_('Produtos'), proteina: ler_('Proteinas') };
  lista.forEach(function (f) {
    if (!ok[f.tipo] || ok[f.tipo].indexOf(f.campo) < 0) return;
    const linha = tabelas[f.tipo].filter(function (x) { return String(x.id) === String(f.id); })[0];
    if (!linha) return;
    const o = {}; o[f.campo] = sim_(f.valor, false);
    atualizarLinha_(f.tipo === 'produto' ? 'Produtos' : 'Proteinas', linha._row, o);
  });
  limparCache_();
  return { ok: true };
}

function salvarConfig_(p, nome) {
  if (!p || typeof p !== 'object' || Array.isArray(p)) return { ok: false, erro: 'Dados inválidos' };
  const cfg = config_();
  Object.keys(p).forEach(function (k) { cfg[k] = p[k]; });
  props_().setProperty('config', JSON.stringify(cfg));
  if (nome && String(nome).trim()) props_().setProperty('nome', String(nome).trim().slice(0, 60));
  limparCache_();
  return { ok: true, config: cfg, nome: nome_() };
}

function relatorio_(ini, fim) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ini) || !/^\d{4}-\d{2}-\d{2}$/.test(fim) || fim < ini) return { ok: false, erro: 'Período inválido' };
  const peds = pedidosEntre_(ini, fim).map(function (p) {
    return { id: p.id, dia: p.dia, numero: p.numero, hora: p.hora, cliente: p.cliente, telefone: p.telefone, turno: p.turno, tipo: p.tipo,
      status: p.status, pagamento: p.pagamento, subtotal: p.subtotal, taxa: p.taxa, desconto: p.desconto, total: p.total,
      itens: p.itens.map(function (i) { return { nome: i.nome, qtd: i.qtd, preco: i.preco }; }) };
  });
  const desp = ler_('Despesas').map(despesa_).filter(function (d) { return d.dia >= ini && d.dia <= fim; })
    .sort(function (a, b) { return a.dia < b.dia ? 1 : -1; });
  return { ok: true, pedidos: peds, despesas: desp };
}

function enviarFoto_(b) {
  if (!b.base64) throw new Error('Nenhuma imagem recebida.');
  if (String(b.base64).length > 8 * 1024 * 1024) throw new Error('Imagem grande demais.');
  const blob = Utilities.newBlob(Utilities.base64Decode(b.base64), b.mime || 'image/jpeg', 'foto-' + Date.now() + '.jpg');
  const arq = pasta_().createFile(blob);
  arq.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + arq.getId() + '&sz=w900';
}

// Ações que só o dono faz (financeiro e senhas): pedem a senha do dono além da senha do painel
const DONO = {
  admin_relatorio: function (b) { return relatorio_(String(b.p_ini), String(b.p_fim)); },
  admin_despesa: function (b) { return comTrava_(function () { return salvarDespesa_(b.p || {}); }); },
  admin_dono_ok: function () { return { ok: true }; },
  admin_senha: function (b) { return trocarSenha_(b.nova); },
  admin_senha_dono: function (b) {
    const nova = String(b.nova || ''); if (nova.length < 6) return { ok: false, erro: 'A senha precisa ter pelo menos 6 dígitos.' };
    props_().setProperty('pin_dono', nova); return { ok: true };
  }
};
function trocarSenha_(nova) {
  nova = String(nova || '');
  if (nova.length < 6) return { ok: false, erro: 'A senha precisa ter pelo menos 6 dígitos.' };
  props_().setProperty('pin', nova);
  return { ok: true };
}

/* =====================================================================
   PEDIDOS: leitura rápida (só a parte da planilha que interessa)
   ===================================================================== */
function pedidosDoDia_(dia) { return pedidosEntre_(dia, dia); }

function pedidosEntre_(ini, fim) {
  const sh = aba_('Pedidos'), n = sh.getLastRow();
  if (n < 2) return [];
  const colDia = ABAS.Pedidos.indexOf('dia') + 1;
  const dias = sh.getRange(2, colDia, n - 1, 1).getValues().map(function (r) { return dia_(r[0]); });
  let a = -1, b = -1;
  for (let i = dias.length - 1; i >= 0; i--) {
    if (dias[i] > fim) continue;
    if (dias[i] < ini) break;
    if (b < 0) b = i;
    a = i;
  }
  if (a < 0) return [];
  return sh.getRange(a + 2, 1, b - a + 1, ABAS.Pedidos.length).getValues()
    .map(function (r, k) { return pedObj_(r, a + 2 + k); })
    .filter(function (p) { return p.dia >= ini && p.dia <= fim; })
    .map(function (p) { delete p._row; delete p.codigo; return p; });
}

function pedidoPorId_(id) {
  const sh = aba_('Pedidos'), n = sh.getLastRow();
  if (n < 2) return null;
  const ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(id)) return pedObj_(sh.getRange(i + 2, 1, 1, ABAS.Pedidos.length).getValues()[0], i + 2);
  }
  return null;
}

function pedObj_(r, row) {
  const o = {}; ABAS.Pedidos.forEach(function (c, i) { o[c] = r[i]; });
  let itens = []; try { itens = JSON.parse(o.itens || '[]'); } catch (e) {}
  return { _row: row, id: Number(o.id), dia: dia_(o.dia), numero: Number(o.numero) || 0, hora: Number(o.hora) || 0,
    criado_em: iso_(o.criado_em), atualizado_em: iso_(o.atualizado_em), turno: String(o.turno), cliente: String(o.cliente),
    telefone: String(o.telefone), tipo: String(o.tipo), endereco: String(o.endereco), agendado: hm_(o.agendado),
    itens: itens, subtotal: num_(o.subtotal), taxa: num_(o.taxa), desconto: num_(o.desconto), total: num_(o.total),
    pagamento: String(o.pagamento), troco: num_(o.troco), cupom: String(o.cupom || ''), fidelidade: sim_(o.fidelidade, false),
    obs: String(o.obs || ''), status: String(o.status || 'Novo'), codigo: String(o.codigo), manual: String(o.codigo).charAt(0) === 'm' };
}

function contarVendidos_(doDia, turno) {
  return doDia.reduce(function (s, p) {
    if (p.turno !== turno || p.status === 'Cancelado') return s;
    return s + p.itens.reduce(function (t, i) { return t + (i.lim ? i.qtd : 0); }, 0);
  }, 0);
}

function vendidos_(dia) {
  let v = CacheService.getScriptCache().get('vend_' + dia);
  if (v) return JSON.parse(v);
  const doDia = pedidosDoDia_(dia);
  v = { almoco: contarVendidos_(doDia, 'almoco'), noite: contarVendidos_(doDia, 'noite') };
  guardarCache_('vend_' + dia, v, 21600);
  return v;
}

/* =====================================================================
   CLIENTES (fidelidade)
   ===================================================================== */
function cliente_(tel) {
  const c = ler_('Clientes').filter(function (x) { return String(x.telefone) === tel; })[0];
  if (!c) return null;
  return { _row: c._row, telefone: tel, nome: String(c.nome), pedidos: num_(c.pedidos), entregues: num_(c.entregues), desde_premio: num_(c.desde_premio) };
}
function registrarCliente_(c, tel, nome, dia, usouPremio) {
  if (c) atualizarLinha_('Clientes', c._row, { nome: seguro_(nome), pedidos: c.pedidos + 1, ultimo_pedido: dia, desde_premio: usouPremio ? 0 : c.desde_premio });
  else aba_('Clientes').appendRow([tel, seguro_(nome), 1, 0, 0, dia]);
}
function ajustarEntregues_(tel, d, pontos) {
  if (!tel) return;
  const c = cliente_(String(tel)); if (!c) return;
  atualizarLinha_('Clientes', c._row, { entregues: Math.max(c.entregues + d, 0), desde_premio: Math.max(c.desde_premio + d * (pontos || 0), 0) });
}
// Quantos pontos de fidelidade um pedido vale — regras que o dono escolhe em Ajustes
function pontosDoPedido_(cfg, p) {
  const f = cfg.fidelidade || {};
  if (p.fidelidade) return 0;                                        // pedido que usou o prêmio não conta
  const tipos = f.tipos || {};
  if (tipos[p.tipo] === false) return 0;                             // tipo de pedido que o dono não quis contar
  if (num_(p.subtotal) < num_(f.minimo)) return 0;                   // abaixo do valor mínimo
  if (f.modo === 'valor') return Math.floor(num_(p.subtotal) / Math.max(num_(f.valor_ponto) || 10, 1));
  if (f.modo === 'itens') return (p.itens || []).reduce(function (s, i) { return s + (i.lim ? Number(i.qtd) || 0 : 0); }, 0);
  return 1;
}

/* =====================================================================
   UTILITÁRIOS
   ===================================================================== */
function props_() { return PropertiesService.getScriptProperties(); }
function config_() {
  let c = {}; try { c = JSON.parse(props_().getProperty('config') || '{}'); } catch (e) {}
  Object.keys(CONFIG_PADRAO).forEach(function (k) { if (c[k] === undefined) c[k] = JSON.parse(JSON.stringify(CONFIG_PADRAO[k])); });
  return c;
}
function nome_() { return props_().getProperty('nome') || 'Felipe Burguer Drinks'; }
function aba_(nome) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
  if (!sh) throw new Error('Aba "' + nome + '" não encontrada. Rode a função setup.');
  return sh;
}
function ler_(nome) {
  const sh = aba_(nome), n = sh.getLastRow(), cab = ABAS[nome];
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, cab.length).getValues().map(function (r, i) {
    const o = { _row: i + 2 }; cab.forEach(function (c, j) { o[c] = r[j]; }); return o;
  }).filter(function (o) { return o[cab[0]] !== '' && o[cab[0]] !== null; });
}
function atualizarLinha_(nome, row, campos) {
  const sh = aba_(nome), cab = ABAS[nome];
  const atual = sh.getRange(row, 1, 1, cab.length).getValues()[0];
  Object.keys(campos).forEach(function (k) { const i = cab.indexOf(k); if (i >= 0) atual[i] = campos[k]; });
  sh.getRange(row, 1, 1, cab.length).setValues([atual]);
}
function prod_(o) {
  return { id: String(o.id), nome: String(o.nome), descricao: String(o.descricao || ''), preco: num_(o.preco),
    categoria: String(o.categoria || 'Outros'), turnos: lista_(o.turnos), qtd_proteinas: parseInt(o.qtd_proteinas, 10) || 0,
    sabores: lista_(o.sabores), conta_limite: sim_(o.conta_limite, false), foto: String(o.foto || '').split('|')[0], fotos: String(o.foto || '').split('|').filter(Boolean), ativo: sim_(o.ativo, true),
    esgotado: sim_(o.esgotado, false), destaque: sim_(o.destaque, false), ordem: parseInt(o.ordem, 10) || 0, opcao_titulo: String(o.opcao_titulo || ''), sabores_off: lista_(o.sabores_off),
    com_acomp: o.com_acomp === '' || o.com_acomp === undefined ? !/por[cç][aã]o/i.test(String(o.nome) + ' ' + String(o.categoria)) : sim_(o.com_acomp, true),   // porção: só a proteína
    preco_noite: num_(o.preco_noite) > 0 ? num_(o.preco_noite) : '' };
}
function prot_(o) {
  let tur = lista_(o.turnos).filter(function (t) { return t === 'almoco' || t === 'noite'; });
  if (!tur.length) tur = ['almoco', 'noite'];   // proteínas antigas: valem nos dois turnos
  return { id: String(o.id), nome: String(o.nome), turnos: tur,
           almoco: tur.indexOf('almoco') >= 0 && sim_(o.almoco, true), noite: tur.indexOf('noite') >= 0 && sim_(o.noite, true),
           esgotado: sim_(o.esgotado, false), ordem: parseInt(o.ordem, 10) || 0 };
}
function cupom_(o) {
  return { _row: o._row, id: String(o.id), codigo: String(o.codigo).toUpperCase(), tipo: o.tipo === 'pct' ? 'pct' : 'valor',
           valor: num_(o.valor), minimo: num_(o.minimo), ativo: sim_(o.ativo, true), usos: parseInt(o.usos, 10) || 0 };
}
function despesa_(o) { return { id: String(o.id), dia: dia_(o.dia), descricao: String(o.descricao || ''), categoria: String(o.categoria || ''), valor: num_(o.valor) }; }
function porOrdem_(a, b) { return (a.ordem - b.ordem) || String(a.nome).localeCompare(String(b.nome)); }
function lista_(v) {
  if (Array.isArray(v)) return v.map(function (x) { return String(x).trim(); }).filter(String);
  return String(v || '').split(',').map(function (x) { return x.trim(); }).filter(String);
}
function sim_(v, padrao) {
  if (v === undefined || v === null || v === '') return padrao;
  if (v === true || v === false) return v;
  return ['true', 'sim', 's', '1', 'verdadeiro'].indexOf(String(v).trim().toLowerCase()) >= 0;
}
function num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const s = String(v || '').trim().replace(/[^\d,.-]/g, '');
  const n = Number(s.indexOf(',') >= 0 ? s.replace(/\./g, '').replace(',', '.') : s);
  return isFinite(n) ? n : 0;
}
function brl_(v) { return 'R$ ' + num_(v).toFixed(2).replace('.', ','); }
function seguro_(s) { s = String(s == null ? '' : s); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function min_(hm) { const m = /^(\d{1,2}):(\d{2})/.exec(String(hm || '')); return m ? +m[1] * 60 + +m[2] : null; }
function tz_(cfg) { return (cfg && cfg.tz) || 'America/Fortaleza'; }
function hoje_(cfg) {
  const tz = tz_(cfg || config_()), d = new Date();
  return { dia: Utilities.formatDate(d, tz, 'yyyy-MM-dd'), hm: Utilities.formatDate(d, tz, 'HH:mm'),
           iso: Utilities.formatDate(d, tz, "yyyy-MM-dd'T'HH:mm:ss") };
}
function dia_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(config_()), 'yyyy-MM-dd');
  return String(v || '').slice(0, 10);
}
function iso_(v) { return v instanceof Date ? v.toISOString() : String(v || ''); }
function hm_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(config_()), 'HH:mm');
  return String(v || '');
}
function turnoAtivo_(cfg, t) { return !!(cfg.turnos && cfg.turnos[t] && cfg.turnos[t].ativo !== false); }
function turnoEm_(cfg, hm) {
  const m = min_(hm), ts = ['almoco', 'noite'];
  for (let i = 0; i < ts.length; i++) {
    const t = ts[i]; if (!turnoAtivo_(cfg, t)) continue;
    const a = min_(cfg.turnos[t].inicio), b = min_(cfg.turnos[t].fim);
    if (a !== null && b !== null && m >= a && m < b) return t;
  }
  return null;
}
function novoId_() { return 'i' + Utilities.getUuid().replace(/-/g, '').slice(0, 9); }   // começa com letra: a planilha não converte em número
function proximoId_() { const n = (parseInt(props_().getProperty('seq'), 10) || 0) + 1; props_().setProperty('seq', String(n)); return n; }
function versao_() {
  const c = CacheService.getScriptCache().get('versao');
  if (c) return c;
  const v = props_().getProperty('versao') || '0';
  CacheService.getScriptCache().put('versao', v, 21600);
  return v;
}
function subirVersao_() {
  const v = String((parseInt(props_().getProperty('versao'), 10) || 0) + 1);
  props_().setProperty('versao', v);
  CacheService.getScriptCache().put('versao', v, 21600);
}
function guardarCache_(k, v, s) { try { CacheService.getScriptCache().put(k, JSON.stringify(v), s); } catch (e) {} }
function limparCache_() { CacheService.getScriptCache().remove('menu'); }
function pasta_() { const it = DriveApp.getFoldersByName(PASTA_FOTOS); return it.hasNext() ? it.next() : DriveApp.createFolder(PASTA_FOTOS); }
function comTrava_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

    return { doPost, setup };
  })();

  let ultimo = localStorage.getItem(CHAVE);
  function recarregar() {                      // outra aba (cliente ou painel) pode ter mudado os dados
    let t = null; try { t = localStorage.getItem(CHAVE); } catch (e) {}
    if (!t || t === ultimo) return;
    try { const o = JSON.parse(t); ultimo = t;
      Object.keys(abas).forEach(k => delete abas[k]); Object.keys(o.abas || {}).forEach(k => abas[k] = new Sheet(o.abas[k]));
      Object.keys(props).forEach(k => delete props[k]); Object.assign(props, o.props || {});
      Object.keys(cache).forEach(k => delete cache[k]);
    } catch (e) {}
  }
  function guardar() { try { ultimo = JSON.stringify({ abas: Object.fromEntries(Object.entries(abas).map(([k, v]) => [k, v.d])), props }); localStorage.setItem(CHAVE, ultimo); } catch (e) {} }
  if (!props.config) { motor.setup(); guardar(); }

  // logo embutida: abrindo o arquivo direto do computador, o navegador bloqueia a imagem na arte
  window.DEMO_LOGO = 'data:image/png;base64,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';
  window.DEMO = {
    async chamar(acao, args) {
      await new Promise(r => setTimeout(r, 250));
      const corpo = Object.assign({ acao }, args || {});
      if (acao === 'admin_foto') {                       // na demonstração a foto fica no próprio navegador
        if (String(corpo.pin) !== String(props.pin)) return { ok: false, erro: 'PIN inválido' };
        return { ok: true, url: 'data:' + (corpo.mime || 'image/jpeg') + ';base64,' + corpo.base64 };
      }
      recarregar();
      const r = JSON.parse(motor.doPost({ postData: { contents: JSON.stringify(corpo) } }).s);
      guardar();
      return r;
    },
    zerar() { localStorage.removeItem(CHAVE); location.reload(); }
  };
  document.addEventListener('DOMContentLoaded', () => {
    const b = document.createElement('div');
    b.setAttribute('role', 'note');
    b.style.cssText = 'position:relative;z-index:9999;background:#ffcc00;color:#111;font:600 12px Montserrat,system-ui,sans-serif;text-align:center;padding:5px 10px';
    b.innerHTML = 'Modo demonstração: os dados ficam só neste navegador. Senha do painel: 123456. <button style="margin-left:8px;font:inherit;text-decoration:underline;background:none;border:0;cursor:pointer" onclick="DEMO.zerar()">Zerar dados</button>';
    document.body.prepend(b);
  });
})();

window.DEMO_FOTOS = {};
