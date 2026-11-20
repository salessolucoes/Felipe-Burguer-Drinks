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
