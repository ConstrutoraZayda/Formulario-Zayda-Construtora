/**
 * ZAYDA CONSTRUTORA — Formulário de materiais no Google Apps Script.
 * Versão 1.1 (28/10/2026): correções marcadas com [ZAYDA 1.1].
 * Versão 1.2 (29/10/2026): confirmação de entrega pelo formulário, marcada com [ZAYDA 1.2].
 * Versão 1.3 (29/10/2026): nome e assinatura de quem recebeu, marcados com [ZAYDA 1.3].
 * Versão 1.4 (29/10/2026): lista de obras para escolher o endereço e ponto de referência, marcados com [ZAYDA 1.4].
 * Versão 1.5 (30/10/2026): busca de pedidos em "Últimas solicitações", marcada com [ZAYDA 1.5].
 * Versão 1.6 (05/11/2026): quanto falta entregar ao abrir um pedido, marcado com [ZAYDA 1.6].
 * Versão 1.7 (07/11/2026): pedido entregue travado, última entrega e situação na lista de pedidos, marcados com [ZAYDA 1.7].
 * Versão 1.8 (09/11/2026): pedido aberto na hora por um índice carregado antes, marcado com [ZAYDA 1.8].
 * Versão 1.9 (09/11/2026): pedido salvo não muda mais, chave contra salvamento duplicado, cancelar/encerrar e textos que viravam data, marcados com [ZAYDA 1.9].
 * Versão 2.0 (09/11/2026): quem entregou cada entrega (ENTREGAS coluna N) e a lista de entregas do pedido, marcados com [ZAYDA 2.0].
 * Versão 2.3 (13/11/2026): Outros fornecedores (FORNECEDORES, TERCEIROS, TERCEIROS_PARCELAS e anexos no Drive), marcado com [ZAYDA 2.3].
 * O restante está igual ao original.
 * Instale este script dentro da planilha fornecida: Extensões > Apps Script.
 * DEPLOY: só permita contas Google da empresa / lista aprovada.
 * Segurança: nenhuma ação aceita usuários sem identidade Google verificável.
 */
const SHEETS = Object.freeze({catalogo:'CATALOGO',pedidos:'PEDIDOS',itens:'ITENS_PEDIDO',auditoria:'AUDITORIA',config:'CONFIG'});
// [ZAYDA 1.2] Aba criada automaticamente na 1ª confirmação de entrega.
const ABA_ENTREGAS = 'ENTREGAS';
const CAB_ENTREGAS = ['ENTREGA_ID','PEDIDO_ID','REVISAO','SKU','PRODUTO','UNIDADE','QTD_ENTREGUE','DATA_ENTREGA','EMAIL_LOGIN','REGISTRADO_EM','OBSERVACOES',
                      'RECEBIDO_POR','ASSINATURA', // [ZAYDA 1.3]
                      'ENTREGUE_POR']; // [ZAYDA 2.0] coluna N: quem levou o material
const MAX_ITENS = 200;
// [ZAYDA 1.1] Fuso fixo de São Paulo, independente do fuso configurado no projeto Apps Script.
const FUSO_ZAYDA = 'America/Sao_Paulo';
const PREFIXO_PEDIDO = 'ZAYDA-';

// [ZAYDA 1.1] Linhas de pedido/item reconhecidas pelo prefixo do ID.
// Antes era getValues().slice(1), que escondia o 1º pedido porque PEDIDOS e ITENS_PEDIDO
// não têm cabeçalho. Assim funciona com ou sem cabeçalho na linha 1.
function registros_(key) {
  // [ZAYDA 2.6] Só as colunas usadas (PEDIDOS A:O, ITENS_PEDIDO A:K), até a última linha com dados: o mesmo resultado do getDataRange, com leitura menor.
  const sh = s_(key), n = sh.getLastRow(), cols = Math.min(key === 'itens' ? 11 : 15, sh.getMaxColumns());
  if (!n) return [];
  return sh.getRange(1, 1, n, cols).getValues().filter(r => String(r[0] || '').indexOf(PREFIXO_PEDIDO) === 0);
}

// [ZAYDA 1.1] Data/hora sempre como texto 'yyyy-MM-dd HH:mm:ss'. Antes, String(Date) gravava
// "Wed Oct 28 2026 13:29:10 GMT-0300 (...)" no CRIADO_EM das revisões, com a hora deslocada.
function dataTexto_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, FUSO_ZAYDA, 'yyyy-MM-dd HH:mm:ss');
  return String(v || '');
}

/** [ZAYDA 1.9] Campo digitado que a planilha transformou em hora ou data ("09:00", "10/11"): volta ao texto,
 *  nunca "Sat Dec 30 1899 ...". Hora sozinha a planilha guarda no dia 30/12/1899. */
function textoCelula_(v) {
  if (v instanceof Date && !isNaN(v)) {
    return Utilities.formatDate(v, FUSO_ZAYDA, 'yyyy-MM-dd') === '1899-12-30'
      ? Utilities.formatDate(v, FUSO_ZAYDA, 'HH:mm') : Utilities.formatDate(v, FUSO_ZAYDA, 'dd/MM/yyyy');
  }
  return String(v === null || v === undefined ? '' : v);
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Zayda | Solicitação de Materiais')
    .setFaviconUrl('https://res.cloudinary.com/dovqcebdt/image/upload/v1793988690/favicon-zayda_jiggwc.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function s_(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Abra este projeto por Extensões > Apps Script na planilha Zayda.');
  const sh = ss.getSheetByName(SHEETS[key]);
  if (!sh) throw new Error('Aba obrigatória não encontrada: ' + SHEETS[key]);
  return sh;
}

function config_() {
  const sh = s_('config'), n = sh.getLastRow(); // [ZAYDA 2.6] só as colunas A:B
  const rows = n > 1 ? sh.getRange(2, 1, n - 1, Math.min(2, sh.getMaxColumns())).getValues() : [];
  const out={}; rows.forEach(r => { if(r[0])out[String(r[0]).trim()] = String(r[1] || '').trim(); });
  return out;
}

function auth_() {
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  const conf = confAcesso_(); // [ZAYDA 1.8] lista de acesso guardada por 5 minutos
  const domain = conf.dominio;
  const allow = conf.permitidos;
  if (!email) throw new Error('Não foi possível identificar seu login Google. Faça login com sua conta da empresa e peça ao administrador para conferir as permissões de implantação.');
  const negado = 'Acesso não autorizado para ' + email + '. Solicite ao administrador inclusão do email da empresa na aba CONFIG. Peça ao administrador para cadastrar seu e-mail.';
  // [ZAYDA 2.1] Com a aba USUARIOS preenchida, vale o cadastro dela (FORMULARIO = Sim e ATIVO = Sim); sem ela, CONFIG!EMAILS_PERMITIDOS.
  const cad = usuarios_();
  if (cad.total) {
    const u = cad.porEmail[email];
    if (!u || !u.formulario || !u.ativo) throw new Error(negado);
    return {email, empresa: conf.empresa, nome: u.nome, empresaUsuario: u.empresa};
  }
  if (!(domain && email.endsWith('@' + domain)) && !allow.includes(email)) {
    throw new Error(negado);
  }
  return {email, empresa: conf.empresa};
}

/** [ZAYDA 2.1] Cadastro da aba USUARIOS (só leitura; a aba é protegida e mantida pelo dono).
 *  A EMAIL · B NOME · C USUARIO · D EMPRESA · E PERFIL · F FORMULARIO · G PAINEL · H ATIVO · I OBSERVACAO.
 *  Guardado por 2 minutos: desativar alguém vale rápido. */
function usuarios_() {
  const chave = 'ZAYDA_USUARIOS_V1'; let cache = null;
  try { cache = CacheService.getScriptCache(); const v = cache.get(chave); if (v) return JSON.parse(v); } catch (e) {}
  const sim = x => x === true || String(x || '').trim().normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase() === 'SIM';
  const out = {total: 0, porEmail: {}};
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('USUARIOS');
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues().forEach(r => {
      const em = String(r[0] || '').trim().toLowerCase();
      if (!em) return;
      out.total++;
      if (!out.porEmail[em]) out.porEmail[em] = {nome: String(r[1] || '').trim(), empresa: String(r[3] || '').trim(), formulario: sim(r[5]), ativo: sim(r[7])};
    });
  }
  try { if (cache) cache.put(chave, JSON.stringify(out), 120); } catch (e) {}
  return out;
}

/* ======================= [ZAYDA 2.2] PRESENÇA NO FORMULÁRIO ======================= */
/** Aba técnica _PRESENCA (A EMAIL · B ULTIMO_FORMULARIO, texto yyyy-MM-dd HH:mm:ss), uma linha por e-mail.
 *  Chamada pelo Index.html enquanto o formulário está aberto e visível. Só grava se a última gravação
 *  daquele e-mail tiver mais de 90 s (marcada no cache); sem a trava na hora, apenas ignora. */
const ABA_PRESENCA = '_PRESENCA';
function presencaFormulario() {
  const email = auth_().email, chave = 'ZAYDA_PRESENCA_' + email;
  let cache = null;
  try { cache = CacheService.getScriptCache(); if (cache.get(chave)) return {ok: true, gravado: false}; } catch (e) {}
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(500)) return {ok: true, gravado: false};
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(ABA_PRESENCA);
    if (!sh) {
      sh = ss.insertSheet(ABA_PRESENCA);
      sh.getRange(1, 1, 1, 2).setValues([['EMAIL', 'ULTIMO_FORMULARIO']]).setFontWeight('bold').setBackground('#005f2a').setFontColor('#ffffff');
      sh.setFrozenRows(1);
      sh.protect().setDescription('Gravada pelo formulário (presença). Não editar.').setWarningOnly(true);
    }
    const agora = Utilities.formatDate(new Date(), FUSO_ZAYDA, 'yyyy-MM-dd HH:mm:ss');
    const n = sh.getLastRow() - 1;
    const emails = n > 0 ? sh.getRange(2, 1, n, 1).getDisplayValues().map(r => String(r[0] || '').trim().toLowerCase()) : [];
    const i = emails.indexOf(email);
    sh.getRange(i >= 0 ? i + 2 : sh.getLastRow() + 1, 1, 1, 2).setNumberFormat('@').setValues([[email, agora]]);
    try { if (cache) cache.put(chave, '1', 90); } catch (e) {}
    return {ok: true, gravado: true};
  } finally {
    lock.releaseLock();
  }
}

function dinheiro_(x, nullable) {
  if (x === '' || x === null || x === undefined) return nullable ? null : 0;
  if (typeof x === 'number') return Number.isFinite(x) ? x : null;
  let s = String(x).replace(/R\$/g,'').trim().replace(/\s/g,'');
  if (s.includes(',')) s = s.replace(/\./g,'').replace(',','.');
  const n=Number(s); return Number.isFinite(n) ? n : null;
}

function catalogo_() {
  const sh=s_('catalogo'), n=sh.getLastRow(); // [ZAYDA 2.6] só as colunas A:L
  const dados=n>1?sh.getRange(2,1,n-1,Math.min(12,sh.getMaxColumns())).getValues():[];
  return dados.filter(r=>r[0] && String(r[11] || '').toUpperCase() !== 'NAO').map(r=>({
    sku:String(r[0]).trim(), produto:String(r[1] || '').trim(), categoria:String(r[2] || 'Outros'),
    subcategoria:String(r[3] || ''), unidade:String(r[4] || 'UN'),
    preco:dinheiro_(r[5],true),statusPreco:String(r[6] || ''),fontePreco:String(r[7] || '')
  })).map(p=>{ p.semPreco = p.preco === null || !(p.preco > 0); if (p.semPreco) p.preco = null; return p; }); // [ZAYDA 2.5] preço vazio: semPreco (o null pode chegar como undefined no navegador)
}

/** [ZAYDA 2.6] versaoCatalogo = versão guardada no aparelho. Igual à atual: não manda os produtos (o aparelho usa o que guardou).
 *  Diferente: manda o catálogo inteiro como texto JSON (catalogoJson) e a versão nova. */
function getBootstrap(versaoCatalogo) {
  const user=auth_();
  const cat=catalogoCache_();
  const out={user,catVersao:cat.versao,recentes:recentes_(),popular:maisPedidos_()}; // [ZAYDA 2.7] popular: sku → nº de pedidos
  if (String(versaoCatalogo||'')!==cat.versao) out.catalogoJson=cat.json;
  return out;
}

function ultimos_() {
  const rows=registros_('pedidos'), map={}; // [ZAYDA 1.1]
  rows.forEach(r=>{
    const id=String(r[0]||''); const revision=Number(r[1]||0);
    if(id && (!map[id] || revision>map[id].revisao)) map[id]=registroPedido_(r); // [ZAYDA 1.8] mesmo registro do índice (coluna O = referência)
  });
  return map;
}

function recentes_() {
  const lista=Object.values(ultimos_()).sort((a,b)=>b.atualizado.localeCompare(a.atualizado)).slice(0,200); // [ZAYDA 1.7] era 30: o botão Ver mais mostra de 10 em 10
  const situacao=situacoes_(lista);
  return lista.map(o=>({id:o.id,atendente:o.atendente,casa:o.casa,total:o.total,semPreco:o.semPreco,atualizado:o.atualizado,situacao:situacao[o.id]}));
}

/** [ZAYDA 1.8] Reserva para pedido que não está no índice do navegador (getIndicePedidos). */
function getPedido(id) {
  auth_();
  const found=ultimos_()[String(id||'').trim()];
  if (!found) throw new Error('Solicitação não encontrada. Confira o código.');
  const rows=registros_('itens'); // [ZAYDA 1.1]
  const itens=rows.filter(r=>String(r[0])===found.id && Number(r[1])===found.revisao).map(r=>({
    sku:String(r[2]||''),nome:String(r[3]||''),categoria:String(r[4]||''),unidade:String(r[5]||''),
    quantidade:Number(r[6]),preco:dinheiro_(r[7],true)
  }));
  const ent=resumoEntregas_(blocoAba_(abaEntregas_(false),2,14),found.id,found.entregador); // [ZAYDA 2.0] entregue, última entrega e lista numa leitura só
  itens.forEach(it=>{it.entregue=ent.entregue[it.sku]||0;it.falta=Math.max(0,Math.round((it.quantidade-it.entregue)*1000)/1000);}); // [ZAYDA 1.6] falta
  return {...found,itens,ultimaEntrega:ent.ultima,cancelamento:cancelamentoDe_(found.id),entregas:ent.lista}; // [ZAYDA 1.7] última entrega · [ZAYDA 1.9] cancelado/encerrado
}

/* ======================= [ZAYDA 1.2] CONFIRMAÇÃO DE ENTREGA ======================= */
function abaEntregas_(criar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA_ENTREGAS);
  if (!sh && criar) {
    sh = ss.insertSheet(ABA_ENTREGAS);
    sh.getRange(1,1,1,CAB_ENTREGAS.length).setValues([CAB_ENTREGAS]).setFontWeight('bold')
      .setBackground('#005f2a').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.protect().setDescription('Gravada pelo formulário (confirmação de entrega). Não editar.').setWarningOnly(true);
  }
  // [ZAYDA 1.3] Aba criada na versão 1.2: acrescenta o cabeçalho das colunas novas.
  if (sh && criar && String(sh.getRange(1,12).getValue()||'')!=='RECEBIDO_POR') {
    sh.getRange(1,12,1,2).setValues([['RECEBIDO_POR','ASSINATURA']]).setFontWeight('bold').setBackground('#005f2a').setFontColor('#ffffff');
  }
  // [ZAYDA 2.0] Coluna N (ENTREGUE_POR) nas abas criadas antes desta versão.
  if (sh && criar) {
    if (sh.getMaxColumns() < 14) sh.insertColumnsAfter(sh.getMaxColumns(), 14 - sh.getMaxColumns());
    if (String(sh.getRange(1,14).getValue()||'')!=='ENTREGUE_POR') sh.getRange(1,14).setValue('ENTREGUE_POR').setFontWeight('bold').setBackground('#005f2a').setFontColor('#ffffff');
  }
  return sh;
}

/** Quantidade já entregue por SKU de um pedido (soma da aba ENTREGAS). */
function entreguePorSku_(pedidoId) {
  const sh = abaEntregas_(false), out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2,1,sh.getLastRow()-1,7).getValues().forEach(r=>{
    if (String(r[0]).indexOf('ENT-')===0 && String(r[1])===pedidoId) { const sku=String(r[3]); out[sku]=(out[sku]||0)+(Number(r[6])||0); } // [ZAYDA 1.6] só IDs ENT-
  });
  return out;
}

/** Registra a entrega dos itens escolhidos. payload = {id, revisao, data:'yyyy-MM-dd', observacoes, itens:[{sku, quantidade}]} */
function concluirEntrega(payload) {
  const user=auth_();
  if (!payload || typeof payload !== 'object') throw new Error('Dados da entrega inválidos.');
  const lock=LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Sistema ocupado. Tente novamente.');
  try {
    const id=safeString_(payload.id,50,'Pedido',true);
    const found=ultimos_()[id];
    if (!found) throw new Error('Solicitação não encontrada.');
    if (Number(payload.revisao)!==found.revisao) throw new Error('Este pedido foi alterado depois de aberto. Reabra-o antes de confirmar a entrega.');
    const fechado=cancelamentoDe_(found.id); // [ZAYDA 1.9]
    if (fechado) throw new Error('Este pedido foi '+(fechado.tipo==='CANCELADO'?'cancelado':'encerrado')+' em '+dataBrCurta_(fechado.dataHora)+'.');
    const data=String(payload.data||'');
    const hoje=Utilities.formatDate(new Date(),FUSO_ZAYDA,'yyyy-MM-dd');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(data) || data>hoje || data<'2020-01-01') throw new Error('Informe uma data de entrega válida (não pode ser futura).');
    const obs=safeString_(payload.observacoes,300,'Observações da entrega',false);
    // [ZAYDA 1.3] Quem recebeu e assinatura (PNG pequeno em base64, guardado na própria planilha).
    const recebido=safeString_(payload.recebidoPor,120,'Recebido por',true);
    const entreguePor=safeString_(payload.entreguePor,120,'Entregue por',true); // [ZAYDA 2.0]
    const assinatura=String(payload.assinatura||'');
    // [ZAYDA 1.4] Assinatura na tela deixou de ser obrigatória (agora é no papel); se vier, é conferida.
    if (assinatura && (!/^data:image\/png;base64,[A-Za-z0-9+\/=]+$/.test(assinatura) || assinatura.length<200)) throw new Error('Assinatura inválida.');
    if (assinatura.length>45000) throw new Error('A assinatura ficou grande demais. Limpe e assine novamente.');
    const doPedido={};
    registros_('itens').filter(r=>String(r[0])===found.id && Number(r[1])===found.revisao)
      .forEach(r=>{doPedido[String(r[2])]={produto:String(r[3]||''),unidade:String(r[5]||'')};});
    if (!Array.isArray(payload.itens) || !payload.itens.length) throw new Error('Marque ao menos um item entregue.');
    const vistos=new Set(), linhas=[];
    const sh=abaEntregas_(true);
    const anteriores=sh.getLastRow()>1 ? sh.getRange(2,2,sh.getLastRow()-1,1).getValues().filter(r=>String(r[0])===found.id).length : 0;
    const now=Utilities.formatDate(new Date(),FUSO_ZAYDA,'yyyy-MM-dd HH:mm:ss');
    payload.itens.forEach(it=>{
      const sku=safeString_(it.sku,60,'SKU',true);
      const p=doPedido[sku];
      if (!p) throw new Error('Material fora deste pedido: '+sku);
      if (vistos.has(sku)) throw new Error('Material repetido: '+sku);
      vistos.add(sku);
      const qtd=dinheiro_(it.quantidade,false);
      if (qtd===null||qtd<=0||qtd>100000||Math.abs(Math.round(qtd*1000)-qtd*1000)>1e-7) throw new Error('Quantidade entregue inválida para '+p.produto);
      const seq=anteriores+linhas.length+1;
      const entregaId='ENT-'+found.id.slice(-8)+'-'+('0'+seq).slice(-2);
      linhas.push([entregaId,found.id,found.revisao,sku,p.produto,p.unidade,qtd,data,user.email,now,obs,recebido,assinatura,entreguePor]);
    });
    const ini=sh.getLastRow()+1;
    sh.getRange(ini,1,linhas.length,1).setNumberFormat('@');
    sh.getRange(ini,8,linhas.length,1).setNumberFormat('@');
    sh.getRange(ini,10,linhas.length,1).setNumberFormat('@');
    sh.getRange(ini,11,linhas.length,2).setNumberFormat('@'); // [ZAYDA 1.9] observações e recebido por como texto
    sh.getRange(ini,14,linhas.length,1).setNumberFormat('@'); // [ZAYDA 2.0] entregue por como texto
    sh.getRange(ini,1,linhas.length,CAB_ENTREGAS.length).setValues(linhas);
    s_('auditoria').appendRow([now,user.email,'ENTREGA',found.id,found.revisao,
      'Entrega de '+linhas.length+' item(ns) em '+data+', recebida por '+recebido+': '+linhas.map(l=>l[0]).join(', ')]);
    SpreadsheetApp.flush();
    marcarVersaoPedidos_(); // [ZAYDA 1.8]
    const entregue=entreguePorSku_(found.id);
    const itensPedido=registros_('itens').filter(r=>String(r[0])===found.id && Number(r[1])===found.revisao);
    const pendentes=itensPedido.filter(r=>(entregue[String(r[2])]||0) < Number(r[6])-1e-9).length;
    return {registrados:linhas.length, ids:linhas.map(l=>l[0]), pendentes, entregue};
  } finally { lock.releaseLock(); }
}

function safeString_(x,max,field,required) {
  const s=String(x===null||x===undefined?'':x).trim();
  if(required && !s) throw new Error('Preencha o campo: '+field);
  // Impede que dados digitados sejam interpretados como fórmulas na planilha.
  if (/^[=+@]/.test(s)) throw new Error('O campo '+field+' não pode começar com =, + ou @.');
  if(s.length>max) throw new Error('Campo muito longo: '+field);
  return s;
}

function savePedido(payload) {
  const user=auth_();
  if (!payload || typeof payload !== 'object') throw new Error('Dados do formulário inválidos.');
  const lock=LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Sistema ocupado. Tente salvar novamente.');
  try {
    // [ZAYDA 1.9] Mesma tentativa chegando de novo (clique duplo, reenvio): não grava e devolve o resultado da primeira.
    const chave=/^[A-Za-z0-9-]{8,80}$/.test(String(payload.chaveSalvar||''))?'ZAYDA_SALVAR_'+payload.chaveSalvar:'';
    let cache=null;
    if(chave){ try{ cache=CacheService.getScriptCache(); const feito=cache.get(chave); if(feito) return JSON.parse(feito); }catch(e){ cache=null; } }
    // [ZAYDA 1.9] Só pedido novo: pedido salvo não pode ser alterado (sem revisões).
    const existingId=safeString_(payload.id,50,'ID',false);
    if(existingId){
      if(ultimos_()[existingId]) throw new Error('Este pedido já foi salvo e não pode ser alterado. Faça um novo pedido.');
      throw new Error('Pedido inválido. Faça um novo pedido.');
    }
    const id='ZAYDA-'+Utilities.formatDate(new Date(),FUSO_ZAYDA,'yyyyMMdd')+'-'+Utilities.getUuid().substring(0,8).toUpperCase();
    const rev=1;
    const atendente=safeString_(payload.atendente,120,'Nome do(a) atendente',true);
    const entregador=safeString_(payload.entregador,200,'Nome do(s) entregador(es)',true);
    const endereco=safeString_(payload.endereco,300,'Endereço de entrega',true);
    const casa=safeString_(payload.casa,120,'Casa/obra',true);
    const observacoes=safeString_(payload.observacoes,800,'Observações',false);
    const referencia=safeString_(payload.referencia,300,'Ponto de referência',false); // [ZAYDA 1.4]
    if(!Array.isArray(payload.itens) || !payload.itens.length || payload.itens.length>MAX_ITENS) throw new Error('Selecione de 1 a '+MAX_ITENS+' materiais.');
    const catalogo=catalogo_(); const bySku={};catalogo.forEach(p=>bySku[p.sku]=p);
    const seen=new Set(), parsed=[], semPreco=[];let total=0,missing=0;
    payload.itens.forEach((it,idx)=>{
      const sku=safeString_(it.sku,60,'SKU',true);
      const product=bySku[sku];
      if(!product)throw new Error('Material desativado ou desconhecido: '+sku);
      if(seen.has(sku))throw new Error('Material repetido no mesmo pedido: '+sku);
      seen.add(sku);
      const qtd=dinheiro_(it.quantidade,false);
      if(qtd===null||qtd<=0||qtd>100000||Math.abs(Math.round(qtd*1000)-qtd*1000)>1e-7) throw new Error('Quantidade inválida para '+product.produto);
      // O preço editado é registrado como cotação da solicitação, não altera o catálogo.
      const preco=dinheiro_(it.preco,true);
      if(preco===null||preco<=0){semPreco.push(product.produto);return;}
      if(preco>10000000)throw new Error('Preço inválido: '+product.produto);
      const subtotal=preco===null?null:Math.round(qtd*preco*100)/100;
      if(subtotal!==null) total+=subtotal;
      parsed.push({sku,product,qtd,preco,subtotal});
    });
    if(semPreco.length) throw new Error('Informe o preço dos materiais a cotar antes de salvar: '+semPreco.join(', '));
    total=Math.round(total*100)/100;
    const now=Utilities.formatDate(new Date(),FUSO_ZAYDA,'yyyy-MM-dd HH:mm:ss');
    const created=now;
    const status=missing?'PENDENTE DE COTACAO':'REGISTRADO';
    const pedidoRow=[[id,rev,created,now,user.email,atendente,entregador,endereco,casa,observacoes,total,missing,status,'FORMULARIO WEB',referencia]]; // [ZAYDA 1.4] coluna O: PONTO_REFERENCIA
    const pSheet=s_('pedidos'), linhaP=pSheet.getLastRow()+1;
    pSheet.getRange(linhaP,6,1,5).setNumberFormat('@'); pSheet.getRange(linhaP,15).setNumberFormat('@'); // [ZAYDA 1.9] texto, para "09:00" não virar hora
    pSheet.getRange(linhaP,1,1,pedidoRow[0].length).setValues(pedidoRow);
    const itemRows=parsed.map(it=>[id,rev,it.sku,it.product.produto,it.product.categoria,it.product.unidade,it.qtd,it.preco,it.subtotal,it.preco===null?'SIM':'NAO',now]);
    const itSheet=s_('itens');itSheet.getRange(itSheet.getLastRow()+1,1,itemRows.length,itemRows[0].length).setValues(itemRows);
    const aud=s_('auditoria');aud.appendRow([now,user.email,'CRIACAO',id,rev,
      'Registro criado com '+parsed.length+' itens; total conhecido R$ '+total.toFixed(2)+'; '+missing+' item(ns) a cotar.']);
    SpreadsheetApp.flush();
    marcarVersaoPedidos_(); // [ZAYDA 1.8]
    const resultado={id,revisao:rev,status,total,semPreco:missing,itens:parsed.length,atualizado:now};
    if(cache) { try{ cache.put(chave, JSON.stringify(resultado), 600); }catch(e){} } // [ZAYDA 1.9] 10 minutos
    return resultado;
  }finally{lock.releaseLock();}
}

function recentesPublicos() { auth_(); return recentes_(); }

/** [ZAYDA 1.7] Situação de cada pedido da lista (Entregue, Parcial ou A entregar), só lendo ITENS_PEDIDO e ENTREGAS.
 *  Mesma regra do quadro no topo do pedido aberto. */
function situacoes_(lista) {
  const revisao={}, itens={}, entregue={}, out={};
  lista.forEach(o=>{revisao[o.id]=o.revisao;});
  registros_('itens').forEach(r=>{const id=String(r[0]);if(revisao[id]===Number(r[1]))(itens[id]=itens[id]||[]).push({sku:String(r[2]),qtd:Number(r[6])||0});});
  const sh=abaEntregas_(false);
  if (sh && sh.getLastRow()>1) sh.getRange(2,1,sh.getLastRow()-1,7).getValues().forEach(r=>{
    if (String(r[0]).indexOf('ENT-')!==0) return;
    const k=String(r[1])+'|'+String(r[3]);entregue[k]=(entregue[k]||0)+(Number(r[6])||0);
  });
  const canc=cancelamentos_(); // [ZAYDA 1.9]
  lista.forEach(o=>{
    if (canc[o.id]) { out[o.id]=canc[o.id].tipo==='CANCELADO'?'Cancelado':'Encerrado'; return; }
    const its=itens[o.id]||[], ent=it=>entregue[o.id+'|'+it.sku]||0;
    const falta=its.filter(it=>Math.round((it.qtd-ent(it))*1000)/1000>0).length, algum=its.some(it=>ent(it)>0);
    out[o.id]=its.length&&!falta?'Entregue':algum?'Parcial':'A entregar';
  });
  return out;
}

/** [ZAYDA 1.7] Última entrega registrada do pedido, para o cabeçalho: data da entrega, quando foi registrada e quem recebeu. */
function ultimaEntrega_(pedidoId) {
  const sh=abaEntregas_(false);
  if (!sh || sh.getLastRow()<2) return null;
  let ultima=null;
  sh.getRange(2,1,sh.getLastRow()-1,Math.min(12,sh.getMaxColumns())).getValues().forEach(r=>{
    if (String(r[0]).indexOf('ENT-')!==0 || String(r[1])!==pedidoId) return;
    const registrado=dataTexto_(r[9]);
    if (!ultima || registrado>=ultima.registrado) ultima={data:dataTexto_(r[7]).slice(0,10),registrado,recebidoPor:textoCelula_(r[11])};
  });
  return ultima;
}

/** [ZAYDA 1.5] Busca pedidos (só a revisão mais recente) por parte do código, atendente, entregador,
 *  obra/endereço, casa ou data (dd/mm ou dd/mm/aaaa). Ignora acentos e maiúsculas; todas as palavras
 *  digitadas precisam aparecer. Até 30 resultados, do mais recente para o mais antigo. */
function buscarPedidos(termo) {
  auth_();
  const norm = t => String(t || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
  const partes = norm(termo).trim().split(/\s+/).filter(Boolean);
  if (!partes.length || norm(termo).trim().length < 2) return [];
  const dataBr = t => { const m = /^(\d{4})-(\d{2})-(\d{2})(?: (\d{2}:\d{2}))?/.exec(t); return m ? {data: m[3] + '/' + m[2] + '/' + m[1], hora: m[4] || ''} : {data: t, hora: ''}; };
  return Object.values(ultimos_())
    .map(o => ({o, c: dataBr(o.criado), a: dataBr(o.atualizado)}))
    .filter(x => {
      const texto = norm([x.o.id, x.o.atendente, x.o.entregador, x.o.endereco, x.o.casa, x.c.data, x.a.data].join(' | '));
      return partes.every(p => texto.indexOf(p) >= 0);
    })
    .sort((x, y) => (y.o.criado || '').localeCompare(x.o.criado || '') || (y.o.atualizado || '').localeCompare(x.o.atualizado || ''))
    .slice(0, 30)
    .map(x => ({id: x.o.id, data: x.c.data, hora: x.c.hora, atendente: x.o.atendente,
                obra: x.o.endereco + (x.o.casa ? ' (' + x.o.casa + ')' : ''), total: x.o.total, semPreco: x.o.semPreco}));
}

/** [ZAYDA 1.4] Obras para as sugestões de endereço. A aba OBRAS desta planilha é preenchida
 *  automaticamente pelo Controle de Obras (não editar à mão). */
function listarObras() {
  auth_();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('OBRAS');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues()
    .filter(r => String(r[0] || '').trim() && String(r[1] || '').trim())
    .map(r => ({id: String(r[0]).trim(), endereco: String(r[1]).trim(), cliente: String(r[2] || '').trim(),
                status: String(r[3] || '').trim(), casa: String(r[4] || '').trim()}));
}

/* ==========================================================================
 * ZAYDA | BLOCO DE INTEGRACAO COM O CONTROLE OFICIAL DE OBRAS (28/10/2026)
 * --------------------------------------------------------------------------
 * Cole este bloco NO FINAL do Code.gs atual. Ele so ACRESCENTA funcoes:
 * doGet, savePedido, getPedido, ultimos_, recentes_, catalogo_, auth_ e as
 * demais funcoes existentes continuam exatamente iguais.
 * Nenhuma destas funcoes roda sozinha: execute-as pelo editor (Executar).
 *
 * Ordem dos cabecalhos = ordem exata em que savePedido grava as colunas:
 *   PEDIDOS      -> linha 146 do Code.gs (pedidoRow, 14 colunas A:N)
 *   ITENS_PEDIDO -> linha 148 do Code.gs (itemRows, 11 colunas A:K)
 * ========================================================================== */
const INTEG = Object.freeze({
  cabPedidos: ['PEDIDO_ID','REVISAO','CRIADO_EM','ATUALIZADO_EM','EMAIL_LOGIN','ATENDENTE','ENTREGADOR',
               'ENDERECO','CASA_OBRA','OBSERVACOES','TOTAL_CONHECIDO','ITENS_SEM_PRECO','STATUS_PEDIDO','ORIGEM'],
  cabItens:   ['PEDIDO_ID','REVISAO','SKU','PRODUTO','CATEGORIA','UNIDADE','QUANTIDADE','PRECO_UNITARIO',
               'SUBTOTAL','SEM_PRECO','REGISTRADO_EM'],
  exportPedidos: 'EXPORT_PEDIDOS',
  exportItens: 'EXPORT_ITENS',
  prefixo: 'ZAYDA-'
});

function integTipoLinha1_(valor, cab) {
  const v = String(valor === null || valor === undefined ? '' : valor).trim();
  if (!v) return 'VAZIA';
  if (v === cab[0]) return 'CABECALHO';
  if (v.indexOf(INTEG.prefixo) === 0) return 'REGISTRO';
  return 'DESCONHECIDA';
}

function integChaveValores_(rows) {
  return JSON.stringify(rows.map(r => r.map(v => (v instanceof Date) ? 'D:' + v.getTime() : v)));
}

/** 1) DIAGNOSTICO (somente leitura). Mostra o estado real das abas e o que a integracao vai enxergar. */
function diagnosticarIntegracao() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rel = {executadoEm: Utilities.formatDate(new Date(), FUSO_ZAYDA, 'yyyy-MM-dd HH:mm:ss')};
  const ler = (key, cab) => {
    const sh = s_(key);
    const vals = sh.getLastRow() ? sh.getDataRange().getValues() : [];
    const linha1 = integTipoLinha1_(vals.length ? vals[0][0] : '', cab);
    const dados = vals.filter(r => String(r[0] || '').indexOf(INTEG.prefixo) === 0);
    const fora = vals.map((r, i) => ({linha: i + 1, a: String(r[0] || '')}))
      .filter(x => x.a && x.a !== cab[0] && x.a.indexOf(INTEG.prefixo) !== 0);
    return {aba: sh.getName(), linha1, totalLinhas: vals.length, registros: dados.length,
            invisiveisAoFormulario: linha1 === 'REGISTRO' ? 1 : 0, foraDoPadrao: fora, dados};
  };
  const p = ler('pedidos', INTEG.cabPedidos), it = ler('itens', INTEG.cabItens);
  const chaves = {}, dup = [];
  p.dados.forEach(r => { const k = r[0] + '|R' + Number(r[1]); if (chaves[k]) dup.push(k); chaves[k] = true; });
  const maxRev = {};
  p.dados.forEach(r => { const id = String(r[0]); maxRev[id] = Math.max(maxRev[id] || 0, Number(r[1])); });
  const itensPorChave = {};
  it.dados.forEach(r => { const k = r[0] + '|R' + Number(r[1]); itensPorChave[k] = (itensPorChave[k] || 0) + 1; });
  const itensOrfaos = Object.keys(itensPorChave).filter(k => !chaves[k]);
  const pedidosSemItens = Object.keys(chaves).filter(k => !itensPorChave[k]);
  const itensVigentes = Object.keys(maxRev).reduce((s, id) => s + (itensPorChave[id + '|R' + maxRev[id]] || 0), 0);
  const aud = s_('auditoria').getDataRange().getValues().slice(1)
    .filter(r => ['CRIACAO', 'REVISAO'].indexOf(String(r[2])) >= 0)
    .map(r => String(r[3]) + '|R' + Number(r[4]));
  const ausentes = aud.filter(k => !chaves[k]);
  const exp = {};
  [[INTEG.exportPedidos, p.registros], [INTEG.exportItens, it.registros]].forEach(([nome, esperado]) => {
    const sh = ss.getSheetByName(nome);
    if (!sh) { exp[nome] = 'NAO EXISTE'; return; }
    const v = sh.getDataRange().getValues();
    const n = v.slice(1).filter(r => String(r[0] || '').indexOf(INTEG.prefixo) === 0).length;
    exp[nome] = {cabecalhoOK: String(v[0][0]) === 'PEDIDO_ID', registros: n, confereComOrigem: n === esperado};
  });
  rel.PEDIDOS = {linha1: p.linha1, registros: p.registros, invisiveisAoFormulario: p.invisiveisAoFormulario, foraDoPadrao: p.foraDoPadrao};
  rel.ITENS_PEDIDO = {linha1: it.linha1, registros: it.registros, invisiveisAoFormulario: it.invisiveisAoFormulario, foraDoPadrao: it.foraDoPadrao};
  rel.pedidosDistintos = Object.keys(maxRev).length;
  rel.revisoesDuplicadas = dup;
  rel.itensVigentes = itensVigentes;
  rel.itensSemCabecalhoDePedido = itensOrfaos;
  rel.revisoesSemItens = pedidosSemItens;
  rel.registradosNaAuditoriaMasAusentes = ausentes;
  rel.abasExportacao = exp;
  Logger.log(JSON.stringify(rel, null, 2));
  return rel;
}

/** 2) ABAS DE EXPORTACAO com cabecalho (nao toca PEDIDOS nem ITENS_PEDIDO).
 *  Formula usa colunas inteiras + filtro pelo prefixo do ID: funciona com ou sem
 *  cabecalho nas abas originais e nao se desloca se uma linha for inserida nelas. */
function instalarAbasExportacao() {
  const user = auth_();
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Sistema ocupado. Tente novamente.');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const alvo = [
      [INTEG.exportPedidos, INTEG.cabPedidos, SHEETS.pedidos, 'N'],
      [INTEG.exportItens, INTEG.cabItens, SHEETS.itens, 'K']
    ];
    const feito = [];
    alvo.forEach(([nome, cab, origem, ultCol]) => {
      let sh = ss.getSheetByName(nome);
      if (sh) {
        const a1 = String(sh.getRange(1, 1).getValue() || '');
        if (a1 && a1 !== cab[0]) throw new Error('A aba ' + nome + ' ja existe com outro conteudo. Nada foi alterado.');
      } else {
        sh = ss.insertSheet(nome);
      }
      sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold')
        .setBackground('#10243B').setFontColor('#FFFFFF');
      sh.getRange(2, 1).setFormula('=IFERROR(FILTER(' + origem + '!A:' + ultCol + ',LEFT(' + origem + '!A:A,6)="' +
                                   INTEG.prefixo + '"),"")');
      sh.setFrozenRows(1);
      const prot = sh.protect().setDescription('Exportacao automatica para o Controle Oficial - nao editar');
      prot.setWarningOnly(true);
      feito.push(nome);
    });
    SpreadsheetApp.flush();
    const now = Utilities.formatDate(new Date(), FUSO_ZAYDA, 'yyyy-MM-dd HH:mm:ss');
    s_('auditoria').appendRow([now, user.email, 'INTEGRACAO', '', '', 'Abas de exportacao instaladas: ' + feito.join(', ')]);
    return {instaladas: feito};
  } finally { lock.releaseLock(); }
}

/** 3) CABECALHO NAS ABAS ORIGINAIS (corrige o ponto cego do formulario).
 *  Insere 1 linha no topo de PEDIDOS e ITENS_PEDIDO e escreve o cabecalho.
 *  Nao apaga, nao reordena e nao renumera nada. Idempotente. Confere os dados
 *  depois de inserir e desfaz a insercao se algo nao bater. Ajusta o RESUMO. */
function inserirCabecalhosOriginais() {
  const user = auth_();
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Sistema ocupado. Tente novamente.');
  const rel = [];
  try {
    [['pedidos', INTEG.cabPedidos], ['itens', INTEG.cabItens]].forEach(([key, cab]) => {
      const sh = s_(key);
      const nome = sh.getName();
      if (sh.getLastRow() === 0) {
        sh.getRange(1, 1, 1, cab.length).setValues([cab]);
        integEstiloCabecalho_(sh, cab.length);
        rel.push(nome + ': aba vazia, cabecalho escrito na linha 1');
        return;
      }
      const tipo = integTipoLinha1_(sh.getRange(1, 1).getValue(), cab);
      if (tipo === 'CABECALHO') { rel.push(nome + ': ja possui cabecalho (nada alterado)'); return; }
      if (tipo !== 'REGISTRO') throw new Error(nome + ': linha 1 com conteudo inesperado (' + tipo + '). Nada foi alterado.');
      const antes = sh.getDataRange().getValues();
      const linhasAntes = sh.getLastRow();
      sh.insertRowBefore(1);
      sh.getRange(1, 1, 1, cab.length).setValues([cab]);
      const depois = sh.getRange(2, 1, linhasAntes, antes[0].length).getValues();
      if (sh.getLastRow() !== linhasAntes + 1 || integChaveValores_(depois) !== integChaveValores_(antes)) {
        sh.deleteRow(1);
        throw new Error(nome + ': conferencia falhou apos inserir o cabecalho; insercao desfeita, nada foi alterado.');
      }
      integEstiloCabecalho_(sh, cab.length);
      if (linhasAntes >= 2) sh.getRange(3, 1, 1, cab.length).copyFormatToRange(sh, 1, cab.length, 2, 2);
      rel.push(nome + ': cabecalho inserido; ' + linhasAntes + ' linha(s) preservada(s) e conferida(s)');
    });
    const resumo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RESUMO');
    if (resumo) {
      [['A9', SHEETS.pedidos, 10000], ['D9', SHEETS.itens, 30000]].forEach(([cel, aba, fim]) => {
        const f = resumo.getRange(cel).getFormula();
        const correta = '=COUNTA(' + aba + '!A2:A' + fim + ')';
        if (f && f !== correta && /^=COUNTA\(/.test(f) && f.indexOf(aba + '!') >= 0) {
          resumo.getRange(cel).setFormula(correta);
          rel.push('RESUMO!' + cel + ': ' + f + ' -> ' + correta);
        }
      });
    }
    SpreadsheetApp.flush();
    marcarVersaoPedidos_(); // [ZAYDA 1.8]
    const now = Utilities.formatDate(new Date(), FUSO_ZAYDA, 'yyyy-MM-dd HH:mm:ss');
    s_('auditoria').appendRow([now, user.email, 'INTEGRACAO', '', '', 'Cabecalhos: ' + rel.join(' | ')]);
    Logger.log(rel.join('\n'));
    return rel;
  } finally { lock.releaseLock(); }
}

function integEstiloCabecalho_(sh, n) {
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#10243B').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, n).protect().setDescription('Cabecalho usado pela integracao - nao apagar').setWarningOnly(true);
}

/* ======================= [ZAYDA 1.8] ABRIR PEDIDO NA HORA =======================
 * O formulário carrega em segundo plano um índice com todos os pedidos, no mesmo formato de getPedido,
 * e confere uma versão leve (PEDIDOS_VERSAO) para saber quando recarregá-lo. getPedido continua como reserva. */
const CHAVE_VERSAO_PEDIDOS = 'PEDIDOS_VERSAO';
function registroPedido_(r) {
  return {id:String(r[0]||''), revisao:Number(r[1]||0),criado:dataTexto_(r[2]),atualizado:dataTexto_(r[3]),email:String(r[4]||''),
    atendente:textoCelula_(r[5]),entregador:textoCelula_(r[6]),endereco:textoCelula_(r[7]),casa:textoCelula_(r[8]), // [ZAYDA 1.9] textoCelula_
    observacoes:textoCelula_(r[9]),total:dinheiro_(r[10],false)||0,semPreco:Number(r[11]||0),status:String(r[12]||''),
    referencia:textoCelula_(r[14])};
}
/** Entregue por SKU, última entrega e [ZAYDA 2.0] a lista de entregas, numa passada só (mesmas regras de entreguePorSku_ e ultimaEntrega_).
 *  Entrega antiga sem ENTREGUE_POR (coluna N) usa o entregador do pedido. */
function resumoEntregas_(linhas, pedidoId, entregadorPedido) {
  const entregue = {}, lista = []; let ultima = null;
  linhas.forEach(r => {
    if (String(r[0]).indexOf('ENT-') !== 0 || String(r[1]) !== pedidoId) return;
    const sku = String(r[3]); entregue[sku] = (entregue[sku] || 0) + (Number(r[6]) || 0);
    const registrado = dataTexto_(r[9]);
    if (!ultima || registrado >= ultima.registrado) ultima = {data: dataTexto_(r[7]).slice(0, 10), registrado, recebidoPor: textoCelula_(r[11])};
    lista.push({data: dataTexto_(r[7]).slice(0, 10), sku, material: String(r[4] || ''), quantidade: Number(r[6]) || 0, unidade: String(r[5] || ''),
      recebidoPor: textoCelula_(r[11]), observacao: textoCelula_(r[10]), entreguePor: textoCelula_(r[13]) || String(entregadorPedido || '')});
  });
  return {entregue, ultima, lista};
}
/** Pedido completo a partir das linhas já lidas: o mesmo objeto que getPedido devolve. */
function montarPedido_(found, linhasItens, ent, cancelamento) {
  const itens = linhasItens.filter(r => String(r[0]) === found.id && Number(r[1]) === found.revisao).map(r => ({
    sku:String(r[2]||''),nome:String(r[3]||''),categoria:String(r[4]||''),unidade:String(r[5]||''),
    quantidade:Number(r[6]),preco:dinheiro_(r[7],true)
  }));
  itens.forEach(it=>{it.entregue=ent.entregue[it.sku]||0;it.falta=Math.max(0,Math.round((it.quantidade-it.entregue)*1000)/1000);});
  return {...found, itens, ultimaEntrega: ent.ultima, cancelamento: cancelamento || null, entregas: ent.lista};
}
/** Uma leitura em bloco: da linha ini até a última, com ncols colunas (ou menos, se a aba tiver menos). */
function blocoAba_(sh, ini, ncols) {
  if (!sh) return [];
  const ult = sh.getLastRow();
  if (ult < ini) return [];
  try { return sh.getRange(ini, 1, ult - ini + 1, ncols).getValues(); }
  catch (e) { return sh.getRange(ini, 1, ult - ini + 1, Math.min(ncols, sh.getMaxColumns())).getValues(); }
}
/** Todos os pedidos (revisão mais alta), cada um igual ao que getPedido devolve, e a versão dos dados.
 *  Lê PEDIDOS, ITENS_PEDIDO e ENTREGAS uma vez cada. Devolve texto JSON (mais rápido de transmitir). */
function getIndicePedidos() {
  auth_();
  const versao = versaoPedidos_(); // lida antes dos dados: se algo for gravado durante a leitura, a próxima conferência acusa
  const atual = {};
  blocoAba_(s_('pedidos'), 1, 15).forEach(r => {
    const id = String(r[0] || '');
    if (id.indexOf(PREFIXO_PEDIDO) !== 0) return;
    if (!atual[id] || Number(r[1] || 0) > atual[id].revisao) atual[id] = registroPedido_(r);
  });
  const itensPorId = {};
  blocoAba_(s_('itens'), 1, 8).forEach(r => { const k = String(r[0] || ''); if (atual[k]) (itensPorId[k] = itensPorId[k] || []).push(r); });
  const entPorId = {};
  blocoAba_(abaEntregas_(false), 2, 14).forEach(r => { const k = String(r[1]); if (atual[k]) (entPorId[k] = entPorId[k] || []).push(r); });
  const canc = cancelamentos_();
  const pedidos = Object.keys(atual).map(id => montarPedido_(atual[id], itensPorId[id] || [], resumoEntregas_(entPorId[id] || [], id, atual[id].entregador), canc[id]));
  return JSON.stringify({versao, pedidos});
}
function versaoPedidos_() {
  return String(PropertiesService.getScriptProperties().getProperty(CHAVE_VERSAO_PEDIDOS) || '');
}
/** Versão dos dados dos pedidos (leve): muda a cada gravação em PEDIDOS, ITENS_PEDIDO ou ENTREGAS. */
function getVersaoPedidos() {
  auth_();
  return versaoPedidos_();
}
/** Chamado depois de gravar em PEDIDOS, ITENS_PEDIDO ou ENTREGAS. Uma falha aqui não desfaz a gravação. */
function marcarVersaoPedidos_() {
  try {
    const d = new Date();
    PropertiesService.getScriptProperties().setProperty(CHAVE_VERSAO_PEDIDOS, Utilities.formatDate(d, FUSO_ZAYDA, 'yyyy-MM-dd HH:mm:ss.SSS'));
  } catch (e) { console.error('Versão dos pedidos não atualizada: ' + e.message); }
}
/** Lista de acesso guardada por 5 minutos; a checagem do e-mail continua em toda chamada (auth_). */
function confAcesso_() {
  const chave = 'ZAYDA_ACESSO_V1'; let cache = null;
  try { cache = CacheService.getScriptCache(); const v = cache.get(chave); if (v) return JSON.parse(v); } catch (e) {}
  const conf = config_();
  const v = {dominio: String(conf.DOMINIO_PERMITIDO || '').replace(/^@/, '').toLowerCase(),
    permitidos: String(conf.EMAILS_PERMITIDOS || '').toLowerCase().split(/[,;\n]/).map(x => x.trim()).filter(Boolean),
    empresa: conf.NOME_EMPRESA || 'Zayda Construtora'};
  try { if (cache) cache.put(chave, JSON.stringify(v), 300); } catch (e) {}
  return v;
}

/* ======================= [ZAYDA 1.9] CANCELAR PEDIDO E ENCERRAR O QUE FALTA =======================
 * Registro em CANCELAMENTOS (uma linha por pedido). PEDIDOS, ITENS_PEDIDO e ENTREGAS não são alterados. */
const ABA_CANCEL = 'CANCELAMENTOS';
const CAB_CANCEL = ['PEDIDO_ID','TIPO','DATA_HORA','EMAIL_LOGIN','MOTIVO','ITENS_NAO_ENTREGUES'];
function abaCancelamentos_(criar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA_CANCEL);
  if (!sh && criar) {
    sh = ss.insertSheet(ABA_CANCEL);
    sh.getRange(1,1,1,CAB_CANCEL.length).setValues([CAB_CANCEL]).setFontWeight('bold').setBackground('#005f2a').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.protect().setDescription('Gravada pelo formulário (cancelar e encerrar pedido). Não editar.').setWarningOnly(true);
  }
  return sh;
}
/** Pedido cancelado ou encerrado, por código: {tipo, dataHora, email, motivo, naoEntregues}. Vale a primeira linha de cada pedido. */
function cancelamentos_() {
  const out = {};
  blocoAba_(abaCancelamentos_(false), 2, 6).forEach(r => {
    const id = String(r[0] || '').trim();
    if (!id || out[id]) return;
    out[id] = {tipo: String(r[1] || '').trim().toUpperCase() === 'ENCERRADO' ? 'ENCERRADO' : 'CANCELADO', dataHora: dataTexto_(r[2]),
      email: String(r[3] || ''), motivo: textoCelula_(r[4]), naoEntregues: textoCelula_(r[5])};
  });
  return out;
}
function cancelamentoDe_(id) { return cancelamentos_()[id] || null; }
function dataBrCurta_(t) { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(t || '')); return m ? m[3] + '/' + m[2] + '/' + m[1] : String(t || ''); }
function cancelarPedido(id, motivo, chave) { return fecharPedido_('CANCELADO', id, motivo, chave); }
function encerrarPedido(id, motivo, chave) { return fecharPedido_('ENCERRADO', id, motivo, chave); }
function fecharPedido_(tipo, idPedido, motivoTxt, chaveTxt) {
  const user = auth_();
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Sistema ocupado. Tente novamente.');
  try {
    const chave = /^[A-Za-z0-9-]{8,80}$/.test(String(chaveTxt || '')) ? 'ZAYDA_FECHAR_' + chaveTxt : '';
    let cache = null;
    if (chave) { try { cache = CacheService.getScriptCache(); const feito = cache.get(chave); if (feito) return JSON.parse(feito); } catch (e) { cache = null; } }
    const id = safeString_(idPedido, 50, 'Pedido', true);
    const found = ultimos_()[id];
    if (!found) throw new Error('Solicitação não encontrada.');
    const motivo = safeString_(motivoTxt, 300, 'Motivo', true);
    if (motivo.length < 5) throw new Error('Informe o motivo (mínimo 5 caracteres).');
    const antes = cancelamentoDe_(id);
    if (antes) throw new Error('Este pedido já foi ' + (antes.tipo === 'CANCELADO' ? 'cancelado' : 'encerrado') + ' em ' + dataBrCurta_(antes.dataHora) + '.');
    const entregue = entreguePorSku_(id), algum = Object.keys(entregue).some(k => entregue[k] > 0);
    const falta = registros_('itens').filter(r => String(r[0]) === id && Number(r[1]) === found.revisao)
      .map(r => ({sku: String(r[2] || ''), nome: String(r[3] || ''), un: String(r[5] || ''), q: Math.round((Number(r[6]) - (entregue[String(r[2])] || 0)) * 1000) / 1000}))
      .filter(x => x.q > 0);
    if (tipo === 'CANCELADO' && algum) throw new Error('Este pedido já tem entrega registrada e não pode ser cancelado. Use "Encerrar o que falta".');
    if (tipo === 'ENCERRADO' && (!algum || !falta.length)) throw new Error('Só é possível encerrar um pedido parcial (com entrega registrada e material faltando).');
    const naoEntregues = falta.map(x => x.sku + ' ' + x.nome + ' ' + String(x.q).replace('.', ',') + ' ' + x.un).join('; ');
    const now = Utilities.formatDate(new Date(), FUSO_ZAYDA, 'yyyy-MM-dd HH:mm:ss');
    const sh = abaCancelamentos_(true), linha = sh.getLastRow() + 1;
    sh.getRange(linha, 1, 1, CAB_CANCEL.length).setNumberFormat('@'); // tudo como texto
    sh.getRange(linha, 1, 1, CAB_CANCEL.length).setValues([[id, tipo, now, user.email, motivo, naoEntregues]]);
    s_('auditoria').appendRow([now, user.email, tipo === 'CANCELADO' ? 'CANCELAMENTO' : 'ENCERRAMENTO', id, found.revisao,
      (tipo === 'CANCELADO' ? 'Pedido cancelado' : 'Pedido encerrado; não entregue: ' + naoEntregues) + '. Motivo: ' + motivo]);
    SpreadsheetApp.flush();
    marcarVersaoPedidos_();
    const resultado = {id, tipo, dataHora: now, email: user.email, motivo, naoEntregues};
    if (cache) { try { cache.put(chave, JSON.stringify(resultado), 600); } catch (e) {} }
    return resultado;
  } finally { lock.releaseLock(); }
}

/* ======================= [ZAYDA 2.3] OUTROS FORNECEDORES =======================
 * Compras e serviços fora da loja BSJ (vidraçaria, marmoraria, pisos, concretagem...), com pagamento único ou parcelado.
 * Abas criadas na primeira utilização (o painel lê por posição; campo novo = coluna nova no fim):
 *   FORNECEDORES        A ID · B NOME · C TIPO_PRINCIPAL · D CPF_CNPJ · E TELEFONE · F OBSERVACAO · G ATIVO · H CRIADO_POR · I CRIADO_EM
 *   TERCEIROS           A ID · B CRIADO_EM · C CRIADO_POR · D OBRA_ID · E FORNECEDOR_ID · F FORNECEDOR_NOME · G TIPO · H DESCRICAO · I QUANTIDADE
 *                       J UNIDADE · K VALOR_TOTAL · L DATA_REFERENCIA · M NOTA_NUMERO · N ANEXO_URL · O FORMA_PAGAMENTO · P CONFERENCIA
 *                       Q CONFERIDO_POR · R CONFERIDO_EM · S MOTIVO · T OBSERVACAO · U ATUALIZADO_EM · V ATUALIZADO_POR
 *   TERCEIROS_PARCELAS  A ID_PARCELA · B ID_LANCAMENTO · C NUMERO · D DESCRICAO · E VALOR · F VENCIMENTO · G PAGO · H DATA_PAGAMENTO
 *                       I FORMA · J COMPROVANTE_URL · K REGISTRADO_POR · L REGISTRADO_EM
 * A conferência é feita no painel (EMPRESA Zayda + PERFIL Admin); aqui só se lança, edita enquanto não conferido e registra pagamentos. */
const TER = Object.freeze({
  abaForn: 'FORNECEDORES', abaLanc: 'TERCEIROS', abaParc: 'TERCEIROS_PARCELAS',
  cabForn: ['ID','NOME','TIPO_PRINCIPAL','CPF_CNPJ','TELEFONE','OBSERVACAO','ATIVO','CRIADO_POR','CRIADO_EM'],
  cabLanc: ['ID','CRIADO_EM','CRIADO_POR','OBRA_ID','FORNECEDOR_ID','FORNECEDOR_NOME','TIPO','DESCRICAO','QUANTIDADE','UNIDADE','VALOR_TOTAL',
            'DATA_REFERENCIA','NOTA_NUMERO','ANEXO_URL','FORMA_PAGAMENTO','CONFERENCIA','CONFERIDO_POR','CONFERIDO_EM','MOTIVO','OBSERVACAO',
            'ATUALIZADO_EM','ATUALIZADO_POR'],
  cabParc: ['ID_PARCELA','ID_LANCAMENTO','NUMERO','DESCRICAO','VALOR','VENCIMENTO','PAGO','DATA_PAGAMENTO','FORMA','COMPROVANTE_URL',
            'REGISTRADO_POR','REGISTRADO_EM'],
  tipos: ['Vidros e esquadrias','Pedras e marmoraria','Pisos e revestimentos','Concretagem','Outros'],
  unidades: ['un','m²','m³','m','kg','verba'],
  formas: ['Pix','Boleto','Cartão','Dinheiro','Transferência','Outro'],
  pasta: 'ZAYDA - Outros fornecedores',
  maxArquivo: 10 * 1024 * 1024
});

/** Usuário que pode usar a tela: ATIVO na aba USUARIOS com FORMULARIO = Sim (sem a aba USUARIOS preenchida, ninguém). */
function terUsuario_() {
  const a = auth_();
  const u = usuarios_().porEmail[a.email];
  if (!u || !u.formulario || !u.ativo) throw new Error('Outros fornecedores: seu e-mail precisa estar ativo na aba USUARIOS com FORMULARIO = Sim.');
  const zayda = String(u.empresa || '').normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toUpperCase() === 'ZAYDA';
  return {email: a.email, nome: u.nome, empresa: u.empresa, zayda};
}
/** Para o botão da barra de topo: só aparece para quem pode usar (o servidor confere de novo em cada ação). */
function terceirosAcesso() {
  try { const u = terUsuario_(); return {pode: true, zayda: u.zayda, email: u.email}; } catch (e) { return {pode: false}; }
}

function terAba_(nome, cab, criar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nome);
  if (!sh && criar) {
    sh = ss.insertSheet(nome);
    sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#005f2a').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    // Só aviso: o formulário grava como o usuário que acessa, e a proteção "só o dono" impediria a gravação.
    sh.protect().setDescription('Gravada pelo formulário (Outros fornecedores). Não editar à mão.').setWarningOnly(true);
  }
  return sh;
}
function terLinhas_(nome, cab) {
  const sh = terAba_(nome, cab, false);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, cab.length).getValues();
}
const terAgora_ = () => Utilities.formatDate(new Date(), FUSO_ZAYDA, 'yyyy-MM-dd HH:mm:ss');
const terHoje_ = () => Utilities.formatDate(new Date(), FUSO_ZAYDA, 'yyyy-MM-dd');
const terSemAc_ = t => String(t || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
function terDataTxt_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, FUSO_ZAYDA, 'yyyy-MM-dd');
  return String(v || '').trim().slice(0, 10);
}
function terData_(v, campo, obrig, min, max) {
  const s = String(v || '').trim();
  if (!s) { if (obrig) throw new Error('Informe a data: ' + campo + '.'); return ''; }
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s), d = m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])) : null;
  if (!d || d.getUTCDate() !== +m[3] || d.getUTCMonth() !== +m[2] - 1) throw new Error('Data inválida: ' + campo + '.');
  if ((min && s < min) || (max && s > max)) throw new Error('Data fora do permitido: ' + campo + '.');
  return s;
}
function terSomaDias_(iso, dias) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso), d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3] + dias));
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}
function terValor_(v, campo, max) {
  const n = typeof v === 'number' ? v : dinheiro_(v, true);
  if (n === null || !Number.isFinite(n) || n <= 0 || n > max) throw new Error('Valor inválido: ' + campo + ' (maior que zero).');
  const c = Math.round(n * 100);
  if (Math.abs(c - n * 100) > 1e-6) throw new Error('Valor com mais de 2 casas decimais: ' + campo + '.');
  return c / 100;
}
function terUrl_(v, campo) {
  const s = String(v || '').trim();
  if (!s) return '';
  if (s.length > 300 || !/^https:\/\/(drive|docs)\.google\.com\/[A-Za-z0-9_\-\/?=&.%]+$/.test(s)) throw new Error('Link inválido: ' + campo + '.');
  return s;
}
function terAuditar_(email, acao, id, texto) {
  s_('auditoria').appendRow([terAgora_(), email, acao, id, '', String(texto).slice(0, 2000)]);
}
function terLock_() {
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Sistema ocupado. Tente novamente.');
  return lock;
}
/** Obras da aba OBRAS (a mesma lista do pedido BSJ). */
function terObras_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('OBRAS');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues()
    .filter(r => String(r[0] || '').trim() && String(r[1] || '').trim())
    .map(r => ({id: String(r[0]).trim(), endereco: String(r[1]).trim(), cliente: String(r[2] || '').trim(), status: String(r[3] || '').trim()}));
}
function terFornecedores_() {
  return terLinhas_(TER.abaForn, TER.cabForn).filter(r => /^FOR-\d+$/.test(String(r[0] || '').trim())).map(r => ({
    id: String(r[0]).trim(), nome: textoCelula_(r[1]), tipo: textoCelula_(r[2]), doc: textoCelula_(r[3]), telefone: textoCelula_(r[4]),
    obs: textoCelula_(r[5]), ativo: String(r[6] || '').trim().toUpperCase() !== 'NAO'}));
}
function terParcelaObj_(r) {
  return {id: String(r[0]), lanc: String(r[1]), numero: Number(r[2]) || 0, descricao: textoCelula_(r[3]), valor: Math.round((Number(r[4]) || 0) * 100) / 100,
    vencimento: terDataTxt_(r[5]), pago: String(r[6] || '').trim().toUpperCase() === 'SIM', dataPagamento: terDataTxt_(r[7]), forma: textoCelula_(r[8]),
    comprovante: String(r[9] || ''), registradoPor: String(r[10] || ''), registradoEm: dataTexto_(r[11])};
}
function terLancObj_(r) {
  return {id: String(r[0]), criadoEm: dataTexto_(r[1]), criadoPor: String(r[2] || ''), obra: String(r[3] || ''), fornecedorId: String(r[4] || ''),
    fornecedor: textoCelula_(r[5]), tipo: textoCelula_(r[6]), descricao: textoCelula_(r[7]), quantidade: r[8] === '' ? null : Number(r[8]),
    unidade: textoCelula_(r[9]), valor: Math.round((Number(r[10]) || 0) * 100) / 100, dataRef: terDataTxt_(r[11]), nota: textoCelula_(r[12]),
    anexo: String(r[13] || ''), forma: String(r[14] || ''), conferencia: String(r[15] || 'AGUARDANDO').trim().toUpperCase() || 'AGUARDANDO',
    conferidoPor: String(r[16] || ''), conferidoEm: dataTexto_(r[17]), motivo: textoCelula_(r[18]), obs: textoCelula_(r[19]),
    atualizadoEm: dataTexto_(r[20]), atualizadoPor: String(r[21] || '')};
}

/** Tudo que a tela precisa: obras, fornecedores ativos e os lançamentos com as parcelas (mais recentes primeiro). Texto JSON. */
function terceirosDados() {
  const u = terUsuario_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (![TER.abaForn, TER.abaLanc, TER.abaParc].every(n => ss.getSheetByName(n))) { // 1ª abertura: cria as abas vazias, com cabeçalho
    const lock = terLock_();
    try { terAba_(TER.abaForn, TER.cabForn, true); terAba_(TER.abaLanc, TER.cabLanc, true); terAba_(TER.abaParc, TER.cabParc, true); SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
  }
  const parc = {};
  terLinhas_(TER.abaParc, TER.cabParc).filter(r => String(r[0] || '').indexOf('TER-') === 0).forEach(r => {
    const p = terParcelaObj_(r); (parc[p.lanc] = parc[p.lanc] || []).push(p); });
  const lancs = terLinhas_(TER.abaLanc, TER.cabLanc).filter(r => String(r[0] || '').indexOf('TER-') === 0).map(r => {
    const l = terLancObj_(r); l.parcelas = (parc[l.id] || []).sort((a, b) => a.numero - b.numero); return l; })
    .sort((a, b) => b.criadoEm.localeCompare(a.criadoEm) || b.id.localeCompare(a.id));
  return JSON.stringify({eu: {email: u.email, nome: u.nome, zayda: u.zayda}, hoje: terHoje_(), obras: terObras_(),
    fornecedores: terFornecedores_().filter(f => f.ativo), lancamentos: lancs, tipos: TER.tipos, unidades: TER.unidades, formas: TER.formas});
}

/** Novo fornecedor. dados = {nome, tipo, doc, telefone, obs}. Nome repetido (sem acento/maiúsculas) não entra. */
function salvarFornecedor(dados) {
  const u = terUsuario_();
  if (!dados || typeof dados !== 'object') throw new Error('Dados do fornecedor inválidos.');
  const lock = terLock_();
  try {
    const nome = safeString_(dados.nome, 120, 'Nome do fornecedor', true);
    if (nome.length < 2) throw new Error('Informe o nome do fornecedor.');
    const tipo = safeString_(dados.tipo, 40, 'Tipo principal', false);
    if (tipo && TER.tipos.indexOf(tipo) < 0) throw new Error('Tipo principal inválido.');
    const doc = safeString_(dados.doc, 30, 'CPF/CNPJ', false), dig = doc.replace(/\D/g, '');
    if (doc && dig.length !== 11 && dig.length !== 14) throw new Error('CPF/CNPJ inválido: use 11 (CPF) ou 14 (CNPJ) números.');
    const tel = safeString_(dados.telefone, 30, 'Telefone', false);
    if (tel && !/^[0-9()+\-.\s]{8,30}$/.test(tel)) throw new Error('Telefone inválido.');
    const obs = safeString_(dados.obs, 300, 'Observação do fornecedor', false);
    const todos = terFornecedores_();
    const igual = todos.find(f => terSemAc_(f.nome) === terSemAc_(nome));
    if (igual) throw new Error('Já existe o fornecedor "' + igual.nome + '"' + (igual.ativo ? '. Escolha-o na lista.' : ' (inativo). Peça para reativar na aba FORNECEDORES.'));
    const n = todos.reduce((m, f) => Math.max(m, Number(f.id.replace(/\D/g, '')) || 0), 0) + 1;
    const id = 'FOR-' + ('000' + n).slice(-Math.max(4, String(n).length));
    const sh = terAba_(TER.abaForn, TER.cabForn, true), linha = sh.getLastRow() + 1, agora = terAgora_();
    sh.getRange(linha, 1, 1, TER.cabForn.length).setNumberFormat('@').setValues([[id, nome, tipo, doc, tel, obs, 'SIM', u.email, agora]]);
    terAuditar_(u.email, 'FORNECEDOR', id, 'Fornecedor cadastrado: ' + nome + (tipo ? ' (' + tipo + ')' : ''));
    SpreadsheetApp.flush();
    return {id, nome, tipo, doc, telefone: tel, obs, ativo: true};
  } finally { lock.releaseLock(); }
}

/** Confere e normaliza as parcelas do lançamento. Devolve as linhas da aba TERCEIROS_PARCELAS. */
function terParcelas_(lista, forma, total, idLanc, antigas, u) {
  if (!Array.isArray(lista) || !lista.length) throw new Error('Informe ao menos 1 parcela.');
  if (lista.length > 24) throw new Error('No máximo 24 parcelas.');
  if (forma === 'UNICO' && lista.length !== 1) throw new Error('Pagamento único tem 1 parcela.');
  const hoje = terHoje_(), agora = terAgora_(), porNum = {};
  (antigas || []).forEach(p => { porNum[p.numero] = p; });
  let soma = 0;
  const linhas = lista.map((p, i) => {
    if (!p || typeof p !== 'object') throw new Error('Parcela inválida.');
    const n = i + 1, rot = 'parcela ' + n;
    const desc = safeString_(p.descricao, 80, 'Descrição da ' + rot, false) || (forma === 'UNICO' ? 'Pagamento único' : 'Parcela ' + n);
    const valor = terValor_(p.valor, 'valor da ' + rot, 1e9);
    soma += Math.round(valor * 100);
    const venc = terData_(p.vencimento, 'vencimento da ' + rot, true, '2020-01-01', '2100-12-31');
    let pago = 'NAO', dataPag = '', fp = '', comp = '', regPor = '', regEm = '';
    if (p.pago) {
      pago = 'SIM';
      dataPag = terData_(p.dataPagamento, 'pagamento da ' + rot, true, '2020-01-01', hoje);
      fp = String(p.forma || '');
      if (TER.formas.indexOf(fp) < 0) throw new Error('Escolha a forma de pagamento da ' + rot + '.');
      comp = terUrl_(p.comprovante, 'comprovante da ' + rot);
      const ant = porNum[n];
      // Pagamento que já estava registrado igual continua com quem/quando registrou.
      if (ant && ant.pago && ant.dataPagamento === dataPag && ant.forma === fp && Math.abs(ant.valor - valor) < 0.005) { regPor = ant.registradoPor; regEm = ant.registradoEm; if (!comp) comp = ant.comprovante; }
      else { regPor = u.email; regEm = agora; }
    }
    return [idLanc + '-P' + n, idLanc, n, desc, valor, venc, pago, dataPag, fp, comp, regPor, regEm];
  });
  if (soma !== Math.round(total * 100)) {
    const dif = (Math.round(total * 100) - soma) / 100;
    throw new Error('A soma das parcelas (R$ ' + (soma / 100).toFixed(2).replace('.', ',') + ') não bate com o valor total (R$ ' + total.toFixed(2).replace('.', ',') +
      '). Diferença: R$ ' + dif.toFixed(2).replace('.', ',') + '.');
  }
  return linhas;
}
function terGravarParcelas_(sh, idLanc, linhas) {
  const n = sh.getLastRow() - 1, ids = n > 0 ? sh.getRange(2, 2, n, 1).getValues().map(r => String(r[0])) : [];
  const minhas = []; ids.forEach((v, i) => { if (v === idLanc) minhas.push(i + 2); });
  const fmt = (row, q) => {
    sh.getRange(row, 1, q, 4).setNumberFormat('@'); sh.getRange(row, 3, q, 1).setNumberFormat('0'); sh.getRange(row, 5, q, 1).setNumberFormat('#,##0.00');
    sh.getRange(row, 6, q, 7).setNumberFormat('@'); };
  linhas.forEach((l, i) => { if (i < minhas.length) { fmt(minhas[i], 1); sh.getRange(minhas[i], 1, 1, l.length).setValues([l]); } });
  const novas = linhas.slice(minhas.length);
  if (novas.length) { const ini = sh.getLastRow() + 1; fmt(ini, novas.length); sh.getRange(ini, 1, novas.length, novas[0].length).setValues(novas); }
  minhas.slice(linhas.length).sort((a, b) => b - a).forEach(row => sh.deleteRow(row)); // parcelas que saíram na edição
}

/** Novo lançamento (sem payload.id) ou edição (com payload.id, só AGUARDANDO/REPROVADO, por quem lançou ou Zayda).
 *  Grava TERCEIROS e TERCEIROS_PARCELAS juntos; se algo falhar na conferência, nada é gravado. */
function salvarTerceiro(payload) {
  const u = terUsuario_();
  if (!payload || typeof payload !== 'object') throw new Error('Dados do lançamento inválidos.');
  const lock = terLock_();
  try {
    const chave = /^[A-Za-z0-9-]{8,80}$/.test(String(payload.chaveSalvar || '')) ? 'ZAYDA_TER_' + payload.chaveSalvar : '';
    let cache = null;
    if (chave) { try { cache = CacheService.getScriptCache(); const feito = cache.get(chave); if (feito) return JSON.parse(feito); } catch (e) { cache = null; } }
    const hoje = terHoje_(), agora = terAgora_();
    const obraId = safeString_(payload.obra, 20, 'Obra', true);
    if (!terObras_().some(o => o.id === obraId)) throw new Error('Obra não encontrada na lista. Escolha a obra novamente.');
    const fornId = safeString_(payload.fornecedorId, 20, 'Fornecedor', true);
    const forn = terFornecedores_().find(f => f.id === fornId && f.ativo);
    if (!forn) throw new Error('Fornecedor não encontrado ou inativo. Escolha o fornecedor novamente.');
    const tipo = String(payload.tipo || '');
    if (TER.tipos.indexOf(tipo) < 0) throw new Error('Escolha o tipo.');
    const desc = safeString_(payload.descricao, 300, 'Descrição', true);
    let qtd = '';
    if (payload.quantidade !== '' && payload.quantidade !== null && payload.quantidade !== undefined) {
      qtd = dinheiro_(payload.quantidade, true);
      if (qtd === null || !(qtd > 0) || qtd > 1e7) throw new Error('Quantidade inválida.');
      qtd = Math.round(qtd * 1000) / 1000;
    }
    const un = String(payload.unidade || '');
    if (un && TER.unidades.indexOf(un) < 0) throw new Error('Unidade inválida.');
    const total = terValor_(payload.valor, 'valor total', 1e9);
    const dataRef = terData_(payload.dataRef, 'data de referência', true, '2020-01-01', terSomaDias_(hoje, 366));
    const nota = safeString_(payload.nota, 60, 'Nº da nota/orçamento', false);
    const obs = safeString_(payload.obs, 800, 'Observação', false);
    const anexo = terUrl_(payload.anexo, 'anexo');
    const forma = payload.forma === 'PARCELADO' ? 'PARCELADO' : payload.forma === 'UNICO' ? 'UNICO' : '';
    if (!forma) throw new Error('Escolha a forma de pagamento.');
    const shL = terAba_(TER.abaLanc, TER.cabLanc, true), shP = terAba_(TER.abaParc, TER.cabParc, true);
    const edit = safeString_(payload.id, 40, 'Lançamento', false);
    let id, linha, criadoEm = agora, criadoPor = u.email, motivoAntigo = '', antigas = [], antes = null;
    if (edit) {
      const n = shL.getLastRow() - 1, ids = n > 0 ? shL.getRange(2, 1, n, 1).getValues().map(r => String(r[0])) : [];
      const i = ids.indexOf(edit);
      if (i < 0) throw new Error('Lançamento não encontrado.');
      linha = i + 2; antes = terLancObj_(shL.getRange(linha, 1, 1, TER.cabLanc.length).getValues()[0]);
      if (antes.conferencia === 'CONFERIDO') throw new Error('Este lançamento já foi conferido e não pode ser alterado. Só é possível registrar pagamento.');
      if (antes.conferencia === 'CANCELADO') throw new Error('Este lançamento foi cancelado.');
      if (antes.criadoPor.toLowerCase() !== u.email && !u.zayda) throw new Error('Só quem lançou ou um usuário da Zayda pode editar este lançamento.');
      if (String(payload.atualizadoEm || '') !== antes.atualizadoEm) throw new Error('Este lançamento foi alterado por outra pessoa depois que você abriu. Feche e abra de novo.');
      id = edit; criadoEm = antes.criadoEm; criadoPor = antes.criadoPor; motivoAntigo = antes.motivo;
      antigas = terLinhas_(TER.abaParc, TER.cabParc).filter(r => String(r[1]) === id).map(terParcelaObj_);
    } else {
      const existentes = new Set(shL.getLastRow() > 1 ? shL.getRange(2, 1, shL.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : []);
      const dia = Utilities.formatDate(new Date(), FUSO_ZAYDA, 'yyyyMMdd'), ab = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
      for (let t = 0; t < 50 && (!id || existentes.has(id)); t++) {
        let s = ''; for (let k = 0; k < 4; k++) s += ab.charAt(Math.floor(Math.random() * ab.length)); id = 'TER-' + dia + '-' + s; }
      if (existentes.has(id)) throw new Error('Não foi possível gerar o código. Tente de novo.');
      linha = shL.getLastRow() + 1;
    }
    const parcelas = terParcelas_(payload.parcelas, forma, total, id, antigas, u);
    // REPROVADO editado volta para AGUARDANDO: limpa quem/quando conferiu; o motivo antigo fica na AUDITORIA.
    const row = [id, criadoEm, criadoPor, obraId, forn.id, forn.nome, tipo, desc, qtd, un, total, dataRef, nota, anexo, forma,
      'AGUARDANDO', '', '', '', obs, agora, u.email];
    shL.getRange(linha, 1, 1, 8).setNumberFormat('@'); shL.getRange(linha, 9, 1, 1).setNumberFormat('0.###'); shL.getRange(linha, 10, 1, 1).setNumberFormat('@');
    shL.getRange(linha, 11, 1, 1).setNumberFormat('#,##0.00'); shL.getRange(linha, 12, 1, 11).setNumberFormat('@');
    shL.getRange(linha, 1, 1, row.length).setValues([row]);
    terGravarParcelas_(shP, id, parcelas);
    const resumo = forn.nome + ' · ' + tipo + ' · R$ ' + total.toFixed(2).replace('.', ',') + ' · obra ' + obraId + ' · ' + parcelas.length + ' parcela(s)';
    if (!edit) terAuditar_(u.email, 'TERCEIRO_CRIACAO', id, 'Lançamento criado: ' + resumo);
    else terAuditar_(u.email, 'TERCEIRO_EDICAO', id, 'Lançamento editado' + (antes.conferencia === 'REPROVADO' ? ' (estava REPROVADO por ' + antes.conferidoPor +
      (antes.conferidoEm ? ' em ' + antes.conferidoEm : '') + '; motivo: ' + (motivoAntigo || '—') + '; voltou para AGUARDANDO)' : '') + ': ' + resumo);
    SpreadsheetApp.flush();
    const resultado = {id, novo: !edit, atualizadoEm: agora};
    if (cache) { try { cache.put(chave, JSON.stringify(resultado), 600); } catch (e) {} }
    return resultado;
  } finally { lock.releaseLock(); }
}

/** Parcela e lançamento a partir do ID_PARCELA (dentro da trava). */
function terParcelaLinha_(idParcela) {
  const idp = safeString_(idParcela, 60, 'Parcela', true);
  const shP = terAba_(TER.abaParc, TER.cabParc, false), shL = terAba_(TER.abaLanc, TER.cabLanc, false);
  if (!shP || !shL || shP.getLastRow() < 2) throw new Error('Parcela não encontrada.');
  const ids = shP.getRange(2, 1, shP.getLastRow() - 1, 1).getValues().map(r => String(r[0])), i = ids.indexOf(idp);
  if (i < 0) throw new Error('Parcela não encontrada. Atualize a tela.');
  const linha = i + 2, p = terParcelaObj_(shP.getRange(linha, 1, 1, TER.cabParc.length).getValues()[0]);
  const lids = shL.getLastRow() > 1 ? shL.getRange(2, 1, shL.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [], j = lids.indexOf(p.lanc);
  if (j < 0) throw new Error('Lançamento da parcela não encontrado.');
  return {shP, linha, p, l: terLancObj_(shL.getRange(j + 2, 1, 1, TER.cabLanc.length).getValues()[0])};
}
/** Marca uma parcela como paga. dados = {data:'aaaa-mm-dd', forma, comprovante}. Vale em AGUARDANDO, REPROVADO ou CONFERIDO. */
function registrarPagamento(idParcela, dados) {
  const u = terUsuario_();
  dados = dados && typeof dados === 'object' ? dados : {};
  const lock = terLock_();
  try {
    const x = terParcelaLinha_(idParcela);
    if (['AGUARDANDO', 'REPROVADO', 'CONFERIDO'].indexOf(x.l.conferencia) < 0) throw new Error('Este lançamento foi cancelado: não aceita pagamento.');
    if (x.p.pago) throw new Error('Esta parcela já está paga.');
    const data = terData_(dados.data, 'data do pagamento', true, '2020-01-01', terHoje_());
    const forma = String(dados.forma || '');
    if (TER.formas.indexOf(forma) < 0) throw new Error('Escolha a forma de pagamento.');
    const comp = terUrl_(dados.comprovante, 'comprovante'), agora = terAgora_();
    x.shP.getRange(x.linha, 7, 1, 6).setNumberFormat('@').setValues([['SIM', data, forma, comp, u.email, agora]]);
    terAuditar_(u.email, 'TERCEIRO_PAGAMENTO', x.l.id, 'Parcela ' + x.p.id + ' (' + x.p.descricao + ', R$ ' + x.p.valor.toFixed(2).replace('.', ',') + ') paga em ' + data + ' por ' + forma + (comp ? ' · comprovante anexado' : ''));
    SpreadsheetApp.flush();
    return {ok: true, id: x.p.id, data, forma, comprovante: comp, registradoPor: u.email, registradoEm: agora};
  } finally { lock.releaseLock(); }
}
/** Desfaz o pagamento de uma parcela: só no mesmo dia em que foi registrado, por quem registrou ou usuário da Zayda. */
function desfazerPagamento(idParcela) {
  const u = terUsuario_();
  const lock = terLock_();
  try {
    const x = terParcelaLinha_(idParcela);
    if (!x.p.pago) throw new Error('Esta parcela não está paga.');
    if (x.l.conferencia === 'CANCELADO') throw new Error('Este lançamento foi cancelado.');
    if (String(x.p.registradoEm).slice(0, 10) !== terHoje_()) throw new Error('Só é possível desfazer o pagamento no mesmo dia em que foi registrado.');
    if (x.p.registradoPor.toLowerCase() !== u.email && !u.zayda) throw new Error('Só quem registrou o pagamento ou um usuário da Zayda pode desfazer.');
    x.shP.getRange(x.linha, 7, 1, 6).setNumberFormat('@').setValues([['NAO', '', '', '', '', '']]);
    terAuditar_(u.email, 'TERCEIRO_PAGAMENTO_DESFEITO', x.l.id, 'Pagamento desfeito na parcela ' + x.p.id + ' (registrado por ' + x.p.registradoPor + ' em ' + x.p.registradoEm +
      ', pago em ' + x.p.dataPagamento + ' por ' + x.p.forma + ')');
    SpreadsheetApp.flush();
    return {ok: true, id: x.p.id};
  } finally { lock.releaseLock(); }
}

/** Pasta do Drive para notas, orçamentos e comprovantes. Uma pasta comum (Script Property TER_PASTA_ID), criada pelo primeiro
 *  usuário e compartilhada com os editores da planilha. Se este usuário não conseguir gravar nela, usa uma pasta dele mesmo
 *  com o mesmo nome, também compartilhada com os editores. */
function terPastaPropria_() {
  const up = PropertiesService.getUserProperties();
  let pasta = null;
  try { const id = up.getProperty('TER_PASTA_ID'); if (id) pasta = DriveApp.getFolderById(id); } catch (e) { pasta = null; }
  if (pasta) return pasta;
  pasta = DriveApp.createFolder(TER.pasta);
  try {
    const eu = String(Session.getActiveUser().getEmail() || '').toLowerCase();
    const editores = SpreadsheetApp.getActiveSpreadsheet().getEditors().map(x => String(x.getEmail() || '').toLowerCase()).filter(e => e && e !== eu);
    if (editores.length) pasta.addEditors(editores);
  } catch (e) { console.error('Pasta não compartilhada: ' + e.message); }
  up.setProperty('TER_PASTA_ID', pasta.getId());
  const sp = PropertiesService.getScriptProperties();
  if (!sp.getProperty('TER_PASTA_ID')) sp.setProperty('TER_PASTA_ID', pasta.getId());
  return pasta;
}
/** Envia foto/PDF da nota, orçamento ou comprovante. arq = {nome, tipo (mime), dados (base64)}. Devolve {url, nome}. */
function anexarArquivoTerceiro(arq) {
  const u = terUsuario_();
  if (!arq || typeof arq !== 'object') throw new Error('Arquivo inválido.');
  const tipo = String(arq.tipo || '').toLowerCase();
  const ok = ['image/jpeg', 'image/png', 'image/webp', 'image/heic', 'image/heif', 'application/pdf'];
  if (ok.indexOf(tipo) < 0) throw new Error('Envie uma foto (JPG, PNG, WEBP, HEIC) ou um PDF.');
  const b64 = String(arq.dados || '');
  if (!b64 || !/^[A-Za-z0-9+\/=]+$/.test(b64)) throw new Error('Arquivo inválido.');
  if (b64.length * 3 / 4 > TER.maxArquivo + 3) throw new Error('O arquivo passa de 10 MB.');
  const bytes = Utilities.base64Decode(b64);
  if (!bytes.length || bytes.length > TER.maxArquivo) throw new Error('O arquivo passa de 10 MB.');
  const limpo = String(arq.nome || 'arquivo').replace(/[\\\/:*?"<>|#%\u0000-\u001f]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 80) || 'arquivo';
  const nome = Utilities.formatDate(new Date(), FUSO_ZAYDA, 'yyyy-MM-dd HHmmss') + ' ' + u.email.split('@')[0] + ' - ' + limpo;
  const blob = Utilities.newBlob(bytes, tipo, nome);
  let arquivo = null;
  const comum = PropertiesService.getScriptProperties().getProperty('TER_PASTA_ID');
  if (comum) { try { arquivo = DriveApp.getFolderById(comum).createFile(blob); } catch (e) { arquivo = null; } }
  if (!arquivo) arquivo = terPastaPropria_().createFile(blob);
  terAuditar_(u.email, 'TERCEIRO_ANEXO', '', 'Arquivo enviado para a pasta "' + TER.pasta + '": ' + nome);
  return {url: arquivo.getUrl(), nome: limpo};
}

/* ======================= [ZAYDA 2.6] ESTABILIDADE =======================
 * Catálogo com versão: o navegador guarda {versão, produtos, categorias}; se a versão dele for a atual, getBootstrap não manda os produtos.
 * A versão (MD5 do catálogo montado) e o catálogo ficam no CacheService em pedaços (< 100 KB cada) por até 10 minutos;
 * editar a aba CATALOGO apaga essa cópia (onEdit), e a próxima abertura monta de novo. */
const CAT_CACHE = 'ZAYDA_CATALOGO_V1', CAT_PEDACO = 30000, CAT_SEG = 600;
function catalogoCache_() {
  let cache = null;
  try {
    cache = CacheService.getScriptCache();
    const info = cache.get(CAT_CACHE);
    if (info) {
      const x = JSON.parse(info), chaves = [];
      for (let i = 0; i < x.n; i++) chaves.push(CAT_CACHE + '_' + i);
      const p = cache.getAll(chaves);
      if (chaves.every(k => typeof p[k] === 'string')) return {versao: x.versao, json: chaves.map(k => p[k]).join('')};
    }
  } catch (e) {}
  const produtos = catalogo_().map(p => ({sku: p.sku, produto: p.produto, categoria: p.categoria, unidade: p.unidade, preco: p.preco, semPreco: p.semPreco}));
  const categorias = [...new Set(produtos.map(x => x.categoria))].sort((a, b) => a.localeCompare(b, 'pt-BR'));
  const json = JSON.stringify({produtos, categorias});
  const versao = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, json, Utilities.Charset.UTF_8).map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
  if (cache) {
    try {
      const pedacos = {}, n = Math.ceil(json.length / CAT_PEDACO);
      for (let i = 0; i < n; i++) pedacos[CAT_CACHE + '_' + i] = json.slice(i * CAT_PEDACO, (i + 1) * CAT_PEDACO);
      cache.putAll(pedacos, CAT_SEG);
      cache.put(CAT_CACHE, JSON.stringify({versao, n}), CAT_SEG);
    } catch (e) { console.error('Catálogo não guardado no cache: ' + e.message); }
  }
  return {versao, json};
}
/** Gatilho simples: alteração na aba CATALOGO apaga a cópia do catálogo (a próxima abertura monta a versão nova). */
function onEdit(e) {
  try {
    if (e && e.range && e.range.getSheet().getName() === SHEETS.catalogo) CacheService.getScriptCache().remove(CAT_CACHE);
  } catch (err) {}
}

/* Proteção das abas do sistema: SOMENTE AVISO (quem editar à mão vê um alerta; o formulário continua gravando).
 * Rode protegerAbasSistema pelo editor (Executar). Aba que já tem proteção é deixada como está (não duplica).
 * A aba USUARIOS não é tocada. removerProtecaoAbasSistema desfaz só as proteções criadas aqui. */
const ABAS_SISTEMA = ['PEDIDOS', 'ITENS_PEDIDO', 'ENTREGAS', 'CANCELAMENTOS', 'CATALOGO', 'AUDITORIA', 'CONFIG', 'OBRAS',
  'FORNECEDORES', 'TERCEIROS', 'TERCEIROS_PARCELAS', '_PRESENCA'];
const DESC_PROTECAO_SISTEMA = 'ZAYDA · aba do sistema: edite só se souber o que está fazendo (proteção só de aviso)';
function donoDaPlanilha_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), dono = ss.getOwner() ? ss.getOwner().getEmail().toLowerCase() : '';
  const eu = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (!dono || eu !== dono) throw new Error('Só o dono da planilha pode rodar esta função (pelo editor do Apps Script).');
  return ss;
}
function protegerAbasSistema() {
  const ss = donoDaPlanilha_(), protegidas = [], mantidas = [], ausentes = [];
  ABAS_SISTEMA.forEach(nome => {
    const sh = ss.getSheetByName(nome);
    if (!sh) { ausentes.push(nome); return; }
    const ps = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    if (ps.length) { mantidas.push(nome); return; }
    sh.protect().setDescription(DESC_PROTECAO_SISTEMA).setWarningOnly(true);
    protegidas.push(nome);
  });
  const r = {protegidas, jaTinhamProtecao: mantidas, naoExistem: ausentes};
  Logger.log(JSON.stringify(r, null, 2));
  return r;
}
function removerProtecaoAbasSistema() {
  const ss = donoDaPlanilha_(), removidas = [];
  ABAS_SISTEMA.forEach(nome => {
    const sh = ss.getSheetByName(nome);
    if (!sh) return;
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => { if (p.getDescription() === DESC_PROTECAO_SISTEMA) { p.remove(); removidas.push(nome); } });
  });
  Logger.log('Proteções removidas: ' + (removidas.join(', ') || 'nenhuma'));
  return {removidas};
}

/* [ZAYDA 2.7] Mais pedidos: em quantos pedidos (revisão vigente) cada material aparece. Usado na ordem da busca e nas
 * sugestões "Mais pedidos". Guardado por 10 minutos no CacheService. */
function maisPedidos_() {
  const chave = 'ZAYDA_MAIS_PEDIDOS_V1'; let cache = null;
  try { cache = CacheService.getScriptCache(); const v = cache.get(chave); if (v) return JSON.parse(v); } catch (e) {}
  const rev = {};
  registros_('pedidos').forEach(r => { const id = String(r[0]); rev[id] = Math.max(rev[id] || 0, Number(r[1]) || 0); });
  const vistos = {}, out = {};
  registros_('itens').forEach(r => {
    const id = String(r[0]), sku = String(r[2] || '').trim();
    if (!sku || Number(r[1]) !== rev[id]) return;
    const k = id + '|' + sku; if (vistos[k]) return; vistos[k] = 1;
    out[sku] = (out[sku] || 0) + 1;
  });
  try { const t = JSON.stringify(out); if (cache && t.length < 90000) cache.put(chave, t, 600); } catch (e) {}
  return out;
}

