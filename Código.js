/**
 * ZAYDA CONSTRUTORA — Formulário de materiais no Google Apps Script.
 * Versão 1.1 (28/10/2026): correções marcadas com [ZAYDA 1.1].
 * Versão 1.2 (29/10/2026): confirmação de entrega pelo formulário, marcada com [ZAYDA 1.2].
 * Versão 1.3 (29/10/2026): nome e assinatura de quem recebeu, marcados com [ZAYDA 1.3].
 * Versão 1.4 (29/10/2026): lista de obras para escolher o endereço e ponto de referência, marcados com [ZAYDA 1.4].
 * O restante está igual ao original.
 * Instale este script dentro da planilha fornecida: Extensões > Apps Script.
 * DEPLOY: só permita contas Google da empresa / lista aprovada.
 * Segurança: nenhuma ação aceita usuários sem identidade Google verificável.
 */
const SHEETS = Object.freeze({catalogo:'CATALOGO',pedidos:'PEDIDOS',itens:'ITENS_PEDIDO',auditoria:'AUDITORIA',config:'CONFIG'});
// [ZAYDA 1.2] Aba criada automaticamente na 1ª confirmação de entrega.
const ABA_ENTREGAS = 'ENTREGAS';
const CAB_ENTREGAS = ['ENTREGA_ID','PEDIDO_ID','REVISAO','SKU','PRODUTO','UNIDADE','QTD_ENTREGUE','DATA_ENTREGA','EMAIL_LOGIN','REGISTRADO_EM','OBSERVACOES',
                      'RECEBIDO_POR','ASSINATURA']; // [ZAYDA 1.3] duas últimas colunas
const MAX_ITENS = 200;
// [ZAYDA 1.1] Fuso fixo de São Paulo, independente do fuso configurado no projeto Apps Script.
const FUSO_ZAYDA = 'America/Sao_Paulo';
const PREFIXO_PEDIDO = 'ZAYDA-';

// [ZAYDA 1.1] Linhas de pedido/item reconhecidas pelo prefixo do ID.
// Antes era getValues().slice(1), que escondia o 1º pedido porque PEDIDOS e ITENS_PEDIDO
// não têm cabeçalho. Assim funciona com ou sem cabeçalho na linha 1.
function registros_(key) {
  return s_(key).getDataRange().getValues().filter(r => String(r[0] || '').indexOf(PREFIXO_PEDIDO) === 0);
}

// [ZAYDA 1.1] Data/hora sempre como texto 'yyyy-MM-dd HH:mm:ss'. Antes, String(Date) gravava
// "Wed Oct 28 2026 13:29:10 GMT-0300 (...)" no CRIADO_EM das revisões, com a hora deslocada.
function dataTexto_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, FUSO_ZAYDA, 'yyyy-MM-dd HH:mm:ss');
  return String(v || '');
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Zayda | Solicitação de Materiais')
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
  const rows = s_('config').getDataRange().getValues().slice(1);
  const out={}; rows.forEach(r => { if(r[0])out[String(r[0]).trim()] = String(r[1] || '').trim(); });
  return out;
}

function auth_() {
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  const conf = config_();
  const domain = String(conf.DOMINIO_PERMITIDO || '').replace(/^@/, '').toLowerCase();
  const allow = String(conf.EMAILS_PERMITIDOS || '').toLowerCase().split(/[,;\n]/).map(x=>x.trim()).filter(Boolean);
  if (!email) throw new Error('Não foi possível identificar seu login Google. Faça login com sua conta da empresa e peça ao administrador para conferir as permissões de implantação.');
  if (!(domain && email.endsWith('@' + domain)) && !allow.includes(email)) {
    throw new Error('Acesso não autorizado para ' + email + '. Solicite ao administrador inclusão do email da empresa na aba CONFIG.');
  }
  return {email, empresa: conf.NOME_EMPRESA || 'Zayda Construtora'};
}

function dinheiro_(x, nullable) {
  if (x === '' || x === null || x === undefined) return nullable ? null : 0;
  if (typeof x === 'number') return Number.isFinite(x) ? x : null;
  let s = String(x).replace(/R\$/g,'').trim().replace(/\s/g,'');
  if (s.includes(',')) s = s.replace(/\./g,'').replace(',','.');
  const n=Number(s); return Number.isFinite(n) ? n : null;
}

function catalogo_() {
  const dados=s_('catalogo').getDataRange().getValues().slice(1);
  return dados.filter(r=>r[0] && String(r[11] || '').toUpperCase() !== 'NAO').map(r=>({
    sku:String(r[0]).trim(), produto:String(r[1] || '').trim(), categoria:String(r[2] || 'Outros'),
    subcategoria:String(r[3] || ''), unidade:String(r[4] || 'UN'),
    preco:dinheiro_(r[5],true),statusPreco:String(r[6] || ''),fontePreco:String(r[7] || '')
  }));
}

function getBootstrap() {
  const user=auth_();
  const produtos=catalogo_();
  const cats=[...new Set(produtos.map(x=>x.categoria))].sort((a,b)=>a.localeCompare(b,'pt-BR'));
  return {user,produtos,categorias:cats,recentes:recentes_()};
}

function ultimos_() {
  const rows=registros_('pedidos'), map={}; // [ZAYDA 1.1]
  rows.forEach(r=>{
    const id=String(r[0]||''); const revision=Number(r[1]||0);
    if(id && (!map[id] || revision>map[id].revisao)) map[id]={
      id, revisao:revision,criado:dataTexto_(r[2]),atualizado:dataTexto_(r[3]),email:String(r[4]||''),
      atendente:String(r[5]||''),entregador:String(r[6]||''),endereco:String(r[7]||''),casa:String(r[8]||''),
      observacoes:String(r[9]||''),total:dinheiro_(r[10],false)||0,semPreco:Number(r[11]||0),status:String(r[12]||''),
      referencia:String(r[14]||'') // [ZAYDA 1.4] coluna O de PEDIDOS
    };
  });
  return map;
}

function recentes_() {
  return Object.values(ultimos_()).sort((a,b)=>b.atualizado.localeCompare(a.atualizado)).slice(0,30)
    .map(o=>({id:o.id,atendente:o.atendente,casa:o.casa,total:o.total,semPreco:o.semPreco,atualizado:o.atualizado}));
}

function getPedido(id) {
  auth_();
  const found=ultimos_()[String(id||'').trim()];
  if (!found) throw new Error('Solicitação não encontrada. Confira o código.');
  const rows=registros_('itens'); // [ZAYDA 1.1]
  const itens=rows.filter(r=>String(r[0])===found.id && Number(r[1])===found.revisao).map(r=>({
    sku:String(r[2]||''),nome:String(r[3]||''),categoria:String(r[4]||''),unidade:String(r[5]||''),
    quantidade:Number(r[6]),preco:dinheiro_(r[7],true)
  }));
  const entregue=entreguePorSku_(found.id); // [ZAYDA 1.2]
  itens.forEach(it=>{it.entregue=entregue[it.sku]||0;});
  return {...found,itens};
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
  return sh;
}

/** Quantidade já entregue por SKU de um pedido (soma da aba ENTREGAS). */
function entreguePorSku_(pedidoId) {
  const sh = abaEntregas_(false), out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2,1,sh.getLastRow()-1,7).getValues().forEach(r=>{
    if (String(r[1])===pedidoId) { const sku=String(r[3]); out[sku]=(out[sku]||0)+(Number(r[6])||0); }
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
    const data=String(payload.data||'');
    const hoje=Utilities.formatDate(new Date(),FUSO_ZAYDA,'yyyy-MM-dd');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(data) || data>hoje || data<'2020-01-01') throw new Error('Informe uma data de entrega válida (não pode ser futura).');
    const obs=safeString_(payload.observacoes,300,'Observações da entrega',false);
    // [ZAYDA 1.3] Quem recebeu e assinatura (PNG pequeno em base64, guardado na própria planilha).
    const recebido=safeString_(payload.recebidoPor,120,'Recebido por',true);
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
      linhas.push([entregaId,found.id,found.revisao,sku,p.produto,p.unidade,qtd,data,user.email,now,obs,recebido,assinatura]);
    });
    const ini=sh.getLastRow()+1;
    sh.getRange(ini,1,linhas.length,1).setNumberFormat('@');
    sh.getRange(ini,8,linhas.length,1).setNumberFormat('@');
    sh.getRange(ini,10,linhas.length,1).setNumberFormat('@');
    sh.getRange(ini,1,linhas.length,CAB_ENTREGAS.length).setValues(linhas);
    s_('auditoria').appendRow([now,user.email,'ENTREGA',found.id,found.revisao,
      'Entrega de '+linhas.length+' item(ns) em '+data+', recebida por '+recebido+': '+linhas.map(l=>l[0]).join(', ')]);
    SpreadsheetApp.flush();
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
    const atual=ultimos_();
    const existingId=safeString_(payload.id,50,'ID',false);
    const old=existingId?atual[existingId]:null;
    if(existingId && !old) throw new Error('Solicitação para edição não encontrada.');
    if(old && Number(payload.revisao)!==old.revisao) throw new Error('Este pedido foi alterado por outra pessoa. Reabra-o para não sobrescrever a revisão mais recente.');
    const id=existingId || 'ZAYDA-'+Utilities.formatDate(new Date(),FUSO_ZAYDA,'yyyyMMdd')+'-'+Utilities.getUuid().substring(0,8).toUpperCase();
    const rev=old?old.revisao+1:1;
    const atendente=safeString_(payload.atendente,120,'Nome do(a) atendente',true);
    const entregador=safeString_(payload.entregador,200,'Nome do(s) entregador(es)',true);
    const endereco=safeString_(payload.endereco,300,'Endereço de entrega',true);
    const casa=safeString_(payload.casa,120,'Casa/obra',true);
    const observacoes=safeString_(payload.observacoes,800,'Observações',false);
    const referencia=safeString_(payload.referencia,300,'Ponto de referência',false); // [ZAYDA 1.4]
    if(!Array.isArray(payload.itens) || !payload.itens.length || payload.itens.length>MAX_ITENS) throw new Error('Selecione de 1 a '+MAX_ITENS+' materiais.');
    const catalogo=catalogo_(); const bySku={};catalogo.forEach(p=>bySku[p.sku]=p);
    const seen=new Set(), parsed=[];let total=0,missing=0;
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
      if(preco!==null && (preco<0||preco>10000000))throw new Error('Preço inválido: '+product.produto);
      if(preco===null)missing++;
      const subtotal=preco===null?null:Math.round(qtd*preco*100)/100;
      if(subtotal!==null) total+=subtotal;
      parsed.push({sku,product,qtd,preco,subtotal});
    });
    total=Math.round(total*100)/100;
    const now=Utilities.formatDate(new Date(),FUSO_ZAYDA,'yyyy-MM-dd HH:mm:ss');
    const created=old?old.criado:now;
    const status=missing?'PENDENTE DE COTACAO':'REGISTRADO';
    const pedidoRow=[[id,rev,created,now,user.email,atendente,entregador,endereco,casa,observacoes,total,missing,status,'FORMULARIO WEB',referencia]]; // [ZAYDA 1.4] coluna O: PONTO_REFERENCIA
    const pSheet=s_('pedidos');pSheet.getRange(pSheet.getLastRow()+1,1,1,pedidoRow[0].length).setValues(pedidoRow);
    const itemRows=parsed.map(it=>[id,rev,it.sku,it.product.produto,it.product.categoria,it.product.unidade,it.qtd,it.preco,it.subtotal,it.preco===null?'SIM':'NAO',now]);
    const itSheet=s_('itens');itSheet.getRange(itSheet.getLastRow()+1,1,itemRows.length,itemRows[0].length).setValues(itemRows);
    const aud=s_('auditoria');aud.appendRow([now,user.email,old?'REVISAO':'CRIACAO',id,rev,
      'Registro '+(old?'editado':'criado')+' com '+parsed.length+' itens; total conhecido R$ '+total.toFixed(2)+'; '+missing+' item(ns) a cotar.']);
    SpreadsheetApp.flush();
    return {id,revisao:rev,status,total,semPreco:missing,itens:parsed.length,atualizado:now};
  }finally{lock.releaseLock();}
}

function recentesPublicos() { auth_(); return recentes_(); }

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
