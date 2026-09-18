// ============================================================
//   TESTES AUTOMATIZADOS — Controle de Devoluções
//
//   Como usar:
//   1. Abra o editor de scripts (Extensões → Apps Script)
//   2. Selecione a função "executarTodosTestes" e clique em Executar
//   3. Veja o resultado no Logger (Exibir → Logs)
//
//   Os testes usam a planilha real mas gravam e apagam dados numa
//   aba sandbox "_TestSandbox" que é criada e destruída automaticamente.
// ============================================================

var _SANDBOX_ABA = '_TestSandbox';

// ── Runner principal ──────────────────────────────────────────
function executarTodosTestes() {
  var resultados = [];
  var funcs = [
    testeGetSS,
    testeFormatacaoData,
    testeContadorConcluidos,
    testeBaterTermos,
    testeSaudeSistema,
    testeSandboxGravacaoLeitura,
    testeScorecardFornecedores,
    testeSLAFornecedores,
    testeNotificarEventoInativo,
    testeCriarTopicosSemCredenciais,
    testeTgSecretValido,
    testeTgNomeUsuario,
    testeTgPendenteMotivo,
    testeExigirModulo,
    testeGetTodasAbasIncluiExtras,
    testeProcessarAprovacaoGravaLancamento,
    testeSalvarSnapshotConfiguracaoNaoDuplicaHistorico,
    testeExecutarBackupRotaciona,
    testeValidarTipoLancamento,
    testeSalvarCCAlertaFuncionaAdmin,
    testeColunasCorrigidasRelatorios,
    testeAlertaLimiteProtecoesThrottle,
    testeModoManutencaoReflecteEmFlagsAdmin,
    testeAuditoriaAdminRoundTrip,
    testeTrilhaAprovacoesRegistraDecisao,
    testeExecutarBackupComESemRotulo,
    testeAlertasCapacidadeAdmin,
    testeConfiguracaoCompletaRoundTripSemChavesSecretas
  ];
  funcs.forEach(function(fn) {
    try {
      var r = fn();
      resultados.push('[OK] ' + fn.name + (r ? ': ' + r : ''));
    } catch(e) {
      resultados.push('[FAIL] ' + fn.name + ': ' + e.message);
    }
  });
  _limparSandbox();
  var resumo = resultados.join('\n');
  Logger.log('\n═══════ RESULTADO DOS TESTES ═══════\n' + resumo + '\n════════════════════════════════════');
  var ok  = resultados.filter(function(r){ return r.indexOf('[OK]')   === 0; }).length;
  var err = resultados.filter(function(r){ return r.indexOf('[FAIL]') === 0; }).length;
  Logger.log('\nTotal: ' + resultados.length + ' | Passou: ' + ok + ' | Falhou: ' + err);
  if (err > 0) throw new Error(err + ' teste(s) falharam. Veja o log acima.');
  return resumo;
}

// ── Helpers ───────────────────────────────────────────────────
function _assert(cond, msg) { if (!cond) throw new Error(msg || 'Asserção falhou'); }
function _assertEquals(a, b, msg) {
  if (a !== b) throw new Error((msg||'assertEquals') + ': esperado "' + b + '", obtido "' + a + '"');
}
function _assertContains(obj, key, msg) {
  if (obj[key] === undefined) throw new Error((msg||'assertContains') + ': chave "' + key + '" ausente em ' + JSON.stringify(obj));
}

function _criarSandbox() {
  var ss = getSS();
  var ws = ss.getSheetByName(_SANDBOX_ABA);
  if (!ws) ws = ss.insertSheet(_SANDBOX_ABA);
  ws.clearContents();
  return ws;
}
function _limparSandbox() {
  try {
    var ss = getSS();
    var ws = ss.getSheetByName(_SANDBOX_ABA);
    if (ws) ss.deleteSheet(ws);
  } catch(_) {}
}

// ── Testes ────────────────────────────────────────────────────

function testeGetSS() {
  var ss = getSS();
  _assert(ss != null, 'getSS() retornou null');
  _assert(typeof ss.getName === 'function', 'getSS() não é Spreadsheet');
  return ss.getName();
}

function testeFormatacaoData() {
  var tz  = Session.getScriptTimeZone();
  var dt  = new Date(2026, 0, 15); // 15/01/2026
  var fmt = Utilities.formatDate(dt, tz, 'dd/MM/yyyy');
  _assertEquals(fmt, '15/01/2026', 'Formatação de data');
  // Verifica que comparação numérica de timestamps ordena corretamente
  var dt2 = new Date(2025, 11, 31); // 31/12/2025
  _assert(dt.getTime() > dt2.getTime(), 'Timestamp 2026 > 2025');
  return fmt;
}

function testeContadorConcluidos() {
  // Salva estado atual
  var props = PropertiesService.getScriptProperties();
  var valorAntes = props.getProperty(_PROP_KEY_CONCLUIDOS) || '0';
  try {
    props.setProperty(_PROP_KEY_CONCLUIDOS, '5');
    _assertEquals(_lerContadorConcluidos(), 5, '_lerContadorConcluidos');
    _incrementarContadorConcluidos();
    _assertEquals(_lerContadorConcluidos(), 6, 'Incrementar contador');
    _decrementarContadorConcluidos(2);
    _assertEquals(_lerContadorConcluidos(), 4, 'Decrementar contador');
    _zerarContadorConcluidos();
    _assertEquals(_lerContadorConcluidos(), 0, 'Zerar contador');
  } finally {
    props.setProperty(_PROP_KEY_CONCLUIDOS, valorAntes);
  }
  return 'ok';
}

function testeBaterTermos() {
  var r1 = _baterTermos(['123', '456'], '123', '999');
  _assert(r1.bate, '_baterTermos deve bater por NFD');
  _assertEquals(r1.termoBateu, '123', 'termo batido deve ser 123');

  var r2 = _baterTermos(['777'], '123', '777');
  _assert(r2.bate, '_baterTermos deve bater por NF');

  var r3 = _baterTermos(['AAA'], '123', '456');
  _assert(!r3.bate, '_baterTermos não deve bater');
  return 'ok';
}

function testeSaudeSistema() {
  var resp = verificarSaudeSistema();
  var d    = JSON.parse(resp);
  _assert(!d.erro, 'verificarSaudeSistema retornou erro: ' + d.erro);
  _assertContains(d, 'status', 'campo status');
  _assertContains(d, 'checks', 'campo checks');
  _assert(Array.isArray(d.checks), 'checks deve ser array');
  _assert(d.checks.length > 0, 'checks não pode ser vazio');
  return 'status=' + d.status + ', checks=' + d.checks.length;
}

function testeSandboxGravacaoLeitura() {
  var ws  = _criarSandbox();
  var val = 'TESTE_' + Date.now();
  ws.getRange(1,1).setValue(val);
  SpreadsheetApp.flush();
  var lido = ws.getRange(1,1).getValue();
  _assertEquals(lido, val, 'Leitura do sandbox');
  return 'sandbox ok';
}

function testeScorecardFornecedores() {
  var resp = obterScorecardFornecedores();
  var d    = JSON.parse(resp);
  _assert(!d.erro, 'obterScorecardFornecedores erro: ' + (d.erro||''));
  _assertContains(d, 'fornecedores', 'campo fornecedores');
  _assert(Array.isArray(d.fornecedores), 'fornecedores deve ser array');
  return 'fornecedores=' + d.fornecedores.length;
}

function testeSLAFornecedores() {
  var resp = obterSLAFornecedores();
  var d    = JSON.parse(resp);
  _assert(!d.erro, 'obterSLAFornecedores erro: ' + (d.erro||''));
  _assertContains(d, 'sla', 'campo sla');
  _assert(Array.isArray(d.sla), 'sla deve ser array');
  return 'fornecedores_com_sla=' + d.sla.length;
}

function testeNotificarEventoInativo() {
  var chaveAnterior = PropertiesService.getScriptProperties().getProperty(_KEY_WEBHOOK_CONF);
  try {
    PropertiesService.getScriptProperties().setProperty(_KEY_WEBHOOK_CONF, JSON.stringify({ ativo: false }));
    var r = notificarEvento('sistema', 'teste');
    _assertEquals(r, null, 'notificarEvento deve retornar null quando config está inativa');
    return 'ok';
  } finally {
    if (chaveAnterior === null) PropertiesService.getScriptProperties().deleteProperty(_KEY_WEBHOOK_CONF);
    else PropertiesService.getScriptProperties().setProperty(_KEY_WEBHOOK_CONF, chaveAnterior);
  }
}

function testeCriarTopicosSemCredenciais() {
  var r = JSON.parse(criarTopicosWebhook({ telegram: { token: '', chatId: '' } }));
  _assertContains(r, 'erro', 'sem token/chatId deve retornar erro');
  return 'ok';
}

function testeTgSecretValido() {
  _assertEquals(_tgSecretValido({ webhookSecret: 'abc' }, 'abc'), true,  'secret correto deve validar');
  _assertEquals(_tgSecretValido({ webhookSecret: 'abc' }, 'xyz'), false, 'secret errado não deve validar');
  _assertEquals(_tgSecretValido({}, ''), false, 'sem secret configurado não deve validar');
  return 'ok';
}

function testeTgNomeUsuario() {
  _assertEquals(_tgNomeUsuario({ username: 'joaosilva' }), '@joaosilva', 'deve preferir username');
  _assertEquals(_tgNomeUsuario({ first_name: 'João' }), 'João', 'sem username usa first_name');
  _assertEquals(_tgNomeUsuario(null), 'alguém', 'sem from usa fallback');
  return 'ok';
}

function testeTgPendenteMotivo() {
  var chaveAnterior = PropertiesService.getScriptProperties().getProperty(_KEY_APROV_AGUARDANDO_MOTIVO);
  try {
    PropertiesService.getScriptProperties().deleteProperty(_KEY_APROV_AGUARDANDO_MOTIVO);
    _tgSalvarPendenteMotivo('ap_1', '-100123', 555);
    var achado = _tgAcharPendenteMotivo(555);
    _assert(achado != null, 'deve achar pendente pelo messageId');
    _assertEquals(achado.item.aprovacaoId, 'ap_1', 'aprovacaoId deve bater');

    var naoAchado = _tgAcharPendenteMotivo(999);
    _assertEquals(naoAchado, null, 'messageId sem correspondência deve retornar null');

    _tgRemoverPendenteMotivo(achado.lista, achado.idx);
    _assertEquals(_tgAcharPendenteMotivo(555), null, 'após remover, não deve mais achar');
    return 'ok';
  } finally {
    if (chaveAnterior === null) PropertiesService.getScriptProperties().deleteProperty(_KEY_APROV_AGUARDANDO_MOTIVO);
    else PropertiesService.getScriptProperties().setProperty(_KEY_APROV_AGUARDANDO_MOTIVO, chaveAnterior);
  }
}

function testeExigirModulo() {
  // _exigirModulo aceita um 3º parâmetro (_emailParaTeste) só para testes:
  // quando fornecido (mesmo string vazia), ele é usado no lugar de
  // Session.getActiveUser().getEmail() — isso permite simular tanto um
  // usuário específico quanto o contexto "sem usuário ativo" (trigger).
  var CARGO_TESTE = '_TesteCargoExigirModulo';
  var EMAIL_TESTE = 'teste.exigirmodulo@example.com';
  var idCargoCriado = null;

  try {
    // Garante que não existe cargo/vínculo de um teste anterior que falhou no meio
    try { removerUsuarioCargo(EMAIL_TESTE); } catch(_) {}

    // 1) Sem cargo nenhum vinculado ao e-mail → deve bloquear
    var r1 = _exigirModulo('lancamento', true, EMAIL_TESTE);
    _assert(!r1.ok, 'sem cargo vinculado, _exigirModulo deveria bloquear');
    _assertContains(JSON.parse(r1.resp), 'erro', 'resposta de bloqueio deve ter campo erro');

    // 2) Cria cargo com o módulo "lancamento" (leitura/escrita) e vincula o e-mail
    var respCargo = JSON.parse(salvarCargo(null, CARGO_TESTE, ['lancamento'], false));
    _assert(!respCargo.erro, 'salvarCargo falhou: ' + (respCargo.erro || ''));
    idCargoCriado = respCargo.id;
    salvarUsuarioCargo(EMAIL_TESTE, idCargoCriado);

    var r2 = _exigirModulo('lancamento', true, EMAIL_TESTE);
    _assert(r2.ok, 'com cargo autorizado para "lancamento", _exigirModulo deveria liberar');

    // 3) Módulo fora do cargo → deve bloquear
    var r3 = _exigirModulo('backup', true, EMAIL_TESTE);
    _assert(!r3.ok, 'módulo "backup" fora do cargo deveria bloquear');

    // 4) Sem e-mail ativo (contexto de trigger/tempo) → não deve bloquear
    var r4 = _exigirModulo('backup', true, '');
    _assert(r4.ok, 'sem usuário ativo (trigger), _exigirModulo não deve bloquear');

    return 'ok';
  } finally {
    try { removerUsuarioCargo(EMAIL_TESTE); } catch(_) {}
    if (idCargoCriado) { try { excluirCargo(idCargoCriado); } catch(_) {} }
  }
}

function testeGetTodasAbasIncluiExtras() {
  // Regressão: várias funções (busca, e-mail, backup, relatórios etc.) usavam
  // a lista fixa ABAS_OPERACIONAIS em vez de _getTodasAbas() e por isso
  // ignoravam abas de fornecedores criadas depois (via "Adicionar Novo
  // Fornecedor"). Este teste garante que _getTodasAbas() sempre inclui as
  // abas extras salvas em PropertiesService, sem duplicar as fixas.
  var props     = PropertiesService.getScriptProperties();
  var chaveRaw  = 'cdv_abas_extras';
  var valorAntes = props.getProperty(chaveRaw);
  var ABA_FAKE  = '_TesteAbaExtraFake';

  try {
    props.setProperty(chaveRaw, JSON.stringify([ABA_FAKE]));

    var todas = _getTodasAbas();
    _assert(todas.indexOf(ABA_FAKE) !== -1, '_getTodasAbas() deveria incluir a aba extra "' + ABA_FAKE + '"');
    ABAS_OPERACIONAIS.forEach(function(nome) {
      _assert(todas.indexOf(nome) !== -1, '_getTodasAbas() deveria incluir a aba fixa "' + nome + '"');
    });
    _assertEquals(todas.length, ABAS_OPERACIONAIS.length + 1, 'total de abas (fixas + extra) sem duplicar');

    // Extra que coincide com uma fixa não deve duplicar
    props.setProperty(chaveRaw, JSON.stringify([ABAS_OPERACIONAIS[0], ABA_FAKE]));
    var todas2 = _getTodasAbas();
    _assertEquals(todas2.length, ABAS_OPERACIONAIS.length + 1, 'extra duplicando nome fixo não deve gerar entrada repetida');

    return 'ok';
  } finally {
    if (valorAntes === null) props.deleteProperty(chaveRaw);
    else props.setProperty(chaveRaw, valorAntes);
  }
}

// ── Testes de regressão — correções desta rodada ───────────────

function testeProcessarAprovacaoGravaLancamento() {
  // Regressão CRÍTICA: _processarAprovacaoInterno chamava
  // _gravarLancamento(item.dados) com aridade errada (a função espera
  // ss, ws, dados, ulPre, nfsPre) — "dados" chegava undefined e a gravação
  // sempre falhava com TypeError, DEPOIS que o item já tinha sido removido
  // da fila de pendentes — ou seja, toda aprovação perdia o lançamento pra
  // sempre (tanto via Configurações quanto via botão do Telegram). Este
  // teste simula uma aprovação completa contra a aba sandbox e confere que
  // a linha realmente aparece gravada, e que a fila fica vazia.
  var ws = _criarSandbox();
  var chaveAnterior = PropertiesService.getScriptProperties().getProperty(_KEY_APROVACOES_PEND);
  var id = 'ap_teste_' + Date.now();
  try {
    var dadosTeste = {
      abaSelecao: _SANDBOX_ABA,
      nf: 'NFTESTE' + Date.now(),
      nfd: '',
      fornecedor: 'Fornecedor Teste',
      tipo: 'Avaria',
      motivo: 'Teste automatizado',
      descricao: 'Item de teste',
      qtd: 2,
      valorUnit: 10
    };
    var item = { id: id, dados: dadosTeste, usuario: 'teste@example.com', ts: new Date().toISOString(), status: 'pendente' };
    PropertiesService.getScriptProperties().setProperty(_KEY_APROVACOES_PEND, JSON.stringify([item]));

    var resp = JSON.parse(_processarAprovacaoInterno(id, true, '', 'teste'));
    _assert(!resp.erro, '_processarAprovacaoInterno retornou erro: ' + (resp.erro || ''));

    var nfGravado = ws.getRange(LINHA_DADOS, COL_NF).getValue();
    _assertEquals(String(nfGravado), dadosTeste.nf, 'NF do lançamento aprovado deveria estar gravada na aba (bug fazia isso falhar sempre)');

    var pendRestante = JSON.parse(PropertiesService.getScriptProperties().getProperty(_KEY_APROVACOES_PEND) || '[]');
    _assertEquals(pendRestante.length, 0, 'fila de aprovações pendentes deveria estar vazia após aprovar com sucesso');

    return 'ok';
  } finally {
    if (chaveAnterior === null) PropertiesService.getScriptProperties().deleteProperty(_KEY_APROVACOES_PEND);
    else PropertiesService.getScriptProperties().setProperty(_KEY_APROVACOES_PEND, chaveAnterior);
    _limparSandbox();
  }
}

function testeSalvarSnapshotConfiguracaoNaoDuplicaHistorico() {
  // Regressão: salvarSnapshotConfiguracao incluía a própria chave
  // _KEY_CONFIG_HISTORICO dentro do snapshot salvo (getProperties() traz
  // todas as propriedades), fazendo cada novo snapshot embutir o histórico
  // inteiro dos anteriores — crescimento exponencial até estourar o limite
  // de tamanho do PropertiesService. Confere que o snapshot salvo NÃO
  // contém a própria chave de histórico dentro de si.
  var chaveAnterior = PropertiesService.getScriptProperties().getProperty(_KEY_CONFIG_HISTORICO);
  try {
    PropertiesService.getScriptProperties().deleteProperty(_KEY_CONFIG_HISTORICO);
    var r1 = JSON.parse(salvarSnapshotConfiguracao('teste 1'));
    _assert(!r1.erro, 'primeiro snapshot falhou: ' + (r1.erro || ''));
    var r2 = JSON.parse(salvarSnapshotConfiguracao('teste 2'));
    _assert(!r2.erro, 'segundo snapshot falhou: ' + (r2.erro || ''));

    var hist = JSON.parse(PropertiesService.getScriptProperties().getProperty(_KEY_CONFIG_HISTORICO) || '[]');
    _assert(hist.length >= 2, 'deveria ter ao menos 2 snapshots no histórico');
    _assert(!hist[0].snapshot.hasOwnProperty(_KEY_CONFIG_HISTORICO), 'snapshot não deveria conter a própria chave de histórico dentro de si (causa do crescimento exponencial)');
    _assert(!hist[1].snapshot.hasOwnProperty(_KEY_CONFIG_HISTORICO), 'snapshot anterior também não deveria conter a própria chave de histórico');

    return 'ok';
  } finally {
    if (chaveAnterior === null) PropertiesService.getScriptProperties().deleteProperty(_KEY_CONFIG_HISTORICO);
    else PropertiesService.getScriptProperties().setProperty(_KEY_CONFIG_HISTORICO, chaveAnterior);
  }
}

function testeExecutarBackupRotaciona() {
  // Funcionalidade nova: antes existia só 1 backup (aba fixa sobrescrita a
  // cada execução) — um backup ruim ou uma restauração no meio apagava a
  // única cópia de segurança. Agora mantém um histórico rotativo de até
  // BACKUPS_MANTIDOS backups. Roda o backup 2x e confere que o histórico
  // cresce e que dá pra listar via listarHistoricoBackups().
  var chaveAnterior = PropertiesService.getScriptProperties().getProperty(_KEY_BACKUP_HISTORICO);
  try {
    PropertiesService.getScriptProperties().deleteProperty(_KEY_BACKUP_HISTORICO);
    var r1 = JSON.parse(executarBackup());
    _assert(!r1.erro, 'primeiro backup falhou: ' + (r1.erro || ''));
    Utilities.sleep(1100); // garante nome de aba com timestamp diferente do 2º backup
    var r2 = JSON.parse(executarBackup());
    _assert(!r2.erro, 'segundo backup falhou: ' + (r2.erro || ''));

    var lista = JSON.parse(listarHistoricoBackups());
    _assert(!lista.erro, 'listarHistoricoBackups retornou erro: ' + (lista.erro || ''));
    _assert(lista.historico.length >= 2, 'histórico deveria ter ao menos 2 entradas após 2 backups');
    _assertEquals(lista.mantidos, BACKUPS_MANTIDOS, 'mantidos deveria refletir BACKUPS_MANTIDOS');

    return 'ok';
  } finally {
    // limpa as abas de backup criadas por este teste
    var hist = [];
    try { hist = JSON.parse(PropertiesService.getScriptProperties().getProperty(_KEY_BACKUP_HISTORICO) || '[]'); } catch (_) {}
    var ss = getSS();
    hist.forEach(function(h) {
      try { var wsB = ss.getSheetByName(h.sheetName); if (wsB) ss.deleteSheet(wsB); } catch (_) {}
    });
    if (chaveAnterior === null) PropertiesService.getScriptProperties().deleteProperty(_KEY_BACKUP_HISTORICO);
    else PropertiesService.getScriptProperties().setProperty(_KEY_BACKUP_HISTORICO, chaveAnterior);
  }
}

function testeValidarTipoLancamento() {
  // Regressão: o campo "Tipo" do lançamento não era validado no servidor —
  // um valor fora da lista podia ser gravado direto via google.script.run,
  // contornando a tela (a validação da própria planilha usa
  // setAllowInvalid(true), ou seja, é só um aviso visual).
  var dadosValidos = { abaSelecao: _SANDBOX_ABA, nf: '1', descricao: 'x', qtd: 1, valorUnit: 1, tipo: 'Avaria' };
  _validarDadosForm(dadosValidos); // não deve lançar

  var dadosInvalidos = { abaSelecao: _SANDBOX_ABA, nf: '1', descricao: 'x', qtd: 1, valorUnit: 1, tipo: 'TipoQueNaoExiste' };
  var lancou = false;
  try { _validarDadosForm(dadosInvalidos); } catch (e) { lancou = true; }
  _assert(lancou, '_validarDadosForm deveria rejeitar um tipo fora de TIPOS_DEVOLUCAO');

  var semTipo = { abaSelecao: _SANDBOX_ABA, nf: '1', descricao: 'x', qtd: 1, valorUnit: 1 };
  _validarDadosForm(semTipo); // tipo ausente/vazio deve continuar tolerado (campo opcional)
  return 'ok';
}

function testeSalvarCCAlertaFuncionaAdmin() {
  // Regressão: obterCCAlerta/salvarCCAlerta chamavam _usuarioEhAdmin(true)
  // sem usar o valor de retorno — um no-op que não bloqueava ninguém. A
  // correção troca isso por "if (!_usuarioEhAdmin()) return erro". Como o
  // teste roda como o dono/admin da planilha, confere que o caminho normal
  // continua funcionando e que a normalização de chave (trim + lowercase)
  // não quebra o ciclo salvar → ler.
  var chaveAnterior = PropertiesService.getScriptProperties().getProperty(_KEY_CC_ALERTA);
  try {
    var rSalvar = JSON.parse(salvarCCAlerta('TesteTipo', 'cc@example.com', 'bcc@example.com'));
    _assert(!rSalvar.erro, 'salvarCCAlerta (admin) não deveria falhar: ' + (rSalvar.erro || ''));

    var rLer = JSON.parse(obterCCAlerta());
    _assert(!rLer.erro, 'obterCCAlerta (admin) não deveria falhar: ' + (rLer.erro || ''));
    _assertContains(rLer.ccAlerta, 'testetipo', 'chave deveria estar normalizada em minúsculo');
    _assertEquals(rLer.ccAlerta.testetipo.cc, 'cc@example.com', 'cc salvo deveria bater');

    return 'ok';
  } finally {
    salvarCCAlerta('TesteTipo', '', ''); // remove a entrada de teste
    if (chaveAnterior === null) PropertiesService.getScriptProperties().deleteProperty(_KEY_CC_ALERTA);
    else PropertiesService.getScriptProperties().setProperty(_KEY_CC_ALERTA, chaveAnterior);
  }
}

function testeColunasCorrigidasRelatorios() {
  // Regressão: obterComparativoPeriodos, obterConquistasMes e
  // previewEmailDevolucao liam colunas erradas e/ou a partir da linha 2 em
  // vez de LINHA_DADOS — pegavam linha de cabeçalho como dado, liam motivo
  // como se fosse tipo, valor unitário como se fosse valor total (e
  // multiplicavam os dois), etc. Registra 1 linha sintética na aba sandbox
  // (temporariamente tratada como aba extra) e confere que as 3 funções
  // leem os campos certos.
  var ws = _criarSandbox();
  var props = PropertiesService.getScriptProperties();
  var chaveExtras = 'cdv_abas_extras';
  var extrasAntes = props.getProperty(chaveExtras);
  try {
    var listaExtras = [];
    try { listaExtras = JSON.parse(extrasAntes || '[]'); } catch (_) {}
    if (listaExtras.indexOf(_SANDBOX_ABA) === -1) listaExtras.push(_SANDBOX_ABA);
    props.setProperty(chaveExtras, JSON.stringify(listaExtras));

    var nfdTeste = 'NFDTESTE' + Date.now();
    var hoje = new Date();
    var linha = new Array(TOTAL_COLUNAS).fill('');
    linha[IDX_NFD] = nfdTeste; linha[IDX_NF] = 'NF' + Date.now(); linha[IDX_DATA] = hoje;
    linha[IDX_FORN] = 'Fornecedor Teste'; linha[IDX_TIPO] = 'Avaria';
    linha[IDX_MOTIVO] = 'Motivo Teste'; linha[IDX_DESC] = 'Descrição do item teste';
    linha[IDX_QTD] = 4; linha[IDX_VL_UNIT] = 25; linha[IDX_VL_TOT] = 100;
    linha[IDX_STATUS] = 'Devolvido';
    ws.getRange(LINHA_DADOS, 1, 1, TOTAL_COLUNAS).setValues([linha]);
    SpreadsheetApp.flush();

    var hojeStr = Utilities.formatDate(hoje, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var comp = JSON.parse(obterComparativoPeriodos(hojeStr, hojeStr, hojeStr, hojeStr));
    _assert(!comp.erro, 'obterComparativoPeriodos erro: ' + (comp.erro || ''));
    _assertEquals(comp.a.qtd, 1, 'comparativo deveria contar o item sintético no período de hoje');
    _assertEquals(comp.a.valor, 100, 'valor do comparativo deveria ser o Vl Total (100), não vlUnit×vlTot');
    _assertEquals(comp.a.devolvidos, 1, 'status "Devolvido" deveria ser contado (lido da coluna certa)');

    var conq = JSON.parse(obterConquistasMes());
    _assert(!conq.erro, 'obterConquistasMes erro: ' + (conq.erro || ''));
    _assert(conq.total >= 1, 'obterConquistasMes deveria contar ao menos o item lançado hoje');

    var prev = JSON.parse(previewEmailDevolucao({ nfds: [nfdTeste], assunto: 'Teste' }));
    _assert(prev.html.indexOf('Fornecedor Teste') !== -1, 'preview deveria trazer o fornecedor correto (antes lia a coluna do Tipo)');
    _assert(prev.html.indexOf('Descrição do item teste') !== -1, 'preview deveria trazer a descrição correta');

    return 'ok';
  } finally {
    if (extrasAntes === null) props.deleteProperty(chaveExtras);
    else props.setProperty(chaveExtras, extrasAntes);
    _limparSandbox();
  }
}

function testeAlertaLimiteProtecoesThrottle() {
  // Regressão: ao atingir LIMITE_PROTECOES, o sistema só gravava uma linha
  // de log que ninguém olha em tempo real — nenhum alerta ativo. Confere
  // que o alerta é registrado no painel de erros e que chamadas repetidas
  // dentro da janela de cache (throttle de 12h) não duplicam o registro.
  var cache = CacheService.getScriptCache();
  var key = 'alerta_limite_protecoes';
  cache.remove(key);
  var errosAntes = JSON.parse(PropertiesService.getScriptProperties().getProperty(_KEY_ERROS_RECENTES) || '[]').length;
  try {
    _alertarLimiteProtecoes(380);
    var errosDepois1 = JSON.parse(PropertiesService.getScriptProperties().getProperty(_KEY_ERROS_RECENTES) || '[]').length;
    _assertEquals(errosDepois1, errosAntes + 1, 'primeira chamada deveria registrar 1 erro novo no painel');

    _alertarLimiteProtecoes(380); // deve ser ignorada pelo throttle
    var errosDepois2 = JSON.parse(PropertiesService.getScriptProperties().getProperty(_KEY_ERROS_RECENTES) || '[]').length;
    _assertEquals(errosDepois2, errosDepois1, 'segunda chamada dentro da janela de cache não deveria registrar de novo');

    return 'ok';
  } finally {
    cache.remove(key);
  }
}

// ── Testes — Painel Admin (16 funcionalidades novas) ───────────

function testeModoManutencaoReflecteEmFlagsAdmin() {
  // Novo (Painel Admin, item 4): salvarModoManutencao deve persistir a flag
  // e obterFlagsAdmin (consolidado com somente-leitura e aprovação) deve
  // refletir o valor atual.
  var chaveAnterior = PropertiesService.getScriptProperties().getProperty(_KEY_MODO_MANUTENCAO);
  try {
    var r1 = JSON.parse(salvarModoManutencao(true));
    _assert(!r1.erro, 'salvarModoManutencao(true) falhou: ' + (r1.erro || ''));
    var f1 = JSON.parse(obterFlagsAdmin());
    _assert(!f1.erro, 'obterFlagsAdmin erro: ' + (f1.erro || ''));
    _assertEquals(f1.modoManutencao, true, 'obterFlagsAdmin deveria refletir modo manutenção ativado');

    var r2 = JSON.parse(salvarModoManutencao(false));
    _assert(!r2.erro, 'salvarModoManutencao(false) falhou: ' + (r2.erro || ''));
    var f2 = JSON.parse(obterFlagsAdmin());
    _assertEquals(f2.modoManutencao, false, 'obterFlagsAdmin deveria refletir modo manutenção desativado');

    return 'ok';
  } finally {
    if (chaveAnterior === null) PropertiesService.getScriptProperties().deleteProperty(_KEY_MODO_MANUTENCAO);
    else PropertiesService.getScriptProperties().setProperty(_KEY_MODO_MANUTENCAO, chaveAnterior);
  }
}

function testeAuditoriaAdminRoundTrip() {
  // Novo (Painel Admin, item 2): _registrarAuditoriaAdmin grava e
  // obterAuditoriaAdmin/limparAuditoriaAdmin leem e limpam corretamente.
  var chaveAnterior = PropertiesService.getScriptProperties().getProperty(_KEY_AUDITORIA_ADMIN);
  try {
    PropertiesService.getScriptProperties().deleteProperty(_KEY_AUDITORIA_ADMIN);
    _registrarAuditoriaAdmin('teste_acao', 'detalhe do teste');

    var r = JSON.parse(obterAuditoriaAdmin());
    _assert(!r.erro, 'obterAuditoriaAdmin erro: ' + (r.erro || ''));
    _assertEquals(r.registros.length, 1, 'deveria ter exatamente 1 registro após 1 chamada');
    _assertEquals(r.registros[0].acao, 'teste_acao', 'ação registrada deveria bater');
    _assertEquals(r.registros[0].detalhes, 'detalhe do teste', 'detalhes registrados deveriam bater');

    var rLimpar = JSON.parse(limparAuditoriaAdmin());
    _assert(!rLimpar.erro, 'limparAuditoriaAdmin erro: ' + (rLimpar.erro || ''));
    var rDepois = JSON.parse(obterAuditoriaAdmin());
    _assertEquals(rDepois.registros.length, 0, 'trilha de auditoria deveria estar vazia após limpar');

    return 'ok';
  } finally {
    if (chaveAnterior === null) PropertiesService.getScriptProperties().deleteProperty(_KEY_AUDITORIA_ADMIN);
    else PropertiesService.getScriptProperties().setProperty(_KEY_AUDITORIA_ADMIN, chaveAnterior);
  }
}

function testeTrilhaAprovacoesRegistraDecisao() {
  // Novo (Painel Admin, item 15): toda aprovação/reprovação processada por
  // _processarAprovacaoInterno precisa deixar um registro na trilha de
  // aprovações (_KEY_TRILHA_APROVACOES), além de gravar o lançamento
  // (já coberto por testeProcessarAprovacaoGravaLancamento).
  var ws = _criarSandbox();
  var chaveAprovPend = PropertiesService.getScriptProperties().getProperty(_KEY_APROVACOES_PEND);
  var chaveTrilha    = PropertiesService.getScriptProperties().getProperty(_KEY_TRILHA_APROVACOES);
  var id = 'ap_trilha_' + Date.now();
  try {
    var dadosTeste = {
      abaSelecao: _SANDBOX_ABA,
      nf: 'NFTRILHA' + Date.now(),
      nfd: '',
      fornecedor: 'Fornecedor Trilha Teste',
      tipo: 'Avaria',
      motivo: 'Teste automatizado',
      descricao: 'Item de teste',
      qtd: 1,
      valorUnit: 5
    };
    var item = { id: id, dados: dadosTeste, usuario: 'trilha-teste@example.com', ts: new Date().toISOString(), status: 'pendente' };
    PropertiesService.getScriptProperties().setProperty(_KEY_APROVACOES_PEND, JSON.stringify([item]));
    PropertiesService.getScriptProperties().deleteProperty(_KEY_TRILHA_APROVACOES);

    var resp = JSON.parse(_processarAprovacaoInterno(id, true, '', 'revisor.teste'));
    _assert(!resp.erro, '_processarAprovacaoInterno retornou erro: ' + (resp.erro || ''));

    var trilha = JSON.parse(obterTrilhaAprovacoesAdmin());
    _assert(!trilha.erro, 'obterTrilhaAprovacoesAdmin erro: ' + (trilha.erro || ''));
    _assert(trilha.registros.length >= 1, 'trilha deveria ter ao menos 1 registro após a aprovação');
    var reg = trilha.registros[0];
    _assertEquals(reg.decisao, 'aprovado', 'decisão registrada deveria ser "aprovado"');
    _assertEquals(reg.nf, dadosTeste.nf, 'NF na trilha deveria bater com a do lançamento aprovado');
    _assertEquals(reg.revisor, 'revisor.teste', 'revisor na trilha deveria bater com quem aprovou');

    return 'ok';
  } finally {
    if (chaveAprovPend === null) PropertiesService.getScriptProperties().deleteProperty(_KEY_APROVACOES_PEND);
    else PropertiesService.getScriptProperties().setProperty(_KEY_APROVACOES_PEND, chaveAprovPend);
    if (chaveTrilha === null) PropertiesService.getScriptProperties().deleteProperty(_KEY_TRILHA_APROVACOES);
    else PropertiesService.getScriptProperties().setProperty(_KEY_TRILHA_APROVACOES, chaveTrilha);
    _limparSandbox();
  }
}

function testeExecutarBackupComESemRotulo() {
  // Novo (Painel Admin, item 12): executarBackup(rotulo) precisa continuar
  // funcionando sem argumento — todo chamador existente (FormBackup.html,
  // reaplicarCoresTodas, _executarBackupAutomatico) invoca assim — e também
  // aceitar uma etiqueta opcional, salvando-a no histórico.
  var chaveAnterior = PropertiesService.getScriptProperties().getProperty(_KEY_BACKUP_HISTORICO);
  try {
    PropertiesService.getScriptProperties().deleteProperty(_KEY_BACKUP_HISTORICO);

    var rSemRotulo = JSON.parse(executarBackup());
    _assert(!rSemRotulo.erro, 'executarBackup() sem argumento falhou: ' + (rSemRotulo.erro || ''));

    Utilities.sleep(1100); // garante nome de aba com timestamp diferente do 2º backup
    var rComRotulo = JSON.parse(executarBackup('backup de teste'));
    _assert(!rComRotulo.erro, 'executarBackup(rotulo) falhou: ' + (rComRotulo.erro || ''));
    _assert(String(rComRotulo.sucesso).indexOf('backup de teste') !== -1, 'mensagem de sucesso deveria mencionar a etiqueta');

    // listarHistoricoBackups() não repassa o campo "rotulo" no mapeamento
    // público (só indice/ts/totalLinhas/resumo/legado) — confere direto na
    // property, que é onde o rótulo realmente fica salvo.
    var histBruto = JSON.parse(PropertiesService.getScriptProperties().getProperty(_KEY_BACKUP_HISTORICO) || '[]');
    _assert(histBruto.length >= 2, 'deveria haver ao menos 2 backups no histórico');
    _assertEquals(histBruto[0].rotulo, 'backup de teste', 'entrada mais recente do histórico deveria ter a etiqueta salva');
    _assert(!histBruto[1].hasOwnProperty('rotulo'), 'backup sem etiqueta não deveria ganhar o campo rotulo (compatibilidade)');

    return 'ok';
  } finally {
    var hist = [];
    try { hist = JSON.parse(PropertiesService.getScriptProperties().getProperty(_KEY_BACKUP_HISTORICO) || '[]'); } catch (_) {}
    var ss = getSS();
    hist.forEach(function(h) {
      try { var wsB = ss.getSheetByName(h.sheetName); if (wsB) ss.deleteSheet(wsB); } catch (_) {}
    });
    if (chaveAnterior === null) PropertiesService.getScriptProperties().deleteProperty(_KEY_BACKUP_HISTORICO);
    else PropertiesService.getScriptProperties().setProperty(_KEY_BACKUP_HISTORICO, chaveAnterior);
  }
}

function testeAlertasCapacidadeAdmin() {
  // Novo (Painel Admin, item 9): confere que cada aba reportada tem os
  // campos esperados e que as abas fixas de fornecedor aparecem na lista.
  var r = JSON.parse(obterAlertasCapacidadeAdmin());
  _assert(!r.erro, 'obterAlertasCapacidadeAdmin erro: ' + (r.erro || ''));
  _assert(Array.isArray(r.abas), 'abas deveria ser array');
  _assert(r.abas.length > 0, 'deveria haver ao menos 1 aba reportada');

  var algumaFornecedor = r.abas.some(function(a) { return ABAS_OPERACIONAIS.indexOf(a.nome) !== -1; });
  _assert(algumaFornecedor, 'deveria incluir ao menos uma das abas fixas de fornecedor');

  r.abas.forEach(function(a) {
    _assertContains(a, 'nome', 'cada entrada deveria ter "nome"');
    _assertContains(a, 'linhas', 'cada entrada deveria ter "linhas"');
    _assert(a.linhas >= 0, 'linhas não deveria ser negativo (' + a.nome + ')');
  });

  return 'abas=' + r.abas.length;
}

function testeConfiguracaoCompletaRoundTripSemChavesSecretas() {
  // Novo (Painel Admin, item 14): a exportação não deve incluir chaves
  // secretas (token/segredo do Telegram), e a importação deve aplicar
  // apenas as chaves da allow-list — o resto (ex.: _KEY_WEBHOOK_CONF) não
  // pode ser tocado.
  var chaveHistAnterior     = PropertiesService.getScriptProperties().getProperty(_KEY_BACKUP_HISTORICO);
  var chaveWebhookAnterior  = PropertiesService.getScriptProperties().getProperty(_KEY_WEBHOOK_CONF);
  try {
    var rExport = JSON.parse(exportarConfiguracaoCompletaAdmin());
    _assert(!rExport.erro, 'exportarConfiguracaoCompletaAdmin erro: ' + (rExport.erro || ''));
    _assert(!!rExport.config, 'campo config ausente na exportação');
    _assert(!rExport.config.hasOwnProperty('webhookConf') && !rExport.config.hasOwnProperty('telegram'),
      'exportação não deveria incluir configuração do Telegram (token/segredo)');

    var rImport = JSON.parse(importarConfiguracaoCompletaAdmin(rExport.config));
    _assert(!rImport.erro, 'importarConfiguracaoCompletaAdmin erro: ' + (rImport.erro || ''));

    var webhookDepois = PropertiesService.getScriptProperties().getProperty(_KEY_WEBHOOK_CONF);
    _assertEquals(webhookDepois, chaveWebhookAnterior, '_KEY_WEBHOOK_CONF não deveria ser alterado pela importação');

    return 'ok';
  } finally {
    // importarConfiguracaoCompletaAdmin chama executarBackup(...) como rede
    // de segurança antes de gravar — remove a aba de backup criada por isso.
    var hist = [];
    try { hist = JSON.parse(PropertiesService.getScriptProperties().getProperty(_KEY_BACKUP_HISTORICO) || '[]'); } catch (_) {}
    var ss = getSS();
    hist.forEach(function(h) {
      try { var wsB = ss.getSheetByName(h.sheetName); if (wsB) ss.deleteSheet(wsB); } catch (_) {}
    });
    if (chaveHistAnterior === null) PropertiesService.getScriptProperties().deleteProperty(_KEY_BACKUP_HISTORICO);
    else PropertiesService.getScriptProperties().setProperty(_KEY_BACKUP_HISTORICO, chaveHistAnterior);
  }
}
