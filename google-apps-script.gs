/**
 * MULHERES CURADAS - Karoline Rodrigues
 * Recebe as inscrições da página, salva na planilha, envia o e-mail
 * de confirmação para a participante e avisa a organizadora.
 *
 * Como instalar/atualizar: veja o arquivo INSTRUCOES.md.
 */

// ===== CONFIGURAÇÕES (edite se quiser) =====
var NOME_EVENTO    = 'Mulheres Curadas';
var REMETENTE      = 'Mulheres Curadas';           // nome que aparece como remetente
var NOME_ABA       = 'Inscrições';                  // aba onde as inscrições são salvas
var NOME_CONFERENCIA = 'Conferência';               // aba com a lista para conferir (check-in)
var NOME_PRESENCA  = 'Presença';                    // aba que guarda o check-in feito pelo site
var CHAVE_LISTA    = 'curadas2026';                 // senha para abrir a lista no site (TROQUE!)
var EMAIL_AVISO    = 'solysprojetos@gmail.com';     // quem recebe o aviso de nova inscrição
var DATA_EVENTO    = '1º DE AGOSTO DE 2026 (SÁBADO), ÀS 18H';
var LOCAL_EVENTO   = 'CC VISÃO PROFÉTICA, AV. DOS MARINHEIROS, 319, CIDADE NOVA, MARACANAÚ, CE';
var NOME_COMENTARIOS = 'Comentários';               // aba com os comentários/depoimentos
var INSCRICOES_ABERTAS = false;                     // false = evento já aconteceu (form encerrado)
var URL_SITE       = 'https://mulherescuradas.institutoabner.com.br/'; // site (link do comentário)
// ===========================================

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var dados = JSON.parse(e.postData.contents);

    // check-in de presença vindo da página do site
    if (dados.acao === 'presenca') {
      if (dados.chave !== CHAVE_LISTA) return resposta({ ok: false, erro: 'nao_autorizado' });
      marcarPresenca_(dados.email, dados.presente === true || dados.presente === 'true');
      return resposta({ ok: true });
    }

    // comentário/depoimento vindo da página (pós-evento)
    if (dados.acao === 'comentario') {
      salvarComentario_(dados.nome, dados.comentario, dados.grupo);
      avisarComentario_(dados);
      return resposta({ ok: true });
    }

    // inscrições encerradas: não aceita novas inscrições
    if (!INSCRICOES_ABERTAS) {
      return resposta({ ok: false, erro: 'inscricoes_encerradas' });
    }

    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName(NOME_ABA) || ss.getSheets()[0];

    if (aba.getLastRow() === 0) {
      aba.appendRow(['Data/Hora', 'Nome', 'WhatsApp', 'E-mail', 'Grupo']);
      aba.getRange(1, 1, 1, 5).setFontWeight('bold');
    }

    aba.appendRow([
      new Date(),
      padronizarNome(dados.nome),
      padronizarWhatsapp(dados.whatsapp),
      padronizarEmail(dados.email),
      padronizarGrupo(dados.grupo)
    ]);

    var total = aba.getLastRow() - 1; // desconta o cabecalho

    if (dados.email) { enviarConfirmacao(dados.nome, dados.email); }
    avisarOrganizadora(dados, total, ss.getUrl());

    return resposta({ ok: true });
  } catch (err) {
    return resposta({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// e-mail para a participante
function enviarConfirmacao(nome, email) {
  var primeiroNome = (nome || '').split(' ')[0] || '';
  var html =
    '<div style="font-family:Arial,sans-serif;max-width:520px;margin:auto;background:#170f0d;color:#f4eef6;border-radius:16px;overflow:hidden">' +
      '<div style="padding:28px 26px;background:linear-gradient(90deg,#af7569,#c88a80,#d9a898);text-align:center">' +
        '<h1 style="margin:0;font-size:24px;color:#fff">Inscrição confirmada!</h1>' +
      '</div>' +
      '<div style="padding:28px 26px">' +
        '<p style="font-size:16px">Olá ' + primeiroNome + ', tudo bem? 💖</p>' +
        '<p style="font-size:15px;line-height:1.6;color:#ddccc2">' +
          'Recebemos a sua inscrição para o <b style="color:#c88a80">' + NOME_EVENTO + '</b>. Está tudo certo! Anota aí:' +
        '</p>' +
        '<div style="background:#221715;border:1px solid rgba(175,117,105,.4);border-radius:12px;padding:16px 20px;margin:16px 0">' +
          '<p style="margin:6px 0;font-size:15px">📅 <b style="color:#c88a80">' + DATA_EVENTO + '</b></p>' +
          '<p style="margin:6px 0;font-size:14px;line-height:1.5;color:#ddccc2">📍 ' + LOCAL_EVENTO + '</p>' +
          '<p style="margin:6px 0;font-size:14px;color:#c88a80">⏰ INICIAREMOS <b>PONTUALMENTE</b>. NÃO SE ATRASE!</p>' +
        '</div>' +
        '<p style="font-size:15px;line-height:1.6;color:#ddccc2">Com carinho,<br>Karoline Rodrigues</p>' +
      '</div>' +
    '</div>';
  MailApp.sendEmail({ to: email, subject: 'Inscrição confirmada 💖 ' + NOME_EVENTO, htmlBody: html, name: REMETENTE });
}

// aviso para a organizadora a cada nova inscricao
function avisarOrganizadora(dados, total, urlPlanilha) {
  var assunto = '🔔 Nova inscrição: ' + (dados.nome || 'sem nome') + ' (total: ' + total + ')';
  var html =
    '<div style="font-family:Arial,sans-serif;max-width:480px;margin:auto">' +
      '<h2 style="color:#af7569">Nova inscrição no ' + NOME_EVENTO + ' 💖</h2>' +
      '<table style="border-collapse:collapse;width:100%;font-size:14px">' +
        linha('Nome', dados.nome) +
        linha('WhatsApp', dados.whatsapp) +
        linha('E-mail', dados.email) +
        linha('Grupo', dados.grupo) +
      '</table>' +
      '<p style="font-size:16px;margin-top:16px">Total de inscritas: <b style="color:#af7569">' + total + '</b></p>' +
      '<p><a href="' + urlPlanilha + '" style="color:#af7569">Abrir a planilha completa</a></p>' +
    '</div>';
  MailApp.sendEmail({ to: EMAIL_AVISO, subject: assunto, htmlBody: html, name: REMETENTE });
}

function linha(rotulo, valor) {
  return '<tr>' +
    '<td style="padding:7px 10px;border:1px solid #eee;background:#faf5f8;font-weight:bold">' + rotulo + '</td>' +
    '<td style="padding:7px 10px;border:1px solid #eee">' + (valor || '-') + '</td>' +
  '</tr>';
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.lista === '1') {
    var saida = (p.chave === CHAVE_LISTA) ? dadosDaLista() : { ok: false, erro: 'nao_autorizado' };
    var json = JSON.stringify(saida);
    if (p.callback) {
      return ContentService.createTextOutput(p.callback + '(' + json + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT); // JSONP (funciona no site sem CORS)
    }
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('OK - servico de inscricoes ativo.');
}

// monta a lista organizada (para a pagina do site)
function dadosDaLista() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var origem = ss.getSheetByName(NOME_ABA) || ss.getSheets()[0];
  var ult    = origem.getLastRow();
  var presenca = lerPresenca_();
  var testes = ['gustavo teste', 'teste', 'solys projetos', 'teste cores e local'];
  var lista = [], grupos = {}, vistos = {};
  if (ult >= 2) {
    var dados = origem.getRange(2, 1, ult - 1, 5).getValues();
    for (var i = 0; i < dados.length; i++) {
      var nome  = padronizarNome(dados[i][1]);
      var email = padronizarEmail(dados[i][3]);
      if (!nome && !email) continue;
      if (testes.indexOf(nome.toLowerCase()) !== -1) continue;
      if (vistos[nome]) continue; // tira nomes repetidos
      vistos[nome] = true;
      var grupo = padronizarGrupo(dados[i][4]);
      var d = dados[i][0];
      var dataStr = (d instanceof Date)
        ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
        : String(d || '');
      lista.push({
        nome: nome, grupo: grupo,
        whatsapp: padronizarWhatsapp(dados[i][2]),
        email: email, data: dataStr,
        presente: presenca[email] === true
      });
      grupos[grupo] = (grupos[grupo] || 0) + 1;
    }
    lista.sort(function (a, b) {
      var s = String(a.grupo).localeCompare(String(b.grupo), 'pt-BR');
      if (s !== 0) return s;
      return String(a.nome).localeCompare(String(b.nome), 'pt-BR');
    });
  }
  return {
    ok: true,
    atualizadoEm: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
    total: lista.length, grupos: grupos, inscricoes: lista
  };
}

// lê o mapa de presenças já marcadas (e-mail -> true)
function lerPresenca_() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(NOME_PRESENCA);
  var mapa = {};
  if (aba && aba.getLastRow() >= 2) {
    var v = aba.getRange(2, 1, aba.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < v.length; i++) {
      var em = String(v[i][0] || '').toLowerCase().trim();
      if (em) mapa[em] = (v[i][1] === true);
    }
  }
  return mapa;
}

// marca/desmarca presença de um e-mail (guarda na aba Presença)
function marcarPresenca_(email, presente) {
  email = String(email || '').toLowerCase().trim();
  if (!email) return;
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(NOME_PRESENCA);
  if (!aba) { aba = ss.insertSheet(NOME_PRESENCA); aba.appendRow(['E-mail', 'Presente', 'Quando']); }
  var last = aba.getLastRow();
  var col = last >= 2 ? aba.getRange(2, 1, last - 1, 1).getValues() : [];
  for (var i = 0; i < col.length; i++) {
    if (String(col[i][0] || '').toLowerCase().trim() === email) {
      aba.getRange(i + 2, 2, 1, 2).setValues([[presente, new Date()]]);
      return;
    }
  }
  aba.appendRow([email, presente, new Date()]);
}

// ===== PADRONIZADORES (deixam cada campo no padrão) =====
// Nome: sem espaços sobrando e em MAIÚSCULO
function padronizarNome(v) {
  return String(v || '').replace(/\s+/g, ' ').trim().toUpperCase();
}
// E-mail: sem espaços e em minúsculo
function padronizarEmail(v) {
  return String(v || '').replace(/\s+/g, '').toLowerCase();
}
// Grupo: casing canônico dos grupos conhecidos
function padronizarGrupo(v) {
  var g = String(v || '').replace(/\s+/g, ' ').trim();
  var mapa = { 'sgroup': 'SGroup', 'solys': 'Solys', 'convidada': 'Convidada' };
  return mapa[g.toLowerCase()] || g;
}
// WhatsApp: (DD) 9XXXX-XXXX  (aceita com/sem +55, espaços, traços, parênteses)
function padronizarWhatsapp(v) {
  var d = String(v || '').replace(/\D/g, '');
  if (d.length > 11 && d.indexOf('55') === 0) d = d.slice(d.length - 11); // remove +55
  if (d.length === 11) return '(' + d.slice(0, 2) + ') ' + d.slice(2, 7) + '-' + d.slice(7);
  if (d.length === 10) return '(' + d.slice(0, 2) + ') ' + d.slice(2, 6) + '-' + d.slice(6);
  return String(v || '').trim(); // formato inesperado: mantém como veio
}

// ===== ORGANIZA A PLANILHA (rode 1x para arrumar tudo que já está lá) =====
// Padroniza todos os campos, remove linhas de teste e duplicadas (por e-mail),
// formata as datas, ordena por nome e deixa o visual no padrão.
function organizarPlanilha() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(NOME_ABA) || ss.getSheets()[0];
  var ult = aba.getLastRow();

  // cabeçalho no padrão
  aba.getRange(1, 1, 1, 5)
     .setValues([['Data/Hora', 'Nome', 'WhatsApp', 'E-mail', 'Grupo']])
     .setFontWeight('bold');
  if (ult < 2) { aba.setFrozenRows(1); return; } // só cabeçalho

  var testes = ['gustavo teste', 'teste', 'solys projetos', 'teste cores e local'];
  var dados  = aba.getRange(2, 1, ult - 1, 5).getValues();
  var limpos = [];
  var vistos = {}; // controle de duplicadas

  for (var i = 0; i < dados.length; i++) {
    var nome  = padronizarNome(dados[i][1]);
    var email = padronizarEmail(dados[i][3]);

    if (!nome && !email) continue;                            // linha vazia
    if (testes.indexOf(nome.toLowerCase()) !== -1) continue;  // teste
    var chave = email || nome.toLowerCase();
    if (vistos[chave]) continue;                              // duplicada
    vistos[chave] = true;

    limpos.push([
      dados[i][0],
      nome,
      padronizarWhatsapp(dados[i][2]),
      email,
      padronizarGrupo(dados[i][4])
    ]);
  }

  // ordena por setor (grupo) -> nome -> data
  limpos.sort(function (a, b) {
    var porSetor = String(a[4]).localeCompare(String(b[4]), 'pt-BR');
    if (porSetor !== 0) return porSetor;
    var porNome = String(a[1]).localeCompare(String(b[1]), 'pt-BR');
    if (porNome !== 0) return porNome;
    return (a[0] instanceof Date ? a[0].getTime() : 0) -
           (b[0] instanceof Date ? b[0].getTime() : 0);
  });

  // regrava o corpo já organizado
  aba.getRange(2, 1, ult - 1, 5).clearContent();
  if (limpos.length) {
    aba.getRange(2, 1, limpos.length, 5).setValues(limpos);
    aba.getRange(2, 1, limpos.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  }

  aba.setFrozenRows(1);
  aba.autoResizeColumns(1, 5);
  SpreadsheetApp.flush();
  Logger.log('Planilha organizada: ' + limpos.length + ' inscrição(ões).');
}

// ===== MENU dentro da planilha (aparece ao abrir o arquivo) =====
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🌸 Mulheres Curadas')
    .addItem('Atualizar lista de conferência', 'montarListaConferencia')
    .addItem('Organizar planilha de inscrições', 'organizarPlanilha')
    .addSeparator()
    .addItem('Atualizar presença oficial (só presentes)', 'atualizarPresencaOficial')
    .addItem('Enviar e-mail de agradecimento', 'enviarAgradecimento')
    .addToUi();
}

// ===== LISTA DE CONFERÊNCIA (check-in) =====
// Cria/atualiza a aba "Conferência" com a lista organizada e uma caixinha
// para marcar quem já foi conferido/chegou. Ao atualizar, NÃO perde as marcas.
function montarListaConferencia() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var origem = ss.getSheetByName(NOME_ABA) || ss.getSheets()[0];
  var ult    = origem.getLastRow();

  // guarda quem já estava marcado (por e-mail) para não perder ao atualizar
  var marcados = {};
  var conf = ss.getSheetByName(NOME_CONFERENCIA);
  if (conf && conf.getLastRow() >= 2) {
    var antigos = conf.getRange(2, 1, conf.getLastRow() - 1, 5).getValues();
    for (var k = 0; k < antigos.length; k++) {
      var em = String(antigos[k][4] || '').toLowerCase().trim();
      if (em && antigos[k][0] === true) marcados[em] = true;
    }
  }

  // lê e organiza as inscrições (padroniza, tira testes e nomes repetidos)
  var testes = ['gustavo teste', 'teste', 'solys projetos', 'teste cores e local'];
  var linhas = [];
  if (ult >= 2) {
    var dados = origem.getRange(2, 1, ult - 1, 5).getValues();
    var vistos = {};
    for (var i = 0; i < dados.length; i++) {
      var nome  = padronizarNome(dados[i][1]);
      var email = padronizarEmail(dados[i][3]);
      if (!nome && !email) continue;                            // linha vazia
      if (testes.indexOf(nome.toLowerCase()) !== -1) continue;  // teste
      if (vistos[nome]) continue;                               // nome repetido
      vistos[nome] = true;
      linhas.push([
        marcados[email] === true,        // A - Conferida (caixinha)
        nome,                            // B - Nome
        padronizarGrupo(dados[i][4]),    // C - Grupo
        padronizarWhatsapp(dados[i][2]), // D - WhatsApp
        email,                           // E - E-mail
        dados[i][0]                      // F - Data/Hora
      ]);
    }
    linhas.sort(function (a, b) {
      var s = String(a[2]).localeCompare(String(b[2]), 'pt-BR'); // setor
      if (s !== 0) return s;
      return String(a[1]).localeCompare(String(b[1]), 'pt-BR');  // nome
    });
  }

  // cria/limpa a aba de conferência
  if (!conf) conf = ss.insertSheet(NOME_CONFERENCIA);
  conf.clear();
  conf.getRange(1, 1, 1, 6)
      .setValues([['Conferida', 'Nome', 'Grupo', 'WhatsApp', 'E-mail', 'Data/Hora']])
      .setFontWeight('bold');

  if (linhas.length) {
    conf.getRange(2, 1, linhas.length, 6).setValues(linhas);
    conf.getRange(2, 1, linhas.length, 1).insertCheckboxes();
    conf.getRange(2, 6, linhas.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  }

  // contadores automáticos (à direita)
  conf.getRange('H1').setValue('Total inscritas').setFontWeight('bold');
  conf.getRange('I1').setFormula('=COUNTA(B2:B)');
  conf.getRange('H2').setValue('Conferidas').setFontWeight('bold');
  conf.getRange('I2').setFormula('=COUNTIF(A2:A,TRUE)');
  conf.getRange('H3').setValue('Faltam').setFontWeight('bold');
  conf.getRange('I3').setFormula('=I1-I2');

  conf.setFrozenRows(1);
  conf.autoResizeColumns(1, 6);
  ss.setActiveSheet(conf);
  SpreadsheetApp.flush();
  Logger.log('Lista de conferência atualizada: ' + linhas.length + ' inscrição(ões).');
}

// ===== DISPARO: e-mail com a arte da semana =====
function enviarBomInicioDeSemana() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName('Inscrições') || ss.getSheets()[0];
  var valores = aba.getDataRange().getValues();
  var IMG_URL = 'https://mulherescuradas.institutoabner.com.br/semana-mulheres-curadas.jpg';
  var imagem  = UrlFetchApp.fetch(IMG_URL).getBlob().setName('semana.jpg');
  var assunto = '💖 Que Deus abençoe a sua semana!';
  var testes  = ['gustavo teste', 'teste', 'solys projetos', 'teste cores e local'];
  var jaEnviei = {}, total = 0;
  for (var i = 1; i < valores.length; i++) {
    var nome  = String(valores[i][1] || '').trim();
    var email = String(valores[i][3] || '').trim();
    var chave = email.toLowerCase();
    if (!email || email.indexOf('@') === -1) continue;
    if (testes.indexOf(nome.toLowerCase()) !== -1) continue;
    if (jaEnviei[chave]) continue;
    jaEnviei[chave] = true;
    var html =
      '<div style="font-family:Arial,sans-serif;max-width:600px;margin:auto">' +
        '<img src="cid:semana" style="width:100%;display:block;border-radius:12px" alt="Mulheres Curadas">' +
      '</div>';
    MailApp.sendEmail({ to: email, subject: assunto, htmlBody: html, name: 'Mulheres Curadas', inlineImages: { semana: imagem } });
    total++;
    Utilities.sleep(300);
  }
  Logger.log('Disparo enviado para ' + total + ' pessoas.');
}

// ===== DISPARO: contagem regressiva "Faltam 5 dias" =====
function enviarFalta5Dias() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName('Inscrições') || ss.getSheets()[0];
  var valores = aba.getDataRange().getValues();
  var IMG_URL = 'https://mulherescuradas.institutoabner.com.br/falta-5-dias-mulheres-curadas.jpg';
  var imagem  = UrlFetchApp.fetch(IMG_URL).getBlob().setName('falta5dias.jpg');
  var assunto = '🌸 Faltam apenas 5 dias para o Mulheres Curadas!';
  var testes  = ['gustavo teste', 'teste', 'solys projetos', 'teste cores e local'];
  var jaEnviei = {}, total = 0;
  for (var i = 1; i < valores.length; i++) {
    var nome  = String(valores[i][1] || '').trim();
    var email = String(valores[i][3] || '').trim();
    var chave = email.toLowerCase();
    if (!email || email.indexOf('@') === -1) continue;
    if (testes.indexOf(nome.toLowerCase()) !== -1) continue;
    if (jaEnviei[chave]) continue;
    jaEnviei[chave] = true;
    var html =
      '<div style="font-family:Arial,sans-serif;max-width:600px;margin:auto;background:#170f0d;color:#f4eef6;border-radius:16px;overflow:hidden">' +
        '<img src="cid:arte" style="width:100%;display:block" alt="Faltam 5 dias - Mulheres Curadas">' +
        '<div style="padding:26px 24px">' +
          '<p style="font-size:18px;text-align:center;color:#c88a80;margin:0 0 18px"><b>🌸 Faltam apenas 5 dias! 🌸</b></p>' +
          '<p style="font-size:15px;line-height:1.7;color:#ddccc2">Está chegando o momento de vivermos um encontro transformador. O <b style="color:#c88a80">Mulheres Curadas</b> foi preparado com muito amor, oração e propósito para cada mulher que deseja experimentar cura, restauração e um novo tempo na presença de Deus.</p>' +
          '<p style="font-size:15px;line-height:1.7;color:#ddccc2">✨ Reserve esta data e permita-se viver essa experiência. Deus tem algo especial preparado para a sua vida!</p>' +
          '<div style="background:#221715;border:1px solid rgba(175,117,105,.4);border-radius:12px;padding:16px 20px;margin:18px 0">' +
            '<p style="margin:6px 0;font-size:15px">📍 <b style="color:#c88a80">Local:</b> CC Visão Profética</p>' +
            '<p style="margin:6px 0;font-size:14px;line-height:1.5;color:#ddccc2">Av. dos Marinheiros, 319 – Cidade Nova, Maracanaú.</p>' +
            '<p style="margin:12px 0 6px;font-size:15px">🕕 <b style="color:#c88a80">Horário:</b> Às 18h00.</p>' +
            '<p style="margin:6px 0;font-size:14px;line-height:1.5;color:#ddccc2">Iniciaremos pontualmente, por isso chegue com antecedência para que possamos começar juntas esse momento tão especial.</p>' +
          '</div>' +
          '<p style="font-size:15px;text-align:center;color:#c88a80;margin:18px 0">💖 Esperamos por você!</p>' +
          '<p style="font-size:14px;line-height:1.6;color:#ddccc2;font-style:italic;text-align:center;border-top:1px solid rgba(175,117,105,.3);padding-top:16px;margin-top:16px">' +
            '"Dar-vos-ei um coração novo e porei dentro de vós um espírito novo."<br>' +
            '<span style="color:#c88a80;font-style:normal">Ezequiel 36:26</span>' +
          '</p>' +
        '</div>' +
      '</div>';
    MailApp.sendEmail({ to: email, subject: assunto, htmlBody: html, name: 'Mulheres Curadas', inlineImages: { arte: imagem } });
    total++;
    Utilities.sleep(300);
  }
  Logger.log('Disparo "Faltam 5 dias" enviado para ' + total + ' pessoas.');
}

// ===== DISPARO: contagem regressiva "Faltam 2 dias" =====
var ASSUNTO_2DIAS = '🌸 Faltam apenas 2 dias para o Mulheres Curadas!';

// limpa lixo no fim do e-mail (ponto, vírgula, espaço) e valida o formato
function limparEmail_(e) {
  return String(e || '').trim().replace(/[.,;\s]+$/, '');
}
function emailValido_(e) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e);
}

// monta o HTML do e-mail (usado pelas duas funções abaixo)
function htmlFalta2Dias_() {
  return '<div style="font-family:Arial,sans-serif;max-width:600px;margin:auto;background:#170f0d;color:#f4eef6;border-radius:16px;overflow:hidden">' +
      '<div style="padding:28px 26px;background:linear-gradient(90deg,#af7569,#c88a80,#d9a898);text-align:center">' +
        '<h1 style="margin:0;font-size:24px;color:#fff">🌸 Faltam apenas 2 dias! 🌸</h1>' +
      '</div>' +
      '<div style="padding:26px 24px">' +
        '<p style="font-size:15px;line-height:1.7;color:#ddccc2">Está chegando o momento de vivermos um encontro transformador. O <b style="color:#c88a80">Mulheres Curadas</b> foi preparado com muito amor, oração e propósito para cada mulher que deseja experimentar cura, restauração e um novo tempo na presença de Deus.</p>' +
        '<p style="font-size:15px;line-height:1.7;color:#ddccc2">✨ Reserve esta data e permita-se viver essa experiência. Deus tem algo especial preparado para a sua vida!</p>' +
        '<div style="background:#221715;border:1px solid rgba(175,117,105,.4);border-radius:12px;padding:16px 20px;margin:18px 0">' +
          '<p style="margin:6px 0;font-size:15px">📍 <b style="color:#c88a80">Local:</b> CC Visão Profética</p>' +
          '<p style="margin:6px 0;font-size:14px;line-height:1.5;color:#ddccc2">Av. dos Marinheiros, 319 – Cidade Nova, Maracanaú.</p>' +
          '<p style="margin:12px 0 6px;font-size:15px">🕕 <b style="color:#c88a80">Horário:</b> Às 18h00.</p>' +
          '<p style="margin:6px 0;font-size:14px;line-height:1.5;color:#ddccc2">Iniciaremos pontualmente, por isso chegue com antecedência para que possamos começar juntas esse momento tão especial.</p>' +
        '</div>' +
        '<p style="font-size:15px;text-align:center;color:#c88a80;margin:18px 0">💖 Esperamos por você!</p>' +
        '<p style="font-size:14px;line-height:1.6;color:#ddccc2;font-style:italic;text-align:center;border-top:1px solid rgba(175,117,105,.3);padding-top:16px;margin-top:16px">' +
          '"Dar-vos-ei um coração novo e porei dentro de vós um espírito novo."<br>' +
          '<span style="color:#c88a80;font-style:normal">Ezequiel 36:26</span>' +
        '</p>' +
      '</div>' +
    '</div>';
}

// Dispara para TODA a planilha (limpa e-mails e pula os invalidos sem travar)
function enviarFalta2Dias() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName('Inscrições') || ss.getSheets()[0];
  var valores = aba.getDataRange().getValues();
  var testes  = ['gustavo teste', 'teste', 'solys projetos', 'teste cores e local'];
  var html = htmlFalta2Dias_();
  var jaEnviei = {}, total = 0, pulados = 0;
  for (var i = 1; i < valores.length; i++) {
    var nome  = String(valores[i][1] || '').trim();
    var email = limparEmail_(valores[i][3]);
    var chave = email.toLowerCase();
    if (!emailValido_(email)) { if (email) { pulados++; Logger.log('E-mail invalido pulado: ' + email + ' (' + nome + ')'); } continue; }
    if (testes.indexOf(nome.toLowerCase()) !== -1) continue;
    if (jaEnviei[chave]) continue;
    jaEnviei[chave] = true;
    try {
      MailApp.sendEmail({ to: email, subject: ASSUNTO_2DIAS, htmlBody: html, name: 'Mulheres Curadas' });
      total++;
    } catch (err) {
      pulados++;
      Logger.log('Falhou para ' + email + ': ' + err);
    }
    Utilities.sleep(300);
  }
  Logger.log('Disparo "Faltam 2 dias": ' + total + ' enviados, ' + pulados + ' pulados.');
}

// Envia SÓ para quem ainda não recebeu (a lista parou no e-mail com erro).
// Use esta função depois da falha para não mandar duplicado para quem já recebeu.
function enviarFalta2DiasRestantes() {
  var restantes = [
    'camilaaraujoribeiro0401@gmail.com', // tinha um ponto extra no fim (corrigido aqui)
    'vitoria.lorinha29@gmail.com',
    'hermanueleandrade@gmail.com',
    'luanepaiva26@icloud.com',
    'claudiana09paiva@gmail.com',
    'cristinavalerio2811@gmail.com',
    'biancadecastro15@gmail.com',
    'amanda.1995diassousa@gmail.com',
    'davidmarinho23@icloud.com'
  ];
  var html = htmlFalta2Dias_();
  var total = 0, pulados = 0;
  for (var i = 0; i < restantes.length; i++) {
    var email = limparEmail_(restantes[i]);
    if (!emailValido_(email)) { pulados++; continue; }
    try {
      MailApp.sendEmail({ to: email, subject: ASSUNTO_2DIAS, htmlBody: html, name: 'Mulheres Curadas' });
      total++;
    } catch (err) {
      pulados++;
      Logger.log('Falhou para ' + email + ': ' + err);
    }
    Utilities.sleep(300);
  }
  Logger.log('Disparo "Faltam 2 dias" (restantes): ' + total + ' enviados, ' + pulados + ' pulados.');
}

// ======================================================================
// ===== PÓS-EVENTO: presença oficial, agradecimento e comentários ======
// ======================================================================

// Normaliza um nome para comparar (MAIÚSCULO, sem acento, espaços únicos).
function normalizarChaveNome_(v) {
  var s = String(v || '').replace(/\s+/g, ' ').trim().toUpperCase();
  s = s.normalize ? s.normalize('NFD').replace(/[\u0300-\u036f]/g, '') : s;
  return s;
}

// Lista OFICIAL de quem esteve PRESENTE (68), conforme a lista de presença
// gerada em 01/08/2026 19:06. Usada para atualizar a planilha e para o
// disparo de agradecimento (só vai para quem esteve presente).
var PRESENTES_OFICIAL = [
  // CONVIDADAS (27)
  'ADRIELLE RODRIGUES FIDELES PINHEIRO', 'AMANDA HULY', 'ANDREA DA SILVA FREITAS',
  'ANTONIA RAQUEL', 'BRUNA KAREN', 'BRUNA LIMA', 'CAMILA DE ARAÚJO RIBEIRO',
  'CLEIDELENE NOGUEIRA NUNES', 'CRISTIANA ARAUJO DA COSTA MONTEIRO',
  'DARLYANE DA SILVA OLIVEIRA', 'ELAINE', 'ELISANGELA SILVA ARAÚJO',
  'ELVIRENE VALÉRIO', 'EMILE MARINHO', 'IVONETE MARINHO', 'JESSIKA SZABO',
  'LIDIANE MONTE', 'LUANA VANESSA', 'MARIA ANETE MARINHO DA SILVA',
  'MARIA DO SOCORRO PEREIRA SAMPAIO', 'MARIA SAYONARA ALVES GURGEL', 'MILENE BARROSO',
  'NATASHA SARAIVA PAULA', 'PATRICIA BRAGA', 'PRISCILA NASCIMENTO COSTA PEREIRA',
  'ROSE MARINHO', 'VITORIA OLIVEIRA ALBUQUERQUE',
  // SGROUP NACIONAL (18)
  'ALICE VITÓRIA', 'AMANDA SOUSA DIAS', 'CLAUDIANA PAIVA LIMA SILVA', 'ELIZABETH BERANEK',
  'ERIKA MARQUES LEITAO', 'GEISSIANE MARTINS DE SOUZA ANOS', 'ISABELLE GOMES',
  'JULIANA DE SOUSA SILVA PAULINO', 'LILIANE DA SILVA BANDEIRA',
  'LUANE PEREIRA DE OLIVEIRA PAIVA', 'MALINDA GAIL ABREU BARROSO', 'MEIRE MARQUES',
  'MILENA VIANA', 'NAYARA ALVES VASCONCELOS', 'SANDRA MOREIRA', 'SARAH BRITO',
  'STHEFFANY ARAUJO', 'THAIZ RHAIANE BEZERRA BARROS',
  // SOLYS GESTÃO ADMINISTRATIVA (9)
  'AGLAY CRISTHIAN DE O. TEIXEIRA', 'ANDREYZA MATOS DE OLIVEIRA', 'CAMILA RODRIGUES MESQUITA',
  'ERISHELDA DE SOUSA DAMASCENO CAVALCANTE', 'FRANCISCA SAMILA DE BRITO SANTOS',
  'KALINA RODRIGUES CARVALHO', 'LIANA HOLANDA', 'NAYELLY ALVES DA SILVA',
  'SARAH CARDOSO SILVA',
  // VOLUNTÁRIOS (14)
  'BRUNA DE JESUS OLIVEIRA', 'CRISTINA MARTINS', 'DAFINY WENDY GOMES LIMA',
  'JOÃO DAVI DE SOUZA NASCIMENTO', 'LUIZ DANILO AZEVEDO',
  'MANOEL EUFRASIO GALVÃO DO NASCIMENTO', 'MARIA DE FÁTIMA DA COSTA ANDRADE',
  'NATHIELY MORAIS', 'NICLAILTON DA COSTA ANDRADE', 'OTNIEL DE OLIVEIRA SOARES',
  'OZIEL DE OLIVEIRA SOARES', 'PABLO PATRICK MENDES DA SILVA', 'PAMELA DANTAS',
  'RAFAELA MENDES'
];

// Atualiza a planilha com a PRESENÇA OFICIAL: marca como presentes só quem
// esteve no evento (lista acima) e ausentes todo o resto — nas abas
// "Presença" (por e-mail, alimenta o site) e "Conferência" (as caixinhas).
function atualizarPresencaOficial() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var origem = ss.getSheetByName(NOME_ABA) || ss.getSheets()[0];
  var ult    = origem.getLastRow();

  // conjunto de nomes presentes (normalizados)
  var presentesNome = {};
  for (var j = 0; j < PRESENTES_OFICIAL.length; j++) {
    presentesNome[normalizarChaveNome_(PRESENTES_OFICIAL[j])] = true;
  }

  // 1) aba "Presença" (por e-mail) — reescreve tudo: presente = true/false
  var aba = ss.getSheetByName(NOME_PRESENCA) || ss.insertSheet(NOME_PRESENCA);
  aba.clear();
  aba.getRange(1, 1, 1, 3).setValues([['E-mail', 'Presente', 'Quando']]).setFontWeight('bold');

  var linhas = [], agora = new Date(), vistos = {}, achadosNome = {}, presentesEmail = 0;
  if (ult >= 2) {
    var dd = origem.getRange(2, 1, ult - 1, 5).getValues();
    for (var k = 0; k < dd.length; k++) {
      var nomeN = normalizarChaveNome_(dd[k][1]);
      var em    = padronizarEmail(dd[k][3]);
      if (!em || vistos[em]) continue;
      vistos[em] = true;
      var presente = presentesNome[nomeN] === true;
      if (presente) { presentesEmail++; achadosNome[nomeN] = true; }
      linhas.push([em, presente, agora]);
    }
  }
  if (linhas.length) aba.getRange(2, 1, linhas.length, 3).setValues(linhas);
  aba.setFrozenRows(1);

  // nomes do PDF que não bateram com ninguém da planilha (para conferência)
  var naoAchados = [];
  for (var n in presentesNome) { if (!achadosNome[n]) naoAchados.push(n); }

  // 2) aba "Conferência" (por nome) — marca as caixinhas dos presentes
  montarListaConferencia(); // garante a lista atual
  var conf = ss.getSheetByName(NOME_CONFERENCIA);
  if (conf && conf.getLastRow() >= 2) {
    var nomes = conf.getRange(2, 2, conf.getLastRow() - 1, 1).getValues(); // coluna B
    var marcas = [];
    for (var r = 0; r < nomes.length; r++) {
      marcas.push([presentesNome[normalizarChaveNome_(nomes[r][0])] === true]);
    }
    conf.getRange(2, 1, marcas.length, 1).setValues(marcas);
  }

  SpreadsheetApp.flush();
  Logger.log('Presença oficial: ' + presentesEmail + ' presentes com e-mail marcados. ' +
             'Nomes do PDF sem correspondência na planilha: ' + (naoAchados.join(' | ') || 'nenhum'));
}

// ===== DISPARO: e-mail de AGRADECIMENTO (só para quem esteve presente) =====
function enviarAgradecimento() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var origem = ss.getSheetByName(NOME_ABA) || ss.getSheets()[0];
  var ult    = origem.getLastRow();

  var presentesNome = {};
  for (var j = 0; j < PRESENTES_OFICIAL.length; j++) {
    presentesNome[normalizarChaveNome_(PRESENTES_OFICIAL[j])] = true;
  }

  var testes  = ['gustavo teste', 'teste', 'solys projetos', 'teste cores e local'];
  var assunto = '💖 Obrigada por viver o Mulheres Curadas com a gente!';
  var jaEnviei = {}, total = 0, pulados = 0;

  if (ult >= 2) {
    var dd = origem.getRange(2, 1, ult - 1, 5).getValues();
    for (var i = 0; i < dd.length; i++) {
      var nomeBruto = String(dd[i][1] || '').trim();
      var nomeN = normalizarChaveNome_(dd[i][1]);
      var email = limparEmail_(dd[i][3]);
      var chave = email.toLowerCase();
      if (presentesNome[nomeN] !== true) continue;                 // só presentes
      if (testes.indexOf(nomeBruto.toLowerCase()) !== -1) continue;
      if (!emailValido_(email)) { if (email) { pulados++; } continue; }
      if (jaEnviei[chave]) continue;
      jaEnviei[chave] = true;
      var primeiroNome = nomeBruto.split(' ')[0] || '';
      try {
        MailApp.sendEmail({ to: email, subject: assunto, htmlBody: htmlAgradecimento_(primeiroNome), name: REMETENTE });
        total++;
      } catch (err) {
        pulados++;
        Logger.log('Falhou para ' + email + ': ' + err);
      }
      Utilities.sleep(300);
    }
  }
  Logger.log('Disparo de agradecimento: ' + total + ' enviados, ' + pulados + ' pulados.');
}

// HTML do e-mail de agradecimento (com botão para deixar o comentário)
function htmlAgradecimento_(primeiroNome) {
  var ola = primeiroNome ? ('Olá ' + primeiroNome + ', ') : 'Olá, ';
  var linkComentario = URL_SITE.replace(/\/+$/, '') + '/#comentario';
  return '<div style="font-family:Arial,sans-serif;max-width:600px;margin:auto;background:#170f0d;color:#f4eef6;border-radius:16px;overflow:hidden">' +
      '<div style="padding:28px 26px;background:linear-gradient(90deg,#af7569,#c88a80,#d9a898);text-align:center">' +
        '<h1 style="margin:0;font-size:24px;color:#fff">💖 Gratidão!</h1>' +
      '</div>' +
      '<div style="padding:26px 24px">' +
        '<p style="font-size:16px;color:#f4eef6">' + ola + 'que alegria ter você com a gente! 🌸</p>' +
        '<p style="font-size:15px;line-height:1.7;color:#ddccc2">Do fundo do nosso coração, <b style="color:#c88a80">muito obrigada</b> por ter vivido o <b style="color:#c88a80">' + NOME_EVENTO + '</b> conosco. Cada presença fez desse momento algo ainda mais especial. Que a cura, a restauração e o novo tempo que Deus iniciou continuem florescendo na sua vida. 🙏</p>' +
        '<div style="background:#221715;border:1px solid rgba(175,117,105,.4);border-radius:12px;padding:20px;margin:20px 0;text-align:center">' +
          '<p style="margin:0 0 14px;font-size:15px;color:#ddccc2">Conta pra gente: <b style="color:#c88a80">como foi viver esse momento?</b> Seu comentário é muito importante para nós. 💬</p>' +
          '<a href="' + linkComentario + '" style="display:inline-block;background:linear-gradient(90deg,#af7569,#c88a80,#d9a898);color:#4a2c26;text-decoration:none;font-weight:bold;font-size:15px;padding:14px 26px;border-radius:12px">💬 Deixar meu comentário</a>' +
        '</div>' +
        '<p style="font-size:14px;line-height:1.6;color:#ddccc2;font-style:italic;text-align:center;border-top:1px solid rgba(175,117,105,.3);padding-top:16px;margin-top:16px">' +
          '"Dar-vos-ei um coração novo e porei dentro de vós um espírito novo."<br>' +
          '<span style="color:#c88a80;font-style:normal">Ezequiel 36:26</span>' +
        '</p>' +
        '<p style="font-size:15px;line-height:1.6;color:#ddccc2;text-align:center;margin-top:18px">Com carinho,<br>Karoline Rodrigues</p>' +
      '</div>' +
    '</div>';
}

// ===== COMENTÁRIOS (depoimentos pós-evento) =====
// Salva um comentário na aba "Comentários".
function salvarComentario_(nome, comentario, grupo) {
  var texto = String(comentario || '').trim();
  if (!texto) return;
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(NOME_COMENTARIOS);
  if (!aba) {
    aba = ss.insertSheet(NOME_COMENTARIOS);
    aba.appendRow(['Data/Hora', 'Nome', 'Grupo', 'Comentário']);
    aba.getRange(1, 1, 1, 4).setFontWeight('bold');
    aba.setFrozenRows(1);
  }
  aba.appendRow([new Date(), padronizarNome(nome), padronizarGrupo(grupo), texto]);
  aba.getRange(aba.getLastRow(), 1).setNumberFormat('dd/MM/yyyy HH:mm');
}

// Avisa a organizadora quando chega um novo comentário.
function avisarComentario_(dados) {
  var comentario = escaparHtml_(dados.comentario);
  var assunto = '💬 Novo comentário: ' + (dados.nome || 'anônimo');
  var html =
    '<div style="font-family:Arial,sans-serif;max-width:480px;margin:auto">' +
      '<h2 style="color:#af7569">Novo comentário — ' + NOME_EVENTO + ' 💖</h2>' +
      '<table style="border-collapse:collapse;width:100%;font-size:14px">' +
        linha('Nome', dados.nome) +
        linha('Grupo', dados.grupo) +
      '</table>' +
      '<p style="font-size:15px;line-height:1.6;margin-top:14px;white-space:pre-wrap;background:#faf5f8;border:1px solid #eee;border-radius:10px;padding:14px">' + comentario + '</p>' +
    '</div>';
  MailApp.sendEmail({ to: EMAIL_AVISO, subject: assunto, htmlBody: html, name: REMETENTE });
}

// escapa caracteres de HTML (para não quebrar o e-mail de aviso)
function escaparHtml_(v) {
  return String(v || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
