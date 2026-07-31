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
var EMAIL_AVISO    = 'solysprojetos@gmail.com';     // quem recebe o aviso de nova inscrição
var DATA_EVENTO    = '1º DE AGOSTO DE 2026 (SÁBADO), ÀS 18H';
var LOCAL_EVENTO   = 'CC VISÃO PROFÉTICA, AV. DOS MARINHEIROS, 319, CIDADE NOVA, MARACANAÚ, CE';
// ===========================================

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var dados = JSON.parse(e.postData.contents);

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

function doGet() {
  return ContentService.createTextOutput('OK - servico de inscricoes ativo.');
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
