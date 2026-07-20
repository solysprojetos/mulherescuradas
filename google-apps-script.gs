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
      dados.nome     || '',
      dados.whatsapp || '',
      dados.email    || '',
      dados.grupo    || ''
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

// ===== Limpa testes e formata as datas (rode 1x) =====
function limparTestesEFormatar() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName('Inscrições') || ss.getSheets()[0];
  var testes = ['gustavo teste', 'teste', 'solys projetos', 'teste cores e local'];
  var valores = aba.getDataRange().getValues();
  for (var i = valores.length - 1; i >= 1; i--) {
    var nome = String(valores[i][1] || '').toLowerCase().trim();
    if (testes.indexOf(nome) !== -1) aba.deleteRow(i + 1);
  }
  var ult = aba.getLastRow();
  if (ult > 1) aba.getRange(2, 1, ult - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');
}

// ===== DISPARO: e-mail com a arte da semana =====
function enviarBomInicioDeSemana() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName(NOME_ABA) || ss.getSheets()[0];
  var valores = aba.getDataRange().getValues();
  var IMG_URL = 'https://mulherescuradas.institutoabner.com.br/semana-mulheres-curadas.jpg';
  var imagem  = UrlFetchApp.fetch(IMG_URL).getBlob().setName('semana.jpg');
  var assunto = '💖 Que Deus abençoe a sua semana!';
  var testes  = ['gustavo teste', 'teste', 'solys projetos', 'teste cores e local'];
  var html =
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:auto;text-align:center">' +
      '<img src="cid:semana" style="width:100%;display:block;border-radius:12px" alt="Mulheres Curadas">' +
      '<p style="font-size:14px;color:#8a7168;margin-top:12px">Com carinho, Mulheres Curadas 💖</p>' +
    '</div>';
  function enviar(dest) {
    MailApp.sendEmail({ to: dest, subject: assunto, htmlBody: html, name: 'Mulheres Curadas', inlineImages: { semana: imagem } });
  }
  var jaEnviei = {}, total = 0;
  enviar(EMAIL_AVISO);
  jaEnviei[EMAIL_AVISO.toLowerCase()] = true;
  for (var i = 1; i < valores.length; i++) {
    var nome  = String(valores[i][1] || '').trim();
    var email = String(valores[i][3] || '').trim();
    var chave = email.toLowerCase();
    if (!email || email.indexOf('@') === -1) continue;
    if (testes.indexOf(nome.toLowerCase()) !== -1) continue;
    if (jaEnviei[chave]) continue;
    jaEnviei[chave] = true;
    enviar(email);
    total++;
    Utilities.sleep(300);
  }
  Logger.log('Disparo enviado para ' + total + ' inscritas (+ copia para voce).');
}
