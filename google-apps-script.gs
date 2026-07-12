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
var DATA_EVENTO    = '1º de agosto de 2026 (sábado), às 18h';
var LOCAL_EVENTO   = 'Av. Coronel Miguel Dias, 1404 — Guararapes, Fortaleza – CE, 60810-160';
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
    '<div style="font-family:Arial,sans-serif;max-width:520px;margin:auto;background:#120a17;color:#f4eef6;border-radius:16px;overflow:hidden">' +
      '<div style="padding:28px 26px;background:linear-gradient(90deg,#ec2a8f,#ff4fa3,#f6b26b);text-align:center">' +
        '<h1 style="margin:0;font-size:24px;color:#fff">Inscrição confirmada!</h1>' +
      '</div>' +
      '<div style="padding:28px 26px">' +
        '<p style="font-size:16px">Olá ' + primeiroNome + ', tudo bem? 💖</p>' +
        '<p style="font-size:15px;line-height:1.6;color:#d9cfe0">' +
          'Recebemos a sua inscrição para o <b style="color:#ff4fa3">' + NOME_EVENTO + '</b>. Está tudo certo! Anota aí:' +
        '</p>' +
        '<div style="background:#241528;border:1px solid rgba(236,42,143,.4);border-radius:12px;padding:16px 20px;margin:16px 0">' +
          '<p style="margin:6px 0;font-size:15px">📅 <b style="color:#ff4fa3">' + DATA_EVENTO + '</b></p>' +
          '<p style="margin:6px 0;font-size:14px;line-height:1.5;color:#d9cfe0">📍 ' + LOCAL_EVENTO + '</p>' +
          '<p style="margin:6px 0;font-size:14px;color:#f6b26b">⏰ Iniciaremos <b>pontualmente</b> — não se atrase!</p>' +
        '</div>' +
        '<p style="font-size:15px;line-height:1.6;color:#d9cfe0">Com carinho,<br>Karoline Rodrigues</p>' +
      '</div>' +
    '</div>';
  MailApp.sendEmail({ to: email, subject: 'Inscrição confirmada 💖 ' + NOME_EVENTO, htmlBody: html, name: REMETENTE });
}

// aviso para a organizadora a cada nova inscricao
function avisarOrganizadora(dados, total, urlPlanilha) {
  var assunto = '🔔 Nova inscrição: ' + (dados.nome || 'sem nome') + ' (total: ' + total + ')';
  var html =
    '<div style="font-family:Arial,sans-serif;max-width:480px;margin:auto">' +
      '<h2 style="color:#ec2a8f">Nova inscrição no ' + NOME_EVENTO + ' 💖</h2>' +
      '<table style="border-collapse:collapse;width:100%;font-size:14px">' +
        linha('Nome', dados.nome) +
        linha('WhatsApp', dados.whatsapp) +
        linha('E-mail', dados.email) +
        linha('Grupo', dados.grupo) +
      '</table>' +
      '<p style="font-size:16px;margin-top:16px">Total de inscritas: <b style="color:#ec2a8f">' + total + '</b></p>' +
      '<p><a href="' + urlPlanilha + '" style="color:#ec2a8f">Abrir a planilha completa</a></p>' +
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
