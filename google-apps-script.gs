/**
 * MULHERES CURADAS - Karoline Rodrigues
 * Recebe as inscrições da página, salva na planilha e envia
 * o e-mail de confirmação automático para a participante.
 *
 * Como instalar: veja o arquivo INSTRUCOES.md (passo a passo).
 */

// ===== CONFIGURAÇÕES (edite se quiser) =====
var NOME_EVENTO   = 'Mulheres Curadas';
var REMETENTE     = 'Mulheres Curadas';        // nome que aparece como remetente
var NOME_ABA      = 'Inscrições';               // aba onde as inscrições são salvas
// ===========================================

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // evita duas inscrições gravarem ao mesmo tempo

    var dados = JSON.parse(e.postData.contents);

    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName(NOME_ABA) || ss.insertSheet(NOME_ABA);

    // cria o cabeçalho na primeira vez
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

    // envia o e-mail de confirmação para a participante
    if (dados.email) {
      enviarConfirmacao(dados.nome, dados.email);
    }

    return resposta({ ok: true });
  } catch (err) {
    return resposta({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function enviarConfirmacao(nome, email) {
  var primeiroNome = (nome || '').split(' ')[0] || '';
  var assunto = 'Inscrição confirmada 💖 ' + NOME_EVENTO;

  var html =
    '<div style="font-family:Arial,sans-serif;max-width:520px;margin:auto;background:#120a17;color:#f4eef6;border-radius:16px;overflow:hidden">' +
      '<div style="padding:28px 26px;background:linear-gradient(90deg,#ec2a8f,#ff4fa3,#f6b26b);text-align:center">' +
        '<h1 style="margin:0;font-size:24px;color:#fff">Inscrição confirmada!</h1>' +
      '</div>' +
      '<div style="padding:28px 26px">' +
        '<p style="font-size:16px">Olá ' + primeiroNome + ', tudo bem? 💖</p>' +
        '<p style="font-size:15px;line-height:1.6;color:#d9cfe0">' +
          'Recebemos a sua inscrição para o <b style="color:#ff4fa3">' + NOME_EVENTO + '</b>. ' +
          'Está tudo certo! Em breve enviaremos os próximos detalhes por aqui e pelo seu WhatsApp.' +
        '</p>' +
        '<p style="font-size:15px;line-height:1.6;color:#d9cfe0">Com carinho,<br>Karoline Rodrigues</p>' +
      '</div>' +
    '</div>';

  MailApp.sendEmail({
    to: email,
    subject: assunto,
    htmlBody: html,
    name: REMETENTE
  });
}

// resposta em JSON
function resposta(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// permite abrir a URL no navegador para testar (mostra "funcionando")
function doGet() {
  return ContentService.createTextOutput('OK - servico de inscricoes ativo.');
}
