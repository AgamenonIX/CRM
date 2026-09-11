// ============================================================
// Cole este código no editor de Apps Script (Extensões → Apps Script
// na sua planilha). Depois implante como Web App (veja o passo a
// passo que te passei) e cole a URL gerada no index.html.
// ============================================================

// Deixe em branco por enquanto — quando souber o e-mail do SESMT,
// coloque aqui entre aspas (ex: "sesmt@empresa.com.br"). Enquanto
// estiver em branco, o script só grava na planilha e NÃO manda e-mail.
const EMAIL_SESMT = "";

function doPost(e) {
  try {
    const item = JSON.parse(e.postData.contents);
    const d = item.dados;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Cria o cabeçalho na primeira execução, se a planilha estiver vazia.
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Data/Hora Registro', 'Risco Crítico', 'Contrato', 'Data', 'Hora', 'Local',
        'Empresa', 'G.A.', 'Coordenação', 'Gestor Vale', 'Gestor do Contrato',
        'Supervisor', 'Encarregado', 'Fiscal', 'Descrição da Ocorrência', 'Ação Imediata'
      ]);
    }

    sheet.appendRow([
      new Date(), d.risco_critico, d.contrato, d.data, d.hora, d.local,
      d.empresa, d.ga, d.coordenacao, d.gestor_vale, d.gestor_contrato,
      d.supervisor, d.encarregado, d.fiscal, d.descricao_ocorrencia, d.acao_imediata
    ]);

    // Manda o e-mail pro SESMT só se o endereço já estiver configurado acima.
    if (EMAIL_SESMT) {
      const anexos = [];
      if (item.colagem_base64) {
        const partes = item.colagem_base64.split(',');
        const bytes = Utilities.base64Decode(partes[partes.length - 1]);
        anexos.push(Utilities.newBlob(bytes, 'image/jpeg', 'fotos_' + item.localId + '.jpg'));
      }
      MailApp.sendEmail({
        to: EMAIL_SESMT,
        subject: 'Ocorrência/RAC — ' + d.risco_critico,
        body: montarTextoResumo(d),
        attachments: anexos
      });
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, erro: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function montarTextoResumo(d) {
  return 'Risco Crítico: ' + d.risco_critico + '\n' +
    'Contrato: ' + d.contrato + '\n' +
    'Data: ' + d.data + '\n' +
    'Hora: ' + d.hora + '\n' +
    'Local: ' + (d.local || '-') + '\n' +
    'Empresa: ' + d.empresa + '\n' +
    'G.A.: ' + d.ga + '\n' +
    'Coordenação: ' + d.coordenacao + '\n' +
    'Gestor Vale: ' + d.gestor_vale + '\n' +
    'Gestor do Contrato: ' + d.gestor_contrato + '\n' +
    'Supervisor: ' + d.supervisor + '\n' +
    'Encarregado: ' + d.encarregado + '\n' +
    'Fiscal: ' + d.fiscal + '\n\n' +
    'Descrição da Ocorrência:\n' + (d.descricao_ocorrencia || '-') + '\n\n' +
    'Ação Imediata Tomada:\n' + (d.acao_imediata || '-');
}
