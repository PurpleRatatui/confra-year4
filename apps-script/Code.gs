// Confra Year 4 — recebe as respostas da página e grava na planilha.
// Cole este arquivo inteiro em Extensões > Apps Script da planilha de respostas.

const ABA = 'Respostas';
const CABECALHO = ['Data/hora', 'Aluno(a)', 'Resposta', 'Adultos', 'Crianças', 'Total', 'Alergias / restrições'];

function aba_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA);
  if (!sh) sh = ss.insertSheet(ABA);
  if (sh.getLastRow() === 0) {
    sh.appendRow(CABECALHO);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, CABECALHO.length).setFontWeight('bold');
  }
  return sh;
}

// Evita que um texto começando com = + - @ vire fórmula na planilha.
function texto_(v, max) {
  let s = String(v == null ? '' : v).trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function numero_(v) {
  const n = parseInt(v, 10);
  return isNaN(n) ? 0 : Math.max(0, Math.min(10, n));
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const aluno = texto_(d.aluno, 80);
    const rotulos = { sim: 'Vai', talvez: 'Talvez', nao: 'Não vai' };
    const resposta = rotulos[d.status];
    if (!aluno || !resposta) return json_({ ok: false, erro: 'dados incompletos' });

    const vai = d.status !== 'nao';
    const adultos = vai ? numero_(d.adultos) : 0;
    const criancas = vai ? numero_(d.criancas) : 0;
    const alergias = vai ? texto_(d.alergias, 200) : '';

    aba_().appendRow([new Date(), aluno, resposta, adultos, criancas, adultos + criancas, alergias]);
    const r = resumo_();
    r.ok = true;
    return json_(r);
  } catch (err) {
    return json_({ ok: false, erro: 'falha ao gravar' });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_(resumo_());
}

// Conta só a resposta mais recente de cada aluno(a), para quem respondeu de novo.
function resumo_() {
  const linhas = aba_().getDataRange().getValues().slice(1);
  const ultima = {};
  linhas.forEach(function (l) {
    const chave = String(l[1]).replace(/^'/, '').trim().toLowerCase();
    if (chave) ultima[chave] = l;
  });
  let pessoas = 0, familias = 0, talvez = 0;
  Object.keys(ultima).forEach(function (k) {
    const l = ultima[k];
    if (l[2] === 'Vai') { pessoas += Number(l[5]) || 0; familias++; }
    else if (l[2] === 'Talvez') { talvez += Number(l[5]) || 0; }
  });
  return { pessoas: pessoas, familias: familias, talvez: talvez };
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
