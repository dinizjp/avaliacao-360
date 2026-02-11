// IDs de campo no Google Form
const FIELD_EVALUATOR = 'entry.2063972894';  // “Seu e-mail institucional” (avaliador)
const FIELD_EVALUATEE  = 'entry.1094000360'; // “E-mail de quem você está avaliando” (avaliado)

// Base da URL do Form, até antes de preencher campos
const FORM_BASE =
  'https://docs.google.com/forms/d/e/1FAIpQLScKQELPzxOUJHwsifD3jDEl1ttsVSqL4WUC9Os5G19WtZsDUXsadw/viewform?usp=pp_url&';

function doGet() {
  const ss     = SpreadsheetApp.getActive();
  const partSh = ss.getSheetByName('Participants');
  const respSh = ss.getSheetByName('RawResponses');

  // --- carregar e normalizar respostas existentes ---
  const respData = [];
  if (respSh.getLastRow() > 1) {
    // lê em um único getRange: col B = avaliado, col C = avaliador
    const rawResponses = respSh.getRange(2, 2, respSh.getLastRow() - 1, 2).getValues();
    rawResponses.forEach(row => {
      const evaluatee = row[0].toString().trim().toLowerCase();
      const evaluator = row[1].toString().trim().toLowerCase();
      respData.push(evaluator + '|' + evaluatee);
    });
  }
  const doneSet = new Set(respData);

  // --- carregar pares e construir mapa de avaliadores com pendentes ---
  const lastP = partSh.getLastRow();
  const rawPairs = lastP > 1
    ? partSh.getRange(2, 1, lastP - 1, 4).getValues()
    : [];
  const pairs = rawPairs.map(p => ({
    evaluator:    p[0].toString().trim().toLowerCase(),
    evaluatorName:p[1],
    evaluatee:    p[2].toString().trim().toLowerCase(),
    evaluateeName:p[3]
  }));

  // monta lista de avaliadores que ainda têm alguém pendente
  const evalStatus = {}; // email -> { name, hasPending }
  pairs.forEach(p => {
    if (!evalStatus[p.evaluator]) {
      evalStatus[p.evaluator] = { name: p.evaluatorName || p.evaluator, hasPending: false };
    }
    const key = p.evaluator + '|' + p.evaluatee;
    if (!doneSet.has(key)) {
      evalStatus[p.evaluator].hasPending = true;
    }
  });
  const evaluators = Object.entries(evalStatus)
    .filter(([_, info]) => info.hasPending) // só quem tem pendente
    .map(([email, info]) => ({ email, name: info.name }));

  // --- monta HTML ---
  const html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <style>
    body { font-family: sans-serif; background: #f4f5f7; margin: 0;
           display: flex; align-items: center; justify-content: center; height: 100vh; }
    .box { background: #fff; padding: 24px; border-radius: 8px;
           box-shadow: 0 2px 8px rgba(0,0,0,0.1);
           width: 90%; max-width: 380px; text-align: center; }
    h1 { margin: 0 0 12px; color: #333; }
    select, button { width: 100%; padding: 10px; font-size: 1em; margin: 8px 0;
                     border-radius: 4px; border: 1px solid #ccc; }
    button { background: #0066cc; color: #fff; border: none; cursor: pointer; }
    button:disabled { background: #99bfe6; cursor: default; }
    button:hover:not(:disabled) { background: #005bb5; }
    .small { font-size: 0.9em; color: #555; margin-top: 4px; }
  </style>
</head>
<body>
  <div class="box">
    <h1>Avaliação 360º</h1>
    <p>Selecione seu nome (avaliador):</p>
    ${evaluators.length === 0
      ? `<p><b>Todos os avaliadores já completaram suas avaliações.</b></p>`
      : `
        <select id="me" onchange="clearPending()">
          <option value="">— Selecione —</option>
          ${evaluators.map(e => `<option value="${e.email}">${e.name}</option>`).join('')}
        </select>
        <button id="btn" onclick="loadPending()">Carregar pendentes</button>
        <div id="pending"></div>
        <p class="small">Somente aparecem no dropdown quem ainda tem avaliações pendentes.</p>
      `}
  </div>
  <script>
    const btn = document.getElementById('btn');
    function clearPending() {
      const cont = document.getElementById('pending');
      if (cont) cont.innerHTML = '';
    }
    function loadPending() {
      const meEl = document.getElementById('me');
      if (!meEl) return;
      const me = meEl.value;
      if (!me) return alert('Selecione seu nome antes de continuar');
      btn.disabled = true;
      btn.textContent = 'Carregando...';
      google.script.run
        .withSuccessHandler(items => {
          showPending(items);
          btn.disabled = false;
          btn.textContent = 'Carregar pendentes';
        })
        .withFailureHandler(err => {
          alert('Erro ao carregar pendentes');
          console.error(err);
          btn.disabled = false;
          btn.textContent = 'Carregar pendentes';
        })
        .getPending(me);
    }
    function showPending(items) {
      const container = document.getElementById('pending');
      if (!container) return;
      if (!items || items.length === 0) {
        container.innerHTML = '<p><b>Você já avaliou todas as pessoas!</b></p>';
        return;
      }
      let html = '<p>Quem você deseja avaliar agora?</p>'
               + '<select id="sel"><option value="">— Selecione —</option>';
      items.forEach(item => {
        html += \`<option value="\${item.evaluatee}">\${item.evaluateeName}</option>\`;
      });
      html += '</select><button onclick="go()">Avaliar</button>';
      container.innerHTML = html;
    }
    function go() {
      const evaluator = document.getElementById('me')?.value;
      const evaluatee = document.getElementById('sel')?.value;
      if (!evaluatee) return alert('Selecione alguém para avaliar');
      if (!evaluator) return alert('Selecione seu nome');
      const url = '${FORM_BASE}'
                + '${FIELD_EVALUATOR}=' + encodeURIComponent(evaluator)
                + '&${FIELD_EVALUATEE}=' + encodeURIComponent(evaluatee);
      window.open(url, '_blank');
    }
  </script>
</body>
</html>`;

  return HtmlService
    .createHtmlOutput(html)
    .setTitle('Avaliação 360º')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Lê ao vivo planilha e retorna lista de {evaluatee, evaluateeName} 
 * para o avaliador selecionado.
 */
function getPending(me) {
  const ss     = SpreadsheetApp.getActive();
  const partSh = ss.getSheetByName('Participants');
  const respSh = ss.getSheetByName('RawResponses');

  // Normaliza pares
  const lastP = partSh.getLastRow();
  const rawPairs = lastP > 1
    ? partSh.getRange(2,1,lastP-1,4).getValues()
    : [];
  const pairs = rawPairs.map(p=>({
    evaluator: p[0].toString().trim().toLowerCase(),
    evaluatee: p[2].toString().trim().toLowerCase(),
    name:      p[3]
  }));

  // Normaliza respostas com uma única chamada
  const respData = [];
  if (respSh.getLastRow() > 1) {
    const rawResponses = respSh.getRange(2,2,respSh.getLastRow()-1,2).getValues(); // [avaliado, avaliador]
    rawResponses.forEach(row => {
      const evaluatee = row[0].toString().trim().toLowerCase();
      const evaluator = row[1].toString().trim().toLowerCase();
      respData.push(evaluator + '|' + evaluatee);
    });
  }
  const doneSet = new Set(respData);
  const meKey   = me.trim().toLowerCase();

  // Filtra quem falta
  return pairs
    .filter(p => p.evaluator === meKey && !doneSet.has(p.evaluator + '|' + p.evaluatee))
    .map(p => ({ evaluatee: p.evaluatee, evaluateeName: p.name }));
}
