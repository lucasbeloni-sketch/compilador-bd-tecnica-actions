/**
 * Botao "Atualizar agora": dispara o robo compilador-bd-tecnica-actions no GitHub Actions
 * (mesmo efeito do botao "Run workflow" da aba Actions).
 *
 * Instalar na planilha CCM - Planejamento - Juazeiro do Norte (Extensoes > Apps Script):
 *   1. Crie um arquivo BotaoRobo.gs e cole este codigo. Salve.
 *   2. Rode "configurarTokenGitHub" UMA VEZ (seletor de funcoes > Executar), autorize e cole
 *      o token fine-grained (so este repo, permissao Actions: Read and write).
 *   3. Na planilha: clique na imagem > tres pontinhos > Atribuir script > acionarRobo
 */

var GH_OWNER    = 'lucasbeloni-sketch';
var GH_REPO     = 'compilador-bd-tecnica-actions';
var GH_WORKFLOW = 'compilar.yml';
var GH_REF      = 'main';

var ABA_STATUS   = 'Carteira_Planejador';
var CELULA_STATUS = 'F3';
var TEXTO_RODANDO = 'Em Atualização';

/** Atribua esta funcao a imagem/botao. */
function acionarRobo() {
  var ui = SpreadsheetApp.getUi();
  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_STATUS);
  var celula = aba.getRange(CELULA_STATUS);

  // Evita fila de cliques: se ja esta rodando, nao dispara de novo.
  if (String(celula.getDisplayValue()).trim() === TEXTO_RODANDO) {
    ui.alert('O robô já está em atualização. Aguarde terminar (' + CELULA_STATUS + ' vai mostrar a data/hora do fim).');
    return;
  }

  var token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) {
    ui.alert('Token do GitHub não configurado. Rode a função "configurarTokenGitHub" no editor do Apps Script.');
    return;
  }

  var url = 'https://api.github.com/repos/' + GH_OWNER + '/' + GH_REPO +
            '/actions/workflows/' + GH_WORKFLOW + '/dispatches';
  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + token,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28'
    },
    payload: JSON.stringify({ ref: GH_REF }),
    muteHttpExceptions: true
  });

  var code = resp.getResponseCode();
  if (code === 204) {
    // O robo regrava isso no inicio e troca pela data/hora no fim.
    celula.setValue(TEXTO_RODANDO);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Robô acionado. Leva uns 4-5 minutos; ' + CELULA_STATUS + ' mostra a data/hora quando terminar.',
      'GitHub Actions', 8
    );
  } else if (code === 401 || code === 403) {
    ui.alert('GitHub recusou o token (' + code + '). Ele pode ter expirado ou estar sem a permissão "Actions: Read and write". Gere outro e rode "configurarTokenGitHub".');
  } else {
    ui.alert('Falha ao acionar o robô (HTTP ' + code + '):\n' + resp.getContentText().slice(0, 500));
  }
}

/** Rode uma vez (e de novo quando o token expirar) pra guardar o token nas Propriedades do Script. */
function configurarTokenGitHub() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Token do GitHub', 'Cole o fine-grained token (github_pat_...):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var token = r.getResponseText().trim();
  if (!token) return;
  PropertiesService.getScriptProperties().setProperty('GITHUB_TOKEN', token);
  ui.alert('Token salvo.');
}
