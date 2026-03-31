/* ========================================================
   app.js — RedispAro v3.0
   Lógica principal + melhorias modernas
   ======================================================== */

/* ===== PARTICLES ===== */
(function () {
  const canvas = document.getElementById('particles-canvas');
  const ctx = canvas.getContext('2d');
  let W, H, particles = [];

  function resize() {
    W = canvas.width = window.innerWidth;
    H = canvas.height = window.innerHeight;
  }

  class Particle {
    constructor() { this.reset(); }
    reset() {
      this.x = Math.random() * W;
      this.y = Math.random() * H;
      this.size = Math.random() * 1.5 + 0.3;
      this.speedX = (Math.random() - 0.5) * 0.3;
      this.speedY = (Math.random() - 0.5) * 0.3;
      this.opacity = Math.random() * 0.4 + 0.05;
      this.color = Math.random() > 0.5 ? '56,120,255' : '0,229,255';
    }
    update() {
      this.x += this.speedX;
      this.y += this.speedY;
      if (this.x < 0 || this.x > W || this.y < 0 || this.y > H) this.reset();
    }
    draw() {
      ctx.beginPath();
      ctx.arc(this.x, this.y, this.size, 0, Math.PI * 2);
      ctx.fillStyle = `rgba(${this.color},${this.opacity})`;
      ctx.fill();
    }
  }

  function init() {
    resize();
    particles = Array.from({ length: 80 }, () => new Particle());
    animate();
  }

  function animate() {
    ctx.clearRect(0, 0, W, H);
    particles.forEach(p => { p.update(); p.draw(); });
    for (let i = 0; i < particles.length; i++) {
      for (let j = i + 1; j < particles.length; j++) {
        const dx = particles[i].x - particles[j].x;
        const dy = particles[i].y - particles[j].y;
        const dist = Math.sqrt(dx * dx + dy * dy);
        if (dist < 100) {
          ctx.beginPath();
          ctx.moveTo(particles[i].x, particles[i].y);
          ctx.lineTo(particles[j].x, particles[j].y);
          ctx.strokeStyle = `rgba(56,120,255,${0.08 * (1 - dist / 100)})`;
          ctx.lineWidth = 0.5;
          ctx.stroke();
        }
      }
    }
    requestAnimationFrame(animate);
  }

  window.addEventListener('resize', resize);
  init();
})();

/* ===== RIPPLE ===== */
document.addEventListener('click', function (e) {
  const btn = e.target.closest('.btn');
  if (!btn) return;
  const ripple = document.createElement('span');
  ripple.className = 'btn-ripple';
  const rect = btn.getBoundingClientRect();
  ripple.style.left = (e.clientX - rect.left - 5) + 'px';
  ripple.style.top  = (e.clientY - rect.top  - 5) + 'px';
  btn.appendChild(ripple);
  setTimeout(() => ripple.remove(), 600);
});

/* ===== PROGRESS BAR ===== */
const progressBar = document.getElementById('progress-bar');
let progressTimer = null;
let progressValue = 0;

function progressStart() {
  progressValue = 0;
  progressBar.style.opacity = '1';
  progressBar.classList.remove('done');
  progressBar.style.width = '0%';
  clearInterval(progressTimer);
  progressTimer = setInterval(() => {
    if (progressValue < 85) {
      progressValue += Math.random() * 8;
      progressBar.style.width = Math.min(progressValue, 85) + '%';
    }
  }, 120);
}

function progressDone() {
  clearInterval(progressTimer);
  progressBar.style.width = '100%';
  setTimeout(() => {
    progressBar.classList.add('done');
    setTimeout(() => {
      progressBar.style.width = '0%';
      progressBar.style.opacity = '1';
    }, 800);
  }, 350);
}

/* ===== TOAST NOTIFICATIONS ===== */
const toastContainer = document.getElementById('toast-container');

function showToast(msg, type = 'info', duration = 4000) {
  const icons = { success: '✓', error: '✕', warning: '!', info: 'i' };
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.innerHTML = `
    <div class="toast-icon">${icons[type] || 'i'}</div>
    <span style="flex:1;color:var(--ink)">${msg}</span>
    <button class="toast-close" aria-label="Fechar">×</button>
    <div class="toast-progress" style="animation-duration:${duration}ms"></div>
  `;

  const close = toast.querySelector('.toast-close');
  close.addEventListener('click', () => removeToast(toast));

  toastContainer.appendChild(toast);

  const timer = setTimeout(() => removeToast(toast), duration);
  toast._timer = timer;

  return toast;
}

function removeToast(toast) {
  clearTimeout(toast._timer);
  toast.classList.add('leaving');
  setTimeout(() => toast.remove(), 320);
}

/* ===== STAT COUNTER ANIMATION ===== */
function animateCount(el, target) {
  const start = parseInt(el.textContent) || 0;
  const duration = 600;
  const startTime = performance.now();
  function step(now) {
    const progress = Math.min((now - startTime) / duration, 1);
    const ease = 1 - Math.pow(1 - progress, 3);
    el.textContent = String(Math.round(start + (target - start) * ease));
    if (progress < 1) requestAnimationFrame(step);
  }
  requestAnimationFrame(step);
}

/* ===== CONSTANTS ===== */
const STATUS_ENVIADO  = new Set(["enviado", "enviada", "sent"]);
const STATUS_RECEBIDO = new Set(["entregue", "delivered"]);
const STATUS_SUCESSO  = new Set(["enviado", "enviada", "entregue", "sent", "delivered"]);
const COLUNAS_SAIDA   = ["Nome", "Telefone", "country code", "Tags", "E-mail", "Cnpj"];

/* ===== STATE ===== */
const state = {
  file1: null, file2: null,
  file1Rows: [], file2Rows: [],
  removedRows: [], remainingRows: [],
  outputFileName: "", outputSheetName: "",
  baseFile: null, baseRows: [],
  baseRemovedRows: [], baseRemainingRows: [],
  baseOutputFileName: "", baseOutputSheetName: "",
  vcfFile: null, vcfRows: [],
  vcfContacts: [], vcfContent: "", vcfOutputFileName: "",
  listSheets: [],
  activeListSheetId: "",
  listSheetCounter: 1,
  comparadorChart: null,
  csvFormatFile: null,
  csvFormatRows: [],
  csvFormatadoRows: [],
  csvErrosRows: [],
  csvFormatOutputFileName: '',
};

const $ = id => document.getElementById(id);

/* ===== THEME ===== */
function inicializarTema() {
  const saved  = localStorage.getItem("comparaErroTema");
  const prefer = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
  const tema   = saved || (prefer ? "dark" : "light");
  aplicarTema(tema);
  $("themeToggle").checked = tema === "light";
}

function aplicarTema(tema) {
  document.body.classList.toggle("light", tema === "light");
  $("themeLabel").textContent = tema === "light" ? "Claro" : "Escuro";
}

$("themeToggle").addEventListener("change", () => {
  const tema = $("themeToggle").checked ? "light" : "dark";
  aplicarTema(tema);
  localStorage.setItem("comparaErroTema", tema);
});

/* ===== NAV TABS ===== */
const sectionMap = {
  comparador: 'sectionComparador',
  tratamento: 'sectionTratamento',
  vcf:        'sectionVcf',
  editor:     'sectionEditor',
  historico:  'sectionHistorico',
};

function switchTab(target) {
  document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));

  const tab = document.querySelector(`.nav-tab[data-section="${target}"]`);
  if (tab) tab.classList.add('active');

  const secId = sectionMap[target];
  if (secId) {
    const sec = $(secId);
    if (sec) sec.classList.add('active');
  }

  if (target === 'historico') renderHistorico();
}

document.querySelectorAll('.nav-tab').forEach(tab => {
  tab.addEventListener('click', () => switchTab(tab.dataset.section));
});

/* ===== KEYBOARD SHORTCUTS ===== */
document.addEventListener('keydown', e => {
  if (e.ctrlKey && !e.shiftKey && !e.altKey) {
    switch (e.key) {
      case '1': e.preventDefault(); switchTab('comparador'); break;
      case '2': e.preventDefault(); switchTab('tratamento'); break;
      case '3': e.preventDefault(); switchTab('vcf');        break;
      case '4': e.preventDefault(); switchTab('editor');     break;
      case '5': e.preventDefault(); switchTab('historico');  break;
      case 'd':
      case 'D':
        e.preventDefault();
        const isLight = document.body.classList.contains('light');
        aplicarTema(isLight ? 'dark' : 'light');
        $("themeToggle").checked = !isLight;
        localStorage.setItem("comparaErroTema", isLight ? 'dark' : 'light');
        showToast(`Tema ${isLight ? 'escuro' : 'claro'} ativado`, 'info', 2000);
        break;
    }
  }
});

/* ===== DRAG & DROP MELHORADO ===== */
function setupDropZone(dropZoneId, inputId, onDrop) {
  const zone = $(dropZoneId);
  if (!zone) return;

  zone.addEventListener('dragover', e => {
    e.preventDefault();
    zone.classList.add('drag-over');
  });

  zone.addEventListener('dragleave', e => {
    if (!zone.contains(e.relatedTarget)) {
      zone.classList.remove('drag-over');
    }
  });

  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    const files = e.dataTransfer.files;
    if (files && files.length) {
      const input = $(inputId);
      if (input && input.multiple) {
        onDrop(Array.from(files));
      } else if (files[0]) {
        onDrop(files[0]);
      }
    }
  });
}

/* ===== FILE PICK HELPERS ===== */
function setupFilePick(btnId, inputId, nameId, dropZoneId, onPick) {
  $(btnId).addEventListener('click', () => $(inputId).click());
  $(inputId).addEventListener('change', e => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const nameEl = $(nameId);
    nameEl.textContent = file.name;
    nameEl.classList.remove('empty');
    if (dropZoneId) $(dropZoneId).classList.add('has-file');
    onPick(file);
  });
}

function setStatus(id, text, type = '') {
  const el = $(id);
  el.textContent = text;
  el.className = 'status-bar' + (type ? ' ' + type : '');
}

function setFileUI(nameId, dropZoneId, file) {
  const nameEl = $(nameId);
  nameEl.textContent = file.name;
  nameEl.classList.remove('empty');
  if (dropZoneId) $(dropZoneId).classList.add('has-file');
}

/* ===== COMPARADOR FILE SETUP ===== */
setupFilePick('btnFile1', 'file1', 'file1Name', 'dropZone1', file => {
  state.file1 = file;
  setStatus('statusMsg', 'Planilha 1 carregada. Selecione a 2ª planilha.');
  $('btnProcessar').disabled = !(state.file1 && state.file2);
  $('btnDownload').disabled = true;
  atualizarPreviewSaidaRedisparo();
});

setupFilePick('btnFile2', 'file2', 'file2Name', 'dropZone2', file => {
  state.file2 = file;
  setStatus('statusMsg', 'Arquivos prontos. Clique em Iniciar Operação.');
  $('btnProcessar').disabled = !(state.file1 && state.file2);
  $('btnDownload').disabled = true;
});

setupDropZone('dropZone1', 'file1', file => {
  state.file1 = file;
  setFileUI('file1Name', 'dropZone1', file);
  setStatus('statusMsg', 'Planilha 1 carregada. Selecione a 2ª planilha.');
  $('btnProcessar').disabled = !(state.file1 && state.file2);
  $('btnDownload').disabled = true;
  atualizarPreviewSaidaRedisparo();
});

setupDropZone('dropZone2', 'file2', file => {
  state.file2 = file;
  setFileUI('file2Name', 'dropZone2', file);
  setStatus('statusMsg', 'Arquivos prontos. Clique em Iniciar Operação.');
  $('btnProcessar').disabled = !(state.file1 && state.file2);
  $('btnDownload').disabled = true;
});

$('chipInput').addEventListener('input', atualizarPreviewSaidaRedisparo);

/* ===== TRATAMENTO FILE SETUP ===== */
setupFilePick('btnBaseFile', 'baseFile', 'baseFileName', 'dropZoneBase', file => {
  state.baseFile = file;
  setStatus('statusBaseMsg', 'Base carregada. Clique em Iniciar Operação.');
  $('btnTratarBase').disabled = false;
  $('btnDownloadBaseTratada').disabled = true;
});

setupDropZone('dropZoneBase', 'baseFile', file => {
  state.baseFile = file;
  setFileUI('baseFileName', 'dropZoneBase', file);
  setStatus('statusBaseMsg', 'Base carregada. Clique em Iniciar Operação.');
  $('btnTratarBase').disabled = false;
  $('btnDownloadBaseTratada').disabled = true;
});

/* ===== VCF FILE SETUP ===== */
setupFilePick('btnVcfFile', 'vcfFile', 'vcfFileName', 'dropZoneVcf', file => {
  state.vcfFile = file;
  Object.assign(state, { vcfRows: [], vcfContacts: [], vcfContent: '', vcfOutputFileName: '' });
  setStatus('statusVcfMsg', 'Planilha carregada. Clique em Converter.');
  $('btnConverterVcf').disabled = false;
  $('btnDownloadVcf').disabled = true;
  ['statVcfTotal', 'statVcfGerados', 'statVcfIgnorados'].forEach(id => $(id).textContent = '0');
});

setupDropZone('dropZoneVcf', 'vcfFile', file => {
  state.vcfFile = file;
  setFileUI('vcfFileName', 'dropZoneVcf', file);
  Object.assign(state, { vcfRows: [], vcfContacts: [], vcfContent: '', vcfOutputFileName: '' });
  setStatus('statusVcfMsg', 'Planilha carregada. Clique em Converter.');
  $('btnConverterVcf').disabled = false;
  $('btnDownloadVcf').disabled = true;
  ['statVcfTotal', 'statVcfGerados', 'statVcfIgnorados'].forEach(id => $(id).textContent = '0');
});

/* ===== CSV FORMAT FILE SETUP ===== */
setupFilePick('btnCsvFormatFile', 'csvFormatFileInput', 'csvFormatFileName', 'dropZoneCsvFormat', file => {
  state.csvFormatFile = file;
  Object.assign(state, { csvFormatRows: [], csvFormatadoRows: [], csvErrosRows: [], csvFormatOutputFileName: '' });
  setStatus('statusCsvFormatMsg', 'Planilha carregada. Clique em Formatar e Analisar.');
  $('btnFormatarCsv').disabled = false;
  $('btnDownloadCsvFormatado').disabled = true;
  ['statCsvTotal','statCsvResolvidos','statCsvErros','statCsvProntos'].forEach(id => $(id).textContent = '0');
});

setupDropZone('dropZoneCsvFormat', 'csvFormatFileInput', file => {
  state.csvFormatFile = file;
  setFileUI('csvFormatFileName', 'dropZoneCsvFormat', file);
  Object.assign(state, { csvFormatRows: [], csvFormatadoRows: [], csvErrosRows: [], csvFormatOutputFileName: '' });
  setStatus('statusCsvFormatMsg', 'Planilha carregada. Clique em Formatar e Analisar.');
  $('btnFormatarCsv').disabled = false;
  $('btnDownloadCsvFormatado').disabled = true;
  ['statCsvTotal','statCsvResolvidos','statCsvErros','statCsvProntos'].forEach(id => $(id).textContent = '0');
});

$('chkCarterizado').addEventListener('change', function () {
  $('helpCarterizado').style.display = this.checked ? '' : 'none';
});

/* ===== EDITOR LISTAS ===== */
$('btnListFiles').addEventListener('click', () => $('listFiles').click());
$('btnAddMoreLists').addEventListener('click', () => $('listFiles').click());
$('listFiles').addEventListener('change', handleListFilePick);
$('sheetPicker').addEventListener('change', handleSheetSelection);
$('sheetNameInput').addEventListener('input', handleSheetRename);
$('btnCreateEmptySheet').addEventListener('click', criarPlanilhaVazia);
$('btnAddHeader').addEventListener('click', handleAddHeader);
$('btnApplyStructureAll').addEventListener('click', aplicarEstruturaPrimeiraEmTodas);
$('btnMergeAllData').addEventListener('click', juntarDadosDeTodasPlanilhas);
$('btnDownloadActiveSheet').addEventListener('click', baixarPlanilhaAtiva);
$('btnDownloadAllSheets').addEventListener('click', baixarTodasPlanilhas);
$('headerEditorBox').addEventListener('click', handleHeaderActions);
$('headerEditorBox').addEventListener('change', handleHeaderRename);
$('listEditorTable').addEventListener('input', handleCellInput);

setupDropZone('dropZoneListas', 'listFiles', files => {
  handleListFileArray(Array.isArray(files) ? files : [files]);
});

/* ===== BOTÕES LIMPAR ===== */
$('btnClearComparador').addEventListener('click', () => {
  state.file1 = null; state.file2 = null;
  state.file1Rows = []; state.file2Rows = [];
  state.removedRows = []; state.remainingRows = [];
  $('file1Name').textContent = 'nenhum arquivo selecionado'; $('file1Name').classList.add('empty');
  $('file2Name').textContent = 'nenhum arquivo selecionado'; $('file2Name').classList.add('empty');
  $('dropZone1').classList.remove('has-file');
  $('dropZone2').classList.remove('has-file');
  $('chipInput').value = '';
  $('btnProcessar').disabled = true;
  $('btnDownload').disabled = true;
  ['statContatos','statEnviados','statRecebidos','statErros'].forEach(id => $(id).textContent = '0');
  setStatus('statusMsg', 'Selecione as duas planilhas para continuar.');
  $('metaInfo').textContent = 'Estrutura esperada: Nome, Telefone, country code, Tags, E-mail, Cnpj.';
  $('removedCount').textContent = '0 registros';
  $('remainingCount').textContent = '0 registros';
  renderTabela('removedTable', []);
  renderTabela('remainingTable', []);
  if (state.comparadorChart) { state.comparadorChart.destroy(); state.comparadorChart = null; }
  $('comparadorChartWrap').classList.add('hidden');
  showToast('Comparador limpo.', 'info', 2000);
});

$('btnClearTratamento').addEventListener('click', () => {
  state.baseFile = null; state.baseRows = [];
  state.baseRemovedRows = []; state.baseRemainingRows = [];
  $('baseFileName').textContent = 'nenhum arquivo selecionado'; $('baseFileName').classList.add('empty');
  $('dropZoneBase').classList.remove('has-file');
  $('btnTratarBase').disabled = true;
  $('btnDownloadBaseTratada').disabled = true;
  ['statBaseTotal','statDuplicados','statInvalidos','statBaseRestantes'].forEach(id => $(id).textContent = '0');
  setStatus('statusBaseMsg', 'Selecione a base para habilitar o tratamento.');
  $('metaTratamentoInfo').textContent = 'Critérios: telefone vazio, incompleto (55+DDD) e duplicados normalizados.';
  $('baseRemovedCount').textContent = '0 registros';
  $('baseRemainingCount').textContent = '0 registros';
  renderTabela('baseRemovedTable', []);
  renderTabela('baseRemainingTable', []);
  showToast('Tratamento limpo.', 'info', 2000);
});

$('btnClearVcf').addEventListener('click', () => {
  state.vcfFile = null; state.vcfRows = [];
  state.vcfContacts = []; state.vcfContent = ''; state.vcfOutputFileName = '';
  $('vcfFileName').textContent = 'nenhum arquivo selecionado'; $('vcfFileName').classList.add('empty');
  $('dropZoneVcf').classList.remove('has-file');
  $('btnConverterVcf').disabled = true;
  $('btnDownloadVcf').disabled = true;
  ['statVcfTotal','statVcfGerados','statVcfIgnorados'].forEach(id => $(id).textContent = '0');
  setStatus('statusVcfMsg', 'Selecione uma planilha para converter.');
  $('metaVcfInfo').textContent = 'Regra: necessário ao menos telefone ou e-mail para gerar contato no VCF.';
  showToast('Conversor VCF limpo.', 'info', 2000);
});

/* ===== ACTION BUTTONS ===== */
$('btnProcessar').addEventListener('click', processar);
$('btnDownload').addEventListener('click', baixarResultado);
$('btnDownloadCsv').addEventListener('click', baixarResultadoCsv);
$('btnTratarBase').addEventListener('click', tratarBaseDisparo);
$('btnDownloadBaseTratada').addEventListener('click', baixarBaseTratada);
$('btnDownloadBaseCsv').addEventListener('click', baixarBaseTratadaCsv);
$('btnConverterVcf').addEventListener('click', processarConversaoVcf);
$('btnDownloadVcf').addEventListener('click', baixarVcf);
$('btnLimparHistorico').addEventListener('click', limparHistorico);
$('btnFormatarCsv').addEventListener('click', formatarParaCSV);
$('btnDownloadCsvFormatado').addEventListener('click', () => {
  if (!state.csvFormatadoRows.length) return;
  downloadCsv(state.csvFormatadoRows, state.csvFormatOutputFileName || 'base_formatada.csv');
});
$('btnClearCsvFormat').addEventListener('click', () => {
  state.csvFormatFile = null;
  Object.assign(state, { csvFormatRows: [], csvFormatadoRows: [], csvErrosRows: [], csvFormatOutputFileName: '' });
  $('csvFormatFileName').textContent = 'nenhum arquivo selecionado';
  $('csvFormatFileName').classList.add('empty');
  $('dropZoneCsvFormat').classList.remove('has-file');
  $('btnFormatarCsv').disabled = true;
  $('btnDownloadCsvFormatado').disabled = true;
  $('avisoRevisao').classList.add('hidden');
  $('gridCsvFormat').style.display = 'none';
  ['statCsvTotal','statCsvResolvidos','statCsvErros','statCsvProntos'].forEach(id => $(id).textContent = '0');
  setStatus('statusCsvFormatMsg', 'Selecione uma planilha para formatar.');
  showToast('Formatador CSV limpo.', 'info', 2000);
});

/* ===== PROCESSAR COMPARADOR ===== */
async function processar() {
  if (!state.file1 || !state.file2) return;
  setStatus('statusMsg', 'Lendo planilhas...');
  $('btnProcessar').classList.add('btn-loading');
  $('btnProcessar').disabled = true;
  progressStart();

  try {
    const [p1, p2] = await Promise.all([readExcelFile(state.file1), readExcelFile(state.file2)]);
    state.file1Rows = mapearColunasComSinonimos(p1.rows, [
      { canonica: "Telefone", obrigatoria: true,  padroes: ["telefone","fone","celular","whatsapp","contato"] },
      { canonica: "Nome",     obrigatoria: false, padroes: ["nome","cliente","contato nome"] },
      { canonica: "country code", obrigatoria: false, padroes: ["country code","countrycode","ddi","codigo pais","codigopais"] },
      { canonica: "Tags",     obrigatoria: false, padroes: ["tags","tag","etiqueta"] },
      { canonica: "E-mail",   obrigatoria: false, padroes: ["e-mail","email","mail"] },
      { canonica: "Cnpj",     obrigatoria: false, padroes: ["cnpj","documento","doc"] },
    ], "primeira");

    state.file2Rows = mapearColunasComSinonimos(p2.rows, [
      { canonica: "Contato",           obrigatoria: true, padroes: ["contato","telefone","fone","celular","whatsapp"] },
      { canonica: "Status de Mensagem",obrigatoria: true, padroes: ["status de mensagem","status mensagem","status","situacao","situação"] },
    ], "segunda");

    const nomeBase1  = removeExtensao(state.file1.name);
    const chip       = obterChipInformado();
    const novoNome   = montarNomeRedisparo(nomeBase1, chip);
    state.outputFileName  = `${novoNome}.xlsx`;
    state.outputSheetName = nomeAbaExcel(novoNome);

    const resultado = compararPlanilhas(state.file1Rows, state.file2Rows, novoNome);
    state.removedRows   = resultado.removidos;
    state.remainingRows = resultado.restantes;

    animateCount($('statContatos'), state.file1Rows.length);
    animateCount($('statEnviados'), resultado.stats.enviados);
    animateCount($('statRecebidos'), resultado.stats.recebidos);
    animateCount($('statErros'), resultado.stats.erros);

    renderTabela('removedTable', state.removedRows);
    renderTabela('remainingTable', state.remainingRows);
    $('removedCount').textContent   = `${state.removedRows.length} removidos`;
    $('remainingCount').textContent = `${state.remainingRows.length} restantes`;
    $('metaInfo').textContent = `Saída: ${state.outputFileName} | Aba: ${state.outputSheetName}`;
    setStatus('statusMsg', 'Processamento concluído com sucesso.', 'success');
    $('btnDownload').disabled = false;
    $('btnDownloadCsv').disabled = false;

    // Gráfico
    renderComparadorChart(resultado.stats);

    // Histórico
    salvarHistorico({
      tipo: 'comparador',
      arquivo: state.file1.name,
      registros: state.file1Rows.length,
      extras: `${resultado.stats.erros} para redisparo`,
    });

    showToast(`Processamento concluído! ${state.remainingRows.length} contatos para redisparo.`, 'success');
    progressDone();

  } catch (err) {
    setStatus('statusMsg', `Erro: ${err.message || err}`, 'error');
    showToast(`Erro: ${err.message || err}`, 'error', 6000);
    progressDone();
  } finally {
    $('btnProcessar').classList.remove('btn-loading');
    $('btnProcessar').disabled = !(state.file1 && state.file2);
  }
}

/* ===== PROCESSAR TRATAMENTO ===== */
async function tratarBaseDisparo() {
  if (!state.baseFile) return;
  setStatus('statusBaseMsg', 'Lendo base do disparo...');
  $('btnTratarBase').classList.add('btn-loading');
  $('btnTratarBase').disabled = true;
  progressStart();

  try {
    const p = await readExcelFile(state.baseFile);
    state.baseRows = mapearColunasComSinonimos(p.rows, [
      { canonica: "Telefone", obrigatoria: true, padroes: ["telefone","fone","celular","whatsapp","contato"] },
    ], "base do disparo");

    const nome = removeExtensao(state.baseFile.name);
    state.baseOutputFileName  = `${nome} tratada.xlsx`;
    state.baseOutputSheetName = nomeAbaExcel(`${nome} tratada`);

    const r = processarTratamentoBase(state.baseRows);
    state.baseRemovedRows   = r.removidos;
    state.baseRemainingRows = r.restantes;

    animateCount($('statBaseTotal'), r.stats.total);
    animateCount($('statDuplicados'), r.stats.duplicados);
    animateCount($('statInvalidos'), r.stats.invalidos);
    animateCount($('statBaseRestantes'), r.stats.restantes);

    $('baseRemovedCount').textContent   = `${r.removidos.length} removidos`;
    $('baseRemainingCount').textContent = `${r.restantes.length} restantes`;
    $('metaTratamentoInfo').textContent = `Saída: ${state.baseOutputFileName} | Aba: ${state.baseOutputSheetName}`;

    renderTabela('baseRemovedTable', state.baseRemovedRows);
    renderTabela('baseRemainingTable', state.baseRemainingRows);
    setStatus('statusBaseMsg', 'Tratamento concluído com sucesso.', 'success');
    $('btnDownloadBaseTratada').disabled = false;
    $('btnDownloadBaseCsv').disabled = false;

    salvarHistorico({
      tipo: 'tratamento',
      arquivo: state.baseFile.name,
      registros: r.stats.total,
      extras: `${r.stats.restantes} válidos, ${r.stats.duplicados} duplicados`,
    });

    showToast(`Tratamento concluído! ${r.stats.restantes} contatos válidos.`, 'success');
    progressDone();

  } catch (err) {
    setStatus('statusBaseMsg', `Erro: ${err.message || err}`, 'error');
    showToast(`Erro: ${err.message || err}`, 'error', 6000);
    progressDone();
  } finally {
    $('btnTratarBase').classList.remove('btn-loading');
    $('btnTratarBase').disabled = !state.baseFile;
  }
}

/* ===== PROCESSAR VCF ===== */
async function processarConversaoVcf() {
  if (!state.vcfFile) return;
  setStatus('statusVcfMsg', 'Lendo planilha...');
  $('btnConverterVcf').classList.add('btn-loading');
  $('btnConverterVcf').disabled = true;
  progressStart();

  try {
    const p = await readExcelFile(state.vcfFile);
    state.vcfRows = mapearColunasComSinonimos(p.rows, [
      { canonica: "Nome",         obrigatoria: false, padroes: ["nome","cliente","contato nome"] },
      { canonica: "Telefone",     obrigatoria: false, padroes: ["telefone","fone","celular","whatsapp","contato"] },
      { canonica: "country code", obrigatoria: false, padroes: ["country code","countrycode","ddi","codigo pais","codigopais"] },
      { canonica: "E-mail",       obrigatoria: false, padroes: ["e-mail","email","mail"] },
    ], "planilha VCF");

    const contatos = [];
    let ignorados = 0;
    for (let i = 0; i < state.vcfRows.length; i++) {
      const c = mapearContatoVcf(state.vcfRows[i], i + 1);
      if (!c) { ignorados++; continue; }
      contatos.push(c);
    }

    state.vcfContacts    = contatos;
    state.vcfContent     = montarVcf(contatos);
    state.vcfOutputFileName = `${removeExtensao(state.vcfFile.name)}.vcf`;

    animateCount($('statVcfTotal'), state.vcfRows.length);
    animateCount($('statVcfGerados'), contatos.length);
    animateCount($('statVcfIgnorados'), ignorados);
    $('metaVcfInfo').textContent = `Saída: ${state.vcfOutputFileName}`;
    setStatus('statusVcfMsg', 'Conversão concluída.', 'success');
    $('btnDownloadVcf').disabled = !contatos.length;

    salvarHistorico({
      tipo: 'vcf',
      arquivo: state.vcfFile.name,
      registros: state.vcfRows.length,
      extras: `${contatos.length} contatos gerados`,
    });

    showToast(`VCF gerado! ${contatos.length} contatos convertidos.`, 'success');
    progressDone();

  } catch (err) {
    setStatus('statusVcfMsg', `Erro: ${err.message || err}`, 'error');
    showToast(`Erro: ${err.message || err}`, 'error', 6000);
    progressDone();
  } finally {
    $('btnConverterVcf').classList.remove('btn-loading');
    $('btnConverterVcf').disabled = !state.vcfFile;
  }
}

/* ===== DOWNLOADS ===== */
function baixarResultado() {
  if (!state.remainingRows.length) return;
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.remainingRows), state.outputSheetName || "Resultado");
  XLSX.writeFile(wb, state.outputFileName || "redisparado.xlsx");
  showToast('Download XLSX iniciado!', 'success', 2500);
}

function baixarResultadoCsv() {
  if (!state.remainingRows.length) return;
  downloadCsv(state.remainingRows, state.outputFileName ? state.outputFileName.replace('.xlsx', '.csv') : 'redisparado.csv');
  showToast('Download CSV iniciado!', 'success', 2500);
}

function baixarBaseTratada() {
  if (!state.baseRemainingRows.length) return;
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.baseRemainingRows, { defval: "" }), state.baseOutputSheetName || "BaseTratada");
  XLSX.writeFile(wb, state.baseOutputFileName || "base tratada.xlsx");
  showToast('Download XLSX iniciado!', 'success', 2500);
}

function baixarBaseTratadaCsv() {
  if (!state.baseRemainingRows.length) return;
  downloadCsv(state.baseRemainingRows, state.baseOutputFileName ? state.baseOutputFileName.replace('.xlsx', '.csv') : 'base tratada.csv');
  showToast('Download CSV iniciado!', 'success', 2500);
}

function baixarVcf() {
  if (!state.vcfContent) return;
  const a = Object.assign(document.createElement('a'), {
    href: URL.createObjectURL(new Blob([state.vcfContent], { type: "text/vcard;charset=utf-8" })),
    download: state.vcfOutputFileName || "contatos.vcf"
  });
  document.body.appendChild(a); a.click(); a.remove();
  showToast('Download VCF iniciado!', 'success', 2500);
}

/* ===== CSV HELPER ===== */
function downloadCsv(rows, filename) {
  if (!rows.length) return;
  const headers = Object.keys(rows[0]);
  const escape  = v => `"${String(v ?? '').replace(/"/g, '""')}"`;
  const lines   = [headers.map(escape).join(',')];
  rows.forEach(row => lines.push(headers.map(h => escape(row[h])).join(',')));
  const blob = new Blob(["\uFEFF" + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
  const a    = Object.assign(document.createElement('a'), {
    href: URL.createObjectURL(blob), download: filename
  });
  document.body.appendChild(a); a.click(); a.remove();
}

/* ===== FORMATAR PARA DISPARO (CSV) ===== */
async function formatarParaCSV() {
  if (!state.csvFormatFile) return;
  const isCarterizado = $('chkCarterizado').checked;
  setStatus('statusCsvFormatMsg', 'Lendo planilha...');
  $('btnFormatarCsv').classList.add('btn-loading');
  $('btnFormatarCsv').disabled = true;
  progressStart();

  try {
    const p = await readExcelFile(state.csvFormatFile);
    const rows = p.rows;
    if (!rows.length) throw new Error('Planilha vazia ou sem dados.');

    const formatados = [];
    const erros      = [];
    let   nCorrigidos = 0;

    for (let i = 0; i < rows.length; i++) {
      const row    = rows[i];
      const linha  = i + 2; // número da linha no arquivo (1=header)
      const correcoes = [];
      const problemas = [];

      // Normaliza todas as chaves da linha para minúsculo sem acento
      const norm = {};
      for (const k of Object.keys(row)) norm[_chaveNorm(k)] = row[k];

      // Monta linha de saída preservando TODAS as colunas normalizadas
      const out = {};
      for (const k of Object.keys(norm)) out[k] = norm[k];

      // ── Telefone ──────────────────────────────────────────
      const telRaw = String(norm['telefone'] ?? norm['fone'] ?? norm['celular'] ?? norm['whatsapp'] ?? '').trim();
      const telResult = _formatarTel(telRaw);
      if (telResult.erro) {
        problemas.push(`Telefone: ${telResult.erro} (original: "${telRaw}")`);
      } else {
        if (telResult.corrigido) correcoes.push(`telefone: "${telRaw}" → "${telResult.valor}"`);
        out['telefone'] = telResult.valor;
      }

      // ── CNPJ ─────────────────────────────────────────────
      const cnpjRaw = String(norm['cnpj'] ?? '').trim();
      if (cnpjRaw) {
        const cnpjResult = _formatarCnpj(cnpjRaw);
        if (cnpjResult.erro) {
          problemas.push(`CNPJ: ${cnpjResult.erro} (original: "${cnpjRaw}")`);
        } else {
          if (cnpjResult.corrigido) correcoes.push(`cnpj: "${cnpjRaw}" → "${cnpjResult.valor}"`);
          out['cnpj'] = cnpjResult.valor;
        }
      }

      // ── Responsavel (carterizado) ─────────────────────────
      if (isCarterizado) {
        const resp = String(norm['responsavel'] ?? norm['responsável'] ?? norm['responsavel'] ?? '').trim();
        if (!resp) problemas.push('responsavel: campo vazio (obrigatório no modo carterizado)');
        else out['responsavel'] = resp;
      }

      if (problemas.length) {
        erros.push({
          _linha: linha,
          _problema: problemas.join(' | '),
          ...out,
        });
      } else {
        if (correcoes.length) {
          nCorrigidos++;
          out['_correcoes'] = correcoes.join(' | ');
        }
        formatados.push(out);
      }
    }

    state.csvFormatadoRows  = formatados;
    state.csvErrosRows      = erros;
    state.csvFormatOutputFileName = `${removeExtensao(state.csvFormatFile.name)}_formatada.csv`;

    const total    = rows.length;
    const nErros   = erros.length;
    const nProntos = formatados.length;

    animateCount($('statCsvTotal'),      total);
    animateCount($('statCsvResolvidos'), nCorrigidos);
    animateCount($('statCsvErros'),      nErros);
    animateCount($('statCsvProntos'),    nProntos);

    $('csvErrosCount').textContent     = `${nErros} registros`;
    $('csvFormatadosCount').textContent = `${nProntos} registros`;
    $('gridCsvFormat').style.display   = '';

    // Aviso de revisão
    const aviso = $('avisoRevisao');
    if (nErros > 0) {
      aviso.classList.remove('hidden', 'aviso-erro');
      aviso.classList.add('aviso-erro');
      $('avisoRevisaoMsg').textContent =
        `⚠ ${nErros} registro(s) com problemas não resolvidos foram separados. ` +
        `Corrija-os manualmente antes do envio.`;
    } else if (nCorrigidos > 0) {
      aviso.classList.remove('hidden', 'aviso-erro');
      $('avisoRevisaoMsg').textContent =
        `${nCorrigidos} registro(s) foram corrigidos automaticamente (coluna _correcoes). ` +
        `Revise antes do envio para garantir que tudo está dentro do padrão.`;
    } else {
      aviso.classList.add('hidden');
    }

    renderTabela('csvErrosTable',     erros);
    renderTabela('csvFormatadosTable', formatados);

    $('metaCsvFormatInfo').textContent =
      `Saída: ${state.csvFormatOutputFileName} · ${nProntos} prontos · ${nCorrigidos} corrigidos · ${nErros} com erro`;

    const ok = nErros === 0;
    setStatus('statusCsvFormatMsg',
      ok ? `Formatação concluída! ${nProntos} registros prontos para envio.`
         : `Formatação concluída com ${nErros} erro(s). Revise os dados apontados.`,
      ok ? 'success' : 'error');

    $('btnDownloadCsvFormatado').disabled = (nProntos === 0);
    showToast(
      ok ? `CSV formatado! ${nProntos} registros prontos.`
         : `${nErros} erro(s) não resolvido(s). Verifique a tabela.`,
      ok ? 'success' : 'error');
    progressDone();

  } catch (err) {
    setStatus('statusCsvFormatMsg', `Erro: ${err.message || err}`, 'error');
    showToast(`Erro: ${err.message || err}`, 'error', 6000);
    progressDone();
  } finally {
    $('btnFormatarCsv').classList.remove('btn-loading');
    $('btnFormatarCsv').disabled = !state.csvFormatFile;
  }
}

/** Normaliza chave de coluna: minúsculo, sem acento, sem espaços extras */
function _chaveNorm(k) {
  return String(k)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().trim().replace(/\s+/g, '_');
}

/**
 * Formata telefone para DDI+DDD+número (ex: 5521999999999).
 * Retorna { valor, corrigido } ou { erro }.
 */
function _formatarTel(raw) {
  const digits = String(raw ?? '').replace(/\D+/g, '');
  if (!digits) return { erro: 'Telefone vazio' };

  // Já começa com 55 (Brasil)
  if (digits.startsWith('55')) {
    if (digits.length === 12 || digits.length === 13) return { valor: digits, corrigido: false };
    if (digits.length < 12) return { erro: `Número muito curto após DDI 55 (${digits.length} dígitos total)` };
    return { erro: `Número muito longo (${digits.length} dígitos)` };
  }

  // 10-11 dígitos sem DDI → adiciona 55
  if (digits.length === 10 || digits.length === 11) return { valor: `55${digits}`, corrigido: true };

  // Menos de 10 → incompleto
  if (digits.length < 10) return { erro: `Número incompleto — falta DDD e/ou DDI (${digits.length} dígito(s))` };

  // Outro DDI (não 55) com comprimento razoável (10-15 dígitos) → aceita como está
  if (digits.length >= 10 && digits.length <= 15) return { valor: digits, corrigido: false };

  return { erro: `Formato não reconhecido (${digits.length} dígitos)` };
}

/**
 * Formata CNPJ para 14 dígitos com zero à esquerda.
 * Retorna { valor, corrigido } ou { erro }.
 */
function _formatarCnpj(raw) {
  const digits = String(raw ?? '').replace(/\D+/g, '');
  if (!digits) return { valor: '', corrigido: false };
  if (digits.length > 14) return { erro: `CNPJ com mais de 14 dígitos após remoção de máscara (${digits.length})` };
  const padded    = digits.padStart(14, '0');
  const corrigido = padded !== digits;
  return { valor: padded, corrigido };
}

/* ===== CHART.JS — COMPARADOR ===== */
function renderComparadorChart(stats) {
  const wrap = $('comparadorChartWrap');
  wrap.classList.remove('hidden');

  const canvas = $('comparadorChart');
  if (state.comparadorChart) {
    state.comparadorChart.destroy();
    state.comparadorChart = null;
  }

  const total = stats.enviados + stats.recebidos + stats.erros;
  if (!total) return;

  const isDark = !document.body.classList.contains('light');
  const textColor = isDark ? '#8ba0c4' : '#4a6090';

  state.comparadorChart = new Chart(canvas, {
    type: 'doughnut',
    data: {
      labels: ['Enviados', 'Entregues', 'Com Erros'],
      datasets: [{
        data: [stats.enviados, stats.recebidos, stats.erros],
        backgroundColor: [
          'rgba(255,184,77,0.8)',
          'rgba(0,229,158,0.8)',
          'rgba(255,80,112,0.8)',
        ],
        borderColor: [
          'rgba(255,184,77,1)',
          'rgba(0,229,158,1)',
          'rgba(255,80,112,1)',
        ],
        borderWidth: 1.5,
        hoverOffset: 6,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '68%',
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: ctx => ` ${ctx.label}: ${ctx.raw} (${Math.round(ctx.raw / total * 100)}%)`,
          },
          backgroundColor: 'rgba(5,9,18,0.92)',
          titleColor: '#e8f0ff',
          bodyColor: '#8ba0c4',
          borderColor: 'rgba(56,120,255,0.25)',
          borderWidth: 1,
        }
      },
    }
  });

  // Atualiza legenda manual
  $('chartLegendEnviados').textContent = stats.enviados;
  $('chartLegendEntregues').textContent = stats.recebidos;
  $('chartLegendErros').textContent = stats.erros;
}

/* ===== CORE LOGIC ===== */
function compararPlanilhas(rows1, rows2, novoNomeBase) {
  const stats = { enviados: 0, recebidos: 0, erros: 0 };
  const sucessoPorContato = new Set();
  const erroPorContato    = new Set();
  const statusPorContato  = new Map();

  for (const row of rows2) {
    const s = normalizarStatus(row["Status de Mensagem"]);
    const k = chaveComparacao(row["Contato"]);
    if (!s) continue;
    if (STATUS_ENVIADO.has(s))  stats.enviados++;
    else if (STATUS_RECEBIDO.has(s)) stats.recebidos++;
    else stats.erros++;
    if (!k) continue;
    const lista = statusPorContato.get(k) || [];
    lista.push(s);
    statusPorContato.set(k, lista);
    if (STATUS_SUCESSO.has(s)) sucessoPorContato.add(k);
    else erroPorContato.add(k);
  }

  const removidos = [], restantes = [];
  for (const row of rows1) {
    const k = chaveComparacao(row["Telefone"]);
    const base = estruturarLinhaSaida(row, novoNomeBase);
    const statusEnc = (statusPorContato.get(k) || []).join(", ");
    if (!k || !statusPorContato.has(k)) {
      removidos.push({ ...base, "_motivo_remocao": "Contato sem status no relatório", "_status_encontrados": statusEnc });
    } else if (sucessoPorContato.has(k)) {
      removidos.push({ ...base, "_motivo_remocao": "Enviado/Entregue", "_status_encontrados": statusEnc });
    } else if (erroPorContato.has(k)) {
      restantes.push(base);
    } else {
      removidos.push({ ...base, "_motivo_remocao": "Sem status de erro", "_status_encontrados": statusEnc });
    }
  }
  return { removidos, restantes, stats };
}

function processarTratamentoBase(rows) {
  const removidos = [], restantes = [];
  const vistos = new Set();
  const stats  = { total: rows.length, duplicados: 0, invalidos: 0, restantes: 0 };

  for (const row of rows) {
    const telNorm = normalizarTelefone(row["Telefone"]);
    const motivo  = classificarTelefoneInvalido(telNorm);
    if (motivo) {
      stats.invalidos++;
      removidos.push({ ...row, "_motivo_remocao": motivo, "_telefone_normalizado": telNorm });
      continue;
    }
    const chave = chaveComparacao(row["Telefone"]);
    if (!chave) {
      stats.invalidos++;
      removidos.push({ ...row, "_motivo_remocao": "Telefone sem chave", "_telefone_normalizado": telNorm });
      continue;
    }
    if (vistos.has(chave)) {
      stats.duplicados++;
      removidos.push({ ...row, "_motivo_remocao": "Duplicado", "_telefone_normalizado": telNorm });
      continue;
    }
    vistos.add(chave);
    restantes.push(row);
  }
  stats.restantes = restantes.length;
  return { removidos, restantes, stats };
}

function classificarTelefoneInvalido(t) {
  if (!t) return "Telefone vazio";
  if (/^55\d{2}$/.test(t)) return "Incompleto (55+DDD)";
  if (t.length < 10) return "Telefone incompleto";
  return "";
}

function estruturarLinhaSaida(row, novoNomeBase) {
  const out = {};
  for (const col of COLUNAS_SAIDA) out[col] = row[col] ?? "";
  out["Tags"] = novoNomeBase;
  return out;
}

function normalizarTelefone(v) { return v == null ? "" : String(v).replace(/\D+/g, ""); }
function chaveComparacao(v)    { const t = normalizarTelefone(v); return t.length >= 11 ? t.slice(-11) : t; }
function normalizarStatus(v)   { return String(v ?? "").trim().toLowerCase(); }

function mapearContatoVcf(row, idx) {
  const nome  = String(row["Nome"]   ?? "").trim();
  const email = String(row["E-mail"] ?? "").trim();
  const tBase = normalizarTelefone(row["Telefone"]);
  const cc    = normalizarTelefone(row["country code"]);
  let tel = tBase;
  if (cc && tBase && !tBase.startsWith(cc)) tel = `${cc}${tBase}`;
  if (!tel && !email) return null;
  return { nome: nome || `Contato ${idx}`, telefone: tel ? `+${tel}` : "", email };
}

function escapeVcf(v) {
  return String(v ?? "")
    .replace(/\\/g, "\\\\")
    .replace(/\n/g, "\\n")
    .replace(/;/g, "\\;")
    .replace(/,/g, "\\,");
}

function montarVcf(contatos) {
  return contatos.map(c => {
    const l = ["BEGIN:VCARD", "VERSION:3.0", `FN:${escapeVcf(c.nome)}`];
    if (c.telefone) l.push(`TEL;TYPE=CELL:${escapeVcf(c.telefone)}`);
    if (c.email)    l.push(`EMAIL;TYPE=INTERNET:${escapeVcf(c.email)}`);
    l.push("END:VCARD");
    return l.join("\r\n");
  }).join("\r\n");
}

function obterChipInformado() { return String($("chipInput").value || "").trim(); }
function montarNomeRedisparo(base, chip) { return chip ? `${base} - ${chip}` : `${base} redisparado`; }

function atualizarPreviewSaidaRedisparo() {
  if (!state.file1) return;
  const nome = montarNomeRedisparo(removeExtensao(state.file1.name), obterChipInformado());
  $("metaInfo").textContent = `Prévia: ${nome}.xlsx | Aba: ${nomeAbaExcel(nome)}`;
}

function nomeAbaExcel(t) { return String(t || "Resultado").replace(/[:\\/*\[\]]/g, "_").trim().slice(0, 31) || "Resultado"; }
function removeExtensao(n) { return n.replace(/\.[^.]+$/, ""); }

/* ===== COLUMN MAPPING ===== */
function normalizarNomeColuna(nome) {
  return String(nome ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function encontrarMelhorColuna(headers, padroes) {
  const alvos = (padroes || []).map(normalizarNomeColuna);
  let melhorHeader = "", melhorScore = -1;
  headers.forEach(header => {
    const norm = normalizarNomeColuna(header);
    let score = -1;
    alvos.forEach(alvo => {
      if (!alvo) return;
      if (norm === alvo)                              score = Math.max(score, 300);
      else if (norm.startsWith(alvo) || alvo.startsWith(norm)) score = Math.max(score, 200);
      else if (norm.includes(alvo)   || alvo.includes(norm))   score = Math.max(score, 100);
    });
    if (score > melhorScore) { melhorScore = score; melhorHeader = header; }
  });
  return melhorScore >= 0 ? melhorHeader : "";
}

function mapearColunasComSinonimos(rows, definicoes, tipo) {
  if (!rows.length) {
    const obrigatorias = definicoes.filter(d => d.obrigatoria).map(d => d.canonica);
    if (obrigatorias.length) throw new Error(`Planilha ${tipo} sem dados para localizar colunas obrigatórias: ${obrigatorias.join(", ")}.`);
    return rows;
  }
  const headers = Array.from(new Set(rows.flatMap(r => Object.keys(r))));
  const mapa = {};
  definicoes.forEach(def => { mapa[def.canonica] = encontrarMelhorColuna(headers, def.padroes || [def.canonica]); });
  const faltando = definicoes.filter(d => d.obrigatoria && !mapa[d.canonica]).map(d => d.canonica);
  if (faltando.length) throw new Error(`Não foi possível localizar coluna(s) obrigatória(s) na ${tipo} planilha: ${faltando.join(", ")}.`);
  return rows.map(row => {
    const out = { ...row };
    definicoes.forEach(def => {
      const origem = mapa[def.canonica];
      out[def.canonica] = origem ? (row[origem] ?? "") : (out[def.canonica] ?? "");
    });
    return out;
  });
}

/* ===== TABLE RENDER COM BUSCA ===== */
function renderTabela(tableId, rows) {
  const table = $(tableId);
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");
  thead.innerHTML = ""; tbody.innerHTML = "";

  if (!rows.length) {
    thead.innerHTML = "<tr><th>Sem dados</th></tr>";
    tbody.innerHTML = "<tr><td>Nenhum registro para exibir.</td></tr>";
    return;
  }

  const headers = Object.keys(rows[0]);
  const headTr  = document.createElement("tr");
  headers.forEach(h => { const th = document.createElement("th"); th.textContent = h; headTr.appendChild(th); });
  thead.appendChild(headTr);

  function renderRows(filtered) {
    tbody.innerHTML = "";
    const frag = document.createDocumentFragment();
    filtered.forEach(row => {
      const tr = document.createElement("tr");
      headers.forEach(h => { const td = document.createElement("td"); td.textContent = row[h] ?? ""; tr.appendChild(td); });
      frag.appendChild(tr);
    });
    tbody.appendChild(frag);
  }

  renderRows(rows);

  // Ligar busca ao input acima (se existir)
  const searchInput = table.closest('.card')?.querySelector('.table-search');
  if (searchInput) {
    searchInput._rows = rows;
    searchInput.oninput = function () {
      const q = this.value.toLowerCase();
      const filtered = q ? rows.filter(row =>
        Object.values(row).some(v => String(v ?? "").toLowerCase().includes(q))
      ) : rows;
      renderRows(filtered);
    };
    searchInput.value = '';
  }
}

/* ===== EDITOR LISTAS ===== */
function gerarNomeUnicoCabecalho(headers, base) {
  const nomeBase = String(base || "Nova Coluna").trim() || "Nova Coluna";
  let nome = nomeBase, idx = 2;
  const usados = new Set(headers);
  while (usados.has(nome)) { nome = `${nomeBase} (${idx})`; idx++; }
  return nome;
}

function gerarNomeUnicoPlanilha(base) {
  const nomeBase = String(base || `Planilha ${state.listSheetCounter}`).trim() || `Planilha ${state.listSheetCounter}`;
  let nome = nomeBase, idx = 2;
  const usados = new Set(state.listSheets.map(s => s.name));
  while (usados.has(nome)) { nome = `${nomeBase} (${idx})`; idx++; }
  return nome;
}

function getPlanilhaAtiva() {
  return state.listSheets.find(s => s.id === state.activeListSheetId) || null;
}

async function handleListFilePick(e) {
  const files = Array.from(e.target.files || []);
  if (!files.length) return;
  await handleListFileArray(files);
  e.target.value = "";
}

async function handleListFileArray(files) {
  setStatus('statusListMsg', 'Carregando planilhas...');
  try {
    for (const file of files) {
      const data = await readWorkbookFile(file);
      for (const aba of data.workbook.SheetNames) {
        const sheet   = data.workbook.Sheets[aba];
        const rowsRaw = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
        const headersSet = new Set();
        rowsRaw.forEach(row => Object.keys(row).forEach(h => headersSet.add(h)));
        const headers = Array.from(headersSet);
        const rows    = rowsRaw.map(row => {
          const out = {};
          headers.forEach(h => out[h] = row[h] ?? "");
          return out;
        });
        const id   = `sheet_${state.listSheetCounter++}`;
        const name = gerarNomeUnicoPlanilha(`${removeExtensao(file.name)} - ${aba}`);
        state.listSheets.push({ id, sourceFile: file.name, name, headers, rows });
        state.activeListSheetId = id;
      }
    }
    $('listFilesInfo').textContent = `${state.listSheets.length} planilha(s) carregada(s)`;
    $('listFilesInfo').classList.remove('empty');
    $('dropZoneListas').classList.add('has-file');
    atualizarEditorListas();
    setStatus('statusListMsg', 'Planilhas carregadas. Edite cabeçalhos, posições e dados.', 'success');
    showToast(`${files.length} arquivo(s) carregado(s) com sucesso.`, 'success');
  } catch (err) {
    setStatus('statusListMsg', `Erro ao carregar: ${err.message || err}`, 'error');
    showToast(`Erro ao carregar: ${err.message || err}`, 'error', 6000);
  }
}

function atualizarEditorListas() {
  const picker = $('sheetPicker');
  picker.innerHTML = '';

  if (!state.listSheets.length) {
    picker.disabled = true;
    picker.innerHTML = '<option value="">Nenhuma planilha carregada</option>';
    $('sheetNameInput').value = '';
    $('btnDownloadActiveSheet').disabled = true;
    $('btnDownloadAllSheets').disabled   = true;
    $('btnApplyStructureAll').disabled   = true;
    $('btnMergeAllData').disabled        = true;
    $('listHeaderCount').textContent     = '0 cabeçalhos';
    $('listRowCount').textContent        = '0 linhas';
    renderHeaderEditor(null);
    renderListTable(null);
    return;
  }

  if (!getPlanilhaAtiva()) state.activeListSheetId = state.listSheets[0].id;

  state.listSheets.forEach(sheet => {
    const opt = document.createElement('option');
    opt.value    = sheet.id;
    opt.textContent = `${sheet.name} (${sheet.rows.length} linhas)`;
    if (sheet.id === state.activeListSheetId) opt.selected = true;
    picker.appendChild(opt);
  });

  const active = getPlanilhaAtiva();
  picker.disabled = false;
  picker.value    = active.id;
  $('sheetNameInput').value           = active.name;
  $('btnDownloadActiveSheet').disabled = false;
  $('btnDownloadAllSheets').disabled   = false;
  $('btnApplyStructureAll').disabled   = state.listSheets.length < 2;
  $('btnMergeAllData').disabled        = state.listSheets.length < 2;
  $('listHeaderCount').textContent     = `${active.headers.length} cabeçalhos`;
  $('listRowCount').textContent        = `${active.rows.length} linhas`;
  renderHeaderEditor(active);
  renderListTable(active);
}

function handleSheetSelection(e) {
  state.activeListSheetId = e.target.value;
  atualizarEditorListas();
}

function handleSheetRename(e) {
  const sheet = getPlanilhaAtiva();
  if (!sheet) return;
  sheet.name = String(e.target.value || "").trim() || "Planilha";
  atualizarEditorListas();
}

function renderHeaderEditor(sheet) {
  const box = $('headerEditorBox');
  box.innerHTML = '';
  if (!sheet) { box.textContent = 'Nenhuma planilha ativa.'; return; }
  if (!sheet.headers.length) { box.textContent = 'Sem cabeçalhos. Adicione um novo cabeçalho.'; return; }

  const frag = document.createDocumentFragment();
  sheet.headers.forEach((header, index) => {
    const item = document.createElement('div');
    item.className = 'header-item';
    item.innerHTML = `
      <input class="text-input" data-action="rename-header" data-index="${index}" value="${escapeHtml(header)}" />
      <div class="header-actions">
        <button class="btn mini" data-action="move-left"      data-index="${index}" ${index === 0 ? 'disabled' : ''}>&lt;</button>
        <button class="btn mini" data-action="move-right"     data-index="${index}" ${index === sheet.headers.length - 1 ? 'disabled' : ''}>&gt;</button>
        <button class="btn mini" data-action="remove-header"  data-index="${index}">X</button>
      </div>
    `;
    frag.appendChild(item);
  });
  box.appendChild(frag);
}

function handleHeaderActions(e) {
  const btn   = e.target.closest('button[data-action]');
  if (!btn) return;
  const sheet = getPlanilhaAtiva();
  if (!sheet) return;
  const idx   = Number(btn.dataset.index);
  if (Number.isNaN(idx)) return;

  if (btn.dataset.action === 'move-left' && idx > 0) {
    const tmp = sheet.headers[idx - 1];
    sheet.headers[idx - 1] = sheet.headers[idx];
    sheet.headers[idx] = tmp;
  } else if (btn.dataset.action === 'move-right' && idx < sheet.headers.length - 1) {
    const tmp = sheet.headers[idx + 1];
    sheet.headers[idx + 1] = sheet.headers[idx];
    sheet.headers[idx] = tmp;
  } else if (btn.dataset.action === 'remove-header') {
    const removed = sheet.headers.splice(idx, 1)[0];
    sheet.rows.forEach(r => delete r[removed]);
  }
  atualizarEditorListas();
}

function handleHeaderRename(e) {
  const input = e.target.closest('input[data-action="rename-header"]');
  if (!input) return;
  const sheet = getPlanilhaAtiva();
  if (!sheet) return;
  const idx   = Number(input.dataset.index);
  if (Number.isNaN(idx) || idx < 0 || idx >= sheet.headers.length) return;
  const antigo = sheet.headers[idx];
  const novo   = String(input.value || "").trim();
  if (!novo || novo === antigo) return;
  const repetido = sheet.headers.some((h, i) => i !== idx && h === novo);
  if (repetido) {
    input.value = antigo;
    setStatus('statusListMsg', `Cabeçalho "${novo}" já existe.`, 'error');
    return;
  }
  sheet.headers[idx] = novo;
  sheet.rows.forEach(row => { row[novo] = row[antigo] ?? ""; delete row[antigo]; });
  atualizarEditorListas();
}

function handleAddHeader() {
  const sheet = getPlanilhaAtiva();
  if (!sheet) { setStatus('statusListMsg', 'Carregue uma planilha antes de adicionar cabeçalho.', 'error'); return; }
  const base   = String($('newHeaderInput').value || "").trim() || "Nova Coluna";
  const header = gerarNomeUnicoCabecalho(sheet.headers, base);
  sheet.headers.push(header);
  sheet.rows.forEach(row => row[header] = "");
  $('newHeaderInput').value = "";
  atualizarEditorListas();
}

function criarPlanilhaVazia() {
  const id   = `sheet_${state.listSheetCounter++}`;
  const name = gerarNomeUnicoPlanilha(`Planilha ${state.listSheetCounter - 1}`);
  state.listSheets.push({ id, sourceFile: "manual", name, headers: [], rows: [] });
  state.activeListSheetId = id;
  atualizarEditorListas();
  setStatus('statusListMsg', 'Planilha vazia criada. Adicione cabeçalhos e linhas.', 'success');
}

function aplicarEstruturaPrimeiraEmTodas() {
  if (state.listSheets.length < 2) { setStatus('statusListMsg', 'Carregue pelo menos duas planilhas.', 'error'); return; }
  const referencia  = state.listSheets[0];
  const headersRef  = [...referencia.headers];
  if (!headersRef.length) { setStatus('statusListMsg', 'A primeira planilha não possui cabeçalhos.', 'error'); return; }

  for (let i = 1; i < state.listSheets.length; i++) {
    const sheet = state.listSheets[i];
    const headersOriginais = [...sheet.headers];
    sheet.rows = sheet.rows.map(row => {
      const nova = {};
      headersRef.forEach((headerRef, idx) => {
        const mesmoNome   = Object.prototype.hasOwnProperty.call(row, headerRef) ? headerRef : "";
        const porPosicao  = headersOriginais[idx] || "";
        if (mesmoNome) nova[headerRef] = row[mesmoNome] ?? "";
        else if (porPosicao && Object.prototype.hasOwnProperty.call(row, porPosicao)) nova[headerRef] = row[porPosicao] ?? "";
        else nova[headerRef] = "";
      });
      return nova;
    });
    sheet.headers = [...headersRef];
  }
  atualizarEditorListas();
  setStatus('statusListMsg', 'Estrutura da primeira planilha aplicada em todas.', 'success');
  showToast('Estrutura aplicada com sucesso!', 'success', 2500);
}

function juntarDadosDeTodasPlanilhas() {
  if (state.listSheets.length < 2) { setStatus('statusListMsg', 'Carregue pelo menos duas planilhas.', 'error'); return; }
  const referencia    = state.listSheets[0];
  const headersRef    = referencia.headers.length ? [...referencia.headers] : [];
  const headersUsados = headersRef.length ? headersRef : Array.from(new Set(state.listSheets.flatMap(s => s.headers)));
  if (!headersUsados.length) { setStatus('statusListMsg', 'Não há cabeçalhos para juntar.', 'error'); return; }

  const linhasUnificadas = [];
  state.listSheets.forEach(sheet => {
    sheet.rows.forEach(row => {
      const nova = {};
      headersUsados.forEach(h => nova[h] = row[h] ?? "");
      linhasUnificadas.push(nova);
    });
  });

  const id   = `sheet_${state.listSheetCounter++}`;
  const name = gerarNomeUnicoPlanilha("Dados Unificados");
  state.listSheets.push({ id, sourceFile: "merge", name, headers: [...headersUsados], rows: linhasUnificadas });
  state.activeListSheetId = id;
  atualizarEditorListas();
  setStatus('statusListMsg', `Dados unidos em "${name}" com ${linhasUnificadas.length} linha(s).`, 'success');
  showToast(`${linhasUnificadas.length} linhas unificadas!`, 'success');
}

function renderListTable(sheet) {
  const table = $('listEditorTable');
  const thead = table.querySelector('thead');
  const tbody = table.querySelector('tbody');
  thead.innerHTML = ''; tbody.innerHTML = '';

  if (!sheet) { thead.innerHTML = '<tr><th>Sem dados</th></tr>'; tbody.innerHTML = '<tr><td>Nenhuma planilha ativa.</td></tr>'; return; }
  if (!sheet.headers.length) { thead.innerHTML = '<tr><th>Sem cabeçalhos</th></tr>'; tbody.innerHTML = '<tr><td>Adicione cabeçalhos para exibir os dados.</td></tr>'; return; }

  const headTr = document.createElement('tr');
  sheet.headers.forEach(h => { const th = document.createElement('th'); th.textContent = h; headTr.appendChild(th); });
  thead.appendChild(headTr);

  if (!sheet.rows.length) {
    const tr = document.createElement('tr');
    const td = document.createElement('td');
    td.colSpan = sheet.headers.length; td.textContent = 'Sem linhas.';
    tr.appendChild(td); tbody.appendChild(tr); return;
  }

  const frag = document.createDocumentFragment();
  sheet.rows.forEach((row, rowIndex) => {
    const tr = document.createElement('tr');
    sheet.headers.forEach(header => {
      const td = document.createElement('td');
      td.className   = 'editable-cell';
      td.contentEditable = 'true';
      td.dataset.row    = String(rowIndex);
      td.dataset.header = header;
      td.textContent    = row[header] ?? "";
      tr.appendChild(td);
    });
    frag.appendChild(tr);
  });
  tbody.appendChild(frag);
}

function handleCellInput(e) {
  const td = e.target.closest('td.editable-cell');
  if (!td) return;
  const sheet = getPlanilhaAtiva();
  if (!sheet) return;
  const rowIndex = Number(td.dataset.row);
  const header   = td.dataset.header;
  if (Number.isNaN(rowIndex) || !header || !sheet.rows[rowIndex]) return;
  sheet.rows[rowIndex][header] = td.textContent ?? "";
}

function montarRowsParaExportacao(sheet) {
  return sheet.rows.map(row => {
    const out = {};
    sheet.headers.forEach(h => out[h] = row[h] ?? "");
    return out;
  });
}

function baixarPlanilhaAtiva() {
  const sheet = getPlanilhaAtiva();
  if (!sheet) return;
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(montarRowsParaExportacao(sheet), { defval: "" }), nomeAbaExcel(sheet.name || "Planilha"));
  XLSX.writeFile(wb, `${sheet.name || "planilha"}.xlsx`);
  showToast('Planilha ativa exportada!', 'success', 2500);
}

function baixarTodasPlanilhas() {
  if (!state.listSheets.length) return;
  const wb = XLSX.utils.book_new();
  state.listSheets.forEach((sheet, i) => {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(montarRowsParaExportacao(sheet), { defval: "" }), nomeAbaExcel(sheet.name || `Planilha ${i + 1}`));
  });
  XLSX.writeFile(wb, "listas_editadas.xlsx");
  showToast('Todas as planilhas exportadas!', 'success', 2500);
}

/* ===== HISTÓRICO ===== */
const HISTORY_KEY  = 'redisparo_historico';
const HISTORY_MAX  = 10;

function salvarHistorico(entry) {
  try {
    const hist = carregarHistorico();
    hist.unshift({
      ...entry,
      id:   Date.now(),
      data: new Date().toLocaleString('pt-BR'),
    });
    localStorage.setItem(HISTORY_KEY, JSON.stringify(hist.slice(0, HISTORY_MAX)));
  } catch (e) { /* localStorage indisponível */ }
}

function carregarHistorico() {
  try {
    return JSON.parse(localStorage.getItem(HISTORY_KEY) || '[]');
  } catch { return []; }
}

function limparHistorico() {
  localStorage.removeItem(HISTORY_KEY);
  renderHistorico();
  showToast('Histórico limpo.', 'info', 2000);
}

function renderHistorico() {
  const container = $('historicoContent');
  const hist = carregarHistorico();

  if (!hist.length) {
    container.innerHTML = `
      <div class="history-empty">
        <div class="history-empty-icon">📋</div>
        <div>Nenhum processamento registrado ainda.</div>
        <div style="margin-top:6px;font-size:11px;opacity:0.6">Os processamentos serão listados aqui automaticamente.</div>
      </div>
    `;
    return;
  }

  const typeLabel = { comparador: 'Comparador', tratamento: 'Tratamento', vcf: 'VCF' };
  const typeClass = { comparador: '', tratamento: 'tratamento', vcf: 'vcf' };

  container.innerHTML = `<div class="history-grid">${
    hist.map((h, i) => `
      <div class="history-card" style="animation-delay:${i * 40}ms">
        <div class="history-card-header">
          <span class="history-type-badge ${typeClass[h.tipo] || ''}">${typeLabel[h.tipo] || h.tipo}</span>
          <span class="history-date">${h.data}</span>
        </div>
        <div class="history-filename" title="${escapeHtml(h.arquivo)}">${escapeHtml(h.arquivo)}</div>
        <div class="history-stats">
          <span class="history-stat"><strong>${h.registros}</strong> registros</span>
          ${h.extras ? `<span class="history-stat">${escapeHtml(h.extras)}</span>` : ''}
        </div>
      </div>
    `).join('')
  }</div>`;
}

/* ===== HELPERS ===== */
function escapeHtml(v) {
  return String(v ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

async function readExcelFile(file) {
  const buffer = await file.arrayBuffer();
  const wb     = XLSX.read(buffer, { type: "array" });
  const sn     = wb.SheetNames[0];
  const rows   = XLSX.utils.sheet_to_json(wb.Sheets[sn], { defval: "", raw: false });
  return { workbook: wb, sheetName: sn, rows };
}

async function readWorkbookFile(file) {
  const buffer = await file.arrayBuffer();
  const wb     = XLSX.read(buffer, { type: "array" });
  return { workbook: wb };
}

/* ===== INIT ===== */
atualizarEditorListas();
inicializarTema();
