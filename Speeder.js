(function(){
  'use strict';

  // ---------------------------------------------------------------
  // Setup
  // ---------------------------------------------------------------
  const PDF_WORKER_URL = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  let pdfWorkerReady = null;

  // Some browsers (notably Safari/iOS) refuse to construct a Worker whose
  // script lives on a different origin than the page, even when the CDN
  // sends CORS headers. Fetching the script and loading it from a
  // same-origin blob URL avoids that restriction.
  function ensurePdfWorker(){
    if(pdfWorkerReady) return pdfWorkerReady;
    if(!window.pdfjsLib) return Promise.reject(new Error('PDF support did not load. Check your connection and try again.'));
    pdfWorkerReady = fetch(PDF_WORKER_URL)
      .then(res => {
        if(!res.ok) throw new Error('network');
        return res.text();
      })
      .then(code => {
        const blob = new Blob([code], { type: 'application/javascript' });
        pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(blob);
      })
      .catch(() => {
        // Fall back to the direct CDN URL — works in browsers that don't
        // enforce same-origin workers (Chrome, Firefox, most desktop).
        pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER_URL;
      });
    return pdfWorkerReady;
  }

  const WORDS_PER_APPROX_PAGE = 250; // used for .txt / .docx page estimation
  const LARGE_RANGE_WARNING_WORDS = 40000; // ~ soft warning threshold

  const el = (id) => document.getElementById(id);

  // Setup screen elements
  const uploadZone = el('uploadZone');
  const fileInput = el('fileInput');
  const pasteToggle = el('pasteToggle');
  const pastePanel = el('pastePanel');
  const textInput = el('textInput');
  const usePasted = el('usePasted');
  const fileInfo = el('fileInfo');
  const fileNameEl = el('fileName');
  const clearFile = el('clearFile');
  const pageRangePanel = el('pageRangePanel');
  const totalPagesEl = el('totalPages');
  const fromPageInput = el('fromPage');
  const toPageInput = el('toPage');
  const chunkWordsEl = el('chunkWords');
  const chunkWarning = el('chunkWarning');
  const startBtn = el('startBtn');
  const toast = el('toast');

  // Screens
  const setupScreen = el('setupScreen');
  const readerScreen = el('readerScreen');

  // Reader elements
  const wordDisplay = el('wordDisplay');
  const rangeLabel = el('rangeLabel');
  const btnPlay = el('btnPlay');
  const playIcon = el('playIcon');
  const btnBack = el('btnBack');
  const btnFwd = el('btnFwd');
  const speedSlider = el('speedSlider');
  const speedVal = el('speedVal');
  const wpmReadout = el('wpmReadout');
  const progressTrack = el('progressTrack');
  const progressFill = el('progressFill');
  const progressHandle = el('progressHandle');
  const stage = document.querySelector('.stage');
  const btnNewDoc = el('btnNewDoc');

  // ---------------------------------------------------------------
  // State
  // ---------------------------------------------------------------
  let pages = [];      // array of arrays-of-words, one entry per "page"
  let words = [];      // flattened words for the active reading session
  let idx = 0;
  let playing = false;
  let timer = null;
  let wpm = 300;

  // ---------------------------------------------------------------
  // Toast
  // ---------------------------------------------------------------
  let toastTimer = null;
  function showToast(msg, isError){
    toast.textContent = msg;
    toast.classList.toggle('error', !!isError);
    toast.classList.remove('hidden');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(()=> toast.classList.add('hidden'), 3200);
  }

  // ---------------------------------------------------------------
  // File parsing
  // ---------------------------------------------------------------
  function chunkTextIntoPages(fullText){
    const allWords = fullText.trim().split(/\s+/).filter(Boolean);
    const result = [];
    for(let i = 0; i < allWords.length; i += WORDS_PER_APPROX_PAGE){
      result.push(allWords.slice(i, i + WORDS_PER_APPROX_PAGE));
    }
    return result.length ? result : [[]];
  }

  function readTxt(file){
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(chunkTextIntoPages(reader.result));
      reader.onerror = () => reject(new Error('Could not read that text file.'));
      reader.readAsText(file);
    });
  }

  async function readPdf(file){
    if(!window.pdfjsLib){
      throw new Error('PDF support did not load. Check your connection and try again.');
    }
    await ensurePdfWorker();
    const buffer = await file.arrayBuffer();
    let doc;
    try{
      doc = await pdfjsLib.getDocument({ data: buffer }).promise;
    }catch(err){
      throw new Error('Could not open that PDF (' + (err && err.message ? err.message : 'unknown error') + ').');
    }
    const result = [];
    for(let p = 1; p <= doc.numPages; p++){
      const page = await doc.getPage(p);
      const content = await page.getTextContent();
      const pageText = content.items.map(it => it.str).join(' ');
      const pageWords = pageText.trim().split(/\s+/).filter(Boolean);
      result.push(pageWords);
    }
    return result.length ? result : [[]];
  }

  async function readDocx(file){
    if(!window.mammoth){
      throw new Error('Word doc support did not load. Check your connection and try again.');
    }
    const buffer = await file.arrayBuffer();
    const out = await mammoth.extractRawText({ arrayBuffer: buffer });
    return chunkTextIntoPages(out.value || '');
  }

  async function parseFile(file){
    const name = file.name.toLowerCase();
    if(name.endsWith('.pdf')) return readPdf(file);
    if(name.endsWith('.docx')) return readDocx(file);
    if(name.endsWith('.txt')) return readTxt(file);
    throw new Error('Unsupported file type. Use .txt, .pdf, or .docx.');
  }

  // ---------------------------------------------------------------
  // Page-range panel wiring
  // ---------------------------------------------------------------
  function updateChunkWordCount(){
    const from = clampPage(parseInt(fromPageInput.value, 10) || 1);
    const to = clampPage(parseInt(toPageInput.value, 10) || 1);
    const lo = Math.min(from, to);
    const hi = Math.max(from, to);
    let count = 0;
    for(let i = lo - 1; i <= hi - 1; i++){
      if(pages[i]) count += pages[i].length;
    }
    chunkWordsEl.textContent = count.toLocaleString();
    chunkWarning.classList.toggle('hidden', count < LARGE_RANGE_WARNING_WORDS);
  }

  function clampPage(n){
    return Math.max(1, Math.min(pages.length || 1, n));
  }

  function showPageRangeUI(){
    totalPagesEl.textContent = pages.length;
    fromPageInput.min = 1;
    fromPageInput.max = pages.length;
    toPageInput.min = 1;
    toPageInput.max = pages.length;
    fromPageInput.value = 1;
    toPageInput.value = pages.length;
    pageRangePanel.classList.remove('hidden');
    startBtn.classList.remove('hidden');
    updateChunkWordCount();
  }

  function resetToSetup(){
    pages = [];
    words = [];
    idx = 0;
    pause();
    pageRangePanel.classList.add('hidden');
    startBtn.classList.add('hidden');
    fileInfo.classList.add('hidden');
    pastePanel.classList.add('hidden');
    fileInput.value = '';
    textInput.value = '';
    readerScreen.classList.add('hidden');
    setupScreen.classList.remove('hidden');
  }

  // ---------------------------------------------------------------
  // Upload zone interactions
  // ---------------------------------------------------------------
  uploadZone.addEventListener('click', () => fileInput.click());
  uploadZone.addEventListener('keydown', (e) => {
    if(e.key === 'Enter' || e.key === ' '){ e.preventDefault(); fileInput.click(); }
  });
  uploadZone.setAttribute('tabindex', '0');
  uploadZone.setAttribute('role', 'button');

  ['dragenter','dragover'].forEach(evt => {
    uploadZone.addEventListener(evt, (e) => {
      e.preventDefault();
      uploadZone.classList.add('drag-over');
    });
  });
  ['dragleave','drop'].forEach(evt => {
    uploadZone.addEventListener(evt, (e) => {
      e.preventDefault();
      uploadZone.classList.remove('drag-over');
    });
  });
  uploadZone.addEventListener('drop', (e) => {
    const file = e.dataTransfer.files && e.dataTransfer.files[0];
    if(file) handleFile(file);
  });

  fileInput.addEventListener('change', () => {
    const file = fileInput.files[0];
    if(file) handleFile(file);
  });

  async function handleFile(file){
    fileNameEl.textContent = file.name;
    fileInfo.classList.remove('hidden');
    pastePanel.classList.add('hidden');
    showToast('Reading ' + file.name + '…');
    try{
      pages = await parseFile(file);
      showPageRangeUI();
      showToast(pages.length + ' pages found.');
    }catch(err){
      showToast(err.message || 'Could not read that file.', true);
      fileInfo.classList.add('hidden');
    }
  }

  clearFile.addEventListener('click', () => {
    fileInput.value = '';
    fileInfo.classList.add('hidden');
    pageRangePanel.classList.add('hidden');
    startBtn.classList.add('hidden');
    pages = [];
  });

  pasteToggle.addEventListener('click', () => {
    pastePanel.classList.toggle('hidden');
    fileInfo.classList.add('hidden');
    pageRangePanel.classList.add('hidden');
    startBtn.classList.add('hidden');
  });

  usePasted.addEventListener('click', () => {
    const text = textInput.value || '';
    if(!text.trim()){
      showToast('Paste some text first.', true);
      return;
    }
    pages = chunkTextIntoPages(text);
    fileNameEl.textContent = 'Pasted text';
    fileInfo.classList.remove('hidden');
    showPageRangeUI();
  });

  fromPageInput.addEventListener('input', updateChunkWordCount);
  toPageInput.addEventListener('input', updateChunkWordCount);

  startBtn.addEventListener('click', () => {
    const from = clampPage(parseInt(fromPageInput.value, 10) || 1);
    const to = clampPage(parseInt(toPageInput.value, 10) || 1);
    const lo = Math.min(from, to);
    const hi = Math.max(from, to);
    words = [];
    for(let i = lo - 1; i <= hi - 1; i++){
      if(pages[i]) words = words.concat(pages[i]);
    }
    if(!words.length){
      showToast('That page range is empty.', true);
      return;
    }
    idx = 0;
    rangeLabel.textContent = pages.length > 1
      ? 'pages ' + lo + '\u2013' + hi + ' of ' + pages.length
      : words.length + ' words';
    setupScreen.classList.add('hidden');
    readerScreen.classList.remove('hidden');
    showAt(0);
    play();
  });

  btnNewDoc.addEventListener('click', resetToSetup);

  // ---------------------------------------------------------------
  // RSVP engine
  // ---------------------------------------------------------------
  function pivotIndex(word){
    const len = word.replace(/[^a-zA-Z0-9]/g, '').length || word.length;
    if(len <= 1) return 0;
    if(len <= 5) return 1;
    if(len <= 9) return 2;
    if(len <= 13) return 3;
    return 4;
  }

  function escapeHtml(s){
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function renderWord(w){
    if(!w){
      wordDisplay.classList.add('placeholder');
      wordDisplay.style.transform = 'translateX(0px)';
      wordDisplay.textContent = 'Ready';
      return;
    }
    wordDisplay.classList.remove('placeholder');
    // Reset any offset from the previous word first, so the measurement
    // below reflects this word's natural (untransformed) layout position
    // rather than being thrown off by whatever shift was left over.
    wordDisplay.style.transform = 'translateX(0px)';
    const p = pivotIndex(w);
    const before = w.slice(0, p);
    const pivotChar = w.slice(p, p + 1);
    const after = w.slice(p + 1);
    wordDisplay.innerHTML =
      '<span>' + escapeHtml(before) + '</span>' +
      '<span class="pivot-char">' + escapeHtml(pivotChar) + '</span>' +
      '<span>' + escapeHtml(after) + '</span>';
    // getBoundingClientRect forces a synchronous layout, so we can measure
    // and correct in the same tick — no rAF delay, no one-frame flash.
    const stageRect = stage.getBoundingClientRect();
    const pivotEl = wordDisplay.querySelector('.pivot-char');
    const pivotRect = pivotEl.getBoundingClientRect();
    const stageCenter = stageRect.left + stageRect.width / 2;
    const pivotCenter = pivotRect.left + pivotRect.width / 2;
    wordDisplay.style.transform = 'translateX(' + (stageCenter - pivotCenter) + 'px)';
  }

  function durationFor(word){
    const base = 60000 / wpm;
    let mult = 1;
    const len = word.length;
    if(len > 6) mult += 0.25;
    if(len > 10) mult += 0.25;
    if(/[.,;:!?]$/.test(word)) mult += 0.6;
    if(word === '') mult = 0.4;
    return base * mult;
  }

  function updateProgress(){
    const pct = words.length ? (idx / words.length) * 100 : 0;
    progressFill.style.width = pct + '%';
    progressHandle.style.left = pct + '%';
  }

  function showAt(i){
    idx = Math.max(0, Math.min(i, words.length ? words.length - 1 : 0));
    renderWord(words[idx] || '');
    updateProgress();
  }

  function tick(){
    if(!playing) return;
    if(idx >= words.length){
      pause();
      return;
    }
    renderWord(words[idx]);
    updateProgress();
    const d = durationFor(words[idx]);
    idx++;
    timer = setTimeout(tick, d);
  }

  function play(){
    if(!words.length) return;
    if(idx >= words.length) idx = 0;
    playing = true;
    playIcon.innerHTML = '<path d="M6 5h4v14H6zM14 5h4v14h-4z"/>';
    tick();
  }

  function pause(){
    playing = false;
    playIcon.innerHTML = '<path d="M8 5v14l11-7z"/>';
    clearTimeout(timer);
  }

  function togglePlay(){
    playing ? pause() : play();
  }

  btnPlay.addEventListener('click', togglePlay);
  wordDisplay.addEventListener('click', togglePlay);

  btnBack.addEventListener('click', () => { pause(); showAt(idx - 6); });
  btnFwd.addEventListener('click', () => { pause(); showAt(idx + 4); });

  speedSlider.addEventListener('input', () => {
    wpm = parseInt(speedSlider.value, 10);
    speedVal.textContent = wpm;
    wpmReadout.textContent = wpm + ' WPM';
  });

  let dragging = false;
  function scrubTo(clientX){
    const rect = progressTrack.getBoundingClientRect();
    let pct = (clientX - rect.left) / rect.width;
    pct = Math.max(0, Math.min(1, pct));
    showAt(Math.round(pct * words.length));
  }
  progressTrack.addEventListener('pointerdown', (e) => {
    dragging = true;
    pause();
    scrubTo(e.clientX);
    progressTrack.setPointerCapture(e.pointerId);
  });
  progressTrack.addEventListener('pointermove', (e) => { if(dragging) scrubTo(e.clientX); });
  progressTrack.addEventListener('pointerup', () => dragging = false);
  progressTrack.addEventListener('pointercancel', () => dragging = false);

  document.addEventListener('keydown', (e) => {
    if(readerScreen.classList.contains('hidden')) return;
    if(e.code === 'Space'){ e.preventDefault(); togglePlay(); }
    if(e.code === 'ArrowRight'){ pause(); showAt(idx + 4); }
    if(e.code === 'ArrowLeft'){ pause(); showAt(idx - 6); }
  });

  if(document.fonts && document.fonts.ready){
    document.fonts.ready.then(() => {
      if(words.length) renderWord(words[idx] || '');
    });
  }

})();
