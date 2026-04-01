import mammoth from 'https://esm.sh/mammoth@1.8.0?bundle';
import TurndownService from 'https://esm.sh/turndown@7.2.0';
import { gfm } from 'https://esm.sh/turndown-plugin-gfm@1.0.2';
import { marked } from 'https://esm.sh/marked@13.0.2';
import DOMPurify from 'https://esm.sh/dompurify@3.1.6';

if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./sw.js').catch(() => {
      // App remains fully functional without offline support.
    });
  });
}

const MAX_FILE_SIZE_MB = 12;
const IMAGE_PLACEHOLDER = '[Hinweis: Im Originaldokument war hier ein Bild. Bilder werden nicht in Markdown umgewandelt.]';

const dom = {
  input: document.getElementById('docxInput'),
  dropzone: document.getElementById('dropzone'),
  fileMeta: document.getElementById('fileMeta'),
  fileName: document.getElementById('fileName'),
  fileSize: document.getElementById('fileSize'),
  status: document.getElementById('status'),
  clearBtn: document.getElementById('clearBtn'),
  copyBtn: document.getElementById('copyBtn'),
  previewBtn: document.getElementById('previewBtn'),
  downloadBtn: document.getElementById('downloadBtn'),
  output: document.getElementById('markdownOutput'),
  preview: document.getElementById('markdownPreview')
};

const state = {
  selectedFile: null,
  markdown: '',
  isPreviewMode: false
};

const UTF8_BOM = '\uFEFF';

const turndownService = new TurndownService({
  headingStyle: 'atx',
  hr: '---',
  bulletListMarker: '-',
  codeBlockStyle: 'fenced',
  emDelimiter: '*'
});

turndownService.use(gfm);

turndownService.addRule('figureToImageNote', {
  filter: ['img'],
  replacement() {
    return `\n${IMAGE_PLACEHOLDER}\n`;
  }
});

setupEvents();
setupEasterEgg();

function setupEasterEgg() {
  let clickCount = 0;
  let resetTimer = null;

  function onEasterClick(btn) {
    if (!btn.disabled) return;
    clickCount++;
    clearTimeout(resetTimer);
    resetTimer = setTimeout(() => {
      clickCount = 0;
    }, 2000);
    if (clickCount >= 5) {
      clickCount = 0;
      triggerEggRain();
    }
  }

  [dom.clearBtn, dom.copyBtn, dom.previewBtn, dom.downloadBtn].forEach((btn) => {
    btn.addEventListener('pointerdown', () => onEasterClick(btn));
  });
}

function triggerEggRain() {
  const symbols = ['🥚', '🐣', '🐰', '🌷', '🌸', '🌺', '🦋'];
  for (let i = 0; i < 28; i++) {
    setTimeout(() => {
      const el = document.createElement('span');
      el.textContent = symbols[Math.floor(Math.random() * symbols.length)];
      el.className = 'egg-rain-item';
      const rotStart = Math.random() * 30 - 15;
      const rotEnd = Math.random() * 360 - 180;
      el.style.cssText = `
        left: ${Math.random() * 100}vw;
        font-size: ${1.2 + Math.random() * 1.8}rem;
        --duration: ${1.6 + Math.random() * 1.8}s;
        --rot-start: ${rotStart}deg;
        --rot-end: ${rotEnd}deg;
      `;
      document.body.appendChild(el);
      el.addEventListener('animationend', () => el.remove());
    }, i * 90);
  }
  setStatus('🐣 Frohe Ostern! 🥚');
}

function setupEvents() {
  dom.input.addEventListener('change', onInputChange);
  dom.clearBtn.addEventListener('click', onClear);
  dom.copyBtn.addEventListener('click', onCopy);
  dom.previewBtn.addEventListener('click', onTogglePreview);
  dom.downloadBtn.addEventListener('click', onDownload);

  ['dragenter', 'dragover'].forEach((eventName) => {
    dom.dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dom.dropzone.classList.add('is-dragover');
    });
  });

  ['dragleave', 'drop'].forEach((eventName) => {
    dom.dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dom.dropzone.classList.remove('is-dragover');
    });
  });

  dom.dropzone.addEventListener('drop', async (event) => {
    const file = event.dataTransfer?.files?.[0];
    if (file) {
      await selectFile(file);
    }
  });
}

async function onInputChange(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }
  await selectFile(file);
}

async function selectFile(file) {
  const lowerName = file.name.toLowerCase();
  const isDocx = lowerName.endsWith('.docx');
  const isWithinSizeLimit = file.size <= MAX_FILE_SIZE_MB * 1024 * 1024;

  if (!isDocx) {
    setError('Nur .docx Dateien werden unterstützt.');
    return;
  }

  if (!isWithinSizeLimit) {
    setError(`Datei ist zu groß. Maximal ${MAX_FILE_SIZE_MB} MB erlaubt.`);
    return;
  }

  state.selectedFile = file;
  state.markdown = '';
  state.isPreviewMode = false;
  dom.output.value = '';
  dom.preview.innerHTML = '';

  dom.fileMeta.hidden = false;
  dom.fileName.textContent = file.name;
  dom.fileSize.textContent = `${(file.size / (1024 * 1024)).toFixed(2)} MB`;

  dom.clearBtn.disabled = false;
  dom.copyBtn.disabled = true;
  dom.previewBtn.disabled = true;
  dom.downloadBtn.disabled = true;
  updateOutputMode();

  setStatus('Datei erkannt. Starte Konvertierung...');
  await onConvert();
}

async function onConvert() {
  if (!state.selectedFile) {
    setError('Bitte zuerst eine .docx Datei auswählen.');
    return;
  }

  try {
    setStatus('Konvertierung läuft...');
    dom.clearBtn.disabled = true;

    const arrayBuffer = await state.selectedFile.arrayBuffer();
    const result = await mammoth.convertToHtml(
      { arrayBuffer },
      {
        includeDefaultStyleMap: true,
        ignoreEmptyParagraphs: false
      }
    );

    const normalizedHtml = normalizeHtmlForMarkdown(result.value || '');
    let markdown = turndownService.turndown(normalizedHtml);
    markdown = postProcessMarkdown(markdown);

    state.markdown = markdown.trim();
    dom.output.value = state.markdown;
    renderPreview();

    dom.copyBtn.disabled = state.markdown.length === 0;
    dom.previewBtn.disabled = state.markdown.length === 0;
    dom.downloadBtn.disabled = state.markdown.length === 0;

    const warningCount = Array.isArray(result.messages) ? result.messages.length : 0;
    if (warningCount > 0) {
      setStatus(`Konvertierung erfolgreich mit ${warningCount} Hinweis(en).`);
    } else {
      setStatus('Konvertierung erfolgreich.');
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unbekannter Fehler';
    setError(`Konvertierung fehlgeschlagen: ${message}`);
  } finally {
    dom.clearBtn.disabled = !state.selectedFile;
  }
}

function postProcessMarkdown(input) {
  return input
    .normalize('NFC')
    .replace(/\u00A0/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/\t/g, '  ')
    .replace(/\n\s+\n/g, '\n\n')
    .trim();
}

function normalizeHtmlForMarkdown(rawHtml) {
  const parser = new DOMParser();
  const documentFromHtml = parser.parseFromString(rawHtml, 'text/html');
  const walker = documentFromHtml.createTreeWalker(documentFromHtml.body, NodeFilter.SHOW_TEXT);

  while (walker.nextNode()) {
    const textNode = walker.currentNode;
    if (textNode.nodeValue) {
      textNode.nodeValue = textNode.nodeValue.replace(/\u00A0/g, ' ').normalize('NFC');
    }
  }

  documentFromHtml.body.querySelectorAll('[alt],[title]').forEach((node) => {
    if (node.hasAttribute('alt')) {
      node.setAttribute('alt', (node.getAttribute('alt') || '').replace(/\u00A0/g, ' ').normalize('NFC'));
    }
    if (node.hasAttribute('title')) {
      node.setAttribute('title', (node.getAttribute('title') || '').replace(/\u00A0/g, ' ').normalize('NFC'));
    }
  });

  return documentFromHtml.body.innerHTML;
}

async function onCopy() {
  if (!state.markdown) {
    setError('Es gibt noch keinen Markdown-Inhalt zum Kopieren.');
    return;
  }

  try {
    await navigator.clipboard.writeText(state.markdown);
    setStatus('Markdown in die Zwischenablage kopiert.');
  } catch {
    dom.output.focus();
    dom.output.select();
    const successful = document.execCommand('copy');
    if (successful) {
      setStatus('Markdown in die Zwischenablage kopiert.');
    } else {
      setError('Kopieren nicht möglich. Bitte manuell markieren und kopieren.');
    }
  }
}

function onDownload() {
  if (!state.markdown) {
    setError('Es gibt noch keinen Markdown-Inhalt zum Download.');
    return;
  }

  const base = (state.selectedFile?.name || 'document')
    .replace(/\.docx$/i, '')
    .replace(/[^a-zA-Z0-9-_ ]/g, '')
    .trim() || 'document';

  const blob = new Blob([UTF8_BOM, state.markdown], { type: 'text/markdown;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `${base}.md`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);

  setStatus('Markdown-Datei heruntergeladen.');
}

function onTogglePreview() {
  if (!state.markdown) {
    setError('Es gibt noch keinen Markdown-Inhalt für die Vorschau.');
    return;
  }

  state.isPreviewMode = !state.isPreviewMode;
  if (state.isPreviewMode) {
    renderPreview();
  }
  updateOutputMode();
}

function renderPreview() {
  if (!state.markdown) {
    dom.preview.innerHTML = '';
    return;
  }

  const html = marked.parse(state.markdown, {
    breaks: true,
    gfm: true,
    mangle: false,
    headerIds: false
  });

  const safeHtml = DOMPurify.sanitize(html, {
    USE_PROFILES: { html: true }
  });

  dom.preview.innerHTML = safeHtml;
}

function updateOutputMode() {
  const showPreview = state.isPreviewMode && state.markdown.length > 0;

  dom.preview.hidden = !showPreview;
  dom.output.hidden = showPreview;

  dom.previewBtn.setAttribute('aria-pressed', showPreview ? 'true' : 'false');
  dom.previewBtn.classList.toggle('is-active', showPreview);
}

function onClear() {
  state.selectedFile = null;
  state.markdown = '';
  state.isPreviewMode = false;

  dom.input.value = '';
  dom.output.value = '';
  dom.preview.innerHTML = '';
  dom.fileMeta.hidden = true;
  dom.fileName.textContent = '-';
  dom.fileSize.textContent = '-';

  dom.clearBtn.disabled = true;
  dom.copyBtn.disabled = true;
  dom.previewBtn.disabled = true;
  dom.downloadBtn.disabled = true;
  updateOutputMode();

  setStatus('');
}

function setStatus(message) {
  dom.status.classList.remove('error');
  dom.status.textContent = message;
}

function setError(message) {
  dom.status.classList.add('error');
  dom.status.textContent = message;
}
