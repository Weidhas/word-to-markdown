import mammoth from 'https://esm.sh/mammoth@1.8.0?bundle';
import TurndownService from 'https://esm.sh/turndown@7.2.0';
import { gfm } from 'https://esm.sh/turndown-plugin-gfm@1.0.2';

const MAX_FILE_SIZE_MB = 12;

const dom = {
  input: document.getElementById('docxInput'),
  dropzone: document.getElementById('dropzone'),
  fileMeta: document.getElementById('fileMeta'),
  fileName: document.getElementById('fileName'),
  fileSize: document.getElementById('fileSize'),
  status: document.getElementById('status'),
  convertBtn: document.getElementById('convertBtn'),
  clearBtn: document.getElementById('clearBtn'),
  copyBtn: document.getElementById('copyBtn'),
  downloadBtn: document.getElementById('downloadBtn'),
  output: document.getElementById('markdownOutput')
};

const state = {
  selectedFile: null,
  markdown: ''
};

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
  replacement(content, node) {
    const alt = node.getAttribute('alt') || 'image';
    const src = node.getAttribute('src') || '';
    if (!src || src.startsWith('data:')) {
      return `\n![${alt}](#embedded-image)\n`;
    }
    return `\n![${alt}](${src})\n`;
  }
});

setupEvents();

function setupEvents() {
  dom.input.addEventListener('change', onInputChange);
  dom.convertBtn.addEventListener('click', onConvert);
  dom.clearBtn.addEventListener('click', onClear);
  dom.copyBtn.addEventListener('click', onCopy);
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

  dom.dropzone.addEventListener('drop', (event) => {
    const file = event.dataTransfer?.files?.[0];
    if (file) {
      selectFile(file);
    }
  });
}

function onInputChange(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }
  selectFile(file);
}

function selectFile(file) {
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
  dom.output.value = '';

  dom.fileMeta.hidden = false;
  dom.fileName.textContent = file.name;
  dom.fileSize.textContent = `${(file.size / (1024 * 1024)).toFixed(2)} MB`;

  dom.convertBtn.disabled = false;
  dom.clearBtn.disabled = false;
  dom.copyBtn.disabled = true;
  dom.downloadBtn.disabled = true;

  setStatus('Datei bereit zur Konvertierung.');
}

async function onConvert() {
  if (!state.selectedFile) {
    setError('Bitte zuerst eine .docx Datei auswählen.');
    return;
  }

  try {
    setStatus('Konvertierung läuft...');
    dom.convertBtn.disabled = true;

    const arrayBuffer = await state.selectedFile.arrayBuffer();
    const result = await mammoth.convertToHtml(
      { arrayBuffer },
      {
        includeDefaultStyleMap: true,
        ignoreEmptyParagraphs: false
      }
    );

    let markdown = turndownService.turndown(result.value || '');
    markdown = postProcessMarkdown(markdown);

    state.markdown = markdown.trim();
    dom.output.value = state.markdown;

    dom.copyBtn.disabled = state.markdown.length === 0;
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
    dom.convertBtn.disabled = !state.selectedFile;
  }
}

function postProcessMarkdown(input) {
  return input
    .replace(/\n{3,}/g, '\n\n')
    .replace(/\t/g, '  ')
    .replace(/\n\s+\n/g, '\n\n')
    .trim();
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

  const blob = new Blob([state.markdown], { type: 'text/markdown;charset=utf-8' });
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

function onClear() {
  state.selectedFile = null;
  state.markdown = '';

  dom.input.value = '';
  dom.output.value = '';
  dom.fileMeta.hidden = true;
  dom.fileName.textContent = '-';
  dom.fileSize.textContent = '-';

  dom.convertBtn.disabled = true;
  dom.clearBtn.disabled = true;
  dom.copyBtn.disabled = true;
  dom.downloadBtn.disabled = true;

  setStatus('Bitte eine .docx Datei auswählen.');
}

function setStatus(message) {
  dom.status.classList.remove('error');
  dom.status.textContent = message;
}

function setError(message) {
  dom.status.classList.add('error');
  dom.status.textContent = message;
}
