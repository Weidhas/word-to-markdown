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
  clearBtn: document.getElementById('clearBtn'),
  copyBtn: document.getElementById('copyBtn'),
  downloadBtn: document.getElementById('downloadBtn'),
  output: document.getElementById('markdownOutput')
};

const state = {
  selectedFile: null,
  markdown: ''
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
  dom.output.value = '';

  dom.fileMeta.hidden = false;
  dom.fileName.textContent = file.name;
  dom.fileSize.textContent = `${(file.size / (1024 * 1024)).toFixed(2)} MB`;

  dom.clearBtn.disabled = false;
  dom.copyBtn.disabled = true;
  dom.downloadBtn.disabled = true;

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

function onClear() {
  state.selectedFile = null;
  state.markdown = '';

  dom.input.value = '';
  dom.output.value = '';
  dom.fileMeta.hidden = true;
  dom.fileName.textContent = '-';
  dom.fileSize.textContent = '-';

  dom.clearBtn.disabled = true;
  dom.copyBtn.disabled = true;
  dom.downloadBtn.disabled = true;

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
