const form = document.querySelector('#report-form');
const submit = document.querySelector('#submit');
const helper = document.querySelector('#helper');
const progress = document.querySelector('#progress');
const message = document.querySelector('#message');
const inputs = [...form.querySelectorAll('input[type="file"]')];
const results = document.querySelector('#results');
const resultsCount = document.querySelector('#results-count');
const summary = document.querySelector('#summary');
const tableHead = document.querySelector('#results-table thead');
const tableBody = document.querySelector('#results-table tbody');
const exportButton = document.querySelector('#export');

let currentFilename = 'reporte.xlsx';

function normalizeKey(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();
}

function refreshFiles() {
  inputs.forEach((input) => {
    const row = input.closest('.upload-row');
    const filename = row.querySelector('em');
    row.classList.toggle('ready', Boolean(input.files[0]));
    filename.textContent = input.files[0]?.name || 'Seleccionar archivo';
  });
  const complete = inputs.every((input) => input.files.length);
  submit.disabled = !complete;
  helper.textContent = complete
    ? 'Todo listo para procesar.'
    : 'Adjunta los tres archivos para continuar.';
}

function clearResults() {
  results.hidden = true;
  tableHead.replaceChildren();
  tableBody.replaceChildren();
  summary.replaceChildren();
}

function formatValue(value) {
  if (value === null || value === undefined || value === '') return '—';
  if (typeof value === 'number') {
    return new Intl.NumberFormat('es-DO', { maximumFractionDigits: 2 }).format(value);
  }
  if (/^\d{4}-\d{2}-\d{2}T/.test(String(value))) {
    const date = new Date(value);
    if (!Number.isNaN(date.getTime())) return date.toLocaleDateString('es-DO');
  }
  return String(value);
}

function renderSummary(columns, rows) {
  summary.replaceChildren();
  const statusColumn = columns.find((column) => normalizeKey(column) === 'status');
  const resultColumn = columns.find((column) => normalizeKey(column).includes('resultado revision'));
  const groupColumn = statusColumn || resultColumn;
  const counts = new Map();

  if (groupColumn) {
    rows.forEach((row) => {
      const label = formatValue(row[groupColumn]);
      counts.set(label, (counts.get(label) || 0) + 1);
    });
  } else {
    counts.set('Registros', rows.length);
  }

  counts.forEach((count, label) => {
    const item = document.createElement('div');
    const value = document.createElement('strong');
    const caption = document.createElement('span');
    value.textContent = count;
    caption.textContent = label;
    item.append(value, caption);
    summary.append(item);
  });
}

function renderTable(payload) {
  const { columnas, registros, total, nombre_archivo: filename } = payload;
  currentFilename = filename || 'reporte.xlsx';
  tableHead.replaceChildren();
  tableBody.replaceChildren();

  const headerRow = document.createElement('tr');
  columnas.forEach((column) => {
    const th = document.createElement('th');
    th.scope = 'col';
    th.textContent = column;
    headerRow.append(th);
  });
  tableHead.append(headerRow);

  registros.forEach((record) => {
    const row = document.createElement('tr');
    columnas.forEach((column) => {
      const cell = document.createElement('td');
      const value = formatValue(record[column]);
      cell.textContent = value;
      cell.title = value === '—' ? '' : value;
      row.append(cell);
    });
    tableBody.append(row);
  });

  resultsCount.textContent = `${total} registros listos para revisar.`;
  renderSummary(columnas, registros);
  results.hidden = false;
  results.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

inputs.forEach((input) => input.addEventListener('change', () => {
  clearResults();
  refreshFiles();
}));

form.querySelectorAll('input[name="tipo"]').forEach((radio) => {
  radio.addEventListener('change', () => {
    document.querySelectorAll('.choice').forEach((choice) => choice.classList.remove('selected'));
    radio.closest('.choice').classList.add('selected');
    clearResults();
  });
});

form.querySelector('#fecha_corte').addEventListener('change', clearResults);

form.addEventListener('submit', async (event) => {
  event.preventDefault();
  message.textContent = '';
  submit.disabled = true;
  submit.textContent = 'Procesando archivos…';
  progress.hidden = false;
  clearResults();

  try {
    const response = await fetch('/procesar', {
      method: 'POST',
      body: new FormData(form),
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(payload.error || 'No se pudo generar el reporte.');
    renderTable(payload);
    helper.textContent = `${payload.total} registros procesados. Revisa la tabla antes de exportar.`;
  } catch (error) {
    message.textContent = error.message;
  } finally {
    progress.hidden = true;
    submit.textContent = 'Ver resultados';
    submit.disabled = !inputs.every((input) => input.files.length);
  }
});

exportButton.addEventListener('click', async () => {
  message.textContent = '';
  exportButton.disabled = true;
  exportButton.textContent = 'Preparando Excel…';

  try {
    const response = await fetch('/exportar', {
      method: 'POST',
      body: new FormData(form),
    });
    if (!response.ok) {
      const payload = await response.json().catch(() => ({}));
      throw new Error(payload.error || 'No se pudo exportar el reporte.');
    }
    const blob = await response.blob();
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = currentFilename;
    link.click();
    URL.revokeObjectURL(link.href);
    helper.textContent = 'Excel exportado correctamente.';
  } catch (error) {
    message.textContent = error.message;
  } finally {
    exportButton.disabled = false;
    exportButton.innerHTML = '<span aria-hidden="true">↓</span> Exportar a Excel';
  }
});

refreshFiles();
