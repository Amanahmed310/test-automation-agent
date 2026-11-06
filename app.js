const settingsForm = document.querySelector('#settings-form');
const fetchForm = document.querySelector('#fetch-form');
const settingsMessage = document.querySelector('#settings-message');
const quickPatForm = document.querySelector('#pat-quick-form');
const quickPatInput = document.querySelector('#pat-quick-input');
const patSetupStatus = document.querySelector('.pat-setup-status');
const quickPatIndicator = document.querySelector('#pat-quick-indicator');
const quickPatHelp = document.querySelector('#pat-quick-help');
const quickPatSave = document.querySelector('#pat-quick-save');
let azurePatConfigured = false;
const fetchError = document.querySelector('#fetch-error');
const resultSection = document.querySelector('#test-result');
const connectionBadge = document.querySelector('#connection-badge');
const ragStatus = document.querySelector('#rag-status');
const ragError = document.querySelector('#rag-error');
const ragRebuildButton = document.querySelector('#rag-rebuild');
const graphMaintenanceStatus = document.querySelector('#graph-maintenance-status');
const graphMaintenanceError = document.querySelector('#graph-maintenance-error');
const graphServerInfo = document.querySelector('#graph-server-info');
const graphCheckButton = document.querySelector('#graph-check');
const graphUpdateButton = document.querySelector('#graph-update');
const serviceCheckButton = document.querySelector('#check-services');
const serviceCheckSummary = document.querySelector('#services-check-summary');
const serviceCheckPanel = document.querySelector('.service-check-panel');
let graphMaintenanceInFlight = false;
let serviceCheckInFlight = false;
let graphMaintenanceTask = 'idle';
let graphDatabasePath = '';
const attachmentInput = document.querySelector('#attachment-input');
const attachmentList = document.querySelector('#attachment-list');
const attachmentError = document.querySelector('#attachment-error');
const usageHistoryKey = 'automation-agent-smart-usage-v1';
const ragPreferenceKey = 'automation-agent-use-local-rag-v1';
const ragToggle = document.querySelector('#use-local-rag');
const feedbackForm = document.querySelector('#feedback-form');
const feedbackImageInput = document.querySelector('#feedback-image');
const feedbackOverlay = document.querySelector('#feedback-overlay');
const feedbackLauncher = document.querySelector('#open-feedback');
const feedbackClose = document.querySelector('#close-feedback');
try {
  const savedRagPreference = localStorage.getItem(ragPreferenceKey);
  if (savedRagPreference !== null) ragToggle.checked = savedRagPreference === 'true';
} catch { /* Keep the default enabled if browser storage is unavailable. */ }
const cliModels = {
  copilot: [
    ['auto', 'Auto (recommended for cost-aware routing)'], ['gpt-5-mini', 'GPT-5 mini'], ['gpt-5.4-mini', 'GPT-5.4 mini'], ['gpt-5.4', 'GPT-5.4'], ['gpt-5.5', 'GPT-5.5'],
    ['gpt-5.6-luna', 'GPT-5.6 Luna'], ['gpt-5.6-terra', 'GPT-5.6 Terra'], ['gpt-6-luna', 'GPT-6 Luna'],
    ['claude-haiku-4.5', 'Claude Haiku 4.5'], ['claude-haiku-5.5', 'Claude Haiku 5.5'], ['claude-sonnet-5.5', 'Claude Sonnet 5.5'], ['claude-opus-5.5', 'Claude Opus 5.5'], ['gemini-3.8-flash', 'Gemini 3.8 Flash'],
  ],
  codex: [
    ['gpt-5.3-codex', 'GPT-5.3 Codex'], ['gpt-5.4', 'GPT-5.4'], ['gpt-5.5', 'GPT-5.5'],
    ['gpt-5.6-luna', 'GPT-5.6 Luna'], ['gpt-5.6-terra', 'GPT-5.6 Terra'], ['gpt-6-luna', 'GPT-6 Luna'],
  ],
};
let fetchedTestCase = null;
let activeGenerationId = null;
let currentGenerationId = null;
let generationStopRequested = false;
const overlayMessages = [
  'A little test magic is getting everything ready…',
  'Teaching the test steps how to click and check…',
  'Gathering the framework patterns your test needs…',
  'Polishing selectors and assertions…',
  'Almost there. The tiny test sprites are checking their work…',
];

function setBadge(connected) {
  azurePatConfigured = Boolean(connected);
  const configured = ['organization', 'project', 'repositoryPath', 'outputPath', 'smartCliCommand']
    .every((name) => String(settingsForm.elements[name]?.value ?? '').trim());
  const ready = Boolean(connected && configured);
  connectionBadge.textContent = ready ? 'SETUP READY' : 'SETUP INCOMPLETE';
  connectionBadge.classList.toggle('ready', ready);
  connectionBadge.classList.toggle('incomplete', !ready);
  document.querySelector('#workspace-status').textContent = ready
    ? 'Saved locally · Ready to use'
    : 'Setup needs attention';
  document.querySelector('#setup-state-indicator').classList.toggle('ready', ready);
  document.querySelector('#setup-state-indicator').classList.toggle('incomplete', !ready);
  document.querySelector('.setup-state-row').classList.toggle('ready', ready);
  document.querySelector('.setup-state-row').classList.toggle('incomplete', !ready);
  document.querySelector('#setup-state-text').textContent = ready
    ? 'Setup is ready. Azure, framework, and output settings are configured.'
    : 'Setup is incomplete. Add the missing Azure, repository, output, or PAT details.';
  setPatStatus(connected);
  return ready;
}

function setPatStatus(connected, message) {
  const saved = Boolean(connected);
  quickPatIndicator.classList.toggle('ready', saved);
  patSetupStatus.classList.toggle('ready', saved);
  quickPatHelp.textContent = message || (saved
    ? 'PAT saved in the operating system credential store.'
    : 'PAT not saved · enter it here to connect Azure DevOps.');
}

function openSettingsOverlay() {
  document.querySelector('#settings-overlay').hidden = false;
  document.querySelector('#close-settings').focus();
}

function closeSettingsOverlay() {
  document.querySelector('#settings-overlay').hidden = true;
  document.querySelector('#open-settings').focus();
}

document.querySelector('#open-settings').addEventListener('click', openSettingsOverlay);
document.querySelector('#close-settings').addEventListener('click', closeSettingsOverlay);
document.querySelector('#settings-overlay').addEventListener('click', (event) => {
  if (event.target.id === 'settings-overlay') closeSettingsOverlay();
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && !document.querySelector('#settings-overlay').hidden) closeSettingsOverlay();
  if (event.key === 'Escape' && !feedbackOverlay.hidden) closeFeedbackOverlay();
});

function openFeedbackOverlay() {
  feedbackOverlay.hidden = false;
  document.querySelector('#feedback-email').focus();
}

function closeFeedbackOverlay() {
  feedbackOverlay.hidden = true;
  feedbackLauncher.focus();
}

feedbackLauncher.addEventListener('click', openFeedbackOverlay);
feedbackClose.addEventListener('click', closeFeedbackOverlay);
feedbackOverlay.addEventListener('click', (event) => {
  if (event.target === feedbackOverlay) closeFeedbackOverlay();
});

async function api(path, options = {}) {
  const response = await fetch(path, {
    ...options,
    headers: { 'content-type': 'application/json', ...(options.headers ?? {}) },
  });
  const responseText = await response.text();
  let data = {};
  try { data = responseText ? JSON.parse(responseText) : {}; } catch { /* Surface the server response below. */ }
  if (!response.ok) {
    const detail = data.error || data.message || responseText.trim();
    throw new Error(detail ? `Upload failed (${response.status}): ${detail}` : `Request failed (${response.status}).`);
  }
  if (!data || typeof data !== 'object') throw new Error('The server returned an invalid response. Restart the Automation Agent and try again.');
  return data;
}

function addText(parent, tag, text, className) {
  const element = document.createElement(tag);
  if (className) element.className = className;
  element.textContent = text;
  parent.append(element);
  return element;
}

function showTestCase(testCase) {
  fetchedTestCase = testCase;
  document.querySelector('#result-id').textContent = `TEST CASE ${testCase.id}`;
  document.querySelector('#result-title').textContent = testCase.title;
  const resultLink = document.querySelector('#result-link');
  resultLink.href = testCase.webUrl;
  const description = document.querySelector('#result-description');
  description.textContent = testCase.description;
  description.hidden = !testCase.description;
  const list = document.querySelector('#steps-list');
  list.replaceChildren();
  testCase.steps.forEach((step, index) => {
    const row = document.createElement('div');
    row.className = 'test-step';
    addText(row, 'span', String(index + 1).padStart(2, '0'), 'step-number');
    const content = document.createElement('div');
    content.className = 'step-content';
    const action = document.createElement('div');
    addText(action, 'strong', 'Action');
    addText(action, 'p', step.action || '—');
    const expected = document.createElement('div');
    addText(expected, 'strong', 'Expected result');
    addText(expected, 'p', step.expected || '—');
    content.append(action, expected);
    row.append(content);
    list.append(row);
  });
  document.querySelector('#steps-count').textContent = `${testCase.steps.length} step${testCase.steps.length === 1 ? '' : 's'}`;
  document.querySelector('#no-steps').hidden = testCase.steps.length > 0;
  testCase.attachments = (testCase.attachments || []).map((attachment) => ({ ...attachment, selected: Boolean(attachment.available) }));
  renderAttachments();
  resultSection.hidden = false;
}

function formatAttachmentSize(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) return 'size unavailable';
  return bytes < 1024 * 1024 ? `${Math.ceil(bytes / 1024)} KB` : `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function renderAttachments() {
  const attachments = fetchedTestCase?.attachments || [];
  attachmentList.replaceChildren();
  document.querySelector('#attachments-count').textContent = attachments.length
    ? `${attachments.filter((attachment) => attachment.available && attachment.selected).length} selected of ${attachments.length}`
    : 'No files';
  for (const attachment of attachments) {
    const row = document.createElement('div');
    row.className = 'attachment-row';
    const label = document.createElement('label');
    label.className = 'attachment-select';
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = Boolean(attachment.available && attachment.selected);
    checkbox.disabled = !attachment.available;
    checkbox.addEventListener('change', () => {
      attachment.selected = checkbox.checked;
      renderAttachments();
    });
    const details = document.createElement('span');
    details.className = 'attachment-details';
    addText(details, 'strong', attachment.name);
    addText(details, 'small', `${attachment.source || 'Added locally'} · ${formatAttachmentSize(attachment.size)}${attachment.note ? ` · ${attachment.note}` : ''}`);
    label.append(checkbox, details);
    row.append(label);
    if (attachment.source === 'Added locally') {
      const remove = addText(row, 'button', 'Remove', 'attachment-remove');
      remove.type = 'button';
      remove.addEventListener('click', () => {
        fetchedTestCase.attachments = fetchedTestCase.attachments.filter((item) => item.id !== attachment.id);
        renderAttachments();
      });
    }
    attachmentList.append(row);
  }
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).split(',')[1] || '');
    reader.onerror = () => reject(new Error(`Could not read ${file.name}.`));
    reader.readAsDataURL(file);
  });
}

attachmentInput.addEventListener('change', async () => {
  attachmentError.hidden = true;
  if (!fetchedTestCase) {
    attachmentError.textContent = 'Fetch a test case before adding files.';
    attachmentError.hidden = false;
    attachmentInput.value = '';
    return;
  }
  const files = [...attachmentInput.files];
  const selectedCount = (fetchedTestCase.attachments || []).filter((item) => item.available && item.selected).length;
  if (selectedCount + files.length > 10) {
    attachmentError.textContent = 'Choose up to 10 attachments per test case.';
    attachmentError.hidden = false;
    attachmentInput.value = '';
    return;
  }
  const totalBytes = (fetchedTestCase.attachments || []).reduce((sum, item) => sum + (item.available && item.selected ? item.size : 0), 0)
    + files.reduce((sum, file) => sum + file.size, 0);
  if (files.some((file) => file.size > 5 * 1024 * 1024) || totalBytes > 20 * 1024 * 1024) {
    attachmentError.textContent = 'Each file must be 5 MB or smaller, and selected files must total 20 MB or less.';
    attachmentError.hidden = false;
    attachmentInput.value = '';
    return;
  }
  for (const file of files) {
    try {
      const data = await api('/api/attachments', { method: 'POST', body: JSON.stringify({ name: file.name, base64: await fileToBase64(file) }) });
      fetchedTestCase.attachments.push({ ...data.attachment, selected: true });
      renderAttachments();
    } catch (error) {
      attachmentError.textContent = error.message;
      attachmentError.hidden = false;
      break;
    }
  }
  attachmentInput.value = '';
});

const wait = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds));

function selectedProvider() {
  const command = settingsForm.elements.smartCliCommand.value.trim().toLowerCase();
  return /(^|[\\/])codex(?:\.exe)?$/.test(command) ? 'codex' : 'copilot';
}

function populateModelOptions(preferred = '') {
  const select = document.querySelector('#smart-model');
  const provider = selectedProvider();
  const providerDefault = provider === 'copilot' ? 'auto' : '__cli_default__';
  const selectedDefault = preferred || providerDefault;
  select.replaceChildren();
  select.add(new Option('CLI default', '__cli_default__'));
  for (const [id, name] of cliModels[provider]) select.add(new Option(name, id));
  select.add(new Option('Enter another model ID…', '__custom__'));
  select.value = cliModels[provider].some(([id]) => id === selectedDefault) || selectedDefault === '__cli_default__' ? selectedDefault : '__custom__';
  document.querySelector('#custom-model-row').hidden = select.value !== '__custom__';
  const customModelInput = document.querySelector('#custom-smart-model');
  customModelInput.value = select.value === '__custom__' ? preferred : '';
  customModelInput.required = select.value === '__custom__';
  document.querySelector('#smart-model-help').textContent = provider === 'copilot'
    ? 'Auto is the default for cost-aware model routing. Exact credit usage still depends on tokens and the model Copilot selects; availability varies by plan and organization.'
    : 'Codex model choices depend on your CLI account and version. CLI default uses your Codex CLI configuration; exact availability may vary.';
}

function chosenModel() {
  const selected = document.querySelector('#smart-model').value;
  return selected === '__custom__' ? document.querySelector('#custom-smart-model').value.trim() : selected;
}

settingsForm.elements.smartCliCommand.addEventListener('change', () => populateModelOptions(''));
document.querySelector('#smart-model').addEventListener('change', () => {
  const isCustom = document.querySelector('#smart-model').value === '__custom__';
  document.querySelector('#custom-model-row').hidden = !isCustom;
  document.querySelector('#custom-smart-model').required = isCustom;
});

async function loadRagStatus() {
  const status = await api('/api/rag/status');
  ragStatus.textContent = status.ready
    ? `Ready · ${status.fileCount} files, ${status.chunkCount} sections · updated ${new Date(status.indexedAt).toLocaleString()}`
    : 'Not built yet. Smart Generate builds it automatically; refresh after framework changes.';
}

async function checkGraphMcp() {
  if (graphMaintenanceInFlight) return;
  graphMaintenanceInFlight = true;
  graphMaintenanceTask = 'checking';
  graphCheckButton.disabled = true;
  graphUpdateButton.disabled = true;
  graphMaintenanceError.hidden = true;
  graphMaintenanceStatus.textContent = 'Starting one temporary MCP process and checking its read-only tools…';
  loadGraphServerInfo();
  try {
    const result = await api('/api/code-review-graph/check', { method: 'POST', body: '{}' });
    graphMaintenanceStatus.textContent = `Connection successful · verified ${result.tools.join(', ')}. This check process has closed; Smart Generate starts its own read-only session.`;
  } catch (error) {
    graphMaintenanceStatus.textContent = 'MCP connection check failed.';
    graphMaintenanceError.textContent = error.message;
    graphMaintenanceError.hidden = false;
  } finally {
    graphMaintenanceInFlight = false;
    graphMaintenanceTask = 'idle';
    loadGraphServerInfo();
    graphCheckButton.disabled = false;
    graphUpdateButton.disabled = false;
  }
}

async function updateCodeReviewGraph() {
  if (graphMaintenanceInFlight) return;
  graphMaintenanceInFlight = true;
  graphMaintenanceTask = 'updating';
  graphCheckButton.disabled = true;
  graphUpdateButton.disabled = true;
  graphMaintenanceError.hidden = true;
  graphMaintenanceStatus.textContent = 'Starting the graph database update…';
  loadGraphServerInfo();
  try {
    const { id, action } = await api('/api/code-review-graph/update', { method: 'POST', body: '{}' });
    graphMaintenanceStatus.textContent = `${action === 'build' ? 'Building' : 'Updating'} the graph database…`;
    while (true) {
      const job = await api(`/api/code-review-graph/update/${encodeURIComponent(id)}`);
      graphMaintenanceStatus.textContent = `${job.stage}${job.output ? ` · ${job.output.trim().slice(-500)}` : ''}`;
      if (job.status === 'failed') throw new Error(job.error || 'Code Review Graph database update failed.');
      if (job.status === 'completed') break;
      await wait(900);
    }
    await loadSettings();
    graphMaintenanceStatus.textContent = `Complete · ${action === 'build' ? 'built' : 'updated'} ${settingsForm.elements.repositoryPath.value.trim()}\\.code-review-graph\\graph.db. MCP remains read-only during generation.`;
  } catch (error) {
    graphMaintenanceStatus.textContent = 'Graph database operation failed.';
    graphMaintenanceError.textContent = error.message;
    graphMaintenanceError.hidden = false;
  } finally {
    graphMaintenanceInFlight = false;
    graphMaintenanceTask = 'idle';
    loadGraphServerInfo();
    graphCheckButton.disabled = false;
    graphUpdateButton.disabled = false;
  }
}

function loadGraphServerInfo() {
  const repositoryPath = settingsForm.elements.repositoryPath.value.trim();
  const databasePath = graphDatabasePath || (repositoryPath ? `${repositoryPath}\\.code-review-graph\\graph.db` : 'not configured');
  const pythonCommand = settingsForm.elements.graphPythonCommand.value.trim() || (navigator.userAgent.includes('Windows') ? 'python' : 'python3');
  const state = graphMaintenanceTask === 'checking' ? 'checking' : graphMaintenanceTask === 'updating' ? 'updating database' : 'stopped (on-demand)';
  graphServerInfo.textContent = `App: ${window.location.origin} · MCP: ${state}, stdio via ${pythonCommand} · Graph DB: ${databasePath}. MCP starts for checks/generation and closes afterward.`;
}

async function rebuildRagIndex() {
  ragRebuildButton.disabled = true;
  ragError.hidden = true;
  ragStatus.textContent = 'Starting local index build…';
  try {
    const { id } = await api('/api/rag/rebuild', { method: 'POST', body: '{}' });
    while (true) {
      const job = await api(`/api/rag/rebuild/${encodeURIComponent(id)}`);
      ragStatus.textContent = `${job.progress}% · ${job.stage}`;
      if (job.status === 'failed') throw new Error(job.error || 'Local index build failed.');
      if (job.status === 'completed') break;
      await wait(500);
    }
    await loadRagStatus();
  } catch (error) {
    ragError.textContent = error.message;
    ragError.hidden = false;
  } finally {
    ragRebuildButton.disabled = false;
  }
}

function setServiceCheckState(service, state, message) {
  const element = document.querySelector(`#service-${service}-state`);
  element.dataset.state = state;
  element.textContent = message;
}

async function checkAndStartServices() {
  if (serviceCheckInFlight || graphMaintenanceInFlight) return;
  serviceCheckInFlight = true;
  serviceCheckButton.disabled = true;
  serviceCheckPanel.dataset.state = 'checking';
  serviceCheckSummary.textContent = 'Updating the Code Review Graph, refreshing RAG, and checking the selected CLI…';
  for (const service of ['mcp', 'rag', 'cli']) setServiceCheckState(service, 'checking', 'Checking…');

  try {
    const data = Object.fromEntries(new FormData(settingsForm).entries());
    data.smartModel = chosenModel();
    delete data.customSmartModel;
    const saved = await api('/api/settings', { method: 'PUT', body: JSON.stringify(data) });
    setBadge(saved.patConfigured);
    settingsMessage.textContent = 'Workspace settings saved locally before the service checks.';
    settingsMessage.classList.add('success');
    settingsMessage.classList.remove('warning');
  } catch (error) {
    for (const service of ['mcp', 'rag', 'cli']) setServiceCheckState(service, 'failed', 'Not run');
    serviceCheckPanel.dataset.state = 'failed';
    serviceCheckSummary.textContent = `Could not save workspace settings: ${error.message}`;
    serviceCheckInFlight = false;
    serviceCheckButton.disabled = false;
    return;
  }

  const checkMcp = async () => {
    try {
      graphMaintenanceInFlight = true;
      graphMaintenanceTask = 'updating';
      graphCheckButton.disabled = true;
      graphUpdateButton.disabled = true;
      graphMaintenanceError.hidden = true;
      graphMaintenanceStatus.textContent = 'Building or updating the Code Review Graph database…';
      loadGraphServerInfo();
      const { id, action } = await api('/api/code-review-graph/update', { method: 'POST', body: '{}' });
      while (true) {
        const job = await api(`/api/code-review-graph/update/${encodeURIComponent(id)}`);
        graphMaintenanceStatus.textContent = job.stage;
        if (job.status === 'failed') throw new Error(job.error || 'Graph database update failed.');
        if (job.status === 'completed') break;
        await wait(900);
      }
      graphMaintenanceStatus.textContent = `Graph database ${action === 'build' ? 'built' : 'updated'}; starting the read-only MCP check…`;
      await loadSettings();
      graphMaintenanceTask = 'checking';
      graphMaintenanceStatus.textContent = 'Starting a temporary read-only MCP process and verifying graph tools…';
      loadGraphServerInfo();
      const result = await api('/api/code-review-graph/check', { method: 'POST', body: '{}' });
      const message = `Connected · ${result.tools.length} tools verified`;
      setServiceCheckState('mcp', 'ready', message);
      graphMaintenanceStatus.textContent = `Connection successful · verified ${result.tools.join(', ')}. The temporary stdio MCP process closed after the check; generation starts its own process.`;
      return true;
    } catch (error) {
      setServiceCheckState('mcp', 'failed', error.message);
      graphMaintenanceStatus.textContent = 'Graph database update or MCP service check failed.';
      graphMaintenanceError.textContent = error.message;
      graphMaintenanceError.hidden = false;
      return false;
    } finally {
      graphMaintenanceInFlight = false;
      graphMaintenanceTask = 'idle';
      loadGraphServerInfo();
      graphCheckButton.disabled = false;
      graphUpdateButton.disabled = false;
    }
  };

  const refreshRag = async () => {
    try {
      ragError.hidden = true;
      ragStatus.textContent = 'Refreshing the local RAG index…';
      const { id } = await api('/api/rag/rebuild', { method: 'POST', body: '{}' });
      while (true) {
        const job = await api(`/api/rag/rebuild/${encodeURIComponent(id)}`);
        ragStatus.textContent = `${job.progress}% · ${job.stage}`;
        setServiceCheckState('rag', 'checking', `${job.progress}%`);
        if (job.status === 'failed') throw new Error(job.error || 'RAG index refresh failed.');
        if (job.status === 'completed') break;
        await wait(500);
      }
      const status = await api('/api/rag/status');
      ragStatus.textContent = `Ready · ${status.fileCount} files, ${status.chunkCount} sections · updated ${new Date(status.indexedAt).toLocaleString()}`;
      setServiceCheckState('rag', 'ready', `Ready · ${status.fileCount} files, ${status.chunkCount} sections`);
      return true;
    } catch (error) {
      setServiceCheckState('rag', 'failed', error.message);
      ragStatus.textContent = 'RAG index check failed.';
      ragError.textContent = error.message;
      ragError.hidden = false;
      return false;
    }
  };

  const checkCli = async () => {
    try {
      const result = await api('/api/cli/check', { method: 'POST', body: '{}' });
      const message = `${result.provider === 'codex' ? 'Codex' : 'Copilot'} · ${result.version}`;
      setServiceCheckState('cli', 'ready', message);
      return true;
    } catch (error) {
      setServiceCheckState('cli', 'failed', error.message);
      return false;
    }
  };

  try {
    const results = await Promise.all([checkMcp(), refreshRag(), checkCli()]);
    const readyCount = results.filter(Boolean).length;
    const allReady = readyCount === results.length;
    serviceCheckPanel.dataset.state = allReady ? 'ready' : 'failed';
    serviceCheckSummary.textContent = allReady
      ? 'All services are ready. The graph database was updated, the MCP handshake passed, RAG was refreshed, and the selected CLI responded to its version check.'
      : `${readyCount} of ${results.length} checks passed. Review the red result above for the specific error.`;
  } finally {
    serviceCheckInFlight = false;
    serviceCheckButton.disabled = false;
    if (!serviceCheckPanel.dataset.state || serviceCheckPanel.dataset.state === 'checking') serviceCheckPanel.dataset.state = 'failed';
  }
}

function formatUsage(job) {
  if (job.mode === 'quick') return 'Quick Generate uses no model credits; it runs locally.';
  const usage = job.usage;
  if (job.provider === 'copilot') {
    if (usage?.aiCredits !== null && usage?.aiCredits !== undefined) {
      return `Actual Copilot AI credits used: ${Number(usage.aiCredits).toFixed(4)}${usage.inputTokens || usage.outputTokens ? ` | Tokens: ${usage.inputTokens || 0} in, ${usage.outputTokens || 0} out` : ''}.`;
    }
    return job.status === 'running'
      ? 'Actual Copilot AI credits will appear when the CLI reports them at the end of this run.'
      : 'The CLI did not report an exact AI-credit figure for this run.';
  }
  if (usage && (usage.inputTokens || usage.outputTokens || usage.totalTokens)) {
    const total = usage.totalTokens ?? ((usage.inputTokens || 0) + (usage.outputTokens || 0));
    return `Codex token usage: ${total} total (${usage.inputTokens || 0} input, ${usage.outputTokens || 0} output). Codex did not report an AI-credit figure.`;
  }
  return job.status === 'running'
    ? 'Codex token usage will appear when the CLI completes this run.'
    : 'Codex did not report token usage for this run; AI-credit usage is unavailable.';
}

function renderUsageHistory() {
  const container = document.querySelector('#usage-history-rows');
  let history = [];
  try { history = JSON.parse(localStorage.getItem(usageHistoryKey) || '[]'); } catch { history = []; }
  container.replaceChildren();
  if (!history.length) {
    const empty = document.createElement('div');
    empty.className = 'usage-history-empty';
    empty.textContent = 'Run Smart Generate with RAG on and off to compare the reported usage.';
    container.append(empty);
    return;
  }
  const header = document.createElement('div');
  header.className = 'usage-history-row usage-history-head';
  for (const label of ['When', 'Test ID', 'Method', 'Provider', 'Model', 'Usage']) {
    const cell = document.createElement('span'); cell.textContent = label; header.append(cell);
  }
  container.append(header);
  for (const entry of history.slice(0, 8)) {
    const row = document.createElement('div'); row.className = 'usage-history-row';
    const usage = entry.provider === 'copilot' && entry.aiCredits !== null && entry.aiCredits !== undefined
      ? `${Number(entry.aiCredits).toFixed(4)} AI credits`
      : entry.provider === 'codex' && entry.totalTokens
        ? `${entry.totalTokens} tokens (not credits)` : 'Usage not reported';
    const graphCalls = Number.isFinite(entry.mcpToolCalls) ? entry.mcpToolCalls : null;
    const graphCallLabel = graphCalls === null ? 'calls not recorded' : `${graphCalls} graph calls`;
    const method = `${entry.useRag ? 'RAG + MCP' : 'MCP only'} (${graphCallLabel})`;
    for (const value of [new Date(entry.timestamp).toLocaleString(), entry.testCaseId, method, entry.provider, entry.model || 'CLI default', usage]) {
      const cell = document.createElement('span'); cell.textContent = value; row.append(cell);
    }
    container.append(row);
    const details = document.createElement('details'); details.className = 'usage-candidates';
    const summary = document.createElement('summary');
    const localCandidates = entry.sourceReferences || [];
    const mcpCandidates = entry.mcpCandidates || [];
    summary.textContent = `Candidates used · local ${localCandidates.length} · graph ${mcpCandidates.length}`;
    details.append(summary);
    const list = document.createElement('ul');
    for (const candidate of entry.testSearchCandidates || []) addText(list, 'li', `Full test-folder search: ${candidate}`);
    for (const candidate of localCandidates) {
      if ((entry.testSearchCandidates || []).includes(candidate)) continue;
      addText(list, 'li', `Local framework source: ${candidate}`);
    }
    for (const candidate of mcpCandidates) addText(list, 'li', `Code Review Graph MCP: ${candidate}`);
    if (!localCandidates.length && !mcpCandidates.length) {
      const note = graphCalls === null
        ? 'Candidate paths and MCP call counts were not recorded for this older run.'
        : graphCalls ? 'MCP calls were observed, but no source paths were reported in the CLI events.' : 'No source candidates were captured; no graph calls were observed for this run.';
      addText(list, 'li', note);
    }
    details.append(list);
    container.append(details);
  }
}

function saveUsageHistory(job) {
  if (job.mode !== 'smart' || !job.id) return;
  let history = [];
  try { history = JSON.parse(localStorage.getItem(usageHistoryKey) || '[]'); } catch { history = []; }
  if (history.some((entry) => entry.id === job.id)) return;
  history.unshift({ id: job.id, timestamp: new Date().toISOString(), testCaseId: job.testCaseId, provider: job.provider, model: job.model || 'CLI default',
    useRag: job.useRag !== false, mcpToolCalls: job.mcpToolCalls || 0, testsScanned: job.testsScanned || 0, testSearchCandidates: job.testSearchCandidates || [], sourceReferences: job.sourceReferences || [], mcpCandidates: job.mcpCandidates || [],
    aiCredits: job.usage?.aiCredits ?? null, totalTokens: job.usage?.totalTokens ?? null });
  try { localStorage.setItem(usageHistoryKey, JSON.stringify(history.slice(0, 8))); } catch { /* Optional browser-local history. */ }
  renderUsageHistory();
}

function renderGeneration(job) {
  const details = document.querySelector('#generation-details');
  const progress = document.querySelector('#generation-progress-bar');
  const state = document.querySelector('#generation-state');
  const error = document.querySelector('#generation-error');
  details.hidden = false;
  document.querySelector('#generation-stage').textContent = job.stage || 'Working…';
  document.querySelector('#generation-percent').textContent = `${job.progress ?? 0}%`;
  progress.style.width = `${Math.max(0, Math.min(100, job.progress ?? 0))}%`;
  state.textContent = job.status === 'running' ? 'IN PROGRESS' : job.status.toUpperCase();
  error.hidden = job.status !== 'failed';
  error.textContent = job.error || '';
  document.querySelector('#generation-output').textContent = job.outputPath ? `Saved to ${job.outputPath}` : '';
  const sources = document.querySelector('#generation-sources');
  const sourceLines = [];
  if (job.scenarioExtraction) sourceLines.push(`Spreadsheet attachments: ${job.scenarioExtraction.status}.`);
  if (job.validation) {
    sourceLines.push(`${job.validation.summary} Model passes: ${job.validation.attempts}.`);
    if (job.validation.issues?.length) sourceLines.push(`Validation: ${job.validation.issues.join(' | ')}`);
  }
  if (job.sourceReferences?.length) sourceLines.push(`Local framework references: ${job.sourceReferences.join(', ')}`);
  if (job.testsScanned) sourceLines.push(`Full tests-folder search: scanned ${job.testsScanned} specs; ranked matches: ${job.testSearchCandidates?.join(', ') || 'none'}.`);
  if (job.mcpCandidates?.length) sourceLines.push(`Code Review Graph MCP candidates: ${job.mcpCandidates.join(', ')}`);
  sources.hidden = !sourceLines.length;
  sources.textContent = sourceLines.join(' | ');
  const context = document.querySelector('#generation-context');
  context.hidden = job.mode !== 'smart';
  const modelLabel = job.model && job.model !== 'CLI default' ? ` Selected model: ${job.model}.` : ' Selected model: CLI default.';
  const baseContext = job.mode !== 'smart' ? '' : job.useRag
    ? `Context mode: Local RAG + Code Review Graph MCP. Graph lookups performed: ${job.mcpToolCalls || 0}; candidate paths: ${job.mcpCandidates?.length || 0}.${job.mcpError ? ` MCP lookup note: ${job.mcpError}` : ''}${modelLabel}`
    : `Context mode: MCP only; local RAG and example retrieval were skipped. Graph lookups performed: ${job.mcpToolCalls || 0}; candidate paths: ${job.mcpCandidates?.length || 0}.${job.mcpError ? ` MCP lookup note: ${job.mcpError}` : ''}${modelLabel}`;
  context.textContent = [baseContext, job.mode === 'smart' && job.attachmentNames?.length ? `Attachments included: ${job.attachmentNames.join(', ')}.` : ''].filter(Boolean).join(' ');
  const usageText = formatUsage(job);
  const usageDisplay = document.querySelector('#generation-usage');
  usageDisplay.textContent = usageText;
  usageDisplay.hidden = job.mode !== 'smart' && job.status === 'running';
  const preview = document.querySelector('#generated-preview');
  preview.hidden = !job.outputPath || !job.script;
  if (!preview.hidden) document.querySelector('#generated-script').textContent = job.script;
  renderRefinement(job);
  const overlay = document.querySelector('#generation-overlay');
  overlay.hidden = job.status !== 'running';
  if (!overlay.hidden) {
    document.querySelector('#overlay-title').textContent = job.mode === 'quick' ? 'Building your test scaffold' : 'Making your smart test';
    document.querySelector('#overlay-percent').textContent = `${job.progress ?? 0}%`;
    document.querySelector('#overlay-progress-bar').style.width = `${Math.max(0, Math.min(100, job.progress ?? 0))}%`;
    document.querySelector('#overlay-stage').textContent = job.stage || 'Getting started';
    const cap = job.provider === 'copilot' && job.creditLimit ? ` Configured soft cap: ${job.creditLimit} AI credits.` : '';
    document.querySelector('#overlay-credits').textContent = `${usageText}${cap}`;
    const stopButton = document.querySelector('#stop-generation');
    stopButton.disabled = generationStopRequested;
    stopButton.textContent = generationStopRequested ? 'Stopping…' : 'Stop generation';
  }
}

function renderRefinement(job) {
  const panel = document.querySelector('#refinement-panel');
  const available = job.status === 'completed' && Boolean(job.outputPath && job.script);
  panel.hidden = !available;
  if (!available) return;
  currentGenerationId = job.id;
  const refinement = job.refinement || { status: 'idle', messages: [] };
  const messages = document.querySelector('#refinement-messages');
  messages.replaceChildren();
  const visibleMessages = refinement.messages || [];
  if (!visibleMessages.length) {
    const greeting = document.createElement('div');
    greeting.className = 'refinement-message assistant';
    greeting.textContent = 'The test is ready. Tell me what to change or which test data to use.';
    messages.append(greeting);
  }
  for (const message of visibleMessages) {
    const item = document.createElement('div');
    item.className = `refinement-message ${message.role === 'user' ? 'user' : 'assistant'}${message.failed ? ' error' : ''}`;
    item.textContent = message.text;
    messages.append(item);
  }
  if (refinement.status === 'running') {
    const pending = document.createElement('div');
    pending.className = 'refinement-message assistant';
    pending.textContent = 'Updating the test…';
    messages.append(pending);
  }
  messages.scrollTop = messages.scrollHeight;
  const input = document.querySelector('#refinement-input');
  const submit = document.querySelector('#refinement-submit');
  input.disabled = refinement.status === 'running';
  submit.disabled = refinement.status === 'running' || !input.value.trim();
  submit.textContent = refinement.status === 'running' ? 'Updating…' : 'Send edit';
}

document.querySelector('#refinement-input').addEventListener('input', () => {
  document.querySelector('#refinement-submit').disabled = !document.querySelector('#refinement-input').value.trim();
});

document.querySelector('#refinement-form').addEventListener('submit', async (event) => {
  event.preventDefault();
  const input = document.querySelector('#refinement-input');
  const message = input.value.trim();
  if (!currentGenerationId || !message || message.length > 5_000) return;
  const submit = document.querySelector('#refinement-submit');
  submit.disabled = true;
  input.disabled = true;
  try {
    await api(`/api/generations/${encodeURIComponent(currentGenerationId)}/refine`, {
      method: 'POST', body: JSON.stringify({ message }),
    });
    input.value = '';
    let job;
    while (true) {
      job = await api(`/api/generations/${encodeURIComponent(currentGenerationId)}`);
      renderGeneration(job);
      if (job.refinement?.status !== 'running') break;
      await wait(700);
    }
  } catch (error) {
    const messages = document.querySelector('#refinement-messages');
    const item = document.createElement('div');
    item.className = 'refinement-message assistant error';
    item.textContent = `Could not update the test: ${error.message}`;
    messages.append(item);
  } finally {
    input.disabled = false;
    submit.disabled = !input.value.trim();
    submit.textContent = 'Send edit';
    input.focus();
  }
});

async function generateTest(mode) {
  if (!fetchedTestCase) return;
  const buttons = [...document.querySelectorAll('.generate-button')];
  buttons.forEach((button) => { button.disabled = true; });
  document.querySelector('#generation-title').textContent = mode === 'quick' ? 'Quick Generate' : 'Smart Generate';
  document.querySelector('#generation-output').textContent = '';
  document.querySelector('#generation-error').hidden = true;
  document.querySelector('#generated-preview').hidden = true;
  document.querySelector('#refinement-panel').hidden = true;
  currentGenerationId = null;
  document.querySelector('#generation-context').hidden = true;
  document.querySelector('#generation-usage').hidden = true;
  generationStopRequested = false;
  const provider = selectedProvider();
  const model = chosenModel();
  renderGeneration({ status: 'running', mode, provider: mode === 'smart' ? provider : 'local', progress: 0, stage: 'Starting generation' });
  const messageElement = document.querySelector('#overlay-message');
  let messageIndex = 0;
  messageElement.textContent = overlayMessages[messageIndex];
  const messageTimer = setInterval(() => {
    messageIndex = (messageIndex + 1) % overlayMessages.length;
    messageElement.textContent = overlayMessages[messageIndex];
  }, 2800);
  try {
    const { id } = await api('/api/generations', {
      method: 'POST',
      body: JSON.stringify({ mode, testCase: fetchedTestCase, additionalInstructions: mode === 'smart' ? document.querySelector('#generation-notes').value.trim() : '', attachmentIds: (fetchedTestCase.attachments || []).filter((attachment) => attachment.available && attachment.selected).map((attachment) => attachment.id), smartCliCommand: settingsForm.elements.smartCliCommand.value.trim() || 'copilot', smartModel: model, useRag: document.querySelector('#use-local-rag').checked }),
    });
    currentGenerationId = id;
    activeGenerationId = id;
    let finalJob;
    while (true) {
      const job = await api(`/api/generations/${encodeURIComponent(id)}`);
      renderGeneration(job);
      if (job.status !== 'running') { finalJob = job; break; }
      await wait(800);
    }
    saveUsageHistory(finalJob);
  } catch (error) {
    renderGeneration({ status: 'failed', progress: 0, stage: 'Generation failed', error: error.message });
  } finally {
    clearInterval(messageTimer);
    activeGenerationId = null;
    generationStopRequested = false;
    document.querySelector('#generation-overlay').hidden = true;
    buttons.forEach((button) => { button.disabled = false; });
  }
}

document.querySelector('#stop-generation').addEventListener('click', async () => {
  if (!activeGenerationId || generationStopRequested) return;
  generationStopRequested = true;
  document.querySelector('#stop-generation').disabled = true;
  document.querySelector('#stop-generation').textContent = 'Stopping…';
  document.querySelector('#overlay-stage').textContent = 'Stopping the CLI process and cleaning up…';
  try {
    await api(`/api/generations/${encodeURIComponent(activeGenerationId)}/cancel`, { method: 'POST', body: '{}' });
  } catch (error) {
    generationStopRequested = false;
    document.querySelector('#generation-error').textContent = `Could not stop generation: ${error.message}`;
    document.querySelector('#generation-error').hidden = false;
  }
});

ragToggle.addEventListener('change', () => {
  try { localStorage.setItem(ragPreferenceKey, String(ragToggle.checked)); } catch { /* The selection still applies to this run. */ }
});

document.querySelectorAll('.generate-button').forEach((button) => {
  button.addEventListener('click', () => generateTest(button.dataset.mode));
});

feedbackImageInput.addEventListener('change', () => {
  const file = feedbackImageInput.files?.[0];
  const name = document.querySelector('#feedback-image-name');
  const status = document.querySelector('#feedback-status');
  status.textContent = '';
  status.removeAttribute('data-state');
  if (!file) {
    name.textContent = 'One image, up to 3 MB.';
    return;
  }
  if (!file.type.startsWith('image/') || file.size > 3 * 1024 * 1024) {
    feedbackImageInput.value = '';
    name.textContent = 'One image, up to 3 MB.';
    status.textContent = 'Choose an image no larger than 3 MB.';
    status.dataset.state = 'error';
    return;
  }
  name.textContent = `${file.name} · ${(file.size / 1024).toFixed(0)} KB`;
});

feedbackForm.addEventListener('submit', async (event) => {
  event.preventDefault();
  const message = document.querySelector('#feedback-message').value.trim();
  const senderEmail = document.querySelector('#feedback-email').value.trim();
  const status = document.querySelector('#feedback-status');
  const submit = document.querySelector('#feedback-submit');
  if (!message) return;
  if (senderEmail && !/^\S+@\S+\.\S+$/.test(senderEmail)) {
    status.textContent = 'Enter a valid reply email address or leave this field blank.';
    status.dataset.state = 'error';
    return;
  }
  submit.disabled = true;
  submit.textContent = 'Opening Outlook…';
  status.textContent = 'Opening Outlook with your feedback…';
  status.removeAttribute('data-state');
  try {
    const file = feedbackImageInput.files?.[0];
    const image = file ? { name: file.name, type: file.type, base64: await fileToBase64(file) } : null;
    await api('/api/feedback', {
      method: 'POST',
      body: JSON.stringify({ message, senderEmail, image, testCaseId: fetchedTestCase?.id || '' }),
    });
    feedbackForm.reset();
    document.querySelector('#feedback-image-name').textContent = 'One image, up to 3 MB.';
    status.textContent = 'Outlook draft opened with your message and image. Review it, then click Send.';
    status.dataset.state = 'success';
  } catch (error) {
    status.textContent = `Could not open Outlook draft: ${error.message}`;
    status.dataset.state = 'error';
  } finally {
    submit.textContent = 'Open Outlook draft';
    submit.disabled = false;
  }
});

document.querySelector('#copy-script').addEventListener('click', async () => {
  const script = document.querySelector('#generated-script').textContent;
  try {
    await navigator.clipboard.writeText(script);
    document.querySelector('#copy-script').textContent = 'Copied';
    setTimeout(() => { document.querySelector('#copy-script').textContent = 'Copy script'; }, 1500);
  } catch {
    document.querySelector('#generation-error').textContent = 'Clipboard access was unavailable. Select and copy the script from the preview.';
    document.querySelector('#generation-error').hidden = false;
  }
});

async function loadSettings() {
  const data = await api('/api/settings');
  const defaults = {
    organization: '',
    project: '',
    repositoryPath: data.repositoryPath || '',
    outputPath: data.outputPath || '',
    smartCliCommand: 'copilot',
    smartModel: '',
    smartCreditLimit: '',
    graphPythonCommand: navigator.userAgent.includes('Windows') ? 'python' : 'python3',
  };
  for (const name of ['organization', 'project', 'repositoryPath', 'outputPath', 'smartCliCommand', 'smartModel', 'smartCreditLimit', 'graphPythonCommand']) {
    const legacyValue = name === 'smartCliCommand' ? data.copilotCommand : undefined;
    settingsForm.elements[name].value = data[name] || legacyValue || defaults[name] || '';
  }
  const graphStatus = document.querySelector('#graph-database-status');
  graphDatabasePath = data.graphDatabasePath || '';
  graphStatus.textContent = data.graphDatabaseFound
    ? `Graph database found: ${data.graphDatabasePath}. Test the MCP connection below; the app starts a fresh read-only MCP session for each Smart Generate run.`
    : `Graph database not found at ${data.graphDatabasePath}. Use Build / update graph to create it, then test the MCP connection.`;
  if (data.settingsStoredLocally) {
    settingsMessage.textContent = 'Workspace settings are saved in the app folder because this run cannot write to the Windows user config folder.';
  }
  populateModelOptions(data.smartModel || '');
  setBadge(data.patConfigured);
  loadGraphServerInfo();
}

quickPatForm.addEventListener('submit', async (event) => {
  event.preventDefault();
  const pat = quickPatInput.value.trim();
  if (!pat) {
    setPatStatus(azurePatConfigured, 'Paste a PAT to save it securely.');
    quickPatInput.focus();
    return;
  }
  quickPatSave.disabled = true;
  quickPatSave.textContent = 'Saving…';
  setPatStatus(false, 'Saving PAT securely…');
  try {
    const saved = await api('/api/azure/connection', {
      method: 'POST',
      body: JSON.stringify({ pat }),
    });
    quickPatInput.value = '';
    setPatStatus(saved.patConfigured, saved.patConfigured ? 'PAT saved in the operating system credential store.' : undefined);
    setBadge(saved.patConfigured);
    settingsMessage.textContent = 'Azure PAT saved securely. Workspace defaults are ready to configure.';
    settingsMessage.classList.add('success');
    settingsMessage.classList.remove('warning');
  } catch (error) {
    setPatStatus(azurePatConfigured, error.message);
  } finally {
    quickPatSave.disabled = false;
    quickPatSave.textContent = 'Save PAT';
  }
});

settingsForm.addEventListener('submit', async (event) => {
  event.preventDefault();
  settingsMessage.textContent = 'Saving…';
  const data = Object.fromEntries(new FormData(settingsForm).entries());
  data.smartModel = chosenModel();
  delete data.customSmartModel;
  try {
    const saved = await api('/api/settings', { method: 'PUT', body: JSON.stringify(data) });
    const settingsStoredLocally = saved.settingsStoredLocally;
    const ready = setBadge(saved.patConfigured);
    settingsMessage.textContent = ready
      ? 'Saved locally. Setup is ready.'
      : 'Saved locally, but setup is incomplete. Review the red status above and complete the missing fields.';
    settingsMessage.classList.toggle('success', ready);
    settingsMessage.classList.toggle('warning', !ready);
    loadRagStatus().catch(() => {
      ragStatus.textContent = 'RAG index status unavailable for this workspace.';
    });
  } catch (error) {
    settingsMessage.textContent = error.message;
    settingsMessage.classList.remove('success', 'warning');
    settingsMessage.classList.add('warning');
  }
});

fetchForm.addEventListener('submit', async (event) => {
  event.preventDefault();
  fetchError.hidden = true;
  resultSection.hidden = true;
  const button = fetchForm.querySelector('button');
  const previous = button.innerHTML;
  button.disabled = true;
  button.innerHTML = '<span>Fetching…</span><span aria-hidden="true">…</span>';
  const fields = new FormData(fetchForm);
  try {
    const data = await api('/api/azure/test-case', {
      method: 'POST',
      body: JSON.stringify({ testId: fields.get('testId') }),
    });
    currentGenerationId = null;
    document.querySelector('#refinement-panel').hidden = true;
    showTestCase(data.testCase);
    setBadge(true);
  } catch (error) {
    fetchError.textContent = error.message;
    fetchError.hidden = false;
  } finally {
    button.disabled = false;
    button.innerHTML = previous;
  }
});

loadSettings().catch((error) => {
  settingsMessage.textContent = `Could not load workspace settings: ${error.message}`;
  setBadge(false);
});
loadRagStatus().catch((error) => {
  ragStatus.textContent = 'Local index status is unavailable.';
  ragError.textContent = error.message;
  ragError.hidden = false;
});
ragRebuildButton.addEventListener('click', rebuildRagIndex);
graphCheckButton.addEventListener('click', checkGraphMcp);
graphUpdateButton.addEventListener('click', updateCodeReviewGraph);
serviceCheckButton.addEventListener('click', checkAndStartServices);
renderUsageHistory();
