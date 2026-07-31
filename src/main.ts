import { loadSettings, saveSettings } from './storage/settingsStore';
import { AppSettings, ChatMessage } from './types';
import { runLLMChatLoop } from './api/llmClient';
import { executeExcelTool } from './excel/tools';
import { getAllSkills } from './skills/skillsStore';

let currentSettings: AppSettings = loadSettings();
let chatMessages: ChatMessage[] = [];
let isProcessing = false;

// DOM Elements
const modelBadgeEl = document.getElementById('model-badge') as HTMLSpanElement;
const activeContextEl = document.getElementById('active-context') as HTMLSpanElement;
const btnRefreshContext = document.getElementById('btn-refresh-context') as HTMLButtonElement;
const btnSettings = document.getElementById('btn-settings') as HTMLButtonElement;
const chatContainer = document.getElementById('chat-container') as HTMLElement;
const statusBar = document.getElementById('status-bar') as HTMLElement;
const statusText = document.getElementById('status-text') as HTMLSpanElement;
const userInput = document.getElementById('user-input') as HTMLTextAreaElement;
const btnSend = document.getElementById('btn-send') as HTMLButtonElement;
const skillsListEl = document.getElementById('skills-list') as HTMLElement;

// Modal Elements
const settingsModal = document.getElementById('settings-modal') as HTMLElement;
const btnCloseSettings = document.getElementById('btn-close-settings') as HTMLButtonElement;
const btnSaveSettings = document.getElementById('btn-save-settings') as HTMLButtonElement;
const settingBaseUrl = document.getElementById('setting-base-url') as HTMLInputElement;
const settingApiKey = document.getElementById('setting-api-key') as HTMLInputElement;
const settingModel = document.getElementById('setting-model') as HTMLInputElement;
const settingMcpServers = document.getElementById('setting-mcp-servers') as HTMLTextAreaElement;
const skillsChecklistEl = document.getElementById('skills-checklist') as HTMLElement;
const settingSystemPrompt = document.getElementById('setting-system-prompt') as HTMLTextAreaElement;

// Initialize Office & App
Office.onReady(() => {
  initUI();
  refreshExcelContext();
});

function initUI() {
  updateSettingsUI();
  renderActiveSkillsBar();

  // Event Listeners
  btnSettings.addEventListener('click', openSettingsModal);
  btnCloseSettings.addEventListener('click', closeSettingsModal);
  btnSaveSettings.addEventListener('click', handleSaveSettings);

  btnRefreshContext.addEventListener('click', refreshExcelContext);

  btnSend.addEventListener('click', handleSendMessage);
  userInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSendMessage();
    }
  });

  // Quick Action Chips
  document.querySelectorAll('.chip-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      const prompt = btn.getAttribute('data-prompt');
      if (prompt) {
        userInput.value = prompt;
        handleSendMessage();
      }
    });
  });
}

function updateSettingsUI() {
  modelBadgeEl.textContent = currentSettings.model || 'gpt-4o';
  settingBaseUrl.value = currentSettings.baseUrl;
  settingApiKey.value = currentSettings.apiKey;
  settingModel.value = currentSettings.model;
  settingMcpServers.value = (currentSettings.mcpServers || []).join('\n');
  settingSystemPrompt.value = currentSettings.systemPrompt;

  renderSkillsChecklist();
}

function renderActiveSkillsBar() {
  skillsListEl.innerHTML = '';
  const allSkills = getAllSkills();
  const activeIds = currentSettings.activeSkillIds || [];
  const active = allSkills.filter((s) => activeIds.includes(s.id));

  if (active.length === 0) {
    const pill = document.createElement('span');
    pill.className = 'skill-pill';
    pill.textContent = 'None';
    skillsListEl.appendChild(pill);
    return;
  }

  for (const s of active) {
    const pill = document.createElement('span');
    pill.className = 'skill-pill';
    pill.textContent = s.name;
    skillsListEl.appendChild(pill);
  }
}

function renderSkillsChecklist() {
  skillsChecklistEl.innerHTML = '';
  const allSkills = getAllSkills();
  const activeIds = currentSettings.activeSkillIds || [];

  for (const skill of allSkills) {
    const label = document.createElement('label');
    label.className = 'skill-checkbox-label';

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.value = skill.id;
    checkbox.checked = activeIds.includes(skill.id);
    checkbox.setAttribute('data-skill-id', skill.id);

    const span = document.createElement('span');
    span.textContent = skill.name;

    label.appendChild(checkbox);
    label.appendChild(span);
    skillsChecklistEl.appendChild(label);
  }
}

function openSettingsModal() {
  updateSettingsUI();
  settingsModal.classList.remove('hidden');
}

function closeSettingsModal() {
  settingsModal.classList.add('hidden');
}

function handleSaveSettings() {
  const selectedSkillIds: string[] = [];
  skillsChecklistEl.querySelectorAll<HTMLInputElement>('input[type="checkbox"]').forEach((cb) => {
    if (cb.checked) {
      selectedSkillIds.push(cb.value);
    }
  });

  const mcpLines = settingMcpServers.value
    .split('\n')
    .map((s) => s.trim())
    .filter((s) => s.length > 0);

  currentSettings = {
    baseUrl: settingBaseUrl.value.trim() || 'https://api.openai.com/v1',
    apiKey: settingApiKey.value.trim(),
    model: settingModel.value.trim() || 'gpt-4o',
    systemPrompt: settingSystemPrompt.value.trim(),
    mcpServers: mcpLines,
    activeSkillIds: selectedSkillIds,
  };

  saveSettings(currentSettings);
  updateSettingsUI();
  renderActiveSkillsBar();
  closeSettingsModal();
}

async function refreshExcelContext() {
  activeContextEl.textContent = 'Updating Excel context...';
  const overviewResult = await executeExcelTool('get_workbook_overview', {});
  if (overviewResult.success && overviewResult.data) {
    const data = overviewResult.data as { activeSheet: string; selectionAddress: string };
    activeContextEl.textContent = `Sheet: "${data.activeSheet}" | Selected: ${data.selectionAddress}`;
  } else {
    activeContextEl.textContent = 'Excel Connected';
  }
}

function setStatus(text: string | null) {
  if (text) {
    statusText.textContent = text;
    statusBar.classList.remove('hidden');
  } else {
    statusBar.classList.add('hidden');
  }
}

async function handleSendMessage() {
  const text = userInput.value.trim();
  if (!text || isProcessing) return;

  userInput.value = '';
  isProcessing = true;
  btnSend.disabled = true;

  // Add system message if starting fresh
  if (chatMessages.length === 0 && currentSettings.systemPrompt.trim()) {
    chatMessages.push({
      id: `sys-${Date.now()}`,
      role: 'system',
      content: currentSettings.systemPrompt,
      timestamp: Date.now(),
    });
  }

  // Include auto-context in user request
  const contextOverview = await executeExcelTool('get_workbook_overview', {});
  let fullUserContent = text;
  if (contextOverview.success && contextOverview.data) {
    const ctx = contextOverview.data as { activeSheet: string; selectionAddress: string };
    fullUserContent += `\n\n[Active Context: Sheet="${ctx.activeSheet}", Selection="${ctx.selectionAddress}"]`;
  }

  const userMsg: ChatMessage = {
    id: `user-${Date.now()}`,
    role: 'user',
    content: fullUserContent,
    timestamp: Date.now(),
  };

  chatMessages.push(userMsg);
  renderChatMessages();

  try {
    chatMessages = await runLLMChatLoop({
      settings: currentSettings,
      messages: chatMessages,
      onMessageUpdate: (updatedMsgs) => {
        chatMessages = updatedMsgs;
        renderChatMessages();
      },
      onStatusChange: (status) => setStatus(status),
    });
  } catch (err) {
    console.error('Error during LLM chat loop:', err);
  } finally {
    isProcessing = false;
    btnSend.disabled = false;
    setStatus(null);
    refreshExcelContext();
  }
}

function renderChatMessages() {
  // Remove welcome card if messages exist
  const welcomeCard = chatContainer.querySelector('.welcome-card');
  if (chatMessages.filter((m) => m.role !== 'system').length > 0 && welcomeCard) {
    welcomeCard.remove();
  }

  // Clear rendered message nodes
  const existingMsgs = chatContainer.querySelectorAll('.message');
  existingMsgs.forEach((m) => m.remove());

  for (const msg of chatMessages) {
    if (msg.role === 'system') continue;

    const msgEl = document.createElement('div');
    msgEl.className = `message ${msg.role}`;

    if (msg.role === 'user') {
      const bubble = document.createElement('div');
      bubble.className = 'msg-bubble';
      const displayText = (msg.content || '').replace(/\n\n\[Active Context: [^\]]+\]$/, '');
      bubble.textContent = displayText;
      msgEl.appendChild(bubble);
    } else if (msg.role === 'assistant') {
      const bubble = document.createElement('div');
      bubble.className = 'msg-bubble';

      if (msg.reasoning) {
        const reasoningEl = document.createElement('div');
        reasoningEl.className = 'msg-reasoning';
        reasoningEl.textContent = `💭 Thinking: ${msg.reasoning}`;
        bubble.appendChild(reasoningEl);
      }

      if (msg.content) {
        const contentEl = document.createElement('div');
        contentEl.innerHTML = formatMarkdown(msg.content);
        bubble.appendChild(contentEl);
      }

      if (msg.tool_calls && msg.tool_calls.length > 0) {
        for (const tc of msg.tool_calls) {
          const toolTag = document.createElement('div');
          toolTag.className = 'tool-tag';
          toolTag.textContent = `⚙️ Called tool: ${tc.function.name}`;
          bubble.appendChild(toolTag);
        }
      }

      msgEl.appendChild(bubble);
    } else if (msg.role === 'tool') {
      const bubble = document.createElement('div');
      bubble.className = 'msg-bubble';
      const toolTag = document.createElement('div');
      toolTag.className = 'tool-tag';
      toolTag.textContent = `✅ Tool "${msg.name}" executed`;
      bubble.appendChild(toolTag);

      msgEl.appendChild(bubble);
    }

    chatContainer.appendChild(msgEl);
  }

  chatContainer.scrollTop = chatContainer.scrollHeight;
}

function formatMarkdown(text: string): string {
  let escaped = text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

  escaped = escaped.replace(/```([\s\S]*?)```/g, '<pre><code>$1</code></pre>');
  escaped = escaped.replace(/`([^`]+)`/g, '<code>$1</code>');
  escaped = escaped.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
  escaped = escaped.replace(/\n/g, '<br>');

  return escaped;
}
