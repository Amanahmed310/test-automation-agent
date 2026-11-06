import { spawn } from 'node:child_process';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const APP_DIR = path.dirname(fileURLToPath(import.meta.url));
const POWERSHELL_SCRIPT = path.join(APP_DIR, 'open-outlook.ps1');

function appleScriptString(value) {
  const lines = String(value ?? '').replace(/\r\n?/g, '\n').split('\n');
  return lines.map((line) => `"${line.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`).join(' & linefeed & ');
}

export function buildOutlookBody({ message, senderEmail = '', testCaseId = '' }) {
  return [
    String(message ?? '').trim(),
    senderEmail ? `Reply to: ${senderEmail}` : '',
    testCaseId ? `Related Azure DevOps test case: ${testCaseId}` : '',
    'Sent from Automation Agent.',
  ].filter(Boolean).join('\n\n');
}

export function buildMacOutlookScript({ message, senderEmail = '', testCaseId = '', imagePath = '' }) {
  const body = buildOutlookBody({ message, senderEmail, testCaseId });
  const attachment = imagePath
    ? `\n  make new attachment at newMessage with properties {file:POSIX file ${appleScriptString(imagePath)}}`
    : '';
  return `tell application "Microsoft Outlook"\n  set newMessage to make new outgoing message with properties {subject:${appleScriptString('Automation Agent feedback')}, content:${appleScriptString(body)}}${attachment}\n  open newMessage\n  activate\nend tell\n`;
}

function runProcess(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'], ...options });
    let stdout = '';
    let stderr = '';
    child.stdout?.setEncoding('utf8').on('data', (chunk) => { stdout += chunk; });
    child.stderr?.setEncoding('utf8').on('data', (chunk) => { stderr += chunk; });
    child.once('error', reject);
    child.once('close', (code) => {
      if (code === 0) resolve({ stdout, stderr });
      else reject(new Error((stderr || stdout).trim() || `${path.basename(command)} exited with code ${code}.`));
    });
  });
}

export function createOutlookDraftLauncher({ platform = process.platform, processRunner = runProcess } = {}) {
  return async function openOutlookDraft({ message, senderEmail = '', testCaseId = '', image = null }) {
    if (platform !== 'win32' && platform !== 'darwin') {
      const error = new Error('Opening an Outlook draft is supported on Windows and macOS.');
      error.code = 'OUTLOOK_UNSUPPORTED_PLATFORM';
      throw error;
    }
    const tempDir = await mkdtemp(path.join(os.tmpdir(), 'automation-agent-feedback-'));
    try {
      const imagePath = image ? path.join(tempDir, image.name) : '';
      if (image) await writeFile(imagePath, image.buffer, { flag: 'wx' });
      if (platform === 'win32') {
        const payloadPath = path.join(tempDir, 'feedback.json');
        await writeFile(payloadPath, JSON.stringify({
          to: '',
          subject: 'Automation Agent feedback',
          body: buildOutlookBody({ message, senderEmail, testCaseId }),
          imagePath,
        }), 'utf8');
        await processRunner('powershell.exe', [
          '-NoLogo', '-NoProfile', '-NonInteractive', '-STA', '-ExecutionPolicy', 'Bypass',
          '-File', POWERSHELL_SCRIPT, payloadPath,
        ]);
      } else {
        const scriptPath = path.join(tempDir, 'open-outlook.applescript');
        await writeFile(scriptPath, buildMacOutlookScript({ message, senderEmail, testCaseId, imagePath }), 'utf8');
        await processRunner('/usr/bin/osascript', [scriptPath]);
      }
      return { opened: true };
    } catch (error) {
      const wrapped = new Error(platform === 'win32'
        ? `Could not open an Outlook compose window. Install/use classic Outlook for Windows and make sure it has been configured. ${error.message}`
        : `Could not open a compose window in Microsoft Outlook for Mac. Make sure Outlook is installed and configured. ${error.message}`);
      wrapped.code = 'OUTLOOK_DRAFT_FAILED';
      throw wrapped;
    } finally {
      await rm(tempDir, { recursive: true, force: true }).catch(() => {});
    }
  };
}
