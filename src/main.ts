/**
 * Document Redaction Add-in
 * Main entry point
 */

import './styles.css';
import { redactDocument, RedactionResult } from './redaction';

// SVG Icons as template strings
const ICONS = {
  shield: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
    <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>
    <path d="M9 12l2 2 4-4"/>
  </svg>`,
  email: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
    <rect x="2" y="4" width="20" height="16" rx="2"/>
    <path d="m22 7-8.97 5.7a1.94 1.94 0 0 1-2.06 0L2 7"/>
  </svg>`,
  phone: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
    <path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"/>
  </svg>`,
  id: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
    <rect x="2" y="4" width="20" height="16" rx="2"/>
    <circle cx="8" cy="10" r="2"/>
    <path d="M8 14a4 4 0 0 0-4 4"/>
    <line x1="14" y1="10" x2="20" y2="10"/>
    <line x1="14" y1="14" x2="18" y2="14"/>
  </svg>`,
  redact: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
    <path d="M2 12s3-7 10-7 10 7 10 7-3 7-10 7-10-7-10-7Z"/>
    <circle cx="12" cy="12" r="3"/>
    <line x1="2" y1="2" x2="22" y2="22"/>
  </svg>`,
  success: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
    <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
    <polyline points="22 4 12 14.01 9 11.01"/>
  </svg>`,
  warning: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
    <path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z"/>
    <line x1="12" y1="9" x2="12" y2="13"/>
    <line x1="12" y1="17" x2="12.01" y2="17"/>
  </svg>`,
  error: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
    <circle cx="12" cy="12" r="10"/>
    <line x1="15" y1="9" x2="9" y2="15"/>
    <line x1="9" y1="9" x2="15" y2="15"/>
  </svg>`,
};

/**
 * Initialize the Office Add-in
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    initializeApp();
  }
});

/**
 * Initialize the application UI
 */
function initializeApp(): void {
  const app = document.getElementById('app');
  if (!app) return;

  app.innerHTML = `
    <div class="app-container">
      <header class="header">
        <div class="header-icon">
          ${ICONS.shield}
        </div>
        <h1>Document Redaction</h1>
        <p>Protect sensitive information with one click</p>
      </header>

      <div class="info-card">
        <h2>What will be redacted</h2>
        <ul class="info-list">
          <li>
            <span class="icon">${ICONS.email}</span>
            <span>Email addresses</span>
          </li>
          <li>
            <span class="icon">${ICONS.phone}</span>
            <span>Phone numbers</span>
          </li>
          <li>
            <span class="icon">${ICONS.id}</span>
            <span>Social Security Numbers (full & partial)</span>
          </li>
        </ul>
      </div>

      <button id="redact-btn" class="redact-btn">
        <span class="btn-content">
          <span class="btn-icon">${ICONS.redact}</span>
          <span>Redact Document</span>
        </span>
      </button>

      <div id="status-container" class="status-container"></div>

      
    </div>
  `;

  // Attach event listener
  const redactBtn = document.getElementById('redact-btn');
  if (redactBtn) {
    redactBtn.addEventListener('click', handleRedactClick);
  }
}

/**
 * Handle the redact button click
 */
async function handleRedactClick(): Promise<void> {
  const btn = document.getElementById('redact-btn') as HTMLButtonElement;
  const statusContainer = document.getElementById('status-container');
  
  if (!btn || !statusContainer) return;

  // Show loading state
  btn.disabled = true;
  btn.classList.add('loading');
  btn.innerHTML = `
    <span class="btn-content">
      <span class="spinner"></span>
      <span>Processing...</span>
    </span>
  `;
  statusContainer.innerHTML = '';

  try {
    const result = await redactDocument();
    displayResult(result, statusContainer);
  } catch (error) {
    displayError(error, statusContainer);
  } finally {
    // Reset button
    btn.disabled = false;
    btn.classList.remove('loading');
    btn.innerHTML = `
      <span class="btn-content">
        <span class="btn-icon">${ICONS.redact}</span>
        <span>Redact Document</span>
      </span>
    `;
  }
}

/**
 * Display the redaction result
 */
function displayResult(result: RedactionResult, container: HTMLElement): void {
  if (result.success) {
    if (result.totalRedacted > 0) {
      container.innerHTML = `
        <div class="status-message success">
          <span class="status-icon">${ICONS.success}</span>
          <div class="status-content">
            <div class="status-title">Redaction Complete</div>
            <div class="status-details">
              Successfully redacted ${result.totalRedacted} item${result.totalRedacted !== 1 ? 's' : ''}.
              ${result.trackingEnabled ? 'Changes are being tracked.' : ''}
            </div>
            <div class="stats">
              <div class="stat-item">
                <div class="stat-value">${result.emailsRedacted}</div>
                <div class="stat-label">Emails</div>
              </div>
              <div class="stat-item">
                <div class="stat-value">${result.phonesRedacted}</div>
                <div class="stat-label">Phone Numbers</div>
              </div>
              <div class="stat-item">
                <div class="stat-value">${result.ssnsRedacted}</div>
                <div class="stat-label">SSNs</div>
              </div>
            </div>
          </div>
        </div>
      `;
    } else {
      container.innerHTML = `
        <div class="status-message warning">
          <span class="status-icon">${ICONS.warning}</span>
          <div class="status-content">
            <div class="status-title">No Sensitive Data Found</div>
            <div class="status-details">
              The document was scanned but no sensitive information was detected.
            </div>
          </div>
        </div>
      `;
    }
  } else {
    container.innerHTML = `
      <div class="status-message error">
        <span class="status-icon">${ICONS.error}</span>
        <div class="status-content">
          <div class="status-title">Redaction Failed</div>
          <div class="status-details">${result.error || 'An unexpected error occurred.'}</div>
        </div>
      </div>
    `;
  }
}

/**
 * Display an error message
 */
function displayError(error: unknown, container: HTMLElement): void {
  const message = error instanceof Error ? error.message : 'An unexpected error occurred.';
  container.innerHTML = `
    <div class="status-message error">
      <span class="status-icon">${ICONS.error}</span>
      <div class="status-content">
        <div class="status-title">Error</div>
        <div class="status-details">${message}</div>
      </div>
    </div>
  `;
}

