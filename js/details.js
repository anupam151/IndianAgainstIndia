/* ==========================================================================
   details.js — Case Dossier Controller & Theme Manager
   ========================================================================== */

// ------------------------------------------
// 1. LIGHT / DARK THEME MANAGER
// ------------------------------------------
const themeToggleBtn = document.getElementById('themeToggleBtn');
const themeIcon = document.getElementById('themeIcon');
const themeLabel = document.getElementById('themeLabel');

const savedTheme = localStorage.getItem('iai_theme') || 'light';
applyThemeUI(savedTheme);

function applyThemeUI(theme) {
    if (theme === 'dark') {
        document.body.setAttribute('data-theme', 'dark');
        if (themeLabel) themeLabel.textContent = 'Light';
        if (themeIcon) {
            themeIcon.innerHTML = '<path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 3v1m0 16v1m9-9h-1M4 12H3m15.364 6.364l-.707-.707M6.343 6.343l-.707-.707m12.728 0l-.707.707M6.343 17.657l-.707.707M16 12a4 4 0 11-8 0 4 4 0 018 0z"></path>';
        }
    } else {
        document.body.removeAttribute('data-theme');
        if (themeLabel) themeLabel.textContent = 'Dark';
        if (themeIcon) {
            themeIcon.innerHTML = '<path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M20.354 15.354A9 9 0 018.646 3.646 9.003 9.003 0 0012 21a9.003 9.003 0 008.354-5.646z"></path>';
        }
    }
}

themeToggleBtn.addEventListener('click', () => {
    const isCurrentlyDark = document.body.getAttribute('data-theme') === 'dark';
    const nextTheme = isCurrentlyDark ? 'light' : 'dark';
    localStorage.setItem('iai_theme', nextTheme);
    applyThemeUI(nextTheme);
});

// ------------------------------------------
// 2. DOSSIER DATA LOADER & RENDERER
// ------------------------------------------
const EXCEL_FILE_PATH = 'data/data.xlsx';

function cleanValue(val, fallback = "Not Specified") {
    if (val === undefined || val === null || val === "") return fallback;
    const str = String(val).trim();
    if (str.toUpperCase() === "NA" || str.toUpperCase() === "N/A" || str.toUpperCase() === "NONE") {
        return fallback;
    }
    return str;
}

async function loadDossier() {
    const idParam = new URLSearchParams(window.location.search).get('id');
    const container = document.getElementById('dossierContainer');

    if (idParam === null) {
        container.innerHTML = `<div class="loading-state">No case file ID specified.</div>`;
        return;
    }

    try {
        // Fetch and parse the Excel database
        const response = await fetch(EXCEL_FILE_PATH);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: "" });

        const row = data[parseInt(idParam)];
        if (!row) {
            container.innerHTML = `<div class="loading-state">Record not found in database.</div>`;
            return;
        }

        // Extract fields from the row
        const initiative = cleanValue(row['What Initiative'], 'Strategic National Initiative');
        const year = cleanValue(row['Year'], 'N/A');
        const delay = cleanValue(row['Delay'], 'N/A');
        const outcome = cleanValue(row['Final Outcome'], 'N/A');
        const opposed = cleanValue(row['Who opposed'], 'N/A');
        const legalChallenges = cleanValue(row['Legal Challenges (if any)'] || row['Legal Challenges'], 'None officially recorded');
        const importance = cleanValue(row['Why important for India'], 'Detailed strategic importance analysis not available.');
        const effect = cleanValue(row['Effect on India for the delay'], 'Detailed macroeconomic and national security effect analysis not available.');

        // Build the Dossier HTML layout
        let html = `
            <div class="dossier-banner">
                <div class="dossier-eyebrow">Case File #${parseInt(idParam) + 1}</div>
                <h1 class="dossier-title">${initiative}</h1>
                
                <div class="dossier-meta-grid">
                    <div class="meta-box meta-year">
                        <div class="meta-label">Initiation Year</div>
                        <div class="meta-value">${year}</div>
                    </div>
                    <div class="meta-box meta-delay">
                        <div class="meta-label">Total Delay</div>
                        <div class="meta-value">${delay}</div>
                    </div>
                    <div class="meta-box meta-status">
                        <div class="meta-label">Final Outcome / Status</div>
                        <div class="meta-value">${outcome}</div>
                    </div>
                </div>
            </div>

            <div class="dossier-section">
                <h3 class="section-heading">
                    <svg width="20" height="20" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"></path>
                    </svg>
                    <span>Opposition &amp; Legal Hurdles</span>
                </h3>
                <div class="legal-callout">
                    <div class="legal-item">
                        <strong>Who Opposed</strong>
                        <p>${opposed}</p>
                    </div>
                    <div class="legal-item">
                        <strong>Legal Challenges &amp; Court Stays</strong>
                        <p>${legalChallenges}</p>
                    </div>
                </div>
            </div>

            <div class="dossier-section">
                <h3 class="section-heading">
                    <svg width="20" height="20" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 10V3L4 14h7v7l9-11h-7z"></path>
                    </svg>
                    <span>Why Important for India</span>
                </h3>
                <p class="section-body">${importance}</p>
            </div>

            <div class="dossier-section">
                <h3 class="section-heading">
                    <svg width="20" height="20" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"></path>
                    </svg>
                    <span>Effect on India for the Delay</span>
                </h3>
                <p class="section-body">${effect}</p>
            </div>

            <div class="dossier-section">
                <h3 class="section-heading">
                    <svg width="20" height="20" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4M14 4h6m0 0v6m0-6L10 14"></path>
                    </svg>
                    <span>Verified Source Documents &amp; Media Links</span>
                </h3>
                <div class="doc-links-grid">
        `;

        // Render up to 4 proof links dynamically
        [1, 2, 3, 4].forEach(num => {
            const link = cleanValue(row[`Proof Link ${num}`], "");
            if (link && link !== "Not Specified") {
                html += `
                    <a href="${link}" target="_blank" rel="noopener noreferrer" class="doc-card-btn">
                        <span>Source Document ${num}</span>
                        <svg fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4M14 4h6m0 0v6m0-6L10 14"></path>
                        </svg>
                    </a>`;
            }
        });

        html += `</div></div>`;
        container.innerHTML = html;

    } catch (error) {
        console.error("Error loading dossier:", error);
        container.innerHTML = `<div class="loading-state" style="color: var(--danger);">Error loading case file dossier. Ensure VS Code Live Server is running.</div>`;
    }
}

loadDossier();