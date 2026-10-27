const EXCEL_FILE_PATH = 'data/India_300_National_Initiatives_History_2_2.xlsx';

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
        container.innerHTML = "<p>No case file specified.</p>"; return;
    }

    try {
        const response = await fetch(EXCEL_FILE_PATH);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: "" });

        const row = data[parseInt(idParam)];
        if (!row) { container.innerHTML = "<p>Record not found in database.</p>"; return; }

        const initiative = cleanValue(row['What Initiative'], 'Strategic Initiative');
        const year = cleanValue(row['Year'], 'N/A');
        const delay = cleanValue(row['Delay'], 'N/A');
        const outcome = cleanValue(row['Final Outcome'], 'N/A');
        const opposed = cleanValue(row['Who opposed'], 'N/A');
        const legalChallenges = cleanValue(row['Legal Challenges (if any)'], 'None officially recorded');
        const importance = cleanValue(row['Why important for India'], 'Detailed strategic importance analysis not available.');
        const effect = cleanValue(row['Effect on India for the delay'], 'Detailed macroeconomic and national security effect analysis not available.');

        let html = `
            <div class="dossier-header-banner">
                <h1 class="dossier-title">${initiative}</h1>
                <div class="meta-badges">
                    <span class="badge badge-info">
                        <svg width="16" height="16" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z"></path></svg>
                        Initiation Year: ${year}
                    </span>
                    <span class="badge badge-danger">
                        <svg width="16" height="16" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>
                        Total Delay: ${delay}
                    </span>
                    <span class="badge badge-neutral">
                        <svg width="16" height="16" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>
                        Status: ${outcome}
                    </span>
                </div>
            </div>

            <div class="content-section">
                <h3>
                    <svg width="22" height="22" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"></path></svg>
                    Opposition & Legal Hurdles
                </h3>
                <div class="legal-highlight-box">
                    <p style="margin-bottom: 12px;"><strong>Who Opposed:</strong> ${opposed}</p>
                    <p><strong>Legal Challenges / Court Stays:</strong> ${legalChallenges}</p>
                </div>
            </div>

            <div class="content-section">
                <h3>
                    <svg width="22" height="22" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 10V3L4 14h7v7l9-11h-7z"></path></svg>
                    Why Important for India
                </h3>
                <p>${importance}</p>
            </div>

            <div class="content-section">
                <h3>
                    <svg width="22" height="22" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"></path></svg>
                    Effect on India for the Delay
                </h3>
                <p>${effect}</p>
            </div>

            <div class="content-section">
                <h3>
                    <svg width="22" height="22" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4M14 4h6m0 0v6m0-6L10 14"></path></svg>
                    Verified Source Documents & Media Links
                </h3>
                <div class="doc-links">
        `;

        const link1 = cleanValue(row['Proof Link 1'], "");
        const link2 = cleanValue(row['Proof Link 2'], "");
        const link3 = cleanValue(row['Proof Link 3'], "");
        const link4 = cleanValue(row['Proof Link 4'], "");

        if (link1 && link1 !== "Not Specified") {
            html += `<a href="${link1}" target="_blank" class="doc-btn"><span>Source Document 1</span><svg fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4M14 4h6m0 0v6m0-6L10 14"></path></svg></a>`;
        }
        if (link2 && link2 !== "Not Specified") {
            html += `<a href="${link2}" target="_blank" class="doc-btn"><span>Source Document 2</span><svg fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4M14 4h6m0 0v6m0-6L10 14"></path></svg></a>`;
        }
        if (link3 && link3 !== "Not Specified") {
            html += `<a href="${link3}" target="_blank" class="doc-btn"><span>Source Document 3</span><svg fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4M14 4h6m0 0v6m0-6L10 14"></path></svg></a>`;
        }
        if (link4 && link4 !== "Not Specified") {
            html += `<a href="${link4}" target="_blank" class="doc-btn"><span>Source Document 4</span><svg fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4M14 4h6m0 0v6m0-6L10 14"></path></svg></a>`;
        }

        html += `</div></div>`;
        container.innerHTML = html;

    } catch (error) {
        console.error(error);
        container.innerHTML = `<p style="color:var(--danger); font-weight:700;">Error loading case file dossier. Ensure Live Server is running.</p>`;
    }
}

loadDossier();