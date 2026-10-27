let globalData = [];
const EXCEL_FILE_PATH = 'data/India_300_National_Initiatives_History_2_2.xlsx';

async function loadDatabase() {
    try {
        const response = await fetch(EXCEL_FILE_PATH);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        
        globalData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        globalData.forEach((row, index) => { row.id = index; });

        populateFilterCheckboxes(globalData);
        renderCards(globalData);
    } catch (error) {
        console.error("Database loading failed:", error);
        document.getElementById('cardList').innerHTML = `
            <div class="one-line-card" style="grid-template-columns: 1fr; text-align: center;">
                <p style="color: var(--danger); font-weight: 700;">Failed to load Excel file. Ensure 'data/India_300_National_Initiatives_History_2.xlsx' exists and Live Server is running.</p>
            </div>`;
    }
}

function cleanValue(val, fallback = "National Interest Decision") {
    if (val === undefined || val === null || val === "") return fallback;
    const str = String(val).trim();
    if (str.toUpperCase() === "NA" || str.toUpperCase() === "N/A" || str.toUpperCase() === "NONE") {
        return fallback;
    }
    return str;
}

// Populate multi-select checkbox lists dynamically from data
function populateFilterCheckboxes(data) {
    const decadeBox = document.getElementById('decadeCheckboxes');
    const yearBox = document.getElementById('yearCheckboxes');
    const oppBox = document.getElementById('oppositionCheckboxes');

    // 1. Years & Decades
    const uniqueYears = [...new Set(data.map(item => String(cleanValue(item['Year'], ''))))].filter(y => y !== '' && y !== 'N/A').sort((a, b) => b - a);
    
    const decadeSet = new Set();
    uniqueYears.forEach(y => {
        const yrNum = parseInt(y);
        if (!isNaN(yrNum)) {
            decadeSet.add(Math.floor(yrNum / 10) * 10);
        }
    });
    const sortedDecades = [...decadeSet].sort((a, b) => b - a);

    decadeBox.innerHTML = '';
    sortedDecades.forEach(dec => {
        decadeBox.innerHTML += `
            <label class="checkbox-label">
                <input type="checkbox" name="decade" value="${dec}"> ${dec}s (${dec} - ${dec + 9})
            </label>`;
    });

    yearBox.innerHTML = '';
    uniqueYears.forEach(year => {
        yearBox.innerHTML += `
            <label class="checkbox-label">
                <input type="checkbox" name="year" value="${year}"> ${year}
            </label>`;
    });

    // 2. Opposition Groups
    let oppSet = new Set();
    data.forEach(item => {
        const rawOpp = cleanValue(item['Who opposed'] || item['Opposed By'], '');
        if (rawOpp && rawOpp !== 'N/A') {
            rawOpp.split(',').forEach(part => {
                const cleaned = part.trim();
                if (cleaned) oppSet.add(cleaned);
            });
        }
    });
    const sortedOpps = [...oppSet].sort();

    oppBox.innerHTML = '';
    sortedOpps.forEach(opp => {
        oppBox.innerHTML += `
            <label class="checkbox-label">
                <input type="checkbox" name="opposition" value="${opp}"> ${opp}
            </label>`;
    });
}

function renderCards(data) {
    const listContainer = document.getElementById('cardList');
    listContainer.innerHTML = ""; 

    if (data.length === 0) {
        listContainer.innerHTML = `
            <div class="one-line-card" style="grid-template-columns: 1fr; text-align: center;">
                <p>No records found matching your filter criteria.</p>
            </div>`;
        return;
    }

    data.forEach(row => {
        const year = cleanValue(row['Year'] || row['year'], "----");
        const initiative = cleanValue(row['What Initiative'] || row['Initiative'], "National Initiative");
        const opposedRaw = cleanValue(row['Who opposed'] || row['Opposed By'], "Opposition Groups");

        let tagsHtml = '';
        opposedRaw.split(',').forEach(name => {
            const trimmed = name.trim();
            if (trimmed) {
                tagsHtml += `<span class="opposition-pill">${trimmed}</span>`;
            }
        });

        const card = document.createElement('div');
        card.className = 'one-line-card';
        card.onclick = () => { window.location.href = `details.html?id=${row.id}`; };

        card.innerHTML = `
            <div class="card-year-pill">${year}</div>
            <div>
                <div class="card-initiative-title">${initiative}</div>
            </div>
            <div>
                <div style="font-size: 0.75rem; font-weight: 700; color: var(--text-secondary); text-transform: uppercase; margin-bottom: 4px; letter-spacing: 0.5px;">Faced Opposition From:</div>
                <div class="opposition-tags-container">${tagsHtml}</div>
            </div>
            <div class="card-action-arrow" title="View Full Case File">
                <svg width="18" height="18" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5l7 7-7 7"></path></svg>
            </div>
        `;
        listContainer.appendChild(card);
    });
}

// Master filter logic triggered ONLY on clicking "Apply Filters"
function applyAllFilters() {
    const query = document.getElementById('searchInput').value.toLowerCase();
    const sortVal = document.getElementById('sortSelect').value;

    // Collect checked values
    const selectedDecades = [...document.querySelectorAll('input[name="decade"]:checked')].map(cb => parseInt(cb.value));
    const selectedYears = [...document.querySelectorAll('input[name="year"]:checked')].map(cb => cb.value);
    const selectedOpps = [...document.querySelectorAll('input[name="opposition"]:checked')].map(cb => cb.value.toLowerCase());

    let filtered = globalData.filter(row => {
        const yearStr = String(cleanValue(row['Year'], ''));
        const yearNum = parseInt(yearStr) || 0;
        const oppStr = String(cleanValue(row['Who opposed'], '')).toLowerCase();

        // 1. Search Query
        const matchesSearch = Object.values(row).some(val => String(val).toLowerCase().includes(query));

        // 2. Multi-Decade Check (OR logic among selected)
        let matchesDecade = true;
        if (selectedDecades.length > 0) {
            matchesDecade = selectedDecades.some(dec => yearNum >= dec && yearNum <= dec + 9);
        }

        // 3. Multi-Year Check (OR logic among selected)
        let matchesYear = true;
        if (selectedYears.length > 0) {
            matchesYear = selectedYears.includes(yearStr);
        }

        // 4. Multi-Opposition Check (OR logic among selected)
        let matchesOpp = true;
        if (selectedOpps.length > 0) {
            matchesOpp = selectedOpps.some(op => oppStr.includes(op));
        }

        return matchesSearch && matchesDecade && matchesYear && matchesOpp;
    });

    // Sorting
    if (sortVal === 'yearAsc') {
        filtered.sort((a, b) => parseInt(cleanValue(a['Year'], 0)) - parseInt(cleanValue(b['Year'], 0)));
    } else if (sortVal === 'yearDesc') {
        filtered.sort((a, b) => parseInt(cleanValue(b['Year'], 0)) - parseInt(cleanValue(a['Year'], 0)));
    }

    renderCards(filtered);

    // Collapse panel after applying
    document.getElementById('advancedFiltersPanel').classList.remove('active');
}

// Clear all filters
function clearAllFilters() {
    document.querySelectorAll('.advanced-filters input[type="checkbox"]').forEach(cb => cb.checked = false);
    document.getElementById('searchInput').value = '';
    applyAllFilters();
}

// Event Listeners
document.getElementById('filterToggleBtn').addEventListener('click', () => {
    const panel = document.getElementById('advancedFiltersPanel');
    panel.classList.toggle('active');
});

document.getElementById('applyFiltersBtn').addEventListener('click', applyAllFilters);
document.getElementById('clearFiltersBtn').addEventListener('click', clearAllFilters);
document.getElementById('searchInput').addEventListener('input', applyAllFilters); // Search remains real-time
document.getElementById('sortSelect').addEventListener('change', applyAllFilters);

loadDatabase();