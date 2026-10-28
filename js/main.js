/* ==========================================================================
   main.js — Dashboard Controller, Multi-Filter Engine & Theme Manager
   ========================================================================== */

// ------------------------------------------
// 1. LIGHT / DARK THEME MANAGER
// ------------------------------------------
const themeToggleBtn = document.getElementById('themeToggleBtn');
const themeIcon = document.getElementById('themeIcon');
const themeLabel = document.getElementById('themeLabel');

// Apply saved theme on initial page load
const savedTheme = localStorage.getItem('iai_theme') || 'light';
applyThemeUI(savedTheme);

function applyThemeUI(theme) {
    if (theme === 'dark') {
        document.body.setAttribute('data-theme', 'dark');
        if (themeLabel) themeLabel.textContent = 'Light';
        if (themeIcon) {
            // Sun Icon for switching back to Light Mode
            themeIcon.innerHTML = '<path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 3v1m0 16v1m9-9h-1M4 12H3m15.364 6.364l-.707-.707M6.343 6.343l-.707-.707m12.728 0l-.707.707M6.343 17.657l-.707.707M16 12a4 4 0 11-8 0 4 4 0 018 0z"></path>';
        }
    } else {
        document.body.removeAttribute('data-theme');
        if (themeLabel) themeLabel.textContent = 'Dark';
        if (themeIcon) {
            // Moon Icon for switching to Dark Mode
            themeIcon.innerHTML = '<path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M20.354 15.354A9 9 0 018.646 3.646 9.003 9.003 0 0012 21a9.003 9.003 0 008.354-5.646z"></path>';
        }
    }
}

// Toggle theme when user clicks the header button
themeToggleBtn.addEventListener('click', () => {
    const isCurrentlyDark = document.body.getAttribute('data-theme') === 'dark';
    const nextTheme = isCurrentlyDark ? 'light' : 'dark';
    localStorage.setItem('iai_theme', nextTheme);
    applyThemeUI(nextTheme);
});

// ------------------------------------------
// 2. DATABASE LOADING & INITIALIZATION
// ------------------------------------------
let globalData = [];
const EXCEL_FILE_PATH = 'data/India_300_National_Initiatives_History_2_2.xlsx';

async function loadDatabase() {
    try {
        // Fetch and parse the Excel spreadsheet via SheetJS
        const response = await fetch(EXCEL_FILE_PATH);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];

        // Convert sheet rows into JSON objects and assign stable IDs
        globalData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        globalData.forEach((row, index) => {
            row.id = index;
        });

        // Build multi-select filter checkboxes and render initial cards
        populateFilterCheckboxes(globalData);
        renderCards(globalData);
        updateDashboardStats(globalData, false);
    } catch (error) {
        console.error("Failed to load Excel database:", error);
        document.getElementById('cardList').innerHTML = `
            <div class="loading-state" style="color: var(--danger);">
                Failed to load '${EXCEL_FILE_PATH}'. Please ensure the file exists and VS Code Live Server is running.
            </div>`;
    }
}

// Helper function to sanitize cell values
function cleanValue(val, fallback = "Unspecified") {
    if (val === undefined || val === null || val === "") return fallback;
    const str = String(val).trim();
    if (str.toUpperCase() === "NA" || str.toUpperCase() === "N/A" || str.toUpperCase() === "NONE") {
        return fallback;
    }
    return str;
}

// ------------------------------------------
// 3. LIVE COUNTER & STATS UPDATER
// ------------------------------------------
function updateDashboardStats(currentData, isFiltered = false) {
    const countEl = document.getElementById('totalRecordsCount');
    const labelEl = document.getElementById('recordsStatLabel');
    const subtextEl = document.getElementById('recordsFilterSubtext');

    // Update the count of currently displayed initiatives (before & after filter)
    countEl.textContent = currentData.length;

    if (isFiltered) {
        labelEl.textContent = "Filtered Initiatives";
        subtextEl.textContent = `Showing ${currentData.length} of ${globalData.length} total records`;
    } else {
        labelEl.textContent = "Verified Initiatives";
        subtextEl.textContent = `Showing all ${globalData.length} records in ledger`;
    }

    // Count distinct opposition groups across the currently displayed dataset
    const oppSet = new Set();
    currentData.forEach(item => {
        const rawOpp = cleanValue(item['Who opposed'] || item['Opposed By'], '');
        if (rawOpp && rawOpp !== 'Unspecified') {
            rawOpp.split(',').forEach(part => {
                const cleaned = part.trim();
                if (cleaned) oppSet.add(cleaned);
            });
        }
    });
    document.getElementById('totalOppGroups').textContent = oppSet.size;
}

// ------------------------------------------
// 4. POPULATE MULTI-SELECT PILL CHECKBOXES
// ------------------------------------------
function populateFilterCheckboxes(data) {
    const decadeBox = document.getElementById('decadeCheckboxes');
    const yearBox = document.getElementById('yearCheckboxes');
    const oppBox = document.getElementById('oppositionCheckboxes');

    // Extract unique years and sort newest to oldest
    const uniqueYears = [...new Set(data.map(item => String(cleanValue(item['Year'], ''))))]
        .filter(y => y !== '' && y !== 'Unspecified')
        .sort((a, b) => parseInt(b) - parseInt(a));

    // Calculate unique decades from years
    const decadeSet = new Set();
    uniqueYears.forEach(y => {
        const yrNum = parseInt(y);
        if (!isNaN(yrNum)) {
            decadeSet.add(Math.floor(yrNum / 10) * 10);
        }
    });

    // Render Decade Checkboxes
    decadeBox.innerHTML = '';
    [...decadeSet].sort((a, b) => b - a).forEach(dec => {
        decadeBox.innerHTML += `
            <label class="filter-pill-label">
                <input type="checkbox" name="decade" value="${dec}">
                <span>${dec}s</span>
            </label>`;
    });

    // Render Specific Year Checkboxes
    yearBox.innerHTML = '';
    uniqueYears.forEach(year => {
        yearBox.innerHTML += `
            <label class="filter-pill-label">
                <input type="checkbox" name="year" value="${year}">
                <span>${year}</span>
            </label>`;
    });

    // Extract and Render Unique Opposition Groups
    const oppSet = new Set();
    data.forEach(item => {
        const rawOpp = cleanValue(item['Who opposed'] || item['Opposed By'], '');
        if (rawOpp && rawOpp !== 'Unspecified') {
            rawOpp.split(',').forEach(part => {
                const cleaned = part.trim();
                if (cleaned) oppSet.add(cleaned);
            });
        }
    });

    oppBox.innerHTML = '';
    [...oppSet].sort().forEach(opp => {
        oppBox.innerHTML += `
            <label class="filter-pill-label">
                <input type="checkbox" name="opposition" value="${opp}">
                <span>${opp}</span>
            </label>`;
    });
}

// ------------------------------------------
// 5. RENDER INITIATIVE CARDS
// ------------------------------------------
function renderCards(data) {
    const listContainer = document.getElementById('cardList');
    listContainer.innerHTML = "";

    if (data.length === 0) {
        listContainer.innerHTML = `
            <div class="loading-state">
                No initiatives match your selected filter criteria. Click "Clear Filters" to reset.
            </div>`;
        return;
    }

    data.forEach(row => {
        const year = cleanValue(row['Year'] || row['year'], "----");
        const initiative = cleanValue(row['What Initiative'] || row['Initiative'], "National Initiative");
        const opposedRaw = cleanValue(row['Who opposed'] || row['Opposed By'], "Opposition Groups");

        // Split comma-separated opposition entities into compact pill tags
        let tagsHtml = '';
        opposedRaw.split(',').forEach(name => {
            const trimmed = name.trim();
            if (trimmed) {
                tagsHtml += `<span class="opp-pill">${trimmed}</span>`;
            }
        });

        const card = document.createElement('div');
        card.className = 'ledger-card';
        card.onclick = () => {
            window.location.href = `details.html?id=${row.id}`;
        };

        card.innerHTML = `
            <div class="card-top-meta">
                <div class="year-badge">${year}</div>
            </div>
            <div class="card-title">${initiative}</div>
            <div class="card-opposition-block">
                <span class="opp-label">Opposed By:</span>
                <div class="opp-pills-wrap">${tagsHtml}</div>
            </div>
            <div class="card-arrow" title="Open Case Dossier">
                <svg width="16" height="16" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M9 5l7 7-7 7"></path>
                </svg>
            </div>
        `;

        listContainer.appendChild(card);
    });
}

// ------------------------------------------
// 6. APPLY & CLEAR FILTER EXECUTION ENGINE
// ------------------------------------------
function applyAllFilters() {
    const query = document.getElementById('searchInput').value.toLowerCase().trim();
    const sortVal = document.getElementById('sortSelect').value;

    // Gather checked multi-select values
    const selectedDecades = [...document.querySelectorAll('input[name="decade"]:checked')].map(cb => parseInt(cb.value));
    const selectedYears = [...document.querySelectorAll('input[name="year"]:checked')].map(cb => cb.value);
    const selectedOpps = [...document.querySelectorAll('input[name="opposition"]:checked')].map(cb => cb.value.toLowerCase());

    // Update active filter badge count on the Filters button
    const totalSelectedFilters = selectedDecades.length + selectedYears.length + selectedOpps.length;
    const badgeEl = document.getElementById('activeFilterBadge');
    if (totalSelectedFilters > 0) {
        badgeEl.style.display = 'inline-block';
        badgeEl.textContent = totalSelectedFilters;
    } else {
        badgeEl.style.display = 'none';
    }

    // Filter dataset
    const filtered = globalData.filter(row => {
        const yearStr = String(cleanValue(row['Year'], ''));
        const yearNum = parseInt(yearStr) || 0;
        const oppStr = String(cleanValue(row['Who opposed'], '')).toLowerCase();

        // 1. Search bar match
        const matchesSearch = query === "" || Object.values(row).some(val => String(val).toLowerCase().includes(query));

        // 2. Multi-Decade match
        const matchesDecade = selectedDecades.length === 0 || selectedDecades.some(dec => yearNum >= dec && yearNum <= dec + 9);

        // 3. Multi-Year match
        const matchesYear = selectedYears.length === 0 || selectedYears.includes(yearStr);

        // 4. Multi-Opposition match
        const matchesOpp = selectedOpps.length === 0 || selectedOpps.some(op => oppStr.includes(op));

        return matchesSearch && matchesDecade && matchesYear && matchesOpp;
    });

    // Apply chronological sorting
    if (sortVal === 'yearAsc') {
        filtered.sort((a, b) => (parseInt(cleanValue(a['Year'], 0)) || 0) - (parseInt(cleanValue(b['Year'], 0)) || 0));
    } else if (sortVal === 'yearDesc') {
        filtered.sort((a, b) => (parseInt(cleanValue(b['Year'], 0)) || 0) - (parseInt(cleanValue(a['Year'], 0)) || 0));
    }

    // Check if any filter or search is currently active
    const isCurrentlyFiltered = query !== "" || totalSelectedFilters > 0;

    // Render updated cards and update live counters
    renderCards(filtered);
    updateDashboardStats(filtered, isCurrentlyFiltered);

    // Automatically collapse the filter drawer once Apply is clicked
    document.getElementById('advancedFiltersPanel').classList.remove('active');
    document.getElementById('filterToggleBtn').classList.remove('active');
}

function clearAllFilters() {
    // Uncheck all checkboxes and clear search input
    document.querySelectorAll('#advancedFiltersPanel input[type="checkbox"]').forEach(cb => {
        cb.checked = false;
    });
    document.getElementById('searchInput').value = '';
    document.getElementById('sortSelect').value = 'default';
    applyAllFilters();
}

// ------------------------------------------
// 7. EVENT LISTENERS
// ------------------------------------------
document.getElementById('filterToggleBtn').addEventListener('click', (e) => {
    e.currentTarget.classList.toggle('active');
    document.getElementById('advancedFiltersPanel').classList.toggle('active');
});

// Checkboxes only apply when clicking "Apply Filters"
document.getElementById('applyFiltersBtn').addEventListener('click', applyAllFilters);
document.getElementById('clearFiltersBtn').addEventListener('click', clearAllFilters);

// Search bar and Sort dropdown trigger live updates
document.getElementById('searchInput').addEventListener('input', applyAllFilters);
document.getElementById('sortSelect').addEventListener('change', applyAllFilters);

// Initialize application
loadDatabase();