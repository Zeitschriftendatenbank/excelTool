// --- Tabs ---
function showTab(id) {
    // Panels
    const panels = document.getElementsByClassName('panel');
    for (let j = 0; j < panels.length; j++) { panels[j].className = 'panel'; }
    const tab = document.getElementById(`tab_${id}`);
    if (tab) tab.className = 'panel active';
    // Buttons
    const btns = document.getElementsByClassName('tab');
    for (let k = 0; k < btns.length; k++) { btns[k].className = 'tab'; }
    const btnId = `btn_${id === 'cfg' ? 'cfg' : id}`;
    const btn = document.getElementById(btnId);
    if (btn) btn.className = 'tab active';
}

// ===== HILFSFUNKTIONEN für Textarea-"Tree" =====
function _getCurrentLine(el) {
    const start = el.selectionStart, val = el.value;
    let lineStart = start; while (lineStart > 0 && val.charAt(lineStart - 1) !== '\n') lineStart--;
    let lineEnd = start; while (lineEnd < val.length && val.charAt(lineEnd) !== '\n') lineEnd++;
    return val.substring(lineStart, lineEnd);
}
// function waehleZeile() removed (duplicate)
function handle_key_press_auswahl(evt) {
    evt = evt || window.event; const code = evt.keyCode || evt.which;
    if (code === 13) { if (evt.preventDefault) evt.preventDefault(); waehleZeile(); return false; }
    return true;
}

// ===== ORIGINAL‑LOGIK (angepasst auf HTML) =====
// Aus k10_excelTabelle_dialog.js – ES3/JScript. Functionality unverändert.

// Globale Variablen wie im Original
const global = {};
let bContentsChanged = false;
let userAuswahlElement;
let selectedIndex = -1;
let arrayTabelle = [];
let userAuswahl = '';
let message = ['', ''];
let form;

async function onLoad() {
    try {
        form = document.getElementById('excelTabelle');
        const startBtn = document.getElementById('idButtonStart');
        if (startBtn && typeof startBtn.focus === 'function') startBtn.focus();
        userAuswahlElement = document.getElementById('idAuswahlZeilen');
        await trennzeichen();
        await separator();
        await loadDefsInDefinitions();
        // always try to read the user's personal definitions (may be absent) without throwing
                if (userAuswahlElement) {
                    // load last used filename from profile (fallback to csvDefinitionUser.txt)
                    var lastFile = await getProfileString('Exceltool', 'LastUserFile', 'csvDefinitionUser.txt');
                    // ensure hidden input for last filename exists so save knows the target
                    var inp = document.getElementById('idSaveAsFileName');
                    if (!inp) { inp = document.createElement('input'); inp.type = 'hidden'; inp.id = 'idSaveAsFileName'; inp.name = 'idSaveAsFileName'; form.appendChild(inp); }
                    inp.value = lastFile;
                userAuswahlElement.value = await getFileContent('ProfD', 'user\\\\' + lastFile, true, true);
                userAuswahl = userAuswahlElement.value;
                // create and sync hidden field used by runScript to send textarea content
                var hidden = document.getElementById('hid_idAuswahlZeilen');
                if (!hidden) { hidden = document.createElement('input'); hidden.type = 'hidden'; hidden.id = 'hid_idAuswahlZeilen'; hidden.name = 'idAuswahlZeilen'; form.appendChild(hidden); }
                hidden.value = escapeForExeScript(userAuswahlElement.value);
                userAuswahlElement.addEventListener('input', function () { hidden.value = escapeForExeScript(this.value); });
                // indicate which user file was loaded (same behaviour as opening via button)
                var lbl = document.getElementById('idLabelAuswahl'); if (lbl) lbl.innerHTML = 'Auswahl geladen: ' + lastFile;
        }
        // read selected table type from profile and load that configuration (falls back to default in loader)
        const tabTyp = await getProfileString('Exceltool', 'Typ_Tabelle', 'Standardtabelle');
        await waehleKonfigurationstabelle(tabTyp);
        // initialize display of user selection column from profile (UseUserConfig: '1' = show)
        try {
            const useUser = await getProfileString('Exceltool', 'UseUserConfig', '0');
            const chk = document.getElementById('checkUser');
            if (chk) chk.checked = (useUser === '1');
            // pass explicit boolean so mapping is unambiguous (1 => show, 0 => hide)
            await displayUser(useUser === '1');
        } catch (e) {
            // non-fatal; continue silently
        }
        const treeBody = document.getElementById('treeBody');
        if (treeBody) {
            treeBody.addEventListener('dblclick', function (e) {
                e = e || window.event;
                const target = e.target || e.srcElement;
                const row = _closestByClass(target, 'rowConfig');
                if (row) { setSelected(+row.dataset.index); waehleZeile(); }
            });
        }
        bContentsChanged = false;
    } catch (error) {
        alert(`Fehler beim Laden des Dialogs:\n${error.message}`);
    }
}

async function onAccept() {
    await frageSpeichern();
    message = ["Bitte warten bis Schlussmeldung angezeigt wird! \n WinIBW zeigt evtl. keine Reaktion bis zum Ende des Downloads."];
    const idErgebnisEl = document.getElementById('idErgebnis');
    if (idErgebnisEl) idErgebnisEl.hidden = false;
    const idPfadEl = document.getElementById('idPfad');
    if (idPfadEl) idPfadEl.innerText = message[1];

    try {
        const report = await runScript('__excelWriteCSV');
        if (!report) {
            alert('Die Liste konnte nicht erstellt werden');
            return;
        }
        message = report.split('\n');
        if (idErgebnisEl) idErgebnisEl.hidden = false;
        if (idPfadEl) idPfadEl.innerText = message[1];
    } catch (error) {
        alert(`Fehler beim Erstellen der Exceltabelle:\n${error.message}`);
    }
}

async function onCancel() {
    await frageSpeichern();
    closeDialog();
}

async function separator(sep) {
    try {
        if (typeof sep === 'undefined' || sep === null) {
            sep = await getProfileString('Exceltool', 'Separator', ',');
            const select = document.getElementById('idSeparator');
            // If parameter provided, set select and return
            const strSeparator = sep;
            if (select) {
                for (let k = 0; k < select.options.length; k++) {
                    if (select.options[k].value === strSeparator) {
                        select.selectedIndex = k;
                        break;
                    }
                }
            }
            return strSeparator;
        }
        await writeProfileString('Exceltool', 'Separator', sep);
    } catch (error) {
        alert(`Fehler beim Setzen der Dateiendung:\n${error.message}`);
    }
}

async function trennzeichen(tr) {
    try {
        const select = document.getElementById('idTextboxTrennzeichen');
        // If no parameter, get from UI or profile
        if (typeof tr === 'undefined' || tr === null) {
            tr = await getProfileString('Exceltool', 'Trennzeichen', ',');
        }
        await writeProfileString('Exceltool', 'Trennzeichen', tr);
        // If parameter provided, set select and return
        const strTrennzeichen = tr;
        if (select) select.value = strTrennzeichen;
        return strTrennzeichen;
    } catch (error) {
        alert(`Fehler beim Setzen des Trennzeichens:\n${error.message}`);
        return '';
    }
}

// ===== Hilfen
async function wikiWinibw() {
    try {
        await runScript('__wikiWinibw');
    } catch (error) {
        alert(`Fehler beim Öffnen der Hilfe:\n${error.message}`);
    }
}
async function wikiAnzeigen2() {
    try {
        await runScript('__wikiAnzeigen2');
    } catch (error) {
        alert(`Fehler beim Öffnen der Konfigurationshilfe:\n${error.message}`);
    }
}
async function wikiAnzeigen3() {
    try {
        await runScript('__wikiAnzeigen3');
    } catch (error) {
        alert(`Fehler beim Öffnen der Trennzeichenhilfe:\n${error.message}`);
    }
}

// ===== Konfig laden (angepasst auf Textarea statt XUL-Tree) =====

async function selectTabelle() {
    const el = document.getElementById('idTabelle');
    const val = el ? el.value : '';
    alert(`Ausgewählte Tabelle: ${val}`);
    try {
        await waehleKonfigurationstabelle(val);
    } catch (error) {
        alert(`Fehler beim Wechseln der Tabelle:\n${error.message}`);
    }
}

// Toggle display of the user-selection column (Auswahl)
// If `forceShow` is provided (boolean), use it instead of reading the checkbox state.
async function displayUser(forceShow) {
    try {
        const chk = document.getElementById('checkUser');
        const textarea = document.getElementById('idAuswahlZeilen');
        let rightCol = textarea ? textarea.closest('.col') : null;
        // fallback: second .row's second .col inside #tab_cfg
        if (!rightCol) {
            const rows = document.querySelectorAll('#tab_cfg .row');
            if (rows && rows.length > 1) {
                const cols = rows[1].getElementsByClassName('col');
                if (cols && cols.length > 1) rightCol = cols[1];
            }
        }

        const show = (typeof forceShow !== 'undefined') ? !!forceShow : !!(chk && chk.checked);
        // ensure checkbox reflects the chosen state
        if (chk) chk.checked = show;

        if (rightCol) rightCol.style.display = show ? '' : 'none';
        if (textarea) {
            textarea.disabled = !show;
            if (show) textarea.focus();
        }
        // persist preference in profile (1 = enabled, 0 = disabled)
        try { await writeProfileString('Exceltool', 'UseUserConfig', show ? '1' : '0'); } catch (e) { /* ignore */ }
    } catch (error) {
        alert('Fehler beim Umschalten der Benutzerauswahl:\n' + error.message);
    }
}

function findValue(o, v) {
    for (let i = 0; i < o.length; i++) {
        if (o[i].value === v) return i;
    }
    return -1;
}

async function waehleKonfigurationstabelle(auswahlTabelle) {
    const tbl = document.getElementById('idTabelle');
    if (tbl) tbl.selectedIndex = findValue(tbl.options, auswahlTabelle);
    await writeProfileString('Exceltool', 'Typ_Tabelle', auswahlTabelle);
    await ladeKonfigurationstabelle(auswahlTabelle);
}

async function ladeKonfigurationstabelle(fileName) {
    const DEFAULT_PATH = 'excelTool\\\\csvDefinition.txt';
    function normalizePath(p) {
        if (!p) return p;
        return p.replace(/\\+/g, '\\');
    }

    let content = '';
    try {
        if (typeof fileName !== 'undefined' && fileName.indexOf('csvDefinition') > -1) {
            content = await getFileContent('ProfD', normalizePath(fileName), true, true);
            if (!content) {
                // try default once
                content = await getFileContent('ProfD', DEFAULT_PATH, true, true);
            }
        } else {
            content = await getFileContent('ProfD', DEFAULT_PATH, true, true);
        }
    } catch (error) {
        alert('Fehler beim Laden der CSV-Definition:\n' + error.message);
        return;
    }

    if (!content) {
        alert('Fehler beim Laden der CSV-Definition: ' + DEFAULT_PATH);
        return;
    }

    arrayTabelle = content.split('\n');
    renderTree(arrayTabelle);
}

async function frageSpeichern() {
    if (bContentsChanged) {
        if (confirm('Änderungen in der Konfiguration speichern?')) await auswahlSpeichern();
    }
}

async function auswahlSpeichern() {
    try {
        // save current textarea to last used filename (use save-as backend)
        var inp = document.getElementById('idSaveAsFileName');
        // ensure current textarea content is passed to backend
        var ta = document.getElementById('idAuswahlZeilen');
        var textVal = ta ? ta.value : '';
        var hiddenTA = document.getElementById('idAuswahlZeilen');
        // create a hidden input for form submission if the textarea isn't considered by runScript
        // (some runtimes only serialize inputs by name)
        var hidden = document.getElementById('hid_idAuswahlZeilen');
        if (!hidden) {
            hidden = document.createElement('input'); hidden.type = 'hidden'; hidden.id = 'hid_idAuswahlZeilen'; hidden.name = 'idAuswahlZeilen'; form.appendChild(hidden);
        }
        hidden.value = escapeForExeScript(textVal);
        if (!inp) {
            inp = document.createElement('input'); inp.type = 'hidden'; inp.id = 'idSaveAsFileName'; inp.name = 'idSaveAsFileName'; form.appendChild(inp);
            inp.value = 'csvDefinitionUser.txt';
        }
        await runScript('__excelWriteAuswahlAs');
        const lbl = document.getElementById('idLabelAuswahl'); if (lbl) lbl.innerHTML = 'Neue Auswahl gespeichert: ' + inp.value;
        bContentsChanged = false;
        await trennzeichen();
    } catch (error) {
        alert(`Fehler beim Speichern der Auswahl:\n${error.message}`);
    }
}

/*
async function auswahlLoeschen() {
    try {
        if (window.confirm('Soll Ihre Auswahl und Ihre persönliche Konfigurationstabelle gelöscht werden?')) {
            if (userAuswahlElement) userAuswahlElement.value = '';
            const lbl = document.getElementById('idLabelAuswahl'); if (lbl) lbl.innerHTML = 'Auswahl gelöscht.';
            await waehleKonfigurationstabelle('Standardtabelle');
            await writeProfileString('Exceltool', 'Typ_Tabelle', 'Standardtabelle');
            bContentsChanged = false;
        }
    } catch (error) {
        alert(`Fehler beim Löschen der Auswahl:\n${error.message}`);
    }
}*/

function _closestByClass(el, className) {
    while (el && el.nodeType === 1) {
        const cn = el.className || '';
        if ((` ${cn} `).indexOf(` ${className} `) !== -1) return el;
        el = el.parentNode;
    }
    return null;
}
// Your original double-click handler target
function waehleZeile() {
    if (selectedIndex < 0 || selectedIndex >= arrayTabelle.length) return;
    const value = arrayTabelle[selectedIndex];
    // Use current textarea contents as base (covers manual clears and deletions)
    const dst = document.getElementById('idAuswahlZeilen');
    const base = dst ? dst.value : userAuswahl;
    userAuswahl = (base ? base + '\n' : '') + value;
    bContentsChanged = true;
    const lbl = document.getElementById('idLabelAuswahl'); if (lbl) lbl.innerHTML = 'Auswahl geändert.';
    // write updated selection into the textarea
    if (dst) {
        dst.value = userAuswahl;
        dst.focus();
        dst.selectionStart = dst.selectionEnd = dst.value.length;
    } else if (userAuswahlElement) {
        userAuswahlElement.value = userAuswahl;
    }
}


function renderTree(data) {
    const body = document.getElementById('treeBody');
    if (!body) return;
    body.innerHTML = '';
    selectedIndex = -1;

    for (let i = 0; i < data.length; i++) {
        const div = document.createElement('div');
        div.className = 'rowConfig row';
        div.setAttribute('role', 'row');
        div.setAttribute('aria-selected', 'false');
        div.dataset.index = i;
        div.textContent = data[i];

        div.addEventListener('click', function (e) {
            const idx = +e.currentTarget.dataset.index;
            setSelected(idx);
        });

        body.appendChild(div);
    }

    // focus for keyboard navigation
    if (data.length) {
        body.focus();
        setSelected(0);
    }
}

function setSelected(idx) {
    const body = document.getElementById('treeBody');
    if (!body) return;
    const rows = body.querySelectorAll('.rowConfig');
    if (!rows.length) return;

    if (selectedIndex >= 0 && rows[selectedIndex]) rows[selectedIndex].setAttribute('aria-selected', 'false');

    selectedIndex = Math.max(0, Math.min(idx, rows.length - 1));
    rows[selectedIndex].setAttribute('aria-selected', 'true');
}


/**
 * Prepare hidden form fields with directory and path values and invoke the backend script to retrieve file content.
 *
 * Ensures a form element accessible via document.getElementsByName('form')[0] exists, then ensures two hidden
 * inputs (id/name 'etDirectory' and 'etFilePath') are present on that form. Sets their values to the provided
 * dir and path arguments, respectively, and finally calls runScript('GetFileContent') returning its result.
 *
 * Note: This function performs DOM mutations (may create and append hidden inputs) and relies on the presence
 * of a global runScript function. If the expected form or runScript are not present, the function may throw.
 *
 * @param {string} dir - The directory value to be written to the hidden input 'etDirectory'.
 * @param {string} path - The path value to be written to the hidden input 'etFilePath'.
 * @returns {*} The value returned by runScript('GetFileContent') — type depends on that implementation.
 * @throws {TypeError} If the form element named 'form' is not present (so appendChild will fail).
 * @throws {ReferenceError} If runScript is not defined in the global scope.
 */
async function getFileContent(dir, path, noComments, noBlanks) {
    //alert('Die Datei ' + path + ' wird geladen...');
    if (typeof noComments === 'undefined') noComments = false;
    if (typeof noBlanks === 'undefined') noBlanks = false;
    let inputDir = document.getElementById('etDirectory');
    if (!inputDir) {
        inputDir = document.createElement('input');
        inputDir.type = 'hidden';
        inputDir.id = 'etDirectory';
        inputDir.name = 'etDirectory';
        form.appendChild(inputDir);
    }
    inputDir.value = dir;

    let inputPath = document.getElementById('etFilePath');
    if (!inputPath) {
        inputPath = document.createElement('input');
        inputPath.type = 'hidden';
        inputPath.id = 'etFilePath';
        inputPath.name = 'etFilePath';
        form.appendChild(inputPath);
    }
    inputPath.value = path;

    let inputNoComments = document.getElementById('noComments');
    if (!inputNoComments) {
        inputNoComments = document.createElement('input');
        inputNoComments.type = 'hidden';
        inputNoComments.id = 'noComments';
        inputNoComments.name = 'noComments';
        form.appendChild(inputNoComments);
    }
    inputNoComments.value = noComments ? '1' : '0';

    let inputNoBlanks = document.getElementById('noBlanks');
    if (!inputNoBlanks) {
        inputNoBlanks = document.createElement('input');
        inputNoBlanks.type = 'hidden';
        inputNoBlanks.id = 'noBlanks';
        inputNoBlanks.name = 'noBlanks';
        form.appendChild(inputNoBlanks);
    }
    inputNoBlanks.value = noBlanks ? '1' : '0';
    try {
        const content = await runScript('__excelGetFileContent');
        return content;
    } catch (error) {
        alert(`Error: ${error.message}`);
        return '';
    }
}


/**
 * Populates the <select> element with id "idTabelle" from a newline-separated list
 * returned by runScript('__excelLoadFilesInDefinitions()').
 *
 * Behavior:
 * - Calls runScript('__excelLoadFilesInDefinitions()') and returns early if the result is falsy.
 * - Splits the returned string on '\n' to obtain candidate filenames.
 * - Trims CR characters and surrounding whitespace from each filename using an ES3-compatible approach.
 * - Skips empty entries and filenames already present as option values in the target select.
 * - For each remaining filename, creates an <option> element, sets both value and text to the filename,
 *   and appends it to the select element.
 * - Returns early if the select element with id "idTabelle" cannot be found.
 *
 * Side effects: Mutates the DOM by appending <option> elements to the select#idTabelle.
 *
 * @function loadDefsInDefinitions
 * @returns {void} No value is returned.
 * @see runScript
 */
async function loadDefsInDefinitions() {
    const defFiles = await runScript('__excelLoadFilesInDefinitions()');
    if (!defFiles) return;
    const defArray = defFiles.split('\n');
    const select = document.getElementById('idTabelle');
    if (!select) return;

    for (let i = 0; i < defArray.length; i++) {
        let fname = defArray[i];
        if (!fname) continue;
        // trim CR and surrounding whitespace (ES3-compatible)
        fname = fname.replace(/\r/g, '').replace(/^\s+|\s+$/g, '');
        if (!fname) continue;

        // skip duplicates
        let exists = false;
        for (let j = 0; j < select.options.length; j++) {
            if (select.options[j].value === fname) { exists = true; break; }
        }
        if (exists) continue;

        const opt = document.createElement('option');
        opt.value = fname;
        opt.text = fname;
        select.appendChild(opt);
    }
}

// Save-as handler: prompt for filename and write user configuration under ProfD\user\
async function auswahlSpeichernAls() {
    try {
        // default to current last-used filename if present
        var defInp = document.getElementById('idSaveAsFileName');
        var defaultName = (defInp && defInp.value) ? defInp.value : 'csvDefinitionUser.txt';
        var fname = window.prompt('Dateiname für die persönliche Konfiguration (ohne Pfad):', defaultName);
        if (!fname) return;
        // sanitize filename: remove any path separators and trim
        fname = String(fname).replace(/[\\/]/g, '').replace(/^\s+|\s+$/g, '');
        if (fname === '') return;
        // ensure extension
        if (fname.indexOf('.') < 0) fname += '.txt';

        // ensure textarea content is passed to backend as well
        var ta = document.getElementById('idAuswahlZeilen');
        var textVal = ta ? ta.value : '';
        var hidden = document.getElementById('hid_idAuswahlZeilen');
        if (!hidden) { hidden = document.createElement('input'); hidden.type = 'hidden'; hidden.id = 'hid_idAuswahlZeilen'; hidden.name = 'idAuswahlZeilen'; form.appendChild(hidden); }
        hidden.value = escapeForExeScript(textVal);

        // create hidden input to pass filename to the backend script
        var inp = document.getElementById('idSaveAsFileName');
        if (!inp) {
            inp = document.createElement('input');
            inp.type = 'hidden'; inp.id = 'idSaveAsFileName'; inp.name = 'idSaveAsFileName';
            form.appendChild(inp);
        }
        inp.value = fname;

        await runScript('__excelWriteAuswahlAs');
        const lbl = document.getElementById('idLabelAuswahl'); if (lbl) lbl.innerHTML = 'Auswahl gespeichert als ' + fname;
        bContentsChanged = false;
    } catch (error) {
        alert('Fehler beim Speichern als:\n' + error);
    }
}

// Open handler: show list of files under ProfD\user and load selected into textarea
async function auswahlOeffnen() {
    try {
        const filesStr = await runScript('__excelLoadFilesInUser()');
        if (!filesStr) { alert('Keine Dateien unter Benutzerkonfiguration gefunden.'); return; }
        const files = filesStr.split('\n').map(function(s) { return s.replace(/\r/g,'').replace(/^\s+|\s+$/g,''); }).filter(function(s){return s!=='';});
        if (!files.length) { alert('Keine Dateien unter Benutzerkonfiguration gefunden.'); return; }

        // create simple modal selector
        const overlay = document.createElement('div');
        overlay.style.position = 'fixed'; overlay.style.left = 0; overlay.style.top = 0; overlay.style.right = 0; overlay.style.bottom = 0;
        overlay.style.background = 'rgba(0,0,0,0.4)'; overlay.style.zIndex = 9999; overlay.style.display = 'flex'; overlay.style.alignItems = 'center'; overlay.style.justifyContent = 'center';

        const box = document.createElement('div');
        box.style.background = '#fff'; box.style.padding = '12px'; box.style.borderRadius = '6px'; box.style.width = '420px'; box.style.maxHeight = '70%'; box.style.overflow = 'auto';

        const title = document.createElement('div'); title.innerText = 'Datei aus Benutzerverzeichnis öffnen'; title.style.fontWeight = 'bold'; title.style.marginBottom = '8px';
        const sel = document.createElement('select'); sel.size = Math.min(10, files.length); sel.style.width = '100%'; sel.style.marginBottom = '8px';
        for (let i = 0; i < files.length; i++) { const o = document.createElement('option'); o.value = files[i]; o.text = files[i]; sel.appendChild(o); }

        const btnRow = document.createElement('div'); btnRow.style.textAlign = 'right';
        const ok = document.createElement('button'); ok.type = 'button'; ok.className = 'button'; ok.innerText = 'Öffnen';
        const cancel = document.createElement('button'); cancel.type = 'button'; cancel.className = 'cancel'; cancel.innerText = 'Abbrechen';
        btnRow.appendChild(cancel); btnRow.appendChild(ok);

        box.appendChild(title); box.appendChild(sel); box.appendChild(btnRow); overlay.appendChild(box); document.body.appendChild(overlay);

        cancel.addEventListener('click', function() { document.body.removeChild(overlay); });
        ok.addEventListener('click', async function() {
            const chosen = sel.value; if (!chosen) { alert('Keine Datei gewählt.'); return; }
            // create hidden input for backend (open) and update save-as hidden input
            var inp = document.getElementById('idOpenFileName');
            if (!inp) { inp = document.createElement('input'); inp.type = 'hidden'; inp.id = 'idOpenFileName'; inp.name = 'idOpenFileName'; form.appendChild(inp); }
            inp.value = chosen;
            var saveInp = document.getElementById('idSaveAsFileName');
            if (!saveInp) { saveInp = document.createElement('input'); saveInp.type = 'hidden'; saveInp.id = 'idSaveAsFileName'; saveInp.name = 'idSaveAsFileName'; form.appendChild(saveInp); }
            saveInp.value = chosen;
            try {
                const content = await runScript('__excelReadUserFile');
                const ta = document.getElementById('idAuswahlZeilen');
                if (ta) { ta.value = content; ta.focus(); }
                    // update hidden field used for runScript submission
                    var hidden = document.getElementById('hid_idAuswahlZeilen');
                    if (!hidden) { hidden = document.createElement('input'); hidden.type = 'hidden'; hidden.id = 'hid_idAuswahlZeilen'; hidden.name = 'idAuswahlZeilen'; form.appendChild(hidden); }
                    hidden.value = escapeForExeScript(content);
                const lbl = document.getElementById('idLabelAuswahl'); if (lbl) lbl.innerHTML = 'Auswahl geladen: ' + chosen;
                bContentsChanged = false;
            } catch (e) { alert('Fehler beim Laden der Datei:\n' + e); }
            document.body.removeChild(overlay);
        });
    } catch (error) {
        alert('Fehler beim Öffnen der Benutzerdatei:\n' + error);
    }
}

// Escape textarea / filename content so the dialog-side serialisation
// (which only escapes quotes) produces a syntactically valid JS literal
// when the string is inlined into the script call by `W4DialogFunctions.js`.
function escapeForExeScript(s) {
    if (s === null || s === undefined) return '';
    var t = String(s)
        .replace(/\\/g, '\\\\')   // backslash -> double-backslash
        .replace(/\r\n/g, '\\x1E')    // CRLF -> \\x1E (record separator)
        .replace(/\r/g, '\\x1E')       // CR -> \\x1E
        .replace(/\n/g, '\\x1E');      // LF -> \\x1E
    return t.replace(/[\u0080-\uFFFF]/g, function (c) {
        var code = c.charCodeAt(0).toString(16).toUpperCase();
        while (code.length < 4) code = '0' + code;
        return '\\u' + code;
    });
}