// --- Tabs ---
function showTab(id) {
    // Panels
    var panels = document.getElementsByClassName('panel');
    for (var j = 0; j < panels.length; j++) { panels[j].className = 'panel'; }
    document.getElementById('tab_' + id).className = 'panel active';
    // Buttons
    var btns = document.getElementsByClassName('tab');
    for (var k = 0; k < btns.length; k++) { btns[k].className = 'tab'; }
    document.getElementById('btn_' + (id == 'cfg' ? 'cfg' : id)).className = 'tab active';
}

// ===== HILFSFUNKTIONEN für Textarea-"Tree" =====
function _getCurrentLine(el) {
    var start = el.selectionStart, val = el.value;
    var lineStart = start; while (lineStart > 0 && val.charAt(lineStart - 1) != '\n') lineStart--;
    var lineEnd = start; while (lineEnd < val.length && val.charAt(lineEnd) != '\n') lineEnd++;
    return val.substring(lineStart, lineEnd);
}
// function waehleZeile() removed (duplicate)
function handle_key_press_auswahl(evt) {
    evt = evt || window.event; var code = evt.keyCode || evt.which;
    if (code == 13) { if (evt.preventDefault) evt.preventDefault(); waehleZeile(); return false; }
    return true;
}

// ===== ORIGINAL‑LOGIK (angepasst auf HTML) =====
// Aus k10_excelTabelle_dialog.js – ES3/JScript. Functionality unverändert.

// Globale Variablen wie im Original
var global = {},
    bContentsChanged = false,
    userAuswahlElement,
    selectedIndex = -1,
    arrayTabelle = [],
    userAuswahl = '',
    message = ['', ''];

function onLoad() {
    document.getElementById('idButtonStart').focus();
    userAuswahlElement = document.getElementById('idAuswahlZeilen');
    trennzeichen();
    separator()
    einstellungKonfigurationstabelle();
    userAuswahlElement.value = getFileContent('ProfD', 'user\\\\csvDefinitionUser.txt', true, true);
    ladeKonfigurationstabelle();
    document.getElementById('treeBody').addEventListener('dblclick', function (e) {
        e = e || window.event;
        var target = e.target || e.srcElement;
        var row = _closestByClass(target, 'rowConfig');
        if (row) { setSelected(+row.dataset.index); waehleZeile(); }
    });
    bContentsChanged = false;
}

function onAccept() {
    frageSpeichern();
    message = ['Bitte warten bis Schlussmeldung angezeigt wird!',  'WinIBW zeigt evtl. keine Reaktion bis zum Ende des Downloads.'];
    document.getElementById('idLabelErgebnis1').innerHTML = message[0];
    document.getElementById('idLabelErgebnis2').innerHTML = message[1];
    document.getElementById('idTextboxPfad').value = '';

    try {
        var message = runScript('__excelWriteCSV')
        if (!message) {
            alert('Die Liste konnte nicht erstellt werden');
            return;
        }
        var report = message.split("\n");
        document.getElementById('idLabelErgebnis1').innerHTML = report[0];
        document.getElementById('idTextboxPfad').value = report[1];
        alert('Die Exceltabelle wurde erstellt:\n' + report[1]);
    } catch (e) {
        alert('Fehler beim Erstellen der Exceltabelle:\n' + e.message);
    }
}

function onCancel() {
    frageSpeichern();
    closeDialog();
}

function separator(sep) {
    if(typeof sep === 'undefined' || sep === null) {
        sep = getProfileString("Exceltool", "Separator", ",");
        var select = document.getElementById('idSeparator');
        // If parameter provided, set select and return
        var strSeparator = sep;
        for (var k = 0; k < select.options.length; k++) {
            if (select.options[k].value === strSeparator) {
                select.selectedIndex = k;
                break;
            }
        }
        return strSeparator;
    }
    writeProfileString("Exceltool", "Separator", sep);
}

function trennzeichen(tr) {
    var select = document.getElementById('idTextboxTrennzeichen');
    // If no parameter, get from UI or profile
    if (typeof tr === 'undefined' || tr === null) {
        tr = getProfileString("Exceltool", "Trennzeichen", ",");
    }
    writeProfileString("Exceltool", "Trennzeichen", tr);
    // If parameter provided, set select and return
    var strTrennzeichen = tr;
    select.value = strTrennzeichen;
    return strTrennzeichen;
}

// ===== Hilfen
function wikiWinibw() { runScript('__wikiWinibw'); }
function wikiAnzeigen2() { runScript('__wikiAnzeigen2'); }
function wikiAnzeigen3() { runScript('__wikiAnzeigen3'); }

// ===== Konfig laden (angepasst auf Textarea statt XUL-Tree) =====

function einstellungKonfigurationstabelle() {
    waehleKonfigurationstabelle(getProfileInt("Exceltool", "Typ_Tabelle", 0));
}

function selectTabelle() {
    var auswahlTabelle = document.getElementById("idTabelle").selectedIndex; //gibt 0 oder 1 aus
    if (1 === auswahlTabelle) {
        alert("Ihre Konfigurationsdatei wird verwendet.");
    } else {
        alert("Die Standard-Konfigurationsdatei wird verwendet.");
    }
    waehleKonfigurationstabelle(auswahlTabelle)
}


function waehleKonfigurationstabelle(auswahlTabelle) {
    //wenn eigene Konfigurationsdatei leer, kann diese nicht ausgewählt werden
    if (auswahlTabelle === 1) {
        userAuswahlElement.disabled = false;
        document.getElementById("idTabelle").selectedIndex = 1;
        writeProfileInt("Exceltool", "Typ_Tabelle", 1);
    } else {
        userAuswahlElement.disabled = true;
        document.getElementById("idTabelle").selectedIndex = 0;
        writeProfileInt("Exceltool", "Typ_Tabelle", 0);
    }
    return auswahlTabelle;
}

function ladeKonfigurationstabelle() {
    var standard = getFileContent('ProfD', 'ttlFiles_zdb\\\\zdb_csvDefinition.txt', true, true);
    if (!standard) {
        alert("Fehler beim Laden der Default-CSV-Definition.");
        onCancel();
    }
    //document.getElementById('idDefault').value = global.default;
    arrayTabelle = standard.split('\n');
    renderTree(arrayTabelle);
}

function frageSpeichern() {
    if (bContentsChanged) {
        if (confirm('Änderungen in der Konfiguration speichern?')) {
            auswahlSpeichern()
        }
    }
}

function auswahlSpeichern() {
    runScript('__excelWriteAuswahl');
    document.getElementById('idLabelAuswahl').innerHTML = 'Neue Auswahl gespeichert.';
    bContentsChanged = false;
    trennzeichen();
    bContentsChanged = false;
}

function auswahlLoeschen() {
    if (window.confirm('Soll Ihre Auswahl und Ihre persönliche Konfigurationstabelle gelöscht werden?')) {
        userAuswahlElement.value = '';
        document.getElementById('idLabelAuswahl').innerHTML = 'Auswahl gelöscht.';
        waehleKonfigurationstabelle(0);
        writeProfileInt("Exceltool", "Typ_Tabelle", 0);
        bContentsChanged = false;
    }
}

function _closestByClass(el, className) {
    while (el && el.nodeType === 1) {
        var cn = el.className || '';
        if ((' ' + cn + ' ').indexOf(' ' + className + ' ') !== -1) return el;
        el = el.parentNode;
    }
    return null;
}
// Your original double-click handler target
function waehleZeile() {
    if (selectedIndex < 0 || selectedIndex >= arrayTabelle.length) return;
    var value = arrayTabelle[selectedIndex];
    userAuswahl = (userAuswahl ? userAuswahl + "\n" : '') + value;
    bContentsChanged = true;
    var lbl = document.getElementById('idLabelAuswahl');
    if (lbl) lbl.innerHTML = 'Auswahl geändert.';
    var dst = document.getElementById('idAuswahlZeilen');
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
    var body = document.getElementById('treeBody');
    body.innerHTML = "";
    selectedIndex = -1;

    for (var i = 0; i < data.length; i++) {
        var div = document.createElement('div');
        div.className = 'rowConfig row';
        div.setAttribute('role', 'row');
        div.setAttribute('aria-selected', 'false');
        div.dataset.index = i;
        div.textContent = data[i];

        div.addEventListener('click', function (e) {
            var idx = +e.currentTarget.dataset.index;
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
    var body = document.getElementById('treeBody');
    var rows = body.querySelectorAll('.rowConfig');
    if (!rows.length) return;

    if (selectedIndex >= 0 && rows[selectedIndex]) {
        rows[selectedIndex].setAttribute('aria-selected', 'false');
    }

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
function getFileContent(dir, path, noComments, noBlanks) {
    if (typeof noComments === 'undefined') noComments = false;
    if (typeof noBlanks === 'undefined') noBlanks = false;
    var form = document.getElementById('excelTabelle');
    var inputDir = document.getElementById('etDirectory');
    if (!inputDir) {
        inputDir = document.createElement('input');
        inputDir.type = 'hidden';
        inputDir.id = 'etDirectory';
        inputDir.name = 'etDirectory';
        form.appendChild(inputDir);
    }
    inputDir.value = dir;

    var inputPath = document.getElementById('etFilePath');

    if (!inputPath) {
        inputPath = document.createElement('input');
        inputPath.type = 'hidden';
        inputPath.id = 'etFilePath';
        inputPath.name = 'etFilePath';
        form.appendChild(inputPath);
    }
    inputPath.value = path;

    var inputNoComments = document.getElementById('noComments');

    if (!inputNoComments) {
        inputNoComments = document.createElement('input');
        inputNoComments.type = 'hidden';
        inputNoComments.id = 'noComments';
        inputNoComments.name = 'noComments';
        form.appendChild(inputNoComments);
    }
    inputNoComments.value = noComments ? '1' : '0';

    var inputNoBlanks = document.getElementById('noBlanks');

    if (!inputNoBlanks) {
        inputNoBlanks = document.createElement('input');
        inputNoBlanks.type = 'hidden';
        inputNoBlanks.id = 'noBlanks';
        inputNoBlanks.name = 'noBlanks';
        form.appendChild(inputNoBlanks);
    }
    inputNoBlanks.value = noBlanks ? '1' : '0';
    try {
        var content = runScript('__getFileContent');
        return content;
    } catch (e) {
        alert("Error: " + e.message);
    }
}

// Init beim Laden
// Call onLoad when the document is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', onLoad);
} else {
    onLoad();
    document.getElementById('idLabelErgebnis1').innerHTML = message[0];
    document.getElementById('idLabelErgebnis2').innerHTML = message[1];
}
