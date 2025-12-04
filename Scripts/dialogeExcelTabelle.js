function excelTabelle() {
    showDialog('ProfD\\excelTool\\dialogExcelTabelle.html', 200, 100, 800, 450);
}

/**
 *
 * Read the contents of a file (line by line) and send the resulting text to a dialog.
 *
 * The function attempts to open a file using utility.newFileInput().openSpecial(dir, "\\" + path).
 * If the file cannot be opened, utility.sentDataToDialog(false) is invoked and the function returns.
 * When opened successfully, the file is read line-by-line. Lines may be conditionally skipped:
 * - lines starting with "//" can be skipped if o.zdbNoComments is truthy,
 * - blank lines can be skipped if o.zdbNoBlanks is truthy.
 * After processing the lines the collected content is delivered via utility.sentDataToDialog(inhalt).
 *
 * Note: This comment documents the intended behavior of the implementation. The current source
 * contains a few implementation issues that affect behavior (for example: duplicate variable
 * declarations overwrite flags, a referenced noBlanksFlag identifier is not defined, and the
 * collected content variable may not be appended to). Those issues should be resolved in code
 * for the function to behave as described here.
 *
 * @param {Object} o - Options object controlling file selection and filtering.
 * @param {string} o.theDir - Directory (special/open context) used by openSpecial.
 * @param {string} o.thePath - Relative path or filename to open (will be prefixed with a backslash).
 * @param {boolean} [o.zdbNoComments=false] - If true, skip lines that begin with "//".
 * @param {boolean} [o.zdbNoBlanks=false] - If true, skip blank/empty lines.
 *
 * @returns {void} This function does not return a value. On success it calls utility.sentDataToDialog(inhalt)
 *                   where inhalt is the concatenated/processed file content; on open failure it calls
 *                   utility.sentDataToDialog(false).
 *
 * @throws {Error} No explicit exceptions are thrown by this function in normal operation; underlying
 *                 utility methods may raise errors depending on their implementations.
 */
function __getFileContent(o) {
    var dir = o.etDirectory,
        path = o.etFilePath,
        noCommentsFlag = o.noComments,
        noBlanksFlag = o.noBlanks,
        zeile,
        inhalt = '',
        defInpFile = utility.newFileInput();
    if (!defInpFile.openSpecial(dir, "\\" + path)) {
        utility.sentDataToDialog(false);
        return;
    }

    for (zeile = ""; !defInpFile.isEOF();) {
        zeile = defInpFile.readLine();
        if (noCommentsFlag == '1' && zeile.substring(0, 2) === "//") {
            continue;
        }
        // filter blank lines (preserve existing behaviour or conditionalize if needed)
        if (noBlanksFlag == '1' && zeile.length === 0) {
            continue;
        }
        inhalt += zeile + "\n";
    }
    defInpFile.close();
    utility.sentDataToDialog(inhalt);
}

function __excelWriteAuswahl(o) {
    var newContents = utility.restoreStringData(o.idAuswahlZeilen),
        out = utility.newFileOutput();
    out.createSpecial('ProfD', 'user\\csvDefinitionUser.txt');
    out.setTruncate(true);
    out.write(newContents);
    out.close();
    utility.sentDataToDialog(newContents);
}

function __wikiWinibw() {
    shellExecute('https://wiki.k10plus.de/x/agDUAw', 'open', '');
}

function __wikiAnzeigen2() {
    shellExecute('https://wiki.k10plus.de/x/agDUAw#Excel-Tabelleerstellen-KonfigurationdesExcel-Werkzeugs', 'open', '');
}


function __wikiAnzeigen3() {
    shellExecute('https://wiki.k10plus.de/x/agDUAw#Excel-Tabelleerstellen-Trennzeichen', 'open', '');
}

var excelVars = {
    strSystem: '',
    csvLevel2: false,
    csvDefinitions: '',
    strSST: '',
    feldSST: '',
    strTrennzeichen: '',
    sbfDescriptor: String.fromCharCode(402),
    separator: ','
};

/**
 * Generates a CSV file from the current selection of titles or records.
 * 
 * This function reads CSV definitions, processes each record, and writes the results to a CSV file
 * in a special directory. It updates the UI with the result and the file path, and copies the path to the clipboard.
 * 
 * @throws {Error} If the script is not called from a valid context (Kurztitelliste or Präsentation eines Titels).
 * @returns {void|boolean} Returns false if CSV definitions cannot be read; otherwise, writes the CSV file and updates the UI.
 */
function __excelWriteCSV(o) {
    //var scr = runScript('excelGetScr()');
    var scr = application.activeWindow.getVariable('scr')
    if ((scr != '8A') && (scr != '7A')) {
        throw new Error(200, 'Das Skript kann nur aus einer Kurztitelliste oder der Präsentation eines Titels aufgerufen werden.');
    }
    excelVars.strSST = o.idTextboxSST;
    excelVars.feldSST = o.idTextboxFeldSST;
    excelVars.strTrennzeichen = activeWindow.getProfileString('Exceltool', 'Trennzeichen', ',');
    var content,
        ctrl,
        cnt,
        header,
        ergebnis = '',
        idx,
        listenPfad,
        satz,
        outval,
        ext;

    //excelVars.msgboxHeader = 'Schreiben einer CSV‑Datei';
    excelVars.strSystem = application.activeWindow.getVariable('P3GCN');
    cnt = parseInt(application.activeWindow.getVariable('P3GSZ'));
    try {
        if (0 === o.idTabelle.selectedIndex) {
            excelVars.csvDefinitions = __readControl(o.idDefault, true);
        } else {
            excelVars.csvDefinitions = __readControl(utility.restoreStringData(o.idAuswahlZeilen), false);
        }
    } catch (e) {
        alert('Fehler in Definition: ' + e.message);
        utility.sentDataToDialog(false);
    }
    content = __replaceDefinitionsWithLookup(excelVars.csvDefinitions);
    if (content === null) {
        utility.sentDataToDialog(false);
    }
    ctrl = __createCtrlArray(content);
    header = __createHeader(ctrl);

    // Verzeichnis listen unter ProfD anlegen + Datei schreiben
    var out = utility.newFileOutput();
    excelVars.separator = getProfileString('Exceltool', 'Separator', ',');
    if (excelVars.separator === ',') {
        ext = '.csv';
    } else if (excelVars.separator === '\t') {
        ext = '.tsv';
    } else {
        ext = '.txt';
    }
    var rel = 'listen\\liste_' + __exceldatumHeute() + ext;
    out.createSpecial('ProfD', rel);
    out.setTruncate(true);
    out.writeLine('\ufeff' + header);
    var outcnt = 0;
    ctrl.cnt = 0;
    for (idx = 1; idx <= cnt; idx++) {
        activeWindow.command('show ' + idx + ' p', false);
        if (activeWindow.status != 'OK') {
            continue;
        }
        satz = '\n' + __getExpansionFromP3VTX();
        satz = satz.replace(/\r/g, '\n');
        satz = satz.replace(/\u001b./g, '');
        outval = __handleRecord(satz, ctrl);
        if (outval !== '') {
            out.writeLine(outval);
            outcnt++;
        }
    }
    listenPfad = getSpecialPath("ProfD", rel);
    out.close();
    application.activeWindow.command('s d', false);
    application.activeWindow.command('s k', false);

    activeWindow.clipboard = listenPfad;

    ergebnis = ctrl.cnt + ' Zeilen für ' + outcnt + ' Titel ausgegeben.';
    if (outcnt != cnt) {
        ergebnis += 'Leider konnten in ' + (cnt - outcnt) +
            " Titeln die gesuchten Felder nicht gefunden werden.\n";
    }
    shellExecute(listenPfad, 'edit', '');
    utility.sentDataToDialog(ergebnis + '\nDie Datei wurde gespeichert unter: <a href="file://' + listenPfad + '">file://' + listenPfad + '</a>');
}


// ===== Kernlogik (unverändert zum Original, nur UI‑Bindungen) =====
function __readControl(inp, must) {
    //alert('readControl called with inp:\n' + inp + '\nmust: ' + must);
    var out = [],
        line,
        tmp,
        cnt = 0,
        idx,
        inArray = [];

    inArray = inp.split("\n");

    for (var i = 0; i < inArray.length; i += 1) {
        line = inArray[i];
        if (line === null) {
            if (must) {
                line = 'interne Definitionsdatei';
            } else {
                line = 'ausgewählte Kommandodatei';
            }
            throw new Error(200, 'Die ' + line + 'kann nicht verarbeitet werden!');
        }
        // skip comments
        if (line.substring(0, 2) == '//') {
            continue;
        }
        // normalize whitespace and tabs
        tmp = line.replace(/\t/g, ' ');
        tmp = tmp.replace(/\s+/g, ' ').replace(/\s+$/, '').replace(/^\s+/, '');
        if (tmp === '') {
            continue;
        }
        idx = tmp.indexOf(':');
        // if a space occurs before the colon -> error
        if (tmp.indexOf(' ') < idx) {
            throw new Error(200, 'Die Spaltenüberschriften dürfen keine Blanks enthalten.\nZeile:\n' + line);
        }

        if (idx < 0) {
            if (must) {
                throw new Error(200, 'Die interne Definitionsdatei ist fehlerhaft!\nZeile:\n' + line);
            }
            out.push(tmp);
            out.push(tmp);
        } else {
            out.push(tmp.substr(0, idx).replace(/\s+$/, '').replace(/^\s+/, ''));
            out.push(tmp.substr(idx + 1).replace(/\s+$/, '').replace(/^\s+/, ''));
        }
        cnt++;
    }
    if (cnt === 0) {
        throw new Error(200, 'Die Definitionsdatei ist leer');
    }
    return out;
}

function __getExpansionFromP3VTX() {
    var satz = application.activeWindow.getVariable('P3VTX');
    satz = satz.replace('<ISBD><TABLE>', '');
    satz = satz.replace('</TABLE>', '');
    satz = satz.replace(/\u001bI|\u001bN/g, '');
    satz = satz.replace(/<BR>/g, '\n');
    satz = satz.replace(/^$/gm, '');
    satz = satz.replace(/^Eingabe:.*$/gm, '');
    satz = satz.replace(/<a[^<]*>/gm, '');
    satz = satz.replace(/<\/a>/gm, '');
    return satz;
}


/**
 * Replaces definition values in the given content array using excelVars.csvDefinitions.
 * For each pair of elements in the array, the first is kept as-is, and the second is replaced
 * with its definition if available. If the definition is not found and the value is not a tag,
 * prompts the user with an informational dialog and may open a help URL.
 * 
 * @param {Array<string>} content - An array of strings, expected to be in key-value pairs.
 * @returns {Array<string>|null} A new array with definitions replaced, or null if the user chooses to read more information.
 */
function __replaceDefinitionsWithLookup(content) {
    var newc = [];
    for (var i = 0; i < content.length; i += 2) {
        var key = content[i];
        var mask = content[i + 1];
        var defval = null;

        // Lookup definition in excelVars.csvDefinitions
        for (var j = 0; j < excelVars.csvDefinitions.length; j += 2) {
            if (excelVars.csvDefinitions[j] === mask) {
                defval = excelVars.csvDefinitions[j + 1];
                break;
            }
        }

        newc.push(key);
        if (defval === null) {
            if (!__checkIfTag(mask)) {
                var p = utility.newPrompter();
                var antwort = p.confirmEx(
                    'Hinweis zur Konfigurationstabelle',
                    'Diese Zeile ist fehlerhaft:\n' + key + ': ' + mask + '\n\nInformationen zur Konfigurationstabelle finden Sie im WinIBW‑Wiki.\nWollen Sie die Informationen jetzt lesen?',
                    'Ja',
                    'Nein',
                    '',
                    '',
                    ''
                );
                if (antwort === 0) {
                    application.shellExecute('https://wiki.k10plus.de/x/agDUAw#Excel-Tabelleerstellen-KonfigurationdesExcel-Werkzeugs', 5, 'open', '');
                }
                return null;
            }
            newc.push(mask);
        } else {
            newc.push(defval);
        }
    }
    return newc;
}


// Valid prefix characters for tag checking
var VALID_PREFIX_CHARS = 'KS';

/**
 * Checks if the given text matches a specific tag pattern.
 *
 * The function validates the structure of the input string according to a set of rules:
 * - Optionally starts with a character from VALID_PREFIX_CHARS (case-insensitive).
 * - Follows with a digit, possibly '2' (sets lev2 flag).
 * - Next three characters must be digits, with the last being an uppercase letter (A-Z).
 * - Optionally, a '/' or 'x' may follow, with further digit checks depending on lev2.
 * - The tag must end with a space character at the expected position.
 *
 * @param {string} text - The string to check for tag validity.
 * @returns {boolean} True if the text matches the tag pattern, false otherwise.
 */
function __checkIfTag(text) {
    var idx = 0,
        lev2;

    if (VALID_PREFIX_CHARS.indexOf(text.charAt(0).toUpperCase()) >= 0) {
        idx = 1;
    }
    lev2 = (text.charAt(idx) == '2');
    // Ensure text is long enough for the next 4 indices
    if (text.length < idx + 4) return false;
    if ((text.charAt(idx) < '0') || ('2' < text.charAt(idx++))) return false;
    if ((text.charAt(idx) < '0') || ('9' < text.charAt(idx++))) return false;
    if ((text.charAt(idx) < '0') || ('9' < text.charAt(idx++))) return false;
    if ((text.charAt(idx) < '@') || ('Z' < text.charAt(idx++))) return false;

    // Check for '/' or 'x' and ensure sufficient length for further checks
    if (text.charAt(idx) == '/') {
        if (lev2) return false;
        if (text.length < idx + 3) return false;
        idx++;
        if ((text.charAt(idx) < '0') || ('9' < text.charAt(idx++))) return false;
        if ((text.charAt(idx) < '0') || ('9' < text.charAt(idx++))) return false;
        if (text.charAt(idx) != ' ') {
            if (text.length < idx + 1) return false;
            if ((text.charAt(idx) < '0') || ('9' < text.charAt(idx++))) return false;
        }
    } else if (text.charAt(idx) == 'x') {
        if (!lev2) return false;
        if (text.length < idx + 3) return false;
        idx++;
        if ((text.charAt(idx) < '0') || ('9' < text.charAt(idx++))) return false;
        if ((text.charAt(idx) < '0') || ('9' < text.charAt(idx++))) return false;
    }

    if (text.length <= idx) return false;
    return (text.charAt(idx) == ' ');
}

function __createCtrlArray(content) {
    var tmpline,
        obj,
        out = [];

    for (var idx = 0; idx < content.length; idx += 2) {
        obj = {};
        obj.col = content[idx];
        obj.def = content[idx + 1];
        obj.val = '';
        obj.adr = 0;
        tmpline = __getSpecial(obj, content[idx + 1]);
        if (tmpline === null) return null;

        tmpline = __getTagInfos(obj, tmpline);
        if (tmpline === null) return null;

        tmpline = __orPartitions(obj, tmpline);
        if (tmpline === null) return null;

        out[idx / 2] = obj;
    }

    out.cnt = 0;
    return out;
}

/**
 * Extracts and sets a "special" character from the start of a line.
 *
 * If the provided tmpline begins with a digit, the function sets ctrl.spec
 * to a single space (' ') and returns tmpline unchanged. Otherwise it takes
 * the first character of tmpline, converts it to uppercase, assigns that
 * value to ctrl.spec, removes the first character from tmpline and returns
 * the remainder.
 *
 * This function mutates the ctrl object by assigning to its `spec` property.
 *
 * @param {{spec: string}} ctrl - Object whose `spec` property will be set.
 * @param {string} tmpline - Input string to inspect and modify.
 * @returns {string} The possibly-modified tmpline (with the first character removed unless it began with a digit).
 */
function __getSpecial(ctrl, tmpline) {
    //alert('getSpecial called with tmpline:\n' + tmpline);
    if (/^[0-9]/.test(tmpline)) {
        ctrl.spec = ' ';
    } else {
        ctrl.spec = tmpline.charAt(0).toUpperCase();
        tmpline = tmpline.substr(1);
    }
    return tmpline;
}

function __getTagInfos(ctrl, tmpline) {
    var idx;
    if (tmpline.charAt(0) == '2') {
        excelVars.csvLevel2 = true;
    }
    if (tmpline.charAt(4) == 'x') {
        //alert(excelVars.sbfDescriptor + ' == ? == ' + String.fromCharCode(402));
        ctrl.xsbf = excelVars.sbfDescriptor + tmpline.substr(4, 3);
        tmpline = tmpline.substr(0, 4) + tmpline.substr(7);
    } else {
        ctrl.xsbf = '';
    }
    idx = tmpline.indexOf(' ');
    ctrl.tag = tmpline.substr(0, idx);
    return tmpline.substr(idx + 1);
}


function __orPartitions(ctrl, tmpline) {
    var termOr = [],
        tmpObj = {},
        idx = 0;
    tmpline = ' ' + tmpline;
    while (tmpline.charAt(0) == ' ') {
        tmpline = __andPartitions(tmpObj, tmpline.substr(1));
        if (tmpline === null) return null;
        termOr[idx++] = tmpObj.termAnd;
    }
    ctrl.data = termOr;
    return tmpline;
}

function __andPartitions(termOr, tmpline) {
    var termAnd = [],
        tmpObj = {},
        idx = 0;
    tmpline = '+' + tmpline;
    while (tmpline.charAt(0) == '+') {
        tmpline = __sbfPart(tmpObj, tmpline.substr(1));
        if (tmpline === null) return null;
        termAnd[idx] = new Object();
        termAnd[idx++] = tmpObj.field;
    }

    termOr.termAnd = termAnd;
    // ensure tmpline is a non-empty string and starts with a space; otherwise signal a parse error
    if (tmpline.length > 0 && tmpline.charAt(0) !== ' ') {
        return null;
    }

    return tmpline;
}

/**
 * Parses a subfield part from the beginning of a token line and stores the result in the provided object.
 *
 * Handles two cases:
 *   1. The line starts with a $ (e.g. "$a"): parses as a simple subfield.
 *   2. The line starts with a quoted string containing a $ (e.g. "\"foo$a\""): parses as a subfield with prefix/suffix.
 * On success, sets obj.field = {pre, sbf, post} and returns the remaining tmpline.
 * On error, returns null.
 *
 * @param {Object} obj - The object to receive the parsed field structure.
 * @param {string} tmpline - The input string to parse.
 * @returns {string|null} The remaining unparsed string, or null on error.
 */
function __sbfPart(obj, tmpline) {
    // Defensive: return null if input is missing or not a string
    if (!tmpline || typeof tmpline !== "string") return null;

    var field = {}; // Will hold the parsed field structure

    // --- Case 1: Simple subfield, e.g. "$a"
    if (tmpline.charAt(0) === '$') {
        // No prefix or postfix, just the subfield code after '$'
        field.pre = '';
        // sbfDescriptor is a special character (e.g. ƒ) used as a subfield marker
        field.sbf = excelVars.sbfDescriptor + tmpline.charAt(1);
        field.post = '';
        // Remove the parsed part ("$a" = 2 chars)
        tmpline = tmpline.substr(2);
        // Store the result in the provided object
        obj.field = field;
        return tmpline;
    }

    // --- Case 2: Quoted string with a subfield, e.g. "\"foo$a\""
    if (tmpline.charAt(0) === '"') {
        // Remove the opening quote
        tmpline = tmpline.substr(1);
        // Find the closing quote
        var quoteIdx = tmpline.indexOf('"');
        // There must be at least one char before $ and $ not at end
        if (quoteIdx < 2) return null;
        // Extract the quoted part (e.g. "foo$a")
        var quoted = tmpline.substr(0, quoteIdx);
        // Remove the quoted part and closing quote from tmpline
        tmpline = tmpline.substr(quoteIdx + 1);

        // Find the $ in the quoted string
        var dollarIdx = quoted.indexOf('$');
        // $ must exist and not be the last character
        if (dollarIdx < 0 || dollarIdx === quoted.length - 1) return null;

        // Prefix is any text before the $
        field.pre = dollarIdx === 0 ? '' : quoted.substr(0, dollarIdx);
        // Subfield marker: sbfDescriptor + the char after $
        var afterDollar = quoted.substr(dollarIdx);
        field.sbf = excelVars.sbfDescriptor + afterDollar.charAt(1);
        // Postfix is any text after the subfield code
        field.post = afterDollar.substr(2);

        // Store the result in the provided object
        obj.field = field;
        return tmpline;
    }

    // --- If neither case matches, input is invalid
    return null;
}

function __createHeader(ctrl) {
    var idx = -1,
        header = '"PPN"' + excelVars.separator + '"EPN"' + excelVars.separator;
    while (++idx < ctrl.length) {
        header += '"' + ctrl[idx].col.replace(/\u0022/g, "'") + '"' + excelVars.separator;
    }
    header = header.replace(/;$/, '');
    return header;
}

function __handleRecord(satz, ctrl) {
    var lineblock = '', tmp_satz, tmp_line, idx, loopcnt, occ; loopcnt = __getMaxOccurrence(satz); for (idx = 1; idx <= loopcnt; idx++) { if (excelVars.strSystem == 'K10plus') { occ = '/0' + idx; if (idx < 10) { occ = '/00' + idx; } if (idx > 99) { occ = '/' + idx; } } else { occ = (idx < 10) ? '/0' + idx : '/' + idx; } tmp_satz = __filterCopy(satz, occ); if (tmp_satz !== '') { tmp_line = __handleRecordPart(tmp_satz, ctrl); if (tmp_line !== '') { lineblock += tmp_line + '\n'; } } }
    lineblock = lineblock.replace(/\n$/, '');
    if (lineblock === '' && excelVars.strSST === '') { tmp_satz = __filterCopy(satz, '/00'); lineblock = __handleRecordPart(tmp_satz, ctrl); }
    return lineblock;
}

function __getMaxOccurrence(satz) {
    var idx = satz.lastIndexOf('\n203@/');
    if (idx < 0) return 0;
    if (excelVars.strSystem == 'K10plus') {
        return (parseInt(satz.substr(idx + 6, 3), 10));
    } else {
        return (parseInt(satz.substr(idx + 6, 2), 10));
    }
}

function __filterCopy(satz, occ) {
    var tmp_satz = '',
        arr,
        idx,
        found = false;
    if (occ == '/00') {
        found = true;
    }
    arr = satz.split('\n');
    for (idx = 0; idx < arr.length; idx++) {
        if (arr[idx].charAt(0) != '2') {
            tmp_satz += arr[idx] + '\n';
        } else {
            if ((excelVars.strSystem == 'K10plus' && arr[idx].substr(4, 4) == occ) || (excelVars.strSystem != 'K10plus' && arr[idx].substr(4, 3) == occ)) {
                tmp_satz += arr[idx] + '\n'; found = true;
            }
        }
    }
    if (!found) {
        tmp_satz = '';
    }
    var regex4800 = new RegExp(excelVars.feldSST + ".+" + excelVars.strSST);
    return (regex4800.test(tmp_satz)) ? tmp_satz : '';
}

function __handleRecordPart(satz, ctrl) {
    var line, idx = -1; while (++idx < ctrl.length) { ctrl[idx].val = ''; ctrl[idx].adr = 0; }
    __createResult(satz, ctrl);
    var str7800 = '';
    line = '\"' + application.activeWindow.getVariable('P3GPP') + '\"' + excelVars.separator;
    idx = satz.indexOf('\n203@');
    if (idx < 0) { line += excelVars.separator; }
    else {
        if (excelVars.strSystem == 'K10plus') { str7800 = satz.substr(idx + 12, satz.length); }
        else { str7800 = satz.substr(idx + 11, satz.length); }
        str7800 = str7800.substring(0, str7800.indexOf('\n'));
        line += '\"' + str7800 + '\"' + excelVars.separator;
    }
    idx = -1; while (++idx < ctrl.length) {
        ctrl[idx].val = ctrl[idx].val.replace(/&amp;/g, '&');
        ctrl[idx].val = ctrl[idx].val.replace(/&lt;/g, '<');
        ctrl[idx].val = ctrl[idx].val.replace(/&gt;/g, '>');
        line += '"' + ctrl[idx].val.replace(/\u0022/g, "'") + excelVars.separator;
    }
    line = line.replace(/;$/, ''); ctrl.cnt++; return line;
}
function __createResult(satz, ctrl) {
    var tag, suche, regex, group, text, idx = -1, w; while (++idx < ctrl.length) { tag = ctrl[idx].tag; suche = tag + '.+' + ctrl[idx].xsbf; regex = new RegExp(suche, 'g'); group = satz.match(regex); if (group) { var tempArray = [], p; if (ctrl[idx].tag == '031N' || ctrl[idx].tag == '231@') { if (satz.indexOf(excelVars.sbfDescriptor + '0') != -1) { text = group[0].split(excelVars.sbfDescriptor + '0 '); for (p = 0; p < text.length; p++) { tempArray[p] = __convertOrText(text[p], ctrl[idx].spec, ctrl[idx].data); } ctrl[idx].val = tempArray.join(excelVars.strTrennzeichen); } else { for (w = 0; w < group.length; w++) { tempArray[w] = __convertOrText(group[w], ctrl[idx].spec, ctrl[idx].data); } } } else { for (w = 0; w < group.length; w++) { tempArray[w] = __convertOrText(group[w], ctrl[idx].spec, ctrl[idx].data); } if (tempArray.length > 1) { ctrl[idx].val = tempArray.join(excelVars.strTrennzeichen); } else { ctrl[idx].val = tempArray[0]; } } } }
    return;
}
function __convertOrText(text, spec, data) { var idx = -1, tmp; while (++idx < data.length) { tmp = __convertText(text, spec, data[idx]); if (tmp !== '') return tmp; } return ''; }
function __convertText(text, spec, andArr) {
    var tmp = '', idx = -1, idxe, jdxa, jdxe, test = false, tmpArray, tmpSf, jdxl; while (++idx < andArr.length) { jdxa = text.indexOf(andArr[idx].sbf); jdxl = text.lastIndexOf(andArr[idx].sbf); if (jdxa >= 0) { if (jdxa != jdxl) { tmpArray = []; while (test === false) { jdxe = text.indexOf(excelVars.sbfDescriptor, jdxa + 1); if (jdxe < 0) { jdxe = text.length; } tmpSf = andArr[idx].pre + text.substr(jdxa + 2, jdxe - jdxa - 2) + andArr[idx].post; tmpArray.push(tmpSf); if (jdxa == jdxl) test = true; jdxa = text.indexOf(andArr[idx].sbf, jdxa + 1); } tmp += tmpArray.join(excelVars.strTrennzeichen); } else { jdxe = text.indexOf(excelVars.sbfDescriptor, jdxa + 1); if (jdxe < 0) { jdxe = text.length; } tmp += andArr[idx].pre + text.substr(jdxa + 2, jdxe - jdxa - 2) + andArr[idx].post; } } }
    if (spec == 'S') { idx = tmp.indexOf('@'); if (idx < 0) { tmp = '@' + tmp; } else { tmp = tmp.substr(idx); } idx = tmp.indexOf('{'); while (idx >= 1) { idxe = tmp.indexOf(' ', idx); if (idxe < 0) { tmp = tmp.substr(0, idx - 1); } else { tmp = tmp.substr(0, idx - 1) + tmp.substr(idxe + 1); } idx = tmp.indexOf('{'); } tmp = tmp.substr(1); }
    else if (spec !== 'K') { tmp = tmp.replace('@', ''); tmp = tmp.replace('{', ''); }
    tmp = tmp.replace(String.fromCharCode(27) + 'N', ''); tmp = tmp.replace(/\s+/g, ' ').replace(/\s+$/, '').replace(/^\s+/, ''); return tmp;
}

function __exceldatumHeute() { var jetzt = new Date(); var jahr = jetzt.getFullYear(); var monat = jetzt.getMonth() + 1; if (monat < 10) { monat = '0' + monat; } var strTag = jetzt.getDate(); if (strTag < 10) { strTag = '0' + strTag; } var stunde = jetzt.getHours(); if (stunde < 10) { stunde = '0' + stunde; } var minute = jetzt.getMinutes(); if (minute < 10) { minute = '0' + minute; } var sekunde = jetzt.getSeconds(); if (sekunde < 10) { sekunde = '0' + sekunde; } return jahr + '_' + monat + '_' + strTag + '_' + stunde + '_' + minute + '_' + sekunde; }


function __objToString(obj) {
    var str = "\n{";
    if (typeof obj === 'object') {
        for (var p in obj) {
            if (obj.hasOwnProperty(p)) {
                str += "    " + p + ':' + __objToString(obj[p]) + ",\n";
            }
        }
    }
    else {
        if (typeof obj == 'string') {
            return '"' + obj + '"';
        }
        else {
            return obj + '';
        }
    }
    return str.substring(0, str.length - 1) + "\n}";
}

// Print character codes and visual representation for a string s in ES3 style
function dumpStringInfo(s) {
    // Print character codes
    alert('dumpStringInfo called with s:\n"' + s + '"\n length: ' + s.length + ' \ntype: ' + typeof s);
    var codes = [];
    var i, ch, code;
    s = s || '';
    for (i = 0; i < s.length; i++) {
        alert('charAt(' + i + '): "' + s.charAt(i) + '" code: ' + s.charCodeAt(i));
        codes.push(s.charCodeAt(i));
    }
    alert('codes:' + codes.join(','));

    // Print visual representation
    var visual = [];
    for (i = 0; i < s.length; i++) {
        ch = s.charAt(i);
        code = s.charCodeAt(i);
        if (ch === ' ') visual.push('[space]');
        else if (ch === '\n') visual.push('[LF]');
        else if (ch === '\r') visual.push('[CR]');
        else if (ch === '\t') visual.push('[TAB]');
        else if (code === 160) visual.push('[NBSP]');
        else if (code === 65279) visual.push('[BOM]');
        else if (code === 8203) visual.push('[ZWSP]');
        else if (code === 0) visual.push('[NUL]');
        else visual.push(ch);
    }
    alert('visual: ' + visual.join(','));
}