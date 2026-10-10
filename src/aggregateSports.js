// Google Apps Script — Aggregate Sports (v6.0 Smart Match)

function aggregateSports() {
    const CONFIG = {
        SOURCE_SPREADSHEET_ID: '', // ID таблиці-джерела
        SCHEDULE_SPREADSHEET_ID: '', // ID таблиці розкладу
        SCHEDULE_SHEET_NAME: 'Календар', // Назва робочого аркуша в таблиці розкладу

        DEST_SPREADSHEET_ID: '', // якщо порожній — використовує активну таблицю
        DEST_SHEET_NAME: 'Aggregated',

        // Словник регіональних маркерів
        GEO_MARKERS: [
            { keyword: '(3 країни)', suffix: '(3 країни)' },
            { keyword: '(8 гео)', suffix: '(8 гео)' },
            { keyword: ' кз', suffix: '(КЗ)' },
            { keyword: ' kz', suffix: '(KZ)' }
        ]
    };

    const sourceSS = SpreadsheetApp.openById(CONFIG.SOURCE_SPREADSHEET_ID);
    const destSS = CONFIG.DEST_SPREADSHEET_ID
    ? SpreadsheetApp.openById(CONFIG.DEST_SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const limitDate = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000); // +7 днів

    let matchMap = new Map();
    const allSheets = sourceSS.getSheets();

    // ==========================================
    // ФАЗА 1: ЗБІР БАЗОВИХ ДАНИХ (MGG: Трансляції)
    // ==========================================
    allSheets.forEach(sheet => {
        const sheetName = sheet.getName();

        if (sheetName === CONFIG.DEST_SHEET_NAME) return;
        if (sheet.isSheetHidden()) return;
        if (sheet.getLastRow() < 2) return;

        let currentGeoSuffix = '';
        const sheetNameLower = sheetName.toLowerCase();

        for (const marker of CONFIG.GEO_MARKERS) {
            if (sheetNameLower.includes(marker.keyword.toLowerCase())) {
                currentGeoSuffix = ` ${marker.suffix}`;
                break;
            }
        }

        const dataRange = sheet.getDataRange();
        const data = dataRange.getDisplayValues();
        const notesData = dataRange.getNotes();

        const headers = data[0].map(h => h.toString().replace(/\s+/g, '').toLowerCase());

        const findHeader = (...possibleNames) => {
            for (const name of possibleNames) {
                const cleanName = name.replace(/\s+/g, '').toLowerCase();
                const index = headers.indexOf(cleanName);
                if (index !== -1) return index;
            }
            return -1;
        };

        const idx = {
            name:  findHeader('название матча', 'назва матчу', 'название события', 'назва події', 'match name', 'event', 'матч'),
                      start: findHeader('начало матча', 'початок матчу', 'начало, utc+3 (киев)', 'начало, utc+2 (киев)', 'time utc', 'start time', 'начало', 'початок', 'time', 'час'),
                      date:  findHeader('date', 'дата'),
                      cdn:   findHeader('cdn id', 'cdn'),
                      mgg:   findHeader('mgg id', 'mgg', 'megogo id')
        };

        if (idx.name === -1 || idx.start === -1 || idx.date === -1) {
            Logger.log(`⚠️ Пропуск: на аркуші "${sheetName}" немає базових заголовків.`);
            return;
        }

        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const rowNotes = notesData[i];
            const dateStr = row[idx.date];
            let matchName = row[idx.name] ? row[idx.name].toString().trim() : '';

            if (!dateStr || !matchName) continue;

            const parts = dateStr.split('.');
            if (parts.length !== 3) continue;

            const matchDate = new Date(parts[2], parts[1] - 1, parts[0]);
            if (isNaN(matchDate.getTime())) continue;

            matchDate.setHours(0, 0, 0, 0);
            if (matchDate < today || matchDate > limitDate) continue;

            const timeObj = parseMatchTime(row[idx.start], matchDate);
            const matchTimeStr = formatTime(timeObj);

            const cdnId = idx.cdn !== -1 ? row[idx.cdn].toString().trim() : '';
            const mggId = idx.mgg !== -1 ? row[idx.mgg].toString().trim() : '';

            let note = '';
            const noteName = idx.name !== -1 && rowNotes[idx.name] ? rowNotes[idx.name].trim() : '';
            const noteTime = idx.start !== -1 && rowNotes[idx.start] ? rowNotes[idx.start].trim() : '';

            if (noteName && noteTime && noteName !== noteTime) {
                note = `${noteTime} | ${noteName}`;
            } else {
                note = noteTime || noteName;
            }

            if (currentGeoSuffix) {
                matchName = matchName + currentGeoSuffix;
            }

            const matchKey = `${dateStr}||${matchTimeStr}||${matchName.toLowerCase()}`;

            if (matchMap.has(matchKey)) {
                let existingMatch = matchMap.get(matchKey);
                if (!existingMatch.cdn && cdnId) existingMatch.cdn = cdnId;
                if (!existingMatch.mggId && mggId) existingMatch.mggId = mggId;
                if (!existingMatch.note && note) existingMatch.note = note;
            } else {
                matchMap.set(matchKey, {
                    matchDate: matchDate,
                    cdn: cdnId,
                    name: matchName,
                    timeObj: timeObj,
                    mggId: mggId,
                    note: note
                });
            }
        }
    });

    let allRows = Array.from(matchMap.values());

    allRows.sort((a, b) => {
        const dateDiff = a.matchDate.getTime() - b.matchDate.getTime();
        if (dateDiff !== 0) return dateDiff;
        return a.timeObj.getTime() - b.timeObj.getTime();
    });

    // ==========================================
    // ФАЗА 2: ГЕНЕРАЦІЯ AUTO ID
    // ==========================================
    allRows.forEach(r => {
        let autoId = '';
        if (r.cdn) {
            let cdnStr = String(r.cdn).trim();
            autoId = cdnStr.length >= 6 ? '0' + cdnStr.slice(-6) : '0' + cdnStr;
        }
        r.autoId = autoId;
        r.channel = '';
        r.engine = '';
        r.commentator = '';
    });

    // ==========================================
    // ФАЗА 3: МЕТЧИНГ З РОЗКЛАДОМ СПОРТУ
    // ==========================================
    if (CONFIG.SCHEDULE_SPREADSHEET_ID && CONFIG.SCHEDULE_SHEET_NAME) {
        try {
            const scheduleSS = SpreadsheetApp.openById(CONFIG.SCHEDULE_SPREADSHEET_ID);
            const scheduleSheet = scheduleSS.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);

            if (scheduleSheet) {
                const schedData = scheduleSheet.getDataRange().getDisplayValues();

                if (schedData.length > 0) {
                    const schedHeaders = schedData[0].map(h => h.toString().toLowerCase().trim());

                    const idxSched = {
                        id: schedHeaders.indexOf('id'),
                        channel: schedHeaders.indexOf('канал'),
                        engine: schedHeaders.indexOf('engine'),
                        commentator: schedHeaders.indexOf('коментатор')
                    };

                    // Створюємо список потрібних нам ID, щоб не сканувати весь архів дарма
                    const targetIds = new Set(allRows.map(r => r.autoId).filter(id => id !== ''));

                    // Йдемо з кінця (найсвіжіші дані), щоб швидше знайти збіги
                    for (let i = schedData.length - 1; i >= 1; i--) {
                        if (targetIds.size === 0) break; // Знайшли всі потрібні матчі, зупиняємось

                        let row = schedData[i];
                        let schedId = idxSched.id !== -1 ? row[idxSched.id].trim() : '';

                        // Якщо ID пустий або його немає в нашому списку — пропускаємо миттєво
                        if (!schedId || !targetIds.has(schedId)) continue;

                        // Перевіряємо на прихованість ТІЛЬКИ ті рядки, де знайшли наш ID (економія API-запитів)
                        if (scheduleSheet.isRowHiddenByUser(i + 1) || scheduleSheet.isRowHiddenByFilter(i + 1)) {
                            continue;
                        }

                        // Знаходимо всі матчі з цим ID і заповнюємо дані
                        let matchedRows = allRows.filter(r => r.autoId === schedId);
                        matchedRows.forEach(targetMatch => {
                            if (idxSched.channel !== -1) targetMatch.channel = formatNomenclature(row[idxSched.channel]);
                            if (idxSched.engine !== -1) targetMatch.engine = formatNomenclature(row[idxSched.engine]);
                            if (idxSched.commentator !== -1) targetMatch.commentator = row[idxSched.commentator].trim();
                        });

                            targetIds.delete(schedId); // Видаляємо знайдений ID з пулу пошуку
                    }
                }
            } else {
                Logger.log(`⚠️ Аркуш "${CONFIG.SCHEDULE_SHEET_NAME}" не знайдено у таблиці розкладу.`);
            }
        } catch (e) {
            Logger.log(`❌ Помилка доступу до таблиці розкладу: ${e.message}`);
        }
    }

    // ==========================================
    // ФАЗА 4: ВИВІД ДАНИХ (Рендер + Порожні рядки)
    // ==========================================
    const outputRows = [];
    let previousDate = null;

    allRows.forEach(r => {
        // Форматуємо дату для зручного порівняння
        let currentDateStr = Utilities.formatDate(r.matchDate, "GMT+3", "dd.MM.yyyy");

        // Якщо це не перший матч і дата відрізняється від попередньої — вставляємо порожній рядок
        if (previousDate !== null && previousDate !== currentDateStr) {
            outputRows.push(['', '', '', '', '', '', '', '', '']);
        }

        // Додаємо сам матч
        outputRows.push([
            r.matchDate,           
            r.cdn,                 
            r.name,                
            formatTime(r.timeObj), 
                r.channel,             
                r.engine,              
                r.mggId,               
                r.commentator,         
                r.autoId               
        ]);

        previousDate = currentDateStr; // Запам'ятовуємо поточну дату для наступного циклу
    });

    const destSheet = destSS.getSheetByName(CONFIG.DEST_SHEET_NAME) || destSS.insertSheet(CONFIG.DEST_SHEET_NAME);
    destSheet.clear();
    destSheet.getRange('1:1').setFontWeight('bold');

    destSheet.appendRow([
        'Date', 'CDN ID', 'Название матча', 'Начало матча',
        'Канал', 'Engine', 'MGG ID', 'Коментатор', 'Auto ID'
    ]);

    if (outputRows.length > 0) {
        destSheet.getRange(2, 1, outputRows.length, 1).setNumberFormat('dd.MM.yyyy');
        destSheet.getRange(2, 9, outputRows.length, 1).setNumberFormat('@'); // Текстовий формат для Auto ID

        const range = destSheet.getRange(2, 1, outputRows.length, outputRows[0].length);
        range.setValues(outputRows);

        destSheet.autoResizeColumns(1, 9);
    }

    Logger.log(`✅ Імпорт завершено. Зведено матчів: ${allRows.length}, згенеровано рядків: ${outputRows.length}.`);
}

function parseMatchTime(timeStr, dateObj) {
    const d = new Date(dateObj);
    if (!timeStr) {
        d.setHours(0, 0, 0, 0);
        return d;
    }
    const match = timeStr.toString().trim().match(/^(\d{1,2}):(\d{2})/);
    if (!match) {
        d.setHours(0, 0, 0, 0);
        return d;
    }
    const hours = parseInt(match[1], 10);
    const minutes = parseInt(match[2], 10);
    d.setHours(hours, minutes, 0, 0);
    return d;
}

function formatTime(dateObj) {
    if (!(dateObj instanceof Date) || isNaN(dateObj)) return '';
    const hours = String(dateObj.getHours()).padStart(2, '0');
    const minutes = String(dateObj.getMinutes()).padStart(2, '0');
    return `${hours}:${minutes}`;
}
function formatNomenclature(str) {
    if (!str) return '';
    // /MF(\d+)/gi шукає "MF" (незалежно від регістру) плюс будь-які цифри після нього
    // 'F$1' замінює це на "F" і повертає знайдені цифри на місце
    return str.toString().trim()
    .replace(/MF(\d+)/gi, 'F$1')
    .replace(/LE(\d+)/gi, 'le$1');
}
