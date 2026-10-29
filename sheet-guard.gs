/**
 * Safety net for deleted sheets.
 *
 * Mom's account cannot delete a sheet that holds protected cells (rows 1-3 are
 * protected on every sheet). This guard covers everything else: whenever a sheet
 * is removed from the file, it checks that all year sheets and Sheet200 still
 * exist and e-mails the file owner the names that are gone. A deleted sheet is
 * restored with File > Version history > See version history.
 *
 * installSheetGuard is run once (setupYear does it); removeSheetGuard undoes it.
 * Needs config.gs.
 */

const GUARD_HANDLER = 'onSheetStructureChange';
const GUARD_MAIL_SUBJECT = 'Анхааруулга: тайлангийн sheet устгагдсан байна';

function installSheetGuard() {
  const book = currentBook();
  if (triggersOf(book).some((trigger) => trigger.getHandlerFunction() === GUARD_HANDLER)) {
    return;
  }
  ScriptApp.newTrigger(GUARD_HANDLER)
    .forSpreadsheet(book)
    .onChange()
    .create();
}

function removeSheetGuard() {
  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === GUARD_HANDLER)
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));
}

function onSheetStructureChange(event) {
  if (!event || event.changeType !== 'REMOVE_GRID') {
    return;
  }
  const spreadsheet = targetBook(event);
  const names = spreadsheet.getSheets().map((sheet) => sheet.getName());
  const missing = findMissingSheets(names, expectedSheetNames());
  if (missing.length === 0) {
    return;
  }
  MailApp.sendEmail(
    Session.getEffectiveUser().getEmail(),
    GUARD_MAIL_SUBJECT,
    `"${spreadsheet.getName()}" файлаас дараах sheet алга болсон байна: ${missing.join(', ')}\n\n` +
      'Сэргээх: File > Version history > See version history.\n' +
      spreadsheet.getUrl()
  );
}

// ---------- pure helpers (no Spreadsheet calls) ----------

function expectedSheetNames() {
  const yearSheets = CHAINS.flatMap((chain) =>
    Array.from({ length: CHAIN_PERIOD_COUNT }, (_, position) => chainSheetName(chain, position))
  );
  return [...yearSheets, ...GUARD_EXTRA_SHEETS, ...[...STOCK_SHEETS, ...COUNT_SHEETS].map((sheet) => sheet.name)];
}

function findMissingSheets(existingNames, expectedNames) {
  const existing = new Set(existingNames.map(canonicalSheetName));
  return expectedNames.filter((name) => !existing.has(name));
}
