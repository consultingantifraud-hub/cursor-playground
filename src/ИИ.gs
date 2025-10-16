// ============================== ИИ.gs ===================================
// Совместимый слой. НИЧЕГО не переопределяет из AMO.gs/Оценка.gs.
// Важно: здесь НЕТ функций getProtalkBase_/protalkAskUrl_ — чтобы не было коллизий с QA-патчами.

// ---- Совместимый доступ к конфигу из листа "БД" ------------------------
function getConfigFromSheet() {
  try {
    var cfg = getCfg_(); // из AMO.gs
    return {
      AMO_TOKEN     : cfg.AMO_TOKEN,
      AMO_SUBDOMAIN : cfg.AMO_HOST,
      SHEET_NAME    : cfg.SHEET_NAME
    };
  } catch (e) {
    console.error('[ИИ.gs] getConfigFromSheet error:', String(e && e.message || e));
    return null;
  }
}

// ---- Необязательные прокси-хелперы ------------------------------------
function II_getLead(leadId) {
  try { return getAmoLead_(getCfg_(), leadId); }
  catch (e) { console.error('[ИИ.gs] II_getLead:', String(e)); return null; }
}
function II_getNotesForLead(leadId) {
  try { return getAmoNotes_(getCfg_(), 'leads', leadId) || []; }
  catch (e) { console.error('[ИИ.gs] II_getNotesForLead:', String(e)); return []; }
}
function II_getTasksForLead(leadId) {
  try { return getAmoTasks_(getCfg_(), leadId); }
  catch (e) { console.error('[ИИ.gs] II_getTasksForLead:', String(e)); return 'Ошибка'; }
}

// ---- Диагностика -------------------------------------------------------
function II_selfcheck() {
  try {
    var cfg = getCfg_();
    console.log('[II] ok, sheet=', cfg.SHEET_NAME, 'host=', cfg.AMO_HOST);
  } catch (e) {
    console.error('[II] selfcheck failed:', String(e && e.message || e));
  }
}