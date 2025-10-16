// ============================== Мониторинг.gs ===========================
// Диагностический вывод «сырых» событий amo в логи (ограниченный по времени).

function logEvents() {
  var cfg = getCfg_();
  try {
    var timeFrom = Math.floor((new Date().getTime() - 150*60*1000)/1000);
    var url = `https://${cfg.AMO_HOST}/api/v4/events?filter[created_at][from]=${timeFrom}&limit=250`;
    var all = [];
    while (url) {
      var resp = UrlFetchApp.fetch(url, { method:'get', headers:{'Authorization':'Bearer '+cfg.AMO_TOKEN}, muteHttpExceptions:true, followRedirects:true });
      var code = resp.getResponseCode(); if (code<200||code>=300) break;
      var json = JSON.parse(resp.getContentText()||'{}');
      all = all.concat(json?._embedded?.events || []);
      url = json?._links?.next?.href || '';
    }
    console.log('[MONITOR] events: ' + all.length);
    console.log(JSON.stringify(all.slice(0,50), null, 2)); // не заспамить логи
  } catch(e){ console.error('[MONITOR] ' + String(e)); }
}