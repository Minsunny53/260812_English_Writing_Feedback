/**
 * English Writing Feedback — Google Sheets study-log webhook
 *
 * Setup:
 *   1. Open a Google Sheet → Extensions ▸ Apps Script.
 *   2. Replace everything with this file, Save.
 *   3. Deploy ▸ New deployment ▸ Web app.
 *      - Execute as: Me
 *      - Who has access: Anyone
 *   4. Copy the /exec URL and register it as the GitHub secret SHEETS_WEBHOOK.
 *
 * Each saved feedback is upserted by its ID, so re-saving an edited
 * review updates the same row instead of adding a duplicate.
 */
function doPost(e){
  var lock = LockService.getScriptLock();
  lock.tryLock(30000);
  try{
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Feedback') || ss.insertSheet('Feedback');
    var HEAD = ['ID','Date','Category','Original','Corrected',
      'Grammar','Clarity','Vocabulary','Coherence','Overall','Edits','Expressions','Summary'];
    if(sheet.getLastRow() === 0){ sheet.appendRow(HEAD); }

    var body = JSON.parse(e.postData.contents);
    var en = body.entry || {};
    var s = en.scores || {};
    var g=+s.grammar||0, c=+s.clarity||0, v=+s.vocabulary||0, co=+s.coherence||0;
    var overall = Math.round((g+c+v+co)/4);
    var edits = (en.segments||[]).filter(function(x){return x && x.cat;}).length;
    var expr = (en.expressions||[]).map(function(x){
      var w = x.written||x.spoken||'';
      var sp = x.spoken && x.written ? ' | spoken: '+x.spoken : '';
      return (x.phrase? x.phrase+' -> ':'') + w + sp;
    }).join('\n');

    var row = [ en.id||'', new Date(en.date||Date.now()), en.category||'',
      en.original||'', en.corrected||'', g, c, v, co, overall, edits, expr, en.summary||'' ];

    // upsert by ID (column 1)
    var id = en.id||'';
    var updated = false;
    if(id && sheet.getLastRow() > 1){
      var ids = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();
      for(var i=0;i<ids.length;i++){
        if(String(ids[i][0]) === String(id)){
          sheet.getRange(i+2,1,1,row.length).setValues([row]);
          updated = true; break;
        }
      }
    }
    if(!updated) sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({ok:true,updated:updated}))
      .setMimeType(ContentService.MimeType.JSON);
  }catch(err){
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }finally{ lock.releaseLock(); }
}
