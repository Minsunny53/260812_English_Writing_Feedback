/**
 * English Writing Feedback — Apps Script backend
 * Does TWO jobs for the web app, both over the same /exec URL:
 *   1) Gemini proxy  (action:"gemini")  — keeps your API key server-side
 *   2) Study-sheet log (action:"save")  — appends/updates a row per feedback
 *
 * SETUP
 *   1. Open a Google Sheet → Extensions ▸ Apps Script, paste this file, Save.
 *   2. Store your Gemini key WITHOUT putting it in any front-end file:
 *        Project Settings (gear) ▸ Script properties ▸ Add script property
 *        Name:  GEMINI_KEY      Value:  AQ.your-key
 *      (Optional)  GEMINI_MODEL   e.g.  gemini-2.5-flash   — overrides the model.
 *   3. Deploy ▸ New deployment ▸ Web app
 *        Execute as: Me      Who has access: Anyone
 *   4. Copy the /exec URL → register it as the GitHub secret APPS_SCRIPT_URL.
 *
 * The key lives only here. The browser only ever sees the /exec URL.
 */

function doPost(e){
  var out = ContentService.createTextOutput().setMimeType(ContentService.MimeType.JSON);
  try{
    var body = JSON.parse(e.postData.contents);
    if(body.action === 'gemini') return out.setContent(JSON.stringify(handleGemini(body)));
    return out.setContent(JSON.stringify(handleSheet(body)));   // action: "save"
  }catch(err){
    return out.setContent(JSON.stringify({ ok:false, error:String(err) }));
  }
}

/* ---- 1) Gemini proxy ---- */
function handleGemini(body){
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('GEMINI_KEY');
  if(!key) return { ok:false, error:'GEMINI_KEY not set in Script Properties' };
  var model = body.model || props.getProperty('GEMINI_MODEL') || 'gemini-2.5-flash';
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent';
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': key },
    payload: JSON.stringify(body.payload || {}),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  var text = res.getContentText();
  if(code < 200 || code >= 300) return { ok:false, status:code, error:text.slice(0,600) };
  var data = null; try{ data = JSON.parse(text); }catch(err){}
  return { ok:true, data:data };
}

/* ---- 2) Study-sheet log (upsert by ID) ---- */
function handleSheet(body){
  var lock = LockService.getScriptLock();
  lock.tryLock(30000);
  try{
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Feedback') || ss.insertSheet('Feedback');
    var HEAD = ['ID','Date','Category','Original','Corrected',
      'Grammar','Clarity','Vocabulary','Coherence','Overall','Edits','Expressions','Summary'];
    if(sheet.getLastRow() === 0){ sheet.appendRow(HEAD); }

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

    var id = en.id||'', updated = false;
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
    return { ok:true, updated:updated };
  }finally{ lock.releaseLock(); }
}
