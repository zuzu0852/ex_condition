/**
 * 瞬間ジャッジ!撮影条件 — スコア集計用 Google Apps Script
 *
 * 使い方
 * 1. Googleスプレッドシートを新規作成し、[拡張機能] → [Apps Script] を開く
 * 2. このファイルの中身を全部貼り付けて保存
 * 3. [デプロイ] → [新しいデプロイ] → 種類「ウェブアプリ」
 *    ・次のユーザーとして実行: 自分
 *    ・アクセスできるユーザー: 全員
 * 4. 表示された「ウェブアプリのURL」(https://script.google.com/macros/s/.../exec)をClaudeに伝える
 *
 * シート
 *  plays     : 1プレイ1行(スコア)
 *  responses : 1問1行(研究用の生データ)
 */

const PLAY_COLS = ['received_at','session_id','user_id','nickname','experience','role','stage',
  'ref_thickness_cm','ref_kv','ref_mas','total_sec','correct','n','penalty_sec','rule_version','played_at','ua_mobile'];
const RESP_COLS = ['session_id','user_id','experience','role','stage','q_index','ref_thickness_cm','ref_kv','ref_mas',
  'patient_thickness_cm','delta_cm','height_cm','weight_kg','bmi','answer','correct_answer','correct','rt_sec','rule_version'];

function sheet_(name, cols){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if(!sh){ sh = ss.insertSheet(name); sh.appendRow(cols); sh.setFrozenRows(1); }
  return sh;
}

function doPost(e){
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    const p = JSON.parse(e.postData.contents);
    if(p.type !== 'play' || !Array.isArray(p.responses) || p.responses.length > 50) return out_({ok:false});
    const nick = String(p.nickname||'').slice(0,20);
    const plays = sheet_('plays', PLAY_COLS);
    // 同じ session_id の二重送信は無視
    const ids = plays.getRange(1,2,Math.max(plays.getLastRow(),1),1).getValues().flat();
    if(ids.indexOf(p.session_id) >= 0) return out_({ok:true, dup:true});
    plays.appendRow(PLAY_COLS.map(c => c==='received_at' ? new Date() : c==='nickname' ? nick : (p[c] ?? '')));
    const resp = sheet_('responses', RESP_COLS);
    const rows = p.responses.map(r => RESP_COLS.map(c => (c in r) ? r[c] : (p[c] ?? '')));
    if(rows.length) resp.getRange(resp.getLastRow()+1, 1, rows.length, RESP_COLS.length).setValues(rows);
    return out_({ok:true});
  } finally { lock.releaseLock(); }
}

// ?stage=knee → その編の全体ランキング(各ユーザーのベストタイム上位10)
function doGet(e){
  const stage = (e.parameter.stage||'').toString();
  const sh = sheet_('plays', PLAY_COLS);
  const v = sh.getDataRange().getValues(); v.shift();
  const idx = c => PLAY_COLS.indexOf(c);
  const best = {}; let count = 0;
  v.forEach(r => {
    if(r[idx('stage')] !== stage) return;
    count++;
    const uid = r[idx('user_id')], t = Number(r[idx('total_sec')]);
    if(!best[uid] || t < best[uid].time){
      best[uid] = {nick:r[idx('nickname')]||'(名無し)', exp:r[idx('experience')], time:t, correct:r[idx('correct')]};
    }
  });
  const top = Object.values(best).sort((a,b)=>a.time-b.time).slice(0,10);
  return out_({stage, count, top});
}

function out_(o){
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
