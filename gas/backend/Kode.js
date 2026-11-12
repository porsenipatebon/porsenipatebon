/**
 * PORSENI 2026 - Backend pendaftaran
 * Pasang pada Apps Script yang terikat ke Spreadsheet database.
 * Setelah mengganti kode, deploy ulang sebagai Web App: Execute as Me, access Anyone.
 */
var PDF_TTD_KOOR_ID = '1fM24qlCUGcKMwjtnGdLr8eDWH6pmpL9b';
var PDF_TTD_SEKRETARIAT_ID = '1CCuOOo6KzkX5kQbZjmAp1KRZ55qCcVwC';
var PDF_STEMPEL_PANITIA_ID = '1ITYlMp7nb9pIUOiL53J469Q6uhfnzEzT';
var NOTIFY = 'fuadsyafi62@gmail.com, wahyu.tyo242@gmail.com, shochibul.789@gmail.com, porsenipatebon@gmail.com';
var SENDER_EMAIL = 'porsenipatebon@gmail.com';
var SENDER_NAME = 'Panitia PORSENI 2026';
var REPLY_TO = 'porsenipatebon@gmail.com';

/**
 * Pengiriman email terpusat.
 * Jika akun eksekusi adalah alamat utama PORSENI, Gmail mengirim dari akun itu.
 * Jika bukan, alamat PORSENI wajib sudah diverifikasi sebagai alias Gmail.
 */
function porseniEscapeHtml_(value) {
  return String(value == null ? '' : value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\"/g, '&quot;').replace(/'/g, '&#39;');
}

/** Membungkus seluruh email keluar dengan identitas visual resmi PORSENI. */
function buildPorseniEmailHtml_(subject, plainBody) {
  // Sebagian template lama menyimpan pemisah baris sebagai karakter \\n; normalkan agar tidak tampil mentah.
  var raw = String(plainBody == null ? '' : plainBody).replace(/\\n/g, '\n').replace(/\r/g, '');
  var lines = raw.split('\n');
  var content = lines.map(function(line) {
    var safeLine = porseniEscapeHtml_(line);
    if (!safeLine.trim()) return '<div style="height:9px;line-height:9px">&nbsp;</div>';
    if (/^\s*(STATUS VERIFIKASI:|TINDAK LANJUT:|RINCIAN BERKAS|CATATAN DAN PERBAIKAN|ALASAN PENOLAKAN|PERHATIAN PENTING|JADWAL BATAS AKHIR)/i.test(line)) {
      return '<div style="margin:18px 0 8px;padding:9px 12px;background:#f3f5f8;border-left:4px solid #b3262d;color:#17243a;font-weight:700;font-size:13px;letter-spacing:.2px">' + safeLine + '</div>';
    }
    if (/^\s*(\d+\.|[-•])\s*/.test(line)) return '<div style="margin:4px 0 4px 12px;line-height:1.65">' + safeLine + '</div>';
    if (/^(Assalamu|Wassalamu|Hormat kami|PANITIA PORSENI|Panitia PORSENI|PAC IPNU|“Unitas|Bersatu untuk)/i.test(line.trim())) {
      return '<div style="line-height:1.65;color:#243247">' + safeLine + '</div>';
    }
    return '<div style="line-height:1.7;color:#334155">' + safeLine + '</div>';
  }).join('');
  return '<!doctype html><html><body style="margin:0;padding:0;background:#f1f3f6;font-family:Arial,Helvetica,sans-serif;color:#243247">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f1f3f6;padding:28px 12px"><tr><td align="center">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:680px;background:#ffffff;border:1px solid #e2e6eb;border-radius:10px;overflow:hidden">' +
    '<tr><td style="padding:22px 28px;background:#14243b;border-bottom:5px solid #c99a32"><div style="font-size:11px;letter-spacing:2px;color:#e8c979;font-weight:700">PEKAN OLAHRAGA DAN SENI</div><div style="font-size:25px;line-height:1.2;color:#ffffff;font-weight:800;margin-top:5px">PORSENI 2026</div><div style="font-size:12px;color:#d5dce6;margin-top:6px">PAC IPNU-IPPNU Kecamatan Patebon</div></td></tr>' +
    '<tr><td style="padding:26px 28px 30px"><div style="font-size:12px;color:#7b8797;margin-bottom:16px">' + porseniEscapeHtml_(subject || 'Informasi resmi panitia') + '</div>' +
    '<div style="font-size:14px">' + content + '</div></td></tr>' +
    '<tr><td style="padding:16px 28px;background:#f7f8fa;border-top:1px solid #e7ebef"><div style="font-size:12px;font-weight:700;color:#26364d">PANITIA PORSENI 2026</div><div style="font-size:12px;color:#64748b;margin-top:4px">PAC IPNU-IPPNU Kecamatan Patebon · Unitas Ad Gloriam</div><div style="font-size:11px;color:#94a3b8;margin-top:9px">Email ini merupakan pemberitahuan resmi dari sistem PORSENI 2026.</div></td></tr>' +
    '</table></td></tr></table></body></html>';
}

function sendPorseniEmail_(mailOptions) {
  var options = Object.assign({}, mailOptions || {});
  if (!options.htmlBody && options.body) options.htmlBody = buildPorseniEmailHtml_(options.subject, options.body);
  var effectiveEmail = '';
  try {
    effectiveEmail = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  } catch (e) {}

  var senderEmail = SENDER_EMAIL.toLowerCase();
  var aliases = GmailApp.getAliases().map(function (email) {
    return String(email || '').trim().toLowerCase();
  });

  if (effectiveEmail !== senderEmail) {
    if (aliases.indexOf(senderEmail) === -1) {
      throw new Error(
        'Pengiriman sebagai ' + SENDER_EMAIL +
        ' belum tersedia. Jalankan Apps Script menggunakan akun ' +
        SENDER_EMAIL + ' atau tambahkan dan verifikasi alamat tersebut ' +
        'sebagai alias pada Gmail akun yang menjalankan Apps Script.'
      );
    }
    options.from = SENDER_EMAIL;
  } else {
    // Akun utama sudah merupakan alamat pengirim; jangan paksa sebagai alias.
    delete options.from;
  }

  options.name = SENDER_NAME;
  options.replyTo = REPLY_TO;

  var to = options.to;
  var subject = options.subject;
  var body = options.body;
  delete options.to;
  delete options.subject;
  delete options.body;

  GmailApp.sendEmail(to, subject, body, options);
}
var MAX_FILES = 40;
var MAX_FILE_BYTES = 10 * 1024 * 1024;
var MAX_TOTAL_BYTES = 25 * 1024 * 1024;
var DRIVE_PATH = ['PORSENI PATEBON', 'PORSENI 2026'];
var KAT = {
 'Duta Pelajar':['Putra','Putri'], 'Khitobah':['Putra','Putri'],
 'LCC':['Kelompok Campuran'], 'Karya ilmiah':['Kelompok Campuran'], 'Baca kitab':['Putra','Putri'],
 'Badminton':['Ganda Campuran'], 'Voli':['Kelompok Putra'], 'Sprint':['Putra','Putri'],
 'Tenis meja':['Putra','Putri'], 'PES PS3':['Tunggal Bebas'],
 'Parade Gelar Budaya':['Kelompok Campuran'], 'Rebana':['Kelompok Bebas'], 'Paduan suara':['Kelompok Campuran'],
 'Kaligrafi':['Putra','Putri'], 'Baca puisi':['Putra','Putri'],
 'MTQ':['Putra','Putri']
};
var CFG = {
 kontingen:{sheet:'Kontingen',folder:'02. BERKAS KONTINGEN',head:['Waktu','No. Registrasi','Nama Kontingen','Asal','Alamat','Ketua','WA Ketua','Email','Cabang Lomba','Jumlah Peserta','Metode Bayar','Tgl Bayar','Berkas (link Drive)','Folder Drive','PDF Formulir (link Drive)','Arsip PDF Panitia (link Drive)','Ketua PR IPNU','WA Ketua IPNU','Ketua PR IPPNU','WA Ketua IPPNU','Tempat Lahir Ketua','Tanggal Lahir Ketua','Instagram Ketua Kontingen','Instagram Ketua PR/PK IPNU','Instagram Ketua PR/PK IPPNU','Instagram PR/PK IPNU-IPPNU'],keys:['waktu','reg','nama','asal','alamat','ketua','wa_ketua','email','cabang','jml','bayar','tgl_bayar','berkas','folder','pdf','archivePdf','ketua_pr_ipnu','wa_ketua_ipnu','ketua_pr_ippnu','wa_ketua_ippnu','tempat_lahir_ketua','tanggal_lahir_ketua','ig_ketua','ig_ketua_ipnu','ig_ketua_ippnu','ig_organisasi']},
 cabang:{sheet:'Cabang',folder:'03. BERKAS CABANG LOMBA',head:['Waktu','No. Registrasi','Nama Kontingen','Cabang Lomba','Kategori','Jumlah Peserta','Daftar Peserta','Tempat Lahir Peserta','Tanggal Lahir Peserta','Alamat Rumah Peserta','WA Peserta','Catatan','Berkas (link Drive)','Folder Drive','PDF Formulir (link Drive)','Arsip PDF Panitia (link Drive)'],keys:['waktu','reg','nama','cabang','kategori','jml','peserta','tempat_lahir_peserta','tanggal_lahir_peserta','alamat_peserta','wa_peserta','catatan','berkas','folder','pdf','archivePdf']}
};
var KONTINGEN = [
 'PK IPNU IPPNU MTS NU 07 PATEBON',
 'PK IPNU IPPNU MTS NU 24 DARUL ULUM PATEBON',
 'PK IPNU IPPNU SMK NU 04 PATEBON',
 'PKPT IPNU IPPNU STIK',
 'PR IPNU IPPNU BANGUNREJO',
 'PR IPNU IPPNU BANGUNSARI',
 'PR IPNU IPPNU BULUGEDE',
 'PR IPNU IPPNU DONOSARI',
 'PR IPNU IPPNU JAMBEARUM I',
 'PR IPNU IPPNU JAMBEARUM II',
 'PR IPNU IPPNU KEBONHARJO I',
 'PR IPNU IPPNU KEBONHARJO II',
 'PR IPNU IPPNU KUMPULREJO',
 'PR IPNU IPPNU LANJI',
 'PR IPNU IPPNU MAGERSARI',
 'PR IPNU IPPNU MARGOSARI',
 'PR IPNU IPPNU PIDODO KULON',
 'PR IPNU IPPNU PIDODO WETAN',
 'PR IPNU IPPNU PURWOSARI',
 'PR IPNU IPPNU PURWOKERTO',
 'PR IPNU IPPNU SUKOLILAN',
 'PR IPNU IPPNU TAMBAKREJO',
 'PR IPNU IPPNU WONOSARI'
];
var PHONE_KEYS=['wa_ketua','wa_pend','wa','wa_ketua_ipnu','wa_ketua_ippnu'];
var PIN_SHEET = 'PIN Kontingen';
var PIN_LENGTH = 6;
var PIN_MAX_ATTEMPTS = 10;
var PIN_WINDOW_MS = 10 * 60 * 1000;
var PIN_LOCK_MS = 10 * 60 * 1000;
var CONTINGENT_SESSION_SECONDS = 3600;
var SECURITY_LOG_SHEET = 'Log Aktivitas Keamanan';



function ensureSecurityLogSheet_(){
 var ss=SpreadsheetApp.getActiveSpreadsheet();if(!ss)throw new Error('Spreadsheet aktif tidak ditemukan.');
 var sh=ss.getSheetByName(SECURITY_LOG_SHEET)||ss.insertSheet(SECURITY_LOG_SHEET);
 var headers=['Waktu','Aktor','Aktivitas','Status','Detail'];
 if(sh.getLastRow()===0){sh.appendRow(headers);sh.setFrozenRows(1);sh.getRange(1,1,1,headers.length).setFontWeight('bold');}
 try{var protections=sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);if(!protections.length){var p=sh.protect();p.setDescription('Log aktivitas keamanan PORSENI — ditulis oleh Apps Script; baca melalui Portal Admin.');p.setWarningOnly(false);var editors=p.getEditors();if(editors&&editors.length)p.removeEditors(editors);if(p.canDomainEdit())p.setDomainEdit(false);}}catch(protectErr){console.warn('Proteksi log keamanan perlu diperiksa: '+protectErr);}
 return sh;
}
function logSecurityEvent_(actor,activity,status,detail){
 try{ensureSecurityLogSheet_().appendRow([new Date(),String(actor||'Sistem'),String(activity||''),String(status||''),String(detail||'').slice(0,1500)]);}catch(e){console.error('Log keamanan gagal: '+e);}
}
function invalidatePublicCache_(kind){
 try{
  var cache=CacheService.getScriptCache();
  if(!kind||kind==='recap'){cache.remove('public_recap_v1');cache.remove('registered_contingents_v1');cache.remove('available_contingents_v1');}
  if(!kind||kind==='news')cache.remove('public_news_v1');
 }catch(e){console.warn('Invalidasi cache publik gagal: '+e);}
}

/* ===== RESET TOTAL PENDAFTARAN — DUA TAHAP, ADMIN ONLY =====
 * Hanya mengosongkan data pendaftaran aktif dan isi dua pohon folder berkas pendaftaran.
 * Sheet, header, akun admin, berita, hasil kejuaraan, finalisasi, dan seluruh log dipertahankan.
 */
var RESET_REGISTRATION_CONFIRMATION_ = 'RESET SELURUH PENDAFTARAN PORSENI 2026';

function countRegistrationDriveFiles_(folder){
 var count=0,bytes=0,files=folder.getFiles();
 while(files.hasNext()){var f=files.next();if(!f.isTrashed()){count++;try{bytes+=f.getSize();}catch(e){}}}
 var folders=folder.getFolders();
 while(folders.hasNext()){var child=folders.next();var sub=countRegistrationDriveFiles_(child);count+=sub.count;bytes+=sub.bytes;}
 return {count:count,bytes:bytes};
}
function registrationResetSnapshot_(){
 var ss=SpreadsheetApp.getActiveSpreadsheet();
 if(!ss)throw new Error('Spreadsheet database tidak ditemukan.');
 var kont=ss.getSheetByName(CFG.kontingen.sheet),cabang=ss.getSheetByName(CFG.cabang.sheet),pin=ss.getSheetByName(PIN_SHEET);
 if(!kont||!cabang||!pin)throw new Error('Sheet Kontingen, Cabang, atau PIN Kontingen tidak ditemukan. Reset dibatalkan.');
 var root=getFolderPath(DRIVE_PATH);
 var kontFolder=getFolder(root,CFG.kontingen.folder),cabangFolder=getFolder(root,CFG.cabang.folder);
 var kontDrive=countRegistrationDriveFiles_(kontFolder),cabangDrive=countRegistrationDriveFiles_(cabangFolder);
 var props=PropertiesService.getScriptProperties().getProperties(),jobCount=0;
 Object.keys(props).forEach(function(k){if(k.indexOf(PDF_MERGE_JOB_PREFIX_)===0)jobCount++;});
 var legacy=0;try{legacy=JSON.parse(props[PDF_MERGE_QUEUE_PROPERTY_]||'[]').length;}catch(e){}
 return {
  sheets:{kontingen:Math.max(0,kont.getLastRow()-1),cabang:Math.max(0,cabang.getLastRow()-1),pin:Math.max(0,pin.getLastRow()-1)},
  drive:{kontingenFiles:kontDrive.count,cabangFiles:cabangDrive.count,totalFiles:kontDrive.count+cabangDrive.count,totalBytes:kontDrive.bytes+cabangDrive.bytes},
  pdfQueue:{jobs:jobCount,legacyJobs:legacy},
  folders:{kontingen:CFG.kontingen.folder,cabang:CFG.cabang.folder}
 };
}
function requireResetAdminTyo_(token){
 var admin=requireAdmin_(token);
 if(String(admin.user||'').trim()!=='Tyo')throw new Error('Akses ditolak. Fitur Reset Total Data hanya dapat digunakan oleh Admin Tyo.');
 return admin;
}
function adminPreviewResetRegistrations_(token){
 var admin=requireResetAdminTyo_(token),snapshot=registrationResetSnapshot_();
 logSecurityEvent_(admin.user,'PREVIEW RESET PENDAFTARAN','BERHASIL',JSON.stringify(snapshot));
 return {ok:true,preview:snapshot,confirmation:RESET_REGISTRATION_CONFIRMATION_,message:'Preview saja. Belum ada data atau berkas yang dihapus.'};
}
function trashFilesInRegistrationTree_(folder){
 var count=0,files=folder.getFiles();
 while(files.hasNext()){var f=files.next();if(!f.isTrashed()){f.setTrashed(true);count++;}}
 var folders=folder.getFolders();
 while(folders.hasNext())count+=trashFilesInRegistrationTree_(folders.next());
 return count;
}
function adminResetRegistrations_(d){
 var admin=requireResetAdminTyo_(d.token);
 if(String(d.confirmText||'')!==RESET_REGISTRATION_CONFIRMATION_)throw new Error('Konfirmasi tidak cocok. Reset tidak dijalankan.');
 var lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  var before=registrationResetSnapshot_(),ss=SpreadsheetApp.getActiveSpreadsheet();
  var kont=ss.getSheetByName(CFG.kontingen.sheet),cabang=ss.getSheetByName(CFG.cabang.sheet),pin=ss.getSheetByName(PIN_SHEET);
  // Kosongkan hanya baris data; seluruh header, format, dan struktur sheet tetap ada.
  [[kont,'kontingen'],[cabang,'cabang'],[pin,'pin']].forEach(function(pair){
   var sh=pair[0],last=sh.getLastRow();if(last>1)sh.deleteRows(2,last-1);
  });
  // Target dibatasi ke dua folder pendaftaran yang telah ditentukan di CFG.
  var root=getFolderPath(DRIVE_PATH);
  var trashedKont=trashFilesInRegistrationTree_(getFolder(root,CFG.kontingen.folder));
  var trashedCabang=trashFilesInRegistrationTree_(getFolder(root,CFG.cabang.folder));
  // Hapus hanya pekerjaan antrean PDF dan status uji PDF; trigger lain tidak disentuh.
  var props=PropertiesService.getScriptProperties(),all=props.getProperties();
  Object.keys(all).forEach(function(k){
   if(k.indexOf(PDF_MERGE_JOB_PREFIX_)===0 || k===PDF_MERGE_QUEUE_PROPERTY_ ||
      k===PDF_MERGE_RUNNING_PROPERTY_ || k.indexOf('PORSENI_TEST_PDF_STATUS_')===0)props.deleteProperty(k);
  });
  clearPdfMergeTriggers_();
  invalidatePublicCache_('recap');
  var summary={before:before,deletedRows:{kontingen:before.sheets.kontingen,cabang:before.sheets.cabang,pin:before.sheets.pin},
   trashedDriveFiles:{kontingen:trashedKont,cabang:trashedCabang,total:trashedKont+trashedCabang},
   clearedPdfJobs:before.pdfQueue.jobs+before.pdfQueue.legacyJobs};
  adminAudit_(admin,'RESET TOTAL PENDAFTARAN','sistem','', 'RESET PORSENI 2026',
   'Reset pendaftaran aktif selesai. '+JSON.stringify(summary));
  logSecurityEvent_(admin.user,'RESET TOTAL PENDAFTARAN','BERHASIL',JSON.stringify(summary));
  return {ok:true,message:'Reset data pendaftaran aktif selesai. Berkas Drive dipindahkan ke Sampah (bukan dihapus permanen).',summary:summary};
 }catch(err){
  logSecurityEvent_(admin.user,'RESET TOTAL PENDAFTARAN','GAGAL',String(err&&err.message||err));
  throw err;
 }finally{lock.releaseLock();}
}

/* ================= PORSENI CONTROL CENTER — GITHUB / APPS SCRIPT =================
 * Seluruh endpoint dibatasi server-side untuk Admin Tyo.
 * Token GitHub disimpan di Script Properties dan tidak pernah dikirim kembali ke browser.
 * Hanya file teks yang diizinkan; credential, file environment, dan konfigurasi clasp dilarang.
 */
var PORSENI_DEVOPS_PREFIX_ = 'PORSENI_DEVOPS_';
var PORSENI_DEVOPS_WORKFLOW_FILE_ = 'deploy-apps-script.yml';
var PORSENI_DEVOPS_MAX_FILE_BYTES_ = 700 * 1024;

function requireDevOpsTyo_(token){ return requireResetAdminTyo_(token); }
function devOpsProps_(){ return PropertiesService.getScriptProperties(); }
function devOpsConfig_(){
 var p=devOpsProps_();
 return {
  owner:p.getProperty(PORSENI_DEVOPS_PREFIX_+'GITHUB_OWNER')||'',
  repo:p.getProperty(PORSENI_DEVOPS_PREFIX_+'GITHUB_REPO')||'',
  branch:p.getProperty(PORSENI_DEVOPS_PREFIX_+'GITHUB_BRANCH')||'main',
  workflow:PORSENI_DEVOPS_WORKFLOW_FILE_,
  tokenConfigured:!!p.getProperty(PORSENI_DEVOPS_PREFIX_+'GITHUB_TOKEN')
 };
}
function devOpsRequireConfig_(){
 var c=devOpsConfig_(),token=devOpsProps_().getProperty(PORSENI_DEVOPS_PREFIX_+'GITHUB_TOKEN');
 if(!c.owner||!c.repo||!c.branch||!token)throw new Error('GitHub belum dikonfigurasi. Isi owner, repository, branch, dan fine-grained token terlebih dahulu.');
 return {config:c,token:token};
}
function devOpsApi_(method,path,payload,token,allow404){
 var headers={Accept:'application/vnd.github+json',Authorization:'Bearer '+token,'X-GitHub-Api-Version':'2022-11-28'};
 var options={method:String(method||'GET').toLowerCase(),headers:headers,muteHttpExceptions:true};
 if(payload!==undefined&&payload!==null){options.contentType='application/json';options.payload=JSON.stringify(payload);}
 var response=UrlFetchApp.fetch('https://api.github.com'+path,options),status=response.getResponseCode(),raw=response.getContentText(),body=null;
 if(raw){try{body=JSON.parse(raw);}catch(e){body={message:raw.slice(0,500)};}}
 if(allow404&&status===404)return {status:status,data:body};
 if(status<200||status>=300){
  var message=body&&body.message?String(body.message):'Permintaan GitHub gagal.';
  if(status===401||status===403)message='GitHub menolak akses. Periksa token dan izin Contents/Actions pada repository.';
  if(status===404)message='Repository, branch, file, atau workflow tidak ditemukan.';
  if(status===409||status===422)message='GitHub menolak perubahan karena konflik atau validasi. Muat ulang file lalu coba kembali.';
  throw new Error('GitHub API ('+status+'): '+message);
 }
 return {status:status,data:body};
}
function devOpsRepoPath_(owner,repo,path){
 return '/repos/'+encodeURIComponent(owner)+'/'+encodeURIComponent(repo)+'/contents/'+String(path).split('/').map(encodeURIComponent).join('/');
}
function devOpsValidatePath_(raw){
 var path=String(raw||'').trim().replace(/\\/g,'/').replace(/^\/+|\/+$/g,'');
 if(!path||path.length>240||path.indexOf('..')>=0||path.indexOf('//')>=0)throw new Error('Path file tidak valid.');
 var lower=path.toLowerCase(),parts=path.split('/');
 if(parts.some(function(x){return !x||x==='.'||x==='..';}))throw new Error('Path file tidak valid.');
 if(/(^|\/)(\.env(?:\.[^/]*)?|\.clasprc\.json|\.clasp\.json|[^/]*(?:credential|secret|token|private[-_]?key)[^/]*)(\/|$)/i.test(path)||/\.(pem|key|p12|pfx|keystore)$/i.test(path))throw new Error('File kredensial atau rahasia tidak boleh dikelola dari portal.');
 if(lower.indexOf('.github/workflows/')===0 && path!=='.github/workflows/'+PORSENI_DEVOPS_WORKFLOW_FILE_)throw new Error('Portal hanya boleh mengelola workflow deployment PORSENI yang telah ditentukan.');
 if(path!=='CNAME'&&!/\.(gs|html|htm|css|js|json|txt|md|yml|yaml|xml|toml|svg)$/i.test(path))throw new Error('Versi ini hanya mendukung file teks kode/konfigurasi (GS, HTML, CSS, JS, JSON, TXT, MD, YAML, XML, TOML, SVG, CNAME).');
 return path;
}
function devOpsSaveConfig_(d){
 var admin=requireDevOpsTyo_(d.token),owner=String(d.owner||'').trim(),repo=String(d.repo||'').trim(),branch=String(d.branch||'main').trim(),providedToken=String(d.githubToken||'').trim();
 if(!/^[A-Za-z0-9-]{1,39}$/.test(owner))throw new Error('Nama owner GitHub tidak valid.');
 if(!/^[A-Za-z0-9_.-]{1,100}$/.test(repo)||repo==='.'||repo==='..')throw new Error('Nama repository tidak valid.');
 if(!/^[A-Za-z0-9._/-]{1,100}$/.test(branch)||branch.indexOf('..')>=0||branch.charAt(0)==='/'||branch.charAt(branch.length-1)==='/')throw new Error('Nama branch tidak valid.');
 var p=devOpsProps_(),secret=providedToken||p.getProperty(PORSENI_DEVOPS_PREFIX_+'GITHUB_TOKEN');
 if(!secret)throw new Error('Masukkan fine-grained GitHub token. Token disimpan di Script Properties dan tidak ditampilkan kembali.');
 var check=devOpsApi_('GET','/repos/'+encodeURIComponent(owner)+'/'+encodeURIComponent(repo),null,secret,false).data;
 if(!check||!check.full_name)throw new Error('GitHub tidak dapat memverifikasi repository tersebut.');
 p.setProperties({
  [PORSENI_DEVOPS_PREFIX_+'GITHUB_OWNER']:owner,
  [PORSENI_DEVOPS_PREFIX_+'GITHUB_REPO']:repo,
  [PORSENI_DEVOPS_PREFIX_+'GITHUB_BRANCH']:branch,
  [PORSENI_DEVOPS_PREFIX_+'GITHUB_TOKEN']:secret
 },false);
 adminAudit_(admin,'KONFIGURASI GITHUB CONTROL CENTER','devops',0,owner+'/'+repo,'Repository tersambung; branch '+branch+'. Token tidak dicatat ke log.');
 return {ok:true,message:'Repository GitHub berhasil diverifikasi dan konfigurasi tersimpan.',status:devOpsConfig_()};
}
function adminDevOpsStatus_(token){
 var admin=requireDevOpsTyo_(token),c=devOpsConfig_();
 return {ok:true,status:c,ready:!!(c.owner&&c.repo&&c.tokenConfigured),message:c.tokenConfigured?'Token tersimpan di server.':'GitHub belum dikonfigurasi.'};
}
function adminDevOpsSaveSetup_(d){return devOpsSaveConfig_(d);}
function adminDevOpsListFiles_(d){
 requireDevOpsTyo_(d.token);var c=devOpsRequireConfig_(),path='/repos/'+encodeURIComponent(c.config.owner)+'/'+encodeURIComponent(c.config.repo)+'/git/trees/'+encodeURIComponent(c.config.branch)+'?recursive=1';
 var tree=devOpsApi_('GET',path,null,c.token,false).data;
 var files=(tree.tree||[]).filter(function(item){return item.type==='blob'&&item.path&&/\.(gs|html|htm|css|js|json|txt|md|yml|yaml|xml|toml|svg)$/i.test(item.path)&&!(item.path.toLowerCase().indexOf('.github/workflows/')===0&&item.path!=='.github/workflows/'+PORSENI_DEVOPS_WORKFLOW_FILE_)&&!/(^|\/)(\.env(?:\.[^/]*)?|\.clasprc\.json|\.clasp\.json|[^/]*(?:credential|secret|token|private[-_]?key)[^/]*)(\/|$)/i.test(item.path)&&!/\.(pem|key|p12|pfx|keystore)$/i.test(item.path);})
  .slice(0,500).map(function(item){return {path:item.path,size:item.size||0,sha:item.sha||''};});
 return {ok:true,files:files,truncated:!!tree.truncated,branch:c.config.branch};
}
function adminDevOpsReadFile_(d){
 requireDevOpsTyo_(d.token);var c=devOpsRequireConfig_(),path=devOpsValidatePath_(d.path);
 var result=devOpsApi_('GET',devOpsRepoPath_(c.config.owner,c.config.repo,path)+'?ref='+encodeURIComponent(c.config.branch),null,c.token,false).data;
 if(result.type!=='file')throw new Error('Path yang dipilih bukan file.');
 if(Number(result.size||0)>PORSENI_DEVOPS_MAX_FILE_BYTES_)throw new Error('File terlalu besar untuk editor portal (maksimal 700 KB).');
 if(!result.content)throw new Error('Isi file tidak tersedia dari GitHub API.');
 var bytes=Utilities.base64Decode(String(result.content).replace(/\s/g,''));
 var content=Utilities.newBlob(bytes).getDataAsString('UTF-8');
 return {ok:true,path:path,content:content,sha:result.sha||'',size:result.size||bytes.length,htmlUrl:result.html_url||''};
}
function devOpsCommitTextFile_(admin,c,path,content,sha,message){
 path=devOpsValidatePath_(path);content=String(content==null?'':content);
 var bytes=Utilities.newBlob(content,'text/plain','source.txt').getBytes();
 if(bytes.length>PORSENI_DEVOPS_MAX_FILE_BYTES_)throw new Error('Ukuran file melewati batas 700 KB.');
 if(!content.trim())throw new Error('Isi file kosong. Perubahan kosong tidak disimpan.');
 var endpoint=devOpsRepoPath_(c.config.owner,c.config.repo,path),current=devOpsApi_('GET',endpoint+'?ref='+encodeURIComponent(c.config.branch),null,c.token,true);
 var currentSha=current.status===404?'':String(current.data.sha||'');
 if(current.status!==404 && (!sha||String(sha)!==currentSha))throw new Error('File di GitHub sudah berubah sejak dibuka. Muat ulang file sebelum menyimpan agar perubahan orang lain tidak tertimpa.');
 if(current.status===404 && sha)throw new Error('File tidak lagi ditemukan di GitHub. Muat ulang daftar file.');
 var payload={message:String(message||('Update '+path)).slice(0,120),content:Utilities.base64Encode(bytes),branch:c.config.branch};
 if(currentSha)payload.sha=currentSha;
 var saved=devOpsApi_('PUT',endpoint,payload,c.token,false).data;
 adminAudit_(admin,'COMMIT FILE GITHUB','devops',0,path,'Commit '+(saved.commit&&saved.commit.sha?String(saved.commit.sha).slice(0,12):'')+' pada '+c.config.owner+'/'+c.config.repo+'.');
 return {ok:true,path:path,sha:saved.content&&saved.content.sha||'',commitSha:saved.commit&&saved.commit.sha||'',commitUrl:saved.commit&&saved.commit.html_url||'',message:'File berhasil di-commit ke GitHub.'};
}
function adminDevOpsSaveFile_(d){var admin=requireDevOpsTyo_(d.token),c=devOpsRequireConfig_();return devOpsCommitTextFile_(admin,c,d.path,d.content,d.sha,String(d.message||''));}

/** Commit beberapa file dalam satu commit Git tree; tidak membuat commit parsial per file. */
function adminDevOpsPublishBundle_(d){
 var admin=requireDevOpsTyo_(d.token),c=devOpsRequireConfig_(),files=d.files||[];
 if(!Array.isArray(files)||!files.length||files.length>8)throw new Error('Pilih 1 sampai 8 file untuk satu batch.');
 var allowed={'index.html':true,'admin.html':true,'gas/backend/Kode.js':true,'gas/backend/appsscript.json':true,'CNAME':true};
 var seen={},prepared=[],total=0,hasGas=false;
 files.forEach(function(f){
  var path=devOpsValidatePath_(f&&f.path),content=String(f&&f.content==null?'':f.content);
  if(!allowed[path])throw new Error('Path batch tidak diizinkan: '+path);
  if(seen[path])throw new Error('File duplikat pada batch: '+path);seen[path]=true;
  if(!content.trim())throw new Error('File kosong tidak dapat dipublikasikan: '+path);
  var bytes=Utilities.newBlob(content,'text/plain',path.split('/').pop()).getBytes();
  if(bytes.length>PORSENI_DEVOPS_MAX_FILE_BYTES_)throw new Error(path+' melebihi batas 700 KB.');
  total+=bytes.length;if(total>1800*1024)throw new Error('Total batch melebihi 1,8 MB.');
  if(path==='gas/backend/Kode.js'||path==='gas/backend/appsscript.json')hasGas=true;
  if(path==='gas/backend/Kode.js'&&(!/function\s+doGet\s*\(/.test(content)||!/function\s+doPost\s*\(/.test(content)))throw new Error('Kode.js tidak lengkap: doGet() dan doPost(e) wajib tersedia.');
  if(path==='gas/backend/appsscript.json'){try{JSON.parse(content);}catch(e){throw new Error('appsscript.json bukan JSON valid.');}}
  prepared.push({path:path,bytes:bytes});
 });
 var head=devOpsApi_('GET','/repos/'+encodeURIComponent(c.config.owner)+'/'+encodeURIComponent(c.config.repo)+'/git/ref/heads/'+encodeURIComponent(c.config.branch),null,c.token,false).data;
 var parentSha=head.object&&head.object.sha;if(!parentSha)throw new Error('Commit branch saat ini tidak dapat dibaca.');
 var parent=devOpsApi_('GET','/repos/'+encodeURIComponent(c.config.owner)+'/'+encodeURIComponent(c.config.repo)+'/git/commits/'+encodeURIComponent(parentSha),null,c.token,false).data;
 var treeItems=prepared.map(function(f){
  var blob=devOpsApi_('POST','/repos/'+encodeURIComponent(c.config.owner)+'/'+encodeURIComponent(c.config.repo)+'/git/blobs',{content:Utilities.base64Encode(f.bytes),encoding:'base64'},c.token,false).data;
  return {path:f.path,mode:'100644',type:'blob',sha:blob.sha};
 });
 var tree=devOpsApi_('POST','/repos/'+encodeURIComponent(c.config.owner)+'/'+encodeURIComponent(c.config.repo)+'/git/trees',{base_tree:parent.tree.sha,tree:treeItems},c.token,false).data;
 var commit=devOpsApi_('POST','/repos/'+encodeURIComponent(c.config.owner)+'/'+encodeURIComponent(c.config.repo)+'/git/commits',{message:'PORSENI Control Center: publish '+prepared.map(function(x){return x.path;}).join(', '),tree:tree.sha,parents:[parentSha]},c.token,false).data;
 devOpsApi_('PATCH','/repos/'+encodeURIComponent(c.config.owner)+'/'+encodeURIComponent(c.config.repo)+'/git/refs/heads/'+encodeURIComponent(c.config.branch),{sha:commit.sha,force:false},c.token,false);
 var gasDispatched=false,workflowUrl='';
 if(hasGas&&d.deployGas===true){
  if(String(d.confirmation||'')!=='SAYA-PAHAM-APPS-SCRIPT'||String(d.productionConfirm||'')!=='DEPLOY-PRODUKSI-PORSENI-2026')throw new Error('Konfirmasi deployment produksi tidak valid. Commit tersimpan, tetapi deployment belum dipicu.');
  var wf='/repos/'+encodeURIComponent(c.config.owner)+'/'+encodeURIComponent(c.config.repo)+'/actions/workflows/'+encodeURIComponent(PORSENI_DEVOPS_WORKFLOW_FILE_)+'/dispatches';
  devOpsApi_('POST',wf,{ref:c.config.branch,inputs:{mode:'deploy-production',confirmation:'SAYA-PAHAM-APPS-SCRIPT'}},c.token,false);
  gasDispatched=true;workflowUrl='https://github.com/'+encodeURIComponent(c.config.owner)+'/'+encodeURIComponent(c.config.repo)+'/actions/workflows/'+encodeURIComponent(PORSENI_DEVOPS_WORKFLOW_FILE_);
 }
 adminAudit_(admin,'PUBLISH BATCH FILE GITHUB','devops',0,c.config.owner+'/'+c.config.repo,'Commit '+String(commit.sha).slice(0,12)+'; files: '+prepared.map(function(x){return x.path;}).join(', ')+'; GAS dispatch: '+gasDispatched);
 return {ok:true,commitSha:commit.sha,commitUrl:commit.html_url||'',files:prepared.map(function(x){return x.path;}),gasDispatched:gasDispatched,workflowUrl:workflowUrl,message:'Seluruh file berhasil diterbitkan dalam satu commit GitHub.'+(gasDispatched?' Deployment Apps Script otomatis telah dipicu.':'')};
}
function devOpsWorkflowText_(branch){
 branch=String(branch||'main').trim();
 if(!/^[A-Za-z0-9._/-]{1,100}$/.test(branch)||branch.indexOf('..')>=0||branch.charAt(0)==='/'||branch.charAt(branch.length-1)==='/')throw new Error('Branch workflow tidak valid.');
 var ref="refs/heads/"+branch;
 return [
  'name: PORSENI Apps Script — Manual Deploy',
  '',
  'on:',
  '  workflow_dispatch:',
  '    inputs:',
  '      mode:',
  '        description: "Pilih tindakan: push-only tidak memperbarui Web App aktif; deploy-production memperbarui deployment aktif."',
  '        required: true',
  '        type: choice',
  '        options:',
  '          - push-only',
  '          - deploy-production',
  '      confirmation:',
  '        description: "Ketik persis: SAYA-PAHAM-APPS-SCRIPT"',
  '        required: true',
  '        type: string',
  '',
  'permissions:',
  '  contents: read',
  '',
  'concurrency:',
  '  group: porseni-apps-script-production',
  '  cancel-in-progress: false',
  '',
  'jobs:',
  '  deploy:',
  '    name: Sync / deploy Apps Script',
  "    if: github.ref == '"+ref+"' && inputs.confirmation == 'SAYA-PAHAM-APPS-SCRIPT'",
  '    runs-on: ubuntu-24.04',
  '    timeout-minutes: 15',
  '    steps:',
  '      - name: Checkout repository',
  '        uses: actions/checkout@v7',
  '',
  '      - name: Setup Node.js',
  '        uses: actions/setup-node@v7',
  '        with:',
  '          node-version: "22"',
  '',
  '      - name: Validate source and required secrets',
  '        env:',
  '          CLASPRC_JSON: ${{ secrets.CLASPRC_JSON }}',
  '          CLASP_JSON: ${{ secrets.CLASP_JSON }}',
  '          GAS_DEPLOYMENT_ID: ${{ secrets.GAS_DEPLOYMENT_ID }}',
  '          DEPLOY_MODE: ${{ inputs.mode }}',
  '        run: |',
  '          set -euo pipefail',
  '          test -n "$CLASPRC_JSON" || { echo "Missing secret CLASPRC_JSON"; exit 1; }',
  '          test -n "$CLASP_JSON" || { echo "Missing secret CLASP_JSON"; exit 1; }',
  '          if [ "$DEPLOY_MODE" = "deploy-production" ]; then',
  '            test -n "$GAS_DEPLOYMENT_ID" || { echo "Missing secret GAS_DEPLOYMENT_ID"; exit 1; }',
  '          fi',
  '          test -f gas/backend/appsscript.json',
  '          test -f gas/backend/Kode.js',
  '          node -e "const fs=require(\'fs\'); const c=JSON.parse(process.env.CLASP_JSON); if(!c.scriptId || c.rootDir !== \'gas/backend\') process.exit(2); const m=JSON.parse(fs.readFileSync(\'gas/backend/appsscript.json\',\'utf8\')); if(!m.dependencies?.libraries?.some(x=>x.userSymbol===\'PDFApp\' && String(x.version)===\'7\')) process.exit(3);"',
  '          node --check gas/backend/Kode.js',
  '',
  '      - name: Install clasp',
  '        run: npm install --global @google/clasp@3.4.1',
  '',
  '      - name: Configure temporary clasp credentials',
  '        env:',
  '          CLASPRC_JSON: ${{ secrets.CLASPRC_JSON }}',
  '          CLASP_JSON: ${{ secrets.CLASP_JSON }}',
  '        run: |',
  '          set -euo pipefail',
  '          printf \'%s\' "$CLASPRC_JSON" > "$HOME/.clasprc.json"',
  '          chmod 600 "$HOME/.clasprc.json"',
  '          printf \'%s\' "$CLASP_JSON" > .clasp.json',
  '          chmod 600 .clasp.json',
  '',
  '      - name: Verify Apps Script target and file scope',
  '        run: |',
  '          set -euo pipefail',
  '          clasp --user tyo show-file-status',
  '',
  '      - name: Create pre-sync backup version in Apps Script',
  '        run: |',
  '          set -euo pipefail',
  '          clasp --user tyo create-version "PORSENI pre-sync backup ${GITHUB_SHA::7}"',
  '',
  '      - name: Push source to Apps Script project',
  '        run: |',
  '          set -euo pipefail',
  '          clasp --user tyo push --force',
  '',
  '      - name: Update existing production Web App deployment',
  "        if: inputs.mode == 'deploy-production'",
  '        env:',
  '          GAS_DEPLOYMENT_ID: ${{ secrets.GAS_DEPLOYMENT_ID }}',
  '        run: |',
  '          set -euo pipefail',
  '          clasp --user tyo update-deployment "$GAS_DEPLOYMENT_ID" --description "PORSENI production ${GITHUB_SHA::7}"',
  '',
  '      - name: Deployment summary',
  '        env:',
  '          DEPLOY_MODE: ${{ inputs.mode }}',
  '        run: |',
  '          if [ "$DEPLOY_MODE" = "push-only" ]; then',
  '            echo "Source synced to Apps Script. Existing Web App deployment was NOT changed."',
  '          else',
  '            echo "Source synced and existing Web App deployment was updated."',
  '          fi'
 ].join('\n')+'\n';
}
function adminDevOpsInstallWorkflow_(d){
 var admin=requireDevOpsTyo_(d.token),c=devOpsRequireConfig_(),path='.github/workflows/'+PORSENI_DEVOPS_WORKFLOW_FILE_;
 var old=devOpsApi_('GET',devOpsRepoPath_(c.config.owner,c.config.repo,path)+'?ref='+encodeURIComponent(c.config.branch),null,c.token,true);
 var sha=old.status===404?'':String(old.data.sha||'');
 var result=devOpsCommitTextFile_(admin,c,path,devOpsWorkflowText_(c.config.branch),sha,'Install PORSENI Apps Script deployment workflow');
 return {ok:true,message:'Workflow deployment berhasil dipasang ke repository. Tambahkan GitHub Actions Secrets sebelum menjalankan deploy.',commitUrl:result.commitUrl,workflow:path};
}
function adminDevOpsDeployGas_(d){
 var admin=requireDevOpsTyo_(d.token),c=devOpsRequireConfig_(),workflow=PORSENI_DEVOPS_WORKFLOW_FILE_;
 var mode=String(d.mode||'').trim(),confirmation=String(d.confirmation||'').trim(),productionConfirm=String(d.productionConfirm||'').trim();
 if(mode!=='push-only'&&mode!=='deploy-production')throw new Error('Pilih mode push-only atau deploy-production.');
 if(confirmation!=='SAYA-PAHAM-APPS-SCRIPT')throw new Error('Konfirmasi workflow tidak cocok. Ketik SAYA-PAHAM-APPS-SCRIPT.');
 if(mode==='deploy-production'&&productionConfirm!=='DEPLOY-PRODUKSI-PORSENI-2026')
  throw new Error('Konfirmasi produksi tidak cocok. Ketik DEPLOY-PRODUKSI-PORSENI-2026.');
 var endpoint='/repos/'+encodeURIComponent(c.config.owner)+'/'+encodeURIComponent(c.config.repo)+'/actions/workflows/'+encodeURIComponent(workflow)+'/dispatches';
 // GitHub workflow_dispatch mengembalikan HTTP 204 tanpa body; input wajib harus dikirim eksplisit.
 devOpsApi_('POST',endpoint,{ref:c.config.branch,inputs:{mode:mode,confirmation:confirmation}},c.token,false);
 var runUrl='https://github.com/'+encodeURIComponent(c.config.owner)+'/'+encodeURIComponent(c.config.repo)+'/actions/workflows/'+encodeURIComponent(workflow);
 adminAudit_(admin,'JALANKAN WORKFLOW GOOGLE APPS SCRIPT','devops',0,c.config.owner+'/'+c.config.repo,
  'Mode '+mode+' dipicu oleh '+admin.user+' pada branch '+c.config.branch+'.');
 return {ok:true,message:mode==='push-only'
   ?'Workflow push-only berhasil dipicu. Source project Apps Script akan disinkronkan, tetapi Web App aktif tidak diperbarui.'
   :'Workflow deployment produksi berhasil dipicu. Backend aktif akan diperbarui jika seluruh langkah workflow berhasil.',
   mode:mode,runUrl:runUrl,workflow:workflow,branch:c.config.branch};
}
function adminDevOpsWorkflowRuns_(d){
 requireDevOpsTyo_(d.token);var c=devOpsRequireConfig_(),workflow=PORSENI_DEVOPS_WORKFLOW_FILE_;
 var endpoint='/repos/'+encodeURIComponent(c.config.owner)+'/'+encodeURIComponent(c.config.repo)+'/actions/workflows/'+encodeURIComponent(workflow)+'/runs?branch='+encodeURIComponent(c.config.branch)+'&per_page=8';
 var data=devOpsApi_('GET',endpoint,null,c.token,false).data||{},rows=data.workflow_runs||[];
 return {ok:true,workflow:workflow,branch:c.config.branch,actionsUrl:'https://github.com/'+encodeURIComponent(c.config.owner)+'/'+encodeURIComponent(c.config.repo)+'/actions/workflows/'+encodeURIComponent(workflow),runs:rows.map(function(r){return {id:r.id,runNumber:r.run_number||'',name:r.name||r.display_title||'PORSENI Apps Script',status:r.status||'',conclusion:r.conclusion||'',event:r.event||'',branch:r.head_branch||'',sha:String(r.head_sha||'').slice(0,7),createdAt:r.created_at||'',updatedAt:r.updated_at||'',url:r.html_url||'',actor:r.actor&&r.actor.login||'',attempt:r.run_attempt||1};})};
}

function adminSecurityLogs_(token){
 var admin=requireAdmin_(token),sh=ensureSecurityLogSheet_();
 var rows=sh.getLastRow()<2?[]:sh.getRange(2,1,sh.getLastRow()-1,5).getValues();
 return {ok:true,logs:rows.slice(-300).reverse().map(function(r){return {waktu:r[0]?Utilities.formatDate(new Date(r[0]),'Asia/Jakarta','dd-MM-yyyy HH:mm:ss'):'',aktor:String(r[1]||''),aktivitas:String(r[2]||''),status:String(r[3]||''),detail:String(r[4]||'')};})};
}
function createContingentSession_(nama){
 var token=Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');
 CacheService.getScriptCache().put('cont_session_'+token,JSON.stringify({nama:nama,lastActivity:Date.now()}),CONTINGENT_SESSION_SECONDS);
 return token;
}
function requireContingentSession_(token){
 token=String(token||'');if(!/^[a-f0-9]{64}$/.test(token))throw new Error('Sesi Portal Kontingen tidak valid. Silakan login kembali.');
 var cache=CacheService.getScriptCache(),key='cont_session_'+token,raw=cache.get(key);
 if(!raw)throw new Error('Sesi Portal Kontingen telah berakhir. Silakan login kembali.');
 var session=JSON.parse(raw),now=Date.now();
 if(!session.lastActivity||now-session.lastActivity>CONTINGENT_SESSION_SECONDS*1000){cache.remove(key);throw new Error('Sesi berakhir karena tidak aktif selama 60 menit. Silakan login kembali.');}
 session.lastActivity=now;cache.put(key,JSON.stringify(session),CONTINGENT_SESSION_SECONDS);return session;
}
function contingentLogout_(token){token=String(token||'');if(/^[a-f0-9]{64}$/.test(token))CacheService.getScriptCache().remove('cont_session_'+token);return {ok:true};}

function doGet(){return out({ok:true,info:'PORSENI 2026 form endpoint',backendVersion:'PORSENI-2026-CONTROL-CENTER-V40-20261111',pinMaxAttempts:PIN_MAX_ATTEMPTS,pinLockMinutes:PIN_LOCK_MS/60000,contingentSessionSeconds:CONTINGENT_SESSION_SECONDS});}
function doPost(e){
 var lock,d=null,stage='baca_request';
 try{
  if(!e||!e.postData||!e.postData.contents)throw new Error('Data permintaan tidak ditemukan');
  stage='parse_json';d=JSON.parse(e.postData.contents);
  if(d.action==='admin_login')return out(adminLogin_(d.password,d.user));
  if(d.action==='admin_logout')return out(adminLogout_(d.token));
  if(d.action==='admin_dashboard')return out(adminDashboard_(d.token));
  if(d.action==='admin_record_files')return out(adminRecordFiles_(d));
  if(d.action==='public_registration_status')return out({ok:true,registration:registrationControls_()});
  if(d.action==='admin_get_registration_controls')return out(adminGetRegistrationControls_(d.token));
  if(d.action==='admin_set_registration_controls')return out(adminSetRegistrationControls_(d));
  if(d.action==='public_news')return out({ok:true,news:getPublicNews_()});
  if(d.action==='admin_news_list')return out(adminNewsList_(d.token));
  if(d.action==='admin_news_save')return out(adminNewsSave_(d));
  if(d.action==='admin_news_image_upload')return out(adminNewsImageUpload_(d));
  if(d.action==='admin_news_delete')return out(adminNewsDelete_(d));
  if(d.action==='admin_set_public_recap_mode')return out(adminSetPublicRecapMode_(d));
  if(d.action==='admin_update_status')return out(adminUpdateStatus_(d));
  if(d.action==='admin_update_record')return out(adminUpdateRecord_(d));
  if(d.action==='admin_replace_file')return out(adminReplaceFile_(d));
   if(d.action==='admin_rebuild_registration_pdf')return out(adminRebuildRegistrationPdf_(d));
  if(d.action==='admin_delete_file')return out(adminDeleteFile_(d));
  if(d.action==='admin_delete_record')return out(adminDeleteRecord_(d));
  if(d.action==='admin_allow_reregistration')return out(adminAllowReregistration_(d));
  if(d.action==='admin_get_results')return out(adminGetResults_(d.token));
  if(d.action==='admin_save_result')return out(adminSaveResult_(d));
  if(d.action==='admin_delete_result')return out(adminDeleteResult_(d));
  if(d.action==='admin_reset_championship_results')return out(adminResetChampionshipResults_(d));
  if(d.action==='admin_finalize_championship')return out(adminFinalizeChampionship_(d));
   if(d.action==='admin_test_championship_pdf')return out(adminTestChampionshipPdf_(d));
  if(d.action==='admin_resend_final_results')return out(adminResendFinalResults_(d));
  if(d.action==='admin_send_email')return out(adminSendEmail_(d));
  if(d.action==='admin_security_logs')return out(adminSecurityLogs_(d.token));
  if(d.action==='admin_preview_reset_registrations')return out(adminPreviewResetRegistrations_(d.token));
  if(d.action==='admin_reset_registrations')return out(adminResetRegistrations_(d));
  if(d.action==='admin_devops_status')return out(adminDevOpsStatus_(d.token));
  if(d.action==='admin_devops_save_setup')return out(adminDevOpsSaveSetup_(d));
  if(d.action==='admin_devops_list_files')return out(adminDevOpsListFiles_(d));
  if(d.action==='admin_devops_read_file')return out(adminDevOpsReadFile_(d));
  if(d.action==='admin_devops_save_file')return out(adminDevOpsSaveFile_(d));
  if(d.action==='admin_devops_publish_bundle')return out(adminDevOpsPublishBundle_(d));
  if(d.action==='admin_devops_install_workflow')return out(adminDevOpsInstallWorkflow_(d));
  if(d.action==='admin_devops_deploy_gas')return out(adminDevOpsDeployGas_(d));
  if(d.action==='admin_devops_workflow_runs')return out(adminDevOpsWorkflowRuns_(d));
  if(d.action==='contingent_logout')return out(contingentLogout_(d.sessionToken));
  if(d.action==='contingent_session_ping'){var activeSession=requireContingentSession_(d.sessionToken);return out({ok:true,nama:activeSession.nama,sessionSeconds:CONTINGENT_SESSION_SECONDS});}
  if(d.action==='contingent_portal_session'){var restored=requireContingentSession_(d.sessionToken),restoredData=getContingentPortalData_(restored.nama);restoredData.sessionSeconds=CONTINGENT_SESSION_SECONDS;return out({ok:true,data:restoredData});}
  if(d.action==='contingent_request_change')throw new Error('Fitur pengajuan perubahan telah dinonaktifkan. Hubungi Sekretariat/Koordinator Perlombaan dan kirimkan berkas terbaru secara manual melalui WhatsApp.');
  if(d.action==='contingent_update_data')throw new Error('Fitur pengajuan perubahan telah dinonaktifkan. Hubungi Sekretariat/Koordinator Perlombaan dan kirimkan berkas terbaru secara manual melalui WhatsApp.');
  if(d.action==='contingent_request_data_change')throw new Error('Fitur pengajuan perubahan telah dinonaktifkan. Hubungi Sekretariat/Koordinator Perlombaan dan kirimkan berkas terbaru secara manual melalui WhatsApp.');
  if(d.action==='admin_apply_change_request')throw new Error('Fitur pengajuan perubahan telah dinonaktifkan. Hubungi Sekretariat/Koordinator Perlombaan dan kirimkan berkas terbaru secara manual melalui WhatsApp.');
  if(d.action==='admin_reset_pin')return out(adminResetPin_(d));
  if(d.action==='admin_test_pdf')return out(adminTestPdf_(d));
      if(d.action==='admin_test_pdf_status')return out(adminTestPdfStatus_(d));
    if(d.action==='admin_test_pdf_file')return out(adminTestPdfFile_(d));
    if(d.action==='admin_retry_pdf_merge')return out(adminRetryPdfMerge_(d));
  if(d.action==='admin_update_change_request')throw new Error('Fitur pengajuan perubahan telah dinonaktifkan. Hubungi Sekretariat/Koordinator Perlombaan dan kirimkan berkas terbaru secara manual melalui WhatsApp.');
  if(d.action==='list_registered_contingents')return out({ok:true,names:getRegisteredContingentNames_()});
  if(d.action==='list_available_contingents')return out({ok:true,names:getAvailableContingentNames_()});
  if(d.action==='public_recap')return out({ok:true,recap:getPublicRecap_()});
  if(d.action==='lookup_cabang'){
   var lookupName=canonicalContingentName_(d.nama),approval=getContingentApproval_(lookupName),events=getRegisteredEvents_(lookupName);
   return out({ok:true,approval:approval,events:events});
  }
  if(d.action==='contingent_portal'){
   var portalName=canonicalContingentName_(d.nama);
   try{validateContingentPin_(portalName,d.pin);}catch(pinErr){logSecurityEvent_(portalName,'LOGIN KONTINGEN','GAGAL',String(pinErr.message||pinErr));throw pinErr;}
   var sessionToken=createContingentSession_(portalName),portalData=getContingentPortalData_(portalName);
   portalData.sessionToken=sessionToken;portalData.sessionSeconds=CONTINGENT_SESSION_SECONDS;
   logSecurityEvent_(portalName,'LOGIN KONTINGEN','BERHASIL','Sesi aktif 60 menit sejak aktivitas terakhir.');
   return out({ok:true,data:portalData});
  }
  if(d.action==='contingent_preview_file')return out(contingentPreviewFile_(d));
  stage='cek_jenis_formulir';var c=CFG[d.type];if(!c)throw new Error('Jenis formulir tidak dikenal');
  stage='validasi_data_dan_lampiran';validate(d);
  assertPdfMergeLibraryReady_(d.files||[]);
  if(d.type==='cabang')d.email=getContingentEmail_(canonicalContingentName_(d.nama))||'';
  if(d.type==='kontingen'){
   stage='cek_kontingen_dan_buat_pin';assertContingentNotRegistered_(canonicalContingentName_(d.nama));
   d.pin=createOrGetContingentPin_(canonicalContingentName_(d.nama));
  }
  d.reg=newRegFor_(d.type);d.waktu=new Date();
  stage='siapkan_folder_drive';var folder=resolveDestination(d,c);d.folder=folder.getUrl();
  stage='simpan_lampiran';var fileLinks=saveAttachments(d.files||[],folder,d);
  stage='buat_pdf_pendaftaran';var pdfFile=createRegistrationPdf(d,folder);assertPrivateIndividualFile_(pdfFile);d.pdf=pdfFile.getUrl();
  stage='gabungkan_lampiran_pdf';try{assertPdfMergeLibraryReady_(d.files||[]);d.pdfMergePending=queuePdfMerge_(pdfFile,d,folder);}catch(mergeQueueErr){d.pdfMergePending=false;logSecurityEvent_(d.nama,'PENGGABUNGAN PDF','GAGAL DIJADWALKAN',d.reg+' | '+String(mergeQueueErr&&mergeQueueErr.message||mergeQueueErr));}
  stage='arsip_pdf';var archiveFile=archivePrintablePdf_(pdfFile,d);d.archivePdf=archiveFile.getUrl();
  d.berkas=['PDF formulir: '+d.pdf].concat(fileLinks).join('\n');
  stage='tulis_google_sheets';lock=LockService.getScriptLock();lock.waitLock(30000);
  try{if(d.type==='kontingen')assertContingentNotRegistered_(canonicalContingentName_(d.nama));appendRecord(c,d);}finally{lock.releaseLock();}
  invalidatePublicCache_('recap');
  stage='catat_log_keamanan';logSecurityEvent_(d.nama,d.type==='kontingen'?'PENDAFTARAN KONTINGEN':'UNGGAH BERKAS CABANG','BERHASIL',(d.cabang||'')+' '+(d.kategori||'')+' · '+d.reg);
  stage='kirim_notifikasi';notify(d);
  return out({ok:true,reg:d.reg,pdf:d.pdf,pdfMergePending:!!d.pdfMergePending,pin:d.type==='kontingen'?d.pin:'',waktu:Utilities.formatDate(new Date(d.waktu),'Asia/Jakarta',"yyyy-MM-dd'T'HH:mm:ss'+07:00'")});
 }catch(err){
  var message=String(err&&err.message?err.message:err),stack=String(err&&err.stack?err.stack:'');
  if(d&&d.type&&d.nama){try{logSecurityEvent_(d.nama,d.type==='kontingen'?'PENDAFTARAN KONTINGEN':'UNGGAH BERKAS CABANG','GAGAL',stage+' | '+message);}catch(logErr){console.error('Gagal mencatat log kegagalan: '+logErr);}}
  console.error('PORSENI doPost gagal pada tahap '+stage+': '+stack);
  return out({ok:false,error:message,stage:stage,diagnosticStack:stack});
 }
}
/** ================= PORTAL PANITIA (ADMIN) =================
 * Set ADMIN_PASSWORD untuk Sochibul/Fuad dan ADMIN_PASSWORD_TYO untuk Tyo.
 * Password Tyo wajib unik dan minimal 12 karakter.
 * Token sesi disimpan di Script Cache.
 */
var ADMIN_SESSION_SECONDS = 21600; // 6 jam
var ADMIN_STATUS_VALUES = ['Menunggu Verifikasi','Terverifikasi','Perlu Revisi','Ditolak'];
var ADMIN_LOG_SHEET = 'Log Admin';
var ADMIN_USERS = ['Tyo','Sochibul','Fuad'];
var RESULT_SHEET = 'Hasil Kejuaraan';
var RESULT_HEADERS = ['Cabang Lomba','Kategori','Juara 1','Juara 2','Juara 3','Diperbarui Pada','Diperbarui Oleh'];
var CHAMPIONSHIP_FINALIZATION_PROPERTY = 'PORSENI_2026_CHAMPIONSHIP_FINALIZATION';
var CHAMPIONSHIP_PDF_FILE_PROPERTY = 'PORSENI_2026_CHAMPIONSHIP_PDF_FILE_ID';

function adminLogin_(password,userName){
 userName=String(userName||'').trim();
 if(ADMIN_USERS.indexOf(userName)<0)throw new Error('Pilih nama petugas yang terdaftar.');
 var props=PropertiesService.getScriptProperties();
 // Tyo wajib memakai kredensial tersendiri; tidak ada fallback ke password bersama.
 var propertyName=userName==='Tyo'?'ADMIN_PASSWORD_TYO':'ADMIN_PASSWORD';
 var configured=props.getProperty(propertyName);
 if(!configured)throw new Error(userName==='Tyo'
  ?'Password khusus Tyo belum dikonfigurasi. Tambahkan Script Property ADMIN_PASSWORD_TYO.'
  :'Password admin bersama belum dikonfigurasi. Tambahkan Script Property ADMIN_PASSWORD.');
 if(userName==='Tyo'){
  if(String(configured).length<12)throw new Error('Password khusus Tyo harus minimal 12 karakter. Perbarui Script Property ADMIN_PASSWORD_TYO.');
  var shared=props.getProperty('ADMIN_PASSWORD');
  if(shared&&constantTimeEquals_(String(configured),String(shared)))throw new Error('Password Tyo harus berbeda dari password admin bersama.');
 }
 var cache=CacheService.getScriptCache(),attemptKey='admin_login_failures_'+userName;
 var attempts=Number(cache.get(attemptKey)||0);
 if(attempts>=5){logSecurityEvent_(userName,'LOGIN ADMIN','DIBLOKIR','Batas percobaan login admin tercapai.');throw new Error('Terlalu banyak percobaan login. Coba lagi 15 menit.');}
 if(!constantTimeEquals_(String(password||''),String(configured))){
  cache.put(attemptKey,String(attempts+1),900);
  logSecurityEvent_(userName,'LOGIN ADMIN','GAGAL','Password admin tidak sesuai.');
  throw new Error('Password admin tidak sesuai.');
 }
 cache.remove(attemptKey);
 var token=Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');
 cache.put('admin_session_'+token,JSON.stringify({user:userName,created:Date.now(),credential:propertyName,credentialFingerprint:userName==='Tyo'?adminCredentialFingerprint_(configured):''}),ADMIN_SESSION_SECONDS);
 logSecurityEvent_(userName,'LOGIN ADMIN','BERHASIL',userName==='Tyo'?'Menggunakan password khusus Tyo.':'Menggunakan password admin bersama.');
 return {ok:true,token:token,expiresIn:ADMIN_SESSION_SECONDS,user:userName};
}
function constantTimeEquals_(a,b){
 if(a.length!==b.length)return false;
 var diff=0;for(var i=0;i<a.length;i++)diff|=a.charCodeAt(i)^b.charCodeAt(i);
 return diff===0;
}
function adminCredentialFingerprint_(value){
 var digest=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(value||''));
 return Utilities.base64EncodeWebSafe(digest);
}
function requireAdmin_(token){
 token=String(token||'');
 if(!/^[a-f0-9]{64}$/.test(token))throw new Error('Sesi admin tidak valid. Silakan login kembali.');
 var raw=CacheService.getScriptCache().get('admin_session_'+token);
 if(!raw)throw new Error('Sesi admin telah berakhir. Silakan login kembali.');
 var admin=JSON.parse(raw);if(ADMIN_USERS.indexOf(admin.user)<0)throw new Error('Sesi lama tidak memiliki identitas petugas. Silakan login kembali.');
 if(admin.user==='Tyo'){
  var dedicated=PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD_TYO');
  var currentFingerprint=dedicated?adminCredentialFingerprint_(dedicated):'';
  if(!dedicated||!admin.credentialFingerprint||!constantTimeEquals_(String(admin.credentialFingerprint),String(currentFingerprint)))
   throw new Error('Sesi Admin Tyo lama/kedaluwarsa. Silakan login ulang dengan password khusus Tyo.');
 }
 return admin;
}
function adminLogout_(token){
 token=String(token||'');if(token)CacheService.getScriptCache().remove('admin_session_'+token);
 return {ok:true};
}
function ensureAdminColumns_(sheet){
 var wanted=['Status Verifikasi','Catatan Panitia','Diverifikasi Pada','Diverifikasi Oleh','ID Data'];
 if(sheet.getLastColumn()===0)return;
 var headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
 wanted.forEach(function(h){if(headers.indexOf(h)<0){sheet.getRange(1,sheet.getLastColumn()+1).setValue(h).setFontWeight('bold');headers.push(h);}});
}
function adminReadSheet_(sheetName,type,includeFiles){
 includeFiles=includeFiles!==false;
 var ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss&&ss.getSheetByName(sheetName);
 if(!sheet||sheet.getLastRow()<2)return [];
 ensureAdminColumns_(sheet);
 var values=sheet.getDataRange().getValues(),headers=values[0],idx={};
 headers.forEach(function(h,i){idx[String(h)]=i;});
 var idIndex=idx['ID Data'];
 if(idIndex===undefined)throw new Error('Kolom ID Data belum tersedia setelah migrasi.');
 var idValues=values.slice(1).map(function(row){return [String(row[idIndex]||'').trim()||Utilities.getUuid()];});
 var idsChanged=false;
 for(var idRow=1;idRow<values.length;idRow++){
  if(!String(values[idRow][idIndex]||'').trim()){values[idRow][idIndex]=idValues[idRow-1][0];idsChanged=true;}
 }
 if(idsChanged)sheet.getRange(2,idIndex+1,idValues.length,1).setValues(idValues);
 function val(row,key){var i=idx[key];return i===undefined?'':row[i];}
 return values.slice(1).map(function(row,i){
  var files=includeFiles?parsePortalFiles_(val(row,'Berkas (link Drive)')):[];
  var record={type:type,rowNumber:i+2,recordId:String(val(row,'ID Data')||''),nomor:String(val(row,'No. Registrasi')||''),nama:String(val(row,'Nama Kontingen')||''),waktu:val(row,'Waktu')?Utilities.formatDate(new Date(val(row,'Waktu')),Session.getScriptTimeZone(),'dd-MM-yyyy HH:mm'):'',status:String(val(row,'Status Verifikasi')||'Menunggu Verifikasi'),catatanAdmin:String(val(row,'Catatan Panitia')||''),pdf:String(val(row,'PDF Formulir (link Drive)')||''),archivePdf:String(val(row,'Arsip PDF Panitia (link Drive)')||''),folder:String(val(row,'Folder Drive')||''),berkas:files};
  if(type==='kontingen'){
   record.asal=String(val(row,'Asal')||'');record.alamat=String(val(row,'Alamat')||'');record.ketua=String(val(row,'Ketua')||'');record.tempatLahirKetua=String(val(row,'Tempat Lahir Ketua')||'');record.tanggalLahirKetua=String(val(row,'Tanggal Lahir Ketua')||'');record.waKetua=String(val(row,'WA Ketua')||'');record.email=String(val(row,'Email')||'');record.cabang=String(val(row,'Cabang Lomba')||'');record.jumlahPeserta=String(val(row,'Jumlah Peserta')||'');record.metodeBayar=String(val(row,'Metode Bayar')||'');record.tanggalBayar=String(val(row,'Tgl Bayar')||'');record.ketuaIpnu=String(val(row,'Ketua PR IPNU')||'');record.ketuaIppnu=String(val(row,'Ketua PR IPPNU')||'');record.igKetua=String(val(row,'Instagram Ketua Kontingen')||'');record.igKetuaIpnu=String(val(row,'Instagram Ketua PR/PK IPNU')||'');record.igKetuaIppnu=String(val(row,'Instagram Ketua PR/PK IPPNU')||'');record.igOrganisasi=String(val(row,'Instagram PR/PK IPNU-IPPNU')||'');
  }else{
   record.cabang=String(val(row,'Cabang Lomba')||'');record.kategori=String(val(row,'Kategori')||'');record.jumlahPeserta=String(val(row,'Jumlah Peserta')||'');record.peserta=String(val(row,'Daftar Peserta')||'');record.tempatLahirPeserta=String(val(row,'Tempat Lahir Peserta')||'');record.tanggalLahir=String(val(row,'Tanggal Lahir Peserta')||'');record.alamatPeserta=String(val(row,'Alamat Rumah Peserta')||'');record.waPeserta=String(val(row,'WA Peserta')||'');record.catatan=String(val(row,'Catatan')||'');
  }
  return record;
 }).filter(function(x){return !!x.nama;});
}
function adminRecordFiles_(d){
 requireAdmin_(d.token);
 var type=String(d.type||''),recordId=String(d.recordId||'').trim();
 if((type!=='kontingen'&&type!=='cabang')||!recordId)throw new Error('Data detail tidak valid.');
 var cfg=CFG[type],ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss&&ss.getSheetByName(cfg.sheet);
 if(!sh||sh.getLastRow()<2)throw new Error('Data tidak ditemukan.');
 var lastCol=sh.getLastColumn(),headers=sh.getRange(1,1,1,lastCol).getValues()[0],idCol=headers.indexOf('ID Data')+1,filesCol=headers.indexOf('Berkas (link Drive)')+1;
 if(!idCol)throw new Error('Kolom ID Data tidak ditemukan.');
 var ids=sh.getRange(2,idCol,sh.getLastRow()-1,1).getValues(),rowNumber=0;
 for(var i=0;i<ids.length;i++){if(String(ids[i][0]||'').trim()===recordId){rowNumber=i+2;break;}}
 if(!rowNumber)throw new Error('Data tidak ditemukan atau sudah dihapus. Muat ulang daftar.');
 var raw=filesCol?sh.getRange(rowNumber,filesCol).getValue():'';
 return {ok:true,recordId:recordId,berkas:parsePortalFiles_(raw)};
}

function getPublicRecapMode_(){
 var mode=PropertiesService.getScriptProperties().getProperty('PUBLIC_RECAP_MODE');
 return mode==='terverifikasi'?'terverifikasi':'terdaftar';
}
function adminSetPublicRecapMode_(d){
 var admin=requireAdmin_(d.token),mode=String(d.mode||'').trim();
 if(['terdaftar','terverifikasi'].indexOf(mode)<0)throw new Error('Pilihan tampilan rekap tidak valid.');
 PropertiesService.getScriptProperties().setProperty('PUBLIC_RECAP_MODE',mode);
 invalidatePublicCache_('recap');
 logSecurityEvent_(admin.user,'UBAH TAMPILAN REKAP PUBLIK','BERHASIL','Mode: '+mode);
 return {ok:true,mode:mode,message:mode==='terdaftar'?'Beranda sekarang menampilkan semua pendaftaran cabang lomba.':'Beranda sekarang hanya menampilkan cabang lomba dari kontingen terverifikasi.'};
}

var NEWS_SHEET_NAME='Kabar PORSENI';
function newsSheet_(){var ss=SpreadsheetApp.getActiveSpreadsheet();if(!ss)throw new Error('Spreadsheet aktif tidak ditemukan.');var sh=ss.getSheetByName(NEWS_SHEET_NAME)||ss.insertSheet(NEWS_SHEET_NAME);if(sh.getLastRow()===0){sh.appendRow(['ID','Tanggal','Kategori','Judul','Ringkasan','Isi Berita','URL Gambar','Status','Diperbarui Oleh','Diperbarui Pada']);sh.setFrozenRows(1);sh.getRange(1,1,1,10).setFontWeight('bold');}return sh;}
function newsRows_(){var sh=newsSheet_(),v=sh.getDataRange().getValues();return v.length<2?[]:v.slice(1).map(function(r,i){return {row:i+2,id:String(r[0]||''),tanggal:r[1] instanceof Date?Utilities.formatDate(r[1],Session.getScriptTimeZone(),'yyyy-MM-dd'):String(r[1]||''),kategori:String(r[2]||''),judul:String(r[3]||''),ringkasan:String(r[4]||''),isi:String(r[5]||''),gambar:String(r[6]||''),status:String(r[7]||'Draft'),oleh:String(r[8]||''),updated:String(r[9]||'')};});}
function getPublicNews_(){
 var cache=CacheService.getScriptCache(),cacheKey='public_news_v1';
 try{var cached=cache.get(cacheKey);if(cached)return JSON.parse(cached);}catch(cacheReadErr){}
 var result=newsRows_().filter(function(n){return n.status==='Terbit';})
  .sort(function(a,b){return String(b.tanggal).localeCompare(String(a.tanggal))||b.row-a.row;})
  .slice(0,30).map(function(n){return {id:n.id,tanggal:n.tanggal,kategori:n.kategori,judul:n.judul,ringkasan:n.ringkasan,isi:n.isi,gambar:n.gambar};});
 try{var payload=JSON.stringify(result);if(payload.length<60000)cache.put(cacheKey,payload,60);}catch(cacheWriteErr){}
 return result;
}
function adminNewsList_(token){requireAdmin_(token);return {ok:true,news:newsRows_().sort(function(a,b){return b.row-a.row;})};}

function adminNewsImageUpload_(d){
 var admin=requireAdmin_(d.token),mime=String(d.mimeType||''),name=String(d.name||'foto-berita').replace(/[^a-zA-Z0-9._-]/g,'_').slice(0,100),data=String(d.data||'');
 if(!/^image\/(jpeg|png|webp)$/.test(mime))throw new Error('Format foto harus JPG, PNG, atau WebP.');
 if(!data||data.length>7*1024*1024)throw new Error('Foto kosong atau melebihi batas 5 MB.');
 var bytes=Utilities.base64Decode(data);if(bytes.length>5*1024*1024)throw new Error('Ukuran foto maksimal 5 MB.');
 var folderName='Kabar PORSENI';var folders=DriveApp.getFoldersByName(folderName),folder=folders.hasNext()?folders.next():DriveApp.createFolder(folderName);
 var blob=Utilities.newBlob(bytes,mime,name),file=folder.createFile(blob).setName(new Date().getTime()+'_'+name);
 file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
 var id=file.getId(),url='https://drive.google.com/uc?export=view&id='+id;
 logSecurityEvent_(admin.user,'UNGGAH FOTO BERITA','BERHASIL',file.getName());
 return {ok:true,id:id,url:url,name:file.getName()};
}
function adminNewsSave_(d){var admin=requireAdmin_(d.token),x=d.news||{},judul=String(x.judul||'').trim(),isi=String(x.isi||'').trim();if(!judul||!isi)throw new Error('Judul dan isi berita wajib diisi.');if(judul.length>180||isi.length>20000)throw new Error('Judul maksimal 180 karakter dan isi maksimal 20.000 karakter.');var status=x.status==='Terbit'?'Terbit':'Draft',tanggal=String(x.tanggal||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd'));if(!/^\d{4}-\d{2}-\d{2}$/.test(tanggal))throw new Error('Tanggal berita tidak valid.');var sh=newsSheet_(),rows=newsRows_(),id=String(x.id||'').trim(),existing=rows.filter(function(n){return n.id===id;})[0],row=[id||Utilities.getUuid(),tanggal,String(x.kategori||'Pengumuman').slice(0,50),judul,String(x.ringkasan||'').slice(0,500),isi,String(x.gambar||'').trim(),status,admin.user||'Admin',new Date()];if(existing)sh.getRange(existing.row,1,1,row.length).setValues([row]);else sh.appendRow(row);invalidatePublicCache_('news');logSecurityEvent_(admin.user,'KELOLA BERITA',status,judul);return {ok:true,message:existing?'Berita berhasil diperbarui.':'Berita berhasil ditambahkan.',news:row[0]};}
function adminNewsDelete_(d){var admin=requireAdmin_(d.token),id=String(d.id||'');if(!id)throw new Error('ID berita tidak ditemukan.');var sh=newsSheet_(),n=newsRows_().filter(function(x){return x.id===id;})[0];if(!n)throw new Error('Berita tidak ditemukan.');sh.deleteRow(n.row);invalidatePublicCache_('news');logSecurityEvent_(admin.user,'HAPUS BERITA','BERHASIL',n.judul);return {ok:true,message:'Berita berhasil dihapus.'};}

function adminDashboard_(token){
 requireAdmin_(token);
 var kont=adminReadSheet_(CFG.kontingen.sheet,'kontingen',false),cabang=adminReadSheet_(CFG.cabang.sheet,'cabang',false);
 var emailByName={},phoneByName={},leaderByName={};kont.forEach(function(r){var key=String(r.nama||'').trim().toLowerCase();if(key&&r.email)emailByName[key]=r.email;if(key&&r.waKetua)phoneByName[key]=r.waKetua;if(key&&r.ketua)leaderByName[key]=r.ketua;});
 cabang.forEach(function(r){var key=String(r.nama||'').trim().toLowerCase();r.email=emailByName[key]||'';r.waKetua=phoneByName[key]||'';r.ketua=leaderByName[key]||'';});
 var all=kont.concat(cabang),counts={total:all.length,menunggu:0,terverifikasi:0,revisi:0,ditolak:0,kontingen:kont.length,cabang:cabang.length};
 all.forEach(function(x){if(x.status==='Terverifikasi')counts.terverifikasi++;else if(x.status==='Perlu Revisi')counts.revisi++;else if(x.status==='Ditolak')counts.ditolak++;else counts.menunggu++;});
 // Permohonan perubahan tidak ditampilkan pada dashboard utama; jangan baca sheet dan
 // jangan lakukan lookup data kontingen berulang untuk setiap tiket saat login.
 return {ok:true,counts:counts,kontingen:kont,cabang:cabang,changeRequests:[],registration:registrationControls_(),modeRekapPublik:getPublicRecapMode_(),updatedAt:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd-MM-yyyy HH:mm:ss')};
}
function resolveAdminRecordRow_(sheet,d){
 if(!sheet)throw new Error('Sheet data tidak ditemukan.');
 var recordId=String(d.recordId||'').trim();
 if(!recordId)throw new Error('ID data tidak tersedia. Muat ulang portal panitia sebelum mengubah data.');
 var headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0],idCol=headers.indexOf('ID Data')+1;
 if(!idCol)throw new Error('Kolom ID Data belum tersedia. Muat ulang dashboard panitia untuk migrasi data.');
 var last=sheet.getLastRow();if(last<2)throw new Error('Data tidak ditemukan. Muat ulang daftar.');
 var ids=sheet.getRange(2,idCol,last-1,1).getValues();
 for(var i=0;i<ids.length;i++)if(String(ids[i][0]||'')===recordId)return i+2;
 throw new Error('Data sudah berubah atau dihapus oleh petugas lain. Muat ulang daftar sebelum melanjutkan.');
}
function adminUpdateStatus_(d){
 var admin=requireAdmin_(d.token),type=String(d.type||''),status=String(d.status||''),note=String(d.note||'').slice(0,1000);
 if(type!=='kontingen'&&type!=='cabang')throw new Error('Jenis data tidak valid.');
 if(ADMIN_STATUS_VALUES.indexOf(status)<0)throw new Error('Status verifikasi tidak valid.');
 if((status==='Perlu Revisi'||status==='Ditolak')&&!String(note||'').trim())throw new Error('Catatan wajib diisi untuk status Perlu Revisi atau Ditolak.');
 var lock=LockService.getScriptLock(),emailPayload=null,result=null;lock.waitLock(20000);
 try{
  var cfg=CFG[type],ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss&&ss.getSheetByName(cfg.sheet);
  if(!sheet)throw new Error('Data pendaftaran tidak ditemukan. Muat ulang daftar.');
  ensureAdminColumns_(sheet);
  var rowNumber=resolveAdminRecordRow_(sheet,d),headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var nameCol=headers.indexOf('Nama Kontingen')+1,statusCol=headers.indexOf('Status Verifikasi')+1,noteCol=headers.indexOf('Catatan Panitia')+1,dateCol=headers.indexOf('Diverifikasi Pada')+1,byCol=headers.indexOf('Diverifikasi Oleh')+1;
  var name=String(sheet.getRange(rowNumber,nameCol).getValue()||'');if(!name)throw new Error('Baris data kosong.');
  var currentStatus=statusCol?String(sheet.getRange(rowNumber,statusCol).getValue()||'Menunggu Verifikasi'):'Menunggu Verifikasi';
  if(currentStatus==='Terverifikasi'&&status!=='Menunggu Verifikasi')throw new Error('Status sudah Terverifikasi. Batalkan verifikasi terlebih dahulu sebelum mengubah status.');
  if(currentStatus!=='Terverifikasi'&&status==='Menunggu Verifikasi')throw new Error('Status saat ini bukan Terverifikasi.');
  var regCol=headers.indexOf('No. Registrasi')+1,branchCol=headers.indexOf('Cabang Lomba')+1,categoryCol=headers.indexOf('Kategori')+1;
  var now=new Date();sheet.getRange(rowNumber,statusCol).setValue(status);sheet.getRange(rowNumber,noteCol).setValue(note);sheet.getRange(rowNumber,dateCol).setValue(now);sheet.getRange(rowNumber,byCol).setValue(admin.user||'Admin');
  invalidatePublicCache_('recap');
  var reg=regCol?sheet.getRange(rowNumber,regCol).getValue():'';
  adminAudit_(admin,'UBAH STATUS VERIFIKASI',type,rowNumber,reg,name+(type==='cabang'?' | '+String(sheet.getRange(rowNumber,branchCol).getValue()||'')+' / '+String(sheet.getRange(rowNumber,categoryCol).getValue()||''):'')+' | '+status+' | '+note);
  emailPayload=currentStatus!==status&&status!=='Menunggu Verifikasi'&&!(type==='kontingen'&&status==='Terverifikasi')?{token:d.token,type:type,recordId:d.recordId,status:status,note:note}:null;
  result={ok:true,message:'Status berhasil diperbarui.',status:status,updatedAt:Utilities.formatDate(now,Session.getScriptTimeZone(),'dd-MM-yyyy HH:mm:ss')};
 }finally{lock.releaseLock();}
 // Email dikirim setelah status tersimpan dan lock dilepas. Kegagalan email tidak membatalkan keputusan admin.
 if(emailPayload){
  try{
   var mail=adminSendEmail_(emailPayload);
   result.emailSent=true;result.recipient=mail.recipient;result.message+=' Pemberitahuan email berhasil dikirim.';
  }catch(mailErr){
   result.emailSent=false;result.emailError=String(mailErr&&mailErr.message||mailErr);
   result.message+=' Namun email belum berhasil dikirim: '+result.emailError;
   try{logSecurityEvent_(admin.user,'NOTIFIKASI EMAIL VERIFIKASI','GAGAL',type+' | '+String(d.recordId||'')+' | '+status+' | '+result.emailError);}catch(ignore){}
  }
 }else{
  result.emailSent=false;
  if(type==='kontingen'&&status==='Terverifikasi'&&currentStatus!=='Terverifikasi'){
   result.message+=' Email verifikasi akan dikirim satu kali setelah PDF formulir dan seluruh lampiran selesai digabungkan.';
   result.emailPending=true;
  }else{
   result.message+=' Status tidak berubah atau masih menunggu verifikasi; email tidak dikirim.';
  }
 }
 return result;
}
function adminAudit_(admin,action,type,rowNumber,reference,details){
 logSecurityEvent_(admin&&admin.user||'Admin','AKTIVITAS ADMIN: '+String(action||''),'BERHASIL',String(type||'')+' | '+String(reference||'')+' | '+String(details||''));
 var ss=SpreadsheetApp.getActiveSpreadsheet(),log=ss.getSheetByName(ADMIN_LOG_SHEET)||ss.insertSheet(ADMIN_LOG_SHEET);
 var logHeaders=['Waktu','Admin','Jenis Data','Baris Sheet','Nomor Registrasi','Nama Kontingen / Referensi','Status / Aksi','Catatan'];
 if(log.getLastRow()===0){log.getRange(1,1,1,logHeaders.length).setValues([logHeaders]).setFontWeight('bold');log.setFrozenRows(1);}
 else{var currentHeaders=log.getRange(1,1,1,Math.max(log.getLastColumn(),logHeaders.length)).getValues()[0];var headerMismatch=logHeaders.some(function(h,i){return currentHeaders[i]!==h;});if(headerMismatch)log.getRange(1,1,1,logHeaders.length).setValues([logHeaders]).setFontWeight('bold');}
 log.appendRow([new Date(),admin.user||'Admin',type||'',rowNumber||'',reference||'',details||'',action||'', '']);
}
function adminUpdateRecord_(d){
 var admin=requireAdmin_(d.token),type=String(d.type||'');
 if(type!=='kontingen'&&type!=='cabang')throw new Error('Jenis data tidak valid.');
 var lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG[type].sheet);
  if(!sheet)throw new Error('Data tidak ditemukan. Muat ulang halaman.');
  ensureAdminColumns_(sheet);
  var rowNumber=resolveAdminRecordRow_(sheet,d),headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var allowed=type==='kontingen'?{asal:'Asal',alamat:'Alamat',ketua:'Ketua',tempatLahirKetua:'Tempat Lahir Ketua',tanggalLahirKetua:'Tanggal Lahir Ketua',waKetua:'WA Ketua',email:'Email',jumlahPeserta:'Jumlah Peserta',cabang:'Cabang Lomba',metodeBayar:'Metode Bayar',tanggalBayar:'Tgl Bayar',ketuaIpnu:'Ketua PR IPNU',ketuaIppnu:'Ketua PR IPPNU'}:{cabang:'Cabang Lomba',kategori:'Kategori',jumlahPeserta:'Jumlah Peserta',peserta:'Daftar Peserta',tempatLahirPeserta:'Tempat Lahir Peserta',tanggalLahir:'Tanggal Lahir Peserta',alamatPeserta:'Alamat Rumah Peserta',waPeserta:'WA Peserta',catatan:'Catatan'};
  var fields=d.fields||{},changed=[];
  Object.keys(fields).forEach(function(key){if(!allowed[key])return;var col=headers.indexOf(allowed[key]);if(col<0)return;var value=String(fields[key]===null?'':fields[key]);if(/^[=+@-]/.test(value))value="'"+value;sheet.getRange(rowNumber,col+1).setValue(value);changed.push(allowed[key]);});
  if(!changed.length)throw new Error('Tidak ada kolom yang dapat diperbarui.');
  var nameCol=headers.indexOf('Nama Kontingen')+1,regCol=headers.indexOf('No. Registrasi')+1,name=String(sheet.getRange(rowNumber,nameCol).getValue()||'');
  // Akses pratinjau berkas kini memakai PIN Portal Kontingen, bukan ACL email Drive.
  invalidatePublicCache_('recap');
  adminAudit_(admin,'EDIT DATA',type,rowNumber,regCol?sheet.getRange(rowNumber,regCol).getValue():'',name+' | '+changed.join(', '));
  return {ok:true,message:'Data berhasil diperbarui: '+changed.join(', ')};
 }finally{lock.releaseLock();}
}
function adminReplaceFile_(d){
 var admin=requireAdmin_(d.token),type=String(d.type||''),label=String(d.label||'').trim(),name=String(d.name||'lampiran').trim(),mime=String(d.mimeType||'application/octet-stream');
 if(type!=='kontingen'&&type!=='cabang')throw new Error('Jenis data tidak valid.');
 if(!label)throw new Error('Label berkas wajib diisi.');if(!d.fileData)throw new Error('Pilih berkas pengganti terlebih dahulu.');
 var bytes=Utilities.base64Decode(String(d.fileData));if(bytes.length>MAX_FILE_BYTES)throw new Error('Unggahan terlalu besar untuk diproses Apps Script.');if(/^ktp(?: peserta\s+\d+)?$/i.test(label))assertFileKind_(bytes,['png','jpg'],label+' wajib berupa gambar JPG/PNG.');
 var lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG[type].sheet);if(!sheet)throw new Error('Data tidak ditemukan. Muat ulang halaman.');
  var rowNumber=resolveAdminRecordRow_(sheet,d),headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var folderCol=headers.indexOf('Folder Drive')+1,filesCol=headers.indexOf('Berkas (link Drive)')+1;if(!folderCol||!filesCol)throw new Error('Kolom folder atau berkas tidak ditemukan.');
  var folderUrl=String(sheet.getRange(rowNumber,folderCol).getValue()||''),match=folderUrl.match(/[-\w]{25,}/);if(!match)throw new Error('Folder Drive data ini tidak ditemukan.');
  var folder=DriveApp.getFolderById(match[0]);assertPrivateFolderTree_(folder);
  var old=String(sheet.getRange(rowNumber,filesCol).getValue()||''),lines=old?old.split(/\r?\n/):[],oldUrls=[];
  lines.forEach(function(line){var ix=line.indexOf(': ');if(ix>=0&&line.slice(0,ix).trim().toLowerCase()===label.toLowerCase())oldUrls.push(line.slice(ix+2).trim());});
  var nameIndex=headers.indexOf('Nama Kontingen'),contingentName=nameIndex>=0?String(sheet.getRange(rowNumber,nameIndex+1).getValue()||'').trim():'';
  var regIndex=headers.indexOf('No. Registrasi'),registration=regIndex>=0?String(sheet.getRange(rowNumber,regIndex+1).getValue()||'').trim():'';
  var branchIndex=headers.indexOf('Cabang Lomba'),categoryIndex=headers.indexOf('Kategori');
  var branch=branchIndex>=0?String(sheet.getRange(rowNumber,branchIndex+1).getValue()||'').trim():'';
  var category=categoryIndex>=0?String(sheet.getRange(rowNumber,categoryIndex+1).getValue()||'').trim():'';
  var preserveName=String(name||'lampiran').replace(/[\\/:*?"<>|\r\n]+/g,'_').slice(0,180)||'lampiran';
  var standardizedName=attachmentFileName_({type:type,nama:contingentName,reg:registration,cabang:branch,kategori:category,label:label,originalName:preserveName});
  var blob=Utilities.newBlob(bytes,detectedMime_(bytes,mime),standardizedName),file=folder.createFile(blob),url=file.getUrl();
  try{assertPrivateIndividualFile_(file);}catch(shareErr){try{file.setTrashed(true);}catch(ignore){}throw shareErr;}
  var found=false;lines=lines.map(function(line){var ix=line.indexOf(': ');if(ix>=0&&line.slice(0,ix).trim().toLowerCase()===label.toLowerCase()){found=true;return label+': '+url;}return line;});if(!found)lines.push(label+': '+url);
  sheet.getRange(rowNumber,filesCol).setValue(lines.join('\n'));
  if(label.toLowerCase()==='pdf formulir'){var pdfCol=headers.indexOf('PDF Formulir (link Drive)')+1;if(pdfCol)sheet.getRange(rowNumber,pdfCol).setValue(url);}
  var regCol=headers.indexOf('No. Registrasi')+1;adminAudit_(admin,'GANTI BERKAS',type,rowNumber,regCol?sheet.getRange(rowNumber,regCol).getValue():'',contingentName+' | '+label+' | '+name);
  return {ok:true,message:'Berkas berhasil diganti. File baru tetap privat dan pratinjau dilakukan melalui PIN Portal Kontingen.',url:url,label:label};
 }finally{lock.releaseLock();}
}
function adminDeleteFile_(d){
 var admin=requireAdmin_(d.token),type=String(d.type||''),label=String(d.label||'').trim(),fileId=String(d.fileId||'').trim();
 if(type!=='kontingen'&&type!=='cabang')throw new Error('Jenis data tidak valid.');
 if(!label||!/^[-\w]{25,}$/.test(fileId))throw new Error('Identitas berkas tidak valid.');
 var lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  var ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss&&ss.getSheetByName(CFG[type].sheet);
  if(!sheet)throw new Error('Data tidak ditemukan. Muat ulang halaman.');
  var rowNumber=resolveAdminRecordRow_(sheet,d),headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var fileCol=headers.indexOf('Berkas (link Drive)')+1,folderCol=headers.indexOf('Folder Drive')+1,regCol=headers.indexOf('No. Registrasi')+1,nameCol=headers.indexOf('Nama Kontingen')+1;
  if(!fileCol||!folderCol)throw new Error('Kolom berkas atau folder tidak ditemukan.');
  var raw=String(sheet.getRange(rowNumber,fileCol).getValue()||''),lines=raw?raw.split(/\r?\n/):[],targetIndex=-1,targetUrl='';
  for(var i=0;i<lines.length;i++){
   var sep=lines[i].indexOf(': ');if(sep<0)continue;
   var currentLabel=lines[i].slice(0,sep).trim(),url=lines[i].slice(sep+2).trim();
   if(currentLabel.toLowerCase()===label.toLowerCase()&&getDriveFileIdFromUrl_(url)===fileId){targetIndex=i;targetUrl=url;break;}
  }
  if(targetIndex<0)throw new Error('Berkas tidak tercatat pada data pendaftaran ini. Muat ulang Portal Admin.');
  var folderUrl=String(sheet.getRange(rowNumber,folderCol).getValue()||''),folderMatch=folderUrl.match(/[-\w]{25,}/);
  if(!folderMatch)throw new Error('Folder penyimpanan berkas tidak ditemukan.');
  var folder=DriveApp.getFolderById(folderMatch[0]);assertPrivateFolderTree_(folder);
  var file=DriveApp.getFileById(fileId);if(file.isTrashed())throw new Error('Berkas sudah berada di Trash Google Drive.');
  var parents=file.getParents(),belongs=false;while(parents.hasNext()){var parent=parents.next();if(parent.getId()===folder.getId())belongs=true;assertPrivateFolderTree_(parent);}
  if(!belongs)throw new Error('Berkas tidak berada di folder milik data pendaftaran ini.');
  var fileName=file.getName(),previousRaw=raw;lines.splice(targetIndex,1);
  try{
   file.setTrashed(true);
   sheet.getRange(rowNumber,fileCol).setValue(lines.filter(function(x){return String(x||'').trim();}).join('\n'));
  }catch(deleteErr){
   try{file.setTrashed(false);}catch(restoreErr){console.error('Pemulihan file dari Trash gagal: '+restoreErr);}
   try{sheet.getRange(rowNumber,fileCol).setValue(previousRaw);}catch(sheetRestoreErr){console.error('Pemulihan daftar file gagal: '+sheetRestoreErr);}
   throw new Error('Penghapusan berkas belum selesai. Sistem mencoba memulihkan file dan daftar berkas. Detail: '+String(deleteErr&&deleteErr.message||deleteErr));
  }
  var contingentName=nameCol?String(sheet.getRange(rowNumber,nameCol).getValue()||''):'';
  adminAudit_(admin,'HAPUS BERKAS',type,rowNumber,regCol?sheet.getRange(rowNumber,regCol).getValue():'',contingentName+' | '+label+' | '+fileName+' | '+fileId);
  return {ok:true,message:'Berkas dipindahkan ke Trash Google Drive dan dihapus dari daftar berkas. Aktivitas tercatat pada Log Aktivitas.',fileName:fileName,label:label};
 }catch(deleteActionError){
  logSecurityEvent_(admin&&admin.user||'Admin','HAPUS BERKAS','GAGAL',String(type||'')+' | '+String(label||'')+' | '+String(fileId||'')+' | '+String(deleteActionError&&deleteActionError.message||deleteActionError));
  throw deleteActionError;
 }finally{lock.releaseLock();}
}

function adminDeleteRecord_(d){
 var admin=requireAdmin_(d.token),type=String(d.type||'');if(type!=='kontingen'&&type!=='cabang')throw new Error('Jenis data tidak valid.');
 var lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  var ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName(CFG[type].sheet);if(!sheet)throw new Error('Data tidak ditemukan. Muat ulang halaman.');
  var rowNumber=resolveAdminRecordRow_(sheet,d),headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0],values=sheet.getRange(rowNumber,1,1,sheet.getLastColumn()).getValues()[0];
  function value(h){var i=headers.indexOf(h);return i<0?'':values[i];}
  var contingentName=String(value('Nama Kontingen')||''),email=type==='kontingen'?String(value('Email')||'').trim().toLowerCase():getContingentEmail_(contingentName);
  var archive=ss.getSheetByName('Arsip Data Terhapus')||ss.insertSheet('Arsip Data Terhapus');
  archiveDeletedSnapshot_(archive,admin,type,headers,values);
  revokeRecordPreviewAccess_(sheet,rowNumber,email);
  adminAudit_(admin,'HAPUS DATA',type,rowNumber,value('No. Registrasi'),contingentName+(type==='cabang'?' | '+value('Cabang Lomba')+' / '+value('Kategori'):''));
  sheet.deleteRow(rowNumber);
  invalidatePublicCache_('recap');
  if(type==='kontingen'){
   // Hapus baris cabang terkait dan PIN dari database aktif; snapshot tetap disimpan untuk pemulihan.
   var child=ss.getSheetByName(CFG.cabang.sheet);
   if(child&&child.getLastRow()>=2){var ch=child.getRange(1,1,1,child.getLastColumn()).getValues()[0],ci=ch.indexOf('Nama Kontingen');for(var r=child.getLastRow();r>=2;r--){var row=child.getRange(r,1,1,child.getLastColumn()).getValues()[0];if(ci>=0&&normalizeEvent_(row[ci])===normalizeEvent_(contingentName)){archiveDeletedSnapshot_(archive,admin,'cabang terkait',ch,row);revokeRecordPreviewAccess_(child,r,email);adminAudit_(admin,'HAPUS CABANG TERKAIT','cabang',r,'',contingentName+' | '+String(row[ch.indexOf('Cabang Lomba')]||'')+' / '+String(row[ch.indexOf('Kategori')]||''));child.deleteRow(r);}}}
   var pinSheet=ss.getSheetByName(PIN_SHEET);if(pinSheet&&pinSheet.getLastRow()>=2){var pv=pinSheet.getDataRange().getValues();for(var pr=pv.length-1;pr>=1;pr--){if(normalizeEvent_(pv[pr][0])===normalizeEvent_(contingentName))pinSheet.deleteRow(pr+1);}}
  }
  protectDeletedArchive_(archive);
  return {ok:true,message:type==='kontingen'?'Data kontingen, seluruh data cabang terkait, dan PIN dihapus dari daftar aktif. Snapshot disimpan di Arsip Data Terhapus; akses pratinjau dokumen kontingen telah dicabut.':'Data cabang lomba dihapus dari daftar aktif. Snapshot disimpan di Arsip Data Terhapus dan akses pratinjau dokumennya telah dicabut.'};
 }finally{lock.releaseLock();}
}
/** Mengarsipkan pendaftaran kontingen yang ditolak agar nama dapat dipakai untuk pendaftaran ulang. */
function adminAllowReregistration_(d){
 var admin=requireAdmin_(d.token),type=String(d.type||'');
 if(type!=='kontingen')throw new Error('Pendaftaran ulang hanya berlaku untuk data kontingen.');
 var lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  var ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName(CFG.kontingen.sheet);
  if(!sheet)throw new Error('Sheet kontingen tidak ditemukan.');
  ensureAdminColumns_(sheet);
  var rowNumber=resolveAdminRecordRow_(sheet,d),headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0],values=sheet.getRange(rowNumber,1,1,sheet.getLastColumn()).getValues()[0];
  function value(h){var i=headers.indexOf(h);return i<0?'':values[i];}
  var name=String(value('Nama Kontingen')||'').trim(),status=String(value('Status Verifikasi')||'');
  if(!name)throw new Error('Nama kontingen tidak ditemukan.');
  if(normalizeVerificationStatus_(status)!=='ditolak')throw new Error('Pendaftaran ulang hanya dapat diaktifkan untuk kontingen berstatus Ditolak.');
  var email=String(value('Email')||'').trim().toLowerCase();
  var archive=ss.getSheetByName('Arsip Data Terhapus')||ss.insertSheet('Arsip Data Terhapus');
  archiveDeletedSnapshot_(archive,admin,'kontingen — diizinkan daftar ulang',headers,values);
  revokeRecordPreviewAccess_(sheet,rowNumber,email);
  var regCol=headers.indexOf('No. Registrasi')+1;
  adminAudit_(admin,'IZINKAN PENDAFTARAN ULANG','kontingen',rowNumber,regCol?value('No. Registrasi'):'',name+' | status sebelumnya: Ditolak');
  sheet.deleteRow(rowNumber);
  invalidatePublicCache_('recap');
  // Arsipkan dan hapus semua data cabang lama agar pendaftaran baru dimulai dari keadaan bersih.
  var child=ss.getSheetByName(CFG.cabang.sheet);
  if(child&&child.getLastRow()>=2){
   var ch=child.getRange(1,1,1,child.getLastColumn()).getValues()[0],ci=ch.indexOf('Nama Kontingen');
   if(ci>=0){for(var r=child.getLastRow();r>=2;r--){var childValues=child.getRange(r,1,1,child.getLastColumn()).getValues()[0];if(normalizeEvent_(childValues[ci])===normalizeEvent_(name)){archiveDeletedSnapshot_(archive,admin,'cabang terkait — pendaftaran ulang',ch,childValues);revokeRecordPreviewAccess_(child,r,email);adminAudit_(admin,'ARSIPKAN CABANG UNTUK DAFTAR ULANG','cabang',r,'',name);child.deleteRow(r);}}}
  }
  // Simpan snapshot PIN sebelum menghapusnya; pendaftaran baru akan memperoleh PIN baru.
  var pinSheet=ss.getSheetByName(PIN_SHEET);
  if(pinSheet&&pinSheet.getLastRow()>=2){var pv=pinSheet.getDataRange().getValues(),ph=pv[0],pn=ph.indexOf('Nama Kontingen');for(var pr=pv.length-1;pr>=1;pr--){if(pn>=0&&normalizeEvent_(pv[pr][pn])===normalizeEvent_(name)){archiveDeletedSnapshot_(archive,admin,'PIN kontingen — pendaftaran ulang',ph,pv[pr]);pinSheet.deleteRow(pr+1);}}}
  clearPinFailures_(name);protectDeletedArchive_(archive);
  return {ok:true,message:'Pendaftaran lama, seluruh berkas cabang terkait, dan PIN telah diarsipkan. Nama kontingen kini tersedia untuk pendaftaran ulang; PIN baru akan dibuat saat pendaftaran berhasil.'};
 }finally{lock.releaseLock();}
}

function ensureResultSheet_(){
 var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName(RESULT_SHEET)||ss.insertSheet(RESULT_SHEET);
 if(sh.getLastRow()===0){sh.appendRow(RESULT_HEADERS);sh.setFrozenRows(1);sh.getRange(1,1,1,RESULT_HEADERS.length).setFontWeight('bold');}
 else{var hs=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];RESULT_HEADERS.forEach(function(h){if(hs.indexOf(h)<0){sh.getRange(1,sh.getLastColumn()+1).setValue(h).setFontWeight('bold');hs.push(h);}});}
 return sh;
}
function getChampionshipResults_(){
 var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss&&ss.getSheetByName(RESULT_SHEET),detail=[],medalMap={};
 if(sh&&sh.getLastRow()>=2){var rows=sh.getDataRange().getValues(),h=rows[0],ix={};h.forEach(function(v,i){ix[String(v)]=i;});
  rows.slice(1).forEach(function(r){var branch=String(r[ix['Cabang Lomba']]||'').trim(),cat=String(r[ix['Kategori']]||'').trim();if(!branch||!cat)return;
   var item={cabang:branch,kategori:cat,juara1:String(r[ix['Juara 1']]||'').trim(),juara2:String(r[ix['Juara 2']]||'').trim(),juara3:String(r[ix['Juara 3']]||'').trim(),updatedAt:r[ix['Diperbarui Pada']]?Utilities.formatDate(new Date(r[ix['Diperbarui Pada']]),Session.getScriptTimeZone(),'dd-MM-yyyy HH:mm'):'',updatedBy:String(r[ix['Diperbarui Oleh']]||'')};detail.push(item);
   [[item.juara1,'g',5],[item.juara2,'s',3],[item.juara3,'b',1]].forEach(function(pair){var name=pair[0];if(!name)return;if(!medalMap[name])medalMap[name]={k:name,g:0,s:0,b:0,points:0};medalMap[name][pair[1]]++;medalMap[name].points+=pair[2];});
  });
 }
 var order=Object.keys(KAT);detail.sort(function(a,b){var ai=order.indexOf(a.cabang),bi=order.indexOf(b.cabang);if(ai!==bi)return (ai<0?999:ai)-(bi<0?999:bi);var cats=KAT[a.cabang]||[];return cats.indexOf(a.kategori)-cats.indexOf(b.kategori);});
 var medals=Object.keys(medalMap).map(function(k){return medalMap[k];}).sort(function(a,b){return b.points-a.points||b.g-a.g||b.s-a.s||b.b-a.b||a.k.localeCompare(b.k,'id');});
 return {detail:detail,medals:medals};
}
function adminGetResults_(token){requireAdmin_(token);return {ok:true,results:getChampionshipResults_(),finalization:getChampionshipFinalization_(),catalog:getChampionshipCatalog_()};}
function getChampionshipCatalog_(){return Object.keys(KAT).map(function(branch){return {cabang:branch,kategori:(KAT[branch]||[]).slice()};});}
function getChampionshipFinalization_(){
 var raw=PropertiesService.getScriptProperties().getProperty(CHAMPIONSHIP_FINALIZATION_PROPERTY);
 if(!raw)return {finalized:false,finalizedAt:'',finalizedBy:'',finalizationId:''};
 try{var d=JSON.parse(raw),at=String(d.finalizedAt||'');return {finalized:d.finalized===true,finalizedAt:at,finalizedBy:String(d.finalizedBy||''),finalizationId:String(d.finalizationId||('legacy-'+at.replace(/[^0-9]/g,'')))};}
 catch(e){return {finalized:false,finalizedAt:'',finalizedBy:'',finalizationId:''};}
}
function invalidateChampionshipFinalization_(admin,reason){
 var state=getChampionshipFinalization_();if(!state.finalized)return;
 retireChampionshipPdf_();
 PropertiesService.getScriptProperties().setProperty(CHAMPIONSHIP_FINALIZATION_PROPERTY,JSON.stringify({finalized:false,finalizedAt:'',finalizedBy:'',lastReopenedAt:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd-MM-yyyy HH:mm:ss'),lastReopenedBy:admin.user||'Admin',reason:reason||'Perubahan hasil'}));
 adminAudit_(admin,'BATAL OTOMATIS FINALISASI','hasil','', '',reason||'Hasil berubah setelah finalisasi.');
}
function buildChampionshipPdfBlob_(state,isTest){
 isTest=!!isTest;
 var data=getChampionshipResults_(),catalog=getChampionshipCatalog_(),byKey={};
 data.detail.forEach(function(x){byKey[normalizeEvent_(x.cabang)+'|'+normalizeEvent_(x.kategori)]=x;});
 var doc=pdfNewDoc_('TEMP_HASIL_RESMI_PORSENI_2026_'+new Date().getTime(),0);
 var body=doc.getBody();
 body.setMarginTop(42).setMarginBottom(42).setMarginLeft(52).setMarginRight(52);
 var navy='#24364B',muted='#666666',cream='#F7F3E8',ivory='#FFFDF7',softGray='#EEF0F2',line='#D8D5CC';
 function center(text,size,bold){var q=body.appendParagraph(text);q.setAlignment(DocumentApp.HorizontalAlignment.CENTER);q.editAsText().setFontFamily('Arial').setFontSize(size||10).setBold(!!bold);return q;}
 function para(text,size,bold){var q=body.appendParagraph(text);q.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);q.setSpacingAfter(5);q.editAsText().setFontFamily('Arial').setFontSize(size||10).setBold(!!bold);return q;}
 function pageTitle(kicker,title,sub){
  center(kicker,9,true);center(title,15,true);if(sub)center(sub,9,false);
  var line=body.appendParagraph('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');line.setAlignment(DocumentApp.HorizontalAlignment.CENTER);line.setSpacingAfter(10);line.editAsText().setFontSize(7).setForegroundColor(navy);
 }
 function styleTable(table,header){
  table.setBorderWidth(0.6).setBorderColor(line);
  for(var r=0;r<table.getNumRows();r++)for(var c=0;c<table.getRow(r).getNumCells();c++){
   var cell=table.getCell(r,c);cell.setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(7).setPaddingRight(7);
   cell.editAsText().setFontFamily('Arial').setFontSize(9).setForegroundColor('#222222');
   if(header&&r===0){cell.setBackgroundColor(cream);cell.editAsText().setBold(true).setForegroundColor(navy);}
   else cell.setBackgroundColor(r%2===0?ivory:softGray);
  }
 }
 if(isTest){var draft=center('DOKUMEN UJI COBA — BELUM RESMI',11,true);draft.editAsText().setForegroundColor('#B3261E');}

 // HALAMAN 1 — SURAT KEPUTUSAN DAN PENETAPAN JUARA UMUM
 pageTitle('PANITIA PELAKSANA PEKAN OLAHRAGA DAN SENI','SURAT KEPUTUSAN','PORSENI 2026 · PAC IPNU-IPPNU PATEBON');
 center('Nomor: ....../SK/PORSENI/2026',9,false);
 center('TENTANG',10,true);
 center('PENETAPAN JUARA UMUM PORSENI 2026',12,true);
 center('PANITIA PELAKSANA PORSENI 2026',10,true);
 para('Dengan memohon rahmat Tuhan Yang Maha Esa, Panitia Pelaksana Pekan Olahraga dan Seni (PORSENI) 2026 PAC IPNU-IPPNU Patebon, setelah memperhatikan dan mempertimbangkan hasil seluruh cabang perlombaan, menetapkan keputusan sebagai berikut:',9,false);
 para('Menimbang: bahwa untuk memberikan kepastian dan pengesahan atas hasil kejuaraan PORSENI 2026, perlu ditetapkan Juara Umum berdasarkan hasil akhir seluruh cabang perlombaan.',9,false);
 para('Memperhatikan: hasil penilaian dan rekapitulasi panitia pada seluruh cabang perlombaan PORSENI 2026.',9,false);
 center('MEMUTUSKAN',11,true);
 para('Menetapkan: KEPUTUSAN PANITIA PELAKSANA PORSENI 2026 TENTANG PENETAPAN JUARA UMUM PORSENI 2026.',9,true);
 para('Menetapkan tiga kontingen sebagai Juara Umum PORSENI 2026 dengan urutan sebagai berikut:',9,false);
 var winnerRows=[['Peringkat','Nama Kontingen']];
 for(var wi=0;wi<3;wi++){var wm=(data.medals||[])[wi];winnerRows.push(['JUARA UMUM '+(wi+1),wm?wm.k:'—']);}
 var wt=body.appendTable(winnerRows);styleTable(wt,true);
 para('Keputusan ini berlaku sejak tanggal ditetapkan. Apabila di kemudian hari terdapat kekeliruan, akan dilakukan perbaikan sebagaimana mestinya.',9,false);
 center(isTest?'Patebon, .................... 2026':'Patebon, .... ................ 2026',9,false);
 var sign=body.appendTable([
  ['Mengetahui,','','Ditetapkan oleh,'],
  ['Koordinator Perlombaan','','Ketua Panitia Pelaksana'],
  ['\n\n\n\n','RUANG STEMPEL','\n\n\n\n'],
  ['Sochibul Mustaqim','','Prasetyo Wahyu A.R.']
 ]);
 sign.setBorderWidth(0);
 for(var sr=0;sr<sign.getNumRows();sr++)for(var sc=0;sc<3;sc++){
  var cell=sign.getCell(sr,sc);cell.setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(4).setPaddingRight(4);
  cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  cell.editAsText().setFontFamily('Arial').setFontSize(sr===2&&sc===1?7:9).setBold(sr===1||sr===3);
  if(sr===2&&sc===1){cell.setBackgroundColor(cream);cell.setBorderColor(line);}
 }

 // LAMPIRAN — HASIL SETIAP CABANG/KATEGORI, TANPA REKAP MEDALI/POIN
 catalog.forEach(function(g){
  body.appendPageBreak();
  pageTitle('LAMPIRAN SURAT KEPUTUSAN PANITIA PELAKSANA PORSENI 2026','HASIL KEJUARAAN CABANG LOMBA',displayBranchName_(g.cabang));
  var rows=[['No.','Kategori','Juara 1','Juara 2','Juara 3']];
  g.kategori.forEach(function(c,ci){var x=byKey[normalizeEvent_(g.cabang)+'|'+normalizeEvent_(c)]||{};rows.push([String(ci+1),String(c),x.juara1||'—',x.juara2||'—',x.juara3||'—']);});
  var bt=body.appendTable(rows);styleTable(bt,true);
 });
 var foot=body.appendParagraph(isTest?'DOKUMEN UJI COBA — BELUM RESMI. Data dapat berubah dan dokumen ini tidak dikirim kepada kontingen.':'Dokumen ini merupakan lampiran resmi Surat Keputusan Penetapan Juara Umum PORSENI 2026 setelah dilakukan finalisasi oleh Panitia Pelaksana.');
 foot.setAlignment(DocumentApp.HorizontalAlignment.CENTER);foot.setSpacingBefore(12);foot.editAsText().setFontFamily('Arial').setFontSize(8).setItalic(true).setForegroundColor(muted);
 doc.saveAndClose();
 var file=DriveApp.getFileById(doc.getId()),blob=file.getAs(MimeType.PDF).setName('Surat Keputusan Hasil Kejuaraan PORSENI 2026.pdf');
 file.setTrashed(true);return blob;
}
function ensureChampionshipPdfFile_(state){
 var props=PropertiesService.getScriptProperties(),fileId=String(props.getProperty(CHAMPIONSHIP_PDF_FILE_PROPERTY)||''),file=null;
 if(fileId){try{file=DriveApp.getFileById(fileId);if(file.isTrashed())file=null;}catch(e){file=null;}}
 if(!file){
  var blob=buildChampionshipPdfBlob_(state,false),folders=DriveApp.getFoldersByName('HASIL RESMI PORSENI 2026'),folder=folders.hasNext()?folders.next():DriveApp.createFolder('HASIL RESMI PORSENI 2026');
  file=folder.createFile(blob.setName('Surat Keputusan Hasil Kejuaraan PORSENI 2026.pdf'));
  props.setProperty(CHAMPIONSHIP_PDF_FILE_PROPERTY,file.getId());
 }
 file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
 return {fileId:file.getId(),downloadUrl:'https://drive.google.com/uc?export=download&id='+file.getId(),previewUrl:'https://drive.google.com/file/d/'+file.getId()+'/view'};
}
function getPublicChampionshipPdf_(){
 var state=getChampionshipFinalization_();
 if(!state.finalized)return {available:false,downloadUrl:'',message:'PDF resmi belum difinalisasi.'};
 var props=PropertiesService.getScriptProperties(),id=String(props.getProperty(CHAMPIONSHIP_PDF_FILE_PROPERTY)||'');
 if(!id)return {available:false,downloadUrl:'',message:'PDF resmi sedang disiapkan.'};
 try{var file=DriveApp.getFileById(id);if(file.isTrashed())return {available:false,downloadUrl:'',message:'PDF resmi sedang disiapkan.'};return {available:true,downloadUrl:'https://drive.google.com/uc?export=download&id='+id,previewUrl:'https://drive.google.com/file/d/'+id+'/view',fileName:file.getName()};}
 catch(e){return {available:false,downloadUrl:'',message:'PDF resmi sedang disiapkan.'};}
}
function retireChampionshipPdf_(){
 var props=PropertiesService.getScriptProperties(),id=String(props.getProperty(CHAMPIONSHIP_PDF_FILE_PROPERTY)||'');
 if(!id)return;
 try{var f=DriveApp.getFileById(id);f.setSharing(DriveApp.Access.PRIVATE,DriveApp.Permission.NONE);f.setTrashed(true);}catch(e){console.warn('PDF hasil lama tidak dapat dinonaktifkan: '+e);}
 props.deleteProperty(CHAMPIONSHIP_PDF_FILE_PROPERTY);
}
function emailFinalResultsToContingents_(state,retryFailedOnly){
 var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName(CFG.kontingen.sheet);
 var result={sent:0,failed:0,skipped:0,details:[]};
 if(!sh||sh.getLastRow()<2)return {sent:0,failed:0,skipped:0,details:['Belum ada data kontingen untuk dikirimi email.']};
 var values=sh.getDataRange().getValues(),headers=values[0],iName=headers.indexOf('Nama Kontingen'),iLeader=headers.indexOf('Ketua'),iEmail=headers.indexOf('Email');
 if(iEmail<0||iName<0)return {sent:0,failed:0,skipped:0,details:['Kolom email atau nama kontingen tidak ditemukan.']};
 var unique={},recipients=[];
 values.slice(1).forEach(function(row){
  var name=String(row[iName]||'').trim(),email=String(row[iEmail]||'').trim().toLowerCase();
  if(!name||!email||!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))return;
  if(!unique[email]){unique[email]={email:email,names:[],leaders:[]};recipients.push(unique[email]);}
  if(unique[email].names.indexOf(name)<0)unique[email].names.push(name);
  var leader=iLeader>=0?String(row[iLeader]||'').trim():'';if(leader&&unique[email].leaders.indexOf(leader)<0)unique[email].leaders.push(leader);
 });
 var allRows=values.slice(1).filter(function(row){return String(row[iName]||'').trim();}).length;
 var validRows=values.slice(1).filter(function(row){var email=String(row[iEmail]||'').trim();return !!String(row[iName]||'').trim()&&/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);}).length;
 result.skipped=Math.max(0,allRows-validRows)+(validRows-recipients.length);
 var finalId=String(state.finalizationId||('legacy-'+String(state.finalizedAt||'').replace(/[^0-9]/g,'')));
 var log=ss.getSheetByName('Log Email Hasil Final')||ss.insertSheet('Log Email Hasil Final');
 var logHeaders=['Waktu','ID Finalisasi','Petugas','Nama Kontingen','Email','Status Pengiriman','Keterangan'];
 if(log.getLastRow()===0){log.appendRow(logHeaders);log.setFrozenRows(1);log.getRange(1,1,1,logHeaders.length).setFontWeight('bold');}
 else{var lh=log.getRange(1,1,1,log.getLastColumn()).getValues()[0];logHeaders.forEach(function(h){if(lh.indexOf(h)<0){log.getRange(1,log.getLastColumn()+1).setValue(h).setFontWeight('bold');lh.push(h);}});}
 var lvalues=log.getDataRange().getValues(),lh2=lvalues[0],idIx=lh2.indexOf('ID Finalisasi'),emailIx=lh2.indexOf('Email'),statusIx=lh2.indexOf('Status Pengiriman'),sentMap={};
 if(idIx>=0&&emailIx>=0&&statusIx>=0){lvalues.slice(1).forEach(function(row){if(String(row[idIx]||'')===finalId&&String(row[statusIx]||'')==='Terkirim')sentMap[String(row[emailIx]||'').trim().toLowerCase()]=true;});}
 // Mencegah email terkirim dua kali jika tombol finalisasi terklik ulang, termasuk log versi sebelumnya.
 if(finalId.indexOf('legacy-')===0){
  var legacy=ss.getSheetByName('Log Email Hasil');if(legacy&&legacy.getLastRow()>=2){var lv=legacy.getDataRange().getValues(),hh=lv[0],li=hh.indexOf('Email'),ls=hh.indexOf('Status Pengiriman'),lt=hh.indexOf('Waktu'),cut=null;var m=String(state.finalizedAt||'').match(/^(\d{2})-(\d{2})-(\d{4})\s+(\d{2}):(\d{2}):(\d{2})$/);if(m)cut=new Date(Number(m[3]),Number(m[2])-1,Number(m[1]),Number(m[4]),Number(m[5]),Number(m[6]));
   if(li>=0&&ls>=0)lv.slice(1).forEach(function(row){var when=lt>=0?new Date(row[lt]):null;if(String(row[ls]||'')==='Terkirim'&&String(row[li]||'').trim()&&(!cut||!when||when>=cut))sentMap[String(row[li]).trim().toLowerCase()]=true;});
  }
 }
 var pending=[];recipients.forEach(function(rec){if(sentMap[rec.email])result.skipped++;else pending.push(rec);});
 if(!recipients.length){result.details.push('Tidak ada alamat email valid di data kontingen.');return result;}
 if(!pending.length){result.details.push('Semua alamat email valid sudah menerima PDF untuk finalisasi ini.');return result;}
 var pdf;
 try{var stored=ensureChampionshipPdfFile_(state);pdf=DriveApp.getFileById(stored.fileId).getBlob().setName('Surat Keputusan Hasil Kejuaraan PORSENI 2026.pdf');}catch(err){result.failed=pending.length;result.details.push('PDF gagal dibuat: '+String(err&&err.message?err.message:err));pending.forEach(function(rec){log.appendRow([new Date(),finalId,state.finalizedBy||'Admin',rec.names.join(', '),rec.email,'Gagal','PDF gagal dibuat']);});return result;}
 pending.forEach(function(rec){
  var status='Terkirim',note='';
  try{
   if(MailApp.getRemainingDailyQuota()<1)throw new Error('Kuota email harian telah habis. Gunakan tombol kirim ulang setelah kuota tersedia.');
   sendPorseniEmail_({
    to:rec.email,
    subject:'Penyampaian Hasil Resmi Juara Umum dan Rekapitulasi Seluruh Cabang Lomba PORSENI 2026',
    body:'Assalamu’alaikum Wr. Wb.\\n\\n'+
      (rec.leaders&&rec.leaders.length
        ?'Yth. Bapak/Ibu '+rec.leaders.join(' dan ')+'\\nSelaku Ketua Kontingen '+rec.names.join(' dan ')+'.'
        :'Yth. Bapak/Ibu Ketua Kontingen '+rec.names.join(' dan ')+'.')+
      '\\n\\nDengan hormat,\\n\\n'+
      'Sehubungan dengan telah dilaksanakannya seluruh rangkaian kegiatan Pekan Olahraga dan Seni (PORSENI) 2026 PAC IPNU-IPPNU Patebon, serta telah selesainya proses rekapitulasi dan finalisasi hasil penilaian seluruh cabang perlombaan, bersama email ini kami sampaikan hasil resmi PORSENI 2026.\\n\\n'+
      'Adapun dokumen yang kami lampirkan memuat informasi sebagai berikut:\\n'+
      '1. Hasil resmi penetapan Juara Umum PORSENI 2026.\\n'+
      '2. Rekapitulasi perolehan juara masing-masing kontingen.\\n'+
      '3. Hasil akhir seluruh cabang perlombaan yang telah melalui proses finalisasi.\\n'+
      '4. Daftar perolehan prestasi kontingen pada masing-masing kategori lomba.\\n\\n'+
      'Dokumen tersebut disampaikan dalam bentuk PDF sebagai arsip dan bahan informasi resmi bagi masing-masing kontingen. Kami berharap hasil yang telah ditetapkan dapat diterima dengan baik dan menjadi bentuk apresiasi atas seluruh usaha, semangat, sportivitas, serta partisipasi para peserta dalam menyukseskan kegiatan PORSENI 2026.\\n\\n'+
      'Kami segenap Panitia PORSENI 2026 PAC IPNU-IPPNU Patebon mengucapkan terima kasih yang sebesar-besarnya kepada seluruh pengurus, pendamping, official, dan peserta kontingen atas partisipasi, dukungan, serta kerja sama yang telah diberikan sejak awal hingga berakhirnya seluruh rangkaian kegiatan.\\n\\n'+
      'Selamat kepada seluruh kontingen yang telah berhasil meraih prestasi. Bagi kontingen yang belum memperoleh hasil sesuai harapan, semoga pengalaman dan perjuangan dalam PORSENI kali ini menjadi motivasi untuk terus berkembang, meningkatkan kemampuan, serta meraih prestasi yang lebih baik pada kesempatan berikutnya.\\n\\n'+
      'Demikian informasi ini kami sampaikan. Atas perhatian dan kerja sama Bapak/Ibu, kami mengucapkan terima kasih.\\n\\n'+
      'Wassalamu’alaikum Wr. Wb.\\n\\n'+
      'Hormat kami,\\n\\n'+
      'PANITIA PORSENI 2026\\n'+
      'PAC IPNU-IPPNU PATEBON\\n\\n'+
      '“Unitas Ad Gloriam”\\n'+
      'Bersatu untuk Meraih Kejayaan.',
    attachments:[pdf],
    name:SENDER_NAME,
    replyTo:REPLY_TO
   });
   result.sent++;
  }catch(err){status='Gagal';note=String(err&&err.message?err.message:err);result.failed++;result.details.push(rec.names.join(', ')+': '+note);}
  log.appendRow([new Date(),finalId,state.finalizedBy||'Admin',rec.names.join(', '),rec.email,status,note]);
 });
 return result;
}
function adminResendFinalResults_(d){
 var admin=requireAdmin_(d.token),state=getChampionshipFinalization_();if(!state.finalized)throw new Error('Juara umum belum difinalisasi. Finalisasi terlebih dahulu sebelum mengirim PDF resmi.');
 state.finalizedBy=state.finalizedBy||admin.user;
 var dispatch=emailFinalResultsToContingents_(state,true);
 adminAudit_(admin,'KIRIM ULANG HASIL FINAL','hasil','',state.finalizationId,'Email ulang: '+dispatch.sent+' terkirim, '+dispatch.failed+' gagal, '+dispatch.skipped+' dilewati.');
 return {ok:true,message:'Pengiriman ulang PDF hasil resmi selesai diproses.',emailDispatch:dispatch,finalization:state};
}

function adminTestChampionshipPdf_(d){
 var admin=requireAdmin_(d.token),jobId=Utilities.getUuid(),folder=testFolder_();
 cleanupTestPdfStatus_();
 var blob=buildChampionshipPdfBlob_({finalized:false,finalizedAt:'',finalizedBy:admin.user||'Panitia'},true);
 var file=folder.createFile(blob.setName('UJI_HASIL_KEJUARAAN_PORSENI_2026_'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd_HHmmss')+'.pdf'));
 var status={at:Date.now(),state:'ready',fileId:file.getId(),previewUrl:'https://drive.google.com/file/d/'+file.getId()+'/preview',downloadUrl:'https://drive.google.com/uc?export=download&id='+file.getId(),name:file.getName(),message:'PDF uji hasil kejuaraan siap dipratinjau.'};
 PropertiesService.getScriptProperties().setProperty('PORSENI_TEST_PDF_STATUS_'+jobId,JSON.stringify(status));
 adminAudit_(admin,'UJI CETAK HASIL KEJUARAAN','hasil','',jobId,'PDF draft dibuat di folder UJI COBA; tanpa finalisasi dan tanpa pengiriman email.');
 return {ok:true,jobId:jobId,state:'ready',message:'PDF uji hasil kejuaraan berhasil dibuat. Dokumen ini belum resmi.'};
}

function adminFinalizeChampionship_(d){
 var admin=requireAdmin_(d.token),finalized=d.finalized===true||String(d.finalized)==='true',props=PropertiesService.getScriptProperties(),state,wasAlreadyFinalized=false;
 if(finalized){
  var lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
   var existing=getChampionshipFinalization_();
   if(existing.finalized){state=existing;wasAlreadyFinalized=true;}
   else{
    var res=getChampionshipResults_(),byKey={};res.detail.forEach(function(x){byKey[normalizeEvent_(x.cabang)+'|'+normalizeEvent_(x.kategori)]=x;});
    var missing=[];Object.keys(KAT).forEach(function(branch){(KAT[branch]||[]).forEach(function(cat){var x=byKey[normalizeEvent_(branch)+'|'+normalizeEvent_(cat)];if(!x||!x.juara1)missing.push(branch+' — '+cat);});});
    if(missing.length)throw new Error('Finalisasi belum dapat dilakukan. Juara 1 belum lengkap pada '+missing.length+' dari 24 kategori: '+missing.join('; '));
    var now=new Date();state={finalized:true,finalizedAt:Utilities.formatDate(now,Session.getScriptTimeZone(),'dd-MM-yyyy HH:mm:ss'),finalizedBy:admin.user||'Admin',finalizationId:Utilities.getUuid()};
    props.setProperty(CHAMPIONSHIP_FINALIZATION_PROPERTY,JSON.stringify(state));
    adminAudit_(admin,'FINALISASI JUARA UMUM','hasil','',state.finalizationId,'Juara umum dan hasil seluruh kategori ditetapkan final.');
   }
  }finally{lock.releaseLock();}
  var publicPdf;try{publicPdf=ensureChampionshipPdfFile_(state);state.pdf=publicPdf;props.setProperty(CHAMPIONSHIP_FINALIZATION_PROPERTY,JSON.stringify(state));}catch(pdfErr){throw new Error('Finalisasi tersimpan, tetapi PDF resmi gagal disimpan: '+String(pdfErr.message||pdfErr));}
  var emailDispatch=emailFinalResultsToContingents_(state,false);
  return {ok:true,message:wasAlreadyFinalized?'Juara umum sudah berstatus final. Email yang belum berhasil akan dicoba kembali; email yang sudah terkirim tidak diulang.':'Juara umum PORSENI 2026 berhasil difinalisasi dan pengiriman PDF ke email kontingen telah diproses.',finalization:state,results:getChampionshipResults_(),emailDispatch:emailDispatch};
 }
 var lock2=LockService.getScriptLock();lock2.waitLock(20000);
 try{var old=getChampionshipFinalization_();retireChampionshipPdf_();var reopened={finalized:false,finalizedAt:'',finalizedBy:'',finalizationId:'',lastReopenedAt:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd-MM-yyyy HH:mm:ss'),lastReopenedBy:admin.user||'Admin'};props.setProperty(CHAMPIONSHIP_FINALIZATION_PROPERTY,JSON.stringify(reopened));adminAudit_(admin,'BATALKAN FINALISASI JUARA UMUM','hasil','',old.finalizationId,'Finalisasi dibuka kembali oleh '+(admin.user||'Admin')+'.');return {ok:true,message:'Finalisasi dibatalkan. Hasil website kembali berstatus sementara.',finalization:reopened,previousFinalization:old,results:getChampionshipResults_()};}
 finally{lock2.releaseLock();}
}
function adminSaveResult_(d){
 var admin=requireAdmin_(d.token),branch=String(d.cabang||'').trim(),category=String(d.kategori||'').trim();
 if(!branch||!category)throw new Error('Cabang lomba dan kategori wajib dipilih.');if(!KAT[branch]||KAT[branch].indexOf(category)<0)throw new Error('Kombinasi cabang dan kategori tidak terdaftar.');
 var winners=[d.juara1,d.juara2,d.juara3].map(function(x){return String(x||'').trim().slice(0,180);});if(!winners.some(Boolean))throw new Error('Isi minimal satu nama kontingen juara.');
 winners=winners.map(function(name){if(!name)return '';var official=KONTINGEN.filter(function(n){return normalizeEvent_(n)===normalizeEvent_(name);})[0];if(!official)throw new Error('Nama kontingen tidak ditemukan dalam daftar resmi: '+name);return official;});
 var filled=winners.filter(Boolean);if(filled.some(function(name,i){return filled.indexOf(name)!==i;}))throw new Error('Satu kontingen tidak boleh mengisi lebih dari satu posisi juara pada kategori yang sama.');
 var lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  var sh=ensureResultSheet_(),values=sh.getDataRange().getValues(),h=values[0],bi=h.indexOf('Cabang Lomba'),ki=h.indexOf('Kategori'),row=-1;
  for(var i=1;i<values.length;i++){if(normalizeEvent_(values[i][bi])===normalizeEvent_(branch)&&normalizeEvent_(values[i][ki])===normalizeEvent_(category)){row=i+1;break;}}
  var now=new Date(),record=[branch,category,winners[0],winners[1],winners[2],now,admin.user||'Admin'];
  if(row<0)sh.appendRow(record);else{var map={};RESULT_HEADERS.forEach(function(key,j){map[key]=record[j];});var newRow=h.map(function(key){return map[key]!==undefined?map[key]:values[row-1][h.indexOf(key)];});sh.getRange(row,1,1,newRow.length).setValues([newRow]);}
  adminAudit_(admin,'SIMPAN HASIL KEJUARAAN','hasil',row,branch+' / '+category,'Juara 1: '+winners[0]+'; Juara 2: '+winners[1]+'; Juara 3: '+winners[2]);
  invalidateChampionshipFinalization_(admin,'Hasil '+branch+' / '+category+' diubah. Perlu finalisasi ulang.');
  return {ok:true,message:'Hasil kejuaraan berhasil disimpan.',results:getChampionshipResults_(),finalization:getChampionshipFinalization_()};
 }finally{lock.releaseLock();}
}
function adminDeleteResult_(d){
 var admin=requireAdmin_(d.token),branch=String(d.cabang||''),category=String(d.kategori||'');
 var lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESULT_SHEET);if(!sh||sh.getLastRow()<2)throw new Error('Belum ada data hasil kejuaraan.');
  var values=sh.getDataRange().getValues(),h=values[0],bi=h.indexOf('Cabang Lomba'),ki=h.indexOf('Kategori'),deleted=false;
  for(var i=values.length-1;i>=1;i--){if(normalizeEvent_(values[i][bi])===normalizeEvent_(branch)&&normalizeEvent_(values[i][ki])===normalizeEvent_(category)){adminAudit_(admin,'HAPUS HASIL','hasil',i+1,branch+' / '+category,'Menghapus hasil juara cabang/kategori.');sh.deleteRow(i+1);deleted=true;break;}}
  if(!deleted)throw new Error('Hasil yang dipilih tidak ditemukan.');
  invalidateChampionshipFinalization_(admin,'Hasil '+branch+' / '+category+' dihapus. Perlu finalisasi ulang.');
  return {ok:true,message:'Hasil kejuaraan dihapus.',results:getChampionshipResults_(),finalization:getChampionshipFinalization_()};
 }finally{lock.releaseLock();}
}

/* ===== RESET HASIL KEJUARAAN — KHUSUS ADMIN TYO ===== */
function adminResetChampionshipResults_(d){
 var admin=requireResetAdminTyo_(d.token);
 var scope=String(d.scope||''),branch=String(d.cabang||'').trim();
 var expected=scope==='all'?'RESET SELURUH HASIL KEJUARAAN PORSENI 2026':'RESET HASIL '+branch.toUpperCase();
 if(scope!=='all'&&scope!=='branch')throw new Error('Lingkup reset hasil tidak valid.');
 if(scope==='branch'&&(!branch||!KAT[branch]))throw new Error('Cabang lomba tidak valid.');
 if(String(d.confirmText||'')!==expected)throw new Error('Konfirmasi tidak cocok. Reset hasil dibatalkan.');
 var lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss&&ss.getSheetByName(RESULT_SHEET);
  if(!sh||sh.getLastRow()<2)throw new Error('Belum ada hasil kejuaraan untuk direset.');
  var values=sh.getDataRange().getValues(),headers=values[0],bi=headers.indexOf('Cabang Lomba');
  if(bi<0)throw new Error('Kolom Cabang Lomba tidak ditemukan. Reset dibatalkan.');
  var rowsToDelete=[];
  for(var i=1;i<values.length;i++){
   if(scope==='all'||normalizeEvent_(values[i][bi])===normalizeEvent_(branch))rowsToDelete.push(i+1);
  }
  if(!rowsToDelete.length)throw new Error(scope==='all'?'Tidak ada hasil kejuaraan.':'Belum ada hasil pada cabang '+branch+'.');
  rowsToDelete.reverse().forEach(function(row){sh.deleteRow(row);});
  // Hapus cache publik agar index.html langsung membaca hasil terbaru.
  invalidatePublicCache_('recap');
  // Hasil berubah: nonaktifkan PDF final dan buka kembali status finalisasi.
  retireChampionshipPdf_();
  var props=PropertiesService.getScriptProperties();
  props.setProperty(CHAMPIONSHIP_FINALIZATION_PROPERTY,JSON.stringify({
   finalized:false,finalizedAt:'',finalizedBy:'',
   lastReopenedAt:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd-MM-yyyy HH:mm:ss'),
   lastReopenedBy:admin.user||'Admin',
   reason:scope==='all'?'Seluruh hasil kejuaraan direset.':'Hasil cabang '+branch+' direset.'
  }));
  var label=scope==='all'?'SELURUH HASIL KEJUARAAN':'HASIL CABANG '+branch;
  adminAudit_(admin,'RESET HASIL KEJUARAAN', 'hasil','',label,'Menghapus '+rowsToDelete.length+' baris hasil kejuaraan.');
  logSecurityEvent_(admin.user,'RESET HASIL KEJUARAAN','BERHASIL',label+'; baris dihapus: '+rowsToDelete.length);
  return {ok:true,message:scope==='all'?'Seluruh hasil kejuaraan berhasil direset.':'Hasil kejuaraan cabang '+branch+' berhasil direset.',deleted:rowsToDelete.length,results:getChampionshipResults_(),finalization:getChampionshipFinalization_()};
 }catch(err){
  logSecurityEvent_(admin.user,'RESET HASIL KEJUARAAN','GAGAL',String(err&&err.message||err));
  throw err;
 }finally{lock.releaseLock();}
}

function adminSendEmail_(d){
 var admin=requireAdmin_(d.token),type=String(d.type||''),status=String(d.status||''),note=String(d.note||'').trim().slice(0,1000);
 if(type!=='kontingen'&&type!=='cabang')throw new Error('Jenis data tidak valid.');
 if(ADMIN_STATUS_VALUES.indexOf(status)<0)throw new Error('Status verifikasi tidak valid.');
 if((status==='Perlu Revisi'||status==='Ditolak')&&!note)throw new Error('Catatan wajib diisi untuk status Perlu Revisi atau Ditolak.');
 var ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss&&ss.getSheetByName(CFG[type].sheet);
 if(!sheet)throw new Error('Data pendaftaran tidak ditemukan. Muat ulang daftar.');
 var lock=LockService.getScriptLock(),rowNumber,values,headers,row;lock.waitLock(20000);
 try{rowNumber=resolveAdminRecordRow_(sheet,d);values=sheet.getDataRange().getValues();headers=values[0];row=values[rowNumber-1];}finally{lock.releaseLock();}
 var idx={};headers.forEach(function(h,i){idx[String(h)]=i;});
 function val(key){return idx[key]===undefined?'':row[idx[key]];}
 var nama=String(val('Nama Kontingen')||'').trim(),reg=String(val('No. Registrasi')||'').trim(),email='';
 if(!nama)throw new Error('Nama kontingen pada baris ini tidak ditemukan.');
 if(type==='kontingen')email=String(val('Email')||'').trim();
 else{
  var parent=ss.getSheetByName(CFG.kontingen.sheet);
  if(parent&&parent.getLastRow()>=2){var pv=parent.getDataRange().getValues(),ph=pv[0],ni=ph.indexOf('Nama Kontingen'),ei=ph.indexOf('Email');
   if(ni>=0&&ei>=0){for(var i=1;i<pv.length;i++){if(String(pv[i][ni]||'').trim().toLowerCase()===nama.toLowerCase()&&String(pv[i][ei]||'').trim()){email=String(pv[i][ei]).trim();break;}}}
  }
 }
 if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))throw new Error('Alamat email penerima tidak tersedia atau tidak valid pada data pendaftaran kontingen.');
 var cabang=String(val('Cabang Lomba')||'').trim(),kategori=String(val('Kategori')||'').trim();
 var jenis=type==='kontingen'?'pendaftaran kontingen':'berkas cabang lomba'+(cabang?' '+cabang:'')+(kategori?' ('+kategori+')':'');
 var subject='[PORSENI 2026] '+status+' - '+nama+(type==='cabang'?(cabang?' - '+cabang:'')+(kategori?' ('+kategori+')':''):'')+(reg?' ('+reg+')':'');
 var body='';
 if(type==='kontingen'&&status==='Terverifikasi'){
  body='Assalamu’alaikum Wr. Wb.\n\nYth. Ketua Kontingen\n'+nama+'\n\nDengan hormat,\n\nPanitia Pekan Olahraga dan Seni (PORSENI) 2026 PAC IPNU-IPPNU Kecamatan Patebon menyampaikan pemberitahuan terkait status verifikasi berkas pendaftaran kontingen yang telah Saudara kirimkan melalui sistem pendaftaran resmi PORSENI 2026.\n\nSTATUS VERIFIKASI: TERVERIFIKASI\n\nAlhamdulillah, berdasarkan hasil pemeriksaan dan verifikasi yang telah dilakukan oleh panitia, berkas pendaftaran kontingen '+nama+' dengan nomor registrasi '+(reg||'-')+' dinyatakan TERVERIFIKASI.\n\nDengan demikian, proses verifikasi administrasi pendaftaran kontingen telah selesai dan dinyatakan memenuhi persyaratan administrasi yang ditetapkan oleh panitia.\n\nTINDAK LANJUT: PENDAFTARAN CABANG LOMBA\n\nSehubungan dengan telah terverifikasinya pendaftaran kontingen, kami mengimbau kepada Ketua Kontingen untuk segera melanjutkan ke tahap pendaftaran cabang lomba dengan memperhatikan ketentuan berikut:\n\n1. Segera mendaftarkan peserta pada cabang lomba yang akan diikuti.\n2. Memastikan kesesuaian data peserta dengan persyaratan masing-masing cabang lomba.\n3. Melengkapi dan mengunggah seluruh berkas persyaratan administrasi peserta.\n4. Memastikan seluruh proses pendaftaran dan pengumpulan berkas telah selesai sebelum batas waktu yang ditentukan.\n\nJADWAL BATAS AKHIR PENDAFTARAN DAN PENGUMPULAN BERKAS\n\nCabang/Kategori — Batas Akhir\nDuta Pelajar — 6 Desember 2026, pukul 23.59 WIB\nVoli — Sesuai jadwal pada data resmi cabang lomba, pukul 23.59 WIB\nBadminton — Sesuai jadwal pada data resmi cabang lomba, pukul 23.59 WIB\nCabang lomba lainnya — 19 Desember 2026, pukul 23.59 WIB\n\n⚠️ PERHATIAN PENTING\n\nKami mengingatkan seluruh Ketua Kontingen agar tidak menunda proses pendaftaran dan pengumpulan berkas hingga mendekati batas akhir.\n\nKhusus cabang lomba Duta Pelajar, Voli, dan Badminton, harap memperhatikan jadwal batas akhir masing-masing sebagaimana tercantum dalam informasi resmi cabang lomba.\n\nKeterlambatan pendaftaran maupun pengumpulan berkas dapat memengaruhi proses administrasi dan keikutsertaan peserta dalam perlombaan. Oleh karena itu, seluruh kontingen diharapkan dapat menyelesaikan setiap tahapan sesuai jadwal yang telah ditetapkan.\n\nKami mengucapkan terima kasih atas perhatian, kerja sama, dan partisipasi seluruh pengurus serta peserta kontingen dalam menyukseskan PORSENI 2026.\n\nDemikian pemberitahuan ini kami sampaikan. Tetap pantau informasi dan pengumuman resmi dari panitia agar tidak tertinggal informasi penting terkait pelaksanaan kegiatan.\n\nWassalamu’alaikum Wr. Wb.\n\nHormat kami,\n\nPanitia PORSENI 2026\nPAC IPNU-IPPNU Kecamatan Patebon\n“Unitas Ad Gloriam”';
 }else if(type==='kontingen'&&status==='Perlu Revisi'){
  body='Assalamu’alaikum Wr. Wb.\n\nYth. Ketua Kontingen '+nama+',\n\nDengan hormat,\n\nPanitia Pekan Olahraga dan Seni (PORSENI) 2026 PAC IPNU-IPPNU Kecamatan Patebon telah melakukan pemeriksaan terhadap berkas pendaftaran kontingen yang Saudara kirimkan melalui sistem pendaftaran resmi PORSENI 2026.\n\nSTATUS VERIFIKASI: PERLU REVISI\n\nBerdasarkan hasil pemeriksaan, terdapat beberapa bagian pada data dan/atau berkas pendaftaran kontingen '+nama+' yang perlu diperbaiki atau dilengkapi sebelum dapat diproses ke tahap verifikasi selanjutnya.\n\nCATATAN DAN PERBAIKAN YANG HARUS DILAKUKAN\n\n'+note+'\n\nTINDAK LANJUT\n\n1. Periksa dengan saksama seluruh catatan dan arahan perbaikan yang disampaikan oleh panitia.\n2. Lengkapi atau perbaiki data maupun dokumen sesuai dengan catatan tersebut.\n3. Kirimkan kembali pembaruan data dan/atau berkas melalui sistem pendaftaran resmi PORSENI 2026 sesuai mekanisme yang tersedia.\n4. Pastikan seluruh informasi dan dokumen yang diperbarui telah benar, jelas, dan sesuai dengan persyaratan pendaftaran.\n\nKami mengimbau agar perbaikan segera dilakukan sehingga proses verifikasi dapat dilanjutkan tanpa menghambat tahapan pendaftaran berikutnya. Status Perlu Revisi bukan merupakan penolakan akhir; berkas akan dapat diperiksa kembali setelah perbaikan disampaikan sesuai ketentuan panitia.\n\nTerima kasih atas perhatian dan kerja sama Saudara.\n\nWassalamu’alaikum Wr. Wb.\n\nHormat kami,\n\nPanitia PORSENI 2026\nPAC IPNU-IPPNU Kecamatan Patebon\n“Unitas Ad Gloriam”\n\nEmail ini dikirim melalui Portal Panitia PORSENI 2026.';
 }else if(type==='kontingen'&&status==='Ditolak'){
  body='Assalamu’alaikum Wr. Wb.\n\nYth. Ketua Kontingen '+nama+',\n\nDengan hormat,\n\nPanitia Pekan Olahraga dan Seni (PORSENI) 2026 PAC IPNU-IPPNU Kecamatan Patebon telah melakukan pemeriksaan terhadap berkas pendaftaran kontingen yang Saudara kirimkan melalui sistem pendaftaran resmi PORSENI 2026.\n\nSTATUS VERIFIKASI: DITOLAK\n\nBerdasarkan hasil pemeriksaan dan pertimbangan administrasi panitia, dengan berat hati kami sampaikan bahwa pendaftaran kontingen '+nama+' dengan nomor registrasi '+(reg||'-')+' dinyatakan DITOLAK dan belum dapat dilanjutkan ke tahap berikutnya.\n\nALASAN PENOLAKAN\n\n'+note+'\n\nKEJELASAN DAN TINDAK LANJUT\n\nKeputusan tersebut ditetapkan berdasarkan hasil pemeriksaan terhadap data dan/atau dokumen pendaftaran serta ketentuan yang berlaku dalam pelaksanaan PORSENI 2026. Mohon Ketua Kontingen memperhatikan alasan penolakan yang telah disampaikan di atas.\n\nApabila ketentuan dan kesempatan pendaftaran ulang masih memungkinkan, silakan menghubungi panitia atau sekretariat PORSENI 2026 untuk memperoleh informasi lebih lanjut mengenai mekanisme yang dapat ditempuh. Pendaftaran ulang hanya dapat dilakukan sesuai persetujuan dan ketentuan resmi panitia.\n\nKami mengucapkan terima kasih atas perhatian, partisipasi, dan kerja sama Saudara dalam kegiatan PORSENI 2026.\n\nWassalamu’alaikum Wr. Wb.\n\nHormat kami,\n\nPanitia PORSENI 2026\nPAC IPNU-IPPNU Kecamatan Patebon\n“Unitas Ad Gloriam”\n\nEmail ini dikirim melalui Portal Panitia PORSENI 2026.';
 }else if(type==='cabang'&&status==='Terverifikasi'){
  body='Assalamu’alaikum Wr. Wb.\n\nYth. Ketua Kontingen '+nama+',\n\nDengan hormat,\n\nPanitia Pekan Olahraga dan Seni (PORSENI) 2026 PAC IPNU-IPPNU Kecamatan Patebon telah melakukan pemeriksaan terhadap berkas cabang lomba yang telah Saudara kirimkan.\n\nSTATUS VERIFIKASI: TERVERIFIKASI\n\nBerdasarkan hasil pemeriksaan, berkas cabang lomba berikut dinyatakan TERVERIFIKASI dan telah memenuhi persyaratan administrasi yang ditetapkan oleh panitia.\n\nRINCIAN BERKAS\n\nNama Kontingen: '+nama+'\nNomor Registrasi: '+(reg||'-')+'\nCabang Lomba: '+(cabang||'-')+'\nKategori: '+(kategori||'-')+'\n\nDengan demikian, berkas tersebut dapat dilanjutkan ke tahapan berikutnya sesuai ketentuan pelaksanaan PORSENI 2026.\n\nMohon Ketua Kontingen tetap memantau informasi dan pengumuman resmi dari panitia terkait jadwal serta teknis pelaksanaan perlombaan.\n\nTerima kasih atas perhatian, kerja sama, dan partisipasi Saudara.\n\nWassalamu’alaikum Wr. Wb.\n\nHormat kami,\n\nPanitia PORSENI 2026\nPAC IPNU-IPPNU Kecamatan Patebon\n“Unitas Ad Gloriam”';
 }else if(type==='cabang'&&status==='Perlu Revisi'){
  body='Assalamu’alaikum Wr. Wb.\n\nYth. Ketua Kontingen '+nama+',\n\nDengan hormat,\n\nPanitia Pekan Olahraga dan Seni (PORSENI) 2026 PAC IPNU-IPPNU Kecamatan Patebon telah melakukan pemeriksaan terhadap berkas cabang lomba yang Saudara kirimkan.\n\nSTATUS VERIFIKASI: PERLU REVISI\n\nBerdasarkan hasil pemeriksaan, terdapat bagian pada dokumen persyaratan yang perlu diperbaiki atau dilengkapi sebelum dapat dinyatakan memenuhi persyaratan administrasi.\n\nRINCIAN BERKAS\n\nNama Kontingen: '+nama+'\nNomor Registrasi: '+(reg||'-')+'\nCabang Lomba: '+(cabang||'-')+'\nKategori: '+(kategori||'-')+'\n\nCATATAN DAN PERBAIKAN YANG HARUS DILAKUKAN\n\n'+note+'\n\nTINDAK LANJUT\n\n1. Periksa seluruh catatan perbaikan dari panitia.\n2. Lakukan perbaikan atau lengkapi dokumen sesuai arahan.\n3. Unggah kembali dokumen yang telah diperbaiki melalui Portal Kontingen PORSENI 2026.\n4. Pastikan dokumen yang dikirimkan sudah benar dan sesuai ketentuan.\n\nKami mengimbau agar perbaikan segera dilakukan sehingga proses verifikasi dapat dilanjutkan.\n\nTerima kasih atas perhatian dan kerja sama Saudara.\n\nWassalamu’alaikum Wr. Wb.\n\nHormat kami,\n\nPanitia PORSENI 2026\nPAC IPNU-IPPNU Kecamatan Patebon\n“Unitas Ad Gloriam”';
 }else if(type==='cabang'&&status==='Ditolak'){
  body='Assalamu’alaikum Wr. Wb.\n\nYth. Ketua Kontingen '+nama+',\n\nDengan hormat,\n\nPanitia Pekan Olahraga dan Seni (PORSENI) 2026 PAC IPNU-IPPNU Kecamatan Patebon telah melakukan pemeriksaan terhadap berkas cabang lomba yang Saudara kirimkan.\n\nSTATUS VERIFIKASI: DITOLAK\n\nBerdasarkan hasil pemeriksaan dan pertimbangan administrasi panitia, berkas cabang lomba berikut dinyatakan DITOLAK dan tidak dapat dilanjutkan ke tahap berikutnya.\n\nRINCIAN BERKAS\n\nNama Kontingen: '+nama+'\nNomor Registrasi: '+(reg||'-')+'\nCabang Lomba: '+(cabang||'-')+'\nKategori: '+(kategori||'-')+'\n\nALASAN PENOLAKAN\n\n'+note+'\n\nKeputusan tersebut ditetapkan berdasarkan hasil pemeriksaan terhadap dokumen dan ketentuan yang berlaku dalam pelaksanaan PORSENI 2026.\n\nApabila terdapat hal yang perlu dikonfirmasi terkait keputusan tersebut, Ketua Kontingen dapat menghubungi panitia melalui kanal komunikasi resmi PORSENI 2026.\n\nKami mengucapkan terima kasih atas perhatian, partisipasi, dan kerja sama Saudara.\n\nWassalamu’alaikum Wr. Wb.\n\nHormat kami,\n\nPanitia PORSENI 2026\nPAC IPNU-IPPNU Kecamatan Patebon\n“Unitas Ad Gloriam”';
 }else if(type==='cabang'&&status==='Menunggu Verifikasi'){
  throw new Error('Status Menunggu Verifikasi tidak mengirim email keputusan.');
 
  body='Assalamu’alaikum Wr. Wb.\n\nYth. Ketua Kontingen '+nama+',\n\nKami dari Panitia PORSENI 2026 PAC IPNU-IPPNU Kecamatan Patebon menyampaikan informasi mengenai '+jenis+'.\n\nSTATUS VERIFIKASI: '+status.toUpperCase()+'.';
  if(status==='Perlu Revisi')body+='\n\nMohon melakukan perbaikan berkas sesuai catatan panitia berikut:';
  else if(status==='Ditolak')body+='\n\nDengan berat hati kami informasikan bahwa berkas pendaftaran belum dapat diterima. Berikut keterangan dari panitia:';
  else body+='\n\nBerkas saat ini masih dalam antrean pemeriksaan. Panitia akan menyampaikan informasi selanjutnya.';
  if(note)body+='\n\nCatatan Panitia:\n'+note;
  body+='\n\nTerima kasih atas perhatian dan kerja samanya.\n\nSalam,\nPanitia PORSENI 2026\nPAC IPNU-IPPNU Kecamatan Patebon\n“Unitas Ad Gloriam”\n\nEmail ini dikirim melalui Portal Panitia PORSENI 2026.';
 }
 if(MailApp.getRemainingDailyQuota()<1)throw new Error('Kuota pengiriman email harian akun Google Apps Script telah habis.');
 var mailOptions={to:email,subject:subject,body:body,name:SENDER_NAME,replyTo:REPLY_TO};
  if(type==='kontingen'&&status==='Terverifikasi'){
   var pdfId=getDriveFileIdFromUrl_(String(val('PDF Formulir (link Drive)')||''));
   if(!pdfId)throw new Error('PDF formulir kontingen belum tersedia. Bangun ulang PDF terlebih dahulu.');
   var pdfFile=DriveApp.getFileById(pdfId),hasPdfAttachment=false;
   parsePortalFiles_(String(val('Berkas (link Drive)')||'')).forEach(function(f){
    if(/^pdf formulir$/i.test(f.label))return;
    var id=getDriveFileIdFromUrl_(f.url);if(!id)return;
    try{if(DriveApp.getFileById(id).getMimeType()===MimeType.PDF)hasPdfAttachment=true;}catch(ignore){}
   });
   if(hasPdfAttachment&&!wasPdfMerged_(pdfFile))throw new Error('PDF lengkap masih dalam proses penggabungan. Tunggu sampai PDF final siap, lalu simpan verifikasi kembali untuk mengirim email beserta lampiran PDF penuh.');
   mailOptions.attachments=[pdfFile.getBlob().setName(pdfFile.getName())];
   mailOptions.body+='\n\nTerlampir PDF formulir pendaftaran kontingen lengkap beserta lampiran berkas yang telah digabungkan.';
  }
  sendPorseniEmail_(mailOptions);
 try{
  var log=ss.getSheetByName('Log Email')||ss.insertSheet('Log Email');
  if(log.getLastRow()===0){log.appendRow(['Waktu','Admin','Jenis Data','Baris Sheet','Nomor Registrasi','Nama Kontingen','Email Tujuan','Status Verifikasi','Subjek','Hasil']);log.setFrozenRows(1);log.getRange(1,1,1,10).setFontWeight('bold');}
  log.appendRow([new Date(),admin.user||'Admin',type,rowNumber,reg,nama,email,status,subject,'Terkirim']);
 }catch(logErr){console.error('Email terkirim, tetapi pencatatan Log Email gagal: '+logErr);}
 return {ok:true,message:'Email berhasil dikirim.',recipient:email,subject:subject,sentAt:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd-MM-yyyy HH:mm:ss')};
}
/** Jalankan satu kali dari editor Apps Script untuk memberikan otorisasi pengiriman email. */
function authorizeEmailService(){
  var aliases = GmailApp.getAliases();
  var effectiveEmail = '';
  try { effectiveEmail = Session.getEffectiveUser().getEmail(); } catch (e) {}
  Logger.log('Akun eksekusi: ' + effectiveEmail);
  Logger.log('Alamat pengirim tujuan: ' + SENDER_EMAIL);
  Logger.log('Alias Gmail terverifikasi: ' + aliases.join(', '));
  Logger.log('Kuota email tersisa: ' + MailApp.getRemainingDailyQuota());
  return {
    effectiveEmail: effectiveEmail,
    senderEmail: SENDER_EMAIL,
    senderReady: String(effectiveEmail || '').toLowerCase() === SENDER_EMAIL.toLowerCase() ||
      aliases.map(function(x){return String(x).toLowerCase();}).indexOf(SENDER_EMAIL.toLowerCase()) >= 0,
    remainingQuota: MailApp.getRemainingDailyQuota()
  };
}

function registrationControls_(){
 var p=PropertiesService.getScriptProperties();
 return {kontingen:p.getProperty('REG_OPEN_KONTINGEN')==='true',cabang:p.getProperty('REG_OPEN_CABANG')==='true',maintenanceCabang:p.getProperty('REG_MAINTENANCE_CABANG')==='true',updatedAt:p.getProperty('REG_OPEN_UPDATED_AT')||'',updatedBy:p.getProperty('REG_OPEN_UPDATED_BY')||''};
}
function adminGetRegistrationControls_(token){requireAdmin_(token);return {ok:true,registration:registrationControls_()};}
function adminSetRegistrationControls_(d){
 var admin=requireAdmin_(d.token),type=String(d.type||''),open=d.open===true;
 if(type!=='kontingen'&&type!=='cabang'&&type!=='maintenance_cabang')throw new Error('Jenis kontrol pendaftaran tidak valid.');
 var p=PropertiesService.getScriptProperties(),key=type==='kontingen'?'REG_OPEN_KONTINGEN':(type==='cabang'?'REG_OPEN_CABANG':'REG_MAINTENANCE_CABANG');
 p.setProperty(key,open?'true':'false');p.setProperty('REG_OPEN_UPDATED_AT',Utilities.formatDate(new Date(),'Asia/Jakarta','dd-MM-yyyy HH:mm:ss'));p.setProperty('REG_OPEN_UPDATED_BY',admin.user||'Admin');
 logSecurityEvent_(admin.user,'KONTROL PENDAFTARAN '+type.toUpperCase(),open?'DIAKTIFKAN':'DINONAKTIFKAN',type==='maintenance_cabang'?'Mode perbaikan cabang lomba.':(open?'Override tanggal aktif; tanggal otomatis diabaikan saat dibuka.':'Kembali mengikuti tanggal otomatis.'));
 var label=type==='kontingen'?'kontingen':(type==='cabang'?'cabang lomba':'mode perbaikan cabang lomba');
 return {ok:true,registration:registrationControls_(),message:type==='maintenance_cabang'?(open?'Mode perbaikan aktif. Pengiriman berkas cabang lomba dihentikan sementara.':'Mode perbaikan dinonaktifkan. Status pendaftaran cabang kembali mengikuti pengaturan pembukaan dan tanggal.'):'Pendaftaran '+label+(open?' berhasil dibuka. Tanggal otomatis diabaikan.':' berhasil ditutup.')};
}
function registrationIsOpen_(type){var c=registrationControls_();return type==='kontingen'?c.kontingen:c.cabang;}

function validate(d){
 ['nama'].forEach(function(k){if(!String(d[k]||'').trim())throw new Error('Nama kontingen wajib diisi.');});
 if(!canonicalContingentName_(d.nama)||KONTINGEN.indexOf(canonicalContingentName_(d.nama))<0)throw new Error('Nama kontingen tidak ada dalam daftar resmi. Silakan pilih dari dropdown.');
 if(d.type==='kontingen'){if(!registrationIsOpen_('kontingen'))assertKontingenWindow_();['ketua','tempat_lahir_ketua','tanggal_lahir_ketua','alamat','wa_ketua','email','ketua_pr_ipnu','wa_ketua_ipnu','ketua_pr_ippnu','wa_ketua_ippnu'].forEach(function(k){d[k]=cleanKontingenField_(k,d[k]);});}
 if(d.type==='kontingen'){d.email=String(d.email||'').trim().toLowerCase();if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(d.email))throw new Error('Alamat email kontingen wajib diisi dengan format yang valid. Email ini digunakan untuk seluruh notifikasi pendaftaran dan cabang lomba.');}
 if(d.type==='cabang'){
  if(registrationControls_().maintenanceCabang)throw new Error('Pendaftaran cabang lomba sedang ditutup sementara karena pemeliharaan/perbaikan sistem. Silakan coba kembali setelah panitia membuka pendaftaran.');
  var canonicalName=canonicalContingentName_(d.nama);
  if(d.sessionToken){var branchSession=requireContingentSession_(d.sessionToken);if(normalizeEvent_(branchSession.nama)!==normalizeEvent_(canonicalName))throw new Error('Sesi portal tidak sesuai dengan nama kontingen.');}else{validateContingentPin_(canonicalName,d.pin);}
  var parentApproval=getContingentApproval_(canonicalName);
  if(parentApproval.status!=='Terverifikasi')throw new Error('Pengumpulan berkas cabang hanya dapat dilakukan setelah pendaftaran kontingen berstatus Terverifikasi. Status saat ini: '+parentApproval.status+'.');
 }
 if(d.type==='cabang'){
  if(!KAT[d.cabang])throw new Error('Cabang lomba tidak terdaftar.');
  if(!registrationIsOpen_('cabang')){var now=new Date(),deadline=d.cabang==='Badminton'?new Date(2026,10,29,23,59,59):(d.cabang==='Duta Pelajar'||d.cabang==='Voli'?new Date(2026,11,6,23,59,59):new Date(2026,11,19,23,59,59));
  if(now<new Date(2026,10,1,0,0,0))throw new Error('Pengumpulan berkas cabang lomba dibuka mulai 1 November 2026.');
  if(now>deadline)throw new Error('Pengumpulan berkas '+d.cabang+' sudah ditutup pada '+(d.cabang==='Badminton'?'29 November 2026':(d.cabang==='Duta Pelajar'||d.cabang==='Voli'?'6 Desember 2026':'19 Desember 2026'))+'.');}
  if(KAT[d.cabang].indexOf(d.kategori)<0)throw new Error('Kategori tidak sesuai dengan cabang lomba.');
  var registeredEvents=getRegisteredEvents_(String(d.nama||'').trim());
  var registeredBranch=registeredEvents.filter(function(ev){return normalizeEvent_(ev.cabang)===normalizeEvent_(d.cabang);})[0];
  if(!registeredBranch||registeredBranch.categories.indexOf(d.kategori)<0)
   throw new Error('Pendaftaran ditolak: cabang/kategori ini tidak tercatat pada pendaftaran kontingen.');
  // Kategori yang sudah terverifikasi dikunci. Kategori yang masih menunggu,
  // perlu revisi, atau ditolak boleh dikirim ulang sebagai pembaruan data peserta.
  var existingBranchRow=findLatestBranchRecord_(canonicalName,d.cabang,d.kategori);
  if(existingBranchRow&&existingBranchRow.status==='Terverifikasi')
   throw new Error('Berkas cabang/kategori ini sudah TERVERIFIKASI dan terkunci. Hubungi Sekretariat dan kirimkan berkas terbaru secara manual melalui WhatsApp.');
  if(!Array.isArray(d.pesertaList)||!d.pesertaList.length)throw new Error('Daftar peserta wajib diisi.');
  validateParticipantRules_(d.cabang,d.pesertaList);
  d.pesertaList.forEach(function(p,i){
   if(String(p.nama||'').trim().length<2)throw new Error('Nama peserta '+(i+1)+' tidak valid.');
   if(String(d.cabang||'')==='Duta Pelajar'){
    if(!String(p.wa||'').trim())throw new Error('No. WhatsApp peserta Duta Pelajar '+(i+1)+' wajib diisi.');
    var ph=normPhone_(p.wa);if(!/^08\d{8,12}$/.test(ph))throw new Error('No. WA peserta '+(i+1)+' tidak valid. Gunakan format 08xxxxxxxxxx.');p.wa=ph;
   }else p.wa='';
  });
  d.jml=d.pesertaList.length; d.peserta=d.pesertaList.map(function(p,i){return (i+1)+'. '+p.nama+' ('+p.gender+')'+(p.role?' - '+p.role:'')+(p.keterangan?' - '+p.keterangan:'');}).join('\n');
 }
 var files=d.files||[];if(!Array.isArray(files)||files.length>MAX_FILES)throw new Error('Jumlah lampiran tidak valid (maksimal '+MAX_FILES+').');
 if(d.type==='kontingen'){
  ['bukti pembayaran','sp / berita acara rat','ktp ketua kontingen','pakta integritas'].forEach(function(requiredLabel){
   if(!files.some(function(f){return String(f.label||'').trim().toLowerCase()===requiredLabel;}))throw new Error('Pendaftaran ditolak: '+requiredLabel.toUpperCase()+' wajib diunggah.');
  });
 }
 if(d.type==='cabang'&&!files.length)throw new Error('Berkas cabang wajib dilampirkan.');
  if(d.type==='cabang'){
   for(var pi=1;pi<=(d.pesertaList||[]).length;pi++){
    if(!files.some(function(f){return String(f.label||'').trim().toLowerCase()==='pas foto peserta '+pi;}))throw new Error('Pas foto peserta '+pi+' wajib dilampirkan.');
    if(!files.some(function(f){return String(f.label||'').trim().toLowerCase()==='ktp peserta '+pi;}))throw new Error('KTP peserta '+pi+' wajib dilampirkan dalam format gambar JPG/PNG.');
   }
  }
  
 if(d.type==='cabang'&&!files.some(function(f){return String(f.label||'').toLowerCase()==='surat tugas dari pr ipnu-ippnu';}))throw new Error('Surat Tugas dari PR IPNU-IPPNU wajib dilampirkan.');
 if(d.type==='cabang'&&d.cabang==='Duta Pelajar'){['cv duta pelajar','berkas pendukung duta pelajar'].forEach(function(lbl){if(!files.some(function(f){return String(f.label||'').trim().toLowerCase()===lbl;}))throw new Error(lbl.toUpperCase()+' wajib dilampirkan.');});}
 if(d.type==='cabang'&&d.cabang==='Karya ilmiah'){
  ['makalah kti (word)','makalah kti (pdf)','presentasi kti (ppt)'].forEach(function(lbl){
   if(!files.some(function(f){return String(f.label||'').trim().toLowerCase()===lbl;}))throw new Error(lbl.toUpperCase()+' wajib diunggah.');
  });
 }
 if(d.type==='cabang'&&d.cabang==='Khitobah'&&!files.some(function(f){return String(f.label||'').trim().toLowerCase()==='naskah khitobah (pdf)';}))throw new Error('Naskah khitobah PDF wajib dilampirkan.');
 var totalBytes=0;files.forEach(function(f){
  if(!f||typeof f.data!=='string'||!f.data)throw new Error('Data lampiran tidak lengkap.');
  var bytes=Utilities.base64Decode(f.data),size=bytes.length,totalBytesBefore=totalBytes;totalBytes+=size;
  if(totalBytes>MAX_TOTAL_BYTES)throw new Error('Total seluruh lampiran maksimal 25 MB.');
  var label=String(f.label||'').toLowerCase(),name=String(f.name||'').toLowerCase();
  var isPdf=bytes.length>=5&&bytes[0]===37&&bytes[1]===80&&bytes[2]===68&&bytes[3]===70&&bytes[4]===45;
  var isOffice=bytes.length>=4&&bytes[0]===80&&bytes[1]===75&&bytes[2]===3&&bytes[3]===4;
  // Gambar, KTP, pas foto, tanda tangan, dan CV Duta disimpan byte asli tanpa konversi/cek signature.
  if(/tanda tangan/.test(label)){assertFileKind_(bytes,['png','jpg'],f.label);return;}
  if(/pas foto/.test(label)){assertFileKind_(bytes,['png','jpg'],f.label);return;}
  if(label==='ktp ketua kontingen'){assertFileKind_(bytes,['png','jpg'],f.label);return;}
   // KTP setiap peserta cabang wajib berupa gambar JPG/PNG.
  // Normalisasi spasi dan kapitalisasi agar label dari frontend tetap dikenali.
  var normalizedAttachmentLabel=String(f.label||'').trim().toLowerCase();
  if(/^ktp peserta \d+$/.test(normalizedAttachmentLabel)){
    assertFileKind_(bytes,['png','jpg'],f.label||'KTP peserta');
    return;
  }
  if(/cv duta pelajar|berkas pendukung duta pelajar|bukti pembayaran/.test(label)){assertFileKind_(bytes,['pdf','png','jpg'],f.label);return;}
  if(label==='makalah kti (word)'){if(!isOffice&&!/\.doc$/i.test(name))throw new Error('Makalah KTI wajib berupa Word.');return;}
  if(label==='makalah kti (pdf)'||label==='naskah khitobah (pdf)'){if(!isPdf)throw new Error('Dokumen wajib berupa PDF.');return;}
  if(label==='presentasi kti (ppt)'){if(!isOffice&&!/\.pptx?$/i.test(name))throw new Error('File presentasi KTI wajib berupa PowerPoint.');return;}
  if(!isPdf)throw new Error('Lampiran '+(f.name||'')+' perlu berupa PDF.');
 });
 // Data identitas adalah input manual. Disimpan apa adanya; verifikasi isi dilakukan panitia di Portal Admin.
 // Tanggal lahir, tempat lahir, alamat, dan usia tidak diproses/ditolak otomatis.
 if(d.type==='kontingen'){
  d.tempat_lahir_ketua=String(d.tempat_lahir_ketua==null?'':d.tempat_lahir_ketua);
  d.tanggal_lahir_ketua=String(d.tanggal_lahir_ketua==null?'':d.tanggal_lahir_ketua);
  d.alamat=String(d.alamat==null?'':d.alamat);
  d.birthDateKetua=d.tanggal_lahir_ketua;
 }
 if(d.type==='cabang'){
  d.tempat_lahir_peserta=(d.pesertaList||[]).map(function(p){return String(p.tempatLahir==null?'':p.tempatLahir);}).join('\n');
  d.tanggal_lahir_peserta=(d.pesertaList||[]).map(function(p,i){return (i+1)+'. '+String(p.nama||'')+': '+String(p.tanggalLahir==null?'':p.tanggalLahir);}).join('\n');
  d.alamat_peserta=(d.pesertaList||[]).map(function(p,i){return (i+1)+'. '+String(p.nama||'')+': '+String(p.alamat==null?'':p.alamat);}).join('\n');
  d.wa_peserta=(d.pesertaList||[]).map(function(p,i){return (i+1)+'. '+String(p.nama||'')+': '+String(p.wa==null?'':p.wa);}).join('\n');
  d.birthDatesPeserta=(d.pesertaList||[]).map(function(p){return String(p.tanggalLahir==null?'':p.tanggalLahir);});
  d.addressesPeserta=(d.pesertaList||[]).map(function(p){return String(p.alamat==null?'':p.alamat);});
  d.peserta=(d.pesertaList||[]).map(function(p,i){return (i+1)+'. '+String(p.nama||'')+' ('+String(p.gender||'')+')'+(p.role?' - '+p.role:'')+' | Tempat lahir: '+String(p.tempatLahir==null?'':p.tempatLahir)+' | Tanggal lahir: '+String(p.tanggalLahir==null?'':p.tanggalLahir)+' | Alamat: '+String(p.alamat==null?'':p.alamat)+' | WA: '+String(p.wa==null?'':p.wa);}).join('\n');
 }
 /* end identity manual */
 if(totalBytes>MAX_TOTAL_BYTES)throw new Error('Total seluruh lampiran maksimal 25 MB.');
 // Tidak ada pemeriksaan otomatis tanggal lahir, usia, tempat lahir, atau alamat. Data disimpan untuk verifikasi manual.

}

/**
 * PIN dibuat otomatis ketika pendaftaran kontingen berhasil.
 * Fungsi ini hanya dipakai untuk memastikan PIN unik per kontingen.
 */
function ensurePinSheet_(){
 var ss=SpreadsheetApp.getActiveSpreadsheet();if(!ss)throw new Error('Spreadsheet aktif tidak ditemukan.');
 var sh=ss.getSheetByName(PIN_SHEET)||ss.insertSheet(PIN_SHEET);
 if(sh.getLastRow()===0){sh.appendRow(['Nama Kontingen','PIN','Dibuat']);sh.setFrozenRows(1);}
 var headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
 ['PIN Salt','PIN Hash'].forEach(function(h){if(headers.indexOf(h)<0){sh.getRange(1,sh.getLastColumn()+1).setValue(h).setFontWeight('bold');headers.push(h);}});
 return {sheet:sh,headers:headers};
}
function getPinPepper_(){
 var props=PropertiesService.getScriptProperties(),key='PORSENI_PIN_PEPPER',value=props.getProperty(key);
 if(!value){value=Utilities.getUuid()+Utilities.getUuid()+Utilities.getUuid();props.setProperty(key,value);}
 return value;
}
function hashContingentPin_(pin,salt){
 var signature=Utilities.computeHmacSha256Signature(String(salt)+'|'+String(pin),getPinPepper_());
 return Utilities.base64EncodeWebSafe(signature);
}
function pinHashEquals_(a,b){a=String(a||'');b=String(b||'');if(a.length!==b.length)return false;var diff=0;for(var i=0;i<a.length;i++)diff|=a.charCodeAt(i)^b.charCodeAt(i);return diff===0;}
function findPinRow_(name){
 var info=ensurePinSheet_(),sh=info.sheet,h=info.headers,values=sh.getDataRange().getValues(),idx={};h.forEach(function(x,i){idx[String(x)]=i;});
 for(var i=1;i<values.length;i++)if(normalizeEvent_(values[i][idx['Nama Kontingen']])===normalizeEvent_(name))return {sheet:sh,row:i+1,values:values[i],idx:idx,headers:h};
 return {sheet:sh,row:0,values:null,idx:idx,headers:h};
}
function createOrGetContingentPin_(name){
 var lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  var found=findPinRow_(name),sh=found.sheet,idx=found.idx;
  var pin=createUniquePin_({}),salt=Utilities.getUuid()+Utilities.getUuid(),hash=hashContingentPin_(pin,salt),now=new Date();
  if(found.row){
   sh.getRange(found.row,idx['PIN']+1).setValue('');
   sh.getRange(found.row,idx['Dibuat']+1).setValue(now);
   sh.getRange(found.row,idx['PIN Salt']+1).setValue(salt);
   sh.getRange(found.row,idx['PIN Hash']+1).setValue(hash);
  }else{
   var row=new Array(found.headers.length).fill('');row[idx['Nama Kontingen']]=name;row[idx['PIN']]='';row[idx['Dibuat']]=now;row[idx['PIN Salt']]=salt;row[idx['PIN Hash']]=hash;sh.appendRow(row);
  }
  return pin;
 }finally{lock.releaseLock();}
}
function createUniquePin_(existing){
 var pin;do{var hex=Utilities.getUuid().replace(/-/g,'').slice(0,12),numeric=parseInt(hex,16)%1000000;pin=('000000'+numeric).slice(-PIN_LENGTH);}while(existing&&Object.keys(existing).some(function(k){return existing[k]===pin;}));return pin;
}
function migrateLegacyContingentPins_(){
 var lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  var info=ensurePinSheet_(),sh=info.sheet,h=info.headers,values=sh.getDataRange().getValues(),iPin=h.indexOf('PIN'),iSalt=h.indexOf('PIN Salt'),iHash=h.indexOf('PIN Hash'),migrated=0,already=0,empty=0;
  for(var r=1;r<values.length;r++){
   var plain=String(values[r][iPin]||''),hash=String(values[r][iHash]||'');
   if(hash){already++;continue;}if(!/^\d{6}$/.test(plain)){empty++;continue;}
   var salt=Utilities.getUuid()+Utilities.getUuid();sh.getRange(r+1,iSalt+1).setValue(salt);sh.getRange(r+1,iHash+1).setValue(hashContingentPin_(plain,salt));sh.getRange(r+1,iPin+1).clearContent();migrated++;
  }
  return {ok:true,migrated:migrated,alreadyHashed:already,skipped:empty,message:'Migrasi selesai. PIN disimpan sebagai HMAC hash dan kolom PIN teks dikosongkan.'};
 }finally{lock.releaseLock();}
}
function validateContingentPin_(name,pin){
 var guard=LockService.getScriptLock();
 guard.waitLock(15000);
 try{return validateContingentPinCore_(name,pin);}finally{guard.releaseLock();}
}
function validateContingentPinCore_(name,pin){
 var supplied=String(pin||'').replace(/\s+/g,''),state=getPinAttemptState_(name),now=Date.now();
 if(state.lockedUntil&&now<state.lockedUntil){var remain=Math.ceil((state.lockedUntil-now)/60000);throw new Error('Akses PIN dikunci sementara. Coba kembali sekitar '+remain+' menit lagi.');}
 var pinShapeValid=/^\d{6}$/.test(supplied);
 var record=findPinRow_(name);if(!record.row)throw new Error('PIN kontingen belum dibuat oleh panitia. Hubungi panitia.');
 var plain=String(record.values[record.idx['PIN']]||''),salt=String(record.values[record.idx['PIN Salt']]||''),stored=String(record.values[record.idx['PIN Hash']]||''),valid=false;
 if(pinShapeValid&&stored&&salt)valid=pinHashEquals_(hashContingentPin_(supplied,salt),stored);
 else if(pinShapeValid&&/^\d{6}$/.test(plain)&&pinHashEquals_(supplied,plain)){
  salt=Utilities.getUuid()+Utilities.getUuid();record.sheet.getRange(record.row,record.idx['PIN Salt']+1).setValue(salt);record.sheet.getRange(record.row,record.idx['PIN Hash']+1).setValue(hashContingentPin_(plain,salt));record.sheet.getRange(record.row,record.idx['PIN']+1).clearContent();valid=true;
 }
 if(!valid){
  var updated=registerPinFailure_(name,state);
  if(updated.count>=PIN_MAX_ATTEMPTS){updated.lockedUntil=Date.now()+PIN_LOCK_MS;PropertiesService.getScriptProperties().setProperty(pinAttemptKey_(name),JSON.stringify(updated));logSecurityEvent_(name,'PERLINDUNGAN PIN','DIKUNCI','10 percobaan salah; dikunci 10 menit.');
   try{if(MailApp.getRemainingDailyQuota()>0)sendPorseniEmail_({to:NOTIFY,subject:'[PORSENI 2026] PIN kontingen terkunci sementara',body:'Kontingen '+name+' mencapai 10 percobaan PIN yang salah. Akses dikunci selama 10 menit.',name:SENDER_NAME,replyTo:REPLY_TO});}catch(mailErr){console.warn(mailErr);}
   throw new Error('PIN salah 10 kali. Akses dikunci sementara selama 10 menit.');
  }
  throw new Error('PIN kontingen tidak cocok. Percobaan '+updated.count+' dari '+PIN_MAX_ATTEMPTS+'.');
 }
 clearPinFailures_(name);return true;
}
function pinAttemptKey_(name){var bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(name||''),Utilities.Charset.UTF_8);return 'pintry_'+Utilities.base64EncodeWebSafe(bytes).slice(0,32);}
function getPinAttemptState_(name){
 var raw=PropertiesService.getScriptProperties().getProperty(pinAttemptKey_(name));if(!raw)return {count:0,first:0,lockedUntil:0};
 try{var o=JSON.parse(raw),now=Date.now();if(o.lockedUntil&&now<Number(o.lockedUntil))return {count:Number(o.count)||PIN_MAX_ATTEMPTS,first:Number(o.first)||now,lockedUntil:Number(o.lockedUntil)};if(o.lockedUntil&&now>=Number(o.lockedUntil)){clearPinFailures_(name);return {count:0,first:0,lockedUntil:0};}if(!o.first||now-Number(o.first)>=PIN_WINDOW_MS){clearPinFailures_(name);return {count:0,first:0,lockedUntil:0};}return {count:Number(o.count)||0,first:Number(o.first)||now,lockedUntil:0};}catch(e){clearPinFailures_(name);return {count:0,first:0,lockedUntil:0};}
}
function registerPinFailure_(name,state){state=state||getPinAttemptState_(name);var now=Date.now();if(!state.first||now-state.first>=PIN_WINDOW_MS)state={count:0,first:now,lockedUntil:0};state.count++;PropertiesService.getScriptProperties().setProperty(pinAttemptKey_(name),JSON.stringify(state));return state;}
function clearPinFailures_(name){PropertiesService.getScriptProperties().deleteProperty(pinAttemptKey_(name));}

function normalizeEvent_(s){return String(s||'').replace(/[\u200B-\u200D\uFEFF]/g,'').replace(/\u00A0/g,' ').trim().toLowerCase().replace(/\s+/g,' ');}
function normalizeVerificationStatus_(s){return String(s||'').replace(/[\u200B-\u200D\uFEFF]/g,'').replace(/\u00A0/g,' ').trim().toLowerCase().replace(/\s+/g,' ');}
function isVerifiedStatus_(s){return normalizeVerificationStatus_(s)==='terverifikasi';}
function canonicalContingentName_(name){var n=normalizeEvent_(name);for(var i=0;i<KONTINGEN.length;i++){if(normalizeEvent_(KONTINGEN[i])===n)return KONTINGEN[i];}return String(name||'').trim();}
function assertContingentNotRegistered_(name){
 var ss=SpreadsheetApp.getActiveSpreadsheet();if(!ss)throw new Error('Database PORSENI tidak dapat diakses.');
 var sh=ss.getSheetByName(CFG.kontingen.sheet);if(!sh||sh.getLastRow()<2)return;
 var lastCol=sh.getLastColumn(),headers=sh.getRange(1,1,1,lastCol).getValues()[0],iNama=headers.indexOf('Nama Kontingen')+1;
 if(!iNama)throw new Error('Kolom Nama Kontingen tidak ditemukan pada sheet Kontingen.');
 var values=sh.getRange(2,iNama,sh.getLastRow()-1,1).getValues(),target=normalizeEvent_(name);
 for(var i=0;i<values.length;i++)if(normalizeEvent_(values[i][0])===target)throw new Error('Kontingen ini sudah terdaftar. Silakan pilih nama kontingen lain atau hubungi panitia jika pendaftaran sebelumnya bermasalah.');
}

function getRegisteredContingentNames_(){
 var cache=CacheService.getScriptCache(),key='registered_contingents_v1';
 try{var cached=cache.get(key);if(cached)return JSON.parse(cached);}catch(cacheErr){}
 var ss=SpreadsheetApp.getActiveSpreadsheet();if(!ss)throw new Error('Database PORSENI tidak dapat diakses.');
 var sh=ss.getSheetByName(CFG.kontingen.sheet);if(!sh||sh.getLastRow()<2)return [];
 var lastCol=sh.getLastColumn(),headers=sh.getRange(1,1,1,lastCol).getValues()[0],iNama=headers.indexOf('Nama Kontingen')+1;
 if(!iNama)throw new Error('Kolom nama kontingen tidak ditemukan.');
 var values=sh.getRange(2,iNama,sh.getLastRow()-1,1).getValues(),found={};
 for(var i=0;i<values.length;i++){var nm=normalizeEvent_(values[i][0]);if(!nm)continue;for(var j=0;j<KONTINGEN.length;j++){if(normalizeEvent_(KONTINGEN[j])===nm){found[KONTINGEN[j]]=true;break;}}}
 var result=KONTINGEN.filter(function(n){return !!found[n];});
 try{cache.put(key,JSON.stringify(result),60);}catch(cacheWriteErr){}
 return result;
}
function getAvailableContingentNames_(){
 var cache=CacheService.getScriptCache(),key='available_contingents_v1';
 try{var cached=cache.get(key);if(cached)return JSON.parse(cached);}catch(cacheErr){}
 var registered=getRegisteredContingentNames_(),taken={};
 registered.forEach(function(name){taken[normalizeEvent_(name)]=true;});
 var result=KONTINGEN.filter(function(name){return !taken[normalizeEvent_(name)];});
 try{cache.put(key,JSON.stringify(result),60);}catch(cacheWriteErr){}
 return result;
}
function getPublicRecap_(){
 var publicCache=CacheService.getScriptCache(),publicCacheKey='public_recap_v1';
 try{var publicCached=publicCache.get(publicCacheKey);if(publicCached)return JSON.parse(publicCached);}catch(publicCacheReadErr){}
 var ss=SpreadsheetApp.getActiveSpreadsheet();
 if(!ss)throw new Error('Database PORSENI tidak dapat diakses.');
 var sh=ss.getSheetByName(CFG.kontingen.sheet),sub=ss.getSheetByName(CFG.cabang.sheet);
 var mode=getPublicRecapMode_(),names=[],seenNames={},approvedNames={},groups={},registeredGroups={},statusCounts={};
 var diagnostic={backendVersion:'PORSENI-2026-V36-PUBLIC-RECAP-MODE-FIX',sheetKontingen:!!sh,sheetCabang:!!sub,totalBarisKontingen:sh?Math.max(0,sh.getLastRow()-1):0,totalBarisCabang:sub?Math.max(0,sub.getLastRow()-1):0};
 if(sh)ensureAdminColumns_(sh);
 if(sub)ensureAdminColumns_(sub);

 if(sh&&sh.getLastRow()>=2){
  var values=sh.getDataRange().getValues(),headers=values[0];
  var iNama=headers.indexOf('Nama Kontingen'),iStatus=headers.indexOf('Status Verifikasi'),iCabang=headers.indexOf('Cabang Lomba');
  if(iNama<0||iStatus<0||iCabang<0)throw new Error('Kolom Nama Kontingen, Status Verifikasi, atau Cabang Lomba tidak ditemukan pada sheet Kontingen.');
  for(var r=1;r<values.length;r++){
   var name=String(values[r][iNama]||'').trim(),status=String(values[r][iStatus]||'').trim();
   var statusKey=normalizeVerificationStatus_(status)||'(kosong)';
   statusCounts[statusKey]=(statusCounts[statusKey]||0)+1;
   if(!name)continue;
   var normName=normalizeEvent_(name),isApproved=isVerifiedStatus_(status);
   if(isApproved){
    approvedNames[normName]=true;
    if(!seenNames[normName]){seenNames[normName]=true;names.push(name);}
   }
   // The registration's selected branches/categories are the source for public branch recap.
   // In "terdaftar" mode include all registrants; in "terverifikasi" mode include only
   // contingents whose registration itself has been verified. Branch document status
   // remains a separate admin checklist and does not control this public inclusion.
   if(mode==='terverifikasi'&&!isApproved)continue;
   var registeredText=String(values[r][iCabang]||'').trim();
   if(!registeredText)continue;
   registeredText.split(';').forEach(function(part){
    part=String(part||'').trim();if(!part)return;
    var match=part.match(/^(.*?)\s*\((.*?)\)\s*$/);
    var branch=match?match[1].trim():part;
    var categories=match?match[2].split(',').map(function(x){return x.trim();}).filter(Boolean):(KAT[branch]||[]);
    if(!branch)return;
    var branchKey=normalizeEvent_(branch);
    if(!groups[branchKey])groups[branchKey]={cabang:branch,kategori:{}};
    categories.forEach(function(cat){
     var catKey=normalizeEvent_(cat);if(!catKey)return;
     if(!groups[branchKey].kategori[catKey])groups[branchKey].kategori[catKey]={nama:cat,kontingen:[],_seen:{},jumlahKontingen:0};
     var item=groups[branchKey].kategori[catKey];
     if(!item._seen[normName]){
      item._seen[normName]=true;
      item.kontingen.push({nama:name,berkasSudah:false,jumlahPeserta:0});
      item.jumlahKontingen++;
     }
    });
   });
  }
 }

 // Preserve verified-contingent names as an independent recap, regardless of branch mode.
 names.sort(function(a,b){return a.localeCompare(b,'id');});

 var cabang=Object.keys(groups).map(function(key){
  var g=groups[key],cats=Object.keys(g.kategori).map(function(ck){
   var x=g.kategori[ck];delete x._seen;
   x.kontingen.sort(function(a,b){return a.nama.localeCompare(b.nama,'id');});
   return x;
  });
  var categoryOrder=KAT[g.cabang]||[];
  cats.sort(function(a,b){var ai=categoryOrder.indexOf(a.nama),bi=categoryOrder.indexOf(b.nama);return (ai<0?999:ai)-(bi<0?999:bi)||a.nama.localeCompare(b.nama,'id');});
  return {cabang:g.cabang,kategori:cats};
 });
 var branchOrder=Object.keys(KAT);
 cabang.sort(function(a,b){var ai=branchOrder.indexOf(a.cabang),bi=branchOrder.indexOf(b.cabang);return (ai<0?999:ai)-(bi<0?999:bi)||a.cabang.localeCompare(b.cabang,'id');});

 var ringkasan=[],ringkasanPendaftaran=[];
 cabang.forEach(function(g){
  g.kategori.forEach(function(k){
   ringkasanPendaftaran.push({cabang:g.cabang,kategori:k.nama,jumlahKontingen:k.jumlahKontingen});
   ringkasan.push({cabang:g.cabang,kategori:k.nama,jumlahPeserta:k.jumlahKontingen});
  });
 });
 diagnostic.statusKontingen=statusCounts;
 diagnostic.kontingenTerverifikasi=names.length;
 diagnostic.jumlahCabangPublik=cabang.length;
 diagnostic.jumlahKategoriPublik=ringkasanPendaftaran.length;
 diagnostic.modeRekapPublik=mode;
 var finalState=getChampionshipFinalization_();finalState.pdf=getPublicChampionshipPdf_();
 var publicResult={kontingen:names,jumlahKontingen:names.length,cabang:cabang,ringkasan:ringkasan,ringkasanPendaftaran:ringkasanPendaftaran,modeRekapPublik:mode,diagnostik:diagnostic,hasil:getChampionshipResults_(),katalogLomba:getChampionshipCatalog_(),finalisasi:finalState};
 try{var publicPayload=JSON.stringify(publicResult);if(publicPayload.length<60000)publicCache.put(publicCacheKey,publicPayload,30);}catch(publicCacheWriteErr){}
 return publicResult;
}
function portalFileRef_(url,label){
 var id=getDriveFileIdFromUrl_(url);
 return id?{label:String(label||'Berkas'),id:id}:null;
}
function getContingentPortalData_(nama){
 var ss=SpreadsheetApp.getActiveSpreadsheet();if(!ss)throw new Error('Database PORSENI tidak dapat diakses.');
 var result={nama:nama,registrasi:null,berkasCabang:[],statusCabang:[]};
 function readRows(sheetName){var sh=ss.getSheetByName(sheetName);if(!sh||sh.getLastRow()<2)return [];var v=sh.getDataRange().getValues(),h=v[0];return v.slice(1).map(function(row){var o={};h.forEach(function(k,i){o[k]=row[i];});return o;});}
 function fileRefs(text){return parsePortalFiles_(text).map(function(x){return portalFileRef_(x.url,x.label);}).filter(Boolean);}
 var kont=readRows(CFG.kontingen.sheet).filter(function(r){return normalizeEvent_(r['Nama Kontingen'])===normalizeEvent_(nama);});
 if(!kont.length)throw new Error('Data pendaftaran kontingen belum ditemukan.');
 var kr=kont[kont.length-1];
 result.dataKontingen=readKontingenData_(nama);
 result.approval={status:String(kr['Status Verifikasi']||'Menunggu Verifikasi'),catatan:String(kr['Catatan Panitia']||''),waktu:kr['Diverifikasi Pada']?Utilities.formatDate(new Date(kr['Diverifikasi Pada']),Session.getScriptTimeZone(),'dd-MM-yyyy HH:mm'):''};
 // Portal menerima ID file individual saja. Folder dan PDF arsip internal tidak pernah dikirim.
 result.registrasi={nomor:String(kr['No. Registrasi']||''),waktu:kr['Waktu']?Utilities.formatDate(new Date(kr['Waktu']),Session.getScriptTimeZone(),'dd-MM-yyyy HH:mm'):'',cabang:String(kr['Cabang Lomba']||''),pdfId:getDriveFileIdFromUrl_(kr['PDF Formulir (link Drive)']||''),mergePending:isMergePending_(String(kr['No. Registrasi']||'')),berkas:fileRefs(kr['Berkas (link Drive)']||'')};
 var rows=readRows(CFG.cabang.sheet).filter(function(r){return normalizeEvent_(r['Nama Kontingen'])===normalizeEvent_(nama);});
 result.berkasCabang=rows.map(function(r){return {nomor:String(r['No. Registrasi']||''),waktu:r['Waktu']?Utilities.formatDate(new Date(r['Waktu']),Session.getScriptTimeZone(),'dd-MM-yyyy HH:mm'):'',cabang:String(r['Cabang Lomba']||''),kategori:String(r['Kategori']||''),jumlah:String(r['Jumlah Peserta']||''),pdfId:getDriveFileIdFromUrl_(r['PDF Formulir (link Drive)']||''),mergePending:isMergePending_(String(r['No. Registrasi']||'')),berkas:fileRefs(r['Berkas (link Drive)']||'')};});
 var events=getRegisteredEvents_(nama);
 result.statusCabang=events.map(function(ev){
  return {cabang:ev.cabang,batasAkhir:branchDeadline_(ev.cabang),kategori:ev.categories.map(function(cat){
   var found=null;for(var i=rows.length-1;i>=0;i--){if(normalizeEvent_(rows[i]['Cabang Lomba'])===normalizeEvent_(ev.cabang)&&normalizeEvent_(rows[i]['Kategori'])===normalizeEvent_(cat)){found=rows[i];break;}}
   var branchReg=found?String(found['No. Registrasi']||''):'';
   return {nama:cat,status:found?'Sudah dikirim':'Belum dikirim',statusVerifikasi:found?String(found['Status Verifikasi']||'Menunggu Verifikasi'):'',catatanPanitia:found?String(found['Catatan Panitia']||''):'',waktuKirim:found&&found['Waktu']?Utilities.formatDate(new Date(found['Waktu']),Session.getScriptTimeZone(),'dd-MM-yyyy HH:mm'):'',nomor:branchReg,pdfId:found?getDriveFileIdFromUrl_(found['PDF Formulir (link Drive)']||''):'',mergePending:found?isMergePending_(branchReg):false};
  })};
 });
 var logSheet=ss.getSheetByName(SECURITY_LOG_SHEET);result.riwayat=[];
 if(logSheet&&logSheet.getLastRow()>1){var logLast=logSheet.getLastRow(),logStart=Math.max(2,logLast-499),logRows=logSheet.getRange(logStart,1,logLast-logStart+1,5).getValues();result.riwayat=logRows.filter(function(x){return normalizeEvent_(x[1])===normalizeEvent_(nama);}).slice(-100).reverse().map(function(x){return {waktu:x[0]?Utilities.formatDate(new Date(x[0]),'Asia/Jakarta','dd-MM-yyyy HH:mm:ss'):'',aktivitas:String(x[2]||''),status:String(x[3]||''),detail:String(x[4]||'')};});}
 return result;
}

// Ambil daftar file yang benar-benar tercatat pada baris kontingen ini.
// Kolom Folder Drive dan Arsip PDF Panitia sengaja tidak pernah diperiksa.
function getContingentOwnedFileIds_(nama){
 var ss=SpreadsheetApp.getActiveSpreadsheet(),allowed={};
 [CFG.kontingen,CFG.cabang].forEach(function(cfg){
  var sh=ss.getSheetByName(cfg.sheet);if(!sh||sh.getLastRow()<2)return;
  var values=sh.getDataRange().getValues(),h=values[0],ni=h.indexOf('Nama Kontingen'),pi=h.indexOf('PDF Formulir (link Drive)'),bi=h.indexOf('Berkas (link Drive)');
  if(ni<0)return;
  for(var r=1;r<values.length;r++){
   if(normalizeEvent_(values[r][ni])!==normalizeEvent_(nama))continue;
   if(pi>=0){var pId=getDriveFileIdFromUrl_(values[r][pi]);if(pId)allowed[pId]=true;}
   if(bi>=0)parsePortalFiles_(values[r][bi]).forEach(function(x){var id=getDriveFileIdFromUrl_(x.url);if(id)allowed[id]=true;});
  }
 });
 return allowed;
}
function contingentPreviewFile_(d){
 var actor='Sesi tidak dikenal';
 try{
  var session=requireContingentSession_(d.sessionToken);actor=session.nama;
  var nama=session.nama,id=String(d.fileId||'').trim();
  if(!/^[-\w]{25,}$/.test(id))throw new Error('Identitas file tidak valid.');
  var allowed=getContingentOwnedFileIds_(nama);if(!allowed[id])throw new Error('Berkas tidak termasuk dokumen milik kontingen ini.');
  var file;try{file=DriveApp.getFileById(id);}catch(e){throw new Error('Berkas tidak ditemukan atau tidak dapat diakses.');}
  if(file.isTrashed())throw new Error('Berkas sudah tidak tersedia.');
  var access='';try{access=file.getSharingAccess();}catch(accessErr){throw new Error('Izin privasi berkas tidak dapat diperiksa.');}
  if(isBroadDriveAccess_(access))throw new Error('Berkas memiliki izin publik/domain. Hubungi panitia untuk pemeriksaan keamanan.');
  var parents=file.getParents(),parentCount=0;while(parents.hasNext()){assertPrivateFolderTree_(parents.next());parentCount++;}
  if(!parentCount)throw new Error('Lokasi penyimpanan berkas tidak dapat diverifikasi.');
  var blob=file.getBlob(),bytes=blob.getBytes(),mime=String(blob.getContentType()||file.getMimeType()||'application/octet-stream').toLowerCase();
  if(bytes.length>26*1024*1024)throw new Error('Ukuran PDF '+(bytes.length/1048576).toFixed(1)+' MB terlalu besar untuk pratinjau di portal (maksimal 26 MB). Hubungi panitia untuk mendapatkan salinannya.');
  logSecurityEvent_(nama,'PRATINJAU DOKUMEN','BERHASIL',file.getName());
  return {ok:true,file:{name:file.getName(),mimeType:mime,size:bytes.length,base64:Utilities.base64Encode(bytes)}};
 }catch(err){
  logSecurityEvent_(actor,'PRATINJAU DOKUMEN','GAGAL',String(err&&err.message?err.message:err));
  throw err;
 }
}

function ensureChangeRequestSheet_(){
 var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss.getSheetByName('Permohonan Perubahan')||ss.insertSheet('Permohonan Perubahan');
 var headers=['Waktu','No. Tiket','Nama Kontingen','No. Registrasi','Jenis Permohonan','Cabang / Kategori','Data / Berkas','Alasan','Status','Catatan Panitia','Diproses Oleh','Diproses Pada','Perubahan Diminta'];
 if(sh.getLastRow()===0){sh.appendRow(headers);sh.setFrozenRows(1);sh.getRange(1,1,1,headers.length).setFontWeight('bold');}
 else{var old=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];headers.forEach(function(h){if(old.indexOf(h)<0){sh.getRange(1,sh.getLastColumn()+1).setValue(h).setFontWeight('bold');old.push(h);}});}
 return sh;
}
function adminReadChangeRequests_(){
 var list=adminReadChangeRequestsRaw_(),cache={};
 list.forEach(function(x){if(x.perubahan&&['Menunggu Tindak Lanjut','Sedang Diproses'].indexOf(x.status)>=0){try{x.lama=cache[x.nama]||(cache[x.nama]=readKontingenData_(x.nama));}catch(e){}}});
 return list;
}
function adminReadChangeRequestsRaw_(){
 var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Permohonan Perubahan');if(!sh||sh.getLastRow()<2)return [];
 var v=sh.getDataRange().getValues(),h=v[0],idx={};h.forEach(function(x,i){idx[String(x)]=i;});
 function val(row,key){return idx[key]===undefined?'':row[idx[key]];}
 return v.slice(1).map(function(row,i){return {rowNumber:i+2,tiket:String(val(row,'No. Tiket')||''),nama:String(val(row,'Nama Kontingen')||''),reg:String(val(row,'No. Registrasi')||''),jenis:String(val(row,'Jenis Permohonan')||''),target:String(val(row,'Cabang / Kategori')||''),field:String(val(row,'Data / Berkas')||''),alasan:String(val(row,'Alasan')||''),status:String(val(row,'Status')||'Menunggu Tindak Lanjut'),catatan:String(val(row,'Catatan Panitia')||''),perubahan:String(val(row,'Perubahan Diminta')||''),waktu:val(row,'Waktu')?Utilities.formatDate(new Date(val(row,'Waktu')),Session.getScriptTimeZone(),'dd-MM-yyyy HH:mm'):''};}).reverse();
}
function contingentRequestChange_(d){
 var session=requireContingentSession_(d.sessionToken),name=session.nama;if(KONTINGEN.indexOf(name)<0)throw new Error('Nama kontingen tidak dikenal.');
 var jenis=String(d.jenis||'').trim(),target=String(d.target||'').trim().slice(0,200),alasan=String(d.alasan||'').trim().slice(0,1000);
 if(['Data kontingen','Berkas cabang lomba'].indexOf(jenis)<0)throw new Error('Jenis permohonan tidak valid.');
 if(jenis==='Data kontingen')throw new Error('Gunakan formulir perubahan data kontingen di Portal Kontingen.');
 if(!target||!alasan)throw new Error('Nama data/berkas dan alasan perubahan wajib diisi.');
 var stateKey='change_req_'+pinAttemptKey_(name),cache=CacheService.getScriptCache();
 if(cache.get(stateKey))throw new Error('Permohonan perubahan baru saja dikirim. Silakan tunggu 5 menit sebelum mengirim kembali.');
 var approval=getContingentApproval_(name);if(approval.status==='Belum Terdaftar')throw new Error('Data kontingen tidak ditemukan.');
 var splitTarget=target.split(' | '),eventTarget=splitTarget.length>1?splitTarget.shift():'',fieldTarget=splitTarget.join(' | ')||target;
 if(jenis==='Data kontingen'&&approval.status!=='Terverifikasi')throw new Error('Data kontingen belum terverifikasi. Perubahan dapat dilakukan langsung sebelum verifikasi panitia.');
 if(jenis==='Berkas cabang lomba'){
  if(!eventTarget)throw new Error('Pilih cabang/kategori yang sudah terverifikasi untuk mengajukan perubahan.');
  var eventParts=eventTarget.split(' — '),branchName=eventParts.shift(),categoryName=eventParts.join(' — '),existing=findLatestBranchRecord_(name,branchName,categoryName);
  if(!existing||existing.status!=='Terverifikasi')throw new Error('Permohonan hanya diperlukan untuk berkas cabang/kategori yang sudah terverifikasi. Berkas yang belum terverifikasi dapat diperbarui langsung melalui Portal Kontingen.');
 }
 var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ensureChangeRequestSheet_(),reg='';
 var parent=ss.getSheetByName(CFG.kontingen.sheet);if(parent&&parent.getLastRow()>=2){var pv=parent.getDataRange().getValues(),ph=pv[0],ni=ph.indexOf('Nama Kontingen'),ri=ph.indexOf('No. Registrasi');for(var i=pv.length-1;i>=1;i--){if(normalizeEvent_(pv[i][ni])===normalizeEvent_(name)){reg=ri>=0?String(pv[i][ri]||''):'';break;}}}
 var ticket='PRM-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd-HHmmss')+'-'+Utilities.getUuid().slice(0,4).toUpperCase();
 sh.appendRow([new Date(),ticket,name,reg,jenis,eventTarget.replace(/^[=+@\-]/,''),fieldTarget.replace(/^[=+@\-]/,''),alasan,'Menunggu Tindak Lanjut','','','']);
 cache.put(stateKey,'1',300);logSecurityEvent_(name,'PERMOHONAN PERUBAHAN','BERHASIL',ticket+' | '+jenis+' | '+target);
 try{if(MailApp.getRemainingDailyQuota()>0)sendPorseniEmail_({to:NOTIFY,subject:'[PORSENI 2026] Permohonan perubahan '+ticket,body:'Permohonan perubahan baru dari '+name+' ('+reg+').\nJenis: '+jenis+'\nTarget: '+target+'\nAlasan: '+alasan+'\n\nSilakan periksa sheet Permohonan Perubahan.',name:SENDER_NAME,replyTo:REPLY_TO});}catch(mailErr){console.warn('Permohonan disimpan, notifikasi email gagal: '+mailErr);}
 return {ok:true,ticket:ticket,message:'Permohonan perubahan telah diterima panitia.'};
}
function adminUpdateChangeRequest_(d){
 var admin=requireAdmin_(d.token),ticket=String(d.ticket||'').trim(),status=String(d.status||'').trim(),note=String(d.note||'').trim().slice(0,1000);
 if(!ticket||['Menunggu Tindak Lanjut','Sedang Diproses','Selesai','Tidak Dapat Dipenuhi'].indexOf(status)<0)throw new Error('Data tindak lanjut tidak valid.');
 var sh=ensureChangeRequestSheet_(),v=sh.getDataRange().getValues(),h=v[0],it=h.indexOf('No. Tiket'),is=h.indexOf('Status'),ic=h.indexOf('Catatan Panitia'),ia=h.indexOf('Diproses Oleh'),id=h.indexOf('Diproses Pada'),row=-1;
 for(var i=1;i<v.length;i++){if(String(v[i][it]||'')===ticket){row=i+1;break;}}
 if(row<0)throw new Error('Nomor tiket permohonan tidak ditemukan.');
 sh.getRange(row,is+1).setValue(status);sh.getRange(row,ic+1).setValue(note);sh.getRange(row,ia+1).setValue(admin.user);sh.getRange(row,id+1).setValue(new Date());
 adminAudit_(admin,'TINDAK LANJUT PERMOHONAN','permohonan',row,ticket,status+' | '+note);
 return {ok:true,message:'Status permohonan '+ticket+' diperbarui.'};
}

function parsePortalFiles_(text){return String(text||'').split(/\r?\n/).map(function(line){var i=line.indexOf(': ');if(i<0)return null;var label=line.slice(0,i).trim(),url=line.slice(i+2).trim();return /^https?:\/\//i.test(url)?{label:label,url:url}:null;}).filter(Boolean);}

function getContingentEmail_(nama){
 var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss&&ss.getSheetByName(CFG.kontingen.sheet);
 if(!sh||sh.getLastRow()<2)return '';
 var v=sh.getDataRange().getValues(),h=v[0],iN=h.indexOf('Nama Kontingen'),iE=h.indexOf('Email');
 if(iN<0||iE<0)return '';
 for(var i=v.length-1;i>=1;i--){if(normalizeEvent_(v[i][iN])===normalizeEvent_(nama)){var email=String(v[i][iE]||'').trim().toLowerCase();return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)?email:'';}}
 return '';
}
function assertPrivateIndividualFile_(file){
 // Tidak menambahkan Viewer email mana pun; preview hanya lewat endpoint PIN.
 if(!file)throw new Error('File pratinjau tidak ditemukan.');
 var access='';try{access=file.getSharingAccess();}catch(e){throw new Error('Izin privasi berkas tidak dapat diperiksa: '+e);}
 if(isBroadDriveAccess_(access))throw new Error('File masih memiliki akses berbagi publik/domain ('+access+'). Pastikan folder induk tetap Restricted/Private.');
 var parents=file.getParents(),count=0;while(parents.hasNext()){assertPrivateFolderTree_(parents.next());count++;}
 if(!count)throw new Error('Lokasi penyimpanan file tidak dapat diverifikasi.');
 return true;
}

function getContingentApproval_(nama){
 var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss&&ss.getSheetByName(CFG.kontingen.sheet);
 if(!sh||sh.getLastRow()<2)return {status:'Belum Terdaftar',catatan:'Pendaftaran kontingen belum ditemukan.',waktu:''};
 var values=sh.getDataRange().getValues(),headers=values[0],iNama=headers.indexOf('Nama Kontingen'),iStatus=headers.indexOf('Status Verifikasi'),iNote=headers.indexOf('Catatan Panitia'),iDate=headers.indexOf('Diverifikasi Pada');
 if(iNama<0)return {status:'Belum Terdaftar',catatan:'Data kontingen belum ditemukan.',waktu:''};
 for(var i=values.length-1;i>=1;i--){
  if(normalizeEvent_(values[i][iNama])!==normalizeEvent_(nama))continue;
  var status=iStatus<0?'Menunggu Verifikasi':String(values[i][iStatus]||'Menunggu Verifikasi');
  return {status:status,catatan:iNote<0?'':String(values[i][iNote]||''),waktu:iDate>=0&&values[i][iDate]?Utilities.formatDate(new Date(values[i][iDate]),Session.getScriptTimeZone(),'dd-MM-yyyy HH:mm'):''};
 }
 return {status:'Belum Terdaftar',catatan:'Pendaftaran kontingen belum ditemukan.',waktu:''};
}

function findLatestBranchRecord_(nama,cabang,kategori){
 var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss&&ss.getSheetByName(CFG.cabang.sheet);
 if(!sh||sh.getLastRow()<2)return null;
 var values=sh.getDataRange().getValues(),h=values[0],ni=h.indexOf('Nama Kontingen'),bi=h.indexOf('Cabang Lomba'),ki=h.indexOf('Kategori'),si=h.indexOf('Status Verifikasi');
 for(var i=values.length-1;i>=1;i--){
  if(normalizeEvent_(values[i][ni])===normalizeEvent_(nama)&&normalizeEvent_(values[i][bi])===normalizeEvent_(cabang)&&normalizeEvent_(values[i][ki])===normalizeEvent_(kategori))
   return {rowNumber:i+1,status:si>=0?String(values[i][si]||'Menunggu Verifikasi'):'Menunggu Verifikasi'};
 }
 return null;
}

function getRegisteredEvents_(nama){
 var ss=SpreadsheetApp.getActiveSpreadsheet();
 if(KONTINGEN.indexOf(nama)<0)throw new Error('Nama kontingen tidak ada dalam daftar resmi.');
 if(!ss)throw new Error('Database PORSENI tidak dapat diakses.');
 var sh=ss.getSheetByName(CFG.kontingen.sheet);
 if(!sh||sh.getLastRow()<2)throw new Error('Kontingen belum terdaftar di database.');
 var values=sh.getDataRange().getValues(),headers=values[0];
 var iNama=headers.indexOf('Nama Kontingen'),iCabang=headers.indexOf('Cabang Lomba');
 if(iNama<0||iCabang<0)throw new Error('Kolom nama/cabang pada database kontingen tidak ditemukan.');
 var registeredText='';
 for(var i=1;i<values.length;i++){
  if(normalizeEvent_(values[i][iNama])===normalizeEvent_(nama)){registeredText=String(values[i][iCabang]||'');break;}
 }
 if(!registeredText)throw new Error('Pendaftaran kontingen belum ditemukan atau belum memilih cabang lomba.');
 var map={};
 registeredText.split(';').forEach(function(part){
  part=String(part||'').trim();if(!part)return;
  var m=part.match(/^(.*?)\s*\((.*?)\)\s*$/),branch=m?m[1].trim():part;
  var cats=m?m[2].split(',').map(function(x){return x.trim();}).filter(Boolean):(KAT[branch]||[]);
  if(!KAT[branch])return;
  if(!map[branch])map[branch]={cabang:branch,categories:[],submitted:[]};
  cats.forEach(function(cat){if(KAT[branch].indexOf(cat)>=0&&map[branch].categories.indexOf(cat)<0)map[branch].categories.push(cat);});
 });
 var sub=ss.getSheetByName(CFG.cabang.sheet);
 if(sub&&sub.getLastRow()>=2){
  var rows=sub.getDataRange().getValues(),h=rows[0],iN=h.indexOf('Nama Kontingen'),iB=h.indexOf('Cabang Lomba'),iK=h.indexOf('Kategori'),iS=h.indexOf('Status Verifikasi');
  if(iN>=0&&iB>=0&&iK>=0){
   for(var r=1;r<rows.length;r++){
    if(normalizeEvent_(rows[r][iN])!==normalizeEvent_(nama))continue;
    var b=String(rows[r][iB]||''),k=String(rows[r][iK]||''),status=iS>=0?String(rows[r][iS]||'Menunggu Verifikasi'):'Menunggu Verifikasi';
    if(status==='Terverifikasi'&&map[b]&&map[b].categories.indexOf(k)>=0&&map[b].submitted.indexOf(k)<0)map[b].submitted.push(k);
   }
  }
 }
 return Object.keys(map).map(function(k){return map[k];});
}
function validateParticipantRules_(branch,list){
 var count=list.length,limit=1;
 if(branch==='Voli'){list.forEach(function(p){p.gender='L';});}
 list.forEach(function(p,i){if(p.gender!=='L'&&p.gender!=='P')throw new Error(branch==='PES PS3'?'Peserta PES PS3 wajib memilih jenis kelamin Putra atau Putri.':'Jenis kelamin peserta '+(i+1)+' wajib dipilih.');});
 if(branch==='PES PS3'&&count!==1)throw new Error('PES PS3 merupakan kategori tunggal: hanya 1 peserta per kontingen.');
 if(branch==='Voli')limit=12;
 else if(branch==='Rebana')limit=15;
 else if(branch==='LCC'||branch==='Karya ilmiah')limit=3;
 else if(branch==='Badminton')limit=2;
 else if(branch==='Paduan suara')limit=11;
 if(count>limit)throw new Error('Jumlah peserta '+branch+' melebihi batas maksimal '+limit+' orang.');
 if(branch==='Badminton'){
  if(count!==2)throw new Error('Badminton ganda campuran wajib tepat 2 orang: 1 putra dan 1 putri.');
  if(list.filter(function(p){return p.gender==='L';}).length!==1||list.filter(function(p){return p.gender==='P';}).length!==1)throw new Error('Badminton ganda campuran wajib 1 putra (L) dan 1 putri (P).');
 }else if(branch==='Voli'){
  if(list.some(function(p){return p.gender!=='L';}))throw new Error('Peserta voli wajib seluruhnya putra (L).');
 }else if(branch==='LCC'||branch==='Karya ilmiah'){
  if(!list.some(function(p){return p.gender==='L';})||!list.some(function(p){return p.gender==='P';}))throw new Error(branch+' wajib terdiri dari peserta putra dan putri.');
 }else if(branch==='Paduan suara'){
  if(count!==11)throw new Error('Paduan suara wajib tepat 11 orang: 1 dirigen dan 10 penyanyi.');
  var dirs=list.filter(function(p){return p.role==='Dirigen';}),singers=list.filter(function(p){return p.role==='Penyanyi';});
  if(dirs.length!==1||singers.length!==10)throw new Error('Paduan suara wajib terdiri dari tepat 1 dirigen dan 10 penyanyi.');
  if(!singers.some(function(p){return p.gender==='L';})||!singers.some(function(p){return p.gender==='P';}))throw new Error('Sepuluh penyanyi wajib campuran putra dan putri.');
 }
}

function newReg(){return 'P26-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd-HHmmss')+'-'+Utilities.getUuid().slice(0,5).toUpperCase();}
function isBroadDriveAccess_(access){
 var value=String(access||'').toUpperCase();
 return value==='ANYONE'||value==='ANYONE_WITH_LINK'||value==='DOMAIN'||value==='DOMAIN_WITH_LINK';
}

/**
 * Jalankan manual SATU KALI setelah deploy pembaruan PIN-preview.
 * Fungsi ini mencabut izin Viewer email ketua yang pernah ditambahkan langsung
 * pada FILE individual sebelumnya. Fungsi tidak mengubah atau membagikan folder.
 */
function removeLegacyContingentFileSharing(){
 var ss=SpreadsheetApp.getActiveSpreadsheet(),report={filesChecked:0,viewersRemoved:0,skipped:0,errors:[]},emailByName={};
 var parent=ss.getSheetByName(CFG.kontingen.sheet);
 if(parent&&parent.getLastRow()>=2){var pv=parent.getDataRange().getValues(),ph=pv[0],pni=ph.indexOf('Nama Kontingen'),pei=ph.indexOf('Email');for(var pr=1;pr<pv.length;pr++){var pn=pni>=0?normalizeEvent_(pv[pr][pni]):'',pe=pei>=0?String(pv[pr][pei]||'').trim().toLowerCase():'';if(pn&&/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(pe))emailByName[pn]=pe;}}
 [CFG.kontingen,CFG.cabang].forEach(function(cfg){
  var sh=ss.getSheetByName(cfg.sheet);if(!sh||sh.getLastRow()<2)return;
  var values=sh.getDataRange().getValues(),h=values[0],ni=h.indexOf('Nama Kontingen'),ei=h.indexOf('Email'),pi=h.indexOf('PDF Formulir (link Drive)'),bi=h.indexOf('Berkas (link Drive)');
  if(ni<0)return;
  for(var r=1;r<values.length;r++){
   var name=String(values[r][ni]||'').trim(),email=(ei>=0?String(values[r][ei]||'').trim().toLowerCase():'')||emailByName[normalizeEvent_(name)]||'';
   if(!name||!email||!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)){report.skipped++;continue;}
   var urls=[];if(pi>=0&&values[r][pi])urls.push(String(values[r][pi]));if(bi>=0)parsePortalFiles_(values[r][bi]).forEach(function(x){urls.push(x.url);});
   var seen={};urls.forEach(function(url){var id=getDriveFileIdFromUrl_(url);if(!id||seen[id])return;seen[id]=true;report.filesChecked++;
    try{DriveApp.getFileById(id).removeViewer(email);report.viewersRemoved++;}
    catch(err){var msg=String(err&&err.message?err.message:err);if(!/not a viewer|does not have|not found|permission/i.test(msg))report.errors.push(name+' | '+id+' | '+msg);}
   });
  }
 });
 return report;
}

function assertPrivateFolderTree_(folder){
 var current=folder,depth=0;
 while(current&&depth<8){
  var access;try{access=current.getSharingAccess();}catch(e){throw new Error('Tidak dapat memeriksa izin folder '+current.getName()+'. Hentikan unggahan sampai izin Drive dapat diperiksa: '+e);}
  if(isBroadDriveAccess_(access))throw new Error('Folder arsip '+current.getName()+' masih memiliki akses publik/domain ('+access+'). Ubah akses folder dan seluruh induknya menjadi Restricted/Private sebelum menerima berkas.');
  var parents=current.getParents();if(!parents.hasNext())break;current=parents.next();depth++;
 }
}
function getDriveFileIdFromUrl_(url){var m=String(url||'').match(/[-\w]{25,}/);return m?m[0]:'';}
function revokeIndividualFileViewer_(url,email){
 email=String(email||'').trim().toLowerCase();var id=getDriveFileIdFromUrl_(url);if(!email||!id)return;
 try{DriveApp.getFileById(id).removeViewer(email);}catch(e){console.warn('Tidak dapat mencabut akses pratinjau '+email+' pada file '+id+': '+e);}
}
function revokeRecordPreviewAccess_(sheet,rowNumber,email){
 if(!sheet||!email)return;
 var headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
 var values=sheet.getRange(rowNumber,1,1,sheet.getLastColumn()).getValues()[0];
 function value(key){var i=headers.indexOf(key);return i<0?'':values[i];}
 var urls=parsePortalFiles_(value('Berkas (link Drive)')).map(function(x){return x.url;});
 urls.push(String(value('PDF Formulir (link Drive)')||''));
 var unique={};urls.forEach(function(url){var id=getDriveFileIdFromUrl_(url);if(id&&!unique[id]){unique[id]=true;revokeIndividualFileViewer_(url,email);}});
}

function archiveDeletedSnapshot_(archive,admin,type,headers,values){
 if(archive.getLastRow()===0){archive.appendRow(['Waktu','Petugas','Jenis Data','No. Registrasi','Nama Kontingen','Snapshot JSON']);archive.setFrozenRows(1);archive.getRange(1,1,1,6).setFontWeight('bold');}
 archive.appendRow([new Date(),admin.user||'Admin',type,String(headers.indexOf('No. Registrasi')>=0?values[headers.indexOf('No. Registrasi')]||'':''),String(headers.indexOf('Nama Kontingen')>=0?values[headers.indexOf('Nama Kontingen')]||'':''),JSON.stringify({headers:headers,values:values})]);
}
function protectDeletedArchive_(sheet){
 try{var protection=sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0]||sheet.protect();protection.setDescription('Arsip data pendaftaran terhapus — akses khusus pemilik database.');protection.setWarningOnly(false);var editors=protection.getEditors();if(editors&&editors.length)protection.removeEditors(editors);if(protection.canDomainEdit())protection.setDomainEdit(false);}catch(e){console.warn('Perlindungan sheet arsip perlu diperiksa manual: '+e);}
}

function auditPorseniDrivePrivacy(){
 var ss=SpreadsheetApp.getActiveSpreadsheet(),log=ss.getSheetByName('Audit Izin Drive')||ss.insertSheet('Audit Izin Drive');
 log.clearContents();log.appendRow(['Waktu Audit','Jenis','Nama','Jalur','Sharing Access','Status']);log.setFrozenRows(1);log.getRange(1,1,1,6).setFontWeight('bold');
 var result={folders:0,files:0,issues:[]},now=new Date();
 function inspect(item,type,path){
  var access='';try{access=item.getSharingAccess();}catch(e){access='TIDAK DAPAT DIPERIKSA: '+String(e);}
  var issue=isBroadDriveAccess_(access)||String(access).indexOf('TIDAK DAPAT DIPERIKSA')===0;
  var label=issue?'PERLU TINDAKAN':'TERBATAS';log.appendRow([now,type,item.getName(),path,String(access),label]);
  if(type==='Folder')result.folders++;else result.files++;
  if(issue)result.issues.push({type:type,name:item.getName(),path:path,access:String(access)});
 }
 function walk(folder,path,depth){
  inspect(folder,'Folder',path);if(depth>12)return;
  var files=folder.getFiles();while(files.hasNext()){var file=files.next();inspect(file,'File',path+'/'+file.getName());}
  var folders=folder.getFolders();while(folders.hasNext()){var child=folders.next();walk(child,path+'/'+child.getName(),depth+1);}
 }
 var root=getFolderPath(DRIVE_PATH);walk(root,'/'+DRIVE_PATH.join('/'),0);
 log.appendRow([now,'RINGKASAN','Folder: '+result.folders+' · File: '+result.files,'','',result.issues.length?'PERLU TINDAKAN':'BERSIH']);
 protectDeletedArchive_(log);
 return {ok:true,folders:result.folders,files:result.files,issueCount:result.issues.length,issues:result.issues};
}

function resolveDestination(d,c){
 var root=getFolderPath(DRIVE_PATH);assertPrivateFolderTree_(root);
 var base=getFolder(root,c.folder);assertPrivateFolderTree_(base);
 var target;
 if(d.type==='kontingen'){
  target=getFolder(base,safe(d.nama).toUpperCase());
 }else{
  var branch=getFolder(base,safe(d.cabang).toUpperCase());
  var category=getFolder(branch,safe(d.kategori).toUpperCase());
  target=getFolder(category,safe(d.nama).toUpperCase());
 }
 assertPrivateFolderTree_(target);
 return target;
}
function displayBranchName_(branch){
 var labels={'Duta Pelajar':'Duta Pelajar','Khitobah':'Khitobah','LCC':'LCC','Karya ilmiah':'Karya Ilmiah','Baca kitab':'Baca Kitab','Badminton':'Badminton','Voli':'Voli','Sprint':'Sprint','Tenis meja':'Tenis Meja','PES PS3':'PES PS3','Parade Gelar Budaya':'Parade Gelar Budaya','Rebana':'Rebana','Paduan suara':'Paduan Suara','Kaligrafi':'Kaligrafi','Baca puisi':'Baca Puisi','MTQ':'MTQ'};
 return labels[branch]||branch;
}
function displayCategoryName_(branch,category){var value=String(category||'').trim();var legacy={'Tunggal Putra':'Putra','Tunggal Putri':'Putri','Beregu Campuran':'Kelompok Campuran','Beregu Bebas':'Kelompok Bebas','Beregu Putra':'Kelompok Putra'};return legacy[value]||value;}
function branchDeadline_(branch){return branch==='Badminton'?'29 November 2026':(branch==='Duta Pelajar'||branch==='Voli'?'6 Desember 2026':'19 Desember 2026');}
function branchSchedule_(branch){
 var m={
  'Duta Pelajar':'Pembekalan pada Jumat, 25 Desember 2026 pagi. Grand Final Duta Pelajar pada Sabtu, 26 Desember 2026.',
  'Badminton':'Road to PORSENI, Minggu 6 Desember 2026. Waktu mengikuti informasi panitia.',
  'Voli':'Road to PORSENI, Minggu 13 Desember 2026. Waktu mengikuti informasi panitia.',
  'Khitobah':'Sabtu, 26 Desember 2026, pukul 13.00–18.00 WIB.',
  'Baca kitab':'Sabtu, 26 Desember 2026, pukul 13.00–18.00 WIB.',
  'Baca puisi':'Sabtu, 26 Desember 2026, pukul 13.00–18.00 WIB.',
  'Sprint':'Sabtu, 26 Desember 2026, pukul 08.00–12.00 WIB.',
  'Tenis meja':'Sabtu, 26 Desember 2026, pukul 08.00–12.00 WIB.',
  'Kaligrafi':'Sabtu, 26 Desember 2026, pukul 08.00–12.00 WIB.',
  'Paduan suara':'Sabtu, 26 Desember 2026, pukul 08.00–12.00 WIB.',
  'MTQ':'Sabtu, 26 Desember 2026, pukul 08.00–12.00 WIB.',
  'Parade Gelar Budaya':'Jumat, 25 Desember 2026, pukul 20.00–23.30 WIB.',
  'Rebana':'Sabtu, 26 Desember 2026, sesi 1 pukul 16.00–18.00 WIB dan sesi 2 pukul 19.30–23.00 WIB.',
  'PES PS3':'Sabtu, 26 Desember 2026, pukul 19.30–23.00 WIB.',
  'LCC':'Sabtu, 26 Desember 2026, pukul 19.30–23.00 WIB.',
  'Karya ilmiah':'Sabtu, 26 Desember 2026, pukul 19.30–23.00 WIB.'
 };
 return m[branch]||'Sabtu, 26 Desember 2026, mengikuti susunan acara PORSENI 2026.';
}
// Jalankan fungsi ini satu kali dari editor Apps Script untuk menyalin formulir PDF lama
// yang sudah terdaftar sebelum fitur arsip panitia ditambahkan.
function archiveExistingPrintablePdfs(){
 var ss=SpreadsheetApp.getActiveSpreadsheet();
 [CFG.kontingen,CFG.cabang].forEach(function(cfg){
  var sh=ss.getSheetByName(cfg.sheet);if(!sh||sh.getLastRow()<2)return;
  var headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var iPdf=headers.indexOf('PDF Formulir (link Drive)');
  var iName=headers.indexOf('Nama Kontingen');
  var iBranch=headers.indexOf('Cabang Lomba');
  var iCategory=headers.indexOf('Kategori');
  var iArchive=headers.indexOf('Arsip PDF Panitia (link Drive)');
  if(iArchive<0){iArchive=headers.length;sh.getRange(1,iArchive+1).setValue('Arsip PDF Panitia (link Drive)').setFontWeight('bold');}
  if(iPdf<0||iName<0) return;
  var rows=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  rows.forEach(function(row,index){
   if(String(row[iArchive]||'').trim())return;
   var url=String(row[iPdf]||'').trim();if(!url)return;
   var match=url.match(/[-\w]{25,}/);if(!match)return;
   try{
    var source=DriveApp.getFileById(match[0]);
    var d={type:cfg===CFG.kontingen?'kontingen':'cabang',nama:String(row[iName]||'Kontingen')};
    if(d.type==='cabang'){d.cabang=iBranch>=0?String(row[iBranch]||'Cabang'): 'Cabang';d.kategori=iCategory>=0?String(row[iCategory]||'Kategori'):'Kategori';}
    var archive=archivePrintablePdf_(source,d);
    sh.getRange(index+2,iArchive+1).setValue(archive.getUrl());
   }catch(err){console.error('Gagal mengarsipkan PDF baris '+(index+2)+' sheet '+cfg.sheet+': '+err);}
  });
 });
 return 'Tautan arsip formulir diselaraskan ke PDF di folder 02 dan 03. Tidak membuat salinan folder 04.';
}

function archivePrintablePdf_(pdfFile,d){
 // PDF formulir sudah dibuat langsung di folder utama 02 atau 03 sesuai jenis pendaftaran.
 // Tidak membuat salinan kedua di folder 04; tautan arsip panitia menunjuk ke PDF yang sama.
 return pdfFile;
}

var KOP_PORSENI_BASE64='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';
/* ===== PDF RESMI — LAYOUT V30 =====
   Kop dipasang di BODY (bukan header Google Docs) pada tabel pembungkus selebar A4 penuh
   (595,28 pt) dengan margin halaman kiri/kanan 0. Margin isi 2 cm dibuat lewat padding sel,
   sehingga posisi kop dan isi konsisten dan tidak bergantung pada indent paragraf/tabel. */
var PDF_PAGE_W = 595.28;
var PDF_PAGE_H = 841.89;
var PDF_SIDE = 56.7;                       // margin isi kiri/kanan = 2 cm
var PDF_CONTENT_W = PDF_PAGE_W - PDF_SIDE * 2; // 481,88 pt
var PDF_KOP_PX_W = 793;                    // lebar A4 dalam piksel Docs (793 px = 594,75 pt)
var PDF_KOP_PX_H = 144;                    // rasio kop 2448 x 444
var PDF_NAVY = '#0B2D72';
var PDF_INK = '#111111';
var ID_MONTHS_ = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];

function idDate_(v, withTime) {
  if (v === undefined || v === null || v === '') return '';
  var s = String(v).trim(), m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m && !withTime) return m[3] + ' ' + ID_MONTHS_[parseInt(m[2], 10) - 1] + ' ' + m[1];
  var dt = (v instanceof Date) ? v : new Date(s);
  if (isNaN(dt.getTime())) return s;
  var p = Utilities.formatDate(dt, 'Asia/Jakarta', 'dd|M|yyyy|HH:mm:ss').split('|');
  var out = p[0] + ' ' + ID_MONTHS_[parseInt(p[1], 10) - 1] + ' ' + p[2];
  return withTime ? out + ', ' + p[3] + ' WIB' : out;
}

function pdfFont_(el, size, bold, color) {
  el.setFontFamily('Times New Roman').setFontSize(size).setBold(!!bold).setItalic(false).setUnderline(false).setForegroundColor(color || PDF_INK);
  return el;
}

/* ============================================================
   PATCH KOP HEADER — ganti DUA fungsi di script Anda:
     1) pdfNewDoc_   (hapus versi lama, tempel versi ini)
     2) pdfPage_     (hapus versi lama, tempel versi ini)
   Fungsi lain tidak perlu disentuh.
   ============================================================ */

function pdfNewDoc_(title, marginTop) {
  var doc = DocumentApp.create(title);
  var body = doc.getBody();

  // Margin atas = tinggi kop (pt) + jarak 12 pt. 144 px × 0,75 = 108 pt  →  120 pt.
  var topMargin = Math.round(PDF_KOP_PX_H * 0.75) + 12;

  body.setPageWidth(PDF_PAGE_W)
      .setPageHeight(PDF_PAGE_H)
      .setMarginTop(topMargin)
      .setMarginBottom(24)
      .setMarginLeft(0)     // 0 supaya kop di header penuh lebar A4
      .setMarginRight(0);

  // ----- Kop di header Google Docs (berulang di setiap halaman) -----
  var header = doc.addHeader();
  var kop = Utilities.newBlob(
    Utilities.base64Decode(KOP_PORSENI_BASE64),
    'image/png',
    'kop-porseni-resmi.png'
  );

  // Pakai paragraf pertama bawaan header (JANGAN appendParagraph: itu menambah baris kosong
  // di atas gambar sehingga kop turun ±15 pt dari tepi atas).
  var hp = header.getParagraphs();
  var p = hp.length ? hp[0] : header.appendParagraph('\u00A0');
  p.setAlignment(DocumentApp.HorizontalAlignment.LEFT)
   .setSpacingBefore(0)
   .setSpacingAfter(0)
   .setLineSpacing(1)
   .setIndentStart(0)
   .setIndentFirstLine(0)
   .setIndentEnd(0);
  var img = p.appendInlineImage(kop);
  // WAJIB lebar DAN tinggi. Tanpa setHeight gambar memakai tinggi aslinya (444 px)
  // sehingga kop membesar dan hurufnya tampak gepeng/memanjang.
  img.setWidth(PDF_KOP_PX_W).setHeight(PDF_KOP_PX_H);
  try { p.editAsText().setFontSize(1); } catch (e) {}   // opsional: rapatkan celah di bawah gambar

  // Rapatkan header ke tepi atas kertas.
  try {
    if (typeof Docs !== 'undefined') {
      Docs.Documents.batchUpdate({
        requests: [{
          updateDocumentStyle: {
            documentStyle: { marginHeader: { magnitude: 0, unit: 'PT' } },
            fields: 'marginHeader'
          }
        }]
      }, doc.getId());
    }
  } catch (e) {
    console.warn('marginHeader tidak dapat diatur (aktifkan layanan Google Docs API): ' + e);
  }

  return doc;
}

// Pembungkus isi: tabel selebar A4 penuh, margin kiri/kanan 2 cm dibuat lewat padding sel.
// (Margin halaman = 0 agar header penuh lebar, jadi body perlu margin buatan.)
function pdfPage_(body) {
  var t = body.appendTable([['']]);

  t.setBorderWidth(0)
   .setBorderColor('#FFFFFF')
   .setColumnWidth(0, PDF_PAGE_W);          // 595,28 pt = lebar A4 penuh

  var cell = t.getCell(0, 0);

  cell.setPaddingTop(0)
      .setPaddingBottom(0)
      .setPaddingLeft(PDF_SIDE)             // 56,7 pt = 2 cm
      .setPaddingRight(PDF_SIDE);

  return cell;
}

function pdfWriter_(cell) {
  var first = true;
  return {
    para: function (text, o) {
      o = o || {};
      var q;
      if (first) { q = cell.getChild(0).asParagraph(); q.setText(text); first = false; }
      else q = cell.appendParagraph(text);
      q.setAlignment(o.align || DocumentApp.HorizontalAlignment.LEFT)
       .setIndentStart(o.indent || 0).setIndentFirstLine(0).setIndentEnd(0)
       .setSpacingBefore(o.before || 0).setSpacingAfter(o.after || 0).setLineSpacing(o.line || 1.15);
      pdfFont_(q.editAsText(), o.size || 11, o.bold, o.color);
      if (o.italic) q.setItalic(true);
      return q;
    },
    table: function (rows) { first = false; return cell.appendTable(rows); },
    image: function (blob, width, height) {
      var p;
      // Jangan pernah menyisipkan teks/paragraf kosong: Google Docs dapat
      // menolak setText('') atau appendParagraph('') dengan error
      // "Tidak dapat menyisipkan elemen teks kosong".
      if (first) {
        // Gunakan paragraf kosong bawaan sel sebagai wadah gambar.
        p = cell.getChild(0).asParagraph();
        first = false;
      } else {
        // Gunakan NBSP sebagai jangkar paragraf non-kosong sebelum gambar.
        p = cell.appendParagraph('\u00A0');
      }
      p.setAlignment(DocumentApp.HorizontalAlignment.CENTER)
       .setSpacingBefore(0).setSpacingAfter(4).setLineSpacing(1);
      var im = p.appendInlineImage(blob);
      if (width && height) im.setWidth(width).setHeight(height);
      return p;
    }
  };
}

// Gaya tabel data: lebar kolom tetap, padding seragam, header berlatar biru muda.
function pdfGrid_(t, widths, o) {
  o = o || {};
  var size = o.size || 11, padV = o.padV === undefined ? 3 : o.padV, padH = o.padH === undefined ? 6 : o.padH;
  t.setBorderWidth(0.75).setBorderColor('#44556B');
  for (var c = 0; c < widths.length; c++) t.setColumnWidth(c, widths[c]);
  for (var r = 0; r < t.getNumRows(); r++) {
    var row = t.getRow(r);
    if (o.minH) row.setMinimumHeight(o.minH);
    for (var c2 = 0; c2 < row.getNumCells(); c2++) {
      var cell = row.getCell(c2), isHead = !!o.header && r === 0;
      cell.setPaddingTop(padV).setPaddingBottom(padV).setPaddingLeft(padH).setPaddingRight(padH);
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      var para = cell.getChild(0).asParagraph();
      para.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1.1).setIndentStart(0).setIndentFirstLine(0).setIndentEnd(0);
      var al = (o.align && o.align[c2]) || 'L';
      para.setAlignment(isHead || al === 'C' ? DocumentApp.HorizontalAlignment.CENTER : DocumentApp.HorizontalAlignment.LEFT);
      var boldCol = o.boldCols && o.boldCols.indexOf(c2) >= 0;
      pdfFont_(cell.editAsText(), size, isHead || boldCol, isHead ? PDF_NAVY : PDF_INK);
      if (isHead) cell.setBackgroundColor('#E3EAF4');
      else if (o.labelCol === c2) cell.setBackgroundColor('#F3F6FA');
    }
  }
}

// Hapus paragraf kosong bawaan di awal dokumen dan kecilkan paragraf penghubung antar tabel/halaman.
function pdfTidyBody_(body) {
  try {
    var first = body.getChild(0);
    if (first.getType() === DocumentApp.ElementType.PARAGRAPH && body.getNumChildren() > 1 && first.asParagraph().getText() === '') first.removeFromParent();
  } catch (e) {}
  for (var i = 0; i < body.getNumChildren(); i++) {
    var ch = body.getChild(i);
    if (ch.getType() === DocumentApp.ElementType.PARAGRAPH) {
      var p = ch.asParagraph();
      if (p.getText() === '') { p.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1).setIndentStart(0).setIndentFirstLine(0); p.editAsText().setFontSize(1); }
    }
  }
}

function pdfSavePdf_(doc, d, folder) {
  doc.saveAndClose();
  var file = DriveApp.getFileById(doc.getId());
  var pdfBlob = file.getAs(MimeType.PDF).setName(pdfName(d)), pdf = folder.createFile(pdfBlob);
  file.setTrashed(true);
  return pdf;
}



function createRegistrationPdf(d, folder) {
  if (d.type === 'cabang') {
    return createBranchRegistrationPdf_(d, folder);
  }

  var doc = pdfNewDoc_('TEMP PORSENI ' + d.reg, 0);
  var body = doc.getBody();

  var CENTER = DocumentApp.HorizontalAlignment.CENTER;
  var RIGHT = DocumentApp.HorizontalAlignment.RIGHT;
  var JUSTIFY = DocumentApp.HorizontalAlignment.JUSTIFY;

  function heading(w, text, before) {
    return w.para(text, {
      size: 12.5,
      bold: true,
      color: PDF_NAVY,
      before: before === undefined ? 8 : before,
      after: 3
    });
  }

  // ===== HALAMAN 1: IDENTITAS KONTINGEN =====
  var w = pdfWriter_(pdfPage_(body));

  w.para('FORMULIR PENDAFTARAN KONTINGEN', {
    size: 17,
    bold: true,
    align: CENTER,
    before: 0,
    after: 13
  });

  heading(w, 'A. IDENTITAS KONTINGEN', 0);

  var tempatTanggal = [
    String(d.tempat_lahir_ketua || '').trim(),
    idDate_(d.tanggal_lahir_ketua)
  ].filter(Boolean).join(', ');

  var ident = [
    ['Nama Kontingen', d.nama || ''],
    ['Nomor Pendaftaran', d.reg || ''],
    ['PIN Pendaftaran', d.pin || 'Sesuai formulir pendaftaran awal'],
    ['Waktu Pendaftaran', idDate_(d.waktu || new Date(), true)],
    ['Nama Ketua Kontingen', d.ketua || ''],
    ['Tempat/Tanggal Lahir Ketua', tempatTanggal],
    ['Alamat Rumah Ketua', d.alamat || ''],
    ['No. WA Ketua Kontingen', d.wa_ketua || ''],
    ['Email Kontingen', d.email || ''],
    ['Nama Ketua IPNU', d.ketua_pr_ipnu || ''],
    ['Nama Ketua IPPNU', d.ketua_pr_ippnu || '']
  ];

  if (d.diperbarui) {
    ident.splice(4, 0, ['Terakhir Diperbarui', d.diperbarui]);
  }

  var kv = w.table(ident);

  pdfGrid_(kv, [165, PDF_CONTENT_W - 165], {
    size: 11.5,
    padV: 5,
    padH: 8,
    minH: 24,
    boldCols: [0],
    labelCol: 0
  });

  // PIN dibuat menonjol
  if (d.pin) {
    kv.getCell(2, 1).editAsText()
      .setBold(true)
      .setForegroundColor(PDF_NAVY)
      .setFontSize(13);
  }

  w.para('\u00A0', { size: 6, after: 8 });

  var box = w.table([['\u00A0']]);
  var bc = box.getCell(0, 0);

  box.setBorderWidth(0.75)
    .setBorderColor('#9DB0C8')
    .setColumnWidth(0, PDF_CONTENT_W);

  bc.setBackgroundColor('#F3F6FA')
    .setPaddingTop(8)
    .setPaddingBottom(8)
    .setPaddingLeft(10)
    .setPaddingRight(10);

  var bw = pdfWriter_(bc);

  bw.para('CATATAN PENTING', {
    size: 10.5,
    bold: true,
    color: PDF_NAVY,
    after: 3
  });

  bw.para(
    '1. Simpan PIN pendaftaran dan jangan bagikan kepada pihak lain. PIN digunakan untuk masuk ke Portal Kontingen dan mengirim berkas cabang lomba.',
    {
      size: 10.5,
      after: 2,
      align: JUSTIFY,
      line: 1.2
    }
  );

  bw.para(
    '2. Data pada formulir ini tidak dapat diubah langsung. Perubahan data dilakukan melalui Portal Kontingen selama pendaftaran belum diverifikasi oleh panitia.',
    {
      size: 10.5,
      after: 2,
      align: JUSTIFY,
      line: 1.2
    }
  );

  bw.para(
    '3. Setelah pendaftaran diverifikasi, pengubahan data harus diajukan terlebih dahulu kepada panitia melalui Portal Kontingen.',
    {
      size: 10.5,
      align: JUSTIFY,
      line: 1.2
    }
  );

  body.appendPageBreak();

  // ===== HALAMAN 2: CABANG PERLOMBAAN =====
  w = pdfWriter_(pdfPage_(body));

  heading(w, 'B. CABANG PERLOMBAAN YANG DIIKUTI', 4);

  w.para(
    'Tanda centang (✓) menunjukkan cabang perlombaan yang diikuti oleh kontingen.',
    {
      size: 11,
      after: 8
    }
  );

  var catalog = [
    ['A. AKADEMIK', 'Duta Pelajar', ['Putra', 'Putri']],
    ['', 'Khitobah', ['Putra', 'Putri']],
    ['', 'LCC', ['Kelompok Campuran']],
    ['', 'Karya Ilmiah', ['Kelompok Campuran']],
    ['', 'Baca Kitab', ['Putra', 'Putri']],
    ['B. OLAHRAGA', 'Badminton', ['Ganda Campuran']],
    ['', 'Voli', ['Kelompok Putra']],
    ['', 'Sprint', ['Putra', 'Putri']],
    ['', 'Tenis Meja', ['Putra', 'Putri']],
    ['', 'PES PS3', ['Tunggal Bebas']],
    ['C. SENI', 'Parade Gelar Budaya', ['Kelompok Campuran']],
    ['', 'Rebana', ['Kelompok Bebas']],
    ['', 'Paduan Suara', ['Kelompok Campuran']],
    ['', 'Kaligrafi', ['Putra', 'Putri']],
    ['', 'Baca Puisi', ['Putra', 'Putri']],
    ['', 'MTQ', ['Putra', 'Putri']]
  ];

  var selected = String(d.cabang || '')
    .split(';')
    .map(function(x) {
      return x.trim().toLowerCase();
    });

  var matrix = [
    ['KELOMPOK', 'LOMBA', 'KATEGORI', 'YA', 'TIDAK']
  ];

  var yesRows = [];

  catalog.forEach(function(x) {
    x[2].forEach(function(k, n) {
      var yes = selected.some(function(v) {
        return (
          v.indexOf(x[1].toLowerCase() + ' (') === 0 &&
          v.indexOf(k.toLowerCase()) >= 0
        ) || (
          v === x[1].toLowerCase() &&
          x[2].length === 1
        );
      });

      matrix.push([
        n === 0 ? x[0] : '',
        n === 0 ? x[1] : '',
        k,
        yes ? '✓' : '',
        yes ? '' : '✓'
      ]);

      yesRows.push(yes);
    });
  });

  var mt = w.table(matrix);

  pdfGrid_(mt, [
    92,
    118,
    132,
    70,
    PDF_CONTENT_W - 412
  ], {
    size: 11,
    padV: 3,
    padH: 6,
    minH: 20,
    header: true,
    boldCols: [0, 1],
    align: ['L', 'L', 'L', 'C', 'C']
  });

  for (var rr = 1; rr < mt.getNumRows(); rr++) {
    mt.getCell(rr, 3).editAsText()
      .setBold(true)
      .setForegroundColor(PDF_NAVY);

    mt.getCell(rr, 4).editAsText()
      .setBold(true)
      .setForegroundColor('#6B7788');

    if (yesRows[rr - 1]) {
      for (var cc = 2; cc < 4; cc++) {
        mt.getCell(rr, cc).setBackgroundColor('#EAF4EC');
      }
    }
  }

  body.appendPageBreak();

  // ===== HALAMAN 3: ADMINISTRASI DAN TANDA TANGAN =====
  w = pdfWriter_(pdfPage_(body));

  heading(w, 'C. KELENGKAPAN ADMINISTRASI', 4);

  w.para('Berkas yang diserahkan kepada panitia:', {
    size: 11,
    after: 5
  });

  [
    'Bukti pembayaran',
    'KTP ketua kontingen',
    'Pakta integritas',
    'SP / berita acara RAT'
  ].forEach(function(label) {
    w.para(
      (findUpload_(d, label) ? '☑  ' : '☐  ') + label,
      {
        size: 11.5,
        bold: true,
        after: 3,
        indent: 14
      }
    );
  });

  heading(w, 'D. PERNYATAAN PENDAFTARAN', 14);

  w.para(
    'Dengan ini kami menyatakan bahwa data yang tercantum dalam formulir pendaftaran kontingen ini telah diisi dengan benar dan dapat dipertanggungjawabkan. Kami bersedia melengkapi persyaratan administrasi serta mengikuti seluruh ketentuan yang telah ditetapkan oleh Panitia Pelaksana PORSENI 2026 PAC IPNU–IPPNU Patebon.',
    {
      size: 11.5,
      align: JUSTIFY,
      line: 1.4,
      after: 6
    }
  );

  w.para(
    'Demikian formulir pendaftaran ini kami buat untuk digunakan sebagaimana mestinya.',
    {
      size: 11.5,
      align: JUSTIFY,
      line: 1.4,
      after: 12
    }
  );

  w.para('Patebon, ' + idDate_(d.waktu || new Date()), {
    size: 11.5,
    align: RIGHT,
    after: 12
  });

  var sg = w.table([
    ['Ketua Kontingen', 'Ketua IPNU', 'Ketua IPPNU'],
    ['\u00A0', '\u00A0', '\u00A0'],
    [
      '( ' + (d.ketua || d.nama || '') + ' )',
      '( ' + (d.ketua_pr_ipnu || '') + ' )',
      '( ' + (d.ketua_pr_ippnu || '') + ' )'
    ]
  ]);

  sg.setBorderWidth(0).setBorderColor('#FFFFFF');

  var sw = PDF_CONTENT_W / 3;

  for (var sc = 0; sc < 3; sc++) {
    sg.setColumnWidth(sc, sw);
  }

  for (var sr = 0; sr < 3; sr++) {
    if (sr === 1) {
      sg.getRow(sr).setMinimumHeight(70);
    }

    for (var sc2 = 0; sc2 < 3; sc2++) {
      var cell = sg.getCell(sr, sc2);

      cell.setPaddingTop(3)
        .setPaddingBottom(3)
        .setPaddingLeft(4)
        .setPaddingRight(4)
        .setVerticalAlignment(
          DocumentApp.VerticalAlignment.CENTER
        );

      var sp = cell.getChild(0).asParagraph();

      sp.setAlignment(CENTER)
        .setSpacingBefore(0)
        .setSpacingAfter(0)
        .setLineSpacing(1.1)
        .setIndentStart(0)
        .setIndentFirstLine(0)
        .setIndentEnd(0);

      pdfFont_(cell.editAsText(), 11.5, true);
    }
  }

  // Tiga area tanda tangan sengaja dibiarkan kosong untuk tanda tangan manual.

  w.para('\u00A0', { size: 6, after: 10 });

  w.para(
    'Dokumen ini dibuat secara elektronik melalui website resmi PORSENI 2026.',
    {
      size: 10,
      italic: true,
      align: CENTER,
      color: '#3A4A63',
      before: 6
    }
  );

  appendImageAttachmentPages_(body, d);
  appendPdfAttachmentIndex_(body, d);
  pdfTidyBody_(body);

  return pdfSavePdf_(doc, d, folder);
}



function findUpload_(d,label){var target=String(label||'').toLowerCase();return (d.files||[]).filter(function(f){return String(f.label||'').toLowerCase()===target;})[0]||null;}

function attachmentOwnerLabel_(label, d) {
  var m=String(label||'').match(/(?:peserta|participant)\s*(\d+)/i);
  if(m){var idx=Number(m[1])-1, person=(d.pesertaList||[])[idx];return person&&person.nama?person.nama:'Peserta '+m[1];}
  if(/duta pelajar/i.test(String(label||''))&&(d.pesertaList||[])[0])return d.pesertaList[0].nama||d.nama||'Peserta Duta Pelajar';
  return d.type==='kontingen'?(d.ketua||d.nama||'Kontingen'):(d.nama||'Kontingen');
}

function appendBranchImageAttachmentPages_(body, d) {
  var people=d.pesertaList||[], files=d.files||[];
  function findImage(label) {
    var target=String(label||'').trim().toLowerCase();
    return files.filter(function(f){return String(f.label||'').trim().toLowerCase()===target;})[0]||null;
  }
  function addGridPages(kind, perPage, cols, rows, imageW, imageH) {
    var items=[];
    people.forEach(function(person,i){
      var label=(kind==='foto'?'Pas foto peserta ':'KTP peserta ')+(i+1);
      var f=findImage(label);
      if(f&&f.data)items.push({file:f,person:person,no:i+1});
    });
    for(var offset=0;offset<items.length;offset+=perPage){
      body.appendPageBreak();
      var host=pdfPage_(body);
      var batch=items.slice(offset,offset+perPage);
      var tableRows=[];
      for(var r=0;r<rows;r++){
        var row=[];
        for(var c=0;c<cols;c++)row.push('');
        tableRows.push(row);
      }
      var grid=host.appendTable(tableRows);
      grid.setBorderWidth(0).setBorderColor('#FFFFFF');
      var cellW=PDF_CONTENT_W/cols;
      for(var c2=0;c2<cols;c2++)grid.setColumnWidth(c2,cellW);
      for(var k=0;k<batch.length;k++){
        var item=batch[k],rr=Math.floor(k/cols),cc=k%cols,cell=grid.getCell(rr,cc);
        cell.setPaddingTop(5).setPaddingBottom(8).setPaddingLeft(5).setPaddingRight(5);
        cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
        try{
          var bytes=Utilities.base64Decode(item.file.data);
          var mime=detectedMime_(bytes,item.file.type||'image/jpeg');
          if(mime!=='image/png'&&mime!=='image/jpeg')continue;
          var p=cell.getChild(0).asParagraph();
          p.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(0).setSpacingAfter(3).setLineSpacing(1);
          var img=p.appendInlineImage(Utilities.newBlob(bytes,mime,item.file.name||'foto.jpg'));
          var ow=img.getWidth(),oh=img.getHeight();
          if(ow>0&&oh>0){
            var scale=Math.min(imageW/ow,imageH/oh);
            img.setWidth(Math.max(1,Math.round(ow*scale))).setHeight(Math.max(1,Math.round(oh*scale)));
          }
          var cap=cell.appendParagraph('Peserta No. '+item.no+' – '+String(item.person.nama||''));
          cap.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(2).setSpacingAfter(0).setLineSpacing(1);
          pdfFont_(cap.editAsText(),8.5,true,PDF_INK);
        }catch(err){
          var fail=cell.getChild(0).asParagraph();fail.setText('Gambar peserta '+item.no+' tidak dapat dimasukkan ke PDF.');
          pdfFont_(fail.editAsText(),8,false,PDF_INK);
        }
      }
    }
  }
  // Pas foto: 3 kolom x 4 baris, ukuran target 4 x 6 cm.
  addGridPages('foto',12,3,4,113.4,170.1);
  // KTP: 2 kolom x 3 baris, ukuran mendekati ukuran kartu asli.
  addGridPages('ktp',6,2,3,243,153);
}

function appendImageAttachmentPages_(body, d) {
  (d.files||[]).forEach(function(f){
    if(!f||!f.data)return;
    var bytes;try{bytes=Utilities.base64Decode(f.data);}catch(e){return;}
    var mime=detectedMime_(bytes,f.type||'application/octet-stream');
    if(mime!=='image/png'&&mime!=='image/jpeg')return;
    body.appendPageBreak();
    var cell=pdfPage_(body);
    var title=cell.getChild(0).asParagraph();
    title.setText('LAMPIRAN: '+String(f.label||f.name||'DOKUMEN').toUpperCase());
    title.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(0).setSpacingAfter(12).setLineSpacing(1.15);
    pdfFont_(title.editAsText(),11,true,PDF_NAVY);
    try{
      var blob=Utilities.newBlob(bytes,mime,String(f.name||'lampiran'));
      var isKtp=String(f.label||'').toLowerCase()==='ktp ketua kontingen';
      var imageHost=cell;
      if(isKtp){
        var frame=cell.appendTable([['']]);
        frame.setBorderWidth(0.75).setBorderColor('#9DB0C8').setColumnWidth(0,PDF_CONTENT_W);
        imageHost=frame.getCell(0,0);
        imageHost.setPaddingTop(10).setPaddingBottom(10).setPaddingLeft(10).setPaddingRight(10);
      }
      var img=imageHost.appendImage(blob),ow=img.getWidth(),oh=img.getHeight();
      // KTP masuk ke kotak A4 dengan rasio asli; tidak dipotong atau diregangkan.
      if(ow>0&&oh>0){var scale=isKtp?Math.min(430/ow,270/oh,1):Math.min(455/ow,610/oh,1);img.setWidth(Math.max(1,Math.round(ow*scale))).setHeight(Math.max(1,Math.round(oh*scale)));}
      img.getParent().asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    }catch(err){
      var note=cell.appendParagraph('LAMPIRAN: '+String(f.label||'DOKUMEN').toUpperCase());
      note.setAlignment(DocumentApp.HorizontalAlignment.CENTER);pdfFont_(note.editAsText(),10,true,PDF_INK);
    }
  });
}
function appendPdfAttachmentIndex_(body,d){
  var pdfs=(d.files||[]).filter(function(f){
    if(!f)return false;
    if(String(f.type||'').toLowerCase()==='application/pdf')return true;
    if(!f.data)return false;
    try{return detectedMime_(Utilities.base64Decode(f.data),f.type||'')==='application/pdf';}catch(e){return false;}
  });
  if(!pdfs.length)return;
  // Satu PDF: halaman pemisah tepat sebelum lampiran. Beberapa PDF: SATU halaman daftar berurutan,
  // karena penggabungan menempelkan semua PDF setelah formulir sesuai urutan daftar.
  if(pdfs.length===1){
    body.appendPageBreak();
    var cell1=pdfPage_(body),p1=cell1.getChild(0).asParagraph();
    p1.setText('LAMPIRAN: '+String(pdfs[0].label||'DOKUMEN').toUpperCase());
    p1.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(0).setSpacingAfter(0);
    pdfFont_(p1.editAsText(),11,true,PDF_NAVY);
    return;
  }
  body.appendPageBreak();
  var cell=pdfPage_(body),title=cell.getChild(0).asParagraph();
  title.setText('DAFTAR LAMPIRAN PDF');
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(0).setSpacingAfter(10);
  pdfFont_(title.editAsText(),11,true,PDF_NAVY);
  var note=cell.appendParagraph('Lampiran PDF berikut dilampirkan pada halaman setelah halaman ini, berurutan sesuai daftar:');
  note.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(0).setSpacingAfter(10);
  pdfFont_(note.editAsText(),10,false,PDF_INK);
  pdfs.forEach(function(f,i){
    var li=cell.appendParagraph((i+1)+'. '+String(f.label||f.name||'Dokumen'));
    li.setAlignment(DocumentApp.HorizontalAlignment.LEFT).setSpacingBefore(0).setSpacingAfter(4);
    pdfFont_(li.editAsText(),11,false,PDF_INK);
  });
}
function appendKeyValueTable(body,rows){var t=body.appendTable(rows);t.setBorderWidth(0.5);for(var i=0;i<rows.length;i++){t.getCell(i,0).editAsText().setBold(true);}}
function pdfNamePart_(value){return String(value==null?'':value).trim().replace(/[\\/:*?"<>|\r\n]+/g,' ').replace(/\s+/g,' ').replace(/_+/g,'-').trim().slice(0,90)||'Tidak Diketahui';}
function pdfName(d){
 var reg=pdfNamePart_(d.reg||'Tanpa Nomor Registrasi'),nama=pdfNamePart_(d.nama||'Kontingen');
 if(d.type==='kontingen')return 'Formulir Pendaftaran Kontingen_'+nama+'_'+reg+'.pdf';
 return 'Form Pengumpulan Berkas_'+pdfNamePart_(d.cabang||'Cabang Lomba')+'_'+pdfNamePart_(d.kategori||'Kategori')+'_'+nama+'_'+reg+'.pdf';
}
function detectedMime_(bytes,fallback){
 if(!bytes||!bytes.length)return fallback||'application/octet-stream';
 var u=function(i){return bytes[i]&255;},ascii=function(a,z){return String.fromCharCode.apply(null,bytes.slice(a,z));};
 if(bytes.length>=5&&ascii(0,5)==='%PDF-')return 'application/pdf';
 if(bytes.length>=8&&u(0)===137&&ascii(1,4)==='PNG')return 'image/png';
 if(bytes.length>=3&&u(0)===255&&u(1)===216&&u(2)===255)return 'image/jpeg';
 if(bytes.length>=6&&/^GIF8[79]a$/.test(ascii(0,6)))return 'image/gif';
 if(bytes.length>=12&&ascii(0,4)==='RIFF'&&ascii(8,12)==='WEBP')return 'image/webp';
 if(bytes.length>=2&&ascii(0,2)==='BM')return 'image/bmp';
 if(bytes.length>=4&&((u(0)===73&&u(1)===73&&u(2)===42&&u(3)===0)||(u(0)===77&&u(1)===77&&u(2)===0&&u(3)===42)))return 'image/tiff';
 if(bytes.length>=4&&u(0)===0&&u(1)===0&&u(2)===1&&u(3)===0)return 'image/x-icon';
 if(bytes.length>=12&&ascii(4,8)==='ftyp'){var brand=ascii(8,12).toLowerCase();if(brand.indexOf('avi')>=0||brand==='avif'||brand==='avis')return 'image/avif';if(brand.indexOf('hei')>=0||brand.indexOf('mif')>=0)return 'image/heic';}
 var head=ascii(0,Math.min(bytes.length,512)).replace(/^\uFEFF/,'').trim().toLowerCase();
 if(head.indexOf('<svg')>=0||head.indexOf('<?xml')===0&&head.indexOf('<svg')>=0)return 'image/svg+xml';
 if(fallback&&String(fallback).toLowerCase().indexOf('image/')===0)return String(fallback).toLowerCase();
 return fallback||'application/octet-stream';
}
// PDF merge uses the PDFApp Google Apps Script library (library key in README).
var PDF_MERGE_QUEUE_PROPERTY_ = 'PORSENI_2026_PDF_MERGE_QUEUE_V1';
function hasPdfUploads_(files){
 return (files||[]).some(function(f){
  if(!f||!f.data)return false;
  try{return detectedMime_(Utilities.base64Decode(f.data),f.type||'')==='application/pdf';}catch(e){return false;}
 });
}
  function pdfMergeReady_(){return !(typeof PDFApp==='undefined' || !PDFApp || typeof PDFApp.mergePDFs!=='function');}
  // V34: tidak lagi menolak pendaftaran. Bila library belum ada, penggabungan tetap diantrekan dan
  // dijalankan otomatis begitu library terpasang (atau lewat tombol "Gabung ulang PDF" di Portal Panitia).
  function assertPdfMergeLibraryReady_(files){
   if(hasPdfUploads_(files) && !pdfMergeReady_()){
    try{logSecurityEvent_('Sistem','PENGGABUNGAN PDF','PERINGATAN','Library PDFApp belum terpasang; pendaftaran tetap diterima dan penggabungan ditunda.');}catch(e){}
   }
  }
/* ===== ANTREAN PENGGABUNGAN PDF — V38 =====
 * Setiap pekerjaan disimpan sebagai Script Property sendiri (PORSENI_PDF_JOB_<id>) dan baru
 * dihapus SETELAH selesai. Dengan begitu:
 *  - batas 6 menit / error di tengah proses tidak lagi menghilangkan pekerjaan yang belum selesai;
 *  - batas 9 KB per properti tidak terlampaui oleh antrean yang panjang;
 *  - trigger "penjaga" (watchdog) selalu terpasang selama proses berjalan, sehingga antrean
 *    otomatis dilanjutkan bila eksekusi terhenti.
 */
var PDF_MERGE_JOB_PREFIX_ = 'PORSENI_PDF_JOB_';
var PDF_MERGE_RUNNING_PROPERTY_ = 'PORSENI_PDF_MERGE_RUNNING';
var PDF_MERGE_BUDGET_MS_ = 4 * 60 * 1000 + 30 * 1000;   // berhenti mengambil pekerjaan baru setelah 4,5 menit
var PDF_MERGE_WATCHDOG_MS_ = 8 * 60 * 1000;             // lebih lama dari batas eksekusi Apps Script (6 menit)
var PDF_MERGED_MARKER_ = 'PORSENI_MERGED';

function listPdfMergeJobs_(){
 var props=PropertiesService.getScriptProperties(),all=props.getProperties(),jobs=[];
 Object.keys(all).forEach(function(k){
  if(k.indexOf(PDF_MERGE_JOB_PREFIX_)!==0)return;
  try{var j=JSON.parse(all[k]);j.qid=k.slice(PDF_MERGE_JOB_PREFIX_.length);jobs.push(j);}
  catch(e){props.deleteProperty(k);}
 });
 jobs.sort(function(a,b){return String(a.createdAt||'').localeCompare(String(b.createdAt||''));});
 return jobs;
}
function savePdfMergeJob_(job){
 var copy={};Object.keys(job).forEach(function(k){if(k!=='qid')copy[k]=job[k];});
 PropertiesService.getScriptProperties().setProperty(PDF_MERGE_JOB_PREFIX_+job.qid,JSON.stringify(copy));
}
function removePdfMergeJob_(qid){PropertiesService.getScriptProperties().deleteProperty(PDF_MERGE_JOB_PREFIX_+qid);}
// Antrean format lama (satu properti berisi array) dipindahkan ke format per-pekerjaan.
function migrateLegacyPdfMergeQueue_(){
 var props=PropertiesService.getScriptProperties(),raw=props.getProperty(PDF_MERGE_QUEUE_PROPERTY_);
 if(!raw)return;
 try{JSON.parse(raw).forEach(function(j){j.qid=Utilities.getUuid().replace(/-/g,'');savePdfMergeJob_(j);});}
 catch(e){console.error('Migrasi antrean PDF lama gagal: '+e);}
 props.deleteProperty(PDF_MERGE_QUEUE_PROPERTY_);
}
function clearPdfMergeTriggers_(){
 ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()==='processPdfMergeQueue_')ScriptApp.deleteTrigger(t);});
}
function ensurePdfMergeTrigger_(delayMs){
 var exists=ScriptApp.getProjectTriggers().some(function(t){return t.getHandlerFunction()==='processPdfMergeQueue_';});
 // Untuk pekerjaan baru, jadwalkan secepat mungkin. Apps Script tetap dapat
 // menjalankan trigger sedikit lebih lambat dari waktu minimum yang diminta.
 // Trigger watchdog tetap menggunakan jeda panjang yang diberikan pemanggil.
 if(!exists)ScriptApp.newTrigger('processPdfMergeQueue_').timeBased().after(Math.max(1000,Number(delayMs)||1000)).create();
}
function setPdfJobStatus_(job,state,message,extra){
 if(!job||!job.jobId)return;
 var st={at:Date.now(),state:state,message:message};
 if(extra)Object.keys(extra).forEach(function(k){st[k]=extra[k];});
 PropertiesService.getScriptProperties().setProperty('PORSENI_TEST_PDF_STATUS_'+job.jobId,JSON.stringify(st));
}
function isStandaloneCompetitionDocumentLabel_(label){
 var normalized=String(label||'').trim().toLowerCase();
 return normalized==='makalah kti (word)' || normalized==='makalah kti (pdf)' ||
        normalized==='presentasi kti (ppt)' || normalized==='naskah khitobah (pdf)' ||
        normalized==='soft file makalah kti (word)' || normalized==='soft file makalah kti (pdf)' ||
        normalized==='presentasi kti (ppt)';
}
function queuePdfMerge_(mainPdf,d,folder,options){
 options=options||{};
 var seenPdfIds={};
 var pdfIds=(d.savedAttachments||[]).filter(function(a){
  var id=String(a&&a.id||'');
  if(isStandaloneCompetitionDocumentLabel_(a&&a.label))return false;
  if(String(a&&a.mime||'').toLowerCase()!=='application/pdf'||!id||seenPdfIds[id])return false;
  seenPdfIds[id]=true;return true;
 }).map(function(a){return String(a.id);});
 if(!pdfIds.length && !options.jobId)return false;
 var jobId=String(options.jobId||''),mainPdfId=mainPdf.getId();
 var lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
  // Idempotensi: jangan membuat pekerjaan ganda untuk PDF formulir yang sama.
  var queued=listPdfMergeJobs_().filter(function(existing){
   return String(existing.mainPdfId||'')===mainPdfId &&
          String(existing.jobId||'')===jobId;
  })[0];
  if(queued){
   if(jobId)setPdfJobStatus_(queued,'processing','Penggabungan PDF sudah terjadwal; pekerjaan duplikat tidak dibuat.');
   ensurePdfMergeTrigger_(1000);
   return jobId||true;
  }
  var job={qid:Utilities.getUuid().replace(/-/g,''),mainPdfId:mainPdfId,pdfIds:pdfIds,folderId:folder.getId(),outputFolderId:options.outputFolderId||folder.getId(),cleanupFolder:!!options.cleanupFolder,jobId:jobId,reg:String(d.reg||''),type:String(d.type||''),nama:String(d.nama||''),cabang:String(d.cabang||''),kategori:String(d.kategori||''),fileName:options.fileName||pdfName(d),oldPdfIds:(options.oldPdfIds||[]).filter(Boolean),attempts:0,waits:0,createdAt:new Date().toISOString()};
  savePdfMergeJob_(job);
  setPdfJobStatus_(job,'processing','Sedang menyatukan formulir dan seluruh lampiran PDF.');
  ensurePdfMergeTrigger_(1000);
 }finally{lock.releaseLock();}
 return jobId||true;
}
function adminTestPdfStatus_(d){
 var admin=requireAdmin_(d.token),id=String(d.jobId||'');
 if(!/^[a-f0-9-]{20,80}$/i.test(id))throw new Error('ID pekerjaan preview tidak valid.');
 var raw=PropertiesService.getScriptProperties().getProperty('PORSENI_TEST_PDF_STATUS_'+id);
 if(!raw)return {ok:false,state:'missing',message:'Status PDF uji coba tidak ditemukan.'};
 var status=JSON.parse(raw);status.ok=true;return status;
}

async function processPdfMergeQueue_(){
 var startedAt=Date.now(),props=PropertiesService.getScriptProperties(),lock=LockService.getScriptLock(),jobs=[];
 if(!lock.tryLock(20000)){ensurePdfMergeTrigger_(2*60*1000);return;}
 try{
  var running=Number(props.getProperty(PDF_MERGE_RUNNING_PROPERTY_)||0);
  if(running&&Date.now()-running<PDF_MERGE_WATCHDOG_MS_)return;   // proses lain masih aktif; ia yang menjadwalkan ulang
  props.setProperty(PDF_MERGE_RUNNING_PROPERTY_,String(Date.now()));
  migrateLegacyPdfMergeQueue_();
  clearPdfMergeTriggers_();
  ensurePdfMergeTrigger_(PDF_MERGE_WATCHDOG_MS_+30*1000);          // penjaga bila eksekusi ini terhenti di tengah jalan
  jobs=listPdfMergeJobs_();
 }finally{lock.releaseLock();}
 var needMore=false,budgetHit=false;
 try{
  for(var i=0;i<jobs.length;i++){
   var job=jobs[i];
   if(Date.now()-startedAt>PDF_MERGE_BUDGET_MS_){needMore=true;budgetHit=true;break;}
   if(!pdfMergeReady_()){
    job.waits=Number(job.waits||0)+1;
    if(job.waits<=30){
     savePdfMergeJob_(job);
     setPdfJobStatus_(job,'processing','Menunggu library PDFApp terpasang.');
     needMore=true;
    }else{
     notifyMergeFailure_(job,'Library PDFApp tidak terpasang setelah 30 kali pengecekan.');
     setPdfJobStatus_(job,'error','Library PDFApp tidak terpasang.');
     removePdfMergeJob_(job.qid);
    }
    continue;
   }
   try{
    var finalFile;
    if(job.finalFileId){
     finalFile=DriveApp.getFileById(job.finalFileId);              // percobaan ulang: pakai hasil gabungan yang sudah ada, jangan menggabung dua kali
    }else{
     var outputFolder=DriveApp.getFolderById(job.outputFolderId||job.folderId);
     // Nama sementara deterministik memungkinkan pemulihan bila eksekusi berhenti
     // setelah createFile() tetapi sebelum finalFileId tersimpan di Script Properties.
     var pendingName='.__PORSENI_MERGE_PENDING__'+job.qid+'.pdf';
     var pendingFiles=outputFolder.getFilesByName(pendingName);
     if(pendingFiles.hasNext()){
      finalFile=pendingFiles.next();
     }else{
      var blobs=[DriveApp.getFileById(job.mainPdfId).getBlob()];
      (job.pdfIds||[]).forEach(function(id){blobs.push(DriveApp.getFileById(id).getBlob());});
      var mergedBlob=await PDFApp.mergePDFs(blobs);
      mergedBlob.setName(pendingName);
      finalFile=outputFolder.createFile(mergedBlob);
      try{finalFile.setDescription(PDF_MERGED_MARKER_+' | '+job.reg+' | QID:'+job.qid+' | '+new Date().toISOString());}catch(descErr){console.warn('Penanda hasil gabungan gagal dipasang: '+descErr);}
     }
     job.finalFileId=finalFile.getId();savePdfMergeJob_(job);
     finalFile.setName(job.fileName||('PORSENI_'+job.reg+'.pdf'));
    }
    var completedName=job.fileName||('PORSENI_'+job.reg+'.pdf');
    if(finalFile.getName()!==completedName)finalFile.setName(completedName);
    assertPrivateIndividualFile_(finalFile);
    updateMergedPdfRecord_(job,finalFile);
     try{sendApprovedContingentPdfEmailAfterMerge_(job,finalFile);}catch(autoMailErr){console.warn("Email PDF kontingen setelah penggabungan gagal: "+autoMailErr);}
    try{DriveApp.getFileById(job.mainPdfId).setTrashed(true);}catch(renameErr){}
    (job.oldPdfIds||[]).forEach(function(oldId){try{if(oldId&&oldId!==finalFile.getId())DriveApp.getFileById(oldId).setTrashed(true);}catch(cleanOldErr){console.warn('PDF lama tidak dapat dipindahkan ke Trash: '+cleanOldErr);}});
    if(job.cleanupFolder){try{DriveApp.getFolderById(job.folderId).setTrashed(true);}catch(cleanErr){console.warn('Folder sementara tidak dapat dibersihkan: '+cleanErr);}}
    setPdfJobStatus_(job,'ready','PDF final siap dipratinjau.',{fileId:finalFile.getId(),previewUrl:'https://drive.google.com/file/d/'+finalFile.getId()+'/preview',downloadUrl:'https://drive.google.com/uc?export=download&id='+finalFile.getId(),name:finalFile.getName()});
    logSecurityEvent_(job.nama,'PENGGABUNGAN PDF','BERHASIL',job.reg+' | '+(job.pdfIds||[]).length+' lampiran PDF | '+finalFile.getId());
    removePdfMergeJob_(job.qid);                                    // dihapus HANYA setelah seluruh langkah selesai
   }catch(err){
    job.attempts=Number(job.attempts||0)+1;job.lastError=String(err&&err.message||err).slice(0,500);
    logSecurityEvent_(job.nama,'PENGGABUNGAN PDF','GAGAL',job.reg+' | percobaan '+job.attempts+' | '+job.lastError);
    if(job.attempts<3){
     savePdfMergeJob_(job);
     setPdfJobStatus_(job,'processing','Percobaan '+job.attempts+' gagal, akan dicoba lagi: '+(job.lastError||''));
     needMore=true;
    }else{
     setPdfJobStatus_(job,'error',job.lastError||'Penggabungan PDF gagal.');
     if(job.jobId){if(job.cleanupFolder){try{DriveApp.getFolderById(job.folderId).setTrashed(true);}catch(cleanErr2){}}}
     else notifyMergeFailure_(job,job.lastError);
     removePdfMergeJob_(job.qid);
    }
   }
  }
 }finally{
  var finishLock=LockService.getScriptLock();
  try{finishLock.waitLock(20000);}catch(lockErr){console.warn('Kunci akhir antrean PDF tidak didapat: '+lockErr);}
  try{
   props.deleteProperty(PDF_MERGE_RUNNING_PROPERTY_);
   clearPdfMergeTriggers_();
   if(needMore||listPdfMergeJobs_().length)ensurePdfMergeTrigger_(budgetHit?60*1000:2*60*1000);
  }finally{try{finishLock.releaseLock();}catch(relErr){}}
 }
}


function sendApprovedContingentPdfEmailAfterMerge_(job,finalFile){
 if(!job||job.type!=='kontingen'||!finalFile)return {sent:false,reason:'Bukan PDF kontingen.'};
 var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss&&ss.getSheetByName(CFG.kontingen.sheet);if(!sh||sh.getLastRow()<2)return {sent:false,reason:'Data kontingen tidak ditemukan.'};
 var values=sh.getDataRange().getValues(),h=values[0],ix={};h.forEach(function(x,i){ix[String(x)]=i;});
 var row=null;
 for(var i=values.length-1;i>=1;i--)if(String(values[i][ix['No. Registrasi']]||'')===String(job.reg||'')){row=values[i];break;}
 if(!row||String(row[ix['Status Verifikasi']]||'')!=='Terverifikasi')return {sent:false,reason:'Belum terverifikasi.'};
 var email=String(row[ix['Email']]||'').trim().toLowerCase(),nama=String(row[ix['Nama Kontingen']]||'').trim(),ketua=String(row[ix['Ketua']]||'').trim();
 if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))return {sent:false,reason:'Email tidak valid.'};
 var log=ss.getSheetByName('Log Email')||ss.insertSheet('Log Email');
 if(log.getLastRow()===0){log.appendRow(['Waktu','Admin','Jenis Data','Baris Sheet','Nomor Registrasi','Nama Kontingen','Email Tujuan','Status Verifikasi','Subjek','Hasil']);log.setFrozenRows(1);log.getRange(1,1,1,10).setFontWeight('bold');}
 var logs=log.getDataRange().getValues(),lh=logs[0],ri=lh.indexOf('Nomor Registrasi'),si=lh.indexOf('Status Verifikasi'),hi=lh.indexOf('Hasil');
 if(ri>=0&&si>=0&&hi>=0&&logs.slice(1).some(function(r){return String(r[ri]||'')===String(job.reg||'')&&String(r[si]||'')==='Terverifikasi'&&String(r[hi]||'')==='Terkirim';}))return {sent:false,reason:'Email sudah terkirim.'};
 var subject='[PORSENI 2026] TERVERIFIKASI - '+nama+' ('+job.reg+')';
 var body='Assalamu’alaikum Wr. Wb.\n\nYth. '+(ketua?'Bapak/Ibu '+ketua+' selaku ':'')+'Ketua Kontingen '+nama+',\n\nAlhamdulillah, pendaftaran kontingen telah TERVERIFIKASI. Terima kasih atas kerja sama dan partisipasinya.\n\nBersama email ini kami lampirkan PDF formulir pendaftaran lengkap beserta seluruh lampiran yang telah digabungkan.\n\nSalam,\nPanitia PORSENI 2026\nPAC IPNU-IPPNU Kecamatan Patebon\n“Unitas Ad Gloriam”';
 try{
  if(MailApp.getRemainingDailyQuota()<1)throw new Error('Kuota email harian habis.');
  sendPorseniEmail_({to:email,subject:subject,body:body,attachments:[finalFile.getBlob().setName(finalFile.getName())],name:SENDER_NAME,replyTo:REPLY_TO});
  log.appendRow([new Date(),'Sistem (setelah PDF selesai)','kontingen','',job.reg,nama,email,'Terverifikasi',subject,'Terkirim']);
  return {sent:true,recipient:email};
 }catch(err){
  log.appendRow([new Date(),'Sistem (setelah PDF selesai)','kontingen','',job.reg,nama,email,'Terverifikasi',subject,'Gagal: '+String(err&&err.message||err)]);
  return {sent:false,reason:String(err&&err.message||err)};
 }
}

function updateMergedPdfRecord_(job,finalFile){
 if(String(job.reg||'').indexOf('-UJI-')>=0 || String(job.reg||'').indexOf('UJI-')===0)return;
 var cfg=job.type==='kontingen'?CFG.kontingen:CFG.cabang,ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss&&ss.getSheetByName(cfg.sheet);
 if(!sh||sh.getLastRow()<2)return;
 var headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0],iReg=headers.indexOf('No. Registrasi'),iPdf=headers.indexOf('PDF Formulir (link Drive)'),iArchive=headers.indexOf('Arsip PDF Panitia (link Drive)'),iFiles=headers.indexOf('Berkas (link Drive)');
 if(iReg<0||iPdf<0)return;
 var values=sh.getRange(2,1,sh.getLastRow()-1,headers.length).getValues(),target=-1;
 for(var i=values.length-1;i>=0;i--)if(String(values[i][iReg]||'')===String(job.reg)){target=i+2;break;}
 if(target<0)return;
 var url=finalFile.getUrl();sh.getRange(target,iPdf+1).setValue(url);
 if(iArchive>=0)sh.getRange(target,iArchive+1).setValue(url);
 if(iFiles>=0){var raw=String(sh.getRange(target,iFiles+1).getValue()||''),lines=raw.split(/\r?\n/);if(lines.length&&/^PDF formulir:/i.test(lines[0]))lines[0]='PDF formulir: '+url;else lines.unshift('PDF formulir: '+url);sh.getRange(target,iFiles+1).setValue(lines.join('\n'));}
}

function attachmentFileName_(d){
 var parts=['Lampiran'];
 if(d.type==='cabang')parts.push(pdfNamePart_(d.cabang||'Cabang Lomba'),pdfNamePart_(d.kategori||'Kategori'));
 parts.push(pdfNamePart_(d.nama||'Kontingen'));
 if(d.reg)parts.push(pdfNamePart_(d.reg));
 parts.push(pdfNamePart_(d.label||'Dokumen'));
 var original=String(d.originalName||'lampiran').replace(/[\\/:*?"<>|\r\n]+/g,'_').replace(/\s+/g,'_').slice(0,150)||'lampiran';
 return parts.join('_')+'_'+original;
}
function saveAttachments(files,folder,d){
 d.savedAttachments=[];
 return files.map(function(f){
  var bytes=Utilities.base64Decode(f.data);
  if(bytes.length>MAX_FILE_BYTES)throw new Error('Lampiran terlalu besar: '+f.name);
  var originalName=String(f.name||'lampiran').replace(/[\\/:*?"<>|\r\n]+/g,'_').slice(0,180)||'lampiran';
  var name=attachmentFileName_({type:d.type,nama:d.nama,reg:d.reg,cabang:d.cabang,kategori:d.kategori,label:f.label,originalName:originalName}),mime=detectedMime_(bytes,f.type);
  var blob=Utilities.newBlob(bytes,mime,name),file=folder.createFile(blob);
  assertPrivateIndividualFile_(file);
  d.savedAttachments.push({label:String(f.label||'Lampiran'),name:name,mime:mime,url:file.getUrl(),id:file.getId()});
  logSecurityEvent_(d.nama,'SIMPAN LAMPIRAN','BERHASIL',String(f.label||'Lampiran')+' | '+String(f.name||name));
  return (f.label||'Lampiran')+': '+file.getUrl();
 });
}
function appendRecord(c,d){var ss=SpreadsheetApp.getActiveSpreadsheet();if(!ss)throw new Error('Spreadsheet aktif tidak ditemukan. Pastikan Apps Script terikat ke database PORSENI.');var sh=ss.getSheetByName(c.sheet)||ss.insertSheet(c.sheet);
 if(sh.getLastRow()===0){sh.appendRow(c.head);sh.setFrozenRows(1);sh.getRange(1,1,1,c.head.length).setFontWeight('bold');}
 else {var existing=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];c.head.forEach(function(h){if(existing.indexOf(h)<0){sh.getRange(1,sh.getLastColumn()+1).setValue(h).setFontWeight('bold');existing.push(h);}});}
 var values=c.keys.map(function(k){var v=d[k];if(v===null||v===undefined)v='';if(k==='waktu')return v;v=String(v).slice(0,4000);if(v&&(/^[=+\-@]/.test(v)||PHONE_KEYS.indexOf(k)>-1))v="'"+v;return v;});
 // Susun ulang sesuai urutan header aktual agar sheet lama tetap kompatibel.
 var headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0],map={};c.keys.forEach(function(k,i){map[c.head[i]]=values[i];});
 if(d.type==='cabang'){
  var ni=headers.indexOf('Nama Kontingen'),bi=headers.indexOf('Cabang Lomba'),ki=headers.indexOf('Kategori'),si=headers.indexOf('Status Verifikasi'),target=-1;
  var lastDataRow=sh.getLastRow(),existingRows=lastDataRow>=2?sh.getRange(2,1,lastDataRow-1,headers.length).getValues():[];
  for(var r=existingRows.length-1;r>=0;r--){var old=existingRows[r];if(normalizeEvent_(old[ni])===normalizeEvent_(d.nama)&&normalizeEvent_(old[bi])===normalizeEvent_(d.cabang)&&normalizeEvent_(old[ki])===normalizeEvent_(d.kategori)){target=r+2;break;}}
  if(target>0){
   var existing=sh.getRange(target,1,1,headers.length).getValues()[0],status=si>=0?String(existing[si]||'Menunggu Verifikasi'):'Menunggu Verifikasi';
   if(status==='Terverifikasi')throw new Error('Berkas telah terverifikasi dan terkunci. Hubungi Sekretariat dan kirimkan berkas terbaru secara manual melalui WhatsApp.');
   var updated=headers.map(function(h,i){return map[h]!==undefined?map[h]:existing[i];});
   ['Status Verifikasi','Catatan Panitia','Diverifikasi Pada','Diverifikasi Oleh'].forEach(function(key){var ix=headers.indexOf(key);if(ix>=0)updated[ix]=key==='Status Verifikasi'?'Menunggu Verifikasi':'';});
   var idIx=headers.indexOf('ID Data');if(idIx>=0&&!updated[idIx])updated[idIx]=Utilities.getUuid();
   sh.getRange(target,1,1,updated.length).setValues([updated]);
   logSecurityEvent_(d.nama,'PERBARUI BERKAS CABANG','BERHASIL',d.cabang+' / '+d.kategori+' · status kembali Menunggu Verifikasi');
   return;
  }
 }
 var row=headers.map(function(h){return map[h]!==undefined?map[h]:'';});sh.appendRow(row);
}
function getFolderPath(path){var cur=DriveApp.getRootFolder();path.forEach(function(n){cur=getFolder(cur,n);});return cur;}
function getFolder(parent,name){var it=parent.getFoldersByName(name);return it.hasNext()?it.next():parent.createFolder(name);}
function safe(t){return String(t||'').replace(/[\\/:*?"<>|\n\r]+/g,' ').replace(/\s+/g,' ').trim().slice(0,120)||'Tanpa Nama';}
function notify(d){
 try{
  var kind=d.type==='kontingen'?'Pendaftaran kontingen':'Pengumpulan berkas cabang lomba';
  var body=kind+'\\nNomor Registrasi: '+d.reg+'\\nKontingen: '+d.nama+'\\nCabang: '+(d.cabang||'-')+'\\nKategori: '+(d.kategori||'-')+'\\nPDF: '+(d.pdfMergePending?'Sedang digabung otomatis. PDF final akan diperbarui pada Portal Kontingen setelah proses selesai.':d.pdf)+'\\nFolder: '+d.folder+'\\nRekap: '+SpreadsheetApp.getActiveSpreadsheet().getUrl();
  // Notifikasi internal panitia tetap dikirim. Email konfirmasi kepada kontingen tidak dikirim saat upload.
  sendPorseniEmail_({to:NOTIFY,subject:'[PORSENI 2026] '+kind+' - '+safe(d.nama)+' ('+d.reg+')',body:body,name:SENDER_NAME,replyTo:REPLY_TO});
 }catch(err){console.error(err);}
}
function out(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}


function porseniBackendHealthCheckV15(){
 var ss=SpreadsheetApp.getActiveSpreadsheet();
 return {ok:true,backendVersion:'PORSENI-2026-CONTROL-CENTER-V40-20261111',kontingenCount:KONTINGEN.length,branchCount:Object.keys(KAT).length,pinMaxAttempts:PIN_MAX_ATTEMPTS,pinWindowMinutes:PIN_WINDOW_MS/60000,pinLockMinutes:PIN_LOCK_MS/60000,contingentSessionSeconds:CONTINGENT_SESSION_SECONDS,securityLogSheet:SECURITY_LOG_SHEET,spreadsheet:!!ss,pdfMergeLibrary:typeof PDFApp!=='undefined'&&!!PDFApp&&typeof PDFApp.mergePDFs==='function'};
}

/* ===================================================================
   V31 — EDIT DATA KONTINGEN, ACC PERUBAHAN, RESET PIN, UJI CETAK PDF
   =================================================================== */
var EDITABLE_FIELDS_ = {
  ketua:              {head:'Ketua',              label:'Nama ketua kontingen',  type:'name'},
  tempat_lahir_ketua: {head:'Tempat Lahir Ketua', label:'Tempat lahir ketua',    type:'place'},
  tanggal_lahir_ketua:{head:'Tanggal Lahir Ketua',label:'Tanggal lahir ketua',   type:'date'},
  alamat:             {head:'Alamat',             label:'Alamat rumah ketua',    type:'address'},
  wa_ketua:           {head:'WA Ketua',           label:'No. WA ketua kontingen',type:'phone'},
  email:              {head:'Email',              label:'Email kontingen',       type:'email'},
  ketua_pr_ipnu:      {head:'Ketua PR IPNU',      label:'Nama Ketua IPNU',       type:'name'},
  wa_ketua_ipnu:      {head:'WA Ketua IPNU',      label:'No. WA Ketua IPNU',     type:'phone'},
  ketua_pr_ippnu:     {head:'Ketua PR IPPNU',     label:'Nama Ketua IPPNU',      type:'name'},
  wa_ketua_ippnu:     {head:'WA Ketua IPPNU',     label:'No. WA Ketua IPPNU',    type:'phone'}
};
var KONTINGEN_WINDOW_ = [new Date(2026,10,1,0,0,0), new Date(2026,10,20,23,59,59)];

function assertKontingenWindow_(){
  var now=new Date();
  if(now<KONTINGEN_WINDOW_[0])throw new Error('Pendaftaran kontingen dibuka mulai 1 November 2026.');
  if(now>KONTINGEN_WINDOW_[1])throw new Error('Pendaftaran kontingen sudah ditutup pada 20 November 2026.');
}
function assertKontingenChangeWindow_(name,changes){
  var now=new Date(),f=findKontingenRow_(name),raw=String(f.values[f.idx['Cabang Lomba']]||'');
  var branches=raw.split(/;\s*/).map(function(x){return x.replace(/\s*\(.*/, '').trim();}).filter(Boolean);
  var deadlines=branches.map(function(b){var dt;if(b==='Badminton')dt=new Date(2026,10,29,23,59,59);else if(b==='Duta Pelajar'||b==='Voli')dt=new Date(2026,11,6,23,59,59);else dt=new Date(2026,11,19,23,59,59);return {branch:b,date:dt};});
  var deadline=deadlines.length?deadlines.reduce(function(a,b){return a.date<b.date?a:b;}).date:new Date(2026,11,19,23,59,59);
  if(now>deadline)throw new Error('Batas perubahan data kegiatan kontingen telah berakhir. Batas yang berlaku: '+Utilities.formatDate(deadline,'Asia/Jakarta','dd MMMM yyyy')+'.');
}
function normPhone_(v){var s=String(v||'').replace(/[\s\-().]/g,'');if(/^\+62/.test(s))s='0'+s.slice(3);else if(/^62/.test(s))s='0'+s.slice(2);return s;}
function cleanKontingenField_(key,val){
  var f=EDITABLE_FIELDS_[key];if(!f)throw new Error('Kolom "'+key+'" tidak dapat diubah.');
  var v=String(val==null?'':val).replace(/[\u200B-\u200D\uFEFF]/g,'').replace(/\s+/g,' ').trim();
  if(!v)throw new Error(f.label+' wajib diisi.');
  if(f.type==='name'&&(v.length<3||v.length>80||!/[A-Za-z]/.test(v)))throw new Error(f.label+' tidak valid (minimal 3 karakter, memuat huruf).');
  if(f.type==='place'&&(v.length<3||v.length>60||!/[A-Za-z]/.test(v)))throw new Error(f.label+' tidak valid (minimal 3 karakter, memuat huruf).');
  if(f.type==='address'&&(v.length<10||v.length>250))throw new Error(f.label+' harus lengkap (minimal 10 karakter).');
  if(f.type==='phone'){v=normPhone_(v);if(!/^08\d{8,12}$/.test(v))throw new Error(f.label+' tidak valid. Gunakan format 08xxxxxxxxxx.');}
  if(f.type==='email'){v=v.toLowerCase();if(v.length>120||!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v))throw new Error(f.label+' tidak valid.');}
  if(f.type==='date'){
    var m=v.match(/^(\d{4})-(\d{2})-(\d{2})$/);if(!m)throw new Error(f.label+' tidak valid.');
    var dt=new Date(Number(m[1]),Number(m[2])-1,Number(m[3]));
    if(dt.getMonth()!==Number(m[2])-1||dt>new Date()||Number(m[1])<1950)throw new Error(f.label+' tidak valid.');
  }
  return v;
}
function sniffFileKind_(bytes){
  if(!bytes||bytes.length<12)return '';
  var b=function(i){return bytes[i]&255;};
  if(b(0)===37&&b(1)===80&&b(2)===68&&b(3)===70)return 'pdf';
  if(b(0)===255&&b(1)===216&&b(2)===255)return 'jpg';
  if(b(0)===137&&b(1)===80&&b(2)===78&&b(3)===71)return 'png';
  if(b(0)===82&&b(1)===73&&b(2)===70&&b(3)===70&&b(8)===87&&b(9)===69&&b(10)===66&&b(11)===80)return 'webp';
  if(b(4)===102&&b(5)===116&&b(6)===121&&b(7)===112)return 'heic';
  return '';
}
function assertFileKind_(bytes,allowed,label){
  var kind=sniffFileKind_(bytes);
  if(allowed.indexOf(kind)<0)throw new Error('Format file "'+label+'" tidak valid. Setelah konversi, berkas gambar harus berupa JPG atau PNG. Silakan unggah ulang atau ekspor gambar sebagai JPG/PNG.');
}
function newRegFor_(type){
  var c='ABCDEFGHJKLMNPQRSTUVWXYZ23456789',r='';for(var i=0;i<4;i++)r+=c.charAt(Math.floor(Math.random()*c.length));
  return (type==='kontingen'?'K':'C')+'-'+Utilities.formatDate(new Date(),'Asia/Jakarta','yyMMdd')+'-'+r;
}

/* ---------- baca / tulis data kontingen ---------- */
function findKontingenRow_(name){
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.kontingen.sheet);
  if(!sh||sh.getLastRow()<2)throw new Error('Data pendaftaran kontingen belum ditemukan.');
  var v=sh.getDataRange().getValues(),h=v[0],idx={};h.forEach(function(x,i){idx[String(x)]=i;});
  for(var i=v.length-1;i>=1;i--){if(normalizeEvent_(v[i][idx['Nama Kontingen']])===normalizeEvent_(name))return {sheet:sh,row:i+1,values:v[i],headers:h,idx:idx};}
  throw new Error('Data pendaftaran kontingen belum ditemukan.');
}
function readKontingenData_(name){
  var f=findKontingenRow_(name),out={};
  Object.keys(EDITABLE_FIELDS_).forEach(function(k){
    var def=EDITABLE_FIELDS_[k],raw=f.values[f.idx[def.head]];
    if(raw instanceof Date)raw=Utilities.formatDate(raw,'Asia/Jakarta','yyyy-MM-dd');
    raw=String(raw==null?'':raw);
    if(def.type==='phone'&&/^8\d{8,12}$/.test(raw))raw='0'+raw;
    out[k]=raw;
  });
  return out;
}
function formatSheetPhone_(v){return "'"+v;}
function applyKontingenChanges_(name,changes,actor){
  if(!changes||typeof changes!=='object')throw new Error('Tidak ada perubahan data.');
  var clean={};Object.keys(changes).forEach(function(k){clean[k]=cleanKontingenField_(k,changes[k]);});
  if(!Object.keys(clean).length)throw new Error('Tidak ada perubahan data.');
  var lock=LockService.getScriptLock(),diffs=[];
  lock.waitLock(30000);
  try{
    var f=findKontingenRow_(name),current=readKontingenData_(name);
    Object.keys(clean).forEach(function(k){if(String(current[k]||'')!==clean[k])diffs.push(k);});
    if(!diffs.length)throw new Error('Data yang dikirim sama dengan data saat ini.');
    diffs.forEach(function(k){
      var def=EDITABLE_FIELDS_[k],col=f.idx[def.head];
      if(col===undefined){col=f.headers.length;f.sheet.getRange(1,col+1).setValue(def.head).setFontWeight('bold');f.headers.push(def.head);f.idx[def.head]=col;}
      f.sheet.getRange(f.row,col+1).setValue(def.type==='phone'?formatSheetPhone_(clean[k]):(/^[=+\-@]/.test(clean[k])?"'"+clean[k]:clean[k]));
    });
    if(Object.prototype.hasOwnProperty.call(clean,'cabang'))invalidatePublicCache_('recap');
  }finally{lock.releaseLock();}
  var pdfInfo=rebuildKontingenPdf_(name,actor);
  logSecurityEvent_(name,'PERUBAHAN DATA KONTINGEN','BERHASIL',(actor||'Kontingen')+' | '+diffs.map(function(k){return EDITABLE_FIELDS_[k].label;}).join(', '));
  return {changed:diffs,pdf:pdfInfo};
}
// Buat ulang PDF formulir dari data terbaru (tanda tangan diambil dari berkas yang tersimpan di Drive).
function rebuildKontingenPdf_(name,actor){
  var f=findKontingenRow_(name),g=function(h){return f.values[f.idx[h]];},cur=readKontingenData_(name);
  var oldPdfUrl=String(g('PDF Formulir (link Drive)')||''),oldId=getDriveFileIdFromUrl_(oldPdfUrl),oldArchiveId=getDriveFileIdFromUrl_(String(g('Arsip PDF Panitia (link Drive)')||'')),folder;
  try{folder=oldId?DriveApp.getFileById(oldId).getParents().next():getFolderPath(DRIVE_PATH.concat([CFG.kontingen.folder]));}catch(e){folder=getFolderPath(DRIVE_PATH.concat([CFG.kontingen.folder]));}
  var originalRefs=parsePortalFiles_(g('Berkas (link Drive)')||'').filter(function(x){return !/^pdf formulir/i.test(x.label);});
  var files=[],saved=[];
  originalRefs.forEach(function(x){
    try{
      var id=getDriveFileIdFromUrl_(x.url);if(!id)return;
      var file=DriveApp.getFileById(id);if(file.isTrashed())return;
      var blob=file.getBlob(),mime=String(blob.getContentType()||file.getMimeType()||'application/octet-stream').toLowerCase();
      if(mime==='application/pdf')saved.push({id:id,label:x.label,name:file.getName(),mime:mime,url:x.url});
      else if(mime==='image/png'||mime==='image/jpeg')files.push({label:x.label,name:file.getName(),type:mime,data:Utilities.base64Encode(blob.getBytes())});
    }catch(readErr){console.warn('Lampiran asli tidak dapat dibaca saat regenerasi PDF: '+x.label+' | '+readErr);}
  });
  var d={type:'kontingen',reg:String(g('No. Registrasi')||''),nama:name,pin:'',cabang:String(g('Cabang Lomba')||''),waktu:g('Waktu')||new Date(),
    ketua:cur.ketua,tempat_lahir_ketua:cur.tempat_lahir_ketua,tanggal_lahir_ketua:cur.tanggal_lahir_ketua,alamat:cur.alamat,wa_ketua:cur.wa_ketua,email:cur.email,
    ketua_pr_ipnu:cur.ketua_pr_ipnu,ketua_pr_ippnu:cur.ketua_pr_ippnu,files:files,savedAttachments:saved,diperbarui:idDate_(new Date(),true)};
  var pdf=createRegistrationPdf(d,folder);assertPrivateIndividualFile_(pdf);
  function setCol(h,val){var c=f.idx[h];if(c!==undefined)f.sheet.getRange(f.row,c+1).setValue(val);}
  setCol('PDF Formulir (link Drive)',pdf.getUrl());
  var berkas=String(g('Berkas (link Drive)')||'').split(/\r?\n/);
  if(berkas.length&&/^PDF formulir:/i.test(berkas[0]))berkas[0]='PDF formulir: '+pdf.getUrl();else berkas.unshift('PDF formulir: '+pdf.getUrl());
  setCol('Berkas (link Drive)',berkas.join('\n'));
  var oldIds=[oldId,oldArchiveId].filter(function(id,i,a){return id&&a.indexOf(id)===i&&id!==pdf.getId();});
  var queued=false;
  if(saved.length){queued=!!queuePdfMerge_(pdf,d,folder,{oldPdfIds:oldIds,fileName:pdfName(d)});}
  if(!queued){
    var archive=archivePrintablePdf_(pdf,d);setCol('Arsip PDF Panitia (link Drive)',archive.getUrl());
    oldIds.forEach(function(id){try{DriveApp.getFileById(id).setTrashed(true);}catch(cleanErr){console.warn('PDF lama tidak dapat dipindahkan ke Trash: '+cleanErr);}});
  }
  return {url:pdf.getUrl(),mergePending:queued};
}

/* ---------- kontingen: ubah data langsung (sebelum verifikasi) ---------- */
function contingentUpdateData_(d){
  var session=requireContingentSession_(d.sessionToken),name=session.nama;assertKontingenChangeWindow_(name,d.changes);if(KONTINGEN.indexOf(name)<0)throw new Error('Nama kontingen tidak dikenal.');
  var approval=getContingentApproval_(name);
  if(approval.status==='Belum Terdaftar')throw new Error('Data kontingen tidak ditemukan.');
  if(approval.status==='Terverifikasi')throw new Error('Data sudah terverifikasi dan terkunci. Ajukan permohonan perubahan data kepada panitia.');
  var res=applyKontingenChanges_(name,d.changes,'Kontingen');
  return {ok:true,message:'Data kontingen diperbarui ('+res.changed.length+' kolom). PDF formulir dibuat ulang.',data:readKontingenData_(name)};
}
/* ---------- kontingen: ajukan perubahan data (setelah verifikasi) ---------- */
function contingentRequestDataChange_(d){
  var session=requireContingentSession_(d.sessionToken),name=session.nama;assertKontingenChangeWindow_(name,d.changes);if(KONTINGEN.indexOf(name)<0)throw new Error('Nama kontingen tidak dikenal.');
  var approval=getContingentApproval_(name);
  if(approval.status!=='Terverifikasi')throw new Error('Data belum terverifikasi. Ubah data langsung melalui portal.');
  var alasan=String(d.alasan||'').trim().slice(0,1000);if(alasan.length<10)throw new Error('Alasan perubahan wajib diisi (minimal 10 karakter).');
  var current=readKontingenData_(name),clean={},keys=Object.keys(d.changes||{});
  if(!keys.length)throw new Error('Isi minimal satu data yang ingin diubah.');
  keys.forEach(function(k){var v=cleanKontingenField_(k,d.changes[k]);if(String(current[k]||'')!==v)clean[k]=v;});
  if(!Object.keys(clean).length)throw new Error('Nilai baru sama dengan data saat ini.');
  var cache=CacheService.getScriptCache(),stateKey='change_req_'+pinAttemptKey_(name);
  if(cache.get(stateKey))throw new Error('Permohonan baru saja dikirim. Tunggu 5 menit sebelum mengirim lagi.');
  var sh=ensureChangeRequestSheet_(),h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0],f=findKontingenRow_(name);
  var ticket='PRM-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd-HHmmss')+'-'+Utilities.getUuid().slice(0,4).toUpperCase();
  var labels=Object.keys(clean).map(function(k){return EDITABLE_FIELDS_[k].label;}).join(', ');
  var map={'Waktu':new Date(),'No. Tiket':ticket,'Nama Kontingen':name,'No. Registrasi':String(f.values[f.idx['No. Registrasi']]||''),'Jenis Permohonan':'Data kontingen','Cabang / Kategori':'','Data / Berkas':labels,'Alasan':alasan,'Status':'Menunggu Tindak Lanjut','Perubahan Diminta':JSON.stringify(clean)};
  sh.appendRow(h.map(function(x){return map[x]!==undefined?map[x]:'';}));
  cache.put(stateKey,'1',300);logSecurityEvent_(name,'PERMOHONAN PERUBAHAN DATA','BERHASIL',ticket+' | '+labels);
  try{if(MailApp.getRemainingDailyQuota()>0)sendPorseniEmail_({to:NOTIFY,subject:'[PORSENI 2026] Permohonan perubahan data '+ticket,body:'Permohonan perubahan data dari '+name+'.\nData: '+labels+'\nAlasan: '+alasan+'\n\nSilakan setujui atau tolak di Portal Panitia.',name:SENDER_NAME,replyTo:REPLY_TO});}catch(e){}
  return {ok:true,ticket:ticket,message:'Permohonan perubahan data diterima. Menunggu persetujuan panitia.'};
}
/* ---------- panitia: setujui & terapkan permohonan perubahan data ---------- */
function adminApplyChangeRequest_(d){
  var admin=requireAdmin_(d.token),ticket=String(d.ticket||'').trim(),note=String(d.note||'').trim().slice(0,1000);
  var sh=ensureChangeRequestSheet_(),v=sh.getDataRange().getValues(),h=v[0],ix={};h.forEach(function(x,i){ix[String(x)]=i;});
  var row=-1;for(var i=1;i<v.length;i++){if(String(v[i][ix['No. Tiket']]||'')===ticket){row=i;break;}}
  if(row<0)throw new Error('Nomor tiket tidak ditemukan.');
  var r=v[row];if(String(r[ix['Jenis Permohonan']])!=='Data kontingen')throw new Error('Hanya permohonan data kontingen yang dapat diterapkan otomatis.');
  if(['Selesai','Tidak Dapat Dipenuhi'].indexOf(String(r[ix['Status']]))>=0)throw new Error('Permohonan ini sudah diproses.');
  var changes;try{changes=JSON.parse(String(r[ix['Perubahan Diminta']]||'{}'));}catch(e){throw new Error('Isi perubahan tidak dapat dibaca.');}
  var name=String(r[ix['Nama Kontingen']]||'');
  var res=applyKontingenChanges_(name,changes,'Panitia: '+admin.user);
  var n=row+1;
  sh.getRange(n,ix['Status']+1).setValue('Selesai');sh.getRange(n,ix['Catatan Panitia']+1).setValue(note||'Disetujui dan diterapkan.');
  sh.getRange(n,ix['Diproses Oleh']+1).setValue(admin.user);sh.getRange(n,ix['Diproses Pada']+1).setValue(new Date());
  adminAudit_(admin,'SETUJUI PERUBAHAN DATA','permohonan',n,ticket,res.changed.join(', '));
  return {ok:true,message:'Perubahan disetujui dan diterapkan. PDF formulir dibuat ulang.'};
}
/* ---------- panitia: reset PIN kontingen ---------- */
function adminResetPin_(d){
  var admin=requireAdmin_(d.token),name=canonicalContingentName_(d.nama);
  if(!name||KONTINGEN.indexOf(name)<0)throw new Error('Kontingen tidak dikenal.');
  findKontingenRow_(name);
  var pin=createOrGetContingentPin_(name);
  clearPinFailures_(name);
  adminAudit_(admin,'RESET PIN KONTINGEN','kontingen','',name,'PIN baru dibuat (tidak dicatat).');
  logSecurityEvent_(name,'RESET PIN','BERHASIL','Oleh '+admin.user);
  return {ok:true,pin:pin,message:'PIN baru dibuat. Sampaikan hanya kepada ketua kontingen. PIN lama tidak berlaku lagi.'};
}

/* ---------- UJI CETAK PDF ---------- */
var SAMPLE_ATTACHMENT_B64_ = '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';
function testFolder_(){return getFolderPath(DRIVE_PATH.concat(['UJI COBA']));}
function createTestImagesPdf_(files,tempFolder,ownerName){
 // Seluruh gambar contoh dikonversi melalui SATU Google Docs/PDF per jenis uji,
 // bukan membuat satu dokumen dan satu konversi PDF untuk setiap gambar.
 if(!files||!files.length)return null;
 var doc=DocumentApp.create('TEMP IMAGE PAGES '+Utilities.getUuid());
 var body=doc.getBody();
 body.setPageWidth(595.28).setPageHeight(841.89).setMarginTop(42).setMarginBottom(42).setMarginLeft(42).setMarginRight(42);
 files.forEach(function(file,index){
  var bytes=Utilities.base64Decode(file.data),mime=detectedMime_(bytes,file.type||'');
  if(mime.indexOf('image/')!==0)throw new Error('Lampiran uji bukan gambar: '+(file.name||''));
  if(index>0)body.appendPageBreak();
  body.appendParagraph(String(file.label||file.name||'Lampiran')+(ownerName?' · '+String(ownerName):'')).setBold(true).setFontSize(12);
  var image=body.appendImage(Utilities.newBlob(bytes,mime,file.name||'lampiran'));
  var maxW=510,maxH=700,ratio=Math.min(maxW/image.getWidth(),maxH/image.getHeight(),1);
  image.setWidth(Math.max(1,Math.round(image.getWidth()*ratio))).setHeight(Math.max(1,Math.round(image.getHeight()*ratio)));
 });
 doc.saveAndClose();
 var docFile=DriveApp.getFileById(doc.getId());
 try{
  var pdfName='IMG_LAMPIRAN_UJI_'+Utilities.getUuid().slice(0,8)+'.pdf';
  return tempFolder.createFile(docFile.getAs(MimeType.PDF).setName(pdfName));
 }finally{docFile.setTrashed(true);}
}
function buildTestPdfJob_(d){
 if(typeof PDFApp==='undefined'||!PDFApp||typeof PDFApp.mergePDFs!=='function')throw new Error('Library PDFApp belum terpasang. Tambahkan PDFApp dengan fungsi mergePDFs sebelum menjalankan uji cetak.');
 var outputFolder=testFolder_(),tempFolder=DriveApp.createFolder('TEMP_PORSENI_UJI_'+Utilities.getUuid()),saved=[];
 try{
  // Uji coba wajib mengikuti jalur produksi: seluruh gambar diproses hanya oleh
  // createRegistrationPdf()/createBranchRegistrationPdf_(). Jangan konversi ulang
  // gambar menjadi PDF terpisah karena gambar cabang sudah ditata di halaman formulir.
  var formPdf=createRegistrationPdf(d,tempFolder);

  // Hanya lampiran PDF asli yang disimpan untuk digabung setelah formulir,
  // sama seperti jalur pendaftaran sebenarnya. Lampiran gambar tidak digabung ulang.
  (d.files||[]).forEach(function(f){
   if(!f||!f.data)return;
   // Lewati gambar sebelum base64Decode: gambar sudah diproses satu kali oleh
   // mesin formulir produksi. Decode hanya PDF yang benar-benar perlu digabung.
   var declaredMime=String(f.type||'').toLowerCase().split(';')[0].trim();
   if(declaredMime==='image/png'||declaredMime==='image/jpeg'||declaredMime==='image/jpg')return;
   var bytes=Utilities.base64Decode(f.data),mime=detectedMime_(bytes,declaredMime);
   if(mime==='application/pdf'){
    var pdfFile=tempFolder.createFile(Utilities.newBlob(bytes,'application/pdf','SRC_'+safe(f.name||'lampiran.pdf')));
    saved.push({label:f.label||f.name,id:pdfFile.getId(),mime:'application/pdf'});
   }else{
    throw new Error('Format lampiran uji tidak didukung untuk PDF: '+(f.name||mime));
   }
  });

  d.savedAttachments=saved;
  var jobId=Utilities.getUuid();
  // jobId tetap memaksa finalisasi ke folder UJI COBA, termasuk ketika tidak ada
  // lampiran PDF. Ini hanya alur penyimpanan hasil uji; isi PDF tetap mengikuti produksi.
  queuePdfMerge_(formPdf,d,tempFolder,{outputFolderId:outputFolder.getId(),cleanupFolder:true,jobId:jobId,fileName:'UJI_'+safe(d.reg)+'_'+safe(d.nama)+'.pdf'});
  return {jobId:jobId,kind:d.cabang==='Duta Pelajar'?'duta':(/^Beregu/i.test(d.kategori||'')?'beregu':(d.type==='kontingen'?'kontingen':'tunggal'))};
 }catch(err){try{tempFolder.setTrashed(true);}catch(cleanErr){}throw err;}
}

function sampleImageAttachment_(label,name){return {label:label,data:SAMPLE_ATTACHMENT_B64_,type:'image/png',name:name||'contoh-lampiran.png'};}
function ujiCetakPdfKontingen(){
 var d={type:'kontingen',reg:'K-UJI-'+Utilities.getUuid().slice(0,6),nama:'PR IPNU IPPNU CONTOH UJI COBA',pin:'123456',waktu:new Date().toISOString(),ketua:'Ahmad Contoh',tempat_lahir_ketua:'Kendal',tanggal_lahir_ketua:'2006-08-17',alamat:'Jl. Contoh No. 1 RT 02 RW 03, Desa Contoh, Patebon, Kendal',wa_ketua:'081234567890',email:'contoh@example.com',ketua_pr_ipnu:'Ahmad Contoh',ketua_pr_ippnu:'Siti Contoh',cabang:'Duta Pelajar; LCC; Voli; Parade Gelar Budaya',files:[sampleImageAttachment_('Bukti pembayaran','bukti-pembayaran.png'),sampleImageAttachment_('KTP ketua kontingen','ktp-ketua.png'),{label:'Pakta integritas',name:'pakta-integritas.pdf',type:'application/pdf',data:'JVBERi0xLjMKJZOMi54gUmVwb3J0TGFiIEdlbmVyYXRlZCBQREYgZG9jdW1lbnQgKG9wZW5zb3VyY2UpCjEgMCBvYmoKPDwKL0YxIDIgMCBSCj4+CmVuZG9iagoyIDAgb2JqCjw8Ci9CYXNlRm9udCAvSGVsdmV0aWNhIC9FbmNvZGluZyAvV2luQW5zaUVuY29kaW5nIC9OYW1lIC9GMSAvU3VidHlwZSAvVHlwZTEgL1R5cGUgL0ZvbnQKPj4KZW5kb2JqCjMgMCBvYmoKPDwKL0NvbnRlbnRzIDcgMCBSIC9NZWRpYUJveCBbIDAgMCAyMDAgMTIwIF0gL1BhcmVudCA2IDAgUiAvUmVzb3VyY2VzIDw8Ci9Gb250IDEgMCBSIC9Qcm9jU2V0IFsgL1BERiAvVGV4dCAvSW1hZ2VCIC9JbWFnZUMgL0ltYWdlSSBdCj4+IC9Sb3RhdGUgMCAvVHJhbnMgPDwKCj4+IAogIC9UeXBlIC9QYWdlCj4+CmVuZG9iago0IDAgb2JqCjw8Ci9QYWdlTW9kZSAvVXNlTm9uZSAvUGFnZXMgNiAwIFIgL1R5cGUgL0NhdGFsb2cKPj4KZW5kb2JqCjUgMCBvYmoKPDwKL0F1dGhvciAoYW5vbnltb3VzKSAvQ3JlYXRpb25EYXRlIChEOjIwMjYxMTA2MDUyMTEyKzAwJzAwJykgL0NyZWF0b3IgKGFub255bW91cykgL0tleXdvcmRzICgpIC9Nb2REYXRlIChEOjIwMjYxMTA2MDUyMTEyKzAwJzAwJykgL1Byb2R1Y2VyIChSZXBvcnRMYWIgUERGIExpYnJhcnkgLSBcKG9wZW5zb3VyY2VcKSkgCiAgL1N1YmplY3QgKHVuc3BlY2lmaWVkKSAvVGl0bGUgKHVudGl0bGVkKSAvVHJhcHBlZCAvRmFsc2UKPj4KZW5kb2JqCjYgMCBvYmoKPDwKL0NvdW50IDEgL0tpZHMgWyAzIDAgUiBdIC9UeXBlIC9QYWdlcwo+PgplbmRvYmoKNyAwIG9iago8PAovRmlsdGVyIFsgL0FTQ0lJODVEZWNvZGUgL0ZsYXRlRGVjb2RlIF0gL0xlbmd0aCAxNTIKPj4Kc3RyZWFtCkdhcFFoMEU9RiwwVVxIM1RccE5ZVF5RS2s/dGM+SVAsO1cjVTFeMjNpaFBFTV8/Q1c0S0lTaTc6cDtXTSNnJClubjRDaFVhPUtzLTtbQD5eO1tDWGo8LGcsWThsUSljJWdtTFFvKXUnTDE9SzIuYVxpRGxfREsoO0BZLipaOjE+aiZKRTZYQFA1R182ZUkjT11RK0E2TH4+ZW5kc3RyZWFtCmVuZG9iagp4cmVmCjAgOAowMDAwMDAwMDAwIDY1NTM1IGYgCjAwMDAwMDAwNjEgMDAwMDAgbiAKMDAwMDAwMDA5MiAwMDAwMCBuIAowMDAwMDAwMTk5IDAwMDAwIG4gCjAwMDAwMDAzOTIgMDAwMDAgbiAKMDAwMDAwMDQ2MCAwMDAwMCBuIAowMDAwMDAwNzIxIDAwMDAwIG4gCjAwMDAwMDA3ODAgMDAwMDAgbiAKdHJhaWxlcgo8PAovSUQgCls8NGFiZWMwMDg2ZGZlNjQ2YjM4YjRkZTI2NGI2Njc4ODM+PDRhYmVjMDA4NmRmZTY0NmIzOGI0ZGUyNjRiNjY3ODgzPl0KJSBSZXBvcnRMYWIgZ2VuZXJhdGVkIFBERiBkb2N1bWVudCAtLSBkaWdlc3QgKG9wZW5zb3VyY2UpCgovSW5mbyA1IDAgUgovUm9vdCA0IDAgUgovU2l6ZSA4Cj4+CnN0YXJ0eHJlZgoxMDIyCiUlRU9GCg=='} ,{label:'SP / berita acara RAT',name:'sp-rat.pdf',type:'application/pdf',data:'JVBERi0xLjMKJZOMi54gUmVwb3J0TGFiIEdlbmVyYXRlZCBQREYgZG9jdW1lbnQgKG9wZW5zb3VyY2UpCjEgMCBvYmoKPDwKL0YxIDIgMCBSCj4+CmVuZG9iagoyIDAgb2JqCjw8Ci9CYXNlRm9udCAvSGVsdmV0aWNhIC9FbmNvZGluZyAvV2luQW5zaUVuY29kaW5nIC9OYW1lIC9GMSAvU3VidHlwZSAvVHlwZTEgL1R5cGUgL0ZvbnQKPj4KZW5kb2JqCjMgMCBvYmoKPDwKL0NvbnRlbnRzIDcgMCBSIC9NZWRpYUJveCBbIDAgMCAyMDAgMTIwIF0gL1BhcmVudCA2IDAgUiAvUmVzb3VyY2VzIDw8Ci9Gb250IDEgMCBSIC9Qcm9jU2V0IFsgL1BERiAvVGV4dCAvSW1hZ2VCIC9JbWFnZUMgL0ltYWdlSSBdCj4+IC9Sb3RhdGUgMCAvVHJhbnMgPDwKCj4+IAogIC9UeXBlIC9QYWdlCj4+CmVuZG9iago0IDAgb2JqCjw8Ci9QYWdlTW9kZSAvVXNlTm9uZSAvUGFnZXMgNiAwIFIgL1R5cGUgL0NhdGFsb2cKPj4KZW5kb2JqCjUgMCBvYmoKPDwKL0F1dGhvciAoYW5vbnltb3VzKSAvQ3JlYXRpb25EYXRlIChEOjIwMjYxMTA2MDUyMTEyKzAwJzAwJykgL0NyZWF0b3IgKGFub255bW91cykgL0tleXdvcmRzICgpIC9Nb2REYXRlIChEOjIwMjYxMTA2MDUyMTEyKzAwJzAwJykgL1Byb2R1Y2VyIChSZXBvcnRMYWIgUERGIExpYnJhcnkgLSBcKG9wZW5zb3VyY2VcKSkgCiAgL1N1YmplY3QgKHVuc3BlY2lmaWVkKSAvVGl0bGUgKHVudGl0bGVkKSAvVHJhcHBlZCAvRmFsc2UKPj4KZW5kb2JqCjYgMCBvYmoKPDwKL0NvdW50IDEgL0tpZHMgWyAzIDAgUiBdIC9UeXBlIC9QYWdlcwo+PgplbmRvYmoKNyAwIG9iago8PAovRmlsdGVyIFsgL0FTQ0lJODVEZWNvZGUgL0ZsYXRlRGVjb2RlIF0gL0xlbmd0aCAxNTIKPj4Kc3RyZWFtCkdhcFFoMEU9RiwwVVxIM1RccE5ZVF5RS2s/dGM+SVAsO1cjVTFeMjNpaFBFTV8/Q1c0S0lTaTc6cDtXTSNnJClubjRDaFVhPUtzLTtbQD5eO1tDWGo8LGcsWThsUSljJWdtTFFvKXUnTDE9SzIuYVxpRGxfREsoO0BZLipaOjE+aiZKRTZYQFA1R182ZUkjT11RK0E2TH4+ZW5kc3RyZWFtCmVuZG9iagp4cmVmCjAgOAowMDAwMDAwMDAwIDY1NTM1IGYgCjAwMDAwMDAwNjEgMDAwMDAgbiAKMDAwMDAwMDA5MiAwMDAwMCBuIAowMDAwMDAwMTk5IDAwMDAwIG4gCjAwMDAwMDAzOTIgMDAwMDAgbiAKMDAwMDAwMDQ2MCAwMDAwMCBuIAowMDAwMDAwNzIxIDAwMDAwIG4gCjAwMDAwMDA3ODAgMDAwMDAgbiAKdHJhaWxlcgo8PAovSUQgCls8NGFiZWMwMDg2ZGZlNjQ2YjM4YjRkZTI2NGI2Njc4ODM+PDRhYmVjMDA4NmRmZTY0NmIzOGI0ZGUyNjRiNjY3ODgzPl0KJSBSZXBvcnRMYWIgZ2VuZXJhdGVkIFBERiBkb2N1bWVudCAtLSBkaWdlc3QgKG9wZW5zb3VyY2UpCgovSW5mbyA1IDAgUgovUm9vdCA0IDAgUgovU2l6ZSA4Cj4+CnN0YXJ0eHJlZgoxMDIyCiUlRU9GCg=='}]};
 return buildTestPdfJob_(d);
}
function ujiCetakPdfDutaPelajar(){
 var list=[{nama:'Aisyah Contoh',gender:'P',tempatLahir:'Kendal',tanggalLahir:'2008-04-12',alamat:'Desa Contoh, Patebon',wa:'081234567891'}];
 var d={type:'cabang',reg:'DUTA-UJI-'+Utilities.getUuid().slice(0,6),nama:'PR IPNU IPPNU CONTOH UJI COBA',cabang:'Duta Pelajar',kategori:'Putri',jml:1,email:'contoh@example.com',waktu:new Date().toISOString(),pesertaList:list,catatan:'Uji format Duta Pelajar dengan formulir cabang lomba umum.',files:[sampleImageAttachment_('Pas foto peserta 1','foto-duta.png'),sampleImageAttachment_('KTP peserta 1','ktp-duta.png'),{label:'CV Duta Pelajar',name:'cv-duta.pdf',type:'application/pdf',data:'JVBERi0xLjMKJZOMi54gUmVwb3J0TGFiIEdlbmVyYXRlZCBQREYgZG9jdW1lbnQgKG9wZW5zb3VyY2UpCjEgMCBvYmoKPDwKL0YxIDIgMCBSCj4+CmVuZG9iagoyIDAgb2JqCjw8Ci9CYXNlRm9udCAvSGVsdmV0aWNhIC9FbmNvZGluZyAvV2luQW5zaUVuY29kaW5nIC9OYW1lIC9GMSAvU3VidHlwZSAvVHlwZTEgL1R5cGUgL0ZvbnQKPj4KZW5kb2JqCjMgMCBvYmoKPDwKL0NvbnRlbnRzIDcgMCBSIC9NZWRpYUJveCBbIDAgMCAyMDAgMTIwIF0gL1BhcmVudCA2IDAgUiAvUmVzb3VyY2VzIDw8Ci9Gb250IDEgMCBSIC9Qcm9jU2V0IFsgL1BERiAvVGV4dCAvSW1hZ2VCIC9JbWFnZUMgL0ltYWdlSSBdCj4+IC9Sb3RhdGUgMCAvVHJhbnMgPDwKCj4+IAogIC9UeXBlIC9QYWdlCj4+CmVuZG9iago0IDAgb2JqCjw8Ci9QYWdlTW9kZSAvVXNlTm9uZSAvUGFnZXMgNiAwIFIgL1R5cGUgL0NhdGFsb2cKPj4KZW5kb2JqCjUgMCBvYmoKPDwKL0F1dGhvciAoYW5vbnltb3VzKSAvQ3JlYXRpb25EYXRlIChEOjIwMjYxMTA2MDUyMTEyKzAwJzAwJykgL0NyZWF0b3IgKGFub255bW91cykgL0tleXdvcmRzICgpIC9Nb2REYXRlIChEOjIwMjYxMTA2MDUyMTEyKzAwJzAwJykgL1Byb2R1Y2VyIChSZXBvcnRMYWIgUERGIExpYnJhcnkgLSBcKG9wZW5zb3VyY2VcKSkgCiAgL1N1YmplY3QgKHVuc3BlY2lmaWVkKSAvVGl0bGUgKHVudGl0bGVkKSAvVHJhcHBlZCAvRmFsc2UKPj4KZW5kb2JqCjYgMCBvYmoKPDwKL0NvdW50IDEgL0tpZHMgWyAzIDAgUiBdIC9UeXBlIC9QYWdlcwo+PgplbmRvYmoKNyAwIG9iago8PAovRmlsdGVyIFsgL0FTQ0lJODVEZWNvZGUgL0ZsYXRlRGVjb2RlIF0gL0xlbmd0aCAxNTIKPj4Kc3RyZWFtCkdhcFFoMEU9RiwwVVxIM1RccE5ZVF5RS2s/dGM+SVAsO1cjVTFeMjNpaFBFTV8/Q1c0S0lTaTc6cDtXTSNnJClubjRDaFVhPUtzLTtbQD5eO1tDWGo8LGcsWThsUSljJWdtTFFvKXUnTDE9SzIuYVxpRGxfREsoO0BZLipaOjE+aiZKRTZYQFA1R182ZUkjT11RK0E2TH4+ZW5kc3RyZWFtCmVuZG9iagp4cmVmCjAgOAowMDAwMDAwMDAwIDY1NTM1IGYgCjAwMDAwMDAwNjEgMDAwMDAgbiAKMDAwMDAwMDA5MiAwMDAwMCBuIAowMDAwMDAwMTk5IDAwMDAwIG4gCjAwMDAwMDAzOTIgMDAwMDAgbiAKMDAwMDAwMDQ2MCAwMDAwMCBuIAowMDAwMDAwNzIxIDAwMDAwIG4gCjAwMDAwMDA3ODAgMDAwMDAgbiAKdHJhaWxlcgo8PAovSUQgCls8NGFiZWMwMDg2ZGZlNjQ2YjM4YjRkZTI2NGI2Njc4ODM+PDRhYmVjMDA4NmRmZTY0NmIzOGI0ZGUyNjRiNjY3ODgzPl0KJSBSZXBvcnRMYWIgZ2VuZXJhdGVkIFBERiBkb2N1bWVudCAtLSBkaWdlc3QgKG9wZW5zb3VyY2UpCgovSW5mbyA1IDAgUgovUm9vdCA0IDAgUgovU2l6ZSA4Cj4+CnN0YXJ0eHJlZgoxMDIyCiUlRU9GCg=='},{label:'Berkas Pendukung Duta Pelajar',name:'pendukung-duta.pdf',type:'application/pdf',data:'JVBERi0xLjMKJZOMi54gUmVwb3J0TGFiIEdlbmVyYXRlZCBQREYgZG9jdW1lbnQgKG9wZW5zb3VyY2UpCjEgMCBvYmoKPDwKL0YxIDIgMCBSCj4+CmVuZG9iagoyIDAgb2JqCjw8Ci9CYXNlRm9udCAvSGVsdmV0aWNhIC9FbmNvZGluZyAvV2luQW5zaUVuY29kaW5nIC9OYW1lIC9GMSAvU3VidHlwZSAvVHlwZTEgL1R5cGUgL0ZvbnQKPj4KZW5kb2JqCjMgMCBvYmoKPDwKL0NvbnRlbnRzIDcgMCBSIC9NZWRpYUJveCBbIDAgMCAyMDAgMTIwIF0gL1BhcmVudCA2IDAgUiAvUmVzb3VyY2VzIDw8Ci9Gb250IDEgMCBSIC9Qcm9jU2V0IFsgL1BERiAvVGV4dCAvSW1hZ2VCIC9JbWFnZUMgL0ltYWdlSSBdCj4+IC9Sb3RhdGUgMCAvVHJhbnMgPDwKCj4+IAogIC9UeXBlIC9QYWdlCj4+CmVuZG9iago0IDAgb2JqCjw8Ci9QYWdlTW9kZSAvVXNlTm9uZSAvUGFnZXMgNiAwIFIgL1R5cGUgL0NhdGFsb2cKPj4KZW5kb2JqCjUgMCBvYmoKPDwKL0F1dGhvciAoYW5vbnltb3VzKSAvQ3JlYXRpb25EYXRlIChEOjIwMjYxMTA2MDUyMTEyKzAwJzAwJykgL0NyZWF0b3IgKGFub255bW91cykgL0tleXdvcmRzICgpIC9Nb2REYXRlIChEOjIwMjYxMTA2MDUyMTEyKzAwJzAwJykgL1Byb2R1Y2VyIChSZXBvcnRMYWIgUERGIExpYnJhcnkgLSBcKG9wZW5zb3VyY2VcKSkgCiAgL1N1YmplY3QgKHVuc3BlY2lmaWVkKSAvVGl0bGUgKHVudGl0bGVkKSAvVHJhcHBlZCAvRmFsc2UKPj4KZW5kb2JqCjYgMCBvYmoKPDwKL0NvdW50IDEgL0tpZHMgWyAzIDAgUiBdIC9UeXBlIC9QYWdlcwo+PgplbmRvYmoKNyAwIG9iago8PAovRmlsdGVyIFsgL0FTQ0lJODVEZWNvZGUgL0ZsYXRlRGVjb2RlIF0gL0xlbmd0aCAxNTIKPj4Kc3RyZWFtCkdhcFFoMEU9RiwwVVxIM1RccE5ZVF5RS2s/dGM+SVAsO1cjVTFeMjNpaFBFTV8/Q1c0S0lTaTc6cDtXTSNnJClubjRDaFVhPUtzLTtbQD5eO1tDWGo8LGcsWThsUSljJWdtTFFvKXUnTDE9SzIuYVxpRGxfREsoO0BZLipaOjE+aiZKRTZYQFA1R182ZUkjT11RK0E2TH4+ZW5kc3RyZWFtCmVuZG9iagp4cmVmCjAgOAowMDAwMDAwMDAwIDY1NTM1IGYgCjAwMDAwMDAwNjEgMDAwMDAgbiAKMDAwMDAwMDA5MiAwMDAwMCBuIAowMDAwMDAwMTk5IDAwMDAwIG4gCjAwMDAwMDAzOTIgMDAwMDAgbiAKMDAwMDAwMDQ2MCAwMDAwMCBuIAowMDAwMDAwNzIxIDAwMDAwIG4gCjAwMDAwMDA3ODAgMDAwMDAgbiAKdHJhaWxlcgo8PAovSUQgCls8NGFiZWMwMDg2ZGZlNjQ2YjM4YjRkZTI2NGI2Njc4ODM+PDRhYmVjMDA4NmRmZTY0NmIzOGI0ZGUyNjRiNjY3ODgzPl0KJSBSZXBvcnRMYWIgZ2VuZXJhdGVkIFBERiBkb2N1bWVudCAtLSBkaWdlc3QgKG9wZW5zb3VyY2UpCgovSW5mbyA1IDAgUgovUm9vdCA0IDAgUgovU2l6ZSA4Cj4+CnN0YXJ0eHJlZgoxMDIyCiUlRU9GCg=='},{label:'Surat Tugas dari PR IPNU-IPPNU',name:'surat-tugas.pdf',type:'application/pdf',data:'JVBERi0xLjMKJZOMi54gUmVwb3J0TGFiIEdlbmVyYXRlZCBQREYgZG9jdW1lbnQgKG9wZW5zb3VyY2UpCjEgMCBvYmoKPDwKL0YxIDIgMCBSCj4+CmVuZG9iagoyIDAgb2JqCjw8Ci9CYXNlRm9udCAvSGVsdmV0aWNhIC9FbmNvZGluZyAvV2luQW5zaUVuY29kaW5nIC9OYW1lIC9GMSAvU3VidHlwZSAvVHlwZTEgL1R5cGUgL0ZvbnQKPj4KZW5kb2JqCjMgMCBvYmoKPDwKL0NvbnRlbnRzIDcgMCBSIC9NZWRpYUJveCBbIDAgMCAyMDAgMTIwIF0gL1BhcmVudCA2IDAgUiAvUmVzb3VyY2VzIDw8Ci9Gb250IDEgMCBSIC9Qcm9jU2V0IFsgL1BERiAvVGV4dCAvSW1hZ2VCIC9JbWFnZUMgL0ltYWdlSSBdCj4+IC9Sb3RhdGUgMCAvVHJhbnMgPDwKCj4+IAogIC9UeXBlIC9QYWdlCj4+CmVuZG9iago0IDAgb2JqCjw8Ci9QYWdlTW9kZSAvVXNlTm9uZSAvUGFnZXMgNiAwIFIgL1R5cGUgL0NhdGFsb2cKPj4KZW5kb2JqCjUgMCBvYmoKPDwKL0F1dGhvciAoYW5vbnltb3VzKSAvQ3JlYXRpb25EYXRlIChEOjIwMjYxMTA2MDUyMTEyKzAwJzAwJykgL0NyZWF0b3IgKGFub255bW91cykgL0tleXdvcmRzICgpIC9Nb2REYXRlIChEOjIwMjYxMTA2MDUyMTEyKzAwJzAwJykgL1Byb2R1Y2VyIChSZXBvcnRMYWIgUERGIExpYnJhcnkgLSBcKG9wZW5zb3VyY2VcKSkgCiAgL1N1YmplY3QgKHVuc3BlY2lmaWVkKSAvVGl0bGUgKHVudGl0bGVkKSAvVHJhcHBlZCAvRmFsc2UKPj4KZW5kb2JqCjYgMCBvYmoKPDwKL0NvdW50IDEgL0tpZHMgWyAzIDAgUiBdIC9UeXBlIC9QYWdlcwo+PgplbmRvYmoKNyAwIG9iago8PAovRmlsdGVyIFsgL0FTQ0lJODVEZWNvZGUgL0ZsYXRlRGVjb2RlIF0gL0xlbmd0aCAxNTIKPj4Kc3RyZWFtCkdhcFFoMEU9RiwwVVxIM1RccE5ZVF5RS2s/dGM+SVAsO1cjVTFeMjNpaFBFTV8/Q1c0S0lTaTc6cDtXTSNnJClubjRDaFVhPUtzLTtbQD5eO1tDWGo8LGcsWThsUSljJWdtTFFvKXUnTDE9SzIuYVxpRGxfREsoO0BZLipaOjE+aiZKRTZYQFA1R182ZUkjT11RK0E2TH4+ZW5kc3RyZWFtCmVuZG9iagp4cmVmCjAgOAowMDAwMDAwMDAwIDY1NTM1IGYgCjAwMDAwMDAwNjEgMDAwMDAgbiAKMDAwMDAwMDA5MiAwMDAwMCBuIAowMDAwMDAwMTk5IDAwMDAwIG4gCjAwMDAwMDAzOTIgMDAwMDAgbiAKMDAwMDAwMDQ2MCAwMDAwMCBuIAowMDAwMDAwNzIxIDAwMDAwIG4gCjAwMDAwMDA3ODAgMDAwMDAgbiAKdHJhaWxlcgo8PAovSUQgCls8NGFiZWMwMDg2ZGZlNjQ2YjM4YjRkZTI2NGI2Njc4ODM+PDRhYmVjMDA4NmRmZTY0NmIzOGI0ZGUyNjRiNjY3ODgzPl0KJSBSZXBvcnRMYWIgZ2VuZXJhdGVkIFBERiBkb2N1bWVudCAtLSBkaWdlc3QgKG9wZW5zb3VyY2UpCgovSW5mbyA1IDAgUgovUm9vdCA0IDAgUgovU2l6ZSA4Cj4+CnN0YXJ0eHJlZgoxMDIyCiUlRU9GCg=='}]};
 return buildTestPdfJob_(d);
}
function ujiCetakPdfTunggal(){
 var list=[{nama:'Peserta Tunggal Contoh',gender:'L',tempatLahir:'Kendal',tanggalLahir:'2007-03-14',alamat:'Jl. Contoh No. 1, Patebon',wa:'081234567892'}];
 var d={type:'cabang',reg:'TUNGGAL-UJI-'+Utilities.getUuid().slice(0,6),nama:'PR IPNU IPPNU CONTOH UJI COBA',cabang:'Khitobah',kategori:'Putra',jml:1,email:'contoh@example.com',waktu:new Date().toISOString(),pesertaList:list,catatan:'Uji format cabang lomba tunggal.',files:[sampleImageAttachment_('Pas foto peserta 1','foto-tunggal.png'),sampleImageAttachment_('KTP peserta 1','ktp-tunggal.png'),{label:'Surat Tugas dari PR IPNU-IPPNU',name:'surat-tugas.pdf',type:'application/pdf',data:'JVBERi0xLjMKJZOMi54gUmVwb3J0TGFiIEdlbmVyYXRlZCBQREYgZG9jdW1lbnQgKG9wZW5zb3VyY2UpCjEgMCBvYmoKPDwKL0YxIDIgMCBSCj4+CmVuZG9iagoyIDAgb2JqCjw8Ci9CYXNlRm9udCAvSGVsdmV0aWNhIC9FbmNvZGluZyAvV2luQW5zaUVuY29kaW5nIC9OYW1lIC9GMSAvU3VidHlwZSAvVHlwZTEgL1R5cGUgL0ZvbnQKPj4KZW5kb2JqCjMgMCBvYmoKPDwKL0NvbnRlbnRzIDcgMCBSIC9NZWRpYUJveCBbIDAgMCAyMDAgMTIwIF0gL1BhcmVudCA2IDAgUiAvUmVzb3VyY2VzIDw8Ci9Gb250IDEgMCBSIC9Qcm9jU2V0IFsgL1BERiAvVGV4dCAvSW1hZ2VCIC9JbWFnZUMgL0ltYWdlSSBdCj4+IC9Sb3RhdGUgMCAvVHJhbnMgPDwKCj4+IAogIC9UeXBlIC9QYWdlCj4+CmVuZG9iago0IDAgb2JqCjw8Ci9QYWdlTW9kZSAvVXNlTm9uZSAvUGFnZXMgNiAwIFIgL1R5cGUgL0NhdGFsb2cKPj4KZW5kb2JqCjUgMCBvYmoKPDwKL0F1dGhvciAoYW5vbnltb3VzKSAvQ3JlYXRpb25EYXRlIChEOjIwMjYxMTA2MDUyMTEyKzAwJzAwJykgL0NyZWF0b3IgKGFub255bW91cykgL0tleXdvcmRzICgpIC9Nb2REYXRlIChEOjIwMjYxMTA2MDUyMTEyKzAwJzAwJykgL1Byb2R1Y2VyIChSZXBvcnRMYWIgUERGIExpYnJhcnkgLSBcKG9wZW5zb3VyY2VcKSkgCiAgL1N1YmplY3QgKHVuc3BlY2lmaWVkKSAvVGl0bGUgKHVudGl0bGVkKSAvVHJhcHBlZCAvRmFsc2UKPj4KZW5kb2JqCjYgMCBvYmoKPDwKL0NvdW50IDEgL0tpZHMgWyAzIDAgUiBdIC9UeXBlIC9QYWdlcwo+PgplbmRvYmoKNyAwIG9iago8PAovRmlsdGVyIFsgL0FTQ0lJODVEZWNvZGUgL0ZsYXRlRGVjb2RlIF0gL0xlbmd0aCAxNTIKPj4Kc3RyZWFtCkdhcFFoMEU9RiwwVVxIM1RccE5ZVF5RS2s/dGM+SVAsO1cjVTFeMjNpaFBFTV8/Q1c0S0lTaTc6cDtXTSNnJClubjRDaFVhPUtzLTtbQD5eO1tDWGo8LGcsWThsUSljJWdtTFFvKXUnTDE9SzIuYVxpRGxfREsoO0BZLipaOjE+aiZKRTZYQFA1R182ZUkjT11RK0E2TH4+ZW5kc3RyZWFtCmVuZG9iagp4cmVmCjAgOAowMDAwMDAwMDAwIDY1NTM1IGYgCjAwMDAwMDAwNjEgMDAwMDAgbiAKMDAwMDAwMDA5MiAwMDAwMCBuIAowMDAwMDAwMTk5IDAwMDAwIG4gCjAwMDAwMDAzOTIgMDAwMDAgbiAKMDAwMDAwMDQ2MCAwMDAwMCBuIAowMDAwMDAwNzIxIDAwMDAwIG4gCjAwMDAwMDA3ODAgMDAwMDAgbiAKdHJhaWxlcgo8PAovSUQgCls8NGFiZWMwMDg2ZGZlNjQ2YjM4YjRkZTI2NGI2Njc4ODM+PDRhYmVjMDA4NmRmZTY0NmIzOGI0ZGUyNjRiNjY3ODgzPl0KJSBSZXBvcnRMYWIgZ2VuZXJhdGVkIFBERiBkb2N1bWVudCAtLSBkaWdlc3QgKG9wZW5zb3VyY2UpCgovSW5mbyA1IDAgUgovUm9vdCA0IDAgUgovU2l6ZSA4Cj4+CnN0YXJ0eHJlZgoxMDIyCiUlRU9GCg=='}]};
 return buildTestPdfJob_(d);
}
function ujiCetakPdfBeregu(){
 var list=[];for(var i=1;i<=5;i++)list.push({nama:'Peserta Beregu '+i,gender:i%2?'L':'P',role:i===1?'Ketua Tim':'Anggota',tempatLahir:'Kendal',tanggalLahir:'2007-0'+(1+i%9)+'-1'+(i%9),alamat:'Jl. Contoh No. '+i+', Patebon',wa:'0812345678'+(10+i)});
 var d={type:'cabang',reg:'BEREGU-UJI-'+Utilities.getUuid().slice(0,6),nama:'PR IPNU IPPNU CONTOH UJI COBA',cabang:'Rebana',kategori:'Kelompok Bebas',jml:list.length,email:'contoh@example.com',waktu:new Date().toISOString(),pesertaList:list,catatan:'Uji format cabang lomba beregu.',files:[sampleImageAttachment_('Pas foto peserta 1','foto-peserta-1.png'),sampleImageAttachment_('KTP peserta 1','ktp-peserta-1.png'),{label:'Surat Tugas dari PR IPNU-IPPNU',name:'surat-tugas.pdf',type:'application/pdf',data:'JVBERi0xLjMKJZOMi54gUmVwb3J0TGFiIEdlbmVyYXRlZCBQREYgZG9jdW1lbnQgKG9wZW5zb3VyY2UpCjEgMCBvYmoKPDwKL0YxIDIgMCBSCj4+CmVuZG9iagoyIDAgb2JqCjw8Ci9CYXNlRm9udCAvSGVsdmV0aWNhIC9FbmNvZGluZyAvV2luQW5zaUVuY29kaW5nIC9OYW1lIC9GMSAvU3VidHlwZSAvVHlwZTEgL1R5cGUgL0ZvbnQKPj4KZW5kb2JqCjMgMCBvYmoKPDwKL0NvbnRlbnRzIDcgMCBSIC9NZWRpYUJveCBbIDAgMCAyMDAgMTIwIF0gL1BhcmVudCA2IDAgUiAvUmVzb3VyY2VzIDw8Ci9Gb250IDEgMCBSIC9Qcm9jU2V0IFsgL1BERiAvVGV4dCAvSW1hZ2VCIC9JbWFnZUMgL0ltYWdlSSBdCj4+IC9Sb3RhdGUgMCAvVHJhbnMgPDwKCj4+IAogIC9UeXBlIC9QYWdlCj4+CmVuZG9iago0IDAgb2JqCjw8Ci9QYWdlTW9kZSAvVXNlTm9uZSAvUGFnZXMgNiAwIFIgL1R5cGUgL0NhdGFsb2cKPj4KZW5kb2JqCjUgMCBvYmoKPDwKL0F1dGhvciAoYW5vbnltb3VzKSAvQ3JlYXRpb25EYXRlIChEOjIwMjYxMTA2MDUyMTEyKzAwJzAwJykgL0NyZWF0b3IgKGFub255bW91cykgL0tleXdvcmRzICgpIC9Nb2REYXRlIChEOjIwMjYxMTA2MDUyMTEyKzAwJzAwJykgL1Byb2R1Y2VyIChSZXBvcnRMYWIgUERGIExpYnJhcnkgLSBcKG9wZW5zb3VyY2VcKSkgCiAgL1N1YmplY3QgKHVuc3BlY2lmaWVkKSAvVGl0bGUgKHVudGl0bGVkKSAvVHJhcHBlZCAvRmFsc2UKPj4KZW5kb2JqCjYgMCBvYmoKPDwKL0NvdW50IDEgL0tpZHMgWyAzIDAgUiBdIC9UeXBlIC9QYWdlcwo+PgplbmRvYmoKNyAwIG9iago8PAovRmlsdGVyIFsgL0FTQ0lJODVEZWNvZGUgL0ZsYXRlRGVjb2RlIF0gL0xlbmd0aCAxNTIKPj4Kc3RyZWFtCkdhcFFoMEU9RiwwVVxIM1RccE5ZVF5RS2s/dGM+SVAsO1cjVTFeMjNpaFBFTV8/Q1c0S0lTaTc6cDtXTSNnJClubjRDaFVhPUtzLTtbQD5eO1tDWGo8LGcsWThsUSljJWdtTFFvKXUnTDE9SzIuYVxpRGxfREsoO0BZLipaOjE+aiZKRTZYQFA1R182ZUkjT11RK0E2TH4+ZW5kc3RyZWFtCmVuZG9iagp4cmVmCjAgOAowMDAwMDAwMDAwIDY1NTM1IGYgCjAwMDAwMDAwNjEgMDAwMDAgbiAKMDAwMDAwMDA5MiAwMDAwMCBuIAowMDAwMDAwMTk5IDAwMDAwIG4gCjAwMDAwMDAzOTIgMDAwMDAgbiAKMDAwMDAwMDQ2MCAwMDAwMCBuIAowMDAwMDAwNzIxIDAwMDAwIG4gCjAwMDAwMDA3ODAgMDAwMDAgbiAKdHJhaWxlcgo8PAovSUQgCls8NGFiZWMwMDg2ZGZlNjQ2YjM4YjRkZTI2NGI2Njc4ODM+PDRhYmVjMDA4NmRmZTY0NmIzOGI0ZGUyNjRiNjY3ODgzPl0KJSBSZXBvcnRMYWIgZ2VuZXJhdGVkIFBERiBkb2N1bWVudCAtLSBkaWdlc3QgKG9wZW5zb3VyY2UpCgovSW5mbyA1IDAgUgovUm9vdCA0IDAgUgovU2l6ZSA4Cj4+CnN0YXJ0eHJlZgoxMDIyCiUlRU9GCg=='}]};
 return buildTestPdfJob_(d);
}
function adminTestPdf_(d){
 var admin=requireAdmin_(d.token),kind=String(d.kind||'all');
 if(!pdfMergeReady_())throw new Error('Uji cetak formulir memerlukan library PDFApp. Library belum terdeteksi pada proyek Apps Script. Pasang/aktifkan library PDFApp yang menyediakan fungsi mergePDFs, lalu deploy ulang.');
 cleanupTestPdfStatus_();
 var jobs={kontingen:ujiCetakPdfKontingen,duta:ujiCetakPdfDutaPelajar,tunggal:ujiCetakPdfTunggal,beregu:ujiCetakPdfBeregu},result={};
 if(kind==='all'){Object.keys(jobs).forEach(function(k){result[k]=jobs[k]();});}
 else if(jobs[kind])result[kind]=jobs[kind]();
 else throw new Error('Jenis uji cetak tidak dikenal.');
 adminAudit_(admin,'UJI CETAK PDF','sistem','','uji coba',kind==='all'?'Empat PDF contoh sedang diproses dan akan disimpan sebagai PDF final di folder UJI COBA.':'PDF uji '+kind+' sedang diproses.');
 return {ok:true,kind:kind,jobs:result,message:'PDF sedang disusun. Preview akan terbuka otomatis setelah seluruh halaman selesai digabungkan.'};
}

function appendDriveSignature_(cell, fileId, maxWidth, maxHeight) {
  if (!fileId) {
    throw new Error('ID file tanda tangan belum diisi.');
  }

  var file = DriveApp.getFileById(fileId);
  var image = cell.appendImage(file.getBlob());

  var originalWidth = image.getWidth();
  var originalHeight = image.getHeight();

  var ratio = Math.min(
    maxWidth / originalWidth,
    maxHeight / originalHeight,
    1
  );

  image.setWidth(Math.round(originalWidth * ratio));
  image.setHeight(Math.round(originalHeight * ratio));

  image.getParent()
    .asParagraph()
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  return image;
}


/**
 * Menambahkan stempel panitia pada area tanda tangan sekretariat.
 * Stempel ditempatkan dalam paragraf yang sama dengan tanda tangan agar
 * tampil berdampingan dan masuk ke area tanda tangan, bukan di bawah nama.
 */
function appendDriveStampBesideSignature_(signatureImage, fileId, maxWidth, maxHeight) {
  if (!signatureImage) throw new Error('Gambar tanda tangan sekretariat tidak ditemukan.');
  if (!fileId) throw new Error('ID Google Drive stempel panitia belum diisi.');

  var stampFile = DriveApp.getFileById(fileId);
  if (stampFile.isTrashed()) throw new Error('File stempel panitia berada di Trash Google Drive.');
  var stampBlob = stampFile.getBlob();
  if (stampBlob.getContentType() !== 'image/png') {
    throw new Error('Stempel panitia harus berupa file PNG.');
  }

  var paragraph = signatureImage.getParent().asParagraph();
  // Cap organisasi ditempatkan di sisi kiri tanda tangan. Sisipkan sebelum
  // gambar tanda tangan agar tidak terdorong ke kanan/keluar area blok.
  var stamp = paragraph.insertInlineImage(paragraph.getChildIndex(signatureImage), stampBlob);
  var width = stamp.getWidth();
  var height = stamp.getHeight();
  var ratio = Math.min(maxWidth / width, maxHeight / height, 1);
  stamp.setWidth(Math.round(width * ratio));
  stamp.setHeight(Math.round(height * ratio));
  return stamp;
}


function createBranchRegistrationPdf_(d, folder) {

  var doc = pdfNewDoc_('TEMP PORSENI CABANG ' + d.reg, 0);
  var body = doc.getBody();

  // Gunakan header A4 standar yang sama dengan formulir Kontingen.
  // pdfNewDoc_ sudah memasang kop pada header dokumen dengan lebar penuh A4
  // dan margin atas yang menyediakan ruang untuk kop. Header otomatis berulang
  // pada setiap halaman formulir; lampiran foto/KTP dan PDF tetap tanpa kop.
  var CENTER = DocumentApp.HorizontalAlignment.CENTER;
  var JUSTIFY = DocumentApp.HorizontalAlignment.JUSTIFY;

  var w = pdfWriter_(pdfPage_(body));

  function heading(text, before) {
    return w.para(text, {
      size: 12,
      bold: true,
      color: PDF_NAVY,
      before: before === undefined ? 8 : before,
      after: 3
    });
  }

  // ==================================================
  // HALAMAN 1: IDENTITAS DAN DATA PESERTA
  // ==================================================

  w.para('FORMULIR PENGUMPULAN BERKAS', {
    size: 15,
    bold: true,
    align: CENTER,
    before: 0,
    after: 1
  });

  w.para(
    ('CABANG LOMBA ' + displayBranchName_(d.cabang)).toUpperCase(),
    {
      size: 14,
      bold: true,
      align: CENTER,
      before: 0,
      after: 1
    }
  );

  w.para(
    ('KATEGORI ' + displayCategoryName_(d.cabang, d.kategori)).toUpperCase(),
    {
      size: 14,
      bold: true,
      align: CENTER,
      before: 0,
      after: 13
    }
  );

  // A. IDENTITAS KONTINGEN
  heading('A. IDENTITAS KONTINGEN DAN PERLOMBAAN', 0);

  var meta = w.table([
    ['Kontingen', d.nama || ''],
    [
      'Jumlah Peserta',
      String(d.jml || ((d.pesertaList || []).length))
    ]
  ]);

  pdfGrid_(meta, [
    150,
    PDF_CONTENT_W - 150
  ], {
    size: 11,
    padV: 4,
    padH: 8,
    minH: 22,
    boldCols: [0],
    labelCol: 0
  });

  // B. DATA PESERTA: tata letak dibedakan menurut kategori tunggal/beregu.
  var isTeam = /^beregu/i.test(String(d.kategori||''));
  var isDuta = String(d.cabang||'') === 'Duta Pelajar';
  var participantRows = isTeam
    ? (isDuta ? [['No.','Nama Peserta','Jenis Kelamin','Tempat/Tanggal Lahir','WhatsApp','Peran','Alamat Rumah']] : [['No.','Nama Peserta','Jenis Kelamin','Tempat/Tanggal Lahir','Peran','Alamat Rumah']])
    : (isDuta ? [['No.','Nama Peserta','Jenis Kelamin','Tempat/Tanggal Lahir','WhatsApp','Alamat Rumah']] : [['No.','Nama Peserta','Jenis Kelamin','Tempat/Tanggal Lahir','Alamat Rumah']]);
  (d.pesertaList || []).forEach(function(p, i) {
    var gender=p.gender === 'L' ? 'Putra' : p.gender === 'P' ? 'Putri' : String(p.gender||'');
    var birth=[String(p.tempatLahir||''),idDate_(p.tanggalLahir)].filter(Boolean).join(', ');
    var base=[String(i+1),String(p.nama||''),gender,birth];
    if(isDuta)base.push(String(p.wa||''));
    if(isTeam)base.push(String(p.role||'Anggota'));
    base.push(String(p.alamat||''));participantRows.push(base);
  });
  heading(isTeam?'B. DATA PESERTA BEREGU / TIM':'B. DATA PESERTA TUNGGAL');
  var pt=w.table(participantRows);
  var widths=isTeam?(isDuta?[22,90,48,78,68,58,PDF_CONTENT_W-364]:[22,100,52,90,62,PDF_CONTENT_W-326]):(isDuta?[24,105,54,105,78,PDF_CONTENT_W-366]:[24,112,58,112,PDF_CONTENT_W-306]);
  var aligns=isTeam?(isDuta?['C','L','C','C','C','C','L']:['C','L','C','C','C','L']):(isDuta?['C','L','C','C','C','L']:['C','L','C','C','L']);
  pdfGrid_(pt,widths,{size:8.5,padV:3,padH:3,header:true,align:aligns});
  if(String(d.cabang||'')==='Duta Pelajar'){
    heading('C. KETERANGAN KHUSUS DUTA PELAJAR');
    w.para('Pengumpulan ini menggunakan format umum formulir cabang lomba. Berkas khusus Duta Pelajar (CV dan berkas pendukung) tercatat sebagai lampiran pada bagian administrasi.',{size:10,align:JUSTIFY,after:3});
    (d.files||[]).filter(function(f){return /duta pelajar/i.test(String(f.label||''));}).forEach(function(f){w.para('• '+String(f.label||'Lampiran Duta Pelajar')+' — '+String(f.name||''),{size:9.5,after:2,indent:10});});
  }

  // ==================================================
  // HALAMAN 2: ADMINISTRASI DAN TANDA TANGAN
  // ==================================================

  body.appendPageBreak();

  // Writer halaman kedua; kop diulang otomatis oleh header standar dokumen.
  w = pdfWriter_(pdfPage_(body));

  // C. CATATAN (jika tersedia)
  var nx = 'C';

  if (d.catatan) {
    heading(nx + '. CATATAN');

    w.para(String(d.catatan), {
      size: 10.5,
      align: JUSTIFY,
      after: 5
    });

    nx = 'D';
  }

  // PERNYATAAN
  heading(nx + '. PERNYATAAN');

  w.para(
    'Data peserta dan dokumen yang disampaikan merupakan data milik kontingen yang bersangkutan. Formulir ini menjadi bukti pengiriman berkas kepada Panitia PORSENI 2026 PAC IPNU–IPPNU Patebon.',
    {
      size: 10.5,
      align: JUSTIFY,
      line: 1.3,
      after: 6
    }
  );

  // ===== JUDUL VERIFIKASI =====
var verifikasi = w.para('DIVERIFIKASI OLEH');

verifikasi
  .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
  .setSpacingBefore(12)
  .setSpacingAfter(12)
  .setLineSpacing(1)
  .setIndentStart(0)
  .setIndentFirstLine(0)
  .setIndentEnd(0);

pdfFont_(verifikasi.editAsText(), 11, true);


// ===== TABEL TANDA TANGAN 2 KOLOM =====
var sg = w.table([
  ['KOOR. PERLOMBAAN', 'SEKRETARIAT'],
  ['\u00A0', '\u00A0'],
  ['( SOCHIBUL MUSTAQIM )', '( A. FUAD SYAFI )']
]);

sg.setBorderWidth(0).setBorderColor('#FFFFFF');

var sw = PDF_CONTENT_W / 2;

for (var sc = 0; sc < 2; sc++) {
  sg.setColumnWidth(sc, sw);
}

// Tinggi area tanda tangan
sg.getRow(1).setMinimumHeight(100);

// Pengaturan seluruh sel
for (var sr = 0; sr < 3; sr++) {
  for (var sc2 = 0; sc2 < 2; sc2++) {

    var cell = sg.getCell(sr, sc2);

    cell.setPaddingTop(3)
      .setPaddingBottom(3)
      .setPaddingLeft(5)
      .setPaddingRight(5)
      .setVerticalAlignment(
        DocumentApp.VerticalAlignment.CENTER
      );

    var sp = cell.getChild(0).asParagraph();

    sp.setAlignment(
        DocumentApp.HorizontalAlignment.CENTER
      )
      .setSpacingBefore(0)
      .setSpacingAfter(0)
      .setLineSpacing(1)
      .setIndentStart(0)
      .setIndentFirstLine(0)
      .setIndentEnd(0);

    pdfFont_(cell.editAsText(), 11, true);
  }
}

// ===== TANDA TANGAN KOOR. PERLOMBAAN =====
appendDriveSignature_(
  sg.getCell(1, 0),
  PDF_TTD_KOOR_ID,
  180,
  85
);

// ===== TANDA TANGAN SEKRETARIAT =====
var sekretariatSignatureImage = appendDriveSignature_(
  sg.getCell(1, 1),
  PDF_TTD_SEKRETARIAT_ID,
  145,
  75
);

// Stempel berada di samping tanda tangan sekretariat, masuk ke area tanda tangan.
appendDriveStampBesideSignature_(
  sekretariatSignatureImage,
  PDF_STEMPEL_PANITIA_ID,
  48,
  48
);

  w.para(
    'Dokumen ini dibuat secara elektronik melalui website resmi PORSENI 2026 | Nomor Pendaftaran: ' +
    d.reg + ' | ' + idDate_(d.waktu || new Date(), true),
    {
      size: 10,
      italic: true,
      align: CENTER,
      color: '#3A4A63',
      before: 5
    }
  );

  appendBranchImageAttachmentPages_(body, d);
  // Surat Tugas dan Lampiran Khusus berformat PDF digabungkan setelah PDF formulir,
  // dengan urutan yang disimpan pada daftar berkas unggahan.
  pdfTidyBody_(body);

  return pdfSavePdf_(doc, d, folder);
}



/* ===================================================================
   V34 — PENGGABUNGAN PDF YANG TAHAN GANGGUAN
   =================================================================== */
function isMergePending_(reg){
  reg=String(reg||'');if(!reg)return false;
  try{
    if(listPdfMergeJobs_().some(function(j){return !j.jobId&&String(j.reg||'')===reg;}))return true;
    var legacy=JSON.parse(PropertiesService.getScriptProperties().getProperty(PDF_MERGE_QUEUE_PROPERTY_)||'[]');
    return legacy.some(function(j){return !j.jobId&&String(j.reg||'')===reg;});
  }catch(e){return false;}
}
// True bila berkas ini adalah hasil penggabungan (lampiran PDF sudah ada di dalamnya).
// Penanda utama: deskripsi berkas. Cadangan untuk berkas lama: log keamanan memuat ID berkas pada status BERHASIL.
function wasPdfMerged_(file){
  try{if(String(file.getDescription()||'').indexOf(PDF_MERGED_MARKER_)===0)return true;}catch(e){}
  try{
    var id=file.getId(),sh=ensureSecurityLogSheet_(),last=sh.getLastRow();
    if(last<2)return false;
    var start=Math.max(2,last-5000),rows=sh.getRange(start,3,last-start+1,3).getValues();
    for(var i=rows.length-1;i>=0;i--){
      if(String(rows[i][0])==='PENGGABUNGAN PDF'&&String(rows[i][1])==='BERHASIL'&&String(rows[i][2]).indexOf(id)>=0)return true;
    }
  }catch(e){console.warn('Pemeriksaan log penggabungan PDF gagal: '+e);}
  return false;
}
function notifyMergeFailure_(job,reason){
  try{
    logSecurityEvent_(job.nama,'PENGGABUNGAN PDF','GAGAL PERMANEN',job.reg+' | '+String(reason||''));
    if(MailApp.getRemainingDailyQuota()>0)sendPorseniEmail_({to:NOTIFY,subject:'[PORSENI 2026] PDF belum tergabung: '+job.nama+' ('+job.reg+')',body:'Penggabungan formulir dan lampiran PDF gagal.\nKontingen: '+job.nama+'\nNo. registrasi: '+job.reg+'\nPenyebab: '+String(reason||'-')+'\n\nBuka Portal Panitia, pilih data tersebut, lalu klik "Gabung ulang PDF".',name:SENDER_NAME,replyTo:REPLY_TO});
  }catch(e){console.error(e);}
}
function cleanupTestPdfStatus_(){
  try{
    var props=PropertiesService.getScriptProperties(),all=props.getProperties(),now=Date.now();
    Object.keys(all).forEach(function(k){
      if(k.indexOf('PORSENI_TEST_PDF_STATUS_')!==0)return;
      var at=0;try{at=Number(JSON.parse(all[k]).at)||0;}catch(e){}
      if(!at||now-at>12*3600*1000)props.deleteProperty(k);
    });
  }catch(e){console.warn('Pembersihan status uji gagal: '+e);}
}
// Pratinjau PDF uji lewat server (tidak bergantung pada login Google di browser admin).
function adminTestPdfFile_(d){
  requireAdmin_(d.token);
  var id=String(d.jobId||'');if(!/^[A-Za-z0-9_-]{6,80}$/.test(id))throw new Error('ID uji tidak valid.');
  var raw=PropertiesService.getScriptProperties().getProperty('PORSENI_TEST_PDF_STATUS_'+id);
  if(!raw)throw new Error('Status uji tidak ditemukan atau sudah kedaluwarsa. Jalankan uji cetak lagi.');
  var st=JSON.parse(raw);if(st.state!=='ready'||!st.fileId)throw new Error('PDF uji belum siap.');
  var file=DriveApp.getFileById(st.fileId),testId=testFolder_().getId(),ok=false,ps=file.getParents();
  while(ps.hasNext()){if(ps.next().getId()===testId)ok=true;}
  if(!ok)throw new Error('Berkas tidak berada di folder UJI COBA.');
  var bytes=file.getBlob().getBytes();
  if(bytes.length>20*1024*1024)throw new Error('PDF uji '+(bytes.length/1048576).toFixed(1)+' MB terlalu besar untuk pratinjau. Buka langsung dari folder UJI COBA di Drive.');
  return {ok:true,file:{name:file.getName(),mimeType:'application/pdf',size:bytes.length,base64:Utilities.base64Encode(bytes)}};
}
// Gabung ulang PDF untuk satu data pendaftaran (mis. library sempat belum terpasang atau gagal 3x).

function adminRebuildRegistrationPdf_(d){
 var admin=requireAdmin_(d.token),type=String(d.type||'');
 if(type!=='kontingen'&&type!=='cabang')throw new Error('Jenis data tidak valid.');
 var ss=SpreadsheetApp.getActiveSpreadsheet(),sh=ss&&ss.getSheetByName(CFG[type].sheet);
 if(!sh)throw new Error('Data pendaftaran tidak ditemukan.');
 var row=resolveAdminRecordRow_(sh,d),headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0],vals=sh.getRange(row,1,1,sh.getLastColumn()).getValues()[0];
 function val(h){var i=headers.indexOf(h);return i<0?'':vals[i];}
 var folderUrl=String(val('Folder Drive')||''),folderId=getDriveFileIdFromUrl_(folderUrl);
 if(!folderId)throw new Error('Folder arsip tidak ditemukan.');
 var folder=DriveApp.getFolderById(folderId);assertPrivateFolderTree_(folder);
 var reg=String(val('No. Registrasi')||''),nama=String(val('Nama Kontingen')||'');
 var links=parsePortalFiles_(String(val('Berkas (link Drive)')||'')).filter(function(x){return !/^pdf formulir$/i.test(x.label);});
 var files=[],savedPdf=[];
 links.forEach(function(x){
  var id=getDriveFileIdFromUrl_(x.url);if(!id)return;
  var f=DriveApp.getFileById(id),mime=String(f.getMimeType()||'').toLowerCase();
  // PDF hanya memerlukan metadata untuk halaman indeks dan antrean merge.
  // Hindari membaca seluruh byte PDF lalu mengubahnya menjadi base64 saat rebuild.
  if(mime===MimeType.PDF){
   if(!isStandaloneCompetitionDocumentLabel_(x.label)){
    files.push({label:x.label,name:f.getName(),type:MimeType.PDF});
    savedPdf.push({id:id,mime:MimeType.PDF,label:x.label,name:f.getName()});
   }
   return;
  }
  // Hanya gambar yang benar-benar ditanam ke Google Docs yang dibaca.
  if(mime==='image/jpeg'||mime==='image/png'){
   var blob=f.getBlob(),bytes=blob.getBytes();
   files.push({label:x.label,name:f.getName(),type:detectedMime_(bytes,mime),data:Utilities.base64Encode(bytes)});
  }
 });
 var rec={type:type,reg:reg,nama:nama,files:files,savedAttachments:savedPdf};
 if(type==='kontingen'){
  rec.asal=String(val('Asal')||'');rec.alamat=String(val('Alamat')||'');rec.ketua=String(val('Ketua')||'');
  rec.tempat_lahir_ketua=String(val('Tempat Lahir Ketua')||'');rec.tanggal_lahir_ketua=String(val('Tanggal Lahir Ketua')||'');
  rec.wa_ketua=String(val('WA Ketua')||'');rec.email=String(val('Email')||'');rec.cabang=String(val('Cabang Lomba')||'');
  rec.jml=String(val('Jumlah Peserta')||'');rec.bayar=String(val('Metode Bayar')||'');rec.tgl_bayar=String(val('Tgl Bayar')||'');
  rec.ketua_pr_ipnu=String(val('Ketua PR IPNU')||'');rec.ketua_pr_ippnu=String(val('Ketua PR IPPNU')||'');
  rec.wa_ketua_ipnu=String(val('WA Ketua IPNU')||'');rec.wa_ketua_ippnu=String(val('WA Ketua IPPNU')||'');
 }else{
  rec.cabang=String(val('Cabang Lomba')||'');rec.kategori=String(val('Kategori')||'');rec.catatan=String(val('Catatan')||'');rec.jml=String(val('Jumlah Peserta')||'');
  var raw=String(val('Daftar Peserta')||'').split(/\r?\n/).filter(Boolean);
  function detailAt(rawText,i){var lines=String(rawText||'').split(/\r?\n/);if(lines.length<=i)return '';var line=lines[i].replace(/^\s*\d+\.\s*/,'');var p=line.indexOf(':');return p>=0?line.slice(p+1).trim():line.trim();}
  rec.pesertaList=raw.map(function(line,i){
   var clean=line.replace(/^\s*\d+\.\s*/,'').trim(),m=clean.match(/^(.*?)\s*\(([LP])\)(?:\s*-\s*(.*))?$/i);
   return {nama:m?m[1].trim():clean,gender:m?m[2].toUpperCase():'',role:m&&m[3]?m[3].trim():'',tempatLahir:detailAt(val('Tempat Lahir Peserta'),i),tanggalLahir:detailAt(val('Tanggal Lahir Peserta'),i),alamat:detailAt(val('Alamat Rumah Peserta'),i),wa:detailAt(val('WA Peserta'),i)};
  });
  rec.peserta=String(val('Daftar Peserta')||'');rec.tempat_lahir_peserta=String(val('Tempat Lahir Peserta')||'');
  rec.tanggal_lahir_peserta=String(val('Tanggal Lahir Peserta')||'');rec.alamat_peserta=String(val('Alamat Rumah Peserta')||'');rec.wa_peserta=String(val('WA Peserta')||'');
 }
 var oldId=getDriveFileIdFromUrl_(String(val('PDF Formulir (link Drive)')||''));
 var newPdf=createRegistrationPdf(rec,folder);assertPrivateIndividualFile_(newPdf);
 var pdfCol=headers.indexOf('PDF Formulir (link Drive)')+1,archiveCol=headers.indexOf('Arsip PDF Panitia (link Drive)')+1,filesCol=headers.indexOf('Berkas (link Drive)')+1;
 if(pdfCol)sh.getRange(row,pdfCol).setValue(newPdf.getUrl());
 if(archiveCol)sh.getRange(row,archiveCol).setValue(newPdf.getUrl());
 if(filesCol){var oldLines=String(val('Berkas (link Drive)')||'').split(/\r?\n/).filter(function(x){return !/^PDF formulir:/i.test(x);});oldLines.unshift('PDF formulir: '+newPdf.getUrl());sh.getRange(row,filesCol).setValue(oldLines.join('\n'));}
 if(savedPdf.length){
  queuePdfMerge_(newPdf,rec,folder,{oldPdfIds:oldId?[oldId]:[]});
  adminAudit_(admin,'BANGUN ULANG PDF DAN JADWALKAN PENGGABUNGAN',type,row,reg,nama+' | '+savedPdf.length+' lampiran PDF');
  return {ok:true,message:'PDF formulir baru dibuat. Penggabungan '+savedPdf.length+' lampiran PDF dijadwalkan; PDF final akan menggantikan versi sebelumnya setelah selesai.'};
 }
 if(oldId&&oldId!==newPdf.getId()){try{DriveApp.getFileById(oldId).setTrashed(true);}catch(e){}}
 adminAudit_(admin,'BANGUN ULANG PDF',type,row,reg,nama+' | PDF baru tanpa lampiran PDF tambahan');
 return {ok:true,message:'PDF berhasil dibangun ulang dengan seluruh data dan lampiran gambar terbaru.'};
}

function adminRetryPdfMerge_(d){
  var admin=requireAdmin_(d.token),type=String(d.type||'');
  if(type!=='kontingen'&&type!=='cabang')throw new Error('Jenis data tidak valid.');
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG[type].sheet);if(!sh)throw new Error('Data tidak ditemukan.');
  var row=resolveAdminRecordRow_(sh,d),headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0],vals=sh.getRange(row,1,1,sh.getLastColumn()).getValues()[0];
  function g(h){var i=headers.indexOf(h);return i<0?'':vals[i];}
  var mainId=getDriveFileIdFromUrl_(g('PDF Formulir (link Drive)'));if(!mainId)throw new Error('PDF formulir tidak tercatat pada data ini.');
  var mainFile=DriveApp.getFileById(mainId),folder=mainFile.getParents().next();
  if(wasPdfMerged_(mainFile)||folder.getFilesByName('DRAFT_SEBELUM_GABUNG_'+mainFile.getName()).hasNext())throw new Error('PDF ini sudah pernah digabung. Gabung ulang dibatalkan agar lampiran tidak ganda.');
  if(isMergePending_(String(g('No. Registrasi')||'')))throw new Error('Penggabungan untuk data ini sudah dijadwalkan. Tunggu beberapa menit.');
  var pdfIds=[],seenPdfIds={};
  parsePortalFiles_(g('Berkas (link Drive)')).forEach(function(x){
    if(/^pdf formulir$/i.test(x.label)||isStandaloneCompetitionDocumentLabel_(x.label))return;
    var id=getDriveFileIdFromUrl_(x.url);if(!id||seenPdfIds[id])return;
    try{if(DriveApp.getFileById(id).getMimeType()===MimeType.PDF){seenPdfIds[id]=true;pdfIds.push(id);}}catch(e){}
  });
  if(!pdfIds.length)throw new Error('Tidak ada lampiran PDF yang perlu digabung pada data ini.');
  var rd={type:type,reg:String(g('No. Registrasi')||''),nama:String(g('Nama Kontingen')||''),cabang:String(g('Cabang Lomba')||''),kategori:String(g('Kategori')||''),savedAttachments:pdfIds.map(function(id){return {id:id,mime:'application/pdf'};})};
  queuePdfMerge_(mainFile,rd,folder);
  adminAudit_(admin,'GABUNG ULANG PDF',type,row,rd.reg,rd.nama+' | '+pdfIds.length+' lampiran PDF');
  return {ok:true,message:'Penggabungan dijadwalkan ('+pdfIds.length+' lampiran PDF). Hasil final muncul di Portal Kontingen dalam 1–3 menit.'};
}
// Jalankan SEKALI dari editor Apps Script: memberi izin trigger/Drive/Mail dan memeriksa library PDFApp.
function siapkanPenggabunganPdf(){
  var laporan={pdfAppTerpasang:pdfMergeReady_(),trigger:ScriptApp.getProjectTriggers().length,kuotaEmail:MailApp.getRemainingDailyQuota(),folderUji:testFolder_().getUrl()};
  Logger.log(JSON.stringify(laporan));
  if(!laporan.pdfAppTerpasang)Logger.log('PERINGATAN: library PDFApp belum terpasang. Pendaftaran tetap berjalan, tetapi lampiran PDF belum akan digabung otomatis.');
  return laporan;
}
