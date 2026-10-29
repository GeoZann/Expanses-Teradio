// Εισαγωγή κινήσεων από αρχεία τραπεζών, σε δύο στρώματα:
// 1. Πηγές: κάθε τράπεζα μετατρέπεται σε κοινή κίνηση {uid, src, d, a, m, type, hint, pend, tr}
//    (a με πρόσημο: αρνητικό = χρέωση).
// 2. Επεξεργασία: διπλοεγγραφές, μεταφορές μεταξύ λογαριασμών, κατηγορία από κανόνες.
var TI=(function(){
var XLSX_URL="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";

/* ---------- Ανάγνωση αρχείου σε πίνακα γραμμών ---------- */

function loadXLSX(){
if(window.XLSX)return Promise.resolve();
return new Promise(function(ok,no){var s=document.createElement("script");s.src=XLSX_URL;s.onload=ok;
s.onerror=function(){no(new Error("Δεν φόρτωσε ο αναγνώστης Excel. Χρειάζεται σύνδεση στο ίντερνετ την πρώτη φορά."))};document.head.appendChild(s)})
}
function decode(buf){
try{return new TextDecoder("utf-8",{fatal:true}).decode(buf).replace(/^\uFEFF/,"")}
catch(e){return new TextDecoder("windows-1253").decode(buf)}
}
function parseCSV(txt){
var first=txt.split(/\r?\n/,1)[0],sep=[";",",","\t"].sort(function(a,b){return first.split(b).length-first.split(a).length})[0];
var rows=[],row=[],f="",q=false;
for(var i=0;i<txt.length;i++){var c=txt[i];
if(q){if(c==='"'){if(txt[i+1]==='"'){f+='"';i++}else q=false}else f+=c}
else if(c==='"')q=true;
else if(c===sep){row.push(f);f=""}
else if(c==="\n"||c==="\r"){if(c==="\r"&&txt[i+1]==="\n")i++;row.push(f);rows.push(row);row=[];f=""}
else f+=c}
if(f||row.length){row.push(f);rows.push(row)}
return rows.filter(function(r){return r.some(function(x){return String(x).trim()})})
}
async function readRows(file){
var buf=await file.arrayBuffer(),b=new Uint8Array(buf.slice(0,8)),head=decode(buf.slice(0,600)).toLowerCase();
var excel=(b[0]===0x50&&b[1]===0x4B)||(b[0]===0xD0&&b[1]===0xCF)||head.indexOf("<table")>=0||head.indexOf("<html")>=0;
if(!excel)return parseCSV(decode(buf));
await loadXLSX();
var wb=XLSX.read(buf,{type:"array"}),ws=wb.Sheets[wb.SheetNames[0]];
return XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:""}).filter(function(r){return r.some(function(x){return String(x).trim()})})
}

/* ---------- Βοηθητικά ---------- */

function pad(n){return String(n).padStart(2,"0")}
// Ημερομηνία από κελί: αριθμός Excel, Date, «24/10/2026 09:03» ή «2026-10-24 09:03:00».
function toDate(v){
if(typeof v==="number"&&v>20000&&v<80000){var x=new Date(Math.round((v-25569)*864e5));return x.getUTCFullYear()+"-"+pad(x.getUTCMonth()+1)+"-"+pad(x.getUTCDate())}
if(v instanceof Date)return v.getFullYear()+"-"+pad(v.getMonth()+1)+"-"+pad(v.getDate());
var s=String(v).trim(),r=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
if(r)return r[1]+"-"+pad(r[2])+"-"+pad(r[3]);
r=s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})/);
if(r)return(r[3].length===2?"20"+r[3]:r[3])+"-"+pad(r[2])+"-"+pad(r[1]);
return null
}
// Ποσό: δέχεται 1.234,56 / -23,00 / 1234.56 / αριθμό.
function toNum(v){
if(typeof v==="number")return v;
var s=String(v).replace(/[\s€]|EUR/g,"").replace(/\u2212/g,"-");
if(!s)return NaN;
var c=s.lastIndexOf(","),d=s.lastIndexOf(".");
if(c>=0&&d>=0)s=c>d?s.replace(/\./g,"").replace(",","."):s.replace(/,/g,"");
else if(c>=0)s=s.replace(",",".");
else if(/^-?\d{1,3}(\.\d{3})+$/.test(s))s=s.replace(/\./g,"");
return parseFloat(s)
}
function col(hdr,names){
var h=hdr.map(function(x){return String(x).trim().toLowerCase()});
for(var i=0;i<names.length;i++){var k=h.indexOf(names[i].toLowerCase());if(k>=0)return k}
return -1
}
function headerRow(rows,name){
for(var i=0;i<Math.min(rows.length,15);i++)if(rows[i].some(function(x){return String(x).trim().toLowerCase()===name.toLowerCase()}))return i;
return -1
}

/* ---------- Πηγές ---------- */

// Πειραιώς (winbank, Excel): Κατηγορία | Περιγραφή Συναλλαγής | Ημ/νία Συναλλαγής | Ημ/νία Αξίας | Αριθμός Λογαριασμού | Αριθμός Παραστατικού | Ποσό | Νόμισμα
function piraeus(rows){
var h=headerRow(rows,"Περιγραφή Συναλλαγής");if(h<0)return null;
var H=rows[h],iD=col(H,["Ημ/νία Συναλλαγής","Ημ/νία Αξίας"]),iT=col(H,["Περιγραφή Συναλλαγής"]),iA=col(H,["Ποσό"]),iP=col(H,["Αριθμός Παραστατικού"]),iC=col(H,["Κατηγορία"]);
if(iD<0||iA<0)return null;
return rows.slice(h+1).map(function(r){
var desc=String(r[iT]).trim(),m=desc.match(/^(.{2,30}?)\s+-\s*(.+)$/),type=m?m[1]:"",merch=(m?m[2]:desc).replace(/\s+PO\d{6,}$/i,"").trim();
var hint=String(r[iC]||"").trim();
return{src:"Πειραιώς",d:toDate(r[iD]),a:toNum(r[iA]),m:merch,type:type,hint:hint==="Χωρίς Κατηγορία"?"":hint,key:"pir|"+String(r[iP]).trim()+"|"+toDate(r[iD])+"|"+r[iA]}
})
}
// Revolut (CSV): Type, Product, Started Date, Completed Date, Description, Amount, Fee, Currency, State, Balance
function revolut(rows){
var h=headerRow(rows,"Started Date");if(h<0)return null;
var H=rows[h],iT=col(H,["Type"]),iP=col(H,["Product"]),iD=col(H,["Started Date"]),iM=col(H,["Description"]),iA=col(H,["Amount"]),iF=col(H,["Fee"]),iC=col(H,["Currency"]),iS=col(H,["State"]);
return rows.slice(h+1).filter(function(r){var st=String(r[iS]).toUpperCase();return st!=="REVERTED"&&st!=="DECLINED"&&st!=="FAILED"}).map(function(r){
var a=parseFloat(r[iA])-(parseFloat(r[iF])||0),cur=String(r[iC]||"EUR").trim(),type=String(r[iT]).trim(),prod=String(r[iP]||"");
return{src:"Revolut",d:toDate(r[iD]),a:a,m:String(r[iM]).trim()+(cur&&cur!=="EUR"?" ("+cur+")":""),type:type,hint:"",pend:String(r[iS]).toUpperCase()==="PENDING",
tr:/^(TOPUP|EXCHANGE)$/i.test(type)||/savings|deposit/i.test(prod),key:"rev|"+r[iD]+"|"+r[iM]+"|"+r[iA]}
})
}
// Alpha (CSV, ;): Ημερομηνία ώρα ; Έμπορος ; Κατηγορία τράπεζας ; Τύπος ; Ποσό ; Κατάσταση
function alpha(rows){
var L=rows.filter(function(r){return r.length>=5&&/^\d{1,2}\/\d{1,2}\/\d{4}/.test(String(r[0]).trim())&&!isNaN(toNum(r[4]))});
if(!L.length)return null;
return L.map(function(r){
return{src:"Alpha",d:toDate(r[0]),a:toNum(r[4]),m:String(r[1]).trim(),type:String(r[3]).trim(),hint:String(r[2]).trim(),pend:/επεξεργασ/i.test(r[5]||""),key:"alpha|"+String(r[0]).trim()+"|"+String(r[1]).trim()+"|"+String(r[4]).trim()}
})
}
var SOURCES=[piraeus,revolut,alpha];

async function readFile(file){
var rows=await readRows(file),L=null;
for(var i=0;i<SOURCES.length&&!L;i++)L=SOURCES[i](rows);
if(!L)throw new Error("Δεν αναγνώρισα τη μορφή του «"+file.name+"». Υποστηρίζονται αρχεία από Alpha, Πειραιώς και Revolut.");
// Μοναδικό ID: ίδιες γραμμές στο ίδιο αρχείο (π.χ. δύο εισιτήρια το ίδιο λεπτό) παίρνουν αύξοντα αριθμό.
var seen={};
return L.filter(function(t){return t.d&&!isNaN(t.a)&&t.a!==0}).map(function(t){
var n=seen[t.key]=(seen[t.key]||0)+1;t.uid=t.key+"|"+n;delete t.key;return t})
}

/* ---------- Επεξεργασία ---------- */

// Κλειδί εμπόρου για τους κανόνες: οι δύο πρώτες λέξεις, χωρίς αριθμούς και σημεία.
function mkey(m){
return String(m).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/\(.*?\)/g," ").replace(/\d{3,}/g," ").replace(/[^a-zα-ω0-9&]+/g," ").trim().split(" ").slice(0,2).join(" ")
}
// Αρχικοί κανόνες με λέξεις-κλειδιά. Ελέγχονται με τη σειρά, κερδίζει ο πρώτος.
var KW=[
["Φαγητό έξω",/wolt|e-?food|mcdonald|starbucks|coffee ?island|everest|goody|γρηγορη|grigori|mikel|cafe|καφε|coffee|restaur|pizza|souvlak|σουβλ|burger|bakery|φουρν|bar\b/],
["Σούπερ μάρκετ",/lidl|sklavenit|σκλαβενιτ|vasilopoul|βασιλοπουλ|\bab\b|masout|μασουτ|my ?market|galaxias|γαλαξιας|kritikos|κρητικος|bazaar|market ?in|carrefour|super ?market|σουπερ/],
["Μεταφορές",/oasa|οασα|athens transport|uber|free ?now|\bbeat\b|taxi|ταξι|shell|\bbp\b|\beko\b|avin|aegean|revoil|elin|ktel|κτελ|hellenic train|trainose|αττικη οδος|attiki odos|parking|smart park|fuel|καυσιμ/],
["Λογαριασμοί",/\bδεη\b|\bdei\b|ppc|cosmote|vodafone|\bnova\b|wind|eydap|ευδαπ|protergia|heron|elpedison|netflix|spotify|youtube|google|apple\.com|icloud|disney/],
["Υγεία",/φαρμακ|pharm|κλινικ|clinic|ιατρ|γιατρ|οδοντ|hospital|νοσοκ/],
["Διασκέδαση",/cinema|village|σινεμα|steam|playstation|ticket|viva\.gr|more\.com|nintendo|xbox/],
["Αγορές",/amazon|skroutz|ikea|zara|h&m|jumbo|public|plaisio|πλαισιο|kotsovolos|κωτσοβολ|e-?shop|aliexpress|temu|shein|decathlon/]
];
// Κατηγορίες της τράπεζας (Alpha) ως τελευταία ένδειξη.
var HINT=[["Σούπερ μάρκετ",/σουπερ|super/i],["Υγεία",/υγει/i],["Λογαριασμοί",/λογαριασμ|τηλεπικ|ενεργ/i],["Μεταφορές",/καυσιμ|μεταφορ|οχημ/i],["Αγορές",/ενδυσ|ηλεκτρον|αγορ/i]];
var TRANSFER=/μεταφορ|transfer|top.?up|exchange|revolut|αναληψ|ανάληψ|\batm\b|πληρωμ\S* καρτ|εξοφλ/i;

function suggest(t,S){
var k=mkey(t.m),txt=String(t.m).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"");
if(S.rules&&S.rules[k])return{c:S.rules[k],why:"rule"};
if(t.a>0)return{c:/μισθ|salary|payroll/.test(txt)?"Μισθός":"Άλλα έσοδα",why:"kw"};
for(var i=0;i<KW.length;i++)if(KW[i][1].test(txt))return{c:KW[i][0],why:"kw"};
for(i=0;i<HINT.length;i++)if(t.hint&&HINT[i][1].test(t.hint))return{c:HINT[i][0],why:"bank"};
return{c:"Άλλα έξοδα",why:"none"}
}
function days(a,b){return Math.abs(new Date(a)-new Date(b))/864e5}

// Από τις γραμμές των αρχείων βγάζει τη λίστα προεπισκόπησης.
function process(L,S){
var have={};
S.tx.forEach(function(t){if(t.uid)have[t.uid]=1});
(S.ign||[]).forEach(function(u){have[u]=1});
var dup=0,items=L.filter(function(t){if(have[t.uid]){dup++;return false}have[t.uid]=1;return true});
items.forEach(function(t){if(!t.tr)t.tr=TRANSFER.test(t.type+" "+t.m)&&!/μισθ|salary|payroll/i.test(t.m)});
// Μεταφορά μεταξύ δικών σου λογαριασμών: ίδιο ποσό, αντίθετο πρόσημο, άλλη τράπεζα, έως 3 ημέρες.
var pool=items.concat(S.tx.filter(function(t){return t.src}).map(function(t){return{src:t.src,d:t.d,a:t.k==="i"?t.a:-t.a,old:1}}));
items.forEach(function(t){
if(t.tr)return;
var o=pool.find(function(x){return x!==t&&x.src!==t.src&&!x.used&&Math.abs(x.a+t.a)<0.005&&days(x.d,t.d)<=3});
if(o){t.tr=true;t.used=true;o.used=true;if(!o.old)o.tr=true}
});
items.forEach(function(t){var s=suggest(t,S);t.k=t.a>0?"i":"e";t.c=s.c;t.why=s.why;t.mk=mkey(t.m);t.on=!t.tr;delete t.used});
items.sort(function(a,b){return a.d<b.d?1:-1});
return{items:items,dup:dup}
}

return{readFile:readFile,process:process,mkey:mkey}
})();
