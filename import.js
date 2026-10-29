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
// \u03A3\u03C4\u03BF\u03B9\u03C7\u03B5\u03AF\u03B1 \u03C4\u03B7\u03C2 \u03C4\u03B5\u03BB\u03B5\u03C5\u03C4\u03B1\u03AF\u03B1\u03C2 \u03B1\u03BD\u03AC\u03B3\u03BD\u03C9\u03C3\u03B7\u03C2, \u03B3\u03B9\u03B1 \u03C4\u03BF \u03BC\u03AE\u03BD\u03C5\u03BC\u03B1 \u03CC\u03C4\u03B1\u03BD \u03B4\u03B5\u03BD \u03B1\u03BD\u03B1\u03B3\u03BD\u03C9\u03C1\u03AF\u03B6\u03B5\u03C4\u03B1\u03B9 \u03AD\u03BD\u03B1 \u03B1\u03C1\u03C7\u03B5\u03AF\u03BF.
var info={};
function decode(buf){
var b=new Uint8Array(buf.slice(0,2)),z=new Uint8Array(buf.slice(0,200)).filter(function(x){return x===0}).length;
if(b[0]===0xFF&&b[1]===0xFE)return info.enc="utf-16le",new TextDecoder("utf-16le").decode(buf);
if(b[0]===0xFE&&b[1]===0xFF)return info.enc="utf-16be",new TextDecoder("utf-16be").decode(buf);
// UTF-16 \u03C7\u03C9\u03C1\u03AF\u03C2 BOM: \u03C0\u03BF\u03BB\u03BB\u03AC \u03BC\u03B7\u03B4\u03B5\u03BD\u03B9\u03BA\u03AC bytes. \u0396\u03C5\u03B3\u03AC \u03AE \u03BC\u03BF\u03BD\u03AC \u03B4\u03B5\u03AF\u03C7\u03BD\u03BF\u03C5\u03BD \u03C4\u03B7 \u03C3\u03B5\u03B9\u03C1\u03AC.
if(z>20){var nb=new Uint8Array(buf.slice(0,200)),odd=0;for(var i=1;i<nb.length;i+=2)if(nb[i]===0)odd++;
info.enc=odd>z/2?"utf-16le":"utf-16be";return new TextDecoder(info.enc).decode(buf)}
try{info.enc="utf-8";return new TextDecoder("utf-8",{fatal:true}).decode(buf).replace(/^\uFEFF/,"")}
catch(e){info.enc="windows-1253";return new TextDecoder("windows-1253").decode(buf)}
}
function parseCSV(txt){
// \u0394\u03B9\u03B1\u03C7\u03C9\u03C1\u03B9\u03C3\u03C4\u03B9\u03BA\u03CC: \u03B1\u03C5\u03C4\u03CC \u03C0\u03BF\u03C5 \u03B4\u03AF\u03BD\u03B5\u03B9 \u03C4\u03B9\u03C2 \u03C0\u03B5\u03C1\u03B9\u03C3\u03C3\u03CC\u03C4\u03B5\u03C1\u03B5\u03C2 \u03C3\u03C4\u03AE\u03BB\u03B5\u03C2 \u03C3\u03B5 \u03BA\u03AC\u03C0\u03BF\u03B9\u03B1 \u03B1\u03C0\u03CC \u03C4\u03B9\u03C2 \u03C0\u03C1\u03CE\u03C4\u03B5\u03C2 \u03B3\u03C1\u03B1\u03BC\u03BC\u03AD\u03C2
// (\u03B7 \u03C0\u03C1\u03CE\u03C4\u03B7 \u03B3\u03C1\u03B1\u03BC\u03BC\u03AE \u03BC\u03C0\u03BF\u03C1\u03B5\u03AF \u03BD\u03B1 \u03B5\u03AF\u03BD\u03B1\u03B9 \u03C4\u03AF\u03C4\u03BB\u03BF\u03C2 \u03BC\u03B5 \u03BB\u03AF\u03B3\u03B1 \u03C0\u03B5\u03B4\u03AF\u03B1).
var head=txt.split(/\r?\n/).slice(0,30),sep=[";","\t",","].map(function(s){return[s,Math.max.apply(null,head.map(function(l){return l.split(s).length}))]})
.sort(function(a,b){return b[1]-a[1]})[0][0];
info.sep=sep==="\t"?"tab":sep;
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
info={size:buf.byteLength};
if(!excel)return parseCSV(decode(buf));
info.enc="excel";
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
// Τίτλοι στηλών χωρίς τόνους, κεφαλαία, εισαγωγικά, αλλαγές γραμμής και διπλά/αόρατα κενά.
function norm(x){return String(x).replace(/﻿/g,"").normalize("NFD").replace(/[̀-ͯ]/g,"").toLowerCase().replace(/["']/g,"").replace(/\s+/g," ").trim()}
// Πρώτα ακριβές ταίριασμα, μετά τίτλος που περιέχει το όνομα (π.χ. «Ποσό (€)»).
function col(hdr,names){
var h=hdr.map(norm),i,k;
for(i=0;i<names.length;i++){k=h.indexOf(norm(names[i]));if(k>=0)return k}
for(i=0;i<names.length;i++){var n=norm(names[i]);for(k=0;k<h.length;k++)if(h[k].indexOf(n)>=0)return k}
return -1
}
function headerRow(rows,name){
var n=norm(name);
for(var i=0;i<Math.min(rows.length,25);i++)if(rows[i].some(function(x){return norm(x).indexOf(n)>=0}))return i;
return -1
}

/* ---------- Πηγές ---------- */

// Πειραιώς (winbank, Excel), δύο μορφές με κοινή στήλη «Περιγραφή Συναλλαγής»:
// κάρτα:      Κατηγορία | Περιγραφή («Αγορά -ΕΜΠΟΡΟΣ PO…») | Ημ/νία Συναλλαγής | Ημ/νία Αξίας | Αριθμός Λογαριασμού | Αριθμός Παραστατικού | Ποσό | Νόμισμα
// λογαριασμός: Κατηγορία | Περιγραφή («ΑΓΟΡΑ ΜΕ ΚΑΡΤΑ») | Ημ/νία Συναλλαγής | Σχόλια / Κωδικός Αναφοράς | Ποσό | Νόμισμα | Προοδευτικό Λογιστικό Υπόλοιπο | Νόμισμα
function piraeus(rows){
var h=headerRow(rows,"Περιγραφή Συναλλαγής");if(h<0)return null;
var H=rows[h],iD=col(H,["Ημ/νία Συναλλαγής","Ημ/νία Αξίας"]),iT=col(H,["Περιγραφή Συναλλαγής"]),iA=col(H,["Ποσό"]),iP=col(H,["Αριθμός Παραστατικού"]),iC=col(H,["Κατηγορία"]),
iN=col(H,["Σχόλια / Κωδικός Αναφοράς","Σχόλια"]),iB=col(H,["Προοδευτικό Λογιστικό Υπόλοιπο"]);
if(iD<0||iA<0)return null;
var L=rows.slice(h+1).map(function(r){
var desc=String(r[iT]).trim(),type,merch,key,hint=String(r[iC]||"").trim();
if(iN>=0){
// «MCDONALD S SPATA23,00 EUR4410…» → «MCDONALD S SPATA», «B/O ΟΝΟΜΑ / F26TI…» → «ΟΝΟΜΑ»
var note=String(r[iN]).trim();type=desc;
merch=note.replace(/\s*\d{1,3}(?:[.,]\d{3})*[.,]\d{2}\s*[A-Z]{3}.*$/,"").replace(/^(B\/O|BEN|ΠΡΟΣ|ΑΠΟ)[:\s]+/i,"").replace(/\s+\/\s+.*$/,"").trim()||desc;
key="pira|"+toDate(r[iD])+"|"+desc+"|"+note+"|"+r[iA]+"|"+r[iB];
}else{
var m=desc.match(/^(.{2,30}?)\s+-\s*(.+)$/);type=m?m[1]:"";merch=(m?m[2]:desc).replace(/\s+PO\d{6,}$/i,"").trim();
key="pir|"+String(r[iP]).trim()+"|"+toDate(r[iD])+"|"+r[iA];
}
return{src:"Πειραιώς",d:toDate(r[iD]),a:toNum(r[iA]),m:merch,type:type,hint:hint==="Χωρίς Κατηγορία"?"":hint,key:key,bal:iB>=0?toNum(r[iB]):NaN}
});
// Υπόλοιπο λογαριασμού: της πιο πρόσφατης γραμμής (το αρχείο μπορεί να είναι με αύξουσα ή φθίνουσα σειρά).
var ok=L.filter(function(t){return t.d&&!isNaN(t.bal)});
if(ok.length){var f=ok[0],l=ok[ok.length-1],x=l.d>f.d?l:f;L.bal={src:"Πειραιώς",b:x.bal,d:x.d}}
return L
}
// Alpha λογαριασμός (CSV): λίγες γραμμές με υπόλοιπα, μετά
// Α/Α ; Ημ/νία ; Αιτιολογία ; Κατάστημα ; Τοκισμός από ; Αρ. συναλλαγής ; Ποσό ; Πρόσημο ποσού (Χ χρέωση / Π πίστωση)
function alphaAcc(rows){
var h=headerRow(rows,"Πρόσημο ποσού");if(h<0)return null;
var H=rows[h],iD=col(H,["Ημ/νία"]),iV=col(H,["Τοκισμός από"]),iM=col(H,["Αιτιολογία"]),iN=col(H,["Αρ. συναλλαγής"]),iA=col(H,["Ποσό"]),iS=col(H,["Πρόσημο ποσού"]);
var L=rows.slice(h+1).map(function(r){
// Μια αγορά με κάρτα περνά στον λογαριασμό μερικές μέρες μετά. Ο «Τοκισμός από» κρατά την πραγματική ημέρα.
var d=toDate(r[iD]),v=iV>=0?toDate(r[iV]):null,m=String(r[iM]).trim();
return{src:"Alpha",d:v&&v<d?v:d,a:toNum(r[iA])*(/^[ΧX]/i.test(String(r[iS]).trim())?-1:1),m:m,type:m,hint:"",key:"alphaa|"+String(r[iN]).trim()+"|"+r[iD]+"|"+r[iA]+"|"+r[iS]}
});
// «Νέο Μεικτό Υπόλοιπο ; ; 8.408,11 ; Π» και «Ημ/νία ; 29/10/2026 13:22» στις πρώτες γραμμές.
var b=null,bd=null;
rows.slice(0,h).forEach(function(r){var k=norm(r[0]);
if(/^ημ\/νια$/.test(k))bd=toDate(r[1]);
if(/νεο.*υπολοιπο/.test(k)){var n=r.slice(1).map(toNum).filter(function(x){return!isNaN(x)})[0],neg=r.some(function(x){return/^[ΧX]$/.test(String(x).trim())});if(n!==undefined)b=neg?-n:n}});
if(b!==null)L.bal={src:"Alpha",b:b,d:bd||L.reduce(function(a,t){return t.d>a?t.d:a},"")};
return L
}
// Revolut (CSV): Type, Product, Started Date, Completed Date, Description, Amount, Fee, Currency, State, Balance
function revolut(rows){
var h=headerRow(rows,"Started Date");if(h<0)return null;
var H=rows[h],iT=col(H,["Type"]),iP=col(H,["Product"]),iD=col(H,["Started Date"]),iE=col(H,["Completed Date"]),iM=col(H,["Description"]),iA=col(H,["Amount"]),iF=col(H,["Fee"]),iC=col(H,["Currency"]),iS=col(H,["State"]),iB=col(H,["Balance"]);
var R=rows.slice(h+1).filter(function(r){var st=String(r[iS]).toUpperCase();return st!=="REVERTED"&&st!=="DECLINED"&&st!=="FAILED"});
var L=R.map(function(r){
var a=parseFloat(r[iA])-(parseFloat(r[iF])||0),cur=String(r[iC]||"EUR").trim(),type=String(r[iT]).trim(),prod=String(r[iP]||"");
return{src:"Revolut",d:toDate(r[iD]),a:a,m:String(r[iM]).trim()+(cur&&cur!=="EUR"?" ("+cur+")":""),type:type,hint:"",pend:String(r[iS]).toUpperCase()==="PENDING",
tr:/^(TOPUP|EXCHANGE)$/i.test(type)||/savings|deposit/i.test(prod),cash:/^ATM$/i.test(type)?1:0,key:"rev|"+r[iD]+"|"+r[iM]+"|"+r[iA]}
});
// Υπόλοιπο: η πιο πρόσφατη ολοκληρωμένη κίνηση σε ευρώ του τρεχούμενου λογαριασμού.
var last=null;
R.forEach(function(r){if(String(r[iS]).toUpperCase()==="COMPLETED"&&String(r[iC]).trim()==="EUR"&&!/savings|deposit/i.test(r[iP]||"")&&r[iB]!==""&&(!last||String(r[iE])>String(last[iE])))last=r});
if(last&&iB>=0)L.bal={src:"Revolut",b:parseFloat(last[iB]),d:toDate(last[iE])};
return L
}
// Alpha κάρτα (CSV, ;): Ημερομηνία ώρα ; Έμπορος ; Κατηγορία τράπεζας ; Τύπος ; Ποσό ; Κατάσταση
function alpha(rows){
var L=rows.filter(function(r){return r.length>=5&&/^\d{1,2}\/\d{1,2}\/\d{4}/.test(String(r[0]).trim())&&!isNaN(toNum(r[4]))});
if(!L.length)return null;
return L.map(function(r){
return{src:"Alpha",d:toDate(r[0]),a:toNum(r[4]),m:String(r[1]).trim(),type:String(r[3]).trim(),hint:String(r[2]).trim(),pend:/επεξεργασ/i.test(r[5]||""),key:"alpha|"+String(r[0]).trim()+"|"+String(r[1]).trim()+"|"+String(r[4]).trim()}
})
}
var SOURCES=[piraeus,alphaAcc,revolut,alpha];

// Επιστρέφει τις κινήσεις του αρχείου και, αν το αρχείο το δίνει, το υπόλοιπο του λογαριασμού (L.bal).
async function readFile(file){
var rows=await readRows(file),L=null;
for(var i=0;i<SOURCES.length&&!L;i++)L=SOURCES[i](rows);
if(!L){
// Τεχνικά στοιχεία χωρίς προσωπικά δεδομένα: μόνο η πρώτη λέξη από τις πρώτες γραμμές.
var cols=rows.reduce(function(a,r){return Math.max(a,r.length)},0),firsts=rows.slice(0,7).map(function(r){return String(r[0]).trim().slice(0,14)}).join(" | ");
throw new Error("Δεν αναγνώρισα τη μορφή του «"+file.name+"». Υποστηρίζονται αρχεία από Alpha, Πειραιώς και Revolut. ["+
info.enc+(info.sep?", διαχωριστικό "+info.sep:"")+", "+info.size+" bytes, "+rows.length+" γραμμές, έως "+cols+" στήλες: "+firsts+"]")}
// Μοναδικό ID: ίδιες γραμμές στο ίδιο αρχείο (π.χ. δύο εισιτήρια το ίδιο λεπτό) παίρνουν αύξοντα αριθμό.
var seen={},out=L.filter(function(t){return t.d&&!isNaN(t.a)&&t.a!==0}).map(function(t){
var n=seen[t.key]=(seen[t.key]||0)+1;t.uid=t.key+"|"+n;delete t.key;delete t.bal;return t});
out.bal=L.bal&&!isNaN(L.bal.b)?L.bal:null;
return out
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
var TRANSFER=/μεταφορ|transfer|top.?up|exchange|revolut|πληρωμ\S* καρτ|εξοφλ/i;
// Αναλήψεις και καταθέσεις μετρητών: δεν είναι έξοδο ή έσοδο, μετακινούν χρήματα από/προς τα μετρητά.
var WITHDRAW=/αναληψ|ανάληψ|\batm\b|cash withdrawal/i,DEPOSIT=/καταθεσ\S* μετρητ|κατάθεσ\S* μετρητ|cash deposit/i;

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
// Όσες είχες αφήσει χωρίς τικ σε παλιότερη εισαγωγή ξαναεμφανίζονται χωριστά (t.old), για να τις επαναφέρεις αν θέλεις.
var PAIR=/μεταφορ|εμβασμ|transfer|top.?up|revolut|iris/i;
function process(L,S){
var have={},ign={};
S.tx.forEach(function(t){if(t.uid)have[t.uid]=1});
(S.ign||[]).forEach(function(u){ign[u]=1});
var dup=0,items=L.filter(function(t){if(have[t.uid]){dup++;return false}have[t.uid]=1;t.old=!!ign[t.uid];return true});
items.forEach(function(t){
var s=t.type+" "+t.m;
if(!t.cash)t.cash=t.a<0&&WITHDRAW.test(s)?1:t.a>0&&DEPOSIT.test(s)?-1:0;
if(!t.tr&&!t.cash)t.tr=TRANSFER.test(s)&&!/μισθ|salary|payroll/i.test(t.m)});
// Μεταφορά μεταξύ δικών σου λογαριασμών: ίδιο ποσό, αντίθετο πρόσημο, άλλη τράπεζα, έως 3 ημέρες,
// και οι δύο πλευρές μοιάζουν με μεταφορά (όχι π.χ. αγορά με κάρτα απέναντι σε έμβασμα από φίλο).
var pool=items.filter(function(t){return!t.cash&&!t.old&&PAIR.test(t.type+" "+t.m)}).concat(S.tx.filter(function(t){return t.src&&t.k!=="t"&&PAIR.test((t.type||"")+" "+t.n)})
.map(function(t){return{src:t.src,d:t.d,a:t.k==="i"?t.a:-t.a,ex:1}}));
items.forEach(function(t){
if(t.tr||t.cash||t.old||pool.indexOf(t)<0)return;
var o=pool.find(function(x){return x!==t&&x.src!==t.src&&!x.used&&Math.abs(x.a+t.a)<0.005&&days(x.d,t.d)<=3});
if(o){t.tr=true;t.used=true;o.used=true;t.pair=o.src+" "+o.d.split("-").reverse().join("/");if(!o.ex){o.tr=true;o.pair=t.src+" "+t.d.split("-").reverse().join("/")}}
});
items.forEach(function(t){var s=suggest(t,S);t.k=t.cash?"t":t.a>0?"i":"e";t.c=s.c;t.why=s.why;t.mk=mkey(t.m);t.on=!t.tr&&!t.old;delete t.used});
items.sort(function(a,b){return a.d<b.d?1:-1});
return{items:items,dup:dup}
}

return{readFile:readFile,process:process,mkey:mkey}
})();
