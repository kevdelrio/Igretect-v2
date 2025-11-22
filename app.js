/* ========= Constantes GEO ========= */
const GEO_HEADER = {
    pouvoir: "S.P.G.E.",
    spgeRef: "xxxx/xx/xxxx",
    operation: "REHABILITATION DE LA STATION D'EPURATION DE RANCE",
    igretecRef: "05 - 61860"
};
/* ========= Helpers ========= */
const $ = s => document.querySelector(s);
const tbl = document.getElementById('tbl');
const tblBody = document.querySelector("#tbl tbody");
const tableWrap = document.querySelector('.table-wrap');
const topScrollbar = document.getElementById('tableScrollTop');
const topScrollbarInner = document.getElementById('tableScrollTopInner');
const compactLabel = document.querySelector('#btnToggleCompact .label');
const notificationBar = document.querySelector("#notificationBar");
const btnParcelles = document.getElementById('btnParcelles');
let partsRaw = [], parcRaw = [], merged = [], base = [];
// NOUVEAU : Pour suivre les noms des fichiers uploadés
let uploadedFiles = { A: [], B: [] };
let manualRowCount = 1;
let compactMode = false;

function setCompactMode(enabled) {
    compactMode = enabled;
    tableWrap.classList.toggle('compact-mode', compactMode);
    if (compactLabel) {
        compactLabel.textContent = compactMode ? 'Vue complète' : 'Vue compacte';
    }
    updateTopScrollbar();
}

function updateTopScrollbar() {
    if (!topScrollbarInner || !tableWrap || !topScrollbar) return;
    const width = Math.max((tbl?.scrollWidth) || tableWrap.scrollWidth, tableWrap.clientWidth);
    topScrollbarInner.style.width = `${width}px`;
    topScrollbar.scrollLeft = tableWrap.scrollLeft;
}

function setupScrollSync() {
    if (!topScrollbar || !tableWrap) return;
    topScrollbar.addEventListener('scroll', () => { tableWrap.scrollLeft = topScrollbar.scrollLeft; });
    tableWrap.addEventListener('scroll', () => { if (topScrollbar) topScrollbar.scrollLeft = tableWrap.scrollLeft; });
    tableWrap.addEventListener('wheel', (e) => {
        if (Math.abs(e.deltaY) > Math.abs(e.deltaX)) {
            tableWrap.scrollLeft += e.deltaY;
        }
    }, { passive: true });
    updateTopScrollbar();
}

function setStatus(message, type = 'info', duration = 0) {
    notificationBar.innerHTML = '';
    notificationBar.className = 'notification';
    notificationBar.style.display = 'flex';
    let iconSvg = '';
    if (type === 'success') {
        notificationBar.classList.add('success');
        iconSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>`;
    }
    notificationBar.innerHTML = `${iconSvg}<span>${message}</span>`;
    if (duration > 0) {
        setTimeout(() => { notificationBar.style.display = 'none'; }, duration);
    }
}
function hideStatus() {
    notificationBar.style.display = 'none';
}
function esc(s) {
    return String(s ?? "").replace(/[&<>"']/g, m => ({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[m]));
}
function normalizeNumberString(input) {
    const raw = String(input ?? "").trim();
    if (!raw) return "";
    const noSpaces = raw.replace(/\s+/g, "");
    const hasComma = noSpaces.includes(',');
    const hasDot = noSpaces.includes('.');
    let normalized = noSpaces;

    if (hasComma && hasDot) {
        if (noSpaces.lastIndexOf('.') > noSpaces.lastIndexOf(',')) {
            normalized = noSpaces.replace(/,/g, '');
        } else {
            normalized = noSpaces.replace(/\./g, '').replace(',', '.');
        }
    } else if (hasComma) {
        const commaCount = (noSpaces.match(/,/g) || []).length;
        normalized = commaCount > 1 ? noSpaces.replace(/,/g, '') : noSpaces.replace(',', '.');
    }

    return normalized.replace(/[^0-9.\-]/g, '');
}

function formatSurfaceDisplay(v) {
    const normalized = normalizeNumberString(v);
    if (!normalized) return "";
    const n = Number(normalized);
    return Number.isFinite(n)
        ? n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
        : esc(v);
}

function fmtNum(v) {
    return formatSurfaceDisplay(v);
}
// La fonction pad n'est plus utilisée pour l'export Excel des surfaces, 
// mais est conservée car elle est utile pour le PDF.
const pad = (num) => String(num).padStart(2, '0');
function readText(file) {
    return new Promise((res, rej) => {
        const r = new FileReader();
        r.onload = () => res(r.result);
        r.onerror = rej;
        r.readAsText(file, 'windows-1252');
    });
}
function stripAccents(s) {
    return String(s ?? "").normalize('NFD').replace(/\p{Diacritic}/gu, '').replace(/œ/gi, 'oe').replace(/æ/gi, 'ae').replace(/\u00A0/g, ' ').replace(/[\u2000-\u200B\u2028\u2029]/g, ' ');
}
function titleCaseSmart(raw) {
    const s = stripAccents(raw).toLowerCase();
    return s.split(/(\s+|[-'’])/).map(tok => {
        if (!tok || /^\s+$/.test(tok) || tok === '-' || tok === "'" || tok === '’') return tok;
        return tok.charAt(0).toUpperCase() + tok.slice(1);
    }).join('').replace(/\s+/g, ' ').trim();
}
function toAscii(s) { return stripAccents(s).replace(/[^\x20-\x7E]/g, ''); }
function toNumber(v) {
    const normalized = normalizeNumberString(v);
    const n = Number(normalized);
    return Number.isFinite(n) ? n : 0;
}
function sumNumbers(...vals) { return vals.reduce((s, v) => s + toNumber(v), 0); }
function m2toHaACa(m2raw) {
    let m2 = Math.round(toNumber(m2raw));
    const ha = Math.floor(m2 / 10000);
    m2 -= ha * 10000;
    const a = Math.floor(m2 / 100);
    m2 -= a * 100;
    return { ha, a, ca: m2 };
}
function pickNumericByRegex(obj, regexes) {
    for (const k of Object.keys(obj || {})) {
        if (regexes.some(r => r.test(String(k).toLowerCase()))) return toNumber(obj[k]);
    }
    return 0;
}
const mapping={"1":"Terre agricole","2":"Pature","3":"Pre","4":"Jardin","5":"Terrain maraicher","6":"Pre alluvial","7":"Pre embouche","8":"Patsart","9":"Bois","10":"Verger haute tige","11":"Verger basse tige","13":"Pepiniere","14":"Exploitation de sapins de Noel","17":"Parc","18":"Terrain de sport","20":"Plaine de jeux","21":"Terrain de camping","22":"Piscine","24":"Point d’eau","25":"Mare","26":"Etang","27":"Lac","28":"Douve","29":"Fosse","30":"Pisciculture","33":"Chemin cadastre","34":"Place","35":"Terrain vain ou vague","36":"Bruyere","38":"Marais","39":"Fagne","41":"Alluvion","42":"Dune","43":"Rempart","44":"Digue","45":"Terril vain ou vague","46":"Terrain d’epandage vain ou vague","49":"Terrain d’epandage exploite","50":"Terrain industriel","51":"Chantier","52":"Quai","54":"Bassin industriel","55":"Chemin de fer","56":"Terril exploite","57":"Carriere","59":"Canal","62":"Tumulus","63":"Borne","67":"Superficie d’un batiment ordinaire","68":"Superficie d’un batiment exceptionnel","69":"Superficie d’un batiment industriel","70":"Terrain","71":"Parking","72":"Champ d’aviation","73":"Terrain militaire","74":"Cimetiere","75":"Oseraie","76":"Bassin ordinaire","77":"Cour","78":"Terrain a batir","79":"Partie de parc de stationnement","80":"Materiel et outillage non-bati","85":"Bassin de decantation","86":"Rouissoir","87":"Bassin de materiel et outillage","164":"Parties communes generales d’un batiment","165":"Parties communes specifiques d’un batiment","166":"Superficie d’un batiment (autre)","170":"Autre non bati","200":"Maison","201":"Baraquement","202":"Taudis","203":"Remise","204":"Garage","205":"Abri","206":"Toilettes","220":"Entite privative#","221":"Entites privatives","222":"Building","223":"Maison#","240":"Ferme","241":"Ecurie","242":"Pigeonnier","243":"Petit elevage","244":"Grand elevage","245":"Serre","246":"Champignonnieres","247":"Batiment rural","260":"Imprimerie","261":"Garage atelier","262":"Forge","263":"Menuiserie","264":"Lavoir","265":"Atelier","280":"Laiterie","281":"Boulangerie","282":"Charcuterie","283":"Abattoir","284":"Fabrique d’aliments a betail","285":"Fabrique de cafe","286":"Brasserie","287":"Fabrique de boissons","288":"Fabrique de tabac","289":"Meunerie","290":"Fabrique de produits alimentaires","300":"Fabrique d’habillement","301":"Usine textile","302":"Fabrique d’articles de cuir","303":"Fabrique de meubles","304":"Fabrique de jouets","305":"Papeterie","306":"Fabrique d’articles usuels","320":"Briqueterie","321":"Cimenterie","322":"Scierie","323":"Fabrique de couleurs","324":"Fabrique de materiaux de construction","340":"Metallurgie","341":"Haut fourneau","342":"Four chaux","343":"Atelier de construction","344":"Fabrique de materiel electrique","345":"Raffinerie de petrole","346":"Usine chimique","347":"Fabrique de caoutchouc","348":"Glaciere","349":"Verrerie","350":"Fabrique de plastique","351":"Fabrique de ceramique","352":"Charbonnage","353":"Centrale electrique","354":"Usine de gaz","355":"Gazometre","356":"Cokerie","357":"Batiment industriel","370":"Hangar","371":"Entrepot","372":"Cabine electrique","373":"Pylone","374":"Cabine a gaz","375":"Cabine","376":"Reservoir","377":"Silo","378":"Centre de recherche","379":"Sechoir","380":"Installation frigorifique","381":"Materiel et outillage bati","382":"Exploitation industrielle#","400":"Banque","401":"Bourse","402":"Batiment de bureaux","403":"Cafe","404":"Hotel","405":"Restaurant","406":"Salle de fetes","407":"Maison de commerce","408":"Grand magasin","409":"Garage depot","410":"Batiment de parking","411":"Station-service","412":"Marche couvert","413":"Salle d’exposition","414":"Kiosque","415":"Batiment pour animaux","420":"Maison communale","421":"Batiment de gouvernement","422":"Palais royal","423":"Batiment de justice","424":"Batiment penitentiaire","425":"Legation","426":"Batiment de police","427":"Batiment militaire","428":"Station","429":"Abri de transports","430":"Cabine telephonique","431":"Batiment de telecommunication","432":"Aeroport","433":"Batiment funeraire","434":"Batiment administratif","440":"Orphelinat","441":"Creche","442":"Atelier protege","443":"Maison de repos","444":"Batiment hospitalier","445":"Etablissement de cure","446":"Batiment d’aide sociale","460":"Batiment scolaire","461":"Universite","462":"Musee","463":"Bibliotheque","480":"Eglise","481":"Chapelle","482":"Couvent","483":"Presbytere","484":"Seminaire","485":"Eveche","486":"Synagogue","487":"Mosquee","488":"Temple","489":"Batiment de culte","500":"Etablissement de bains","501":"Installation sportive","502":"Home vacances","503":"Habitation de vacances","504":"Maison de jeunes","505":"Theatre","506":"Salle de spectacles","507":"Centre culturel","508":"Cinema","509":"Casino","510":"Point de vue","520":"Ruines","521":"Souterrain","522":"Pavillon","523":"Chateau","524":"Batiment historique","525":"Monument","526":"Moulin a vent","527":"Moulin a eau","528":"Chateau d’eau","529":"Captage d’eau","530":"Installation d’epuration","531":"Traitement des immondices","532":"Cave #","533":"Chambre #","534":"Studio #","535":"Commerce #","536":"Bureau #","537":"Appartement #","538":"Parties communes generales non baties","539":"Autre bati","540":"Autre non bati","541":"Droit sup./emph.","542":"Entite d’exploitation #","543":"Garage box #","544":"Parking couvert #","545":"Parking non-couvert #","546":"Reserve #","547":"Vitrine #","549":"Entite privative diverse #","550":"Cabine #","551":"Partie commune specifique non-batie","552":"Superficie et parties communes","553":"Parties communes","554":"Panneau solaire","555":"Superficie batie","556":"Volume entier","557":"Volume limite","558":"Volume semi-limite","559":"Domaine public"};

$("#btnGo").addEventListener("click", fuse);
$("#btnRefresh").addEventListener("click", refreshApp);
$("#q").addEventListener("input", () => render(base));
$("#btnExcel").addEventListener("click", exportExcel);
$("#btnPDF").addEventListener("click", exportPDF);
btnParcelles?.addEventListener("click", exportParcellesCSV);
$("#btnAddRow").addEventListener("click", addManualRow);
$("#btnToggleCompact").addEventListener("click", () => setCompactMode(!compactMode));
$("#selectAll").addEventListener("click", toggleSelectAll);
document.addEventListener('paste', (e) => {
    if (!(tableWrap.contains(document.activeElement) || document.activeElement === document.body)) return;
    const text = e.clipboardData?.getData('text');
    if (text && text.includes('\t')) {
        e.preventDefault();
        addRowsFromClipboard(text);
    }
});
window.addEventListener('load', () => {
    setupDropZones();
    initSortable();
    setupScrollSync();
    window.addEventListener('resize', updateTopScrollbar);
    setCompactMode(window.innerWidth < 1200);
});

function refreshApp() {
    partsRaw = []; parcRaw = []; merged = []; base = [];
    uploadedFiles = { A: [], B: [] }; // Réinitialiser le suivi des fichiers
    ['A', 'B'].forEach(tag => {
        const dz = $(`#dropZone${tag}`);
        dz.querySelector('.drop-zone__prompt').style.display = 'block';
        const fileInfo = dz.querySelector('.drop-zone__file-info');
        fileInfo.style.display = 'none';
        fileInfo.innerHTML = ''; // Vider la liste des fichiers
        dz.querySelector('.drop-zone__input').value = '';
    });
    $("#q").value = ""; $("#selectAll").checked = false;
    tblBody.innerHTML = "";
    $("#counters").textContent = "";
    hideStatus();
    $("#btnGo").disabled = true; $("#btnExcel").disabled = true; $("#btnPDF").disabled = true; if (btnParcelles) btnParcelles.disabled = true;
    updateTopScrollbar();
    localStorage.removeItem('cadastralDataBackup');
}
function setupDropZones() {
    document.querySelectorAll('.drop-zone').forEach(dz => {
        const input = dz.querySelector('.drop-zone__input');
        const tag = dz.id.endsWith('A') ? 'A' : 'B';
        
        dz.addEventListener('click', () => input.click());
        input.addEventListener('change', () => { if(input.files.length) handleFiles(input.files, tag) });

        dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('dragover'); });
        dz.addEventListener('dragleave', () => dz.classList.remove('dragover'));
        dz.addEventListener('drop', e => {
            e.preventDefault();
            dz.classList.remove('dragover');
            if (e.dataTransfer.files.length) {
                input.files = e.dataTransfer.files;
                handleFiles(e.dataTransfer.files, tag);
            }
        });
    });
}
function saveState() {
    if (base && base.length > 0) {
        try { localStorage.setItem('cadastralDataBackup', JSON.stringify(base)); } 
        catch (err) { console.error("Erreur de sauvegarde:", err); }
    } else { localStorage.removeItem('cadastralDataBackup'); }
}

// MODIFIÉ : Gère une liste de fichiers au lieu d'un seul
async function handleFiles(files, tag) {
    if (!files || files.length === 0) return;
    const dz = $(`#dropZone${tag}`);
    const prompt = dz.querySelector('.drop-zone__prompt');
    const fileInfo = dz.querySelector('.drop-zone__file-info');

    let totalRowsAdded = 0;
    setStatus(`Lecture de ${files.length} fichier(s)...`);

    for (const file of files) {
        // Éviter les doublons
        if (uploadedFiles[tag].includes(file.name)) continue;

        const rows = await readTolerant(file);
        if (!rows.length) continue;

        const kind = guessKind(rows);
        const targetTag = (kind === "parcelles") ? 'B' : 'A';
        
        // Concaténer les nouvelles données
        if (targetTag === 'A') {
            partsRaw = partsRaw.concat(rows);
        } else {
            parcRaw = parcRaw.concat(rows);
        }
        totalRowsAdded += rows.length;
        uploadedFiles[targetTag].push(file.name);
    }
    
    // Mettre à jour l'UI avec la liste complète des fichiers
    const currentList = uploadedFiles[tag];
    const totalLines = (tag === 'A') ? partsRaw.length : parcRaw.length;

    const summaryHtml = `<span>✅ ${currentList.length} fichier(s) chargé(s) (${totalLines} lignes au total)</span>`;
    const fileListHtml = '<ul>' + currentList.map(name => `<li>${esc(name)}</li>`).join('') + '</ul>';

    fileInfo.innerHTML = summaryHtml + fileListHtml;
    prompt.style.display = 'none';
    fileInfo.style.display = 'block';

    hideStatus();
    if (totalRowsAdded > 0) {
      setStatus(`${totalRowsAdded} nouvelles lignes ajoutées.`, 'success', 3000);
    }
    checkReady();
}

function hasIdf(rows) { return Array.isArray(rows) && rows.length > 0 && rows.some(r => String(r.propertySituationIdf || "").trim() !== ""); }
function checkReady() {
    const ok = hasIdf(partsRaw) && hasIdf(parcRaw);
    $("#btnGo").disabled = !ok;
    if (btnParcelles) btnParcelles.disabled = parcRaw.length === 0;
    if (ok) setStatus("Fichiers prêts à être fusionnés.", 'success', 3000);
}
async function readTolerant(file){const ext=(file.name.split('.').pop()||'').toLowerCase();if(ext==="csv")return parseCSV(await readText(file));const buf=await file.arrayBuffer();const wb=XLSX.read(buf,{type:"array"});let all=[];for(const name of wb.SheetNames){const ws=wb.Sheets[name];all=all.concat(sheetToObjects(ws));}return all;}
function parseCSV(text){text=text.replace(/\r/g,'');const headLine=text.split('\n')[0]||"";const sep=headLine.includes('\t')?'\t':(headLine.includes(';')?';':',');const lines=text.split('\n').filter(l=>l.trim().length);const rows=lines.map(l=>splitCSV(l,sep));return rowsToObjects(rows);}
function splitCSV(line,sep){const out=[];let cur="";let inQ=false;for(let i=0;i<line.length;i++){const c=line[i];if(c==='"'){inQ=!inQ;continue;}if(c===sep&&!inQ){out.push(cur);cur="";}else cur+=c;}out.push(cur);return out.map(s=>s.trim());}
function sheetToObjects(ws){const arr=XLSX.utils.sheet_to_json(ws,{header:1,defval:"",raw:false});if(!arr.length)return[];return rowsToObjects(arr);}
function rowsToObjects(arr){const idx=findHeaderRow(arr);const headers=normalizeHeaders(arr[idx]||arr[0]);const data=arr.slice(idx+1);return data.map(row=>{const o={};headers.forEach((h,i)=>o[h]=String(row[i]??"").trim());return o;});}
function findHeaderRow(arr){const target="propertysituationidf";let best=0,found=0;for(let i=0;i<arr.length;i++){const row=arr[i];const score=row.reduce((s,c)=>s+(String(c).toLowerCase().replace(/\s+/g,'')===target?5:(String(c).toLowerCase().includes("property")?1:0)),0);if(score>best){best=score;found=i;}if(row.some(c=>String(c).toLowerCase().replace(/\s+/g,'')===target))return i;}return found;}
function normalizeHeaders(hs){const seen=new Set();return hs.map((h,i)=>{let k=String(h||`col${i+1}`).trim();const kl=k.toLowerCase().replace(/\s+/g,'').replace(/[^\w]/g,'');const map={propertysituationidf:"propertySituationIdf",divcad:"divCad",primarynumber:"primaryNumber",bisnumber:"bisNumber",exponentletter:"exponentLetter",exponentnumber:"exponentNumber",partnumber:"partNumber",surfaceverif:"surfaceVerif",surfacenottaxable:"surfaceNotTaxable",surfacetaxable:"surfaceTaxable",zipcode:"zipCode",municipality:"municipality",firstname:"firstname",name:"name",right:"right",street:"street",number:"number",boxnumber:"boxNumber",country:"country",pays:"country",officialid:"officialId",partytype:"partyType"};if(map[kl])k=map[kl];let base=k,n=1;while(seen.has(k)){k=base+"_"+(++n);}seen.add(k);return k;});}
function guessKind(rows){const cols=new Set(Object.keys(rows[0]||{}).map(c=>c.toLowerCase()));const parcHints=["capakey","section","divcad","nature","surfacetaxable","primarynumber"];const partHints=["name","firstname","right","street","municipality","zipcode","boxnumber","managedby","officialid","partytype","country","pays"];const parcScore=parcHints.filter(h=>cols.has(h)).length;const partScore=partHints.filter(h=>cols.has(h)).length;return parcScore>=partScore?"parcelles":"parties";}
const RIGHT_PRIORITY={'PP':1,'US':2,'NP':3,'GEST':9};
function rightInfo(raw){const s=String(raw||"").trim();if(s==='')return{abbr:'GEST',label:'Gestion',rank:RIGHT_PRIORITY.GEST};const t0=stripAccents(s).toLowerCase();if(/\busufru/i.test(t0))return{abbr:'US',label:'Usufruitier',rank:RIGHT_PRIORITY.US};if(/\bnue?\s*propri/i.test(t0))return{abbr:'NP',label:'Nue-propriete',rank:RIGHT_PRIORITY.NP};if(/\bpleine\s*propri/i.test(t0)||/\bpp\b/.test(t0))return{abbr:'PP',label:'Pleine propriete',rank:RIGHT_PRIORITY.PP};if(/emphyteo/i.test(t0))return{abbr:'EMPH',label:'Emphyteose',rank:4};if(/superficie/i.test(t0))return{abbr:'SUP',label:'Droit de superficie',rank:5};if(/indivis/i.test(t0))return{abbr:'IND',label:'Indivision',rank:6};return{abbr:s.toUpperCase().slice(0,6)||'AUTRE',label:stripAccents(s),rank:7};}
function extractShare(obj){const num=firstNumber(obj.partsNumerator,obj.numerator,obj.num,obj.nbParts,obj.partsNum);const den=firstNumber(obj.partsDenominator,obj.denominator,obj.den,obj.totalParts,obj.partsDen);if(num&&den)return`${num}/${den}`;const pct=firstText(obj.percent,obj.sharePct,obj.pourcentage);if(pct)return pct.toString().replace('%','').trim()+' %';const frac=firstText(obj.share,obj.quote,obj.fraction,obj.parts);if(frac&&/[\/]/.test(frac))return frac;const r=String(obj.right||"");const m1=r.match(/(\d+)\s*\/\s*(\d+)/);if(m1)return`${m1[1]}/${m1[2]}`;const m2=r.match(/(\d+)\s*%/);if(m2)return`${m2[1]} %`;return"";}
function firstNumber(...vals){for(const v of vals){const n=Number(String(v??"").replace(',','.'));if(Number.isFinite(n)&&n>0)return n;}return null;}
function firstText(...vals){for(const v of vals){if(v!=null&&String(v).trim()!=="")return String(v).trim();}return"";}
function buildAddr(p){const parts=[p.street,p.number,p.boxNumber?("bte "+p.boxNumber):"",[p.zipCode,p.municipility||p.municipality].filter(Boolean).join(" ")].filter(Boolean);return titleCaseSmart(parts.join(", ").replace(/\s+,/g,",").replace(/,\s*,/g,", "));}
function formatAddrWithCountry(addr,country){const c=String(country||'').trim().toUpperCase();const a=(addr||'').trim();if(!c||c==='BE')return a||'—';return a?`${a} (${c})`:`Adresse non renseignee (${c})`;}
function ownersMap(rows){const map=new Map();for(const r of rows){const idf=(r.propertySituationIdf||"").trim();if(!idf)continue;const first=titleCaseSmart(r.firstname||"");const last=titleCaseSmart(r.name||"");const name=(first||last)?`${first} ${last}`.trim():(r.officialId||"(inconnu)");const ri=rightInfo(r.right);const share=extractShare(r);const item={name,first,last,officialId:String(r.officialId||"").trim(),partyType:titleCaseSmart(r.partyType||""),right:ri.label,rightAbbr:ri.abbr,rightRank:ri.rank,share,street:titleCaseSmart(r.street||""),number:String(r.number||"").trim(),zipCode:String(r.zipCode||"").trim(),municipality:titleCaseSmart(r.municipility||r.municipality||""),country:String(r.country||r.Country||r.pays||r.Pays||"").trim().toUpperCase(),addr:buildAddr(r)};const cur=map.get(idf)||{list:[],date:""};cur.list.push(item);cur.date=r.dateSituation||cur.date;map.set(idf,cur);}for(const v of map.values()){v.list.sort((a,b)=>(a.rightRank-b.rightRank)||a.name.localeCompare(b.name,'fr'));}return map;}
function natureLabel(code){if(code==null)return"";const raw=String(code).trim();const numeric=Number(raw);if(!Number.isNaN(numeric)&&mapping[String(numeric)])return mapping[String(numeric)];if(mapping[raw])return mapping[raw];const up=raw.toUpperCase();if(mapping[up])return mapping[up];return"";}

function recalcRow(row, changedKey) {
    const surfTax = toNumber(row.surfaceTaxable);
    const surfNotTax = toNumber(row.surfaceNotTaxable);
    const areaCandidate = surfTax + surfNotTax;
    if (areaCandidate > 0) row.areaM2 = areaCandidate;
    const empPPRaw = row.empPP_m2;
    const empPP = toNumber(empPPRaw);
    const hasEmpPP = empPPRaw !== '' && empPPRaw != null && !Number.isNaN(empPP);
    if (changedKey !== 'excedent_m2') {
        const hasArea = row.areaM2 !== '' && row.areaM2 != null && Number.isFinite(toNumber(row.areaM2));
        row.excedent_m2 = hasArea && hasEmpPP ? Math.max(0, toNumber(row.areaM2) - empPP) : '';
    }
}

function applyEdit(row, key, value) {
    const numericFields = new Set(['surfaceTaxable', 'surfaceNotTaxable', 'areaM2', 'empPP_m2', 'empSS_m2', 'excedent_m2', 'servitudePrincipale', 'empPPJudiciaire']);
    const cleanValue = (value ?? '').toString().trim();
    row[key] = numericFields.has(key) ? normalizeNumberString(cleanValue) : cleanValue;
    if (key === 'nature') {
        const lbl = natureLabel(cleanValue);
        if (lbl) row.natureLabel = lbl;
    }
    if (!row.propertySituationIdf) {
        row.propertySituationIdf = `MAN-${manualRowCount}`;
    }
    recalcRow(row, key);
    saveState();
    render(base);
}

  function editableCell(row, key, opts = {}) {
      const td = document.createElement('td');
      td.classList.add('editable-cell');
      if (opts.small) td.classList.add('small');
      if (opts.advanced) td.classList.add('advanced-col');
      if (opts.align === 'right') td.style.textAlign = 'right';
      const box = document.createElement('div');
    box.className = 'edit-box';
    box.contentEditable = true;
    box.dataset.key = key;
    const rawValue = row[key] ?? '';
    box.textContent = opts.formatSurface ? formatSurfaceDisplay(rawValue) : rawValue;
    box.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); box.blur(); } });
    box.addEventListener('focus', () => { if (opts.formatSurface) { box.textContent = row[key] ?? ''; } });
    box.addEventListener('blur', () => applyEdit(row, key, box.textContent));
    td.appendChild(box);
    return td;
}

function fuse(){
    if (!partsRaw.length || !parcRaw.length) { setStatus("Il manque un fichier.", 'error', 4000); return; }
    const omap = ownersMap(partsRaw);
    const out = []; let skipped = 0;
    for(const p of parcRaw){
        const idf=(p.propertySituationIdf||"").trim();
        if(!idf){skipped++;continue;}
        const ow=omap.get(idf)||{list:[],date:""};
        const numero=[p.primaryNumber,p.bisNumber?("bis "+p.bisNumber):"",(p.exponentLetter||"")+(p.exponentNumber||"")].filter(Boolean).join(" ").trim();
        const surfM2=toNumber(p.surfaceVerif)||sumNumbers(p.surfaceTaxable,p.surfaceNotTaxable)||toNumber(p.surfaceTotal)||toNumber(p.parcelleSurface)||0;
        out.push({_selected:false,propertySituationIdf:idf,ownersList:ow.list,countriesNonBE:[...new Set((ow.list||[]).map(o=>String(o.country||'').toUpperCase()).filter(c=>c&&c!=='BE'))].join(', ')||'—',capakey:stripAccents(p.capakey||""),nature:stripAccents(p.nature||""),natureLabel:stripAccents(natureLabel(p.nature)),surfaceTaxable:p.surfaceTaxable||"",surfaceNotTaxable:p.surfaceNotTaxable||"",areaM2:surfM2,empPP_m2:'',empSS_m2:'',excedent_m2:'',servitudePrincipale:'',zoneLocation:'',empPPJudiciaire:'',divCad:stripAccents(p.divCad||""),section:stripAccents(p.section||""),number:stripAccents(numero),partNumber:stripAccents(p.partNumber||""),dateSituation:stripAccents(ow.date||p.dateSituation||"")});
    }
    merged=out; base=out.slice(0); render(base); saveState();
    $("#btnExcel").disabled = false; $("#btnPDF").disabled = false;
    setStatus(`Fusion terminée: ${merged.length} lignes`, 'success', 4000);
    $("#counters").textContent = `Parties: ${partsRaw.length} · Parcelles: ${parcRaw.length} · Fusion: ${merged.length}`;
}

function render(rows) {
    const q = $("#q").value.trim().toLowerCase();
    const filtered = q ? rows.filter(r => (Object.values(r).join(' ')+' '+r.ownersList.map(o=>Object.values(o).join(' ')).join(' ')).toLowerCase().includes(q)) : rows;
    tblBody.innerHTML = "";
    filtered.forEach(r => {
        const tr=document.createElement("tr");
        tr.className="draggable";
        tr.dataset.id=r.propertySituationIdf;

        const ownersHtml=r.ownersList.length>0?r.ownersList.map(owner=>`<div><b>${esc(owner.name)}</b><div class="owner-info"><span class="tag">${esc(owner.right)}</span><br><small>ID: ${esc(owner.officialId||'—')} | Type: ${esc(owner.partyType||'—')}</small></div></div>`).join('<hr class="owner-separator">'):'<div><b>—</b><div class="owner-info"><span class="tag">Gestion</span></div></div>';
        const sharesHtml=r.ownersList.length>0?r.ownersList.map(o=>`<div>${esc(o.share)||'—'}</div>`).join('<hr class="owner-separator">'):'—';
        const addressesHtml=r.ownersList.length>0?r.ownersList.map(o=>`<div>${esc(o.addr)||'—'}</div>`).join('<hr class="owner-separator">'):'—';
        const countriesHtml=r.ownersList.length>0?r.ownersList.map(owner=>{const c=String(owner.country||'').toUpperCase();return`<div>${c&&c!=='BE'?esc(c):'—'}</div>`;}).join('<hr class="owner-separator">'):'—';

        const orderTd=document.createElement('td');
        orderTd.textContent=base.indexOf(r)+1;
        const selectTd=document.createElement('td');
        selectTd.className='select-col';
        const checkbox=document.createElement('input');
        checkbox.type='checkbox';
        checkbox.dataset.idf=r.propertySituationIdf;
        checkbox.checked=!!r._selected;
        checkbox.addEventListener('change',e=>{r._selected=e.target.checked;$("#selectAll").checked=base.length>0&&base.every(item=>item._selected);saveState();});
        selectTd.appendChild(checkbox);

        tr.appendChild(orderTd);
        tr.appendChild(selectTd);
        tr.appendChild(editableCell(r,'propertySituationIdf'));

        const ownerTd=document.createElement('td'); ownerTd.innerHTML=ownersHtml; tr.appendChild(ownerTd);
        const shareTd=document.createElement('td'); shareTd.innerHTML=sharesHtml; tr.appendChild(shareTd);
        const addrTd=document.createElement('td'); addrTd.innerHTML=addressesHtml; tr.appendChild(addrTd);
        const countryTd=document.createElement('td'); countryTd.innerHTML=countriesHtml; tr.appendChild(countryTd);

        tr.appendChild(editableCell(r,'capakey'));
        tr.appendChild(editableCell(r,'nature'));
        tr.appendChild(editableCell(r,'natureLabel'));
        tr.appendChild(editableCell(r,'surfaceTaxable',{small:true,align:'right',formatSurface:true}));
        const empPPTd = editableCell(r,'empPP_m2',{small:true,align:'right',advanced:true,formatSurface:true});
        empPPTd.classList.add('compact-col');
        tr.appendChild(empPPTd);

        const excedentTd=document.createElement('td');
        excedentTd.className='compact-col advanced-col';
        excedentTd.innerHTML=`<span class="computed-value">${fmtNum(r.excedent_m2)}</span>`;
        tr.appendChild(excedentTd);

        const servitudeTd=editableCell(r,'servitudePrincipale',{small:true,align:'right',advanced:true,formatSurface:true});
        servitudeTd.classList.add('compact-col');
        tr.appendChild(servitudeTd);

        const zoneLocTd=editableCell(r,'zoneLocation',{advanced:true});
        zoneLocTd.style.minWidth='120px';
        tr.appendChild(zoneLocTd);

        const empJudTd=editableCell(r,'empPPJudiciaire',{small:true,align:'right',advanced:true,formatSurface:true});
        empJudTd.classList.add('compact-col');
        tr.appendChild(empJudTd);

        tblBody.appendChild(tr);
    });
    $("#selectAll").checked = base.length > 0 && base.every(item => item._selected);
    updateTopScrollbar();
}

function initSortable() {
    new Sortable(tblBody, {
        animation: 150,
        ghostClass: 'sortable-ghost',
        chosenClass: 'sortable-chosen',
        fallbackClass: 'sortable-fallback',
        forceFallback: true,
        scroll: true,
        scrollSensitivity: 100,
        scrollSpeed: 15,
        onEnd: function (evt) {
            if (evt.oldIndex === evt.newIndex) return;
            const [movedItem] = base.splice(evt.oldIndex, 1);
            base.splice(evt.newIndex, 0, movedItem);
            render(base);
            saveState();
        },
    });
}

function toggleSelectAll(e){const isChecked=e.target.checked;base.forEach(r=>r._selected=isChecked);render(base);saveState();}
function makeEmptyRow() {
    return {
        _selected: false,
        propertySituationIdf: `MAN-${manualRowCount++}`,
        ownersList: [],
        countriesNonBE: '—',
        capakey: '',
        nature: '',
        natureLabel: '',
        surfaceTaxable: '',
        surfaceNotTaxable: '',
        areaM2: 0,
        empPP_m2: 0,
        empSS_m2: 0,
        excedent_m2: 0,
        servitudePrincipale: '',
        zoneLocation: '',
        empPPJudiciaire: '',
        divCad: '',
        section: '',
        number: '',
        partNumber: '',
        dateSituation: ''
    };
}

function addManualRow(showStatus = true) {
    const newRow = makeEmptyRow();
    base.push(newRow);
    render(base);
    saveState();
    $("#btnExcel").disabled = false; $("#btnPDF").disabled = false;
    if (showStatus) setStatus('Nouvelle ligne manuelle ajoutée.', 'success', 2000);
    const lastEditable = tblBody.querySelector('tr:last-child .editable-cell .edit-box');
    if (lastEditable) lastEditable.focus();
    return newRow;
}

const PASTE_ORDER = ['propertySituationIdf','capakey','nature','natureLabel','surfaceTaxable','empPP_m2','servitudePrincipale','zoneLocation','empPPJudiciaire'];

function addRowsFromClipboard(text) {
    const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    if (!lines.length) return;
    let added = 0;
    for (const line of lines) {
        const cols = line.split('\t');
        if (!cols.length) continue;
        const row = makeEmptyRow();
        PASTE_ORDER.forEach((key, idx) => {
            const val = cols[idx];
            if (val != null && val !== '') row[key] = val.trim();
        });
        recalcRow(row);
        base.push(row);
        added++;
    }
    if (added > 0) {
        render(base);
        saveState();
        $("#btnExcel").disabled = false; $("#btnPDF").disabled = false;
        setStatus(`${added} ligne(s) collée(s) depuis Excel.`, 'success', 2500);
    }
}

function exportParcellesCSV() {
    if (!parcRaw.length) { setStatus("Aucune donnée parcelle à exporter.", 'error', 3000); return; }
    const headers = Array.from(new Set(parcRaw.flatMap(r => Object.keys(r))));
    const escapeVal = (val) => `"${String(val ?? '').replace(/"/g, '""')}"`;
    const lines = [headers.join(';')];
    parcRaw.forEach(row => {
        lines.push(headers.map(h => escapeVal(row[h])).join(';'));
    });
    const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
    const stamp = new Date().toISOString().slice(0,10).replace(/-/g,'');
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `parcelles_${stamp}.csv`;
    a.click();
    URL.revokeObjectURL(a.href);
    setStatus('Fichier des parcelles créé.', 'success', 3000);
}

// =========================================================================
// =================== FONCTION EXPORTEXCEL MODIFIÉE =======================
// =========================================================================
async function exportExcel() {
    const selectedRows = base.filter(r => r._selected);
    const rowsToExport = selectedRows.length > 0 ? selectedRows : base;
    if (!rowsToExport.length) { setStatus("Rien à exporter", 'error', 3000); return; }
    try {
        const grey='FFEAEAEA',light='FFF9F9F9',wb=new ExcelJS.Workbook(),ws=wb.addWorksheet('TABLEAU DES EMPRISES',{properties:{defaultRowHeight:20}});
        ws.columns=[{width:8},{width:10},{width:9},{width:14},{width:20},{width:5},{width:5},{width:5},{width:22},{width:16},{width:10},{width:18},{width:24},{width:10},{width:5},{width:5},{width:5},{width:5},{width:5},{width:8},{width:14},{width:16},{width:16}];
        const C=(r,c)=>ws.getCell(r,c),center={vertical:'middle',horizontal:'center',wrapText:true},left={vertical:'middle',horizontal:'left',wrapText:true},borderThin={top:{style:'thin'},left:{style:'thin'},right:{style:'thin'},bottom:{style:'thin'}};
        let r=1;
        const topFill={type:'pattern',pattern:'solid',fgColor:{argb:grey}};
        const fillRow=(rowIdx)=>{for(let col=1;col<=23;col++){const cell=C(rowIdx,col);cell.fill=topFill;cell.border=borderThin;}};
        ws.mergeCells(r,1,r,23);C(r,1).value="TABLEAU DES EMPRISES";C(r,1).alignment=center;C(r,1).font={bold:true,size:16};fillRow(r);r++;
        ws.mergeCells(r,1,r,23);C(r,1).value="D'après mesurage";C(r,1).alignment=center;C(r,1).font={italic:true,size:11,color:{argb:'FF444444'}};fillRow(r);r++;
        ws.mergeCells(r,1,r,23);C(r,1).value=`POUVOIR EXPROPRIANT : ${GEO_HEADER.pouvoir}`;C(r,1).alignment=center;C(r,1).font={bold:true,size:12};fillRow(r);r++;
        ws.mergeCells(r,1,r,6);C(r,1).value=`D. SPGE : ${GEO_HEADER.spgeRef}`;C(r,1).alignment=left;C(r,1).font={bold:true};
        ws.mergeCells(r,7,r,14);C(r,7).value=GEO_HEADER.operation;C(r,7).alignment=center;C(r,7).font={bold:true};
        ws.mergeCells(r,15,r,23);C(r,15).value=`D. IGRETEC : ${GEO_HEADER.igretecRef}`;C(r,15).alignment=left;C(r,1).font={bold:true};r+=2;
        const headTop=r,headMid=r+1,headBot=r+2;
        ws.mergeCells(headTop,1,headBot,1);C(headTop,1).value="N° d'ordre";ws.mergeCells(headTop,2,headTop,8);C(headTop,2).value="INFORMATION CADASTRALE";ws.mergeCells(headMid,2,headBot,2);C(headMid,2).value="Division";ws.mergeCells(headMid,3,headBot,3);C(headMid,3).value="Section";ws.mergeCells(headMid,4,headBot,4);C(headMid,4).value="Parcelle";ws.mergeCells(headMid,5,headBot,5);C(headMid,5).value="Nature";ws.mergeCells(headMid,6,headMid,8);C(headMid,6).value="Contenance";C(headBot,6).value="Ha";C(headBot,7).value="A";C(headBot,8).value="Ca";
        ws.mergeCells(headTop,9,headTop,14);C(headTop,9).value="Coordonnées des propriétaires";["Nom","Prénom","Code postal","Commune","Rue","Numéro"].forEach((t,i)=>{ws.mergeCells(headMid,9+i,headBot,9+i);C(headMid,9+i).value=t;});
        ws.mergeCells(headTop,15,headMid,17);C(headTop,15).value="Emprise en pleine propriété";C(headBot,15).value="Ha";C(headBot,16).value="A";C(headBot,17).value="Ca";
        ws.mergeCells(headTop,18,headMid,20);C(headTop,18).value="Excédent d'emprise";C(headBot,18).value="Ha";C(headBot,19).value="A";C(headBot,20).value="Ca";
        ws.mergeCells(headTop,21,headBot,21);C(headTop,21).value="Servitude principale (m²)";
        ws.mergeCells(headTop,22,headBot,22);C(headTop,22).value="Zone de location";
        ws.mergeCells(headTop,23,headBot,23);C(headTop,23).value="Emprise PP (procédure judiciaire)";
        for(let rowNum=headTop;rowNum<=headBot;rowNum++){for(let colNum=1;colNum<=23;colNum++){const cell=ws.getCell(rowNum,colNum);cell.font={bold:true};cell.alignment=center;cell.border=borderThin;if(rowNum===headTop)cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:grey}};else cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:light}};}}
        r=headBot+1;
        rowsToExport.forEach((it,index)=>{const{ha,a,ca}=m2toHaACa(it.areaM2);const owners=(it.ownersList&&it.ownersList.length)?it.ownersList:[{last:"",first:"",zipCode:"",municipality:"",street:"",number:""}];
        const ownerLasts=owners.map(o=>(o.last||"").toString());
        const ownerFirsts=owners.map(o=>(o.first||"").toString());
        const allFirstEmpty=ownerFirsts.every(v=>String(v).trim()==="");

        const hasEmpPP = it.empPP_m2 !== '' && it.empPP_m2 != null && Number.isFinite(toNumber(it.empPP_m2));
        const { ha: empHa, a: empA, ca: empCa } = hasEmpPP ? m2toHaACa(it.empPP_m2) : { ha: '', a: '', ca: '' };

        const rowData=[index+1,it.divCad||"",it.section||"",it.number||"",it.natureLabel||it.nature||"",ha,a,ca,ownerLasts.join('\n'),ownerFirsts.join('\n'),owners.map(o=>(o.zipCode||"").toString()).join('\n'),owners.map(o=>o.municipality||"").join('\n'),owners.map(o=>o.street||"").join('\n'),owners.map(o=>(o.number||"").toString()).join('\n'),empHa,empA,empCa,'', '', '',it.servitudePrincipale||"",it.zoneLocation||"",it.empPPJudiciaire||""];
        const addedRow=ws.addRow(rowData);
        const rowNumber = addedRow.number;
        if(allFirstEmpty){
            ws.mergeCells(rowNumber,9,rowNumber,10);
            const mergedCell=ws.getCell(rowNumber,9);
            mergedCell.value=ownerLasts.join('\n');
            mergedCell.alignment=left;
            mergedCell.border=borderThin;
            ws.getCell(rowNumber,10).border=borderThin;
        }

        addedRow.eachCell((cell, colNumber)=>{
            cell.border=borderThin;
            cell.alignment={vertical:'middle',horizontal:(colNumber>=9&&colNumber<=14)?'left':'center',wrapText:true};

            const surfaceColumns = [6, 7, 8, 15, 16, 17, 18, 19, 20, 21, 23];
            if (surfaceColumns.includes(colNumber)) {
                cell.numFmt = '00';
            }
        });
        const totalM2 = `(F${rowNumber}*10000)+(G${rowNumber}*100)+H${rowNumber}`;
        const empriseM2 = `(O${rowNumber}*10000)+(P${rowNumber}*100)+Q${rowNumber}`;
        const diffM2 = `MAX(0,${totalM2}-${empriseM2})`;
        const hasInputs = `AND(F${rowNumber}<>"",G${rowNumber}<>"",H${rowNumber}<>"",O${rowNumber}<>"",P${rowNumber}<>"",Q${rowNumber}<>"")`;
        ws.getCell(`R${rowNumber}`).value = { formula: `IF(${hasInputs},INT(${diffM2}/10000),"")` };
        ws.getCell(`S${rowNumber}`).value = { formula: `IF(${hasInputs},INT(MOD(${diffM2},10000)/100),"")` };
        ws.getCell(`T${rowNumber}`).value = { formula: `IF(${hasInputs},MOD(${diffM2},100),"")` };
        r++;});
        ws.mergeCells(r,1,r,23);C(r,1).value="Excédent d'emprise = contenance - emprise en pleine propriété";C(r,1).alignment={...left,wrapText:false};C(r,1).font={italic:true,color:{argb:'FF555555'}};r+=2;
        ws.mergeCells(r,1,r,23);C(r,1).value=GEO_HEADER.operation;C(r,1).alignment=center;C(r,1).font={bold:true};
        const buffer=await wb.xlsx.writeBuffer();const blob=new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheet.sheet'});const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='tableau_des_emprises.xlsx';document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(a.href);
        setStatus('Export Excel terminé avec succès.', 'success', 4000);
    } catch (err) { console.error(err); setStatus("Erreur Excel: "+err.message, 'error'); }
}

const PDF_CONFIG={FRAME_PADDING_Y:30,MARGIN_BETWEEN_CARDS:10,SECTION_SPACING:1,OWNER_SEPARATOR_SPACING:10,TITLE_BADGE_HEIGHT:22,SPACE_AFTER_TITLE:12};
function pdfUseFont(doc){if(doc.setCharSpace)doc.setCharSpace(0);doc.setFont('helvetica','normal');}
function pdfText(doc,s,maxWidth){const t=toAscii(s);return maxWidth?doc.splitTextToSize(t,maxWidth):t;}
function calculateContentHeight(doc,r){const C=PDF_CONFIG;let height=0;const valueMaxWidth=doc.internal.pageSize.getWidth()-95-170;doc.setFont('helvetica','normal');const lineHeight9=doc.setFontSize(9).getLineHeight()*1.1;height+=C.TITLE_BADGE_HEIGHT+C.SPACE_AFTER_TITLE;height+=3*lineHeight9;height+=C.SECTION_SPACING;height+=C.TITLE_BADGE_HEIGHT+C.SPACE_AFTER_TITLE;if(r.countriesNonBE&&r.countriesNonBE!=='—'){height+=doc.setFontSize(8).getTextDimensions(pdfText(doc,`Pays (≠ BE) : ${r.countriesNonBE}`)).h+2;}
const list=(r.ownersList?.length?r.ownersList:[{name:"—",officialId:"",partyType:"",right:"Gestion",share:"",addr:"",country:""}]);list.forEach((owner,index)=>{const ownerData={"Propriétaire":titleCaseSmart(owner.name||'—'),"ID National":stripAccents(owner.officialId||'—'),"Type":stripAccents(owner.partyType||'—'),"Droit":stripAccents(owner.right||'—'),"Quote-part":stripAccents(owner.share||'—'),"Adresse":formatAddrWithCountry(titleCaseSmart(owner.addr||''),owner.country)};doc.setFontSize(9);for(const value of Object.values(ownerData)){const dims=doc.getTextDimensions(toAscii(value),{maxWidth:valueMaxWidth,fontSize:9});height+=dims.h+2;}
if(index<list.length-1){height+=C.OWNER_SEPARATOR_SPACING;}});return height;}
async function exportPDF(){const selectedRows=base.filter(r=>r._selected);const rows=selectedRows.length>0?selectedRows:base;if(!rows.length)return;const{jsPDF}=window.jspdf;const doc=new jsPDF({unit:'pt',format:'a4'});pdfUseFont(doc);const generationDate=new Date().toLocaleDateString('fr-BE',{day:'2-digit',month:'2-digit',year:'numeric'});const C=PDF_CONFIG;const pageHeight=doc.internal.pageSize.getHeight();const topMargin=40;const bottomMargin=80;const pageBreakThreshold=pageHeight-bottomMargin;let currentY=topMargin;for(let i=0;i<rows.length;i++){const r=rows[i];const contentHeight=calculateContentHeight(doc,r);const totalBlockHeight=contentHeight+C.FRAME_PADDING_Y;if(currentY+totalBlockHeight>pageBreakThreshold&&i>0){footer(doc,generationDate);doc.addPage();pdfUseFont(doc);currentY=topMargin;}
const frameStartY=currentY;doc.setLineWidth(1.2);doc.setDrawColor(255,122,0);doc.setFillColor(255,252,248);doc.roundedRect(40,frameStartY,doc.internal.pageSize.getWidth()-80,totalBlockHeight,8,8,'FD');doc.setLineWidth(0.2);doc.setDrawColor(0);currentY+=C.FRAME_PADDING_Y/2;const parcelInfoEndY=sectionParcelInfo(doc,r,currentY);sectionOwners(doc,r,parcelInfoEndY+C.SECTION_SPACING);currentY=frameStartY+totalBlockHeight+C.MARGIN_BETWEEN_CARDS;}
footer(doc,generationDate);doc.save('fiches_condensees.pdf');setStatus('Export PDF terminé avec succès.', 'success', 4000);}
function sectionTitle(doc,text,y){const W=doc.internal.pageSize.getWidth();const boxW=W-110;const h=PDF_CONFIG.TITLE_BADGE_HEIGHT;doc.setDrawColor(255,217,191);doc.setFillColor(255,244,234);doc.roundedRect(55,y,boxW,h,6,6,'FD');doc.setFont(undefined,'bold');doc.setFontSize(11);doc.setTextColor(0);doc.text(pdfText(doc,text),65,y+15);return y+h+PDF_CONFIG.SPACE_AFTER_TITLE;}
function sectionParcelInfo(doc,r,startY){let currentY=sectionTitle(doc,`Informations parcelle — Capakey: ${r.capakey}`,startY);doc.setFontSize(9);const lineHeight=doc.getLineHeight()*1.4;const x1=65,x2=140,x3=250,x4=330,x5=440,x6=480;const colW1=x3-x2,colW2=x5-x4,colW3=doc.internal.pageSize.getWidth()-55-x6;doc.setFont(undefined,'bold');doc.setTextColor(100);doc.text('Division:',x1,currentY);doc.text('Section:',x3,currentY);doc.text('Date:',x5,currentY);doc.setFont(undefined,'normal');doc.setTextColor(0);doc.text(pdfText(doc,r.divCad||'—',colW1),x2,currentY);doc.text(pdfText(doc,r.section||'—',colW2),x4,currentY);doc.text(pdfText(doc,r.dateSituation||'—',colW3),x6,currentY);currentY+=lineHeight;doc.setFont(undefined,'bold');doc.setTextColor(100);doc.text('Numéro:',x1,currentY);doc.text('Nature:',x3,currentY);doc.setFont(undefined,'normal');doc.setTextColor(0);doc.text(pdfText(doc,r.number||'—',colW1),x2,currentY);doc.text(pdfText(doc,r.natureLabel||r.nature||'—',colW2+colW3),x4,currentY);currentY+=lineHeight;const{ha,a,ca}=m2toHaACa(r.areaM2);const formattedSurface=`${pad(ha)} ha  ${pad(a)} a  ${pad(ca)} ca`;doc.setFont(undefined,'bold');doc.setTextColor(100);doc.text('Surface totale:',x1,currentY);doc.setFont(undefined,'normal');doc.setTextColor(0);doc.text(pdfText(doc,formattedSurface),x2,currentY);currentY+=lineHeight;return currentY;}
function sectionOwners(doc,r,startY){let currentY=sectionTitle(doc,'Propriétaires & droits',startY);const C=PDF_CONFIG;const W=doc.internal.pageSize.getWidth();const xLabel=65,xValue=170;const valueMaxWidth=W-85-xValue;if(r.countriesNonBE&&r.countriesNonBE!=='—'){doc.setFont(undefined,'italic');doc.setFontSize(8);doc.setTextColor(80);currentY+=doc.getTextDimensions(pdfText(doc,`Pays (≠ BE) : ${r.countriesNonBE}`)).h+2;}
const list=(r.ownersList?.length?r.ownersList:[{name:"—",officialId:"",partyType:"",right:"Gestion",share:"",addr:"",country:""}]);list.forEach((owner,index)=>{const ownerData={"Propriétaire":titleCaseSmart(owner.name||'—'),"ID National":stripAccents(owner.officialId||'—'),"Type":stripAccents(owner.partyType||'—'),"Droit":stripAccents(owner.right||'—'),"Quote-part":stripAccents(owner.share||'—'),"Adresse":formatAddrWithCountry(titleCaseSmart(owner.addr||''),owner.country)};for(const[label,value]of Object.entries(ownerData)){doc.setFontSize(9);doc.setFont(undefined,'bold');doc.setTextColor(100);const labelDims=doc.getTextDimensions(`${label}:`);doc.setFont(undefined,'normal');doc.setTextColor(0);const valueDims=doc.getTextDimensions(toAscii(value),{maxWidth:valueMaxWidth,fontSize:9});const lineY=currentY+labelDims.h/2;doc.text(`${label}:`,xLabel,lineY);doc.text(toAscii(value),xValue,lineY);currentY+=valueDims.h+3;}
if(index<list.length-1){currentY+=C.OWNER_SEPARATOR_SPACING/2-2;doc.setDrawColor(234,234,234);doc.line(xLabel,currentY,W-85,currentY);currentY+=C.OWNER_SEPARATOR_SPACING/2;}});return currentY;}
function footer(doc,generationDate){const W=doc.internal.pageSize.getWidth(),H=doc.internal.pageSize.getHeight();doc.setDrawColor(234,234,234);doc.line(40,H-64,W-40,H-64);const disclaimer="⚠ AVERTISSEMENT — Ce document n’est pas un extrait cadastral officiel. Il a été généré automatiquement et peut contenir des erreurs. Veuillez toujours vérifier les données sources (Capakey, géométrie, surfaces) auprès des services compétents.";doc.setFillColor(245,245,245);doc.roundedRect(40,H-60,W-80,40,6,6,'F');doc.setFont(undefined,'italic');doc.setFontSize(8);doc.setTextColor(80);doc.text(doc.splitTextToSize(toAscii(disclaimer),W-100),50,H-46);doc.setFont(undefined,'normal');doc.setFontSize(8);doc.setTextColor(120);doc.text("Document généré localement — sans valeur légale — impression A4",40,H-14);if(generationDate){doc.text(`Généré le ${generationDate}`,W-40,H-14,{align:'right'});}}
