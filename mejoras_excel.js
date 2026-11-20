/* Mejoras del editor de Excel · Oficios de exoneración (Municipalidad de Pococí)
   Se carga DESPUÉS del script principal (antes de </body>). Reemplaza funciones globales:
   evaluarFormula, ejecutarFuncion (envoltorio), recalcularHoja, aplicarEdicionCelda. */
(function(){
"use strict";
const ERR=c=>({__err:c});
const isErr=v=>!!v&&typeof v==="object"&&!Array.isArray(v)&&v.__err!==undefined;
const ERRS=/^#(DIV\/0!|N\/A|VALUE!|REF!|NAME\?|NUM!|NULL!)$/;
const ORIG=window.ejecutarFuncion;
const LAZY=new Set(["IFERROR","SI.ERROR","IFNA","SI.ND","ISERROR","ESERROR","ISNA","IF","SI","IFS","SI.CONJUNTO","SWITCH","ELEGIR.V","CHOOSE","ELEGIR"]);

/* ---------- valores ---------- */
function celdaVal(sh,a){
 const c=sh[a];
 if(!c||c.v===undefined||c.v===null||c.v==="")return "";
 if(c.t==="e"||(typeof c.v==="string"&&ERRS.test(c.v)))return ERR(c.w&&ERRS.test(c.w)?c.w:(ERRS.test(c.v)?c.v:"#VALUE!"));
 if(c.f&&c.t==="s"&&(c.v==="TRUE"||c.v==="FALSE"))return c.v==="TRUE";
 return c.v;
}
function num(v){
 if(Array.isArray(v))v=v[0];
 if(isErr(v))return v;
 if(v===""||v==null)return 0;
 if(typeof v==="number")return v;
 if(typeof v==="boolean")return v?1:0;
 const n=Number(String(v).trim().replace(",","."));
 return Number.isNaN(n)?ERR("#VALUE!"):n;
}
function bin(a,b,f){const x=num(a);if(isErr(x))return x;const y=num(b);if(isErr(y))return y;return f(x,y)}
function txt(v){if(Array.isArray(v))v=v[0];return typeof v==="boolean"?(v?"TRUE":"FALSE"):String(v??"")}
function cmp(a,b){
 if(Array.isArray(a))a=a[0];if(Array.isArray(b))b=b[0];
 if(isErr(a))return a;if(isErr(b))return b;
 if(a===""&&typeof b==="number")a=0;if(b===""&&typeof a==="number")b=0;
 if(typeof a==="string"&&typeof b==="string"){a=a.toLowerCase();b=b.toLowerCase()}
 return a<b?-1:a>b?1:0;
}
const primerError=a=>Array.isArray(a)?a.find(isErr):(isErr(a)?a:null);
const mat=r=>Array.isArray(r)?{a:r,R:r._r||r.length,C:r._c||1}:{a:[r],R:1,C:1};
const at=(m,i,j)=>m.a[i*m.C+j];
const eq=(a,b)=>typeof a==="string"&&typeof b==="string"?a.toLowerCase()===b.toLowerCase():a===b;

/* ---------- funciones (búsqueda, errores, lógicas) + delegación al motor original ---------- */
function buscar(val,vec,tipo){ /* devuelve índice 0-based o -1 */
 if(tipo===0){for(let k=0;k<vec.length;k++)if(eq(vec[k],val))return k;return -1}
 let f=-1;
 for(let k=0;k<vec.length;k++){
  const v=vec[k];if(v===""||typeof v!==typeof val)continue;
  const c=cmp(v,val);
  if(tipo>0?c<=0:c>=0)f=k;else break;
 }
 return f;
}
function llamar(name,args,ws){
 if(!LAZY.has(name)){for(const a of args){const e=primerError(a);if(e)return e}}
 const A=args.map(a=>a===undefined?"":a);
 switch(name){
  case "IFERROR":case "SI.ERROR":return isErr(A[0])?A[1]:A[0];
  case "IFNA":case "SI.ND":return isErr(A[0])&&A[0].__err==="#N/A"?A[1]:A[0];
  case "ISERROR":case "ESERROR":return isErr(A[0]);
  case "ISNA":case "ESND":return isErr(A[0])&&A[0].__err==="#N/A";
  case "IF":case "SI":{const c=Array.isArray(A[0])?A[0][0]:A[0];if(isErr(c))return c;return c?(A.length>1?A[1]:true):(A.length>2?A[2]:false)}
  case "IFS":case "SI.CONJUNTO":{for(let p=0;p+1<A.length;p+=2){if(isErr(A[p]))return A[p];if(A[p])return A[p+1]}return ERR("#N/A")}
  case "SWITCH":case "ELEGIR.V":{for(let p=1;p+1<A.length;p+=2)if(eq(A[0],A[p]))return A[p+1];return A.length%2===0?A[A.length-1]:ERR("#N/A")}
  case "CHOOSE":case "ELEGIR":{const n=num(A[0]);return isErr(n)?n:(A[n]!==undefined&&n>=1?A[n]:ERR("#VALUE!"))}
  case "VLOOKUP":case "BUSCARV":case "HLOOKUP":case "BUSCARH":{
   const v=name==="VLOOKUP"||name==="BUSCARV",m=mat(A[1]),idx=num(A[2]);
   if(isErr(idx))return idx;
   const largo=v?m.C:m.R;if(idx<1||idx>largo)return ERR("#REF!");
   const vec=[];for(let k=0;k<(v?m.R:m.C);k++)vec.push(v?at(m,k,0):at(m,0,k));
   const k=buscar(A[0],vec,A.length>3&&!A[3]?0:1);
   return k<0?ERR("#N/A"):(v?at(m,k,idx-1):at(m,idx-1,k));
  }
  case "XLOOKUP":case "BUSCARX":{
   const k=buscar(A[0],flat1(A[1]),0);
   return k<0?(A.length>3?A[3]:ERR("#N/A")):flat1(A[2])[k];
  }
  case "LOOKUP":case "BUSCAR":{
   const vec=flat1(A[1]),k=buscar(A[0],vec,1);
   return k<0?ERR("#N/A"):(A[2]!==undefined?flat1(A[2]):vec)[k];
  }
  case "MATCH":case "COINCIDIR":{
   const t=A.length>2?num(A[2]):1,k=buscar(A[0],flat1(A[1]),t===0?0:t>0?1:-1);
   return k<0?ERR("#N/A"):k+1;
  }
  case "INDEX":case "INDICE":{
   const m=mat(A[0]);let r=A.length>1?num(A[1]):1,c=A.length>2?num(A[2]):1;
   if(isErr(r))return r;if(isErr(c))return c;
   if(m.R===1&&A.length<3){c=r;r=1}
   return r<1||c<1||r>m.R||c>m.C?ERR("#REF!"):at(m,r-1,c-1);
  }
 }
 const r=ORIG(name,A,ws);
 if(r===null||r===undefined)throw new Error("Función no soportada: "+name);
 if(typeof r==="number"&&!Number.isFinite(r))return ERR("#NUM!");
 return r;
}
const flat1=a=>Array.isArray(a)?a:[a];
window.ejecutarFuncion=function(name,args,ws){return llamar(String(name).toUpperCase(),args,ws)};

/* ---------- evaluador: referencias entre hojas, errores como valores, % y rangos con forma ---------- */
const RE=/"(?:[^"]|"")*"|(?:'(?:[^']|'')+'|[A-Za-zÀ-ÿ_][\wÀ-ÿ.]*)!|\$?[A-Z]{1,3}\$?\d+(?![\w(])|\d+(?:\.\d+)?(?:E[+-]?\d+)?|[A-Z_À-ÿ][\wÀ-ÿ.]*|<=|>=|<>|[()+\-*\/^&,:;=<>%]/gyi;
const CELL=/^\$?[A-Z]{1,3}\$?\d+$/i;
const limpia=t=>t.replace(/\$/g,"").toUpperCase();
function evaluar(formula,ws){
 const src=String(formula||"").replace(/^=/,"").trim(),tk=[];
 let pos=0,m;
 while(pos<src.length){
  if(/\s/.test(src[pos])){pos++;continue}
  RE.lastIndex=pos;m=RE.exec(src);if(!m)return null;
  tk.push(m[0]);pos=RE.lastIndex;
 }
 let i=0;
 const hojaDe=t=>{let n=t.slice(0,-1);if(n[0]==="'")n=n.slice(1,-1).replace(/''/g,"'");
  const k=Object.keys(WBF.Sheets).find(s=>s.toLowerCase()===n.toLowerCase());return k?WBF.Sheets[k]:null};
 function ref(sh){
  const t=tk[i++];if(!CELL.test(t||""))throw new Error("ref");
  const a=limpia(t);
  if(tk[i]!==":")return sh?celdaVal(sh,a):ERR("#REF!");
  i++;const t2=tk[i++];if(!CELL.test(t2||""))throw new Error("rango");
  if(!sh)return ERR("#REF!");
  const r=XLSX.utils.decode_range(a+":"+limpia(t2));
  if((r.e.r-r.s.r+1)*(r.e.c-r.s.c+1)>200000)return ERR("#VALUE!");
  const out=[];
  for(let y=r.s.r;y<=r.e.r;y++)for(let x=r.s.c;x<=r.e.c;x++)out.push(celdaVal(sh,XLSX.utils.encode_cell({r:y,c:x})));
  out._r=r.e.r-r.s.r+1;out._c=r.e.c-r.s.c+1;return out;
 }
 function primario(){
  const t=tk[i];if(t===undefined)throw new Error("fin");
  if(t==="+"||t==="-"){i++;const v=num(primario());return isErr(v)?v:(t==="-"?-v:v)}
  if(t==="("){i++;const v=comparacion();if(tk[i++]!==")")throw new Error(")");return v}
  if(t[0]==='"'){i++;return t.slice(1,-1).replace(/""/g,'"')}
  if(t.endsWith("!")){i++;return ref(hojaDe(t))}
  if(CELL.test(t))return ref(ws);
  if(/^\d/.test(t)){i++;return Number(t)}
  if(/^[A-Z_À-ÿ]/i.test(t)){
   i++;const name=t.toUpperCase();
   if(tk[i]!=="("){
    if(name==="TRUE"||name==="VERDADERO")return true;
    if(name==="FALSE"||name==="FALSO")return false;
    return ERR("#NAME?");
   }
   i++;const args=[];
   while(tk[i]!==")"){
    if(i>=tk.length)throw new Error(")");
    if(tk[i]===","||tk[i]===";"){args.push("");i++;continue}
    args.push(comparacion());
    if(tk[i]===","||tk[i]===";")i++;else if(tk[i]!==")")throw new Error("arg");
   }
   i++;return llamar(name,args,ws);
  }
  throw new Error("token");
 }
 const pct=()=>{let v=primario();while(tk[i]==="%"){i++;v=bin(v,0,x=>x/100)}return v};
 const pot=()=>{let v=pct();while(tk[i]==="^"){i++;v=bin(v,pct(),(a,b)=>{const p=a**b;return Number.isFinite(p)?p:ERR("#NUM!")})}return v};
 const prod=()=>{let v=pot();while(tk[i]==="*"||tk[i]==="/"){const op=tk[i++];v=bin(v,pot(),(a,b)=>op==="*"?a*b:(b===0?ERR("#DIV/0!"):a/b))}return v};
 const suma=()=>{let v=prod();while(tk[i]==="+"||tk[i]==="-"){const op=tk[i++];v=bin(v,prod(),(a,b)=>op==="+"?a+b:a-b)}return v};
 const cat=()=>{let v=suma();while(tk[i]==="&"){i++;const r=suma();v=isErr(v)?v:isErr(r)?r:txt(v)+txt(r)}return v};
 function comparacion(){
  const v=cat(),op=tk[i];
  if(!["=","<>",">","<",">=","<="].includes(op))return v;
  i++;const c=cmp(v,cat());if(isErr(c))return c;
  return op==="="?c===0:op==="<>"?c!==0:op===">"?c>0:op==="<"?c<0:op===">="?c>=0:c<=0;
 }
 try{const v=comparacion();return i===tk.length?v:null}catch(e){return null}
}
window.evaluarFormula=evaluar;

/* ---------- recálculo de todo el libro (incluye fórmulas entre hojas) ---------- */
function aplicar(cell,val,hoja,addr){
 let v=Array.isArray(val)?val[0]:val,t,nv;
 if(v===undefined)v="";
 if(isErr(v)){t="s";nv=v.__err}
 else if(typeof v==="number"){if(Number.isFinite(v)){t="n";nv=v}else{t="s";nv="#NUM!"}}
 else if(typeof v==="boolean"){t="s";nv=v?"TRUE":"FALSE"}
 else{t="s";nv=String(v)}
 if(cell.v===nv&&cell.t===t)return false;
 cell.v=nv;cell.t=t;delete cell.w;
 try{cell.w=XLSX.utils.format_cell(cell)}catch(e){}
 CELDAS_MODIFICADAS.add(hoja+"!"+addr);
 return true;
}
window.recalcularHoja=function(){
 const items=[];
 for(const nombre of WBF.SheetNames){
  const ws=WBF.Sheets[nombre];if(!ws||!ws["!ref"])continue;
  for(const a of Object.keys(ws)){if(a[0]==="!")continue;const c=ws[a];if(c&&c.f)items.push({nombre,ws,a,c})}
 }
 for(let pass=0;pass<Math.min(30,items.length+1);pass++){
  let cambio=false;
  for(const it of items){
   const v=evaluar(it.c.f,it.ws);if(v===null)continue;
   if(aplicar(it.c,v,it.nombre,it.a))cambio=true;
  }
  if(!cambio)break;
 }
};

/* ---------- edición: ya no se pierden fórmulas ni formatos al solo hacer clic ---------- */
window.aplicarEdicionCelda=function(el){
 const ws=WBF.Sheets[HOJA_EXCEL],a=el.dataset.cell,c=ws[a];
 const nuevo=el.innerText.replace(/\r?\n/g," ").trim();
 const mostrado=valorCelda(c).trim();
 if(nuevo===mostrado||(c&&c.f&&nuevo==="="+c.f))return;
 guardarValorCelda(a,nuevo);
};
const tb=document.getElementById("tb");
tb.addEventListener("focusin",e=>{
 const td=e.target.closest&&e.target.closest("[data-cell]");if(!td)return;
 const c=WBF&&WBF.Sheets[HOJA_EXCEL]&&WBF.Sheets[HOJA_EXCEL][td.dataset.cell];
 if(c&&c.f&&td.textContent.trim()!=="="+c.f)td.textContent="="+c.f;
});
function enfocar(a){
 const n=tb.querySelector('[data-cell="'+a+'"]');if(!n)return;
 seleccionarCelda(a);n.focus();
 const r=document.createRange();r.selectNodeContents(n);
 const s=getSelection();s.removeAllRanges();s.addRange(r);
}
tb.addEventListener("keydown",e=>{
 const td=e.target.closest&&e.target.closest("[data-cell]");if(!td)return;
 const k=e.key;if(!["Enter","Tab","Escape","ArrowUp","ArrowDown"].includes(k))return;
 e.preventDefault();
 const p=XLSX.utils.decode_cell(td.dataset.cell),c=WBF.Sheets[HOJA_EXCEL][td.dataset.cell];
 if(k==="Escape"){td.textContent=valorCelda(c);td.blur();return}
 const atras=e.shiftKey||k==="ArrowUp";
 const sig=XLSX.utils.encode_cell(k==="Tab"?{r:p.r,c:Math.max(0,p.c+(e.shiftKey?-1:1))}:{r:Math.max(0,p.r+(atras?-1:1)),c:p.c});
 td.blur();enfocar(sig);
});
tb.addEventListener("paste",e=>{
 if(!(e.target.closest&&e.target.closest("[data-cell]")))return;
 e.preventDefault();
 document.execCommand("insertText",false,(e.clipboardData.getData("text/plain")||"").replace(/\r?\n/g," "));
});
})();
