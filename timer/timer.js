// Логіка «Таймера PRO» (timer/index.html). Окремий файл, щоб CSP сторінки міг заборонити вбудовані скрипти.
"use strict";
const $ = (s,r)=> (r||document).querySelector(s);
const $$ = (s,r)=> Array.from((r||document).querySelectorAll(s));
const KEY = "trainingTimer.v1";
/* declared early: the boot timers below read it before the display code runs */
const IS_DISPLAY = location.hash === "#display";
/* Embedded in the "Веб на слайді" PowerPoint add-in as timer/index.html?embed=1.
   The add-in page is the parent on the same origin: it hands over this slide's
   settings (HOST.cfg, a JSON string), says whether PowerPoint is editing or
   showing the slide (HOST.view) and stores changes in the presentation (HOST.save).
   Without a host the file works exactly as the standalone timer. */
const EMBED = /[?&]embed=1(&|$)/.test(location.search);
const HOST = (function(){
  if(!EMBED || window.parent===window) return null;
  try{ const h = window.parent.WOS_TIMER_HOST; return h && typeof h.save==="function" ? h : null; }catch(e){ return null; }
})();
function hostView(){ try{ return HOST ? HOST.view : ""; }catch(e){ return ""; } }
const clamp = (v,a,b)=> Math.max(a,Math.min(b,v));

/* ================= config ================= */
const DEF = {
  title:"Тренінг", mode:"countdown", theme:"dtek-light", vis:"ring",
  bgFill:true, tenths:true, autoNext:true, overtime:false, loop:false,
  bells:"", bellsOver:0, protect:0, step:30, msgHold:6, warnColor:"", dangerColor:"", noRed:false,
  msgs:["Завершуйте","Час питань","Гучніше, будь ласка","Переходимо далі"],
  amb:"none", ambVol:.4,
  logo:{on:false, url:"", size:44, pos:"tr", inv:false, plate:false, rad:18},
  warn:60, danger:10,
  sound:true, volume:.6, pack:"beep", tickFrom:5, flash:false, voice:false, wake:true,
  musicOn:true, musicLoop:true, musicVol:.35,
  noise:false, noiseTh:.35,
  buzzOn:false, buzzAnswer:5,
  showScore:false, showClock:true, showEnd:true, noRepeat:true, teamCount:3,
  cd:{dur:"10:00", label:"Робота"},
  iv:{prep:"0:10", work:"0:40", rest:"0:20", rounds:8, wlabel:"Робота", rlabel:"Пауза", restlast:false, wcolor:"", rcolor:"", rotation:false},
  pg:{items:[
    {name:"Знайомство", dur:"5:00", color:"", note:""},
    {name:"Теорія", dur:"15:00", color:"", note:""},
    {name:"Робота в групах", dur:"20:00", color:"#2fd07a", note:"Обговоріть у групах і запишіть три висновки"},
    {name:"Презентації", dur:"15:00", color:"", note:""},
    {name:"Підсумки", dur:"5:00", color:"", note:""}
  ]},
  tn:{dur:"3:00", order:"list"},
  ck:{bank:"5:00", inc:0, delay:0, glass:false, periods:0, pdur:30},
  roster:[{name:"Команда 1",score:0},{name:"Команда 2",score:0},{name:"Команда 3",score:0}]
};
const BUILTIN = {
  "Помодоро 25 / 5 хв":{mode:"interval",vis:"disc",iv:{prep:"0:00",work:"25:00",rest:"5:00",rounds:4,wlabel:"Фокус",rlabel:"Перерва",restlast:true,wcolor:"",rcolor:"#2fd07a",rotation:false},autoNext:true,warn:120,danger:10},
  "Табата 20 / 10 с x8":{mode:"interval",vis:"disc",iv:{prep:"0:10",work:"0:20",rest:"0:10",rounds:8,wlabel:"Робота",rlabel:"Відпочинок",restlast:false,wcolor:"#ff4d5e",rcolor:"#2fd07a",rotation:false},autoNext:true,warn:0,danger:5,tickFrom:3},
  "Мозковий штурм 10 хв":{mode:"countdown",vis:"liquid",cd:{dur:"10:00",label:"Генеруємо ідеї"},warn:120,danger:15,autoNext:false},
  "Виступи по 3 хв":{mode:"turns",vis:"disc",tn:{dur:"3:00",order:"list"},autoNext:false,overtime:true,warn:60,danger:15},
  "Воркшоп 60 хв":{mode:"program",vis:"disc",autoNext:false,warn:120,danger:20},
  "Кава-брейк 15 хв":{mode:"countdown",vis:"sand",cd:{dur:"15:00",label:"Перерва"},warn:180,danger:30,autoNext:false},
  "Швидкий нетворкінг":{mode:"interval",vis:"traffic",iv:{prep:"0:15",work:"4:00",rest:"0:30",rounds:8,wlabel:"Розмова",rlabel:"Змінюємо стіл",restlast:false,wcolor:"",rcolor:"#ffd23f",rotation:true},autoNext:true,warn:30,danger:10},
  "Квіз для команд":{mode:"countdown",vis:"flip",cd:{dur:"0:30",label:"Питання"},autoNext:false,overtime:false,warn:0,danger:10,tickFrom:5}
};
const COLORS = [["","Авто"],["#4f9cff","Синій"],["#2fd07a","Зелений"],["#ffd23f","Жовтий"],["#ff9f43","Помаранч"],["#ff4d5e","Червоний"],["#a56bff","Фіолет"],["#25d0d8","Бірюза"]];
/* Hue and saturation, so the switch below can tell an alarm colour apart from
   a theme that never had one. */
function hsOf(c){
  const p = hx(c), r = p[0]/255, g = p[1]/255, b = p[2]/255;
  const mx = Math.max(r,g,b), mn = Math.min(r,g,b), d = mx - mn;
  let h = 0;
  if(d){
    if(mx===r) h = (g-b)/d + (g<b ? 6 : 0);
    else if(mx===g) h = (b-r)/d + 2;
    else h = (r-g)/d + 4;
    h *= 60;
  }
  return [h, mx ? d/mx : 0];
}
function isRed(c){
  const s = hsOf(c);
  return s[1] > .35 && (s[0] < 22 || s[0] > 330);
}
/* Violet leads: it is the furthest thing from an alarm that still reads as a
   different state, and it clashes with neither the usual blue accent nor an
   amber warning. Teal and blue stand behind it for themes that are already
   violet. The pick has to stay clear of the colours it is told to avoid, or
   the last seconds would look exactly like the minutes before them. */
const CALM = ["#a56bff","#25d0d8","#2fd07a","#4f9cff"];
function calmColor(avoid){
  const bad = (avoid||[]).filter(Boolean).map(function(c){ return hsOf(c)[0]; });
  bad.push(hsOf(PAL.accent || "#4f9cff")[0]);
  let best = CALM[0], bestGap = -1;
  for(let i=0;i<CALM.length;i++){
    const h = hsOf(CALM[i])[0];
    let gap = 360;
    for(let j=0;j<bad.length;j++){
      let d = Math.abs(h - bad[j]); if(d > 180) d = 360 - d;
      if(d < gap) gap = d;
    }
    if(gap >= 45) return CALM[i];            /* clear of everything, and first in preference */
    if(gap > bestGap){ bestGap = gap; best = CALM[i]; }
  }
  return best;                               /* nothing is clear, so take the furthest away */
}
/* One gate for every colour that reaches the stage, so the switch means what
   its label says: a red the trainer put on a single segment goes too, not just
   the finish. The other segments have already claimed their colours, so it has
   to land where none of them sit: turning a red work period into the same
   green the rest periods use would merge the two halves of an interval. */
function segColor(c){
  if(!(cfg.noRed && c && isRed(c))) return c;
  const taken = [];
  for(let i=0;i<S.program.length;i++){
    const k = S.program[i].color;
    if(k && !isRed(k)) taken.push(k);
  }
  taken.push(PAL.danger);
  return calmColor(taken);
}
function colorOptions(v){
  return COLORS.map(function(c){ return '<option value="'+c[0]+'"'+(c[0]===(v||"")?" selected":"")+'>'+c[1]+'</option>'; }).join("");
}
const STACKED = ["sand","liquid","candle","dots","traffic","flip","energy","coal","code"];

function deepMerge(base, over){
  const out = Array.isArray(base) ? base.slice() : Object.assign({}, base);
  for(const k in over){
    const v = over[k];
    if(v && typeof v==="object" && !Array.isArray(v) && base && typeof base[k]==="object" && !Array.isArray(base[k])) out[k]=deepMerge(base[k],v);
    else out[k] = (v && typeof v==="object") ? JSON.parse(JSON.stringify(v)) : v;
  }
  return out;
}
function loadStore(){
  let s; try{ s = JSON.parse(localStorage.getItem(KEY)||"{}") || {}; }catch(e){ s = {}; }
  if(HOST){
    /* settings belong to the slide; the preset library stays on this computer */
    let c = null; try{ c = JSON.parse(HOST.cfg || "null"); }catch(e){ c = null; }
    return {cfg: c && typeof c==="object" ? c : {}, presets: s.presets || {}};
  }
  return s;
}
function saveStore(){
  if(HOST){
    try{
      const old = JSON.parse(localStorage.getItem(KEY)||"{}") || {};
      localStorage.setItem(KEY, JSON.stringify({cfg:old.cfg||{}, presets:presets}));
    }catch(e){}
    try{ HOST.save(JSON.stringify(cfg)); }catch(e){}
    return;
  }
  try{ localStorage.setItem(KEY, JSON.stringify({cfg:cfg, presets:presets})); }catch(e){}
}
const store = loadStore();
/* Налаштування приходять із файлу презентації, тобто з чужих рук: список учасників
   приводимо до {name: рядок, score: число}, щоб у сторінку не потрапила розмітка. */
function cleanRoster(c){
  const r = Array.isArray(c.roster) ? c.roster : [];
  c.roster = r.filter(function(p){ return p && typeof p==="object"; }).slice(0,200).map(function(p){
    const n = Number(p.score);
    return {name:String(p.name==null ? "" : p.name).slice(0,120), score:isFinite(n) ? Math.round(n) : 0};
  });
  return c;
}
let cfg = cleanRoster(deepMerge(DEF, store.cfg||{}));
let presets = store.presets || {};
/* runs before the first render: everything downstream assumes "M:SS" */
migrateDurs(cfg);
Object.keys(presets).forEach(function(n){ migrateDurs(presets[n]); });

/* ================= time helpers ================= */
function parseDur(str, defUnit){
  if(str==null) return 0;
  let s = String(str).trim().toLowerCase().replace(",",".");
  if(!s) return 0;
  /* "@13:00" means run until that time on the wall clock, the way an agenda
     pins a break to a moment rather than to a length */
  if(s[0]==="@"){
    const p = s.slice(1).split(":");
    const hh = parseInt(p[0],10), mm = parseInt(p[1]||"0",10);
    if(isNaN(hh)) return 0;
    const now = new Date(), t = new Date(now);
    t.setHours(hh, isNaN(mm)?0:mm, 0, 0);
    if(t <= now) t.setDate(t.getDate()+1);
    return t - now;
  }
  let unit = defUnit || "min";
  if(!s.includes(":")){
    if(/[сs]$/.test(s)) unit = "sec";
    else if(/(хв|m)$/.test(s)) unit = "min";
    s = s.replace(/[^\d.]+$/,"").trim();
  }
  if(s.includes(":")){
    const p = s.split(":").map(function(x){ return parseFloat(x)||0; });
    while(p.length<3) p.unshift(0);
    return ((p[0]*3600)+(p[1]*60)+p[2])*1000;
  }
  const n = parseFloat(s)||0;
  return unit==="sec" ? n*1000 : n*60000;
}
function pad(n){ return n<10 ? "0"+n : ""+n; }
function fmt(ms, opt){
  opt = opt||{};
  const neg = ms < 0;
  const a = Math.abs(ms);
  const t = opt.ceil ? Math.ceil(a/1000 - 1e-6) : Math.floor(a/1000);
  const h = Math.floor(t/3600), m = Math.floor((t%3600)/60), s = t%60;
  let out = h>0 ? h+":"+pad(m)+":"+pad(s) : pad(m)+":"+pad(s);
  if(opt.tenths && h===0 && a < 10000) out = pad(m)+":"+pad(Math.floor(a/1000))+"."+Math.floor((a%1000)/100);
  return (neg?"+":"")+out;
}
function fmtShort(ms){
  const t = Math.round(ms/1000), m = Math.floor(t/60), s = t%60;
  return m>0 ? (s ? m+":"+pad(s) : m+" хв") : s+" с";
}
function fmtInput(ms){
  const t = Math.round(ms/1000);
  return t%60===0 ? String(t/60) : Math.floor(t/60)+":"+pad(t%60);
}
function hhmm(d){ return pad(d.getHours())+":"+pad(d.getMinutes()); }

/* ================= palette ================= */
let PAL = {};
function refreshPal(){
  const cs = getComputedStyle(document.body);
  ["accent","ok","warn","danger","ringbg","txt","muted","bg","panel","line"].forEach(function(k){
    PAL[k] = cs.getPropertyValue("--"+k).trim() || "#888";
  });
}

/* ================= themes =================
   Six themes are hand-written in the CSS above. The rest are generated here
   from the design-skill packs, which publish only raw tokens: primary,
   secondary, surface, text, the three state colours and three font families.
   Surfaces, borders and muted text are derived, and every colour is pushed
   toward black or white until it clears a contrast ratio against its own
   background, so no skill palette can ship unreadable text. Generating the
   rules beats hand-writing seventy CSS blocks nobody could maintain. */
const TBASE = [["dtek-light","ДТЕК світла"],["dtek","ДТЕК темна"],["light","Світла"],["dark","Темна"],["corporate","Корпоративна"],
  ["mono","Контраст"],["neon","Неон"],["fantasy","Фентезі"]];
/* key, name, primary, secondary, surface, text, ok, warn, danger, font, display, mono */
const SKINS = [
  ["agentic","Агентик","#FF5701","#F6F6F1","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Playfair Display","Playfair Display","JetBrains Mono"],
  ["ant","Ant Design","#1677ff","#8B5CF6","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Plus Jakarta Sans","Plus Jakarta Sans","JetBrains Mono"],
  ["artistic","Артистична","#3B82F6","#8B5CF6","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Limelight","Limelight","JetBrains Mono"],
  ["basic","Базова","#A855F7","#0A1829","#FFFFFF","#0A1829","#16A34A","#D97706","#DC2626","Nunito","Oswald","JetBrains Mono"],
  ["bento","Бенто","#FAD4C0","#80A1C1","#FFF5E6","#111827","#16A34A","#D97706","#DC2626","Inter","Inter","JetBrains Mono"],
  ["bold","Смілива","#0077BC","#009866","#111111","#111827","#16A34A","#D97706","#DC2626","Archivo Black","Archivo Black","JetBrains Mono"],
  ["brutalism","Бруталізм","#DD614C","#DAA144","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Darker Grotesque","Darker Grotesque","JetBrains Mono"],
  ["cafe","Кав'ярня","#5D4432","#E9E3DD","#F9F7F5","#3E2B1E","#16A34A","#D97706","#DC2626","Poppins","Poppins","JetBrains Mono"],
  ["claude","Claude","#141413","#FAF9F6","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Anthropic Sans","Anthropic Sans","JetBrains Mono"],
  ["claymorphism","Пластилін","#3B82F6","#FFFFFF","#FFFFFF","#1C398E","#16A34A","#D97706","#DC2626","Montserrat","Poppins","JetBrains Mono"],
  ["clean","Чиста","#3B82F6","#8B5CF6","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Roboto","Poppins","Inconsolata"],
  ["codex","Кодекс","#000000","#ffffff","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Open Sans","Open Sans","JetBrains Mono"],
  ["colorful","Барвиста","#3B82F6","#8B5CF6","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Inter","Inter","JetBrains Mono"],
  ["contemporary","Сучасна","#C800DF","#E60076","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Jost","Jost","Overpass Mono"],
  ["cosmic","Космос","#3B82F6","#8B5CF6","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Audiowide","Audiowide","JetBrains Mono"],
  ["creative","Креативна","#3B82F6","#8B5CF6","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Bangers","Bangers","IBM Plex Mono"],
  ["dithered","Дизеринг","#3B82F6","#8B5CF6","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Open Sans","Space Grotesk","IBM Plex Mono"],
  ["doodle","Дудл","#49B6E5","#263D5B","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Delius Swash Caps","Delius Swash Caps","JetBrains Mono"],
  ["dramatic","Драматична","#8B5CF6","#F43F5E","#09090B","#FAFAFA","#16A34A","#D97706","#DC2626","Outfit","Outfit","JetBrains Mono"],
  ["editorial","Редакційна","#111111","#f1f1f1","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Gelasio","Gelasio","Ubuntu Mono"],
  ["enterprise","Корпорація","#0C5CAB","#0a4a8a","#09090b","#fafafa","#10b981","#f59e0b","#ef4444","IBM Plex Sans","IBM Plex Sans","IBM Plex Sans"],
  ["expressive","Виразна","#db2777","#2563eb","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","IBM Plex Mono","IBM Plex Mono","IBM Plex Mono"],
  ["fiction","Фікшн","#222222","#FFE9CE","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Cossette Texte","Cossette Texte","JetBrains Mono"],
  ["flat","Пласка","#F2673C","#8B5CF6","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Inter","Inter","JetBrains Mono"],
  ["friendly","Дружня","#F2D9DC","#D9F2D8","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Noto Serif Display","Noto Serif Display","Space Mono"],
  ["futuristic","Футуризм","#3B82F6","#8B5CF6","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Roboto","Audiowide","Anonymous Pro"],
  ["geometric","Геометрія","#3B82F6","#8B5CF6","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Inter","Inter","JetBrains Mono"],
  ["glassmorphism","Скло","#1856FF","#3A344E","#FFFFFF","#141414","#07CA6B","#E89558","#EA2143","Plus Jakarta Sans","Plus Jakarta Sans","JetBrains Mono"],
  ["gradient","Градієнт","#990FFA","#E60076","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Montserrat","Space Grotesk","JetBrains Mono"],
  ["immersive","Занурення","#00592B","#0023D1","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Oswald","Oswald","JetBrains Mono"],
  ["impeccable","Бездоганна","#CC8800","#C55221","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Chakra Petch","Chakra Petch","JetBrains Mono"],
  ["levels","Рівні","#27272A","#8B5CF6","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Inter","Inter","JetBrains Mono"],
  ["lingo","Лінго","#58cc02","#ce82ff","#FFFFFF","#3c3c3c","#58cc02","#ffc800","#ff4b4b","Nunito","Nunito","JetBrains Mono"],
  ["material","Матеріал","#6442D6","#C8B3FD","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Inter","Roboto","Fira Code"],
  ["matrix","Матриця","#2DB58A","#0B0C14","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Space Mono","Space Mono","Space Mono"],
  ["minimal","Мінімал","#0C0C09","#312C85","#F4F4F1","#0C0C09","#16A34A","#D97706","#DC2626","Open Sans","Inter","Inconsolata"],
  ["modern","Модерн","#553F83","#111111","#553F83","#ffffff","#16A34A","#D97706","#DC2626","IBM Plex Serif","IBM Plex Serif","JetBrains Mono"],
  ["mono","Термінал","#37F712","#00A6F4","#E7E5E4","#78716B","#00A63D","#FE9900","#FF2157","Space Mono","Space Mono","JetBrains Mono"],
  ["neobrutalism","Необруталізм","#FDC800","#432DD7","#FBFBF9","#1C293C","#16A34A","#D97706","#DC2626","Inter","Inter","JetBrains Mono"],
  ["neon","Неон світлий","#BBF351","#00BCFF","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Roboto","STIX Two Text","Source Code Pro"],
  ["neumorphism","Неоморфізм","#006666","#F1F2F5","#E7E5E4","#1E2938","#00A63D","#FE9900","#FF2157","Space Mono","Space Mono","JetBrains Mono"],
  ["pacman","Pac-Man","#2A3FE5","#F4B9B0","#000000","#111827","#16A34A","#D97706","#DC2626","Press Start 2P","Press Start 2P","Space Mono"],
  ["paper","Папір","#111111","#8B5CF6","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Roboto","Montserrat","PT Mono"],
  ["perspective","Перспектива","#00BD7D","#00BD7D","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Poppins","Oswald","JetBrains Mono"],
  ["power","Потуга","#FAFAFA","#FAFAFA","#000000","#ffffff","#16A34A","#D97706","#DC2626","Oswald","Oswald","JetBrains Mono"],
  ["premium","Преміум","#3B82F6","#8B5CF6","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Inter","Inter","JetBrains Mono"],
  ["professional","Професійна","#FECE14","#000000","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Poppins","Poppins","IBM Plex Mono"],
  ["pulse","Пульс","#EA580B","#F59E0B","#FDBA74","#EA580C","#16A34A","#D97706","#DC2626","Limelight","Limelight","JetBrains Mono"],
  ["refined","Витончена","#3B82F6","#8B5CF6","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Playfair Display","Playfair Display","JetBrains Mono"],
  ["retro","Ретро","#3B82F6","#8B5CF6","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Macondo","Macondo","JetBrains Mono"],
  ["riso","Різографія","#F237A1","#2C40A7","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Space Grotesk","Space Grotesk","Overpass Mono"],
  ["roku","Roku","#9333ea","#a855f7","#FFFFFF","#09090b","#10b981","#f59e0b","#ef4444","Inter","Inter","JetBrains Mono"],
  ["sega","Sega","#4502FF","#FFDA14","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","VT323","VT323","JetBrains Mono"],
  ["shadcn","shadcn","#000000","#111111","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Geist","Geist","Fira Code"],
  ["sketch","Ескіз","#1DAD97","#F4EDE0","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Delicious Handrawn","Delicious Handrawn","JetBrains Mono"],
  ["skeumorphism","Скевоморфізм","#FA3C00","#F08321","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Roboto","Germania One","JetBrains Mono"],
  ["sleek","Елегантна","#3B82F6","#8B5CF6","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Inter","Inter","JetBrains Mono"],
  ["spacious","Простір","#3B82F6","#8B5CF6","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Open Sans","Montserrat","IBM Plex Mono"],
  ["square","Квадрат","#3B82F6","#8B5CF6","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Google Sans","Google Sans","Anonymous Pro"],
  ["stitch","Stitch","#072C2C","#FF5F03","#EDEADE","#111827","#16A34A","#D97706","#DC2626","Ubuntu","Oswald","Ubuntu Mono"],
  ["storytelling","Сторітелінг","#3B82F6","#8B5CF6","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","Inter","Abril Fatface","JetBrains Mono"],
  ["terracotta","Теракота","#C56A3C","#F3E9D8","#FFFFFF","#111827","#16A34A","#D97706","#DC2626","DM Serif Display","DM Serif Display","JetBrains Mono"],
  ["tetris","Тетріс","#1C202B","#7107E7","#DFE7FF","#1C398E","#16A34A","#D97706","#DC2626","Bangers","Bangers","JetBrains Mono"],
  ["vibrant","Яскрава","#7C61D4","#EAAE87","#FFFFFF","#2F281D","#16A34A","#D97706","#DC2626","Noto Sans","Fascinate","Fira Code"],
  ["vintage","Вінтаж","#008080","#C0C0C0","#C0C0C0","#000000","#16A34A","#D97706","#DC2626","Silkscreen","Silkscreen","JetBrains Mono"]
];
/* corner radius where the skill's visual style clearly calls for one */
const RAD = {brutalism:2,neobrutalism:2,pacman:0,tetris:4,vintage:0,matrix:2,square:0,
  shadcn:6,codex:6,editorial:4,minimal:4,claymorphism:22,friendly:20,lingo:18,doodle:18,
  bento:18,cafe:14,glassmorphism:16,neumorphism:18,sketch:16,skeumorphism:10};
/* families whose digits are too stylised to read a countdown at a glance */
const DECOR = /Limelight|Bangers|New Rocker|Press Start|VT323|Silkscreen|Fascinate|Abril|Macondo|Germania|Delicious|Delius|Audiowide|Archivo Black|Cossette|Darker Grotesque|Chakra/i;
const SERIF = /Playfair|Gelasio|Serif|Abril|Macondo|Limelight|New Rocker|Germania|Fascinate|STIX|Cossette|Noto Serif/i;
const MONOF = /Mono|Code|Silkscreen|VT323|Press Start|Inconsolata/i;
/* families Google Fonts does not host — the local stack takes over silently */
const NOGF = {"Anthropic Sans":1,"Google Sans":1,"Cossette Texte":1};

function hx(c){
  c = String(c).replace("#","");
  if(c.length===3) c = c[0]+c[0]+c[1]+c[1]+c[2]+c[2];
  return [parseInt(c.slice(0,2),16),parseInt(c.slice(2,4),16),parseInt(c.slice(4,6),16)];
}
function hex(a){
  return "#"+a.map(function(v){ v=Math.max(0,Math.min(255,Math.round(v))); return (v<16?"0":"")+v.toString(16); }).join("");
}
function mix2(a,b,t){ const A=hx(a),B=hx(b); return hex([0,1,2].map(function(i){ return A[i]+(B[i]-A[i])*t; })); }
function lum(c){
  const A = hx(c).map(function(v){ v/=255; return v<=.03928 ? v/12.92 : Math.pow((v+.055)/1.055,2.4); });
  return .2126*A[0]+.7152*A[1]+.0722*A[2];
}
function cr(a,b){ const x=lum(a), y=lum(b); return (Math.max(x,y)+.05)/(Math.min(x,y)+.05); }
function fit(col,bg,need){
  if(cr(col,bg)>=need) return col;
  const to = lum(bg)>.4 ? "#000000" : "#ffffff";
  let out = col;
  for(let t=.06;t<=1.001;t+=.06){ out = mix2(col,to,t); if(cr(out,bg)>=need) return out; }
  return out;
}
function stack(n){
  n = String(n||"").split(",")[0].trim();
  if(!n) return '"Segoe UI",system-ui,Roboto,Arial,sans-serif';
  const tail = MONOF.test(n) ? 'ui-monospace,"Cascadia Mono",Consolas,monospace'
             : SERIF.test(n) ? 'Georgia,"Times New Roman",serif'
             : '"Segoe UI",system-ui,Roboto,Arial,sans-serif';
  return '"'+n+'",'+tail;
}
function skinVars(s){
  const surf=s[4], dark=lum(surf)<.4, ink=dark?"#ffffff":"#000000";
  const bg     = dark ? surf : "#ffffff";   /* light skins sit on white, like the slide */
  const panel  = dark ? mix2(surf,"#ffffff",.08) : "#ffffff";
  const panel2 = dark ? mix2(surf,"#ffffff",.05) : mix2(surf,s[5],.03);
  const line   = mix2(bg,ink,dark?.20:.14);
  /* a token like #111827 on a black surface is a skill default, not a choice:
     start from plain ink instead of dragging it up into grey */
  const txt    = cr(s[5],bg)<3 ? fit(ink,bg,8) : fit(s[5],bg,8);
  const accent = fit(s[2],bg,3.4);
  let accent2  = fit(s[3],bg,3.2);
  if(cr(accent2,accent)<1.25) accent2 = fit(s[6],bg,3.2);   // secondary too close to primary
  return {
    bg:bg, panel:panel, panel2:panel2, line:line,
    txt:txt, muted:fit(mix2(txt,bg,.42),bg,4.6),
    accent:accent, accent2:accent2,
    onAccent: cr(accent,"#000000")>=cr(accent,"#ffffff") ? "#000000" : "#ffffff",
    ok:fit(s[6],bg,3.2), warn:fit(s[7],bg,3.2), danger:fit(s[8],bg,3.2),
    ringbg: mix2(bg,ink,dark?.14:.10),
    shadow: dark ? "0 12px 30px rgba(0,0,0,.5)" : "0 8px 24px rgba(20,30,60,.10)",
    r: (RAD[s[0]]===undefined ? 12 : RAD[s[0]])+"px",
    font: stack(s[9]),
    "font-display": stack(s[10]),
    "font-mono": stack(s[11]),
    "font-num": DECOR.test(String(s[10]).split(",")[0]) ? stack(s[11]) : stack(s[10])
  };
}
/* Two skins share a key with a built-in (neon, mono); their generated rule came
   later in the cascade and silently replaced the built-in look, so they are dropped. */
const BASEKEY = {}; TBASE.forEach(function(b){ BASEKEY[b[0]] = 1; });
const SKINS_OK = SKINS.filter(function(s){ return !BASEKEY[s[0]]; });
/* Skins named after other brands and products stay working for old slides but are
   not offered in the list: a DTEK trainer tool should not suggest them. */
const SKIN_HIDDEN = {claude:1, sega:1, pacman:1, roku:1, ant:1, shadcn:1, stitch:1};
/* Calm, readable skins with a plain sans for the digits: these suit a training room. */
const SKIN_PICK = ["clean","minimal","paper","editorial","levels","material","flat","perspective","enterprise"];
const SKMAP = {};
SKINS_OK.forEach(function(s){ SKMAP[s[0]] = s; });
(function buildThemeCSS(){
  let css = "";
  SKINS_OK.forEach(function(s){
    const v = skinVars(s);
    css += '[data-theme="'+s[0]+'"]{' + Object.keys(v).map(function(k){ return "--"+k+":"+v[k]; }).join(";") + "}\n";
  });
  const el = document.createElement("style");
  el.textContent = css;
  document.head.appendChild(el);
})();
/* Webfonts load only for the theme in use, and with no weight axis: asking a
   family for a weight it lacks fails the whole request, and synthetic bold is
   fine here. Offline every request simply fails and the local stack shows. */
/* Шрифти базових тем і добірки «Для тренінгу» описані у fonts/fonts.css і лежать на цьому ж сайті;
   решта скінів показується системними шрифтами зі свого запасного стека. До Google Fonts сторінка не звертається. */
function loadThemeFonts(){}
function themeName(k){
  const b = TBASE.filter(function(x){ return x[0]===k; })[0];
  return b ? b[1] : (SKMAP[k] ? SKMAP[k][1] : k);
}
/* Seventy-two themes, some light and some dark: the form controls, the
   scrollbars and the mobile browser chrome all follow from the resolved
   background rather than from a hard-coded value. */
function applyScheme(){
  const dark = lum(PAL.bg || "#000000") < .4;
  document.documentElement.style.colorScheme = dark ? "dark" : "light";
  let m = document.querySelector('meta[name="theme-color"]');
  if(!m){ m = document.createElement("meta"); m.setAttribute("name","theme-color"); document.head.appendChild(m); }
  m.setAttribute("content", PAL.bg || "#000000");
  /* scrollbars, number spinners and the open state of a native select are
     painted by the browser, and it only knows which way to paint them from
     this. Seventy-two themes means deriving it, not declaring it once. */
  document.documentElement.style.colorScheme = lum(PAL.bg || "#000000") < .5 ? "dark" : "light";
}
/* Red is the default alarm, not a requirement: some groups read it as threat
   rather than urgency. Overriding the token, not just the canvas colour, keeps
   the chips, the log and the screen flash in step with the choice. */
function applyStateColors(){
  const b = document.body;
  b.style.removeProperty("--warn"); b.style.removeProperty("--danger");
  refreshPal();                       /* PAL now holds the theme's own background */
  if(cfg.warnColor) b.style.setProperty("--warn", fit(cfg.warnColor, PAL.bg, 3.2));
  if(cfg.dangerColor) b.style.setProperty("--danger", fit(cfg.dangerColor, PAL.bg, 3.2));
  /* The switch touches only what is actually red: a theme whose finish is
     already yellow, like the DTEK one, comes through untouched, and a colour
     the trainer picked by hand always wins over it. */
  if(cfg.noRed){
    const used = [];
    if(!cfg.dangerColor && isRed(PAL.danger)){
      const c = calmColor(used); used.push(c);
      b.style.setProperty("--danger", fit(c, PAL.bg, 3.2));
    }
    if(!cfg.warnColor && isRed(PAL.warn)){
      b.style.setProperty("--warn", fit(calmColor(used), PAL.bg, 3.2));
    }
  }
  refreshPal();
  applyScheme();
}
function applyTheme(k){
  if(!SKMAP[k] && !TBASE.some(function(x){ return x[0]===k; })) k = "dtek-light";
  cfg.theme = k;
  loadThemeFonts(k);
  document.body.dataset.theme = k;
  applyStateColors();
}
(function fillThemes(){
  const sel = $("#o_theme");
  const g1 = document.createElement("optgroup"); g1.label = "Для тренінгу";
  TBASE.forEach(function(b){
    const o = document.createElement("option"); o.value=b[0]; o.textContent=b[1]; g1.appendChild(o);
  });
  SKIN_PICK.forEach(function(k){
    const s = SKMAP[k]; if(!s) return;
    const o = document.createElement("option"); o.value=s[0]; o.textContent=s[1]; g1.appendChild(o);
  });
  const light = document.createElement("optgroup"); light.label = "Інші світлі";
  const dark  = document.createElement("optgroup"); dark.label  = "Інші темні";
  SKINS_OK.slice().sort(function(a,b){ return a[1].localeCompare(b[1],"uk"); }).forEach(function(s){
    if(SKIN_HIDDEN[s[0]] || SKIN_PICK.indexOf(s[0])!==-1) return;
    const o = document.createElement("option"); o.value=s[0]; o.textContent=s[1];
    (lum(s[4])<.4 ? dark : light).appendChild(o);
  });
  sel.innerHTML = "";
  sel.appendChild(g1); sel.appendChild(light); sel.appendChild(dark);
})();

/* ================= audio ================= */
let AC = null;
function actx(){
  if(!AC){ const C = window.AudioContext||window.webkitAudioContext; if(!C) return null; AC = new C(); }
  if(AC.state==="suspended") AC.resume();
  return AC;
}
function tone(freq, dur, delay, type, vol){
  if(!cfg.sound) return;
  const c = actx(); if(!c) return;
  const t = c.currentTime + (delay||0);
  const o = c.createOscillator(), g = c.createGain();
  o.type = type||"sine"; o.frequency.setValueAtTime(freq,t);
  g.gain.setValueAtTime(0.0001,t);
  g.gain.exponentialRampToValueAtTime(Math.max(0.0002, cfg.volume*(vol==null?1:vol)), t+0.012);
  g.gain.exponentialRampToValueAtTime(0.0001, t+dur);
  o.connect(g); g.connect(c.destination);
  o.start(t); o.stop(t+dur+0.05);
}
const PACKS = {
  beep:{
    seg:function(){ tone(680,.15,0,"sine",.8); tone(1020,.28,.15,"sine",.75); },
    end:function(){ [523,659,784,1046].forEach(function(f,i){ tone(f,.45,i*.15,"sine",.8); }); },
    warn:function(){ tone(700,.22,0,"triangle",.55); },
    tick:function(){ tone(1180,.05,0,"square",.28); }
  },
  gong:{
    seg:function(){ tone(196,1.3,0,"sine",.9); tone(294,1.0,0,"sine",.35); },
    end:function(){ tone(131,2.4,0,"sine",1); tone(196,2.0,.06,"sine",.5); tone(262,1.6,.12,"sine",.28); },
    warn:function(){ tone(262,.8,0,"sine",.5); },
    tick:function(){ tone(880,.05,0,"sine",.2); }
  },
  chime:{
    seg:function(){ [1319,1568].forEach(function(f,i){ tone(f,.55,i*.12,"triangle",.6); }); },
    end:function(){ [1047,1319,1568,2093].forEach(function(f,i){ tone(f,.7,i*.13,"triangle",.55); }); },
    warn:function(){ tone(1047,.4,0,"triangle",.4); },
    tick:function(){ tone(1568,.04,0,"triangle",.18); }
  },
  soft:{
    seg:function(){ tone(440,.45,0,"sine",.5); tone(660,.35,.1,"sine",.3); },
    end:function(){ [440,554,659].forEach(function(f,i){ tone(f,.85,i*.2,"sine",.5); }); },
    warn:function(){ tone(392,.35,0,"sine",.35); },
    tick:function(){ tone(660,.05,0,"sine",.15); }
  }
};
function pk(){ return PACKS[cfg.pack] || PACKS.beep; }
const sndTick = function(){ pk().tick(); };
const sndWarn = function(){ pk().warn(); };
const sndSeg  = function(){ pk().seg(); if(cfg.flash) flashScreen(); };
const sndEnd  = function(){ pk().end(); if(cfg.flash) flashScreen(); };
/* ================= voice clips =================
   Nothing is synthesised here. The systems these trainings run on have no
   Ukrainian voice, and a wrong-language robot reading a stage name is worse
   than silence. The trainer records the phrases once and assigns a file to
   each event instead. Clips live in their own localStorage key as data URLs,
   so they survive a reload, stay out of the settings export, and each one is
   size-capped to keep the quota safe. */
const CLIPKEY = "trainingTimer.clips.v1";
const CLIP_MAX = 700*1024;          /* one file, before base64 */
const CLIP_TOTAL = 3.6*1024*1024;   /* everything together, after base64 */
const CLIP_EVENTS = [["evt:seg","Старт етапу"],["evt:warn","Попередження"],["evt:end","Час вийшов"]];
let clips = {};
try{ clips = JSON.parse(localStorage.getItem(CLIPKEY)||"{}"); }catch(e){ clips = {}; }
const clipAudio = {};
function clipsBytes(){ let n=0; for(const k in clips) n += clips[k].length; return n; }
function clipKB(s){ return Math.round(s*0.75/1024); }   /* data URL back to real bytes */
function clipsSave(){
  try{ localStorage.setItem(CLIPKEY, JSON.stringify(clips)); return true; }
  catch(e){ toast("Сховище браузера переповнене, файл не збережено"); return false; }
}
function playClip(key){
  const d = clips[key]; if(!d) return false;
  let a = clipAudio[key];
  if(!a){ a = new Audio(d); clipAudio[key] = a; }
  try{ a.pause(); a.currentTime = 0; }catch(e){}
  a.volume = clamp(cfg.volume+.2,0,1);
  a.play().catch(function(){});
  return true;
}
/* a clip recorded for this exact name wins; the generic one covers the rest */
function announceSeg(name){
  if(!cfg.voice) return;
  if(!playClip("seg:"+name)) playClip("evt:seg");
}
function clipEvent(key){ if(cfg.voice) playClip(key); }
function clipNames(){
  const out = [], seen = {};
  function add(n){ if(n && !seen[n]){ seen[n]=1; out.push(n); } }
  S.program.forEach(function(s){ add(s.name); });
  cfg.pg.items.forEach(function(s){ add(s.name); });
  cfg.roster.forEach(function(p){ add(p.name); });
  /* a clip whose stage was renamed still shows up, so it can be removed */
  Object.keys(clips).forEach(function(k){ if(k.indexOf("seg:")===0) add(k.slice(4)); });
  return out;
}
let clipTarget = "";
function renderClips(){
  const box = $("#clipList"); if(!box) return;
  box.innerHTML = "";
  const rows = CLIP_EVENTS.concat(clipNames().map(function(n){ return ["seg:"+n, n]; }));
  rows.forEach(function(r){
    const key = r[0], has = !!clips[key];
    const row = document.createElement("div"); row.className = "item";
    const t = document.createElement("span");
    t.style.cssText = "flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap";
    t.textContent = r[1]; row.appendChild(t);
    const sz = document.createElement("span");
    sz.className = "hint"; sz.style.flex = "none";
    sz.textContent = has ? clipKB(clips[key].length)+" КБ" : "нема файлу";
    row.appendChild(sz);
    const play = document.createElement("button");
    play.className = "btn sm"; play.title = "Прослухати"; play.setAttribute("aria-label","Прослухати: "+r[1]);
    play.innerHTML = '<svg viewBox="0 0 24 24" width="12" height="12" fill="currentColor"><path d="M6 4l14 8-14 8z"/></svg>';
    play.disabled = !has;
    play.onclick = function(){ playClip(key); };
    row.appendChild(play);
    const pick = document.createElement("button");
    pick.className = "btn sm"; pick.textContent = has ? "Замінити" : "Файл";
    pick.onclick = function(){ clipTarget = key; $("#clipFile").click(); };
    row.appendChild(pick);
    const drop = document.createElement("button");
    drop.className = "btn sm danger"; drop.textContent = "×";
    drop.title = "Прибрати файл"; drop.setAttribute("aria-label","Прибрати файл: "+r[1]);
    drop.disabled = !has;
    drop.onclick = function(){ delete clips[key]; delete clipAudio[key]; clipsSave(); renderClips(); };
    row.appendChild(drop);
    box.appendChild(row);
  });
  $("#clipInfo").textContent = "Зайнято " + clipKB(clipsBytes()) + " КБ приблизно з 2700 КБ. Один файл до 700 КБ. Формати ті, що грає браузер: mp3, ogg, m4a, wav.";
}
let flashT2 = 0;
function flashScreen(){
  const f = $("#flashOvl");
  f.classList.remove("go"); void f.offsetWidth; f.classList.add("go");
  clearTimeout(flashT2); flashT2 = setTimeout(function(){ f.classList.remove("go"); }, 2900);
}

/* ================= music ================= */
let musicUrl = null, fadeT = 0;
function musicStart(){
  const a = $("#bgm");
  if(!musicUrl || !cfg.musicOn) return;
  clearInterval(fadeT);
  a.loop = !!cfg.musicLoop; a.volume = cfg.musicVol;
  a.play().catch(function(){});
}
function musicPause(){ clearInterval(fadeT); $("#bgm").pause(); }
function musicFadeOut(){
  const a = $("#bgm");
  if(a.paused) return;
  clearInterval(fadeT);
  fadeT = setInterval(function(){
    a.volume = Math.max(0, a.volume - .05);
    if(a.volume <= .001){ clearInterval(fadeT); a.pause(); a.currentTime = 0; a.volume = cfg.musicVol; }
  }, 70);
}

/* A quest master and a facilitator both want the same closing number: how far
   the session drifted from the plan, and how often it ran over. */
function logSummary(){
  const el = $("#logSum"); if(!el) return;
  if(!S.log.length){ el.textContent = ""; return; }
  let plan = 0, act = 0, over = 0;
  S.log.forEach(function(e){ plan += e.planned; act += e.actual; if(e.actual - e.planned > 3000) over++; });
  const d = act - plan;
  el.textContent = "Етапів " + S.log.length + " · план " + fmt(plan) + " · факт " + fmt(act) +
    " · різниця " + (d>=0 ? "+" : "−") + fmt(Math.abs(d)) + " · з перебором " + over;
}

/* ================= bells inside a segment =================
   Debate timers ring at fixed moments, not only at zero: a bell at the last
   minute means something different from a bell at the end. The field takes a
   list of moments counted back from the end, so "1хв, 30с" rings twice. */
function bellList(){
  return String(cfg.bells||"").split(/[,;]+/).map(function(s){ return parseDur(s.trim(),"sec"); })
    .filter(function(ms){ return ms>0; });
}
function bellsFor(seg){
  return bellList().filter(function(ms){ return ms < seg.ms; });
}

/* ================= messages on the stage =================
   A facilitator needs to say "wrap up" without saying it out loud. The line
   sits over the timer for a few seconds and never touches the clock. */
let msgT = 0;
function showMsg(text){
  const el = $("#stageMsg");
  el.textContent = text || "";
  el.classList.toggle("on", !!text);
  clearTimeout(msgT);
  if(text) msgT = setTimeout(function(){ el.classList.remove("on"); }, Math.max(2000, cfg.msgHold*1000));
}
function renderMsgs(){
  const box = $("#msgList"); if(!box) return;
  box.innerHTML = "";
  cfg.msgs.forEach(function(m,i){
    const row = document.createElement("div"); row.className = "item";
    const key = document.createElement("span"); key.className = "key";
    key.textContent = i<9 ? "⇧"+(i+1) : "·";
    key.setAttribute("aria-hidden","true");    /* the badge repeats the shortcut, it is not a name */
    row.appendChild(key);
    const inp = document.createElement("input");
    inp.className = "inp"; inp.style.flex = "1"; inp.value = m;
    inp.setAttribute("aria-label", "Репліка " + (i+1) + (i<9 ? ", Shift плюс " + (i+1) : ""));
    inp.oninput = function(){ cfg.msgs[i] = inp.value; persist(true); };
    row.appendChild(inp);
    const send = document.createElement("button");
    send.className = "btn sm"; send.textContent = "На екран";
    send.onclick = function(){ showMsg(cfg.msgs[i]); };
    row.appendChild(send);
    const del = document.createElement("button");
    del.className = "btn sm danger"; del.textContent = "×"; del.setAttribute("aria-label","Прибрати репліку");
    del.onclick = function(){ cfg.msgs.splice(i,1); renderMsgs(); persist(); };
    row.appendChild(del);
    tameInputs(row);
    box.appendChild(row);
  });
}

/* ================= ambient sound =================
   Generated, not played from a file: brown noise through a moving filter is
   enough for rain or wind, and the tool stays one self-contained file. */
let ambNodes = null;
function ambBuffer(c){
  const len = 3*c.sampleRate, buf = c.createBuffer(1,len,c.sampleRate), d = buf.getChannelData(0);
  let last = 0;
  for(let i=0;i<len;i++){
    const w = Math.random()*2-1;
    last = (last + .02*w)/1.02;
    d[i] = last*3.2;
  }
  return buf;
}
function ambStart(){
  if(!cfg.amb || cfg.amb==="none" || ambNodes) return;
  const c = actx(); if(!c) return;
  const src = c.createBufferSource(); src.buffer = ambBuffer(c); src.loop = true;
  const f = c.createBiquadFilter(), g = c.createGain();
  const lfo = c.createOscillator(), lfoGain = c.createGain();
  if(cfg.amb==="rain"){ f.type="bandpass"; f.frequency.value=1400; f.Q.value=.6; lfo.frequency.value=.13; lfoGain.gain.value=420; }
  else if(cfg.amb==="wind"){ f.type="lowpass"; f.frequency.value=480; f.Q.value=.9; lfo.frequency.value=.06; lfoGain.gain.value=260; }
  else { f.type="lowpass"; f.frequency.value=180; f.Q.value=.5; lfo.frequency.value=.04; lfoGain.gain.value=60; }
  lfo.connect(lfoGain); lfoGain.connect(f.frequency);
  src.connect(f); f.connect(g); g.connect(c.destination);
  g.gain.value = 0;
  g.gain.linearRampToValueAtTime(clamp(cfg.ambVol,0,1)*.5, c.currentTime+1.2);
  src.start(); lfo.start();
  ambNodes = {src:src, g:g, lfo:lfo, c:c};
}
function ambStop(){
  const n = ambNodes; if(!n) return;
  ambNodes = null;
  try{
    n.g.gain.cancelScheduledValues(n.c.currentTime);
    n.g.gain.setValueAtTime(n.g.gain.value, n.c.currentTime);
    n.g.gain.linearRampToValueAtTime(0, n.c.currentTime+1);
    setTimeout(function(){ try{ n.src.stop(); n.lfo.stop(); }catch(e){} }, 1200);
  }catch(e){}
}

/* ================= noise meter ================= */
let micStream = null, analyser = null, micBuf = null;
function noiseStart(){
  if(analyser) return;
  if(!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia){
    cfg.noise=false; $("#o_noise").checked=false; toast("Мікрофон недоступний у цьому браузері"); return;
  }
  navigator.mediaDevices.getUserMedia({audio:true}).then(function(s){
    const c = actx(); if(!c) return;
    micStream = s;
    const src = c.createMediaStreamSource(s);
    analyser = c.createAnalyser(); analyser.fftSize = 512;
    micBuf = new Uint8Array(analyser.fftSize);
    src.connect(analyser);
    $("#noiseBox").classList.add("on");
  }).catch(function(){
    cfg.noise=false; $("#o_noise").checked=false;
    toast("Мікрофон заблоковано. Відкрийте файл через локальний сервер.");
  });
}
function noiseStop(){
  if(micStream){ micStream.getTracks().forEach(function(t){ t.stop(); }); micStream=null; }
  analyser = null; $("#noiseBox").classList.remove("on");
}
function noiseLevel(){
  if(!analyser) return 0;
  analyser.getByteTimeDomainData(micBuf);
  let sum = 0;
  for(let i=0;i<micBuf.length;i++){ const v=(micBuf[i]-128)/128; sum += v*v; }
  return clamp(Math.sqrt(sum/micBuf.length)*3.2, 0, 1);
}

/* ================= state ================= */
const S = {
  program:[], idx:0, segLeft:0, anchor:0, running:false,
  swBase:0, laps:[], log:[], segStart:0,
  lastSec:null, warned:false, ended:false, picked:[], raf:0, bells:[], overBell:0,
  vizRatio:1, vizColor:"#4f9cff", vizPhase:"run", vizDone:false, vizDirty:true, ck:null
};
const Q = { locked:null, out:[], deadline:0, t:0, question:1 };

function buildProgram(){
  const m = cfg.mode, out = [];
  if(m==="countdown"){
    out.push({name:(cfg.cd.label||"Таймер"), ms:Math.max(1000,parseDur(cfg.cd.dur,"min")), kind:"work"});
  } else if(m==="interval"){
    const prep = parseDur(cfg.iv.prep,"sec"), work = parseDur(cfg.iv.work,"sec"), rest = parseDur(cfg.iv.rest,"sec");
    const rounds = clamp(parseInt(cfg.iv.rounds)||1,1,99);
    if(prep>0) out.push({name:"Підготовка", ms:prep, kind:"prep"});
    for(let r=1;r<=rounds;r++){
      out.push({name:(cfg.iv.wlabel||"Робота"), ms:Math.max(1000,work), color:cfg.iv.wcolor, kind:"work", round:r, of:rounds});
      const last = r===rounds;
      if(rest>0 && (!last || cfg.iv.restlast)) out.push({name:(cfg.iv.rlabel||"Пауза"), ms:rest, color:cfg.iv.rcolor, kind:"rest", round:r, of:rounds});
    }
  } else if(m==="program"){
    cfg.pg.items.forEach(function(it){
      const ms = parseDur(it.dur,"min");
      if(ms>0) out.push({name:it.name||"Етап", ms:ms, color:it.color, note:it.note, kind:"work"});
    });
    if(!out.length) out.push({name:"Етап", ms:300000, kind:"work"});
  } else if(m==="turns"){
    const ms = Math.max(1000,parseDur(cfg.tn.dur,"min"));
    let names = cfg.roster.map(function(p){ return p.name; }).filter(Boolean);
    if(!names.length) names = ["Учасник 1","Учасник 2","Учасник 3"];
    if(cfg.tn.order==="random") names = shuffle(names.slice());
    names.forEach(function(n,i){ out.push({name:n, ms:ms, kind:"work", round:i+1, of:names.length}); });
  }
  return out;
}
function shuffle(a){
  for(let i=a.length-1;i>0;i--){ const j=Math.floor(Math.random()*(i+1)); const t=a[i]; a[i]=a[j]; a[j]=t; }
  return a;
}
function totalMs(){ return S.program.reduce(function(s,x){ return s+x.ms; },0); }
function doneMs(){
  let s=0; for(let i=0;i<S.idx;i++) s+=S.program[i].ms;
  const cur = S.program[S.idx];
  if(cur) s += clamp(cur.ms - left(), 0, cur.ms);
  return s;
}
function restMs(){
  let r = Math.max(0,left());
  for(let i=S.idx+1;i<S.program.length;i++) r += S.program[i].ms;
  return r;
}
function compile(){
  if(cfg.mode==="clock"){ ckInit(); return; }
  S.program = buildProgram();
  S.idx = 0; S.segLeft = S.program.length ? S.program[0].ms : 0;
  S.anchor = performance.now(); S.lastSec=null; S.warned=false; S.bells=[]; S.overBell=0; S.ended=false;
  S.running = false; S.segStart = 0; S.log = [];
  renderLog(); render(); renderTimeline(); renderClips();
}
function left(){
  if(cfg.mode==="stopwatch") return 0;
  return S.running ? S.segLeft - (performance.now()-S.anchor) : S.segLeft;
}
function elapsedSW(){ return S.swBase + (S.running ? performance.now()-S.anchor : 0); }

/* ================= session log ================= */
function leaveSeg(){
  const seg = S.program[S.idx];
  if(!seg || !S.segStart) return;
  const actual = Date.now() - S.segStart;
  S.segStart = 0;
  if(actual < 500) return;
  S.log.push({name:seg.name, planned:seg.ms, actual:actual});
  renderLog();
}
function enterSeg(){ S.segStart = S.running ? Date.now() : 0; }

/* ================= controls ================= */
function start(){
  if(S.running) return;
  actx();
  /* The clock keeps its own anchor. Resuming has to move it to now, or the
     whole paused stretch is billed to whoever was on the move; and it must not
     fall into the program branch below, where an empty program would rebuild
     the boards and hand everybody a full bank back. */
  if(cfg.mode==="clock"){
    if(!S.ck) ckInit();
    S.ck.anchor = performance.now();
  } else if(cfg.mode!=="stopwatch"){
    if(!S.program.length) compile();
    if(S.ended){ resetAll(); }
  }
  S.running = true; S.anchor = performance.now();
  if(!S.segStart) S.segStart = Date.now();
  requestWake(); musicStart(); ambStart();
  if(cfg.mode!=="stopwatch"){
    const seg = S.program[S.idx];
    if(seg && Math.abs(S.segLeft-seg.ms) < 50) announceSeg(seg.name);
  }
  tickLoop(); render();
}
function pause(){
  if(!S.running) return;
  if(cfg.mode==="clock") ckStep();          /* bank the sliver since the last tick */
  else if(cfg.mode==="stopwatch") S.swBase = elapsedSW();
  else S.segLeft = left();
  S.running = false; releaseWake(); musicPause(); ambStop(); render();
}
function toggle(){ S.running ? pause() : start(); }
function resetAll(){
  S.running=false; releaseWake(); musicPause(); ambStop();
  if(cfg.mode==="clock"){ ckInit(); return; }
  if(cfg.mode==="stopwatch"){ S.swBase=0; S.laps=[]; renderLaps(); render(); return; }
  compile();
}
function jump(i){
  if(cfg.mode==="stopwatch") return;
  if(i<0 || i>=S.program.length) return;
  leaveSeg();
  S.idx = i; S.segLeft = S.program[i].ms; S.anchor = performance.now();
  S.lastSec=null; S.warned=false; S.bells=[]; S.overBell=0; S.ended=false;
  enterSeg(); render(); renderTimeline();
}
function nextSeg(){
  if(cfg.mode==="clock"){ ckSwitch(1); return; }
  if(cfg.mode==="stopwatch"){ lap(); return; }
  if(S.idx < S.program.length-1) jump(S.idx+1);
  else { leaveSeg(); S.ended=true; S.segLeft=0; S.running=false; releaseWake(); musicFadeOut(); ambStop(); render(); }
}
function prevSeg(){
  if(cfg.mode==="clock"){ ckSwitch(-1); return; }
  if(cfg.mode!=="stopwatch") jump(Math.max(0,S.idx-1));
}
function addTime(ms){
  if(cfg.mode==="clock"){
    const p = ckActive(); if(!p) return;
    p.ms = Math.max(0, p.ms + ms); if(p.ms>0){ p.out=false; }
    render(); return;
  }
  if(cfg.mode==="stopwatch"){
    S.swBase = Math.max(0, elapsedSW()+ms);
    if(S.running) S.anchor = performance.now();
    render(); return;
  }
  const seg = S.program[S.idx]; if(!seg) return;
  const cur = left() + ms;
  seg.ms = Math.max(1000, seg.ms + ms);
  S.segLeft = Math.max(0, cur); S.anchor = performance.now();
  S.ended=false; render(); renderTimeline();
}
function setSegTime(ms){
  const seg = S.program[S.idx]; if(!seg || ms<=0) return;
  seg.ms = ms; S.segLeft = ms; S.anchor = performance.now();
  S.ended=false; S.warned=false; S.lastSec=null;
  render(); renderTimeline();
}
function lap(){
  if(cfg.mode!=="stopwatch") return;
  const t = elapsedSW();
  const prev = S.laps.length ? S.laps[S.laps.length-1].t : 0;
  const nm = (cfg.roster[S.laps.length]||{}).name || ("Коло "+(S.laps.length+1));
  S.laps.push({t:t, d:t-prev, name:nm});
  sndSeg(); renderLaps(); toast(nm+": "+fmt(t));
}
function blitz(sec){
  cfg.mode = "countdown"; cfg.cd.dur = Math.floor(sec/60)+":"+pad(sec%60); cfg.cd.label = "Бліц";
  syncUI(); setMode("countdown"); start(); toast("Бліц пішов");
}

/* ================= engine ================= */
function tickLoop(){
  cancelAnimationFrame(S.raf);
  /* Відлік веде таймер нижче (20 тактів на секунду); окремий покадровий цикл тут більше не потрібен. */
  paced();
}
/* rAF freezes when the window is hidden or occluded, so a timer drives the
   clock too; step() reads performance.now(), so extra calls are harmless. */
if(!IS_DISPLAY) setInterval(function(){ if(S.running) paced(); }, 50);
/* Повне оновлення екрана не частіше 20 разів на секунду: цифри міняються щонайбільше 10 разів,
   а два джерела тактів (кадри й таймер) разом давали 80 перемальовувань і зайве навантаження на процесор. */
let lastStep = 0;
function paced(){
  const n = performance.now();
  if(n - lastStep < 45) return;
  lastStep = n; step();
}

function step(){
  if(cfg.mode==="clock"){ ckStep(); render(); return; }
  if(cfg.mode!=="stopwatch"){
    let guard = 0;
    while(S.running && guard++ < 200){
      const l = left();
      if(l > 0) break;
      const overshoot = -l;
      if(!S.ended) onSegEnd();
      const last = S.idx >= S.program.length-1;
      if(cfg.autoNext && (!last || cfg.loop)){
        leaveSeg();
        S.idx = last ? 0 : S.idx+1;
        S.segLeft = S.program[S.idx].ms - overshoot; S.anchor = performance.now();
        S.lastSec=null; S.warned=false; S.bells=[]; S.overBell=0; S.ended=false;
        enterSeg(); segChanged();
      } else if(cfg.overtime){
        break;
      } else {
        leaveSeg(); S.running = false; S.segLeft = 0; releaseWake(); musicFadeOut(); ambStop(); break;
      }
    }
    audioCues(left());
  }
  render();
}
function onSegEnd(){
  S.ended = true;
  const last = S.idx >= S.program.length-1;
  if(last && !cfg.loop){ sndEnd(); clipEvent("evt:end"); musicFadeOut(); ambStop(); }
  else sndSeg();
}
/* The name must differ from the live-region announce() further down: two
   function declarations with one name leave only the last, and this one was
   losing, so an automatic hand-over played no clip at all. */
function segChanged(){
  const seg = S.program[S.idx];
  if(seg) announceSeg(seg.name);
  renderTimeline();
}
function audioCues(l){
  const seg = S.program[S.idx]; if(!seg) return;
  const secs = Math.ceil(l/1000);
  if(!S.warned && cfg.warn>0 && l <= cfg.warn*1000 && l > 0 && seg.ms > cfg.warn*1000){
    S.warned = true; sndWarn(); clipEvent("evt:warn");
  }
  /* every bell fires once per segment; S.bells holds the ones already rung */
  bellsFor(seg).forEach(function(ms){
    if(l <= ms && l > 0 && S.bells.indexOf(ms)<0){ S.bells.push(ms); sndWarn(); }
  });
  if(cfg.bellsOver>0 && l < 0){
    const n = Math.floor(-l/(cfg.bellsOver*1000));
    if(n > S.overBell){ S.overBell = n; sndWarn(); }
  }
  if(S.lastSec !== secs){
    if(cfg.tickFrom>0 && secs>0 && secs<=cfg.tickFrom && S.lastSec!==null) sndTick();
    S.lastSec = secs;
  }
}

/* ================= chess-clock mode =================
   Borrowed from board-game clocks, where the interesting part is not the
   countdown but what happens at a handover: an increment rewards a short turn
   (Fischer), a delay makes the first seconds free (Bronstein), an hourglass
   moves the time you spend to the other side, and byo-yomi periods keep a
   speaker alive past zero. Each of those is a different conversation format. */
function ckPeople(){
  let names = cfg.roster.map(function(p){ return p.name; }).filter(Boolean);
  if(names.length < 2) names = ["Учасник 1","Учасник 2"];
  if(cfg.ck.glass) names = names.slice(0,2);   /* sand only falls between two */
  return names;
}
function ckInit(){
  const bank = Math.max(1000, parseDur(cfg.ck.bank,"min"));
  const per = clamp(parseInt(cfg.ck.periods)||0, 0, 9);
  S.ck = {
    list: ckPeople().map(function(n){ return {name:n, ms:bank, per:per, pms:0, turns:0, out:false}; }),
    i:0, anchor:performance.now(), delayLeft:0
  };
  S.running = false; S.ended = false;
  render(); ckBoard();
}
function ckActive(){ return S.ck && S.ck.list[S.ck.i]; }
function ckFlag(p){
  p.out = true;
  S.running = false; releaseWake(); musicFadeOut(); ambStop();
  sndEnd(); clipEvent("evt:end");
  showMsg("Час вичерпано: " + p.name);
  if(S.ck.list.filter(function(x){ return !x.out; }).length) ckSwitch(1);
}
function ckStep(){
  if(!S.ck) return;
  const now = performance.now();
  let dt = now - S.ck.anchor;
  S.ck.anchor = now;
  const p = ckActive();
  if(!p || !S.running || dt <= 0) return;
  if(S.ck.delayLeft > 0){            /* the free head of a turn */
    const use = Math.min(S.ck.delayLeft, dt);
    S.ck.delayLeft -= use; dt -= use;
    if(dt <= 0) return;
  }
  if(cfg.ck.glass && S.ck.list.length > 1){
    const other = S.ck.list[(S.ck.i+1) % S.ck.list.length];
    other.ms += dt;                  /* what you spend, the other side gains */
  }
  if(p.ms > 0){
    p.ms -= dt;
    if(p.ms <= 0){
      p.ms = 0; sndWarn();
      if(p.per > 0) p.pms = cfg.ck.pdur*1000; else ckFlag(p);
    }
  } else if(p.per > 0){
    p.pms -= dt;
    if(p.pms <= 0){
      p.per -= 1;
      if(p.per > 0){ p.pms = cfg.ck.pdur*1000; sndWarn(); }
      else { p.pms = 0; ckFlag(p); }
    }
  }
}
function ckSwitch(dir){
  if(!S.ck || !S.ck.list.length) return;
  ckStep();
  const p = ckActive();
  if(p && S.running && !p.out){
    if(p.ms > 0 && cfg.ck.inc > 0) p.ms += cfg.ck.inc*1000;
    if(p.ms <= 0 && p.per > 0) p.pms = cfg.ck.pdur*1000;   /* a finished turn refills the period */
    p.turns++;
  }
  const n = S.ck.list.length;
  let k = S.ck.i;
  for(let s=0;s<n;s++){
    k = (k + (dir<0 ? -1 : 1) + n) % n;
    if(!S.ck.list[k].out) break;
  }
  ckTo(k);
}
function ckTo(k){
  if(!S.ck.list[k] || S.ck.list[k].out) return;
  S.ck.i = k;
  S.ck.delayLeft = Math.max(0, cfg.ck.delay)*1000;
  S.ck.anchor = performance.now();
  sndSeg(); announceSeg(S.ck.list[k].name);
  render(); ckBoard();
}
function ckBoard(){
  const box = $("#ckBoard"); if(!box || !S.ck) return;
  box.innerHTML = "";
  S.ck.list.forEach(function(p,i){
    const d = document.createElement("button");
    d.className = "ck" + (i===S.ck.i ? " cur" : "") + (p.out ? " out" : "");
    d.type = "button";
    const b = document.createElement("b"); b.textContent = p.name;
    const t = document.createElement("i");
    t.textContent = p.ms > 0 ? fmt(p.ms) : (p.per > 0 ? fmt(p.pms) + " · " + p.per : "0:00");
    d.appendChild(b); d.appendChild(t);
    d.setAttribute("aria-label", "Передати хід: " + p.name);
    d.disabled = p.out;
    d.onclick = function(){ ckTo(i); };
    box.appendChild(d);
  });
}
function renderClock(){
  const stage = $("#stage");
  stage.className = "stage vis-plain lay-stack";
  const p = ckActive();
  const inPer = p && p.ms <= 0 && p.per > 0;
  const shown = !p ? 0 : (p.ms > 0 ? p.ms : p.pms);
  const txt = fmt(shown, {ceil:true, tenths: cfg.tenths && shown < 10000 && S.running});
  $("#timeText").textContent = txt;
  $("#segName").textContent = p ? p.name : "Годинник";
  $("#phaseTag").textContent = !p ? "" : p.out ? "Час вичерпано"
    : inPer ? ("Овертайм, лишилось " + p.per) : (S.ck.delayLeft > 0 ? "Затримка" : "");
  const parts = [];
  if(p) parts.push("Ходів " + p.turns);
  if(cfg.ck.inc > 0) parts.push("+" + cfg.ck.inc + " с за хід");
  if(cfg.ck.glass) parts.push("пісочний годинник");
  $("#metaText").textContent = parts.join("  ·  ");
  $("#segNote").textContent = ""; $("#rotPrompt").textContent = "";
  const col = !p || p.out ? PAL.danger
    : shown <= (cfg.danger||10)*1000 ? PAL.danger
    : inPer ? PAL.warn : PAL.accent;
  stage.style.setProperty("--phase", col);
  S.vizRatio = 0; S.vizColor = col;
  S.vizPhase = (!p || p.out || shown <= (cfg.danger||10)*1000) ? "danger" : inPer ? "warn" : "run";
  S.vizDone = !!(p && p.out);
  $("#bgFill").style.transform = "scaleY(0)";
  $("#sessBar").style.width = "0%";
  $("#endAt").textContent = "";
  const now = new Date();
  $("#wallClock").textContent = cfg.showClock ? (pad(now.getHours())+":"+pad(now.getMinutes())+":"+pad(now.getSeconds())) : "";
  const b = $("#btnStart");
  b.querySelector("span").textContent = S.running ? "Пауза" : "Старт";
  b.querySelector("svg").innerHTML = S.running ? '<path d="M8 4v16M16 4v16"/>' : '<path d="M6 4l14 8-14 8z"/>';
  $("#btnPrev").disabled = false;
  stepLabels();
  $("#tbSound").classList.toggle("off", !cfg.sound);
  $("#scoreBox").classList.toggle("on", cfg.showScore);
  const len = txt.length;
  $("#timeText").style.fontSize = "calc(var(--d)*" + (len>7 ? .155 : len>5 ? .19 : .235) + ")";
  ckBoard();
  pushDisplay();
}

/* ================= canvas visuals ================= */
const cvs = $("#viz"), ctx = cvs.getContext("2d");
function fitCanvas(){
  const dpr = Math.min(2, window.devicePixelRatio||1);
  const w = Math.round(cvs.clientWidth*dpr), h = Math.round(cvs.clientHeight*dpr);
  if(w>0 && h>0 && (cvs.width!==w || cvs.height!==h)){ cvs.width=w; cvs.height=h; }
  return {w:cvs.width, h:cvs.height, dpr:dpr};
}
function drawViz(t){
  const v = cfg.vis;
  if(v==="ring" || v==="bar" || v==="plain" || v==="flip"){ cvs.style.display="none"; return; }
  cvs.style.display = "block";
  const g = fitCanvas(); if(!g.w) return;
  const w = g.w, h = g.h, ratio = clamp(S.vizRatio,0,1), col = S.vizColor;
  ctx.clearRect(0,0,w,h);
  if(v==="disc") drawDisc(w,h,ratio,col);
  else if(v==="chrono") drawChrono(w,h,ratio,col);
  else if(v==="square") drawSquare(w,h,ratio,col);
  else if(v==="energy") drawEnergy(w,h,ratio,col,t);
  else if(v==="coal") drawCoal(w,h,ratio,col,t);
  else if(v==="code") drawCode(w,h,ratio,col,t);
  else if(v==="sand") drawSand(w,h,ratio,col,t);
  else if(v==="liquid") drawLiquid(w,h,ratio,col,t);
  else if(v==="candle") drawCandle(w,h,ratio,col,t);
  else if(v==="dots") drawDots(w,h,ratio,col);
  else if(v==="traffic") drawTraffic(w,h,ratio,col);
}
function drawDisc(w,h,ratio,col){
  const cx=w/2, cy=h/2, R=Math.min(w,h)/2-2, r=R*0.58, A=-Math.PI/2;
  ctx.beginPath(); ctx.arc(cx,cy,R,0,Math.PI*2); ctx.arc(cx,cy,r,0,Math.PI*2,true);
  ctx.fillStyle = PAL.ringbg; ctx.fill("evenodd");
  if(ratio<=0) return;
  ctx.beginPath();
  ctx.arc(cx,cy,R,A,A+ratio*Math.PI*2,false);
  ctx.arc(cx,cy,r,A+ratio*Math.PI*2,A,true);
  ctx.closePath(); ctx.fillStyle=col; ctx.fill();
}
/* ================= energy scene =================
   One picture instead of three: the turbine makes the energy on the left, the
   line carries it across the middle, the building spends it on the right. The
   rotor turns and the impulse travels only while the clock runs, the line
   energises as the time is spent, and the floors light bottom-up behind it, so
   the same elapsed fraction is told three ways at once. */
const REDUCED = window.matchMedia ? window.matchMedia("(prefers-reduced-motion: reduce)") : {matches:false};
function motionScale(){ return REDUCED.matches ? .25 : 1; }
let ePhase = 0, ePulse = 0, eLast = 0;
function bez(u, x0,y0, mx,my, x1,y1){
  const m = 1-u;
  return [m*m*x0 + 2*m*u*mx + u*u*x1, m*m*y0 + 2*m*u*my + u*u*y1];
}
function drawEnergy(w,h,ratio,col,t){
  const dt = eLast ? Math.min(120, t - eLast) : 0;
  eLast = t;
  if(S.running){
    ePhase += dt * (0.0013 + (1-ratio)*0.0026) * motionScale();
    ePulse = (ePulse + dt/2300*motionScale()) % 1;
  }
  const done = 1 - clamp(ratio,0,1);
  /* The screen fill sits behind this canvas, so the structures are read against
     two backgrounds during one countdown: the plain stage above the fill line
     and the tinted stage below it. They therefore take the strongest neutral
     the theme already owns — its muted tone, or its text colour when that muted
     tone sits too close to the background. On a dark stage that lands on light
     grey or white, on a light stage on near black. Windows go the other way,
     in the background colour, so they read as openings in a lit wall. */
  const ink = PAL.bg;
  const edge = cr(PAL.muted, PAL.bg) >= 3 ? PAL.muted : PAL.txt;
  const ground = h*0.90;
  const body = function(path){ ctx.fillStyle = edge; path(); ctx.fill(); };

  ctx.lineCap = "round"; ctx.lineJoin = "round";
  ctx.strokeStyle = edge; ctx.lineWidth = Math.max(2, h*0.018);
  ctx.beginPath(); ctx.moveTo(w*0.03, ground); ctx.lineTo(w*0.97, ground); ctx.stroke();

  /* --- turbine --- */
  const tx = w*0.13, hubY = h*0.24, R = h*0.21, tw = Math.max(3, h*0.026);
  body(function(){
    ctx.beginPath();
    ctx.moveTo(tx - tw, ground); ctx.lineTo(tx - tw*0.34, hubY);
    ctx.lineTo(tx + tw*0.34, hubY); ctx.lineTo(tx + tw, ground);
    ctx.closePath();
  });

  /* --- building --- */
  const FLOORS = 6, COLS = 3;
  const bw = Math.min(w*0.22, h*0.52), bh = h*0.66;
  /* not hx: that is the hex parser this file already owns, and a local of the
     same name would put it in the temporal dead zone for the whole function */
  const bldX = w*0.90 - bw, bldY = ground - bh;
  body(function(){ ctx.beginPath(); ctx.rect(bldX, bldY, bw, bh); });
  body(function(){ ctx.beginPath(); ctx.rect(bldX - bw*0.06, bldY - bh*0.03, bw*1.12, bh*0.03); });
  const padX = bw*0.14, padY = bh*0.08;
  const gx = (bw - padX*2)/(COLS*2 - 1), gy = (bh - padY*2)/(FLOORS*2 - 1);
  const litFull = Math.max(1, Math.floor(done*FLOORS));
  const frac = clamp(done*FLOORS - Math.floor(done*FLOORS), 0, 1);
  for(let fl=0; fl<FLOORS; fl++){
    const fy = bldY + bh - padY - gy*(fl*2+1);
    const a = fl < litFull ? 1 : (fl === litFull ? frac : 0);
    for(let c=0;c<COLS;c++){
      const wx = bldX + padX + gx*c*2;
      ctx.fillStyle = ink;                       /* a dark window in a lit wall */
      ctx.fillRect(wx, fy, gx, gy);
      if(a > 0){
        ctx.save(); ctx.globalAlpha = a; ctx.fillStyle = col;
        ctx.fillRect(wx, fy, gx, gy); ctx.restore();
      }
    }
  }
  ctx.fillStyle = ink;
  ctx.fillRect(bldX + bw*0.42, ground - gy*1.1, bw*0.16, gy*1.1);

  /* --- the line between them --- */
  const ax = tx + tw*0.5, ay = hubY + h*0.03;
  const bx = bldX, by = bldY + bh*0.18;
  const mx = (ax+bx)/2, my = Math.max(ay,by) + h*0.17;
  const N = 44, pts = [];
  for(let i=0;i<=N;i++) pts.push(bez(i/N, ax,ay, mx,my, bx,by));
  const lw = Math.max(3, h*0.028);
  const line = function(from,to,style){
    ctx.strokeStyle = style; ctx.lineWidth = lw;
    ctx.beginPath(); ctx.moveTo(pts[from][0], pts[from][1]);
    for(let i=from+1;i<=to;i++) ctx.lineTo(pts[i][0], pts[i][1]);
    ctx.stroke();
  };
  line(0, N, edge);
  const fill = Math.round(done*N);
  if(fill > 0) line(0, fill, col);

  /* the rotor sits above the line, the impulse above everything */
  ctx.save();
  ctx.translate(tx, hubY); ctx.rotate(ePhase); ctx.fillStyle = col;
  for(let i=0;i<3;i++){
    ctx.rotate(Math.PI*2/3);
    ctx.beginPath();
    ctx.moveTo(0, -R*0.11);
    ctx.quadraticCurveTo(R*0.55, -R*0.17, R, -R*0.02);
    ctx.quadraticCurveTo(R*0.55, R*0.05, 0, R*0.11);
    ctx.closePath(); ctx.fill();
  }
  ctx.restore();
  ctx.fillStyle = col;
  ctx.beginPath(); ctx.arc(tx, hubY, R*0.09, 0, Math.PI*2); ctx.fill();

  const p = bez(ePulse, ax,ay, mx,my, bx,by);
  const tail = bez(Math.max(0, ePulse-0.08), ax,ay, mx,my, bx,by);
  ctx.strokeStyle = col; ctx.lineWidth = lw*0.95;
  ctx.beginPath(); ctx.moveTo(tail[0], tail[1]); ctx.lineTo(p[0], p[1]); ctx.stroke();
  const g = ctx.createRadialGradient(p[0], p[1], 0, p[0], p[1], lw*2.6);
  g.addColorStop(0, col); g.addColorStop(1, "rgba(0,0,0,0)");
  ctx.fillStyle = g;
  ctx.beginPath(); ctx.arc(p[0], p[1], lw*2.6, 0, Math.PI*2); ctx.fill();
}

/* ================= pit, wagons, station =================
   The same countdown told as a supply chain. The winder turns over the shaft on
   the left and loaded wagons run the middle; that movement only says the clock
   is going. The reading is the chimney, which takes one more band of paint for
   every eighth of the time that is spent, so the stack fills as the time
   empties and can be counted from the back of the room. When the time is up the
   plant fires: smoke leaves the stack, the hall lights, the line to the pylon
   carries current, and the wagons come back empty.

   The wagons are drawn before the buildings, so they come out from behind the
   shaft house and go in behind the station instead of stopping in mid air. And
   every horizontal measure is capped against the width as well as the height:
   the box is a wide strip on a laptop and nearly square on a phone held
   upright, and without the caps the pit head would grow into the track. */
let cwPhase = 0, cwBurn = 0, cwLast = 0;
function drawCoal(w,h,ratio,col,t){
  const dt = cwLast ? Math.min(120, t - cwLast) : 0;
  cwLast = t;
  const done = 1 - clamp(ratio,0,1);
  const fin = !!S.vizDone;
  if(S.running || fin) cwPhase = (cwPhase + dt*0.000105*motionScale()) % 1;
  if(fin) cwBurn = (cwBurn + dt*0.00040*motionScale()) % 1;

  /* the same reading as the energy scene: structures in the strongest neutral
     the theme owns, openings in the background colour, coal and current in the
     phase colour, so all three survive the screen fill rising behind them */
  const ink = PAL.bg;
  const edge = cr(PAL.muted, PAL.bg) >= 3 ? PAL.muted : PAL.txt;
  const ground = h*0.86;
  const solid = function(path){ ctx.fillStyle = edge; path(); ctx.fill(); };

  const R = Math.min(h*0.135, w*0.042), pitX = w*0.095, hubY = h*0.22;
  const stX = w*0.685, hw0 = Math.min(w*0.036, h*0.13), hw1 = hw0*0.66;
  const stTop = h*0.20, lip = Math.max(3, h*0.022);
  const hallW = w*0.170, hallH = h*0.32;
  const hallX = stX + hw0, hallY = ground - hallH;

  ctx.lineCap = "round"; ctx.lineJoin = "round";
  ctx.strokeStyle = edge; ctx.lineWidth = Math.max(2, h*0.016);
  ctx.beginPath(); ctx.moveTo(w*0.02, ground); ctx.lineTo(w*0.99, ground); ctx.stroke();
  ctx.lineWidth = Math.max(1.5, h*0.009);
  ctx.beginPath();
  for(let x=w*0.05; x<=stX; x+=Math.max(9, w*0.024)){
    ctx.moveTo(x, ground); ctx.lineTo(x, ground + h*0.035);
  }
  ctx.stroke();

  /* --- the wagons --- */
  const cars = 3, carW = Math.min(w*0.10, h*0.36), carH = h*0.15;
  const wr = Math.max(2.5, h*0.030);
  /* both ends of the run start inside a building, so a wagon is never seen
     appearing out of nothing: it leaves the shaft house and enters the plant */
  const from = pitX + R*1.5 - carW, to = stX - hw0*0.4;
  for(let i=0;i<cars;i++){
    const x = from + ((cwPhase + i/cars) % 1)*(to - from);
    const y = ground - wr*2 - carH, q = Math.max(2, carW*0.10);
    solid(function(){
      ctx.beginPath();
      ctx.moveTo(x, y); ctx.lineTo(x + carW, y);
      ctx.lineTo(x + carW*0.88, y + carH); ctx.lineTo(x + carW*0.12, y + carH);
      ctx.closePath();
    });
    ctx.fillStyle = fin ? ink : col;      /* heaped on the way out, empty coming back */
    ctx.beginPath();
    ctx.moveTo(x + q, y + q*0.7);
    if(fin) ctx.lineTo(x + carW - q, y + q*0.7);
    else ctx.quadraticCurveTo(x + carW/2, y - carH*0.36, x + carW - q, y + q*0.7);
    ctx.lineTo(x + carW*0.88 - q*0.6, y + carH - q);
    ctx.lineTo(x + carW*0.12 + q*0.6, y + carH - q);
    ctx.closePath(); ctx.fill();
    [0.28, 0.72].forEach(function(u){
      ctx.fillStyle = edge;
      ctx.beginPath(); ctx.arc(x + carW*u, ground - wr, wr, 0, Math.PI*2); ctx.fill();
      ctx.fillStyle = ink;
      ctx.beginPath(); ctx.arc(x + carW*u, ground - wr, wr*0.34, 0, Math.PI*2); ctx.fill();
    });
  }

  /* --- the pit head --- */
  /* the shaft house is tall enough to swallow a loaded wagon, heap and all,
     so nothing is seen poking over its roof while it waits out of sight */
  solid(function(){ ctx.beginPath(); ctx.rect(pitX - R*1.5, ground - h*0.30, R*3.0, h*0.30); });
  ctx.strokeStyle = edge; ctx.lineWidth = Math.max(3, h*0.024);
  ctx.beginPath();
  ctx.moveTo(pitX - R*0.55, ground); ctx.lineTo(pitX - R*0.55, hubY);
  ctx.moveTo(pitX + R*0.55, ground); ctx.lineTo(pitX + R*0.55, hubY);
  ctx.moveTo(pitX + R*1.35, ground - h*0.30); ctx.lineTo(pitX + R*0.30, hubY + R*0.45);
  ctx.stroke();
  ctx.save();
  ctx.translate(pitX, hubY); ctx.rotate(cwPhase*Math.PI*12);
  ctx.strokeStyle = col; ctx.lineWidth = Math.max(2, h*0.020);
  ctx.beginPath(); ctx.arc(0, 0, R, 0, Math.PI*2); ctx.stroke();
  ctx.lineWidth = Math.max(1.5, h*0.011);
  ctx.beginPath();
  for(let i=0;i<3;i++){
    const a = i*Math.PI/3, sx = Math.cos(a)*R, sy = Math.sin(a)*R;
    ctx.moveTo(-sx, -sy); ctx.lineTo(sx, sy);
  }
  ctx.stroke();
  ctx.restore();
  ctx.fillStyle = col;
  ctx.beginPath(); ctx.arc(pitX, hubY, R*0.18, 0, Math.PI*2); ctx.fill();

  /* --- the chimney, which is the reading --- */
  const bands = 8, bandH = (ground - stTop)/bands;
  const flue = function(){
    ctx.beginPath();
    ctx.moveTo(stX - hw1, stTop); ctx.lineTo(stX + hw1, stTop);
    ctx.lineTo(stX + hw0, ground); ctx.lineTo(stX - hw0, ground);
    ctx.closePath();
  };
  flue(); ctx.fillStyle = PAL.ringbg; ctx.fill();
  ctx.save();
  flue(); ctx.clip();
  ctx.fillStyle = col;
  for(let i=0;i<bands;i++){
    const u = clamp(done*bands - i, 0, 1);
    if(u > 0) ctx.fillRect(stX - hw0, ground - i*bandH - bandH*u, hw0*2, bandH*u);
  }
  ctx.strokeStyle = ink; ctx.lineWidth = Math.max(1.5, h*0.009);
  ctx.beginPath();
  for(let i=1;i<bands;i++){
    ctx.moveTo(stX - hw0, ground - i*bandH); ctx.lineTo(stX + hw0, ground - i*bandH);
  }
  ctx.stroke();
  ctx.restore();
  flue(); ctx.strokeStyle = edge; ctx.lineWidth = Math.max(2, h*0.013); ctx.stroke();
  solid(function(){ ctx.beginPath(); ctx.rect(stX - hw1*1.6, stTop - lip, hw1*3.2, lip); });

  /* --- the hall --- */
  solid(function(){ ctx.beginPath(); ctx.rect(hallX, hallY, hallW, hallH); });
  solid(function(){ ctx.beginPath(); ctx.rect(hallX - hallW*0.04, hallY - lip, hallW*1.08, lip); });
  for(let i=0;i<4;i++){
    ctx.fillStyle = fin ? col : ink;      /* the hall lights when the furnace does */
    ctx.fillRect(hallX + hallW*(0.11 + i*0.21), hallY + hallH*0.24, hallW*0.13, hallH*0.28);
  }

  /* --- and the current that leaves --- */
  const pyX = w*0.945, armY = h*0.38, armW = Math.min(w*0.020, h*0.075);
  ctx.strokeStyle = edge; ctx.lineWidth = Math.max(2, h*0.013);
  ctx.beginPath();
  ctx.moveTo(pyX - armW*1.4, ground); ctx.lineTo(pyX - armW*0.30, armY - h*0.08);
  ctx.lineTo(pyX + armW*0.30, armY - h*0.08); ctx.lineTo(pyX + armW*1.4, ground);
  ctx.moveTo(pyX - armW*1.05, armY); ctx.lineTo(pyX + armW*1.05, armY);
  ctx.moveTo(pyX - armW*1.25, armY + h*0.09); ctx.lineTo(pyX + armW*1.25, armY + h*0.09);
  ctx.stroke();

  const ax = hallX + hallW, ay = hallY - lip, bx = pyX - armW;
  const mx = (ax + bx)/2, my = Math.max(ay, armY) + h*0.09;
  const pts = [];
  for(let i=0;i<=20;i++) pts.push(bez(i/20, ax, ay, mx, my, bx, armY));
  pts.push([pyX + armW, armY]);
  pts.push([w, armY - h*0.02]);           /* and on out of the frame */
  ctx.strokeStyle = fin ? col : edge;
  ctx.lineWidth = Math.max(2, h*0.012);
  ctx.globalAlpha = fin ? 1 : 0.5;
  ctx.beginPath(); ctx.moveTo(pts[0][0], pts[0][1]);
  for(let i=1;i<pts.length;i++) ctx.lineTo(pts[i][0], pts[i][1]);
  ctx.stroke();
  ctx.globalAlpha = 1;

  if(fin){
    const p2 = pts[Math.min(pts.length-1, Math.floor(((cwBurn*1.6) % 1)*pts.length))];
    const glow = ctx.createRadialGradient(p2[0], p2[1], 0, p2[0], p2[1], h*0.06);
    glow.addColorStop(0, col); glow.addColorStop(1, "rgba(0,0,0,0)");
    ctx.fillStyle = glow;
    ctx.beginPath(); ctx.arc(p2[0], p2[1], h*0.06, 0, Math.PI*2); ctx.fill();
    for(let i=0;i<6;i++){
      const u = (cwBurn + i/6) % 1;
      ctx.globalAlpha = 0.34*(1-u);
      ctx.fillStyle = edge;
      ctx.beginPath();
      ctx.arc(stX + u*w*0.045, stTop - lip - u*h*0.17, hw1*(0.85 + u*2.1), 0, Math.PI*2);
      ctx.fill();
    }
    ctx.globalAlpha = 1;
  }
}

/* ================= desk, wires, monitor =================
   A countdown told as a release. Somebody types at the desk on the left, the
   cables in the middle carry what they type, and the screen on the right fills
   with it line by line as the time is spent. Only the screen carries the
   reading; the typing and the packets just say the clock is going. When the
   time is up the build lands and the screen says so, in English, because that
   is the language the word ships in.

   The code is drawn as stroked marks, not as characters. Real glyphs turn to
   mush at the size this box gets on a phone, and the point is that it reads as
   code from the back of the room, not that anyone can read it. The marks are
   chosen by position, never at random, so the screen holds still between
   frames instead of shimmering. */
/* Six lines of a joke every programmer has lived: the loop fixes one bug and
   makes two, and the last line ships it untested, right before the screen says
   the release is complete. Lines stay under about twenty characters so the
   text still fits when the box is narrow. */
const CODELINES = [
  "while (time > 0) {",
  "  coffee++;",
  "  fixBug();",
  "  bugs += 2;  // oops",
  "}",
  "ship(); // TODO: test"
];
const MONO = "ui-monospace, SFMono-Regular, Menlo, Consolas, monospace";
let cdPhase = 0, cdShow = 0, cdLast = 0;
function codeMark(x, y, s, k){
  const a = s*0.5;
  ctx.beginPath();
  if(k===0){ ctx.moveTo(x+s, y-a); ctx.lineTo(x+s*0.15, y); ctx.lineTo(x+s, y+a); }
  else if(k===1){ ctx.moveTo(x, y-a); ctx.lineTo(x+s*0.85, y); ctx.lineTo(x, y+a); }
  else if(k===2){ ctx.moveTo(x+s*0.8, y-a); ctx.lineTo(x+s*0.3, y-a); ctx.lineTo(x+s*0.3, y+a); ctx.lineTo(x+s*0.8, y+a); }
  else if(k===3){ ctx.moveTo(x+s*0.85, y-a); ctx.lineTo(x+s*0.15, y+a); }
  else if(k===4){ ctx.moveTo(x, y-a*0.45); ctx.lineTo(x+s, y-a*0.45); ctx.moveTo(x, y+a*0.45); ctx.lineTo(x+s, y+a*0.45); }
  else if(k===5){ ctx.moveTo(x+s*0.5, y); ctx.lineTo(x+s*0.52, y); }
  else if(k===6){ ctx.rect(x+s*0.12, y-a*0.75, s*0.76, a*1.5); }
  else if(k===7){ ctx.moveTo(x, y+a*0.3); ctx.quadraticCurveTo(x+s*0.5, y-a*1.15, x+s, y+a*0.3); }
  else if(k===8){ ctx.moveTo(x+s*0.5, y-a); ctx.lineTo(x+s*0.5, y+a); }
  else { ctx.moveTo(x+s*0.5, y-a); ctx.lineTo(x+s*0.5, y+a); ctx.moveTo(x, y); ctx.lineTo(x+s, y); }
  ctx.stroke();
}
function drawCode(w,h,ratio,col,t){
  const dt = cdLast ? Math.min(120, t - cdLast) : 0;
  cdLast = t;
  const done = 1 - clamp(ratio,0,1);
  const fin = !!S.vizDone;
  const dpr = cvs.clientWidth ? w/cvs.clientWidth : 1;
  if(S.running && !fin) cdPhase = (cdPhase + dt*0.00042*motionScale()) % 1;
  cdShow = !fin ? 0 : REDUCED.matches ? 1 : Math.min(1, cdShow + dt/620);

  const ink = PAL.bg;
  const edge = cr(PAL.muted, PAL.bg) >= 3 ? PAL.muted : PAL.txt;
  const ground = h*0.86;
  const solid = function(path){ ctx.fillStyle = edge; path(); ctx.fill(); };

  ctx.lineCap = "round"; ctx.lineJoin = "round";
  ctx.strokeStyle = edge; ctx.lineWidth = Math.max(2, h*0.016);
  ctx.beginPath(); ctx.moveTo(w*0.02, ground); ctx.lineTo(w*0.99, ground); ctx.stroke();

  /* --- the chair --- */
  const seatY = ground - h*0.20, seatT = Math.max(3, h*0.026);
  const chX = w*0.045, chW = Math.min(w*0.070, h*0.25);
  const post = Math.max(3, Math.min(w*0.011, h*0.038));
  solid(function(){ ctx.beginPath(); ctx.rect(chX, seatY, chW, seatT); });
  solid(function(){ ctx.beginPath(); ctx.rect(chX, ground - h*0.46, post, h*0.26); });
  ctx.strokeStyle = edge; ctx.lineWidth = post*0.8;
  ctx.beginPath();
  ctx.moveTo(chX + chW*0.45, seatY + seatT); ctx.lineTo(chX + chW*0.45, ground - h*0.02);
  ctx.moveTo(chX + chW*0.10, ground - h*0.02); ctx.lineTo(chX + chW*0.80, ground - h*0.02);
  ctx.stroke();

  /* --- and whoever is in it. Everything about the figure is measured off the
         head, so it keeps its proportions when the box goes from a wide strip
         on a laptop to something nearly square on a phone --- */
  const headR = Math.min(h*0.058, w*0.021);
  const hipX = chX + chW*0.62, shY = ground - h*0.45, shX = hipX + headR*0.9;
  ctx.strokeStyle = edge; ctx.lineWidth = headR*1.05;
  ctx.beginPath();
  ctx.moveTo(hipX, seatY - headR*0.1);
  ctx.lineTo(chX + chW + headR*1.4, seatY - headR*0.1);
  ctx.lineTo(chX + chW + headR*1.9, ground - headR*0.3);
  ctx.stroke();
  solid(function(){
    ctx.beginPath();
    ctx.moveTo(hipX - headR*0.85, seatY);
    ctx.lineTo(shX - headR*0.75, shY);
    ctx.lineTo(shX + headR*0.75, shY);
    ctx.lineTo(hipX + headR*0.85, seatY);
    ctx.closePath();
  });
  solid(function(){ ctx.beginPath(); ctx.arc(shX + headR*0.30, shY - headR*1.2, headR, 0, Math.PI*2); });

  /* --- the desk and the keyboard --- */
  const dx0 = chX + chW + headR*0.8, dx1 = w*0.315, dw = dx1 - dx0;
  const deskY = ground - h*0.27, deskT = Math.max(3, h*0.028);
  solid(function(){ ctx.beginPath(); ctx.rect(dx0, deskY, dw, deskT); });
  ctx.strokeStyle = edge; ctx.lineWidth = Math.max(2, h*0.018);
  ctx.beginPath();
  ctx.moveTo(dx1 - dw*0.07, deskY + deskT); ctx.lineTo(dx1 - dw*0.07, ground);
  ctx.stroke();
  const kbW = Math.min(dw*0.44, h*0.32), kbT = Math.max(2, h*0.016);
  /* the keyboard sits at the near end of the desk, not the far one: put it
     across the table and the arm has to stretch half again its own length */
  const kbX = dx0 + dw*0.13;
  solid(function(){ ctx.beginPath(); ctx.rect(kbX, deskY - kbT, kbW, kbT); });
  ctx.fillStyle = ink;
  for(let i=0;i<4;i++){
    ctx.fillRect(kbX + kbW*(0.10 + i*0.22), deskY - kbT*0.72, kbW*0.13, Math.max(1, kbT*0.30));
  }
  const bob = (S.running && !fin) ? Math.sin(cdPhase*Math.PI*24)*h*0.014*motionScale() : 0;
  ctx.strokeStyle = edge; ctx.lineWidth = headR*0.78;
  ctx.beginPath();
  ctx.moveTo(shX + headR*0.5, shY + headR*0.6);
  ctx.quadraticCurveTo(shX + headR*1.6, shY + h*0.09, kbX + kbW*0.26, deskY - kbT - headR*0.4 + bob);
  ctx.stroke();

  /* --- the cables --- */
  const ax = dx1, monW = Math.min(w*0.30, h*0.95);
  const monH = Math.min(h*0.52, monW*0.62), monCx = w*0.785;
  const monX = monCx - monW/2, monY = ground - h*0.16 - monH;
  const lw = Math.max(2, h*0.011);
  for(let i=0;i<3;i++){
    const ay = deskY - h*0.005 + i*lw*1.6;
    const by = monY + monH*(0.30 + i*0.20);
    const mx = (ax + monX)/2, my = deskY + h*0.10 + i*h*0.02;
    ctx.strokeStyle = edge; ctx.lineWidth = lw; ctx.globalAlpha = 0.55;
    ctx.beginPath(); ctx.moveTo(ax, ay);
    for(let j=1;j<=24;j++){ const p = bez(j/24, ax, ay, mx, my, monX, by); ctx.lineTo(p[0], p[1]); }
    ctx.stroke(); ctx.globalAlpha = 1;
    if(!S.running || fin) continue;
    ctx.strokeStyle = col; ctx.lineWidth = lw*1.5;
    for(let k=0;k<2;k++){
      const u = (cdPhase*2.6 + i*0.21 + k*0.5) % 1;
      const p = bez(u, ax, ay, mx, my, monX, by);
      const q = bez(Math.max(0, u-0.07), ax, ay, mx, my, monX, by);
      ctx.beginPath(); ctx.moveTo(q[0], q[1]); ctx.lineTo(p[0], p[1]); ctx.stroke();
    }
  }

  /* --- the monitor, which is the reading --- */
  const bezel = Math.max(3, Math.min(w, h)*0.016);
  solid(function(){                                  /* the stand */
    ctx.beginPath();
    ctx.moveTo(monCx - w*0.010, monY + monH); ctx.lineTo(monCx + w*0.010, monY + monH);
    ctx.lineTo(monCx + w*0.020, ground - h*0.02); ctx.lineTo(monCx - w*0.020, ground - h*0.02);
    ctx.closePath();
  });
  solid(function(){ ctx.beginPath(); ctx.rect(monCx - w*0.042, ground - h*0.03, w*0.084, Math.max(3, h*0.030)); });
  solid(function(){ roundRect(monX, monY, monW, monH, bezel*1.1); });
  const sx = monX + bezel, sy = monY + bezel, sw = monW - bezel*2, sh = monH - bezel*2;
  ctx.fillStyle = ink; ctx.fillRect(sx, sy, sw, sh);

  ctx.save();
  ctx.beginPath(); ctx.rect(sx, sy, sw, sh); ctx.clip();
  const LINES = CODELINES.length, lh = sh/(LINES + 1.4), pad = sw*0.07;
  if(fin){
    ctx.save();
    ctx.beginPath(); ctx.rect(sx, sy, sw*cdShow, sh); ctx.clip();
    ctx.fillStyle = col; ctx.textAlign = "center"; ctx.textBaseline = "middle";
    let fs = sh*0.25;
    ctx.font = "700 " + fs.toFixed(1) + "px Inter, system-ui, sans-serif";
    const wide = ctx.measureText("COMPLETE").width;
    if(wide > sw*0.80){
      fs = fs*sw*0.80/wide;
      ctx.font = "700 " + fs.toFixed(1) + "px Inter, system-ui, sans-serif";
    }
    ctx.fillText("RELEASE", sx + sw/2, sy + sh*0.5 - fs*0.58);
    ctx.fillText("COMPLETE", sx + sw/2, sy + sh*0.5 + fs*0.58);
    ctx.restore();
  } else {
    const typed = clamp(done, 0, 1)*LINES;
    const room = sw - pad*2;
    let fp = lh*0.62;
    ctx.font = fp.toFixed(1) + "px " + MONO;
    let widest = 0;
    for(let i=0;i<LINES;i++) widest = Math.max(widest, ctx.measureText(CODELINES[i]).width);
    if(widest > room){ fp = fp*room/widest; ctx.font = fp.toFixed(1) + "px " + MONO; }
    /* Code has to be readable to be worth printing. Under about six device
       pixels a line of it is a smudge, so a box that small gets the marks
       instead: the same texture, and the count of lines still reads. */
    const legible = fp/dpr >= 6;
    ctx.textAlign = "left"; ctx.textBaseline = "middle";
    let caretX = sx + pad, caretY = sy + lh;
    for(let i=0;i<LINES;i++){
      const f = clamp(typed - i, 0, 1);
      if(f <= 0) break;
      const y = sy + lh*(1.0 + i);
      caretY = y;
      if(legible){
        const vis = CODELINES[i].slice(0, Math.ceil(CODELINES[i].length*f));
        const cut = vis.indexOf("//");
        const head = cut < 0 ? vis : vis.slice(0, cut);
        ctx.fillStyle = col; ctx.fillText(head, sx + pad, y);
        if(cut >= 0){                                /* comments step back a tone */
          ctx.fillStyle = edge;
          ctx.fillText(vis.slice(cut), sx + pad + ctx.measureText(head).width, y);
        }
        caretX = sx + pad + ctx.measureText(vis).width;
      } else {
        const gs = lh*0.60, n = 5 + ((i*5 + 3) % 6), indent = (i%3)*gs*1.1;
        const upto = Math.ceil(n*f);
        ctx.lineWidth = Math.max(1, gs*0.15);
        for(let j=0;j<upto;j++){
          const k = (i*7 + j*3 + i*j) % 10;
          ctx.strokeStyle = (k%3===0) ? edge : col;
          codeMark(sx + pad + indent + j*gs*1.25, y, gs, k);
        }
        caretX = sx + pad + indent + upto*gs*1.25;
      }
    }
    if(S.running && Math.floor(t/460)%2===0){
      ctx.strokeStyle = col; ctx.lineWidth = Math.max(1.5, lh*0.10);
      ctx.beginPath();
      ctx.moveTo(caretX, caretY - lh*0.28); ctx.lineTo(caretX, caretY + lh*0.28);
      ctx.stroke();
    }
  }
  ctx.restore();
}

/* A square indicator for identities built on square geometry: an outer frame
   that empties around the perimeter and an inner field that drains from the
   top. It is a geometric shape, not a logo element, and must not be used as
   one — brand guides forbid inventing variants of their marks. */
function drawSquare(w,h,ratio,col){
  const s = Math.min(w,h) - 4, x = (w-s)/2, y = (h-s)/2;
  const t = Math.max(6, s*0.085);
  const a = x + t/2, b = y + t/2, side = s - t;
  const path = function(){
    ctx.beginPath();
    ctx.moveTo(a + side/2, b);
    ctx.lineTo(a + side, b);
    ctx.lineTo(a + side, b + side);
    ctx.lineTo(a, b + side);
    ctx.lineTo(a, b);
    ctx.closePath();
  };
  ctx.lineWidth = t; ctx.lineCap = "butt"; ctx.lineJoin = "miter";
  path(); ctx.strokeStyle = PAL.ringbg; ctx.setLineDash([]); ctx.stroke();
  if(ratio > 0){
    const per = 4*side;
    path();
    ctx.strokeStyle = col;
    ctx.setLineDash([per*ratio, per]);
    ctx.stroke();
    ctx.setLineDash([]);
  }
  const ix = x + t*1.55, iy = y + t*1.55, is = s - t*3.1;
  if(is > 8){
    ctx.lineWidth = Math.max(2, t*0.22);
    ctx.strokeStyle = PAL.ringbg;
    ctx.strokeRect(ix, iy, is, is);
    const fh = is*ratio;
    if(fh > 2){                       /* the core reads as a level line, not a wash */
      ctx.strokeStyle = col;
      ctx.beginPath();
      ctx.moveTo(ix, iy + (is - fh)); ctx.lineTo(ix + is, iy + (is - fh));
      ctx.stroke();
    }
  }
}
/* Three rings at once, the way a chronograph stacks its counters: the outer one
   is this segment, the middle one the round, the inner one the whole session.
   One glance answers "how long left" and "how far through the day". */
function drawChrono(w,h,ratio,col){
  const cx=w/2, cy=h/2, R=Math.min(w,h)/2-2, A=-Math.PI/2;
  const seg = S.program[S.idx];
  const roundV = seg && seg.of ? ((seg.round-1) + (1-ratio))/seg.of
    : (S.program.length ? (S.idx + (1-ratio))/S.program.length : 0);
  const tm = totalMs();
  const sessV = tm>0 ? clamp(doneMs()/tm,0,1) : 0;
  const rings = [
    [R*0.93, R*0.11, ratio, col],
    [R*0.75, R*0.075, 1-clamp(roundV,0,1), PAL.accent2 || PAL.accent],
    [R*0.62, R*0.055, 1-sessV, PAL.muted]
  ];
  rings.forEach(function(r){
    ctx.beginPath(); ctx.arc(cx,cy,r[0],0,Math.PI*2);
    ctx.strokeStyle = PAL.ringbg; ctx.lineWidth = r[1]; ctx.stroke();
    if(r[2] <= 0) return;
    ctx.beginPath(); ctx.arc(cx,cy,r[0],A,A+r[2]*Math.PI*2,false);
    ctx.strokeStyle = r[3]; ctx.lineWidth = r[1]; ctx.lineCap = "round"; ctx.stroke();
  });
}

/* ================= logo =================
   A trainer runs this under someone's brand, so the stage carries a mark. The
   image can be embedded from a file, which lands in its own storage key as a
   data URL and survives a reload offline, or pointed at by a link when the
   room has network. The file wins when both are set. */
const LOGOKEY = "trainingTimer.logo.v1";
const LOGO_MAX = 400*1024;
let logoData = "";
try{ logoData = localStorage.getItem(LOGOKEY) || ""; }catch(e){ logoData = ""; }
function logoSave(){
  try{
    if(logoData) localStorage.setItem(LOGOKEY, logoData); else localStorage.removeItem(LOGOKEY);
    return true;
  }catch(e){ toast("Сховище браузера переповнене, файл не збережено"); return false; }
}
function logoSrc(){ return logoData || String(cfg.logo.url||"").trim(); }
/* Brand files usually ship dark-on-transparent, and most themes here are dark.
   Sample the image once on load and, when the mark would disappear into the
   stage, stand it on the white plate the guides ask for. */
function logoAutoFit(src){
  const im = new Image();
  im.onload = function(){
    try{
      const c = document.createElement("canvas");
      const w = c.width = 64, hh = c.height = Math.max(1, Math.round(64*(im.height||1)/(im.width||1)));
      const g = c.getContext("2d");
      g.drawImage(im, 0, 0, w, hh);
      const d = g.getImageData(0,0,w,hh).data;
      let sum = 0, n = 0;
      for(let i=0;i<d.length;i+=4){
        if(d[i+3] < 40) continue;                       /* transparent pixels say nothing */
        sum += (d[i]*.2126 + d[i+1]*.7152 + d[i+2]*.0722); n++;
      }
      if(!n) return;
      const dark = (sum/n) < 110;
      const bgDark = lum(PAL.bg || "#000000") < .4;
      if(dark && bgDark){                               /* a dark mark on a dark stage */
        /* the plate, not the monochrome switch: it is what the guides ask for
           and it keeps the accent colour the mark was drawn with */
        cfg.logo.plate = true;
        $("#o_logoPlate").checked = cfg.logo.plate;
        persist(); renderLogo();
        toast("Логотип темний на темному тлі, поставив його на білу підкладку");
      }
    }catch(e){}
  };
  im.src = src;
}
function renderLogo(){
  const box = $("#logoBox");
  const src = logoSrc();
  box.className = "logo " + (cfg.logo.pos||"tr") + (cfg.logo.plate ? " plate" : "")
    + ((cfg.logo.on && src) ? " on" : "");
  box.style.setProperty("--logoh", Math.max(16, cfg.logo.size|0) + "px");
  box.style.setProperty("--logorad", String(clamp(isFinite(cfg.logo.rad) ? +cfg.logo.rad : 18, 0, 50)));
  let img = $("#logoImg");
  if(!src){ if(img) img.remove(); }
  else {
    if(!img){
      img = document.createElement("img");
      img.id = "logoImg"; img.alt = ""; img.setAttribute("aria-hidden","true");
      box.appendChild(img);
    }
    if(img.getAttribute("src") !== src) img.setAttribute("src", src);
    /* a mark standing on a white plate has to keep its own colours: painting
       it white as well would leave nothing to look at */
    img.style.filter = (cfg.logo.inv && !cfg.logo.plate) ? "brightness(0) invert(1)" : "";
  }
  const inv = $("#o_logoInv");
  if(inv) inv.disabled = !!cfg.logo.plate;
  const info = $("#logoInfo");
  if(info){
    const kb = logoData.length*0.75/1024;
    info.textContent = logoData ? ("Вбудовано файл, " + (kb<10 ? kb.toFixed(1) : Math.round(kb)) + " КБ")
      : (cfg.logo.url ? "Береться за посиланням" : "Файл не обрано");
  }
}

/* ================= second window =================
   A projector wants the clock and nothing else, while the laptop keeps the
   controls. The second window runs no engine of its own: it mirrors whatever
   the controller posts, so the two can never drift apart. */
/* The mirror is the same document on the same origin, so it takes frames only
   from its own opener. Opened straight from a file the origin is opaque and
   postMessage refuses it as a target, so the wildcard stays for that one case
   and the source check carries the guarantee instead. */
const DISP_ORIGIN = (location.origin && location.origin !== "null") ? location.origin : "*";
let dispWin = null;
function openDisplay(){
  if(EMBED) return;   /* PowerPoint itself is what the projector shows */
  if(dispWin && !dispWin.closed){ dispWin.focus(); return; }
  dispWin = window.open(location.href.split("#")[0]+"#display", "ttDisplay", "width=1280,height=720");
  if(!dispWin) toast("Браузер заблокував нове вікно. Дозвольте спливаючі вікна для цієї сторінки");
  else toast("Друге вікно відкрито. Перетягніть його на проєктор і натисніть F");
}
function pushDisplay(){
  if(!dispWin || dispWin.closed) return;
  try{
    dispWin.postMessage({
      tt:1, run:S.running,
      time:$("#timeText").textContent, name:$("#segName").textContent,
      meta:$("#metaText").textContent, note:$("#segNote").textContent,
      tag:$("#phaseTag").textContent, rot:$("#rotPrompt").textContent,
      msg:$("#stageMsg").classList.contains("on") ? $("#stageMsg").textContent : "",
      phase:$("#stage").style.getPropertyValue("--phase"),
      ratio:S.vizRatio, vphase:S.vizPhase, vdone:S.vizDone,
      nored:cfg.noRed, wcol:cfg.warnColor, dcol:cfg.dangerColor,
      vis:cfg.vis, theme:cfg.theme, lay:$("#stage").className,
      fill:$("#bgFill").style.transform, title:$("#stageTitle").textContent,
      clock:$("#wallClock").textContent, endAt:$("#endAt").textContent,
      sess:$("#sessBar").style.width, font:$("#timeText").style.fontSize,
      logo:{on:cfg.logo.on, src:logoSrc(), size:cfg.logo.size, pos:cfg.logo.pos, inv:cfg.logo.inv,
        plate:cfg.logo.plate, rad:cfg.logo.rad}
    }, DISP_ORIGIN);
  }catch(e){}
}
function initDisplay(){
  document.body.classList.add("present");
  document.body.dataset.display = "1";
  window.addEventListener("message", function(e){
    if(e.source !== window.opener) return;
    if(DISP_ORIGIN !== "*" && e.origin !== DISP_ORIGIN) return;
    const d = e.data; if(!d || d.tt!==1) return;
    /* the mirror runs no engine, but the drawings ask whether the clock is
       going: without this the rotor and the falling sand stood still here */
    S.running = !!d.run;
    /* the second window paints its own traffic light and its own chips out of
       these tokens, so it has to be told the same colour rules */
    if(cfg.noRed!==!!d.nored || cfg.warnColor!==(d.wcol||"") || cfg.dangerColor!==(d.dcol||"")){
      cfg.noRed = !!d.nored; cfg.warnColor = d.wcol||""; cfg.dangerColor = d.dcol||"";
      applyStateColors();
    }
    if(d.theme && d.theme!==cfg.theme){ cfg.theme=d.theme; applyTheme(d.theme); }
    const stage = $("#stage");
    if(d.lay && stage.className!==d.lay) stage.className = d.lay;
    stage.style.setProperty("--phase", d.phase || PAL.accent);
    $("#timeText").textContent = d.time;
    $("#timeText").style.fontSize = d.font || "";
    $("#segName").textContent = d.name;
    $("#metaText").textContent = d.meta;
    $("#segNote").textContent = d.note;
    $("#phaseTag").textContent = d.tag;
    $("#rotPrompt").textContent = d.rot;
    $("#stageTitle").textContent = d.title;
    $("#wallClock").textContent = d.clock;
    $("#endAt").textContent = d.endAt;
    $("#sessBar").style.width = d.sess;
    $("#bgFill").style.transform = d.fill;
    $("#stageMsg").textContent = d.msg || "";
    $("#stageMsg").classList.toggle("on", !!d.msg);
    if(d.logo){
      cfg.logo = {on:d.logo.on, url:d.logo.src, size:d.logo.size, pos:d.logo.pos, inv:d.logo.inv,
        plate:!!d.logo.plate, rad:isFinite(d.logo.rad) ? d.logo.rad : 18};
      logoData = "";                 /* the mirror shows what it is sent, it stores nothing */
      renderLogo();
    }
    if(cfg.vis!==d.vis) cfg.vis = d.vis;
    S.vizRatio = d.ratio; S.vizColor = d.phase || PAL.accent;
    /* the traffic light picks its lamp from the phase and the pit head lights
       its furnace at the finish; neither can be read back from the colour */
    S.vizPhase = d.vphase || "run"; S.vizDone = !!d.vdone;
    if(cfg.vis==="flip") renderFlip(d.time);
    drawViz(performance.now());
  });
}

function drawLiquid(w,h,ratio,col,t){
  const cx=w/2, cy=h/2, R=Math.min(w,h)/2-2;
  ctx.save();
  ctx.beginPath(); ctx.arc(cx,cy,R,0,Math.PI*2); ctx.fillStyle=PAL.ringbg; ctx.fill();
  ctx.clip();
  const lvl = cy + R - ratio*2*R;
  const amp = R*0.045, k = 2*Math.PI/(R*1.05);
  ctx.beginPath(); ctx.moveTo(cx-R, h);
  for(let x=cx-R; x<=cx+R; x+=3){
    const y = lvl + Math.sin((x)*k + t/620)*amp + Math.sin((x)*k*0.6 - t/900)*amp*0.6;
    ctx.lineTo(x,y);
  }
  ctx.lineTo(cx+R,h); ctx.closePath();
  ctx.fillStyle = col; ctx.globalAlpha = .85; ctx.fill();
  ctx.globalAlpha = .35; ctx.translate(0, R*0.05); ctx.fill();
  ctx.restore();
  ctx.beginPath(); ctx.arc(cx,cy,R,0,Math.PI*2);
  ctx.strokeStyle = PAL.line; ctx.lineWidth = Math.max(2,R*0.02); ctx.stroke();
}
function drawSand(w,h,ratio,col,t){
  const cx=w/2, cy=h/2, R=Math.min(w,h)/2-2;
  const hw = R*0.72, y0 = cy-R*0.92, y1 = cy+R*0.92, neck = R*0.05;
  ctx.strokeStyle = PAL.line; ctx.lineWidth = Math.max(2,R*0.025);
  ctx.beginPath();
  ctx.moveTo(cx-hw,y0); ctx.lineTo(cx+hw,y0); ctx.lineTo(cx+neck,cy); ctx.lineTo(cx+hw,y1);
  ctx.lineTo(cx-hw,y1); ctx.lineTo(cx-neck,cy); ctx.closePath(); ctx.stroke();
  ctx.fillStyle = col;
  // top chamber: sand rests near the neck, area fraction = ratio
  const k = Math.sqrt(clamp(ratio,0,1));
  const ty = cy - k*(cy-y0);
  const tw = neck + (hw-neck)*k;
  ctx.beginPath(); ctx.moveTo(cx-tw,ty); ctx.lineTo(cx+tw,ty); ctx.lineTo(cx+neck,cy); ctx.lineTo(cx-neck,cy); ctx.closePath(); ctx.fill();
  // bottom chamber mound
  const f = 1-clamp(ratio,0,1);
  const k2 = Math.sqrt(1-f);
  const by = cy + k2*(y1-cy);
  const bw = neck + (hw-neck)*k2;
  ctx.beginPath(); ctx.moveTo(cx-bw,by); ctx.lineTo(cx+bw,by); ctx.lineTo(cx+hw,y1); ctx.lineTo(cx-hw,y1); ctx.closePath(); ctx.fill();
  // stream and grains
  if(S.running && ratio>0){
    ctx.fillRect(cx-R*0.012, cy, R*0.024, by-cy);
    for(let i=0;i<7;i++){
      const p = ((t/380)+i/7)%1;
      const y = cy + p*(by-cy);
      const x = cx + Math.sin(t/120+i)*R*0.02;
      ctx.beginPath(); ctx.arc(x,y,R*0.014,0,Math.PI*2); ctx.fill();
    }
  }
}
function drawCandle(w,h,ratio,col,t){
  const cx=w/2, cy=h/2, R=Math.min(w,h)/2-2;
  const bw = R*0.44, base = cy+R*0.86, full = R*1.45;
  const ch = Math.max(R*0.06, full*clamp(ratio,0,1));
  const top = base-ch;
  // holder
  ctx.fillStyle = PAL.line;
  ctx.beginPath(); ctx.moveTo(cx-bw*1.5,base+R*0.1); ctx.lineTo(cx+bw*1.5,base+R*0.1);
  ctx.lineTo(cx+bw*0.9,base); ctx.lineTo(cx-bw*0.9,base); ctx.closePath(); ctx.fill();
  // body
  ctx.fillStyle = col; ctx.globalAlpha = .9;
  ctx.fillRect(cx-bw, top, bw*2, ch);
  ctx.globalAlpha = 1;
  // drips
  ctx.fillStyle = col;
  for(let i=0;i<3;i++){
    const dx = cx-bw + bw*2*(0.2+i*0.3);
    const dl = R*0.10 + (i*7 % 5)*R*0.03 + Math.abs(Math.sin(i*2.1))*R*0.08;
    ctx.beginPath(); ctx.moveTo(dx-R*0.035, top); ctx.lineTo(dx+R*0.035, top);
    ctx.lineTo(dx+R*0.035, top+dl); ctx.arc(dx, top+dl, R*0.035, 0, Math.PI); ctx.closePath(); ctx.fill();
  }
  // wick and flame
  const fl = 1 + Math.sin(t/90)*0.10 + Math.sin(t/37)*0.05;
  ctx.strokeStyle = PAL.txt; ctx.lineWidth = Math.max(1.5,R*0.014);
  ctx.beginPath(); ctx.moveTo(cx, top); ctx.lineTo(cx, top-R*0.07); ctx.stroke();
  const fh = R*0.30*fl, fw = R*0.11*fl;
  const grad = ctx.createRadialGradient(cx, top-R*0.07-fh*0.45, fw*0.15, cx, top-R*0.07-fh*0.4, fh*0.8);
  grad.addColorStop(0, "#fff8d0"); grad.addColorStop(.45, "#ffc84a"); grad.addColorStop(1, "rgba(255,120,20,0)");
  ctx.fillStyle = grad;
  ctx.beginPath();
  ctx.moveTo(cx, top-R*0.07-fh);
  ctx.quadraticCurveTo(cx+fw, top-R*0.07-fh*0.35, cx, top-R*0.07);
  ctx.quadraticCurveTo(cx-fw, top-R*0.07-fh*0.35, cx, top-R*0.07-fh);
  ctx.fill();
}
function drawDots(w,h,ratio,col){
  const n = 10, R = Math.min(w,h), gap = R/(n+1), rr = gap*0.30;
  const lit = Math.ceil(clamp(ratio,0,1)*n*n);
  const ox = (w-(n-1)*gap)/2, oy = (h-(n-1)*gap)/2;
  for(let i=0;i<n*n;i++){
    const r0 = Math.floor(i/n), c0 = i%n;
    const x = ox + c0*gap, y = oy + r0*gap;
    const on = (n*n-i) <= lit;
    ctx.beginPath(); ctx.arc(x,y,rr,0,Math.PI*2);
    ctx.fillStyle = on ? col : PAL.ringbg;
    ctx.globalAlpha = on ? 1 : .7; ctx.fill(); ctx.globalAlpha = 1;
  }
}
function drawTraffic(w,h,ratio,col){
  const R = Math.min(w,h);
  const bw = R*0.42, bh = R*0.96, x = (w-bw)/2, y = (h-bh)/2;
  const rr = bw*0.16;
  ctx.fillStyle = PAL.panel; ctx.strokeStyle = PAL.line; ctx.lineWidth = Math.max(2,R*0.012);
  roundRect(x,y,bw,bh,rr); ctx.fill(); ctx.stroke();
  const cx = w/2, rad = bw*0.32, step = bh/3.4;
  const cy0 = y + step*0.62;
  const on = S.running || S.vizRatio<1;
  /* The lit lamp follows the phase the engine reports, not the phase guessed
     back from the colour: a palette is free to give run, warn and finish the
     same value, and the DTEK guide does exactly that. */
  const active = !on ? -1 : (S.vizPhase==="danger" ? 2 : S.vizPhase==="warn" ? 1 : 0);
  [[PAL.ok,0],[PAL.warn,1],[PAL.danger,2]].forEach(function(p,i){
    const yy = cy0 + i*step;
    ctx.beginPath(); ctx.arc(cx,yy,rad,0,Math.PI*2);
    if(active===i){
      ctx.fillStyle = p[0]; ctx.shadowColor = p[0]; ctx.shadowBlur = rad*0.9;
    } else { ctx.fillStyle = PAL.ringbg; ctx.shadowBlur = 0; }
    ctx.fill(); ctx.shadowBlur = 0;
  });
}
function roundRect(x,y,w,h,r){
  ctx.beginPath();
  ctx.moveTo(x+r,y); ctx.lineTo(x+w-r,y); ctx.quadraticCurveTo(x+w,y,x+w,y+r);
  ctx.lineTo(x+w,y+h-r); ctx.quadraticCurveTo(x+w,y+h,x+w-r,y+h);
  ctx.lineTo(x+r,y+h); ctx.quadraticCurveTo(x,y+h,x,y+h-r);
  ctx.lineTo(x,y+r); ctx.quadraticCurveTo(x,y,x+r,y); ctx.closePath();
}
/* Only four of the visuals actually move. The others are repainted by render(),
   twenty times a second while the clock runs and once a second while it waits,
   so redrawing all of them on every frame spent a laptop's battery for a whole
   training day and changed nothing on screen. */
const ANIMATED = ["energy","sand","liquid","candle","coal","code"];
/* Two scenes have something to say after the clock stops: the furnace only
   lights when the time is spent, and the release only lands then. They keep
   their frames through the finish while every other one goes still. A finished
   timer can sit on a projector for an hour, though, so they run for half a
   minute and then hold the last frame rather than spending a battery on an
   empty room. */
const FINISH_LIVE = ["coal","code"];
let cwEnd = 0;
function vizMoving(){
  if(!S.vizDone || FINISH_LIVE.indexOf(cfg.vis) < 0) cwEnd = 0;
  else if(!cwEnd) cwEnd = performance.now();
  if(ANIMATED.indexOf(cfg.vis) < 0) return false;
  return S.running || (cwEnd > 0 && performance.now() - cwEnd < 30000);
}
function vizLoop(t){
  vizRaf = 0;
  if(vizMoving()) drawViz(t||0);
  else if(S.vizDirty){ S.vizDirty = false; drawViz(t||0); }
  if(cfg.noise && analyser){
    const lv = noiseLevel();
    const bar = $("#noiseBar");
    bar.style.transform = "scaleY(" + lv.toFixed(3) + ")";
    $("#noiseBox").classList.toggle("loud", lv > cfg.noiseTh);
  }
  /* Цикл кадрів іде лише поки є що малювати: PowerPoint не ховає блок, що зійшов зі слайда,
     тож безперервний requestAnimationFrame тримав би процесор зайнятим у кожному блоці з таймером. */
  if(vizMoving() || S.vizDirty || (cfg.noise && analyser)) vizRaf = requestAnimationFrame(vizLoop);
}
var vizRaf = 0;
function vizKick(){ if(!vizRaf) vizRaf = requestAnimationFrame(vizLoop); }

/* ================= flip clock ================= */
/* A Solari board does not snap to the new character, it runs the reel through
   the ones between. Three intermediate flaps are enough to read as mechanical
   without holding the number back.

   Three rules keep the reel honest. The digit it is travelling towards is kept
   on the element, so the twenty-a-second render pass never restarts a roll that
   is already on its way to that same digit. Every roll carries a deadline in
   real time. And the reel only runs at all when the pace of this function says
   it can finish: a hidden window has its timers clamped to one second, which is
   slower than the countdown itself, so a reel driven by its own timer there
   either chases the clock forever or never advances. That is what turned the
   board to nonsense in a projector window left in the background. When the pace
   is too slow the digit simply appears, which is what a reader needs anyway. */
const ROLL_MS = 260, FLIP_LIVE_MS = 250;
function rollDigit(el, to, live){
  if(el._to === to) return;               /* already on its way there */
  clearInterval(el._roll); el._roll = 0;
  el._to = to;
  const flap = function(ch){
    el.textContent = ch;
    el.classList.remove("turn"); void el.offsetWidth; el.classList.add("turn");
  };
  const snap = function(){ clearInterval(el._roll); el._roll = 0; flap(to); };
  const from = parseInt(el.textContent,10), t = parseInt(to,10);
  if(!live || isNaN(from) || isNaN(t) || REDUCED.matches){ snap(); return; }
  const t0 = performance.now();
  let cur = from, steps = 0;
  el._roll = setInterval(function(){
    if(performance.now() - t0 >= ROLL_MS){ snap(); return; }
    cur = (cur + 9) % 10;                 /* reels run down, like a falling count */
    steps++;
    if(steps >= 3 || cur === t){ snap(); return; }
    flap(String(cur));
  }, 45);
}
let flipAt = 0;
function renderFlip(str){
  const box = $("#flipClock");
  const now = performance.now();
  const live = flipAt > 0 && (now - flipAt) < FLIP_LIVE_MS;
  flipAt = now;
  const sig = str.length+"|"+str.replace(/\d/g,"9");
  if(box.dataset.sig !== sig){
    box.innerHTML = "";
    for(const ch of str){
      const el = document.createElement("span");
      if(/\d/.test(ch)){ el.className = "fd"; el.textContent = ch; el._to = ch; }
      else { el.className = "fsep"; el.textContent = ch; }
      box.appendChild(el);
    }
    box.dataset.sig = sig;
    return;
  }
  const kids = box.children;
  for(let i=0;i<str.length;i++){
    const el = kids[i]; if(!el) continue;
    if(el.className.indexOf("fd")>=0){
      if(el._to !== str[i]) rollDigit(el, str[i], live);
    } else if(el.textContent !== str[i]) el.textContent = str[i];
  }
}

/* ================= render ================= */
const CIRC = 2*Math.PI*45;
function render(){
  vizKick();
  if(cfg.mode==="clock"){ renderClock(); return; }
  const stage = $("#stage");
  const lay = STACKED.indexOf(cfg.vis)>=0 ? "lay-stack" : "lay-over";
  stage.className = "stage vis-"+cfg.vis+" "+lay;
  let ratio = 0, phaseCol = PAL.accent, vphase = "run", vdone = false;

  if(cfg.mode==="stopwatch"){
    const e = elapsedSW();
    const txt = fmt(e, {tenths: cfg.tenths && e<600000});
    $("#timeText").textContent = txt;
    if(cfg.vis==="flip") renderFlip(txt);
    $("#segName").textContent = S.running ? "Секундомір" : (e>0 ? "Пауза" : "Готово до старту");
    $("#metaText").textContent = S.laps.length ? ("Кіл: "+S.laps.length+"  ·  останнє "+fmt(S.laps[S.laps.length-1].d)) : "";
    $("#segNote").textContent = ""; $("#rotPrompt").textContent = ""; $("#phaseTag").textContent = "";
    ratio = (e % 60000)/60000;
    $("#sessBar").style.width = (ratio*100)+"%";
    $("#endAt").textContent = "";
  } else {
    const seg = S.program[S.idx];
    if(!seg){ compile(); return; }
    const l = left();
    const over = l < 0;
    ratio = clamp(l/seg.ms, 0, 1);
    const phase = over ? "danger"
      : (cfg.danger>0 && l<=cfg.danger*1000) ? "danger"
      : (cfg.warn>0 && l<=cfg.warn*1000 && seg.ms>cfg.warn*1000) ? "warn" : "run";
    phaseCol = phase==="danger" ? PAL.danger : phase==="warn" ? PAL.warn : (seg.color ? fit(segColor(seg.color), PAL.bg, 3.2) : PAL.accent);
    vphase = phase;
    /* the drawings that have something to say about the finish itself ask here,
       not at ratio zero: a stopwatch passes through zero every minute */
    vdone = l <= 0;
    if(phase!=="run") stage.classList.add(phase);
    /* the state must survive a colour-blind viewer and a washed-out projector */
    const prot = cfg.protect>0 && !over && (seg.ms-l <= cfg.protect*1000 || l <= cfg.protect*1000);
    stage.classList.toggle("protect", !!prot && S.running);
    $("#phaseTag").textContent = (S.ended && !S.running) ? ""
      : over ? "Овертайм"
      : phase==="danger" ? "Останні секунди"
      : prot ? "Захищений час, без питань"
      : phase==="warn" ? "Час спливає" : "";

    const rot = cfg.mode==="interval" && cfg.iv.rotation;
    const txt = fmt(l, {ceil:!over, tenths: cfg.tenths && !over && l<10000 && S.running});
    $("#timeText").textContent = txt;
    if(cfg.vis==="flip") renderFlip(txt);
    $("#timeText").style.visibility = (rot && seg.kind==="work" && S.running) ? "hidden" : "";
    $("#rotPrompt").textContent = (rot && seg.kind==="rest") ? seg.name : "";
    $("#segName").textContent = over ? (seg.name+" · овертайм") : (S.ended && !S.running ? "Завершено" : seg.name);
    $("#segNote").textContent = seg.note || "";

    const parts = [];
    if(S.program.length>1) parts.push("Етап "+(S.idx+1)+"/"+S.program.length);
    if(seg.round && seg.of) parts.push("Раунд "+seg.round+"/"+seg.of);
    const tm = totalMs();
    if(tm>0) parts.push(fmt(doneMs())+" / "+fmt(tm));
    $("#metaText").textContent = parts.join("  ·  ");

    /* One spoken status per meaningful change. A live region that fires every
       second is unusable, so the key is the segment plus the phase. */
    const ak = S.idx+"|"+phase+"|"+(S.running?1:0)+"|"+(S.ended?1:0);
    if(ak!==annKey){
      annKey = ak;
      announce(S.ended ? "Завершено"
        : !S.running ? ("Пауза. "+seg.name+", лишилось "+human(l))
        : over ? "Овертайм"
        : phase==="danger" ? "Останні секунди"
        : phase==="warn" ? ("Час спливає, лишилось "+human(l))
        : (seg.name+", "+human(l)));
    }
    $("#sessBar").style.width = (tm>0 ? clamp(doneMs()/tm,0,1)*100 : 0)+"%";

    if(cfg.showEnd){
      const r = restMs();
      $("#endAt").textContent = r>0 ? ("Фініш ≈ "+hhmm(new Date(Date.now()+r))) : "";
    } else $("#endAt").textContent = "";
  }

  stage.style.setProperty("--phase", phaseCol);
  S.vizRatio = ratio; S.vizColor = phaseCol; S.vizPhase = vphase; S.vizDone = vdone;

  $("#bgFill").style.transform = "scaleY(" + (cfg.bgFill ? (1-ratio) : 0).toFixed(4) + ")";

  const now = new Date();
  $("#wallClock").textContent = cfg.showClock ? (pad(now.getHours())+":"+pad(now.getMinutes())+":"+pad(now.getSeconds())) : "";

  $("#prg").style.strokeDasharray = CIRC;
  $("#prg").style.strokeDashoffset = CIRC*(1-ratio);
  $("#bar").style.width = (ratio*100)+"%";

  // keep long values (hours, tenths) inside the ring
  const len = $("#timeText").textContent.length;
  /* the disc leaves a hole of .58 and the chronograph's inner counter closes at
     .62, so digits set inside them need a smaller basis than a bare ring does */
  const hole = (cfg.vis==="disc" || cfg.vis==="chrono") ? .8 : 1;
  $("#timeText").style.fontSize = "calc(var(--d)*" + ((len>7 ? .155 : len>5 ? .19 : .235)*hole).toFixed(3) + ")";

  const b = $("#btnStart");
  b.querySelector("span").textContent = S.running ? "Пауза" : (S.ended ? "Спочатку" : "Старт");
  b.querySelector("svg").innerHTML = S.running ? '<path d="M8 4v16M16 4v16"/>' : '<path d="M6 4l14 8-14 8z"/>';
  $("#tbSound").classList.toggle("off", !cfg.sound);
  $("#tbScore").classList.toggle("on", cfg.showScore);
  $("#tbSound").setAttribute("aria-pressed", cfg.sound ? "true" : "false");
  $("#tbScore").setAttribute("aria-pressed", cfg.showScore ? "true" : "false");
  $("#scoreBox").classList.toggle("on", cfg.showScore);
  $("#btnPrev").disabled = cfg.mode==="stopwatch";   /* nothing to step back to */
  stepLabels();

  // rAF can be frozen when the window is occluded, so draw from here too
  drawViz(performance.now());
  pushDisplay();
}

function stepLabels(){
  const s = Math.max(1, cfg.step|0);
  const txt = s%60===0 ? (s/60)+" хв" : s+"с";
  $("#btnPlus").textContent = "+"+txt;
  $("#btnMinus").textContent = "−"+txt;
  $("#btnPlus").title = "Додати "+txt;
  $("#btnMinus").title = "Відняти "+txt;
}
function renderTimeline(){
  const tl = $("#timeline");
  if(cfg.mode==="stopwatch" || S.program.length<2){ tl.innerHTML=""; return; }
  tl.innerHTML = "";
  S.program.forEach(function(seg,i){
    /* a real button, not a clickable div: the strip is how you jump between
       stages, and it has to be reachable from the keyboard like everything else */
    const c = document.createElement("button");
    c.type = "button";
    c.className = "chip" + (i===S.idx ? " cur" : i<S.idx ? " done" : "");
    c.innerHTML = (seg.color ? '<span class="pip" aria-hidden="true" style="background:'+esc(segColor(seg.color))+'"></span>' : "")+
      '<span>'+esc(seg.name)+'</span><span class="t">'+fmtShort(seg.ms)+'</span>';
    if(i===S.idx) c.setAttribute("aria-current","step");
    c.onclick = function(){ jump(i); };
    tl.appendChild(c);
  });
  const cur = tl.children[S.idx];
  if(cur && cur.scrollIntoView) cur.scrollIntoView({block:"nearest",inline:"center",behavior:"smooth"});
}
function esc(s){ return String(s).replace(/[&<>"]/g, function(c){ return ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"})[c]; }); }

function renderScore(){
  const rows = cfg.roster.map(function(p,i){ return {p:p, i:i}; })
    .sort(function(a,b){ return (b.p.score||0)-(a.p.score||0); });
  const max = rows.length ? (rows[0].p.score||0) : 0;
  $("#scoreRows").innerHTML = rows.map(function(r){
    const out = Q.out.indexOf(r.i)>=0 ? " out" : "";
    const lead = (r.p.score||0)===max && max>0 ? " lead" : "";
    return '<div class="r'+lead+out+'"><span>'+esc(r.p.name)+'</span><b>'+esc(r.p.score||0)+'</b></div>';
  }).join("") || '<div class="r"><span style="color:var(--muted)">Список порожній</span></div>';
}
function renderLaps(){
  const box = $("#lapList"); box.innerHTML = "";
  if(!S.laps.length) box.innerHTML = '<p class="hint">Кіл ще немає. Клавіша L фіксує коло, поки секундомір іде.</p>';
  S.laps.slice().reverse().forEach(function(lp,i){
    const d = document.createElement("div");
    d.className = "logrow";
    d.innerHTML = '<span class="d">'+(S.laps.length-i)+'</span><span class="n">'+esc(lp.name)+'</span>'+
                  '<span class="dl">'+fmt(lp.d)+'</span><span class="d">'+fmt(lp.t)+'</span>';
    box.appendChild(d);
  });
}
function renderLog(){
  logSummary();
  const box = $("#logList"); box.innerHTML = "";
  if(!S.log.length){ box.innerHTML = '<p class="hint">Поки порожньо. Записи з\'являться, щойно перший етап завершиться.</p>'; return; }
  S.log.forEach(function(e){
    const diff = e.actual - e.planned;
    const cls = Math.abs(diff) < 3000 ? "" : (diff>0 ? " over" : " under");
    const sign = diff>0 ? "+" : "−";
    const d = document.createElement("div");
    d.className = "logrow";
    d.innerHTML = '<span class="n">'+esc(e.name)+'</span><span class="d">'+fmt(e.planned)+'</span>'+
                  '<span class="dl'+cls+'">'+sign+fmt(Math.abs(diff))+'</span>';
    box.appendChild(d);
  });
}

/* ================= editors ================= */
let dragI = null;
function renderProgram(){
  const box = $("#pgList"); box.innerHTML="";
  if(!cfg.pg.items.length) box.innerHTML = '<p class="hint">Етапів ще немає. Додайте перший — він одразу зʼявиться на таймлайні під таймером.</p>';
  cfg.pg.items.forEach(function(it,i){
    const row = document.createElement("div");
    row.className="item col"; row.draggable=true;
    row.innerHTML =
      '<div class="r1">'+
        '<span class="grab" aria-hidden="true">::</span>'+
        '<input class="inp" style="flex:1" value="'+esc(it.name)+'" data-k="name" aria-label="Назва етапу '+(i+1)+'">'+
        '<span class="durf compact" data-dur="1"></span>'+
        '<span class="acts">'+
          '<button class="x mv" data-mv="-1" aria-label="Пересунути етап '+(i+1)+' вище"'+(i===0?" disabled":"")+'>&uarr;</button>'+
          '<button class="x mv" data-mv="1" aria-label="Пересунути етап '+(i+1)+' нижче"'+(i===cfg.pg.items.length-1?" disabled":"")+'>&darr;</button>'+
          '<button class="x" aria-label="Видалити етап '+(i+1)+'">&times;</button>'+
        '</span>'+
      '</div>'+
      '<div class="r1">'+
        '<span class="atwrap"></span>'+
        '<select class="inp col-sel" data-k="color" aria-label="Колір етапу '+(i+1)+'">'+colorOptions(it.color)+'</select>'+
        '<input class="inp note" placeholder="нотатка для екрана…" value="'+esc(it.note||"")+'" data-k="note" aria-label="Нотатка етапу '+(i+1)+'">'+
      '</div>';
    $$("input,select", row).forEach(function(el){
      const k = el.dataset.k; if(!k) return;
      el.oninput = el.onchange = function(){
        cfg.pg.items[i][k] = el.value;
        persist(true); updatePgTotal();
        if(!S.running && cfg.mode==="program") compile();
      };
    });
    durInputs(row.querySelector("[data-dur]"),
      function(){ return cfg.pg.items[i].dur; },
      function(v){
        cfg.pg.items[i].dur = v; persist(true); updatePgTotal();
        if(!S.running && cfg.mode==="program") compile();
      },
      {unit:"min", at:true, compact:true, label:"Тривалість етапу", atHost:row.querySelector(".atwrap")});
    tameInputs(row);
    $$(".mv", row).forEach(function(b){
      b.onclick = function(){
        const to = i + (+b.dataset.mv);
        if(to < 0 || to >= cfg.pg.items.length) return;
        const moved = cfg.pg.items.splice(i,1)[0];
        cfg.pg.items.splice(to,0,moved);
        renderProgram(); persist(); if(!S.running) compile();
        const again = $$(".item .mv")[to*2 + (+b.dataset.mv > 0 ? 1 : 0)];
        if(again && !again.disabled) again.focus();
      };
    });
    $$(".x:not(.mv)", row)[0].onclick = function(){ cfg.pg.items.splice(i,1); renderProgram(); persist(); if(!S.running) compile(); };
    row.addEventListener("dragstart", function(e){ dragI=i; row.classList.add("dragging"); e.dataTransfer.effectAllowed="move"; });
    row.addEventListener("dragend", function(){ row.classList.remove("dragging"); $$(".item.over").forEach(function(x){ x.classList.remove("over"); }); });
    row.addEventListener("dragover", function(e){ e.preventDefault(); row.classList.add("over"); });
    row.addEventListener("dragleave", function(){ row.classList.remove("over"); });
    row.addEventListener("drop", function(e){
      e.preventDefault(); row.classList.remove("over");
      if(dragI===null || dragI===i) return;
      const moved = cfg.pg.items.splice(dragI,1)[0];
      cfg.pg.items.splice(i,0,moved); dragI=null;
      renderProgram(); persist(); if(!S.running) compile();
    });
    box.appendChild(row);
  });
  updatePgTotal();
}
function updatePgTotal(){
  const t = cfg.pg.items.reduce(function(s,x){ return s+parseDur(x.dur,"min"); },0);
  $("#pgTotal").textContent = fmt(t);
}
function renderRoster(){
  const box = $("#rsList"); box.innerHTML="";
  if(!cfg.roster.length) box.innerHTML = '<p class="hint">Список порожній. Додайте людей або команди — і запрацюють черга виступів, табло, жеребкування та поділ на команди.</p>';
  cfg.roster.forEach(function(p,i){
    const row = document.createElement("div"); row.className="item";
    row.innerHTML =
      (i<9 ? '<span class="key" aria-hidden="true">'+(i+1)+'</span>' : '<span class="key" aria-hidden="true">·</span>')+
      '<input class="inp" style="flex:1" value="'+esc(p.name)+'" aria-label="Імʼя учасника '+(i+1)+'">'+
      '<button class="btn sm" data-d="-1" aria-label="Мінус бал: '+esc(p.name||("учасник "+(i+1)))+'">−</button>'+
      '<input class="inp num" value="'+esc(p.score||0)+'" aria-label="Бали: '+esc(p.name||("учасник "+(i+1)))+'">'+
      '<button class="btn sm" data-d="1" aria-label="Плюс бал: '+esc(p.name||("учасник "+(i+1)))+'">+</button>'+
      '<button class="x" aria-label="Видалити: '+esc(p.name||("учасник "+(i+1)))+'">&times;</button>';
    const ins = row.querySelectorAll("input");
    ins[0].oninput = function(){ cfg.roster[i].name = ins[0].value; persist(true); renderScore(); };
    ins[1].oninput = function(){ cfg.roster[i].score = parseInt(ins[1].value)||0; persist(true); renderScore(); };
    $$("[data-d]", row).forEach(function(b){
      b.onclick = function(){
        cfg.roster[i].score = (cfg.roster[i].score||0) + parseInt(b.dataset.d);
        ins[1].value = cfg.roster[i].score; persist(); renderScore();
      };
    });
    tameInputs(row);
    row.querySelector(".x").onclick = function(){ cfg.roster.splice(i,1); renderRoster(); persist(); if(cfg.mode==="turns" && !S.running) compile(); };
    box.appendChild(row);
  });
  renderScore(); updatePickState(); updateBuzzState();
}
function updatePickState(){
  const all = cfg.roster.filter(function(p){ return p.name; }).length;
  $("#pickState").textContent = cfg.noRepeat
    ? ("У списку "+all+", вже витягли "+S.picked.length+".")
    : ("У списку "+all+", повтори дозволені.");
}
function updateBuzzState(){
  $("#buzzState").textContent = cfg.buzzOn
    ? ("Питання "+Q.question+", вибуло "+Q.out.length)
    : "Вимкнено";
}

/* ================= option binding ================= */
function bind(id, path, type){
  const el = $(id); if(!el) return;
  const get = function(){ return path.split(".").reduce(function(o,k){ return o[k]; }, cfg); };
  const set = function(v){
    const ks = path.split("."); let o = cfg;
    for(let i=0;i<ks.length-1;i++) o = o[ks[i]];
    o[ks[ks.length-1]] = v;
  };
  if(type==="check"){ el.checked = !!get(); el.onchange = function(){ set(el.checked); onCfg(path); }; }
  else if(type==="num"){ el.value = get(); el.oninput = function(){ set(parseFloat(el.value)||0); onCfg(path); }; }
  else { el.value = get(); el.oninput = el.onchange = function(){ set(el.value); onCfg(path); }; }
}
const RECOMPILE = ["ck.bank","ck.periods","ck.glass","cd.dur","cd.label","iv.prep","iv.work","iv.rest","iv.rounds","iv.wlabel","iv.rlabel","iv.restlast","iv.wcolor","iv.rcolor","tn.dur","tn.order"];
function onCfg(path){
  persist(true);
  if(path==="theme"){ applyTheme(cfg.theme); }
  if(path==="noRepeat") updatePickState();
  if(path==="buzzOn"){ Q.out=[]; Q.locked=null; updateBuzzState(); renderScore(); }
  if(path==="musicOn"){ if(cfg.musicOn && S.running) musicStart(); else musicPause(); }
  if(path==="amb"){ ambStop(); if(S.running) setTimeout(ambStart,300); }
  if(path.indexOf("logo")===0) renderLogo();
  if(path==="warnColor" || path==="dangerColor" || path==="noRed"){ applyStateColors(); render(); renderTimeline(); }
  if(path==="msgHold" || path==="bells") persist(true);
  if(path==="musicVol"){ $("#bgm").volume = cfg.musicVol; }
  if(path==="musicLoop"){ $("#bgm").loop = !!cfg.musicLoop; }
  if(path==="noise"){ cfg.noise ? noiseStart() : noiseStop(); }
  if(RECOMPILE.indexOf(path)>=0 && !S.running) compile();
  render(); renderTimeline();
}
let saveT=0;
function persist(debounce){
  clearTimeout(saveT);
  if(debounce) saveT = setTimeout(saveStore, 400); else saveStore();
}
function bindAll(){
  bindDur("#cd_dur","cd.dur","min",true); bind("#cd_label","cd.label");
  bindDur("#iv_prep","iv.prep","sec"); bindDur("#iv_work","iv.work","sec"); bindDur("#iv_rest","iv.rest","sec");
  bind("#iv_rounds","iv.rounds","num"); bind("#iv_wlabel","iv.wlabel"); bind("#iv_rlabel","iv.rlabel");
  bind("#iv_restlast","iv.restlast","check"); bind("#iv_rotation","iv.rotation","check");
  bind("#iv_wcolor","iv.wcolor"); bind("#iv_rcolor","iv.rcolor");
  bindDur("#tn_dur","tn.dur","min"); bind("#tn_order","tn.order");
  bindDur("#ck_bank","ck.bank","min"); bind("#ck_inc","ck.inc","num"); bind("#ck_delay","ck.delay","num");
  bind("#ck_periods","ck.periods","num"); bind("#ck_pdur","ck.pdur","num"); bind("#ck_glass","ck.glass","check");
  bind("#o_autoNext","autoNext","check"); bind("#o_overtime","overtime","check"); bind("#o_loop","loop","check");
  bind("#o_warn","warn","num"); bind("#o_danger","danger","num");
  bind("#o_bells","bells"); bind("#o_bellsOver","bellsOver","num");
  bind("#o_warnColor","warnColor"); bind("#o_dangerColor","dangerColor"); bind("#o_noRed","noRed","check");
  bind("#o_protect","protect","num"); bind("#o_step","step","num"); bind("#o_msgHold","msgHold","num");
  bind("#o_amb","amb"); bind("#o_ambVol","ambVol","num");
  bind("#o_logoOn","logo.on","check"); bind("#o_logoUrl","logo.url");
  bind("#o_logoPos","logo.pos"); bind("#o_logoSize","logo.size","num"); bind("#o_logoInv","logo.inv","check");
  bind("#o_logoPlate","logo.plate","check"); bind("#o_logoRad","logo.rad","num");
  bind("#o_theme","theme"); bind("#o_vis","vis"); bind("#o_bgFill","bgFill","check");
  bind("#o_tenths","tenths","check"); bind("#o_showScore","showScore","check");
  bind("#o_showClock","showClock","check"); bind("#o_showEnd","showEnd","check");
  bind("#o_sound","sound","check"); bind("#o_volume","volume","num"); bind("#o_pack","pack");
  bind("#o_tickFrom","tickFrom","num"); bind("#o_flash","flash","check");
  bind("#o_voice","voice","check"); bind("#o_wake","wake","check");
  bind("#o_musicOn","musicOn","check"); bind("#o_musicLoop","musicLoop","check"); bind("#o_musicVol","musicVol","num");
  bind("#o_noise","noise","check"); bind("#o_noiseTh","noiseTh","num");
  bind("#o_buzzOn","buzzOn","check"); bind("#o_buzzAnswer","buzzAnswer","num");
  bind("#o_noRepeat","noRepeat","check"); bind("#tmCount","teamCount","num");
}

/* ================= modes ================= */
const MODE_TITLE = {countdown:"Зворотний відлік", interval:"Інтервали і раунди", program:"Програма етапів", turns:"Черга виступів", clock:"Шаховий годинник", stopwatch:"Секундомір"};
function setMode(m){
  cfg.mode = m;
  document.body.dataset.mode = m;
  $$("#modes button").forEach(function(b){
    const on = b.dataset.mode===m;
    b.classList.toggle("on", on); b.setAttribute("aria-pressed", on ? "true" : "false");
  });
  $$("[data-cfg]").forEach(function(d){ d.style.display = d.dataset.cfg===m ? "" : "none"; });
  $("#modeTitle").textContent = MODE_TITLE[m];
  if(m==="stopwatch"){ S.running=false; S.swBase=0; S.program=[]; S.log=[]; renderLaps(); renderLog(); }
  else compile();
  persist(); render(); renderTimeline();
}

/* ================= duration fields =================
   Every duration is entered as minutes and seconds. A bare number used to mean
   minutes in one mode and seconds in another, which is how a "25 / 5" pomodoro
   preset ended up running for twenty-five seconds. Stored values are now always
   "M:SS", so nothing depends on which field it came from. The same pair of
   boxes becomes hours and minutes when the field is switched to a wall-clock
   target, which is how "@13:00" stays reachable without a second control. */
function durParts(v, defUnit){
  const s = String(v==null ? "" : v).trim();
  if(s[0]==="@"){
    const p = s.slice(1).split(":");
    return {at:true, a:clamp(parseInt(p[0],10)||0,0,23), b:clamp(parseInt(p[1],10)||0,0,59)};
  }
  const ms = parseDur(s, defUnit||"min");
  return {at:false, a:Math.floor(ms/60000), b:Math.round((ms%60000)/1000)};
}
function durText(p){
  if(p.at) return "@"+pad(clamp(p.a|0,0,23))+":"+pad(clamp(p.b|0,0,59));
  let a = Math.max(0,p.a|0), b = Math.max(0,p.b|0);
  if(b >= 60){ a += Math.floor(b/60); b = b%60; }
  return a+":"+pad(b);
}
function durInputs(host, get, set, opts){
  /* atHost may be an element, not a selector, when a caller has one at hand */
  const $$$ = function(x){ return typeof x==="string" ? $(x) : x; };
  opts = opts || {};
  const unit = opts.unit || "min";
  const draw = function(){
    const p = durParts(get(), unit);
    host.className = "durf" + (opts.compact ? " compact" : "");
    host.innerHTML = "";
    const box = function(key, val, max, lab){
      const w = document.createElement("span"); w.className = "durbox";
      const i = document.createElement("input");
      i.className = "inp"; i.type = "number"; i.min = "0"; i.max = String(max);
      i.value = String(val);
      i.setAttribute("aria-label", opts.label ? (opts.label+", "+lab) : lab);
      i.oninput = function(){
        const cur = durParts(get(), unit);
        cur[key] = parseInt(i.value,10) || 0;
        set(durText(cur));
      };
      const u = document.createElement("b"); u.textContent = lab;
      w.appendChild(i); w.appendChild(u);
      return w;
    };
    host.appendChild(box("a", p.a, p.at ? 23 : 999, p.at ? "год" : "хв"));
    host.appendChild(box("b", p.b, 59, p.at ? "хв" : "с"));
    const atHost = opts.atHost ? $$$(opts.atHost) : host;
    if(opts.atHost) atHost.innerHTML = "";
    if(opts.at){
      const t = document.createElement("button");
      t.type = "button"; t.className = "btn sm at" + (p.at ? " primary" : "");
      t.textContent = p.at ? "до часу" : (opts.compact ? "трив." : "тривалість");
      t.title = p.at ? "Рахувати до часу на годиннику. Натисніть, щоб задати тривалість"
                     : "Задана тривалість. Натисніть, щоб рахувати до часу на годиннику";
      t.onclick = function(){
        const cur = durParts(get(), unit);
        if(cur.at){ set(durText({at:false, a:cur.a, b:cur.b})); }
        else {
          const n = new Date();
          set(durText({at:true, a:(n.getHours()+1)%24, b:0}));
        }
        draw();
      };
      atHost.appendChild(t);
    }
  };
  draw();
}
function bindDur(sel, path, unit, allowAt){
  const host = $(sel); if(!host) return;
  const tame = function(){ if(typeof tameInputs==="function") tameInputs(host); };
  const get = function(){ return path.split(".").reduce(function(o,k){ return o[k]; }, cfg); };
  const set = function(v){
    const ks = path.split("."); let o = cfg;
    for(let i=0;i<ks.length-1;i++) o = o[ks[i]];
    o[ks[ks.length-1]] = v;
    onCfg(path);
  };
  durInputs(host, get, set, {unit:unit, at:!!allowAt, label:host.dataset.label||""});
  tame();
}
/* Settings written before durations became explicit still carry a bare number
   whose unit depended on its field. Convert once, on load. */
function migrateDurs(c){
  if(!c) return c;
  const fix = function(v, unit){
    const s = String(v==null ? "" : v);
    if(!s || s[0]==="@" || s.indexOf(":")>=0) return s;
    const ms = parseDur(s, unit);
    return Math.floor(ms/60000)+":"+pad(Math.round((ms%60000)/1000));
  };
  if(c.cd) c.cd.dur = fix(c.cd.dur,"min");
  if(c.iv){ c.iv.prep = fix(c.iv.prep,"sec"); c.iv.work = fix(c.iv.work,"sec"); c.iv.rest = fix(c.iv.rest,"sec"); }
  if(c.tn) c.tn.dur = fix(c.tn.dur,"min");
  if(c.ck) c.ck.bank = fix(c.ck.bank,"min");
  if(c.pg && c.pg.items) c.pg.items.forEach(function(it){ it.dur = fix(it.dur,"min"); });
  return c;
}

/* ================= presets ================= */
const PRESET_KEYS = ["mode","vis","bgFill","tenths","autoNext","overtime","loop","warn","danger","tickFrom","pack","cd","iv","pg","tn","ck"];
function snapshotPreset(){
  const o = {};
  PRESET_KEYS.forEach(function(k){ o[k] = JSON.parse(JSON.stringify(cfg[k])); });
  return o;
}
/* only a preset of your own can be deleted, so the button says as much */
function syncPresetBtns(){ $("#prDel").disabled = $("#prSel").value.indexOf("u:")!==0; }
function fillPresets(){
  const sel = $("#prSel");
  sel.innerHTML = '<option value="">— шаблон заняття —</option>';
  const g1 = document.createElement("optgroup"); g1.label = "Готові";
  Object.keys(BUILTIN).forEach(function(n){ const o=document.createElement("option"); o.value="b:"+n; o.textContent=n; g1.appendChild(o); });
  sel.appendChild(g1);
  const names = Object.keys(presets);
  if(names.length){
    const g2 = document.createElement("optgroup"); g2.label = "Мої";
    names.forEach(function(n){ const o=document.createElement("option"); o.value="u:"+n; o.textContent=n; g2.appendChild(o); });
    sel.appendChild(g2);
  }
  syncPresetBtns();
}
function applyPreset(v){
  if(!v) return;
  const isB = v.slice(0,2)==="b:", name = v.slice(2);
  const data = isB ? BUILTIN[name] : presets[name];
  if(!data) return;
  cfg = cleanRoster(deepMerge(cfg, data));
  syncUI(); setMode(cfg.mode); persist();
  toast("Пресет «"+name+"» застосовано");
}
function syncUI(){
  applyTheme(cfg.theme);
  $("#stageTitle").textContent = cfg.title;
  $("#o_warnColor").innerHTML = colorOptions(cfg.warnColor);
  $("#o_dangerColor").innerHTML = colorOptions(cfg.dangerColor);
  $("#iv_wcolor").innerHTML = colorOptions(cfg.iv.wcolor);
  $("#iv_rcolor").innerHTML = colorOptions(cfg.iv.rcolor);
  bindAll(); renderProgram(); renderRoster(); renderLaps(); renderLog();
  $$("[data-cfg]").forEach(function(d){ d.style.display = d.dataset.cfg===cfg.mode ? "" : "none"; });
  renderLogo();
  document.body.dataset.mode = cfg.mode;
  $$("#modes button").forEach(function(b){
    const on = b.dataset.mode===cfg.mode;
    b.classList.toggle("on", on); b.setAttribute("aria-pressed", on ? "true" : "false");
  });
  $("#modeTitle").textContent = MODE_TITLE[cfg.mode];
}

/* ================= misc ================= */
let toastT=0;
function toast(msg){
  const t = $("#toast"); t.textContent = msg; t.classList.add("on");
  clearTimeout(toastT); toastT = setTimeout(function(){ t.classList.remove("on"); }, 2200);
}
let wakeLock=null;
function requestWake(){
  if(EMBED) return;   /* a slide show keeps the screen on by itself */
  if(!cfg.wake || !navigator.wakeLock) return;
  navigator.wakeLock.request("screen").then(function(w){ wakeLock=w; }).catch(function(){});
}
function releaseWake(){ if(wakeLock){ try{ wakeLock.release(); }catch(e){} wakeLock=null; } }
document.addEventListener("visibilitychange", function(){ if(!document.hidden && S.running) requestWake(); });
/* the canvas backing store is resized from its own box, and resizing clears it */
window.addEventListener("resize", function(){ S.vizDirty = true; vizKick(); });

function togglePresent(){
  /* during the slide show the side panel must never come back in front of the audience */
  if(EMBED && hostView()==="read") return;
  if(EMBED){
    const open = document.body.classList.toggle("drawer");
    $("#tbFull").setAttribute("aria-expanded", open ? "true" : "false");
    if(open){ const f = document.querySelector(".side .tabs button.on") || $("#drawerClose"); if(f) f.focus(); }
    else $("#tbFull").focus();
    S.vizDirty = true; setTimeout(render, 80); return;
  }
  const on = !document.body.classList.contains("present");
  document.body.classList.toggle("present", on);
  if(on){ if(document.documentElement.requestFullscreen) document.documentElement.requestFullscreen().catch(function(){}); }
  else if(document.fullscreenElement && document.exitFullscreen) document.exitFullscreen().catch(function(){});
  setTimeout(render, 80);
}
function download(name, mime, text){
  if(EMBED){ copyOut(name, text); return; }
  const blob = new Blob([text], {type:mime});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob); a.download = name; a.click();
  setTimeout(function(){ URL.revokeObjectURL(a.href); }, 1000);
}

/* ================= overlays ================= */
function openOvl(sub, html, buttons){
  $("#ovlSub").textContent = sub;
  $("#ovlMain").innerHTML = html;
  const box = $("#ovlBtns"); box.innerHTML = "";
  buttons.forEach(function(b){
    const el = document.createElement("button");
    el.className = "btn" + (b.cls ? " "+b.cls : "");
    el.textContent = b.label; el.onclick = b.fn;
    box.appendChild(el);
  });
  ovlPrev = document.activeElement;
  $("#ovl").classList.add("on");
  /* focus moves into the dialog and comes back to where it started on close */
  const first = box.querySelector("button") || $("#ovl .ovlbox");
  setTimeout(function(){ if(first) first.focus(); }, 0);
}
let ovlPrev = null;
/* The browser confirm and prompt boxes are blocked inside PowerPoint add-ins, so questions
   go through the same overlay as the picker and the teams */
function askConfirm(text, okLabel, fn){
  openOvl("Підтвердіть", '<div class="ask"></div>', [
    {label:okLabel||"Так", cls:"primary", fn:function(){ closeOvl(); fn(); }},
    {label:"Скасувати", fn:closeOvl}
  ]);
  $("#ovlMain .ask").textContent = text;
}
function askText(title, hint, value, multi, fn){
  openOvl(title, '<div class="ask"></div>' + (multi ? '<textarea class="inp askin" rows="8"></textarea>' : '<input class="inp askin" type="text">'), [
    {label:"Готово", cls:"primary", fn:done},
    {label:"Скасувати", fn:closeOvl}
  ]);
  const inp = $("#ovlMain .askin");
  $("#ovlMain .ask").textContent = hint || "";
  inp.value = value || "";
  inp.onkeydown = function(e){
    if(e.key==="Escape"){ e.preventDefault(); closeOvl(); }
    else if(e.key==="Enter" && (!multi || e.ctrlKey)){ e.preventDefault(); done(); }
  };
  setTimeout(function(){ inp.focus(); if(!multi) inp.select(); }, 10);
  function done(){ const v = inp.value; closeOvl(); fn(v); }
}
/* No file downloads inside PowerPoint: the text goes to the clipboard, and if
   that is refused too it is shown ready to select and copy by hand */
function copyOut(name, text){
  function manual(){
    openOvl("Скопіюйте вручну", '<div class="ask"></div><textarea class="inp askin" rows="10" readonly></textarea>', [{label:"Закрити", cls:"primary", fn:closeOvl}]);
    $("#ovlMain .ask").textContent = "Виділіть текст (Ctrl+A), скопіюйте (Ctrl+C) і вставте у файл " + name;
    const t = $("#ovlMain .askin"); t.value = text;
    setTimeout(function(){ t.focus(); t.select(); }, 10);
  }
  if(navigator.clipboard && navigator.clipboard.writeText){
    navigator.clipboard.writeText(text).then(function(){ toast("Скопійовано. Вставте у файл " + name); }, manual);
  } else manual();
}
function closeOvl(){
  clearTimeout(spinT); clearInterval(Q.t);
  $("#ovl").classList.remove("on");
  if(ovlPrev && ovlPrev.focus) ovlPrev.focus();
  ovlPrev = null;
}

let spinT = 0;
function availablePool(){
  const all = cfg.roster.map(function(p){ return p.name; }).filter(Boolean);
  return cfg.noRepeat ? all.filter(function(n){ return S.picked.indexOf(n)<0; }) : all;
}
function openPicker(){
  if(!availablePool().length){ toast("Список порожній або всіх уже обрано"); return; }
  actx();
  openOvl("Жеребкування", '<div id="pickName">—</div>', [
    {label:"Ще раз", fn:spin, cls:"primary"},
    {label:"Скинути список", fn:function(){ S.picked=[]; updatePickState(); toast("Жеребкування скинуто"); }},
    {label:"Закрити", fn:closeOvl}
  ]);
  spin();
}
function spin(){
  const pool = availablePool();
  const el = $("#pickName"); if(!el) return;
  if(!pool.length){ el.textContent = "Усіх обрано"; $("#ovlSub").textContent = "Скиньте список, щоб почати знову"; return; }
  el.classList.add("spin");
  $("#ovlSub").textContent = "Обираємо…";
  let n = 0;
  const total = 22 + Math.floor(Math.random()*8);
  clearTimeout(spinT);
  const run = function(){
    el.textContent = pool[Math.floor(Math.random()*pool.length)];
    tone(500+Math.random()*400,.03,0,"square",.18);
    n++;
    if(n<total) spinT = setTimeout(run, 45 + Math.pow(n/total,3)*180);
    else{
      const w = pool[Math.floor(Math.random()*pool.length)];
      el.textContent = w; el.classList.remove("spin");
      if(cfg.noRepeat) S.picked.push(w);
      $("#ovlSub").textContent = cfg.noRepeat ? ("Залишилось у списку: "+(pool.length-1)) : "Жеребкування";
      updatePickState(); sndSeg(); announceSeg(w);
    }
  };
  run();
}
function openTeams(){
  const names = cfg.roster.map(function(p){ return p.name; }).filter(Boolean);
  if(names.length < 2){ toast("Додайте людей у ростер"); return; }
  const n = clamp(parseInt(cfg.teamCount)||3, 2, Math.max(2,names.length));
  const draw = function(){
    const mixed = shuffle(names.slice());
    const groups = [];
    for(let i=0;i<n;i++) groups.push([]);
    mixed.forEach(function(nm,i){ groups[i%n].push(nm); });
    $("#ovlMain").innerHTML = '<div class="teams">' + groups.map(function(g,i){
      return '<div class="team"><h5>Команда '+(i+1)+'</h5>'+g.map(function(x){ return '<div>'+esc(x)+'</div>'; }).join("")+'</div>';
    }).join("") + '</div>';
    sndSeg();
  };
  openOvl("Поділ на "+n+" команди", "", [
    {label:"Перемішати ще раз", fn:draw, cls:"primary"},
    {label:"Закрити", fn:closeOvl}
  ]);
  draw();
}
function openAgenda(){
  if(cfg.mode==="stopwatch" || !S.program.length){ toast("Огляд доступний для програми, раундів і черги"); return; }
  const base = Date.now() - doneMs();
  let acc = 0;
  const rows = S.program.map(function(s,i){
    const st = new Date(base+acc); acc += s.ms;
    const en = new Date(base+acc);
    const cls = i===S.idx ? " cur" : i<S.idx ? " past" : "";
    return '<div class="ag'+cls+'"><span class="when">'+hhmm(st)+" – "+hhmm(en)+'</span>'+
      '<span class="what">'+esc(s.name)+(s.note ? '<span class="nt">'+esc(s.note)+'</span>' : "")+'</span>'+
      '<span class="how">'+fmtShort(s.ms)+'</span></div>';
  }).join("");
  openOvl("Огляд програми · разом "+fmt(totalMs()),
    '<div class="agenda">'+rows+'</div>',
    [{label:"Закрити", fn:closeOvl, cls:"primary"}]);
}

/* ================= quiz buzzer ================= */
function newQuestion(){
  Q.locked = null; Q.out = []; Q.question++;
  clearInterval(Q.t); closeOvl();
  updateBuzzState(); renderScore();
  tone(880,.12,0,"sine",.5); tone(1180,.18,.12,"sine",.5);
  toast("Питання "+Q.question+", усі кнопки активні");
}
function buzz(i){
  if(!cfg.buzzOn || Q.locked!==null) return;
  const p = cfg.roster[i];
  if(!p || !p.name || Q.out.indexOf(i)>=0) return;
  actx();
  Q.locked = i; Q.deadline = Date.now() + Math.max(2,cfg.buzzAnswer)*1000;
  sndSeg(); announceSeg(p.name);
  openOvl("Питання "+Q.question,
    '<div id="buzzName">'+esc(p.name)+'</div><div id="buzzCount">'+cfg.buzzAnswer+'</div>',
    [
      {label:"Правильно", fn:buzzRight, cls:"primary"},
      {label:"Помилка", fn:buzzWrong, cls:"danger"},
      {label:"Нове питання", fn:newQuestion}
    ]);
  clearInterval(Q.t); Q.t = setInterval(buzzTick, 100);
}
function buzzTick(){
  const el = $("#buzzCount"); if(!el){ clearInterval(Q.t); return; }
  const l = Q.deadline - Date.now();
  el.textContent = Math.max(0, Math.ceil(l/1000));
  if(l <= 0){ clearInterval(Q.t); buzzWrong(true); }
}
function buzzRight(){
  const i = Q.locked; if(i===null) return;
  cfg.roster[i].score = (cfg.roster[i].score||0) + 1;
  clearInterval(Q.t); Q.locked = null;
  renderRoster(); persist(); sndEnd(); closeOvl();
  toast(cfg.roster[i].name+" отримує бал");
  Q.out = []; Q.question++; updateBuzzState();
}
function buzzWrong(timeout){
  const i = Q.locked; if(i===null) return;
  clearInterval(Q.t); Q.out.push(i); Q.locked = null;
  renderScore(); updateBuzzState(); closeOvl();
  tone(180,.4,0,"sawtooth",.5);
  toast(cfg.roster[i].name + (timeout===true ? " не встигла" : " вибуває з цього питання"));
}

/* ================= accessibility =================
   The stage is watched on a projector but driven from a keyboard, so every
   control needs a name, the tabs need arrow keys, the overlay needs a focus
   trap, and phase changes need one spoken status each — not one per tick. */
function human(ms){
  const s = Math.max(0,Math.round(ms/1000)), m = Math.floor(s/60), r = s%60;
  return ((m ? m+" хв" : "") + ((r||!m) ? " "+r+" с" : "")).trim();
}
let annKey = "";
function announce(txt){
  const el = $("#live"); if(el) el.textContent = txt;
}
function syncTabs(){
  $$("#tabs button[data-tab]").forEach(function(b){
    const on = b.classList.contains("on");
    b.setAttribute("aria-selected", on ? "true" : "false");
    b.tabIndex = on ? 0 : -1;          /* roving tabindex: one stop for the set */
  });
}
function selectTab(b){
  $$("#tabs button[data-tab]").forEach(function(x){ x.classList.toggle("on", x===b); });
  $$(".pane").forEach(function(p){ p.classList.toggle("on", p.dataset.pane===b.dataset.tab); });
  $("#scroll").scrollTop = 0;
  syncTabs();
}
/* Names, labels and notes are session content, not profile data: the browser
   should neither autofill them nor underline them in red. */
function tameInputs(root){
  $$("input", root||document).forEach(function(i){
    const t = (i.type||"text").toLowerCase();
    if(t==="file" || t==="checkbox" || t==="radio" || t==="range") return;
    if(!i.getAttribute("autocomplete")) i.setAttribute("autocomplete","off");
    if(t!=="number" && !i.hasAttribute("spellcheck")) i.setAttribute("spellcheck","false");
    if(t==="number" && !i.getAttribute("inputmode")) i.setAttribute("inputmode","numeric");
  });
}
(function wireA11y(){
  /* glyphs are decorative: the button's own label carries the meaning */
  $$("button svg,a svg,.brand svg").forEach(function(s){
    s.setAttribute("aria-hidden","true"); s.setAttribute("focusable","false");
  });
  /* an icon-only button has no text, so its tooltip becomes its name */
  $$("button[title]").forEach(function(b){
    if(!b.getAttribute("aria-label") && !b.textContent.trim()) b.setAttribute("aria-label", b.getAttribute("title"));
  });
  $$("#tabs button[data-tab]").forEach(function(b){
    b.setAttribute("role","tab"); b.id = "tab-"+b.dataset.tab;
    b.setAttribute("aria-controls","pane-"+b.dataset.tab);
  });
  $$(".pane").forEach(function(p){
    p.id = "pane-"+p.dataset.pane;
    p.setAttribute("role","tabpanel"); p.setAttribute("aria-labelledby","tab-"+p.dataset.pane);
  });
  $$("#modes button[data-mode]").forEach(function(b){ b.setAttribute("aria-pressed","false"); });
  /* A visible caption that is not tied to its control is a label only by
     proximity: a screen reader may read it, and may not. Each .f block holds
     one caption and one control, so they are bound here rather than by hand on
     forty inputs. The minutes-and-seconds pairs are built later and carry two
     boxes, so there the caption names the group and each box says which half
     of it it is. */
  let autoId = 0;
  $$(".f").forEach(function(f){
    const lab = f.querySelector("label"); if(!lab || lab.getAttribute("for")) return;
    const ctl = f.querySelector("input,select,textarea");
    if(ctl){
      if(!ctl.id) ctl.id = "ctl-" + (++autoId);
      lab.setAttribute("for", ctl.id);
      return;
    }
    if(!lab.id) lab.id = "lab-" + (++autoId);
    f.setAttribute("role","group");
    f.setAttribute("aria-labelledby", lab.id);
  });
  syncTabs();
  tameInputs();
})();
$("#tabs").onkeydown = function(e){
  const bs = $$("#tabs button[data-tab]");
  const i = bs.indexOf(document.activeElement); if(i<0) return;
  let j = -1;
  if(e.key==="ArrowRight"||e.key==="ArrowDown") j = (i+1)%bs.length;
  else if(e.key==="ArrowLeft"||e.key==="ArrowUp") j = (i-1+bs.length)%bs.length;
  else if(e.key==="Home") j = 0;
  else if(e.key==="End") j = bs.length-1;
  if(j<0) return;
  e.preventDefault(); e.stopPropagation();   /* keep ArrowUp away from the +30 s hotkey */
  selectTab(bs[j]); bs[j].focus();
};
/* the clock doubles as an edit control, so it answers Enter and Space */
$("#timeText").onkeydown = function(e){
  if(e.key==="Enter"||e.key===" "){ e.preventDefault(); $("#timeText").click(); }
};
$("#ovl").addEventListener("keydown", function(e){
  if(e.key!=="Tab") return;
  const f = $$("#ovl button, #ovl [tabindex]:not([tabindex='-1'])").filter(function(x){ return x.offsetParent!==null; });
  if(!f.length) return;
  const first = f[0], last = f[f.length-1];
  if(e.shiftKey && document.activeElement===first){ e.preventDefault(); last.focus(); }
  else if(!e.shiftKey && document.activeElement===last){ e.preventDefault(); first.focus(); }
});

/* ================= events ================= */
$("#tabs").onclick = function(e){
  const b = e.target.closest("button[data-tab]"); if(b) selectTab(b);
};
$("#modes").onclick = function(e){
  const b = e.target.closest("button[data-mode]"); if(!b) return;
  setMode(b.dataset.mode);
};
$$(".card>h3").forEach(function(h){
  /* the heading stays a heading; the button role goes on a span inside it */
  const btn = document.createElement("span");
  btn.className = "cardbtn";
  while(h.firstChild) btn.appendChild(h.firstChild);
  h.appendChild(btn);
  btn.setAttribute("role","button");
  btn.tabIndex = 0;
  const sync = function(){ btn.setAttribute("aria-expanded", h.parentElement.classList.contains("closed") ? "false" : "true"); };
  /* not toggle: that name belongs to start/pause at the top of the file */
  const toggleCard = function(){ h.parentElement.classList.toggle("closed"); sync(); };
  h.onclick = toggleCard;
  btn.onkeydown = function(e){ if(e.key==="Enter" || e.key===" "){ e.preventDefault(); toggleCard(); } };
  sync();
});

$("#btnStart").onclick = toggle;
$("#btnReset").onclick = function(){
  if(S.running){ askConfirm("Таймер іде. Скинути на початок?", "Скинути", resetAll); return; }
  resetAll();
};
$("#btnNext").onclick = nextSeg;
$("#btnPrev").onclick = prevSeg;
$("#btnPlus").onclick = function(){ addTime(cfg.step*1000); };
$("#btnMinus").onclick = function(){ addTime(-cfg.step*1000); };
$("#swLap").onclick = lap;
$("#btnAgenda").onclick = openAgenda;
$("#tbAgenda").onclick = openAgenda;

$("#timeText").onclick = function(){
  if(cfg.mode==="stopwatch") return;
  const seg = S.program[S.idx]; if(!seg) return;
  askText("Новий час", "Етап «"+seg.name+"»", fmtInput(seg.ms), false, function(v){
    const ms = parseDur(v, cfg.mode==="interval" ? "sec" : "min");
    if(ms>0) setSegTime(ms); else toast("Не вдалося розібрати час");
  });
};
$("#stageTitle").oninput = function(){ cfg.title = $("#stageTitle").textContent.trim(); persist(true); };
$("#stageTitle").onkeydown = function(e){ if(e.key==="Enter"){ e.preventDefault(); $("#stageTitle").blur(); } };

$("#tbSound").onclick = function(){ cfg.sound=!cfg.sound; $("#o_sound").checked=cfg.sound; persist(); render(); toast(cfg.sound?"Звук увімкнено":"Звук вимкнено"); };
$("#tbScore").onclick = function(){ cfg.showScore=!cfg.showScore; $("#o_showScore").checked=cfg.showScore; persist(); render(); };
/* The button walks the seven built-ins only: cycling seventy-two themes one
   click at a time would be useless. Everything else is picked from the list. */
$("#tbTheme").onclick = function(){
  const list = TBASE.map(function(t){ return t[0]; });
  const i = list.indexOf(cfg.theme);
  applyTheme(list[(i+1)%list.length]);
  $("#o_theme").value = cfg.theme; persist(); render(); toast("Тема: "+themeName(cfg.theme));
};
$("#tbFull").onclick = togglePresent;
$("#tbWindow").onclick = openDisplay;
$("#sndTest").onclick = function(){ actx(); sndSeg(); setTimeout(sndTick,500); setTimeout(sndEnd,900); };
$("#clipFile").onchange = function(){
  const file = this.files[0]; this.value = "";
  if(!file || !clipTarget) return;
  if(file.size > CLIP_MAX){ toast("Файл на "+Math.round(file.size/1024)+" КБ завеликий. Потрібно до 700 КБ"); return; }
  const rd = new FileReader();
  rd.onerror = function(){ toast("Не вдалося прочитати файл"); };
  rd.onload = function(){
    const key = clipTarget, prev = clips[key];
    clips[key] = String(rd.result);
    /* put the old clip back rather than leave the store half written */
    if(clipsBytes() > CLIP_TOTAL){
      if(prev) clips[key] = prev; else delete clips[key];
      toast("Разом файли не вміщаються. Приберіть зайві або візьміть коротші");
    } else if(!clipsSave()){
      if(prev) clips[key] = prev; else delete clips[key];
    } else {
      delete clipAudio[key];
      toast("Файл призначено");
    }
    renderClips();
  };
  rd.readAsDataURL(file);
};

$("#msgAdd").onclick = function(){ cfg.msgs.push("Нова репліка"); renderMsgs(); persist(); };
$("#msgHide").onclick = function(){ showMsg(""); };

$("#styleDtek").onclick = function(){
  cfg.theme = "dtek-light"; cfg.vis = "square"; cfg.bgFill = true;
  $("#o_theme").value = cfg.theme; $("#o_vis").value = cfg.vis; $("#o_bgFill").checked = true;
  applyTheme(cfg.theme); persist(); render(); toast("Стиль ДТЕК застосовано");
};
$("#logoPick").onclick = function(){ $("#logoFile").click(); };
$("#logoDrop").onclick = function(){ logoData = ""; logoSave(); renderLogo(); };
$("#logoFile").onchange = function(){
  const file = this.files[0]; this.value = "";
  if(!file) return;
  if(file.size > LOGO_MAX){ toast("Файл на "+Math.round(file.size/1024)+" КБ завеликий. Потрібно до 400 КБ"); return; }
  const rd = new FileReader();
  rd.onerror = function(){ toast("Не вдалося прочитати файл"); };
  rd.onload = function(){
    const prev = logoData;
    logoData = String(rd.result);
    if(!logoSave()){ logoData = prev; }
    else {
      cfg.logo.on = true; $("#o_logoOn").checked = true; persist();
      toast("Логотип вбудовано");
      logoAutoFit(logoData);
    }
    renderLogo();
  };
  rd.readAsDataURL(file);
};

$("#musPick").onclick = function(){ $("#musFile").click(); };
$("#musFile").onchange = function(){
  const f = $("#musFile").files[0]; if(!f) return;
  if(musicUrl) URL.revokeObjectURL(musicUrl);
  musicUrl = URL.createObjectURL(f);
  const a = $("#bgm"); a.src = musicUrl; a.loop = !!cfg.musicLoop; a.volume = cfg.musicVol;
  $("#musName").textContent = f.name;
  toast("Музику підключено");
};

$("#pgAdd").onclick = function(){ cfg.pg.items.push({name:"Новий етап",dur:"5:00",color:"",note:""}); renderProgram(); persist(); if(!S.running) compile(); };
$("#pgBulk").onclick = function(){
  const cur = cfg.pg.items.map(function(i){ return i.name+" | "+i.dur; }).join("\n");
  askText("Етапи списком", "Один етап на рядок:  Назва | тривалість (10 = хв, 45s = сек, 7:30 = хв:сек). Ctrl+Enter — готово", cur, true, function(t){
    cfg.pg.items = t.split("\n").map(function(l){ return l.trim(); }).filter(Boolean).map(function(l){
      const p = l.split("|");
      return {name:(p[0]||"Етап").trim(), dur:(p[1]||"5").trim(), color:"", note:""};
    });
    renderProgram(); persist(); if(!S.running) compile();
  });
};
$("#rsAdd").onclick = function(){ cfg.roster.push({name:"Учасник "+(cfg.roster.length+1),score:0}); renderRoster(); persist(); renderClips(); if(cfg.mode==="turns" && !S.running) compile(); };
$("#rsBulk").onclick = function(){
  askText("Учасники списком", "Один учасник або команда на рядок. Ctrl+Enter — готово", cfg.roster.map(function(p){ return p.name; }).join("\n"), true, function(t){
  const old = {}; cfg.roster.forEach(function(p){ old[p.name] = p.score||0; });
  cfg.roster = t.split("\n").map(function(s){ return s.trim(); }).filter(Boolean).map(function(n){ return {name:n, score:old[n]||0}; });
  S.picked = []; Q.out = []; renderRoster(); persist(); if(cfg.mode==="turns" && !S.running) compile();
  });
};
$("#rsZero").onclick = function(){
  const any = cfg.roster.some(function(p){ return (p.score||0)!==0; });
  function zero(){ cfg.roster.forEach(function(p){ p.score=0; }); renderRoster(); persist(); }
  if(any) askConfirm("Обнулити всі бали?", "Обнулити", zero); else zero();
};
$("#rsShuffle").onclick = function(){ cfg.roster = shuffle(cfg.roster.slice()); renderRoster(); persist(); if(cfg.mode==="turns" && !S.running) compile(); toast("Ростер перемішано"); };
$("#btnPick").onclick = openPicker;
$("#btnTeams").onclick = openTeams;
$("#btnNewQ").onclick = newQuestion;
$("#ovl").onclick = function(e){ if(e.target===$("#ovl")) closeOvl(); };

$("#logCsv").onclick = function(){
  if(!S.log.length){ toast("Журнал порожній"); return; }
  const rows = [["Етап","План, с","Факт, с","Різниця, с"]].concat(S.log.map(function(e){
    return ['"'+String(e.name).replace(/"/g,'""')+'"', Math.round(e.planned/1000), Math.round(e.actual/1000), Math.round((e.actual-e.planned)/1000)];
  }));
  download("session-"+new Date().toISOString().slice(0,10)+".csv", "text/csv;charset=utf-8", "﻿"+rows.map(function(r){ return r.join(";"); }).join("\n"));
};
/* the log is the only record of how the session actually ran, and nothing
   rebuilds it: ask before it goes */
$("#logClear").onclick = function(){
  function clear(){ S.log=[]; renderLog(); }
  if(S.log.length) askConfirm("Очистити журнал сесії? Записи не відновлюються.", "Очистити", clear); else clear();
};

$("#prSel").onchange = function(){ applyPreset($("#prSel").value); syncPresetBtns(); };
$("#prSave").onclick = function(){
  askText("Новий пресет", "Назва пресета", cfg.title || "Мій пресет", false, function(n){
    n = n.trim(); if(!n) return;
    presets[n] = snapshotPreset(); saveStore(); fillPresets(); $("#prSel").value = "u:"+n; toast("Збережено");
  });
};
$("#prDel").onclick = function(){
  const v = $("#prSel").value;
  if(v.slice(0,2)!=="u:"){ toast("Готові пресети видалити не можна"); return; }
  const nm = v.slice(2);
  askConfirm("Видалити пресет «"+nm+"»? Це не можна скасувати.", "Видалити", function(){
    delete presets[nm]; saveStore(); fillPresets(); toast("Видалено");
  });
};
$("#prExport").onclick = function(){
  download("training-timer-"+new Date().toISOString().slice(0,10)+".json", "application/json", JSON.stringify({cfg:cfg,presets:presets},null,2));
};
$("#prImport").onclick = function(){ $("#prFile").click(); };
$("#prFile").onchange = function(){
  const f = $("#prFile").files[0]; if(!f) return;
  const r = new FileReader();
  r.onload = function(){
    try{
      const d = JSON.parse(r.result);
      if(d.cfg) cfg = cleanRoster(migrateDurs(deepMerge(DEF, d.cfg)));
      if(d.presets){ presets = Object.assign(presets, d.presets); Object.keys(presets).forEach(function(n){ migrateDurs(presets[n]); }); }
      saveStore(); syncUI(); fillPresets(); setMode(cfg.mode); toast("Імпортовано");
    }catch(e){ toast("Не вдалося прочитати файл"); }
  };
  r.readAsText(f); $("#prFile").value = "";
};
$("#prReset").onclick = function(){
  askConfirm("Скинути всі налаштування та пресети?", "Скинути", function(){
    localStorage.removeItem(KEY); cfg = deepMerge(DEF,{}); presets = {};
    syncUI(); fillPresets(); setMode(cfg.mode); saveStore(); toast("Скинуто");
  });
};

/* Shortcuts are read from the physical key rather than from the character it
   produces. This is a Ukrainian tool, and on a Cyrillic layout the N key
   reports "т", which left every letter shortcut dead exactly when the trainer
   was typing in Ukrainian. Anything held with Ctrl, Alt or Cmd belongs to the
   browser, so Ctrl+R reloads and Ctrl+P prints without also driving the timer. */
const KEYS = {
  KeyR:function(){ $("#btnReset").click(); },
  KeyN:nextSeg, KeyP:prevSeg, KeyF:togglePresent, KeyD:openDisplay,
  KeyM:function(){ $("#tbSound").click(); }, KeyT:function(){ $("#tbScore").click(); },
  KeyA:openAgenda, KeyJ:openPicker, KeyG:openTeams, KeyQ:newQuestion, KeyL:lap,
  KeyX:function(){ showMsg(""); }
};
document.addEventListener("keydown", function(e){
  if(e.ctrlKey || e.metaKey || e.altKey) return;
  const t = e.target;
  if(t && (t.tagName==="INPUT"||t.tagName==="TEXTAREA"||t.tagName==="SELECT"||t.isContentEditable)) return;
  /* a focused button owns Space and Enter; the global shortcuts stay out of it */
  if((e.code==="Space"||e.key==="Enter") && t && t.closest && t.closest("button,a,[role='button']")) return;
  const digit = /^Digit[1-9]$/.test(e.code) ? +e.code.slice(5) : (/^[1-9]$/.test(e.key) ? +e.key : 0);
  /* shift plus a digit is the message, so the quiz lockout keeps the bare digits */
  if(e.shiftKey && digit && cfg.msgs[digit-1]){ e.preventDefault(); showMsg(cfg.msgs[digit-1]); return; }
  if(!e.shiftKey && cfg.buzzOn && digit){ e.preventDefault(); buzz(digit-1); return; }
  if(e.code==="Space"){
    e.preventDefault();
    if(cfg.mode==="clock" && S.running) ckSwitch(1); else toggle();
    return;
  }
  if(e.key==="ArrowUp"){ e.preventDefault(); addTime(cfg.step*1000); return; }
  if(e.key==="ArrowDown"){ e.preventDefault(); addTime(-cfg.step*1000); return; }
  if(e.key==="Escape"){
    if($("#stageMsg").classList.contains("on")) showMsg("");
    else if($("#ovl").classList.contains("on")) closeOvl();
    else if(EMBED ? document.body.classList.contains("drawer") : document.body.classList.contains("present")) togglePresent();
    return;
  }
  /* the physical key first, then the character, for keyboards the map misses */
  const fn = KEYS[e.code] || KEYS["Key"+e.key.toUpperCase()];
  if(fn) fn();
});

/* ================= quick durations ================= */
[1,2,3,5,7,10,15,20,30,45,60].forEach(function(m){
  const b = document.createElement("button");
  b.textContent = m + " хв";
  b.onclick = function(){
    cfg.cd.dur = m+":00"; bindDur("#cd_dur","cd.dur","min",true);
    persist(); if(!S.running) compile(); else setSegTime(m*60000);
  };
  $("#cdQuick").appendChild(b);
});
(function(){
  const b = document.createElement("button");
  b.className = "hot"; b.textContent = "Бліц 90 с";
  b.onclick = function(){ blitz(90); };
  $("#cdQuick").appendChild(b);
})();

/* ================= boot ================= */
if(EMBED){
  document.body.classList.add("embed");
  cfg.noise = false;   /* the add-in has no microphone access */
  $("#prExport").textContent = "Копіювати JSON";
  $("#logCsv").textContent = "Копіювати CSV";
  $("#tbFull").title = "Налаштування (F)";
  $("#tbFull").setAttribute("aria-label", "Налаштування");
  $("#tbFull").setAttribute("aria-expanded", "false");
  $("#tbFull").innerHTML = '<svg viewBox="0 0 24 24"><path d="M4 6h10M18 6h2M4 12h4M12 12h8M4 18h12M20 18h0"/><circle cx="16" cy="6" r="2"/><circle cx="10" cy="12" r="2"/><circle cx="18" cy="18" r="2"/></svg>';
  $("#drawerClose").onclick = function(){ if(document.body.classList.contains("drawer")) togglePresent(); };
  /* a click on the stage beside the open drawer closes it */
  $("#stage").addEventListener("pointerdown", function(e){
    if(document.body.classList.contains("drawer") && !e.target.closest(".stagebar")) togglePresent();
  });
  /* PowerPoint switches between editing and the slide show. Both get the stage on
     the whole block; the editor opens the settings as a drawer. */
  window.WOS_TIMER_VIEW = function(v){
    document.body.classList.add("present");
    document.body.classList.toggle("edit", v!=="read");
    if(v==="read") document.body.classList.remove("drawer");
    S.vizDirty = true; setTimeout(render, 80);
  };
  let idleT = 0;
  const wake = function(){
    document.body.classList.remove("idle");
    clearTimeout(idleT);
    idleT = setTimeout(function(){ document.body.classList.add("idle"); }, 3000);
  };
  ["pointermove","pointerdown","focusin"].forEach(function(ev){ document.addEventListener(ev, wake, true); });
  wake();
  window.WOS_TIMER_VIEW(hostView());
}
applyTheme(cfg.theme);
syncUI(); fillPresets(); setMode(cfg.mode); renderScore(); updateBuzzState(); renderClips(); renderMsgs(); renderLogo();
if(cfg.noise) noiseStart();
if(!IS_DISPLAY) setInterval(function(){ if(!S.running) render(); }, 1000);
render(); renderTimeline();
if(IS_DISPLAY) initDisplay();
vizKick();
