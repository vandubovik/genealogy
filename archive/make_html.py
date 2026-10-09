# -*- coding: utf-8 -*-
"""Генератор HTML-страницы со списком дел из arch.xlsx.
Читает НАРБ и НИАБ, сохраняет data.json и arch.html рядом с xlsx.
"""
import openpyxl, json, datetime, pathlib, html, sys

SRC = r"C:\Ivan\gen\arch.xlsx"
OUT_DIR = r"C:\Ivan\Sync\Drop\tree\site"
OUT_JSON = OUT_DIR + r"\data.json"
OUT_HTML = OUT_DIR + r"\arch.html"
OUT_CSS = OUT_DIR + r"\style.css"

wb = openpyxl.load_workbook(SRC, data_only=True)

def cell(v):
    if v is None:
        return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%d.%m.%Y")
    return str(v)

def looks_like_header(row, nrows):
    vals = [c for c in row if c.strip()]
    if len(vals) < 2:
        return False
    # все короткие строки без цифр в начале
    return all(len(v) < 40 and not v[:1].isdigit() for v in vals)

def extract(sheet_name):
    ws = wb[sheet_name]
    rows = [[cell(c.value) for c in r] for r in ws.iter_rows()]
    # убрать полностью пустые строки
    rows = [r for r in rows if any(x.strip() for x in r)]
    cases = []
    cur = None
    current_cols = []
    for r in rows:
        first = r[0].strip()
        col1 = r[1].strip() if len(r) > 1 else ""
        if col1 == "ф" or first.lower() == "архив":
            current_cols = [c for c in r[5:]]
            while current_cols and not current_cols[-1].strip():
                current_cols.pop()
            continue
        is_case_header = first in ("НАРБ", "НИАБ") and col1.isdigit()
        if first and not is_case_header:
            # '-' в первом столбце — разделитель: дела без категории
            cur = {"section": True, "title": ("Без категории" if first == "-" else first), "archive": "", "f": "", "o": "", "d": "",
                   "note": "", "cols": [], "rows": []}
            cases.append(cur)
            continue
        if is_case_header:
            cur = {"section": False, "title": r[4].strip() if len(r) > 4 else "",
                   "archive": r[0].strip(), "f": r[1].strip(), "o": r[2].strip(), "d": r[3].strip(),
                   "note": "", "cols": list(current_cols), "rows": []}
            extra = [c for c in r[5:] if c.strip()]
            if extra:
                cur["note"] = "; ".join(extra)
            cases.append(cur)
            continue
        if cur is None:
            continue
        if any(x.strip() for x in r):
            trimmed = r[5:]
            while trimmed and not trimmed[-1].strip():
                trimmed.pop()
            if current_cols:
                trimmed = trimmed[:len(current_cols)] + [""] * max(0, len(current_cols) - len(trimmed))
            if any(t.strip() for t in trimmed):
                cur["rows"].append(trimmed)
    # удалить полностью пустые колонки у каждой таблицы; заполнить «лист» во все строки
    for c in cases:
        if c["rows"] and c["cols"]:
            n = len(c["cols"])
            keep = [i for i in range(n) if c["cols"][i].strip() or any(row[i].strip() for row in c["rows"])]
            c["cols"] = [c["cols"][i] for i in keep]
            c["rows"] = [[row[i] for i in keep] for row in c["rows"]]
            for i, col in enumerate(c["cols"]):
                if col.strip().lower().startswith("лист"):
                    last = ""
                    for row in c["rows"]:
                        if row[i].strip():
                            last = row[i]
                        elif last:
                            row[i] = last
    return cases

data = {"НАРБ": extract("НАРБ"), "НИАБ": extract("НИАБ")}
pathlib.Path(OUT_JSON).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

CSS = """/* ===== Общие переменные ===== */
:root{
  --border:#d8dee6;   /* цвет рамок */
  --head:#f2f5f9;     /* цвет шапок таблиц */
  --accent:#2b6cb0;   /* синий акцент (сортировка, фокус) */
}

*{box-sizing:border-box}

body{
  margin:0;
  padding:24px 24px 80px 24px;
  font-family:system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;
  font-size:15px;
  color:#1a202c;
  background:#f7f9fc;
}

h1{font-size:20px;margin:0 0 16px}

/* Подсветка найденного при поиске */
b.hl{font-weight:700;color:#e11d2e}

/* Панель инструментов: поиск + кнопки */
.toolbar{
  display:flex;
  gap:12px;
  align-items:center;
  flex-wrap:wrap;
  margin-bottom:12px;
}
input[type=search]{
  padding:8px 12px;
  border:1px solid var(--border);
  border-radius:8px;
  font:inherit;
  min-width:280px;
  background:#fff;
}
input[type=search]:focus{
  outline:2px solid var(--accent);
  outline-offset:1px;
  border-color:transparent;
}
button{
  padding:7px 12px;
  border:1px solid var(--border);
  border-radius:8px;
  background:#fff;
  font:inherit;
  cursor:pointer;
}
button:hover{background:var(--head)}

/* Заголовки секций (НАРБ, НИАБ, категории) */
h2{font-size:16px;margin:24px 0 10px}

/* Карточка вокруг каждой таблицы секции */
.table-wrap{
  overflow:auto;
  background:#fff;
  border:1px solid var(--border);
  border-radius:10px;
  box-shadow:0 1px 2px rgba(16,24,40,.05);
  margin-bottom:24px;
}

/* Основная таблица дел */
table.cases{
  border-collapse:separate;
  border-spacing:0;
  width:100%;
  font-size:14px;
}
table.cases th,table.cases td{
  padding:7px 10px;
  border-bottom:1px solid var(--border);
  text-align:left;
  vertical-align:top;
  overflow-wrap:anywhere;
  word-break:break-word;
}
table.cases > thead > tr > th:nth-child(-n+4){width:62px}
table.cases > thead > tr > th{
  position:sticky;
  top:0;
  z-index:1;
  background:var(--head);
  white-space:nowrap;
  font-weight:600;
}
table.cases tbody tr:last-child td{border-bottom:none}

/* Раскрываемые дела — жирным; без деталей — серым */
tr.section td{background:var(--head);font-weight:bold}
tr.has td{font-weight:bold;color:#1a202c}
tr.has td.note{font-weight:normal;color:#5a6472}
tr.nodata td{color:#5a6472}
tr.clickable{cursor:pointer}
tr.clickable td{transition:background .1s}
tr.clickable:hover td{background:#eef4fb}

/* Строка раскрытой вложенной таблицы */
tr.dets td{
  background:#fff;
  padding:10px 14px;
  border-top:1px solid var(--border);
}

/* Вложенная таблица */
table.data{
  border-collapse:separate;
  border-spacing:0;
  width:auto;
  min-width:50%;
  font-size:13px;
  background:#fff;
}
table.data th,table.data td{
  padding:4px 9px;
  border:1px solid var(--border);
  text-align:left;
  vertical-align:top;
  overflow-wrap:anywhere;
  word-break:break-word;
}
table.data th{
  background:var(--head);
  white-space:nowrap;
  cursor:pointer;
  user-select:none;
  font-weight:600;
}
table.data tbody tr:last-child td{border-bottom:1px solid var(--border)}

/* Индикаторы сортировки: ↕ — не сортировали, ▲ / ▼ — направление */
table.data th.sorth::after{content:' ↕';opacity:.4;font-size:11px;margin-left:3px}
table.data th.sorth.asc::after{content:' ▲';opacity:1;color:var(--accent)}
table.data th.sorth.desc::after{content:' ▼';opacity:1;color:var(--accent)}

/* Примечание и пустые состояния */
.note{color:#5a6472;font-size:13px}
.empty{color:#aab2bd;text-align:center;padding:40px}
"""
pathlib.Path(OUT_CSS).write_text(CSS, encoding="utf-8")

HTML = r"""<!DOCTYPE html>
<html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Список дел</title><link rel="stylesheet" href="style.css?v=8"></head><body>
<h1>Список дел</h1>
<div class="toolbar">
<input type="search" id="q" placeholder="Поиск по всем делам…">
<button id="expand">Раскрыть все</button><button id="collapse">Свернуть все</button><button id="unsort">Сбросить сортировку</button>
</div>
<div id="app"><p class="empty">Загрузка data.json… (если не загрузилось — запустите из этой папки: python -m http.server 8000 и откройте http://localhost:8000/arch.html)</p></div>
<script>
fetch('data.json').then(r=>r.json()).then(render).catch(e=>{
  document.getElementById('app').innerHTML='<p class="empty">Не удалось загрузить data.json. Откройте страницу через локальный сервер: python -m http.server 8000 → http://localhost:8000/arch.html</p>';
});
function esc(s){return String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}
function nestedTable(rows, cols){
  if(!rows.length) return '';
  let head='';
  let bodyRows=rows;
  if(cols && cols.length){head='<tr>'+cols.map((c,i)=>`<th data-i="${i}" class="sorth">${esc(c)}</th>`).join('')+'</tr>';}
  // колонка «лист»: пустые значения берут предыдущее; подряд идущие одинаковые — показываем один раз
  const width=Math.max((cols||[]).length,...bodyRows.map(r=>r.length),0);
  const body=bodyRows.map((r,idx)=>{const rr=[...r];while(rr.length<width)rr.push('');return `<tr data-o="${idx}">`+rr.map(c=>`<td>${esc(c)}</td>`).join('')+'</tr>';}).join('');
  return `<table class="data sortable"><thead>${head}</thead><tbody>${body}</tbody></table>`;
}
function caseRow(c){
  const tds=`<td>${esc(c.archive)}</td><td>${esc(c.f)}</td><td>${esc(c.o)}</td><td>${esc(c.d)}</td><td>${esc(c.title)}</td><td class="note">${esc(c.note)}</td>`;
  const rowsText=c.rows.map(r=>r.join(' ')).join(' ');
  const search=esc((c.archive+c.f+c.o+c.d+c.title+c.note+rowsText).toLowerCase());
  if(!c.rows.length) return `<tr data-search="${search}" class="nodata">${tds}</tr>`;
  return `<tr data-search="${search}" class="clickable has" data-open="">${tds}</tr><tr class="dets" data-search="${search}" style="display:none"><td colspan="6">${nestedTable(c.rows,c.cols)}</td></tr>`;
}
function render(data){
  let h='';
  for(const sec of Object.keys(data)){
    h+=`<h2>${sec}</h2><div class="table-wrap"><table class="cases"><thead><tr><th>Архив</th><th>Фонд</th><th>Опись</th><th>Дело</th><th>Название дела</th><th>Примечание</th></tr></thead><tbody>`;
    for(const c of data[sec]){
      if(c.section){h+=`<tr data-search="${esc((c.title+c.note).toLowerCase())}" class="section"><td colspan="6">${esc(c.title)}</td></tr>`;continue;}
      h+=caseRow(c);
    }
    h+='</tbody></table></div>';
  }
  document.getElementById('app').innerHTML=h;
  document.querySelectorAll('tr.clickable').forEach(tr=>{
    tr.addEventListener('click',()=>{
      const d=tr.nextElementSibling;
      const open=tr.dataset.open==='1';
      tr.dataset.open=open?'':'1';
      tr.classList.toggle('open',!open);
      d.style.display=open?'none':'';
    });
  });
  function applySearch(){
    const q=document.getElementById('q').value.toLowerCase().trim();
    document.querySelectorAll('tr[data-search]').forEach(tr=>{
      const show=!q||tr.dataset.search.includes(q);
      if(tr.classList.contains('dets')){
        const caseRow=tr.previousElementSibling;
        tr.style.display=(show && caseRow.style.display!=='none') ? (q? '' : (caseRow.dataset.open==='1'?'':'none')) : 'none';
      } else tr.style.display=show?'':'none';
    });
    document.querySelectorAll('table.data tbody tr').forEach(tr=>{
      tr.style.display=(!q||tr.innerText.toLowerCase().includes(q))?'':'none';
    });
    document.querySelectorAll('table.cases td, table.data td').forEach(td=>{
      if(td.querySelector('table'))return;
      if(td.dataset.orig===undefined) td.dataset.orig=td.innerHTML;
      if(!q){td.innerHTML=td.dataset.orig;return;}
      const re=new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),'giu');
      td.innerHTML=td.dataset.orig.replace(re,'<b class="hl">$&</b>');
    });
    document.querySelectorAll('tr.dets').forEach(d=>{
      if(q && d.style.display!=='none'){
        const any=[...d.querySelectorAll('table.data tbody tr')].some(tr=>tr.style.display!=='none');
        if(!any) d.style.display='none';
      }
    });
    document.querySelectorAll('table.cases').forEach(t=>{
      const vis=[...t.querySelectorAll('tr[data-search]')].some(tr=>tr.style.display!=='none');
      t.style.display=(!q||vis)?'':'none';
    });
    document.querySelectorAll('h2').forEach(h=>{
      const t=h.nextElementSibling;
      h.style.display=(!q||t.style.display!=='none')?'':'none';
    });
  }
  document.querySelectorAll('#q')[0].addEventListener('input',applySearch);
  document.getElementById('expand').onclick=()=>{document.querySelectorAll('tr.dets').forEach(d=>d.style.display='');document.querySelectorAll('tr.clickable').forEach(t=>{t.classList.add('open');t.dataset.open='1';});applySearch();};
  document.getElementById('collapse').onclick=()=>{document.querySelectorAll('tr.dets').forEach(d=>d.style.display='none');document.querySelectorAll('tr.clickable').forEach(t=>{t.classList.remove('open');t.dataset.open='';});applySearch();};
  function toKey(s){
    let m=s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/); if(m) return new Date(+m[3],+m[2]-1,+m[1]).getTime();
    m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); if(m) return new Date(+m[1],+m[2]-1,+m[3]).getTime();
    return null;
  }
  document.querySelectorAll('table.sortable').forEach(t=>{
    t.querySelectorAll('th[data-i]').forEach(th=>{
      th.addEventListener('click',()=>{
        const i=+th.dataset.i, tb=t.querySelector('tbody'), rows=[...tb.rows];
        const vals=rows.map(r=>r.cells[i]?r.cells[i].innerText.trim():'').filter(v=>v);
        const mode=vals.every(v=>toKey(v)!==null)?'date':(vals.every(v=>/^[\d.,\-+\s]+$/.test(v))?'num':'str');
        const same=(t.dataset.dirKey===String(i));
        const dir=(same&&t.dataset.dir==='asc')?-1:1;
        rows.sort((a,b)=>{
          const x=a.cells[i]?a.cells[i].innerText.trim():'', y=b.cells[i]?b.cells[i].innerText.trim():'';
          if(mode==='date') return ((toKey(x)||0)-(toKey(y)||0))*dir;
          if(mode==='num') return ((parseFloat(x.replace(',','.'))||0)-(parseFloat(y.replace(',','.'))||0))*dir;
          return dir*x.localeCompare(y,'ru');
        });
        t.dataset.dir=(dir===1)?'asc':'desc';t.dataset.dirKey=String(i);
        t.querySelectorAll('th.sorth').forEach(h=>h.classList.remove('asc','desc'));
        th.classList.add(t.dataset.dir);
        rows.forEach(r=>tb.appendChild(r));
      });
    });
  });
  document.getElementById('unsort').onclick=()=>document.querySelectorAll('table.sortable').forEach(t=>{
    const tb=t.querySelector('tbody');
    [...tb.rows].sort((a,b)=>(+a.dataset.o||0)-(+b.dataset.o||0)).forEach(r=>tb.appendChild(r));
    delete t.dataset.dir;delete t.dataset.dirKey;
    t.querySelectorAll('th.sorth').forEach(h=>h.classList.remove('asc','desc'));
  });
}
</script></body></html>"""
pathlib.Path(OUT_HTML).write_text(HTML, encoding="utf-8")
print("CSS сохранён:", OUT_CSS)
print("JSON сохранён:", OUT_JSON); print("HTML сохранён:", OUT_HTML)
for k, v in data.items():
    print(k, len(v), "блоков")
