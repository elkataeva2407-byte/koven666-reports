from pathlib import Path
import json, re, html, logging, zipfile, os, subprocess, threading, time as time_module, secrets, urllib.parse, base64, urllib.request, urllib.error, asyncio
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from datetime import datetime, time, timedelta
from openpyxl import load_workbook
from telegram import Update, InputFile, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, ContextTypes, filters

BASE = Path(__file__).resolve().parent

# GitHub settings: read DIRECTLY from .env, never from Windows environment.
ENV_PATH = BASE / '.env'

def read_local_env():
    vals = {}
    if ENV_PATH.exists():
        for line in ENV_PATH.read_text(encoding='utf-8-sig').splitlines():
            line=line.strip()
            if not line or line.startswith('#') or '=' not in line:
                continue
            k,v=line.split('=',1)
            k=k.strip(); v=v.strip().strip('"').strip("'")
            vals[k]=v
    return vals

_ENV = read_local_env()
TEMPLATE_PATH = BASE / 'template.html'
DATA_DIR = BASE / 'data'; DATA_DIR.mkdir(exist_ok=True)
REPORTS_DIR = BASE / 'reports'; REPORTS_DIR.mkdir(exist_ok=True)
ACCESS_PATH = DATA_DIR / 'access.json'
HISTORY_PATH = DATA_DIR / 'history_all.json'
RATING_NAMES = [
    'rating_source.xlsx',
    'НК Расчет рейтинга 3.4 — динамика премий и переходов (1).xlsx',
    'НК Расчет рейтинга 3.4 — динамика премий и переходов.xlsx',
    'НК Расчет рейтинга 3.4 — динамика премий и переходов (2).xlsx',
]
CLUSTER_GROUPS = {
    '3':['9','10','11','12','41'],
    '5':['17','18','19','20','42'],
    '6':['21','22','23','24','45'],
    '8':['29','31','32','46','47'],
    '9':['33','34','35','36','49'],
}
ALL_GROUPS = [g for gs in CLUSTER_GROUPS.values() for g in gs]
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger=logging.getLogger('coven666')

# Публикация общего HTML в GitHub Pages.
# Участникам ничего устанавливать не нужно: бот обновляет index.html
# в отдельном GitHub-репозитории, а Telegram отправляет одну и ту же ссылку всем.
#
# Один раз задаются переменные Windows:
#   GITHUB_TOKEN   — fine-grained token с Contents: Read and write для репозитория
#   GITHUB_OWNER   — имя владельца GitHub
#   GITHUB_REPO    — имя репозитория Pages
#   GITHUB_BRANCH  — обычно main
#   GITHUB_PAGES_URL — полный URL Pages, например https://owner.github.io/koven666/
GITHUB_API = 'https://api.github.com'
GITHUB_TOKEN = _ENV.get('GITHUB_TOKEN', '').strip()
GITHUB_OWNER = _ENV.get('GITHUB_OWNER', '').strip()
GITHUB_REPO = _ENV.get('GITHUB_REPO', '').strip()
GITHUB_BRANCH = _ENV.get('GITHUB_BRANCH', 'main').strip() or 'main'
GITHUB_PAGES_URL = _ENV.get('GITHUB_PAGES_URL', '').strip().rstrip('/')
PUBLISHED_NAME = 'index.html'

def github_request(method, path, payload=None):
    # Always read GitHub credentials fresh from the local .env.
    vals = read_local_env()
    token = vals.get('GITHUB_TOKEN', '').strip()
    owner = vals.get('GITHUB_OWNER', '').strip()
    repo = vals.get('GITHUB_REPO', '').strip()
    branch = vals.get('GITHUB_BRANCH', 'main').strip() or 'main'
    if not token or not owner or not repo:
        raise RuntimeError('GitHub API: в .env не заполнены GITHUB_TOKEN/GITHUB_OWNER/GITHUB_REPO')
    url = GITHUB_API + path
    data = json.dumps(payload).encode('utf-8') if payload is not None else None
    headers = {
        'Accept': 'application/vnd.github+json',
        'Authorization': f'Bearer {token}',
        'X-GitHub-Api-Version': '2022-11-28',
        'User-Agent': 'KOVEN-666-Bot',
        'Content-Type': 'application/json',
    }
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=300) as resp:
            raw = resp.read()
            return resp.status, json.loads(raw.decode('utf-8')) if raw else {}
    except urllib.error.HTTPError as e:
        body = e.read().decode('utf-8', errors='replace')
        raise RuntimeError(f'GitHub API {e.code}: {body[:1200]}') from e
    except Exception as e:
        raise RuntimeError(f'Ошибка GitHub API: {e}') from e

def publish_to_github(report_path):
    """Publish current HTML as index.html in GitHub Pages.
    First tries a direct PUT (important for an empty repository). If the file
    already exists, fetches its SHA and updates it.
    """
    vals = read_local_env()
    token = vals.get('GITHUB_TOKEN', '').strip()
    owner = vals.get('GITHUB_OWNER', '').strip()
    repo = vals.get('GITHUB_REPO', '').strip()
    branch = vals.get('GITHUB_BRANCH', 'main').strip() or 'main'
    pages = vals.get('GITHUB_PAGES_URL', '').strip().rstrip('/')
    if not token or not owner or not repo:
        raise RuntimeError('GitHub Pages не настроен: проверь GITHUB_TOKEN/GITHUB_OWNER/GITHUB_REPO в .env')
    size = report_path.stat().st_size
    if size >= 100 * 1024 * 1024:
        raise RuntimeError('HTML больше 100 МБ — GitHub не принимает такой файл через Contents API.')
    api_path = f'/repos/{urllib.parse.quote(owner)}/{urllib.parse.quote(repo)}/contents/index.html'
    logger.info('Публикация HTML в GitHub Pages: %.1f МБ', size / 1024 / 1024)
    content = base64.b64encode(report_path.read_bytes()).decode('ascii')
    payload = {
        'message': f'Обновление КОВЕН 666 — {datetime.now():%d.%m.%Y %H:%M}',
        'content': content,
        'branch': branch,
    }
    # 1) Try create/update without SHA. For a fresh repo this is the correct call.
    try:
        github_request('PUT', api_path, payload)
    except RuntimeError as first_error:
        if 'GitHub API 422' not in str(first_error):
            raise
        # 2) File already exists: get SHA and update it.
        _, existing = github_request('GET', api_path + f'?ref={urllib.parse.quote(branch)}')
        sha = existing.get('sha')
        if not sha:
            raise RuntimeError('GitHub API: не удалось получить SHA существующего index.html')
        payload['sha'] = sha
        github_request('PUT', api_path, payload)
    url = pages + '/' if pages else f'https://{owner}.github.io/{repo}/'
    logger.info('HTML опубликован: %s', url)
    return url

def safe_num(v):
    try:
        if v is None or v=='': return 0.0
        return float(v)
    except: return 0.0

def excel_seconds(v):
    """Convert Excel/OpenPyXL time values to seconds.
    Excel durations are commonly loaded as datetime.timedelta; time values
    can be datetime.time; numeric values are Excel day fractions.
    """
    if v is None or v == '': return 0
    if isinstance(v, timedelta):
        return max(0, int(round(v.total_seconds())))
    if isinstance(v, time):
        return v.hour*3600 + v.minute*60 + v.second + v.microsecond/1_000_000
    if isinstance(v, datetime):
        return v.hour*3600 + v.minute*60 + v.second + v.microsecond/1_000_000
    if isinstance(v, str):
        m=re.match(r'^\s*(\d{1,3}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?\s*$', v)
        if m:
            h,mi,se=map(int,(m.group(1),m.group(2),m.group(3) or 0))
            return h*3600+mi*60+se
    x=safe_num(v)
    return x*86400 if 0 < x < 2 else x

def group_from_name(name):
    m=re.search(r'КЦ\s*УД\s*(\d+)', str(name or ''))
    return m.group(1) if m and m.group(1) in ALL_GROUPS else None

def cluster_for_group(g):
    for c,gs in CLUSTER_GROUPS.items():
        if g in gs:return c
    return None

def gp_assignment(group, date, base_data):
    candidates=[]
    for x in base_data.get('group_gp_history',[]):
        if str(x.get('group_id'))!=str(group): continue
        vf=str(x.get('valid_from') or '1900-01-01')
        vt=x.get('valid_to')
        if vf<=date and (not vt or str(vt)>=date): candidates.append(x)
    if candidates: return candidates[-1]
    for x in base_data.get('group_gp_history',[]):
        if str(x.get('group_id'))==str(group): return x
    return {'gp_id':'','gp_name':'ГП'}

def load_template_data():
    if not TEMPLATE_PATH.exists(): raise FileNotFoundError('Не найден template.html рядом с bot.py')
    text=TEMPLATE_PATH.read_text(encoding='utf-8')
    text=text.replace('id="borya-data"','id="coven-data"').replace("getElementById('borya-data')","getElementById('coven-data')")
    text=re.sub(r'<script[^>]*id=["\']coven-promotion-js["\'][^>]*>.*?</script>', '', text, flags=re.I|re.S)
    m=re.search(r'<script[^>]*id=["\']coven-data["\'][^>]*>', text, re.I)
    if not m: raise ValueError('В template.html не найден блок coven-data')
    b=text.find('</script>',m.end())
    if b<0: raise ValueError('В template.html не закрыт блок coven-data')
    return text, json.loads(text[m.end():b]), m.end(), b

def find_rating():
    for n in RATING_NAMES:
        for p in [BASE/n, DATA_DIR/n]:
            if p.exists(): return p
    return None

def load_rating():
    p=find_rating()
    if not p: return {'_rules': {}}
    wb=load_workbook(p,read_only=True,data_only=True)
    result={}; rules={}
    if 'Расчет рейтинга' in wb.sheetnames:
        ws=wb['Расчет рейтинга']; rows=ws.iter_rows(values_only=True); h=list(next(rows)); ix={str(v):i for i,v in enumerate(h) if v is not None}
        for r in rows:
            name=str(r[ix.get('ФИО')] or '').strip(); g=str(r[ix.get('Группа')] or '').strip()
            if not name or g not in ALL_GROUPS: continue
            gp=str(r[ix['ГП рейтинга']] or '').strip() if 'ГП рейтинга' in ix else ''
            price=safe_num(r[ix['Цена звонка']]) if 'Цена звонка' in ix else 0
            if gp and price: rules.setdefault(gp,{}).update({'price':price})
            cat=safe_num(r[ix['Категория']]) if 'Категория' in ix else 0
            result[name]={
                'income_usp':safe_num(r[ix['Доход за УСП']]) if 'Доход за УСП' in ix else None,
                'rank_in_gp':safe_num(r[ix['Место в ГП']]) if 'Место в ГП' in ix else None,
                'gp_population':safe_num(r[ix['Всего в ГП']]) if 'Всего в ГП' in ix else None,
                'rank_level':int(cat) if cat else None,'category':int(cat) if cat else None,
                'sv_premium_amount':safe_num(r[ix['Выплата СВ']]) if 'Выплата СВ' in ix else 0,
                'nk_premium_amount':safe_num(r[ix['Выплата НК']]) if 'Выплата НК' in ix else 0,
                'next_rank':None,'next_threshold':None,'amount_to_next_category':None,
                'category_boundaries':{},'gp_name':gp,
            }
            if gp:
                rules.setdefault(gp,{}).setdefault('premium',{})
                rules[gp]['premium'][int(cat)]={'sv':safe_num(r[ix['Выплата СВ']]),'nk':safe_num(r[ix['Выплата НК']])}
        # Exact thresholds are reconstructed from the workbook's calculated categories.
        by_gp={}
        for name,rr in result.items(): by_gp.setdefault(rr.get('gp_name',''),[]).append(rr)
        for gp,bucket in by_gp.items():
            mins={c:min((safe_num(x.get('income_usp')) for x in bucket if x.get('rank_level')==c),default=None) for c in range(1,6)}
            rules.setdefault(gp,{})['thresholds']={str(k):v for k,v in mins.items() if v is not None}
            for rr in bucket:
                c=rr.get('rank_level'); inc=safe_num(rr.get('income_usp'))
                rr['category_boundaries']={str(k):v for k,v in mins.items() if v is not None}
                if c and c>1 and mins.get(c-1) is not None:
                    rr['next_rank']=c-1; rr['next_threshold']=mins[c-1]
                    rr['amount_to_next_category']=round(max(0,mins[c-1]-inc),2)
                elif c==1:
                    rr['next_rank']=None; rr['next_threshold']=None; rr['amount_to_next_category']=0
    wb.close(); logger.info('Рейтинг/правила загружены: %s записей, %s ГП',len(result),len(rules)); result['_rules']=rules; return result

def financial_rating_from_history(history, current_date, rating_rules, reference_snapshot=None):
    """Calculate the financial 1–5 rating for the current calendar month.
    The workbook remains the authoritative reference for the 01–12.09 snapshot;
    from the next uploaded day onward the same formula is applied to the accumulated daily Excel history.
    """
    if reference_snapshot and current_date == reference_snapshot.get('date'):
        return [dict(x) for x in reference_snapshot.get('ratings',[])]
    month_start=current_date[:8]+'01'
    facts={}
    for r in history:
        d=str(r.get('metric_date') or '')
        if d<month_start or d>current_date: continue
        oid=str(r.get('operator_id') or r.get('full_name') or '')
        if not oid: continue
        x=facts.get(oid)
        if x is None:
            x={'operator_id':r.get('operator_id'),'full_name':r.get('full_name',''),'group_id':str(r.get('group_id') or ''),'cluster_id':r.get('cluster_id'),'gp_id':r.get('gp_id',''),'gp_name':r.get('gp_name',''),'calls':0.0,'income':0.0,'latest':r}
            facts[oid]=x
        x['calls']+=safe_num(r.get('calls')); x['income']+=safe_num(r.get('income'))
        if d>=str(x['latest'].get('metric_date') or ''): x['latest']=r
    by_gp={}
    rules=rating_rules or {}
    for x in facts.values():
        r=x['latest']; g=str(x['group_id']); gp=str(r.get('gp_name') or x.get('gp_name') or '')
        if g not in ALL_GROUPS or not gp: continue
        price=safe_num(rules.get(gp,{}).get('price',3.5 if 'Недвижимость' in gp else 3.4))
        usp=round(x['income']-x['calls']*price,2)
        item={k:r.get(k) for k in ['cluster_id','gp_id','gp_name','group_id']}
        item.update({'operator_id':x['operator_id'],'full_name':x['full_name'],'income_usp':usp,'calls_month':x['calls'],'income_month':x['income']})
        by_gp.setdefault(gp,[]).append(item)
    ratings=[]
    for gp,bucket in by_gp.items():
        bucket.sort(key=lambda z:(-safe_num(z.get('income_usp')), str(z.get('full_name','')), str(z.get('operator_id',''))))
        n=len(bucket); base=n//5; rem=n%5; bounds={}; idx=0
        for c in range(1,6):
            size=base+(1 if c<=rem else 0); end=idx+size
            for i in range(idx,end):
                bucket[i]['rank_in_gp']=i+1; bucket[i]['gp_population']=n; bucket[i]['category']=c; bucket[i]['rank_level']=c
            idx=end
        for c in range(1,6):
            vals=[safe_num(x['income_usp']) for x in bucket if x['category']==c]
            if vals: bounds[c]=min(vals)
        gp_rule=rules.get(gp,{})
        premiums=gp_rule.get('premium',{})
        for x in bucket:
            c=x['category']; nr=c-1 if c>1 else None; threshold=bounds.get(nr) if nr else None
            x['next_rank']=nr; x['next_threshold']=threshold
            x['amount_to_next_category']=round(max(0,threshold-x['income_usp']),2) if threshold is not None else 0
            p=premiums.get(c,{})
            x['sv_premium_amount']=safe_num(p.get('sv',2000 if c==1 else 1300 if c==2 else 600 if c==3 else 0))
            x['nk_premium_amount']=safe_num(p.get('nk',1100 if c==1 else 100 if c==2 else 50 if c==3 else 0))
            x['category_boundaries']={str(k):v for k,v in bounds.items()}
            x['metric_date']=current_date
            ratings.append(x)
    return ratings

def parse_date(name):
    m=re.findall(r'(\d{2})[.\-_](\d{2})[.\-_](\d{4})',name)
    if not m:
        m=re.findall(r'(\d{4})[.\-_](\d{2})[.\-_](\d{2})',name)
        return f'{m[-1][0]}-{m[-1][1]}-{m[-1][2]}' if m else datetime.now().strftime('%Y-%m-%d')
    return f'{m[-1][2]}-{m[-1][1]}-{m[-1][0]}'

def read_daily(path, date, template_data, rating):
    wb=load_workbook(path,read_only=True,data_only=True); ws=wb[wb.sheetnames[0]]; it=ws.iter_rows(values_only=True); h=list(next(it)); ix={str(v):i for i,v in enumerate(h) if v is not None}
    req=['ID','ФИО','Смен','Рабочее время','Время в звонках','За период','Успешные','Недозвон','Перезвон','Всего','Категория 1-4']
    miss=[x for x in req if x not in ix]
    if miss: wb.close(); raise ValueError('В Excel отсутствуют столбцы: '+', '.join(miss))
    out=[]
    for r in it:
        name=str(r[ix['ФИО']] or '').strip(); g=group_from_name(name); shifts=safe_num(r[ix['Смен']])
        if not g or shifts<=0: continue
        a=gp_assignment(g,date,template_data); c=cluster_for_group(g); rr=rating.get(name)
        out.append({
            'metric_date':date,'operator_id':str(r[ix['ID']] or name),'full_name':name,
            'group_id':g,'cluster_id':c,'gp_id':a.get('gp_id',''),'gp_name':a.get('gp_name',''),
            'calls':safe_num(r[ix['Всего']]),'shifts':shifts,'productive_seconds':excel_seconds(r[ix['Рабочее время']]),'working_seconds':excel_seconds(r[ix['Рабочее время']]),'break_seconds':excel_seconds(r[ix['Время перерывов']]) if 'Время перерывов' in ix else 0,'waiting_seconds':excel_seconds(r[ix['Время ожидания заявки']]) if 'Время ожидания заявки' in ix else 0,
            'income':safe_num(r[ix['За период']]),'successful':safe_num(r[ix['Успешные']]),
            'not_actual':safe_num(r[ix['Недозвон']]),'callbacks':safe_num(r[ix['Перезвон']]),
            'talk_seconds':excel_seconds(r[ix['Время в звонках']]),'workplace_pct':safe_num(r[ix['% у рабочего места']]) if '% у рабочего места' in ix else None,'connect_pct':safe_num(r[ix['% дозвона до человека']]) if '% дозвона до человека' in ix else None,'success_connect_pct':safe_num(r[ix['% успешных по дозвону']]) if '% успешных по дозвону' in ix else None,'success_total_pct':safe_num(r[ix['% успешных общий']]) if '% успешных общий' in ix else None,'crm_category':safe_num(r[ix['Категория 1-4']]) or None,
            'income_usp':rr['income_usp'] if rr else None,'rank_in_gp':rr['rank_in_gp'] if rr else None,
            'gp_population':rr['gp_population'] if rr else None,'rank_level':rr['rank_level'] if rr else None,
            'amount_to_next_category':rr.get('amount_to_next_category') if rr else None,
            'next_rank':rr.get('next_rank') if rr else None,
            'next_threshold':rr.get('next_threshold') if rr else None,
            'category_boundaries':rr.get('category_boundaries',{}) if rr else {},
        })
    wb.close(); return out

def load_history():
    if not HISTORY_PATH.exists(): return []
    try:return json.loads(HISTORY_PATH.read_text(encoding='utf-8'))
    except:return []

def save_history(h): HISTORY_PATH.write_text(json.dumps(h,ensure_ascii=False),encoding='utf-8')

def merge_history(base_daily,current):
    # Start from authoritative reference history; current date is replaced by uploaded Excel.
    cur_dates={str(x['metric_date']) for x in current}; by={(str(x['metric_date']),str(x['operator_id'])):x for x in base_daily if str(x['metric_date']) not in cur_dates}
    for x in current: by[(str(x['metric_date']),str(x['operator_id']))]=x
    h=list(by.values()); h.sort(key=lambda x:(str(x['metric_date']),str(x['operator_id']))); return h

def build_payload(base, history, current_date, rating=None):
    rating=rating or {}; rules=rating.get('_rules',{})
    # The 01–12.09 workbook snapshot is authoritative for that closed calculation.
    ref=[]
    for name,rr in rating.items():
        if name=='_rules': continue
        if not isinstance(rr,dict): continue
        g=next((str(x.get('group_id')) for x in history if str(x.get('full_name','')).strip()==name and str(x.get('group_id')) in ALL_GROUPS),None)
        if not g:
            m=re.search(r'КЦ\s*УД\s*(\d+)',name); g=m.group(1) if m and m.group(1) in ALL_GROUPS else None
        if not g: continue
        gp_name=rr.get('gp_name') or gp_assignment(g,current_date,base).get('gp_name','')
        a=gp_assignment(g,current_date,base)
        ref.append({'metric_date':current_date,'operator_id':name,'full_name':name,'group_id':g,'cluster_id':cluster_for_group(g),'gp_id':a.get('gp_id',''),'gp_name':gp_name,'income_usp':rr.get('income_usp'),'rank_in_gp':rr.get('rank_in_gp'),'gp_population':rr.get('gp_population'),'rank_level':rr.get('rank_level'),'category':rr.get('category'),'amount_to_next_category':rr.get('amount_to_next_category'),'next_rank':rr.get('next_rank'),'next_threshold':rr.get('next_threshold'),'category_boundaries':rr.get('category_boundaries',{}),'sv_premium_amount':rr.get('sv_premium_amount',0),'nk_premium_amount':rr.get('nk_premium_amount',0)})
    # Use the bundled workbook only for 12.09. For later dates calculate from accumulated daily Excel.
    ref_snapshot={'date':'2026-09-12','ratings':ref}
    ratings=financial_rating_from_history(history,current_date,rules,ref_snapshot if current_date=='2026-09-12' and ref else None)
    # If the reference workbook is for 12.09 but some operator rows are absent from daily history, keep them in financial rating.
    latest={}; per={}
    for r in history:
        if str(r.get('metric_date'))<=current_date:
            latest[str(r.get('operator_id'))]=r; per.setdefault(str(r.get('operator_id')),[]).append(r)
    by_name={str(r.get('full_name','')).strip():r for r in ratings}
    operators=[]
    for oid,r in latest.items():
        rs=sorted(per[oid],key=lambda x:str(x.get('metric_date','')),reverse=True)[:7]
        calls=sum(safe_num(x.get('calls')) for x in rs); suc=sum(safe_num(x.get('successful')) for x in rs); shifts=sum(safe_num(x.get('shifts')) for x in rs)
        rr=by_name.get(str(r.get('full_name','')).strip())
        operators.append({'operator_id':oid,'full_name':r.get('full_name'),'group_id':r.get('group_id'),'cluster_id':r.get('cluster_id'),'gp_id':r.get('gp_id'),'gp_name':r.get('gp_name'),'rank_in_gp':rr.get('rank_in_gp') if rr else None,'gp_population':rr.get('gp_population') if rr else None,'income_usp':rr.get('income_usp') if rr else None,'amount_to_next_category':rr.get('amount_to_next_category') if rr else None,'category_boundaries':rr.get('category_boundaries',{}) if rr else {},'transition':rr.get('amount_to_next_category') if rr else None,'rank_level':rr.get('rank_level') if rr else None,'financial_rank':rr.get('rank_level') if rr else None,'category':rr.get('category') if rr else None,'crm_category':r.get('crm_category'),'next_rank':rr.get('next_rank') if rr else None,'next_threshold':rr.get('next_threshold') if rr else None,'conversion_7_shifts':suc/calls if calls else None,'calls_7_shifts':calls,'successful_7_shifts':suc,'shifts_7':shifts})
    rating_history=[{k:r.get(k) for k in ['metric_date','operator_id','group_id','gp_id','gp_name','rank_in_gp','gp_population','income_usp','amount_to_next_category','rank_level','crm_category']} | {'category':r.get('rank_level'),'crm_category':r.get('crm_category'),'run_id':0,'rank_observed_at':r.get('metric_date'),'department_member':True} for r in history if r.get('rank_level') is not None]
    data=dict(base)
    data.update({'source':'coven666_daily_excel_accumulated','audience':'nk','cluster_id':'5','report_level':'department','clusters':['3','5','6','8','9'],'cluster_groups':CLUSTER_GROUPS,'groups':ALL_GROUPS,'group_cluster':{g:c for c,gs in CLUSTER_GROUPS.items() for g in gs},'daily':history,'department_daily':history,'gp_daily':history,'operators':operators,'rating_history':rating_history,'department_rating_history':rating_history,'premium_ratings':ratings,'rating_snapshot':ratings,'rating_rules':rules,'premium_snapshots':base.get('premium_snapshots',[]),'rating_run_id':'daily-accumulated','rating_period':{'from':current_date[:8]+'01','to':current_date},'range':{'min':min([str(x.get('metric_date')) for x in history if x.get('metric_date')]+[current_date]),'max':max([str(x.get('metric_date')) for x in history if x.get('metric_date')]+[current_date])},'default_period':{'from':current_date,'to':current_date},'today':current_date,'department_name':'КОВЕН 666 · НК · Все кластеры','premium_data_source':'Excel за день → накопительная история → финансовый рейтинг 1–5 по ГП; CRM 1–4 отдельно'})
    return data

def inject_promotion_section(template):
    return re.sub(r'<script[^>]*id=["\']coven-promotion-js["\'][^>]*>.*?</script>', '', template, flags=re.I|re.S)

def inject_rating_summary(template, payload, current_date):
    current=[r for r in payload.get('premium_ratings',[]) if str(r.get('metric_date') or '')==str(current_date) and str(r.get('group_id') or '') in ALL_GROUPS]
    if not current:
        current=[r for r in payload.get('daily',[]) if str(r.get('metric_date') or '')==str(current_date) and str(r.get('group_id') or '') in ALL_GROUPS]
    CALL_GROUPS={'Солянка МЮ':['11','17','21','29'],'Солянка Недвижимость':['12','18','23','36'],'Солянка Прочее':['22','34','41','42','47'],'СПК':['10','20','24','32','35','45','46','49'],'Партнеры солянка':['9','19','31','33']}
    group_to_gp={g:gp for gp,groups in CALL_GROUPS.items() for g in groups}
    by_group={}
    for r in current:
        g=str(r.get('group_id') or '')
        if not g: continue
        gp=group_to_gp.get(g,str(r.get('gp_name') or r.get('gp_id') or 'ГП'))
        try: rank=int(float(r.get('rank_level'))) if r.get('rank_level') not in (None,'') else None
        except: rank=None
        rec=by_group.setdefault(g,{'group':g,'gp':gp,'people':0,'ranks':{i:0 for i in range(1,6)},'usp':0.0,'rank_sum':0.0,'rank_n':0})
        rec['people']+=1; rec['usp']+=safe_num(r.get('income_usp'))
        if rank in rec['ranks']: rec['ranks'][rank]+=1; rec['rank_sum']+=rank; rec['rank_n']+=1
    def team_row(rec):
        total=rec['people']; rr=rec['ranks']; top=rr[1]+rr[2]+rr[3]
        cells=''.join(f'<td class="rs-r{i}"><b>{rr[i]}</b><small>{(rr[i]/total*100 if total else 0):.0f}%</small></td>' for i in range(1,6))
        avg=rec['rank_sum']/rec['rank_n'] if rec['rank_n'] else 0
        return f'<tr><td><b>Группа {html.escape(rec["group"])} · {html.escape(rec["gp"])}</b></td><td>{total}</td>{cells}<td><b>{top}</b><small>{(top/total*100 if total else 0):.0f}%</small></td><td class="money-usp"><b>{rec["usp"]:,.0f} ₽</b></td><td>{avg:.2f}</td></tr>'
    buttons=[]; panels=[]
    for idx,gp in enumerate(CALL_GROUPS):
        groups=[g for g in CALL_GROUPS[gp] if g in by_group]
        if not groups: continue
        pid=f'gp-panel-{idx}'
        buttons.append(f'<button type="button" class="gp-switch{" active" if not buttons else ""}" data-gp-panel="{pid}">{html.escape(gp)}</button>')
        rows=''.join(team_row(by_group[g]) for g in groups)
        display='block' if not panels else 'none'
        panels.append(f'<div id="{pid}" class="gp-panel" style="display:{display};"><div class="section-head"><div><p class="eyebrow">ГП</p><h2>{html.escape(gp)}</h2></div><p>Сравнение всех команд внутри ГП</p></div><div class="panel"><div class="scroll"><table class="rating-table"><thead><tr><th>Команда</th><th>Чел.</th><th>Ранг 1</th><th>Ранг 2</th><th>Ранг 3</th><th>Ранг 4</th><th>Ранг 5</th><th>ТОП 1–3</th><th>Доля ТОП 1–3</th><th>Доход за УСП</th><th>Средний финансовый ранг</th></tr></thead><tbody>{rows}</tbody></table></div></div></div>')
    buttons_html=''.join(buttons) or '<span>Нет данных</span>'; panels_html=''.join(panels) or '<p>Нет данных за выбранную дату.</p>'
    css='''<style id="coven-rating-summary-css">#rating-summary .summary-hero{background:linear-gradient(120deg,#edf8f0,#f9fcf9);border:1px solid #cee5d4;border-radius:14px;padding:22px;margin:12px 0 18px}#rating-summary .rating-table{width:100%;border-collapse:separate;border-spacing:0;font-size:12px;min-width:1180px}#rating-summary .rating-table th,#rating-summary .rating-table td{padding:10px 9px;border-bottom:1px solid var(--line);text-align:right;white-space:nowrap}#rating-summary .rating-table th:first-child,#rating-summary .rating-table td:first-child{text-align:left}#rating-summary .rating-table th{background:#fbfdfb;color:var(--muted);font-size:9px;text-transform:uppercase}#rating-summary .rating-table .rs-r1{background:#eaf7ee;color:#176c42}#rating-summary .rating-table .rs-r2{background:#edf5ff;color:#286aa8}#rating-summary .rating-table .rs-r3{background:#fff6df;color:#946c19}#rating-summary .rating-table .rs-r4{background:#fdeeed;color:#b44842}#rating-summary .rating-table .rs-r5{background:#f3eafa;color:#7b3e9e}#rating-summary .rating-table td small{display:block;color:var(--muted);font-size:9px;margin-top:2px}#rating-summary .rating-table .money-usp{background:#f5fbf7;color:var(--strong)}#rating-summary .gp-switch{border:1px solid #cfe0d4;background:#f7faf8;color:#254d37;border-radius:10px;padding:9px 14px;margin:0 8px 10px 0;cursor:pointer;font-weight:700}#rating-summary .gp-switch.active{background:#dff2e5;border-color:#70ad82}</style>'''
    js='''<script id="coven-gp-switch-js">(function(){var root=document.getElementById('rating-summary');if(!root)return;var bs=root.querySelectorAll('[data-gp-panel]');function show(id){root.querySelectorAll('.gp-panel').forEach(function(p){p.style.display=p.id===id?'block':'none';});root.querySelectorAll('.gp-switch').forEach(function(b){b.classList.toggle('active',b.getAttribute('data-gp-panel')===id);});}bs.forEach(function(b){b.addEventListener('click',function(){show(b.getAttribute('data-gp-panel'));});});})();</script>'''
    section=f'<section id="rating-summary" class="view"><div class="summary-hero"><p class="eyebrow">Рейтинг · ГП</p><h1>ГП · команды</h1><p>Открой ГП и сравни все команды внутри него по одинаковым показателям.</p><p class="summary-note">Данные на {html.escape(current_date)}. Доход считается как сумма <b>Доход за УСП</b> всех сотрудников команды.</p></div><div class="panel" style="margin-bottom:16px"><div style="padding:14px 16px"><p class="eyebrow">Выберите ГП</p>{buttons_html}</div></div>{panels_html}</section>'
    template=template.replace('<button type="button" data-view-target="premiums">Премии</button>','<button type="button" data-view-target="premiums">Премии</button><button type="button" data-view-target="rating-summary">ГП · команды</button>',1)
    template=template.replace('<section id="premiums" class="view">',section+'<section id="premiums" class="view">',1)
    template=template.replace('</head>',css+'</head>',1); template=template.replace('</body>',js+'</body>',1)
    return template

def render(template, payload, current_date):
    template = template.replace("getElementById('borya-data')", "getElementById('coven-data')").replace('getElementById("borya-data")', 'getElementById("coven-data")').replace('id="borya-data"', 'id="coven-data"')
    template = inject_rating_summary(template, payload, current_date)
    template = inject_promotion_section(template)
    marker = '<script id="coven-data" type="application/json">'
    a = template.find(marker)
    if a < 0: raise ValueError('В шаблоне не найден coven-data')
    a += len(marker); b = template.find('</script>', a)
    if b < 0: raise ValueError('В шаблоне не закрыт coven-data')
    txt = template[:a] + json.dumps(payload, ensure_ascii=False, separators=(',', ':')) + template[b:]
    txt = re.sub(r'<script[^>]*id=["\']coven-promotion-js["\'][^>]*>.*?</script>', '', txt, flags=re.I|re.S)
    return txt.replace('<title>Отчёт НК · Подразделение Солянки</title>', f'<title>КОВЕН 666 · НК · Все кластеры · {current_date}</title>', 1)

def generate(path):
    template,base,_,_=load_template_data(); date=parse_date(path.name); rating=load_rating(); current=read_daily(path,date,base,rating)
    hist_ref=[r for r in base.get('daily',[]) if r.get('metric_date')]
    prior=load_history()
    seed=hist_ref+prior
    history=merge_history(seed,current)
    save_history(history)
    payload=build_payload(base,history,date,rating); out=REPORTS_DIR/f'КОВЕН_666_НК_Все_кластеры_{date.replace("-","_")}.html'; out.write_text(render(template,payload,date),encoding='utf-8')
    return out,current

# Закрытый режим: только администратор и добавленные участники.
# Только администратор может загружать Excel, либо временный загрузчик,
# которому администратор выдал право на 7/14 дней (или другой срок).
ADMIN_ID=int(os.getenv('ADMIN_ID','0') or 0)

def load_access():
    data={'participants':[], 'uploaders':{}}
    try:
        if ACCESS_PATH.exists():
            x=json.loads(ACCESS_PATH.read_text(encoding='utf-8'))
            if isinstance(x,dict): data.update(x)
    except Exception:
        logger.exception('Не удалось прочитать access.json')
    # Старый ALLOWED_USERS остаётся совместимым: считаем их участниками.
    env_users={int(x.strip()) for x in os.getenv('ALLOWED_USERS','').split(',') if x.strip().isdigit()}
    data['participants']=sorted(set(int(x) for x in data.get('participants',[]) if str(x).isdigit()) | env_users | ({ADMIN_ID} if ADMIN_ID else set()))
    up={}
    for k,v in (data.get('uploaders') or {}).items():
        try:
            until=str(v);
            if until: up[str(int(k))]=until
        except Exception: pass
    data['uploaders']=up
    return data

def save_access(data):
    ACCESS_PATH.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')

ACCESS=load_access()

def refresh_access():
    global ACCESS
    ACCESS=load_access()
    return ACCESS

def is_admin(update):
    return bool(ADMIN_ID and update.effective_user and update.effective_user.id==ADMIN_ID and is_private(update))

def authorized(update):
    if not is_private(update): return False
    uid=getattr(update.effective_user,'id',None)
    data=refresh_access(); ok=bool(uid is not None and uid in set(data.get('participants',[])))
    if not ok: logger.warning('Заблокирован пользователь Telegram ID=%s',uid)
    return ok

def can_upload(update):
    if not authorized(update): return False
    uid=str(update.effective_user.id)
    if is_admin(update): return True
    until=refresh_access().get('uploaders',{}).get(uid)
    if not until: return False
    try: return datetime.now() < datetime.fromisoformat(until)
    except Exception: return False

def access_help_text():
    return ('🔐 Управление доступом\n\n'
            '/add_user ID — добавить участника\n'
            '/remove_user ID — убрать участника\n'
            '/grant_upload ID 7 — дать загрузку Excel на 7 дней\n'
            '/grant_upload ID 14 — дать загрузку Excel на 14 дней\n'
            '/revoke_upload ID — забрать право загрузки\n'
            '/users — список доступа')

def is_private(update):
    return bool(update.effective_chat and update.effective_chat.type == 'private')

def authorized(update):
    if not is_private(update): return False
    uid=getattr(update.effective_user,'id',None)
    ok=bool(uid is not None and uid in set(refresh_access().get('participants',[])))
    if not ok:
        logger.warning('Заблокирован пользователь Telegram ID=%s', uid)
    return ok

async def start(update:Update,context:ContextTypes.DEFAULT_TYPE):
    if not authorized(update): return
    await update.message.reply_text('🔮 КОВЕН 666\n\n' + ('У тебя есть право загружать дневной Excel.\n' if can_upload(update) else 'Ты можешь просматривать отчёты. Загрузка Excel доступна только администратору или назначенному загрузчику.\n') + '\nОтчёт по кластерам 3, 5, 6, 8, 9. Кластер 6 содержит расширенную аналитику.')

async def receive_excel(update:Update,context:ContextTypes.DEFAULT_TYPE):
    if not can_upload(update): return
    doc=update.message.document
    if not doc:return
    name=doc.file_name or 'report.xlsx'; path=DATA_DIR/name
    await update.message.reply_text('📥 Получил Excel.\n🔮 Формирую общий HTML-отчёт по всем кластерам...')
    try:
        tg=await doc.get_file(); await tg.download_to_drive(custom_path=str(path)); logger.info('Excel скачан. Генерация HTML вынесена в отдельный поток, Telegram не блокируется.'); out,records=await asyncio.to_thread(generate,path)
        c6=[r for r in records if str(r.get('cluster_id'))=='6']; c6_calls=sum(safe_num(r.get('calls')) for r in c6); c6_success=sum(safe_num(r.get('successful')) for r in c6); c6_shifts=sum(safe_num(r.get('shifts')) for r in c6); c6_work=sum(safe_num(r.get('working_seconds')) for r in c6)
        c6_conv=(c6_success/c6_calls*100) if c6_calls else 0; c6_work_shift=(c6_work/c6_shifts) if c6_shifts else 0
        slogans=['Ну что, красотки, смотрим отчёт и летим к звёздам ✨','Сегодня не ждём чуда — сами делаем результат 🚀','Звёзды сами не зажигаются. Работаем и забираем своё ⭐','Ещё один день — ещё один шаг к топу 🔥','Красотки, держим темп и забираем лучший результат 💫','План увидели — цель поставили — результат сделали 💪','Сегодня работаем так, чтобы завтра собой гордиться 🌟']
        try: slogan=slogans[(int(parse_date(name)[-2:])-1)%len(slogans)]
        except Exception: slogan=slogans[0]
        def fmt_hms(sec):
            sec=max(0,int(round(sec))); h=sec//3600; m=(sec%3600)//60; ss=sec%60; return f'{h:02d}:{m:02d}:{ss:02d}'
        report_url=await asyncio.to_thread(publish_to_github,out)
        text=(f'✅ HTML-отчёт готов.\n\n📅 Дата: {parse_date(name)}\n👥 Операторов в 5 кластерах: {len(records)}\n\n6 кластер:\nКонверсия — {c6_conv:.2f}%\nРабочее время / смену — {fmt_hms(c6_work_shift)}\n\n{slogan}')
        recipients=sorted(refresh_access().get('participants',[])) or [getattr(update.effective_user,'id',0)]
        keyboard=InlineKeyboardMarkup([[InlineKeyboardButton('📊 ОТКРЫТЬ HTML-ОТЧЁТ',url=report_url)]])
        for uid in recipients:
            if uid:
                try: await context.bot.send_message(chat_id=uid,text=text,reply_markup=keyboard)
                except Exception: logger.exception('Не удалось отправить отчёт пользователю %s',uid)
    except Exception as e:
        logger.exception('Ошибка создания/публикации отчёта'); await update.message.reply_text('❌ Не удалось опубликовать HTML-отчёт.\n\nПричина: '+repr(e))


async def add_user(update:Update,context:ContextTypes.DEFAULT_TYPE):
    if not is_admin(update): return
    if not context.args or not context.args[0].isdigit():
        await update.message.reply_text('Использование: /add_user 123456789'); return
    uid=int(context.args[0]); data=refresh_access(); data['participants']=sorted(set(data.get('participants',[]))|{uid}); save_access(data)
    await update.message.reply_text(f'✅ Участник {uid} добавлен. Он может получать отчёты, но не может загружать Excel.')

async def remove_user(update:Update,context:ContextTypes.DEFAULT_TYPE):
    if not is_admin(update): return
    if not context.args or not context.args[0].isdigit():
        await update.message.reply_text('Использование: /remove_user 123456789'); return
    uid=int(context.args[0]); data=refresh_access(); data['participants']=[x for x in data.get('participants',[]) if int(x)!=uid]; data['uploaders'].pop(str(uid),None); save_access(data)
    await update.message.reply_text(f'🚫 Участник {uid} удалён из доступа.')

async def grant_upload(update:Update,context:ContextTypes.DEFAULT_TYPE):
    if not is_admin(update): return
    if len(context.args)<2 or not context.args[0].isdigit() or not context.args[1].isdigit():
        await update.message.reply_text('Использование: /grant_upload 123456789 7'); return
    uid=int(context.args[0]); days=max(1,min(int(context.args[1]),31)); data=refresh_access(); data['participants']=sorted(set(data.get('participants',[]))|{uid}); until=datetime.now()+timedelta(days=days); data['uploaders'][str(uid)]=until.isoformat(timespec='seconds'); save_access(data)
    await update.message.reply_text(f'✅ {uid} назначен загрузчиком на {days} дней. До {until:%d.%m.%Y %H:%M}.')

async def revoke_upload(update:Update,context:ContextTypes.DEFAULT_TYPE):
    if not is_admin(update): return
    if not context.args or not context.args[0].isdigit():
        await update.message.reply_text('Использование: /revoke_upload 123456789'); return
    uid=int(context.args[0]); data=refresh_access(); data['uploaders'].pop(str(uid),None); save_access(data)
    await update.message.reply_text(f'✅ Право загрузки у {uid} снято. Доступ к отчётам остаётся.')

async def users_cmd(update:Update,context:ContextTypes.DEFAULT_TYPE):
    if not is_admin(update): return
    data=refresh_access(); now=datetime.now(); lines=['👥 Участники:']
    for uid in data.get('participants',[]):
        role='админ' if int(uid)==ADMIN_ID else 'участник'
        until=data.get('uploaders',{}).get(str(uid));
        if until:
            try:
                dt=datetime.fromisoformat(until); role += f', загрузка до {dt:%d.%m.%Y %H:%M}' if dt>now else ', загрузка истекла'
            except: pass
        lines.append(f'• {uid} — {role}')
    await update.message.reply_text('\\n'.join(lines))

async def access_cmd(update:Update,context:ContextTypes.DEFAULT_TYPE):
    if not is_admin(update): return
    await update.message.reply_text(access_help_text())

async def my_id(update:Update,context:ContextTypes.DEFAULT_TYPE):
    if not authorized(update): return
    uid=getattr(update.effective_user,'id',None)
    await update.message.reply_text(f'Ваш Telegram ID: {uid}')

async def help_cmd(update:Update,context:ContextTypes.DEFAULT_TYPE):
    if not authorized(update): return
    await update.message.reply_text('📊 КОВЕН 666\n\nExcel → общий HTML по кластерам 3/5/6/8/9.\nВнутри — группы, ГП, динамика, сотрудники и рейтинги.\nДля Кластера 6: CRM 1–4, финансовый ранг 1–5, 7 смен, переходы и сумма до следующего ранга.')

async def reports(update:Update,context:ContextTypes.DEFAULT_TYPE):
    if not authorized(update): return
    fs=sorted(REPORTS_DIR.glob('*.html'),key=lambda p:p.stat().st_mtime,reverse=True)
    await update.message.reply_text('📊 HTML-ОТЧЁТЫ\n\n'+('\n'.join(f'{i}. {p.name}' for i,p in enumerate(fs[:20],1)) if fs else 'Пока нет отчётов.'))

def main():
    token=__import__('os').getenv('BOT_TOKEN','YOUR_BOT_TOKEN')
    env = read_local_env()
    if env.get('GITHUB_TOKEN') and env.get('GITHUB_OWNER') and env.get('GITHUB_REPO'):
        logger.info('GitHub Pages настроен: %s', env.get('GITHUB_PAGES_URL', '').rstrip('/'))
    else:
        logger.warning('GitHub Pages не настроен: задайте GITHUB_TOKEN, GITHUB_OWNER и GITHUB_REPO.')
    app=Application.builder().token(token).build(); app.add_handler(CommandHandler('start',start)); app.add_handler(CommandHandler('help',help_cmd)); app.add_handler(CommandHandler('id',my_id)); app.add_handler(CommandHandler('add_user',add_user)); app.add_handler(CommandHandler('remove_user',remove_user)); app.add_handler(CommandHandler('grant_upload',grant_upload)); app.add_handler(CommandHandler('revoke_upload',revoke_upload)); app.add_handler(CommandHandler('users',users_cmd)); app.add_handler(CommandHandler('access',access_cmd)); app.add_handler(CommandHandler('reports',reports)); app.add_handler(MessageHandler(filters.Document.ALL,receive_excel))
    print('==================================='); print('🔮 КОВЕН 666 ЗАПУЩЕН'); print('Telegram → Excel → ОБЩИЙ HTML'); print('Кластеры 3 · 5 · 6 · 8 · 9'); print('WEB → общий отчёт публикуется в GitHub Pages'); print('==================================='); app.run_polling()
if __name__=='__main__': main()
