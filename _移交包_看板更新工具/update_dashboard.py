#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
服饰中心创意供给监控看板 · 通用更新脚本（v4.5）
作者：傻福🦞 (代 Javi)

用法：
    python3 update_dashboard.py 0506              # 用 0506 命名的 CSV 更新
    python3 update_dashboard.py 0506 --date 2026-05-05   # 自定义看板上的日期
    python3 update_dashboard.py 0506 --config /path/config.txt   # 自定义配置路径

兼容性：
- 自动适配文件名：「九图XXXX.csv」/「朋友圈九图XXXX.csv」、「视频号XXXX.csv」/「视频号流量XXXX.csv」等
- 自动适配CSV列结构变化（动态定位"消耗(万元)"列）
- 配置通过 config.txt 注入，不再hard-code路径
"""
import csv, json, os, re, sys, argparse
from datetime import datetime, timedelta

try:
    import openpyxl
except ImportError:
    print("❌ 缺少依赖 openpyxl，请运行：pip3 install openpyxl")
    sys.exit(1)

# ===========================================================================
# 配置加载
# ===========================================================================
def load_config(config_path):
    """从 config.txt 读取配置"""
    if not os.path.exists(config_path):
        print(f"❌ 配置文件不存在：{config_path}")
        print(f"   请参考 config.example.txt 创建你自己的 config.txt")
        sys.exit(1)
    cfg = {}
    with open(config_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#'): continue
            if '=' in line:
                k, v = line.split('=', 1)
                cfg[k.strip()] = v.strip()
    required = ['SRC_DIR', 'HTML_FILE', 'KA_FILE']
    for k in required:
        if k not in cfg:
            print(f"❌ 配置缺失：{k}")
            sys.exit(1)
    return cfg

# ===========================================================================
# 命令行参数
# ===========================================================================
parser = argparse.ArgumentParser(description='服饰中心创意供给监控看板更新脚本')
parser.add_argument('mmdd', help='文件名日期（4位），如 0506')
parser.add_argument('--date', help='看板显示的数据日期（默认=文件名日期-1天），如 2026-05-05')
parser.add_argument('--config', default=None, help='配置文件路径（默认=同目录下 config.txt）')
args = parser.parse_args()

mmdd = args.mmdd
if not (len(mmdd) == 4 and mmdd.isdigit()):
    print(f"❌ 日期格式错误：{mmdd}，应为 4 位数字如 0506")
    sys.exit(1)

# 配置文件路径
script_dir = os.path.dirname(os.path.abspath(__file__))
config_path = args.config or os.path.join(script_dir, 'config.txt')
cfg = load_config(config_path)

SRC = cfg['SRC_DIR']
HTML = cfg['HTML_FILE']
KA_FILE = cfg['KA_FILE']

# 数据日期：文件名日期 - 1天
if args.date:
    new_date = args.date
else:
    year = datetime.now().year
    file_date = datetime(year, int(mmdd[:2]), int(mmdd[2:]))
    data_date = file_date - timedelta(days=1)
    new_date = data_date.strftime('%Y-%m-%d')

print(f"📅 文件名日期: {mmdd}  →  看板数据日期: {new_date}")
print(f"📁 源数据目录: {SRC}")
print(f"📄 看板文件:   {HTML}")
print()

# ===========================================================================
# 通用工具
# ===========================================================================
def fmt_wan(v):
    try: n = float(v)
    except: return '-'
    return f"{n:,.1f}"

def fmt_int(v):
    try: n = float(v)
    except: return '-'
    return f"{int(round(n)):,}"

def fmt_pct_pair(v):
    if str(v) in ('∞','-∞'): return ('∞','up')
    try: n = float(v)
    except: return ('0%','flat')
    cls = 'up' if n>0 else ('down' if n<0 else 'flat')
    sign = '+' if n>0 else ''
    return (f"{sign}{n:.0f}%", cls)

def parse_num(v):
    if v in ('~','-','',None): return None
    try: return float(v)
    except: return None

def read_csv_(path):
    with open(path,'r',encoding='utf-8-sig') as f:
        return list(csv.reader(f))

def find_file(src_dir, mmdd, candidates):
    """从多个候选文件名里找到第一个存在的"""
    for name in candidates:
        full = os.path.join(src_dir, name.format(mmdd=mmdd))
        if os.path.exists(full):
            return full
    print(f"❌ 找不到文件，候选：{[c.format(mmdd=mmdd) for c in candidates]}")
    sys.exit(1)

def find_cost_col(header):
    """动态定位'消耗(万元)'列"""
    for i, h in enumerate(header):
        if h.strip() == '消耗(万元)':
            return i
    return None

def wk_pct_fmt(v):
    """格式化周同比值"""
    if str(v) in ('∞','-∞'): return '∞'
    try: n = float(v)
    except: return '-'
    sign = '+' if n>0 else ''
    return f"{sign}{n:.0f}%"

# ===========================================================================
# 文件名兼容（支持多种命名）
# ===========================================================================
F_HY     = find_file(SRC, mmdd, ['分行业{mmdd}.csv'])
F_HY_WK  = find_file(SRC, mmdd, ['分行业周同比{mmdd}.csv'])
F_LK     = find_file(SRC, mmdd, ['分链路{mmdd}.csv'])
F_LK_WK  = find_file(SRC, mmdd, ['分链路周同比{mmdd}.csv'])
F_PYQ    = find_file(SRC, mmdd, ['九图{mmdd}.csv', '朋友圈九图{mmdd}.csv'])
F_PYQ_WK = find_file(SRC, mmdd, ['九图周同比{mmdd}.csv', '朋友圈九图周同比{mmdd}.csv'])
F_SPH    = find_file(SRC, mmdd, ['视频号{mmdd}.csv', '视频号流量{mmdd}.csv'])
F_SPH_WK = find_file(SRC, mmdd, ['视频号周同比{mmdd}.csv', '视频号流量周同比{mmdd}.csv'])
F_HB     = find_file(SRC, mmdd, ['红黑榜{mmdd}.csv'])

# ===========================================================================
# KA 白名单
# ===========================================================================
wb = openpyxl.load_workbook(KA_FILE, data_only=True)
ws = wb['Sheet1']
KA = set()
for i,row in enumerate(ws.iter_rows(values_only=True)):
    if i==0: continue
    if row and row[0]: KA.add(str(row[0]).strip())
print(f'KA 白名单客户数: {len(KA)}')

# ===========================================================================
# 视图 01 · Summary（分行业整体行）+ 周同比
# ===========================================================================
hy_rows = read_csv_(F_HY)
all_row = hy_rows[1]
hy_wk = read_csv_(F_HY_WK)
all_wk = hy_wk[1]
hy_wk_by_name = {r[0]: r for r in hy_wk[1:] if r and r[0]}

lk_wk = read_csv_(F_LK_WK)
lk_wk_by_name = {r[0]: r for r in lk_wk[1:] if r and r[0]}

def fmt_card(label, sublabel, val, pct_raw, wk_pct_raw, typ='int'):
    v_fmt = fmt_wan(val) if typ=='wan' else fmt_int(val)
    try: pct = float(pct_raw)
    except: pct = 0
    change = f"{'+' if pct>0 else ''}{pct:.0f}%"
    cls = 'up' if pct>0 else ('down' if pct<0 else 'flat')
    card = {'label':label,'sublabel':sublabel,'value':v_fmt,'change':change,'cls':cls,
            'wk': wk_pct_fmt(wk_pct_raw)}
    if typ=='wan': card['unit']='万'
    return card

v1_summary = [
    fmt_card('服饰竞价消耗（不包含全域通）','消耗（万元）',all_row[1],all_row[2], all_wk[2], 'wan'),
    fmt_card('曝光创意唯一性ID数','有曝光的去重创意数',all_row[7],all_row[8], all_wk[8], 'int'),
    fmt_card('消耗创意唯一性ID数','有消耗的去重创意数',all_row[9],all_row[10], all_wk[10], 'int'),
    fmt_card('在线创意唯一性ID数','在线去重创意数',all_row[11],all_row[12], all_wk[12], 'int'),
    fmt_card('新建创意唯一性ID数','新建去重创意数',all_row[3],all_row[4], all_wk[4], 'int'),
    fmt_card('新建创意中新增唯一性ID数','新建全新创意数',all_row[5],all_row[6], all_wk[6], 'int'),
]

# ===========================================================================
# 视图01 模块A · 朋友圈九图（动态列定位，兼容多种CSV结构）
# ===========================================================================
pyq_rows = read_csv_(F_PYQ)
pyq_wk_rows = read_csv_(F_PYQ_WK)
pyq_cost_col = find_cost_col(pyq_rows[0]) or 2
pyq_wk_cost_col = find_cost_col(pyq_wk_rows[0]) or 2
pyq_all = pyq_rows[1]
pyq_wk_all = pyq_wk_rows[1]
pyq_summary = [
    fmt_card('朋友圈九图消耗','消耗（万元）', pyq_all[pyq_cost_col], pyq_all[pyq_cost_col+1], pyq_wk_all[pyq_wk_cost_col+1], 'wan'),
    fmt_card('朋友圈九图曝光创意','曝光创意唯一性ID数', pyq_all[pyq_cost_col+2], pyq_all[pyq_cost_col+3], pyq_wk_all[pyq_wk_cost_col+3], 'int'),
]

# ===========================================================================
# 视图01 模块B · 视频号（动态列定位）
# ===========================================================================
sph_rows = read_csv_(F_SPH)
sph_wk_rows = read_csv_(F_SPH_WK)
sph_cost_col = find_cost_col(sph_rows[0]) or 1
sph_wk_cost_col = find_cost_col(sph_wk_rows[0]) or 1
sph_all = sph_rows[1]
sph_wk_all = sph_wk_rows[1]
sph_summary = [
    fmt_card('视频号流量消耗','消耗（万元）', sph_all[sph_cost_col], sph_all[sph_cost_col+1], sph_wk_all[sph_wk_cost_col+1], 'wan'),
    fmt_card('视频号曝光创意','曝光创意唯一性ID数', sph_all[sph_cost_col+2], sph_all[sph_cost_col+3], sph_wk_all[sph_wk_cost_col+3], 'int'),
]

# ===========================================================================
# 视图01 模块C · 分链路明细
# ===========================================================================
def build_6metric_row(r, wk=None):
    cost_r, cost_c = fmt_pct_pair(r[2])
    new_r, new_c = fmt_pct_pair(r[4])
    fresh_r, fresh_c = fmt_pct_pair(r[6])
    expo_r, expo_c = fmt_pct_pair(r[8])
    used_r, used_c = fmt_pct_pair(r[10])
    online_r, online_c = fmt_pct_pair(r[12])
    d = {
        'name': r[0],
        'cost': fmt_wan(r[1]), 'cost_r': cost_r, 'cost_c': cost_c,
        'expo': fmt_int(r[7]), 'expo_r': expo_r, 'expo_c': expo_c,
        'used': fmt_int(r[9]), 'used_r': used_r, 'used_c': used_c,
        'online': fmt_int(r[11]), 'online_r': online_r, 'online_c': online_c,
        'new': fmt_int(r[3]), 'new_r': new_r, 'new_c': new_c,
        'fresh': fmt_int(r[5]), 'fresh_r': fresh_r, 'fresh_c': fresh_c,
    }
    if wk:
        d['cost_w']   = wk_pct_fmt(wk[2])
        d['new_w']    = wk_pct_fmt(wk[4])
        d['fresh_w']  = wk_pct_fmt(wk[6])
        d['expo_w']   = wk_pct_fmt(wk[8])
        d['used_w']   = wk_pct_fmt(wk[10])
        d['online_w'] = wk_pct_fmt(wk[12])
    return d

lk_rows = read_csv_(F_LK)
v1_links = []
for r in lk_rows[2:]:
    if not r or not r[0]: continue
    wk = lk_wk_by_name.get(r[0])
    v1_links.append(build_6metric_row(r, wk))

# ===========================================================================
# 视图02 · KA 红黑榜
# ===========================================================================
hb_rows = read_csv_(F_HB)
ka_by_name = {}
for r in hb_rows[1:]:
    if len(r) < 14: continue
    ind, name = r[0], r[1]
    if not name or name == '整体' or name not in KA: continue
    name = name.strip()
    if name not in ka_by_name:
        ka_by_name[name] = []
    ka_by_name[name].append(r)

def merge_metric(rows, val_idx, pct_idx):
    curr_sum = 0.0; prev_sum = 0.0
    has_val = False
    for r in rows:
        v = parse_num(r[val_idx])
        if v is None: continue
        has_val = True
        curr_sum += v
        pct_raw = r[pct_idx]
        if str(pct_raw) == '∞': continue
        pct = parse_num(pct_raw)
        if pct is None:
            prev_sum += v
        else:
            ratio = 1 + pct/100
            if ratio == 0: continue
            prev_sum += v / ratio
    if not has_val: return None, None
    if prev_sum == 0:
        return curr_sum, ('∞' if curr_sum > 0 else 0)
    merged_pct = (curr_sum - prev_sum) / prev_sum * 100
    return curr_sum, merged_pct

def pick_industry(rows):
    for r in rows:
        if r[0] and r[0] != '其他': return r[0]
    return rows[0][0] if rows else ''

ka_records = []
for name, rows in ka_by_name.items():
    ind = pick_industry(rows)
    cost_v, cost_p = merge_metric(rows, 2, 3)
    new_v, new_p = merge_metric(rows, 4, 5)
    fresh_v, fresh_p = merge_metric(rows, 6, 7)
    expo_v, expo_p = merge_metric(rows, 8, 9)
    used_v, used_p = merge_metric(rows, 10, 11)
    online_v, online_p = merge_metric(rows, 12, 13)
    if cost_v is None: continue
    ka_records.append({
        'ind': ind, 'name': name,
        'cost': cost_v, 'cost_p': cost_p,
        'new': new_v, 'new_p': new_p,
        'fresh': fresh_v, 'fresh_p': fresh_p,
        'expo': expo_v, 'expo_p': expo_p,
        'used': used_v, 'used_p': used_p,
        'online': online_v, 'online_p': online_p,
    })
print(f'KA 客户合并后记录数: {len(ka_records)}')

def fmt_ka(rec):
    def fp(p):
        if p is None: return ('-','flat')
        if p == '∞': return ('∞','up')
        return fmt_pct_pair(p)
    def fn(v, wan=False):
        if v is None: return '-'
        return fmt_wan(v) if wan else fmt_int(v)
    cr,cc = fp(rec['cost_p']); nr,nc = fp(rec['new_p'])
    fr,fc = fp(rec['fresh_p']); er,ec = fp(rec['expo_p'])
    ur,uc = fp(rec['used_p']); orr,oc = fp(rec['online_p'])
    return {
        'ind': rec['ind'], 'name': rec['name'],
        'cost': fn(rec['cost'],True), 'cost_r': cr, 'cost_c': cc,
        'new': fn(rec['new']), 'new_r': nr, 'new_c': nc,
        'fresh': fn(rec['fresh']), 'fresh_r': fr, 'fresh_c': fc,
        'expo': fn(rec['expo']), 'expo_r': er, 'expo_c': ec,
        'used': fn(rec['used']), 'used_r': ur, 'used_c': uc,
        'online': fn(rec['online']), 'online_r': orr, 'online_c': oc,
    }

fresh_list = [r for r in ka_records if r['fresh'] is not None]
fresh_top = [fmt_ka(r) for r in sorted(fresh_list, key=lambda x:-x['fresh'])[:10]]
fresh_bot = [fmt_ka(r) for r in sorted([r for r in fresh_list if r['fresh']>0], key=lambda x:x['fresh'])[:10]]

online_list = [r for r in ka_records if r['online'] is not None]
online_top = [fmt_ka(r) for r in sorted(online_list, key=lambda x:-x['online'])[:10]]
online_bot = [fmt_ka(r) for r in sorted([r for r in online_list if r['online']>0], key=lambda x:x['online'])[:10]]

v2 = {'fresh_top': fresh_top, 'fresh_bot': fresh_bot, 'online_top': online_top, 'online_bot': online_bot}

# ===========================================================================
# 视图03 · 分赛道
# ===========================================================================
v3 = []
for r in hy_rows[2:]:
    if not r or not r[0]: continue
    wk = hy_wk_by_name.get(r[0])
    v3.append(build_6metric_row(r, wk))
def cost_float(r):
    try: return float(str(r['cost']).replace(',',''))
    except: return 0
v3.sort(key=lambda x: -cost_float(x))

# ===========================================================================
# 写入HTML
# ===========================================================================
with open(HTML,'r',encoding='utf-8') as f:
    html = f.read()

m = re.search(r'const DATA = (\{.+?\});\n', html, re.DOTALL)
if not m:
    print("❌ HTML中找不到 const DATA = {...}; 结构，请检查 index.html 是否完整")
    sys.exit(1)
data = json.loads(m.group(1))

data[new_date] = {
    'date': new_date,
    'v1': {'pyq_summary': pyq_summary, 'sph_summary': sph_summary, 'summary': v1_summary, 'links': v1_links},
    'v2': v2, 'v3': v3
}

new_json = json.dumps(data, ensure_ascii=False, separators=(',',':'))
html_new = html.replace(m.group(0), f'const DATA = {new_json};\n')

# 更新日期选择器
all_dates = sorted(data.keys(), reverse=True)
opts_html = '\n'.join([
    f'        <option value="{dt}"{" selected" if dt==new_date else ""}>{dt}</option>'
    for dt in all_dates
])
html_new = re.sub(
    r'<select id="dateSelect" onchange="renderDate\(this\.value\)">[\s\S]*?</select>',
    f'''<select id="dateSelect" onchange="renderDate(this.value)">
{opts_html}
      </select>''',
    html_new
)

with open(HTML,'w',encoding='utf-8') as f:
    f.write(html_new)

# ===========================================================================
# 摘要 + 预警
# ===========================================================================
def pct_val(s):
    try: return float(str(s).replace('%','').replace('+',''))
    except: return None

print(f'\n✅ 已写入 {new_date} 数据')
print(f'   v1: pyq={len(pyq_summary)} sph={len(sph_summary)} summary={len(v1_summary)} links={len(v1_links)}')
print(f'   v2: fresh_top={len(fresh_top)} fresh_bot={len(fresh_bot)} online_top={len(online_top)} online_bot={len(online_bot)}')
print(f'   v3 赛道数: {len(v3)}')

print(f'\n【大盘 6 指标】')
for s in v1_summary:
    unit = s.get('unit','')
    print(f"  {s['label']}: {s['value']}{unit} (日{s['change']}/周{s['wk']})")

print(f'\n【核心板块（含全域通）】')
for s in pyq_summary + sph_summary:
    unit = s.get('unit','')
    print(f"  {s['label']}: {s['value']}{unit} (日{s['change']}/周{s['wk']})")

print(f'\n【视图01 分链路预警 曝光||消耗创意日环比<-10%，黑名单=其他】')
LINK_BL = {'其他'}
for r in v1_links:
    if r['name'] in LINK_BL: continue
    warns = []
    ep = pct_val(r['expo_r']); up = pct_val(r['used_r'])
    if ep is not None and ep < -10: warns.append(f'曝光{r["expo_r"]}')
    if up is not None and up < -10: warns.append(f'消耗创意{r["used_r"]}')
    if warns: print(f'  ⚠️ {r["name"]}: {", ".join(warns)}')

print(f'\n【视图03 赛道预警 在线日环比<-10%，黑名单=服饰配件/钟表，<0.1万隐藏】')
TRACK_BL = {'服饰配件','钟表'}
shown, hidden = 0, 0
for r in v3:
    c = cost_float(r)
    if c < 0.1: hidden += 1; continue
    shown += 1
    if r['name'] in TRACK_BL: continue
    op = pct_val(r['online_r'])
    if op is not None and op < -10:
        print(f'  ⚠️ {r["name"]}: 在线={r["online"]} 环比={r["online_r"]}')
print(f'  显示{shown}个赛道，隐藏{hidden}个（消耗<0.1万）')

print(f'\n🦞 看板更新完成！下一步：git add index.html && git commit -m "update: {new_date}" && git push')
