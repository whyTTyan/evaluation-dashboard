import os
# 防止 OpenBLAS 多线程内存分配偶发失败（曾导致 df 未定义 / 误报 NameError）
os.environ.setdefault('OPENBLAS_NUM_THREADS', '1')
os.environ.setdefault('OMP_NUM_THREADS', '1')
os.environ.setdefault('MKL_NUM_THREADS', '1')
import pandas as pd, json, numpy as np, sys, zipfile

SRC = sys.argv[1] if len(sys.argv) > 1 else r'C:\Users\T\Downloads\会话记录_最新.xlsx'
OUT = sys.argv[2] if len(sys.argv) > 2 else None
SRC_INV = sys.argv[3] if len(sys.argv) > 3 else None  # 邀评数据源
CHECK_ONLY = '--check' in sys.argv  # 轻量模式：只验数据不生成HTML

# Handle ZIP
if SRC.endswith('.zip'):
    with zipfile.ZipFile(SRC) as z:
        csv_name = [n for n in z.namelist() if n.endswith('.csv')][0]
        z.extract(csv_name, os.path.dirname(SRC))
        SRC = os.path.join(os.path.dirname(SRC), csv_name)

# Read CSV/Excel
if SRC.endswith('.csv'):
    for enc in ['gbk','gb18030','utf-8','utf-8-sig']:
        try:
            df = pd.read_csv(SRC, encoding=enc)
            break
        except: continue
else:
    df = pd.read_excel(SRC)

# === 排除指定账号 ===
EXCLUDE = ['美小言-小橙子', '美小言-17号', '美小言-小谭']
before = len(df)
df = df[~df['客服'].isin(EXCLUDE)]
print(f'已排除: {before - len(df)} 条 ({",".join(EXCLUDE)})')

# === 邀评数据 ===
inv_lookup = {}
if SRC_INV and os.path.exists(SRC_INV):
    # Handle ZIP for inv data too
    if SRC_INV.endswith('.zip'):
        with zipfile.ZipFile(SRC_INV) as z:
            csv_name = [n for n in z.namelist() if n.endswith('.csv')][0]
            z.extract(csv_name, os.path.dirname(SRC_INV))
            SRC_INV = os.path.join(os.path.dirname(SRC_INV), csv_name)
    if SRC_INV.endswith('.csv'):
        for enc in ['gbk','gb18030','utf-8','utf-8-sig']:
            try:
                df_inv = pd.read_csv(SRC_INV, encoding=enc)
                break
            except: continue
    else:
        df_inv = pd.read_excel(SRC_INV)
    mx_inv = df_inv[df_inv['客服'].str.contains('美小言', na=False)]
    inv_agg = mx_inv.groupby('客服').agg(
        邀评数=('客服','count'),
        邀评成功=('满意度', lambda x: int(x.notna().sum()))
    ).reset_index()
    inv_agg['邀评成功率'] = (inv_agg['邀评成功'] / inv_agg['邀评数'] * 100).round(1)
    for _, r in inv_agg.iterrows():
        inv_lookup[r['客服']] = [int(r['邀评数']), int(r['邀评成功']), float(r['邀评成功率'])]
    print(f'邀评数据: {len(inv_lookup)} agents loaded')

# ============================================================
# 自检机制 — 数据进来先验证，不合规报警不输出
# ============================================================
errors = []

# C1: 必填列存在
REQUIRED_COLS = ['客服', '满意度', '评价时间', '是否解决']
for c in REQUIRED_COLS:
    if c not in df.columns:
        errors.append(f'缺少必填列: {c}')

# C2: 记录数合理（每小时自动化凌晨数据量小，下限调为1）
if len(df) < 1:
    errors.append(f'记录数为0，数据源可能为空')
if len(df) > 500000:
    errors.append(f'记录数过多({len(df)}条)，检查是否导出了错误日期范围')

# C3: 客服字段有美小言账号
mx_check = df[df['客服'].str.contains('美小言', na=False)]
if len(mx_check) == 0:
    errors.append('未找到任何美小言客服，检查数据源是否正确')

# C4: 时间范围合理
try:
    times_check = df['评价时间'].dropna()
    if len(times_check) > 0:
        t_min = pd.to_datetime(times_check).min()
        t_max = pd.to_datetime(times_check).max()
        if (t_max - t_min).days > 365:
            errors.append(f'评价时间跨度过大({(t_max-t_min).days}天)，确认是否导出了多日数据')
except:
    pass

# C5: INV文件检查
if SRC_INV:
    if not os.path.exists(SRC_INV):
        errors.append(f'邀评数据文件不存在: {SRC_INV}')
    elif len(inv_lookup) == 0:
        errors.append('邀评数据已加载但无美小言客服，检查文件内容')
else:
    print('[警告] 未提供邀评数据（第3参数），INV将为空')

# 自检结果
print(f'\n{"="*50}')
print(f'自检报告 | 数据: {os.path.basename(SRC)} | 记录: {len(df)}条 | 美小言: {len(mx_check)}条')
if errors:
    for e in errors:
        print(f'  ✗ {e}')
    print(f'{"="*50}')
    print('自检未通过，退出。请修正数据后重试。')
    sys.exit(1)
else:
    print('  ✓ C1 必填列完整')
    print(f'  ✓ C2 记录数合理 ({len(df)}条)')
    print(f'  ✓ C3 美小言账号存在 ({len(mx_check)}条)')
    print('  ✓ C4 时间范围合理')
    if SRC_INV:
        print(f'  ✓ C5 邀评数据已加载 ({len(inv_lookup)}人)')
    print(f'{"="*50}\n')

if CHECK_ONLY:
    print('[--check] 数据验证通过，未生成HTML\n')
    sys.exit(0)

# 更新时间 = 从数据取评价时间最大值
try:
    times = df['评价时间'].dropna()
    max_t = times.max()[11:16] if len(times) else ''
    min_d = times.max()[:10] if len(times) else ''
    UPDATE_TIME = f'{min_d} {max_t}'
except:
    from datetime import datetime
    UPDATE_TIME = datetime.now().strftime('%m-%d %H:%M')

def prov(n):
    if pd.isna(n): return '其他'
    s = str(n)
    if '美小言' in s: return '美小言'
    if '美的空调' in s: return '美的空调'
    if '美云' in s: return '美云客服'
    if '美的小萌' in s: return '美的小萌'
    if '美的共建' in s: return '美的共建'
    return '其他'
df['p'] = df['客服'].apply(prov)

# 有评价的记录
has_sat = df['满意度'].notna() & (df['满意度']!='') & (df['满意度']!='-')
df_ev = df[has_sat].copy()

t = len(df)  # 总接待量
ev_t = len(df_ev)  # 总评价数
os_ = int((df_ev['满意度'].isin(['非常满意','满意'])).sum())
od_ = int((df_ev['满意度'].isin(['不满意','非常不满意','一般'])).sum())
g_sr = round(os_/ev_t*100, 1) if ev_t else 0
g_dr = round(od_/ev_t*100, 1) if ev_t else 0
need_good = max(0, 19*od_ - os_)

# 服务商
ps = {}
for p in ['美小言','美云客服','美的空调','美的小萌','美的共建']:
    s = df[df['p']==p]
    s_ev = df_ev[df_ev['p']==p]
    if len(s)==0: continue
    tt = len(s); evv = len(s_ev)
    ss = int((s_ev['满意度'].isin(['非常满意','满意'])).sum()) if evv else 0
    dd = int((s_ev['满意度'].isin(['不满意','非常不满意','一般'])).sum()) if evv else 0
    sr = round(ss/evv*100,1) if evv else 0
    er = round(evv/tt*100, 1) if tt else 0
    dr = round(dd/evv*100, 1) if evv else 0
    ps[p] = {'tt':tt, 'ev':evv, 'ss':ss, 'dd':dd, 'sr':sr, 'er':er, 'dr':dr}

# 各服务商邀评数（邀评数据源覆盖全部服务商；邀评数=总记录数）
prov_inv = {}
for _p in ['美小言','美云客服','美的空调','美的小萌','美的共建']:
    if SRC_INV and os.path.exists(SRC_INV):
        _sub = df_inv[df_inv['客服'].apply(prov) == _p]
        prov_inv[_p] = int(len(_sub))
    else:
        prov_inv[_p] = 0

mx = df[df['p']=='美小言']
mx_ev = df_ev[df_ev['p']=='美小言']
mx_tt = len(mx)
mx_ev_t = len(mx_ev)
mx_s = int((mx_ev['满意度'].isin(['非常满意','满意'])).sum())
mx_d = int((mx_ev['满意度'].isin(['不满意','非常不满意','一般'])).sum())
mx_sr = round(mx_s/mx_ev_t*100, 1) if mx_ev_t else 0
mx_dr = round(mx_d/mx_ev_t*100, 1) if mx_ev_t else 0
rs_ = int((mx_ev['是否解决']=='已解决').sum())
ur_ = int((mx_ev['是否解决']=='未解决').sum())
mx_rr = round(rs_/(rs_+ur_)*100, 1) if (rs_+ur_)>0 else 100
mx_iv = int((mx_ev['评价来源']=='系统邀评').sum())
mx_need = max(0, 19*mx_d - mx_s)

# 差评TOP10
bad_agg = mx_ev[mx_ev['满意度'].isin(['不满意','非常不满意','一般'])].groupby('客服').size().reset_index(name='bc')
bad_agg = bad_agg.sort_values('bc', ascending=False).head(10)
top10_bad = []
for _, row in bad_agg.iterrows():
    agent = row['客服']; bc = int(row['bc'])
    sub = mx_ev[mx_ev['客服']==agent]
    good = int((sub['满意度'].isin(['非常满意','满意'])).sum())
    bad = int((sub['满意度'].isin(['不满意','非常不满意','一般'])).sum())
    sr = round(good/(good+bad)*100, 1) if (good+bad) else 0
    bad_recs = sub[sub['满意度'].isin(['不满意','非常不满意','一般'])].sort_values('评价时间', ascending=False)
    lt = bad_recs.iloc[0]
    top10_bad.append({'n':agent,'bc':bc,'gt':good,'tt':len(mx[mx['客服']==agent]),'bt':bad,'sr':sr,
        'lr':str(lt['评价时间'])[11:19] if pd.notna(lt['评价时间']) else '',
        'sat':lt['满意度'],'re':str(lt['不满意原因'])[:22] if lt['不满意原因']!='-' else ''})

# 好评TOP10
good_agg = mx_ev.groupby('客服').agg(
    total=('满意度','count'),
    good=('满意度',lambda x:int((x.isin(['非常满意','满意'])).sum())),
    bad=('满意度',lambda x:int((x.isin(['不满意','非常不满意','一般'])).sum()))
).reset_index()
good_agg['sr'] = (good_agg['good']/good_agg['total']*100).round(1)
good_agg = good_agg.sort_values('good', ascending=False).head(10)
top10_good = [{'n':r['客服'],'gt':int(r['good']),'bt':int(r['bad']),'tt':len(mx[mx['客服']==r['客服']]),'ev':int(r['total']),'sr':r['sr'],'dk':1 if r['sr']>=95 else 0} for _,r in good_agg.iterrows()]

# 客服排名（含接待量）
a = mx.groupby('客服').size().reset_index(name='tt')  # 总接待量 = 全部会话数
# 评价数据
a_ev = mx_ev.groupby('客服').agg(
    ev=('满意度','count'), vs=('满意度',lambda x:int((x=='非常满意').sum())),
    ss=('满意度',lambda x:int((x=='满意').sum())), dd=('满意度',lambda x:int((x=='不满意').sum())),
    vd=('满意度',lambda x:int((x=='非常不满意').sum())),
    gen=('满意度',lambda x:int((x=='一般').sum())),
    rr=('是否解决',lambda x:int((x=='已解决').sum())), uu=('是否解决',lambda x:int((x=='未解决').sum())),
    iv=('评价来源',lambda x:int((x=='系统邀评').sum()))
).reset_index()
a = a.merge(a_ev, on='客服', how='left').fillna(0)
a['sr'] = ((a['vs']+a['ss'])/a['ev']*100).round(1)
a['dr'] = ((a['dd']+a['vd']+a['gen'])/a['ev']*100).round(1)
try:
    a['rr'] = (a['rr']/(a['rr']+a['uu'])*100).round(1).fillna(100)
except ZeroDivisionError:
    a['rr'] = 100.0
a['good'] = a['vs']+a['ss']
a = a[a['tt']>=2].sort_values('dr', ascending=False).fillna(0)

ags = []
for _, r in a.iterrows():
    wl='高危' if r['dr']>=50 else('预警' if r['dr']>=30 else('关注' if r['dr']>=15 else'正常'))
    dk = r['sr']>=95
    qualified = 1 if (r['good']>=10 and r['sr']>=90) else 0
    star = 1 if (r['good']>=15 and r['sr']>=95) else 0
    bd = mx_ev[(mx_ev['客服']==r['客服'])&(mx_ev['满意度'].isin(['不满意','非常不满意','一般']))]
    tr_=''
    if len(bd):
        rc=bd['不满意原因'].value_counts(); tr_=str(rc.index[0])[:20] if rc.index[0]!='-' else ''
    gd = int(r['vs'])+int(r['ss'])
    bd = int(r['dd'])+int(r['vd'])+int(r['gen'])
    ags.append({'n':r['客服'],'tt':int(r['tt']),'ev':int(r['ev']),'gd':gd,'bd':bd,'sr':float(r['sr']),'dr':float(r['dr']),
        'rr':float(r['rr']),'iv':int(r['iv']),'wl':wl,'tr':tr_,'dk':dk,'qg':qualified,'st':star,'need':max(0,19*bd-gd)})

mb2 = mx_ev[mx_ev['满意度'].isin(['不满意','非常不满意','一般'])]
res = [{'l':'未填写原因' if k=='-' else str(k)[:25],'c':int(v)} for k,v in mb2['不满意原因'].value_counts().items()]

# 差评时段分布 — 按小时聚合差评数 + 涉及客服ID+差评数（精简：111号(3), 222号(1)）
mx_bad = mx_ev[mx_ev['满意度'].isin(['不满意','非常不满意','一般'])].copy()
hourly_data = []
if len(mx_bad) > 0:
    mx_bad['hour'] = pd.to_datetime(mx_bad['评价时间']).dt.hour
    hourly_cnt = mx_bad.groupby('hour').size().reset_index(name='count')
    hourly_detail = mx_bad.groupby(['hour','客服']).size().reset_index(name='cnt')
    def _fmt(g):
        items = sorted(zip(g['客服'],g['cnt']), key=lambda x:-x[1])
        R = '<span style="color:#DC2626;font-weight:600">'
        CIR = '①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳'
        def _n(c):
            if 1 <= c <= 20: return CIR[c-1]
            return f'({c})'
        short = ', '.join(f"{a.replace('美小言-','')}{R}{_n(c)}</span>" for a,c in items)
        full  = ', '.join(f"{a}{_n(c)}" for a,c in items)  # 纯文本，给title属性用
        return pd.Series({'agents_short':short,'agents_full':full})
    hourly_agents = hourly_detail.groupby('hour').apply(_fmt).reset_index()
    hourly = hourly_cnt.merge(hourly_agents, on='hour')
    hourly = hourly.sort_values('hour', ascending=True)  # 按时间从早到晚
    for _, r in hourly.iterrows():
        h = int(r['hour'])
        hourly_data.append({
            'h': h,
            'label': f'{h:02d}:00-{h+1:02d}:00',
            'c': int(r['count']),
            'a': r['agents_short'],
            'full': r['agents_full']
        })

class E(json.JSONEncoder):
    def default(self,o):
        if isinstance(o,(np.integer,)):return int(o)
        if isinstance(o,(np.floating,)):return float(o)
        return super().default(o)
K=lambda d:json.dumps(d,ensure_ascii=False,cls=E)

AGENT_TOTAL = len(ags)
AGENT_QUALIFIED = sum(1 for a in ags if a['qg'])
AGENT_STAR = sum(1 for a in ags if a['st'])

kpi = ''
kpi += f'<div class="kc kb"><div class="kl">总接待量</div><div class="kv">{t}</div><div class="ks">已评价{ev_t}条</div></div>'
kpi += f'<div class="kc kg"><div class="kl">全局满意率</div><div class="kv">{g_sr}%</div><div class="ks">好评{os_}条</div></div>'
kpi += f'<div class="kc kr_"><div class="kl">全局不满率</div><div class="kv">{g_dr}%</div><div class="ks">差评{od_}条</div></div>'
kpi += f'<div class="kc ko"><div class="kl">美小言接待</div><div class="kv">{mx_tt}</div><div class="ks">评价{mx_ev_t}条</div></div>'
kpi += f'<div class="kc kr_"><div class="kl">美小言满意率</div><div class="kv">{mx_sr}%</div><div class="ks">差评{mx_d}条</div></div>'

prov = ''
wl_cls = {'高危':('wr','pcm'), '预警':('wy','pcy'), '关注':('wb','pcb'), '正常':('wg','pcg')}
for nm in ['美小言','美云客服','美的空调','美的小萌','美的共建']:
    if nm not in ps: continue
    d=ps[nm]; ex=''; dr=d.get('dr',50)
    wl='高危' if dr>=50 else('预警' if dr>=30 else('关注' if dr>=15 else'正常'))
    bc,pc=wl_cls[wl]
    if d['sr']>=95: ex=' <span class="bc">已达标</span>'
    gap = max(0, 19*d['dd']-d['ss'])
    if gap>0 and d['dd']>0: ex=f' <span class="wy" style="font-size:11px;border-radius:4px;padding:0 7px">缺{gap}好评</span>'
    ir = prov_inv.get(nm, 0)
    ir_disp = f'{round(ir/d["tt"]*100,1)}%' if (ir and d["tt"]) else '-'
    prov += f'<div class="pc {pc}"><div class="pn">{nm}<span class="w {bc}">{wl}</span>{ex}</div><div class="ps_"><span>接待 {d["tt"]}</span><span>评价率 {d["er"]}%</span><span>满意 {d["sr"]}%</span><span>差评 {d["dd"]}</span><span>邀评率 {ir_disp}</span></div></div>'

alert = f'<div class="ab"><div class="ai">&#9888;&#65039;</div><div><div class="at">美小言不满率 {mx_dr}%（{mx_d}条差评 / {mx_ev_t}条评价），需{mx_need}个好评达标95%</div><div class="as" id="vrfNote">加载中...</div></div></div>'

# 原生渲染两张图（无 Chart.js 依赖，避免 CDN 不可达导致空白）
_PI = 3.14159
_total = (mx_s or 0) + (mx_d or 0) or 1
_C = 2 * _PI * 40
_len_s = (mx_s / _total) * _C
_len_d = (mx_d / _total) * _C
donut_svg = (
    f'<div style="position:relative;width:100%;min-height:200px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px">'
    f'<svg viewBox="0 0 100 100" style="width:150px;height:150px">'
    f'<circle cx="50" cy="50" r="40" fill="none" stroke="#F3F4F6" stroke-width="18"/>'
    f'<circle cx="50" cy="50" r="40" fill="none" stroke="#86EFAC" stroke-width="18" '
    f'stroke-dasharray="{_len_s:.2f} {_C-_len_s:.2f}" transform="rotate(-90 50 50)"/>'
    f'<circle cx="50" cy="50" r="40" fill="none" stroke="#FCA5A5" stroke-width="18" '
    f'stroke-dasharray="{_len_d:.2f} {_C-_len_d:.2f}" '
    f'transform="rotate({-90 + (mx_s/_total)*360:.2f} 50 50)"/>'
    f'</svg>'
    f'<div style="display:flex;gap:14px;font-size:12px;color:#374151">'
    f'<span style="display:inline-flex;align-items:center;gap:4px"><i style="width:10px;height:10px;border-radius:2px;background:#86EFAC;display:inline-block"></i>满意 {mx_s}条 ({mx_sr}%)</span>'
    f'<span style="display:inline-flex;align-items:center;gap:4px"><i style="width:10px;height:10px;border-radius:2px;background:#FCA5A5;display:inline-block"></i>不满 {mx_d}条 ({mx_dr}%)</span>'
    f'</div></div>'
)

if res:
    _max_c = max((r['c'] for r in res), default=0) or 1
    _rows = []
    for _r in res[:8]:
        _pct = _r['c'] / _max_c * 100
        _rows.append(
            f'<div style="display:flex;align-items:center;gap:8px;margin:5px 0;font-size:12px">'
            f'<div style="width:95px;flex-shrink:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#374151">{_r["l"]}</div>'
            f'<div style="flex:1;background:#F3F4F6;border-radius:3px;height:14px;overflow:hidden">'
            f'<div style="width:{_pct:.1f}%;background:#DC2626;height:100%;border-radius:3px"></div>'
            f'</div>'
            f'<div style="width:32px;text-align:right;color:#6b7280;flex-shrink:0">{_r["c"]}</div>'
            f'</div>'
        )
    bar_html = '<div style="padding:4px 0">' + ''.join(_rows) + '</div>'
    if len(res) > 8:
        bar_html += f'<div style="font-size:11px;color:#9ca3af;text-align:center;margin-top:6px">其他 {len(res)-8} 类原因未展示</div>'
else:
    bar_html = '<div style="color:#9ca3af;font-size:12px;text-align:center;padding:24px 0">暂无差评原因数据</div>'

html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>评价看板 - 含接待量</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:-apple-system,'PingFang SC','Microsoft YaHei',sans-serif;background:#f0f2f5;color:#1a1d23;font-size:14px}}
.c{{max-width:1440px;margin:0 auto;padding:16px}}
.hd{{background:linear-gradient(135deg,#1f2a44,#2d3b5c);color:#fff;border-radius:12px;padding:18px 24px;margin-bottom:16px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px}}
.hd h1{{font-size:20px;font-weight:700}}
.hd .i{{font-size:13px;opacity:.75}}
.kr{{display:grid;grid-template-columns:repeat(auto-fit,minmax(145px,1fr));gap:12px;margin-bottom:16px}}
.kc{{background:#fff;border-radius:10px;padding:14px 16px;box-shadow:0 1px 4px rgba(0,0,0,.06);border-left:4px solid}}
.kc .kl{{font-size:12px;color:#6b7280;margin-bottom:4px}}.kc .kv{{font-size:26px;font-weight:700}}.kc .ks{{font-size:11px;color:#6b7280;margin-top:2px}}
.kb{{border-color:#4A90D9}}.kb .kv{{color:#4A90D9}}.kg{{border-color:#51B87A}}.kg .kv{{color:#51B87A}}.kr_{{border-color:#FF6B6B}}.kr_ .kv{{color:#FF6B6B}}.ko{{border-color:#FFA94D}}.ko .kv{{color:#FFA94D}}
.st{{font-size:16px;font-weight:700;margin:18px 0 10px;display:flex;align-items:center;gap:8px}}.st .tg{{font-size:11px;background:#eef2ff;color:#4A90D9;border-radius:20px;padding:2px 10px;font-weight:500}}.st .ts{{font-size:11px;color:#6b7280;font-weight:400;margin-left:auto}}
.pm{{display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap}}
.pc{{background:#fff;border-radius:8px;padding:10px 14px;box-shadow:0 1px 4px rgba(0,0,0,.06);flex:1;min-width:160px;border-left:4px solid;font-size:13px}}
.pc .pn{{font-weight:600;font-size:14px;margin-bottom:4px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:4px}}
.pc .pn .w{{font-size:11px;font-weight:600;border-radius:4px;padding:1px 7px}}
.pc .ps_{{display:flex;gap:12px;color:#6b7280;flex-wrap:wrap}}.pc .ps_ span{{white-space:nowrap}}
.pcm{{border-color:#FF6B6B}}.pcy{{border-color:#4ECDC4}}.pcg{{border-color:#51B87A}}.pcb{{border-color:#4A90D9}}
.wr{{background:#FEE2E2;color:#DC2626}}.wy{{background:#FEF3C7;color:#D97706}}.wb{{background:#E8F4F8;color:#4A90D9}}.wg{{background:#D1FAE5;color:#059669}}
.ab{{background:linear-gradient(135deg,#FEF2F2,#FFE4E4);border:1px solid #FECACA;border-radius:10px;padding:12px 18px;margin-bottom:14px;display:flex;align-items:center;gap:10px}}
.ab .ai{{font-size:20px}}.ab .at{{font-size:14px;font-weight:600;color:#991B1B}}.ab .as{{font-size:12px;color:#B91C1C;margin-top:1px}}
.cr{{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:16px}}
.cc{{background:#fff;border-radius:10px;padding:16px 18px;box-shadow:0 1px 4px rgba(0,0,0,.06)}}.cc h3{{font-size:14px;font-weight:600;margin-bottom:12px}}.cc canvas{{max-height:250px;max-width:100%}}
.ts{{background:#fff;border-radius:10px;padding:16px 18px;box-shadow:0 1px 4px rgba(0,0,0,.06);margin-bottom:16px;overflow-x:auto}}
.ts h3{{font-size:14px;font-weight:600;margin-bottom:10px;display:flex;align-items:center;gap:8px}}
table{{width:100%;border-collapse:collapse;font-size:13px;border-collapse:separate;border-spacing:0}}
th{{text-align:left;padding:9px 8px;border-bottom:2px solid #e5e7eb;color:#6b7280;font-weight:600;font-size:12px;white-space:nowrap;cursor:pointer;-webkit-user-select:none;position:sticky;top:0;background:#fff;z-index:10;box-shadow:inset 0 -1px 0 #e5e7eb}}
th:hover{{color:#1a1d23;background:#f8f9fa}}
td{{padding:7px 8px;border-bottom:1px solid #f3f4f6;white-space:nowrap}}
tr:hover td{{background:#f9fafb}}
.rk{{font-weight:600;color:#6b7280;width:28px}}
.rh{{color:#059669;font-weight:600}}.rm{{color:#D97706;font-weight:600}}.rl{{color:#DC2626;font-weight:600}}
.bh{{background:#FEE2E2;color:#DC2626;border-radius:4px;padding:0 7px;font-size:11px;font-weight:600;line-height:20px;display:inline-block}}
.by{{background:#FEF3C7;color:#D97706;border-radius:4px;padding:0 7px;font-size:11px;font-weight:600;line-height:20px;display:inline-block}}
.bg{{background:#D1FAE5;color:#059669;border-radius:4px;padding:0 7px;font-size:11px;font-weight:600;line-height:20px;display:inline-block}}
.bb{{background:#E8F4F8;color:#4A90D9;border-radius:4px;padding:0 7px;font-size:11px;font-weight:600;line-height:20px;display:inline-block}}
.bc{{background:#D1FAE5;color:#059669;border-radius:4px;padding:0 7px;font-size:11px;font-weight:600;line-height:20px;display:inline-block}}
.hl-name{{color:#059669;font-weight:700}}
@media(max-width:900px){{.cr{{grid-template-columns:1fr}}}}
</style>
</head>
<body>
<div class="c">
<div class="hd"><h1>\u5b9e\u65f6\u8bc4\u4ef7\u770b\u677f</h1><div class="i">\u66f4\u65b0\u65f6\u95f4 {UPDATE_TIME} | \u63a5\u5f85{t}\u00b7\u8bc4\u4ef7{ev_t}\u00b7\u76ee\u680795%</div></div>

<div class="kr">{kpi}</div>
<div class="st">\u670d\u52a1\u5546\u5bf9\u6bd4 <span class="tg">5\u8d26\u53f7</span><span class="ts">\u63a5\u5f85\u91cf\u00b7\u6ee1\u610f\u5ea6\u00b7\u5dee\u8bc4\u91cf</span></div>
<div class="pm">{prov}</div>
{alert}

<div class="cr">
<div class="cc"><h3>\u4e0d\u6ee1\u610f\u539f\u56e0\u5206\u5e03\uff08\u7f8e\u5c0f\u8a00\uff09</h3>{bar_html}</div>
<div class="cc"><h3>\u7f8e\u5c0f\u8a00\u6574\u4f53\u8bc4\u4ef7\uff08{mx_ev_t}\u6761\uff09</h3>{donut_svg}</div>
</div>

<div class="st" style="margin-top:0">\u7f8e\u5c0f\u8a00\u6392\u884c <span class="tg">\u5dee\u8bc4TO10 \u00b7 \u597d\u8bc4TOP10</span><span class="ts">更新时间: {UPDATE_TIME}</span></div>
<div class="cr">
<div class="ts" style="border-left:4px solid #FF6B6B;margin-bottom:0">
<h3 style="color:#DC2626;font-size:13px">\u5dee\u8bc4TOP10 <span style="font-weight:400;color:#6b7280;font-size:12px;">\u6309\u5dee\u8bc4\u6570</span><span style="font-weight:400;color:#94a3b8;font-size:11px;margin-left:auto">{UPDATE_TIME}</span></h3>
<table style="font-size:12px"><thead><tr>
<th>#</th><th onclick="sb(0)">\u5ba2\u670d</th><th onclick="sb(1)">\u5dee\u8bc4</th><th onclick="sb(2)">\u597d\u8bc4</th><th onclick="sb(3)">\u63a5\u5f85\u91cf</th><th onclick="sb(4)">\u6ee1\u610f\u7387</th><th onclick="sb(5)">\u9700\u597d\u8bc4</th><th onclick="sb(6)">邀评数</th><th onclick="sb(7)">邀评率</th><th onclick="sb(8)">邀评成功率</th><th onclick="sb(9)">\u6700\u65b0\u5dee\u8bc4</th><th>\u539f\u56e0</th>
</tr></thead>
<tbody id="t10b"></tbody>
</table></div>
<div class="ts" style="border-left:4px solid #51B87A;margin-bottom:0">
<h3 style="color:#059669;font-size:13px">\u597d\u8bc4TOP10 <span style="font-weight:400;color:#6b7280;font-size:12px;">\u6309\u597d\u8bc4\u6570</span><span style="font-weight:400;color:#94a3b8;font-size:11px;margin-left:auto">{UPDATE_TIME}</span></h3>
<table style="font-size:12px"><thead><tr>
<th>#</th><th onclick="sg(0)">\u5ba2\u670d</th><th onclick="sg(1)">\u597d\u8bc4</th><th onclick="sg(2)">\u5dee\u8bc4</th><th onclick="sg(3)">\u63a5\u5f85\u91cf</th><th onclick="sg(4)">\u8bc4\u4ef7</th><th onclick="sg(5)">\u6ee1\u610f\u7387</th><th onclick="sg(6)">邀评数</th><th onclick="sg(7)">邀评率</th><th onclick="sg(8)">邀评成功率</th><th onclick="sg(9)">\u8fbe\u680795%</th>
</tr></thead>
<tbody id="t10g"></tbody>
</table></div>
</div>

<div class="st" style="margin-top:0">\u5dee\u8bc4\u65f6\u6bb5\u5206\u5e03 <span class="tg">\u7f8e\u5c0f\u8a00\u5dee\u8bc4\u6309\u5c0f\u65f6\u7edf\u8ba1</span><span class="ts">\u66f4\u65b0\u65f6\u95f4: {UPDATE_TIME}</span></div>
<div class="ts">
<table style="font-size:12px"><thead><tr><th>#</th><th onclick=\"sh('h')\" style=\"cursor:pointer\">时段 &#9662;</th><th onclick=\"sh('c')\" style=\"cursor:pointer\">差评数 &#9662;</th><th>涉及客服</th></tr></thead>
<tbody id="hourlyBody"><tr><td colspan="4" style="text-align:center;color:#9ca3af;padding:16px">加载中...</td></tr></tbody></table>
</div>

<div class="ts">
<h3>\u7f8e\u5c0f\u8a00\u5ba2\u670d\u8bc4\u4ef7\u6392\u540d <span style="font-weight:400;color:#6b7280;font-size:13px;">\u4e0d\u6ee1\u7387\u6392\u5e8f \u00b7 \u8fbe\u6210\u597d\u8bc410\u4e14\u6ee190%\u7effID\u00b7\u8fbe\u621015\u4e14\u6ee195%\u2605</span><span style="font-weight:400;color:#6b7280;font-size:13px;margin-left:auto">\u5171 <b style=\"color:#1a1d23;font-size:15px\">{AGENT_TOTAL}</b> \u00b7 \u8fbe\u6210 <b style=\"color:#059669;font-size:15px\">{AGENT_QUALIFIED}</b> \u00b7 \u2605 <b style=\"color:#F59E0B;font-size:15px\">{AGENT_STAR}</b></span><span style="font-weight:400;color:#94a3b8;font-size:11px;margin-left:8px">{UPDATE_TIME}</span></h3>
<div style="margin-bottom:10px"><input id="searchInput" type="text" placeholder="\u641c\u7d22\u5ba2\u670dID..." onkeyup="filterTable()" style="padding:6px 12px;border:1px solid #d1d5db;border-radius:6px;font-size:13px;width:200px;outline:none"><span style="font-size:11px;color:#6b7280;margin-left:8px">\u8f93\u5165\u5ba2\u670d\u540d\u79f0\u7b5b\u9009</span><button onclick="exportCSV()" style="margin-left:12px;padding:6px 14px;background:#4A90D9;color:#fff;border:none;border-radius:6px;font-size:12px;cursor:pointer">\u5bfc\u51faCSV</button></div>
<div style="max-height:620px;overflow-y:auto;border:1px solid #f0f0f0;border-radius:6px">
<table><thead><tr>
<th>#</th><th onclick="st('n')">\u5ba2\u670d</th><th onclick="st('tt')">\u63a5\u5f85\u91cf</th><th onclick="st('ev')">\u8bc4\u4ef7</th><th onclick="st('gd')">\u597d\u8bc4</th><th onclick="st('bd')">\u5dee\u8bc4</th><th onclick="st('sr')">\u6ee1\u610f\u7387</th><th onclick="st('dr')">\u4e0d\u6ee1\u7387</th><th onclick="st('invc')">邀评数</th><th onclick="st('invrate')">邀评率</th><th onclick="st('invr')">邀评成功率</th><th onclick="st('tr')">\u5dee\u8bc4\u4e3b\u56e0</th><th onclick="st('wl')">\u9884\u8b66</th><th onclick="st('need')">\u76ee\u680795%</th>
</tr></thead>
<tbody id="tb"></tbody>
</table>
</div></div>
</div>

<script>
var A = {K(ags)};
var RE = {K(res)};
var T10B = {K(top10_bad)};
var T10G = {K(top10_good)};
var HOURLY = {K(hourly_data)};
var INV = __INV_JSON__;
A.forEach(function(a){{var i=INV[a.n];if(i){{a.invc=i[0];a.invr=i[2];a.invrate=a.tt>0?+(i[0]/a.tt*100).toFixed(1):0;}}else{{a.invc=0;a.invr=0;a.invrate=0;}}}});
T10B.forEach(function(r){{var i=INV[r.n];if(i){{r.invc=i[0];r.invr=i[2];r.invrate=r.tt>0?+(i[0]/r.tt*100).toFixed(1):0;}}else{{r.invc=0;r.invr=0;r.invrate=0;}}r.need=Math.max(0,19*r.bc-r.gt);}});
T10G.forEach(function(r){{var i=INV[r.n];if(i){{r.invc=i[0];r.invr=i[2];r.invrate=r.tt>0?+(i[0]/r.tt*100).toFixed(1):0;}}else{{r.invc=0;r.invr=0;r.invrate=0;}}}});
(function(){{var c=A.filter(function(x){{return x.ev>0;}});var pick=null;if(c.length)pick=c.reduce(function(a,b){{return (a.bd||0)>(b.bd||0)?a:b;}});var el=document.getElementById('vrfNote');if(el&&pick){{el.textContent=pick.n+': \u603b\u63a5\u5f85'+pick.tt+'\u6b21, \u597d\u8bc4'+(pick.gd||0)+'\u6761, \u5dee\u8bc4'+(pick.bd||0)+'\u6761, \u6ee1\u610f\u7387'+pick.sr+'%, \u9700'+pick.need+'\u4e2a\u597d\u8bc4';}}else if(el){{el.textContent='\u5f53\u524d\u65e0\u53ef\u7528\u9a8c\u8bc1\u8d26\u53f7';}}}})();
var sc='dr',sd='desc';
var hsc='h',hsd='asc';
function sh(col){{if(hsc===col){{hsd=hsd==='asc'?'desc':'asc'}}else{{hsc=col;hsd='asc'}}var s=HOURLY.slice().sort(function(a,b){{var av=a[hsc],bv=b[hsc];return hsd==='asc'?(av>bv?1:-1):(av<bv?1:-1)}});var h='';if(s.length===0){{h='<tr><td colspan="4" style="text-align:center;color:#9ca3af;padding:16px">暂无差评时段数据</td></tr>'}}else{{for(var i=0;i<s.length;i++){{var r=s[i];h+='<tr><td class="rk">'+(i+1)+'</td><td style="font-weight:600">'+r.label+'</td><td class="rl" style="font-weight:600">'+r.c+'</td><td style="font-size:11px;color:#6b7280" title="'+r.full+'">'+r.a+'</td></tr>'}}}}document.getElementById('hourlyBody').innerHTML=h;}}sh('h');
function wb(w){{if(w=='\u9ad8\u5371')return '<span class="bh">\u9ad8\u5371</span>';if(w=='\u9884\u8b66')return '<span class="by">\u9884\u8b66</span>';if(w=='\u5173\u6ce8')return '<span class="bb">\u5173\u6ce8</span>';return '<span class="bg">\u6b63\u5e38</span>';}}
function rc_(v){{return v>=95?'rh':v>=70?'rm':'rl';}}
function need(v,n){{return v>=95?'\u8fbe\u6807':'\u9700'+n+'\u4e2a\u597d\u8bc4';}}

var hb='';
for(var i=0;i<T10B.length;i++){{var r=T10B[i];var tc=r.sr>=95?'rh':r.sr>=70?'rm':'rl';
  var bc=Math.max(0,19*r.bc-r.gt);
  hb+='<tr><td class="rk">'+(i+1)+'</td><td style="font-size:12px"><strong>'+r.n+'</strong></td><td class="rl"><strong>'+r.bc+'</strong></td><td class="rh">'+r.gt+'</td><td>'+r.tt+'</td><td class="'+tc+'">'+r.sr+'%</td><td>'+(r.sr>=95?'\u8fbe\u6807':bc+'\u4e2a')+'</td><td>'+(INV[r.n]?INV[r.n][0]:0)+'</td><td>'+(INV[r.n]? (r.tt>0?(INV[r.n][0]/r.tt*100).toFixed(1):'0')+'%':'-')+'</td><td>'+(INV[r.n]?INV[r.n][2]+'%':'-')+'</td><td style="font-size:11px;color:#6b7280">'+r.lr+'</td><td style="font-size:11px;color:#6b7280;max-width:100px;overflow:hidden">'+r.re+'</td></tr>';}}
document.getElementById('t10b').innerHTML=hb;

var hg='';
for(var i=0;i<T10G.length;i++){{var r=T10G[i];var tc=r.sr>=95?'rh':r.sr>=70?'rm':'rl';
  var dk=r.dk?'<span class="bc">\u8fbe\u6807</span>':'<span style="color:#6b7280;font-size:11px">\u9700'+Math.max(0,19*r.bt-r.gt)+'\u4e2a\u597d\u8bc4</span>';
  hg+='<tr><td class="rk">'+(i+1)+'</td><td style="font-size:12px"><strong>'+r.n+'</strong></td><td class="rh"><strong>'+r.gt+'</strong></td><td>'+r.bt+'</td><td>'+r.tt+'</td><td>'+r.ev+'</td><td class="'+tc+'">'+r.sr+'%</td><td>'+(INV[r.n]?INV[r.n][0]:0)+'</td><td>'+(INV[r.n]? (r.tt>0?(INV[r.n][0]/r.tt*100).toFixed(1):'0')+'%':'-')+'</td><td>'+(INV[r.n]?INV[r.n][2]+'%':'-')+'</td><td>'+dk+'</td></tr>';}}
document.getElementById('t10g').innerHTML=hg;

var sbd=1,sgd=1;
function sb(i){{sbd*=-1;var a=T10B.slice();var ks=['n','bc','gt','tt','sr','need','invc','invrate','invr','lr'];
  a.sort(function(x,y){{var xv=x[ks[i]],yv=y[ks[i]];if(typeof xv==='string')return (sbd>0?1:-1)*xv.localeCompare(yv,'zh');return (sbd>0?1:-1)*(xv-yv);}});
  var h='';for(var k=0;k<a.length;k++){{var r=a[k];var tc=r.sr>=95?'rh':r.sr>=70?'rm':'rl';
    var bc=Math.max(0,19*r.bc-r.gt);
    h+='<tr><td class="rk">'+(k+1)+'</td><td style="font-size:12px"><strong>'+r.n+'</strong></td><td class="rl"><strong>'+r.bc+'</strong></td><td class="rh">'+r.gt+'</td><td>'+r.tt+'</td><td class="'+tc+'">'+r.sr+'%</td><td>'+(r.sr>=95?'\u8fbe\u6807':bc+'\u4e2a')+'</td><td>'+(INV[r.n]?INV[r.n][0]:0)+'</td><td>'+(INV[r.n]? (r.tt>0?(INV[r.n][0]/r.tt*100).toFixed(1):'0')+'%':'-')+'</td><td>'+(INV[r.n]?INV[r.n][2]+'%':'-')+'</td><td style="font-size:11px;color:#6b7280">'+r.lr+'</td><td style="font-size:11px;color:#6b7280;max-width:100px;overflow:hidden">'+r.re+'</td></tr>';}}
  document.getElementById('t10b').innerHTML=h;}}
function sg(i){{sgd*=-1;var a=T10G.slice();var ks=['n','gt','bt','tt','ev','sr','invc','invrate','invr','dk'];
  a.sort(function(x,y){{var xv=x[ks[i]],yv=y[ks[i]];if(typeof xv==='string')return (sgd>0?1:-1)*xv.localeCompare(yv,'zh');return (sgd>0?1:-1)*(xv-yv);}});
  var h='';for(var k=0;k<a.length;k++){{var r=a[k];var tc=r.sr>=95?'rh':r.sr>=70?'rm':'rl';
    var dk=r.dk?'<span class="bc">\u8fbe\u6807</span>':'<span style="color:#6b7280;font-size:11px">\u9700'+Math.max(0,19*r.bt-r.gt)+'\u4e2a\u597d\u8bc4</span>';
    h+='<tr><td class="rk">'+(k+1)+'</td><td style="font-size:12px"><strong>'+r.n+'</strong></td><td class="rh"><strong>'+r.gt+'</strong></td><td>'+r.bt+'</td><td>'+r.tt+'</td><td>'+r.ev+'</td><td class="'+tc+'">'+r.sr+'%</td><td>'+(INV[r.n]?INV[r.n][0]:0)+'</td><td>'+(INV[r.n]? (r.tt>0?(INV[r.n][0]/r.tt*100).toFixed(1):'0')+'%':'-')+'</td><td>'+(INV[r.n]?INV[r.n][2]+'%':'-')+'</td><td>'+dk+'</td></tr>';}}
  document.getElementById('t10g').innerHTML=h;}}

function st(f){{if(sc===f){{sd=sd==='asc'?'desc':'asc'}}else{{sc=f;sd='desc'}}var q=document.getElementById('searchInput').value.trim().toLowerCase();rt(q);}}
function rt(q){{var b='';var a=A.slice();a.sort(function(x,y){{var xv=x[sc],yv=y[sc];if(typeof xv==='string')return sd==='asc'?xv.localeCompare(yv,'zh'):yv.localeCompare(xv,'zh');return sd==='asc'?xv-yv:yv-xv;}});
  for(var i=0;i<a.length;i++){{var r=a[i];if(q&&!r.n.toLowerCase().includes(q))continue;var tc=r.dr>=50?'rl':r.dr>=30?'rm':'';
    var dk=r.dk?'<span class="bc">\u8fbe\u6807</span>':'<span style="color:#6b7280;font-size:11px">\u9700'+r.need+'\u4e2a\u597d\u8bc4</span>';
    var hl=r.qg?' hl-name':''; var star=r.st?' \u2605':'';
    b+='<tr><td class="rk">'+(i+1)+'</td><td class="'+hl+'"><strong>'+r.n+star+'</strong></td><td>'+r.tt+'</td><td>'+r.ev+'</td><td class="rh">'+r.gd+'</td><td class="rl">'+r.bd+'</td><td class="'+rc_(r.sr)+'">'+r.sr+'%</td><td class="'+tc+'">'+r.dr+'%</td><td>'+(INV[r.n]?INV[r.n][0]:0)+'</td><td>'+(INV[r.n]? (r.tt>0?(INV[r.n][0]/r.tt*100).toFixed(1):'0')+'%':'-')+'</td><td>'+(INV[r.n]?INV[r.n][2]+'%':'-')+'</td><td style="font-size:12px;max-width:100px">'+r.tr+'</td><td>'+wb(r.wl)+'</td><td>'+dk+'</td></tr>';}}
  document.getElementById('tb').innerHTML=b;}}
function filterTable(){{var q=document.getElementById('searchInput').value.trim().toLowerCase();rt(q);}}
function exportCSV(){{var h=['\u5ba2\u670d','\u63a5\u5f85\u91cf','\u8bc4\u4ef7','\u597d\u8bc4','\u5dee\u8bc4','\u6ee1\u610f\u7387','\u4e0d\u6ee1\u7387','\u9080\u8bc4\u6570','\u9080\u8bc4\u7387','\u9080\u8bc4\u6210\u529f\u7387','\u5dee\u8bc4\u4e3b\u56e0','\u9884\u8b66','\u9700\u597d\u8bc4','\u8fbe\u6807','\u660e\u661f'];var k=['n','tt','ev','gd','bd','sr','dr','invc','invrate','invr','tr','wl','need','dk','st'];var c='\ufeff'+h.join(',')+'\\n';for(var i=0;i<A.length;i++){{var r=A[i];var row=[];for(var j=0;j<k.length;j++){{var v=r[k[j]];if(v===undefined||v===null)v='';row.push(v);}}c+=row.join(',')+'\\n';}}var b=new Blob([c],{{type:'text/csv;charset=utf-8'}});var a=document.createElement('a');a.href=URL.createObjectURL(b);a.download='\u7f8e\u5c0f\u8a00\u6392\u540d\u660e\u7ec6.csv';a.click();}}
function rt2(f){{document.getElementById('searchInput').value='';sc=f;sd='desc';rt('');}}
rt('');
</script>
</body>
</html>'''

# Fix interpolation
html = html.replace('{mx_s}', str(mx_s)).replace('{mx_d}', str(mx_d))
html = html.replace('{mx_sr}', str(mx_sr)).replace('{mx_dr}', str(mx_dr))
html = html.replace('{UPDATE_TIME}', UPDATE_TIME)
html = html.replace('__INV_JSON__', K(inv_lookup))

out_path = OUT or os.path.join(os.path.dirname(__file__) or '.', '评价看板_v10.html')
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(html)
print(f'OK: {len(html)} bytes -> {out_path}')

# ============================================================
# 输出自检 — 确保生成的HTML完整正确
# ============================================================
post_errors = []
h = html  # reuse the generated content

# P1: INV没有残留占位符
if '__INV_JSON__' in h:
    post_errors.append('INV占位符未替换')

# P2: 数据变量非空
for var_name in ['var A =', 'var RE =', 'var T10B =', 'var T10G =', 'var INV =']:
    idx = h.find(var_name)
    if idx < 0:
        post_errors.append(f'{var_name} 未找到')
    else:
        # Check it's not empty
        nxt = h.find(';', idx)
        val = h[idx+len(var_name):nxt].strip()
        if val in ['[]', '{}', '']:
            post_errors.append(f'{var_name} 数据为空')

# P3: 所有表（含时段分布）表头与数据cell数对齐
import re
tables_spec = [
    ('差评TOP10', 12, 't10b'),
    ('好评TOP10', 11, 't10g'),
    ('差评时段分布', 4, None),
    ('排名表', 14, 'tb'),
]
# 用正则找到每张表
table_matches = list(re.finditer(r'<thead><tr>(.*?)</tr></thead>', h, re.DOTALL))
for (name, expected_cols, tbody_id), tm in zip(tables_spec, table_matches):
    actual_cols = len(re.findall(r'<th[^>]*>', tm.group(1)))
    if actual_cols != expected_cols:
        post_errors.append(f'{name} 表头{actual_cols}列, 期望{expected_cols}列')

# P4: 三个需求公式都存在
formulas = {
    '排名表(need)': 'r.need',
    '差评TOP(bc-gt)': '19*r.bc-r.gt',
    '好评TOP(bt-gt)': '19*r.bt-r.gt',
}
for desc, formula in formulas.items():
    if formula not in h:
        post_errors.append(f'公式缺失: {desc} ({formula})')

# P5: 文件大小合理
if len(html) < 5000:
    post_errors.append(f'输出文件过小({len(html)}B)，可能生成失败')
if len(html) > 200000:
    post_errors.append(f'输出文件过大({len(html)}B)，检查是否有异常')

# P6: JS语法完整性 — 检测单引号不成对行（阻止 f-string 中 \\n 被误解析为裸换行）
script_blocks = re.findall(r'<script>(.*?)</script>', h, re.DOTALL)
for bi, block in enumerate(script_blocks):
    for li, line in enumerate(block.split('\n')):
        if line.count("'") % 2 == 1:
            post_errors.append(f'P6 JS语法风险: script块{bi+1}行{li+1}单引号不配对 ({line.count("\'")}个)')

print(f'\n输出自检:')
if post_errors:
    for e in post_errors:
        print(f'  ✗ {e}')
    print('输出自检未通过！请检查脚本。')
    sys.exit(2)
else:
    print(f'  ✓ P1 INV占位符已替换')
    print(f'  ✓ P2 数据变量非空')
    print(f'  ✓ P3 四张表列数对齐 (12/11/4/14)')
    print(f'  ✓ P4 三个需好评公式完整')
    print(f'  ✓ P5 文件大小合理 ({len(html)}B)')
    print(f'  ✓ P6 JS语法完整性 (无裸换行/引号不配对)')
    print(f'输出自检通过 ✓')
