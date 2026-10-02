import pickle,json,collections,datetime,sys,unicodedata,os
# Confere, item a item, os números que o painel mostrou (JSON gerado por extrair_painel.js)
# contra os valores calculados DIRETO das abas de origem (sem passar pelo CSV).
# Uso: python3 conferir_painel.py painel.json   (precisa de planilha.pkl)
# Mesmas observações de nomes de aba do simular_gerarcsv.py.
from simular_gerarcsv import nz,s,hdrmap
d=pickle.load(open(os.environ.get('PLANILHA_PKL','planilha.pkl'),'rb'))
ui=json.load(open(sys.argv[1]))
C=collections.Counter
def norm(t): return unicodedata.normalize('NFD',t).encode('ascii','ignore').decode().lower().strip()
T=d['Tratamento']; th=hdrmap(T[0])
RE=d['RE 08.09']; rh=hdrmap(RE[0]); recargo={s(r[rh['Matrícula']]):s(r[rh['Cargo']]) for r in RE[1:]}
E=[]
for r in T[1:]:
    if not s(r[th['Nome']]) or not isinstance(r[th['Data']],datetime.datetime): continue
    if s(r[th['Check']]).upper()=='S': continue
    sit=s(r[th['Situação']])
    if sit in ('','Sem alteração') or norm(sit)=='problema horario': continue
    g=s(r[th['Gestor']]); g='Sem gestor' if (not g or g.startswith('#')) else g
    mat=s(r[th['Matricula']]); cg=s(r[th['Cargo']])
    if not cg or cg.startswith('#'): cg=recargo.get(mat,'') or 'Sem cargo'
    E.append(dict(nome=s(r[th['Nome']]),mat=mat,sit=sit,data=r[th['Data']],gestor=g,cargo=cg,
       extra=(nz(r[th['Extras']])+nz(r[th['Extra 100%']]))*1440, falta=nz(r[th['Faltas']])*1440))
res=[]
def chk(nome,ok,det=''): res.append((nome,ok,det)); print(('OK   ' if ok else 'FALHA')+' '+nome+(' -> '+det if det else ''))
def tips(lst): return {t.rsplit(': ',1)[0]: int(t.rsplit(': ',1)[1].replace('.','')) for t in lst}
kp={k[0]:k for k in ui['kpis']}
chk('KPI Ocorrências no período', kp['Ocorrências no período'][1]==str(len(E)), f"painel {kp['Ocorrências no período'][1]} / fonte {len(E)}")
chk('KPI Colaboradores envolvidos', kp['Colaboradores envolvidos'][1]==str(len({e['nome'] for e in E})), f"painel {kp['Colaboradores envolvidos'][1]} / fonte {len({e['nome'] for e in E})}")
exp=C(e['sit'] for e in E); got={a:int(b) for a,b in ui['tipo']}
chk('Ocorrências por tipo', got==dict(exp), f"painel {got} / fonte {dict(exp)}" if got!=dict(exp) else str(dict(exp)))
lab=['Seg','Ter','Qua','Qui','Sex','Sáb','Dom']; exp={lab[e['data'].weekday()]:0 for e in E}
expd=C(lab[e['data'].weekday()] for e in E); got=tips(ui['diaSemana']); got={k:v for k,v in got.items() if v}
chk('Ocorrências por dia da semana', got==dict(expd), str(dict(expd)))
expd=C(e['data'].strftime('%d/%m') for e in E); got=tips(ui['porData'])
chk('Ocorrências por data', got==dict(expd), f"{len(got)} dias, soma {sum(got.values())}")
def top(counter,n):
    items=counter.most_common(); 
    if len(items)<=n: return dict(items)
    cut=items[n-1][1]; return {k:v for k,v in items if v>cut}, cut
def cmp_top(nome,got,counter,n):
    gotd=tips(got); full=dict(counter)
    ok = all(full.get(k)==v for k,v in gotd.items()) and len(gotd)==min(n,len(full))
    mins=min(gotd.values()) if gotd else 0
    ok = ok and all(v<=mins for k,v in full.items() if k not in gotd)
    chk(nome, ok, f"top {len(gotd)} confere com a fonte" if ok else f"painel {gotd} / fonte {counter.most_common(n+2)}")
cmp_top('Ocorrências por gestor (top 12)', ui['gestor'], C(e['gestor'] for e in E), 12)
cmp_top('Ocorrências por cargo (top 12)', ui['cargo'], C(e['cargo'] for e in E), 12)
cmp_top('Colaboradores com mais ocorrências (top 10)', ui['top'], C(e['nome'] for e in E), 10)
chk('Histórico (nº de linhas)', ui['histCount'].startswith(str(len(E))+' '), ui['histCount'])
# banco de horas / hora extra: minutos de extra (Cartão 50%+100% quando existir, senão Tratamento)
cur=d['Cartão atual']; ch=hdrmap(cur[1]); cart=collections.defaultdict(lambda:[0,0,0])
for r in cur[2:]:
    m=s(r[ch['Matrícula']]); dt=r[ch['DT']]
    if m and isinstance(dt,datetime.datetime):
        k=(m,dt.date()); cart[k][0]+=nz(r[ch['BH']])*1440; cart[k][1]+=nz(r[ch['50%']])*1440; cart[k][2]+=nz(r[ch['100%']])*1440
hx=C()
for e in E:
    bh,e50,e100=cart.get((e['mat'],e['data'].date()),(0,0,0))
    if e['extra']>0: hx[e['nome']]+=round(e50 if e50>0 else e['extra'])
    if e100>0: hx[e['nome']]+=round(e100)
def hm(t):
    t=t.split(': ',1)[1]; h=m=0
    for p in t.split():
        if p.endswith('min'): m=int(p[:-3])
        elif p.endswith('h'): h=int(p[:-1])
    return h*60+m
gotb=collections.Counter()
for t in ui['banco']: gotb[t.split(' · ')[0]]+=hm(t)
ok=all(hx[k]==v for k,v in gotb.items()) and len(gotb)==min(10,len(hx))
chk('Banco de horas e hora extra por colaborador', ok, f"{len(gotb)} colaboradores, minutos batem" if ok else f"painel {dict(gotb)} / fonte {hx.most_common(10)}")
# extra e falta mesmo dia
cards=['Cartão atual','Cartão Agosto','Cartão Julho']; tot=C(); mes=C(); nomes={}
for n in cards:
    rows=d[n]; h=hdrmap(rows[1])
    for r in rows[2:]:
        m=s(r[h['Matrícula']])
        if m and s(r[h['Banco de horas e extra no mesmo dia']]).lower()=='sim':
            tot[m]+=1; 
            if n=='Cartão atual': mes[m]+=1
        if m and m not in nomes: nomes[m]=s(r[h['Nome']])
byname={nomes[m]:(mes[m],tot[m]) for m in tot}
ok=all(byname.get(r[0])==(int(r[2]),int(r[3].split()[0])) for r in ui['extraFalta'])
mn=min(int(r[3].split()[0]) for r in ui['extraFalta'])
ok=ok and all(v[1]<=mn for k,v in byname.items() if k not in {r[0] for r in ui['extraFalta']})
alert=all((('⚠' in r[3])==(int(r[3].split()[0])>=10)) for r in ui['extraFalta'])
chk('Extra e falta no mesmo dia (mês / 3 meses / alerta 10+)', ok and alert, f"{sum(1 for v in byname.values() if v[1]>=10)} colaboradores com 10+ na fonte")
# curtas
cu=C()
for r in cur[2:]:
    tol=s(r[ch['Tolerância < 15min']]); m=s(r[ch['Matrícula']])
    if tol and m and isinstance(r[ch['DT']],datetime.datetime):
        if 'falta' in tol.lower(): cu[s(r[ch['Nome']])]+=1
        if 'extra' in tol.lower(): cu[s(r[ch['Nome']])]+=1
gotc={r[0]:int(r[-1]) for r in ui['curtas']}
ok=all(cu[k]==v for k,v in gotc.items())
chk('Ocorrências curtas (< 15 min)', ok, "confere" if ok else "painel "+str({k:(v,cu[k]) for k,v in gotc.items() if cu[k]!=v})+" (painel, fonte)")
# pausas
P=d['Pausas Térmicas']; ph=hdrmap(P[0]); src={}
for r in P[1:]:
    if not s(r[ph['Matrícula']]): break
    src[s(r[ph['Colaborador']])]=r
ok=len(ui['pausas'])==len(src)
mism=[]
for row in ui['pausas']:
    r=src.get(row[0]); hdr=ui.get('pausasHdr')
    vals=[int(nz(r[ph[c]])) for c in ['Pausa corretas','Pausa menor 0:20','Trabalho correto 1:40','Trabalho maior 1:40','Trabalho menor 1:40','Marcações Ímpares']]
    shown=[int(x) for x in (row[2:4]+row[-4:] if len(row)==8 else row[2:4]+row[-4:])]
    if vals!=shown: mism.append((row[0],shown,vals))
chk('Pausas Térmicas (tabela)', ok and not mism, f"{len(ui['pausas'])} colaboradores" + (f"; divergências {mism[:3]}" if mism else ''))
# demora
Pn=d['Pontonet']; pnh=hdrmap(Pn[0]); aus={}
for r in Pn[1:]:
    m=s(r[pnh['Matrícula']]); dt=r[pnh['Data falta']]
    if m and isinstance(dt,datetime.datetime): aus[(m,dt.date())]=(r[pnh['Avaliado em:']], s(r[pnh['Situação atual']]).upper()=='INTEGRADO')
g=collections.defaultdict(list)
for e in E:
    av,integ=aus.get((e['mat'],e['data'].date()),(None,False))
    if integ and isinstance(av,datetime.datetime):
        h=(av-e['data']).total_seconds()/3600
        if h>=0: g[e['gestor']].append(h)
def fd(h): return (f"{h:.1f}h" if h<24 else f"{h/24:.1f} dias").replace('.',',')
ok=True; bad=[]
for row in ui['demoraG']:
    l=g.get(row[0],[]); exp=[str(len(l)), fd(sum(l)/len(l)) if l else '-', fd(max(l)) if l else '-']
    if row[1:]!=exp: ok=False; bad.append((row,exp))
n=sum(len(v) for v in g.values())
chk('Demora no PontoNet por gestor', ok and len(ui['demoraG'])==len(g), f"{n} ocorrências tratadas; KPI diz {kp['Demora média no PontoNet'][2]}" + (f"; {bad[:3]}" if bad else ''))
allh=[h for v in g.values() for h in v]
chk('KPI Demora média', kp['Demora média no PontoNet'][1]==fd(sum(allh)/len(allh)), f"painel {kp['Demora média no PontoNet'][1]} / fonte {fd(sum(allh)/len(allh))}")
