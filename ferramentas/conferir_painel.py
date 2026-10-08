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
def tips(lst):  # ignora os complementos " (...)" e " · clique para ver o colaborador"
    lst=[t.split(' · ')[0].rsplit(': ',1) for t in lst]
    return {n: int(v.split(' (')[0].replace('.','')) for n,v in lst}
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
# banco de horas / hora extra: Cartão atual inteiro, por colaborador (extra = 50%+100%, banco = BH)
def hm(t):
    h=m=0
    for p in t.replace('\u2212','').replace('+','').split():
        if p.endswith('min'): m=int(p[:-3])
        elif p.endswith('h'): h=int(p[:-1])
    return h*60+m
src=collections.defaultdict(lambda:[0.0,0.0]); nomeMat={}
for r in cur[2:]:
    m=s(r[ch['Matrícula']])
    if not m: continue
    src[m][0]+=(nz(r[ch['50%']])+nz(r[ch['100%']]))*1440; src[m][1]+=nz(r[ch['BH']])*1440
    nomeMat.setdefault(m, s(r[ch['Nome']]))
fonte={nomeMat[m]:(round(v[0]),round(v[1])) for m,v in src.items() if round(v[0])>0 or round(v[1])>0}
gotb={}
for t in ui['banco']:
    p=t.split(' · ')
    if len(p)==4: gotb[p[0]]=(hm(p[1].split(': ')[1]), hm(p[2].split(': ')[1]), p[3].split(': ')[1])
okv=all(fonte.get(k)==v[:2] for k,v in gotb.items())
oks=all((v[0]-v[1]==0 and v[2]=='0h') or (v[2].startswith('+')==(v[0]>v[1]) and hm(v[2])==abs(v[0]-v[1])) for v in gotb.values())
corte=min(sum(v[:2]) for v in gotb.values()) if gotb else 0
okt=len(gotb)==min(10,len(fonte)) and all(sum(v)<=corte for k,v in fonte.items() if k not in gotb)
chk('Banco de horas e hora extra por colaborador (Cartão atual)', okv and oks and okt,
    f"top {len(gotb)} de {len(fonte)} colaboradores; extra, banco e saldo batem com o Cartão" if okv and oks and okt else f"painel {gotb} / fonte {sorted(fonte.items(), key=lambda kv:-sum(kv[1]))[:10]}")
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
    l=g.get(row[0].split('⚠')[0].strip(),[]); exp=[str(len(l)), fd(sum(l)/len(l)) if l else '-', fd(max(l)) if l else '-']
    if row[1:]!=exp: ok=False; bad.append((row,exp))
n=sum(len(v) for v in g.values())
chk('Demora no PontoNet por gestor', ok and len(ui['demoraG'])==len(g), f"{n} ocorrências tratadas; KPI diz {kp['Demora média no PontoNet'][2]}" + (f"; {bad[:3]}" if bad else ''))
allh=[h for v in g.values() for h in v]
chk('KPI Demora média', kp['Demora média no PontoNet'][1]==fd(sum(allh)/len(allh)), f"painel {kp['Demora média no PontoNet'][1]} / fonte {fd(sum(allh)/len(allh))}")
# lista do ORG: pessoas únicas do ORG (o aviso de "fora do ORG" saiu do painel a pedido do RH)
if 'ORG' in d and ui.get('orgResumo'):
    def _sa(t): return unicodedata.normalize('NFD',t).encode('ascii','ignore').decode().upper()
    O=d['ORG']; oh={_sa(s(x)):i for i,x in enumerate(O[0]) if s(x)}
    mats=set(); nomesO=set()
    for r in O[1:]:
        nm=s(r[oh['COLABORADOR']]); m=s(r[oh['MATRICULA']])
        if nm and m and not m.startswith('#'): mats.add(m); nomesO.add(_sa(nm))
    semmat={_sa(s(r[oh['COLABORADOR']])) for r in O[1:] if s(r[oh['COLABORADOR']]) and (not s(r[oh['MATRICULA']]) or s(r[oh['MATRICULA']]).startswith('#'))} - nomesO
    total=len(mats)+len(semmat)
    okn=ui['orgResumo'].split()[0]==str(total) and ui['orgLinhas']==total
    chk('Lista de colaboradores (ORG)', okn, f"{total} no ORG" if okn else f"painel {ui['orgResumo']!r} {ui['orgLinhas']} / fonte {total}")
# marcações digitadas: total, pessoas, por motivo e top 20, direto da aba formatada
abaDig=next((n for n,rw in d.items() if rw and all(k in hdrmap(rw[0]) for k in ('Matrícula','Data','Hora','Motivo','Justificativa'))), None)
if abaDig and ui.get('digKpis'):
    rw=d[abaDig]; hd=hdrmap(rw[0])
    linhasD=[r for r in rw[1:] if s(r[hd['Matrícula']]) and isinstance(r[hd['Data']],datetime.datetime)
             and not ('Origem' in hd and s(r[hd['Origem']]).upper()=='E')]
    mot=C(s(r[hd['Motivo']]) for r in linhasD); pes=C(s(r[hd['Colaborador']]) for r in linhasD)
    k=ui['digKpis'].split('\n')
    okk=k[1]==str(len(linhasD)) and k[4]==str(len({s(r[hd['Matrícula']]) for r in linhasD})) and ui['digLinhas']==len(linhasD)
    okm={a:int(b) for a,b in ui['digMotivo']}==dict(mot)
    gt=tips(ui['digTop']); corte=min(gt.values()) if gt else 0
    okt=all(pes[n]==v for n,v in gt.items()) and len(gt)==min(20,len(pes)) and all(v<=corte for n,v in pes.items() if n not in gt)
    chk('Marcações digitadas (total, pessoas, motivo, top 20)', okk and okm and okt,
        f"{len(linhasD)} batidas de {len({s(r[hd['Matrícula']]) for r in linhasD})} pessoas; {dict(mot)}" if okk and okm and okt else f"painel {k} {ui['digMotivo']} {gt} / fonte {len(linhasD)} {dict(mot)} {pes.most_common(10)}")
# absenteísmo (atestados): dias por ciclo, pessoas por ciclo, encostados em folga e reincidentes,
# recalculados direto dos 3 Cartões (dia a dia), sem passar pelo CSV
if ui.get('absEvol'):
    def _sa2(t): return unicodedata.normalize('NFD',t).encode('ascii','ignore').decode().lower()
    dias=collections.defaultdict(dict); nomesA={}
    for n in ['Cartão Julho','Cartão Agosto','Cartão atual']:
        rw=d[n]; h=hdrmap(rw[1])
        for r in rw[2:]:
            m=s(r[h['Matrícula']]); dt=r[h['DT']]
            if m and isinstance(dt,datetime.datetime):
                dias[m].setdefault(dt.date(), s(r[h['Descrição Marcação']])); nomesA.setdefault(m, s(r[h['Nome']]))
    def ciclo(x): return (x.year+(1 if x.month==12 and x.day>=28 else 0), (x.month % 12)+1 if x.day>=28 else x.month)
    porC=collections.defaultdict(lambda:[0,set()]); eps=[]
    fol=lambda t: any(w in _sa2(t) for w in ('dsr','compensado','folga','feriado'))
    abaAf=next((n for n,rw in d.items() if rw and all(k in hdrmap(rw[0]) for k in ('Matrícula','Situação','Início','Término'))), None)
    if abaAf:   # fonte = relatório de afastamentos do sistema (1 registro = 1 atestado)
        rw=d[abaAf]; ha=hdrmap(rw[0]); um=datetime.timedelta(days=1)
        for r in rw[1:]:
            m=s(r[ha['Matrícula']]); ini=r[ha['Início']]
            if not m or not isinstance(ini,datetime.datetime) or 'atest' not in _sa2(s(r[ha['Situação']])): continue
            fim=r[ha['Término']] if isinstance(r[ha['Término']],datetime.datetime) else ini
            x=ini.date()
            while x<=fim.date(): porC[ciclo(x)][0]+=1; porC[ciclo(x)][1].add(m); x+=um
            dd=dias.get(m,{}); eps.append((m,ini.date(),fim.date(),fol(dd.get(ini.date()-um,'')) or fol(dd.get(fim.date()+um,''))))
            nomesA[m]=s(r[ha['Colaborador']])
    for m,dd in (dias.items() if not abaAf else []):
        at=sorted(k for k,v in dd.items() if 'atest' in _sa2(v))
        for k in at: porC[ciclo(k)][0]+=1; porC[ciclo(k)][1].add(m)
        i=0
        while i<len(at):
            ini=fim=at[i]
            while i+1<len(at) and (at[i+1]-fim).days==1: i+=1; fim=at[i]
            um=datetime.timedelta(days=1)
            eps.append((m,ini,fim,fol(dd.get(ini-um,'')) or fol(dd.get(fim+um,'')))); i+=1
    esp=[[str(v[0]),str(len(v[1]))] for k,v in sorted(porC.items())]
    got=[[r[1],r[3]] for r in ui['absEvol']]
    emend=sum(1 for e in eps if e[3])
    rk=collections.defaultdict(lambda:[0,set()])
    for m,ini,fim,_ in eps:
        rk[m][0]+=1
        x=ini
        while x<=fim: rk[m][1].add(ciclo(x)); x+=datetime.timedelta(days=1)
    reinc=sorted(nomesA[m] for m,v in rk.items() if v[0]>=3 or len(v[1])>=2)
    ok=got==esp and ui['absEmenda'].startswith(f"{emend} de {len(eps)} ") and ui['absRankN']==len(rk) and sorted(ui['absReinc'])==reinc
    chk('Absenteísmo (atestados: dias/pessoas por ciclo, encostados em folga, reincidentes)', ok,
        f"ciclos {esp}; {emend} de {len(eps)} atestados encostam em folga; {len(reinc)} reincidentes" if ok else f"painel {got} {ui['absEmenda']} {ui['absRankN']} {ui['absReinc']} / fonte {esp} {emend}/{len(eps)} {len(rk)} {reinc}")
