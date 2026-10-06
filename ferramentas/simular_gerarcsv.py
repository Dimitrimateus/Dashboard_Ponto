import pickle, datetime, collections, unicodedata, sys, os
# Reproduz em Python a lógica do GerarCSVPonto.bas (versão de 06/10) a partir das abas
# de origem já carregadas por carregar_planilha.py. Serve para conferir o CSV sem Excel.
# Escrito para a planilha de 28/08 a 27/09 (aba única "Tratamento", "RE 08.09",
# "Pontonet", "Pausas Térmicas", "Cartão atual/Agosto/Julho"): na planilha com
# "Tratamento - Internos/Externos" ou outra RE, ajuste os nomes de aba abaixo.
# Saída: simulacao.pkl (lista de linhas) - use exportar_csv() para gerar o .csv.
d=pickle.load(open(os.environ.get('PLANILHA_PKL','planilha.pkl'),'rb'))
def nz(v):
    if isinstance(v,datetime.timedelta): return v.total_seconds()/86400
    if isinstance(v,datetime.time): return (v.hour*3600+v.minute*60+v.second)/86400
    if isinstance(v,(int,float)) and not isinstance(v,bool): return float(v)
    return 0.0
def s(v):
    if v is None: return ''
    if isinstance(v,str): v=v.replace('\r',' ').replace('\n',' ')
    if isinstance(v,float) and v.is_integer(): v=int(v)
    return str(v).strip()
def hdrmap(row): 
    m={}
    for i,h in enumerate(row):
        h=s(h)
        if h and h not in m: m[h]=i
    return m
def vround(x):  # VBA Round = banker's
    return round(x)
EXCL_PROB = '--exclui-problema' in sys.argv
CHECK_S_TOTAL = '--check-s-total' in sys.argv
def norm(t): return unicodedata.normalize('NFD',t).encode('ascii','ignore').decode().lower().strip()

RE=d['RE 08.09']; rh=hdrmap(RE[0])
dictRE={}
for r in RE[1:]:
    m=s(r[rh['Matrícula']])
    if m and m not in dictRE: dictRE[m]=(s(r[rh['Nome Unidade']]), s(r[rh['Cargo']]))
def setorcargo(m):
    if m in dictRE: a,b=dictRE[m]; return (a or 'Sem setor', b)
    return ('Sem setor','')
# cartao sheets
def cart(name):
    rows=d[name]; h=hdrmap(rows[1]); return rows,h
cards=[]
for n in d:
    if n.lower().startswith(('cartão','cartao')):
        rows,h=cart(n)
        if 'Matrícula' in h and 'DT' in h:
            mx=max((r[h['DT']] for r in rows[2:] if isinstance(r[h['DT']],datetime.datetime)),default=None)
            cards.append((mx,n))
cards.sort(reverse=True); cards=cards[:3]
cur=cards[0][1]; rows,h=cart(cur)
dictCart={}
for r in rows[2:]:
    m=s(r[h['Matrícula']]); dt=r[h['DT']]
    if m and isinstance(dt,datetime.datetime):
        k=(m,dt.date()); v=(nz(r[h['BH']])*1440, nz(r[h['50%']])*1440, nz(r[h['100%']])*1440)
        if k in dictCart: o=dictCart[k]; dictCart[k]=(o[0]+v[0],o[1]+v[1],o[2]+v[2])
        else: dictCart[k]=v
# pontonet
P=d['Pontonet']; ph=hdrmap(P[0]); dictAus={}
for r in P[1:]:
    m=s(r[ph['Matrícula']]); dt=r[ph['Data falta']]
    if m and isinstance(dt,datetime.datetime):
        dictAus[(m,dt.date())]=(s(r[ph['Justificativa']]), r[ph['Avaliado em:']], s(r[ph['Situação atual']]).upper()=='INTEGRADO')
T=d['Tratamento']; th=hdrmap(T[0])
out=[]; gestorPor={}
def W(dt,nome,mat,gestor,setor,cargo,tipo,sit,status,dur,trat=None,**kw):
    out.append(dict(data=dt,colaborador=nome,matricula=mat,gestor=gestor,setor=setor,cargo=cargo,tipo=tipo,situacao=sit,status=status,duracao=dur,tratativa=trat,**kw))
def mapst(t): return 'Regularizado' if t.upper()=='TRATADO' else 'Pendente'
for r in T[1:]:
    mat=s(r[th['Matricula']]); nome=s(r[th['Nome']]); gestor=s(r[th['Gestor']]); gestor='' if gestor.startswith('#') else gestor; sit=s(r[th['Situação']]); st=s(r[th['Status']]); dt=r[th['Data']]; chk=s(r[th['Check']])
    if nome=='' or not isinstance(dt,datetime.datetime): continue
    if gestor and mat and mat not in gestorPor: gestorPor[mat]=gestor
    fb,f50,f100=dictCart.get((mat,dt.date()),(0,0,0))
    setor,cargo=setorcargo(mat); ct=s(r[th['Cargo']]) 
    if ct: cargo=ct
    jus,av,integ=dictAus.get((mat,dt.date()),('',None,False))
    temTrat = integ and isinstance(av,datetime.datetime)
    status=mapst(st); sitF = jus or sit
    if chk.upper()=='S': continue
    if sit in ('','Sem alteração'): continue
    if norm(sit)=='problema horario': continue
    if f100>0:
        W(dt,nome,mat,gestor,setor,cargo,'Hora Extra 100%',sitF,status,vround(f100))
    me=nz(r[th['Extras']])*1440; m100=nz(r[th['Extra 100%']])*1440; mf=nz(r[th['Faltas']])*1440
    te=(me+m100)>0; tf=mf>0
    mE = f50 if f50>0 else me+m100; mF = fb if fb>0 else mf
    if te: W(dt,nome,mat,gestor,setor,cargo,'Hora Extra',sitF,status,vround(mE))
    if tf: W(dt,nome,mat,gestor,setor,cargo,sit,sitF,status,-vround(mF),av if temTrat else None)
    if not tf: W(dt,nome,mat,gestor,setor,cargo,sit,sitF,status,0,av if temTrat else None)
# ORG (LerListaORG): matrícula repetida = 1ª linha; matrícula em erro entra sem matrícula,
# a não ser que o nome já exista com matrícula válida. Completa o gestor de quem não está na Tratamento.
def semac(t): return unicodedata.normalize('NFD',t).encode('ascii','ignore').decode().upper()
listaORG=[]
if 'ORG' in d:
    Og=d['ORG']; oh={semac(s(x)):i for i,x in enumerate(Og[0]) if s(x)}
    def oc(r,k): 
        v=s(r[oh[k]]) if k in oh else ''
        return '' if v.startswith('#') else v
    vm=set(); vn=set()
    for passo in (1,2):
        for r in Og[1:]:
            nm=s(r[oh['COLABORADOR']])
            if not nm or nm.startswith('#'): continue
            m=oc(r,'MATRICULA')
            if passo==1 and m and m not in vm:
                vm.add(m); vn.add(semac(nm)); listaORG.append((m,nm,oc(r,'DESCRICAO'),oc(r,'CARGO'),oc(r,'GESTOR')))
            elif passo==2 and not m and semac(nm) not in vn:
                vn.add(semac(nm)); listaORG.append(('',nm,oc(r,'DESCRICAO'),oc(r,'CARGO'),oc(r,'GESTOR')))
for m,nm,ar,ca,g in listaORG:
    if m and g and m not in gestorPor: gestorPor[m]=g
# auditoria
nomes={}; cont=collections.defaultdict(dict)
for _,n in cards:
    rows2,h2=cart(n)
    for r in rows2[2:]:
        m=s(r[h2['Matrícula']])
        if not m: continue
        if s(r[h2['Banco de horas e extra no mesmo dia']]).lower()=='sim':
            cont[m][n]=cont[m].get(n,0)+1
        if m not in nomes and s(r[h2['Nome']]): nomes[m]=s(r[h2['Nome']])
dmax=cards[0][0]
for m,c in cont.items():
    tot=sum(c.values()); 
    if tot>0:
        se,ca=setorcargo(m)
        W(dmax,nomes.get(m,'Matrícula '+m),m,gestorPor.get(m,''),se,ca,'Auditoria Extra e Falta (3 Meses)','','Pendente',0,mesAtual=c.get(cur,0),total3=tot)
# curtas
for r in rows[2:]:
    tol=s(r[h['Tolerância < 15min']])
    if not tol: continue
    m=s(r[h['Matrícula']]); dt=r[h['DT']]
    if not m or not isinstance(dt,datetime.datetime): continue
    nome=s(r[h['Nome']]) or 'Matrícula '+m; se,ca=setorcargo(m); g=gestorPor.get(m,'')
    fb=nz(r[h['BH']])*1440; ex=(nz(r[h['50%']])+nz(r[h['100%']]))*1440
    if 'falta' in tol.lower(): W(dt,nome,m,g,se,ca,'Falta < 15min','Falta < 15min','Pendente',-vround(fb))
    if 'extra' in tol.lower(): W(dt,nome,m,g,se,ca,'Extra < 15min','Extra < 15min','Pendente',vround(ex))
# resumo de horas do Cartão atual (GerarResumoHorasCartao): extra = 50% + 100%, banco = BH
hrs={}; nomesH={}
for r in rows[2:]:
    m=s(r[h['Matrícula']])
    if not m: continue
    ex=(nz(r[h['50%']])+nz(r[h['100%']]))*1440; bh=nz(r[h['BH']])*1440
    if m in hrs: hrs[m]=(hrs[m][0]+ex, hrs[m][1]+bh)
    else: hrs[m]=(ex,bh); nomesH[m]=s(r[h['Nome']])
for m,(ex,bh) in hrs.items():
    if vround(ex)>0 or vround(bh)>0:
        se,ca=setorcargo(m)
        W(dmax,nomesH[m] or 'Matrícula '+m,m,gestorPor.get(m,''),se,ca,'Resumo Horas Cartão','','Pendente',0,hx=vround(ex),hb=vround(bh))
# pausas
Pz=d['Pausas Térmicas']; pz=hdrmap(Pz[0])
for r in Pz[1:]:
    m=s(r[pz['Matrícula']])
    if not m: break
    se,ca=setorcargo(m)
    W(dmax,s(r[pz['Colaborador']]),m,gestorPor.get(m,''),se,s(r[pz['Cargo']]) or ca,'Resumo Pausas Térmicas','','Pendente',0,
      pc=nz(r[pz['Pausa corretas']]),pm=nz(r[pz['Pausa menor 0:20']]),pM=nz(r[pz['Pausa maior 0:20']]),tc=nz(r[pz['Trabalho correto 1:40']]),tM=nz(r[pz['Trabalho maior 1:40']]),tm=nz(r[pz['Trabalho menor 1:40']]),imp=nz(r[pz['Marcações Ímpares']]))
# marcações digitadas (GerarLinhasDigitadas): aba achada pelo cabeçalho; data = dia + hora (minuto)
for nome_aba,rows_d in d.items():
    if not rows_d: continue
    hd=hdrmap(rows_d[0])
    if all(k in hd for k in ('Matrícula','Data','Hora','Motivo','Justificativa')):
        for r in rows_d[1:]:
            m=s(r[hd['Matrícula']]); dt=r[hd['Data']]
            if not m or not isinstance(dt,datetime.datetime): continue
            if 'Origem' in hd and s(r[hd['Origem']]).upper()=='E': continue   # Origem "E" é ignorada
            hr=r[hd['Hora']]
            if isinstance(hr,datetime.time): dt=dt+datetime.timedelta(minutes=round(hr.hour*60+hr.minute+hr.second/60+hr.microsecond/6e7))
            se,caRE=setorcargo(m)
            W(dt,s(r[hd['Colaborador']]) or 'Matrícula '+m,m,gestorPor.get(m,''),se,caRE or s(r[hd['Cargo']]),'Marcação Digitada',
              s(r[hd['Motivo']]),'Pendente',0,just=s(r[hd['Justificativa']]),orig=s(r[hd['Origem']]) if 'Origem' in hd else '')
        break
for m,nm,ar,ca,g in listaORG:
    se,caRE=setorcargo(m) if m else ('Sem setor','')
    if se=='Sem setor' and ar: se=ar
    W(dmax,nm,m,g,se,ca or caRE,'Cadastro ORG',ar,'Pendente',0)
pickle.dump(out,open('simulacao.pkl','wb'))
print('cartoes',cards); print(len(out), collections.Counter(o['tipo'] for o in out))

def exportar_csv(linhas, caminho='dados_TODOS.csv'):
    """Mesmo formato do ExportarLinhasParaCSV: vírgula, UTF-8 com BOM, datas dd/mm/aaaa."""
    cab=["data","colaborador","matricula","gestor","setor","cargo","tipo_ocorrencia","situacao","status","duracao_minutos","destino_horas_extra","data_tratativa_pontonet","horas_excedentes","ocorrencias_extra_falta_mes_atual","ocorrencias_extra_falta_3_meses","pausas_corretas","pausas_menor_20min","pausas_maior_20min","trabalho_correto_140","trabalho_maior_140","trabalho_menor_140","pausas_marcacoes_impares","grupo","minutos_hora_extra_cartao","minutos_banco_horas_cartao","justificativa_marcacao","origem_marcacao"]
    def t(v):
        if v is None: return ''
        if isinstance(v,datetime.datetime): return v.strftime('%d/%m/%Y') if v.hour==0 and v.minute==0 else v.strftime('%d/%m/%Y %H:%M')
        if isinstance(v,float) and v.is_integer(): v=int(v)
        return str(v)
    q=lambda x: '"'+x.replace('"','""')+'"' if (',' in x or '"' in x or '\n' in x) else x
    out=[','.join(cab)]
    for o in linhas:
        v=[o['data'],o['colaborador'],o['matricula'],o['gestor'],o['setor'],o['cargo'],o['tipo'],o['situacao'],o['status'],o['duracao'],'',o['tratativa'],'',o.get('mesAtual'),o.get('total3'),o.get('pc'),o.get('pm'),o.get('pM'),o.get('tc'),o.get('tM'),o.get('tm'),o.get('imp'),o.get('grupo',''),o.get('hx'),o.get('hb'),o.get('just'),o.get('orig')]
        out.append(','.join(q(t(x)) for x in v))
    open(caminho,'w',encoding='utf-8-sig').write('\n'.join(out)+'\n')

if __name__=='__main__':
    exportar_csv(out)
    print('dados_TODOS.csv gerado com', len(out), 'linhas')
