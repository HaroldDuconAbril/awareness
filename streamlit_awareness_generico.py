
import io, re, json, unicodedata
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BANKS=['Banco Agrario','Bancolombia','Davivienda','Banco de Bogotá','BBVA','Banco Caja Social','Banco Popular','Bancamía','Banco de Occidente','Banco Mundo Mujer']
BANK_ALIASES={'Banco Agrario':['banco agrario','agrario'],'Bancolombia':['bancolombia'],'Davivienda':['davivienda'],'Banco de Bogotá':['banco de bogota','banco de bogotá','banco bogota','banco bogotá'],'BBVA':['bbva','bbvva','bvva','bbwa','bva'],'Banco Caja Social':['banco caja social','caja social'],'Banco Popular':['banco popular','popular'],'Bancamía':['bancamia','bancamía','banca mia','banca mía'],'Banco de Occidente':['banco de occidente','occidente'],'Banco Mundo Mujer':['banco mundo mujer','mundo mujer','banco de la mujer','fundacion de la mujer','fundación de la mujer']}
BANK_NORMALIZATIONS=[{'pattern':r'(?i)Banco\s+de\s+Bogota','replacement':'Banco de Bogotá'},{'pattern':r'(?i)(?<!Banco\s)Caja\s+Social','replacement':'Banco Caja Social'},{'pattern':r'(?i)Bancamia','replacement':'Bancamía'}]
NEG={'','0','no','nan','none','false','ninguno','ninguna','no se','no sé','no recuerdo','no aplica'}

def norm(x):
    if pd.isna(x): return ''
    x=str(x).strip().lower(); x=unicodedata.normalize('NFKD',x).encode('ascii','ignore').decode('ascii')
    return re.sub(r'\s+',' ',re.sub(r'[^a-z0-9\s]+',' ',x)).strip()

def find_col(df,prefix):
    p=norm(prefix)
    for c in df.columns:
        if norm(c).startswith(p): return c
    return None

def find_cols(df,txt):
    return [c for p in [x.strip() for x in txt.splitlines() if x.strip()] for c in [find_col(df,p)] if c]

def clean(x,norms):
    if pd.isna(x): return x
    s=str(x)
    for it in norms:
        try: s=re.sub(it.get('pattern',''),it.get('replacement',''),s)
        except Exception: pass
    return s

def contains(x,entity,aliases):
    t=norm(x)
    if t in NEG: return False
    for a in aliases.get(entity,[entity]):
        aa=norm(a)
        if aa and re.search(r'(^|\s)'+re.escape(aa)+r'(\s|$)',t): return True
    return False

def entity_from_col(col,entities):
    t=norm(str(col).replace('\n',' '))
    for e in entities:
        if norm(e) in t: return e
    return None

def safe(s): return re.sub(r'[\\/*?:\[\]]',' ',str(s)).strip()[:31] or 'Hoja'

def build(df, demo, cfg, entities, aliases, norms, expected):
    tom = find_col(df, cfg['tom'])
    esp = find_cols(df, cfg['esp'])
    ayu = find_cols(df, cfg['ayu'])

    raw = [demo.get('edad'), demo.get('departamento'), demo.get('estrato'), tom] + esp + ayu
    raw = [c for c in raw if c]

    if expected and len(raw) != expected:
        raise ValueError(
            f"{cfg['name']}: se esperaban {expected} columnas crudas y se detectaron {len(raw)}"
        )

    if not tom or not esp or not ayu:
        raise ValueError(f"{cfg['name']}: faltan columnas TOM/Espontáneo/Ayudado")

    rdf = df[raw].apply(lambda col: col.map(lambda x: clean(x, norms)))
    out = rdf.copy()
    ind = {}

    amap = {e: None for e in entities}

    for c in ayu:
        e = entity_from_col(c, entities)
        if e:
            amap[e] = c

    for e in entities:
        T = rdf[tom].apply(lambda x: contains(x, e, aliases))

        E = pd.Series(False, index=df.index)
        for c in esp:
            E = E | rdf[c].apply(lambda x: contains(x, e, aliases))

        if amap.get(e):
            A = rdf[amap[e]].apply(lambda x: contains(x, e, aliases))
        else:
            
            A = pd.Series(False, index=df.index)

        ind[(e, 'TOM')] = T.astype(int)
        ind[(e, 'Espontaneo')] = ((~T) & E).astype(int)
        ind[(e, 'Ayudado')] = ((~T) & (~E) & A).astype(int)

        out[(e, 'TOM')] = ind[(e, 'TOM')]
        out[(e, 'Espontaneo')] = ind[(e, 'Espontaneo')]
        out[(e, 'Ayudado')] = ind[(e, 'Ayudado')]

    return out, pd.DataFrame(ind), raw

def tab(ind,df,col,val,entities):
    m=(df[col].fillna('Sin dato')==val) if col else pd.Series(True,index=df.index); base=int(m.sum())
    rows=[]
    for e in entities:
        t=int(ind[(e,'TOM')][m].sum()); s=int(ind[(e,'Espontaneo')][m].sum()); a=int(ind[(e,'Ayudado')][m].sum()); awa=t+s+a
        rows.append({'Entidad':e,'TOM':t,'TOM %':t/base if base else 0,'Espontaneo':s,'Espontaneo %':s/base if base else 0,'Ayudado':a,'Ayudado %':a/base if base else 0,'AWA':awa,'AWA %':awa/base if base else 0})
    return pd.DataFrame(rows)

def sections(ind,df,demo,name,entities):
    sec=[(f'Resumen general {name}',tab(ind,df,None,None,entities))]
    for key,label in [('sexo','sexo'),('estrato','estrato')]:
        col=demo.get(key)
        if col:
            for v in pd.Series(df[col].fillna('Sin dato')).drop_duplicates(): sec.append((f'Detalle por {label} - {v}',tab(ind,df,col,v,entities)))
    return sec

def dept_sections(ind,df,demo,entities):
    col=demo.get('departamento')
    return [] if not col else [(str(v),tab(ind,df,col,v,entities)) for v in pd.Series(df[col].fillna('Sin dato')).drop_duplicates()]

def demo_sections(df,demo):
    out=[('Base',pd.DataFrame({'Indicador':['Total encuestas procesadas'],'Valor':[len(df)]}))]
    for k,l in [('sexo','Sexo'),('edad','Edad'),('departamento','Departamento'),('estrato','Estrato'),('ingreso','Ingreso')]:
        col=demo.get(k)
        if col:
            t=df[col].fillna('Sin dato').value_counts().rename_axis(l).reset_index(name='Cantidad'); t['%']=t['Cantidad']/len(df)
            out.append((f'Demográfico - {l}',t))
    return out

def style(ws):
    thin=Side(style='thin',color='BFBFBF'); border=Border(left=thin,right=thin,top=thin,bottom=thin)
    for row in ws.iter_rows():
        for cell in row: cell.border=border; cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    for c in range(1,ws.max_column+1): ws.column_dimensions[get_column_letter(c)].width=18

def write_sections(wb,name,secs):
    ws=wb.create_sheet(safe(name)); r=1; tf=PatternFill('solid',fgColor='C6E0B4'); hf=PatternFill('solid',fgColor='4472C4')
    for title,df in secs:
        ws.cell(r,1).value=title; ws.cell(r,1).fill=tf; ws.cell(r,1).font=Font(bold=True); r+=1
        for c,col in enumerate(df.columns,1): ws.cell(r,c).value=col; ws.cell(r,c).fill=hf; ws.cell(r,c).font=Font(bold=True,color='FFFFFF')
        r+=1
        for rec in df.itertuples(index=False):
            for c,val in enumerate(rec,1):
                ws.cell(r,c).value=val
                if df.columns[c-1].endswith('%'): ws.cell(r,c).number_format='0%'
            r+=1
        r+=2
    style(ws)

def write_main(wb,name,out,raw,entities):
    ws=wb.create_sheet(safe(name)); cols=list(out.columns)
    for c,col in enumerate(cols,1): ws.cell(1,c).value=str(col)
    for r,row in enumerate(out.itertuples(index=False),2):
        for c,val in enumerate(row,1): ws.cell(r,c).value=val
    style(ws)

def make_excel(df,demo,cfg1,cfg2,entities,aliases,norms,expected):
    out1,ind1,raw1=build(df,demo,cfg1,entities,aliases,norms,expected); out2,ind2,raw2=build(df,demo,cfg2,entities,aliases,norms,expected)
    wb=Workbook(); wb.remove(wb.active)
    write_main(wb,cfg1['name'],out1,raw1,entities); write_main(wb,cfg2['name'],out2,raw2,entities)
    write_sections(wb,'Resumen Demográficos',demo_sections(df,demo)); write_sections(wb,f"Resumen {cfg1['name']}",sections(ind1,df,demo,cfg1['name'],entities)); write_sections(wb,f"Resumen {cfg2['name']}",sections(ind2,df,demo,cfg2['name'],entities))
    write_sections(wb,f"Deptos {cfg1['name']}",dept_sections(ind1,df,demo,entities)); write_sections(wb,f"Deptos {cfg2['name']}",dept_sections(ind2,df,demo,entities))
    bio=io.BytesIO(); wb.save(bio); return bio.getvalue()

st.set_page_config(page_title='Awareness flexible',layout='wide'); st.title('Generador flexible de Awareness / Recordación')
mode=st.sidebar.radio('Tipo de estudio',['Bancos','Personalizado'])
upload=st.sidebar.file_uploader('Archivo base .xlsx',type='xlsx'); rows=st.sidebar.number_input('Filas a evaluar (0=todas)',0,step=50); expected=st.sidebar.number_input('Columnas crudas esperadas (0=no validar)',0,value=17)
ents_default=BANKS if mode=='Bancos' else ['Entidad 1','Entidad 2','Entidad 3']; aliases_default=BANK_ALIASES if mode=='Bancos' else {'Entidad 1':['entidad 1'],'Entidad 2':['entidad 2'],'Entidad 3':['entidad 3']}; norms_default=BANK_NORMALIZATIONS if mode=='Bancos' else []
with st.expander('Entidades, alias y normalizaciones',expanded=mode=='Personalizado'):
    ents_txt=st.text_area('Entidades, una por línea','\n'.join(ents_default)); aliases_txt=st.text_area('Alias/condiciones JSON',json.dumps(aliases_default,ensure_ascii=False,indent=2),height=260); norms_txt=st.text_area('Normalizaciones JSON',json.dumps(norms_default,ensure_ascii=False,indent=2),height=150)
entities=[x.strip() for x in ents_txt.splitlines() if x.strip()]; aliases=json.loads(aliases_txt); norms=json.loads(norms_txt)
with st.expander('Preguntas y demográficos'):
    c1,c2,c3=st.columns(3)
    with c1:
        sexo=st.text_input('Sexo','F1 '); edad=st.text_input('Edad','F2 '); depto=st.text_input('Departamento/Región/Zona','F3 '); estrato=st.text_input('Estrato/NSE/Segmento','F4 '); ingreso=st.text_input('Ingreso/Otro','F5 ')
    with c2:
        a1n=st.text_input('Nombre análisis 1','AWA PUB' if mode=='Bancos' else 'Awareness 1'); a1t=st.text_input('TOM análisis 1','P2 '); a1e=st.text_area('Espontáneo análisis 1','P2_1_1\nP2_1_2\nP2_1_3'); a1a=st.text_area('Ayudado análisis 1','\n'.join([f'P4_{i} ' for i in range(1,11)]))
    with c3:
        a2n=st.text_input('Nombre análisis 2','AWA Marca' if mode=='Bancos' else 'Awareness 2'); a2t=st.text_input('TOM análisis 2','P1 '); a2e=st.text_area('Espontáneo análisis 2','P1_1_1\nP1_1_2\nP1_1_3'); a2a=st.text_area('Ayudado análisis 2','\n'.join([f'P3_{i} ' for i in range(1,11)]))
if not upload: st.info('Carga un archivo para iniciar.'); st.stop()
df=pd.read_excel(upload,engine='openpyxl');
if rows: df=df.head(int(rows)).copy()
demo={'sexo':find_col(df,sexo),'edad':find_col(df,edad),'departamento':find_col(df,depto),'estrato':find_col(df,estrato),'ingreso':find_col(df,ingreso)}
cfg1={'name':a1n,'tom':a1t,'esp':a1e,'ayu':a1a}; cfg2={'name':a2n,'tom':a2t,'esp':a2e,'ayu':a2a}
st.subheader('Validación'); st.write('Demográficos detectados:',demo); st.write('Columnas TOM detectadas:',find_col(df,a1t),find_col(df,a2t))
if st.button('Generar Excel',type='primary'):
    try:
        data=make_excel(df,demo,cfg1,cfg2,entities,aliases,norms,int(expected)); st.success('Archivo generado')
        st.download_button('Descargar resultado',data,file_name='Resultado_Awareness_Flexible.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e: st.exception(e)
