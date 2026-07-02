# -*- coding: utf-8 -*-
"""Generador Awareness Final Limpio v6

Hojas finales:
- AWA PUB
- AWA Marca
- Resumen Demográficos
- Resumen AWA PUB
- Resumen AWA Marca
- Deptos AWA PUB
- Deptos AWA Marca

En Resumen AWA PUB / Marca:
- Resumen general
- Tabla independiente por cada sexo
- Tabla independiente por cada estrato

En Deptos AWA PUB / Marca:
- Tabla independiente por cada departamento, con TOM, Espontaneo, Ayudado y AWA.
"""

import os, re, unicodedata
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BRANDS = ['Banco Agrario','Bancolombia','Davivienda','Banco de Bogotá','BBVA','Banco Caja Social','Banco Popular','Bancamía','Banco de Occidente','Banco Mundo Mujer']
EXACT_HEADERS = {'AWA PUB': [[None, None, None, '=E512', None, None, None, 'Ayudado', None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None], ['PREGUNTAS DEMOGRAFICAS', None, None, 'TOM PUB', 'Espontaneo PUB', None, None, 'Banco Agrario', 'Bancolombia', 'Davivienda', 'Banco de Bogotá', 'BBVA', 'Banco Caja Social', 'Banco Popular', 'Bancamía', 'Banco de Occidente', 'Banco Mundo Mujer', '=+H2', None, None, '=+I2', None, None, '=+J2', None, None, '=+K2', None, None, '=+L2', None, None, '=+M2', None, None, '=+N2', None, None, '=+O2', None, None, '=+P2', None, None, '=+Q2', None, None], ['F2 ¿En qué rango de edad te encuentras? [OCULTA]', 'F3 ¿En qué departamento vives?', 'F4 ¿Según los recibos de servicios públicos de tu hogar, ¿a qué estrato perteneces?', 'P2 ¿De qué bancos has visto, escuchado o leído publicidad en el último mes?', 'P2_1_1 Además de los que ya mencionaste, ¿qué otros bancos recuerdas haber visto o escuchado en publicidad durante el último mes?', 'P2_1_2 Además de los que ya mencionaste, ¿qué otros bancos recuerdas haber visto o escuchado en publicidad durante el último mes?', 'P2_1_3 Además de los que ya mencionaste, ¿qué otros bancos recuerdas haber visto o escuchado en publicidad durante el último mes?', 'P4_1 Del siguiente listado de marcas de bancos, ¿de cuáles has escuchado, leído o visto publicidad en el último mes? (Recuerda que puedes marcar más de una opción) Banco Agrario', 'P4_2 Del siguiente listado de marcas de bancos, ¿de cuáles has escuchado, leído o visto publicidad en el último mes? (Recuerda que puedes marcar más de una opción) Bancolombia', 'P4_3 Del siguiente listado de marcas de bancos, ¿de cuáles has escuchado, leído o visto publicidad en el último mes? (Recuerda que puedes marcar más de una opción) Davivienda', 'P4_4 Del siguiente listado de marcas de bancos, ¿de cuáles has escuchado, leído o visto publicidad en el último mes? (Recuerda que puedes marcar más de una opción) Banco de Bogotá', 'P4_5 Del siguiente listado de marcas de bancos, ¿de cuáles has escuchado, leído o visto publicidad en el último mes? (Recuerda que puedes marcar más de una opción) BBVA', 'P4_6 Del siguiente listado de marcas de bancos, ¿de cuáles has escuchado, leído o visto publicidad en el último mes? (Recuerda que puedes marcar más de una opción) Banco Caja Social', 'P4_7 Del siguiente listado de marcas de bancos, ¿de cuáles has escuchado, leído o visto publicidad en el último mes? (Recuerda que puedes marcar más de una opción) Banco Popular', 'P4_8 Del siguiente listado de marcas de bancos, ¿de cuáles has escuchado, leído o visto publicidad en el último mes? (Recuerda que puedes marcar más de una opción) Bancamía', 'P4_9 Del siguiente listado de marcas de bancos, ¿de cuáles has escuchado, leído o visto publicidad en el último mes? (Recuerda que puedes marcar más de una opción) Banco de Occidente', 'P4_10 Del siguiente listado de marcas de bancos, ¿de cuáles has escuchado, leído o visto publicidad en el último mes? (Recuerda que puedes marcar más de una opción) Banco Mundo Mujer', 'TOM', 'Espontaneo', 'Ayudado', 'TOM', 'Espontaneo', 'Ayudado', 'TOM', 'Espontaneo', 'Ayudado', 'TOM', 'Espontaneo', 'Ayudado', 'TOM', 'Espontaneo', 'Ayudado', 'TOM', 'Espontaneo', 'Ayudado', 'TOM', 'Espontaneo', 'Ayudado', 'TOM', 'Espontaneo', 'Ayudado', 'TOM', 'Espontaneo', 'Ayudado', 'TOM', 'Espontaneo', 'Ayudado']], 'AWA Marca': [[None, None, None, '=E517', None, None, None, 'Ayudado', None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None], ['PREGUNTAS DEMOGRAFICAS', None, None, 'TOM', 'Espontaneo', None, None, 'Banco Agrario', 'Bancolombia', 'Davivienda', 'Banco de Bogotá', 'BBVA', 'Banco Caja Social', 'Banco Popular', 'Bancamía', 'Banco de Occidente', 'Banco Mundo Mujer', '=+H2', None, None, '=+I2', None, None, '=+J2', None, None, '=+K2', None, None, '=+L2', None, None, '=+M2', None, None, '=+N2', None, None, '=+O2', None, None, '=+P2', None, None, '=+Q2', None, None], ['F2 ¿En qué rango de edad te encuentras? [OCULTA]', 'F3 ¿En qué departamento vives?', 'F4 ¿Según los recibos de servicios públicos de tu hogar, ¿a qué estrato perteneces?', 'P1 ¿Cuál es el primer banco que recuerdas?', 'P1_1_1 ¿Qué otros bancos conoces?', 'P1_1_2 ¿Qué otros bancos conoces?', 'P1_1_3 ¿Qué otros bancos conoces?', 'P3_1 Del siguiente listado de bancos, ¿cuáles conoces?\n(Recuerda que puedes marcar más de una opción) Banco Agrario', 'P3_2 Del siguiente listado de bancos, ¿cuáles conoces?\n(Recuerda que puedes marcar más de una opción) Bancolombia', 'P3_3 Del siguiente listado de bancos, ¿cuáles conoces?\n(Recuerda que puedes marcar más de una opción) Davivienda', 'P3_4 Del siguiente listado de bancos, ¿cuáles conoces?\n(Recuerda que puedes marcar más de una opción) Banco de Bogotá', 'P3_5 Del siguiente listado de bancos, ¿cuáles conoces?\n(Recuerda que puedes marcar más de una opción) BBVA', 'P3_6 Del siguiente listado de bancos, ¿cuáles conoces?\n(Recuerda que puedes marcar más de una opción) Banco Caja Social', 'P3_7 Del siguiente listado de bancos, ¿cuáles conoces?\n(Recuerda que puedes marcar más de una opción) Banco Popular', 'P3_8 Del siguiente listado de bancos, ¿cuáles conoces?\n(Recuerda que puedes marcar más de una opción) Bancamía', 'P3_9 Del siguiente listado de bancos, ¿cuáles conoces?\n(Recuerda que puedes marcar más de una opción) Banco de Occidente', 'P3_10 Del siguiente listado de bancos, ¿cuáles conoces?\n(Recuerda que puedes marcar más de una opción) Banco Mundo Mujer', 'TOM', 'Espontaneo', 'Ayudado', 'TOM', 'Espontaneo', 'Ayudado', 'TOM', 'Espontaneo', 'Ayudado', 'TOM', 'Espontaneo', 'Ayudado', 'TOM', 'Espontaneo', 'Ayudado', 'TOM', 'Espontaneo', 'Ayudado', 'TOM', 'Espontaneo', 'Ayudado', 'TOM', 'Espontaneo', 'Ayudado', 'TOM', 'Espontaneo', 'Ayudado', 'TOM', 'Espontaneo', 'Ayudado']]}
NEG = {'', '0', 'no', 'nan', 'none', 'false', 'ninguno', 'ninguna', 'no se', 'no sé', 'no recuerdo', 'ningun otro', 'ningún otro', 'no aplica'}
ALIASES = {
    'Banco Agrario':['banco agrario','banco agrario de colombia','agrario'],
    'Bancolombia':['bancolombia'],
    'Davivienda':['davivienda'],  # Davibank/DaviBank NO se cuenta como Davivienda
    'Banco de Bogotá':['banco de bogota','banco de bogotá','banco bogota','banco bogotá'],
    'BBVA':['bbva','bbvva','bvva','bbwa','bva','bvvwa'],
    'Banco Caja Social':['banco caja social','caja social','cajas social','caja sosial'],
    'Banco Popular':['banco popular','popular'],
    'Bancamía':['bancamia','bancamía','banca mia','banca mía'],
    'Banco de Occidente':['banco de occidente','occidente'],
    'Banco Mundo Mujer':['banco mundo mujer','mundo mujer','banco de la mujer','fundacion de la mujer','fundación de la mujer','de la mujer']
}

def norm(x):
    if pd.isna(x): return ''
    x=str(x).strip().lower()
    x=unicodedata.normalize('NFKD',x).encode('ascii','ignore').decode('ascii')
    x=re.sub(r'[^a-z0-9\s]+',' ',x)
    return re.sub(r'\s+',' ',x).strip()

def clean_text(x):
    if pd.isna(x): return x
    s=str(x)
    s=re.sub(r'(?i)Banco\s+de\s+Bogota','Banco de Bogotá',s)
    s=re.sub(r'(?i)(?<!Banco\s)Caja\s+Social','Banco Caja Social',s)
    s=re.sub(r'(?i)Bancamia','Bancamía',s)
    return s

def find_col(df,prefix):
    p=norm(prefix)
    for c in df.columns:
        if norm(c).startswith(p): return c
    return None

def find_cols(df,prefixes):
    return [c for p in prefixes for c in [find_col(df,p)] if c is not None]

def contains_brand(x,brand):
    t=norm(x)
    if t in NEG: return False
    for a in ALIASES.get(brand,[brand]):
        aa=norm(a)
        if aa and re.search(r'(^|\s)'+re.escape(aa)+r'(\s|$)', t): return True
    return False

def aided_brand(col):
    t=norm(str(col).replace('\n',' '))
    for b in BRANDS:
        if norm(b) in t: return b
    if 'caja social' in t: return 'Banco Caja Social'
    if 'bancamia' in t: return 'Bancamía'
    if 'banco de bogota' in t: return 'Banco de Bogotá'
    return None

def detect(df):
    d={}
    for f in ['F1','F2','F3','F4','F5']:
        d[f]=find_col(df,f+' ')
    d['P1']=find_col(df,'P1 '); d['P2']=find_col(df,'P2 ')
    d['P1e']=find_cols(df,['P1_1_1','P1_1_2','P1_1_3'])
    d['P2e']=find_cols(df,['P2_1_1','P2_1_2','P2_1_3'])
    d['P3a']=find_cols(df,[f'P3_{i} ' for i in range(1,11)])
    d['P4a']=find_cols(df,[f'P4_{i} ' for i in range(1,11)])
    return d

def build_awareness(df,d,kind):
    if kind=='PUB':
        raw=[d['F2'],d['F3'],d['F4'],d['P2']]+d['P2e']+d['P4a']; tom=d['P2']; esps=d['P2e']; ayud=d['P4a']
    else:
        raw=[d['F2'],d['F3'],d['F4'],d['P1']]+d['P1e']+d['P3a']; tom=d['P1']; esps=d['P1e']; ayud=d['P3a']
    raw=[c for c in raw if c is not None]
    if len(raw)!=17:
        raise ValueError(f'{kind} debe tener 17 columnas crudas; detectadas {len(raw)}: {raw}')
    rawdf=df[raw].applymap(clean_text)
    out=rawdf.copy(); ind={}
    ayumap={b:None for b in BRANDS}
    for c in ayud:
        b=aided_brand(c)
        if b in ayumap: ayumap[b]=c
    for b in BRANDS:
        T=rawdf[tom].apply(lambda x: contains_brand(x,b))
        E=pd.Series(False,index=df.index)
        for c in esps:
            E = E | rawdf[c].apply(lambda x: contains_brand(x,b))
        A=rawdf[ayumap[b]].apply(lambda x: contains_brand(x,b)) if ayumap.get(b) else pd.Series(False,index=df.index)
        ind[(b,'TOM')]=T.astype(int)
        ind[(b,'Espontaneo')]=((~T)&E).astype(int)
        ind[(b,'Ayudado')]=((~T)&(~E)&A).astype(int)
        out[(b,'TOM')]=ind[(b,'TOM')]
        out[(b,'Espontaneo')]=ind[(b,'Espontaneo')]
        out[(b,'Ayudado')]=ind[(b,'Ayudado')]
    return out, pd.DataFrame(ind)

def style_cells(ws):
    thin=Side(style='thin', color='BFBFBF')
    border=Border(left=thin,right=thin,top=thin,bottom=thin)
    for row in ws.iter_rows():
        for cell in row:
            cell.border=border
            cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    for c in range(1,ws.max_column+1):
        ws.column_dimensions[get_column_letter(c)].width=18

def style_main(ws):
    style_cells(ws)
    fill=PatternFill('solid',fgColor='D9EAF7')
    for r in [1,2,3]:
        for c in range(1,ws.max_column+1):
            ws.cell(r,c).fill=fill
            ws.cell(r,c).font=Font(bold=True)
    ws.freeze_panes='A4'

def write_main(wb,name,dfout):
    ws=wb.create_sheet(name)
    for r in range(1,4):
        for c in range(1,48):
            ws.cell(r,c).value=EXACT_HEADERS[name][r-1][c-1]
    for rr,row in enumerate(dfout.itertuples(index=False), start=4):
        for cc,val in enumerate(row,start=1):
            ws.cell(rr,cc).value=val
    style_main(ws)

def table_by_category(ind, df, d, key, category):
    col=d[key]
    mask=(df[col].fillna('Sin dato')==category)
    base=int(mask.sum())
    rows=[]
    for b in BRANDS:
        tom=int(ind[(b,'TOM')][mask].sum())
        esp=int(ind[(b,'Espontaneo')][mask].sum())
        ayu=int(ind[(b,'Ayudado')][mask].sum())
        awa=tom+esp+ayu
        rows.append({'Marca':b,'TOM':tom,'TOM %':tom/base if base else 0,'Espontaneo':esp,'Espontaneo %':esp/base if base else 0,'Ayudado':ayu,'Ayudado %':ayu/base if base else 0,'AWA':awa,'AWA %':awa/base if base else 0})
    return pd.DataFrame(rows)

def summary(ind,n):
    rows=[]
    for b in BRANDS:
        tom=int(ind[(b,'TOM')].sum()); esp=int(ind[(b,'Espontaneo')].sum()); ayu=int(ind[(b,'Ayudado')].sum()); awa=tom+esp+ayu
        rows.append({'Marca':b,'TOM':tom,'TOM %':tom/n,'Espontaneo':esp,'Espontaneo %':esp/n,'Ayudado':ayu,'Ayudado %':ayu/n,'AWA':awa,'AWA %':awa/n})
    return pd.DataFrame(rows)

def sex_dept_table(df,d):
    t=pd.crosstab(df[d['F3']].fillna('Sin dato'), df[d['F1']].fillna('Sin dato'))
    t['Total']=t.sum(axis=1)
    return t.reset_index().rename(columns={d['F3']:'Departamento'})

def demo_sections(df,d):
    n=len(df)
    sections=[('Base', pd.DataFrame({'Indicador':['Total encuestas procesadas'],'Valor':[n]}))]
    for name,key in [('Sexo','F1'),('Edad','F2'),('Departamento','F3'),('Estrato','F4'),('Ingreso','F5')]:
        col=d.get(key)
        if col:
            t=df[col].fillna('Sin dato').value_counts().rename_axis(name).reset_index(name='Cantidad')
            t['%']=t['Cantidad']/n
            sections.append(('Demográfico - '+name,t))
    sections.append(('Participación Sexo por Departamento', sex_dept_table(df,d)))
    return sections

def resumen_sections(ind, df, d, titulo, n):
    sections=[(titulo, summary(ind,n))]
    if d.get('F1'):
        for sexo in list(pd.Series(df[d['F1']].fillna('Sin dato')).drop_duplicates()):
            sections.append((f'Detalle por sexo - {sexo}', table_by_category(ind,df,d,'F1',sexo)))
    if d.get('F4'):
        for estrato in list(pd.Series(df[d['F4']].fillna('Sin dato')).drop_duplicates()):
            sections.append((f'Detalle por estrato - {estrato}', table_by_category(ind,df,d,'F4',estrato)))
    return sections

def dept_sections(ind, df, d):
    sections=[]
    for dep in list(pd.Series(df[d['F3']].fillna('Sin dato')).drop_duplicates()):
        sections.append((str(dep), table_by_category(ind,df,d,'F3',dep)))
    return sections

def write_sections(wb,name,sections,percent_cols=None):
    ws=wb.create_sheet(name)
    r=1
    title_fill=PatternFill('solid',fgColor='C6E0B4')
    head_fill=PatternFill('solid',fgColor='4472C4')
    thin=Side(style='thin', color='BFBFBF')
    border=Border(left=thin,right=thin,top=thin,bottom=thin)
    percent_cols=percent_cols or []
    for title,df in sections:
        ws.cell(r,1).value=title
        ws.cell(r,1).fill=title_fill
        ws.cell(r,1).font=Font(bold=True)
        r+=1
        for c,col in enumerate(df.columns,start=1):
            ws.cell(r,c).value=col
            ws.cell(r,c).fill=head_fill
            ws.cell(r,c).font=Font(bold=True,color='FFFFFF')
        r+=1
        for rec in df.itertuples(index=False):
            for c,val in enumerate(rec,start=1):
                cell=ws.cell(r,c)
                cell.value=val
                cell.border=border
                if df.columns[c-1] in percent_cols:
                    cell.number_format='0%'
            r+=1
        r+=2
    style_cells(ws)

def build_workbook(input_file, output_file):
    df=pd.read_excel(input_file, engine='openpyxl')
    d=detect(df)
    n=len(df)
    pub,pubi=build_awareness(df,d,'PUB')
    marca,marcai=build_awareness(df,d,'MARCA')
    wb=Workbook(); wb.remove(wb.active)
    write_main(wb,'AWA PUB',pub)
    write_main(wb,'AWA Marca',marca)
    pct=['%','TOM %','Espontaneo %','Ayudado %','AWA %']
    write_sections(wb,'Resumen Demográficos',demo_sections(df,d),pct)
    write_sections(wb,'Resumen AWA PUB',resumen_sections(pubi,df,d,'Resumen general AWA PUB',n),pct)
    write_sections(wb,'Resumen AWA Marca',resumen_sections(marcai,df,d,'Resumen general AWA Marca',n),pct)
    write_sections(wb,'Deptos AWA PUB',dept_sections(pubi,df,d),pct)
    write_sections(wb,'Deptos AWA Marca',dept_sections(marcai,df,d),pct)
    wb.save(output_file)
    print('Listo:',output_file)

if __name__=='__main__':
    try:
        import tkinter as tk
        from tkinter import filedialog
        root=tk.Tk(); root.withdraw()
        input_path=filedialog.askopenfilename(title='Selecciona el Excel base',filetypes=[('Excel','*.xlsx *.xls')])
        if not input_path: raise SystemExit('No se seleccionó archivo')
    except Exception:
        input_path=input('Ruta del archivo Excel: ').strip().strip('"')
    build_workbook(input_path, os.path.join(os.path.dirname(input_path) or '.', 'Resultado_Awareness_Final_Limpio_v6.xlsx'))
