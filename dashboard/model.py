"""Modelo unitário: item contratual 1:N condições de custo, sem alocar quantidades."""
from pathlib import Path
import hashlib
import math
import unicodedata
import pandas as pd
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT.parent / 'Planilha de Quantidades.xlsx'
KEY = ['codigo_composicao']
NAN = float('nan')

def normalized(value):
    return ' '.join(unicodedata.normalize('NFKD', str(value)).encode('ascii', 'ignore').decode().upper().split())

def identifier(value):
    """Não remove pontuação ou zeros significativos de identificadores textuais."""
    if pd.isna(value) or not str(value).strip():
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return str(int(value)) if float(value).is_integer() else str(value)
    return str(value).strip()

def divide(a, b):
    return a / b if pd.notna(a) and pd.notna(b) and b != 0 else NAN

def complete_sum(values):
    return values.sum() if len(values) and values.notna().all() else NAN

def sheet_to_frame(book, sheet_name):
    """Leitura leve das abas de dados da nova origem, sem carregar as abas extensas."""
    if sheet_name not in book.sheetnames:
        raise ValueError(f'Aba {sheet_name} não encontrada.')
    sheet = book[sheet_name]
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        raise ValueError(f'Aba {sheet_name} está vazia.')
    columns = [normalized(value) for value in rows[0]]
    if len(set(columns)) != len(columns):
        raise ValueError(f'Cabeçalhos duplicados na aba {sheet_name}.')
    frame = pd.DataFrame(rows[1:], columns=columns)
    frame['linha_excel'] = range(2, len(rows) + 1)
    frame['aba'] = sheet_name
    return frame.dropna(subset=columns, how='all')

def read_source(path=SOURCE):
    path = Path(path)
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        cost = sheet_to_frame(book, 'Tabela_Custo')
        revenue = sheet_to_frame(book, 'Tabela_Receita')
    finally:
        book.close()
    common = {
        'No DO PRECO': 'numero_preco', 'NO DO PRECO': 'numero_preco', 'N DO PRECO': 'numero_preco',
        'CODIGO COMPOSICAO': 'codigo_composicao'
    }
    cost = cost.rename(columns={**common, 'CPU':'cpu','CODINSUMO':'codigo','INSUMO':'insumo',
        'CLASSIFICACAO':'classe','LOCAL':'local','CUSTO UNIT':'preco_b6',
        'CUSTO UNIT AJUSTADA':'preco_b8','INCIDENCIA':'inc_b6','INCIDENCIA AJUSTADA':'inc_b8',
        'INCIDENCIA REVISAO':'inc_revisao','CUSTO INSUMO':'salvo_b6','CUSTO INSUMO AJUSTADO':'salvo_b8',
        'DELTA CUSTO INSUMO':'salvo_delta'})
    revenue = revenue.rename(columns={**common,'CPU':'descricao','PRECO UNITARIO':'receita_unitaria',
        'RECEITA PREVISTA':'receita_prevista','QUANTIDADE PREVISTA':'quantidade_contratada',
        'QUANTIDADE CONTRATADA':'quantidade_contratada','QUANTIDADE EXECUTADA':'quantidade_executada',
        'SALDO QUANTIDADE':'saldo','UNIDADE':'unidade'})
    for frame, required in [(cost, KEY+['cpu','codigo','insumo','classe','local','preco_b6','preco_b8','inc_b6','inc_b8','salvo_b6','salvo_b8']),
                            (revenue, KEY+['descricao','receita_unitaria','quantidade_contratada','saldo'])]:
        missing = set(required)-set(frame)
        if missing:
            raise ValueError(f'Colunas obrigatórias ausentes: {sorted(missing)}')
        if frame.columns.duplicated().any():
            raise ValueError('Mais de uma coluna representa o mesmo conceito; padronize os cabeçalhos.')
    for field in ['quantidade_executada','unidade','receita_prevista']:
        if field not in revenue: revenue[field] = None
    for field in ['inc_revisao','salvo_delta']:
        if field not in cost: cost[field] = None
    # Campos de compatibilidade para a apresentação existente. A ligação de
    # dados continua usando exclusivamente Código Composição.
    cost['item'] = cost['codigo_composicao']
    revenue['item'] = revenue['codigo_composicao']
    return build_model(cost, revenue, digest)

def build_model(cost, revenue, digest=''):
    cost, revenue = cost.copy(), revenue.copy()
    issues = []
    def record(frame, mask, level, problem, field):
        for _, row in frame.loc[mask].iterrows():
            issues.append({'Nivel':level,'Ocorrência':problem,'Campo':field,
                'Nº DO PREÇO':row.get('numero_preco'),'Item':row.get('item'),
                'CPU / descrição':row.get('cpu',row.get('descricao')),
                'Aba':row.get('aba'),'Linha Excel':row.get('linha_excel')})
    for frame in [cost, revenue]:
        for key in KEY:
            frame[key] = frame[key].map(identifier)
            record(frame,frame[key].isna(),'Problema','Chave incompleta',key)
    cost_numeric = ['preco_b6','preco_b8','inc_b6','inc_b8','inc_revisao','salvo_b6','salvo_b8','salvo_delta']
    revenue_numeric = ['receita_unitaria','receita_prevista','quantidade_contratada','quantidade_executada','saldo']
    for frame, fields in [(cost,cost_numeric),(revenue,revenue_numeric)]:
        for field in fields:
            original=frame[field]
            frame[field]=pd.to_numeric(original,errors='coerce').replace([float('inf'),-float('inf')],NAN)
            record(frame,original.notna() & frame[field].isna(),'Problema','Valor não numérico ou infinito',field)
    for field in ['cpu','codigo','insumo','classe','local']:
        empty=cost[field].isna() | cost[field].astype(str).str.strip().eq('')
        record(cost,empty,'Problema' if field!='local' else 'Atenção','Campo relevante vazio',field)
        cost.loc[empty,field]='Não informado'
    for field in ['descricao','receita_unitaria','quantidade_contratada','saldo']:
        record(revenue,revenue[field].isna(),'Problema','Campo relevante vazio',field)
    revenue['quantidade_executada']=revenue.quantidade_contratada-revenue.saldo
    revenue['executado_pct']=[divide(e,q)*100 for e,q in zip(revenue.quantidade_executada,revenue.quantidade_contratada)]
    for field in ['unidade']:
        record(revenue,revenue[field].isna(),'Opcional','Não informado na base',field)
    record(cost,cost.inc_revisao.isna(),'Opcional','Incidência sem revisão','inc_revisao')
    for base in ['b6','b8']:
        product=cost[f'preco_{base}']*cost[f'inc_{base}']
        saved=cost[f'salvo_{base}']
        cost[base]=product.where(product.notna(),saved)
        record(cost,product.isna() & saved.notna(),'Atenção','Fatores incompletos; custo salvo preservado',base)
        record(cost,cost[base].isna(),'Problema','Custo indisponível; total da CPU bloqueado',base)
        record(cost,product.notna() & saved.notna() & (product-saved).abs().gt(.01),'Atenção','Produto difere do custo salvo; usando preço × incidência',base)
    cost['delta']=cost.b8-cost.b6
    record(cost,cost.salvo_delta.notna() & (cost.delta-cost.salvo_delta).abs().gt(.01),'Atenção','Delta difere do valor salvo','delta')
    record(cost,cost.duplicated(KEY+['cpu','codigo'],keep=False),'Atenção','Código de insumo repetido na mesma CPU e chave','codigo')
    cost['diferenca_inc']=cost.inc_b8-cost.inc_b6
    cost['diferenca_inc_pct']=[divide(a,b)*100 for a,b in zip(cost.diferenca_inc,cost.inc_b6)]
    cost['diferenca_revisao']=cost.inc_revisao-cost.inc_b6
    cost['diferenca_revisao_pct']=[divide(a,b)*100 for a,b in zip(cost.diferenca_revisao,cost.inc_b6)]

    valid_key=revenue[KEY].notna().all(axis=1)
    duplicate=revenue.duplicated(KEY,keep=False) & valid_key
    revenue['chave_unica']=valid_key & ~duplicate
    record(revenue,duplicate,'Problema','Chave composta duplicada; associação bloqueada','Nº DO PREÇO + Item')
    for field in ['quantidade_contratada','quantidade_executada','saldo','receita_unitaria']:
        record(revenue,revenue[field].lt(0),'Problema','Valor negativo',field)
    record(revenue,revenue.saldo.gt(revenue.quantidade_contratada),'Atenção','Saldo maior que quantidade contratada','saldo')
    discrepancy=(revenue.quantidade_executada+revenue.saldo-revenue.quantidade_contratada).abs().gt(.01)
    record(revenue,discrepancy,'Atenção','Executado + saldo difere do contratado','quantidades')
    price_difference=(revenue.receita_unitaria*revenue.quantidade_contratada-revenue.receita_prevista).abs().gt(.01)
    record(revenue,price_difference & revenue.receita_prevista.notna(),'Atenção','Preço × quantidade difere da receita prevista','receita_unitaria')

    # Quantidades permanecem exclusivamente no catálogo de itens. Não entram no join.
    summary=cost.groupby(KEY+['cpu'],dropna=False,sort=False).agg(
        b6=('b6',complete_sum),b8=('b8',complete_sum),
        local=('local',lambda s:'; '.join(dict.fromkeys(s.astype(str)))),
        linhas=('linha_excel','size')).reset_index()
    summary['item'] = summary['codigo_composicao']
    summary['delta']=summary.b8-summary.b6
    summary['delta_pct']=[(divide(b,a)-1)*100 for a,b in zip(summary.b6,summary.b8)]
    unique=revenue.loc[revenue.chave_unica,KEY+['numero_preco','descricao','unidade','receita_unitaria']].copy()
    summary=summary.merge(unique,on=KEY,how='left',validate='many_to_one',indicator=True)
    summary['associacao']=summary['_merge'].map({'both':'Chave composta correspondente','left_only':'Sem correspondência única','right_only':'Sem custo'}).astype(str)
    summary=summary.drop(columns='_merge')
    summary['fcd_b6']=[divide(p,c) if pd.notna(c) and c>0 and pd.notna(p) and p>=0 else NAN for p,c in zip(summary.receita_unitaria,summary.b6)]
    summary['fcd_b8']=[divide(p,c) if pd.notna(c) and c>0 and pd.notna(p) and p>=0 else NAN for p,c in zip(summary.receita_unitaria,summary.b8)]
    existing=set(map(tuple,unique[KEY].values.tolist()))
    record(cost,~cost[KEY].apply(tuple,axis=1).isin(existing),'Problema','CPU sem correspondência única na Receita','Código Composição')
    cost_keys=set(map(tuple,cost.loc[cost[KEY].notna().all(axis=1),KEY].values.tolist()))
    record(revenue,~revenue[KEY].apply(tuple,axis=1).isin(cost_keys),'Atenção','Item de receita sem CPU de custo','Código Composição')
    counts=summary.loc[summary[KEY].notna().all(axis=1)].groupby(KEY).size().to_dict()
    revenue['cpus_relacionadas']=[counts.get(getattr(row, 'codigo_composicao'), 0) for row in revenue.itertuples()]
    quality=pd.DataFrame(issues,columns=['Nivel','Ocorrência','Campo','Nº DO PREÇO','Item','CPU / descrição','Aba','Linha Excel'])
    return {'cost':cost,'revenue':revenue,'cpus':summary,'quality':quality,'sha256':digest}

def decomposition(detail):
    groups=detail.groupby('classe',sort=False)[['b6','b8','delta']].agg(complete_sum).reset_index()
    totals={field:complete_sum(detail[field]) for field in ['b6','b8','delta']}
    for name,field in [('participacao_b6','b6'),('participacao_b8','b8'),('participacao_delta','delta')]:
        groups[name]=[divide(v,totals[field])*100 for v in groups[field]]
    return groups

def strategy_change(detail):
    labor=detail[detail.classe.map(normalized).eq('MAO DE OBRA')]
    third=detail[detail.classe.map(normalized).eq('SERVICOS DE TERCEIROS')]
    return not labor.empty and not third.empty and complete_sum(labor.inc_b8)<complete_sum(labor.inc_b6) and complete_sum(third.b8)>complete_sum(third.b6)

def resource_costs(detail):
    """Agrega pelo nome exibido, preservando códigos apenas como informação secundária."""
    resources=detail.groupby('insumo',sort=False).agg(
        b6=('b6',complete_sum),b8=('b8',complete_sum),delta=('delta',complete_sum),
        codigo=('codigo',lambda s:', '.join(dict.fromkeys(s.astype(str))))).reset_index()
    for base in ['b6','b8']:
        total=complete_sum(detail[base])
        resources['participacao_'+base]=[divide(v,total)*100 for v in resources[base]]
    return resources
