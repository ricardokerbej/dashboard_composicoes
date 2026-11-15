from datetime import datetime
import textwrap
import pandas as pd
import plotly.express as px
import streamlit as st
from model import SOURCE, KEY, read_source, decomposition, strategy_change, normalized

st.set_page_config(page_title='B8 · Análise de composições',page_icon='◈',layout='wide')
st.markdown('''<style>
.stApp {background:#f5f7fa;color:#183247}
[data-testid="stSidebar"] {background:#102e42;color:#f5f7fa}
[data-testid="stSidebar"] h2,[data-testid="stSidebar"] label p,[data-testid="stSidebar"] [data-testid="stCaptionContainer"] {color:#eaf0f5}
[data-testid="stSidebar"] button p {color:#183247}
[data-testid="stMetric"] {background:white;border:1px solid #dce5ec;border-radius:10px;padding:14px}
[data-testid="stMetricValue"] {font-size:clamp(18px,2vw,28px)}
[data-testid="stMetricLabel"] p {white-space:normal}
@media (min-width:641px) and (max-width:1100px) {
[data-testid="stHorizontalBlock"] {flex-wrap:wrap}
[data-testid="stHorizontalBlock"] > [data-testid="stColumn"] {min-width:45%;flex:1 1 45%}}
h1,h2,h3 {letter-spacing:-.02em}
.eyebrow {color:#217b85;letter-spacing:.13em;font-size:12px;font-weight:700}
</style>''',unsafe_allow_html=True)

@st.cache_data(show_spinner='Lendo itens e composições…')
def load(stamp, size):
    return read_source()

def fmt(value, decimals=2, prefix=''):
    if pd.isna(value): return 'Não informado'
    return prefix+f'{value:,.{decimals}f}'.replace(',','X').replace('.',',').replace('X','.')

def resource_costs(detail):
    """Agrega os custos por nome do insumo para os gráficos da CPU."""
    resources=detail.groupby('insumo',sort=False).agg(
        b6=('b6','sum'),b8=('b8','sum'),delta=('delta','sum'),
        codigo=('codigo',lambda values:', '.join(dict.fromkeys(values.astype(str))))).reset_index()
    for base in ['b6','b8']:
        total=detail[base].sum()
        resources['participacao_'+base]=resources[base].map(lambda value: value/total*100 if pd.notna(total) and total != 0 else float('nan'))
    return resources

LABELS={'numero_preco':'Nº DO PREÇO','item':'Item','descricao':'Descrição contratual','cpu':'CPU de custo',
        'local':'Local / anotação da fonte','unidade':'Unidade','receita_unitaria':'Receita R$/un',
        'quantidade_contratada':'Quantidade contratada','quantidade_executada':'Quantidade executada',
        'executado_pct':'Executado %',
        'saldo':'Saldo do item','cpus_relacionadas':'CPUs relacionadas','b6':'B6 R$/un','b8':'B8 preliminar R$/un',
        'delta':'Delta R$/un','delta_pct':'Delta %','fcd_b6':'F/CD B6','fcd_b8':'F/CD B8 preliminar',
        'classe':'Classificação','codigo':'Código insumo','insumo':'Insumo','linha_excel':'Linha Excel',
        'aba':'Aba de origem','inc_b6':'Incidência B6','inc_b8':'Incidência ajustada','inc_revisao':'Incidência revisão',
        'diferenca_inc':'Diferença ajustada − B6','diferenca_inc_pct':'Diferença ajustada %',
        'diferenca_revisao':'Diferença revisão − B6','diferenca_revisao_pct':'Diferença revisão %',
        'participacao_b6':'Participação custo B6 %','participacao_b8':'Participação custo B8 %',
        'participacao_delta':'Participação na variação %','associacao':'Correspondência',
        'preco_b6':'Preço insumo B6','preco_b8':'Preço insumo ajustado',
        'salvo_b6':'Custo B6 salvo','salvo_b8':'Custo ajustado salvo'}

def grid(frame, download=None):
    view=frame.rename(columns=LABELS)
    # Exibição pt-BR nas tabelas: ponto para milhar e vírgula para decimais.
    # A cópia usada no download permanece numérica para não prejudicar o CSV.
    export_view=view.copy()
    numeric_cols=view.select_dtypes('number').columns.tolist()
    for name in numeric_cols:
        view[name]=view[name].map(lambda value: '' if pd.isna(value) else fmt(value))
    configs={}
    for name in ['CPU de custo','Descrição contratual','Local / anotação da fonte','Insumo','Ocorrência','CPU / descrição']:
        if name in view: configs[name]=st.column_config.TextColumn(width='medium')
    st.dataframe(view,hide_index=True,width='stretch',column_config=configs)
    if download:
        st.download_button('Baixar tabela (CSV)',export_view.to_csv(index=False,sep=';',decimal=',').encode('utf-8-sig'),download+'.csv','text/csv',key=download)

def chart(frame,x,y,color=None,share=None):
    if frame.empty: return
    fig=px.bar(frame,x=x,y=y,color=color,orientation='h',barmode='group',labels=LABELS,
               text=share,hover_data={share:':.2f'} if share else None,
               color_discrete_sequence=['#217b85','#98afc1','#d5a441'],
               color_discrete_map={'Redução':'#217b85','Aumento':'#bc613d','Sem variação':'#98afc1'})
    names=frame[y].drop_duplicates().tolist()
    fig.update_yaxes(categoryorder='array',categoryarray=names[::-1],tickmode='array',tickvals=names,
                     ticktext=['<br>'.join(textwrap.wrap(str(name),32)) for name in names])
    if share: fig.update_traces(texttemplate='%{text:.1f}%',textposition='auto')
    fig.update_layout(template='plotly_white',height=max(300,min(1100,120+frame[y].nunique()*95)),
                      margin=dict(l=0,r=15,t=15,b=55),legend=dict(orientation='h',y=-.2),legend_title_text='',
                      font=dict(color='#183247'))
    st.plotly_chart(fig,width='stretch')

try:
    stat=SOURCE.stat()
    data=load(stat.st_mtime_ns,stat.st_size)
except (OSError,ValueError,KeyError) as exc:
    st.error(f'Falha na leitura de {SOURCE.name}: {exc}')
    st.stop()
c,r,s,q=(data[k] for k in ['cost','revenue','cpus','quality'])

with st.sidebar:
    st.markdown('## B8 / LINHA 2')
    st.caption('ANÁLISE DE COMPOSIÇÕES B6 × B8')
    page=st.radio('Página',['Visão Executiva','Análise Comparativa','Produtividade','Qualidade dos dados'])
    st.divider()
    contract=None
    if st.button('Atualizar base'):
        st.cache_data.clear()
        st.rerun()
    st.caption(f'Fonte: {SOURCE.name}')
    st.caption('Arquivo salvo em '+datetime.fromtimestamp(stat.st_mtime).strftime('%d/%m/%Y %H:%M'))

st.markdown('<div class="eyebrow">METRÔ LINHA 2 · REPLANEJAMENTO B8</div>',unsafe_allow_html=True)
st.title(page)

if page=='Qualidade dos dados':
    st.caption('Auditoria da base completa. Campos opcionais não são contabilizados como problemas.')
    cols=st.columns(3)
    for col,level in zip(cols,['Problema','Atenção','Opcional']): col.metric(level,int(q.Nivel.eq(level).sum()))
    levels=st.multiselect('Nível', ['Problema','Atenção','Opcional'],default=['Problema','Atenção'])
    grid(q[q.Nivel.isin(levels)],'qualidade')
    st.subheader('Chaves da Receita')
    grid(r[KEY+['descricao','chave_unica','cpus_relacionadas','linha_excel']],'chaves_receita')
    st.caption('A associação usa o Código Composição. Um código da Receita pode atender várias CPUs. Se o código aparecer mais de uma vez na Receita, a associação afetada fica indisponível.')
    with st.expander('Rastreabilidade da leitura'):
        st.code(data['sha256'],language=None)
        st.caption('SHA-256 da planilha original. Leitura somente de consulta; nenhuma alteração na fonte.')
    st.stop()

scope=s.copy()
item_row=None
if contract is not None:
    item_row=r.loc[contract]
    scope=scope[scope.numero_preco.eq(item_row.numero_preco)&scope.item.eq(item_row['item'])]

def item_cards(row):
    st.subheader(f'{row.numero_preco} · Item {row["item"]}')
    st.write(row.descricao if pd.notna(row.descricao) else 'Descrição não informada')
    if not row.chave_unica:
        st.error('Chave composta duplicada na Receita. Valores contratuais não exibidos como referência única; consulte Qualidade dos dados.')
        return
    st.caption('Unidade: '+(str(row.unidade) if pd.notna(row.unidade) else 'não informada na fonte'))
    for col,(label,value,prefix) in zip(st.columns(3),[
        ('Receita unitária',row.receita_unitaria,'R$ '),('Quantidade contratada',row.quantidade_contratada,''),('Quantidade executada',row.quantidade_executada,'')]):
        col.metric(label,fmt(value,prefix=prefix))
    for col,(label,value) in zip(st.columns(3),[('Saldo do item contratual',row.saldo),('Executado %',row.executado_pct),('CPUs relacionadas',row.cpus_relacionadas)]):
        col.metric(label,fmt(value,0 if label=='CPUs relacionadas' else 2))
    st.caption('Quantidades exibidas uma única vez, no item contratual. Não representam uma alocação por CPU.')

comparison_cols=['cpu','local','b6','b8','delta','delta_pct','receita_unitaria','fcd_b6','fcd_b8']

if scope.empty:
    st.info('Nenhuma CPU disponível na base selecionada.')
    st.stop()
scope=scope.reset_index(drop=True)
chosen=st.selectbox('CPU de custo',list(scope.index),format_func=lambda idx:scope.loc[idx,'cpu']+' · '+str(scope.loc[idx,'numero_preco'])+' / '+str(scope.loc[idx,'item']))
cpu=scope.loc[chosen]
detail=c[c.cpu.eq(cpu.cpu)&c.numero_preco.fillna('').eq(cpu.numero_preco if pd.notna(cpu.numero_preco) else '')&c.item.fillna('').eq(cpu['item'] if pd.notna(cpu['item']) else '')]
if page=='Visão Executiva':
    st.subheader(cpu.cpu)
    st.caption('Local / anotações: '+cpu.local)
    st.caption(f'Nº DO PREÇO: {cpu.numero_preco} · Item: {cpu["item"]} · '+(str(cpu.descricao) if pd.notna(cpu.descricao) else 'Sem correspondência única na receita'))
    for col,(label,field) in zip(st.columns(4),[('B6 R$/un','b6'),('B8 preliminar R$/un','b8'),('Delta R$/un','delta'),('Delta %','delta_pct')]):
        col.metric(label,fmt(cpu[field]))
    for col,(label,field) in zip(st.columns(3),[('Receita unitária','receita_unitaria'),('F/CD B6','fcd_b6'),('F/CD B8','fcd_b8')]):
        col.metric(label,fmt(cpu[field],prefix='R$ ' if field=='receita_unitaria' else ''))
    if strategy_change(detail): st.warning('Mudança de estratégia de execução — possível transferência de mão de obra própria para serviços de terceiros. A redução de incidência não comprova ganho de produtividade.')
    if (detail.preco_b6.isna()|detail.inc_b6.isna()).any(): st.caption('Há fatores B6 vazios; os custos salvos no Excel foram preservados. Detalhes em Qualidade dos dados.')

    groups=decomposition(detail).sort_values('b8',ascending=False,na_position='last')
    resources=resource_costs(detail)

    st.subheader('Composição do custo por grupo')
    comparison=pd.concat([groups[['classe',base,'participacao_'+base]].rename(columns={base:'R$/un','participacao_'+base:'Participação %'}).assign(Base=label)
                          for base,label in [('b6','B6'),('b8','B8 preliminar')]],ignore_index=True)
    chart(comparison,'R$/un','classe','Base',share='Participação %')
    st.caption('Grupos ordenados pelo custo B8. Percentuais representam a participação no total da respectiva base.')
    grid(groups[['classe','b6','b8','participacao_b6','participacao_b8']],'estrutura_custo')

    st.subheader('Principais insumos do custo')
    top_cost=resources.sort_values('b8',ascending=False,na_position='last').head(10)
    chart(top_cost,'b8','insumo',share='participacao_b8')
    st.caption('Até 10 insumos, em ordem decrescente de custo B8. Participação sobre o custo B8 total da CPU; nomes repetidos são agregados.')
    grid(top_cost[['insumo','b6','b8','participacao_b6','participacao_b8']],'principais_custos')

elif page=='Análise Comparativa':
    groups=decomposition(detail).sort_values('b8',ascending=False,na_position='last')
    resources=resource_costs(detail)

    st.subheader('Variação por grupo de custo')
    variation=groups.loc[groups.delta.abs().sort_values(ascending=False).index].copy()
    variation['Efeito']=variation.delta.map(lambda x:'Indisponível' if pd.isna(x) else 'Redução' if x<0 else 'Aumento' if x>0 else 'Sem variação')
    chart(variation,'delta','classe','Efeito')
    grid(variation[['classe','delta','participacao_delta']],'variacao_grupos')
    st.caption('Delta negativo = redução; positivo = aumento. Participação na variação = delta do grupo / delta total; pode ser negativa ou maior que 100%. Com total zero, fica indisponível.')

    st.subheader('Principais insumos da variação')
    top=resources.loc[resources.delta.abs().sort_values(ascending=False).index].head(10).copy()
    top['Efeito']=top.delta.map(lambda x:'Indisponível' if pd.isna(x) else 'Redução' if x<0 else 'Aumento' if x>0 else 'Sem variação')
    chart(top,'delta','insumo','Efeito')
    grid(top[['insumo','b6','b8','delta']],'variacao_insumos')

    st.subheader('Tabela detalhada da composição')
    grid(detail[['insumo','classe','codigo','local','preco_b6','preco_b8','inc_b6','inc_b8','inc_revisao','b6','b8','delta','salvo_b6','salvo_b8','aba','linha_excel']],'insumos')

else:
    labor=detail[detail.classe.map(normalized).eq('MAO DE OBRA')]
    st.subheader('Incidências de mão de obra')
    st.caption('Diferenças em relação ao B6. Incidência de revisão é independente da ajustada e não substitui automaticamente a base de custos.')
    if labor.empty:
        st.info('Esta CPU não tem insumos classificados como mão de obra.')
    else:
        bars=labor.melt(id_vars='insumo',value_vars=['inc_b6','inc_b8','inc_revisao'],var_name='Base',value_name='Incidência').dropna(subset=['Incidência'])
        bars['Base']=bars.Base.map({'inc_b6':'B6','inc_b8':'Ajustada','inc_revisao':'Revisão'})
        chart(bars,'Incidência','insumo','Base')
        grid(labor[['insumo','local','inc_b6','inc_b8','inc_revisao','diferenca_inc','diferenca_inc_pct','diferenca_revisao','diferenca_revisao_pct','linha_excel']],'produtividade')
        if labor.inc_revisao.isna().all(): st.caption('Incidência revisão ainda não preenchida: campo opcional nesta etapa.')
