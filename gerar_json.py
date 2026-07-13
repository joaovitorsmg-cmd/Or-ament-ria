# -*- coding: utf-8 -*-
# gerar_json.py -- Agroquima Painel Orcamentario
# Execute: python gerar_json.py  (ou de 2 cliques neste arquivo)
import pandas as pd, json, os, glob, sys

DESPESA_FILE   = 'Mapa_Despesa.xlsx'
REGIONAIS_FILE = 'Base_regionais.xlsx'
OUT            = 'aqm_data.json'


def encontrar_arquivo(nome_esperado, pistas):
    """Retorna o caminho do arquivo; lanca FileNotFoundError se nao achar."""
    if os.path.exists(nome_esperado):
        return nome_esperado
    candidatos = [f for f in glob.glob('*.xlsx') if not os.path.basename(f).startswith('~$')]
    for f in candidatos:
        chave = f.lower().replace(' ', '').replace('_', '')
        if all(p in chave for p in pistas):
            print("[INFO] '{}' nao encontrado -- usando '{}' no lugar (nome parecido).".format(nome_esperado, f))
            return f
    print("[ERRO] Nenhum arquivo parecido com '{}' encontrado nesta pasta.".format(nome_esperado))
    if candidatos:
        print("       Arquivos .xlsx encontrados aqui:")
        for f in candidatos:
            print("         - " + f)
    else:
        print("       Nenhum arquivo .xlsx encontrado nesta pasta.")
    print("       Coloque o arquivo certo aqui, ou renomeie para '{}'.".format(nome_esperado))
    raise FileNotFoundError(nome_esperado)


def main():
    despesa_path   = encontrar_arquivo(DESPESA_FILE,   ['despesa'])
    regionais_path = encontrar_arquivo(REGIONAIS_FILE, ['regio'])

    d = pd.read_excel(despesa_path)
    r = pd.read_excel(regionais_path)

    # ==== PERIODO: usar SEMPRE ANO/MES/DIA (colunas I/J/K), NUNCA DATA_LANCAMENTO ====
    # DATA_LANCAMENTO/DATA_APROVACAO sao apenas a data em que a contabilidade
    # fechou/aprovou o lote no sistema -- nao representam a competencia do
    # lancamento. A competencia real e dada por ANO, MES e DIA.
    # ANO_REFERENCIA: deixe None para usar automaticamente o ano mais recente
    # presente na base (evita somar 2025 + 2026 juntos). Ou defina um ano fixo.
    ANO_REFERENCIA = None
    ano_alvo = ANO_REFERENCIA if ANO_REFERENCIA else int(d['ANO'].max())
    print("[INFO] Ano de referencia do painel: {} (defina ANO_REFERENCIA no topo do script para mudar)".format(ano_alvo))
    d = d[d['ANO'] == ano_alvo].copy()

    fl  = dict(zip(r['FILIAL'], r['LOJA']))
    fr  = dict(zip(r['FILIAL'], r['REGIONAL']))
    fa  = dict(zip(r['FILIAL'], r['AUDITOR']))
    fs  = dict(zip(r['FILIAL'], r['SIGLA']))
    fg  = dict(zip(r['FILIAL'], r['GERENTE']))
    fsu = dict(zip(r['FILIAL'], r['SUPERVISOR']))
    fe  = dict(zip(r['FILIAL'], r['ESTADO']))

    d['LOJA']       = d['ID_FILIAL'].map(fl)
    d['REGIONAL']   = d['ID_FILIAL'].map(fr)
    d['AUDITOR']    = d['ID_FILIAL'].map(fa)
    d['SIGLA']      = d['ID_FILIAL'].map(fs).fillna(d['ID_FILIAL'].apply(lambda x: 'F{}'.format(x)))
    d['GERENTE']    = d['ID_FILIAL'].map(fg)
    d['SUPERVISOR'] = d['ID_FILIAL'].map(fsu)
    d['ESTADO']     = d['ID_FILIAL'].map(fe)
    for col in ['LOJA', 'REGIONAL', 'AUDITOR', 'GERENTE', 'SUPERVISOR', 'ESTADO']:
        d[col] = d[col].fillna('N/D')

    # ==== CLASSIFICACAO: FATURAMENTO x DESPESA ====
    # Faturamento (Realizado) = SOMA de VLR_REALIZADO SOMENTE dos lancamentos
    # cujo GRUPO_CONTA esteja listado abaixo. Tudo que nao estiver nesta lista
    # entra como despesa. Se surgir outro grupo de receita (ex.: devolucoes,
    # vendas de servicos), adicione-o aqui.
    GRUPOS_RECEITA = [
        '01-VENDAS DE MERCADORIAS',
    ]

    def _norm(s):
        return str(s).strip().upper()

    grupos_receita_norm = {_norm(g) for g in GRUPOS_RECEITA}
    d['_GC_NORM'] = d['GRUPO_CONTA'].apply(_norm)

    fat  = d[ d['_GC_NORM'].isin(grupos_receita_norm)].copy()
    desp = d[~d['_GC_NORM'].isin(grupos_receita_norm)].copy()

    # Diagnostico: mostra como cada GRUPO_CONTA da base foi classificado e o
    # valor de VLR_REALIZADO somado em cada um. Use isso para conferir se o
    # faturamento total bate com o esperado -- se houver um grupo de receita
    # faltando na lista acima, ele vai aparecer somado dentro de "DESPESA".
    _resumo = d.groupby('GRUPO_CONTA')['VLR_REALIZADO'].agg(['count', 'sum']).reset_index()
    _resumo['CLASSIFICACAO'] = _resumo['GRUPO_CONTA'].apply(
        lambda g: 'RECEITA' if _norm(g) in grupos_receita_norm else 'DESPESA')
    print("\n[INFO] Classificacao por GRUPO_CONTA (confira se esta correto):")
    for _, row in _resumo.sort_values('sum', ascending=False).iterrows():
        print("  [{:8s}] {:45s} | {:6d} lanc. | R$ {:,.2f}".format(
            row['CLASSIFICACAO'], row['GRUPO_CONTA'], int(row['count']), row['sum']))
    print()

    if fat.empty:
        raise RuntimeError(
            "Nenhum lancamento casou com GRUPOS_RECEITA -- faturamento ficaria zerado. "
            "Confira os nomes exatos em GRUPO_CONTA na planilha.")

    mes_nomes = {1:'Jan',2:'Fev',3:'Mar',4:'Abr',5:'Mai',6:'Jun',
                 7:'Jul',8:'Ago',9:'Set',10:'Out',11:'Nov',12:'Dez'}

    # Periodo exibido no cabecalho = meses do ano de referencia que ja tem
    # faturamento realizado (> 0). Meses futuros com apenas orcamento ficam
    # fora do rotulo, mas continuam no grafico de sazonalidade.
    meses_com_realizado = sorted(fat[fat['VLR_REALIZADO'] > 0]['MES'].unique())
    if meses_com_realizado:
        periodo_str = "{}-{} {}".format(
            mes_nomes[min(meses_com_realizado)],
            mes_nomes[max(meses_com_realizado)],
            ano_alvo)
    else:
        periodo_str = str(ano_alvo)

    # Divisao segura: evita NaN/Infinity quando o denominador e 0 (ex.: mes
    # futuro sem faturamento realizado ainda). NaN/Infinity nao sao JSON valido
    # e quebram o carregamento do painel no navegador (JSON.parse falha).
    def pct(numer, denom):
        return (numer / denom * 100).replace([float('inf'), float('-inf')], 0).fillna(0).round(2)

    def pct_s(n, d_val):
        return round(n / d_val * 100, 2) if d_val else 0.0

    # ==== JANELA DE COMPARACAO (Realizado x Planejado) ====
    # A planilha traz o orcamento do ANO INTEIRO (Jan-Dez), mas o Realizado so
    # existe ate o ultimo mes ja fechado (ex.: Jan-Mai). Comparar Realizado
    # desses 5 meses contra o Planejado de 12 meses gera desvios gigantes e
    # incorretos (sempre parece que ficou muito abaixo do plano). fat_comp/
    # desp_comp restringem os DOIS lados (Realizado e Planejado) aos mesmos
    # meses, para que Desvio % seja uma comparacao valida.
    fat_comp  = fat[ fat['MES'].isin(meses_com_realizado)].copy()
    desp_comp = desp[desp['MES'].isin(meses_com_realizado)].copy()

    fat_fil = fat_comp.groupby(
        ['ID_FILIAL', 'SIGLA', 'LOJA', 'REGIONAL', 'AUDITOR', 'GERENTE', 'SUPERVISOR', 'ESTADO']
    ).agg(FAT_REAL=('VLR_REALIZADO', 'sum'), FAT_PREV=('VLR_PREVISAO', 'sum')).reset_index()
    desp_fil = desp_comp.groupby('ID_FILIAL').agg(
        DESP_REAL=('VLR_REALIZADO', 'sum'), DESP_PREV=('VLR_PREVISAO', 'sum')).reset_index()
    filiais = fat_fil.merge(desp_fil, on='ID_FILIAL', how='left').fillna(0)
    filiais['PERC_DESP']   = pct(filiais['DESP_REAL'], filiais['FAT_REAL'])
    filiais['DESVIO_FAT']  = pct(filiais['FAT_REAL']  - filiais['FAT_PREV'],  filiais['FAT_PREV'])
    filiais['DESVIO_DESP'] = pct(filiais['DESP_REAL'] - filiais['DESP_PREV'], filiais['DESP_PREV'])
    filiais = filiais.sort_values('FAT_REAL', ascending=False)

    # meses: usa o ano inteiro (fat/desp, nao fat_comp/desp_comp) -- a
    # comparacao ja e mes-a-mes, entao meses futuros (Realizado=0) sao validos.
    fat_mes  = fat.groupby('MES').agg(FAT_REAL=('VLR_REALIZADO','sum'), FAT_PREV=('VLR_PREVISAO','sum')).reset_index()
    desp_mes = desp.groupby('MES').agg(DESP_REAL=('VLR_REALIZADO','sum'), DESP_PREV=('VLR_PREVISAO','sum')).reset_index()
    meses = fat_mes.merge(desp_mes, on='MES').fillna(0)
    meses['NOME'] = meses['MES'].map(mes_nomes)
    meses['PERC_DESP'] = pct(meses['DESP_REAL'], meses['FAT_REAL'])

    # filiais_mes: igual a "meses", mas quebrado por filial -- necessario para
    # o filtro de Periodo na Visao Geral funcionar combinado com Regional/Auditor.
    fat_fil_mes  = fat.groupby(['ID_FILIAL','MES']).agg(
        FAT_REAL=('VLR_REALIZADO','sum'), FAT_PREV=('VLR_PREVISAO','sum')).reset_index()
    desp_fil_mes = desp.groupby(['ID_FILIAL','MES']).agg(
        DESP_REAL=('VLR_REALIZADO','sum'), DESP_PREV=('VLR_PREVISAO','sum')).reset_index()
    filiais_mes = fat_fil_mes.merge(desp_fil_mes, on=['ID_FILIAL','MES'], how='outer').fillna(0)
    filiais_mes['ID_FILIAL'] = filiais_mes['ID_FILIAL'].astype(int)
    filiais_mes['MES']       = filiais_mes['MES'].astype(int)

    fat_reg  = fat_comp.groupby('REGIONAL').agg(FAT_REAL=('VLR_REALIZADO','sum'), FAT_PREV=('VLR_PREVISAO','sum')).reset_index()
    desp_reg = desp_comp.groupby('REGIONAL').agg(DESP_REAL=('VLR_REALIZADO','sum'), DESP_PREV=('VLR_PREVISAO','sum')).reset_index()
    regionais_out = fat_reg.merge(desp_reg, on='REGIONAL').fillna(0)
    regionais_out['PERC_DESP'] = pct(regionais_out['DESP_REAL'], regionais_out['FAT_REAL'])

    fat_aud  = fat_comp.groupby('AUDITOR').agg(FAT_REAL=('VLR_REALIZADO','sum'), FAT_PREV=('VLR_PREVISAO','sum')).reset_index()
    desp_aud = desp_comp.groupby('AUDITOR').agg(DESP_REAL=('VLR_REALIZADO','sum'), DESP_PREV=('VLR_PREVISAO','sum')).reset_index()
    auditores = fat_aud.merge(desp_aud, on='AUDITOR').fillna(0)
    auditores['PERC_DESP'] = pct(auditores['DESP_REAL'], auditores['FAT_REAL'])

    grupos = (desp_comp.groupby('GRUPO_CONTA')
              .agg(DESP_REAL=('VLR_REALIZADO','sum'), DESP_PREV=('VLR_PREVISAO','sum'))
              .reset_index().sort_values('DESP_REAL', ascending=False))
    grupos['DESVIO'] = pct(grupos['DESP_REAL'] - grupos['DESP_PREV'], grupos['DESP_PREV'])

    grupos_filial = (desp_comp.groupby(['ID_FILIAL','GRUPO_CONTA'])
                    .agg(DESP_REAL=('VLR_REALIZADO','sum'), DESP_PREV=('VLR_PREVISAO','sum'))
                    .reset_index())
    grupos_filial['ID_FILIAL'] = grupos_filial['ID_FILIAL'].astype(int)

    tot_fat_real     = fat_comp['VLR_REALIZADO'].sum()
    tot_fat_prev     = fat_comp['VLR_PREVISAO'].sum()
    tot_desp_real    = desp_comp['VLR_REALIZADO'].sum()
    tot_desp_prev    = desp_comp['VLR_PREVISAO'].sum()
    # Orcamento do ano inteiro (Jan-Dez), so para referencia/contexto.
    tot_fat_prev_ano  = fat['VLR_PREVISAO'].sum()
    tot_desp_prev_ano = desp['VLR_PREVISAO'].sum()

    payload = {
        'meta': {
            'periodo':          periodo_str,
            'ano':              int(ano_alvo),
            'filiais_count':    int(len(filiais)),
            'tot_fat_real':     round(tot_fat_real,  2),
            'tot_fat_prev':     round(tot_fat_prev,  2),
            'tot_desp_real':    round(tot_desp_real, 2),
            'tot_desp_prev':    round(tot_desp_prev, 2),
            'tot_fat_prev_ano': round(tot_fat_prev_ano,  2),
            'tot_desp_prev_ano':round(tot_desp_prev_ano, 2),
            'perc_desp_real':   pct_s(tot_desp_real, tot_fat_real),
            'perc_desp_plan':   pct_s(tot_desp_prev, tot_fat_prev),
            'resultado':        round(tot_fat_real - tot_desp_real, 2),
            'dev_fat':          pct_s(tot_fat_real  - tot_fat_prev,  tot_fat_prev),
            'dev_desp':         pct_s(tot_desp_real - tot_desp_prev, tot_desp_prev),
        },
        'filiais':      filiais.to_dict('records'),
        'meses':        meses.to_dict('records'),
        'filiais_mes':  filiais_mes.to_dict('records'),
        'regionais':    regionais_out.to_dict('records'),
        'auditores':    auditores.to_dict('records'),
        'grupos':       grupos.to_dict('records'),
        'grupos_filial':grupos_filial.to_dict('records'),
    }

    with open(OUT, 'w', encoding='utf-8') as fout:
        json.dump(payload, fout, ensure_ascii=False, default=str)

    print("\n[OK] {} gerado com sucesso!".format(OUT))
    print("[OK] Grupos de receita: {}".format(', '.join(GRUPOS_RECEITA)))
    print("[OK] Fat: R$ {:,.0f}  |  Desp: R$ {:,.0f}  |  Resultado: R$ {:,.0f}".format(
        tot_fat_real, tot_desp_real, tot_fat_real - tot_desp_real))


if __name__ == '__main__':
    # Roda tudo dentro de um try/except e SEMPRE espera ENTER no final --
    # assim, se der 2 cliques no arquivo no Windows, a janela do console
    # nao fecha sozinha antes de dar tempo de ler o resultado (ou o erro).
    try:
        main()
        print("\nConcluido! Arraste o aqm_data.json no painel (aba Atualizar Base).")
    except FileNotFoundError as e:
        print("\n[ERRO] Arquivo nao encontrado: {}".format(e))
        print("       Verifique se todos os arquivos estao na mesma pasta que este .py")
    except Exception as e:
        print("\n" + "=" * 60)
        print("[ERRO] Falha ao gerar o aqm_data.json:")
        print("    {}: {}".format(type(e).__name__, e))
        print("=" * 60)
    input("\nPressione ENTER para fechar esta janela...")
