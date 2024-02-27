def custo_unitario_execucao(sheet_planilha_comp, linha_inicial,
                            linha_total,
                            coluna_totais_composicao, coluna_valor_string):
    linha_custo_horario_execucao = (
        sheet_planilha_comp
        [f'{coluna_totais_composicao}{linha_total+1}'].value)
    if (linha_custo_horario_execucao == 'Custo Horário da Execução:'):
        sheet_planilha_comp
        [f'{coluna_valor_string}{linha_total+1}'].value = (
            '=Subtotal(9,' + f'{coluna_valor_string}{linha_inicial}:{
                coluna_valor_string}{linha_total}' + ")"
        )
