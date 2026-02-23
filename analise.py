import pandas as pd
import numpy as np
import os

PASTA_DOWNLOADS = os.path.join(os.getcwd(), "downloads")
ARQUIVO_CSV = os.path.join(PASTA_DOWNLOADS, "relatorio_chats_atualizado.csv")
ARQUIVO_EXCEL = os.path.join(PASTA_DOWNLOADS, "relatorio_chats_pronto.xlsx")

def formatar_tempo_exato(td):
    if pd.isna(td): return ""
    segundos_totais = max(0, int(td.total_seconds()))
    horas, resto = divmod(segundos_totais, 3600)
    minutos, segundos = divmod(resto, 60)
    return f"{horas:02d}:{minutos:02d}:{segundos:02d}"

def analisar_e_limpar_dados():
    print("📊 A iniciar a Estruturação Cronológica Baseada em Tempos...")

    try:
        df = pd.read_csv(ARQUIVO_CSV, sep=',', low_memory=False)

        # 1. LIMPEZA DE TEXTOS (Evitar notação científica)
        colunas_texto = ['Telefone do contato', 'Id do atendimento', 'Id do cliente', 'CPF do contato']
        for col in colunas_texto:
            if col in df.columns:
                df[col] = df[col].fillna(-1).astype(str).str.replace(r'\.0$', '', regex=True)
                df[col] = df[col].replace('-1', '')

        # 2. TRATAMENTO DAS DATAS (Ponto de partida para a verdade)
        dt_chegada = pd.to_datetime(df.get('Data de criação do chat'), errors='coerce').dt.tz_localize(None)
        dt_resposta = pd.to_datetime(df.get('Data de primeira resposta'), errors='coerce').dt.tz_localize(None)
        dt_fim = pd.to_datetime(df.get('Data de finalização do chat'), errors='coerce').dt.tz_localize(None)

        # ========================================================
        # 3. CONTEXTO DA CHEGADA (Horários e Expediente)
        # ========================================================
        horas = dt_chegada.dt.hour
        
        df['Período do Dia'] = np.select(
            [(horas >= 0) & (horas < 6), (horas >= 6) & (horas < 12), (horas >= 12) & (horas < 18), (horas >= 18) & (horas <= 23)],
            ['Madrugada', 'Manhã', 'Tarde', 'Noite'], default='Desconhecido'
        )

        def verificar_expediente(dt):
            if pd.isna(dt): return "Desconhecido"
            if dt.dayofweek >= 5: return "NÃO (Fim de Semana)" 
            
            minutos_do_dia = dt.hour * 60 + dt.minute
            inicio_expediente = 8 * 60 + 1  # 08:01
            fim_expediente = 17 * 60 + 59   # 17:59
            
            if inicio_expediente <= minutos_do_dia <= fim_expediente:
                return "SIM"
            return "NÃO (Fora do Horário)"
            
        df['Dentro do Expediente?'] = dt_chegada.apply(verificar_expediente)

        # ========================================================
        # 4. AVALIAÇÃO DA FILA E DO ATENDIMENTO (O Cronómetro)
        # ========================================================
        def avaliar_espera(row, data_chegada, data_resposta, data_fim):
            # NOVO: LÓGICA DE QUEM FOI IGNORADO
            if pd.isna(data_resposta): 
                if pd.isna(data_fim):
                    return "⏳ Na Fila (Ainda em Aberto)"
                
                # Se foi fechado sem resposta, calcula quanto tempo o cliente ficou no vácuo
                delta_vacuo = (data_fim - data_chegada).total_seconds()
                if data_fim.date() > data_chegada.date():
                    return "⚠️ Vácuo até o Dia Seguinte (Fechado sem resposta)"
                
                minutos_vacuo = delta_vacuo / 60
                
                # Diferencia se foi o robô que fechou ou se foi um humano que ignorou
                if pd.isna(row.get('Atendente')) or str(row.get('Atendente')).strip() == '':
                    return f"🤖 Sistema/Robô (Encerrado após {int(minutos_vacuo)} min)"
                else:
                    return f"👻 Vácuo Total (Fechado após {int(minutos_vacuo)} min)"

            # LÓGICA NORMAL DE QUEM FOI RESPONDIDO
            delta = (data_resposta - data_chegada).total_seconds()
            
            if data_resposta.date() > data_chegada.date(): 
                return "⚠️ Passou para o Dia Seguinte"
            
            minutos = delta / 60
            if minutos <= 5: return "🟢 Rápido (< 5 min)"
            elif minutos <= 15: return "🟡 Aceitável (5 a 15 min)"
            else: return "🟠 Demorado (> 15 min)"
            
        df['Avaliação da Espera'] = [avaliar_espera(row, c, r, f) for row, c, r, f in zip(df.to_dict('records'), dt_chegada, dt_resposta, dt_fim)]

        def diagnosticar_conversa(row, resp, fim):
            if pd.isna(row.get('Atendente')): 
                return "🤖 Retido no Robô"
            if pd.isna(resp): 
                return "👻 Ignorado (Atendente nunca respondeu)"
            
            if pd.notna(fim):
                tempo_conversa_seg = (fim - resp).total_seconds()
                if tempo_conversa_seg < 60: 
                    return "⚡ Fechamento Imediato (Sem diálogo longo)"
                else:
                    return "✅ Atendimento com Interação"
            
            return "⏳ Em Andamento"

        df['Diagnóstico da Conversa'] = [diagnosticar_conversa(row, r, f) for row, r, f in zip(df.to_dict('records'), dt_resposta, dt_fim)]

        df['Status Final'] = np.where(dt_fim.notna(), "Encerrado", "Em Aberto")

        # ========================================================
        # 5. CÁLCULO EXATO DE TEMPOS (A NOVA LÓGICA MATADORA)
        # ========================================================
        # Se teve 1ª Resposta, mede até ela. Se não teve, mede até o Encerramento!
        data_limite_espera = dt_resposta.fillna(dt_fim)
        
        df['Tempo de Espera (Fila)'] = (data_limite_espera - dt_chegada).apply(formatar_tempo_exato)
        
        # O Tempo de Conversa só existe se realmente houve conversa (1ª Resposta não nula)
        # Caso contrário, fica vazio, pois a pessoa apenas esperou na fila até fecharem.
        df['Tempo de Conversa (Atendimento)'] = np.where(
            dt_resposta.notna(), 
            (dt_fim - dt_resposta).apply(formatar_tempo_exato), 
            ""
        )
        
        df['Tempo Total (Início ao Fim)'] = (dt_fim - dt_chegada).apply(formatar_tempo_exato)

        # ========================================================
        # 6. CRIAÇÃO DAS DATAS E HORAS SEPARADAS
        # ========================================================
        df['1. Data de Entrada'] = dt_chegada.dt.normalize()
        df['1. Hora de Entrada'] = dt_chegada.dt.time

        df['2. Data da 1ª Resposta'] = dt_resposta.dt.normalize()
        df['2. Hora da 1ª Resposta'] = dt_resposta.dt.time

        df['3. Data de Encerramento'] = dt_fim.dt.normalize()
        df['3. Hora de Encerramento'] = dt_fim.dt.time

        # ========================================================
        # 7. ORDENAÇÃO CRONOLÓGICA PERFEITA
        # ========================================================
        ordem_historia = [
            'Id do atendimento',
            'Cliente',
            'Telefone do contato',
            
            # A CHEGADA
            '1. Data de Entrada',
            '1. Hora de Entrada',
            'Período do Dia',
            'Dentro do Expediente?',
            
            # A DISTRIBUIÇÃO
            'Houve redirecionamento',
            'Departamento do Chat',
            'Atendente',
            'Tempo de Espera (Fila)',
            'Avaliação da Espera',
            
            # O INÍCIO DO ATENDIMENTO
            '2. Data da 1ª Resposta',
            '2. Hora da 1ª Resposta',
            
            # O DECORRER DA CONVERSA
            'Diagnóstico da Conversa',
            'Tempo de Conversa (Atendimento)',
            
            # O FIM DO ATENDIMENTO
            '3. Data de Encerramento',
            '3. Hora de Encerramento',
            'Fechado por',
            'Tempo Total (Início ao Fim)',
            'Status Final',
            
            # TABULAÇÃO
            'Motivo do serviço',
            'Motivo do fechamento'
        ]

        colunas_finais = [col for col in ordem_historia if col in df.columns]
        df_final = df[colunas_finais].copy()

        # ========================================================
        # 8. EXPORTAÇÃO EXCEL (Tabela Limpa)
        # ========================================================
        print("🎨 A construir a Tabela Excel...")
        writer = pd.ExcelWriter(
            ARQUIVO_EXCEL, 
            engine='xlsxwriter', 
            datetime_format='dd/mm/yyyy',
            engine_kwargs={'options': {'strings_to_urls': False}}
        )

        nome_aba = 'Relatorio_Chats'
        df_final.to_excel(writer, index=False, header=False, startrow=1, sheet_name=nome_aba)
        
        workbook = writer.book
        ws = writer.sheets[nome_aba]
        ws.set_tab_color('#FF8C00') 
        
        (max_r, max_c) = df_final.shape
        if max_r > 0:
            ws.add_table(0, 0, max_r, max_c - 1, {
                'columns': [{'header': str(c)} for c in df_final.columns],
                'style': 'Table Style Medium 9',
                'name': 'Tab_Chats'
            })
        else:
            ws.write_row(0, 0, df_final.columns)

        ws.ignore_errors({'number_stored_as_text': 'A1:XFD1048576'})

        fmt_hora = workbook.add_format({'num_format': 'hh:mm:ss'})
        fmt_central = workbook.add_format({'align': 'center'})

        for i, col in enumerate(df_final.columns):
            try: tamanho = int(df_final[col].fillna("").astype(str).str.len().max())
            except: tamanho = 10
            
            largura = min(max(tamanho, len(str(col))) + 2, 45)

            if "Hora " in col or "Tempo " in col:
                ws.set_column(i, i, largura, fmt_hora)
            elif "Houve" in col:
                ws.set_column(i, i, largura, fmt_central)
            else:
                ws.set_column(i, i, largura)

        writer.close()
        print(f"🎉 SUCESSO! A base de dados agora cobre os casos de vácuo perfeitamente.")
        print(f"Abra o ficheiro em: {ARQUIVO_EXCEL}")

    except Exception as e:
        print(f"❌ Ocorreu um erro crítico no script: {e}")

if __name__ == "__main__":
    analisar_e_limpar_dados()