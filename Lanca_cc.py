import streamlit as st
import pandas as pd
import win32com.client
import time
from io import BytesIO

# Configuração da página
st.set_page_config(
    page_title="SAP - Carta de Correção",
    page_icon="📝",
    layout="wide"
)

# Título
st.title("📝 Automatização de Carta de Correção SAP - by Djalma")
st.markdown("---")

# Classe para interação com SAP
class SAPConnector:
    def __init__(self):
        self.session = None
        self.connection = None
        self.sap_gui_auto = None
        
    def connect(self):
        """Conecta ao SAP GUI"""
        try:
            self.sap_gui_auto = win32com.client.GetObject("SAPGUI")
            if not type(self.sap_gui_auto) == win32com.client.CDispatch:
                return False
            
            application = self.sap_gui_auto.GetScriptingEngine
            if not type(application) == win32com.client.CDispatch:
                return False
            
            self.connection = application.Children(0)
            self.session = self.connection.Children(0)
            return True
        except Exception as e:
            st.error(f"Erro ao conectar ao SAP: {str(e)}")
            return False
    
    def process_carta_correcao(self, doc_num, texto_correcao):
        """Processa uma carta de correção no SAP"""
        try:
            # Navega para a transação J1BNFE
            self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/N J1BNFE"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            
            # Preenche o número do documento
            self.session.findById("wnd[0]/usr/txtDOCNUM-LOW").Text = str(doc_num)
            self.session.findById("wnd[0]/usr/ctxtDATE0-LOW").Text = ""
            self.session.findById("wnd[0]/usr/ctxtBUKRS-LOW").Text = "1000"
            
            # Executa a busca
            self.session.findById("wnd[0]").sendVKey(8)
            time.sleep(0.5)
            
            # Seleciona o registro
            self.session.findById("wnd[0]/usr/cntlNFE_CONTAINER/shellcont/shell").currentCellColumn = ""
            self.session.findById("wnd[0]/usr/cntlNFE_CONTAINER/shellcont/shell").selectedRows = "0"
            
            # Abre menu de carta de correção
            self.session.findById("wnd[0]/mbar/menu[4]/menu[0]/menu[0]").Select()
            time.sleep(0.5)
            
            # Insere o texto da correção
            self.session.findById("wnd[1]/usr/cntlTEXTEDITOR1/shellcont/shell").Text = texto_correcao
            self.session.findById("wnd[1]/usr/cntlTEXTEDITOR1/shellcont/shell").setSelectionIndexes(len(texto_correcao), len(texto_correcao))
            
            # Confirma
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            time.sleep(0.5)
            
            # Captura a mensagem de retorno
            status = self.session.findById("wnd[0]/sbar/pane[0]").Text
            
            return status
            
        except Exception as e:
            return f"Erro: {str(e)}"
    
    def disconnect(self):
        """Desconecta do SAP"""
        self.session = None
        self.connection = None
        self.sap_gui_auto = None

# Função para converter DataFrame para Excel
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Resultado')
    return output.getvalue()

# Interface Streamlit
def main():
    # Sidebar com instruções
    with st.sidebar:
        st.header("ℹ️ Instruções")
        st.markdown("""
        ### Como usar:
        1. **Faça upload** do arquivo Excel com os dados
        2. O arquivo deve conter:
           - **Coluna A**: Número do Documento
           - **Coluna B**: Texto da Correção
        3. Clique em **Processar** para executar
        4. **Baixe** o resultado com os status
        
        ### Formato do arquivo:
        | Documento | Texto Correção |
        |-----------|----------------|
        | 66693215  | Texto aqui...  |
        | 66693216  | Texto aqui...  |
        """)
        
        st.markdown("---")
        st.info("💡 **Dica**: Certifique-se de que o SAP GUI está aberto e logado antes de processar.")
    
    # Área principal
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("📤 Upload do Arquivo")
        uploaded_file = st.file_uploader(
            "Selecione o arquivo Excel",
            type=['xlsx', 'xls'],
            help="Arquivo deve conter as colunas: Documento e Texto Correção"
        )
    
    if uploaded_file is not None:
        try:
            # Lê o arquivo
            df = pd.read_excel(uploaded_file)
            
            # Valida as colunas
            if len(df.columns) < 2:
                st.error("❌ O arquivo deve conter pelo menos 2 colunas (Documento e Texto Correção)")
                return
            
            # Renomeia as colunas para padronizar
            df.columns = ['Documento', 'Texto_Correcao'] + list(df.columns[2:])
            
            # Remove linhas vazias
            df = df.dropna(subset=['Documento'])
            
            # Adiciona coluna de status se não existir
            if 'Status' not in df.columns:
                df['Status'] = ''
            
            # Exibe preview dos dados
            st.subheader("👀 Preview dos Dados")
            st.dataframe(df.head(10), use_container_width=True)
            st.info(f"📊 Total de registros: {len(df)}")
            
            # Botão de processamento
            col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 2])
            
            with col_btn1:
                process_button = st.button("🚀 Processar", type="primary", use_container_width=True)
            
            with col_btn2:
                clear_button = st.button("🗑️ Limpar Dados", use_container_width=True)
            
            # Processamento
            if process_button:
                # Conecta ao SAP
                sap = SAPConnector()
                
                with st.spinner("🔌 Conectando ao SAP..."):
                    if not sap.connect():
                        st.error("❌ Falha ao conectar ao SAP. Verifique se o SAP GUI está aberto e logado.")
                        return
                
                st.success("✅ Conectado ao SAP com sucesso!")
                
                # Processa cada linha
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                results = []
                
                for idx, row in df.iterrows():
                    doc_num = row['Documento']
                    texto = row['Texto_Correcao']
                    
                    status_text.text(f"Processando documento {idx + 1}/{len(df)}: {doc_num}")
                    
                    # Processa no SAP
                    status = sap.process_carta_correcao(doc_num, texto)
                    
                    # Atualiza o DataFrame
                    df.at[idx, 'Status'] = status
                    results.append({
                        'Documento': doc_num,
                        'Status': status
                    })
                    
                    # Atualiza barra de progresso
                    progress_bar.progress((idx + 1) / len(df))
                    
                    # Pequena pausa entre requisições
                    time.sleep(0.5)
                
                # Desconecta do SAP
                sap.disconnect()
                
                status_text.empty()
                progress_bar.empty()
                
                # Exibe resultados
                st.success("✅ Processamento concluído!")
                
                st.subheader("📊 Resultados")
                st.dataframe(df, use_container_width=True)
                
                # Análise de resultados
                col_result1, col_result2, col_result3 = st.columns(3)
                
                success_count = df['Status'].str.contains('sucesso|êxito', case=False, na=False).sum()
                error_count = df['Status'].str.contains('erro|falha', case=False, na=False).sum()
                
                with col_result1:
                    st.metric("Total Processado", len(df))
                with col_result2:
                    st.metric("Sucesso", success_count)
                with col_result3:
                    st.metric("Erros", error_count)
                
                # Download do resultado
                excel_data = to_excel(df)
                st.download_button(
                    label="📥 Download Resultado (Excel)",
                    data=excel_data,
                    file_name=f"resultado_carta_correcao_{time.strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            # Limpar dados
            if clear_button:
                df['Status'] = ''
                st.success("✅ Dados limpos com sucesso!")
                st.rerun()
                
        except Exception as e:
            st.error(f"❌ Erro ao processar arquivo: {str(e)}")
    
    else:
        # Exibe mensagem quando não há arquivo
        st.info("👆 Faça upload de um arquivo Excel para começar")
        
        # Template de exemplo
        with st.expander("📝 Baixar Template de Exemplo"):
            template_df = pd.DataFrame({
                'Documento': ['66693215', '66693216', '66693217'],
                'Texto_Correcao': [
                    'EM VOLUMES TRANSPORTADOS, EM PESO, CONSIDERAR: 1,50KG',
                    'CORREÇÃO DE DADOS CADASTRAIS',
                    'AJUSTE DE VALORES'
                ]
            })
            
            st.dataframe(template_df, use_container_width=True)
            
            template_excel = to_excel(template_df)
            st.download_button(
                label="📥 Download Template",
                data=template_excel,
                file_name="template_carta_correcao.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()