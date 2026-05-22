import streamlit as st
import pandas as pd
import pdfplumber
import os
import re

from io import BytesIO
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage

# ============================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================

st.set_page_config(
    page_title="Passold - Controle de Saldos",
    layout="wide"
)

# ============================================================
# LOGIN DO SISTEMA
# ============================================================

USUARIOS = {
    "Cenilda": "1234",
    "Joice": "568279",
    "pcp": "2026"
}

if "logado" not in st.session_state:
    st.session_state.logado = False

if "usuario" not in st.session_state:
    st.session_state.usuario = ""

if not st.session_state.logado:

    st.markdown("""
        <div style='text-align:center; padding-top:80px;'>
            <h1 style='color:#1e3a8a;'>🔐 PASSOLD</h1>
            <h3 style='color:#475569;'>Sistema de Gestão de Saldos</h3>
        </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])

    with col2:

        usuario = st.text_input("Usuário", placeholder="Digite seu usuário")
        senha = st.text_input("Senha", type="password", placeholder="Digite sua senha")

        if st.button("Entrar", use_container_width=True):

            if usuario in USUARIOS and USUARIOS[usuario] == senha:
                st.session_state.logado = True
                st.session_state.usuario = usuario
                st.rerun()
            else:
                st.error("Usuário ou senha inválidos.")

    st.stop()

# ============================================================
# CSS
# ============================================================

st.markdown("""
    <style>
        .reportview-container { background: #f8fafc; }
        h1, h2, h3 { color: #0f172a; font-family: 'Segoe UI', Arial, sans-serif; }
        .stButton>button {
            background-color: #1e3a8a;
            color: white;
            border-radius: 4px;
            border: none;
            font-weight: 600;
            padding: 0.5rem 2rem;
        }
        .stButton>button:hover { background-color: #1d4ed8; color: white; }
        .stTabs [data-baseweb="tab"] { font-size: 14px; font-weight: 600; color: #475569; }
        .stTabs [data-baseweb="tab"][aria-selected="true"] {
            color: #1e3a8a;
            border-bottom-color: #1e3a8a;
        }
    </style>
""", unsafe_allow_html=True)

# ============================================================
# TOPO
# ============================================================

col_topo1, col_topo2 = st.columns([8, 1])

with col_topo1:
    st.title("Sistema de Gestão de Saldos e Romaneios")
    st.write("Plataforma de Monitoramento e Liberação de Ordens de Produção")

with col_topo2:
    st.write("")
    st.info(f"👤 {st.session_state.usuario}")
    if st.button("🚪 Sair"):
        st.session_state.logado = False
        st.session_state.usuario = ""
        st.rerun()

st.markdown("---")

# ============================================================
# BANCO
# ============================================================

BANCO_DADOS = "banco_ops.xlsx"

if not os.path.exists(BANCO_DADOS):
    colunas = [
        "OP", "Obra", "Projeto", "Tipo_Cod",
        "Descricao", "Medida", "Qtd_Total", "Qtd_Enviada", "Saldo"
    ]
    pd.DataFrame(columns=colunas).to_excel(BANCO_DADOS, index=False)

# ============================================================
# TABS
# ============================================================

aba1, aba2, aba3 = st.tabs([
    "Importação de O.P.",
    "Emissão de Romaneio Parcial",
    "Painel Geral de Saldos"
])

# ============================================================
# ABA 1
# ============================================================

with aba1:

    st.subheader("Registro e Integração de PDF de Produção")

    col_input1, col_input2, col_input3, col_input4 = st.columns(4)

    with col_input1:
        numero_op = st.text_input("Número da OP:", placeholder="Ex: 1100")

    with col_input2:
        obra_input = st.text_input("Nome da Obra:", placeholder="Ex: CETOR DUO")

    with col_input3:
        projeto_input = st.text_input("Projeto:", placeholder="Ex: 931")

    with col_input4:
        descricao_manual = st.text_input("Descrição Manual:", placeholder="Opcional")

    arquivo_pdf = st.file_uploader("Selecione o PDF:", type=["pdf"])

    if arquivo_pdf is not None:

        st.info("PDF carregado com sucesso.")

        # ====================================================
        # DEBUG
        # ====================================================

        if st.button("🔍 Ver texto bruto do PDF"):

            with pdfplumber.open(arquivo_pdf) as pdf:
                for i, pagina in enumerate(pdf.pages):
                    texto = pagina.extract_text()
                    with st.expander(f"Página {i+1}"):
                        if texto:
                            for j, linha in enumerate(texto.split("\n")):
                                st.code(f"[{j:03d}] {repr(linha)}")
                        else:
                            st.warning("Nenhum texto encontrado.")

        # ====================================================
        # PROCESSAMENTO
        # ====================================================

        if st.button("Processar Documento Técnico"):

            if not numero_op or not obra_input:
                st.error("Preencha Número da OP e Obra.")

            else:

                itens_extraidos = []

                CABECALHO_LINHA = "Tipo Quantidade L H Peso Unit"

                IGNORAR_PREFIXOS = (
                    "Obs:", "TOTAIS", "Emitido", "Obra:", "Endereço:",
                    "Cliente:", "Tratamento:", "PESO =", "Itens da Obra"
                )

                def is_linha_subtotal(partes):
                    if not partes:
                        return False
                    try:
                        return (
                            all(
                                p.replace(',', '').replace('.', '').isdigit()
                                for p in partes
                            )
                            and len(partes) <= 4
                        )
                    except:
                        return False

                def is_linha_item(partes):
                    if len(partes) < 5:
                        return False
                    tipo = partes[0]
                    if not re.match(r'^[A-Za-z][A-Za-z0-9\-]+$', tipo):
                        return False
                    try:
                        qtd = int(partes[1])
                        l = int(partes[2])
                        h = int(partes[3])
                        peso = float(partes[4].replace(',', '.'))
                        return qtd > 0 and l > 0 and h > 0 and peso > 0
                    except:
                        return False

                def is_linha_descricao(partes):
                    if not partes:
                        return False
                    return bool(re.match(r'^[A-Za-z0-9]+-[A-Za-z0-9\-]+$', partes[0]))

                if descricao_manual.strip():
                    descricao_atual = descricao_manual.upper().strip()
                else:
                    descricao_atual = "SEM DESCRIÇÃO"

                with pdfplumber.open(arquivo_pdf) as pdf:

                    for pagina in pdf.pages:

                        texto = pagina.extract_text()

                        if not texto:
                            continue

                        for linha in texto.split("\n"):

                            linha = linha.strip()

                            if not linha:
                                continue

                            if any(linha.startswith(pref) for pref in IGNORAR_PREFIXOS):
                                continue

                            if CABECALHO_LINHA in linha:
                                continue

                            partes = linha.split()

                            if is_linha_subtotal(partes):
                                continue

                            if not descricao_manual.strip() and is_linha_descricao(partes):
                                descricao_atual = " ".join(partes[1:]).upper()
                                continue

                            if is_linha_item(partes):
                                try:
                                    tipo_cod = partes[0].upper()
                                    qtd = int(partes[1])
                                    medida = f"{partes[2]}x{partes[3]}"

                                    item_mapeado = {
                                        "OP": str(numero_op).strip(),
                                        "Obra": obra_input.upper().strip(),
                                        "Projeto": projeto_input.strip(),
                                        "Tipo_Cod": tipo_cod,
                                        "Descricao": descricao_atual,
                                        "Medida": medida,
                                        "Qtd_Total": qtd,
                                        "Qtd_Enviada": 0,
                                        "Saldo": qtd
                                    }

                                    itens_extraidos.append(item_mapeado)

                                except:
                                    continue

                if itens_extraidos:

                    df_novos = pd.DataFrame(itens_extraidos)

                    try:
                        df_banco = pd.read_excel(BANCO_DADOS)

                        if not df_banco.empty and "OP" in df_banco.columns:
                            df_banco = df_banco[
                                df_banco["OP"].astype(str) != str(numero_op).strip()
                            ]

                        df_final = pd.concat([df_banco, df_novos], ignore_index=True)

                    except:
                        df_final = df_novos

                    df_final.to_excel(BANCO_DADOS, index=False)

                    st.success(f"OP {numero_op} importada com {len(df_novos)} itens.")

                else:
                    st.error("Nenhum item encontrado no PDF.")

# ============================================================
# ABA 2
# ============================================================

with aba2:

    st.subheader("Ordem de Separação e Carregamento Parcial")

    if os.path.exists(BANCO_DADOS):

        try:
            df_banco = pd.read_excel(BANCO_DADOS)
        except:
            df_banco = pd.DataFrame()

        if not df_banco.empty and "OP" in df_banco.columns:

            ops_disponiveis = df_banco["OP"].unique()

            op_selecionada = st.selectbox("Selecione a OP:", ops_disponiveis)

            itens_op = df_banco[
                df_banco["OP"].astype(str) == str(op_selecionada)
            ].copy()

            col_cab1, col_cab2 = st.columns(2)

            with col_cab1:
                digitado_por = st.text_input("Digitado por:", value="JOICE")

            with col_cab2:
                endereco_obra = st.text_input("Endereço da Obra:")

            st.write("---")

            lista_liberacao = []

            for index, row in itens_op.iterrows():

                col1, col2, col3 = st.columns(3)

                codigo_item = row.get('Tipo_Cod', 'COD')
                descricao_item = row.get('Descricao', 'Sem Descrição')
                medida_item = row.get('Medida', 'Não informada')
                saldo_item = row.get('Saldo', 0)
                qtd_total_item = row.get('Qtd_Total', 0)

                with col1:
                    st.write(f"**{codigo_item}** — {descricao_item} ({medida_item})")

                with col2:
                    st.write(f"Total: {qtd_total_item} | Saldo: {saldo_item}")

                with col3:
                    qtd_saida = st.number_input(
                        f"Qtd saída {index}",
                        min_value=0,
                        max_value=int(saldo_item),
                        value=0,
                        key=f"saida_{index}"
                    )

                    if qtd_saida > 0:
                        lista_liberacao.append({
                            "index": index,
                            "item": row,
                            "qtd_saida": qtd_saida
                        })

            if len(lista_liberacao) > 0:

                st.markdown("---")
                st.write("### Itens Selecionados")

                dados_resumo = [
                    {
                        "Código": lib['item'].get('Tipo_Cod', 'COD'),
                        "Descrição": lib['item'].get('Descricao', ''),
                        "Quantidade": lib['qtd_saida']
                    }
                    for lib in lista_liberacao
                ]

                st.table(dados_resumo)

                if st.button("Executar Baixa e Emitir Romaneio"):

                    df_banco_atual = pd.read_excel(BANCO_DADOS)

                    for lib in lista_liberacao:

                        idx = lib["index"]

                        df_banco_atual.at[idx, "Qtd_Enviada"] = (
                            int(df_banco_atual.at[idx, "Qtd_Enviada"])
                            + lib["qtd_saida"]
                        )

                        df_banco_atual.at[idx, "Saldo"] = (
                            int(df_banco_atual.at[idx, "Saldo"])
                            - lib["qtd_saida"]
                        )

                    df_banco_atual.to_excel(BANCO_DADOS, index=False)

                    # ============================================================
                    # GERAR ROMANEIO EXCEL
                    # ============================================================

                    wb = Workbook()
                    ws = wb.active
                    ws.sheet_view.showGridLines = False
                    ws.title = "Romaneio"

                    bd_fina = Side(style='thin', color="000000")
                    borda_padrao = Border(
                        left=bd_fina, right=bd_fina,
                        top=bd_fina, bottom=bd_fina
                    )

                    fill_cabecalho = PatternFill(
                        start_color="F2F2F2",
                        end_color="F2F2F2",
                        fill_type="solid"
                    )

                    # ========================================================
                    # COLUNAS
                    # ========================================================

                    ws.column_dimensions['A'].width = 12
                    ws.column_dimensions['B'].width = 18
                    ws.column_dimensions['C'].width = 45
                    ws.column_dimensions['D'].width = 18
                    ws.column_dimensions['E'].width = 25

                    # ========================================================
                    # CABEÇALHO SUPERIOR
                    # ========================================================

                    ws.merge_cells("A1:B3")
                    ws.merge_cells("C1:D3")
                    ws.merge_cells("E1:E3")

                    ws["C1"] = "Comprovante de Entrega de Material"
                    ws["C1"].font = Font(name="Arial", size=14, bold=True)
                    ws["C1"].alignment = Alignment(horizontal="center", vertical="center")

                    # ========================================================
                    # LOGOS
                    # ========================================================

                    if os.path.exists("Imagem1.png"):
                        img1 = OpenpyxlImage("Imagem1.png")
                        img1.width = 110
                        img1.height = 45
                        ws.add_image(img1, "A1")

                    if os.path.exists("Imagem2.png"):
                        img2 = OpenpyxlImage("Imagem2.png")
                        img2.width = 110
                        img2.height = 45
                        ws.add_image(img2, "E1")

                    # ========================================================
                    # DADOS CABEÇALHO
                    # ========================================================

                    primeiro_item = lista_liberacao[0]['item']

                    ws["A4"] = "Nº:"
                    ws["B4"] = str(op_selecionada)

                    ws["D4"] = "DIGITADO POR"
                    ws["E4"] = digitado_por.upper()

                    ws["A5"] = "Data:"
                    ws["B5"] = datetime.now().strftime('%d/%m/%Y %H:%M')

                    ws["A6"] = "Obra:"
                    ws["B6"] = str(primeiro_item.get('Obra', '')).upper()

                    ws["A7"] = "Nº Projeto:"
                    ws["B7"] = str(primeiro_item.get('Projeto', ''))

                    ws["A8"] = "Endereço da Obra:"
                    ws["B8"] = endereco_obra.upper()

                    # ========================================================
                    # FORMATAÇÃO CABEÇALHO
                    # ========================================================

                    for r in range(4, 9):

                        if r == 4:
                            ws.merge_cells("B4:C4")
                        else:
                            ws.merge_cells(
                                start_row=r, start_column=2,
                                end_row=r, end_column=3
                            )

                        ws.cell(row=r, column=1).font = Font(name="Arial", size=10, bold=True)

                        for c in range(1, 6):
                            ws.cell(row=r, column=c).border = borda_padrao

                    # ========================================================
                    # TÍTULOS TABELA
                    # ========================================================

                    titulos = [
                        "QTD",
                        "COD",
                        "DESCRIÇÃO",
                        "CONF. INTERNA",
                        "OBSERVAÇÕES"
                    ]

                    for col_num, titulo in enumerate(titulos, 1):
                        celula = ws.cell(row=10, column=col_num)
                        celula.value = titulo
                        celula.font = Font(bold=True)
                        celula.fill = fill_cabecalho
                        celula.alignment = Alignment(horizontal="center", vertical="center")
                        celula.border = borda_padrao

                    # ========================================================
                    # ITENS
                    # ========================================================

                    linha_excel = 11

                    for lib in lista_liberacao:

                        item = lib["item"]

                        ws.cell(linha_excel, 1, lib["qtd_saida"])
                        ws.cell(linha_excel, 2, item["Tipo_Cod"])
                        ws.cell(linha_excel, 3, item["Descricao"])
                        ws.cell(linha_excel, 4, item["Medida"])

                        for col in range(1, 6):
                            ws.cell(linha_excel, col).border = borda_padrao

                        linha_excel += 1

                    # ============================================================
                    # ASSINATURAS
                    # ============================================================

                    linha_ass = linha_excel + 3

                    # Aviso superior
                    ws.merge_cells(
                        start_row=linha_ass, start_column=1,
                        end_row=linha_ass, end_column=5
                    )
                    ws.cell(row=linha_ass, column=1).value = (
                        "Favor conferir todos os termos descritos neste romaneio antes de assinar. "
                        "Verificar se os ACM estão em perfeito estado."
                    )
                    ws.cell(row=linha_ass, column=1).font = Font(name="Arial", size=9, italic=True)

                    # Segunda linha de aviso
                    linha_ass += 1
                    ws.merge_cells(
                        start_row=linha_ass, start_column=1,
                        end_row=linha_ass, end_column=5
                    )
                    ws.cell(row=linha_ass, column=1).value = (
                        "Não serão aceitas reclamações após recebimento da mercadoria."
                    )
                    ws.cell(row=linha_ass, column=1).font = Font(name="Arial", size=9, italic=True)

                    # Texto de recebimento
                    linha_ass += 3
                    ws.merge_cells(
                        start_row=linha_ass, start_column=1,
                        end_row=linha_ass, end_column=5
                    )
                    ws.cell(row=linha_ass, column=1).value = (
                        "Recebi da Fachadas Passold as mercadorias acima relacionadas."
                    )
                    ws.cell(row=linha_ass, column=1).font = Font(name="Arial", size=10)

                    # Conferência interna
                    linha_ass += 2
                    ws.cell(row=linha_ass, column=1).value = "Conferência Interna:"
                    ws.cell(row=linha_ass, column=1).font = Font(bold=True)

                    for col in range(1, 4):
                        ws.cell(row=linha_ass + 1, column=col).border = Border(
                            bottom=Side(style='thin')
                        )
                    ws.cell(row=linha_ass + 1, column=4).value = "____/____/______"

                    # Motorista
                    linha_motorista = linha_ass + 4
                    for col in range(1, 4):
                        ws.cell(row=linha_motorista, column=col).border = Border(
                            bottom=Side(style='thin')
                        )
                    ws.cell(row=linha_motorista, column=4).value = "____/____/______"
                    ws.cell(row=linha_motorista + 1, column=1).value = "Nome Motorista (Data)"
                    ws.cell(row=linha_motorista + 1, column=1).font = Font(bold=True)

                    # Recebedor na obra
                    linha_receb = linha_motorista + 4
                    for col in range(1, 4):
                        ws.cell(row=linha_receb, column=col).border = Border(
                            bottom=Side(style='thin')
                        )
                    ws.cell(row=linha_receb, column=4).value = "____/____/______"
                    ws.cell(row=linha_receb + 1, column=1).value = "Nome Recebedor Obra (Data)"
                    ws.cell(row=linha_receb + 1, column=1).font = Font(bold=True)

                    # Líder
                    linha_lider = linha_receb + 4
                    for col in range(1, 4):
                        ws.cell(row=linha_lider, column=col).border = Border(
                            bottom=Side(style='thin')
                        )
                    ws.cell(row=linha_lider, column=4).value = "____/____/______"
                    ws.cell(row=linha_lider + 1, column=1).value = "Nome Líder (Data)"
                    ws.cell(row=linha_lider + 1, column=1).font = Font(bold=True)

                    # ============================================================
                    # SALVAR E DISPONIBILIZAR
                    # ============================================================

                    buffer = BytesIO()
                    wb.save(buffer)
                    buffer.seek(0)

                    st.success("Romaneio gerado com sucesso.")

                    st.download_button(
                        label="⬇️ Baixar Romaneio",
                        data=buffer,
                        file_name=f"Romaneio_OP_{op_selecionada}.xlsx",
                        mime=(
                            "application/vnd.openxmlformats-"
                            "officedocument.spreadsheetml.sheet"
                        )
                    )

# ============================================================
# ABA 3
# ============================================================

with aba3:

    st.subheader("Painel Geral de Saldos")

    if os.path.exists(BANCO_DADOS):

        try:
            df_ver = pd.read_excel(BANCO_DADOS)
            st.dataframe(df_ver, use_container_width=True)
        except:
            st.error("Erro ao abrir banco.")