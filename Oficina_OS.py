import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QGroupBox, QLabel, QLineEdit, QTextEdit, QPushButton,
    QListWidget, QMessageBox, QFileDialog, QSizePolicy, QComboBox,
    QStyle,  # Importado QStyle para usar ícones padrão do sistema
    QScrollArea
)
from PyQt5.QtGui import QFont, QPainter, QPageLayout, QPageSize, QTextOption, QPixmap, QDoubleValidator, QIntValidator
from PyQt5.QtCore import Qt, QDateTime, QRectF, QSizeF, QPointF
import os
import requests
import subprocess
import platform
import base64
from PIL import Image  # Importado Pillow para redimensionamento
import tempfile  # Importado para criar arquivos temporários
import atexit  # Para garantir a limpeza de arquivos temporários

# --- IMPORTS PARA JINJA2 E WEASYPRINT ---
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML, CSS

# --- FIM DOS IMPORTS ---

# --- Configurações Globais ---
ARQUIVO_EXCEL = "Ordens_de_Servico.xlsx"
ARQUIVO_LOGO = os.path.join("resources", "logo.png")  # Caminho ajustado para pasta resources
PASTA_OS_CLIENTES = "OS_Clientes"  # Esta pasta ainda é usada para a depuração, mas não para salvar PDFs
HTML_TEMPLATE_FILE = "os_template.html"

INFO_OFICINA = {
    "nome": "DEMONSTRAÇÃO OFICINA v6.0 230117 011216",
    "endereco": "Estrada do barro vermelho 341 - Rocha Miranda - RIO DEJANEIRO-RJ",
    "cnpj": "CNPJ 48.969.894/0001-59",
    "telefone": "(21) 99757-0103 / 97125-0490"
}

# Lista global para manter referências a arquivos temporários para limpeza
_temp_files_to_clean = []


def _cleanup_temp_files():
    """Função para limpar arquivos temporários ao sair do programa."""
    for temp_file_path in _temp_files_to_clean:
        try:
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)
                print(f"DEBUG: Arquivo temporário removido: {temp_file_path}", file=sys.stderr)
        except Exception as e:
            print(f"ERRO: Não foi possível remover o arquivo temporário {temp_file_path}: {e}", file=sys.stderr)


# Registrar a função de limpeza para ser chamada ao final do programa
atexit.register(_cleanup_temp_files)


class OficinaOSApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gerador de Ordem de Serviço - Oficina")
        self.setGeometry(100, 100, 1200, 800)  # Aumenta a largura inicial para 3 colunas
        self.setMinimumSize(1000, 750)  # Aumenta o mínimo para 3 colunas caberem

        # --- CSS BÁSICO INLINE TOTALMENTE COMENTADO PARA ALTERAR MANUALMENTE ---
        # Removido self.setStyleSheet() para estilos mínimos
        # Agora, apenas os estilos essenciais para visibilidade e estrutura.
        # TODAS AS REGRAS ABAIXO ESTÃO COMENTADAS. DESCOMENTE E EDITE CONFORME SUA NECESSIDADE.
        self.setStyleSheet("""
            /*
            QWidget {
                background-color: #f8f8f8; /* Fundo geral muito claro */
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 10pt;
                color: #333333;
            }
            QGroupBox {
                border: 1px solid #e0e0e0;
                border-radius: 6px;
                margin-top: 12px;
                padding: 8px;
                background-color: #ffffff;
                box-shadow: none;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 6px;
                color: #007bff;
                font-weight: bold;
                font-size: 11pt;
            }
            QLabel {
                color: #555555;
                font-weight: normal;
                padding: 1px 0;
            }
            QLineEdit, QTextEdit, QComboBox {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 5px 7px;
                background-color: #ffffff;
                color: #000000; /* CRÍTICO: Cor do texto digitado PRETA para garantir visibilidade */
            }
            QLineEdit::placeholder, QTextEdit::placeholder {
                color: #a0a0a0;
            }
            QLineEdit:hover, QTextEdit:hover, QComboBox:hover {
                border: 1px solid #b0b0b0;
            }
            QLineEdit:focus, QTextEdit:focus, QComboBox:focus {
                border: 1px solid #007bff;
                outline: none;
            }
            QPushButton {
                background-color: #007bff;
                color: #ffffff;
                border: none;
                padding: 7px 14px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QPushButton:pressed {
                background-color: #004085;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
            QPushButton#btnSalvar { background-color: #28a745; }
            QPushButton#btnSalvar:hover { background-color: #218838; }
            QPushButton#btnLimpar { background-color: #6c757d; }
            QPushButton#btnLimpar:hover { background-color: #5a6268; }
            QPushButton#btnDeletar { background-color: #dc3545; }
            QPushButton#btnDeletar:hover { background-color: #c82333; }
            QPushButton#btnGerarPDF { background-color: #17a2b8; }
            QPushButton#btnGerarPDF:hover { background-color: #138496; }
            QPushButton#btnSair { background-color: #6f42c1; }
            QPushButton#btnSair:hover { background-color: #563d7c; }

            QStatusBar {
                background-color: #e0e0e0;
                color: #333333;
                font-size: 9pt;
                padding: 3px;
            }
            QListWidget {
                border: 1px solid #c0c0c0;
                border-radius: 5px;
                background-color: #ffffff;
                color: #222222;
                padding: 5px;
            }
        """)

        try:
            os.makedirs(PASTA_OS_CLIENTES, exist_ok=True)
            print(f"Pasta '{PASTA_OS_CLIENTES}' verificada/criada com sucesso.")
        except Exception as e:
            QMessageBox.critical(self, "Erro de Pasta",
                                 f"Não foi possível criar a pasta '{PASTA_OS_CLIENTES}': {e}\nVerifique as permissões.")
            print(f"Erro ao criar pasta: {e}", file=sys.stderr)

        self.df_os = self._carregar_dados_os()
        self.itens_pecas_servicos_cache = []

        self.env = Environment(loader=FileSystemLoader('.'))
        self.env.filters['format_money'] = self._format_money_filter
        self.env.filters['km_format'] = self._km_format_filter
        self.env.filters['default_if_nan'] = self._default_if_nan_filter

        self._criar_interface()
        self._gerar_novo_id_os()

    # --- Filtros Jinja2 ---
    def _format_money_filter(self, value):
        try:
            if pd.isna(value) or value is None:
                value = 0.0
            val = float(value)
            return f"{val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        except (ValueError, TypeError):
            return f"{0.00:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

    def _km_format_filter(self, value):
        try:
            if pd.isna(value) or value is None:
                value = ""
            clean_text = ''.join(filter(str.isdigit, str(value)))
            if clean_text:
                return f"{int(clean_text):,}".replace(',', '.')
            return ""
        except (ValueError, TypeError):
            return ""

    def _default_if_nan_filter(self, value):
        if pd.isna(value) or value is None:
            return ""
        return str(value)

    # --- Fim dos Filtros Jinja2 ---

    def _carregar_dados_os(self):
        """Carrega os dados das OSs do arquivo Excel ou cria e inicializa um DataFrame vazio."""
        if os.path.exists(ARQUIVO_EXCEL):
            try:
                converters = {
                    'Numero_OS': str,
                    'KM_Atual_Veiculo': lambda x: int(str(x).replace('.', '').replace(',', '')) if str(x).replace('.',
                                                                                                                  '').replace(
                        ',', '').isdigit() else pd.NA,
                    'Total_Itens': lambda x: float(str(x).replace('.', '').replace(',', '.')) if isinstance(x,
                                                                                                            str) and ',' in x else float(
                        x) if x is not None else pd.NA,
                    'Valor_Total_Final': lambda x: float(str(x).replace('.', '').replace(',', '.')) if isinstance(x,
                                                                                                                  str) and ',' in x else float(
                        x) if x is not None else pd.NA,
                    'Deslocamento': lambda x: float(str(x).replace('.', '').replace(',', '.')) if isinstance(x,
                                                                                                             str) and ',' in x else float(
                        x) if x is not None else pd.NA,
                    'Desconto_Geral': lambda x: float(str(x).replace('.', '').replace(',', '.')) if isinstance(x,
                                                                                                               str) and ',' in x else float(
                        x) if x is not None else pd.NA,
                }
                df = pd.read_excel(ARQUIVO_EXCEL, converters=converters)
                print(f"Dados carregados de {ARQUIVO_EXCEL}")

                # CORREÇÃO PARA GARANTIR QUE Numero_OS SEJA STRING E STRIPADO AO CARREGAR
                df['Numero_OS'] = df['Numero_OS'].astype(str).str.strip()

                colunas_esperadas = self._get_expected_columns()
                for col in colunas_esperadas:
                    if col not in df.columns:
                        df[col] = pd.NA
                return df
            except Exception as e:
                QMessageBox.critical(self, "Erro de Leitura",
                                     f"Erro ao carregar o arquivo Excel: {e}\nUm novo arquivo será criado e o erro será gravado no console.")
                print(f"Detalhes do erro ao carregar Excel: {e}", file=sys.stderr)
                return self._criar_dataframe_vazio_e_salvar()
        else:
            print(f"Arquivo {ARQUIVO_EXCEL} não encontrado. Criando e inicializando novo DataFrame.")
            return self._criar_dataframe_vazio_e_salvar()

    def _get_expected_columns(self):
        return [
            "Numero_OS", "Data_OS", "Hora_OS",
            "Nome_Cliente", "Endereco_Cliente", "Numero_Imovel_Cliente", "Bairro_Cliente", "Cidade_Cliente",
            "UF_Cliente", "CEP_Cliente", "Telefone_Cliente", "CPF_CNPJ_Cliente",
            "Placa_Veiculo", "Marca_Veiculo", "Modelo_Veiculo", "Cor_Veiculo", "Ano_Veiculo", "KM_Atual_Veiculo",
            "Combustivel_Veiculo", "Box_Veiculo",
            "Problema_Informado", "Problema_Constatado", "Servico_Executado",
            "Detalhes_Itens",
            "Total_Itens",
            "Deslocamento", "Desconto_Geral", "Valor_Total_Final",
            "Responsavel", "Situacao_Atual",
            "Condicoes_Pagamento"
        ]

    def _criar_dataframe_vazio_e_salvar(self):
        colunas = self._get_expected_columns()
        df = pd.DataFrame(columns=colunas)
        try:
            df.to_excel(ARQUIVO_EXCEL, index=False)
            print(f"Archivo Excel vacío '{ARQUIVO_EXCEL}' creado con éxito.")
        except Exception as e:
            QMessageBox.critical(self, "Erro de Escritura",
                                 f"No fue posible crear el archivo Excel vacío: {e}. Verifique los permisos de la carpeta.")
            print(f"Detalles del error al crear Excel vacío: {e}", file=sys.stderr)
        return df

    def _gerar_novo_id_os(self):
        if not self.df_os.empty:
            numeros_validos = self.df_os['Numero_OS'].astype(str).apply(lambda x: ''.join(filter(str.isdigit, x)))
            ultimos_numeros = [int(n) for n in numeros_validos if n]
            if ultimos_numeros:
                novo_id = max(ultimos_numeros) + 1
            else:
                novo_id = 1
        else:
            novo_id = 1
        self.entry_numero_os.setText(str(novo_id).zfill(6))
        self.entry_numero_os.setReadOnly(True)

    def _limpar_campos(self):
        for entry in self.findChildren(QLineEdit):
            entry.clear()
        for text_edit in self.findChildren(QTextEdit):
            text_edit.clear()

        # Define o índice 0 para o placeholder ""
        self.combo_situacao_atual.setCurrentIndex(0)
        self.combo_condicoes_pagamento.setCurrentIndex(0)

        self.listbox_itens.clear()
        self.itens_pecas_servicos_cache = []

        self.label_total_itens.setText("Total Itens: R$ 0,00")
        self.label_valor_total.setText("Valor Total: R$ 0,00")
        self.label_data.setText(QDateTime.currentDateTime().toString("dd/MM/yyyy hh:mm:ss"))
        self._gerar_novo_id_os()

    def _criar_interface(self):
        main_layout = QVBoxLayout(self)

        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)

        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_widget.setMinimumSize(800, 1000)

        scroll_area.setWidget(content_widget)
        main_layout.addWidget(scroll_area, 1)

        # --- TOP SECTION: Logo, Info Oficina, Dados OS/Busca (Horizontal) ---
        header_layout = QHBoxLayout()
        content_layout.addLayout(header_layout)
        content_layout.setStretchFactor(header_layout, 0)

        logo_info_layout = QVBoxLayout()
        header_layout.addLayout(logo_info_layout)
        header_layout.setStretchFactor(logo_info_layout, 0)

        self.logo_label = QLabel()
        if os.path.exists(ARQUIVO_LOGO):
            pixmap = QPixmap(ARQUIVO_LOGO)
            if not pixmap.isNull():
                self.logo_label.setPixmap(pixmap.scaled(80, 80, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            else:
                self.logo_label.setText("LOGO")
        else:
            self.logo_label.setText("LOGO")
        logo_info_layout.addWidget(self.logo_label, alignment=Qt.AlignTop | Qt.AlignLeft)

        oficina_info_label = QLabel(
            f"<b>{INFO_OFICINA['nome']}</b><br>"
            f"{INFO_OFICINA['endereco']}<br>"
            f"CNPJ: {INFO_OFICINA['cnpj']} | Tel: {INFO_OFICINA['telefone']}"
        )
        oficina_info_label.setTextFormat(Qt.RichText)
        oficina_info_label.setFont(QFont("Arial", 8))
        logo_info_layout.addWidget(oficina_info_label, alignment=Qt.AlignTop | Qt.AlignLeft)
        logo_info_layout.addStretch(1)

        os_info_group = QGroupBox("Dados da OS")
        os_info_layout = QGridLayout()
        os_info_group.setLayout(os_info_layout)
        header_layout.addWidget(os_info_group)
        header_layout.setStretchFactor(os_info_group, 1)

        os_info_layout.addWidget(QLabel("Número da OS:"), 0, 0, Qt.AlignLeft)
        self.entry_numero_os = QLineEdit()
        self.entry_numero_os.setFixedWidth(80)
        os_info_layout.addWidget(self.entry_numero_os, 0, 1, Qt.AlignLeft)
        os_info_layout.setColumnStretch(1, 0)

        os_info_layout.addWidget(QLabel("Data:"), 0, 2, Qt.AlignLeft)
        self.label_data = QLabel(QDateTime.currentDateTime().toString("dd/MM/yyyy hh:mm:ss"))
        os_info_layout.addWidget(self.label_data, 0, 3, Qt.AlignLeft)
        os_info_layout.setColumnStretch(3, 0)

        os_info_layout.addWidget(QLabel("Buscar OS por ID:"), 1, 0, Qt.AlignLeft)
        self.entry_busca_os = QLineEdit()
        self.entry_busca_os.setFixedWidth(80)
        os_info_layout.addWidget(self.entry_busca_os, 1, 1, Qt.AlignLeft)

        btn_buscar = QPushButton("Buscar")
        btn_buscar.clicked.connect(self._buscar_os)
        btn_buscar.setIcon(self.style().standardIcon(QStyle.SP_FileDialogToParent))
        os_info_layout.addWidget(btn_buscar, 1, 2, 1, 2, Qt.AlignLeft)
        os_info_layout.setColumnStretch(1, 1)

        # --- Layout Horizontal para Dados do Cliente e Dados do Veículo ---
        main_content_top_horizontal_layout = QHBoxLayout()
        content_layout.addLayout(main_content_top_horizontal_layout)
        content_layout.setStretchFactor(main_content_top_horizontal_layout, 2)

        # --- Grupo Dados do Cliente (Grid com 2 colunas de campos, conforme solicitado) ---
        cliente_group = QGroupBox("Dados do Cliente")
        cliente_layout = QGridLayout()
        cliente_group.setLayout(cliente_layout)
        main_content_top_horizontal_layout.addWidget(cliente_group, 1)

        # Configurar ícone para Dados do Cliente
        cliente_group.setTitle("👤 Dados do Cliente")  # Usando emoji Unicode para o ícone.

        self.entries_cliente = {}
        client_field_positions = {
            "Nome": (0, 0), "Endereço": (1, 0), "Número": (2, 0), "Bairro": (3, 0), "Cidade": (4, 0),
            "UF": (0, 2), "CEP": (1, 2), "Telefone": (2, 2), "CPF/Cnpj": (3, 2)
        }
        campos_cliente_display_order = ["Nome", "Endereço", "Número", "Bairro", "Cidade", "UF", "CEP", "Telefone",
                                        "CPF/Cnpj"]

        for campo_display_name in campos_cliente_display_order:
            row, col_start = client_field_positions[campo_display_name]
            field_name_internal = campo_display_name.lower().replace('/', '_')

            cliente_layout.addWidget(QLabel(f"{campo_display_name}:"), row, col_start, Qt.AlignLeft)
            entry = QLineEdit()
            entry.setPlaceholderText(f"Digite o {campo_display_name.lower()}...")
            entry.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            entry.setMinimumWidth(150)
            self.entries_cliente[field_name_internal] = entry
            cliente_layout.addWidget(entry, row, col_start + 1, Qt.AlignLeft)

            # --- VALIDADORES NUMÉRICOS (SEM inputMask) ---
            if field_name_internal == "telefone":
                entry.setValidator(QIntValidator())
                # Adiciona formatação para Telefone ao digitar
                entry.textChanged.connect(lambda text, e=entry: self._formatar_telefone_cpf_cnpj(e, "telefone"))
            elif field_name_internal == "cpf_cnpj":
                entry.setValidator(QIntValidator())
                # Adiciona formatação para CPF/CNPJ ao digitar
                entry.textChanged.connect(lambda text, e=entry: self._formatar_telefone_cpf_cnpj(e, "cpf_cnpj"))
            elif field_name_internal == "cep":
                entry.setValidator(QIntValidator())
                entry.editingFinished.connect(self._autopreencher_cep)
            elif field_name_internal == "número":
                entry.setValidator(QIntValidator())
            # --- FIM DOS VALIDADORES ---

        cliente_layout.setColumnStretch(1, 1)
        cliente_layout.setColumnStretch(3, 1)

        # Grupo Dados do Veículo
        veiculo_group = QGroupBox("Dados do Veículo")
        veiculo_layout = QGridLayout()
        veiculo_group.setLayout(veiculo_layout)
        main_content_top_horizontal_layout.addWidget(veiculo_group, 1)  # Adiciona ao layout horizontal

        # Configurar ícone para Dados do Veículo
        veiculo_group.setTitle("🚗 Dados do Veículo")  # Usando emoji Unicode para o ícone.

        self.entries_veiculo = {}
        vehicle_field_positions = {
            "Placa": (0, 0), "Ano": (0, 2),
            "Marca": (1, 0), "KM Atual": (1, 2),
            "Modelo": (2, 0), "Combustível": (2, 2),
            "Cor": (3, 0), "Box": (3, 2)
        }
        campos_veiculo_display_order = ["Placa", "Marca", "Modelo", "Cor", "Ano", "KM Atual", "Combustível", "Box"]

        for campo_display_name in campos_veiculo_display_order:
            row, col_start = vehicle_field_positions[campo_display_name]
            field_name_internal = campo_display_name.lower().replace(' ', '_')

            veiculo_layout.addWidget(QLabel(f"{campo_display_name}:"), row, col_start, Qt.AlignLeft)
            entry = QLineEdit()
            entry.setPlaceholderText(f"Digite o {campo_display_name.lower()}...")
            entry.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
            entry.setMinimumWidth(100)
            self.entries_veiculo[field_name_internal] = entry
            veiculo_layout.addWidget(entry, row, col_start + 1, Qt.AlignLeft)

            if field_name_internal == "ano":
                entry.setValidator(QIntValidator(1900, QDateTime.currentDateTime().date().year() + 5))
            elif field_name_internal == "km_atual":
                entry.setValidator(QIntValidator(0, 999999999))
                entry.textChanged.connect(self._formatar_quilometragem)
            # Adiciona as opções para Combustível e Box
            elif field_name_internal == "combustível":
                combo = QComboBox()
                combo.setPlaceholderText("Selecione...")
                combo.addItems(["", "Gasolina", "Etanol", "Flex", "Diesel", "GNV", "Elétrico", "Híbrido"])  #
                veiculo_layout.removeWidget(entry)  # Remove o QLineEdit temporário
                entry.deleteLater()  # Deleta o QLineEdit
                self.entries_veiculo[field_name_internal] = combo  # Substitui pela referência do combobox
                veiculo_layout.addWidget(combo, row, col_start + 1, Qt.AlignLeft)
            elif field_name_internal == "box":
                combo = QComboBox()
                combo.setPlaceholderText("Selecione...")
                combo.addItems(["", "Box 1", "Box 2", "Box 3", "Box 4", "Pátio"])  #
                veiculo_layout.removeWidget(entry)  # Remove o QLineEdit temporário
                entry.deleteLater()  # Deleta o QLineEdit
                self.entries_veiculo[field_name_internal] = combo  # Substitui pela referência do combobox
                veiculo_layout.addWidget(combo, row, col_start + 1, Qt.AlignLeft)

        veiculo_layout.setColumnStretch(1, 1)
        veiculo_layout.setColumnStretch(3, 1)

        # --- Grupo Problemas e Serviços - Campos LADO A LADO (3 colunas) ---
        problemas_group = QGroupBox("Problemas e Serviços")
        problemas_layout = QGridLayout()  # Usar QGridLayout para 3 colunas lado a lado
        problemas_group.setLayout(problemas_layout)
        content_layout.addWidget(problemas_group)  # Adiciona diretamente ao content_layout, abaixo dos outros grupos
        content_layout.setStretchFactor(problemas_group, 2)  # Dá mais peso para expandir verticalmente

        # Configurar ícone para Problemas e Serviços
        problemas_group.setTitle("🛠️ Problemas e Serviços")  # Usando emoji Unicode para o ícone.

        # Problema Informado (colunas 0 e 1)
        problemas_layout.addWidget(QLabel("Problema Informado:"), 0, 0, Qt.AlignTop | Qt.AlignLeft)
        self.text_problema_informado = QTextEdit()
        self.text_problema_informado.setPlaceholderText("Descreva o problema informado pelo cliente...")
        self.text_problema_informado.setMinimumHeight(60)
        self.text_problema_informado.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        problemas_layout.addWidget(self.text_problema_informado, 1, 0, 1, 2)  # Ocupa a linha 1, colunas 0 e 1

        # Problema Constatado (colunas 2 e 3)
        problemas_layout.addWidget(QLabel("Problema Constatado:"), 0, 2, Qt.AlignTop | Qt.AlignLeft)
        self.text_problema_constatado = QTextEdit()
        self.text_problema_constatado.setPlaceholderText("Descreva o problema constatado após a inspeção...")
        self.text_problema_constatado.setMinimumHeight(60)
        self.text_problema_constatado.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        problemas_layout.addWidget(self.text_problema_constatado, 1, 2, 1, 2)  # Ocupa a linha 1, colunas 2 e 3

        # Serviço Executado (colunas 4 e 5)
        problemas_layout.addWidget(QLabel("Serviço Executado:"), 0, 4, Qt.AlignTop | Qt.AlignLeft)
        self.text_servico_executado = QTextEdit()
        self.text_servico_executado.setPlaceholderText("Detalhe os serviços realizados...")
        self.text_servico_executado.setMinimumHeight(60)
        self.text_servico_executado.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        problemas_layout.addWidget(self.text_servico_executado, 1, 4, 1, 2)  # Ocupa a linha 1, colunas 4 e 5

        # Definir stretch para as colunas dos QTextEdit
        problemas_layout.setColumnStretch(1, 1)  # Campo 1
        problemas_layout.setColumnStretch(3, 1)  # Campo 2
        problemas_layout.setColumnStretch(5, 1)  # Campo 3

        # --- Seção de Itens (Peças e Serviços) ---
        itens_group = QGroupBox("Itens (Peças e Serviços)")
        itens_layout = QGridLayout()
        itens_group.setLayout(itens_layout)
        content_layout.addWidget(itens_group)
        content_layout.setStretchFactor(itens_group, 1)  # Dá um peso para expandir verticalmente (lista de itens)

        # Adicionar campo "Tipo" (ComboBox)
        itens_layout.addWidget(QLabel("Tipo:"), 0, 0, Qt.AlignLeft)
        self.combo_item_tipo = QComboBox()
        self.combo_item_tipo.setPlaceholderText("Selecione...")
        self.combo_item_tipo.addItems(["Mão de obra", "Peça", "Serviço"])  # Opções do tipo de item
        itens_layout.addWidget(self.combo_item_tipo, 0, 1, Qt.AlignLeft)
        itens_layout.setColumnStretch(1, 0)  # Coluna fixa para o tipo

        itens_layout.addWidget(QLabel("Ref:"), 0, 2, Qt.AlignLeft)  # Ajustar para col 2
        self.entry_item_ref = QLineEdit()
        self.entry_item_ref.setPlaceholderText("Código/Ref")  # Alterado conforme imagem
        self.entry_item_ref.setFixedWidth(80)  # Ajustado para 80px
        itens_layout.addWidget(self.entry_item_ref, 0, 3, Qt.AlignLeft)
        itens_layout.setColumnStretch(3, 0)  # Não expande esta coluna (fixa)

        itens_layout.addWidget(QLabel("Desc:"), 0, 4, Qt.AlignLeft)  # Ajustar para col 4
        self.entry_item_desc = QLineEdit()
        self.entry_item_desc.setPlaceholderText("Descrição do Item/Serviço")  # Alterado conforme imagem
        itens_layout.addWidget(self.entry_item_desc, 0, 5, Qt.AlignLeft)
        itens_layout.setColumnStretch(5, 1)  # Permite expansão da Descrição

        itens_layout.addWidget(QLabel("Val Unit:"), 0, 6, Qt.AlignLeft)  # Ajustar para col 6
        self.entry_item_valor = QLineEdit()
        self.entry_item_valor.setPlaceholderText("0,00")
        itens_layout.addWidget(self.entry_item_valor, 0, 7, Qt.AlignLeft)
        itens_layout.setColumnStretch(7, 0)  # Não expande

        itens_layout.addWidget(QLabel("Qtd:"), 0, 8, Qt.AlignLeft)  # Ajustar para col 8
        self.entry_item_qtd = QLineEdit()
        self.entry_item_qtd.setPlaceholderText("1")  # Padrão para 1
        itens_layout.addWidget(self.entry_item_qtd, 0, 9, Qt.AlignLeft)
        itens_layout.setColumnStretch(9, 0)

        itens_layout.addWidget(QLabel("Desc. (%):"), 0, 10, Qt.AlignLeft)  # Novo campo Desconto
        self.entry_item_desc_perc = QLineEdit()
        self.entry_item_desc_perc.setPlaceholderText("0")
        self.entry_item_desc_perc.setValidator(QIntValidator(0, 100))  # Desconto de 0 a 100%
        itens_layout.addWidget(self.entry_item_desc_perc, 0, 11, Qt.AlignLeft)
        itens_layout.setColumnStretch(11, 0)

        btn_add_item = QPushButton("Adicionar Serviço")  # Alterado conforme imagem (Adicionar Serviço)
        btn_add_item.clicked.connect(self._adicionar_item)
        btn_add_item.setIcon(self.style().standardIcon(QStyle.SP_DialogApplyButton))
        itens_layout.addWidget(btn_add_item, 0, 12, 1, 2, Qt.AlignLeft)  # Ocupa 2 colunas para o botão
        itens_layout.setColumnStretch(12, 0)

        self.listbox_itens = QListWidget()
        # Ajustar para ocupar a largura total do grid de itens
        itens_layout.addWidget(self.listbox_itens, 1, 0, 1, 14)  # Ajusta colspan para todas as colunas
        itens_layout.setRowStretch(1, 1)
        btn_rem_item = QPushButton("Remover Item Selecionado")
        btn_rem_item.setIcon(self.style().standardIcon(QStyle.SP_DialogCancelButton))
        btn_rem_item.clicked.connect(self._remover_item)
        itens_layout.addWidget(btn_rem_item, 2, 0, 1, 14)  # Ajusta colspan

        # --- Seção de Totais e Finais ---
        finais_group = QGroupBox("Totais e Finais")
        finais_layout = QGridLayout()
        finais_group.setLayout(finais_layout)
        content_layout.addWidget(finais_group)
        content_layout.setStretchFactor(finais_group, 0)

        self.entries_finais = {}

        finais_layout.addWidget(QLabel("Responsável:"), 0, 0, Qt.AlignLeft)
        entry_responsavel = QLineEdit()
        entry_responsavel.setPlaceholderText("Nome do responsável")
        entry_responsavel.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.entries_finais["responsável"] = entry_responsavel
        finais_layout.addWidget(entry_responsavel, 0, 1, Qt.AlignLeft)

        finais_layout.addWidget(QLabel("Situação Atual:"), 1, 0, Qt.AlignLeft)
        self.combo_situacao_atual = QComboBox()
        self.combo_situacao_atual.setPlaceholderText("Selecione a situação")
        # Adiciona opções da imagem dd99c1.png
        self.combo_situacao_atual.addItems(
            ["", "Orçamento", "Aprovado", "Em Andamento", "Aguardando Peças", "Finalizado", "Entregue"])  #
        finais_layout.addWidget(self.combo_situacao_atual, 1, 1, Qt.AlignLeft)
        self.entries_finais["situação_atual"] = self.combo_situacao_atual

        finais_layout.addWidget(QLabel("Condições de Pagamento:"), 2, 0, Qt.AlignLeft)
        self.combo_condicoes_pagamento = QComboBox()
        self.combo_condicoes_pagamento.setPlaceholderText("Selecione a condição")
        # Adiciona opções da imagem dd9d26.png
        self.combo_condicoes_pagamento.addItems(
            ["", "À Vista", "PIX", "Cartão Crédito", "Cartão Débito", "Dinheiro", "Boleto", "Parcelado"])  #
        finais_layout.addWidget(self.combo_condicoes_pagamento, 2, 1, Qt.AlignLeft)
        self.entries_finais["condições_de_pagamento"] = self.combo_condicoes_pagamento

        finais_layout.setColumnStretch(1, 1)

        self.label_total_itens = QLabel("Total Itens: R$ 0,00")
        self.label_total_itens.setFont(QFont("Arial", 10, QFont.Bold))
        finais_layout.addWidget(self.label_total_itens, 0, 2, Qt.AlignRight)

        self.label_valor_total = QLabel("Valor Total: R$ 0,00")
        self.label_valor_total.setFont(QFont("Arial", 12, QFont.Bold))
        finais_layout.addWidget(self.label_valor_total, 1, 2, Qt.AlignRight)
        finais_layout.setColumnStretch(2, 1)

        # --- Seção de Botões de Ação (fixa na parte inferior, fora do scroll) ---
        button_layout = QHBoxLayout()
        self.layout().addLayout(button_layout)
        self.layout().setStretchFactor(button_layout, 0)

        btn_salvar = QPushButton("Salvar/Atualizar OS")
        btn_salvar.clicked.connect(self._salvar_os)
        btn_salvar.setObjectName("btnSalvar")
        btn_salvar.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        button_layout.addWidget(btn_salvar)

        btn_limpar = QPushButton("Limpar Campos")
        btn_limpar.clicked.connect(self._limpar_campos)
        btn_limpar.setObjectName("btnLimpar")
        btn_limpar.setIcon(self.style().standardIcon(QStyle.SP_DialogResetButton))
        button_layout.addWidget(btn_limpar)

        btn_deletar = QPushButton("Deletar OS Atual")
        btn_deletar.clicked.connect(self._deletar_os)
        btn_deletar.setObjectName("btnDeletar")
        btn_deletar.setIcon(self.style().standardIcon(QStyle.SP_TrashIcon))
        button_layout.addWidget(btn_deletar)

        btn_imprimir = QPushButton("Gerar e Visualizar OS (PDF)")
        btn_imprimir.clicked.connect(self._imprimir_os_pdf)
        btn_imprimir.setObjectName("btnGerarPDF")
        btn_imprimir.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))
        button_layout.addWidget(btn_imprimir)

        btn_sair = QPushButton("Sair")
        btn_sair.clicked.connect(self.close)
        btn_sair.setObjectName("btnSair")
        button_layout.addWidget(btn_sair)

    def _formatar_quilometragem(self, text):
        cursor_pos = self.entries_veiculo["km_atual"].cursorPosition()
        original_text_len = len(text)

        clean_text = ''.join(filter(str.isdigit, text))

        if not clean_text:
            self.entries_veiculo["km_atual"].setText("")
            return

        try:
            formatted_km = f"{int(clean_text):,}".replace(',', '.')
            self.entries_veiculo["km_atual"].setText(formatted_km)

            new_text_len = len(formatted_km)
            len_diff = new_text_len - original_text_len
            self.entries_veiculo["km_atual"].setCursorPosition(cursor_pos + len_diff)

        except ValueError:
            pass

    def _formatar_telefone_cpf_cnpj(self, entry_widget, field_type):
        current_text = entry_widget.text()
        clean_text = ''.join(filter(str.isdigit, current_text))
        formatted_text = ""
        cursor_pos = entry_widget.cursorPosition()
        len_diff = 0  # Diferença no comprimento para ajustar a posição do cursor

        if field_type == "telefone":
            if len(clean_text) > 11:  # Limita a 11 dígitos
                clean_text = clean_text[:11]

            if len(clean_text) > 2:
                formatted_text += f"({clean_text[:2]}) "
                if len(clean_text) > 7:
                    formatted_text += f"{clean_text[2:7]}-{clean_text[7:]}"
                else:
                    formatted_text += clean_text[2:]
            else:
                formatted_text = clean_text

            # Ajusta a posição do cursor
            if len(current_text) < len(formatted_text):
                len_diff = len(formatted_text) - len(current_text)
            elif len(current_text) > len(formatted_text):
                len_diff = -(len(current_text) - len(formatted_text))  # Diminuiu o texto

        elif field_type == "cpf_cnpj":
            if len(clean_text) > 14:  # Limita a 14 dígitos
                clean_text = clean_text[:14]

            if len(clean_text) <= 11:  # CPF
                if len(clean_text) > 9:
                    formatted_text = f"{clean_text[:3]}.{clean_text[3:6]}.{clean_text[6:9]}-{clean_text[9:]}"
                elif len(clean_text) > 6:
                    formatted_text = f"{clean_text[:3]}.{clean_text[3:6]}.{clean_text[6:]}"
                elif len(clean_text) > 3:
                    formatted_text = f"{clean_text[:3]}.{clean_text[3:]}"
                else:
                    formatted_text = clean_text
            else:  # CNPJ
                if len(clean_text) > 12:
                    formatted_text = f"{clean_text[:2]}.{clean_text[2:5]}.{clean_text[5:8]}/{clean_text[8:12]}-{clean_text[12:]}"
                elif len(clean_text) > 8:
                    formatted_text = f"{clean_text[:2]}.{clean_text[2:5]}.{clean_text[5:8]}/{clean_text[8:]}"
                elif len(clean_text) > 5:
                    formatted_text = f"{clean_text[:2]}.{clean_text[2:5]}.{clean_text[5:]}"
                elif len(clean_text) > 2:
                    formatted_text = f"{clean_text[:2]}.{clean_text[2:]}"
                else:
                    formatted_text = clean_text

            # Ajusta a posição do cursor
            if len(current_text) < len(formatted_text):
                len_diff = len(formatted_text) - len(current_text)
            elif len(current_text) > len(formatted_text):
                len_diff = -(len(current_text) - len(formatted_text))

        entry_widget.setText(formatted_text)
        # Ajusta a posição do cursor apenas se o texto não for o mesmo
        if current_text != formatted_text:
            entry_widget.setCursorPosition(cursor_pos + len_diff)

    def _formatar_valor_monetario(self):
        sender = self.sender()
        text = sender.text().strip().replace(',', '.')
        try:
            value = float(text)
            formatted_value = f"{value:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            sender.setText(formatted_value)
        except ValueError:
            sender.setText(f"{0.00:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
            QMessageBox.warning(self, "Formato Inválido", "Valor monetário inválido. Use apenas números.")

    def _adicionar_item(self):
        tipo = self.combo_item_tipo.currentText().strip()  # Novo campo "Tipo"
        referencia = self.entry_item_ref.text().strip()
        descricao = self.entry_item_desc.text().strip()
        valor_str = self.entry_item_valor.text().strip().replace(',', '.')
        qtd_str = self.entry_item_qtd.text().strip()
        desc_perc_str = self.entry_item_desc_perc.text().strip()  # Novo campo Desconto (%)

        if not tipo or not descricao or not valor_str or not qtd_str:
            QMessageBox.warning(self, "Entrada Inválida", "Por favor, preencha todos os campos do item.")
            return

        try:
            valor_unitario = float(valor_str)
            quantidade = int(qtd_str)
            desconto_percentual = float(desc_perc_str) if desc_perc_str else 0.0  # Converte desconto

            if valor_unitario <= 0 or quantidade <= 0:
                QMessageBox.warning(self, "Entrada Inválida", "Valor unitário e quantidade devem ser maiores que zero.")
                return
            if not (0 <= desconto_percentual <= 100):
                QMessageBox.warning(self, "Entrada Inválida", "Desconto percentual deve estar entre 0 e 100.")
                return

        except ValueError:
            QMessageBox.warning(self, "Entrada Inválida", "Valores numéricos inválidos. Use apenas números.")
            return

        valor_total_item_sem_desc = valor_unitario * quantidade
        valor_total_item = valor_total_item_sem_desc * (1 - (desconto_percentual / 100))

        item_data = {
            "tipo": tipo,  # Adicionado tipo
            "referencia": referencia,
            "descricao": descricao,
            "uni": "un",  # Unidade padrão
            "valor": valor_unitario,
            "quantia": quantidade,
            "desc": desconto_percentual,  # Salva o percentual
            "valor_total": valor_total_item
        }
        self.itens_pecas_servicos_cache.append(item_data)
        self.listbox_itens.addItem(
            f"Tipo: {tipo} | Ref: {referencia} - {descricao} | Qtde: {quantidade} x R${valor_unitario:.2f} | Desc: {desconto_percentual:.0f}% = R${valor_total_item:.2f}")

        self.combo_item_tipo.setCurrentIndex(0)  # Limpa o tipo
        self.entry_item_ref.clear()
        self.entry_item_desc.clear()
        self.entry_item_valor.setText("0,00")  # Reseta para 0,00
        self.entry_item_qtd.setText("1")  # Reseta para 1
        self.entry_item_desc_perc.setText("0")  # Reseta desconto para 0

        self._atualizar_totais()

    def _remover_item(self):
        try:
            selected_row = self.listbox_itens.currentRow()
            if selected_row != -1:
                self.listbox_itens.takeItem(selected_row)
                del self.itens_pecas_servicos_cache[selected_row]
                self._atualizar_totais()
            else:
                QMessageBox.warning(self, "Seleção Inválida", "Por favor, selecione um item para remover.")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Ocorreu um erro ao remover o item: {e}")

    def _atualizar_totais(self):
        total_itens = sum(item['valor_total'] for item in self.itens_pecas_servicos_cache)
        # Por enquanto, sem deslocamento ou desconto geral que não seja por item
        valor_total_final = total_itens

        self.label_total_itens.setText(
            f"Total Itens: R$ {total_itens:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        self.label_valor_total.setText(
            f"Valor Total: R$ {valor_total_final:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))

    def _autopreencher_cep(self):
        cep = self.entries_cliente["cep"].text().strip()  # Remove o hífen da máscara para buscar
        print(f"DEBUG: Autopreencher CEP chamado para: '{cep}'", file=sys.stderr)
        if len(cep) == 8 and cep.isdigit():
            url = f"https://viacep.com.br/ws/{cep}/json/"
            try:
                response = requests.get(url, timeout=5)
                response.raise_for_status()
                data = response.json()
                print(f"DEBUG: Resposta da ViaCEP: {data}", file=sys.stderr)

                if "erro" not in data:
                    self.entries_cliente["endereço"].setText(data.get("logradouro", "") or "")
                    if not self.entries_cliente["número"].text().strip():
                        self.entries_cliente["número"].clear()
                    self.entries_cliente["bairro"].setText(data.get("bairro", "") or "")
                    self.entries_cliente["cidade"].setText(data.get("localidade", "") or "")
                    self.entries_cliente["uf"].setText(data.get("uf", "") or "")
                else:
                    QMessageBox.warning(self, "CEP Inválido", "CEP não encontrado ou inválido.")
                    self.entries_cliente["endereço"].clear()
                    self.entries_cliente["bairro"].clear()
                    self.entries_cliente["cidade"].clear()
                    self.entries_cliente["uf"].clear()
            except requests.exceptions.RequestException as e:
                QMessageBox.critical(self, "Erro de Conexão",
                                     f"Não foi possível consultar o CEP: {e}\nVerifique sua conexão com a internet.")
                print(f"Erro de conexão no autopreencher CEP: {e}", file=sys.stderr)
            except Exception as e:
                QMessageBox.critical(self, "Erro Inesperado", f"Ocorreu um erro ao autopreencher o CEP: {e}")
                print(f"Erro inesperado no autopreencher CEP: {e}", file=sys.stderr)
        elif len(cep) > 0 and (len(cep) != 8 or not cep.isdigit()):
            QMessageBox.warning(self, "CEP Inválido", "CEP deve conter 8 dígitos numéricos.")
            self.entries_cliente["endereço"].clear()
            self.entries_cliente["bairro"].clear()
            self.entries_cliente["cidade"].clear()
            self.entries_cliente["uf"].clear()

    def _coletar_dados_form(self):
        dados = {
            "Numero_OS": self.entry_numero_os.text(),
            "Data_OS": self.label_data.text().split(' ')[0],
            "Hora_OS": self.label_data.text().split(' ')[1],
            "Nome_Cliente": self.entries_cliente["nome"].text(),
            "Endereco_Cliente": self.entries_cliente["endereço"].text(),
            "Numero_Imovel_Cliente": self.entries_cliente["número"].text(),
            "Bairro_Cliente": self.entries_cliente["bairro"].text(),
            "Cidade_Cliente": self.entries_cliente["cidade"].text(),
            "UF_Cliente": self.entries_cliente["uf"].text(),
            "CEP_Cliente": self.entries_cliente["cep"].text(),
            "Telefone_Cliente": self.entries_cliente["telefone"].text(),
            "CPF_CNPJ_Cliente": self.entries_cliente["cpf_cnpj"].text(),
            "Placa_Veiculo": self.entries_veiculo["placa"].text(),
            "Marca_Veiculo": self.entries_veiculo["marca"].text(),
            "Modelo_Veiculo": self.entries_veiculo["modelo"].text(),
            "Cor_Veiculo": self.entries_veiculo["cor"].text(),
            "Ano_Veiculo": self.entries_veiculo["ano"].text(),
            "KM_Atual_Veiculo": self.entries_veiculo["km_atual"].text().replace('.', ''),
            "Combustivel_Veiculo": self.entries_veiculo["combustível"].currentText() if isinstance(
                self.entries_veiculo["combustível"], QComboBox) else self.entries_veiculo["combustível"].text(),
            "Box_Veiculo": self.entries_veiculo["box"].currentText() if isinstance(self.entries_veiculo["box"],
                                                                                   QComboBox) else self.entries_veiculo[
                "box"].text(),
            "Problema_Informado": self.text_problema_informado.toPlainText(),
            "Problema_Constatado": self.text_problema_constatado.toPlainText(),
            "Servico_Executado": self.text_servico_executado.toPlainText(),
            "Detalhes_Itens": "; ".join([
                                            f"Tipo: {item['tipo']} | Ref: {item['referencia']} | Desc: {item['descricao']} | Qtd: {item['quantia']} | Val: {item['valor']:.2f} | Desc(%): {item['desc']:.0f} | Total: {item['valor_total']:.2f}"
                                            for item in self.itens_pecas_servicos_cache]),
            "Total_Itens": sum(item['valor_total'] for item in self.itens_pecas_servicos_cache),
            "Deslocamento": 0.00,
            "Desconto_Geral": 0.00,
            "Responsavel": self.entries_finais["responsável"].text(),
            "Situacao_Atual": self.combo_situacao_atual.currentText(),
            "Condicoes_Pagamento": self.combo_condicoes_pagamento.currentText()
        }
        dados["Valor_Total_Final"] = dados["Total_Itens"] + dados["Deslocamento"] - dados["Desconto_Geral"]

        dados["Itens_Pecas_Servicos"] = self.itens_pecas_servicos_cache

        return dados

    def _preencher_campos_form(self, dados_os_dict):
        self._limpar_campos()

        def get_display_value(key, default_value=""):
            value = dados_os_dict.get(key, default_value)
            if pd.isna(value) or value is None:
                return ""
            return str(value)

        self.entry_numero_os.setReadOnly(False)
        self.entry_numero_os.setText(get_display_value("Numero_OS"))
        self.entry_numero_os.setReadOnly(True)

        self.label_data.setText(f"{get_display_value('Data_OS')} {get_display_value('Hora_OS')}")

        self.entries_cliente["nome"].setText(get_display_value("Nome_Cliente"))
        self.entries_cliente["endereço"].setText(get_display_value("Endereco_Cliente"))
        self.entries_cliente["número"].setText(get_display_value("Numero_Imovel_Cliente"))
        self.entries_cliente["bairro"].setText(get_display_value("Bairro_Cliente"))
        self.entries_cliente["cidade"].setText(get_display_value("Cidade_Cliente"))
        self.entries_cliente["uf"].setText(get_display_value("UF_Cliente"))
        self.entries_cliente["cep"].setText(get_display_value("CEP_Cliente"))
        self.entries_cliente["telefone"].setText(get_display_value("Telefone_Cliente"))
        self.entries_cliente["cpf_cnpj"].setText(get_display_value("CPF_CNPJ_Cliente"))

        self.entries_veiculo["placa"].setText(get_display_value("Placa_Veiculo"))
        self.entries_veiculo["marca"].setText(get_display_value("Marca_Veiculo"))
        self.entries_veiculo["modelo"].setText(get_display_value("Modelo_Veiculo"))
        self.entries_veiculo["cor"].setText(get_display_value("Cor_Veiculo"))
        self.entries_veiculo["ano"].setText(get_display_value("Ano_Veiculo"))

        km_atual_raw = dados_os_dict.get("KM_Atual_Veiculo", "")
        if pd.isna(km_atual_raw) or km_atual_raw is None:
            self.entries_veiculo["km_atual"].setText("")
        else:
            self._formatar_quilometragem(str(km_atual_raw))

        # Para ComboBox, setar o texto diretamente ou o índice
        if isinstance(self.entries_veiculo["combustível"], QComboBox):
            self.entries_veiculo["combustível"].setCurrentText(get_display_value("Combustivel_Veiculo"))
        else:
            self.entries_veiculo["combustível"].setText(get_display_value("Combustivel_Veiculo"))

        if isinstance(self.entries_veiculo["box"], QComboBox):
            self.entries_veiculo["box"].setCurrentText(get_display_value("Box_Veiculo"))
        else:
            self.entries_veiculo["box"].setText(get_display_value("Box_Veiculo"))

        self.text_problema_informado.setText(get_display_value("Problema_Informado"))
        self.text_problema_constatado.setText(get_display_value("Problema_Constatado"))
        self.text_servico_executado.setText(get_display_value("Servico_Executado"))

        self.combo_situacao_atual.setCurrentText(get_display_value("Situacao_Atual"))
        self.combo_condicoes_pagamento.setCurrentText(get_display_value("Condicoes_Pagamento"))

        self.itens_pecas_servicos_cache = []
        self.listbox_itens.clear()

        # Correção para popular itens_pecas_servicos_cache corretamente
        # Se os dados.Detalhes_Itens for uma string semi-estruturada, precisamos parsear de volta
        itens_str = get_display_value("Detalhes_Itens")
        if itens_str:
            for item_entry_str in itens_str.split('; '):
                if item_entry_str.strip():
                    try:
                        # Exemplo de parse: "Tipo: Peça | Ref: 1 | Desc: Pastilhas | Qtd: 2 | Val: 50.00 | Desc(%): 0 | Total: 100.00"
                        parts = {k.strip(): v.strip() for k, v in
                                 (item.split(': ', 1) for item in item_entry_str.split(' | '))}

                        item_data = {
                            "tipo": parts.get("Tipo", "N/A"),
                            "referencia": parts.get("Ref", "N/A"),
                            "descricao": parts.get("Desc", "N/A"),
                            "uni": "un",  # Assumindo unidade padrão ao carregar
                            "valor": float(
                                parts["Val"].replace('R$', '').replace(',', '').strip()) if "Val" in parts else 0.0,
                            "quantia": int(parts["Qtd"].strip()) if "Qtd" in parts else 0,
                            "desc": float(parts["Desc(%)"].replace('%', '').strip()) if "Desc(%)" in parts else 0.0,
                            "valor_total": float(
                                parts["Total"].replace('R$', '').replace(',', '').strip()) if "Total" in parts else 0.0,
                        }
                        self.itens_pecas_servicos_cache.append(item_data)
                        self.listbox_itens.addItem(
                            f"Tipo: {item_data['tipo']} | Ref: {item_data['referencia']} - {item_data['descricao']} | Qtde: {item_data['quantia']} x R${item_data['valor']:.2f} | Desc: {item_data['desc']:.0f}% = R${item_data['valor_total']:.2f}"
                        )
                    except Exception as e:
                        print(f"Erro ao parsear item do Excel durante o carregamento: '{item_entry_str}' - {e}",
                              file=sys.stderr)
                        self.listbox_itens.addItem(item_entry_str)  # Adiciona como texto bruto se falhar

        if self.itens_pecas_servicos_cache and self.listbox_itens.count() > 0 and not self.listbox_itens.item(
                0).text().startswith("Tipo:"):  # Ajusta a verificação
            QMessageBox.information(self, "Aviso de Itens",
                                    "Itens carregados. A edição de itens complexos pode exigir a remoção e adição.")

        self.entries_finais["responsável"].setText(get_display_value("Responsavel"))

        self._atualizar_totais()

    def _buscar_os(self):
        os_id_busca = self.entry_busca_os.text().strip()

        if not os_id_busca:
            QMessageBox.warning(self, "Campo Vazio", "Por favor, digite o número da OS para buscar.")
            return

        if os_id_busca.isdigit():
            os_id_busca = str(os_id_busca).zfill(6)
        print(f"DEBUG: Buscando OS com ID formatado: '{os_id_busca}'", file=sys.stderr)

        os_encontrada = self.df_os[self.df_os['Numero_OS'] == os_id_busca]

        if not os_encontrada.empty:
            dados_os_dict = os_encontrada.iloc[0].to_dict()
            self._preencher_campos_form(dados_os_dict)
            QMessageBox.information(self, "OS Encontrada", f"Ordem de Serviço {os_id_busca} carregada com sucesso!")
            self.entry_busca_os.clear()
        else:
            QMessageBox.warning(self, "OS Não Encontrada",
                                f"Ordem de Serviço {os_id_busca} não encontrada no arquivo Excel.")
            self._limpar_campos()

    def _salvar_os(self):
        dados_os_coletados = self._coletar_dados_form()

        if not dados_os_coletados["Nome_Cliente"] or not dados_os_coletados["Placa_Veiculo"]:
            QMessageBox.warning(self, "Dados Incompletos",
                                "Nome do Cliente e Placa do Veículo são obrigatórios para salvar.")
            return

        dados_salvar = dados_os_coletados.copy()

        # CORREÇÃO PARA GARANTIR TIPOS DE DADOS COMPATÍVEIS PARA O PANDAS
        # Isso é crucial para evitar o FutureWarning de dtype e possíveis problemas de busca.
        # Converter para o tipo numérico adequado ou pd.NA se vazio.

        # KM_Atual_Veiculo
        km_limpo = ''.join(filter(str.isdigit, str(dados_salvar['KM_Atual_Veiculo'] or '')))
        try:
            dados_salvar['KM_Atual_Veiculo'] = int(km_limpo) if km_limpo else pd.NA
        except ValueError:
            dados_salvar['KM_Atual_Veiculo'] = pd.NA

        # Ano_Veiculo
        ano_str = str(dados_salvar['Ano_Veiculo'] or '')
        try:
            dados_salvar['Ano_Veiculo'] = int(ano_str) if ano_str.isdigit() else pd.NA
        except ValueError:
            dados_salvar['Ano_Veiculo'] = pd.NA

        # Numero_Imovel_Cliente (campo Número)
        num_imovel_str = str(dados_salvar['Numero_Imovel_Cliente'] or '')
        try:
            dados_salvar['Numero_Imovel_Cliente'] = int(num_imovel_str) if num_imovel_str.isdigit() else pd.NA
        except ValueError:
            dados_salvar['Numero_Imovel_Cliente'] = pd.NA

        # Conversão de Valor Total de Itens e Valor Total Final (garante float)
        campos_monetarios = ['Total_Itens', 'Valor_Total_Final', 'Deslocamento', 'Desconto_Geral']
        for campo in campos_monetarios:
            valor_str = str(dados_salvar[campo]).replace('.', '').replace(',',
                                                                          '.')  # Remove pontos de milhar e substitui vírgula por ponto decimal
            try:
                dados_salvar[campo] = float(valor_str) if valor_str else pd.NA
            except ValueError:
                dados_salvar[campo] = pd.NA  # Define como NA se a conversão falhar

        if dados_salvar["Numero_OS"].isdigit():
            dados_salvar["Numero_OS"] = str(dados_salvar["Numero_OS"]).zfill(6)
        print(f"DEBUG: Salvando OS com ID formatado: '{dados_salvar['Numero_OS']}'", file=sys.stderr)

        df_nova_os_linha = pd.DataFrame([dados_salvar])

        for col in self._get_expected_columns():
            if col not in df_nova_os_linha.columns:
                df_nova_os_linha[col] = pd.NA

        current_os_id = str(dados_os_coletados['Numero_OS'])
        os_existente_idx = self.df_os[self.df_os['Numero_OS'].astype(str) == current_os_id].index

        try:
            if not os_existente_idx.empty:
                idx = os_existente_idx[0]
                for col in df_nova_os_linha.columns:
                    if col not in self.df_os.columns:
                        self.df_os[col] = pd.NA
                    self.df_os.at[idx, col] = df_nova_os_linha.at[0, col]
                QMessageBox.information(self, "OS Atualizada",
                                        f"Ordem de Serviço {current_os_id} atualizada com sucesso no Excel!")
            else:
                self.df_os = pd.concat([self.df_os, df_nova_os_linha], ignore_index=True)
                QMessageBox.information(self, "OS Salva",
                                        f"Ordem de Serviço {current_os_id} salva com sucesso no Excel!")

            self.df_os.to_excel(ARQUIVO_EXCEL, index=False)
            print(f"Dados salvos em {ARQUIVO_EXCEL}")

        except Exception as e:
            QMessageBox.critical(self, "Erro ao Salvar", f"Não foi possível salvar os dados no Excel: {e}")
            print(f"Detalhes do erro ao salvar Excel: {e}", file=sys.stderr)

    def _deletar_os(self):
        current_os_id = self.entry_numero_os.text().strip()

        id_to_delete = current_os_id if self.entry_numero_os.isReadOnly() and current_os_id else None

        if not id_to_delete:
            search_os_id = self.entry_busca_os.text().strip()
            if search_os_id:
                id_to_delete = search_os_id
            else:
                QMessageBox.warning(self, "Deletar OS", "Nenhuma OS carregada ou ID de busca para deletar.")
                return

        if id_to_delete.isdigit():
            id_to_delete = str(id_to_delete).zfill(6)

        reply = QMessageBox.question(self, 'Deletar OS',
                                     f"Tem certeza que deseja deletar a OS {id_to_delete}?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            os_existente_idx = self.df_os[self.df_os['Numero_OS'].astype(str) == id_to_delete].index
            if not os_existente_idx.empty:
                self.df_os = self.df_os.drop(os_existente_idx).reset_index(drop=True)
                try:
                    self.df_os.to_excel(ARQUIVO_EXCEL, index=False)
                    QMessageBox.information(self, "Deletar OS",
                                            f"Ordem de Serviço {id_to_delete} deletada com sucesso do Excel!")
                    self._limpar_campos()
                except Exception as e:
                    QMessageBox.critical(self, "Erro ao Deletar", f"Não foi possível deletar a OS do Excel: {e}")
                    print(f"Detalhes do erro ao deletar Excel: {e}", file=sys.stderr)
            else:
                QMessageBox.warning(self, "Deletar OS", f"Ordem de Serviço {id_to_delete} no encontrada para deletar.")

    def _imprimir_os_pdf(self):
        """Gera um archivo PDF usando la plantilla HTML y WeasyPrint."""
        self._salvar_os()

        dados_os = self._coletar_dados_form()

        if not dados_os["Numero_OS"] or not dados_os["Nome_Cliente"] or not dados_os["Placa_Veiculo"]:
            QMessageBox.warning(self, "Datos Mínimos",
                                "Número de OS, Nombre del Cliente y Placa del Vehículo son obligatorios para imprimir.")
            return

        # --- CORREÇÃO DEL NOMBRE DEL ARCHIVO PDF: Elimina caracteres inválidos de Windows ---
        # Garantir que a data e hora estejam limpas e formatadas para o nome do arquivo.
        formatted_date_for_filename = QDateTime.currentDateTime().toString("yyyy-MM-dd")  # Ex: 2025-06-11
        formatted_time_for_filename = QDateTime.currentDateTime().toString("hh_mm_ss").replace(':', '_')  # Ex: 13_30_45

        # O Numero_OS já deve estar limpo ao ser carregado/salvo, mas um extra para segurança.
        safe_os_number = "".join(c for c in dados_os['Numero_OS'] if c.isalnum() or c == '_')

        # Usar tempfile para criar um arquivo temporário que será limpo automaticamente
        # 'delete=False' para Windows, porque o subprocesso pode manter o arquivo aberto
        temp_file = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        filename_full_path = temp_file.name
        temp_file.close()  # Fechar o handle para que o WeasyPrint possa escrever

        # Registrar o arquivo temporário para ser excluído quando o programa sair
        _temp_files_to_clean.append(filename_full_path)
        # --- FIN DE LA CORREÇÃO ---

        # Depuração para verificar o caminho do logo e o nome do arquivo PDF
        logo_absolute_path = os.path.abspath(ARQUIVO_LOGO) if os.path.exists(ARQUIVO_LOGO) else None
        print(f"DEBUG: Logo path enviado para o template: {logo_absolute_path}", file=sys.stderr)
        print(f"DEBUG: Nome do archivo PDF gerado (temporário): {filename_full_path}", file=sys.stderr)

        # CÓDIGO PARA LER IMAGEM E CODIFICAR EM BASE64
        logo_base64_data = None
        if logo_absolute_path and os.path.exists(logo_absolute_path):
            try:
                # Abrir a imagem e redimensionar para a largura desejada (ex: 25mm = ~94 pixels a 96 DPI)
                # WeasyPrint lida com mm, mas o redimensionamento em pixels é mais comum com Pillow.
                # 25mm a 96 DPI (pixels por polegada) é 25 / 25.4 * 96 = ~94.5 pixels. Vamos usar 100 pixels de largura.
                target_width_px = 100
                img = Image.open(logo_absolute_path)
                original_width, original_height = img.size

                # Calcular nova altura mantendo a proporção
                new_height_px = int((target_width_px / original_width) * original_height)

                img_resized = img.resize((target_width_px, new_height_px),
                                         Image.Resampling.LANCZOS)  # Melhor filtro de redimensionamento

                # Salvar a imagem redimensionada em um buffer para codificação Base64
                import io
                buf = io.BytesIO()
                img_resized.save(buf, format="PNG")  # Salvar como PNG para manter qualidade e transparência
                logo_base64_data = base64.b64encode(buf.getvalue()).decode('utf-8')
                buf.close()

                print(f"DEBUG: Logo Base64 data (first 50 chars): {logo_base64_data[:50]}...", file=sys.stderr)
            except Exception as e:
                print(f"ERROR: No fue posible codificar/redimensionar el logo en Base64: {e}", file=sys.stderr)
                logo_base64_data = None
        else:
            print(f"ALERTA: Archivo de logo no encontrado en: {ARQUIVO_LOGO}", file=sys.stderr)

        try:
            template = self.env.get_template(HTML_TEMPLATE_FILE)

            template_data = {
                'dados': dados_os,
                'info_oficina': INFO_OFICINA,
                'logo_base64': logo_base64_data  # Pasa la string Base64 para el template
            }

            html_content = template.render(template_data)

            HTML(string=html_content, base_url=os.getcwd()).write_pdf(filename_full_path)

            QMessageBox.information(self, "PDF Generado",
                                    f"Orden de Servicio guardada en:\n{filename_full_path}\nSerá abierta para visualización.")

            try:
                print(f"Intentando abrir el PDF: {filename_full_path}", file=sys.stderr)
                if platform.system() == "Windows":
                    subprocess.run(["start", "", filename_full_path], shell=True, check=True)
                elif platform.system() == "Darwin":
                    subprocess.run(["open", filename_full_path], check=True)
                else:
                    subprocess.run(["xdg-open", filename_full_path], check=True)
                print("PDF abierto con éxito.", file=sys.stderr)
            except FileNotFoundError:
                QMessageBox.warning(self, "Visor no encontrado",
                                    "No se pudo encontrar un programa para abrir PDFs. Instale un visor o verifique el PATH.")
                print(f"Error: Visor de PDF no encontrado para '{platform.system()}'", file=sys.stderr)
            except subprocess.CalledProcessError as e:
                QMessageBox.warning(self, "Error al abrir PDF",
                                    f"El comando para abrir el PDF falló (código {e.returncode}). Error: {e.stderr.decode()}")
                print(f"Error en el subproceso al abrir PDF: {e}", file=sys.stderr)
            except Exception as e:
                QMessageBox.warning(self, "Error al abrir PDF",
                                    f"No se pudo abrir el PDF automáticamente. Por favor, ábralo manualmente desde: {filename_full_path}\nError: {e}")
                print(f"Error inesperado al intentar abrir PDF: {e}", file=sys.stderr)

        except Exception as e:
            QMessageBox.critical(self, "Error en la Generación del PDF",
                                 f"Ocurrió un error al generar el PDF con Weasyprint: {e}\nVerifique la plantilla HTML y la configuración de las bibliotecas.")
            print(f"Error detallado en la generación del PDF con Weasyprint: {e}", file=sys.stderr)
            # A remoção do arquivo temporário é garantida pelo atexit.register, mesmo em caso de erro aqui.


# --- Ejecución de la Aplicación ---
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OficinaOSApp()
    window.show()
    sys.exit(app.exec_())
