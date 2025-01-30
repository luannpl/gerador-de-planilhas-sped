from PySide6.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QVBoxLayout, QMainWindow
from PySide6 import QtWidgets, QtGui, QtCore
import pandas as pd
import sys
import os
from utils.icone import usar_icone
from utils.resource import resource_path

tabela0000 = []
tabela0150 = []
tabela0200 = []
tabelac100 = []
tabelac170 = []

def gerar_tabelas(progresso):
    msg = QtWidgets.QMessageBox()
    msg.setIcon(QtWidgets.QMessageBox.Information)
    msg.setText("Selecionar pasta ou arquivo?")
    msg.setWindowTitle("Selecionar")
    usar_icone(msg)
    btn_pasta = msg.addButton("Pasta", QtWidgets.QMessageBox.AcceptRole)
    btn_arquivo = msg.addButton("Arquivo", QtWidgets.QMessageBox.AcceptRole)
    msg.exec()
    progresso.setValue(0)
    
    if msg.clickedButton() == btn_arquivo:
        nomes_arquivo, _ = QFileDialog.getOpenFileNames(None, "Selecione o arquivo de texto", "", "Arquivos de Texto (*.txt);;Todos os Arquivos (*)")
    
    elif msg.clickedButton() == btn_pasta:
        pasta_selecionada = QFileDialog.getExistingDirectory(None, "Selecione uma pasta")
        if pasta_selecionada:
            nomes_arquivo = [os.path.join(pasta_selecionada, f) for f in os.listdir(pasta_selecionada) if f.endswith(".txt")]
    
    if not nomes_arquivo:
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Warning)
        msg.setText("Nenhum arquivo selecionado ou nenhum arquivo encontrado na pasta.")
        msg.setWindowTitle("Erro")
        usar_icone(msg)
        msg.exec()
        return

    tabela0000.clear()
    tabela0150.clear()
    tabela0200.clear()
    tabelac100.clear()
    tabelac170.clear()

    total_arquivos = len(nomes_arquivo)
    cod_fornecedores_vistos = set()
    cod_item_vistos = set() 
    mapa_cod_item_descr = {}

    try:
        for i, nome_arquivo in enumerate(nomes_arquivo):
            with open(nome_arquivo, 'r', errors='ignore') as arquivo:
                filial = None
                periodo = None
                for linha in arquivo:
                    if not (linha.startswith('|0000|') or linha.startswith('|0150|') or linha.startswith('|0200|') or linha.startswith('|C100|') or linha.startswith('|C170|')):
                        continue
                    if linha.startswith('|0000|'):
                        dados = linha.split('|')
                        filial = dados[7][8:12]
                        periodo = f"{dados[4][2:4]}/{dados[4][4:]}"
                        print(f"Filial: {filial}")
                        print(f"Período: {periodo}")
                        tabela0000.append({
                            'reg': dados[1],
                            'cod_ver': dados[2],
                            'cod_fin': dados[3],
                            'dt_ini': dados[4],
                            'dt_fin': dados[5],
                            'nome': dados[6],
                            'cnpj': dados[7],
                            'cpf': dados[8],
                            'uf': dados[9],
                            'ie': dados[10],
                            'cod_mun': dados[11],
                            'im': dados[12],
                            'suframa': dados[13],
                            'ind_perfil': dados[14],
                            'ind_ativ': dados[15],
                            'filial': filial,
                            'periodo': periodo
                        })

                    if linha.startswith('|0150|'):
                        dados = linha.split('|')
                        cod_fornecedor = dados[2]
                        if cod_fornecedor in cod_fornecedores_vistos:
                            continue

                        cod_fornecedores_vistos.add(cod_fornecedor) 
                        tabela0150.append({
                            'reg': dados[1],
                            'cod_fornecedor': cod_fornecedor,
                            'nome_fornecedor': dados[3],
                            'cod_pais': dados[4],
                            'cnpj': dados[5],
                            'cpf': dados[6],
                            'ie': dados[7],
                            'cod_municipio': dados[8],
                            'suframa': dados[9],
                            'endereco': dados[10],
                            'num': dados[11],
                            'compl': dados[12],
                            'bairro': dados[13],
                            'periodo': periodo
                        })
                    
                    if linha.startswith('|0200|'):
                        dados = linha.split('|')
                        cod_item = dados[2]
                        cod_item = cod_item.lstrip('0')
                        if cod_item in cod_item_vistos:
                            continue
                        cod_item_vistos.add(cod_item)
                        mapa_cod_item_descr[cod_item] = dados[3]
                        tabela0200.append({
                            'reg': dados[1],
                            'cod_item': cod_item,
                            'descr_item': dados[3],
                            'cod_barra': dados[4],
                            'cod_ant_item': dados[5],
                            'unid_inv': dados[6],
                            'tipo_item': dados[7],
                            'cod_ncm': dados[8],
                            'ex_ipi': dados[9],
                            'cod_gen': dados[10],
                            'cod_lst': dados[11],
                            'aliq_icms': dados[12],
                            'cest': dados[13],
                            'periodo': periodo
                        })
                    

                    if linha.startswith('|C100|'):
                        dados = linha.split('|')
                        tabelac100.append({
                            'reg': dados[1],
                            'ind_oper': dados[2],
                            'ind_emit': dados[3],
                            'cod_part': dados[4],
                            'cod_mod': dados[5],
                            'cod_sit': dados[6],
                            'ser': dados[7],
                            'num_doc': dados[8],
                            'chv_nfe': dados[9],
                            'dt_doc': dados[10],
                            'dt_e_s': dados[11],
                            'vl_doc': dados[12],
                            'ind_pag': dados[13],
                            'vl_desc': dados[14],
                            'vl_abat_nt': dados[15],
                            'vl_merc': dados[16],
                            'ind_frt': dados[17],
                            'vl_frt': dados[18],
                            'vl_seg': dados[19],
                            'vl_out_da': dados[20],
                            'vl_bc_icms': dados[21],
                            'vl_icms': dados[22],
                            'vl_bc_icms_st': dados[23],
                            'vl_icms_st': dados[24],
                            'vl_ipi': dados[25],
                            'vl_pis': dados[26],
                            'vl_cofins': dados[27],
                            'vl_pis_st': dados[28],
                            'vl_cofins_st': dados[29],
                            'filial': filial,
                            'periodo': periodo
                            
                        })
                        c100_atual = tabelac100[-1]  # Último item inserido
                        ind_oper, cod_part, num_doc, chv_nfe = c100_atual["ind_oper"], c100_atual["cod_part"], c100_atual["num_doc"], c100_atual["chv_nfe"]
                    
                    if linha.startswith('|C170|'):
                        dados = linha.split('|')
                        cod_item = dados[3]
                        cod_item = cod_item.lstrip('0')
                        descr_compl = mapa_cod_item_descr.get(cod_item, dados[4])
                        tabelac170.append({
                            'reg': dados[1],
                            'num_item': dados[2],
                            'cod_item': dados[3],
                            'descr_compl': descr_compl,
                            'qtd': dados[5],
                            'unid': dados[6],
                            'vl_item': dados[7],
                            'vl_desc': dados[8],
                            'ind_mov': dados[9],
                            'cst_icms': dados[10],
                            'cfop': dados[11],
                            'cod_nat': dados[12],
                            'vl_bc_icms': dados[13],
                            'aliq_icms': dados[14],
                            'vl_icms': dados[15],
                            'vl_bc_icms_st': dados[16],
                            'aliq_st': dados[17],
                            'vl_icms_st': dados[18],
                            'ind_apur': dados[19],
                            'cst_ipi': dados[20],
                            'cod_enq': dados[21],
                            'vl_bc_ipi': dados[22],
                            'aliq_ipi': dados[23],
                            'vl_ipi': dados[24],
                            'cst_pis': dados[25],
                            'vl_bc_pis': dados[26],
                            'aliq_pis': dados[27],
                            'quant_bc_pis': dados[28],
                            'aliq_pis_reais': dados[29],
                            'vl_pis': dados[30],
                            'cst_cofins': dados[31],
                            'vl_bc_cofins': dados[32],
                            'aliq_cofins': dados[33],
                            'quant_bc_cofins': dados[34],
                            'aliq_cofins_reais': dados[35],
                            'vl_cofins': dados[36],
                            'cod_cta': dados[37],
                            'vl_abat_nt': dados[38],
                            'ind_oper': ind_oper,
                            'cod_part': cod_part,
                            'num_doc': num_doc,
                            'chv_nfe': chv_nfe,
                            'filial': filial,
                            'periodo': periodo
                        })

            progresso.setValue(int(((i + 1) / total_arquivos) * 100))
            print(mapa_cod_item_descr[cod_item])
   
                
        msg_salvar = QtWidgets.QMessageBox()
        progresso.setValue(100)
        msg_salvar.setIcon(QtWidgets.QMessageBox.Information)
        msg_salvar.setText("Deseja salvar o arquivo gerado?")
        msg_salvar.setWindowTitle("Salvar Arquivo")
        usar_icone(msg_salvar)
        btn_salvar = msg_salvar.addButton("Salvar", QtWidgets.QMessageBox.AcceptRole)
        msg_salvar.addButton("Cancelar", QtWidgets.QMessageBox.RejectRole)
        msg_salvar.exec()

        if msg_salvar.clickedButton() == btn_salvar:
            progresso.setValue(0)
            nome_arquivo_salvar, _ = QFileDialog.getSaveFileName(None, "Salvar Arquivo", "", "Arquivos Excel (*.xlsx);;Todos os Arquivos (*)")
        
            if nome_arquivo_salvar:
                with pd.ExcelWriter(nome_arquivo_salvar) as writer:
                    df = pd.DataFrame(tabela0000)
                    df.to_excel(writer, sheet_name='0000', index=False)
                    progresso.setValue(5)
                    df = pd.DataFrame(tabela0150)
                    df.to_excel(writer, sheet_name='0150', index=False)
                    progresso.setValue(15)
                    df = pd.DataFrame(tabela0200)
                    df.to_excel(writer, sheet_name='0200', index=False)
                    progresso.setValue(40)
                    df = pd.DataFrame(tabelac100)
                    df.to_excel(writer, sheet_name='C100', index=False)
                    progresso.setValue(70)
                    df = pd.DataFrame(tabelac170)
                    df.to_excel(writer, sheet_name='C170', index=False)
                    progresso.setValue(85)

                msg = QtWidgets.QMessageBox()
                progresso.setValue(100)
                msg.setIcon(QtWidgets.QMessageBox.Information)
                msg.setText(f"Arquivo salvo com sucesso em {nome_arquivo_salvar}!")
                msg.setWindowTitle("Sucesso")
                usar_icone(msg)
                msg.exec()
                progresso.setValue(0)
            else:
                msg = QtWidgets.QMessageBox()
                msg.setIcon(QtWidgets.QMessageBox.Warning)
                msg.setText("Nenhum arquivo selecionado.")
                msg.setWindowTitle("Erro")
                usar_icone(msg)
                msg.exec()

        else:
            progresso.setValue(0)
            msg = QtWidgets.QMessageBox()
            msg.setIcon(QtWidgets.QMessageBox.Warning)
            msg.setText("Operação cancelada.")
            msg.setWindowTitle("Erro")
            usar_icone(msg)
            msg.exec()
    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")


def main():
    app = QtWidgets.QApplication(sys.argv)

    janela = QtWidgets.QMainWindow()
    usar_icone(janela)

    janela.setWindowTitle("Gerador de Tabelas")
    janela.setGeometry(100, 100, 800, 600)

    widget_central = QtWidgets.QWidget()
    janela.setCentralWidget(widget_central)

    layout = QtWidgets.QVBoxLayout(widget_central)

    botao_frame = QtWidgets.QHBoxLayout()
    layout.addLayout(botao_frame)

    botao_frame = QtWidgets.QHBoxLayout()
    layout.addLayout(botao_frame)

    botoes = [
        ("Inserir Arquivos SPEDS", lambda: gerar_tabelas(progresso)),
    ]

    for texto, funcao in botoes:
        botao = QtWidgets.QPushButton(texto)
        botao.clicked.connect(funcao)
        botao.setFont(QtGui.QFont("Arial", 14))
        botao.setStyleSheet(""" 
            QPushButton {
                background-color: #001F3F;
                color: white;
                border: none;
                padding: 10px 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #191970;
            }
        """)
        botao.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        botao_frame.addWidget(botao)
    
    imagem_placeholder = QtWidgets.QLabel()
    imagem_placeholder.setPixmap(QtGui.QPixmap(resource_path("images\\logo.png")).scaled(350, 350, QtCore.Qt.KeepAspectRatio))
    imagem_placeholder.setAlignment(QtCore.Qt.AlignCenter)
    
    vbox_layout = QtWidgets.QVBoxLayout()
    vbox_layout.addWidget(imagem_placeholder)
    vbox_layout.setAlignment(QtCore.Qt.AlignCenter)

    stack_layout = QtWidgets.QStackedLayout()
    layout.addLayout(stack_layout)

    container_widget = QtWidgets.QWidget()
    container_widget.setLayout(vbox_layout)

    stack_layout.addWidget(container_widget)
    stack_layout.setCurrentWidget(container_widget)

    progresso = QtWidgets.QProgressBar()
    progresso.setRange(0, 100)
    progresso.setValue(0)
    layout.addWidget(progresso)

    janela.showMaximized()

    app.exec()

if __name__ == '__main__':
    main()
