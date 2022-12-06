import sys, time, re
import pandas as pd

from validate_docbr import CPF
from PyQt5.QtWidgets import QApplication, QLabel, QWidget
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QAbstractTableModel, Qt, QModelIndex

validacpf = CPF()

class Banco():
	#Inicio da classe Banco, realiza a leitura das tabelas usadas como banco de dados
	def __init__(self):
		try:
			self.ferramentas = pd.read_excel('ferramentas.xlsx')
		except:
			self.ferramentas = {'Id': [], 'Descrição': [], 'Fabricante': [], 'Tensão': [], 'PN': [], 'Tamanho': [], 'UnidadeMedida': [], 'Tipo': [], 'Material': [], 'TempoReserva': []}
			self.ferramentas = pd.DataFrame(data = self.ferramentas)
			self.ferramentas.to_excel('ferramentas.xlsx', index=False)
			self.ferramentas = pd.read_excel('ferramentas.xlsx')
		try:
			self.tecnicos = pd.read_excel('tecnicos.xlsx', dtype={'CPF': 'str', 'Contato': 'str'})
		except:
			self.tecnicos = {'CPF': [],'Nome': [],'Contato': [],'Turno': [],'Equipe': []}
			self.tecnicos = pd.DataFrame(data = self.tecnicos)
			self.tecnicos.to_excel('tecnicos.xlsx', index=False)
			self.tecnicos = pd.read_excel('tecnicos.xlsx')

	#Função responsavel por adicionar as informações dos Tecnicos no banco de dados
	def addTecnico(self):
		if cpf.text() != '' and nome.text != '' and equipe.text != '' and tel2.text != '':
			dados = [cpf.text(), nome.text(), tel2.text(), turno.currentText(), equipe.text()]
		else:
			return

		#Atualizar itme já existente
		if dados[0] in self.tecnicos['CPF'].values:
			
			linha = self.tecnicos['CPF'].values.tolist().index(dados[0])
			dados[0] = ''.join(re.findall(r"\d", dados[0]))
			
			valida = validacpf.validate(dados[0])
			if valida == False:
				return
			if turno.currentText() == 'Celular':
				dados[2] = dados[2].split(':')[2]
				if len(dados[2]) != 9:
					print('Telefone incorreto')
					return
			if turno.currentText() == 'Rádio':
				dados[2] = dados[2].split(':')[1]
				
			#Adiciona os dados do Tecnico no banco de dados
			self.tecnicos.loc[linha, 'CPF'] = dados[0]
			self.tecnicos.loc[linha, 'Nome'] = dados[1]
			self.tecnicos.loc[linha, 'Contato'] = str(dados[2])
			self.tecnicos.loc[linha, 'Turno'] = dados[3]
			self.tecnicos.loc[linha, 'Equipe'] = dados[4]
			model = PandasModel(banco.tecnicos)
			tecnicoTable.setModel(model)

		#Adiciona novo item
		else:
			dados[0] = ''.join(re.findall(r"\d", dados[0]))
			
			valida = validacpf.validate(dados[0])
			if valida == False:
				print(valida)
				return
			if turno.currentText() == 'Celular':
				dados[2] = dados[2].split(':')[2]
				if len(dados[2]) != 9:
					print('Telefone incorreto')
					return
			if turno.currentText() == 'Rádio':
				dados[2] = dados[2].split(':')[1]
				
			#Adiciona os dados do Tecnico no banco de dados
			linha = len(self.tecnicos['CPF'])
			self.tecnicos.loc[linha, 'CPF'] = dados[0]
			self.tecnicos.loc[linha, 'Nome'] = dados[1]
			self.tecnicos.loc[linha, 'Contato'] = dados[2]
			self.tecnicos.loc[linha, 'Turno'] = dados[3]
			self.tecnicos.loc[linha, 'Equipe'] = dados[4]
			model = PandasModel(self.tecnicos)
			tecnicoTable.setModel(model)
		
		#Salva o banco de dados
		self.tecnicos.to_excel('tecnicos.xlsx', index=False)
	
	#Função responsavel por adicionar as informações das Ferramentas no banco de dados
	def addFerramenta(self):
		if ID.text() != '' and Tensão.text() != '' and descricao.text() != '' and fabricante.text() != '' and PN.text() != '' and tipo.text() != '' and material.text() != '' and tamanho.text() != '' and UC.text() != '' and tempoReserva.text() != '':
			dados = [ID.text(), Tensão.text(), descricao.text(), fabricante.text(), PN.text(), tipo.text(), material.text(), tamanho.text(), UC.text(), tempoReserva.text()]
		else:
			return
		

		#Atualizar itme já existente
		if str(dados[0]) in self.ferramentas['Id'].values:
			linha = self.ferramentas['Id'].values.tolist().index(dados[0])
				
			#Adiciona os dados do Tecnico no banco de dados
			self.ferramentas.loc[linha, 'Id'] = str(dados[0])
			self.ferramentas.loc[linha, 'Tensão'] = str(dados[1])
			self.ferramentas.loc[linha, 'Descrição'] = str(dados[2])
			self.ferramentas.loc[linha, 'Fabricante'] = str(dados[3])
			self.ferramentas.loc[linha, 'PN'] = str(dados[4])
			self.ferramentas.loc[linha, 'Tipo'] = str(dados[5])
			self.ferramentas.loc[linha, 'Material'] = str(dados[6])
			self.ferramentas.loc[linha, 'Tamanho'] = str(dados[7])
			self.ferramentas.loc[linha, 'UnidadeMedida'] = str(dados[8])
			self.ferramentas.loc[linha, 'TempoReserva'] = str(dados[9])
			model = PandasModel(banco.ferramentas)
			ferramentasTable.setModel(model)

		#Adiciona novo item
		else:
			#Adiciona os dados do Tecnico no banco de dados
			linha = len(self.ferramentas['Id'])
			self.ferramentas.loc[linha, 'Id'] = str(dados[0])
			self.ferramentas.loc[linha, 'Tensão'] = str(dados[1])
			self.ferramentas.loc[linha, 'Descrição'] = str(dados[2])
			self.ferramentas.loc[linha, 'Fabricante'] = str(dados[3])
			self.ferramentas.loc[linha, 'PN'] = str(dados[4])
			self.ferramentas.loc[linha, 'Tipo'] = str(dados[5])
			self.ferramentas.loc[linha, 'Material'] = str(dados[6])
			self.ferramentas.loc[linha, 'Tamanho'] = str(dados[7])
			self.ferramentas.loc[linha, 'UnidadeMedida'] = str(dados[8])
			self.ferramentas.loc[linha, 'TempoReserva'] = str(dados[9])
			model = PandasModel(self.ferramentas)
			ferramentasTable.setModel(model)
		
		#Salva o banco de dados
		self.ferramentas.to_excel('ferramentas.xlsx', index=False)

	def removeTecnico(self):
		for i in range(len(self.tecnicos)):
			if ''.join(re.findall(r"\d", cpf.text())) == self.tecnicos['CPF'][i]:
				self.tecnicos = self.tecnicos.drop(i)
				break
		model = PandasModel(self.tecnicos)
		tecnicoTable.setModel(model)

		#Salva o banco de dados
		self.tecnicos.to_excel('tecnicos.xlsx', index=False)

	def removeFerramenta(self):
		for i in range(len(self.ferramentas)):
			if str(ID.text()) == str(self.ferramentas['Id'][i]):
				self.ferramentas = self.ferramentas.drop(i)
				break
		model = PandasModel(self.ferramentas)
		ferramentasTable.setModel(model)

		#Salva o banco de dados
		self.ferramentas.to_excel('ferramentas.xlsx', index=False)

	#Apaga todas as informações de ambos os banco de dados
	def reset(self):
		ferramentas = {'Id': [], 'Descrição': [], 'Fabricante': [], 'Tensão': [], 'PN': [], 'Tamanho': [], 'UnidadeMedida': [], 'Tipo': [], 'Material': [], 'TempoReserva': []}
		ferramentas = pd.DataFrame(data = ferramentas)
		ferramentas.to_excel('ferramentas.xlsx', index=False)
		ferramentas = pd.read_excel('ferramentas.xlsx')
		tecnicos = {'CPF': [],'Nome': [],'Contato': [],'Turno': [],'Equipe': []}
		tecnicos = pd.DataFrame(data = tecnicos)
		tecnicos.to_excel('tecnicos.xlsx', index=False)
		tecnicos = pd.read_excel('tecnicos.xlsx')

banco = Banco()

app = QApplication([])

def mostrarTecnicos():
	addFerramentasFrame.hide()
	addTecnicoFrame.show()

def mostrarFerramentas():
	addTecnicoFrame.hide()
	addFerramentasFrame.show()

def treeMediaClicked(index):
	try:
		dados = [i.data() for i in tecnicoTable.selectedIndexes()]
		cpf.setText(dados[0])
		nome.setText(dados[1])
		if len(dados[2]) < 9:
			tel1.setCurrentIndex(1)
		else:
			tel1.setCurrentIndex(0)
		tel2.setText(dados[2])
		if dados[3] == 'Manhã':
			turno.setCurrentIndex(0)
		elif dados[3] == 'Tarde':
			turno.setCurrentIndex(1)
		elif dados[3] == 'Noite':
			turno.setCurrentIndex(2)
		equipe.setText(dados[4])
	except:
		dados = [i.data() for i in ferramentasTable.selectedIndexes()]
		ID.setText(str(dados[0]))
		Tensão.setText(dados[3])
		descricao.setText(dados[1])
		fabricante.setText(dados[2])
		PN.setText(dados[4])
		tipo.setText(dados[7])
		material.setText(dados[8])
		tamanho.setText(dados[5])
		UC.setText(dados[6])
		tempoReserva.setText(dados[9])


class PandasModel(QAbstractTableModel):
	"""A model to interface a Qt view with pandas dataframe """

	def __init__(self, dataframe: pd.DataFrame, parent=None):
		QAbstractTableModel.__init__(self, parent)
		self._dataframe = dataframe

	def rowCount(self, parent=QModelIndex()) -> int:
		""" Override method from QAbstractTableModel

		Return row count of the pandas DataFrame
		"""
		if parent == QModelIndex():
			return len(self._dataframe)

		return 0

	def columnCount(self, parent=QModelIndex()) -> int:
		"""Override method from QAbstractTableModel

		Return column count of the pandas DataFrame
		"""
		if parent == QModelIndex():
			return len(self._dataframe.columns)
		return 0

	def data(self, index: QModelIndex, role=Qt.ItemDataRole):
		"""Override method from QAbstractTableModel

		Return data cell from the pandas DataFrame
		"""
		if not index.isValid():
			return None

		if role == Qt.DisplayRole:
			return str(self._dataframe.iloc[index.row(), index.column()])

		return None

	def headerData(
		self, section: int, orientation: Qt.Orientation, role: Qt.ItemDataRole
	):
		"""Override method from QAbstractTableModel

		Return dataframe index as vertical header data and columns as horizontal header data.
		"""
		if role == Qt.DisplayRole:
			if orientation == Qt.Horizontal:
				return str(self._dataframe.columns[section])

			if orientation == Qt.Vertical:
				return str(self._dataframe.index[section])

		return None

style = """
	QLineEdit { 
		background-color: #fff;
		border-style: outset;
		border-width: 1px;
		border-radius: 5px;
		border-color: #909090;
	}

	QLineEdit::focus {
		border-style: outset;
		border-width: 1px;
		border-radius: 5px;
		border-color: navy;
	}

	QFrame {
		border: 1px;
		height: 100%;
	}

	QFrame#menu {
		background-color: #f0f0f0;	
		border-style: outset;
		border-width: 1px;
		border-color: #000;
	}

	QLabel {
		
	}

	QPushButton {
		background-color: #fff;
		border-style: outset;
		border-width: 1px;
		border-radius: 10px;
		border-color: #909090;
	}
	QPushButton:pressed {
		border-style: inset;
	}

	QComboBox {
		
	}
	"""

style2 = """
	QPushButton {
		background-color: #fff;
		border-style: outset;
		border-width: 1px;
		border-radius: 0px;
		border-color: #000;
	}
	QPushButton:pressed {
		background-color: #d0d0d0;
		border-style: inset;
		border-width: 1px;
	}
"""

# ---- Inicio ----
window = QWidget()
window.setWindowTitle("PyQt App")
window.setGeometry(100, 100, 1000, 600)
window.setMaximumSize(1000, 600)
window.setMinimumSize(1000, 600)
window.setStyleSheet(style)


# ---- Menu ----
menuFrame = QtWidgets.QFrame(window)
menuFrame.setObjectName("menu")
menuFrame.setGeometry(-1, -1, 1002, 45)
menuFrame.setStyleSheet(style2)

addTecnicoMenu = QtWidgets.QPushButton(menuFrame)
addTecnicoMenu.setText('Tecnico')
addTecnicoMenu.setGeometry(QtCore.QRect(0, 0, 100, 45))
addTecnicoMenu.clicked.connect(mostrarTecnicos)

addFerramentasMenu = QtWidgets.QPushButton(menuFrame)
addFerramentasMenu.setText('Ferramentas')
addFerramentasMenu.setGeometry(QtCore.QRect(99, 0, 100, 45))
addFerramentasMenu.clicked.connect(mostrarFerramentas)


# ---- addFerramenta ----
addFerramentasFrame = QtWidgets.QFrame(window)
addFerramentasFrame.setFrameShape(QtWidgets.QFrame.StyledPanel)
addFerramentasFrame.move(0, 40)

IDLabel = QtWidgets.QLabel(addFerramentasFrame)
IDLabel.setText('ID')
IDLabel.move(10, 10)
ID = QtWidgets.QLineEdit(addFerramentasFrame)
ID.setGeometry(QtCore.QRect(10, 30, 130, 25))
ID.setAttribute(QtCore.Qt.WA_MacShowFocusRect,0);

TensãoLabel = QtWidgets.QLabel(addFerramentasFrame)
TensãoLabel.setText('Tensão')
TensãoLabel.move(155, 10)
Tensão = QtWidgets.QLineEdit(addFerramentasFrame)
Tensão.setGeometry(QtCore.QRect(150, 30, 90, 25))
Tensão.setAttribute(QtCore.Qt.WA_MacShowFocusRect,0);

descricaoLabel = QtWidgets.QLabel(addFerramentasFrame)
descricaoLabel.setText('Descrição')
descricaoLabel.move(10, 70)
descricao = QtWidgets.QLineEdit(addFerramentasFrame)
descricao.setGeometry(QtCore.QRect(10, 90, 231, 25))
descricao.setAttribute(QtCore.Qt.WA_MacShowFocusRect,0);

fabricanteLabel = QtWidgets.QLabel(addFerramentasFrame)
fabricanteLabel.setText('Fabricante')
fabricanteLabel.move(10, 130)
fabricante = QtWidgets.QLineEdit(addFerramentasFrame)
fabricante.setGeometry(QtCore.QRect(10, 150, 231, 25))
fabricante.setAttribute(QtCore.Qt.WA_MacShowFocusRect,0);

PNLabel = QtWidgets.QLabel(addFerramentasFrame)
PNLabel.setText('PartNumber')
PNLabel.move(10, 190)
PN = QtWidgets.QLineEdit(addFerramentasFrame)
PN.setGeometry(QtCore.QRect(10, 210, 231, 25))
PN.setAttribute(QtCore.Qt.WA_MacShowFocusRect,0);

tipoLabel = QtWidgets.QLabel(addFerramentasFrame)
tipoLabel.setText('Tipo')
tipoLabel.move(10, 250)
tipo = QtWidgets.QLineEdit(addFerramentasFrame)
tipo.setGeometry(QtCore.QRect(10, 270, 231, 25))
tipo.setAttribute(QtCore.Qt.WA_MacShowFocusRect,0);

materialLabel = QtWidgets.QLabel(addFerramentasFrame)
materialLabel.setText('Material')
materialLabel.move(10, 310)
material = QtWidgets.QLineEdit(addFerramentasFrame)
material.setGeometry(QtCore.QRect(10, 330, 231, 25))
material.setAttribute(QtCore.Qt.WA_MacShowFocusRect,0);

tamanhoLabel = QtWidgets.QLabel(addFerramentasFrame)
tamanhoLabel.setText('Tamanho')
tamanhoLabel.move(10, 370)
tamanho = QtWidgets.QLineEdit(addFerramentasFrame)
tamanho.setGeometry(QtCore.QRect(10, 390, 80, 25))
tamanho.setAttribute(QtCore.Qt.WA_MacShowFocusRect,0);

UCLabel = QtWidgets.QLabel(addFerramentasFrame)
UCLabel.setText('UC')
UCLabel.move(100, 370)
UC = QtWidgets.QLineEdit(addFerramentasFrame)
UC.setGeometry(QtCore.QRect(100, 390, 80, 25))
UC.setAttribute(QtCore.Qt.WA_MacShowFocusRect,0);

tempoReservaLabel = QtWidgets.QLabel(addFerramentasFrame)
tempoReservaLabel.setText('Tempo (dias)')
tempoReservaLabel.move(190, 370)
tempoReserva = QtWidgets.QLineEdit(addFerramentasFrame)
tempoReserva.setGeometry(QtCore.QRect(190, 390, 80, 25))
tempoReserva.setAttribute(QtCore.Qt.WA_MacShowFocusRect,0);

btnAddFerramenta = QtWidgets.QPushButton(addFerramentasFrame)
btnAddFerramenta.setText('Adicionar/Editar')
btnAddFerramenta.setGeometry(QtCore.QRect(10, 440, 231, 25))
btnAddFerramenta.clicked.connect(banco.addFerramenta)

btnDelFerramenta = QtWidgets.QPushButton(addFerramentasFrame)
btnDelFerramenta.setText('Remover')
btnDelFerramenta.setGeometry(QtCore.QRect(10, 480, 231, 25))
btnDelFerramenta.clicked.connect(banco.removeFerramenta)

ferramentasTable = QtWidgets.QTreeView(addFerramentasFrame)
ferramentasTable.setGeometry(QtCore.QRect(280, 30, 600, 500))
ferramentasTable.clicked.connect(treeMediaClicked)

model = PandasModel(banco.ferramentas)
ferramentasTable.setModel(model)


# ---- addTecnico ----
addTecnicoFrame = QtWidgets.QFrame(window)
addTecnicoFrame.setFrameShape(QtWidgets.QFrame.StyledPanel)
addTecnicoFrame.move(0, 40)

cpfLabel = QtWidgets.QLabel(addTecnicoFrame)
cpfLabel.setText('CPF')
cpfLabel.move(10, 10)
cpf = QtWidgets.QLineEdit(addTecnicoFrame)
cpf.setGeometry(QtCore.QRect(10, 30, 130, 25))
cpf.setAttribute(QtCore.Qt.WA_MacShowFocusRect,0);

turnoLabel = QtWidgets.QLabel(addTecnicoFrame)
turnoLabel.setText('Turno')
turnoLabel.move(155, 10)
turno = QtWidgets.QComboBox(addTecnicoFrame)
turno.setGeometry(QtCore.QRect(150, 32, 100, 25))
turno.addItems(['Manhã', 'Tarde', 'Noite'])

nomeLabel = QtWidgets.QLabel(addTecnicoFrame)
nomeLabel.setText('Nome Completo')
nomeLabel.move(10, 70)
nome = QtWidgets.QLineEdit(addTecnicoFrame)
nome.setGeometry(QtCore.QRect(10, 90, 231, 25))
nome.setMaxLength(40)
nome.setAttribute(QtCore.Qt.WA_MacShowFocusRect,0);

tel1Label = QtWidgets.QLabel(addTecnicoFrame)
tel1Label.setText('Contato')
tel1Label.move(10, 130)
tel1 = QtWidgets.QComboBox(addTecnicoFrame)
tel1.setGeometry(QtCore.QRect(5, 152, 100, 25))
tel1.addItems(['Celular', 'Rádio'])

tel2Label = QtWidgets.QLabel(addTecnicoFrame)
tel2Label.setText('Número')
tel2Label.move(110, 130)
tel2 = QtWidgets.QLineEdit(addTecnicoFrame)
tel2.setGeometry(QtCore.QRect(110, 150, 130, 25))
tel2.setMaxLength(9)
tel2.setAttribute(QtCore.Qt.WA_MacShowFocusRect,0);

equipeLabel = QtWidgets.QLabel(addTecnicoFrame)
equipeLabel.setText('Equipe')
equipeLabel.move(10, 190)
equipe = QtWidgets.QLineEdit(addTecnicoFrame)
equipe.setGeometry(QtCore.QRect(10, 210, 231, 25))
equipe.setMaxLength(30)
equipe.setAttribute(QtCore.Qt.WA_MacShowFocusRect,0);

btnAddTecnico = QtWidgets.QPushButton(addTecnicoFrame)
btnAddTecnico.setText('Adicionar/Editar')
btnAddTecnico.setGeometry(QtCore.QRect(10, 260, 231, 25))
btnAddTecnico.clicked.connect(banco.addTecnico)

btnDelTecnico = QtWidgets.QPushButton(addTecnicoFrame)
btnDelTecnico.setText('Remover')
btnDelTecnico.setGeometry(QtCore.QRect(10, 300, 231, 25))
btnDelTecnico.clicked.connect(banco.removeTecnico)

tecnicoTable = QtWidgets.QTreeView(addTecnicoFrame)
tecnicoTable.setGeometry(QtCore.QRect(280, 30, 600, 500))
tecnicoTable.clicked.connect(treeMediaClicked)

model = PandasModel(banco.tecnicos)
tecnicoTable.setModel(model)

window.show()
mostrarFerramentas()

app.exec()


