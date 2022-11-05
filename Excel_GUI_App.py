import ctypes
import openpyxl
import time as ti
from openpyxl.utils import get_column_letter
from sys                  import(argv, exit)
from PyQt5.QtWidgets      import(QGridLayout, QFrame, QLabel, QPushButton, QLineEdit , QFileDialog)
from PyQt5.Qt             import(QApplication, QMainWindow)
from PyQt5.QtCore         import(Qt)
from PyQt5.QtGui          import(QFont)


class PROJECT(QMainWindow):

	def __init__(self, Developer):
		QMainWindow.__init__(self)

		self.Developer = Developer

		self.screen_width  = ctypes.windll.user32.GetSystemMetrics(0)
		self.screen_height = ctypes.windll.user32.GetSystemMetrics(1)
		self.width  = 1400
		self.height = 900

		self.setWindowTitle('PROJECT')
		#self.setWindowIcon(QIcon('path\\here'))         ### -> to set window icon
		#self.setWindowFlags(Qt.WindowMinimizeButtonHint | Qt.WindowCloseButtonHint)          ### -> search about window flags but not neccessary now
		self.setGeometry(int((self.screen_width-self.width)/2),            # -> window x position
                         int((self.screen_height-self.height)/2),          # -> window y position
                         self.width,                                       # -> window width
                         self.height)                                      # -> window height
		#self.setFixedSize(self.width, self.height)         ### -> to keep the size fixed
		self.GUI()
		#self.Connect()
		self.Open_File()

	def GUI(self):
		self.frame = QFrame(self)
		self.frame.setFixedSize(self.width,self.height)
		self.grid = QGridLayout(self.frame)
		#self.OpenFile = QPushButton("Choose File")
		#self.grid.addWidget(self.OpenFile,0,0,Qt.AlignCenter)




		#self.OpenFile.setFixedSize(120,30)



	#def Connect(self):
		#self.OpenFile.clicked.connect(lambda : self.Open_File())
		


	def Open_File(self):
		fileName = QFileDialog.getOpenFileName(self,str("Open Excel File"),"", str("Excel Files (*.xlsx)"))[0]
		if fileName == '': exit()
		self.wb = openpyxl.load_workbook(fileName)
		self.wb.active
		self.OpenSheetWindow()
			

	def OpenSheetWindow(self):
		pass



if __name__ == '__main__':
	root = QApplication(argv)
	root.setStyle('Fusion')
	
	MyProject = PROJECT("Abdo Alshoki")
	MyProject.show()

	exit(root.exec_())