from kiwoom.kiwoom import *

import sys
from PyQt5.QtWidgets import *

class Ui_class():
    def __init__(self):
        print('ui class')
    
        self.app = QApplication(sys.argv)
        
        self.kiwoom = Kiwoom() 
        
        self.app.exec_() #프로그램이 끝나지 않고 계속 돌아가게 해주는 코드
    
    