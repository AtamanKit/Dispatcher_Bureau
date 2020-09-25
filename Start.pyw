from PySide2.QtGui import QPixmap
from PySide2.QtWidgets import QSplashScreen, QApplication, QMainWindow
import sys, subprocess
import threading
import time


def splashFunc():
    app = QApplication(sys.argv)
    # global exeControl = False
    pixmap = QPixmap("ataman_splash.png")
    global splash
    splash = QSplashScreen(pixmap)
    splash.show()
    splash.showMessage("Loading...")
    app.exec_()

def splashClose():
    time.sleep(30)
    splash.close()

def subprossFunc():
    subprocess.run('Biroul_dispecerului_nosplash.exe')
    splash.close()

threading.Thread(target=splashFunc).start()
threading.Thread(target=splashClose).start()
threading.Thread(target=subprossFunc).start()