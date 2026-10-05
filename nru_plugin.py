from qgis.PyQt.QtWidgets import QAction
from qgis.PyQt.QtGui import QIcon
import os

from .nru_dialog import NRUDialog


class NRUPlugin:

    def __init__(self, iface):
        self.iface = iface

    def initGui(self):
        self.action = QAction("Générer NRU", self.iface.mainWindow())
        self.action.triggered.connect(self.run)
        self.iface.addPluginToMenu("NRU Generator", self.action)
        self.iface.addToolBarIcon(self.action)

    def unload(self):
        self.iface.removePluginMenu("NRU Generator", self.action)
        self.iface.removeToolBarIcon(self.action)

    def run(self):
        self.dialog = NRUDialog(self.iface)
        self.dialog.exec_()