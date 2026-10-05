# -*- coding: utf-8 -*-
"""
NRU Generator Plugin - Point d'entrée
"""

def classFactory(iface):
    from .nru_plugin import NRUPlugin
    return NRUPlugin(iface)