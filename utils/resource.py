import os
import sys

def resource_path(relative_path):
    """Obtem o caminho correto para o recurso em qualquer ambiente (Python ou executável)"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)