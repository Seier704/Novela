# -*- coding: utf-8 -*-
"""
Compilador general de la novela 'El camino del reencarnado'.
Compila todos los capítulos presentes en la carpeta 'capitulos/*.txt'.
"""
import os
import sys

NOVELA_DIR = os.path.dirname(os.path.abspath(__file__))
if NOVELA_DIR not in sys.path:
    sys.path.insert(0, NOVELA_DIR)

from agregar_capitulo import compilar_libro_completo

if __name__ == "__main__":
    compilar_libro_completo(NOVELA_DIR)
