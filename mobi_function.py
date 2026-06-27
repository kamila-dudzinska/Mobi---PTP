# -*- coding: utf-8 -*-
"""
Created on Fri Jun 26 19:51:32 2026

@author: Kamila Dudzińska
"""

#funkcja zabezpiecza dzielenie przez zero 
def divide_z(a, b, default=0):
    """
    Robimy funkcję dzielenia z zabezpieczeniem dzielenia przez zero.
    
    """
    try:
        # Sprawdzenie typu danych
        if not isinstance(a, (int, float)) or not isinstance(b, (int, float)):
            raise TypeError("Oba argumenty muszą być liczbami.")
        
        return a / b
    except ZeroDivisionError:
        return default
    except TypeError as e:
        print(f"Błąd: {e}")
        return default
    