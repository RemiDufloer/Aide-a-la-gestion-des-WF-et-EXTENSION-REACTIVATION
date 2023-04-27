# -*- coding: utf-8 -*-
"""
Created on Tue Feb 21 18:02:23 2023

@author: starx
"""

import sys

# choix utilisateur 


def demande_a_utilisateur(message,minimum = None, maximum = None, fd=sys.stdout):
    while True:
        try:
            print(message, file=fd)
            value = int(input(""))
        except ValueError:
            print("Vous devez rentrer un entier", file=fd)
            continue

        if minimum is not None and value < minimum :
            print("Vous devez choisir un nombre supérieur à {0}.".format(minimum), file=fd)
            continue
        elif maximum  is not None and value > maximum :
            print("Vous devez choisir un nombre inférieur à {0}.".format(maximum), file=fd)
            continue
        else:
            break
    return value

    
def demande_a_utilisateur_string(message,minimum = None, maximum = None, fd=sys.stdout):
    while True:
        try:
            print(message, file=fd)
            value = str(input(""))
        except ValueError:
            print("Vous devez rentrer un string", file=fd)
            continue
        break
    return value


# def choix( module, fd=sys.stdout):
#     if module == 1 :
#     elif module == 2 :
        