#%% HOLIIII

import math

#Funcion de la uo
def uo(colums, API, T):
    """

    Parameters

    colums
        Correlación usada (Beggs y robinson o Glaso)
    API
        Gravedas API
    T
        Temperatura en °R
    Returns
        uo en cp


    """
    z = 3.0324 - (0.02023 * API)
    y = 10**z
    x = (y) * (T-460)**(-1.163)
    a1=(10.313 * math.log10(T-460)) - 36.447
    if colums == "Beggs & Robinson":
        uo= 10**x - 1
    else:
        uo= 3.141 * (10**10) * (T-460)**(-3.414) * (math.log10(API))**a1
    return uo


#Funcion PB
def Pb(colums, Rs, Yg, T, API=None, Yo=None):
    """
    Parameters
    ----------
    colums
        Correlación usada
    Rs
        Solubilidad del gas en PCN/BN
    Yg
        Solubilidad del gas
    T
        Temperatura en °R
    API
        Gravedad API
    Yo
        Gravedad específica del petroleo
    Returns
    Pb en psi
    -------
    """
    a= 0.00091 * (T-460) - 0.0125*(API)
    a1=5.38088 * 10**(-3)
    b=0.715082
    c=-1.87784
    d=3.1437
    e=1.32657
    if colums == "Standing":
        Pb = 18.2 * (((Rs / Yg) ** 0.83) * 10 ** a - 1.4)
    else:
        Pb=(a1 * Rs**b) * (Yg**c) * (Yo**d) * (T**e)
    return  Pb

