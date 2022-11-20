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

---------
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