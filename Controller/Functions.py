

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