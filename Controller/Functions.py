
#Funcion para encontrar el Rs
def Rs(colums, P, API, T=None, Yg=None, Yo=None):
    """"
    Parameters
    ---------------
    colums:
        Correlación usada
    P:
        Presión del sistema en psi
    API:
        Gravedad API
    T:
        Temperatura del sistema en °R
    Yg:
        Gravedad específica del gas en superficie
    Yo:
        Gravedad específica del crudo en superficie
    Returns
        float number  -> Rs
    """
    x = 0.0125 * API - 0.00091 * (T - 460)
    a = 185.843208
    b = 1.877840
    c = -3.1437
    d = -1.32657
    e = 1.398441
    if colums == "Standing":
        Rs = (Yg) * (((P / 18.2) + 1.4) * (10 ** x)) ** 1.2048
    else:
        Rs = (a * (Yg ** b) * (Yo ** c) * (T ** d) * P) ** e
    return Rs