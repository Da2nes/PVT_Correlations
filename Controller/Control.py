#%% Import libraries and functions
import xlwings as xw
import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from PVT_Correlations.Model.Functions import Bo
from PVT_Correlations.Model.Functions import Pb
from PVT_Correlations.Model.Functions import Rs
from PVT_Correlations.Model.Functions import vo

#%% Create sheet, variable and data names

# Names for sheets
SHEET_SUMMARY = "Datos"
SHEET_RESULTS = "Resultados"

# Name of columns for distribution definitions
VARIABLES = "Variables"
VALORES = "Valores"
PARAMETROS = "Parametros"
CORRELACION = "Correlacion"

# Name of Data
STOC_VALUES = "df_bo_calculator"
# Result cells # Call range cells from MS Excel
BO_STANDING = "Bo_Standing"
BO_AL_MARHOUN = "Bo_Al_Marhoun"
RS_STANDING = "Rs_Standing"
RS_AL_MARHOUN = "Rs_Al_Marhoun"
PB_STANDING = "Pb_Standing"
PB_AL_MARHOUN = "Pb_Al_Marhoun"
UO_BEAL = "uo_Beal"
UO_GLASO = "uo_Glaso"
VALUES = "Valores"
CORRELACION_S = "Standing"
CORRELACION_AL = "AL_Marhoun"
CORRELACION_B = "Beggs & Robinson"
CORRELACION_G = "Glaso"


