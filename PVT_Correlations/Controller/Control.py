# %% Import libraries and functions
import xlwings as xw
import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from PVT_Correlations.Model.Functions import Bo
from PVT_Correlations.Model.Functions import Pb
from PVT_Correlations.Model.Functions import Rs
from PVT_Correlations.Model.Functions import uo

# %% Create sheet, variable and data names

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
VALUES = "valores"
CORRELACION_S = "Standing"
CORRELACION_AL = "AL_Marhoun"
CORRELACION_B = "Beggs & Robinson"
CORRELACION_G = "Glaso"


# %% P ###
def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[SHEET_SUMMARY]

    df_TVD = sheet[STOC_VALUES].options(pd.DataFrame, index=False, expand="table").value
    input_col_names = df_TVD["Valores"].to_list()
    Rs_value, Yo_value, Yg_value, T_value, P_value, API_value, BASIO, BASIO2 = tuple(
        input_col_names
    )
    inpt_idx = [CORRELACION_S, CORRELACION_AL]
    result_Bo = {}
    result_Pb = {}
    result_Rs = {}
    for col in inpt_idx:
        print(col)
        result_Bo[col] = Bo(col, Rs_value, Yg_value, Yo_value, T_value)
        result_Pb[col] = Pb(col, Rs_value, Yg_value, T_value, API_value, Yo_value)
        result_Rs[col] = Rs(col, P_value, API_value, T_value, Yg_value, Yo_value)
    inpt_idx2 = [CORRELACION_B, CORRELACION_G]
    result_uo = {}

    for col2 in inpt_idx2:
        result_uo[col2] = uo(col2, API_value, T_value)

    PVT_summary_results = [
        result_Bo[CORRELACION_S],
        result_Bo[CORRELACION_AL],
        result_Pb[CORRELACION_S],
        result_Pb[CORRELACION_AL],
        result_Rs[CORRELACION_S],
        result_Rs[CORRELACION_AL],
        result_uo[CORRELACION_B],
        result_uo[CORRELACION_G],
    ]
    sheet[BO_STANDING].options(transpose=True).value = PVT_summary_results
    print(PVT_summary_results)

    for col, idx in inpt_idx2:
        results_dict[col] = Bo("Standing", "AL_Marhoun")

    # Resultado del Bo por Standing & AL-Marhoun:
    sheet[BO_STANDING].value = results_Bo[CORRELACION_S]  # Bo("Standing", Rs_value, Yg_value, Yo_value, T_value)
    sheet[BO_AL_MARHOUN].value = results_Bo[CORRELACION_AL]  # Bo("AL_MARHOUN", Rs_value, Yg_value, Yo_value, T_value)

    # Resultado de Pb por Standing & AL-Marhoun:
    sheet[PB_STANDING].value = results_Pb[
        CORRELACION_S]  # Pb("Standing", Rs_value, Yg_value, T_value, API_value, Yo_value)
    sheet[PB_AL_MARHOUN].value = results_Pb[
        CORRELACION_AL]  # Pb("AL_MARHOUN", Rs_value, Yg_value, T_value, API_value, Yo_value)

    # Resultado del Rs por Standing & AL-Marhoun:
    sheet[RS_STANDING].value = results_Rs[
        CORRELACION_S]  # Rs("Standing", P_value, API_value, T_value, Yg_value, Yo_value)
    sheet[RS_AL_MARHOUN].value = results_Rs[
        CORRELACION_AL]  # Rs("AL_MARHOUN", P_value, API_value, T_value, Yg_value, Yo_value)

    # Caluclo de la uo por Beal & Glaso
    sheet[UO_BEAL].value = uo("Beal", API_value, T_value)
    sheet[UO_GLASO].value = uo("Glaso", API_value, T_value)

    PVT_summary_result = [
        results_Bo[CORRELACION_S],
        results_Bo[CORRELACION_AL],
        results_Pb[CORRELACION_S],
        results_Pb[CORRELACION_AL],
        results_Rs[CORRELACION_S],
        results_Rs[CORRELACION_AL],
        results_uo[CORRELACION_B],
        results_uo[CORRELACION_G],
    ]
    sheet[BO_STANDING].options(transpose=True).value = PVT_summary_result
    print(PVT_summary_result)

    if __name__ == "__main__":
        xw.Book("Control.xlsm").set_mock_caller()
        main()

    # %% Graficas Bo
    df_resultsBo = pd.DataFrame(results_Bo)

    eng_formatter = ticker.EngFormatter()
    sns.set_style("white")
    fig = plt.figure(figsize=(8, 6))
    ax = sns.histplot(df_resultsBo[CORRELACION_S], color="lightgray", kde=True)
    ax.xaxis.set_major_formatter(eng_formatter)
    plot = sheet.pictures.add(fig, name="Histograma", left=sheet.range("J1").left)
    print(plot)


if __name__ == "__main__":
    xw.Book("Control.xlsm").set_mock_caller()
    main()
