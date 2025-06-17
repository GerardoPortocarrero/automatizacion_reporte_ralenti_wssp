import importlib
import graphics as myg

# REPORTES
def report_configuration(project_address, df, report):
    # CONFIGURACION GENERAL
    bar_width = 10
    bar_height = 7
    bar_label = 17
    bar_fontsize = 21
    bar_color = '#E41A1C'

    # CONFIGURACION DE GRAFICO
    group_by = report
    indicator = 'TIEMPO RALENTI.'
    title = group_by.upper()

    myg.barh_graphic_v2(
        df,
        group_by,
        indicator,
        bar_width,
        bar_height,
        bar_label,
        bar_fontsize,
        bar_color
    )

# Funcion principal
def main(project_address, df, document):
    importlib.reload(myg)
    
    # Calling graphic generators
    for report in document['reports']:
        report_configuration(project_address, df, report)