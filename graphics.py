import seaborn as sns
import pandas as pd
import matplotlib.pyplot as plt
import os

def barh_graphic_v1(
        df, 
        group_by, 
        indicator, 
        date, 
        locacion,
        bar_width,
        bar_height,
        bar_label,
        bar_fontsize, 
        bar_color
    ):
    
    # Convertir tiempo a minutos
    df['Minutos Ralenti'] = df[indicator].dt.total_seconds() / 60

    # Ordenar por tiempo
    df = df.sort_values('Minutos Ralenti')

    # Crear figura y eje
    fig, ax = plt.subplots(figsize=(bar_width, bar_height))

    # Gráfico de barras horizontales
    bars = ax.barh(df[group_by], df['Minutos Ralenti'], color=bar_color)

    # Etiquetas sobre las barras
    for bar in bars:
        width = bar.get_width()
        ax.text(width + 0.5, bar.get_y() + bar.get_height() / 2,
                f"{width:.1f} min", va='center', fontsize=bar_fontsize, color='black')

    # Estética de ejes y etiquetas
    ax.set_xlabel("")  # Sin etiqueta
    ax.set_ylabel("")  # Sin etiqueta
    ax.set_xticks(range(0, int(df['Minutos Ralenti'].max()) + 2, 2))
    ax.tick_params(axis='x', labelsize=bar_label)
    ax.tick_params(axis='y', labelsize=bar_label)
    ax.grid(axis='x', linestyle='--', alpha=0.3)

    # ✅ Eliminar bordes del gráfico
    for spine in ax.spines.values():
        spine.set_visible(False)

    # ✅ Título visualmente centrado (ajustado a ojo si es necesario)
    fig.text(0.5, 0.96, f"RALENTÍ {locacion} {date}",
            ha='center', va='top', fontsize=bar_fontsize)

    # Deja espacio arriba para el título
    plt.tight_layout(rect=[0, 0, 1, 0.93])

    plt.show()

def barh_graphic_v2(
        project_address,
        df, 
        group_by, 
        indicator,
        bar_width,
        bar_height,
        bar_label,
        bar_fontsize, 
        bar_color
    ):
    
    # Convertir tiempo a minutos
    df['Minutos Ralenti'] = df[indicator].dt.total_seconds() / 60

    # Ordenar por tiempo
    df = df.sort_values('Minutos Ralenti', ascending=True)

    # Crear figura y eje
    fig, ax = plt.subplots(figsize=(bar_width, bar_height))

    # Gráfico de barras horizontales
    bars = ax.barh(df[group_by], df['Minutos Ralenti'], color=bar_color)

    # Etiquetas sobre o dentro de las barras
    for bar in bars:
        width = bar.get_width()
        # Si la barra es suficientemente ancha, poner el texto dentro, en blanco
        if width > 4:
            ax.text(width - 0.5, bar.get_y() + bar.get_height() / 2,
                    f"{width:.1f} min", va='center', ha='right',
                    fontsize=bar_fontsize, color='white')
        else:
            ax.text(width + 0.3, bar.get_y() + bar.get_height() / 2,
                    f"{width:.1f} min", va='center', ha='left',
                    fontsize=bar_fontsize, color='black')

    # Estética general
    ax.set_xlabel("")
    ax.set_ylabel("")
    ax.set_xticks(range(0, int(df['Minutos Ralenti'].max()) + 2, 2))
    ax.tick_params(axis='x', labelsize=bar_label)
    ax.tick_params(axis='y', labelsize=bar_label)
    ax.grid(axis='x', linestyle='--', alpha=0.3)

    # Eliminar todos los bordes del gráfico
    for spine in ax.spines.values():
        spine.set_visible(False)

    # Ajuste para espacio del título
    plt.tight_layout(rect=[0, 0, 1, 0.92])

    # Nombre de archivo
    file_name = f'{group_by}.png'

    # Guardar en alta calidad (opcional)
    plt.savefig(os.path.join(project_address, file_name), dpi=300, bbox_inches='tight', facecolor='white')
    #plt.show()