import pandas as pd

# Eliminar columnas duplicadas
def delete_duplicated_columns(df):
    return df.loc[:, ~df.columns.duplicated(keep='first')]

# Rellenar los valores faltantes en la columna LOCACION
def full_fill_data(df):
    df["LOCACION"] = df["LOCACION"].ffill()
    return df

# Eliminar todas las columnas excepto las relevantes
def get_relevant_columns(df, file):
    return df[file['relevant_columns']]

# Ajustar valores
def adjust_values(df):
    df = df.fillna(0)
    return df

# Eliminar locaciones 'SIN GRUPO'
def delete_sin_grupo_locations(df):
    return df[df['LOCACION'] != 'SIN GRUPO']

# Convertir tiempo a deltatime
def make_date_deltatime(df):
    df['TIEMPO RALENTI.'] = pd.to_timedelta(df['TIEMPO RALENTI.'])
    return df

# Conservar las filas con las siguientes locaciones
def get_specific_location(df, locacion):
    df = df[df['LOCACION'] == locacion]
    return df

# Preparar archivo para analizar
def file_processing(df, file, locacion):
    df = delete_duplicated_columns(df)

    df = get_relevant_columns(df, file)

    df = full_fill_data(df)

    df = adjust_values(df)

    df = delete_sin_grupo_locations(df)

    df = make_date_deltatime(df)

    df = get_specific_location(df, locacion)

    return df