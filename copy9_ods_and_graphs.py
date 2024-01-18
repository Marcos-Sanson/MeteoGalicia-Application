"""
Author: Marcos Sanson
Year: 2023
Project Title: PROYECTO DE INFORMACIÓN METEOROLÓGICA

Overview:
This script reorganizes data from a CSV file into an ODS file and creates bar graphs for specific years.

Dependencies:
- pandas
- pyexcel_ods3
- numpy
- matplotlib

Usage:
1. Replace the input_file_path and output_file_path variables in the example usage section.
2. Run the script.

Note: Ensure that the required dependencies are installed before running the script.
"""


import calendar
import locale
import pandas as pd
import pyexcel_ods3 as ods
import numpy as np
import matplotlib.pyplot as plt


def reorganize_ods(input_file, output_file):
    """Reorganize data from a CSV file and save the result to a new ODS file.

    Args:
        input_file (str): The path to the input CSV file.
        output_file (str): The path to save the reorganized ODS file.

    Returns:
        pd.DataFrame: The reorganized data as a DataFrame.
    """
    # Set the language to Spanish
    locale.setlocale(locale.LC_TIME, 'es_ES.utf8')

    # Read the input ODS file
    data = ods.get_data(input_file)

    # Get the data from the first sheet of the input file
    sheet_name = list(data.keys())[0]
    df = pd.DataFrame(data[sheet_name][1:], columns=data[sheet_name][0])

    # Find the date column name based on its data type
    date_column = df.select_dtypes(include=['datetime']).columns[0]

    # Find the data column name based on its position in the DataFrame
    data_column = df.columns[-1]

    # Convert -9999 to NaN (missing value)
    df[data_column] = df[data_column].replace(-9999, float('nan'))

    # Extract the year from the date column
    df['Year'] = pd.to_datetime(df[date_column]).dt.year

    # Calculate the yearly sum excluding missing data and -9999 numbers
    df['Suma Anual'] = df.groupby(df['Year'])[data_column].apply(lambda x: np.sum(x[x != ''] != '-9999'))

    # Calculate the monthly mean excluding missing data and -9999 numbers
    df.loc['Media Mensual'] = df.groupby(pd.to_datetime(df[date_column]).dt.month)[data_column].apply(
        lambda x: np.mean(x[x != ''] != '-9999'))

    # Reorganize the data
    df_reorganized = pd.pivot_table(df, values=data_column, index='Year',
                                     columns=pd.to_datetime(df[date_column]).dt.month)
    df_reorganized = df_reorganized.rename_axis(None, axis=1)

    # Add labels in the first row with the names of the months
    month_names = [calendar.month_name[i].capitalize() for i in range(1, 13)]
    df_reorganized.loc[-1] = month_names
    df_reorganized.sort_index(inplace=True)

    # Add the yearly sum to the right of the spreadsheet with a label
    df_reorganized['Suma Anual'] = df.groupby(df['Year'])[data_column].sum()
    df_reorganized.loc[-1, 'Suma Anual'] = 'Suma Anual'

    # Add the monthly mean at the bottom of each column
    df_reorganized.loc['Media Mensual'] = df.groupby(pd.to_datetime(df[date_column]).dt.month)[data_column].mean()

    # Calculate the mean of the yearly sum data excluding missing data and -9999 numbers
    mean_yearly_sum = df_reorganized.iloc[2:, -1]
    mean_yearly_sum = mean_yearly_sum[mean_yearly_sum != '']
    mean_yearly_sum = mean_yearly_sum[mean_yearly_sum != '-9999']
    mean_yearly_sum = mean_yearly_sum.astype(float).mean()

    # Move the mean value to the bottom of the "Suma Anual" column
    df_reorganized.loc[df_reorganized.index[-1], 'Suma Anual'] = float(mean_yearly_sum)

    # Save the reorganized data to a new ODS file
    reorganized_data = df_reorganized.reset_index().values.tolist()

    # Replace the first cell of the output file with the value from the second cell of the input file
    reorganized_data[0][0] = data[sheet_name][0][1]

    ods.save_data(output_file, {sheet_name: reorganized_data})

    print(f"\nODS file reorganization completed!\n"
          f"Input file: {input_file}\n"
          f"Output file: {output_file}\n")

    return df_reorganized


def create_graph(input_file_path, output_file_path, year):
    """Create and display a bar graph based on reorganized data for a specific year.

    Args:
        input_file_path (str): The path to the input CSV file.
        output_file_path (str): The path to save the reorganized ODS file.
        year (int): The selected year for the graph.

    Returns:
        None
    """
    try:
        # Get the y-axis label from the first cell of the second column of the input file
        data = ods.get_data(input_file_path)
        sheet_name = list(data.keys())[0]
        title_label = data[sheet_name][0][1]  # Assuming the label is in the first row and second column

        df_reorganized = reorganize_ods(input_file_path, output_file_path)

        # Check if the selected year is in the DataFrame
        if year not in df_reorganized.index:
            raise KeyError(f"Year {year} not found in the data.")

        selected_data = df_reorganized.loc[year]
        months = selected_data.index[:-1]
        values = selected_data[:-1].astype(float)

        # Map month numbers to capitalized month names
        month_names = [calendar.month_name[i].capitalize() for i in range(1, 13)]
        months = [month_names[int(month) - 1] for month in months]

        # Replace invalid or missing values with np.nan
        values = np.where(np.isnan(values), np.nan, values)

        # Create an index array for evenly spaced x-axis tick positions
        index = np.arange(len(months))

        bar_width = 0.8  # Adjust the bar width as desired

        plt.bar(index, values, width=bar_width)
        plt.xlabel('Meses')  # Translate 'Months' to 'Meses'

        y_axis_label = title_label
        if y_axis_label == 'Chuvia':
            y_axis_label += ' (lluvia) [litros/m2]'  # Add unit for 'Chuvia'

        plt.ylabel(y_axis_label)  # Set y-axis label

        # Modify the title only if title_label is 'Chuvia'
        if title_label == 'Chuvia':
            plt.title(f'{title_label} (lluvia) en el año {year}')  # Translate and set graph title
        else:
            plt.title(f'{title_label} en el año {year}')  # Translate and set graph title

        plt.xticks(index, months, rotation=45, ha='right')  # Set x-axis ticks and rotate labels

        # Add numbers on top of each bar
        for i, v in enumerate(values):
            if not np.isnan(v):
                plt.text(i, v, str(v), ha='center', va='bottom', fontweight='normal')

        # Adjust the subplot spacing
        plt.subplots_adjust(left=0.125, right=0.9, top=0.9, bottom=0.225)

        # Add padding above each bar to accommodate text labels
        padding_factor = 0.1  # Adjust the padding as desired
        valid_values = values[~np.isnan(values)]
        min_value = np.min(valid_values)
        max_value = np.max(valid_values)
        y_padding = (max_value - min_value) * padding_factor
        plt.ylim(bottom=0, top=max_value + y_padding + np.max(valid_values) * 0.1)

        # Add extra empty space for the first and last months
        padding_months = 0.7  # Adjust the padding as desired
        plt.xlim(left=-padding_months, right=len(months) - 1 + padding_months)

        plt.show()

    except KeyError as e:
        print(f"Error: {e}")
        print("Please select a valid year present in the data.")


# Example usage
input_file_path = 'Número de días de helada_mensuales_test1/resultadoCSV_Columnas.csv'     # Replace with the path to your input ODS file
output_file_path = 'Número de días de helada_mensuales_test1/test1_Número de días de helada_mensuales_Rios_final.ods'     # Replace with the desired path for the output ODS file

input_file_path = 'Chuvia_mensual_test2/resultadoCSV_Columnas.csv'     # Replace with the path to your input ODS file
output_file_path = 'Chuvia_mensual_test2/test3_Chuvia_mensual.ods'     # Replace with the desired path for the output ODS file

reorganize_ods(input_file_path, output_file_path)

print("\nODS file reorganization completed!\n")


# Example usage
input_file_path = 'Chuvia_mensual_test1/resultadoCSV_Columnas.csv'     # Replace with the path to your input ODS file
output_file_path = 'Chuvia_mensual_test1/Chuvia_mensual_Rios_final.ods'     # Replace with the desired path for the output ODS file

# Example usage
input_file_path = 'Horas_de_sol_mensual_test1/resultadoCSV_Columnas.csv'     # Replace with the path to your input ODS file
output_file_path = 'Horas_de_sol_mensual_test1/graph_Horas_de_sol_mensual_final.ods'     # Replace with the desired path for the output ODS file

# Example usage
input_file_path = 'Horas_de_frío_mensuales_test1/resultadoCSV_Columnas.csv'     # Replace with the path to your input ODS file
output_file_path = 'Horas_de_frío_mensuales_test1/graph_Horas_de_frío_mensuales.ods'     # Replace with the desired path for the output ODS file

# Example usage
input_file_path = 'Chuvia_mensual_test2/resultadoCSV_Columnas.csv'
output_file_path = 'Chuvia_mensual_test2/copy4_Chuvia_mensual_Rios_final.ods'


create_graph(input_file_path, output_file_path, year=2023)

