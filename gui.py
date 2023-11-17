"""
Author: Marcos Sanson
Year: 2023
Project Title: PROYECTO DE INFORMACIÓN METEOROLÓGICA

Overview:
This application allows users to process meteorological data and create graphs
based on selected CSV and ODS files. It provides features for dark mode toggling,
language switching between English and Spanish, file browsing, ODS file creation,
and graph generation for a specific year.

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

import tkinter as tk
from tkinter import filedialog, ttk
import calendar
import locale
import pandas as pd
import pyexcel
import pyexcel_io
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

LIGHT_STYLE = {
    "bg": "white",
    "fg": "black",
    "highlightbackground": "white",
    "highlightcolor": "white",
}

DARK_STYLE = {
    "bg": "#2E2E2E",
    "fg": "white",
    "highlightbackground": "#2E2E2E",
    "highlightcolor": "#2E2E2E",
}

DEFAULT_BG_COLOR = "white"
def toggle_style(dark_mode):
    """Toggle the GUI style between dark mode and light mode.

    This function adjusts the Tkinter root window's palette to switch between dark mode and light mode.
    It sets background colors and other style attributes based on the provided `dark_mode` flag.

    Args:
        dark_mode (bool): A boolean flag indicating whether dark mode should be enabled.

    Returns:
        None
    """
    style = DARK_STYLE if dark_mode else LIGHT_STYLE
    root.tk_setPalette(background=style.get("bg", DEFAULT_BG_COLOR), **style)


class MeteorologicalApp:
    """GUI application/blueprint for meteorological data processing and graph creation."""
    def __init__(self, master):
        """Initialize the MeteorologicalApp.

        Args:
            master (Tk): The Tkinter master window.

        Attributes:
            master (Tk): The Tkinter master window.
            current_language (str): The current language of the application ('english' or 'spanish').
            dark_mode (bool): The current state of the dark mode (True for enabled, False for disabled).
            english_content (dict): English content for the application, including titles, labels, and instructions.
            spanish_content (dict): Spanish content for the application, including titles, labels, and instructions.
            dark_mode_button (ttk.Button): Button to toggle dark mode.
            language_button (ttk.Button): Button to switch between English and Spanish languages.
            title_label (tk.Label): Label displaying the title of the application.
            instruction_label (tk.Label): Label displaying instructional text.
            label_create_ods (tk.Label): Label for the section to create an ODS file.
            label_input (tk.Label): Label for selecting an input CSV file.
            input_button (tk.Button): Button to browse and select an input CSV file.
            input_file_label (tk.Label): Label displaying the selected input CSV file.
            create_ods_button (tk.Button): Button to create an ODS file.
            label_create_graph (tk.Label): Label for the section to create a graph.
            label_output (tk.Label): Label for selecting an input ODS file for creating a graph.
            output_button (tk.Button): Button to browse and select an input ODS file.
            output_file_label (tk.Label): Label displaying the selected input ODS file for the graph.
            label_year (tk.Label): Label for entering the year for graph creation.
            year_entry (tk.Entry): Entry widget for entering the year.
            create_graph_button (tk.Button): Button to create a graph.
            message_label (tk.Label): Label displaying messages in the lower left corner.
            info_label (tk.Label): Label displaying information about the graph.
            output_ods_file: The selected ODS file for graph creation (initialized to None).

        Returns:
            None
        """
        self.master = master
        self.current_language = 'english'  # Default language
        self.dark_mode = False  # Initialize dark mode state

        # English content
        self.english_content = {
            'title': "Meteorological Data Application",
            'dark_mode_button': "Toggle Dark Mode",
            'language_button': "Cambiar Idioma\n "
                               "a Español/Castellano",
            'instructions': "Instructions:\n"
                            "1. Select a CSV input file (resultadoCSV_Columnas).\n"
                            "2. Click 'Create ODS File' to save the reorganized data into a readable spreadsheet.\n"
                            "3. Click 'Create Graph' to generate a graph using the ODS file.",
            'create_ods_section': "Create an ODS File (Spreadsheet File With Data):",
            'select_csv_label': "Select an Input CSV File:",
            'browse_button': "Browse",
            'selected_csv_label': "Selected CSV Input File: ",
            'create_ods_button': "Create ODS File",
            'create_graph_section': "Create a Graph:",
            'select_ods_label': "Select an Input ODS File (to create a graph):",
            'selected_ods_label': "Selected ODS Input File (this file will be used to generate a graph): ",
            'enter_year_label': "Enter Year:",
            'create_graph_button': "Create Graph",
            'graph_information': "Graph Information:\n"
                                 "The generated graph represents meteorological data for the selected year.\n"
                                 "You can customize the appearance of the graph using the matplotlib interface.\n"
                                 "To save the graph as an image file (e.g., PNG, JPEG), you can use the 'Save'\n"
                                 "option in the matplotlib plot window that appears when the graph is generated.\n"
                                 "Do not generate more than one graph at a time. If necessary, you can save a\n"
                                 "graph as an image, close it, then create a new graph of a different year.",
            'message_label': "Message",
            'author_label': "Developed by Marcos Sanson - 2023",
        }

        # Spanish content
        self.spanish_content = {
            'title': "Aplicación de Datos Meteorológicos",
            'dark_mode_button': "Alternar Modo Oscuro",
            'language_button': "Switch Language\n"
                               "to English",
            'instructions': "Instrucciones:\n"
                            "1. Seleccione un archivo CSV de entrada (resultadoCSV_Columnas).\n"
                            "2. Haga clic en 'Crear archivo ODS' para guardar los datos reorganizados en un formato legible.\n"
                            "3. Haga clic en 'Crear gráfica' para crear un gráfico usando el archivo ODS.",
            'create_ods_section': "Crear un archivo ODS (archivo en un formato legible):",
            'select_csv_label': "Seleccione un archivo CSV de entrada:",
            'browse_button': "Buscar",
            'selected_csv_label': "Archivo CSV de entrada seleccionado:",
            'create_ods_button': "Crear archivo ODS",
            'create_graph_section': "Crear una gráfica:",
            'select_ods_label': "Seleccione un archivo ODS de entrada (para crear una gráfica):",
            'selected_ods_label': "Archivo ODS de entrada seleccionado (este archivo se usará para generar una gráfica): N/A",
            'enter_year_label': "Ingresa el año:",
            'create_graph_button': "Crear gráfica",
            'graph_information': "Información de la gráfica:\n"
                                 "La gráfica generada representa datos meteorológicos para el año seleccionado.\n"
                                 "Puede personalizar la apariencia de la gráfica usando la interfaz de matplotlib.\n"
                                 "Para guardar la gráfica como un archivo de imagen (por ejemplo, PNG, JPEG), puede usar\n"
                                 "la opción 'Guardar' en la ventana de la gráfica de matplotlib que aparece cuando se genera la gráfica.\n"
                                 "No cree más de una gráfica a la vez. Si es necesario, puede guardar una gráfica como\n"
                                 "imagen, cerrarla y luego crear una nueva gráfica de un año diferente.",
            'message_label': "Mensaje",
            'author_label': "Desarrollado por Marcos Sanson - 2023",
        }

        # Add a label for displaying me (the author) and year of application development
        self.author_label = tk.Label(master, text=self.get_text('author_label'), font=("Helvetica", 10))
        self.author_label.pack(side='right', anchor='se', padx=10, pady=10)

        master.title(self.get_text('title'))
        master.geometry("1250x700")

        # Add a dark mode toggle button
        self.dark_mode_button = ttk.Button(master, text=self.get_text('dark_mode_button'), command=self.toggle_dark_mode)
        self.dark_mode_button.pack(side='right', anchor='nw', padx=10, pady=10)

        # Add a language switch button
        self.language_button = ttk.Button(master, text=self.get_text('language_button'), command=self.switch_language)
        self.language_button.pack(side='left', anchor='nw', padx=10, pady=10)

        # Add a title label
        self.title_label = tk.Label(master, text=self.get_text('title'), font=("Helvetica", 16))
        self.title_label.pack(pady=10)

        # Add an instruction section
        self.instruction_label = tk.Label(master, text=self.get_text('instructions'), anchor='w', justify='left', font=("Helvetica", 12))
        self.instruction_label.pack(anchor='w', pady=10, padx=10)

        # Section for creating an ODS file
        self.label_create_ods = tk.Label(master, text=self.get_text('create_ods_section'), font=("Helvetica", 14, "bold"))
        self.label_create_ods.pack(anchor='w', pady=10, padx=10)

        self.label_input = tk.Label(master, text=self.get_text('select_csv_label'))
        self.label_input.pack(anchor='w', pady=2, padx=10)

        self.input_button = tk.Button(master, text=self.get_text('browse_button'), command=self.browse_input)
        self.input_button.pack(anchor='w', padx=10)

        self.input_file_label = tk.Label(master, text=self.get_text('selected_csv_label'), anchor='w')
        self.input_file_label.pack(anchor='w', pady=10, padx=10)

        self.create_ods_button = tk.Button(master, text=self.get_text('create_ods_button'), command=self.create_ods)
        self.create_ods_button.pack(anchor='w', pady=5, padx=10)

        # Section for creating a graph
        self.label_create_graph = tk.Label(master, text=self.get_text('create_graph_section'), font=("Helvetica", 14, "bold"))
        self.label_create_graph.pack(anchor='w', pady=10, padx=10)

        self.label_output = tk.Label(master, text=self.get_text('select_ods_label'))
        self.label_output.pack(anchor='w', pady=2, padx=10)

        self.output_button = tk.Button(master, text=self.get_text('browse_button'), command=self.browse_output)
        self.output_button.pack(anchor='w', padx=10)

        self.output_file_label = tk.Label(master, text=self.get_text('selected_ods_label'), anchor='w')
        self.output_file_label.pack(anchor='w', pady=5, padx=10)

        self.label_year = tk.Label(master, text=self.get_text('enter_year_label'))
        self.label_year.pack(anchor='w', pady=2, padx=10)

        self.year_entry = tk.Entry(master)
        self.year_entry.pack(anchor='w', pady=2, padx=10)

        self.create_graph_button = tk.Button(master, text=self.get_text('create_graph_button'), command=self.create_gui_graph)
        self.create_graph_button.pack(anchor='w', pady=10, padx=10)

        # Message label in the lower left
        self.message_label = tk.Label(master, text="", fg="green", font=("Helvetica", 14))
        self.message_label.pack(side='left', anchor='sw', padx=10, pady=10)

        self.output_ods_file = None  # Add this line to initialize the variable

        # Hide the unnecessary output file selection section
        self.label_output.pack_forget()  # Hide the label
        self.output_button.pack_forget()  # Hide the Browse button

        # Information section after year input
        self.info_label = tk.Label(master, text=self.get_text('graph_information'), anchor='w', justify='left', font=("Helvetica", 12))
        self.info_label.pack(anchor='w', pady=10)

    def toggle_dark_mode(self):
        """Toggle dark mode and update the UI."""
        # Toggle dark mode and update the UI
        self.dark_mode = not self.dark_mode
        self.update_dark_mode()

    def update_dark_mode(self):
        """Update UI elements based on dark mode state."""
        # Update UI elements based on dark mode state
        if self.dark_mode:
            # Dark mode settings
            style = {'background': '#2d2d2d', 'foreground': 'white'}
        else:
            # Light mode settings
            style = {'background': 'white', 'foreground': 'black'}

        # Apply style to the root window
        self.master.tk_setPalette(**style)

    def switch_language(self):
        """Toggle between English and Spanish and update the UI."""
        # Toggle between English and Spanish
        if self.current_language == 'english':
            self.current_language = 'spanish'
        else:
            self.current_language = 'english'

        # Update the UI with the new language
        self.update_language()

    def update_language(self):
        """Update UI elements based on the current language."""
        self.title_label.config(text=self.get_text('title'))
        self.dark_mode_button.config(text=self.get_text('dark_mode_button'))
        self.language_button.config(text=self.get_text('language_button'))
        self.instruction_label.config(text=self.get_text('instructions'))
        self.label_create_ods.config(text=self.get_text('create_ods_section'))
        self.label_input.config(text=self.get_text('select_csv_label'))
        self.input_button.config(text=self.get_text('browse_button'))
        self.input_file_label.config(text=self.get_text('selected_csv_label'))
        self.create_ods_button.config(text=self.get_text('create_ods_button'))
        self.label_create_graph.config(text=self.get_text('create_graph_section'))
        self.label_output.config(text=self.get_text('select_ods_label'))
        self.output_button.config(text=self.get_text('browse_button'))
        self.output_file_label.config(text=self.get_text('selected_ods_label'))
        self.label_year.config(text=self.get_text('enter_year_label'))
        self.create_graph_button.config(text=self.get_text('create_graph_button'))
        self.info_label.config(text=self.get_text('graph_information'))
        self.author_label.config(text=self.get_text('author_label'))

    def get_text(self, key):
        """Get text based on the current language and key.

        Args:
            key (str): The key for the desired text.

        Returns:
            str: The text corresponding to the key.
        """
        # Get the text based on the current language and key
        if self.current_language == 'english':
            return self.english_content.get(key, f'{key}')
        elif self.current_language == 'spanish':
            return self.spanish_content.get(key, f'{key}')

    def browse_input(self):
        """Open a file dialog to select an input CSV file."""
        self.input_file = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        # Update the input file label
        self.input_file_label.config(text=f"{self.get_text('selected_csv_label')}: {self.input_file}")

    def browse_output(self):
        """Open a file dialog to select an output ODS file for graph creation."""
        self.output_file = filedialog.asksaveasfilename(defaultextension=".ods", filetypes=[("ODS files", "*.ods")])
        # Update the output file label
        self.output_file_label.config(text=f"{self.get_text('selected_ods_label')}: {self.output_file}")

    def create_ods(self):
        """Create an ODS file based on the selected input CSV file."""
        input_file = getattr(self, "input_file", None)

        if not input_file:
            self.show_message(f"Error: {self.get_text('select_input_file')}", "red")
            return

        # Automatically prompt the user to save a new ODS file
        output_file = filedialog.asksaveasfilename(defaultextension=".ods", filetypes=[("ODS files", "*.ods")])

        # Update the output file label
        self.output_file_label.config(text=f"{self.get_text('selected_ods_label')}: {output_file}")

        if not output_file:
            # User canceled the file save operation
            return

        # Continue with the script execution
        reorganize_ods(input_file, output_file)
        self.output_ods_file = output_file  # Update the instance variable
        self.show_message(self.get_text('ods_created'), "green")

    def create_gui_graph(self):
        """Create a graph based on the selected input CSV and ODS files."""
        input_file = getattr(self, "input_file", None)
        output_file = self.output_ods_file  # Use the stored path

        if not input_file:
            self.show_message(f"Error: {self.get_text('select_ods_file')}", "red")
            return

        if not output_file:
            self.show_message(f"Error: {self.get_text('ods_not_found')}", "red")
            return

        # Check if the year is a valid integer
        try:
            year = int(self.year_entry.get())
        except ValueError:
            self.show_message(f"Error: {self.get_text('enter_valid_year')}", "red")
            return

        df_reorganized = reorganize_ods(input_file, output_file)

        # Check if the selected year is in the DataFrame
        if year not in df_reorganized.index:
            self.show_message(f"Error: {self.get_text('year_not_found')} {year}", "red")
            return

        # Continue with the script execution
        create_graph(input_file, output_file, year)
        self.show_message(self.get_text('graph_created'), "green")

    def show_message(self, message, color):
        """Display a message in the UI.

        Args:
            message (str): The message to be displayed.
            color (str): The color of the message text.
        """
        # Temporarily forget the message label to prevent shifting
        self.message_label.pack_forget()

        # Update the message and color
        self.message_label.config(text=f"{self.get_text('message_label')}: {message}", fg=color)

        # Re-display the message label
        self.message_label.pack(side='left', anchor='sw', padx=10, pady=10)

# Create the main application window
root = tk.Tk()

# Create an instance of the MeteorologicalApp class
app = MeteorologicalApp(root)

# Start the Tkinter event loop
root.mainloop()
