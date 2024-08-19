import pandas as pd

def agregar_persona(fil):
    idx = fil[2].title()
    nombre = fil[3].title()
    apellido = fil[4].title()
    equipo = fil[4].title()
    print(nombre)

def excel_sheet_to_csv(excel_file, sheet_name, csv_file):
    # Cargar la hoja específica del archivo de Excel en un DataFrame
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    df = df.iloc[7:]
    df = df.drop(df.index[-1])

    # Eliminar la segunda línea del DataFrame
    if len(df) > 1:
        df = df.drop(df.index[0])

    matrix = df.values.tolist()
    for i in matrix:
        agregar_persona(i)
    
    
    # Especificar el encabezado
    header1 = ["Name", "Given Name", "Additional Name", "Family Name", "Yomi Name", "Given Name Yomi", "Additional Name Yomi", 
              "Family Name Yomi", "Name Prefix", "Name Suffix", "Initials", "Nickname", "Short Name", "Maiden Name", 
              "Birthday", "Gender", "Location", "Billing Information", "Directory Server", "Mileage", "Occupation", 
              "Hobby", "Sensitivity", "Priority", "Subject", "Notes", "Group Membership", "E-mail 1 - Type", 
              "E-mail 1 - Value", "E-mail 2 - Type", "E-mail 2 - Value", "IM 1 - Type", "IM 1 - Service", 
              "IM 1 - Value", "Phone 1 - Type", "Phone 1 - Value", "Phone 2 - Type", "Phone 2 - Value", 
              "Phone 3 - Type", "Phone 3 - Value", "Phone 4 - Type", "Phone 4 - Value", "Address 1 - Type", 
              "Address 1 - Formatted", "Address 1 - Street", "Address 1 - City", "Address 1 - PO Box", 
              "Address 1 - Region", "Address 1 - Postal Code", "Address 1 - Country", "Address 1 - Extended Address", 
              "Organization 1 - Type", "Organization 1 - Name", "Organization 1 - Yomi Name", "Organization 1 - Title", 
              "Organization 1 - Department", "Organization 1 - Symbol", "Organization 1 - Location", 
              "Organization 1 - Job Description", "Website 1 - Type", "Website 1 - Value"]
    
    # Guardar el encabezado y el DataFrame como archivo CSV
    with open(csv_file, 'w', newline='') as file:
        file.write(",".join(header1) + "\n")
        df.to_csv(file, index=False, header=False)
    
    print(f'Se ha convertido la hoja "{sheet_name}" del archivo "{excel_file}" a "{csv_file}" exitosamente.')

if __name__ == "__main__":
    # Nombre del archivo de Excel y CSV
    excel_file = 'input.xls'
    sheet_name = 'Report'  # Nombre de la segunda hoja o índice de la hoja (0-indexed)
    csv_file = 'output.csv'
    
    # Llamar a la función para convertir la segunda hoja
    excel_sheet_to_csv(excel_file, sheet_name, csv_file)
