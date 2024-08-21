import pandas as pd
from datetime import datetime, timedelta

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
class Persona:
    def __init__(self):
        self.nivelLinea = ""
        self.nivelSupervisor = ""
        self.id = ""
        self.nombre = ""
        self.apellido = ""
        self.nivelEquipo = ""
        self.volumenTotal = []
        self.vp = []
        self.volumenCompradoPersonalmente = []
        self.royalities = []
        self.correoElectronico = ""
        self.pais = ""
        self.patrocinador = ""
        self.cuentaLineaDescendenteVisible = ""
        self.metodoCalificacionSP = ""
        self.idPatrocinador = ""
        self.nombreLocalizado = ""
        self.apellidoLocalizado = ""
        self.nombreConyugue = ""
        self.totalVT = ""
        self.pvTotal = ""
        self.ppvTotal = ""
        self.volumenTotalLineaDescendente = ""
        self.totalRO = ""
        self.vt = []  # Agregado
        self.vap = [] # Agregado
        self.vld = [] # Agregado
        self.awt = ""
        self.recalificado = ""
        self.fechaCalificacionSupervisor = ""
        self.fechaCalificacionProductor = ""
        self.fechaAniversario = ""
        self.fechaSolicitud = ""
        self.añosHerbalife = ""
        self.cumpleanos = ""
        self.fechaLimiteRenovacion = ""
        self.domicilio = ""
        self.ciudad = ""
        self.provincia = ""
        self.codigoPostal = ""
        self.telefono1 = ""
        self.telefono2 = ""
        self.etiquetas = []
        self.siglas = ""
        self.numeroPVV = []


def cargarSubject(persona = Persona()):
    subject = "Patrocinador: " + persona.patrocinador 
    subject = subject + "\nID: " + persona.id 
    subject = subject + "\nNivel de Linea: " + persona.nivelLinea
    subject = subject + "\nPais: " + persona.pais
    subject = subject + "\nNivel: " + persona.nivelEquipo + " (" + persona.siglas + ")"
    subject = subject + "\nDia de Registro: " + datetime.strptime(persona.fechaSolicitud, "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y") + " (" + persona.añosHerbalife + " años en Herbalife)"
    subject = subject + "\nRenovacion: " + datetime.strptime(persona.fechaLimiteRenovacion, "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y")
    if datetime.strptime(persona.fechaLimiteRenovacion, "%Y-%m-%d %H:%M:%S") < datetime.now():
         subject = subject + " (licencia expirada)"
    subject = subject + "\nCumpleaños: " + datetime.strptime(persona.cumpleanos, "%Y-%m-%d %H:%M:%S").strftime("%d/%m")
    subject = subject + ("\nFecha de calificación a supervisor: " + datetime.strptime(persona.fechaCalificacionSupervisor, "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y") if persona.fechaCalificacionSupervisor != 'nan' else "")
    subject = subject + ("\nFecha de calificación a productor calificado: " + datetime.strptime(persona.fechaCalificacionProductor, "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y") if persona.fechaCalificacionProductor != 'nan' else "")
    subject = subject + ("\nCalificación supervisor: " + persona.metodoCalificacionSP if persona.metodoCalificacionSP != "No disponible" else "") 
    subject = subject + ("\nHa recalificado supervisor: " + persona.recalificado if persona.recalificado != "nan" else "") 
    subject = subject + ("\nHa recalificado AWT: " + persona.awt if persona.awt != "nan" else "") 
    subject = subject + ("\nNombre del cónyuge: "  + persona.nombreConyugue if persona.nombreConyugue != "  " else "") 
    subject = subject + "\nVolumen Comprado Personalmente (PPV): "
    for i in persona.volumenCompradoPersonalmente:
        subject = subject + "\n       * " + i
    subject = subject + "\nEtiquetas:"
    for i in persona.etiquetas:
        subject = subject + "\n   * " + i
    return subject

def cargar_csv(persona = Persona()):
    notes = cargarSubject(persona)
    row = {
            "Name": persona.nombre + cargar_emojis(persona) + " " + persona.apellido + " (" + persona.siglas + "-" + persona.nivelLinea + ")" + " " + datetime.strptime(persona.fechaSolicitud, "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y"),
            "Given Name": persona.nombre,
            "Additional Name": "",
            "Family Name": persona.apellido + " (" + persona.siglas + "-" + persona.nivelLinea + ")" + " " + datetime.strptime(persona.fechaSolicitud, "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y"),
            "Yomi Name": "",
            "Given Name Yomi": "",
            "Additional Name Yomi": "",
            "Family Name Yomi": "",
            "Name Prefix": "",
            "Name Suffix": "",
            "Initials": "",
            "Nickname": "",
            "Short Name": "",
            "Maiden Name": "",
            "Birthday": datetime.strptime(persona.cumpleanos, "%Y-%m-%d %H:%M:%S").strftime("--%m-%d"),
            "Gender": "",
            "Location": persona.pais,
            "Billing Information": "",
            "Directory Server": "",
            "Mileage": "",
            "Occupation": "",
            "Hobby": "",
            "Sensitivity": "",
            "Priority": "",
            "Subject": "",
            "Notes": notes,
            "Group Membership": "* My Contacts",
            "E-mail 1 - Type": "* Home",
            "E-mail 1 - Value": persona.correoElectronico,
            "E-mail 2 - Type": "",
            "E-mail 2 - Value": "",
            "IM 1 - Type": "",
            "IM 1 - Service": "",
            "IM 1 - Value": "",
            "Phone 1 - Type": ("Personal" if persona.telefono1 != "nan" else ""),
            "Phone 1 - Value": (persona.telefono1 if persona.telefono1 != "nan" else ""),
            "Phone 2 - Type": ("Work" if persona.telefono2 != "nan" else ""),
            "Phone 2 - Value": (persona.telefono2 if persona.telefono2 != "nan" else ""),
            "Phone 3 - Type": "",
            "Phone 3 - Value": "",
            "Phone 4 - Type": "",
            "Phone 4 - Value": "",
            "Address 1 - Type": "Home",
            "Address 1 - Formatted": persona.domicilio + "\n" + persona.ciudad + "\n" + persona.provincia + "\n" + persona.codigoPostal,
            "Address 1 - Street": persona.domicilio,
            "Address 1 - City": persona.ciudad,
            "Address 1 - PO Box": "",
            "Address 1 - Region": persona.provincia,
            "Address 1 - Postal Code": persona.codigoPostal,
            "Address 1 - Country": persona.pais,
            "Address 1 - Extended Address": "",
            "Organization 1 - Type": "",
            "Organization 1 - Name": "",
            "Organization 1 - Yomi Name": "",
            "Organization 1 - Title": persona.nivelEquipo,
            "Organization 1 - Department": "",
            "Organization 1 - Symbol": "",
            "Organization 1 - Location": "",
            "Organization 1 - Job Description": "",
            "Website 1 - Type": "",
            "Website 1 - Value": ""
        }
    return row

def cargar_etiquetas(persona = Persona()):
    etiquetas = []
    con = 0
    etiquetas.append("members20#")
    if datetime.strptime(persona.cumpleanos, "%Y-%m-%d %H:%M:%S").month == datetime.now().month:
         etiquetas.append("birthday")
    if datetime.strptime(persona.fechaLimiteRenovacion, "%Y-%m-%d %H:%M:%S") < datetime.now():
         etiquetas.append("sin_licencia")
    elif datetime.strptime(persona.fechaLimiteRenovacion, "%Y-%m-%d %H:%M:%S") - timedelta(months=3) > datetime.now():
         etiquetas.append("renovar_licencia")
    for i in persona.numeroPVV:
        if i == "0":
            con = con + 1
        else:
            break
    if con > 0:
        etiquetas.append(str(con) + "_mes_sin_pedir")
    return etiquetas


def cargar_emojis(persona = Persona()):
    emojis = "💚"
    for i in persona.etiquetas:
        if i == "sin_licencia":
            emojis = emojis + "❌"
        elif i.endswith("_mes_sin_pedir"):
            num = i.replace("_mes_sin_pedir", "")
            for j in num:
                emojis = emojis + calcNum(j)
        elif i == "birthday":
            emojis = emojis + "🎂"
        elif i == "renovar_licencia":
            emojis = emojis + "❗"
    return emojis

def calcNum(j):
    match j:
        case "0":
            return "0️⃣"
        case "1":
            return "1️⃣"
        case "2":
            return "2️⃣"
        case "3":
            return "3️⃣"
        case "4":
            return "4️⃣"
        case "5":
            return "5️⃣"
        case "6":
            return "6️⃣"
        case "7":
            return "7️⃣"
        case "8":
            return "8️⃣"
        case "9":
            return "9️⃣"
        
def cargarsiglas(nombre = ""):
    nombre = nombre.upper()
    match nombre:
        case "EQUIPO MILLONARIO":
            return "NR"
        case "DISTRIBUIDOR":
            return "DS"
        case "CONSULTOR SENIOR":
            return "CS"
        case "PRODUCTOR CALIFICADO":
            return "QP"
        case "SUPERVISOR":
            return "SP"
        case "EQUIPO MUNDIAL":
            return "WT"
        case "EQUIPO MUNDIAL ACTIVO":
            return "WTA"
        case "GET":
            return "GET"
        case "GET2500":
            return "GET2500"
        case "MILLONARIO TEAM":
            return "MT"
        case "MILLONARIO TEAM 7500":
            return "MT7500"
        case "PRESIDENTE":
            return "PT"
        case _:
            return "??"

def agregar_persona(title, fil):
    persona = Persona()
    for i, celd in enumerate(title):
        if celd == "Nivel de la línea":
            persona.nivelLinea = fil[i] + "L"
        elif celd == "Nivel de Supervisor":
            persona.nivelSupervisor = fil[i]
        elif celd == "ID":
            persona.id = fil[i]
        elif celd == "Nombre":
            persona.nombre = fil[i].title()
        elif celd == "Apellidos":
            persona.apellido = fil[i].title()
        elif celd == "Nivel del Equipo":
            persona.nivelEquipo = fil[i]
            persona.siglas = cargarsiglas(fil[i])
        elif celd.endswith("Volumen Total"):
            date = celd.replace("  Volumen Total", "")
            persona.volumenTotal.append(date + ": " + fil[i])
        elif celd.endswith("VP"):
            date = celd.replace("  VP", "")
            persona.vp.append(date + ": " + fil[i])
        elif celd.endswith("Volumen Comprado Personalmente"):
            date = celd.replace("  Volumen Comprado Personalmente", "")
            persona.volumenCompradoPersonalmente.append(date + ": " + fil[i])
            persona.numeroPVV.append(fil[i])
        elif celd.endswith("Royalties"):
            date = celd.replace("  Royalties", "")
            persona.royalities.append(date + ": " + fil[i])
        elif celd == "Correo electrónico":
            persona.correoElectronico = fil[i].lower()
        elif celd == "País":
            persona.pais = fil[i]
        elif celd == "Patrocinador":
            persona.patrocinador = fil[i].title()
        elif celd == "Cuenta de la Línea Descendente Visible":
            persona.cuentaLineaDescendenteVisible = fil[i]
        elif celd == "Método de Calificación de SP":
            persona.metodoCalificacionSP = fil[i]
        elif celd == "Número ID del Patrocinador":
            persona.idPatrocinador = fil[i]
        elif celd == "Nombre (Localizado)":
            persona.nombreLocalizado = fil[i].title()
        elif celd == "Apellidos (Localizados)":
            persona.apellidoLocalizado = fil[i].title()
        elif celd == "Nombre del cónyuge":
            persona.nombreConyugue = fil[i].title()
        elif celd == "Total VT":
            persona.totalVT = fil[i]
        elif celd == "PV Total":
            persona.pvTotal = fil[i]
        elif celd == "PPV Total":
            persona.ppvTotal = fil[i]
        elif celd == "Volumen Total de Línea Descendente":
            persona.volumenTotalLineaDescendente = fil[i]
        elif celd == "Total RO":
            persona.totalRO = fil[i]
        elif celd.endswith(" VT (Volumen documentado/para calificar)"): #desde aca
            date = celd.replace(" VT (Volumen documentado/para calificar)", "")
            persona.vt.append(date + ": " + fil[i])
        elif celd.endswith(" VP (Volumen documentado/para calificar)"):
            date = celd.replace(" VP (Volumen documentado/para calificar)", "")
            persona.vp.append(date + ": " + fil[i])
        elif celd.endswith(" VAP (Volumen documentado/para calificar)"):
            date = celd.replace(" VAP (Volumen documentado/para calificar)", "")
            persona.vap.append(date + ": " + fil[i])
        elif celd.endswith(" VLD (Volumen documentado/para calificar)"):
            date = celd.replace(" VLD (Volumen documentado/para calificar)", "")
            persona.vld.append(date + ": " + fil[i])
        elif celd == "Calificante a AWT (año fiscal actual)":
            persona.awt = fil[i]
        elif celd == "¿Ha Recalificado?":
            persona.recalificado = fil[i]
        elif celd == "Fecha de Calificación a Supervisor":
            persona.fechaCalificacionSupervisor = fil[i]
        elif celd == "Fecha de Calificación de Productor Calificado":
            persona.fechaCalificacionProductor = fil[i]
        elif celd == "Fecha de Aniversario":
            persona.fechaAniversario = fil[i]
        elif celd == "Fecha de la Solicitud":
            persona.fechaSolicitud = fil[i]
        elif celd == "Años en Herbalife":
            persona.añosHerbalife = fil[i]
        elif celd == "Cumpleaños":
            persona.cumpleanos = fil[i]
        elif celd == "Fecha límite de renovación":
            persona.fechaLimiteRenovacion = fil[i]
        elif celd == "Domicilio 1":
            persona.domicilio = fil[i].title()
        elif celd == "Ciudad":
            persona.ciudad = fil[i].title()
        elif celd == "Provincia":
            persona.provincia = fil[i].title()
        elif celd == "Código Postal":
            persona.codigoPostal = fil[i]
        elif celd == "Teléfono preferente":
            persona.telefono1 = fil[i]
        elif celd == "Teléfono Tardes":
            persona.telefono2 = fil[i]
    persona.etiquetas = cargar_etiquetas(persona)
    return cargar_csv(persona)

def excel_sheet_to_csv(excel_file, sheet_name, file_path):
    
    # Guardar el encabezado y el DataFrame como archivo CSV
    df_csv = pd.DataFrame(columns=header1)
    # Cargar la hoja específica del archivo de Excel en un DataFrame
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    df = df.iloc[7:-1]
    df = df.astype(str)

    # Eliminar la segunda línea del DataFrame
    title = df.iloc[0].tolist()
    matrix = df.iloc[1:].values.tolist()
    rows = []
    for i in matrix:
        rows.append(agregar_persona(title, i))
        df_csv = pd.DataFrame(rows, columns=header1)
        df_csv.to_csv(file_path, index=False)


    # Guardar el DataFrame como archivo CSV en el path especificado
def toCsv(excel_file, file_path):
    # Nombre del archivo de Excel y CSV
    sheet_name = 'Report'  # Nombre de la segunda hoja o índice de la hoja (0-indexed)
    
    # Llamar a la función para convertir la segunda hoja
    excel_sheet_to_csv(excel_file, sheet_name, file_path)
