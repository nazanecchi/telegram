import io
import os
import uuid
from config import *
from xlsToCsv.xlsToCsv import toCsv
import shutil

def usuariosConPermisos():
    with open('users.txt', 'r') as file:
            lines = file.readlines()
        
        # Eliminar los saltos de línea (\n) al final de cada línea
    lines = [line.strip() for line in lines]
        
    return lines

def soySuperUsuario(message):
    if message.from_user.username == "nazarenonecchi" or message.from_user.username == "Mario_Infantehb" or message.from_user.username == "fanymario":
        return True
    else:
        return False
    
def eliminar(message):
    if soySuperUsuario(message):
        usuario = message.text.split('/eliminar ')[1]
        if usuario[0] == '@':
                usuario = message.text.split('@')[1]
        print("/" + usuario + "/")
        if usuario in usuariosConPermisos():
            if any(char.isspace() for char in usuario):
                bot.reply_to(message, "Formato de usuario invalido")
                return
            try:
                with open('users.txt', 'r') as file:
                    content = file.read()
                new_content = content.replace(usuario+"\n", '')
                with open('users.txt', 'w') as file:
                    file.write(new_content)
                bot.reply_to(message, "Usuario eliminado")
            except:
                bot.reply_to(message, "Error en la eliminacion, intentelo mas tarde")
        else:
            bot.reply_to(message, "Ese usuario no tiene permisos")


def agregar(message):
    if soySuperUsuario(message):
        usuario = message.text.split('/agregar ')[1]
        if usuario[0] == '@':
                usuario = message.text.split('@')[1]
        print("/" + usuario + "/")
        if not  usuario in usuariosConPermisos():
            if any(char.isspace() for char in usuario):
                bot.reply_to(message, "Formato de usuario invalido")
                return
            try:
                with open('users.txt', 'a') as file:
                    file.write(usuario + '\n')
                bot.reply_to(message, "Usuario agregado")
            except:
                bot.reply_to(message, "Error en la insercion, intentelo mas tarde")
        else:
            bot.reply_to(message, "Este usuario ya tiene permisos")

def usuarios(message):
    if soySuperUsuario(message):
        try:
            with open('users.txt', 'r') as file:
                lines = file.read()
            bot.reply_to(message, lines)
        except:
            bot.reply_to(message, "No hay usuarios con permisos")


def help(message):
    if soySuperUsuario(message):
        bot.reply_to(message, "/agregar nombreUsuario\n/eliminar nombreUsuario\n/usuarios")

def es_archivo_excel(file_name):
    # Lista de extensiones de archivo que corresponden a formatos de Excel
    extensiones_excel = ['.xls', '.xlsx', '.xlsm', '.xlsb', '.odf', '.ods', '.odt']
    
    # Verifica si la extensión del archivo está en la lista
    return any(file_name.lower().endswith(ext) for ext in extensiones_excel)

def transformar_documento(message):
    if (message.from_user.username in usuariosConPermisos()) or soySuperUsuario(message):
        file_name = message.document.file_name
        
        if es_archivo_excel(file_name):
            file_name = file_name.split('.')[0]+".csv"
            file_info = bot.get_file(message.document.file_id)
            downloaded_file = bot.download_file(file_info.file_path)
            file_stream = io.BytesIO(downloaded_file)
            try:
                folder_id = uuid.uuid4().hex
                folder_path = os.path.join('temp_files', folder_id)
                file_path = os.path.join(folder_path, file_name)
                os.makedirs(folder_path, exist_ok=True)
                toCsv(file_stream, file_path)
                with open(file_path, 'rb') as file:
                    bot.send_document(message.chat.id, file,  reply_to_message_id=message.message_id)
                shutil.rmtree(folder_path)
            except Exception as e:
                print(e)
                bot.reply_to(message, "Archivo invalido")
            
        else:
            bot.reply_to(message, "Este archivo no es un archivo de Excel.")