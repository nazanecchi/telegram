import time
from config import *
import telebot
from comandos import *
from flask import Flask, request
from pyngrok import ngrok, conf

web_server = Flask(__name__)

@web_server.route('/', methods = ['POST'])
def webhook():
    if(request.headers.get("content-type") == "application/json"):
        update = telebot.types.Update.de_json(request.stream.read().decode('utf-8'))
        bot.process_new_updates([update])
        return "OK", 200

@bot.message_handler(commands=['agregar'])
def handle_agregar_usuario(message):
    agregar(message)

@bot.message_handler(commands=['help'])
def handle_agregar_usuario(message):
    help(message)

@bot.message_handler(commands=['eliminar'])
def handle_agregar_usuario(message):
    eliminar(message)
    
@bot.message_handler(commands=['usuarios'])
def handle_agregar_usuario(message):
    usuarios(message)

# Manejador para documentos
@bot.message_handler(content_types=['document'])
def manejar_documento(message):
    transformar_documento(message)

if __name__ == '__main__':
    print("Iniciando bot")
    conf.get_default().config_path = "./config_ngrok.yml"
    conf.get_default().region = "sa"
    ngrok.set_auth_token(ngrok_token)
    ngrok_tunel = ngrok.connect(5000, bind_tls=True)
    ngrok_url = ngrok_tunel.public_url
    print("URL NRGOK: ", ngrok_url)
    bot.remove_webhook()
    time.sleep(1)
    bot.set_webhook(url = ngrok_url)
    web_server.run(host="0.0.0.0", port=5000)
