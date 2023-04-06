import subprocess
import time
# import webbrowser
import telebot
import os
from ctypes import cast, POINTER
import ctypes
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
# from pynput.keyboard import Key, Controller

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

ctypes.windll.kernel32.SetConsoleTitleW("King of bots")

re_start = ["shutdown", "-f", "-r", "-t", "30"]
shut_down = ["shutdown", "-f", "-s", "-t", "30"]


def shutdown(self):
    import subprocess
    subprocess.call(shut_down)


def restart(self):
    import subprocess
    subprocess.call(re_start)


token = '6176587339:AAEoi7OpgOFgwZXqjSCpZun-uyCzaEhvAq4'
bot = telebot.TeleBot(token)
chat_id = '-1001820758497'

'''text = "*Привет2*"
msg = bot.send_message(chat_id, text, parse_mode='MarkdownV2')
msg_id = msg.message_id'''

doc = 'C:/Users/user190717/Desktop/заказы.xlsx'
doc1 = 'C:/Users/user190717/Desktop/долги.xlsx'
doc2 = 'C:/Users/user190717/Desktop/приход.pdf'
doc3 = 'C:/Users/user190717/Desktop/пополнение товара.xlsx'


@bot.message_handler(content_types=['text'])
def get_text_messages(message):
    if message.text == "заказы" or message.text == "Заказы":
        bot.send_document(chat_id=chat_id, document=open(doc, 'rb'))

        '''elif 'громкость' in message.text or 'Громкость' in message.text:
            if '100' in message.text:
                volume.SetMasterVolumeLevel(-0.0, None)
            elif '50' in message.text:
                volume.SetMasterVolumeLevel(-10.0, None)
            elif '70' in message.text:
                volume.SetMasterVolumeLevel(-5.0, None)
            elif '30' in message.text:
                volume.SetMasterVolumeLevel(-18.0, None)
            elif '0.0' in message.text:
                volume.SetMasterVolumeLevel(-60.0, None)'''

    elif message.text == "1с" or message.text == "1С":
        os.startfile(r'C:/Program Files (x86)/1cv82/common/1cestart.exe')

        '''elif message.text == "лол" or message.text == "Лол":
            keyboard = Controller()
            keyboard.press(Key.cmd)
            keyboard.press(Key.home)
            keyboard.release(Key.home)
            keyboard.release(Key.cmd)
            os.startfile(r'C:/Users/user190717/Desktop/файлы/VID_20230110_000245_285.mp4')'''

    elif message.text == "закрыть" or message.text == "Закрыть":
        subprocess.call(["taskkill", "/F", "/IM", "1cv8.exe"])
        subprocess.call(["taskkill", "/F", "/IM", "1cv8s.exe"])

        '''elif message.text == "закрыть браузер" or message.text == "Закрыть браузер":
            subprocess.call(["taskkill", "/F", "/IM", "chrome.exe"])
    
        elif message.text == "открыть браузер" or message.text == "Открыть браузер":
            os.startfile(r'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe')
    
        elif message.text == "яндекс" or message.text == "Яндекс":
            webbrowser.open('https://ya.ru')'''

    elif message.text == "выкл" or message.text == "Выкл":
        shutdown(shut_down)

    elif message.text == "перезагрузка" or message.text == "Перезагрузка":
        restart(re_start)

    elif message.text == "товары" or message.text == "Товары":
        bot.send_document(chat_id=chat_id, document=open(doc3, 'rb'))

    elif message.text == "фото" or message.text == 'Фото':
        bot.send_photo(chat_id=chat_id, photo='https://traveltimes.ru/wp-content/uploads/2021/05/204_mini.jpg')

    elif message.text == "долги" or message.text == 'Долги':
        bot.send_document(chat_id=chat_id, document=open(doc1, 'rb'))

    elif message.text == "приход" or message.text == 'Приход':
        bot.send_document(chat_id=chat_id, document=open(doc2, 'rb'))

    elif message.text == "/help@chirag700bot":
        m1 = bot.send_message(chat_id=chat_id, text='*Напиши \'заказы\', \'долги\', \'приход\', \'товары\''
                                               ' или нажми /orders, /loans, /sells, /refill'
                                               ' чтобы выгрузить данные\.\n'
                                               'Нажми /search для поиска\.*', parse_mode='MarkdownV2')
        ms_id = m1.message_id
        time.sleep(7)
        bot.delete_message(chat_id=chat_id, message_id=ms_id)

    elif message.text == "/search@chirag700bot":
        m = bot.send_message(chat_id=chat_id, text='*Нажми:\n'
                                               '\t\t            \#заказ\n\n'
                                               '\t\t            \#долг\n\n'
                                               '\t\t            \#уддолг\n\n'
                                               '\t\t            \#списание*', parse_mode='MarkdownV2')
        ms_id = m.message_id
        time.sleep(7)
        bot.delete_message(chat_id=chat_id, message_id=ms_id)

    elif message.text == "/loans@chirag700bot":
        bot.send_document(chat_id=chat_id, document=open(doc1, 'rb'))

    elif message.text == "/orders@chirag700bot":
        bot.send_document(chat_id=chat_id, document=open(doc, 'rb'))

    elif message.text == "/refill@chirag700bot":
        bot.send_document(chat_id=chat_id, document=open(doc3, 'rb'))

    elif message.text == "/sells@chirag700bot":
        bot.send_document(chat_id=chat_id, document=open(doc2, 'rb'))

    elif message.text == "/shutdown@chirag700bot":
        shutdown(shut_down)

    elif message.text == "/close1c@chirag700bot":
        subprocess.call(["taskkill", "/F", "/IM", "1cv8.exe"])
        subprocess.call(["taskkill", "/F", "/IM", "1cv8s.exe"])

    else:
        bot.send_message(chat_id=chat_id, text="Я тебя не понимаю. Напиши /help.")


@bot.message_handler(content_types=['document'])
def handle_docs_photo(message):
    file_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_info.file_path)

    src = 'C:/Users/user190717/Desktop/' + message.document.file_name
    with open(src, 'wb') as new_file:
        new_file.write(downloaded_file)

    bot.reply_to(message, "Пожалуй, я сохраню это")


@bot.message_handler(content_types=['audio'])
def audio_processing(message):
    file_info = bot.get_file(message.audio.file_id)
    downloaded_file = bot.download_file(file_info.file_path)
    src = 'C:/Users/user190717/Desktop/' + str(message.audio.file_name) + '.mp3'

    with open(src, 'wb') as new_file:
        new_file.write(downloaded_file)

    bot.reply_to(message, "Пожалуй, я сохраню это")


@bot.message_handler(content_types=['photo'])
def download_image(message):
    fileID = message.photo[-1].file_id
    file_info = bot.get_file(fileID)
    downloaded_file = bot.download_file(file_info.file_path)

    src = 'C:/Users/user190717/Desktop/' + f'Photo{message.id}.jpg'

    with open(src, 'wb') as new_file:
        new_file.write(downloaded_file)

    bot.reply_to(message, "Пожалуй, я сохраню это")


bot.polling(none_stop=True, interval=0)
