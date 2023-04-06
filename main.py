import ctypes
'''import cryptocode
from past.builtins import raw_input'''
from functions import add_func, search_func, sub_func, clean_func, \
    del_func, bar_func, edit_func, sort_func, ord_func, orc_func, \
    del_orc_func, all_orc_func, start_bot, end_bot, end_app  # path_passwd
import py_win_keyboard_layout as pw
import keyboard
from telebot.apihelper import ApiTelegramException
import time

pw.change_foreground_window_keyboard_layout(0x00000419)
ctypes.windll.kernel32.SetConsoleTitleW("CHIRAG LOANS & ORDERS")

'''while True:
    with open(path_passwd, 'r') as f:
        password = f.read()
        dec_passwd = cryptocode.decrypt(password, 'wow')

    p = raw_input(' Пароль: ')

    if p == '0':
        print('\n' + ' Выход...')
        break

    if p == str(dec_passwd):
        break
    else:
        print(' Неверный пароль! Попробуйте снова')
'''
while True:
    try:
        '''if p == '0':
            break'''
        print('\n' + '\t' + '<<====================| ГЛАВНАЯ |====================>>')
        print('\n'
              '        \'1\' ДОЛГИ\n'
              '        \'2\' ЗАКАЗЫ\n'
              '    ---------------------\n'
              '        \'и\' ИЗМЕНИТЬ\n'
              '        \'о\' ОЧИСТИТЬ\n'
              '    ---------------------\n'
              '        \'д\' ДОБАВИТЬ\n'
              '        \'у\' УДАЛИТЬ\n'
              '        \'т\' ТАБЛИЦА')

        while True:
            if keyboard.is_pressed('q'):
                print(' <--')
                break
            if keyboard.is_pressed('1'):
                while True:
                    if keyboard.is_pressed('q'):
                        break
                    print('\n' + '\t' + '<<====================| ДОЛГИ |====================>>')
                    print('\n'
                          '        \'3\' ДОБАВИТЬ\n'
                          '        \'4\' ПОИСК\n'
                          '        \'5\' СОРТИРОВКА\n'
                          '        \'с\' СПИСАТЬ\n'
                          '        \'7\' УДАЛИТЬ\n'
                          '        \'г\' ГРАФИК\n'
                          '        \'й + з\' НАЗАД')
                    if keyboard.is_pressed('q'):
                        break
                    while True:
                        if keyboard.is_pressed('q'):
                            break
                        if keyboard.is_pressed('3'):
                            print('\n' + '\t' + '<<====================| Добавление |====================>>\n')
                            add_func()
                            break
                        if keyboard.is_pressed('4') or keyboard.is_pressed('ctrl + f'):
                            print('\n' + '\t' + '<<====================| Поиск |====================>>\n')
                            search_func()
                            break
                        if keyboard.is_pressed('5'):
                            print('\n' + '\t' + '<<====================| Сортировка |====================>>\n')
                            sort_func()
                            break
                        if keyboard.is_pressed('c'):
                            print('\n' + '\t' + '<<====================| Списание |====================>>\n')
                            sub_func()
                            break
                        if keyboard.is_pressed('7'):
                            print('\n' + '\t' + '<<====================| Удаление |====================>>\n')
                            del_func()
                            break
                        if keyboard.is_pressed('u'):
                            print('\n' + '\t' + '<<====================| График |====================>>\n')
                            bar_func()
                            break

            if keyboard.is_pressed('j'):
                clean_func()
                break
            if keyboard.is_pressed('b'):
                print('\n' + '\t' + '<<====================| Изменение |====================>>\n')
                edit_func()
                break
            if keyboard.is_pressed('2'):
                ord_func()
                break
            if keyboard.is_pressed("l"):
                print('\n' + '\t' + '<<====================| Добавление товара |====================>>\n')
                orc_func()
                break
            if keyboard.is_pressed("n"):
                print('\n' + '\t' + '<<=================| Таблица товаров для пополнения |================>>\n')
                all_orc_func()
                break
            if keyboard.is_pressed("e"):
                print('\n' + '\t' + '<<====================| Удаление товара |====================>>\n')
                del_orc_func()
                break
            if keyboard.is_pressed("ctrl + i"):
                start_bot()
                print('\n' + ' Бот запущен!')
                break
            if keyboard.is_pressed("ctrl + m"):
                end_bot()
                print('\n' + ' Бот выключен!')
                break
            if keyboard.is_pressed("ctrl + z"):
                print('\n' + ' Выход из программы...\n')
                time.sleep(1.5)
                end_app()
                break
        print('\n' + ' Нажмите (з) ')
        if keyboard.wait('p'):
            continue

    except ValueError:
        print('\n' + ' ValueError! ПОПРОБУЙТЕ СНОВА')

    except ModuleNotFoundError:
        print('\n' + ' ModuleNotFoundError! ПОПРОБУЙТЕ СНОВА')

    except SyntaxError:
        print('\n' + ' SyntaxError! ПОПРОБУЙТЕ СНОВА')

    except FileExistsError:
        print('\n' + ' FileExistsError! ПОПРОБУЙТЕ СНОВА')

    except FileNotFoundError:
        print('\n' + ' FileNotFoundError! ПОПРОБУЙТЕ СНОВА')

    except UnicodeDecodeError:
        print('\n' + ' UnicodeDecodeError! ПОПРОБУЙТЕ СНОВА')

    except TypeError:
        print('\n' + ' TypeError! ПОПРОБУЙТЕ СНОВА')

    except PermissionError:
        print('\n' + ' PermissionError! ПОПРОБУЙТЕ СНОВА')

    except KeyError:
        print('\n' + ' KeyError! ПОПРОБУЙТЕ СНОВА')

    except AttributeError:
        print('\n' + ' AttributeError! ПОПРОБУЙТЕ СНОВА')

    except ConnectionError:
        print('\n' + ' ConnectionError! ПОПРОБУЙТЕ СНОВА')

    except ApiTelegramException:
        print('\n' + ' ApiTelegramExceptionError! ПОПРОБУЙТЕ СНОВА')

    continue
