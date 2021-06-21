import vk_api
from vk_api.longpoll import VkLongPoll, VkEventType
from vk_api.keyboard import VkKeyboard
from toks import main_token


vk_session = vk_api.VkApi(token=main_token)
session_api = vk_session.get_api()
longpoll = VkLongPoll(vk_session)

import openpyxl

my_path = "123.xlsx"
my_wb_obj = openpyxl.load_workbook(my_path)
my_sheet_obj = my_wb_obj.active
string = ''

from vk_api.keyboard import VkKeyboard, VkKeyboardColor
def sender(id, text, keybord=None) :
    post = {'user_id': id, 'message': text, 'random_id': 0}
    if keybord != None:
        post["keyboard"] = keybord.get_keyboard()
    else:
        post = post
    vk_session.method('messages.send', post)

def readerEx(msg):
    c = 1
    p = 1
    for cell in my_sheet_obj['A']:
        if str(msg) == str(cell.value):
            for c in range(2, 9):
                string = string + str((my_sheet_obj.cell(row=p, column=c)).value) + ' '

            sender(id, string + '')
        string = ''
        p = p + 1
        c = 1
        print(p)
        # this reader


def button():
     keyboard = VkKeyboard(one_time=False)
     buttosn = ["Пондельник", "Вторник", "Среда" , "Четверг" , "Пятница" , "Суббота" ]
     buttosn_color = [VkKeyboardColor.PRIMARY,VkKeyboardColor.PRIMARY,VkKeyboardColor.PRIMARY,VkKeyboardColor.PRIMARY,VkKeyboardColor.PRIMARY,VkKeyboardColor.PRIMARY]

     trg = 0
     for btn, btn_c in zip(buttosn, buttosn_color):
         trg = trg + 1
         keyboard.add_button(btn, btn_c)
         if trg == 2:
             if(btn != 'Суббота'):
                 keyboard.add_line()
             trg = 0
     sender(id, "Выберите день недели :", keyboard)



for event in longpoll.listen():
    if event.type == VkEventType.MESSAGE_NEW:
         if event.to_me:
            msg = event.text.lower()
            id = event.user_id
            flag_f = False
            for cell in my_sheet_obj['A']:
                if str(msg) == str(cell.value):
                    flag_f = True
                    break

            if(flag_f == True):
                 from toks import days
                 button()
                 if(msg == days[0]):
                     sender(id, "Расписание на понедельник :")
                 #readerEx(msg)


            else :
                sender(id , 'Группа не найдена.Повторите')






