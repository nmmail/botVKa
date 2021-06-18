import vk_api
from vk_api.longpoll import VkLongPoll, VkEventType
from vk_api.keyboard import VkKeyboard
from toks import main_token
import openpyxl

my_path = "123.xlsx"
my_wb_obj = openpyxl.load_workbook(my_path)
my_sheet_obj = my_wb_obj.active
string = ''

vk_session = vk_api.VkApi(token=main_token)
session_api = vk_session.get_api()
longpoll = VkLongPoll(vk_session)

def sender(id, text) :
    vk_session.method('messages.send', {'user_id': id, 'message': text, 'random_id': 0})

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
                for cell in my_sheet_obj['A']:
                    if str(cell.value) == str(msg):
                        for row in range(1,8):
                            string = string + sender(id, str(row))
                        sender(id, string)
                        string = ''
            else :
                sender(id , 'Группа не найдена.Повторите')







