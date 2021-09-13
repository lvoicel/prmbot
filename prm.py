import datetime, time, telebot, yadisk, xlrd, xlwt
from telebot import types
from xlutils.copy import copy
bot = telebot.TeleBot('token');
ya = yadisk.YaDisk(token="token")

print("Токен: ",ya.check_token())

class Orders(object):
    def __init__(self, text, works, files):
        self.text = {}
        self.works = {}
        self.files = {}
        
    def set_text(self, text):
        self.text = text


@bot.message_handler(content_types=['text'])
def start(message):
    text_d = get_disk()
    bot.send_message(
        chat_id = message.from_user.id,
        text=text_d['text'],
        reply_markup = menu())
        

def get_disk(): # Получаем заявки
    a = list(ya.listdir("/Производство/ЗАЯВКИ В РАБОТУ/ЗАЯВКИ В РАБОТЕ", fields = ['name', 'type', 'size', 'path']))
    text = ''
    dict = {}
    num = ''
    now = datetime.datetime.today().strftime('%d.%m.%Y')
    text += f'\U0001F4C5 Сегодня {now}\n\n'
    for i in range(len(a)):
        name = a[i].name
        text += '\U0001F539 ' + name + '\n'
        if a[i].type == 'file':
            break
        elif name[name.find("П")+2:name.find("П")+5] != ' - ':
            num = name[name.find("П"):name.find("П")+6]
        elif name[name.find("П")+1].isdigit() != True:
            num = name
        else:
            num = name[name.find("П"):name.find("П")+2] + '-' + name[name.find("П")+5:name.find("П")+8]
        dict[num] = a[i].name
    dict['text'] = text
    return dict
    

def get_order(x): # Получаем файлы из заявки
    dir = list(ya.listdir("/Производство/ЗАЯВКИ В РАБОТУ/ЗАЯВКИ В РАБОТЕ/", fields = ['name', 'path', 'type']))
    order_path = ''
    c = 0
    p = 1
    id = ''
    for i in range(len(dir)):
        if dir[i].type == 'file':
            continue
        if x in dir[i].name:
            order_path = dir[i].path
        elif order_path == '':
            id = x[0:2] + ' ' + '-' + ' ' + x[3:6]
            if id in dir[i].name:
                order_path = dir[i].path
    order = list(ya.listdir(order_path, fields = ['name', 'type', 'path', 'file']))
    for i in range(len(order)):
        print('iter: ', i, order[i].name)
        if order[i].name.find('.xlsx') >= 0:
            break
        elif order[i].name.find('.xls') >= 0 and order[i].name.find('П') >= 0:
            if c > 0:
                break
            c += 1
            print('xls: ', order[i].name)
            ya.download(order[i].path, order[i].name)
            text = parse_xls(order[i].name, order_path)
            files = Orders(text, {}, {})
            files.set_text(text)
        if order[i].name.find('.pdf') >= 0:
            ya.download(order[i].path, order[i].name)
            name = f"pdf{p}"
            files.files[name] = order[i].path
            p += 1
    if c == 0 and p == 1:
        files = Orders({}, {}, {})
        files.set_text({'order': x, 'works': ''})
    return files
            
    
def parse_xls(xls, path):
    result = {}
    works = {}
    print('Parsing: ', xls)
    wb = xlrd.open_workbook(xls)
    sheet = wb.sheet_by_index(0)
    if xls[xls.find("П")+2:xls.find("П")+5] != ' - ':
        result['order'] = xls[xls.find("П"):xls.find("П")+6]
    else: 
        result['order'] = xls[xls.find("П"):xls.find("П")+2] + '-' + xls[xls.find("П")+5:xls.find("П")+8]
    result['manager'] = sheet.cell_value(1, 9)
    wrong_date = sheet.cell_value(5, 6)
    if (isinstance(wrong_date, str)):
        result['date'] = wrong_date
    else:
        y, m, d, h, i, s = xlrd.xldate_as_tuple(wrong_date, wb.datemode)
        result['date'] = "{0}.{1}.{2}".format(d, m, y)  
    result['customer'] = sheet.cell_value(3, 6)
    i = 15
    c = 1
    while True:
        item = []
        if type(sheet.cell_value(i, 7)) != float:
            break
        item.append(int(sheet.cell_value(i, 7)))
        if sheet.cell_value(i, 8) == 'X':
            item.append('8')
        elif sheet.cell_value(i, 9) == 'X':
            item.append('9')
        elif sheet.cell_value(i, 10) == 'X':
            item.append('10')
        elif sheet.cell_value(i, 11) == 'X':
            item.append('11')
        elif sheet.cell_value(i, 12) == 'X':
            item.append('12')
        elif sheet.cell_value(i, 13) == 'X':
            item.append('13')
        item.append('')
        works[sheet.cell_value(i, 1)] = item
        c += 1
        i += 1
    result['works'] = works
    result['xls'] = xls
    result['xls_o'] = path
    return result
    
def write_xls(id, work, order): # Вносим данные в xls
    x = 8
    mark = 'X' 
    # Антон - 1226725096
    if id == 242171872: # Артём
        xf = xlwt.easyxf('align: wrap on, vertical center, horizontal center;'
            'borders: left thin, right thin, top thin, bottom thin;'
            'pattern: pattern solid, pattern_fore_colour yellow, pattern_back_colour yellow'
            )
        x = 8
    # elif id == '':# Руслан
        # xf = xlwt.easyxf('align: wrap on, vertical center, horizontal center;'
            # 'borders: left thin, right thin, top thin, bottom thin;'
            # 'pattern: pattern solid, pattern_fore_colour purple, pattern_back_colour purple'
            # )
        # x = 9
    elif id == 1184599004:# Сергей
        xf = xlwt.easyxf('align: wrap on, vertical center, horizontal center;'
            'borders: left thin, right thin, top thin, bottom thin;'
            'pattern: pattern solid, pattern_fore_colour blue, pattern_back_colour blue'
            )
        x = 10
    # elif id == '':# Андрей
        # xf = xlwt.easyxf('align: wrap on, vertical center, horizontal center;'
            # 'borders: left thin, right thin, top thin, bottom thin;'
            # 'pattern: pattern solid, pattern_fore_colour green, pattern_back_colour green'
            # )
        # x = 11
    elif id == 1371484606:# Максим
        xf = xlwt.easyxf('align: wrap on, vertical center, horizontal center;'
            'borders: left thin, right thin, top thin, bottom thin;'
            'pattern: pattern solid, pattern_fore_colour green, pattern_back_colour green'
            )
        x = 12
    else:
        xf = xlwt.easyxf('align: wrap on, vertical center, horizontal center;'
            'borders: left thin, right thin, top thin, bottom thin;'
            'pattern: pattern solid, pattern_fore_colour white, pattern_back_colour white'
            )
        x = 13
    
    wb = xlrd.open_workbook(order.text['xls'], formatting_info=True)
    xl = copy(wb)
    sheet = xl.get_sheet(0)
    sheet.write(15 + int(work), x, mark, xf)
    i = 15
    c = 1
    xl.save(order.text['xls'])
    ya.upload(order.text['xls'], order.text['xls_o'] + '/' + order.text['xls'], overwrite=True)
    
    
def order_mark(order): # Клавиатура для отметки в заказах
    keyboard = []
    keyrow = []
    c = 0
    r = 0
    for i in range(len(order.text['works'])):
        if c > 6:
            c = 0
            r += 1
            keyboard.append(keyrow)
            keyrow = []
        b = 'w|' + str(i) + '|' + order.text['order']
        keyrow.append(types.InlineKeyboardButton(str(i+1), callback_data=b))
        c += 1;
    keyboard.append(keyrow)
    keyboard.append([types.InlineKeyboardButton('\U000021A9', callback_data='now')])
    markup = types.InlineKeyboardMarkup(keyboard)
    return markup
    

def order_menu(dict): # Клавиатура файлов заявки
    keyboard = []
    keyrow = []
    for i in dict.files:
        if i.startswith('pdf'):
            name_rew = dict.files[i][::-1]
            name = name_rew[name_rew.find("fdp."):name_rew.find("/")]
            name = name[::-1]
            keyrow.append(types.InlineKeyboardButton(f'{name}', callback_data=f's|{name}|'))
    keyboard.append(keyrow)
    keyrow = []
    keyrow.append(types.InlineKeyboardButton('Изделия', callback_data='w|r|'+ dict.text['order']))
    keyrow.append(types.InlineKeyboardButton('Отметить', callback_data='w|w|'+ dict.text['order']))
    keyboard.append(keyrow)
    keyboard.append([types.InlineKeyboardButton('\U000021A9', callback_data='now')])
    markup = types.InlineKeyboardMarkup(keyboard)
    return markup
    
    
def order_key(dict): # Клавиатура заявок
    keyboard = []
    keyrow = []
    r = 0;
    c = 0;
    for i in dict:
        if i == 'text':
            break
        elif c > 3:
            c = 0
            r += 1
            keyboard.append(keyrow)
            keyrow = []
        b = 'o|' + i
        keyrow.append(types.InlineKeyboardButton(i, callback_data=b))
        c += 1;
    keyboard.append(keyrow)
    keyboard.append([types.InlineKeyboardButton('\U000021A9', callback_data='now')])
    markup = types.InlineKeyboardMarkup(keyboard)
    return markup
    

def menu(): 
    keyboard = [
        [types.InlineKeyboardButton('Заявки', callback_data='now')]
        ]
    markup = types.InlineKeyboardMarkup(keyboard)
    return markup
    

@bot.callback_query_handler(func=lambda call: True)
def callback_worker(call):
    if call.data.startswith('o'):
        call_data, x = call.data.split('|')
        order = get_order(x)
    elif call.data.startswith('w'):
        call_data, x, ord = call.data.split('|')
        order = get_order(ord)
    elif call.data.startswith('s'):
        call_data, fl, ord = call.data.split('|')
    else:
        call_data = call.data
    print(call.message.chat.id)
    user=call.message.chat.id
    msg = call.message
    if call_data == 'now':
        text_d = get_disk()
        bot.edit_message_text(
            chat_id=call.message.chat.id, 
            message_id=call.message.message_id,
            text=text_d['text'],
            parse_mode="HTML",
            reply_markup=order_key(text_d));
    elif call_data == 'o':
        icons = ['\U000027A1 ', '\U0001F468 ', '\U0001F4C5 ', '\U0001F464 ']
        text_o = ''
        c = 0
        print('call o: ', order.text)
        for i in order.text:
            if i == 'works':
                break
            text_o += icons[c] + order.text[i] + '\n'
            c += 1
        bot.edit_message_text(
            chat_id=call.message.chat.id, 
            message_id=call.message.message_id,
            text= text_o,
            parse_mode="HTML",
            reply_markup=order_menu(order))
    elif call_data == 'w':
        icons = ['\U000027A1 ', '\U0000274C ', '\U00002705 ', '\U0001F7E9 ', '\U0001F7E2 ']
        order_t = icons[0] + order.text['order'] + '\n\n'
        works = order.text['works']
        text_w = order_t
        print('call w: ', order.text)
        c = 1
        for i in works:
            icon = icons[1]
            if works[i][1] == '8':
                icon = icons[3]
                if works[i][2] != '':
                    icon = icons[2]
            elif works[i][1] == '12':
                icon = icons[4]
                if works[i][2] != '':
                    icon = icons[2]
            elif works[i][1] != '':
                icon = icons[2]
            text_w += icon + str(c) + '_' + i + ': ' + str(works[i][0]) + '\n'
            c += 1
        if x == 'w':
            bot.edit_message_text(
                chat_id=call.message.chat.id, 
                message_id=call.message.message_id,
                text=text_w,
                parse_mode="HTML",
                reply_markup=order_mark(order));
        elif x == 'r':
            bot.edit_message_text(
                chat_id=call.message.chat.id, 
                message_id=call.message.message_id,
                text=text_w,
                parse_mode="HTML",
                reply_markup=order_menu(order));
        else:
            print(x)
            write_xls(call.message.chat.id, x, order)
            order.text = parse_xls(order.text['xls'], order.text['xls_o'])
            bot.edit_message_text(
                chat_id=call.message.chat.id, 
                message_id=call.message.message_id,
                text=text_w,
                parse_mode="HTML",
                reply_markup=order_mark(order));
    elif call_data == 's':
            flsnd = open(f'{fl}', 'rb')
            bot.send_document(
                chat_id=call.message.chat.id, 
                data=flsnd);

            

bot.polling(none_stop=True, interval=0)

