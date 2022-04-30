import telebot
import openpyxl
from groups.ISP2 import isp_101_52_00
from groups.ISP2 import isp_202_52_00
from groups.ISP2 import isp_203_52_00
from groups.ISP2 import isp_204_52_00
from groups.ISP1 import isp_105_52_00
from groups.ISP1 import isp_104_52_00
from groups.ISP1 import isp_102_52_00
from groups.ISP1 import isp_103_52_00
import setings

from datetime import datetime, date, time

bot = telebot.TeleBot(setings.TOKEN_TELEGRAM)
wb = openpyxl.load_workbook(setings.PATH_TO_TABLE)
sheet = wb['СПО']

def day_lessons(message):
	wb = openpyxl.load_workbook('base.xlsx')
	sheet = wb['List']
	B = sheet.max_row
	cells = sheet['A1':f'B{B}']
	for id, gr in cells:
		if id.value == str(message.chat.id):
			if gr.value == 'ИСПк-101':
				less = isp_101_52_00
				less.handle_text(message, bot)
			elif gr.value == 'ИСПк-202':
				less = isp_202_52_00
				less.handle_text(message, bot)
			elif gr.value == 'ИСПк-203':
				less = isp_203_52_00
				less.handle_text(message, bot)
			elif gr.value == 'ИСПк-204':
				less = isp_204_52_00
				less.handle_text(message, bot)
			elif gr.value == 'ИСПк-104':
				less = isp_104_52_00
				less.handle_text(message, bot)
			elif gr.value == 'ИСПк-105':
				less = isp_105_52_00
				less.handle_text(message, bot)
			elif gr.value == 'ИСПк-102':
				less = isp_102_52_00
				less.handle_text(message, bot)
			elif gr.value == 'ИСПк-103':
				less = isp_103_52_00
				less.handle_text(message, bot)

def base_group(bot, group, message):
		bot.send_message(message.chat.id, f'Активная группа: {message.text.strip()}-52-00')
		tumbler = 0
		i = 0
		wb = openpyxl.load_workbook('base.xlsx')
		sheet = wb['List']
		B = sheet.max_row
		cells = sheet['A1':f'B{B}']
		for id, gr in cells:
			i += 1
			if id.value == str(message.chat.id):
				tumbler = 1
				sheet[f'B{i}'] = f'{group}'
			wb.save('base.xlsx')
		if tumbler == 0:
			sheet[f'B{B + 1}'] = f'{group}'
			sheet[f'A{B + 1}'] = f'{message.chat.id}'
			wb.save('base.xlsx')

# Команда start
@bot.message_handler(commands=["start"])
def start(m, res=False):
		bot.send_message(m.chat.id, 'Бот был написан Zarnii и Sitomple\nДля групп ИСПк 2 курса \n'
									'Напишите Help')
		# Добавляем две кнопки
		markup = telebot.types.ReplyKeyboardMarkup(resize_keyboard=True)
		item1 = telebot.types.KeyboardButton("Help")
		item2 = telebot.types.KeyboardButton("Сегодня")
		item3 = telebot.types.KeyboardButton("Завтра")
		item4 = telebot.types.KeyboardButton("Послезавтра")
		item5 = telebot.types.KeyboardButton("Неделя")
		markup.add(item1, item2, item3)
		markup.add(item4, item5)
		bot.send_message(m.chat.id, 'Выберите день, на который хотите получить расписание:',  reply_markup=markup)


@bot.message_handler(content_types=["text"])
def handle_text(message):
	if (message.text.strip().lower() == 'help') or (message.text.strip().lower() == 'расписание'):
		bot.send_message(message.chat.id, 'Напишите свою группу \nПример: ИСПк-204')

	if (message.text.strip() == 'ИСПк-101-51-00') or (message.text.strip() == 'ИСПк-101'):
		group = 'ИСПк-101'
		base_group(bot, group, message)

	elif (message.text.strip() == 'ИСПк-202-52-00') or (message.text.strip() == 'ИСПк-202'):
		group = 'ИСПк-202'
		base_group(bot, group, message)

	elif (message.text.strip() == 'ИСПк-203-52-00') or (message.text.strip() == 'ИСПк-203'):
		group = 'ИСПк-203'
		base_group(bot, group, message)

	elif (message.text.strip() == 'ИСПк-204-52-00') or (message.text.strip() == 'ИСПк-204'):
		group = 'ИСПк-204'
		base_group(bot, group, message)
	elif (message.text.strip() == 'ИСПк-105-52-00') or (message.text.strip() == 'ИСПк-105'):
		group = 'ИСПк-105'
		base_group(bot, group, message)
	elif (message.text.strip() == 'ИСПк-104-52-00') or (message.text.strip() == 'ИСПк-104'):
		group = 'ИСПк-104'
		base_group(bot, group, message)
	elif (message.text.strip() == 'ИСПк-102-52-00') or (message.text.strip() == 'ИСПк-102'):
		group = 'ИСПк-102'
		base_group(bot, group, message)
	elif (message.text.strip() == 'ИСПк-103-52-00') or (message.text.strip() == 'ИСПк-103'):
		group = 'ИСПк-103'
		base_group(bot, group, message)


	if message.text.strip() == 'Сегодня':
		day_lessons(message)

	elif message.text.strip() == 'Завтра':
		day_lessons(message)

	if message.text.strip() == 'Послезавтра':
		day_lessons(message)

	if message.text.strip() == 'Неделя':
		day_lessons(message)

bot.infinity_polling()