import telebot
import openpyxl
from groups import setings
from datetime import datetime, date, time
wb = openpyxl.load_workbook(setings.PATH_TO_TABLE)
sheet = wb['СПО']

def handle_text(message,bot):
	if message.text.strip() == 'Сегодня':
		day = datetime.today().weekday()
		week_study(message, day, bot)

	elif message.text.strip() == 'Завтра':
		day = datetime.today().weekday() + 1
		if day < 6:
			week_study(message, day, bot)

	elif message.text.strip() == 'Послезавтра':
		day = datetime.today().weekday() + 2
		if day < 6:
			week_study(message, day, bot)

	elif message.text.strip() == 'Неделя':
			week_study(message, 9, bot)

#
def Monday(message, day,bot):
	bot.send_message(message.chat.id, "Понедельник:")
	cells = sheet['AX26':'BB32']
	for time, lesson, a, b, room, in cells:
		if (time.value != None ) & (lesson.value != None ) & (room.value != None ) :
			bot.send_message(chat_id=message.chat.id,
							 text=f"Время: {time.value} \nПара: {lesson.value} \nВ аудитории: {room.value}")

def Tuesday(message, day,bot):
	bot.send_message(message.chat.id, "Вторник:")
	cells = sheet['AX33':'BB40']
	for time, lesson, a, b, room, in cells:
		if (time.value != None) & (lesson.value != None) & (room.value != None):
			bot.send_message(chat_id=message.chat.id,
							 text=f"Время: {time.value} \nПара: {lesson.value} \nВ аудитории: {room.value}")

def Wednesday(message, day,bot):
	bot.send_message(message.chat.id, "Среда:")
	cells = sheet['AX41':'BB47']
	for time, lesson, a, b, room, in cells:
		if (time.value != None) & (lesson.value != None) & (room.value != None):
			bot.send_message(chat_id=message.chat.id,
							 text=f"Время: {time.value} \nПара: {lesson.value} \nВ аудитории: {room.value}")

def Thursday(message, day,bot):
	bot.send_message(message.chat.id, "Четверг:")
	cells = sheet['AX48':'BB54']
	for time, lesson, a, b, room, in cells:
		if (time.value != None) & (lesson.value != None) & (room.value != None):
			bot.send_message(chat_id=message.chat.id,
							 text=f"Время: {time.value} \nПара: {lesson.value} \nВ аудитории: {room.value}")

def Friday(message, day,bot):
	bot.send_message(message.chat.id, "Пятница:")
	cells = sheet['AX55':'BB61']
	for time, lesson, a, b, room, in cells:
		if (time.value != None) & (lesson.value != None) & (room.value != None):
			bot.send_message(chat_id=message.chat.id,
							 text=f"Время: {time.value} \nПара: {lesson.value} \nВ аудитории: {room.value}")

def Saturday(message, day,bot):
	bot.send_message(message.chat.id, "Суббота:")
	cells = sheet['AX62':'BB67']
	for time, lesson, a, b, room, in cells:
		if (time.value != None) & (lesson.value != None) & (room.value != None):
			bot.send_message(chat_id=message.chat.id,
							 text=f"Время: {time.value} \nПара: {lesson.value} \nВ аудитории: {room.value}")

def week_study(message, day, bot):
	if day == 0:
		Monday(message, day,bot)

	elif day == 1:
		Tuesday(message, day,bot)

	elif day == 2:
		Wednesday(message, day,bot)

	elif day == 3:
		Thursday(message, day,bot)

	elif day == 4:
		Friday(message, day,bot)

	elif day == 5:
		Saturday(message, day,bot)

	elif day == 9:

		Monday(message, day,bot)
		bot.send_message(chat_id=message.chat.id, text="---------------------")

		Tuesday(message, day,bot)
		bot.send_message(chat_id=message.chat.id, text="---------------------")

		Wednesday(message, day,bot)
		bot.send_message(chat_id=message.chat.id, text="---------------------")

		Thursday(message, day,bot)
		bot.send_message(chat_id=message.chat.id, text="---------------------")

		Friday(message, day,bot)
		bot.send_message(chat_id=message.chat.id,text="---------------------")


		Saturday(message, day,bot)

	else :
		bot.send_message(message.chat.id, "Вы ошиблись с командой, попробуйте ещё раз.")