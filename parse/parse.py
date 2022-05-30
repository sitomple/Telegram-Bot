import requests
import schedule
import json
import os

token = 'e9f84467b77bfda80c832b0ca7254f96b2ab82329417b23d61b091186da68c56787339f7fa78e3261aec8'

group_name = 'kollegevyatsu'
path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))


def get_timetable():
	url = f'https://api.vk.com/method/wall.get?domain={group_name}&count=15&access_token={token}&v=5.131'
	r = requests.get(url)
	src = r.json()

	# Проверка пути
	if os.path.exists(f'{path}/bot/lessons'):
		print(f"Директория с именем {group_name} уже существует!")
	else:
		os.mkdir(f'{path}/bot/lessons')
	
	#print(posts)


	with open(f'{path}/bot/lessons/{group_name}.json', 'w', encoding = "utf-8") as file:
		json.dump(src, file, indent = 4, ensure_ascii = False)

	attachments = []
	urls = []
	posts = src['response']['items']

	# Ищет посты
	for post in posts:
		try:
			# Нет смысла искать старое расписание, а новое будет первым, поэтому в списке один элемент
			if len(urls) == 0:
				post = post["attachments"]
				# Если в посте есть 'xlsx'
				if post[0]['type'] == 'doc' and post[0]['doc']['ext'] == 'xlsx':
					u = post[0]['doc']['url']
					urls.append(u)
		except Exception:
			# print('err')
			pass
	
	# print(urls)

	# Запись

	try:
		with open(f'{path}/bot/lessons/lessons.xlsx', 'wb') as file:
			r = requests.get(urls[0], allow_redirects = True)
			file.write(r.content)
			os.remove(f'{path}/bot/lessons/{group_name}.json')
	except Exception:
		# print('err')
		pass


if __name__ == "__main__":
	while True:
		schedule.every().minute.at(":17").do(get_timetable())