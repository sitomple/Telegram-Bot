def Friday():
	import os
	import openpyxl
	

	path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
	book = openpyxl.open(f"{path}/lessons/lessons.xlsx", read_only = True)
	sheet = book.active

	NUM_LIST = [54,55,56,57,58,59,60]

	teacher_name = input('Введите свою фамилию: ')


	#ИСПк-102-52-00
	for num in NUM_LIST:
		teacher = sheet[f'K{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'H{num}'].value
				lesson = sheet[f'I{num}'].value
				room = sheet[f'L{num}'].value
				group = sheet[f'I24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ИСПк-103-52-00
	for num in NUM_LIST:
		teacher = sheet[f'O{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'H{num}'].value
				lesson = sheet[f'M{num}'].value
				room = sheet[f'P{num}'].value
				group = sheet[f'M24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ИСПк-104-52-00
	for num in NUM_LIST:
		teacher = sheet[f'S{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'H{num}'].value
				lesson = sheet[f'Q{num}'].value
				room = sheet[f'T{num}'].value
				group = sheet[f'Q24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')

	#ИСПк-105-52-00
	for num in NUM_LIST:
		teacher = sheet[f'Y{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'V{num}'].value
				lesson = sheet[f'W{num}'].value
				room = sheet[f'Z{num}'].value
				group = sheet[f'W24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ИСПк-202-52-00
	for num in NUM_LIST:
		teacher = sheet[f'AC{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'V{num}'].value
				lesson = sheet[f'AA{num}'].value
				room = sheet[f'AD{num}'].value
				group = sheet[f'AA24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ИСПк-203-52-00
	for num in NUM_LIST:
		teacher = sheet[f'AG{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'V{num}'].value
				lesson = sheet[f'AE{num}'].value
				room = sheet[f'AH{num}'].value
				group = sheet[f'AE24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ИСПк-204-52-00
	for num in NUM_LIST:
		teacher = sheet[f'AM{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'AJ{num}'].value
				lesson = sheet[f'AK{num}'].value
				room = sheet[f'AN{num}'].value
				group = sheet[f'AK24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ИСПк-104-51-00
	for num in NUM_LIST:
		teacher = sheet[f'AQ{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'AJ{num}'].value
				lesson = sheet[f'AO{num}'].value
				room = sheet[f'AR{num}'].value
				group = sheet[f'AO24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ИСПк-302-52-00
	for num in NUM_LIST:
		teacher = sheet[f'AU{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'AJ{num}'].value
				lesson = sheet[f'AS{num}'].value
				room = sheet[f'AV{num}'].value
				group = sheet[f'AS24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ИСПк-303-52-00
	for num in NUM_LIST:
		teacher = sheet[f'BA{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'AX{num}'].value
				lesson = sheet[f'AY{num}'].value
				room = sheet[f'BB{num}'].value
				group = sheet[f'AY24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ИСПк-304-52-00
	for num in NUM_LIST:
		teacher = sheet[f'BE{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'AX{num}'].value
				lesson = sheet[f'BC{num}'].value
				room = sheet[f'BF{num}'].value
				group = sheet[f'BC24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ИСПк-305-52-00
	for num in NUM_LIST:
		teacher = sheet[f'BI{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'AX{num}'].value
				lesson = sheet[f'BG{num}'].value
				room = sheet[f'BJ{num}'].value
				group = sheet[f'BG24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ИСПк-201-52-00
	for num in NUM_LIST:
		teacher = sheet[f'BO{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'BL{num}'].value
				lesson = sheet[f'BM{num}'].value
				room = sheet[f'BP{num}'].value
				group = sheet[f'BM24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ИСПк-401/301-51/52-00
	for num in NUM_LIST:
		teacher = sheet[f'BS{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'BL{num}'].value
				lesson = sheet[f'BQ{num}'].value
				room = sheet[f'BT{num}'].value
				group = sheet[f'BQ24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ИСПк-402/403-52-00
	for num in NUM_LIST:
		teacher = sheet[f'BW{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'BL{num}'].value
				lesson = sheet[f'BU{num}'].value
				room = sheet[f'BX{num}'].value
				group = sheet[f'BU24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ЗИОк-102-52-00
	for num in NUM_LIST:
		teacher = sheet[f'CC{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'BZ{num}'].value
				lesson = sheet[f'CA{num}'].value
				room = sheet[f'CD{num}'].value
				group = sheet[f'CA24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ЗИОк-202/101-52/51-00
	for num in NUM_LIST:
		teacher = sheet[f'CG{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'BZ{num}'].value
				lesson = sheet[f'CE{num}'].value
				room = sheet[f'CH{num}'].value
				group = sheet[f'CE24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ЗИОк-302/201-52/51-00
	for num in NUM_LIST:
		teacher = sheet[f'CK{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'BZ{num}'].value
				lesson = sheet[f'CI{num}'].value
				room = sheet[f'CL{num}'].value
				group = sheet[f'CI24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ЗИОк-302/201-52/51-00
	for num in NUM_LIST:
		teacher = sheet[f'CK{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'BZ{num}'].value
				lesson = sheet[f'CI{num}'].value
				room = sheet[f'CL{num}'].value
				group = sheet[f'CI24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ФКк-102/104-52-00
	for num in NUM_LIST:
		teacher = sheet[f'CW{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'CT{num}'].value
				lesson = sheet[f'CU{num}'].value
				room = sheet[f'CX{num}'].value
				group = sheet[f'CU24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ФКк-103-52-00
	for num in NUM_LIST:
		teacher = sheet[f'DA{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'CT{num}'].value
				lesson = sheet[f'CY{num}'].value
				room = sheet[f'DB{num}'].value
				group = sheet[f'CY24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ФКк-202-52-00
	for num in NUM_LIST:
		teacher = sheet[f'DK{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'DH{num}'].value
				lesson = sheet[f'DI{num}'].value
				room = sheet[f'DL{num}'].value
				group = sheet[f'DI24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ФКк-203/101-52/51-00
	for num in NUM_LIST:
		teacher = sheet[f'DO{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'DH{num}'].value
				lesson = sheet[f'DM{num}'].value
				room = sheet[f'DP{num}'].value
				group = sheet[f'DM24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ФКк-301/201-52/51-00
	for num in NUM_LIST:
		teacher = sheet[f'DY{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'DV{num}'].value
				lesson = sheet[f'DW{num}'].value
				room = sheet[f'DZ{num}'].value
				group = sheet[f'DW24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ФКк-302-52-00
	for num in NUM_LIST:
		teacher = sheet[f'EC{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'DV{num}'].value
				lesson = sheet[f'EA{num}'].value
				room = sheet[f'ED{num}'].value
				group = sheet[f'EA24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ФКк-401/402/304-52-00
	for num in NUM_LIST:
		teacher = sheet[f'EM{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'EJ{num}'].value
				lesson = sheet[f'EK{num}'].value
				room = sheet[f'EN{num}'].value
				group = sheet[f'EK24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ПСОк-103/104-52-00
	for num in NUM_LIST:
		teacher = sheet[f'EQ{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'EJ{num}'].value
				lesson = sheet[f'EO{num}'].value
				room = sheet[f'ER{num}'].value
				group = sheet[f'EO24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ПСОк-203-52-00
	for num in NUM_LIST:
		teacher = sheet[f'FA{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'EX{num}'].value
				lesson = sheet[f'EY{num}'].value
				room = sheet[f'FB{num}'].value
				group = sheet[f'EY24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ПСОк-204-52-00
	for num in NUM_LIST:
		teacher = sheet[f'FE{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'EX{num}'].value
				lesson = sheet[f'FC{num}'].value
				room = sheet[f'FF{num}'].value
				group = sheet[f'FC24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ПСОк-101/102-51-00
	for num in NUM_LIST:
		teacher = sheet[f'FI{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'EX{num}'].value
				lesson = sheet[f'FG{num}'].value
				room = sheet[f'FJ{num}'].value
				group = sheet[f'FG24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ПСОк-303-52-00
	for num in NUM_LIST:
		teacher = sheet[f'FO{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'FL{num}'].value
				lesson = sheet[f'FM{num}'].value
				room = sheet[f'FP{num}'].value
				group = sheet[f'FM24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ПСОк-304-52-00
	for num in NUM_LIST:
		teacher = sheet[f'FS{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'FL{num}'].value
				lesson = sheet[f'FQ{num}'].value
				room = sheet[f'FT{num}'].value
				group = sheet[f'FQ24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ПСОк-201/202-51-00
	for num in NUM_LIST:
		teacher = sheet[f'FW{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'FL{num}'].value
				lesson = sheet[f'FU{num}'].value
				room = sheet[f'FX{num}'].value
				group = sheet[f'FU24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ЭКНк-102-52-00
	for num in NUM_LIST:
		teacher = sheet[f'GC{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'FZ{num}'].value
				lesson = sheet[f'GA{num}'].value
				room = sheet[f'GD{num}'].value
				group = sheet[f'GA24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ФНк-102-52-00
	for num in NUM_LIST:
		teacher = sheet[f'GG{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'FZ{num}'].value
				lesson = sheet[f'GE{num}'].value
				room = sheet[f'GH{num}'].value
				group = sheet[f'GE24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ЭКНк-202/101-52/51-00
	for num in NUM_LIST:
		teacher = sheet[f'GK{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'FZ{num}'].value
				lesson = sheet[f'GI{num}'].value
				room = sheet[f'GL{num}'].value
				group = sheet[f'GI24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ФНк-202-52-00
	for num in NUM_LIST:
		teacher = sheet[f'GQ{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'GN{num}'].value
				lesson = sheet[f'GO{num}'].value
				room = sheet[f'GR{num}'].value
				group = sheet[f'GO24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')


	#ФНк-303/201-52/51-00
	for num in NUM_LIST:
		teacher = sheet[f'GU{num}'].value
		if teacher != None:
			teacher = teacher[:-5]  
			if teacher_name == teacher:
				time = sheet[f'GN{num}'].value
				lesson = sheet[f'GS{num}'].value
				room = sheet[f'GV{num}'].value
				group = sheet[f'GS24'].value
				return(f'{teacher_name} время: {time}, дисциплина: {lesson}, аудитория: {room}, группа: {group}')

