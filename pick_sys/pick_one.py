#!/usr/bin/env python3

import openpyxl
import json
import string
import random
import string
from tkinter import ttk
import tkinter as tk

total_student_num = 0


def loadStudentFile(path, fine):
	switch_bit = -1
	id_index = 0
	student = {
		'name': '',
		'gender': '',
		'id':-1,
		'picked': -1,
		'temp_picked_time':-1,
		'total_picked_time':-1
	}

	student_list = []

	try:
		file_obj = openpyxl.load_workbook(path)
		sheet = file_obj.get_sheet_by_name('Sheet1')

		for row in sheet:
			for cell in row:
				if switch_bit == -1:
					student['name'] = cell.value
				else:
					student['gender'] = cell.value

				switch_bit *= -1
				student['picked'] = 0
				student['temp_picked_time'] = 0
				student['total_picked_time'] = 0
				student['id'] = id_index

			tmp = student
			id_index += 1

			student_list.append(tmp.copy())

		#total_student_num = id_index				
		#print(student_list)

		with open ('PickRecord.json','w') as new_file:
			new_file.write(json.dumps(student_list))

		fine = 1
		return fine

	except:
		print("檔案載入錯誤！")

cur_mode = 0
first_time = 0
picked_num = 0
chosen_one = -1
def action():

	print("可以打help看看有什麼指令可以用喔！")
	while True:
		if first_time == 0:
			try:
				jstudent = json.loads(open('./PickRecord.json').read())

				first_time += 1
				#for student in jstudent:
				#	print(student['name'])

			except:
				print("檔案還未建立，請將學生資料放入現在的資料夾")
				print("（檔案格式限定: xlsx）")
				file_name = input("請輸入檔名: ")
				fine = 0
				fine = loadStudentFile(file_name, fine)#'./Test_readxml.xlsx')
				if fine == 1:
					print("檔案載入完成！")
				continue

		total_student_num = jstudent[-1]['id']+1
		cmd = input(">>>")

		if cmd == "chmod":
			print("模式: 0-> 抽出不放回, 非0-> 抽出放回")
			print("當前模式：", cur_mode)

			new_mode = input("改變模式為(0/1)：")
			cur_mode = int(new_mode)

		elif cmd == "s":
			if total_student_num == picked_num and cur_mode == 0:
				print("籤被抽完摟～")
				print("輸入reset重置籤筒")
				continue

			while True:

				chosen_one = random.randrange(total_student_num)
				if cur_mode == 0:
					#check is picked or not
					if jstudent[chosen_one]['picked'] == 1:
						continue
					else:
						jstudent[chosen_one]['picked'] = 1
						picked_num += 1
				break

			print("！！恭喜  乂 ", jstudent[chosen_one]['name']," 乂  中獎！！")
			jstudent[chosen_one]['temp_picked_time'] += 1
			jstudent[chosen_one]['total_picked_time'] += 1
			


		elif cmd == "show":
			print("模式：", cur_mode)
			print("(0-> 抽出不放回, 非0-> 抽出放回)")
			print("總人數\t", total_student_num)

			if cur_mode == 0:
				print("已抽出人數：", picked_num)
				for unpicked in jstudent:
					if unpicked['picked'] == 1:
						print(unpicked['name'],"\t", end = "")

				print()

				print("未抽出人數：", total_student_num - picked_num)
				for unpicked in jstudent:
					if unpicked['picked'] == 0:
						print(unpicked['name'],"\t", end = "")

				print()

			else:
				for student in jstudent:
					print(student['name'],"(",
						student['temp_picked_time'],"): ", end = "")

					for i in range(student['temp_picked_time']):
						print("蓮", end="")

					print()

		elif cmd == "static":
			for student in jstudent:
				print(student['name'],"(",
					student['total_picked_time'],"): ", end = "")

				for i in range(student['total_picked_time']):
					print("蓮", end="")

				print()

		elif cmd == "reset":
			picked_num = 0
			for student in jstudent:
				student['picked'] = 0
				student['temp_picked_time'] = 0

		elif cmd == "help":
			print("以下為可用指令")
			print("chmod: \t\t更改抽籤模式")
			print("s: \t\t開始抽籤")
			print("show: \t\t列出籤桶還剩哪幾支籤")
			print("static: \t統計從抽出第一支籤開始，每個人的抽籤紀錄")
			print("reset: \t\t將抽出的籤放回籤桶")
			print("q: \t\t結束抽籤")


		elif cmd == "q":
			for student in jstudent:
				student['temp_picked_time'] = 0
			break

		else:
			print("輸入錯誤，輸入help可查看可輸入指令")

	#reset temp data
	for student in jstudent:
		student['picked'] = 0
		student['temp_picked_time'] = 0

	with open ('PickRecord.json','w') as new_file:
		new_file.write(json.dumps(jstudent))


	#loadStudentFile('./Test_readxml.xlsx')
class Application(tk.Frame):
		
		def __init__(self, master=None):
			#global varibles define scope
			self.n = 0

			#
			super().__init__(master)
			master.title("Chosen System")

			self.grid()
			self.create_windows()

		def create_windows(self):
			self.msg_box = tk.Label(self)
			self.msg_box["text"] = "Hello World!\n"
			self.msg_box.config(fg="red")
			self.msg_box.grid(column=0, row=0, ipady=20)


			self.num = tk.Label(self, text=self.n)
			self.num.grid(column=0, row=1)
			

			self.test_buttom = tk.Button(self, text="Add number")
			self.test_buttom["fg"] = "red"
			self.test_buttom["command"] = self.add_num
			self.test_buttom.grid(column=1, row=0)


			
		def add_num(self):
			self.n += 1
			self.num["text"] = self.n

if __name__ == '__main__':

	root = tk.Tk()
	app = Application(master=root)
	app.mainloop()
