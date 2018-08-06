#!/usr/bin/env python3

import time
import openpyxl
import json
import string
import random
import string
from tkinter import ttk
import tkinter as tk
from tkinter.messagebox import showinfo

def popup_showinfo(info):
	showinfo("Window", str(info))

class Application(tk.Frame):
		
		def __init__(self, master=None):
			#global varibles define scope
			self.n = 0
			self.cur_mode = 1
			self.first_time = 0
			self.picked_num = 0
			self.chosen_one = -1
			self.total_student_num = 0

			#				

			super().__init__(master)
			self.total_student_num = jstudent[-1]['id']+1
			master.title("Chosen System")
			master.geometry('640x670')
			self.create_windows()
			self.grid()

		def create_windows(self):
			
			self.mode_string = "模式: 抽完不放回"


			self.mode_lab = tk.Label(self, text=self.mode_string, font=("Courier", 20), pady=15)
			self.mode_lab.grid(column=1, row=1)

			self.empty = tk.Label(self, text="未抽", 
				font=("Courier", 20))
			self.empty.grid(column=0, row=3, sticky="w")

			self.empty = tk.Label(self, text="已抽出", 
				font=("Courier", 20))
			self.empty.grid(column=0, row=5, sticky="w")

			self.num = tk.Label(self, text="Mr. X", 
				font=("Courier", 50), fg="orange")
			self.num.grid(column=1, row=2)

			self.pre = tk.Label(self, text="恭喜", 
				font=("Courier", 30), fg="red")
			self.pre.grid(column=0, row=2)

			self.suf = tk.Label(self, text="中獎！！", 
				font=("Courier", 30), fg="red")
			self.suf.grid(column=2, row=2)

			self.input_entry = tk.Entry(self)
			self.input_entry.grid(column=4, row=1)


			self.start_buttom = tk.Button(self, text="開抽")
			self.start_buttom["fg"] = "red"
			self.start_buttom["command"] = self.start
			self.start_buttom.config(font=("Courier", 30))
			self.start_buttom.grid(column=0, row=0, rowspan=1)


			self.chmod_buttom = tk.Button(self, text="換模式")
			self.chmod_buttom["command"] = self.change_mode
			self.chmod_buttom.grid(column=1, row=0, rowspan=1)


			self.static_buttom = tk.Button(self, text="歷史數據")
			self.static_buttom["fg"] = "red"
			self.static_buttom["command"] = self.add_num
			self.static_buttom.grid(column=2, row=0, rowspan=1)


			self.reset_buttom = tk.Button(self, text="重設")
			self.reset_buttom["fg"] = "red"
			self.reset_buttom["command"] = self.reset
			self.reset_buttom.grid(column=3, row=0, rowspan=1)


			self.load_buttom = tk.Button(self, text="載入檔案")
			self.load_buttom["fg"] = "red"
			self.load_buttom["command"] = self.loadStudentFile
			self.load_buttom.grid(column=4, row=0, rowspan=1)


			self.leave_buttom = tk.Button(self, text="離開")
			self.leave_buttom["command"] = root.destroy
			self.leave_buttom.grid(column=4, row=7)


			self.chosing_pool = tk.Text(self, height="10", width="50", 
				borderwidth=3, relief="solid")
			self.chosing_pool.config(font=("Courier", 20))
			self.show_pool()
			self.chosing_pool.grid(column=0, row=4, columnspan=5)

			self.chosed_pool = tk.Text(self, height="10", width="50", 
				borderwidth=3, relief="solid")
			self.chosed_pool.config(font=("Courier", 20))
			self.show_choosed()
			self.chosed_pool.grid(column=0, row=6, columnspan=5)

		def show_pool(self):
			self.chosing_pool.delete(1.0, "end")

			for student in jstudent:
				if student["picked"] == 0:	
					self.chosing_pool.insert('insert', student["name"])
					self.chosing_pool.insert('insert'," ")

		def show_choosed(self):
			self.chosed_pool.delete(1.0, "end")

			for student in jstudent:
				if student["picked"] != 0:	
					self.chosed_pool.insert('insert', student["name"])
					self.chosed_pool.insert('insert'," ")
			
		def end(self):
			self.reset()
			root.destroy

		def reset(self):
			self.picked_num = 0

			for student in jstudent:
				student['picked'] = 0
				student['temp_picked_time'] = 0

			self.show_pool()
			self.show_choosed()

		def add_num(self):
			try:
				self.n += 1
				self.num["text"] = self.n
			except:
				popup_showinfo("test success!")

		def start(self):
			#print("total_student_num:", self.total_student_num)
			#print("picked_num", self.picked_num)
			can_pick = True
			if self.total_student_num == self.picked_num and self.cur_mode == 1:
				can_pick = False
				popup_showinfo("籤被抽完摟～\n 請按重設鈕")


			while can_pick:

				self.chosen_one = random.randrange(self.total_student_num)
				if self.cur_mode == 1:
					#check is picked or not
					if jstudent[self.chosen_one]['picked'] == 1:
						continue
					else:
						jstudent[self.chosen_one]['picked'] = 1
						self.picked_num += 1
						#print("picked_num", self.picked_num)
				break

			jstudent[self.chosen_one]['temp_picked_time'] += 1
			jstudent[self.chosen_one]['total_picked_time'] += 1
			self.num["text"] = jstudent[self.chosen_one]['name']
			self.show_pool()
			self.show_choosed()

		def change_mode(self):

			self.cur_mode *= -1
			self.reset()

			if self.cur_mode == 1:
				self.mode_string = "模式:抽完不放回"

			else:
				self.mode_string = "模式:抽完放回"

			self.mode_lab["text"] = self.mode_string

		def loadStudentFile(self):
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

			self.path = self.input_entry.get()

			try:
				file_obj = openpyxl.load_workbook(self.path)
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

				student_n = 0

				for student in jstudent:
					student_n += 1

				for n in range(student_n):
					jstudent.pop(0)

				jstudent_temp = json.loads(open('./PickRecord.json').read())

				#jstudent.append(jstudent_temp)
				for student in jstudent_temp:
					jstudent.append(student)

				self.total_student_num = jstudent[-1]['id']+1

				popup_showinfo("檔案載入成功！")

				
				self.show_pool()
				self.show_choosed()

			except:
				#print("檔案載入錯誤！")
				popup_showinfo("檔案載入錯誤！")

		

if __name__ == '__main__':
	root = tk.Tk()
	havntLoaded = True
	try:
		jstudent = json.loads(open('./PickRecord.json').read())
		havntLoaded = False
	except:
		popup_showinfo("檔案還未建立，請將學生資料放入現在的資料夾、並按載入檔案")
		jstudent = [{"id":0, "name":"", "picked":0}]

	app = Application(master=root)
	app.mainloop()

