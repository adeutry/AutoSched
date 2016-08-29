import openpyxl
from openpyxl.styles import Side

class Schedule:

	def __init__(self):
		self.wb = openpyxl.load_workbook('204Fall2016.xlsx')
		self.sheet = self.wb.get_sheet_by_name('Rm 204')
		self.day_dict = {'M' : 'C', 'T' : 'D', 'W' : 'E', 'TR' : 'F', 'F' : 'G'}

	def add_section(self, section):
		start_min = section.start % 100
		start_hour = (section.start - start_min)/100

		end_min = section.end % 100
		end_hour = (section.end - end_min)/100

		start_block = (start_hour-8)*2

		if start_min >= 30:
			start_block += 1

		end_block = (end_hour-8)*2

		if end_min >= 30:
			end_block += 1

		med_border = self.sheet['C9'].border.top
		med_side = Side(style = 'medium')
		gray_fill = self.sheet['C10'].fill
		self.sheet['C10'].fill = self.sheet['C11'].fill

		#first row
		index = self.day_dict[section.day] + str(start_block)
		self.sheet[index].border.left = med_side
		self.sheet[index].border.right = med_side
		self.sheet[index].border.top = med_side
		self.sheet[index].value = section.num

		#middle rows

		#end row

		

class Section:

	def __init__(self, day, num, instructor, start, end):
		self.day = day
		self.num = num
		self.instructor = instructor
		self.start = start
		self.end = end


sched = Schedule()
sec = Section('M', 2000, 'Mr. Meme', 1200, 1400)
sched.add_section(sec)
