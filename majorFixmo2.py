import xlrd
import sys
file_name =input("Enter the name of Excel sheet no need for extension:  ")
file_location = "C:/Users/Jeffrey/Desktop/"+file_name+".xlsx"
work_book = xlrd.open_workbook(file_location)
sheet = work_book.sheet_by_index(1)
formated = input("Enter the destination name no need for extension(doesn't need to exist):  ")
file = open("c:/Users/Jeffrey/Desktop/"+formated+".txt","w")
number_of_students = sheet.nrows
for i in range(1, number_of_students):
	id = sheet.cell_value(i,0)
	last = sheet.cell_value(i,1)
	first = sheet.cell_value(i,2)
	hold = sheet.cell_value(i,3)
	ccsf = sheet.cell_value(i,4)
	sl = sheet.cell_value(i,5)
	id.replace(" ", "")
	last.replace(" ", "")
	first.replace(" ", "")
	sl.replace(" ", "")
	if(ccsf != ""):
		new_String ="\"" +id+" : "+last+", "+first+" \" <"+ccsf+">"
	elif(sl !=""):
		new_String ="\"" +id+" : "+last+", "+first+" <"+sl+">"
	num = int((i/(number_of_students))*100)
	sys.stdout.write("Write progress: %d%%   \r" %num  )
	sys.stdout.flush()
	file.write(new_String+"\n")
sys.stdout.write("write progress: 100 %")
sys.stdout.flush()
file.close()