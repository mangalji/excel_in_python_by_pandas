# # from openpyxl import Workbook
# from openpyxl import load_workbook
# # import openpyxl

# #workbook = Workbook()
# #sheet = workbook.active

# #sheet["A1"] = "hello"
# #sheet["B1"] = "world!"

# # workbook.save(filename="hello_world.xlsx")

# workbook = load_workbook(filename="AttendRecord.xlsx")
# wb = workbook.active

# names = []

# for row in range(0, wb.max_row):
# 	for col in wb.iter_cols(1, wb.max_column):
# 		# print(row)
# 		# print(col)
# 		# if col[row] == 'L'+row:
# 		# 	names.append(col[row].value)
# 		# a = col[row]
# 		# print(a)
# 		# print(type(a))
# 		# print(type(col[row].value))
# 		# print(col[row].value)
# 		if col[row] == None:
# 			# print(col[row].value)
# 			continue
# 		else:
# 			print(col[row].value)


# # total_sheets = workbook.sheetnames
# # # print(total_sheets)

# # sheet = workbook.active
# # # print(sheet)

# # title = sheet.title
# # # print(title)

# # # print(sheet["I1"].value)
# # # print(sheet.cell(row=5,column=12))
# # # print(sheet.cell(row=17,column=12).value)

# # # sheet.cell()

# # # print(sheet["A1:Z10"])

# # # for i in range(5,78,3):
# # # 	for row in sheet.iter_rows(min_row=i,min_col=12,values_only=True):
# # # 		print(row)

# # length = 0


# # for i in range(5,78,2):
# # 	print(i)
# # 	print("")
# # 	for value in sheet.iter_rows(min_row=i,max_row=77,min_col=12,max_col=12,values_only=True):
# # 		print(value)

# # # for column in sheet.iter_cols(min_row=1,max_row=20,min_col=1,max_col=30):
# # # 	print(column)




import pandas as pd
from datetime import datetime, timedelta

#step 1 load excel without header
file_name = "AttendRecord.xlsx"
df = pd.read_excel(file_name, header=None)
# print(df)
len(df)

#step 2 find all "userID" rows
#column index 1 (and column) contains text "userID:"

user_rows = df.index[df[1] == "UserID:"].tolist()
print(user_rows)
summaries = []

#step 3 process each user block

for user_id_row in user_rows:
	print(user_id_row)
	#(a) find user id (in column 3 - index 3)
	user_id = df.loc[user_id_row, 3]
	print(user_id)
	#(b) Days row and time row
	employee_name = df.loc[user_id_row, 11] 
	print(employee_name)
	day_row = user_id_row + 1  #row with 1,2,3,4,....30
	print(day_row)
	time_row = user_id_row + 2 # row with time data like "12:20\n18:16"
	print(time_row)
	#safety: if file ends unexpectedly, skip
	if time_row >= len(df):
		continue

	header_values = df.loc[day_row]
	print(header_values)

	# (c) find columns that actually have day numbers

	day_cols = []
	for col_indx ,val in header_values.items():
		# if the cell is a number (1,2,3,...)
		if isinstance(val, (int,float)) and not pd.isna(val):
			day_cols.append(col_indx)

	present_days = 0
	total_duration = timedelta(0)

	# (d) loop on each day column for this user

	for col in day_cols:
		cell = df.loc[time_row,col]

		#cell example: "12:20\n18:16" OR NaN
		if isinstance(cell,str) and cell.strip():
			parts = [p.strip() for p in cell.split('\n') if p.strip()]

			try:
				t_in = datetime.strptime(parts[0],"%H:%M")
			except:
				continue

			default_out_time = datetime.strptime("20:00","%H:%M")

			if len(parts) >= 2:
				try:
					#first time = in, last time = out
					# t_in = datetime.strptime(parts[0], "%H:%M")
					t_out = datetime.strptime(parts[-1], "%H:%M")
				except:
					t_out = default_out_time
			else:
				t_out = default_out_time
					#if out is smaller than in, assume it crossed midnight 

					# if t_out < t_in:
					# 	t_out += timedelta(days=1)

			if t_out < t_in:
    			t_out += timedelta(days=1)

			duration = t_out - t_in
			total_duration += duration
			present_days += 1

	summaries.append({
		"UserID":user_id,
		"PresentDays":present_days,
		"Name":employee_name,
		"TotalHours": total_duration
		})		 
# step4 create a summary table

summary_df = pd.DataFrame(summaries)

#convert "TotalHours" (timedelta) into hours:minutes for readibility
summary_df["TotalHours_str"] = summary_df["TotalHours"].apply(
	lambda td: f"{td.days * 24 + td.seconds // 3600:02d}:{(td.seconds % 3600)//60:02d}")

print(summary_df)

# ---------- STEP 5: Save to Excel/CSV ----------

summary_df.to_excel("attendence_summary.xlsx",index=False)
# summary_df.to_csv("attendence_summary.csv", index=False)
print("\nSummary saved as 'attendence_summary.xlsx' and 'attendence_summary.csv'")




















