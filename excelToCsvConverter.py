#! python3
# excelToCsvConverter.py - Converts xlsx spreadsheets to csv files.

import csv, openpyxl, os

for foldername, subfolders, filenames in os.walk(os.getcwd()):

	# Find and open each xlsx in foldernames and subfolders.
	for filename in filenames:
		if (filename.endswith('.xlsx')):
		
			xlsxPath = os.path.join(foldername, filename)
			workbook = openpyxl.load_workbook(xlsxPath)
			
			print('Loading %s ...' % filename)
			
			# Cycle through sheets in each of the workbooks.
			for sheetName in workbook.get_sheet_names():
			
				# Create new csv for each sheet.
				excelFile = filename[:-5]
				sheetTitle = sheetName
				
				outputName = excelFile + '_' + sheetTitle + '.csv'
				outputFile = open(outputName, 'w', newline='')
				outputWriter = csv.writer(outputFile)
				
				sheet = workbook.get_sheet_by_name(sheetName)

				# Cycle through rows and columns.
				for i in range(1, sheet.max_row + 1):

					# New row holds list of values to be appended later.
					newRow = []

					for j in range(1, sheet.max_column + 1):
			
						# Append cell to string list.
						newRow.append(sheet.cell(row=i, column=j).value)
						
					# Append string list to csv as row.
					outputWriter.writerow(newRow)

				# Confirm and close csv file.
				print('%s successfully saved as %s.' % (filename, outputName))
				
				outputFile.close()