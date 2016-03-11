import os
import re
import fileinput
import io
import xlsxwriter

temp_memory = []
temp_memory_1 = []
counter  = 0
inital_row  = 1
column_number = []
index  = 0
init = 0
final = 0

#for second sheet
inital_row_2 = 1

#Source folder of the files
#dirtocheck = "r'D:\Original_Files\TP_B777_0420_218.cpp"
dirtocheck = r'F:\SOURCE'

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('Test_Procedure_Links.xlsx')
worksheet1 = workbook.add_worksheet('Script_LINKS')
worksheet2 = workbook.add_worksheet('TP_In_Header')
worksheet3 = workbook.add_worksheet('TP_In_Footer')

####################### This is for 1st sheet pointer ############################
row = 0
col = 0
worksheet1.set_row(0, 30)
worksheet1.set_column(0,1, 30)

# To Format the shells for 1st sheet
cell_format = workbook.add_format({'bold': True})

worksheet1.write(row, col,"Test Script", cell_format)
worksheet1.write(row, col + 1, "TP Links", cell_format)

###################### This is for 2nd sheet pointer ##############################
row_1 = 0
col_1 = 0
worksheet2.set_row(0, 30)
worksheet2.set_column(0,1, 40)

# To Format the shells for 2nd sheet
cell_format = workbook.add_format({'bold': True})

worksheet2.write(row, col,"Test Script", cell_format)
worksheet2.write(row, col + 1, "Links Mismatch Between Header & Script", cell_format)

###################### This is for 3rd sheet pointer ##############################

row_2 = 0
col_2 = 0
worksheet3.set_row(0, 30)
worksheet3.set_column(0,1, 40)

# To Format the shells for 2nd sheet
cell_format = workbook.add_format({'bold': True})

worksheet3.write(row, col,"Test Script", cell_format)
worksheet3.write(row, col + 1, "Links Mismatch Between Footer & Script", cell_format)


###################### Main program Starts here #####################################

for root, _, files in os.walk(dirtocheck):
    for f in files:
        counter  = 0
        counter_1 = 0
        del temp_memory[:]
        del temp_memory_1[:]
        fullpath = os.path.join(root, f)
        final = fullpath
        f_old = open(final,'r')

        for i, line in enumerate(f_old):
            if re.search('[/*]\s*LINK:\s*',line, re.IGNORECASE):
                del column_number[:] 
                column_number = list(line[:])
                line_array = line
                for i in range (1, len(column_number)):
                    if ((column_number[i]) == 'T'):
                        init = i

# complicated logic used in order to get the tc id from all the files, irrespective of the way author used it #
                    if (re.search("\s",(column_number[i]))):
                        if ((column_number[i+1]) == '*') and ((column_number[i+2]) == '/'):
                                final = i
                                break
                    elif ((column_number[i]) == '*'):
                        if ((column_number[i+1]) == '/'):
                            final = i
                            break

#slicing of the line in order to get the TC id #
                a = line_array[init:final]
                temp_memory.insert(counter, a)
                counter = counter + 1
            else:
                z = 0

####################### To get TC ids in the header part #####################
            if re.search('[#]\s*TEST CASE/S\s*:', line, re.IGNORECASE):
                counter_1 = 1
            elif (counter_1 == 1):
                k = line.split()
                print k
                if (len(k) == 1):
                    counter_1 = 0
                else:
                    temp_memory_1.insert(index, k[1])
            else:
                z = 0

######################## To get TC ids in the Footer part ######################
##            if re.search('[#]\s*[List]\s*[of]', line, re.IGNORECASE):
##                counter_2 = 1
##            elif (counter_1 == 1):
##                k = line.split()
##                print k
##                if (len(k) == 1):
##                    counter_1 = 0
##                else:
##                    temp_memory_1.insert(index, k[1])
##            else:
##                z = 0

                    
                

        arry_size =  len(temp_memory)
        arry_size_2 = len(temp_memory_1)
        arry1_size = max(arry_size,arry_size_2)
        f_old.close()

        row = inital_row
        col = 0

        worksheet1.write(row, col,f)
        worksheet1.write(row, col + 1, temp_memory[0])
        # Iterate over the data and write it out row by row. #
        row = inital_row + 1
        col = 0
        for i in range (1,arry_size):
            file_contents = ([f,temp_memory[i]])
            worksheet1.write(row, col,     file_contents[0])
            worksheet1.write(row, col + 1, file_contents[1])
            row += 1
        inital_row = row

# all the data is for 2nd sheet - This is to compare to provide those id present in header and not in script #
        col_1 = 0
        row_1 = inital_row_2
        for j in range(0,arry1_size):
            result = temp_memory[j] in temp_memory_1
            print temp_memory[j]
            
            print result
            if (result == False):
                print "coming twice"
                worksheet2.write(row_1, col_1, f)
                worksheet2.write(row_1, col_1 +1, temp_memory[j])
                row_1 +=1
            else:
                z = 0
        inital_row_2 = row_1

        
print temp_memory
print temp_memory_1

workbook.close()
