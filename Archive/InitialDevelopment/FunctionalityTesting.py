from datetime import date
import calendar
#Received an error on NASDAQ 4238 on Truecar (TRUE)
input = True
url = "https://roic.ai/company/" + str(input) #Cast it.
print(url)

#Ensuring functionality before replacing component in main Phase 1 file
input = "this is lucas-luwa's first test. I like hiking, investing and /playing tennis. ()%"
approvedCharacters = {'-', ',', '%', ' ', '/', '(', ')'}
offset = 0
#Works on 5.27.23 
for i in range(0, len(input)-1):
    if (input[i].isnumeric() or offset > 0 or input[i] == '-' or input[i] == ','  or input[i] == '%' or input[i] == ' ' or input[i] == '/' or input[i] == '(' or input[i] == ')'): 
        print("OLDTRUE")
    else: 
        print("OLDFALSE")
    if(input[i].isnumeric() or offset > 0 or input[i] in approvedCharacters): 
        print("NEWTRUE")
    else: 
        print("NEWFALSE")
    print("")

#Testing for new naming convention
year, month, day = str(date.today()).split('-')
print(year,calendar.month_name[int(month)], day)
print(str(calendar.month_name[int(month)]) + str(year) + "RawDataV")

