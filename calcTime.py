from pandas import DataFrame
import xlwt
import datetime

day = "1 Shanbe"
fileName = "9633.time"

file = open(fileName, "r")
s = file.read()
index = s.find(day)
index = s.find("\n", index + 1)
index = s.find("\n", index + 1)
startIndex = s.find("\n", index + 1)
endIndex = s.find("Summary", index + 1)
lines = s[startIndex:endIndex].splitlines()
data = [[]]
for i in range(1, len(lines) - 2):
    data.append([])
    parts = lines[i].split(' ')
    times = parts[0].split('-')
    data[i].append(times[0])
    data[i].append(times[1])
    data[i].append(parts[1])
wb = xlwt.Workbook()
sheet = fileName.split(".")

ws = wb.add_sheet("Scoreboard")

fmt = xlwt.Style.easyxf("""
font: name Liberation Sans;
alignment: horiz centre;
""", num_format_str='[HH]:MM')

date_format = xlwt.XFStyle()
date_format.num_format_str = 'hh:mm'

num_fmt = xlwt.Style.easyxf("""
font: name Liberation Sans;
alignment: horiz centre;
""", num_format_str='0')
# for writing times
for i in range(1, len(data)):
    time = data[i][0].split(':')
    print(str(i))
    ws.write(i, 1, datetime.time(int(time[0]), int(time[1])), fmt)
    time = data[i][1].split(':')
    ws.write(i, 2, datetime.time(int(time[0]), int(time[1])), fmt)
    ws.write(i, 3, int(data[i][2]), num_fmt)
    ind = str(i + 1)
    for j in range(1, 6):
        ws.write(i, 3 + j, xlwt.Formula("IF(D" + ind + "=" + str(j) + ",C" + ind + "-B" + ind + ",0)"), fmt)

numOfRow = len(data)

# Writing headers
headers = ['Start', 'End', 'Type', 'Work', 'Futuristic', 'Entertainment', 'Other', 'Critical']

for i in range(0, len(headers)):
    ws.write(0, i + 1, headers[i], fmt)

# Writing SUMs
ws.write(numOfRow, 4, xlwt.Formula("SUM(E2:E" + str(numOfRow) + ")"), fmt)
ws.write(numOfRow, 5, xlwt.Formula("SUM(F2:F" + str(numOfRow) + ")"), fmt)
ws.write(numOfRow, 6, xlwt.Formula("SUM(G2:G" + str(numOfRow) + ")"), fmt)
ws.write(numOfRow, 7, xlwt.Formula("SUM(H2:H" + str(numOfRow) + ")"), fmt)
ws.write(numOfRow, 8, xlwt.Formula("SUM(I2:I" + str(numOfRow) + ")"), fmt)
ws.write(numOfRow, 9, xlwt.Formula("SUM(E" + str(numOfRow + 1) + ":I" + str(numOfRow + 1) + ")"), fmt)

ws = wb.add_sheet(sheet[0])
wb.save('test.ods')
