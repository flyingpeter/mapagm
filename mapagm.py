"""
Take all comments from the Datasheet and save them to separate files
"""

import openpyxl
import re

month = 'Março19'

wb = openpyxl.load_workbook('mapagm.xlsx', data_only=True)
ws = wb[month]

webspace = ['<html>',
            '<head>',
            '<title>MapaGM</title>',
            '<link rel="stylesheet" type="text/css" href="style.css">',
            '</head>',
            '<body>',
            '<table>',
            '<tr>'
            ]


for row in ws.iter_rows(min_row=1, min_col=1, max_row=51, max_col=32):
    for cell in row:
        with open(f'./Month/{month}/{str(cell)[(str(cell).find(".")+1):(len(str(cell))-1)]}.txt', 'w') as myfile:
            if cell.value == None:
                myfile.write("")
                webspace.append('<td>')
                webspace.append("")
            else:
                if str(cell.value)[0:4] == "2019":
                    myfile.write(str(cell.value)[8:10])
                    webspace.append('<td width="25px">')
                    webspace.append(str(cell.value)[8:10])
                else:
                    if str(cell)[16:17] == "A" and re.match(r'^\d{2}\:\d{2}\:\d{2}$',str(cell.value)):
                        webspace.append('<td width="75px">')
                        myfile.write(str(cell.value)[0:5])
                        webspace.append(str(cell.value)[0:5])
                    else:
                        webspace.append('<td class="content">')
                        webspace.append(f'<a href ="" class="tip">{str(cell.value)}'
                                        f'<span>{str(cell.comment)}</span></a>')
            webspace.append('</td>')

    webspace.append('</tr>')

webspace.append('</body>')
webspace.append('</html>')


webspace.insert(11, f"<td colspan='31'>{month[0:len(month)-2]}</td></tr><tr><td>")
for i in range(13,109):
    del webspace[13]
webspace[12] = ""


with open(f'MapaGM.html', 'w') as htmlgenerator:
    for line in webspace:
        print(line)
        htmlgenerator.write(line)
