### pulling it all together
print "Hello xxxx ! would you like to play a game?"
print " "
print 'Working.....'
print '...'
import os
import win32com.client
import time
import shutil

#xxx denotes user must change

emerlist = []
folder = "xxx"
outlook = win32com.client.Dispatch('Outlook.Application').GetNameSpace("MAPI")
inbox = outlook.GetDefaultFolder(6)
cbyd = inbox.Folders("CBYD")

messages = cbyd.Items  # change to cbyd.Items for cbyd folder
message = messages.GetLast()
att = message.Attachments
attnum = att.count
date = time.strftime("%x")
  

print "there are %s attachments this morning!" %attnum


for i in os.listdir(folder):
    i = os.path.join(folder, i)
    os.remove(i)
    
print "Folder Cleared"

## gets attachments and moves to folder
i=1
while i <= attnum:
    attachment = att.Item(i)
    name = attachment.FileName
    fpath = os.path.join(folder, name)
    attachment.SaveAsFile(fpath)
    i+=1

## converts to a txt file
for x in os.listdir(folder):
    if x == "Thumbs.db":
        break
    else:
        
        p = os.path.join(folder,x)
        a = r"%s" %p
        #msg = open(p)
        msg = outlook.OpenShareditem(p)
        q = msg.body
        spath = p.replace('.msg', '.txt')
        xtxt = x.replace('.msg', '.txt')
        openpath = os.path.join(folder, x)
        openpath = openpath.replace('.msg', '.txt')
        ntxt = open(openpath, 'w')
        ntxt.write(q)
        ntxt.close()
        del msg
del outlook    
# need to grab correct information

xl = win32com.client.Dispatch('Excel.Application')
xl.Visible=1
workbook = xl.Workbooks.Open("xxx")
ws = workbook.Worksheets("All Tickets")


for u in range(1,1000000):
    v = xl.Cells(u,1).Value
    if v == None:
        wx = u
        break

for x in os.listdir(folder):
    sx = os.path.splitext(x)
    filename = sx[0]
    ext = sx[1]
    if ext == '.msg':
        continue
    elif filename[:17] == 'Free Form Request':
        continue
    elif filename[:18] == 'Good Night Message':
        continue
    elif ext == '.txt':
        opath = os.path.join(folder,x)
        ofile = open(opath, 'r')
        lines = ofile.readlines()
        remote0 = lines[5][7:14]
        remote1 = lines[6][7:14]
        remote2 = lines[7][7:14]
        remote3 = lines[8][7:14]
        remote4 = lines[9][7:14]
        remote5 = lines[10][7:14]
        if (remote0 or remote1) == 'ROUTINE':
            ##Date recieved
            if lines[7][0:4] == 'TIME':
                rec = lines[7][19:29]
                xl.Cells(wx,8).Value = rec
            elif lines[8][0:4] == 'TIME':
                rec = lines[8][19:29]
                xl.Cells(wx,8).Value = rec
            elif lines[9][0:4] == 'TIME':
                rec = lines[9][19:29]
                xl.Cells(wx,8).Value = rec
            elif lines[10][0:4] == 'TIME':
                rec = lines[10][19:29]
                xl.Cells(wx,8).Value = rec
            ## REQUEST NUMBER
            if lines[9][0:10] == 'REQUEST NO':
                requestno = lines[9][13:]
                xl.Cells(wx,1).Value = requestno
            elif lines[10][0:10] == 'REQUEST NO':
                requestno = lines[10][13:]
                xl.Cells(wx,1).Value = requestno
            elif lines[11][0:10] == 'REQUEST NO':
                requestno = lines[11][13:]
                xl.Cells(wx,1).Value = requestno
            elif lines[12][0:10] == 'REQUEST NO':
                requestno = lines[12][13:]
                xl.Cells(wx,1).Value = requestno
            ## lat and long
            if lines[11][0:8] == 'LATITUDE':
                lat = lines[11][10:19]
                lon = lines[11][31:45]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            elif lines[12][0:8] == 'LATITUDE':
                lat = lines[12][10:19]
                lon = lines[12][31:45]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            elif lines[13][0:8] == 'LATITUDE':
                lat = lines[13][10:19]
                lon = lines[13][31:45]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            elif lines[14][0:8] == 'LATITUDE':
                lat = lines[14][10:19]
                lon = lines[14][31:45]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            ## town
            if lines[13][0:4] == 'TOWN':
                town = lines[13][10:25]
                xl.Cells(wx,4).Value = town
            elif lines[14][0:4] == 'TOWN':
                town = lines[14][10:25]
                xl.Cells(wx,4).Value = town
            elif lines[15][0:4] == 'TOWN':
                town = lines[15][10:25]
                xl.Cells(wx,4).Value = town
            elif lines[16][0:4] == 'TOWN':
                town = lines[16][10:25]
                xl.Cells(wx,4).Value = town
            ## Address
            if lines[15][0:7] == 'ADDRESS':
                add = lines[15][10:20]
                xl.Cells(wx,5).Value = add
            elif lines[16][0:7] == 'ADDRESS':
                add = lines[16][10:20]
                xl.Cells(wx,5).Value = add
            elif lines[17][0:7] == 'ADDRESS':
                add = lines[17][10:20]
                xl.Cells(wx,5).Value = add
            elif lines[18][0:7] == 'ADDRESS':
                add = lines[19][10:20]
                xl.Cells(wx,5).Value = add
            ## Street
            if lines[16][0:6] == 'STREET':
                street = lines[16][10:40]
                xl.Cells(wx,6).Value = street
            elif lines[17][0:6] == 'STREET':
                street = lines[17][10:40]
                xl.Cells(wx,6).Value = street
            elif lines[18][0:6] == 'STREET':
                street = lines[18][10:40]
                xl.Cells(wx,6).Value = street
            elif lines[19][0:6] == 'STREET':
                street = lines[19][10:40]
                xl.Cells(wx,6).Value = street
            ## work description
            if lines[21][0:4] == 'TYPE':
                work = lines[21][17:]
                xl.Cells(wx,12).Value = work
            elif lines[22][0:4] == 'TYPE':
                work = lines[22][17:]
                xl.Cells(wx,12).Value = work
            elif lines[23][0:4] == 'TYPE':
                work = lines[23][17:]
                xl.Cells(wx,12).Value = work
            elif lines[24][0:4] == 'TYPE':
                work = lines[24][17:]
                xl.Cells(wx,12).Value = work
            ## remarks
            if lines[25][0:7] == 'REMARKS':
                remarks = lines[26]
                xl.Cells(wx,13).Value = remarks
            elif lines[26][0:7] == 'REMARKS':
                remarks = lines[27]
                xl.Cells(wx,13).Value = remarks
            elif lines[27][0:7] == 'REMARKS':
                remarks = lines[28]
                xl.Cells(wx,13).Value = remarks
            elif lines[28][0:7] == 'REMARKS':
                remarks = lines[29]
                xl.Cells(wx,13).Value = remarks
            ## start date
            if lines[28][0:5] == 'START':
                startdate = lines[28][16:26]
                xl.Cells(wx,7).Value = startdate
            elif lines[29][0:5] == 'START':
                startdate = lines[29][16:26]
                xl.Cells(wx,7).Value = startdate
            elif lines[30][0:5] == 'START':
                startdate = lines[30][16:26]
                xl.Cells(wx,7).Value = startdate
            elif lines[31][0:5] == 'START':
                startdate = lines[31][16:26]
                xl.Cells(wx,7).Value = startdate
            ## caller
            if lines[30][0:6] == 'CALLER':
                caller = lines[30][16:40]
                xl.Cells(wx,9).Value = caller
            elif lines[31][0:6] == 'CALLER':
                caller = lines[31][16:40]
                xl.Cells(wx,9).Value = caller
            elif lines[32][0:6] == 'CALLER':
                caller = lines[32][16:40]
                xl.Cells(wx,9).Value = caller
            elif lines[33][0:6] == 'CALLER':
                caller = lines[33][16:40]
                xl.Cells(wx,9).Value = caller
            ##number
            if lines[32][0:5] == 'PHONE':
                number = lines[32][16:30]
                xl.Cells(wx,10).Value = number
            elif lines[33][0:5] == 'PHONE':
                number = lines[33][16:30]
                xl.Cells(wx,10).Value = number
            elif lines[34][0:5] == 'PHONE':
                number = lines[34][16:30]
                xl.Cells(wx,10).Value = number
            elif lines[35][0:5] == 'PHONE':
                number = lines[35][16:30]
                xl.Cells(wx,10).Value = number
            ## contractor
            if lines[36][0:10] == 'CONTRACTOR':
                contractor = lines[36][16:50]
                xl.Cells(wx,11).Value = contractor
            elif lines[37][0:10] == 'CONTRACTOR':
                contractor = lines[37][16:50]
                xl.Cells(wx,11).Value = contractor
            elif lines[38][0:10] == 'CONTRACTOR':
                contractor = lines[38][16:50]
                xl.Cells(wx,11).Value = contractor
            elif lines[39][0:10] == 'CONTRACTOR':
                contractor = lines[39][16:50]
                xl.Cells(wx,11).Value = contractor
            ws.Range("2:1000000").RowHeight = 15
            ws.Range('2:1000000').VerticalAlignment = 1
            ws.Range('2:1000000').HorizontalAlignment = 1
            wx = wx + 1
            ofile.close()
        elif remote2 == 'ROUTINE':
            ## daterecieved
            if lines[8][0:4] == 'TIME':
                rec = lines[8][19:29]
                xl.Cells(wx,8).Value = rec
            elif lines[9][0:4] == 'TIME':
                rec = lines[9][19:29]
                xl.Cells(wx,8).Value = rec
            elif lines[10][0:4] == 'TIME':
                rec = lines[10][19:29]
                xl.Cells(wx,8).Value = rec
            elif lines[11][0:4] == 'TIME':
                rec = lines[11][19:29]
                xl.Cells(wx,8).Value = rec
            ## REQUEST NUMBER
            if lines[10][0:10] == 'REQUEST NO':
                requestno = lines[10][13:]
                xl.Cells(wx,1).Value = requestno
            elif lines[11][0:10] == 'REQUEST NO':
                requestno = lines[11][13:]
                xl.Cells(wx,1).Value = requestno
            elif lines[12][0:10] == 'REQUEST NO':
                requestno = lines[12][13:]
                xl.Cells(wx,1).Value = requestno
            elif lines[13][0:10] == 'REQUEST NO':
                requestno = lines[13][13:]
                xl.Cells(wx,1).Value = requestno
            ## lat and long
            if lines[12][0:8] == 'LATITUDE':
                lat = lines[12][10:19]
                lon = lines[12][31:45]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            elif lines[13][0:8] == 'LATITUDE':
                lat = lines[13][10:19]
                lon = lines[13][31:45]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            elif lines[14][0:8] == 'LATITUDE':
                lat = lines[14][10:19]
                lon = lines[14][31:45]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            elif lines[15][0:8] == 'LATITUDE':
                lat = lines[15][10:19]
                lon = lines[15][31:45]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            ## town
            if lines[14][0:4] == 'TOWN':
                town = lines[14][10:25]
                xl.Cells(wx,4).Value = town
            elif lines[15][0:4] == 'TOWN':
                town = lines[15][10:25]
                xl.Cells(wx,4).Value = town
            elif lines[16][0:4] == 'TOWN':
                town = lines[16][10:25]
                xl.Cells(wx,4).Value = town
            elif lines[17][0:4] == 'TOWN':
                town = lines[17][10:25]
                xl.Cells(wx,4).Value = town
            ## Address
            if lines[16][0:7] == 'ADDRESS':
                add = lines[16][10:20]
                xl.Cells(wx,5).Value = add
            elif lines[17][0:7] == 'ADDRESS':
                add = lines[17][10:20]
                xl.Cells(wx,5).Value = add
            elif lines[18][0:7] == 'ADDRESS':
                add = lines[18][10:20]
                xl.Cells(wx,5).Value = add
            elif lines[19][0:7] == 'ADDRESS':
                add = lines[19][10:20]
                xl.Cells(wx,5).Value = add
            ## Street
            if lines[17][0:6] == 'STREET':
                street = lines[17][10:40]
                xl.Cells(wx,6).Value = street
            elif lines[18][0:6] == 'STREET':
                street = lines[18][10:40]
                xl.Cells(wx,6).Value = street
            elif lines[19][0:6] == 'STREET':
                street = lines[19][10:40]
                xl.Cells(wx,6).Value = street
            elif lines[20][0:6] == 'STREET':
                street = lines[20][10:40]
                xl.Cells(wx,6).Value = street
            ## work description
            if lines[22][0:4] == 'TYPE':
                work = lines[22][17:]
                xl.Cells(wx,12).Value = work
            elif lines[23][0:4] == 'TYPE':
                work = lines[23][17:]
                xl.Cells(wx,12).Value = work
            elif lines[24][0:4] == 'TYPE':
                work = lines[24][17:]
                xl.Cells(wx,12).Value = work
            elif lines[25][0:4] == 'TYPE':
                work = lines[25][17:]
                xl.Cells(wx,12).Value = work
            ## remarks
            if lines[26][0:7] == 'REMARKS':
                remarks = lines[27]
                xl.Cells(wx,13).Value = remarks
            elif lines[27][0:7] == 'REMARKS':
                remarks = lines[28]
                xl.Cells(wx,13).Value = remarks
            elif lines[28][0:7] == 'REMARKS':
                remarks = lines[29]
                xl.Cells(wx,13).Value = remarks
            elif lines[29][0:7] == 'REMARKS':
                remarks = lines[30]
                xl.Cells(wx,13).Value = remarks
            ## start date
            if lines[29][0:5] == 'START':
                startdate = lines[29][16:26]
                xl.Cells(wx,7).Value = startdate
            elif lines[30][0:5] == 'START':
                startdate = lines[30][16:26]
                xl.Cells(wx,7).Value = startdate
            elif lines[31][0:5] == 'START':
                startdate = lines[31][16:26]
                xl.Cells(wx,7).Value = startdate
            elif lines[32][0:5] == 'START':
                startdate = lines[32][16:26]
                xl.Cells(wx,7).Value = startdate
            ## caller
            if lines[31][0:6] == 'CALLER':
                caller = lines[31][16:40]
                xl.Cells(wx,9).Value = caller
            elif lines[32][0:6] == 'CALLER':
                caller = lines[32][16:40]
                xl.Cells(wx,9).Value = caller
            elif lines[33][0:6] == 'CALLER':
                caller = lines[33][16:40]
                xl.Cells(wx,9).Value = caller
            elif lines[34][0:6] == 'CALLER':
                caller = lines[34][16:40]
                xl.Cells(wx,9).Value = caller
            ##number
            if lines[33][0:5] == 'PHONE':
                number = lines[33][16:30]
                xl.Cells(wx,10).Value = number
            elif lines[34][0:5] == 'PHONE':
                number = lines[34][16:30]
                xl.Cells(wx,10).Value = number
            elif lines[35][0:5] == 'PHONE':
                number = lines[35][16:30]
                xl.Cells(wx,10).Value = number
            elif lines[36][0:5] == 'PHONE':
                number = lines[36][16:30]
                xl.Cells(wx,10).Value = number
            ## contractor
            if lines[37][0:10] == 'CONTRACTOR':
                contractor = lines[37][16:50]
                xl.Cells(wx,11).Value = contractor
            elif lines[38][0:10] == 'CONTRACTOR':
                contractor = lines[38][16:50]
                xl.Cells(wx,11).Value = contractor
            elif lines[39][0:10] == 'CONTRACTOR':
                contractor = lines[39][16:50]
                xl.Cells(wx,11).Value = contractor
            elif lines[40][0:10] == 'CONTRACTOR':
                contractor = lines[40][16:50]
                xl.Cells(wx,11).Value = contractor
            ws.Range("2:1000000").RowHeight = 15
            ws.Range('2:1000000').VerticalAlignment = 1
            ws.Range('2:1000000').HorizontalAlignment = 1
            wx = wx + 1
            ofile.close()
        elif remote3 == 'ROUTINE':
            ## daterecieved
            if lines[9][0:4] == 'TIME':
                rec = lines[9][19:29]
                xl.Cells(wx,8).Value = rec
            elif lines[10][0:4] == 'TIME':
                rec = lines[10][19:29]
                xl.Cells(wx,8).Value = rec
            elif lines[11][0:4] == 'TIME':
                rec = lines[11][19:29]
                xl.Cells(wx,8).Value = rec
            elif lines[12][0:4] == 'TIME':
                rec = lines[12][19:29]
                xl.Cells(wx,8).Value = rec
            ## REQUEST NUMBER
            if lines[11][0:10] == 'REQUEST NO':
                requestno = lines[11][13:]
                xl.Cells(wx,1).Value = requestno
            elif lines[12][0:10] == 'REQUEST NO':
                requestno = lines[12][13:]
                xl.Cells(wx,1).Value = requestno
            elif lines[13][0:10] == 'REQUEST NO':
                requestno = lines[13][13:]
                xl.Cells(wx,1).Value = requestno
            elif lines[14][0:10] == 'REQUEST NO':
                requestno = lines[14][13:]
                xl.Cells(wx,1).Value = requestno
            ## lat and long
            if lines[13][0:8] == 'LATITUDE':
                lat = lines[13][10:19]
                lon = lines[13][31:45]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            elif lines[14][0:8] == 'LATITUDE':
                lat = lines[14][10:19]
                lon = lines[14][31:45]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            elif lines[15][0:8] == 'LATITUDE':
                lat = lines[15][10:19]
                lon = lines[15][31:45]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            elif lines[16][0:8] == 'LATITUDE':
                lat = lines[16][10:19]
                lon = lines[16][31:45]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            ## town
            if lines[15][0:4] == 'TOWN':
                town = lines[15][10:25]
                xl.Cells(wx,4).Value = town
            elif lines[16][0:4] == 'TOWN':
                town = lines[16][10:25]
                xl.Cells(wx,4).Value = town
            elif lines[17][0:4] == 'TOWN':
                town = lines[17][10:25]
                xl.Cells(wx,4).Value = town
            elif lines[18][0:4] == 'TOWN':
                town = lines[18][10:25]
                xl.Cells(wx,4).Value = town
            ## Address
            if lines[17][0:7] == 'ADDRESS':
                add = lines[17][10:20]
                xl.Cells(wx,5).Value = add
            elif lines[18][0:7] == 'ADDRESS':
                add = lines[18][10:20]
                xl.Cells(wx,5).Value = add
            elif lines[19][0:7] == 'ADDRESS':
                add = lines[19][10:20]
                xl.Cells(wx,5).Value = add
            elif lines[20][0:7] == 'ADDRESS':
                add = lines[20][10:20]
                xl.Cells(wx,5).Value = add
            ## Street
            if lines[18][0:6] == 'STREET':
                street = lines[18][10:40]
                xl.Cells(wx,6).Value = street
            elif lines[19][0:6] == 'STREET':
                street = lines[19][10:40]
                xl.Cells(wx,6).Value = street
            elif lines[20][0:6] == 'STREET':
                street = lines[20][10:40]
                xl.Cells(wx,6).Value = street
            elif lines[21][0:6] == 'STREET':
                street = lines[21][10:40]
                xl.Cells(wx,6).Value = street
            ## work description
            if lines[23][0:4] == 'TYPE':
                work = lines[23][17:]
                xl.Cells(wx,12).Value = work
            elif lines[24][0:4] == 'TYPE':
                work = lines[24][17:]
                xl.Cells(wx,12).Value = work
            elif lines[25][0:4] == 'TYPE':
                work = lines[25][17:]
                xl.Cells(wx,12).Value = work
            elif lines[26][0:4] == 'TYPE':
                work = lines[26][17:]
                xl.Cells(wx,12).Value = work
            ## remarks
            if lines[27][0:7] == 'REMARKS':
                remarks = lines[28]
                xl.Cells(wx,13).Value = remarks
            elif lines[28][0:7] == 'REMARKS':
                remarks = lines[29]
                xl.Cells(wx,13).Value = remarks
            elif lines[29][0:7] == 'REMARKS':
                remarks = lines[30]
                xl.Cells(wx,13).Value = remarks
            elif lines[30][0:7] == 'REMARKS':
                remarks = lines[31]
                xl.Cells(wx,13).Value = remarks
            ## start date
            if lines[30][0:5] == 'START':
                startdate = lines[30][16:26]
                xl.Cells(wx,7).Value = startdate
            elif lines[31][0:5] == 'START':
                startdate = lines[31][16:26]
                xl.Cells(wx,7).Value = startdate
            elif lines[32][0:5] == 'START':
                startdate = lines[32][16:26]
                xl.Cells(wx,7).Value = startdate
            elif lines[33][0:5] == 'START':
                startdate = lines[33][16:26]
                xl.Cells(wx,7).Value = startdate
            ## caller
            if lines[32][0:6] == 'CALLER':
                caller = lines[32][16:40]
                xl.Cells(wx,9).Value = caller
            elif lines[33][0:6] == 'CALLER':
                caller = lines[33][16:40]
                xl.Cells(wx,9).Value = caller
            elif lines[34][0:6] == 'CALLER':
                caller = lines[34][16:40]
                xl.Cells(wx,9).Value = caller
            elif lines[35][0:6] == 'CALLER':
                caller = lines[35][16:40]
                xl.Cells(wx,9).Value = caller
            ##number
            if lines[34][0:5] == 'PHONE':
                number = lines[34][16:30]
                xl.Cells(wx,10).Value = number
            elif lines[35][0:5] == 'PHONE':
                number = lines[35][16:30]
                xl.Cells(wx,10).Value = number
            elif lines[36][0:5] == 'PHONE':
                number = lines[36][16:30]
                xl.Cells(wx,10).Value = number
            elif lines[37][0:5] == 'PHONE':
                number = lines[37][16:30]
                xl.Cells(wx,10).Value = number
            ## contractor
            if lines[38][0:10] == 'CONTRACTOR':
                contractor = lines[38][16:50]
                xl.Cells(wx,11).Value = contractor
            elif lines[39][0:10] == 'CONTRACTOR':
                contractor = lines[39][16:50]
                xl.Cells(wx,11).Value = contractor
            elif lines[40][0:10] == 'CONTRACTOR':
                contractor = lines[40][16:50]
                xl.Cells(wx,11).Value = contractor
            elif lines[41][0:10] == 'CONTRACTOR':
                contractor = lines[41][16:50]
                xl.Cells(wx,11).Value = contractor
            ws.Range("2:1000000").RowHeight = 15
            ws.Range('2:1000000').VerticalAlignment = 1
            ws.Range('2:1000000').HorizontalAlignment = 1
            wx = wx + 1
            ofile.close()
        elif remote4 == 'ROUTINE':
            ## daterecieved
            if lines[10][0:4] == 'TIME':
                rec = lines[10][19:29]
                xl.Cells(wx,8).Value = rec
            elif lines[11][0:4] == 'TIME':
                rec = lines[11][19:29]
                xl.Cells(wx,8).Value = rec
            elif lines[12][0:4] == 'TIME':
                rec = lines[12][19:29]
                xl.Cells(wx,8).Value = rec
            elif lines[13][0:4] == 'TIME':
                rec = lines[13][19:29]
                xl.Cells(wx,8).Value = rec
            ## REQUEST NUMBER
            if lines[12][0:10] == 'REQUEST NO':
                requestno = lines[12][13:]
                xl.Cells(wx,1).Value = requestno
            elif lines[13][0:10] == 'REQUEST NO':
                requestno = lines[13][13:]
                xl.Cells(wx,1).Value = requestno
            elif lines[14][0:10] == 'REQUEST NO':
                requestno = lines[14][13:]
                xl.Cells(wx,1).Value = requestno
            elif lines[15][0:10] == 'REQUEST NO':
                requestno = lines[15][13:]
                xl.Cells(wx,1).Value = requestno
            ## lat and long
            if lines[14][0:8] == 'LATITUDE':
                lat = lines[14][10:19]
                lon = lines[14][31:45]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            elif lines[15][0:8] == 'LATITUDE':
                lat = lines[15][10:19]
                lon = lines[15][31:45]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            elif lines[16][0:8] == 'LATITUDE':
                lat = lines[16][10:19]
                lon = lines[16][31:45]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            elif lines[17][0:8] == 'LATITUDE':
                lat = lines[17][10:19]
                lon = lines[17][31:45]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            ## town
            if lines[16][0:4] == 'TOWN':
                town = lines[16][10:25]
                xl.Cells(wx,4).Value = town
            elif lines[17][0:4] == 'TOWN':
                town = lines[17][10:25]
                xl.Cells(wx,4).Value = town
            elif lines[18][0:4] == 'TOWN':
                town = lines[18][10:25]
                xl.Cells(wx,4).Value = town
            elif lines[19][0:4] == 'TOWN':
                town = lines[19][10:25]
                xl.Cells(wx,4).Value = town
            ## Address
            if lines[18][0:7] == 'ADDRESS':
                add = lines[18][10:20]
                xl.Cells(wx,5).Value = add
            elif lines[19][0:7] == 'ADDRESS':
                add = lines[19][10:20]
                xl.Cells(wx,5).Value = add
            elif lines[20][0:7] == 'ADDRESS':
                add = lines[20][10:20]
                xl.Cells(wx,5).Value = add
            elif lines[21][0:7] == 'ADDRESS':
                add = lines[21][10:20]
                xl.Cells(wx,5).Value = add
            ## Street
            if lines[19][0:6] == 'STREET':
                street = lines[19][10:40]
                xl.Cells(wx,6).Value = street
            elif lines[20][0:6] == 'STREET':
                street = lines[20][10:40]
                xl.Cells(wx,6).Value = street
            elif lines[21][0:6] == 'STREET':
                street = lines[21][10:40]
                xl.Cells(wx,6).Value = street
            elif lines[22][0:6] == 'STREET':
                street = lines[22][10:40]
                xl.Cells(wx,6).Value = street
            ## work description
            if lines[24][0:4] == 'TYPE':
                work = lines[24][17:]
                xl.Cells(wx,12).Value = work
            elif lines[25][0:4] == 'TYPE':
                work = lines[25][17:]
                xl.Cells(wx,12).Value = work
            elif lines[26][0:4] == 'TYPE':
                work = lines[26][17:]
                xl.Cells(wx,12).Value = work
            elif lines[27][0:4] == 'TYPE':
                work = lines[27][17:]
                xl.Cells(wx,12).Value = work
            ## remarks
            if lines[28][0:7] == 'REMARKS':
                remarks = lines[29]
                xl.Cells(wx,13).Value = remarks
            elif lines[29][0:7] == 'REMARKS':
                remarks = lines[30]
                xl.Cells(wx,13).Value = remarks
            elif lines[30][0:7] == 'REMARKS':
                remarks = lines[31]
                xl.Cells(wx,13).Value = remarks
            elif lines[31][0:7] == 'REMARKS':
                remarks = lines[32]
                xl.Cells(wx,13).Value = remarks
            ## start date
            if lines[31][0:5] == 'START':
                startdate = lines[31][16:26]
                xl.Cells(wx,7).Value = startdate
            elif lines[32][0:5] == 'START':
                startdate = lines[32][16:26]
                xl.Cells(wx,7).Value = startdate
            elif lines[33][0:5] == 'START':
                startdate = lines[33][16:26]
                xl.Cells(wx,7).Value = startdate
            elif lines[34][0:5] == 'START':
                startdate = lines[34][16:26]
                xl.Cells(wx,7).Value = startdate
            ## caller
            if lines[33][0:6] == 'CALLER':
                caller = lines[33][16:40]
                xl.Cells(wx,9).Value = caller
            elif lines[34][0:6] == 'CALLER':
                caller = lines[34][16:40]
                xl.Cells(wx,9).Value = caller
            elif lines[35][0:6] == 'CALLER':
                caller = lines[35][16:40]
                xl.Cells(wx,9).Value = caller
            elif lines[36][0:6] == 'CALLER':
                caller = lines[36][16:40]
                xl.Cells(wx,9).Value = caller
            ##number
            if lines[35][0:5] == 'PHONE':
                number = lines[35][16:30]
                xl.Cells(wx,10).Value = number
            elif lines[36][0:5] == 'PHONE':
                number = lines[36][16:30]
                xl.Cells(wx,10).Value = number
            elif lines[37][0:5] == 'PHONE':
                number = lines[37][16:30]
                xl.Cells(wx,10).Value = number
            elif lines[38][0:5] == 'PHONE':
                number = lines[38][16:30]
                xl.Cells(wx,10).Value = number
            ## contractor
            if lines[39][0:10] == 'CONTRACTOR':
                contractor = lines[39][16:50]
                xl.Cells(wx,11).Value = contractor
            elif lines[40][0:10] == 'CONTRACTOR':
                contractor = lines[40][16:50]
                xl.Cells(wx,11).Value = contractor
            elif lines[41][0:10] == 'CONTRACTOR':
                contractor = lines[41][16:50]
                xl.Cells(wx,11).Value = contractor
            elif lines[42][0:10] == 'CONTRACTOR':
                contractor = lines[42][16:50]
                xl.Cells(wx,11).Value = contractor
            ws.Range("2:1000000").RowHeight = 15
            ws.Range('2:1000000').VerticalAlignment = 1
            ws.Range('2:1000000').HorizontalAlignment = 1
            wx = wx + 1
            ofile.close()
        elif remote5 == 'ROUTINE':
            ## daterecieved
            if lines[11][0:4] == 'TIME':
                rec = lines[11][19:29]
                xl.Cells(wx,8).Value = rec
            elif lines[12][0:4] == 'TIME':
                rec = lines[12][19:29]
                xl.Cells(wx,8).Value = rec
            elif lines[13][0:4] == 'TIME':
                rec = lines[13][19:29]
                xl.Cells(wx,8).Value = rec
            elif lines[14][0:4] == 'TIME':
                rec = lines[14][19:29]
                xl.Cells(wx,8).Value = rec
            ## REQUEST NUMBER
            if lines[13][0:10] == 'REQUEST NO':
                requestno = lines[13][13:]
                xl.Cells(wx,1).Value = requestno
            elif lines[14][0:10] == 'REQUEST NO':
                requestno = lines[14][13:]
                xl.Cells(wx,1).Value = requestno
            elif lines[15][0:10] == 'REQUEST NO':
                requestno = lines[15][13:]
                xl.Cells(wx,1).Value = requestno
            elif lines[16][0:10] == 'REQUEST NO':
                requestno = lines[16][13:]
                xl.Cells(wx,1).Value = requestno
            ## lat and long
            if lines[15][0:8] == 'LATITUDE':
                lat = lines[15][10:19]
                lon = lines[15][31:45]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            elif lines[16][0:8] == 'LATITUDE':
                lat = lines[16][10:19]
                lon = lines[16][31:45]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            elif lines[17][0:8] == 'LATITUDE':
                lat = lines[17][10:19]
                lon = lines[17][31:45]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            elif lines[18][0:8] == 'LATITUDE':
                lat = lines[18][10:19]
                lon = lines[18][31:45]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            ## town
            if lines[17][0:4] == 'TOWN':
                town = lines[17][10:25]
                xl.Cells(wx,4).Value = town
            elif lines[18][0:4] == 'TOWN':
                town = lines[18][10:25]
                xl.Cells(wx,4).Value = town
            elif lines[19][0:4] == 'TOWN':
                town = lines[19][10:25]
                xl.Cells(wx,4).Value = town
            elif lines[20][0:4] == 'TOWN':
                town = lines[20][10:25]
                xl.Cells(wx,4).Value = town
            ## Address
            if lines[19][0:7] == 'ADDRESS':
                add = lines[19][10:20]
                xl.Cells(wx,5).Value = add
            elif lines[20][0:7] == 'ADDRESS':
                add = lines[20][10:20]
                xl.Cells(wx,5).Value = add
            elif lines[21][0:7] == 'ADDRESS':
                add = lines[21][10:20]
                xl.Cells(wx,5).Value = add
            elif lines[22][0:7] == 'ADDRESS':
                add = lines[22][10:20]
                xl.Cells(wx,5).Value = add
            ## Street
            if lines[20][0:6] == 'STREET':
                street = lines[20][10:40]
                xl.Cells(wx,6).Value = street
            elif lines[21][0:6] == 'STREET':
                street = lines[21][10:40]
                xl.Cells(wx,6).Value = street
            elif lines[22][0:6] == 'STREET':
                street = lines[22][10:40]
                xl.Cells(wx,6).Value = street
            elif lines[23][0:6] == 'STREET':
                street = lines[23][10:40]
                xl.Cells(wx,6).Value = street
            ## work description
            if lines[25][0:4] == 'TYPE':
                work = lines[25][17:]
                xl.Cells(wx,12).Value = work
            elif lines[26][0:4] == 'TYPE':
                work = lines[26][17:]
                xl.Cells(wx,12).Value = work
            elif lines[27][0:4] == 'TYPE':
                work = lines[27][17:]
                xl.Cells(wx,12).Value = work
            elif lines[28][0:4] == 'TYPE':
                work = lines[28][17:]
                xl.Cells(wx,12).Value = work
            ## remarks
            if lines[29][0:7] == 'REMARKS':
                remarks = lines[30]
                xl.Cells(wx,13).Value = remarks
            elif lines[30][0:7] == 'REMARKS':
                remarks = lines[31]
                xl.Cells(wx,13).Value = remarks
            elif lines[31][0:7] == 'REMARKS':
                remarks = lines[32]
                xl.Cells(wx,13).Value = remarks
            elif lines[32][0:7] == 'REMARKS':
                remarks = lines[33]
                xl.Cells(wx,13).Value = remarks
            ## start date
            if lines[32][0:5] == 'START':
                startdate = lines[32][16:26]
                xl.Cells(wx,7).Value = startdate
            elif lines[33][0:5] == 'START':
                startdate = lines[33][16:26]
                xl.Cells(wx,7).Value = startdate
            elif lines[34][0:5] == 'START':
                startdate = lines[34][16:26]
                xl.Cells(wx,7).Value = startdate
            elif lines[35][0:5] == 'START':
                startdate = lines[35][16:26]
                xl.Cells(wx,7).Value = startdate
            ## caller
            if lines[34][0:6] == 'CALLER':
                caller = lines[34][16:40]
                xl.Cells(wx,9).Value = caller
            elif lines[35][0:6] == 'CALLER':
                caller = lines[35][16:40]
                xl.Cells(wx,9).Value = caller
            elif lines[36][0:6] == 'CALLER':
                caller = lines[36][16:40]
                xl.Cells(wx,9).Value = caller
            elif lines[37][0:6] == 'CALLER':
                caller = lines[37][16:40]
                xl.Cells(wx,9).Value = caller
            ##number
            if lines[36][0:5] == 'PHONE':
                number = lines[36][16:30]
                xl.Cells(wx,10).Value = number
            elif lines[37][0:5] == 'PHONE':
                number = lines[37][16:30]
                xl.Cells(wx,10).Value = number
            elif lines[38][0:5] == 'PHONE':
                number = lines[38][16:30]
                xl.Cells(wx,10).Value = number
            elif lines[39][0:5] == 'PHONE':
                number = lines[39][16:30]
                xl.Cells(wx,10).Value = number
            ## contractor
            if lines[40][0:10] == 'CONTRACTOR':
                contractor = lines[40][16:50]
                xl.Cells(wx,11).Value = contractor
            elif lines[41][0:10] == 'CONTRACTOR':
                contractor = lines[41][16:50]
                xl.Cells(wx,11).Value = contractor
            elif lines[42][0:10] == 'CONTRACTOR':
                contractor = lines[42][16:50]
                xl.Cells(wx,11).Value = contractor
            elif lines[43][0:10] == 'CONTRACTOR':
                contractor = lines[43][16:50]
                xl.Cells(wx,11).Value = contractor
            ws.Range("2:1000000").RowHeight = 15
            ws.Range('2:1000000').VerticalAlignment = 1
            ws.Range('2:1000000').HorizontalAlignment = 1
            wx = wx + 1
            ofile.close()
        else:
            print "This is an emergency %s" %filename
            emerlist.append(filename)
            ## daterecieved
            if lines[6][0:4] == 'TIME':
                rec = lines[6][19:29]
                xl.Cells(wx,8).Value = rec
            elif lines[7][0:4] == 'TIME':
                rec = lines[7][19:29]
                xl.Cells(wx,8).Value = rec
            elif lines[8][0:4] == 'TIME':
                rec = lines[8][19:29]
                xl.Cells(wx,8).Value = rec
            elif lines[9][0:4] == 'TIME':
                rec = lines[9][19:29]
                xl.Cells(wx,8).Value = rec
            ## REQUEST NUMBER
            if lines[9][0:10] == 'REQUEST NO':
                requestno = lines[9][13:25]
                xl.Cells(wx,1).Value = requestno
            elif lines[10][0:10] == 'REQUEST NO':
                requestno = lines[10][13:25]
                xl.Cells(wx,1).Value = requestno
            elif lines[11][0:10] == 'REQUEST NO':
                requestno = lines[11][13:25]
                xl.Cells(wx,1).Value = requestno
            elif lines[12][0:10] == 'REQUEST NO':
                requestno = lines[12][13:25]
                xl.Cells(wx,1).Value = requestno
            ## lat and long
            if lines[11][0:8] == 'LATITUDE':
                lat = lines[11][10:19]
                lon = lines[11][31:41]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            elif lines[12][0:8] == 'LATITUDE':
                lat = lines[12][10:19]
                lon = lines[12][31:41]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            elif lines[13][0:8] == 'LATITUDE':
                lat = lines[13][10:19]
                lon = lines[13][31:41]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            elif lines[14][0:8] == 'LATITUDE':
                lat = lines[14][10:19]
                lon = lines[14][31:41]
                xl.Cells(wx,2).Value = lat
                xl.Cells(wx,3).Value = lon
            ## town
            if lines[13][0:4] == 'TOWN':
                town = lines[13][10:22]
                xl.Cells(wx,4).Value = town
            elif lines[14][0:4] == 'TOWN':
                town = lines[14][10:22]
                xl.Cells(wx,4).Value = town
            elif lines[15][0:4] == 'TOWN':
                town = lines[15][10:22]
                xl.Cells(wx,4).Value = town
            elif lines[16][0:4] == 'TOWN':
                town = lines[16][10:22]
                xl.Cells(wx,4).Value = town
            ## Address
            if lines[15][0:7] == 'ADDRESS':
                add = lines[15][10:15]
                xl.Cells(wx,5).Value = add
            elif lines[16][0:7] == 'ADDRESS':
                add = lines[16][10:15]
                xl.Cells(wx,5).Value = add
            elif lines[17][0:7] == 'ADDRESS':
                add = lines[17][10:15]
                xl.Cells(wx,5).Value = add
            elif lines[18][0:7] == 'ADDRESS':
                add = lines[18][10:15]
                xl.Cells(wx,5).Value = add
            ## Street
            if lines[16][0:6] == 'STREET':
                street = lines[16][10:40]
                xl.Cells(wx,6).Value = street
            elif lines[17][0:6] == 'STREET':
                street = lines[17][10:40]
                xl.Cells(wx,6).Value = street
            elif lines[18][0:6] == 'STREET':
                street = lines[18][10:40]
                xl.Cells(wx,6).Value = street
            elif lines[19][0:6] == 'STREET':
                street = lines[19][10:40]
                xl.Cells(wx,6).Value = street
            ## work description
            if lines[21][0:4] == 'TYPE':
                work = lines[21][17:]
                xl.Cells(wx,12).Value = work
            elif lines[22][0:4] == 'TYPE':
                work = lines[22][17:]
                xl.Cells(wx,12).Value = work
            elif lines[23][0:4] == 'TYPE':
                work = lines[23][17:]
                xl.Cells(wx,12).Value = work
            elif lines[24][0:4] == 'TYPE':
                work = lines[24][17:]
                xl.Cells(wx,12).Value = work
            ## remarks
            if lines[26][0:7] == 'REMARKS':
                remarks = lines[27]
                xl.Cells(wx,13).Value = remarks
            elif lines[27][0:7] == 'REMARKS':
                remarks = lines[28]
                xl.Cells(wx,13).Value = remarks
            elif lines[28][0:7] == 'REMARKS':
                remarks = lines[29]
                xl.Cells(wx,13).Value = remarks
            elif lines[29][0:7] == 'REMARKS':
                remarks = lines[30]
                xl.Cells(wx,13).Value = remarks
            ## start date
            if lines[28][0:5] == 'START':
                startdate = lines[28][16:26]
                xl.Cells(wx,7).Value = startdate
            elif lines[29][0:5] == 'START':
                startdate = lines[29][16:26]
                xl.Cells(wx,7).Value = startdate
            elif lines[30][0:5] == 'START':
                startdate = lines[30][16:26]
                xl.Cells(wx,7).Value = startdate
            elif lines[31][0:5] == 'START':
                startdate = lines[31][16:26]
                xl.Cells(wx,7).Value = startdate
            ## caller
            if lines[30][0:6] == 'CALLER':
                caller = lines[30][16:40]
                xl.Cells(wx,9).Value = caller
            elif lines[31][0:6] == 'CALLER':
                caller = lines[31][16:40]
                xl.Cells(wx,9).Value = caller
            elif lines[32][0:6] == 'CALLER':
                caller = lines[32][16:40]
                xl.Cells(wx,9).Value = caller
            elif lines[33][0:6] == 'CALLER':
                caller = lines[33][16:40]
                xl.Cells(wx,9).Value = caller
            ##number
            if lines[32][0:5] == 'PHONE':
                number = lines[32][16:30]
                xl.Cells(wx,10).Value = number
            elif lines[33][0:5] == 'PHONE':
                number = lines[33][16:30]
                xl.Cells(wx,10).Value = number
            elif lines[34][0:5] == 'PHONE':
                number = lines[34][16:30]
                xl.Cells(wx,10).Value = number
            elif lines[35][0:5] == 'PHONE':
                number = lines[35][16:30]
                xl.Cells(wx,10).Value = number
            ## contractor
            if lines[36][0:10] == 'CONTRACTOR':
                contractor = lines[36][16:50]
                xl.Cells(wx,11).Value = contractor
            elif lines[37][0:10] == 'CONTRACTOR':
                contractor = lines[37][16:50]
                xl.Cells(wx,11).Value = contractor
            elif lines[38][0:10] == 'CONTRACTOR':
                contractor = lines[38][16:50]
                xl.Cells(wx,11).Value = contractor
            elif lines[39][0:10] == 'CONTRACTOR':
                contractor = lines[39][16:50]
                xl.Cells(wx,11).Value = contractor
            xl.Cells(wx,14).Value = 'EMERGENCY'
            ws.Range("2:1000000").RowHeight = 15
            ws.Range('2:1000000').VerticalAlignment = 1
            ws.Range('2:1000000').HorizontalAlignment = 1
            wx = wx + 1
            emersavepath = "xxx"
            qi = os.path.join(emersavepath, x)
            shutil.copy2(opath,qi)
            
            ofile.close()

for i in os.listdir(folder):
    i = os.path.splitext(i)
    for q in emerlist:
        if q == i[0]:
            continue
        else:
            d = os.path.join(i[0]+i[1])
            d = os.path.join(folder,d)
            os.remove(d)
            



    
workbook.Save()

del xl
del workbook
del ws





print "DUN!!!!"

    


        
