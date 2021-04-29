#python 2.7
#Read number from excel sheet and throw them at L2L


import xlrd                    # module to interface with excel
import requests                # module to upload the data to L2L
from datetime import datetime  # module for time stamps
import time
import base64

file_location = "C:/Users/QABRFPLC/Desktop/L2L_Numbers.xlsm"  # location of excel file on local computer
workbook = xlrd.open_workbook(file_location)
sheet = workbook.sheet_by_index(0)
server = sheet.cell_value(1, 2)
API_key = sheet.cell_value(1, 3)
site = sheet.cell_value(1, 4)
enable_proxies = sheet.cell_value(1, 5)
proxy_username = sheet.cell_value(1, 6)
proxy_password = sheet.cell_value(1, 7)
proxy_ip = sheet.cell_value(1, 8)
proxies = {'https': 'https://' + str(proxy_username) + ':' + str(base64.b64decode(proxy_password)) + '@' + str(proxy_ip)}
Api_End_Point = 'https://' + server + '.leading2lean.com/api/1.0/pitchdetails/record_details/'
Scrap_End_Point = 'https://' + server + '.leading2lean.com/api/1.0/scrapdetail/add_scrap_detail/'
url = Api_End_Point + "?auth=" + API_key
Scrap_Url = Scrap_End_Point + "?auth=" + API_key
DATETIME_SECONDS_STRING_FORMAT = "%Y-%m-%d %H:%M:%S"
firstPass = True


class Line:
    def __init__(self, name, locrow):
        self.name = name
        self.loga = 0
        self.logb = 0
        self.end = 0
        self.oldloga = 0
        self.oldlogb = 0
        self.oldend = 0
        self.diff = 0
        self.actual = 0
        self.scrap = 0
        self.response = ""
        self.locrow = locrow
        self.loclogacol = 3
        self.loclogbcol = 4
        self.locendcol = 10
        self.locbotcol = 2
        self.linecode = name
        self.scrap_machines = sheet.cell_value(locrow, 11)
        self.bot = ""
        self.therm = False
        self.getnew()

    def zero(self):
        print("line class zero called")
        self.oldloga = 0
        self.oldlogb = 0
        self.oldend = 0
        self.loga = 0
        self.logb = 0
        self.end = 0

    def getnew(self):
        workbook = xlrd.open_workbook(file_location)
        sheet = workbook.sheet_by_index(0)
        self.bot = sheet.cell_value(self.locrow, self.locbotcol)
        if (sheet.cell_value(self.locrow, self.loclogacol) == 42) and (sheet.cell_value(self.locrow, self.locendcol) == 42):
            print(self.name + " No link, PLC most likely off.")
        else:
            self.loga = sheet.cell_value(self.locrow, self.loclogacol)
            if sheet.cell_value(self.locrow, self.loclogbcol) == "":
                self.logb = 0
            else:
                self.logb = sheet.cell_value(self.locrow, self.loclogbcol)
                self.therm = True
            self.end = sheet.cell_value(self.locrow, self.locendcol)

    def sendem(self, startingtime, currenttime):
        self.maths()
        parameters = {
            'site': int(site),
            'linecode': self.linecode,  # 'newCodeTest',
            'start': startingtime,
            'end': currenttime,
            'actual': self.actual,
            'scrap': self.scrap,
            'operator_count': 1,
            'productcode': self.bot,
        }
        if self.scrap_machines:
            del parameters['scrap']
        try:
            if self.actual != 0 or self.scrap != 0:
                if enable_proxies:
                    self.response = requests.post(url, proxies=proxies, data=parameters)
                else:
                    self.response = requests.post(url, data=parameters)
                self.response = self.response.json().get("success")
                if self.response:
                    print(self.name + " send was good.")
                    self.oldloga = self.loga
                    self.oldlogb = self.logb
                    self.oldend = self.end
                return self.response
            else:
                numbers_ = (self.name + " No change in numbers")
                print(numbers_)
        except Exception as e:
            print(e)

    def maths(self):
        if firstPass:
            self.oldloga = self.loga
            self.oldlogb = self.logb
            self.oldend = self.end
        self.getnew()
        self.actual = self.end - self.oldend
        oldchangediff = (self.oldloga + self.oldlogb) - self.oldend
        changediff = (self.loga + self.logb) - self.end
        self.scrap = changediff - oldchangediff
        if self.therm:
            if (self.oldloga - self.loga) >= 1 and (self.oldlogb - self.logb) == 0:
                self.bot = "Side A"
            elif (self.oldlogb - self.logb) >= 1 and (self.oldloga - self.loga) == 0:
                self.bot = "Side B"
            else:
                self.bot = "Side A & B"


class Machine(Line):  # TODO write number force to zero function
    def __init__(self, name, locrow):
        Line.__init__(self, name, locrow)
        self.trimerinloc = 5
        self.flamerinloc = 6
        self.leakinloc = 7
        self.visioninloc = 8
        self.trimerin = 0
        self.flamerin = 0
        self.leakin = 0
        self.visionin = 0
        self.trimerinold = 0
        self.flamerinold = 0
        self.leakinold = 0
        self.visioninold = 0
        self.wheelscrap = 0
        self.trimerscrap = 0
        self.flamerscrap = 0
        self.leakscrap = 0
        self.visionscrap = 0
        self.sendwheel = 0
        self.sendtrimer = 0
        self.sendflamer = 0
        self.sendleak = 0
        self.sendvision = 0
        self.wheelscrapold = 0
        self.trimerscrapold = 0
        self.flamerscrapold = 0
        self.leakscrapold = 0
        self.visionscrapold = 0
        self.machine = {}
        self.machinelist = []
        for m in range(12, 17):
            self.machinelist.append(sheet.cell_value(self.locrow, m))
        self.mget()

    def mzero(self):
        self.zero()
        self.trimerin = 0
        self.flamerin = 0
        self.leakin = 0
        self.visionin = 0
        self.trimerinold = 0
        self.flamerinold = 0
        self.leakinold = 0
        self.visioninold = 0
        self.trimerin = 0
        self.flamerin = 0
        self.leakin = 0
        self.visionin = 0
        self.wheelscrap = 0
        self.wheelscrapold = 0
        self.trimerscrap = 0
        self.trimerscrapold = 0
        self.flamerscrap = 0
        self.flamerscrapold = 0
        self.leakscrap = 0
        self.leakscrapold = 0
        self.visionscrap = 0
        self.visionscrapold = 0

    def mget(self):
        workbook = xlrd.open_workbook(file_location)
        sheet = workbook.sheet_by_index(0)
        self.trimerin = sheet.cell_value(self.locrow, self.trimerinloc)
        self.flamerin = sheet.cell_value(self.locrow, self.flamerinloc)
        self.leakin = sheet.cell_value(self.locrow, self.leakinloc)
        self.visionin = sheet.cell_value(self.locrow, self.visioninloc)
        self.wheelscrap = (self.loga + self.logb) - self.trimerin
        self.trimerscrap = self.trimerin - self.flamerin
        self.flamerscrap = self.flamerin - self.leakin
        self.leakscrap = self.leakin - self.visionin
        self.visionscrap = self.visionin - self.end

    def mmaths(self):  # This doesnt send correct scrap counts
        self.mget()
        if firstPass:
            #self.trimerinold = self.trimerin
            #self.flamerinold = self.flamerin
            #self.leakinold = self.leakin
            #self.visioninold = self.visionin
            self.wheelscrapold = self.wheelscrap
            self.trimerscrapold = self.trimerscrap
            self.flamerscrapold = self.flamerscrap
            self.leakscrapold = self.leakscrap
            self.visionscrapold = self.visionscrap

        self.sendwheel = self.wheelscrap - self.wheelscrapold
        self.sendtrimer = self.trimerscrap - self.trimerscrapold
        self.sendflamer = self.flamerscrap - self.flamerscrapold
        self.sendleak = self.leakscrap - self.leakscrapold
        self.sendvision = self.visionscrap - self.visionscrapold

        # TODO write auto-grab of scrap categories
        self.machine = {
            self.machinelist[0]: self.sendwheel,
            self.machinelist[1]: self.sendtrimer,
            self.machinelist[2]: self.sendflamer,
            self.machinelist[3]: self.sendleak,
            self.machinelist[4]: self.sendvision
        }

    def msend(self, start, end):
        self.sendem(start, end)
        self.mmaths()
        i = 0
        for machines in self.machinelist:
            parameters = {
                'site': int(site),
                'linecode': self.linecode,  # 'newCodeTest',
                'scrapcategory': machines,
                'scrap': self.machine.get(machines),
                'scrap_datetime': end,
            }
            try:
                if enable_proxies:
                    self.response = requests.post(Scrap_Url, proxies=proxies, data=parameters)
                else:
                    self.response = requests.post(Scrap_Url, data=parameters)
                self.response = self.response.json().get("success")
                if i == 0 and self.response:
                    print(self.name + ' Wheel good')
                    self.wheelscrapold = self.wheelscrap
                if i == 1 and self.response:
                    print(self.name + ' Trim good')
                    self.trimerscrapold = self.trimerscrap
                if i == 2 and self.response:
                    print(self.name + ' Flamer good')
                    self.flamerscrapold = self.flamerscrap
                if machines == 'Leak' and self.response:
                    print(self.name + ' Leak good')
                    self.leakscrapold = self.leakscrap
                if machines == 'Vision' and self.response:
                    print(self.name + ' Vision good')
                    self.visionscrapold = self.visionscrap
                i += 1
            except Exception as e:
                print(e)


starttime = datetime.now().strftime(DATETIME_SECONDS_STRING_FORMAT)
lineList = []
linedict = {}


for i in range(sheet.nrows):  # init lines and info in excel sheet
    if "Test" in sheet.row(i)[0].value:  # for real run change Test to BLWM
        lineList.append(sheet.cell_value(i, 1))
        if sheet.cell_value(i, 11):
            linedict[sheet.cell_value(i, 1)] = Machine(sheet.cell_value(i, 0), i)
        else:
            linedict[sheet.cell_value(i, 1)] = Line(sheet.cell_value(i, 0), i)


x = 1
while x == 1:
    while True:
        l2 = [
            linedict['L2'].trimerin,
            linedict['L2'].trimerinold,
            linedict['L2'].trimerscrap,
            linedict['L2'].trimerscrapold,
            linedict['L2'].flamerin,
            linedict['L2'].flamerinold,
            linedict['L2'].flamerscrap,
            linedict['L2'].flamerscrapold,
        ]
    #try:
        secs = datetime.now().strftime("%S")
        mins = datetime.now().strftime("%M")
        hours = datetime.now().strftime("%H")
        if int(hours) == 18 or int(hours) == 6:
            if int(mins) == 30:
                for a in lineList:  # runs through list of lines for repeated action instead of writing them out
                    if "Machine" in str(linedict[a].__class__):
                        linedict[a].mzero()
                    else:
                        linedict[a].zero()
        if 30 > int(mins) >= 27:
            minutes_to_30 = 29 - int(mins)
            minutes_to_seconds = minutes_to_30 * 60
            secs_to_go = 60 - int(secs)
            seconds_to_30 = secs_to_go + minutes_to_seconds
            pause_time = seconds_to_30
            print(seconds_to_30)
        else:
            pause_time = 180
        print(starttime)
        time.sleep(pause_time)
        endtime = datetime.now().strftime(DATETIME_SECONDS_STRING_FORMAT)
        '''
        for line in lineList:
            #linedict[line].sendem(starttime, endtime)
            print(linedict[line].name)
            if linedict[line].scrap_machines:
                linedict[line].msend(starttime, endtime)
            else:
                linedict[line].sendem(starttime, endtime)
        '''
        #linedict['L2'].sendem(starttime, endtime)
        linedict['L2'].msend(starttime, endtime)
        starttime = endtime
        firstPass = False
    #except Exception as e:
        #print(e)
