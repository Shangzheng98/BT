# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import datetime
import xml.etree.ElementTree as ET

import numpy as np
import pandas as pd
import requests
import xlrd

MonthAndDay = ['09-21', '09-22', '09-23', '10-01', '10-09']  # '09-21', '09-22', '09-23', '10-01', '10-09'
YEAR = 2020

Routes = []
Buses = []


class Route:

    def __init__(self, name):
        self.name = name
        self.buses = []


class BusStop:
    def __init__(self, name, code):
        self.stopName = name
        self.stopCode = code


class Bus:
    def __init__(self, Id, length):
        self.busId = Id
        self.busLength = length


def fileNameFormater(routeName, date, year):
    return "/home/jsz/PycharmProject/Input/data/" + routeName + "-" + date + "-" + str(
        year) + "-filtered.xlsx"


def createLoopTimeSheet(file):
    d = {}
    for r in Routes:
        d.update({r.name: []})
    for date in MonthAndDay:

        for i in range(len(Routes)):
            if Routes[i].name == "HWA" or Routes[i].name == "MSA":
                continue
            if Routes[i].name == "CAS" and date == "09-23":
                continue
            if Routes[i].name == "CBD" and date == '09-23':
                continue
            if Routes[i].name == "CBD" and date == "10-09":
                continue
            if Routes[i].name == "PHD" and date == "09-21":
                continue
            if Routes[i].name == "PHD" and date == "10-01":
                continue
            if Routes[i].name == "PRO" and date == "09-21":
                continue
            if Routes[i].name == "PRO" and date == "09-22":
                continue
            if Routes[i].name == "PRO" and date == "09-23":
                continue
            if Routes[i].name == "PRO" and date == "10-01":
                continue
            if Routes[i].name == "TOM" and date == "09-21":
                continue
            if Routes[i].name == "TOM" and date == "09-22":
                continue
            if Routes[i].name == "TOM" and date == "10-01":
                continue
            print(date + "loop time ", Routes[i].name)
            workBook = xlrd.open_workbook(fileNameFormater(Routes[i].name, date, YEAR))
            arr = []
            for k in range(workBook.nsheets):
                data = workBook.sheet_by_index(k)
                currentStartStopName = ''
                startTime = None
                endTime = None
                loopTime = []
                for j in range(data.nrows):
                    if j == 0:
                        continue
                    row = data.row(j)
                    if row[2].value == "Blacksburg Transit" and Routes[i].name == "CRC": continue
                    if row[2].value == "Tall Oaks/Colonial Sbnd" and Routes[i].name == "HWA": continue
                    if row[2].value == "Hethwood Square on Hethwood" and Routes[i].name == "HWB": continue
                    if (row[2].value == "Professional Park Nbnd" or row[2].value == "Fairfax/Ellett Ebnd") and Routes[
                        i].name == "MSS": continue
                    if row[2].value == "LewisGale Hospital Montgomery" and Routes[i].name == "TTT": continue

                    if row[4].value == "Y" and currentStartStopName == '':
                        currentStartStopName = row[2].value
                        startTime = datetime.datetime(*xlrd.xldate_as_tuple(row[5].value, workBook.datemode))
                    elif row[4].value == "Y" and currentStartStopName != '' and row[2].value == currentStartStopName:
                        endTime = datetime.datetime(*xlrd.xldate_as_tuple(row[5].value, workBook.datemode))

                    if startTime is not None and endTime is not None:
                        oneLoopTime = endTime - startTime
                        oneLoopTimeMin = divmod(oneLoopTime.total_seconds(), 60)[0]
                        if oneLoopTimeMin < 10:
                            continue
                        d[Routes[i].name].append(oneLoopTimeMin)
                        startTime = None
                        endTime = None
                        currentStartStopName = ""
    # print(d)
    for key in d:
        dataset = d[key]
        if len(dataset) == 0:
            d[key] = ["invalid data"]
        else:
            ave = np.ceil(np.average(dataset))
            d[key] = [ave]
    df = pd.DataFrame.from_dict(d, orient='index')
    df = df.transpose()
    df.to_excel(file, sheet_name='Loop Time')


def createBusNumberSheet(file):
    d = {}
    for r in Routes:
        d.update({r.name: []})
    for date in MonthAndDay:
        for i in range(len(Routes)):
            if Routes[i].name == "HWA" or Routes[i].name == "MSA":
                continue
            if Routes[i].name == "CAS" and date == "09-23":
                continue
            if Routes[i].name == "CBD" and date == '09-23':
                continue
            if Routes[i].name == "CBD" and date == "10-09":
                continue
            if Routes[i].name == "PHD" and date == "09-21":
                continue
            if Routes[i].name == "PHD" and date == "10-01":
                continue
            if Routes[i].name == "PRO" and date == "09-21":
                continue
            if Routes[i].name == "PRO" and date == "09-22":
                continue
            if Routes[i].name == "PRO" and date == "09-23":
                continue
            if Routes[i].name == "PRO" and date == "10-01":
                continue
            if Routes[i].name == "TOM" and date == "09-21":
                continue
            if Routes[i].name == "TOM" and date == "09-22":
                continue
            if Routes[i].name == "TOM" and date == "10-01":
                continue
            workBook = xlrd.open_workbook(fileNameFormater(Routes[i].name, date, YEAR))
            arr = []
            for k in range(workBook.nsheets):
                data = workBook.sheet_by_index(k)
                # print(Routes[i].name, data.name)
                # bus = searchBusID(data.name)
                # type = -1
                # if bus.busLength == 35:
                #     type = 1
                # elif bus.busLength == 40:
                #     type = 2
                # elif bus.busLength == 60:
                #     type = 3
                # else:
                #     type = -1
                d[Routes[i].name].append(data.name)
            # print(arr)
    for key in d:
        dateset = d[key]
        dateset = list(dict.fromkeys(dateset))
        one = "1 "
        two = "2 "
        three = "3 "
        for bid in dateset:
            bus = searchBusID(bid)
            if bus.busLength == 35:
                one += (bid + " ")
            elif bus.busLength == 40:
                two += (bid + " ")
            elif bus.busLength == 60:
                three += (bid + " ")
        arr = [one, two, three]
        temp = []
        for e in arr:
            if len(e) > 2:
                string = e[:2] + '(' + e[2:] + ')'
                temp.append(string)

        d[key] = temp
    df = pd.DataFrame.from_dict(d, orient='index')
    df = df.transpose()
    df.to_excel(file, sheet_name='Bus Number')


def searchBusID(busId):
    for bus in Buses:
        if bus.busId == int(busId):
            return bus

    return None


class Stop:

    def __init__(self, stopCode, stopName):
        self.stopCode = stopCode
        self.stopName = stopName

    def toString(self):
        return str(self.stopCode) + ',' + self.stopName


def createBusDic():
    busFile = xlrd.open_workbook("/home/jsz/PycharmProject/Input/data/Simple Fleet List.xlsx")
    data = busFile.sheet_by_index(0)
    for i in range(data.nrows):

        row = data.row(i)
        if i == 0 or i == 1:
            continue

        busId = int(row[0].value)

        slice_object = slice(2)
        busLength = int(row[3].value[slice_object])

        newBus = Bus(busId, busLength)
        Buses.append(newBus)


def getStopInfo(name):
    stopList = {}

    StopNameResponse = requests.get(
        "http://www.bt4uclassic.org/webservices/bt4u_webservice.asmx/GetScheduledStopNames",
        "routeShortName=" + name)
    root = ET.fromstring(StopNameResponse.content)
    for ele in root.iter("ScheduledStops"):
        stopName = ele.find("StopName").text
        stopCode = int(ele.find("StopCode").text)
        stop = Stop(stopCode, stopName)
        arr = []
        stopList.update({stopCode: arr})

    return stopList


def createPassengerOnboardSheet(startTime, endTime, index, date):
    filewriter = pd.ExcelWriter(
        'PassengerOnboard' + " " + date + "-" + str(index) + "-" + str(index + 1) + '.xlsx',
        engine='xlsxwriter')
    for r in Routes:
        if r.name == "HWA" or r.name == "MSA":
            continue
        if r.name == "CAS" and date == "09-23":
            continue
        if r.name == "CBD" and date == '09-23':
            continue
        if r.name == "CBD" and date == "10-09":
            continue
        if r.name == "PHD" and date == "09-21":
            continue
        if r.name == "PHD" and date == "10-01":
            continue
        if r.name == "PRO" and date == "09-21":
            continue
        if r.name == "PRO" and date == "09-22":
            continue
        if r.name == "PRO" and date == "09-23":
            continue
        if r.name == "PRO" and date == "10-01":
            continue
        if r.name == "TOM" and date == "09-21":
            continue
        if r.name == "TOM" and date == "09-22":
            continue
        if r.name == "TOM" and date == "10-01":
            continue
        print(r.name)
        d = {}
        workBook = xlrd.open_workbook(fileNameFormater(r.name, date, YEAR))

        stoplist = getStopInfo(r.name)
        for k in range(workBook.nsheets):
            # print(stoplist)
            data = workBook.sheet_by_index(k)
            # passengerSet = []

            for j in range(data.nrows):
                if j == 0:
                    continue
                rowData = data.row(j)
                strRecordTime = datetime.datetime(*xlrd.xldate_as_tuple(rowData[5].value, workBook.datemode)).strftime(
                    '%H:%M:%S')
                recordTime = datetime.datetime.strptime(strRecordTime, '%H:%M:%S').time()
                if endTime > recordTime > startTime:
                    arr = stoplist[int(rowData[3].value)]
                    arr.append(int(rowData[6].value))

        sumArray = []
        for key in stoplist:
            if len(stoplist[key]) != 0:
                ave = np.ceil(np.average(stoplist[key]))
            else:
                ave = 0
            sumArray.append(ave)

        d.update({"ave": sumArray})
        # print(d)

        stopList = {}

        StopNameResponse = requests.get(
            "http://www.bt4uclassic.org/webservices/bt4u_webservice.asmx/GetScheduledStopNames",
            "routeShortName=" + r.name)
        root = ET.fromstring(StopNameResponse.content)
        for ele in root.iter("ScheduledStops"):
            stopName = ele.find("StopName").text
            stopCode = ele.find("StopCode").text
            stop = Stop(stopCode, stopName)
            arr = []
            stopList.update({stopCode + " " + stopName: arr})

        d.update({"stopCode": stopList})
        df = pd.DataFrame.from_dict(d, orient='index')
        df = df.transpose()
        df.to_excel(filewriter, sheet_name=r.name)
    filewriter.save()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    routesName = ['CAS', 'CRC', "CBD", "HXP", "HWA", "HWB",
                  "HDG", "MSS", "TOM", "UMS", "MSN", "PHD", "PRB", "PRO", "MSA", "TTT", "UCB"]

    for e in routesName:
        route = Route(e)
        Routes.append(route)

    createBusDic()
    writer = pd.ExcelWriter('scenario.xlsx', engine='xlsxwriter')

    createLoopTimeSheet(writer)
    createBusNumberSheet(writer)
    writer.save()

    count = 0
    for date in MonthAndDay:
        for i in range(7, 21):
            count += 1
            print(count)
            timeStart = datetime.time(i, 0)
            timeEnd = datetime.time(i + 1, 0, 0)
            createPassengerOnboardSheet(timeStart, timeEnd, i, date)

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
