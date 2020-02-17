# coding=UTF-8
import serial
import os, sys
import re
import collections
import xlwt

MODE = 0

if(len(sys.argv)> 1):
    if(sys.argv[1] == "gen" or sys.argv[1] == "G" or sys.argv[1] == "g"):
        MODE = 1
        print("##### Generate Excel #####")
else:
    print("##### Input Mode #####")

COM_NUM = 'COM106'
BAUDRATE = 115200
DIRS = 'uart_data'

def Help():
    print("#####################")
    print("Usage:")
    print("1.Input chip info: read_com.py ")
    print("2.Generate Excel: read_com.py g/G/gen")
    print("#####################")
def CheckDir():
    # set directory
    if not os.path.exists(DIRS):
        os.makedirs(DIRS)

def ConfigFileName():
    # set filename
    return (DIRS+'\\chip')

def GetFileList(dir, ftype,fileList):
    if os.path.isfile(dir):
        if ftype in dir:
            fileList.append(dir)
    #recurrence find
    elif os.path.isdir(dir):
        for s in os.listdir(dir):
            newDir=os.path.join(dir,s)
            GetFileList(newDir, ftype, fileList)
    return fileList

def GetInfo(filelist):
    chipdict= collections.OrderedDict()
    for fl in filelist:
        id = fl[fl.find('chip_')+5:-4]
        dict = {}
        lvt = svt = waferid = xloc = yloc = 0
        with open(fl,"r") as f:
            for line in f:
                lvt_patten = re.search(r'LVT ([\d]+) vs TYP ([\d]+).*', line, re.M | re.I)
                if(lvt_patten):
                    lvt = lvt_patten.group(1)
                svt_patten = re.search(r'SVT ([\d]+) vs TYP ([\d]+).*', line, re.M | re.I)
                if(svt_patten):
                    svt = svt_patten.group(1)
                waferid_patten = re.search(r'Wafer ID : ([\d]+).*', line, re.M | re.I)
                if(waferid_patten):
                    waferid = waferid_patten.group(1)
                xloc_patten = re.search(r'X location : ([\d]+).*', line, re.M | re.I)
                if(xloc_patten):
                    xloc = xloc_patten.group(1)
                yloc_patten = re.search(r'Y location : ([\d]+).*', line, re.M | re.I)
                if(yloc_patten):
                    yloc = yloc_patten.group(1)
        dict['LVT'] = lvt
        dict['SVT'] = svt
        dict['WAFERID'] = waferid
        dict['XLOC'] = xloc
        dict['YLOC'] = yloc
        chipdict[id] = dict
    return (chipdict)

def StyleSetting():
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = "SimSun"
    font.height = 20 * 11
    font.bold = False
    font.colour_index = 0x01
    style.font = font
    pat = xlwt.Pattern()
    pat.pattern = xlwt.Pattern.SOLID_PATTERN
    pat.pattern_fore_colour = xlwt.Style.colour_map['dark_green']
    style.pattern = pat
    borders = xlwt.Borders()
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    style.borders = borders
    return style

def WriteExcel(chipdict):
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('chip_dro')
    worksheet.col(0).width = 10 * 256
    worksheet.col(3).width = 14*256
    style = StyleSetting()
    worksheet.write(0, 0, 'chip_num', style)
    worksheet.write(0, 1, 'LVT', style)
    worksheet.write(0, 2, 'SVT', style)
    worksheet.write(0, 3, 'WAFER_ID', style)
    worksheet.write(0, 4, 'X_LOC', style)
    worksheet.write(0, 5, 'Y_LOC', style)
    row = 0
    for k,v in chipdict.items():
        row += 1
        col = 0
        worksheet.write(row, col, k)
        for k2, v2 in v.items():
            col += 1
            worksheet.write(row, col, v2)
    workbook.save(DIRS+'\\dro.xls')
    print("Excel Generate Done!")

def Process(filename):
    # set com
    selComPort = COM_NUM
    selBaudRate = BAUDRATE
    ser = serial.Serial(port=selComPort, baudrate=selBaudRate, bytesize=8, stopbits=1, timeout=5)

    # process
    while(1):
        chip_id = input("Please input Chip ID (X to exit) : ")
        if chip_id == "X" or chip_id == "x":
            break
        fn = open(filename + "_" + str(chip_id) + ".txt", 'w')
        fn.write('CHIP_NUM = ' + chip_id+ "\n")
        cmd1 = 'reboot 0\n'
        ser.write(cmd1.encode('utf-8'))

        # read uart
        timecnt = 0
        while True:
            count = ser.inWaiting()
            timecnt += 1
            if (count == 0 and timecnt > 100000):
                break
            if count > 0:
                timecnt = 0
                data = ser.readline()
                if data != b'':
                    print(str(data)[2:])
                    fn.write(str(data)[2:] + "\n")
                else:
                    break
        fn.flush()
        fn.close()

        print("Write Chip Info Done!")
    ser.close()
    print("Exit!")

if __name__ == '__main__':
    Help()
    CheckDir()
    if(MODE == 0):
        filename = ConfigFileName()
        Process(filename)
    else:
        fileList = GetFileList(DIRS,".txt",[])
        datadict = GetInfo(fileList)
        WriteExcel(datadict)
