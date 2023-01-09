#!/usr/local/bin/python3

########## Imports ##########

from barcode import Code128
from barcode.writer import ImageWriter

import os, subprocess
import os.path
import glob
import time

import labels
from reportlab.graphics import shapes
from reportlab.platypus import Image
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFont
from reportlab.lib import colors

from PIL import Image

import re
import string



########## Variables ##########

# This is the first of the two-year range (ex 2017-2018 would be '2017')
currentYear = 2022

base_path = os.path.dirname(__file__)
desktop_path = os.path.expanduser("~/Desktop")

registerFont(TTFont('Ostrich Sans Heavy', os.path.join(base_path, 'resources/OstrichSans-Heavy.ttf')))
registerFont(TTFont('SerifRegular', os.path.join(base_path, 'resources/SourceSerifPro-Regular.ttf')))
registerFont(TTFont('SerifBold', os.path.join(base_path, 'resources/SourceSerifPro-Bold.ttf')))

isBorders = False

skipStart = [1, 1]


sheetSpecWidth = 215.9
sheetSpecHeight = 279.4
sheetSpecLeftMargin = 4
sheetSpecTopMargin = 13.5

sheetSpecColumns = 3
sheetSpecRows = 10

sheetSpecColumnGap = 3.556
sheetSpecRowGap = 0

sheetSpecLabelWidth = 66.675
sheetSpecLabelHeight = 25.4
sheetSpecLabelCornerRadius = 2


specs = labels.Specification(
    sheetSpecWidth,
    sheetSpecHeight,
    sheetSpecColumns,
    sheetSpecRows,
    sheetSpecLabelWidth,
    sheetSpecLabelHeight,
    top_margin=sheetSpecTopMargin,
    left_margin=sheetSpecLeftMargin,
    column_gap=sheetSpecColumnGap,
    row_gap=sheetSpecRowGap,
    corner_radius=sheetSpecLabelCornerRadius
    )


########## Functions ##########

class studentUser:

    def __init__(self, input, currentYear):

        validity = True
        userName = "USER NOT FOUND"
        fullName = "Name Not Found"
        firstName = "FNameError"
        lastName = "LNameError"
        gradYear = "0000"
        school = "None"

        try:
            userName = re.findall("[a-z]+[0-9]*", input)[0]
        except:
            validity = False

        if validity:

            p = subprocess.Popen(
                f"dscl /Active\ Directory/CHATHAM-AD/All\ Domains/ read /Users/{userName}", 
                stdout= subprocess.PIPE,
                shell=True
                )
            (data, err) = p.communicate()
            p_status = p.wait()
            data = data.decode('utf-8')

            # Full Name
            try:
                fullName = re.findall("RealName:[\n]?[ ]?[A-Za-z0-9'\-\s]+\n", data)[0]
                fullName = re.sub("^RealName:[\n]?[ ]?", "", fullName)
                fullName = re.sub("[\n]$", "", fullName)
            except:
                validity = False

            # First Name
            try:
                firstName = re.findall("FirstName:[\n]?[ ]?[A-Za-z0-9'\-\s]+\n", data)[0]
                firstName = re.sub("^FirstName:[\n]?[ ]?", "", firstName)
                firstName = re.sub("[\n]$", "", firstName)
            except:
                validity = False

            # Last Name
            try:
                lastName = re.findall("LastName:[\n]?[ ]?[A-Za-z0-9'\-\s]+\n", data)[0]
                lastName = re.sub("^LastName:[\n]?[ ]?", "", lastName)
                lastName = re.sub("[\n]$", "", lastName)
            except:
                validity = False

            # Grade Level
            try:
                gradYear = re.findall("FAXNumber:[\n]?[ ]?[0-9]+\n", data)[0]
                gradYear = re.sub("^FAXNumber:[\n]?[ ]?", "", gradYear)
                gradYear = re.sub("[\n]$", "", gradYear)
                try:
                    gradeCalc = (12 - int(gradYear)) + currentYear + 1
                    gradYear = str(gradeCalc).zfill(4)
                except:
                    gradYear = "0000"
            except:
                validity = False

            # School Code
            try:
                school = re.findall("dsAttrTypeNative:extensionAttribute8:[\n]?[ ]?[A-Za-z0-9]+\n", data)[0]
                school = re.sub("^dsAttrTypeNative:extensionAttribute8:[\n]?[ ]?", "", school)
                school = re.sub("[\n]$", "", school)
            except:
                validity = False


        self.userName = userName

        self.fullName = fullName
        self.firstName = firstName
        self.lastName = lastName
        self.gradYear = gradYear
        self.school = school
        self.valid = validity

        #if not validity:
        #    print(f"WARNING: Invalid Tag -| {input} |-")


def draw_label(label, width, height, obj):

    userNamePath = os.path.join(base_path, f"temp/{obj.userName}.png")
    with open(userNamePath, "wb") as f:
        Code128(obj.userName, writer=ImageWriter()).write(f, options={"write_text": False})
    im = Image.open(userNamePath)

    label.add(shapes.Image(23, 10, width=150, height=25, path=userNamePath))

    userName = ' '.join( (obj.userName)[i] for i in range(0, len(obj.userName)) )
    userNameText = shapes.String(96, 12, userName, fontName="Ostrich Sans Heavy", fontSize=8, textAnchor='middle')
    userNameTextBounds = userNameText.getBounds()

    backRect = shapes.Rect(
        userNameTextBounds[0],
        userNameTextBounds[1] - 2,
        userNameTextBounds[2] - userNameTextBounds[0],
        userNameTextBounds[3] - userNameTextBounds[1] - 1,
    )
    backRect.fillColor = colors.white
    backRect.strokeColor = colors.white
    backRect.strokeWidth = 6
    backRect.strokeLineJoin = 1
    label.add(backRect)

    label.add(userNameText)


    lastName = obj.lastName
    fontSize = 20
    escape = False
    while(not escape):
        nameText = shapes.String(10, 38, lastName, fontName="SerifBold", fontSize=fontSize, textAnchor='start')
        nameTextBounds = nameText.getBounds()
        if(nameTextBounds[2] - nameTextBounds[0] > 170):
            fontSize = fontSize - 1
        else:
            escape = True
    label.add(nameText)

    firstName = obj.firstName
    fontSize = 10
    escape = False
    while(not escape):
        nameText = shapes.String(11, 55, firstName, fontName="SerifRegular", fontSize=fontSize, textAnchor='start')
        nameTextBounds = nameText.getBounds()
        if(nameTextBounds[2] - nameTextBounds[0] > 100):
            fontSize = fontSize - 1
        else:
            escape = True
    label.add(nameText)


    label.add(shapes.String(180, 55, obj.school + " " + obj.gradYear, fontName="SerifRegular", fontSize=10, textAnchor='end'))

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


########## BEGIN SCRIPT ##########

print("")

studentList = input(f"{bcolors.HEADER}Enter student Usernames separated by commas: {bcolors.ENDC}")
studentList = studentList.split(",")
studentList = [s.strip() for s in studentList]


skipStartInput = input(f"{bcolors.HEADER}Enter the first available label or press enter for A1: {bcolors.ENDC}")


sheet = labels.Sheet(specs, draw_label, border=isBorders)


skipStartInputSearch = re.search("^[a-zA-Z][0-9]+$", skipStartInput)
if (not skipStartInputSearch == None):

    newString = ""
    for i in range(skipStartInputSearch.span(0)[0], skipStartInputSearch.span(0)[0] + skipStartInputSearch.span(0)[1]):
        newString = newString + skipStartInput[i].upper()

    skipStart[0] = string.ascii_uppercase.find(newString[0]) + 1
    skipStart[1] = int(newString[1:])
    if (skipStart[0] > sheetSpecColumns):
        skipStart[0] = sheetSpecColumns
        print(f"{bcolors.WARNING}WARNING: Label column out of range, reducing to {sheetSpecColumns}{bcolors.ENDC}")
    if (skipStart[1] > sheetSpecRows):
        skipStart[1] = sheetSpecRows
        print(f"{bcolors.WARNING}WARNING: Label row out of range, reducing to {sheetSpecRows}{bcolors.ENDC}")


    print(f"{bcolors.OKCYAN}Beginning labels at {skipStart}{bcolors.ENDC}")

else:
    print(f"{bcolors.WARNING}Defaulting label start to A1{bcolors.ENDC}")

print("")

usedLabels = [[]]
isFound = False
for i in range(1, sheetSpecRows+1):
    for j in range(1, sheetSpecColumns+1):
        if (not isFound):
            if (j == skipStart[0] and i == skipStart[1]):
                isFound = True
            else:
                if (isFound == False):
                    usedLabels.append([i, j])
usedLabels.pop(0)
sheet.partial_page(1, usedLabels)

generatedSomething = False
generatedError = False
listOfErrors = []
for i in range(len(studentList)):
    studentStore = studentUser(studentList[i], currentYear)
    if studentStore.valid:
        sheet.add_label(studentStore)
        print(f">  Generated [{studentStore.userName}] - {studentStore.fullName}")
        generatedSomething = True
    else:
        print(f"{bcolors.FAIL}WARNING: Failed to generate tag [{studentStore.userName}]{bcolors.ENDC}")

print("")

if(generatedSomething):
    sheet.save(os.path.join(desktop_path, "Labels.pdf"))
    print(f"{bcolors.OKGREEN}Document Saved as Labels.pdf on Desktop{bcolors.ENDC}")
    print(f"{bcolors.OKCYAN}{sheet.label_count} label{'s' if sheet.label_count > 1 else ''} output on {sheet.page_count} page{'s' if sheet.page_count > 1 else ''}.{bcolors.ENDC}")
else:
    print(f"{bcolors.FAIL}ERROR: No Labels Generated{bcolors.ENDC}")

try:
    filelist = glob.glob(os.path.join(base_path, "temp", "*"))
    for f in filelist:
        os.remove(f)
except:
    print(f"{bcolors.WARNING}WARNING: Unable to complete cleanup{bcolors.ENDC}")

time.sleep(5)
print("")
