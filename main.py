import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import xlwt
from xlwt import Workbook
import random
from colorConstants import colors as clr


def chooseADocument():
    isRFP = True
    data = None
    experimentNumber = "First Experiment"
    chosenDoc = input("Select a document with the relevant data to analyze\n"
                      "1. Experiment 1 OD\n"
                      "2. Experimenet 2 OD\n"
                      "3. Experiment 2 RFP\n")
    match chosenDoc:
        case '1':
            isRFP = False
            data = pd.read_excel('OD_Data.xlsx', index_col=0)
            experimentNumber = "First Experiment"
        case '2':
            isRFP = False
            data = pd.read_excel('YV_Exp2.xlsx', index_col=0)
            experimentNumber = "Second Experiment"
        case '3':
            isRFP = True
            data = pd.read_excel('YV_Exp2_RFP.xlsx', index_col=0)
            experimentNumber = "Second Experiment"
        case _:
            print("Invalid input")
    dataList = (data, isRFP, experimentNumber)
    return dataList


dataFrame, isRFP, experimentNumber = chooseADocument()


class Variable:
    instances = []

    def __init__(self, variableName, locationInDataframe, isInduced, variableData, variableAvg, stdErr):
        self.variableData = variableData
        self.locationInDataFrame = locationInDataframe
        self.variableName = variableName
        self.isInduced = isInduced
        self.instances.append(self)
        self.variableAvg = variableAvg
        self.stdErr = stdErr

    def getVariableData(self):
        pool = 0
        poolLen = 0
        for location in self.locationInDataFrame:
            self.variableData.append(dataFrame.loc[location])
        for item in self.variableData:
            pool += item.sum()
            poolLen += len(item)
        average = pool/poolLen
        standardDev = np.std(self.variableData)
        standardError = standardDev/(np.sqrt(len(self.variableData)))
        self.variableData.clear()
        self.variableData.append(standardError)
        self.variableData.append(average)
        self.stdErr.append(standardDev)
        print(f'The average of {self.variableName} is {self.variableData[0]}\nThe standard deviation is {self.variableData[1]}\n'
              f'The standard error is {standardError}')
        return self.variableData

    def getAverages(self):
        current = 0
        x = 0
        while x < dataFrame.shape[1]:
            for items in self.locationInDataFrame:
                current += dataFrame.loc[items].iloc[x]
            average = current/4
            self.variableAvg.append(average)
            current = 0
            x += 1
        return self.variableAvg


nonInducedMedia = Variable('NonInduced Media', ["A1", "B1", "C1", "D1"], False, [], [], [])
nonInducedNegativeControl = Variable('NonInduced Negative Control', ["A2", "B2", "C2", "D2"], False, [], [], [])
nonInducedPositiveControl = Variable('NonInduced Positive Control', ["A3", "B3", "C3", "D3"], False, [], [], [])
nonInducedLonOne = Variable('NonInduced Lon RFP', ["A4", "B4", "C4", "D4"], False, [], [], [])
nonInducedLonTwo = Variable('NonInduced Lon Cam', ["A5", "B5", "C5", "D5"], False, [], [], [])
nonInducedMazOne = Variable('NonInduced Maz RFP', ["A6", "B6", "C6", "D6"], False, [], [], [])
nonInducedMazTwo = Variable('NonInduced Maz Cam', ["A7", "B7", "C7", "D7"], False, [], [], [])
nonInducedBlank = Variable('nonInduced Blank', ["A8", "B8", "C8", "D8"], False, [], [], [])

InducedMedia = Variable('Media Induced', ["E1", "F1", "G1", "H1"], True, [], [], [])
InducedNegativeControl = Variable('Negative Control Induced', ["E2", "F2", "G2", "H2"], True, [], [], [])
InducedPositiveControl = Variable('Positive Control Induced', ["E3", "F3", "G3", "H3"], True, [], [], [])
InducedLonOne = Variable('Lon RFP Induced', ["E4", "F4", "G4", "H4"], True, [], [], [])
InducedLonTwo = Variable('Lon Cam Induced', ["E5", "F5", "G5", "H5"], True, [], [], [])
InducedMazOne = Variable('Maz RFP Induced', ["E6", "F6", "G6", "H6"], True, [], [], [])
InducedMazTwo = Variable('Maz Cam Induced', ["E7", "F7", "G7", "H7"], True, [], [], [])
InducedBlank = Variable('Blank', ["E8", "F8", "G8", "H8"], True, [], [], [])


def getAllData():
    for item in Variable.instances:
        item.getVariableData()


def writeToNewSheet():
    averagesWorkBook = Workbook()
    sheetOne = averagesWorkBook.add_sheet('Sheet 2', cell_overwrite_ok=True)
    style = xlwt.easyxf('font: bold 1')
    sheetOne.write(0, 0, 'Time', style)
    width = 100
    startStatement = "AVERAGES"
    centeredStatement = startStatement.center(width)
    print(centeredStatement)
    for item in Variable.instances:
        row = 1
        column = Variable.instances.index(item) + 1
        itemLoc = 0
        # print(item.getAverages())
        item.getAverages()
        while itemLoc < len(item.variableAvg):
            sheetOne.write(row, column, f'{item.variableAvg[itemLoc]}')
            row += 1
            itemLoc += 1
        column += 1
    endStatement = "END AVERAGES"
    centeredEndStatement = endStatement.center(width)
    print(centeredEndStatement)
    x = 1
    while x < len(Variable.instances) + 1:
        for item in Variable.instances:
            sheetOne.write(0, x, f'{item.variableName}', style)
            x += 1
    starCol = 0
    endCol = 12
    while starCol < endCol:
        if starCol == 0:
            sheetOne.col(starCol).width = 1500
            starCol += 1
        elif starCol == 1:
            sheetOne.col(starCol).width = 3500
            starCol += 1
        elif starCol == 2 or 3 or 7 or 8 or 9:
            sheetOne.col(starCol).width = 6000
            starCol += 1
        elif starCol == 4 or 5 or 6:
            sheetOne.col(starCol).width = 2000
            starCol += 1
    # Write time values into column 0 labeled ' Time '
    timeFrame_values_in_minutes = []
    timeframe = dataFrame.iloc[0]
    for value in timeframe:
        timeFrame_values_in_minutes.append(round((value / 60)))
        startVal = 1
    valNum = 0
    EndVal = len(timeFrame_values_in_minutes) + 1
    while startVal < EndVal:
        sheetOne.write(startVal, 0, timeFrame_values_in_minutes[valNum])
        startVal += 1
        valNum += 1

    averagesWorkBook.save('Averages Workbook.xlsx')


def rgb_to_hex(color):
    return '%02x%02x%02x' % color


def generateRandomColor():
    colorOne = "#" + (rgb_to_hex(random.choice(list(clr.values()))))
    return colorOne


def graphControlAverages():
    fig, ax = plt.subplots(figsize=(10, 10), facecolor='white')
    ax.set_facecolor('white')
    ax.tick_params(labelcolor='black')
    if isRFP:
        ylimit = 1200
        yax = "RFP"
    else:
        ylimit = 1
        yax = "OD"
    ax.set_title(f'Control {yax} Averages {experimentNumber}', color='black')
    ax.set_xlabel('Time (minutes)', color='black', fontsize=20)
    ax.set_ylabel(yax, color='black', fontsize=20)
    ax.set_xlim(0, 1120)
    ax.set_ylim(0, ylimit)
    dataFrameTwo = pd.read_excel('Averages Workbook.xlsx', index_col=0)
    for item in Variable.instances:
        if item.variableName == "NonInduced Positive Control":
            plt.plot(dataFrameTwo[f'{item.variableName}'], label=f'{item.variableName}',
                     color='blue', alpha=1,
                     linestyle='dashed')
        elif item.variableName == "Positive Control Induced":
            plt.plot(dataFrameTwo[f'{item.variableName}'], label=f'{item.variableName}',
                     color='red', alpha=1)
        elif item.variableName == "Negative Control Induced":
            plt.plot(dataFrameTwo[f'{item.variableName}'], label=f'{item.variableName}',
                     color='black', alpha=1)
        elif item.variableName == "NonInduced Negative Control":
            plt.plot(dataFrameTwo[f'{item.variableName}'], label=f'{item.variableName}',
                     color='purple', alpha=1,
                     linestyle='dashed')
        # dataFrameTwo.plot(y=f'{item.variableName}', use_index=True)
        # print(item.variableName)
    ax.legend(facecolor='white', framealpha=1, fontsize='large', shadow=True, labelspacing=2)
    ax.patch.set_edgecolor('black')
    ax.patch.set_linewidth('1')
    plt.show()


def graphLonAverages():
    fig, ax = plt.subplots(figsize=(10, 10), facecolor='white')
    ax.set_facecolor('white')
    ax.tick_params(labelcolor='black')
    if isRFP:
        ylimit = 1200
        yax = "RFP"
    else:
        ylimit = 1
        yax = "OD"
    ax.set_title(f'Lon {yax} Averages {experimentNumber}', color='black')
    ax.set_xlabel('Time (minutes)', color='black', fontsize=20)
    ax.set_ylabel(yax, color='black', fontsize=20)
    ax.set_xlim(0, 1120)
    ax.set_ylim(0, ylimit)
    dataFrameTwo = pd.read_excel('Averages Workbook.xlsx', index_col=0)
    for item in Variable.instances:
        if item.variableName == "NonInduced Lon RFP":
            plt.plot(dataFrameTwo[f'{item.variableName}'], label=f'{item.variableName}',
                     color=generateRandomColor(), alpha=1,
                     linestyle='dashed')
        elif item.variableName == "NonInduced Lon Cam":
            plt.plot(dataFrameTwo[f'{item.variableName}'], label=f'{item.variableName}',
                     color=generateRandomColor(), alpha=1,
                     linestyle='dashdot')
        elif item.variableName == "Lon RFP Induced":
            plt.plot(dataFrameTwo[f'{item.variableName}'], label=f'{item.variableName}',
                     color=generateRandomColor(), alpha=1,
                     linestyle='solid')
        elif item.variableName == "Lon Cam Induced":
            plt.plot(dataFrameTwo[f'{item.variableName}'], label=f'{item.variableName}',
                     color=generateRandomColor(), alpha=1,
                     linestyle='dotted')
        # dataFrameTwo.plot(y=f'{item.variableName}', use_index=True)
        # print(item.variableName)
    ax.legend(facecolor='white', framealpha=1, fontsize='large', shadow=True, labelspacing=2)
    ax.patch.set_edgecolor('black')
    ax.patch.set_linewidth('1')
    plt.show()


def graphMazAverages():
    fig, ax = plt.subplots(figsize=(10, 10), facecolor='white')
    ax.set_facecolor('white')
    ax.tick_params(labelcolor='black')
    if isRFP:
        ylimit = 1200
        yax = "RFP"
    else:
        ylimit = 1
        yax = "OD"
    ax.set_title(f'Maz {yax} Averages {experimentNumber}', color='black')
    ax.set_xlabel('Time', color='black', fontsize=20)
    ax.set_ylabel(yax, color='black', fontsize=20)
    ax.set_xlim(0, 1120)
    ax.set_ylim(0, ylimit)
    dataFrameTwo = pd.read_excel('Averages Workbook.xlsx', index_col=0)
    for item in Variable.instances:
        match item.variableName:
            case "NonInduced Maz RFP":
                plt.plot(dataFrameTwo[f'{item.variableName}'], label=f'{item.variableName}',
                     color=generateRandomColor(), alpha=1,
                     linestyle='dashed')
            case "NonInduced Maz Cam":
                plt.plot(dataFrameTwo[f'{item.variableName}'], label=f'{item.variableName}',
                     color=generateRandomColor(), alpha=1,
                     linestyle='dashed')
            case "Maz RFP Induced":
                plt.plot(dataFrameTwo[f'{item.variableName}'], label=f'{item.variableName}',
                     color=generateRandomColor(), alpha=1)
            case "Maz Cam Induced":
                plt.plot(dataFrameTwo[f'{item.variableName}'], label=f'{item.variableName}',
                     color=generateRandomColor(), alpha=1)
        # dataFrameTwo.plot(y=f'{item.variableName}', use_index=True)
        # print(item.variableName)
    ax.legend(facecolor='white', framealpha=1, fontsize='large', shadow=True, labelspacing=2)
    ax.patch.set_edgecolor('black')
    ax.patch.set_linewidth('1')
    plt.show()


def graphMediaAverages():
    fig, ax = plt.subplots(figsize=(10, 10), facecolor='white')
    ax.set_facecolor('white')
    ax.tick_params(labelcolor='black')
    if isRFP:
            ylimit = 1200
            yax = "RFP"
    else:
        ylimit = 1
        yax = "OD"
    ax.set_title(f'Media {yax} Averages {experimentNumber}', color='black')
    ax.set_xlabel('Time (minutes)', color='black', fontsize=20)
    ax.set_ylabel(yax, color='black', fontsize=20)
    ax.set_xlim(0, 1120)
    ax.set_ylim(0, ylimit)
    dataFrameTwo = pd.read_excel('Averages Workbook.xlsx', index_col=0)
    for item in Variable.instances:
        if item.variableName == "NonInduced Media":
            plt.plot(dataFrameTwo[f'{item.variableName}'], label=f'{item.variableName}',
                     color=generateRandomColor(), alpha=1,
                     linestyle='dashed')
        elif item.variableName == "Media Induced":
            plt.plot(dataFrameTwo[f'{item.variableName}'], label=f'{item.variableName}',
                     color=generateRandomColor(), alpha=1)
        # dataFrameTwo.plot(y=f'{item.variableName}', use_index=True)
        # print(item.variableName)
    ax.legend(facecolor='white', framealpha=1, fontsize='large', shadow=True, labelspacing=2)
    ax.patch.set_edgecolor('black')
    ax.patch.set_linewidth('1')
    plt.show()


def graphBlank():
    fig, ax = plt.subplots(figsize=(10, 10), facecolor='white')
    ax.set_facecolor('white')
    ax.tick_params(labelcolor='black')
    if isRFP:
            ylimit = 1200
            yax = "RFP"
    else:
        ylimit = 1
        yax = "OD"
    ax.set_title(f'Blank {yax} Averages {experimentNumber}', color='black')
    ax.set_xlabel('Time (minutes)', color='black', fontsize=20)
    ax.set_ylabel(yax, color='black', fontsize=20)
    ax.set_xlim(0, 1120)
    ax.set_ylim(0, ylimit)
    dataFrameTwo = pd.read_excel('Averages Workbook.xlsx', index_col=0)

    for item in Variable.instances:
        match item.variableName:
            case "Blank":
                plt.plot(dataFrameTwo[f'{item.variableName}'], label=f'{item.variableName}',
                         color=generateRandomColor(), alpha=1,
                         linestyle='dashed')
        # dataFrameTwo.plot(y=f'{item.variableName}', use_index=True)
        # print(item.variableName)
    ax.legend(facecolor='white', framealpha=1, fontsize='large', shadow=True, labelspacing=2)
    ax.patch.set_edgecolor('black')
    ax.patch.set_linewidth('1')
    plt.show()


def graphAllData():
    getAllData()
    timeFrame_values_in_minutes = []
    timeframe = dataFrame.iloc[0]

    fig, ax = plt.subplots(figsize=(8, 8))

    timeFrame_values_in_minutes = [round(values/60) for values in timeframe]

    ax.set_title('Induced Variables')
    ax.set_xlabel('Time')
    ax.set_ylabel('OD')
    plt.xticks(range(len(timeFrame_values_in_minutes)),
               timeFrame_values_in_minutes)
    ax.xaxis.set_tick_params(pad=2)
    ax.set_xticks(ax.get_xticks()[::2])
    ax.set_ylim(0, 1)
    ax.set_xlim(1, 1)

    for item in Variable.instances:
        if item.variableName == 'Positive Control Induced':
            ax.plot(item.variableData, label=f'{item.variableName}', color='blue')
        # if item.isInduced:
        #     # print(f'Induced average for {item.variableName} {item.variableAvg}')
        #     ax.plot(item.variableData, label=f'{item.variableName}', color='blue')
        # elif not item.isInduced:
            print("\n")
            # print(f'Induced average for {item.variableName} {item.variableAvg}')
            ax.plot(item.variableData, label=f'{item.variableName}', color='red')
    plt.show()


def graphAllAverages():
    graphBlank()
    graphMediaAverages()
    graphLonAverages()
    graphMazAverages()
    graphControlAverages()
# writeToNewSheet()


def displayMenu():
    print("Select a graph by typing the corresponding number below.")
    selectedOption = input(""
                           "1. Graph Blanks\n"
                           "2. Graph Controls\n"
                           "3. Graph Maz\n"
                           "4. Graph Media\n"
                           "5. Graph Lon\n"
                           "6. Graph All\n")
    print(f"Selected Opt = {selectedOption}")
    return selectedOption


def selectDataToGraph(option):
    selected = option
    writeToNewSheet()
    match selected:
        case '1':
            graphBlank()
            print("Printing Blanks")
        case '2':
            graphControlAverages()
            print("Printing Controls")
        case '3':
            graphMazAverages()
            print("Printing Maz")
        case '4':
            print("Printing Media")
            graphMediaAverages()
        case '5':
            print("Printing Lon")
            graphLonAverages()
        case '6':
            print("Printing all")
            graphAllAverages()
        case _:
            print("invalid input")


selectDataToGraph(displayMenu())
