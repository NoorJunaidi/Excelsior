import sys
from PySide2.QtWidgets import QApplication, QDialog, QTextBrowser, QFileDialog, QAction, QTableWidget, \
    QPushButton, QTableWidgetItem
from PySide2.QtUiTools import QUiLoader
from PySide2.QtCore import QFile, SIGNAL, SLOT, Qt
from PySide2.QtGui import QFont, QColor
import controlWords
import invokeTools
import malScanner
import operator


class Form(QDialog):

    def setRawTextToFileContent(self, fileName, textBrowser):
        with open(fileName, "r") as myfile:
            try:
                data = myfile.read()
                textBrowser.setPlainText(data)
            except:
                return -1

    ############# Control Words start #############
    def setUncommonControlWords(self, fileName, tbControlWords):
        self.controlWordsReset()

        uncommonControlWords = controlWords.findUncommonControlWords(
            fileName)  # returns a counter object; contains all uncommon control word with count

        # if not rtf file
        if uncommonControlWords == -1:
            return

        tbControlWords.setRowCount(len(uncommonControlWords))

        # Populate the table
        for iter, word in enumerate(uncommonControlWords):
            # word length
            newItemLen = QTableWidgetItem(str(len(word)))

            # control word with checkbox
            newItemWord = QTableWidgetItem(word)
            newItemWord.setFlags(Qt.ItemIsUserCheckable |
                                 Qt.ItemIsEnabled)
            newItemWord.setCheckState(Qt.Unchecked)

            # word count
            newItemCount = QTableWidgetItem(str(uncommonControlWords[word]))

            # add to table
            tbControlWords.setItem(iter, 0, newItemLen)
            tbControlWords.setItem(iter, 1, newItemWord)
            tbControlWords.setItem(iter, 2, newItemCount)

        tbControlWords.resizeColumnsToContents()

    # To be used with add to whitelist button
    def addWordsToTempWhitelist(self, item):
        # buggy; clicked box if already checked, will duplicate
        if item.checkState() == Qt.Checked:
            self.newWhitelistList.append(item.text())
        else:
            self.newWhitelistList.remove(item.text())

    def controlWordsReset(self):
        tbControlWords = self.window.findChild(QTableWidget, 'tbControlWords')
        tbControlWords.clearContents()

    def addWordsToWhitelist(self):
        # if no file loaded, exit
        if self.openedFile == "":
            return

        # Append templist to file (call func in controlWords.py)
        controlWords.addToWhiteList(self.newWhitelistList)

        # Reset the display
        self.controlWordsReset()
        tbControlWords = self.window.findChild(QTableWidget, 'tbControlWords')
        self.setUncommonControlWords(self.openedFile, tbControlWords)

        # clear the temp list
        self.newWhitelistList.clear()

    ############# Control words end #############

    ############# oledump Summary start #############

    def setSummary(self, filename, tbSummary):
        self.summaryReset()

        # returns dictionary of summary
        resultList = invokeTools.oledumpMetadata(filename)
        # if not ole file
        if resultList == -1:
            return

        # set table row num
        # tbSummary.setRowCount(len(resultDict))

        # fill table
        # iterate through results

        # change from dict to list format (no keys)
        for index, line in enumerate(resultList):
            tbSummary.insertRow(index)  # add row
            # column 1 Attribute
            newItemeAtt = QTableWidgetItem(line[0])
            tbSummary.setItem(index, 0, newItemeAtt)
            # column 2 Value
            newItemeVal = QTableWidgetItem(line[1])
            tbSummary.setItem(index, 1, newItemeVal)


        # old code **********************************************************************
        # for attribute, value in resultDict.items():
        #     tbSummary.insertRow(counter)  # add row
        #     if value == "TITLE234":
        #         if counter != 0:  # for title start of table no need extra space
        #             tbSummary.insertRow(counter)  # add row
        #             counter += 1
        #
        #         # column 1 Attribute
        #         newItemeAtt = QTableWidgetItem(attribute)
        #         tbSummary.setItem(counter, 0, newItemeAtt)
        #         # make bold
        #         font = QFont()
        #         font.setBold(True)
        #         newItemeAtt.setFont(font)
        #
        #         counter += 1
        #         continue
        #
        #     # column 1 Attribute
        #     newItemeAtt = QTableWidgetItem(attribute)
        #     tbSummary.setItem(counter, 0, newItemeAtt)
        #     # column 2 Value
        #     newItemeVal = QTableWidgetItem(value)
        #     tbSummary.setItem(counter, 1, newItemeVal)
        #
        #     counter += 1
        # old code **********************************************************************

        tbSummary.resizeColumnsToContents()

    def summaryReset(self):
        tbSummary = self.window.findChild(QTableWidget, 'tbSummary')
        tbSummary.clearContents()
        tbSummary.setRowCount(0)

    ############# oledump Summary end #############

    ############# oledump Macros start #############
    def setFoundMacros(self, fileName, tbMacroDetails, tbMacroDecoded):
        # call main macro function
        self.foundMacros = invokeTools.oledumpMacro(fileName)
        # TODO order my M then streamIndex
        self.foundMacros = sorted(self.foundMacros, key=operator.itemgetter(1, 0))

        # set the MacroDetails details
        # find number of rows
        noOfRows = len(self.foundMacros)
        # set number of rows in the table
        tbMacroDetails.setRowCount(noOfRows)

        # populate the table
        for iter, macroDetails in enumerate(self.foundMacros):
            # foundMacros = [streamNo, macro type, size(bytes), stream name, decodedMacro]
            # each cell
            newItemStreamIndex = QTableWidgetItem(macroDetails[0])
            tbMacroDetails.setItem(iter, 0, newItemStreamIndex)

            # stream name
            newItemStreamName = QTableWidgetItem(macroDetails[3])
            tbMacroDetails.setItem(iter, 1, newItemStreamName)

            # stream type
            # if attribute only, "Attributes only"
            if macroDetails[1] == "m":
                newItemType = QTableWidgetItem("Attributes Only")

            # if has macro, "Macro" + red box
            else:
                newItemType = QTableWidgetItem("Macro Found")
                # make darkRed background
                newBackgroundColor = QColor("red")
                newItemType.setBackgroundColor(newBackgroundColor)
                # make white text
                newTextColor = QColor("white")
                newItemType.setTextColor(newTextColor)
            tbMacroDetails.setItem(iter, 3, newItemType)

            # bytes
            newItemBytes = QTableWidgetItem(macroDetails[2] + " bytes")
            tbMacroDetails.setItem(iter, 2, newItemBytes)

        # set the raw macro decoded textBrowser for first one
        tbMacroDecoded.setPlainText(self.foundMacros[0][4])
        tbMacroDetails.resizeColumnsToContents()

    def changeMacroTextbox(self, row, column):
        # get textbrowser
        tbMacroDecoded = self.window.findChild(QTextBrowser, 'tbMacroDecoded')
        tbMacroDecoded.setPlainText(self.foundMacros[row][4])

    ############# oledump Macros end #############

    ############# OfficeMalScanner start #############
    def setMalScanner(self, filename, tbScanResults, tbExtractedMacro):
        # run main malScanner function
        self.yaraMatches = malScanner.mainMalScanner(filename)

        # TODO if return -1; ignore?
        if self.yaraMatches == -1:
            return

        # TODO
        # Count no of rows first
        noOfRows = 0
        for record in self.yaraMatches:
            # if empty +1
            if not record[3]:
                noOfRows += 1
            # if not empty count number of vba files found
            else:
                noOfRows += len(record[3])
        # set number of rows in the table
        tbScanResults.setRowCount(noOfRows)

        # populate the table
        rowCounter = 0
        # to display the first extracted VBA to textbox
        isDisplayed = False
        for record in self.yaraMatches:
            # if there is vba record, make entry for each vba
            if record[3]:
                # loop thorough vba items
                for vbaFile in record[3]:
                    # set the text of the extracted macro text box
                    if not isDisplayed:
                        with open(vbaFile, "r") as myfile:
                            try:
                                data = myfile.read()
                                tbExtractedMacro.setPlainText(data)
                                isDisplayed = True
                            except:
                                continue

                    # File of interest
                    newItemInterest = QTableWidgetItem(record[0])
                    tbScanResults.setItem(rowCounter, 0, newItemInterest)

                    # Is malicious?
                    # if true, "Malicious", darkRed back, white fore
                    if record[4] == True:
                        newItemIsMal = QTableWidgetItem("Malicious")
                        # make darkRed background
                        newBackgroundColor = QColor("red")
                        newItemIsMal.setBackgroundColor(newBackgroundColor)
                        # make white text
                        newTextColor = QColor("white")
                        newItemIsMal.setTextColor(newTextColor)


                    # if false, "Non Malicious", "green" back, regular fore
                    else:
                        newItemIsMal = QTableWidgetItem("Non Malicious")
                        # make darkRed background
                        newBackgroundColor = QColor("lawngreen")
                        newItemIsMal.setBackgroundColor(newBackgroundColor)

                    tbScanResults.setItem(rowCounter, 1, newItemIsMal)

                    # File format
                    newItemFormat = QTableWidgetItem(record[1][0].rule)
                    tbScanResults.setItem(rowCounter, 2, newItemFormat)

                    # Vba file location
                    newItemExtractedLoc = QTableWidgetItem(vbaFile)
                    tbScanResults.setItem(rowCounter, 3, newItemExtractedLoc)

                    rowCounter += 1

            # if no vba record, only do 1 line
            else:
                # File of interest
                newItemInterest = QTableWidgetItem(record[0])
                tbScanResults.setItem(rowCounter, 0, newItemInterest)

                # Is malicious?
                # if true, "Malicious", darkRed back, white fore
                if record[4] == True:
                    newItemIsMal = QTableWidgetItem("Malicious")
                    # make darkRed background
                    newBackgroundColor = QColor("red")
                    newItemIsMal.setBackgroundColor(newBackgroundColor)
                    # make white text
                    newTextColor = QColor("white")
                    newItemIsMal.setTextColor(newTextColor)


                # if false, "Non Malicious", "green" back, regular fore
                else:
                    newItemIsMal = QTableWidgetItem("Non Malicious")
                    # make darkRed background
                    newBackgroundColor = QColor("lawngreen")
                    newItemIsMal.setBackgroundColor(newBackgroundColor)

                tbScanResults.setItem(rowCounter, 1, newItemIsMal)

                # File format
                newItemFormat = QTableWidgetItem(record[1][0].rule)
                tbScanResults.setItem(rowCounter, 2, newItemFormat)

                # Vba file location
                newItemExtractedLoc = QTableWidgetItem('')
                tbScanResults.setItem(rowCounter, 3, newItemExtractedLoc)

                rowCounter += 1
        tbScanResults.resizeColumnsToContents()

    def changeMalScannerTextbox(self, row, column): # row and col retrieved; play with yaraMatches[]
        # get column 3
        tbScanResults = self.window.findChild(QTableWidget, 'tbScanResults')
        cellWithPathToMacro = tbScanResults.item(row, 3)
        pathToMacro = cellWithPathToMacro.text()

        # use path to change textBrowser
        tbExtractedMacro = self.window.findChild(QTextBrowser, 'tbExtractedMacro')

        try:
            with open(pathToMacro, "r") as myfile:
                data = myfile.read()
                tbExtractedMacro.setPlainText(data)
        except:
            tbExtractedMacro.setPlainText("")



    ############# OfficeMalScanner end #############

    # The main set up function
    def openFile(self):
        fileName = QFileDialog.getOpenFileName(self, "Open Document", "C:\\",
                                               "All Files (*.*)")
        self.openedFile = fileName[0]

        # # Run all auto functions when a file is loaded.
        # Oledump Metadata
        tbSummary = self.window.findChild(QTableWidget, 'tbSummary')
        self.setSummary(fileName[0], tbSummary)

        # Control words
        tbControlWords = self.window.findChild(QTableWidget, 'tbControlWords')
        self.setUncommonControlWords(fileName[0], tbControlWords)

        # Raw text
        fileContentBox = self.window.findChild(QTextBrowser, 'tbFileContent')
        self.setRawTextToFileContent(fileName[0], fileContentBox)

        # TODO macro
        tbMacroDetails = self.window.findChild(QTableWidget, 'tbMacroDetails')
        tbMacroDecoded = self.window.findChild(QTextBrowser, 'tbMacroDecoded')
        self.setFoundMacros(fileName[0], tbMacroDetails, tbMacroDecoded)

        # # TODO MalScanner
        # tbScanResults = self.window.findChild(QTableWidget, 'tbScanResults')
        # tbExtractedMacro = self.window.findChild(QTextBrowser, 'tbExtractedMacro')
        # self.setMalScanner(fileName[0], tbScanResults, tbExtractedMacro)

    def __init__(self, ui_file, parent=None):
        ### Required
        super(Form, self).__init__(parent)
        ui_file = QFile(ui_file)
        ui_file.open(QFile.ReadOnly)

        loader = QUiLoader()
        self.window = loader.load(ui_file)
        ui_file.close()
        ###

        # Custom variables
        self.openedFile = ""
        self.newWhitelistList = []
        self.foundMacros = []
        self.macrosDecoded = []
        self.yaraMatches = []

        ## catch the elements to change later
        actionOpen = self.window.findChild(QAction, 'actionOpen')
        btnControlWords = self.window.findChild(QPushButton, 'btnControlWords')
        tbControlWords = self.window.findChild(QTableWidget, 'tbControlWords')
        # tbScanResults = self.window.findChild(QTableWidget, 'tbScanResults')
        tbMacroDetails = self.window.findChild(QTableWidget, 'tbMacroDetails')

        ## Connects all action button to their slots (functions)
        self.connect(actionOpen, SIGNAL("triggered()"), self, SLOT("openFile()"))
        # self.connect(tbControlWords, SIGNAL("itemClicked"), self, SLOT("addWordsToTempWhitelist"))
        tbControlWords.itemClicked.connect(self.addWordsToTempWhitelist)  # alt to above; working but buggy ^
        self.connect(btnControlWords, SIGNAL("clicked()"), self, SLOT("addWordsToWhitelist()"))
        # tbScanResults.cellClicked.connect(self.changeMalScannerTextbox)
        tbMacroDetails.cellClicked.connect(self.changeMacroTextbox)

        self.window.show()  # Required


if __name__ == '__main__':
    app = QApplication(sys.argv)
    form = Form('main.ui')
    sys.exit(app.exec_())
