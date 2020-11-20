import errno
import os
import sys
from PySide2.QtWidgets import QApplication, QDialog, QTextBrowser, QFileDialog, QAction, QTableWidget, \
    QPushButton, QTableWidgetItem, QComboBox, QMessageBox
from PySide2.QtUiTools import QUiLoader
from PySide2.QtCore import QFile, SIGNAL, SLOT
from PySide2.QtGui import QColor
import invokeTools, vbaMain, powershellMain
import operator
from inspect import signature


class Form(QDialog):

    def setRawTextToFileContent(self, fileName, textBrowser):
        with open(fileName, "r") as myfile:
            try:
                data = myfile.read()
                textBrowser.setPlainText(data)
            except:
                return -1

    ############# oledump Summary start #############

    def setMetadata(self, filename, tbMetadata):
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

        for index, line in enumerate(resultList):
            tbMetadata.insertRow(index)  # add row
            # column 1 Attribute
            newItemeAtt = QTableWidgetItem(line[0])
            tbMetadata.setItem(index, 0, newItemeAtt)
            # column 2 Value
            newItemeVal = QTableWidgetItem(line[1])
            tbMetadata.setItem(index, 1, newItemeVal)

        tbMetadata.resizeColumnsToContents()

    def summaryReset(self):
        tbMetadata = self.window.findChild(QTableWidget, 'tbSummary')
        tbMetadata.clearContents()
        tbMetadata.setRowCount(0)

    ############# oledump Summary end #############

    ############# oledump Macros start #############
    def setFoundMacros(self, fileName, tbMacroDetails, tbMacroDecoded):
        # call main macro function
        self.foundMacros = invokeTools.oledumpMacro(fileName)

        # if no macros found, exit
        if self.foundMacros == []:
            return

        # order my M then streamIndex
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

    def exportMacrosToFolder(self):
        # if no macros, leave function
        if not self.foundMacros:
            return

        folderName = QFileDialog.getExistingDirectory(self, "Select folder to export to", "C:\\")

        for item in self.foundMacros:
            newFilePath = folderName + "/" + item[3]

            if not os.path.exists(os.path.dirname(newFilePath)):
                try:
                    os.makedirs(os.path.dirname(newFilePath))
                except OSError as exc:  # Guard against race condition
                    if exc.errno != errno.EEXIST:
                        raise

            f = open(newFilePath, "w")
            f.write(item[4])
            f.close()

    ############# oledump Macros end #############

    ############# VBA deobsfuscation start #############
    # main setup VBA deobf function
    def setVBADeobfPage(self):
        tbAvailObf = self.window.findChild(QTableWidget, 'tbAvailObf')  # available operations
        tbCurrObf = self.window.findChild(QTableWidget, 'tbCurrObf')  # applied operations

        self.populateAvailObfTable(tbAvailObf)

        # have column 3 invisible
        tbAvailObf.setColumnHidden(2, True)
        tbCurrObf.setColumnHidden(2, True)

    # add items to found macros combobox
    def populateFoundMacroComboBox(self, cbDecodedMac, tbInputVba):
        # delete all first; in case of second file run
        cbDecodedMac.clear()

        # if no macros, exit
        if self.foundMacros == []:
            return

        # [3] = name to fill in combobox
        # [4] = macro to fill in tbInputVba
        for item in self.foundMacros:
            cbDecodedMac.addItem(item[3])

        # set text for first macro
        tbInputVba.setPlainText(self.foundMacros[0][4])

    def populateAvailObfTable(self, tbAvailObf):
        # find number of rows
        # count number of 0s in [2]
        noOfRows = 0
        for item in self.ALLVBADEOBFOPERATIONS:
            if item[2] == 0:
                noOfRows += 1

        # set number of rows in the table
        tbAvailObf.setRowCount(noOfRows)

        # populate the table
        currRow = 0
        for iter, item in enumerate(self.ALLVBADEOBFOPERATIONS):
            # if operation need document, skip; add operation when document is opened
            if item[2] == 0:
                # item = [display name, func name, isDocRequired]
                displayName = item[0]
                functionName = item[1]
                # operation display name
                newItemOperationDisplayName = QTableWidgetItem(displayName)
                tbAvailObf.setItem(currRow, 0, newItemOperationDisplayName)

                # add + "button"
                newItemAddButton = QTableWidgetItem("+")
                newItemAddButton.setTextAlignment(0x0004 | 0x0080)  # AlignHCenter | AlignVCenter
                tbAvailObf.setItem(currRow, 1, newItemAddButton)

                # function name; invisible
                newItemFuncName = QTableWidgetItem(functionName)
                tbAvailObf.setItem(currRow, 2, newItemFuncName)

                currRow += 1

        tbAvailObf.resizeColumnsToContents()

    # add operation to tbCurrObf when + is clicked
    def addVBAOperation(self, row, column):
        # add operation only when + is clicked
        if column == 1:
            tbAvailObf = self.window.findChild(QTableWidget, 'tbAvailObf')
            displayNameItem = tbAvailObf.item(row, 0)
            displayName = displayNameItem.text()
            funcNameItem = tbAvailObf.item(row, 2)
            funcName = funcNameItem.text()

            tbCurrObf = self.window.findChild(QTableWidget, 'tbCurrObf')

            # count and add another row
            rowCount = tbCurrObf.rowCount()
            tbCurrObf.insertRow(rowCount)

            # operation display name
            newItemOperationDisplayName = QTableWidgetItem(displayName)
            tbCurrObf.setItem(rowCount, 0, newItemOperationDisplayName)

            # to add remove "button"
            newItemAddButton = QTableWidgetItem("-")
            newItemAddButton.setTextAlignment(0x0004 | 0x0080)  # AlignHCenter | AlignVCenter
            tbCurrObf.setItem(rowCount, 1, newItemAddButton)

            # function name; invisible
            newItemFuncName = QTableWidgetItem(funcName)
            tbCurrObf.setItem(rowCount, 2, newItemFuncName)

            # ADD ^ and V BUTTONS
            newItemUpArrow = QTableWidgetItem("▲")
            newItemUpArrow.setTextAlignment(0x0004 | 0x0080)  # AlignHCenter | AlignVCenter
            tbCurrObf.setItem(rowCount, 3, newItemUpArrow)

            newItemDownArrow = QTableWidgetItem("▼")
            newItemDownArrow.setTextAlignment(0x0004 | 0x0080)  # AlignHCenter | AlignVCenter
            tbCurrObf.setItem(rowCount, 4, newItemDownArrow)

            tbCurrObf.resizeColumnsToContents()

    # remove operation from tbCurrObf when - is clicked
    def actionVBAOperation(self, row, column):
        tbCurrObf = self.window.findChild(QTableWidget, 'tbCurrObf')  # applied operations
        displayNameItem = tbCurrObf.item(row, 0)
        displayName = displayNameItem.text()
        funcNameItem = tbCurrObf.item(row, 2)
        funcName = funcNameItem.text()

        # remove operation only when - is clicked
        if column == 1:
            tbCurrObf = self.window.findChild(QTableWidget, 'tbCurrObf')  # applied operations
            tbCurrObf.removeRow(row)
            # tbCurrObf.resizeColumnsToContents()
            return


        # up arrow; move operation up
        elif column == 3:
            # if at top no need to move
            if row == 0:
                return

            # remove current row
            tbCurrObf.removeRow(row)

            # readding row here
            # add row into row - 1
            insertRow = row - 1
            tbCurrObf.insertRow(insertRow)

            # operation display name
            newItemOperationDisplayName = QTableWidgetItem(displayName)
            tbCurrObf.setItem(insertRow, 0, newItemOperationDisplayName)

            # to add remove "button"
            newItemAddButton = QTableWidgetItem("-")
            newItemAddButton.setTextAlignment(0x0004 | 0x0080)  # AlignHCenter | AlignVCenter
            tbCurrObf.setItem(insertRow, 1, newItemAddButton)

            # function name; invisible
            newItemFuncName = QTableWidgetItem(funcName)
            tbCurrObf.setItem(insertRow, 2, newItemFuncName)

            # ADD ^ and V BUTTONS
            newItemUpArrow = QTableWidgetItem("▲")
            newItemUpArrow.setTextAlignment(0x0004 | 0x0080)  # AlignHCenter | AlignVCenter
            tbCurrObf.setItem(insertRow, 3, newItemUpArrow)

            newItemDownArrow = QTableWidgetItem("▼")
            newItemDownArrow.setTextAlignment(0x0004 | 0x0080)  # AlignHCenter | AlignVCenter
            tbCurrObf.setItem(insertRow, 4, newItemDownArrow)
            return

        # down arrow; move operation down
        elif column == 4:
            # if at last no need to move
            if row == tbCurrObf.rowCount() - 1:
                return

            # remove current row
            tbCurrObf.removeRow(row)

            # reading row here
            # add row into row - 1
            insertRow = row + 1
            tbCurrObf.insertRow(insertRow)

            # operation display name
            newItemOperationDisplayName = QTableWidgetItem(displayName)
            tbCurrObf.setItem(insertRow, 0, newItemOperationDisplayName)

            # to add remove "button"
            newItemAddButton = QTableWidgetItem("-")
            newItemAddButton.setTextAlignment(0x0004 | 0x0080)  # AlignHCenter | AlignVCenter
            tbCurrObf.setItem(insertRow, 1, newItemAddButton)

            # function name; invisible
            newItemFuncName = QTableWidgetItem(funcName)
            tbCurrObf.setItem(insertRow, 2, newItemFuncName)

            # ADD ^ and V BUTTONS
            newItemUpArrow = QTableWidgetItem("▲")
            newItemUpArrow.setTextAlignment(0x0004 | 0x0080)  # AlignHCenter | AlignVCenter
            tbCurrObf.setItem(insertRow, 3, newItemUpArrow)

            newItemDownArrow = QTableWidgetItem("▼")
            newItemDownArrow.setTextAlignment(0x0004 | 0x0080)  # AlignHCenter | AlignVCenter
            tbCurrObf.setItem(insertRow, 4, newItemDownArrow)
            return

    # apply operations to inputTextBrowser MAIN
    def applyVBAOps(self):
        # get operations from tbCurrObf
        tbCurrObf = self.window.findChild(QTableWidget, 'tbCurrObf')  # applied operations
        rowCount = tbCurrObf.rowCount()

        # get input text
        tbInputVba = self.window.findChild(QTextBrowser, 'tbInputVba')
        originalText = tbInputVba.toPlainText()

        # loop through selected operations
        for index in range(rowCount):
            # get the function name to call
            item = tbCurrObf.item(index, 2)
            funcName = item.text()

            # run the functions in order
            # apply operation one by one on input TextBrowser
            method_to_call = getattr(vbaMain, funcName)
            sig = signature(method_to_call)
            params = sig.parameters
            numOfParam = len(params)

            # most operations only need originalText
            if numOfParam == 1:
                originalText = method_to_call(originalText)

            # replaceFormVariablesWithValue needs filename + originalText
            elif numOfParam == 2:
                originalText = method_to_call(self.openedFile, originalText)

        # put the result into
        tbOutputVba = self.window.findChild(QTextBrowser, 'tbOutputVba')
        tbOutputVba.setPlainText(originalText)

    # changes input textbrowser when combobox is changed. For combobox signal.
    def updateInputVBATextBrowser(self, index):
        if self.foundMacros == []:
            return

        tbInputVba = self.window.findChild(QTextBrowser, 'tbInputVba')
        tbInputVba.setPlainText(self.foundMacros[index][4])

    # save deobf VBA output to file
    def saveOutputVBA(self):
        # if empty textbox, leave function
        tbOutputVba = self.window.findChild(QTextBrowser, 'tbOutputVba')
        output = tbOutputVba.toPlainText().strip()
        if output == "":
            return

        folderName = QFileDialog.getSaveFileName(self, "Select file to export to", "C:\\",
                                                 "Text (*.txt);;All Files (*)")
        toCreate = folderName[0]
        if toCreate == "":  # exits if x or cancel button pressed
            return

        f = open(toCreate, "w")
        f.write(output)
        f.close()

        # POPUP
        msgBox = QMessageBox()
        msgBox.setText("Output saved to " + toCreate)
        msgBox.exec_()

    # Copies text in output to input VBA
    # save to self.macros when button clicked
    def copyOutputToInputVBA(self):
        cbDecodedMac = self.window.findChild(QComboBox, 'cbDecodedMac')
        currentIndex = cbDecodedMac.currentIndex()
        # if no macros loaded, exit func
        if currentIndex == -1:
            return

        tbOutputVba = self.window.findChild(QTextBrowser, 'tbOutputVba')
        tbInputVba = self.window.findChild(QTextBrowser, 'tbInputVba')

        # save even if output empty
        if tbOutputVba.toPlainText().strip() == "":
            toSave = tbInputVba.toPlainText()
            self.foundMacros[currentIndex][4] = toSave

        # replace toCopy to self.foundMacros; changes are kept
        else:
            toCopy = tbOutputVba.toPlainText()
            tbInputVba.setPlainText(toCopy)
            tbOutputVba.setPlainText("")
            self.foundMacros[currentIndex][4] = toCopy

    def closePopup(self):
        self.popup.close()

    # display helpful hints for the analyst
    def displayVBAHints(self):
        outputText = ""

        # display vba hints
        # only if there is a file opened
        if self.openedFile != "":
            olevbaHintsresult = vbaMain.getOlevbaHints(self.openedFile)
            outputText += olevbaHintsresult
            outputText += "\n"

            # display where auto exec keyword located
            autoExecResult = vbaMain.locateAutoExec(olevbaHintsresult, self.foundMacros)
            autoExecLocationsStr = ""
            for locList in autoExecResult:
                autoExecLocationsStr += "\n" + locList[0] + " is found in " + locList[1]

            outputText += autoExecLocationsStr

            # display part
            popupFileName = "DialogSusPatterns.ui"
            ui_file = QFile(popupFileName)
            ui_file.open(QFile.ReadOnly)

            loader = QUiLoader()
            self.popup = loader.load(ui_file)

            # set the close button
            btnCloseSusPatterns = self.popup.findChild(QPushButton, 'btnCloseSusPatterns')
            btnCloseSusPatterns.clicked.connect(self.closePopup)

            # set the textbox and button
            tbPatternOutput = self.popup.findChild(QTextBrowser, 'tbPatternOutput')
            tbPatternOutput.setPlainText(outputText)

            self.popup.exec_()

    ############# VBA deobsfuscation end #############

    ############# PowerShell deobsfuscation start #############
    # set up main PS deobf page
    def setPowershellDeobfPage(self):
        # need to populate Powershell available operations
        tbPSAvailOps = self.window.findChild(QTableWidget, 'tbPSAvailOps')
        tbPSCurrOps = self.window.findChild(QTableWidget, 'tbPSCurrOps')

        self.populateAvailPSObfTable(tbPSAvailOps)

        # have column 3 invisible
        tbPSAvailOps.setColumnHidden(2, True)
        tbPSCurrOps.setColumnHidden(2, True)

    # populate avail PS ops with operation and + button
    def populateAvailPSObfTable(self, tbPSAvailOps):
        # find number of rows
        # count number of 0s in [2]
        noOfRows = 0
        for item in self.ALLPOWERSHELLDEOBFOPERATIONS:
            if item[2] == 0:
                noOfRows += 1

        # set number of rows in the table
        tbPSAvailOps.setRowCount(noOfRows)

        # populate the table
        currRow = 0
        for iter, item in enumerate(self.ALLPOWERSHELLDEOBFOPERATIONS):
            if item[2] == 0:
                # item = [display name, func name, isDocRequired]
                displayName = item[0]
                functionName = item[1]
                # operation display name
                newItemOperationDisplayName = QTableWidgetItem(displayName)
                tbPSAvailOps.setItem(currRow, 0, newItemOperationDisplayName)

                # add "button"
                newItemAddButton = QTableWidgetItem("+")
                newItemAddButton.setTextAlignment(0x0004 | 0x0080)  # AlignHCenter | AlignVCenter
                tbPSAvailOps.setItem(currRow, 1, newItemAddButton)

                # function name; invisible
                newItemFuncName = QTableWidgetItem(functionName)
                tbPSAvailOps.setItem(currRow, 2, newItemFuncName)

                currRow += 1

        tbPSAvailOps.resizeColumnsToContents()

    # add operation to tbPSCurrObf when + is clicked
    def addPSOperation(self, row, column):
        # add operation only when + is clicked
        if column == 1:
            tbPSAvailOps = self.window.findChild(QTableWidget, 'tbPSAvailOps')
            displayNameItem = tbPSAvailOps.item(row, 0)
            displayName = displayNameItem.text()
            funcNameItem = tbPSAvailOps.item(row, 2)
            funcName = funcNameItem.text()

            tbPSCurrOps = self.window.findChild(QTableWidget, 'tbPSCurrOps')

            # count and add another row
            rowCount = tbPSCurrOps.rowCount()
            tbPSCurrOps.insertRow(rowCount)

            # operation display name
            newItemOperationDisplayName = QTableWidgetItem(displayName)
            tbPSCurrOps.setItem(rowCount, 0, newItemOperationDisplayName)

            # to add remove "button"
            newItemAddButton = QTableWidgetItem("-")
            newItemAddButton.setTextAlignment(0x0004 | 0x0080)  # AlignHCenter | AlignVCenter
            tbPSCurrOps.setItem(rowCount, 1, newItemAddButton)

            # function name; invisible
            newItemFuncName = QTableWidgetItem(funcName)
            tbPSCurrOps.setItem(rowCount, 2, newItemFuncName)

            # ADD ^ and V BUTTONS
            newItemUpArrow = QTableWidgetItem("▲")
            newItemUpArrow.setTextAlignment(0x0004 | 0x0080)  # AlignHCenter | AlignVCenter
            tbPSCurrOps.setItem(rowCount, 3, newItemUpArrow)

            newItemDownArrow = QTableWidgetItem("▼")
            newItemDownArrow.setTextAlignment(0x0004 | 0x0080)  # AlignHCenter | AlignVCenter
            tbPSCurrOps.setItem(rowCount, 4, newItemDownArrow)

            tbPSCurrOps.resizeColumnsToContents()

    # remove operation from tbPSCurrObf when - is clicked
    def actionPSOperation(self, row, column):
        tbPSCurrOps = self.window.findChild(QTableWidget, 'tbPSCurrOps')  # applied operations
        displayNameItem = tbPSCurrOps.item(row, 0)
        displayName = displayNameItem.text()
        funcNameItem = tbPSCurrOps.item(row, 2)
        funcName = funcNameItem.text()

        # remove operation only when - is clicked
        if column == 1:
            tbPSCurrOps.removeRow(row)
            # tbPSCurrOps.resizeColumnsToContents()
            return

        # up arrow; move operation up
        elif column == 3:
            # if at top no need to move
            if row == 0:
                return

            # remove current row
            tbPSCurrOps.removeRow(row)

            # readding row here
            # add row into row - 1
            insertRow = row - 1
            tbPSCurrOps.insertRow(insertRow)

            # operation display name
            newItemOperationDisplayName = QTableWidgetItem(displayName)
            tbPSCurrOps.setItem(insertRow, 0, newItemOperationDisplayName)

            # to add remove "button"
            newItemAddButton = QTableWidgetItem("-")
            newItemAddButton.setTextAlignment(0x0004 | 0x0080)  # AlignHCenter | AlignVCenter
            tbPSCurrOps.setItem(insertRow, 1, newItemAddButton)

            # function name; invisible
            newItemFuncName = QTableWidgetItem(funcName)
            tbPSCurrOps.setItem(insertRow, 2, newItemFuncName)

            # ADD ^ and V BUTTONS
            newItemUpArrow = QTableWidgetItem("▲")
            newItemUpArrow.setTextAlignment(0x0004 | 0x0080)  # AlignHCenter | AlignVCenter
            tbPSCurrOps.setItem(insertRow, 3, newItemUpArrow)

            newItemDownArrow = QTableWidgetItem("▼")
            newItemDownArrow.setTextAlignment(0x0004 | 0x0080)  # AlignHCenter | AlignVCenter
            tbPSCurrOps.setItem(insertRow, 4, newItemDownArrow)
            return

        # down arrow; move operation down
        elif column == 4:
            # if at last no need to move
            if row == tbPSCurrOps.rowCount() - 1:
                return

            # remove current row
            tbPSCurrOps.removeRow(row)

            # reading row here
            # add row into row - 1
            insertRow = row + 1
            tbPSCurrOps.insertRow(insertRow)

            # operation display name
            newItemOperationDisplayName = QTableWidgetItem(displayName)
            tbPSCurrOps.setItem(insertRow, 0, newItemOperationDisplayName)

            # to add remove "button"
            newItemAddButton = QTableWidgetItem("-")
            newItemAddButton.setTextAlignment(0x0004 | 0x0080)  # AlignHCenter | AlignVCenter
            tbPSCurrOps.setItem(insertRow, 1, newItemAddButton)

            # function name; invisible
            newItemFuncName = QTableWidgetItem(funcName)
            tbPSCurrOps.setItem(insertRow, 2, newItemFuncName)

            # ADD ^ and V BUTTONS
            newItemUpArrow = QTableWidgetItem("▲")
            newItemUpArrow.setTextAlignment(0x0004 | 0x0080)  # AlignHCenter | AlignVCenter
            tbPSCurrOps.setItem(insertRow, 3, newItemUpArrow)

            newItemDownArrow = QTableWidgetItem("▼")
            newItemDownArrow.setTextAlignment(0x0004 | 0x0080)  # AlignHCenter | AlignVCenter
            tbPSCurrOps.setItem(insertRow, 4, newItemDownArrow)

    # apply operations to tbInputPS textBrowser MAIN
    def applyPSOps(self):
        # get operations from tbCurrObf
        tbPSCurrOps = self.window.findChild(QTableWidget, 'tbPSCurrOps')  # applied operations
        rowCount = tbPSCurrOps.rowCount()

        # get input text
        tbInputPS = self.window.findChild(QTextBrowser, 'tbInputPS')
        originalText = tbInputPS.toPlainText()

        # loop through selected operations
        for index in range(rowCount):
            # get the function name to call
            item = tbPSCurrOps.item(index, 2)
            funcName = item.text()

            # run the functions in order
            # apply operation one by one on input TextBrowser
            method_to_call = getattr(powershellMain, funcName)
            sig = signature(method_to_call)
            params = sig.parameters
            numOfParam = len(params)

            # most operations only need originalText
            if numOfParam == 1:
                originalText = method_to_call(originalText)

            # only if somehow need document; so far none
            # replaceFormVariablesWithValue needs filename + originalText
            elif numOfParam == 2:
                originalText = method_to_call(self.openedFile, originalText)

        # put the result into
        tbOutputPS = self.window.findChild(QTextBrowser, 'tbOutputPS')
        tbOutputPS.setPlainText(originalText)

    # save deobf PS output to file
    def saveOutputPS(self):
        # if empty textbox, leave function
        tbOutputPS = self.window.findChild(QTextBrowser, 'tbOutputPS')
        output = tbOutputPS.toPlainText().strip()
        if output == "":
            return

        folderName = QFileDialog.getSaveFileName(self, "Select file to export to", "C:\\",
                                                 "Text (*.txt);;All Files (*)")
        toCreate = folderName[0]
        if toCreate == "":  # exits if x or cancel button pressed
            return

        f = open(toCreate, "w")
        f.write(output)
        f.close()

        # POPUP
        msgBox = QMessageBox()
        msgBox.setText("Output saved to " + toCreate)
        msgBox.exec_()

    # Copies text in output to input PS
    def copyOutputToInputPS(self):
        tbOutputPS = self.window.findChild(QTextBrowser, 'tbOutputPS')
        tbInputPS = self.window.findChild(QTextBrowser, 'tbInputPS')

        if tbOutputPS.toPlainText().strip() == "":
            return

        toCopy = tbOutputPS.toPlainText()
        tbInputPS.setPlainText(toCopy)
        tbOutputPS.setPlainText("")

    ############# PowerShell deobsfuscation end #############

    def addDocumentDependentOps(self, tbAvailObf, tbPSAvailOps):
        # find number of rows
        # count number of 0s in [2]
        noOfAdditionalRows = 0
        for item in self.ALLVBADEOBFOPERATIONS:
            if item[2] == 1:
                noOfAdditionalRows += 1

        # set number of rows in the table
        currentRow = tbAvailObf.rowCount()
        tbAvailObf.setRowCount(currentRow + noOfAdditionalRows)

        # populate the table
        for iter, item in enumerate(self.ALLVBADEOBFOPERATIONS):
            # if operation need document, skip; add operation when document is opened
            if item[2] == 1:
                # item = [display name, func name, isDocRequired]
                displayName = item[0]
                functionName = item[1]
                # operation display name
                newItemOperationDisplayName = QTableWidgetItem(displayName)
                tbAvailObf.setItem(currentRow, 0, newItemOperationDisplayName)

                # add + "button"
                newItemAddButton = QTableWidgetItem("+")
                newItemAddButton.setTextAlignment(0x0004 | 0x0080)  # AlignHCenter | AlignVCenter
                tbAvailObf.setItem(currentRow, 1, newItemAddButton)

                # function name; invisible
                newItemFuncName = QTableWidgetItem(functionName)
                tbAvailObf.setItem(currentRow, 2, newItemFuncName)

                currentRow += 1

        tbAvailObf.resizeColumnsToContents()
        self.docDependantIsLoaded = 1
        return

    # The main set up function
    def openFile(self):
        fileName = QFileDialog.getOpenFileName(self, "Open Document", "C:\\",
                                               "All Files (*.*)")
        self.openedFile = fileName[0]

        # # Run all auto functions when a file is loaded.
        # Oledump Metadata
        tbMetadata = self.window.findChild(QTableWidget, 'tbSummary')
        self.setMetadata(fileName[0], tbMetadata)

        # macro
        tbMacroDetails = self.window.findChild(QTableWidget, 'tbMacroDetails')
        tbMacroDecoded = self.window.findChild(QTextBrowser, 'tbMacroDecoded')
        self.setFoundMacros(fileName[0], tbMacroDetails, tbMacroDecoded)

        # handle VBA combobox also
        cbDecodedMac = self.window.findChild(QComboBox, 'cbDecodedMac')  # combobox of found macros
        tbInputVba = self.window.findChild(QTextBrowser, 'tbInputVba')
        self.populateFoundMacroComboBox(cbDecodedMac, tbInputVba)

        # add farm variables operation to VBA on document load; only if no file already loaded
        if self.docDependantIsLoaded == 0:
            tbAvailObf = self.window.findChild(QTableWidget, 'tbAvailObf')  # available operations
            tbPSAvailOps = self.window.findChild(QTableWidget, 'tbPSAvailOps')
            self.addDocumentDependentOps(tbAvailObf, tbPSAvailOps)

        # clear vba outputbox
        tbOutputVba = self.window.findChild(QTextBrowser, 'tbOutputVba')
        tbOutputVba.setPlainText("")

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
        self.foundMacros = []
        self.docDependantIsLoaded = 0

        # stores all vba deobf operations for adding and removing
        # [display name, func name, need document to function = 1]
        self.ALLVBADEOBFOPERATIONS = [
            ["Remove forced newlines", "removeForcedNewlines", 0],
            ["Replace Form Variables with value", "replaceFormVariablesWithValue", 1],
            ["Replace initialised variables with values", "replaceVariableWithValue", 0],
            ["Concatenate broken strings", "concatBrokenStrings", 0],
            ["Resolve add/subtract numbers", "resolveAddSubtractNumber", 0],
            ["Separate Functions with newlines", "separateFunctions", 0],
            ["Resolve Split()", "resolveSplit", 0],
            ["Resolve Trim()", "resolveTrim", 0],
            ["Resolve CVar()", "resolveCVar", 0],
            ["Resolve Join()", "resolveJoin", 0],
            ["Resolve Char()", "resolveChar", 0],
            ["Resolve Mid()", "resolveMid", 0],
            ["Remove junk code: Remove time wasting structures (Experimental)", "experimental1RemoveJunkCode", 0],
            ["Remove junk code: Remove x = y = z (Experimental)", "experimental2RemoveJunkCode", 0],
            ["Remove junk code: Remove number + number + number (Experimental)", "experimental4RemoveJunkCode", 0],
            ["Remove junk code: Remove all comments", "experimental3RemoveJunkCode", 0],
            ["Remove junk code: Remove Dim arrays (Experimental)", "experimental5RemoveJunkCode", 0],
            ["Remove junk code: Remove Unused Variables (Experimental)", "experimental6RemoveUnusedVariables", 0]
        ]

        # stores all vba deobf operations for adding and removing
        # [display name, func name, need document to function = 1]
        self.ALLPOWERSHELLDEOBFOPERATIONS = [
            ["Decode base64", "decodeBase64", 0],
            ["Create new lines", "createNewLines", 0],
            ["Remove unused variables", "removeUnusedVariables", 0],
            ["Remove useless escape characters", "removeUselessEscapeChars", 0],
            ["Remove surrounding brackets from strings", "removeBracketFromStrings", 0],
            ["Concatenate broken strings", "concatBrokenStrings", 0],
            ["Resolve char", "resolveChar", 0],
            ["Resolve split", "resolveSplit", 0],
            ["Resolve format operator", "resolveFormatOperator", 0],
            ["Resolve .replace() operator (.replace() )", "resolveReplace", 0],
            ["Resolve -Replace operator (-Replace/CReplace )", "resolveCReplace", 0],
            ["Extract value from array (e.g. [array]'val1','val2')[0] )", "extractValueFromArray", 0],
            ["Replace initialised variables with values", "replaceVariableWithValue", 0],
            ["Replace initialised variables with values (Aggressive)", "replaceVariableWithValueAggressive", 0]
        ]

        # populate deobf tables on startup
        # VBA deobfuscation page
        self.setVBADeobfPage()
        # PowerShell deobfuscation page
        self.setPowershellDeobfPage()

        ## catch the elements to change later
        actionOpen = self.window.findChild(QAction, 'actionOpen')  # menubar select file to open
        tbMacroDetails = self.window.findChild(QTableWidget, 'tbMacroDetails')  # table to change macro textbrowser
        btnExportMacros = self.window.findChild(QPushButton, 'btnExportMacros')  # button export macros to folder

        # VBA obf page
        cbDecodedMac = self.window.findChild(QComboBox, 'cbDecodedMac')  # combobox of found macros
        tbAvailObf = self.window.findChild(QTableWidget, 'tbAvailObf')  # available operations VBA
        tbCurrObf = self.window.findChild(QTableWidget, 'tbCurrObf')  # applied operations VBA
        btnApplyOps = self.window.findChild(QPushButton, 'btnApplyOps')  # button apply operations
        btnSaveOutput = self.window.findChild(QPushButton, 'btnSaveOutput')  # button save output to file
        btnOutputToInputVBA = self.window.findChild(QPushButton, 'btnOutputToInputVBA')  # button copy output to input
        btnSusPatterns = self.window.findChild(QPushButton, 'btnSusPatterns')  # button sus patterns

        # PowerShell obf page
        tbPSAvailOps = self.window.findChild(QTableWidget, 'tbPSAvailOps')  # available operations PowerShell
        tbPSCurrOps = self.window.findChild(QTableWidget, 'tbPSCurrOps')  # applied operations PowerShell
        btnApplyOpsPS = self.window.findChild(QPushButton, 'btnApplyOpsPS')  # button apply operations PowerShell
        btnSaveAsPS = self.window.findChild(QPushButton, 'btnSaveAsPS')  # button save output to file
        btnOutputToInputPS = self.window.findChild(QPushButton, 'btnOutputToInputPS')  # button copy output to input

        ## Connects all action button to their slots (functions)
        self.connect(actionOpen, SIGNAL("triggered()"), self, SLOT("openFile()"))
        tbMacroDetails.cellClicked.connect(self.changeMacroTextbox)
        btnExportMacros.clicked.connect(self.exportMacrosToFolder)

        cbDecodedMac.currentIndexChanged.connect(self.updateInputVBATextBrowser)
        tbAvailObf.cellClicked.connect(self.addVBAOperation)
        tbCurrObf.cellClicked.connect(self.actionVBAOperation)
        btnApplyOps.clicked.connect(self.applyVBAOps)
        btnSaveOutput.clicked.connect(self.saveOutputVBA)
        btnOutputToInputVBA.clicked.connect(self.copyOutputToInputVBA)
        btnSusPatterns.clicked.connect(self.displayVBAHints)

        tbPSAvailOps.cellClicked.connect(self.addPSOperation)
        tbPSCurrOps.cellClicked.connect(self.actionPSOperation)
        btnApplyOpsPS.clicked.connect(self.applyPSOps)
        btnSaveAsPS.clicked.connect(self.saveOutputPS)
        btnOutputToInputPS.clicked.connect(self.copyOutputToInputPS)

        self.window.show()  # Required


if __name__ == '__main__':
    app = QApplication(sys.argv)
    form = Form('main.ui')
    sys.exit(app.exec_())
