import re
from subprocess import Popen, PIPE
import os


# REMOVE FORCED NEWLINES ---------------------------------------------------------------------------------------
def removeForcedNewlines(originalText):
    # space underscore whitespace or comment
    regex = r'\s_[\s\']+'

    originalText = re.sub(regex, r'', originalText)
    return originalText


def getLocationWdkeyList():
    ROOT_DIR = os.path.dirname(os.path.abspath(__file__))  # This is the Project Root
    ASCII_PATH = os.path.join(ROOT_DIR, 'Imports/wdKey.txt')  # TODO CHANGE LATER IN FINAL PROJECT
    return ASCII_PATH


def getLocationAsciiList():
    ROOT_DIR = os.path.dirname(os.path.abspath(__file__))  # This is the Project Root
    ASCII_PATH = os.path.join(ROOT_DIR, 'Imports/asciiTable.txt')  # TODO CHANGE LATER IN FINAL PROJECT
    return ASCII_PATH


# RESOLVE CHR, CHRW, CHRB WITH KEY VALUES (WDKEYS) ---------------------------------------------------------------------
def resolveChar(originalText):
    wdKeyPath = getLocationWdkeyList()
    wdkeyValues = []

    # ascii and wdkeys to be used later
    # list of wdkey values
    with open(wdKeyPath, "r") as myfile:
        for line in myfile:
            keyValuePair = line.strip().split(":")
            wdkeyValues.append(keyValuePair)

    # list of ascii values
    asciiPath = getLocationAsciiList()
    asciiTableValues = []

    # list of string values
    with open(asciiPath, "r") as myfile:
        for line in myfile:
            decValuePair = line.strip().split(":")
            # space is not stored in the txt file. Add space here
            if decValuePair[0] == "32":
                decValuePair[1] = " "
            asciiTableValues.append(decValuePair)

    # catch ChrW(wdKeyP); group 1 is all, group 2 is value inside ()
    regex = r'(chr[WB]?\s*\(\s*(.*?)\s*\))'

    listOfChar = re.findall(regex, originalText, flags=re.I)
    for tuple in listOfChar:
        charFunc = tuple[0]
        charValue = tuple[1]

        # see if value is a number
        if charValue.isnumeric():
            # asciiValue will be the char
            asciiValue = ""
            # find value of dec from list
            for asciiLine in asciiTableValues:
                # compare found value with asciiTable
                if asciiLine[0] == charValue:
                    # save value
                    asciiValue = asciiLine[1]
                    break

            # sub value with the ascii char
            originalText = originalText.replace(charFunc, "\"" + asciiValue + "\"")

        # if value is keystroke
        else:
            wdKeyValue = ""
            # change wdkey to value
            for wdKeyPair in wdkeyValues:
                # found wdKey value
                if charValue == wdKeyPair[0]:
                    wdKeyValue = wdKeyPair[1]
                    break

            # change wdKeyValue to char
            asciiValue = ""
            # find value of dec from list
            for asciiLine in asciiTableValues:
                # compare found value with asciiTable
                if asciiLine[0] == wdKeyValue:
                    # save value
                    asciiValue = asciiLine[1]
                    break

            # sub value with the ascii char
            originalText = originalText.replace(charFunc, "\"" + asciiValue + "\"")

    return originalText


# SAVE VBA FORM VARIABLES TO LIST
# START FORM VARIABLES ---------------------------------------------------------------------------------------
def getLocationOlevba():
    ROOT_DIR = os.path.dirname(os.path.abspath(__file__))  # This is the Project Root
    OLEVBA_PATH = os.path.join(ROOT_DIR, 'Imports\\Olevba\\olevba.exe')  # requires `import os`
    return OLEVBA_PATH


def callOleVba(filename):
    OLEVBA_PATH = getLocationOlevba()

    p = Popen([OLEVBA_PATH,
               filename], stdin=PIPE, stdout=PIPE, stderr=PIPE)
    output, err = p.communicate()
    rc = p.returncode  # optional?
    return output  # returns in bytes


# testing extraction of vba form variables
def getVbaFormVariables(filename):
    output = callOleVba(filename)

    encList = ['ascii',
               'UTF-8',
               'utf_8_sig',
               'windows-1252',
               'latin_1']

    outputDecoded = ""

    for encoding in encList:
        try:
            outputDecoded = output.decode(encoding=encoding)
            break
        except:
            continue

    regex = r'VBA FORM Variable \"(.*)\" IN \'.*\' - OLE stream: u\'(.*)\'\s*\n- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -\s*\n(.*)'
    # [variable, parent, value]
    # want to make [parent.variable, value]
    result = re.findall(regex, outputDecoded)

    resolvedVariableValue = []
    for tuple in result:
        variable = tuple[0]
        # macro/parent; get the parent
        parent = tuple[1].rsplit('/', 1)[1]
        combined = parent + "." + variable
        # if value is empty
        if tuple[2] == "-------------------------------------------------------------------------------":
            value = ""
        else:
            value = tuple[2]

        variableValuePair = [combined, value]
        resolvedVariableValue.append(variableValuePair)

    return resolvedVariableValue


# UTILITY Escape characters to be used in python regex
def escapePythonRegex(stringToEscape):
    # escape variable
    charToEsc = ['\\', '+', '-', '.', ',', '(', ')', '{', '}', '[', ']', '<', '>', '$', '\"', '*', '?', '!', '=',
                 ':',
                 '|', '#', '^']
    # craft regex to replace
    escapedString = stringToEscape
    for char in charToEsc:
        escapedString = escapedString.replace(char, '\\' + char)

    return escapedString


# sub form variables
def replaceFormVariablesWithValue(filename, originalText):
    resolvedVariableValue = getVbaFormVariables(filename)

    # replace all occurances of form variables in the originalText
    for item in resolvedVariableValue:
        variable = item[0]
        value = item[1]

        # escape variable
        variableEscaped = escapePythonRegex(variable)

        regexPerVariable = "(" + variableEscaped + ")" + "(\\s)"

        # Dtcqcidgf.Xrqtcysa.GroupName not substituted
        # may need to regex perfect match; no extra .anything at the back
        # want to sub group one only; ignore any space or newlines after the variable
        toreplace = "\"" + value + "\"\\g<2>"
        originalText = re.sub(regexPerVariable, toreplace, originalText)

    return originalText


# END FORM VARIABLES ---------------------------------------------------------------------------------------

# concat broken strings
def concatBrokenStrings(originalText):
    # catch 'string1'+'string2"
    regex = r"['\"](.*?)['\"]\s*\+\s*['\"](.*?)['\"]"

    match = re.search(regex, originalText)
    # loop until no more broken strings found
    while match:
        originalText = re.sub(regex, r'"\g<1>\g<2>"', originalText)
        match = re.search(regex, originalText)

    return originalText


# resolve raw math
def resolveAddSubtractNumber(originalText):
    # find number +/- number
    regex = r'((\d+)\s*([+-])\s*(\d+))'
    match = re.search(regex, originalText)
    # loop until no more match found
    while match:
        # add/sub then replace
        fullStatement = match.group(1)
        firstNum = int(match.group(2))
        operator = match.group(3)
        secondNum = int(match.group(4))

        if operator == "+":
            resultNum = firstNum + secondNum
        if operator == "-":
            resultNum = firstNum - secondNum

        originalText = originalText.replace(fullStatement, str(resultNum))
        match = re.search(regex, originalText)

    return originalText


# CAPTURE ALL INITIALISED VARIABLES AND VALUES
def getInitialisedVariablesAndValues(originalText):
    regex = r'\s(([^\.\s][\w]*)\s*=\s*(.*))'
    # grp1: whole line, grp2: variableName, grp3: whole value
    match = re.findall(regex, originalText)

    return match

# SUB IN VALUES IF VARIABLE IS STR/INT(?) ="string" or =digitOnly
# Aggressive
# This one can be slow ~ 40 sec in some cases
# TODO edit regex a bit to include toReplace.create(...)
# TODO Refer to Fcbxktofsye.Create(....)
# TODO REMEMBER TO BACKUP CURRENT REGEX FIRST
def replaceVariableWithValue(originalText):
    initialisedVariablesAndValues = getInitialisedVariablesAndValues(originalText)

    for tuple in initialisedVariablesAndValues:
        variable = tuple[1]

        # loop through each variable; regrab value as it will change as it gets replaced
        # The regrab value part
        # escape the variable for regex
        escapedVariable = escapePythonRegex(variable)

        # search for updated value
        valueRegex = "\\s((" + escapedVariable + ")\\s*=\\s*(.*))"
        valueMatch = re.search(valueRegex, originalText)
        # whole statement grp 1, variable: 2, value: 3
        value = valueMatch.group(3)
        escapedValue = escapePythonRegex(value)

        # replace variables with updated values
        toReplaceRegex = "([\\s(]+)(" + escapedVariable + ")([\\s),.]+)" + "(?!\\s*?=\\s*?" + escapedValue + ")"
        subRegex = "\\g<1>" + value + "\\g<3>"

        originalText = re.sub(toReplaceRegex, subRegex, originalText)

    return originalText


# RESOLVE SPLIT; ignores if there is extra variables
def resolveSplit(originalText):
    regex = r'(Split\s*\((.*?)[\"\'](.*?)[\"\']\s*(.*?),\s*[\"\'](.*?)[\"\']\s*\))'
    match = re.findall(regex, originalText)

    for tuple in match:
        wholeStatement = tuple[0]
        toSplit = tuple[2]
        delimiter = tuple[4]

        # check if there are extra variables in the split() function
        beforeExtra = tuple[1].strip()
        afterExtra = tuple[3].strip()

        # if no extra variables in split() function
        if beforeExtra == "" and afterExtra == "":
            # split the function and replace
            splittedString = toSplit.split(delimiter)
            # remove empty values
            splittedString = [x for x in splittedString if x]

            # build vba array syntax
            listLen = len(splittedString)
            vbaArray = "Array("
            for index, item in enumerate(splittedString):
                # add item to the array
                vbaArray += "\"" + item + "\""
                # add comma if not last item
                if index < listLen - 1:
                    vbaArray += ", "
            vbaArray += ")"

            # replace original split() with vbaArray
            originalText = originalText.replace(wholeStatement, vbaArray)

    return originalText


# resolve trim()
def resolveTrim(originalText):
    regex = r'(([LR]?)Trim\s*\(\s*(.*?)[\'\"]([\s\S]*?)[\'\"](.*?)\))'
    # group 3 and 5 must be empty; extra variables check
    match = re.findall(regex, originalText)

    for tuple in match:
        # check if group 2 and 4 are empty
        fullStatement = tuple[0]
        beforeCheck = tuple[2].strip()
        afterCheck = tuple[4].strip()

        if beforeCheck == "" and afterCheck == "":
            toTrim = tuple[3]

            # check which trim: L, R or both
            direction = tuple[1]
            if direction == "L":
                toTrim = toTrim.lstrip()
            elif direction == "R":
                toTrim = toTrim.rstrip()
            else:
                toTrim = toTrim.strip()

            # replace original statement
            originalText = originalText.replace(fullStatement, "\"" + toTrim + "\"")

    return originalText


# resolve cvar()
# Known issue: if there is + extra variable in the function, will wrongly catch. User must manually clean
def resolveCVar(originalText):
    # detect of string or int
    # if string add "", if int no need to add ""; remove CVar() in the end
    # see if start and end with "" or '';
    # if no "" and not number means variable. Ignore this case
    regex = r'(CVar\s*\(\s*([\'\"]?[\s\S]*?[\'\"]?)\s*\))'
    match = re.findall(regex, originalText)

    for tuple in match:
        fullStatement = tuple[0]
        toConvert = tuple[1]

        # case 1: "This is a String"
        if (toConvert.startswith("\"") or toConvert.startswith("\'")) and (
                toConvert.endswith("\"") or toConvert.endswith("\'")):
            # If string replace with ""
            originalText = originalText.replace(fullStatement, toConvert)

        else:
            # case 2: 500 number
            if toConvert.isdecimal():
                # If number replace without ""
                originalText = originalText.replace(fullStatement, toConvert)

            # case 3: HelloVariable
            else:
                # If variable don't do anything
                continue

    return originalText


# resolve Join()
def resolveJoin(originalText):
    # search for Join(Array(), delimeter)
    regex = r'(Join\s*\(\s*Array\(((?:[\'\"].*?[\'\"]\s*,?\s*)+)\s*\)\s*,\s*([\'\"].*[\'\"]\s*)\))'
    match = re.findall(regex, originalText)

    # grp1: full statement, 2: toJoin, 3: delimeter
    for tuple in match:
        fullStatement = tuple[0]
        toJoin = tuple[1]

        # strip delimeter of "" or '' to get value (although usually it's empty
        delimeter = tuple[2]
        delimeter = delimeter.strip()
        delimeter = delimeter.strip("\"")
        delimeter = delimeter.strip("\'")

        # split toJoin then join, mind the ""
        splitList = toJoin.split(",")

        # add enumerate to change original list
        for index, item in enumerate(splitList):
            # remove start and trail "" or ''
            item = item.strip()
            item = item.strip("\"")
            strippedItem = item.strip("\'")
            splitList[index] = strippedItem

        # USE THE DELIMETER, USE DELIMITER.JOIN(splitList)
        toJoinResult = delimeter.join(splitList)

        # now replace original statement with new result
        originalText = originalText.replace(fullStatement, "\"" + toJoinResult + "\"")

    return originalText


# vba Mid()
# TODO VERIFY CORRENT ANSWER
def resolveMid(originalText):
    # find Mid("zahajUZomiAjVADEAMAA4ADMdpokrTZ", 14, 11)
    regex = r'(Mid\(\"([\s\S]*?)\"\s*,\s*(\d*)\s*,\s*(\d*)\))'
    match = re.findall(regex, originalText)

    for tuple in match:
        fullStatement = tuple[0]
        toMid = tuple[1]
        startIndex = int(tuple[2]) - 1  # VBA mid func starts with 1
        length = int(tuple[3])

        afterMid = toMid[startIndex:startIndex + (length)]

        # now replace original statement with new result
        originalText = originalText.replace(fullStatement, "\"" + afterMid + "\"")

    return originalText


# capture olevba hints ---------------------------------------------------------------------------------------
def getOlevbaHints(filename):
    oleHints = ""

    output = callOleVba(filename)

    encList = ['ascii',
               'UTF-8',
               'utf_8_sig',
               'windows-1252',
               'latin_1']

    outputDecoded = "No hints found."

    for encoding in encList:
        try:
            outputDecoded = output.decode(encoding=encoding)
            break
        except:
            continue

    regex = r'(\+-[\s\S]*Description[\s\S]*-\+)'
    result = re.search(regex, outputDecoded)
    if result == None:
        return outputDecoded

    oleHints = result.group(1)

    return oleHints  # returns whole box as string


# find which macro file contains the autoexec keywords
def locateAutoExec(olevbaHintsresult, foundMacros):
    autoExecLocations = []
    regex = r'AutoExec\s*\|\s*(\w*)\s*'
    result = re.findall(regex, olevbaHintsresult)

    # result may have more than one autoexec keyword; but usually there will only be 1
    for keyword in result:
        # item is a string (autoexec name)
        # search for string in all found macros [3]: macro name, [4] macro string
        for macroList in foundMacros:
            if keyword in macroList[4]:
                newLocation = [keyword, macroList[3]]
                autoExecLocations.append(newLocation)

    return autoExecLocations


# UTILITY MAYBE USEFUL
# CAPTURE FUNCTION NAMES ONLY
def getFunctionNames(originalText):
    regex = r'Function +(.*) *\(\)'
    match = re.findall(regex, originalText)

    return match


#  CAPTURE ALL FUNCTIONS AND CONTENTS
def getFunctionsAndContents(originalText):
    regex = r'(Function +(.*) *\(\)[\s\S]*?\s*End Function)'
    match = re.findall(regex, originalText)  # grp1: whole func, grp2: func name

    return match


# CAPTURE ALL INITIALISED VARIABLES NAMES ONLY
def getInitialisedVariables(originalText):
    regex = r'\s([^\.\s][\w]*)(?=\s+\=)'
    match = re.findall(regex, originalText)

    return match


# CAPTURE ALL LOOPS
def getAllLoops(originalText):
    regex = r'\s*(For\s*[\s\S]*?Next)'  # catch For/Next loop; can't catch nested
    match = re.findall(regex, originalText)

    return match


# CAPTURE ALL IFS
def getAllIfs(originalText):
    regex = r'\s*(If\s*[\s\S]*?End If)'  # catch If/End If; can't catch nested
    match = re.findall(regex, originalText)

    return match


# utility for experimental1RemoveJunkCode
def removeSusStruc(strucStr, originalText):
    # sus is determined by (no of sus func / no of lines) > 80%?
    # suspicious math "time wasting" functions
    timeWastingFunc = ["*", "-", "<", ">", "/", "CBool", "CByte", "CCur", "CDate", "CDbl", "CDec", "CInt", "CLng",
                       "CLngLng", "CLngPtr", "CSng", "CStr", "CVar", "Abs", "Atn", "Cos", "Exp", "Log", "Rnd", "Sgn",
                       "Sin", "Sqr", "Tan", "Acos", "Asin", "Atan", "Atan2", "BigMul", "Ceiling", "Cosh", "DivRem",
                       "Floor", "Log10", "Max", "Min", "Pow", "Round", "Sign", "Sinh", "Sqrt", "Tanh", "Truncate"]
    # number of lines in structure
    noOfLines = strucStr.count('\n') + 1

    # total "time wasting" functions in structure
    mathCounter = 0
    for func in timeWastingFunc:
        mathCounter += strucStr.count(func)

    # if more than 80% time wasting, remove
    if (mathCounter / noOfLines) > 0.8:
        originalText = originalText.replace(strucStr, "")

    return originalText


# separate functions with newlines (Readability)
def separateFunctions(originalText):
    regex = r'(Function[\s\S]*?End Function)'
    originalText = re.sub(regex, "\\n\\n\\g<1>", originalText)

    return originalText


# remove "time wasting" structures
def experimental1RemoveJunkCode(originalText):
    # find all loops then remove sus loops; then if statements
    # sus is determined by (no of sus func / no of lines) > 80%
    matchedLoops = getAllLoops(originalText)
    for item in matchedLoops:
        # item is the entire loop as a string
        originalText = removeSusStruc(item, originalText)

    matchedIfs = getAllIfs(originalText)
    for item in matchedIfs:
        # item is the entire loop as a string
        originalText = removeSusStruc(item, originalText)

    # cleanup; remove empty lines
    lines = originalText.split("\n")
    non_empty_lines = [line for line in lines if line.strip() != ""]
    string_without_empty_lines = ""
    for line in non_empty_lines:
        string_without_empty_lines += line + "\n"

    return string_without_empty_lines


# remove x = y = z
def experimental2RemoveJunkCode(originalText):
    # split into individual line first (regex limitation workaround)
    splittedText = originalText.splitlines()

    # find x = y = z in each line
    regex = r'[\s\S]*?\s*=\s*[\s\S]*?\s*=\s*[\s\S]*?$'
    for line in splittedText:
        match = re.search(regex, line)
        if match:
            originalText = originalText.replace(line, "")

    # cleanup; remove empty lines
    lines = originalText.split("\n")
    non_empty_lines = [line for line in lines if line.strip() != ""]
    string_without_empty_lines = ""
    for line in non_empty_lines:
        string_without_empty_lines += line + "\n"

    return string_without_empty_lines


# remove comments
def experimental3RemoveJunkCode(originalText):
    regex = r"(^[\s\S]*?)'[\s\S]*?$"
    originalText = re.sub(regex, r"\g<1>", originalText, flags=re.MULTILINE)

    # cleanup; remove empty lines
    lines = originalText.split("\n")
    non_empty_lines = [line for line in lines if line.strip() != ""]
    string_without_empty_lines = ""
    for line in non_empty_lines:
        string_without_empty_lines += line + "\n"

    return string_without_empty_lines


# remove lines with number + number + number
def experimental4RemoveJunkCode(originalText):
    # split into individual line first (regex limitation workaround)
    splittedText = originalText.splitlines()

    regex = r'\d+\s[+-]\s*\d+\s*[+-]\s*\d+'
    for line in splittedText:
        match = re.search(regex, line)
        if match:
            originalText = originalText.replace(line, "")

    # cleanup; remove empty lines
    lines = originalText.split("\n")
    non_empty_lines = [line for line in lines if line.strip() != ""]
    string_without_empty_lines = ""
    for line in non_empty_lines:
        string_without_empty_lines += line + "\n"

    return string_without_empty_lines


# remove Dim(array) that is unused
# Note very specific to one type of sample.
# Assumes Dim/ReDim is not used in any useful part of code.
def experimental5RemoveJunkCode(originalText):
    # split into individual line first (regex limitation workaround)
    splittedText = originalText.splitlines()

    # group 1 contains variable to search line and remove
    regex = r'Dim\s+(\w*)\s*\(\s*\d*\s*\)'
    listOfVariables = re.findall(regex, originalText)

    # remove duplicates
    listOfVariables = list(dict.fromkeys(listOfVariables))

    # Every line that contains variable, remove
    for line in splittedText:
        for variable in listOfVariables:
            if variable in line:
                originalText = originalText.replace(line, "")
                break

    # cleanup; remove empty lines
    lines = originalText.split("\n")
    non_empty_lines = [line for line in lines if line.strip() != ""]
    string_without_empty_lines = ""
    for line in non_empty_lines:
        string_without_empty_lines += line + "\n"

    return string_without_empty_lines


# remove unused initialised variables (Experimental)
def experimental6RemoveUnusedVariables(originalText):
    # split into individual line first (regex limitation workaround)
    splittedText = originalText.splitlines()

    # find variable = anything
    regex = r'(\w+?)\s=\s[\s\S]*?'
    for line in splittedText:
        match = re.search(regex, line)
        if match:
            variable = match.group(1)
            variableStripped = variable.strip()
            # find instances of variable in originalText
            noOfInst = originalText.count(variableStripped)

            # if appear once means unused; remove
            if noOfInst == 1:
                originalText = originalText.replace(line, "")

    # cleanup; remove empty lines
    lines = originalText.split("\n")
    non_empty_lines = [line for line in lines if line.strip() != ""]
    string_without_empty_lines = ""
    for line in non_empty_lines:
        string_without_empty_lines += line + "\n"

    return string_without_empty_lines

# TODO note that "" and 0 are same. So showwindow = "" also works

# import invokeTools

# path = "C:\\Users\\user\\Desktop\\Deobsfuscation\\Test.txt"
# working = ""
# with open(path, "r") as myfile:
#     working = myfile.read()
#
# working = experimental6RemoveUnusedVariables(working)
# print(working)
# print("---------------------------------------------------------\n")


# print("Before:")
# print(working)
# print("---------------------------------------------------------\n")


# foundMacros = invokeTools.oledumpMacro(path)
#
# oleVbaOutput = getOlevbaHints(path)
# listOfAutoExec = locateAutoExec(oleVbaOutput, foundMacros)
#
# # import time
#
# start_time = time.time()
#
# # print("Before:")
# # print(working)
# # print("---------------------------------------------------------")
#
# working = experimental1RemoveJunkCode(working)
# print(working)
# print("---------------------------------------------------------")


# filename = "C:\\Users\\user\\Desktop\\ST_28546448"
# path = "C:\\Users\\user\\Desktop\\Deobsfuscation\\rawvbanojunkremovedunusedvar.txt"
# working = ""
# with open(path, "r") as myfile:
#     working = myfile.read()
#
# import time
#
# start_time = time.time()
#
# print("Before:")
# print(working)
# print("---------------------------------------------------------")
#
# working = removeForcedNewlines(working)
# working = resolveChar(working)
# working = replaceFormVariablesWithValue(filename, working)
# working = concatBrokenStrings(working)
# working = replaceVariableWithValue(working)
# working = resolveTrim(working)
# working = resolveCVar(working)
# working = concatBrokenStrings(working)
# working = resolveSplit(working)
# working = resolveJoin(working)
# working = concatBrokenStrings(working)
#
# print("After:")
# print(working)
# print("\n--- %s seconds ---" % (time.time() - start_time))
