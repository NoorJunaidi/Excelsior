import re
import base64
import os


# working
def createNewLines(originalText):
    # if ;notNewLine replace with ;\n
    replacement1 = r";\n"  # replace ;notnewline with ;newline
    ans = re.sub(r';\s*', replacement1, originalText)  # regex to find ; to create new line

    replacement2 = r"\g<1>\n{\n"  # replace {
    ans = re.sub(r'([^\s\$]){(?!\d+})(?=)', replacement2,
                 ans)  # regex to find { that is not newline, variable or format

    replacement3 = r"\g<1>\n}\n\g<2>"  # replace }
    ans = re.sub(r'([\W])}(.)', replacement3, ans)  # regex to find } to create new line; ignores {num} format

    return ans


def getAllVariables(originalText):
    # catch $variable = anything;
    # old reges: (\$[\w{}]+)\s*=
    listOfVariablesAndValues = re.findall(r'(\$[\w{}]+)\s*=([\s\S]*?)[;\n]',
                                          originalText)  # regex to find all $variable and ${variable}

    return listOfVariablesAndValues


def removeUnusedVariables(originalText):
    listOfVariablesAndValues = getAllVariables(originalText)

    # if variable doesnt appear again, delete
    toRemove = []  # the variables to be removed
    linesInText = originalText.splitlines()

    # find all variables to be removed; appear only once
    # powershell is not case sensitive; need to lower originalText
    originalTextLowered = originalText.lower()

    for variable in listOfVariablesAndValues:
        numOfOccurances = originalTextLowered.count(variable[0].lower())

        if numOfOccurances == 1:
            toRemove.append(variable)

    # TODO replace using regex not line-by-line

    # TODO need to find better way
    # TODO PROBLEM WITH BELOW: CANT DETECT MULTILINE VARIABLES AND VARAIBLES WITHOUT ; AT THE SAME TIME
    # (\$[\w{}]+)\s*=[\s\S]*?[;\n]

    # match line with created regex, then remove
    # regex is to find variable=initialize; new regex created for each variable
    for index, line in enumerate(linesInText):
        for variable in toRemove:
            regex = '\\' + variable[0] + '[^}]+'

            match = re.search(regex, line)
            if match:
                # remove the line
                lineToReplace = match.group(0)
                lineResult = line.replace(lineToReplace, "")
                linesInText[index] = lineResult

    # rejoin all line and return
    joinedOutput = ""
    for line in linesInText:
        if line != "":
            joinedOutput += line + "\n"

    return joinedOutput


# remove useless escape characters
def removeUselessEscapeChars(originalText):
    # searches for all escapes that has no meaning in Powershell
    regex = r"(`([^0abfnrtv`#'\"]))"  # returns 2 groups; (whole escape, just the char)

    match = re.findall(regex, originalText)

    for tuple in match:
        lineToReplace = tuple[0]
        replacementLine = tuple[1]
        originalText = originalText.replace(lineToReplace, replacementLine)

    return originalText


# search and decode base64
def decodeBase64(originalText):
    # regex looks for base64 string
    regex = r"([A-Za-z0-9+\/]{4}){3,}([A-Za-z0-9+\/]{2}==|[A-Za-z0-9+\/]{3}=)?"
    match = re.search(regex, originalText)

    if match:
        toBeDecoded = match.group(0)
        decoded = base64.b64decode(toBeDecoded)

        try:
            decoded = decoded.decode("utf-16")
            originalText = originalText.replace(toBeDecoded, decoded)

        except:
            return originalText

    return originalText


# concat random broken strings
def concatBrokenStrings(originalText):
    # catch 'string1'+'string2"
    regex = r"['\"](.*?)['\"]\s*\+\s*['\"](.*?)['\"]"

    match = re.search(regex, originalText)
    # loop until no more broken strings found
    while match:
        originalText = re.sub(regex, r"'\g<1>\g<2>'", originalText)
        match = re.search(regex, originalText)

    return originalText


#  RESOLVE [CHAR]
def getLocationAsciiList():
    ROOT_DIR = os.path.dirname(os.path.abspath(__file__))  # This is the Project Root
    ASCII_PATH = os.path.join(ROOT_DIR, 'Imports/asciiTable.txt')  # TODO CHANGE LATER IN FINAL PROJECT
    return ASCII_PATH


def resolveChar(originalText):
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

    # detect [char]42 statement, accounts for spacing in between; group1: [char]42, 2:42
    regex = r"(\[\s*?char\s*?\]\s*?\(?(\d+)?)"
    listOfChar = re.findall(regex, originalText, flags=re.IGNORECASE)  # non case sensitive

    # tuple contains ([char]42, 42] format
    for tuple in listOfChar:
        toReplace = tuple[0]
        charNumber = tuple[1]
        asciiValue = ""

        # find value of dec from list
        for asciiLine in asciiTableValues:
            # compare found value with asciiTable
            if asciiLine[0] == charNumber:
                # save value
                asciiValue = asciiLine[1]
                break

        replacementStr = "\"" + str(asciiValue) + "\""
        originalText = originalText.replace(toReplace, replacementStr)
        # originalText = re.sub(regex, "\'" + str(asciiValue) + "\'", originalText)

    return originalText


# RESOLVE SPLIT() W/ DELIMITER
# TODO sometimes split resolves but does not actually split; will still create "list"
def resolveSplit(originalText):
    # First capture string and delimiter
    # search for "" quotes.split(); group 1 long string to split, group 2 delimiter including '' if have
    regex = r"([\"'](.*)[\"']\.[\"']?split[\"']?.*\((.*)\))"
    listToSplit = re.findall(regex, originalText, re.IGNORECASE)

    # for each element to split, get string and delimiter, then split, then replace
    for tuple in listToSplit:
        stringToSplit = tuple[1]
        delimiter = tuple[2]

        # strip starting and trailing quotes
        delimiter = delimiter.strip("\'\"")
        # print(delimiter)

        # # OLD remove start and trailing quotations from delimiter
        # if (delimiter.startswith('\'') and delimiter.endswith('\'')) or (
        #         delimiter.startswith('\"') and delimiter.endswith('\"')):
        #     delimiter = delimiter[1:-1]

        # split the string into a list
        splittedList = stringToSplit.split(delimiter)

        # construct list in powershell format
        powershellList = "@("
        list_len = len(splittedList) - 1  # to find last element, different output
        for index, item in enumerate(splittedList):
            powershellList += "\"" + item + "\""
            # if not last member add comma ,
            if index != list_len:
                powershellList += ","
            # if last member, add bracket )
            else:
                powershellList += ")"

        # sub back in original text; replace old string with formatted array
        originalText = originalText.replace(tuple[0], powershellList)

    return originalText


# SUB IN VALUES IF VARIABLE IS STR/INT(?) ="string" or =digitOnly
# CONSERVATIVE VERSION; ONLY DO INITIALISED STR AND INT
# This one does not include ${variable} format
# TODO UPDATE REGEX WITH ${VAR}; SEE GETALLVARIABLES() maybe
def replaceVariableWithValue(originalText):
    # detects variable and value; value may be string or int (returns 2 groups, var and val)
    # OLD regex = r"(\$.*?)\s*?=\s*?(['\"].*?['\"]|\d*)\s*?;"
    regex = r"(\$.*?)\s*?=\s*?(['\"][^\'\"]*?['\"])\s*?;"
    match = re.findall(regex, originalText)

    # for every variable value found, replace the text with value
    for tuple in match:
        variable = tuple[0]
        value = tuple[1].strip()

        # regex to find all other variables except the original which initialized it; craft new regex for each
        # uses negative lookahead ?!
        # \$Nnyjthcrzjoyv(?!\s*?=\s*?'937') example
        regex2 = "\\" + variable + "(?!\\s*?=\\s*?" + value + ")"
        originalText = re.sub(regex2, value, originalText)

    return originalText


# AGGRESSIVE VARIABLE REPLACEMENT. ALL-IN VERSION OF ABOVE ^
# AGGRESSIVE VERSION; TAKES ALL
# UPDATE REGEX WITH ${VAR}; SEE GETALLVARIABLES()
def replaceVariableWithValueAggressive(originalText):
    charToEsc = ['\\', '+', '-', '.', ',', '(', ')', '{', '}', '[', ']', '<', '>', '$', '\"', '*', '?', '!', '=', ':',
                 '|', '#', '^']

    # detects variable and value; value may be string or int (returns 2 groups, var and val)
    regex = r"(\$[\w{}]+)\s*=([\s\S]*?)[;\n]"
    match = re.findall(regex, originalText)

    # for every variable value found, replace the text with value
    # AFTER EVERY REPLACE, NEED TO GET NEW VALUE OF VARIABLE
    # BECAUSE VARIABLES WOULD HAVE BEEN SUBBED TOO
    for tuple in match:
        variable = tuple[0]

        # escape the variable to generate new regex to grab updated value
        variableEscaped = re.escape(variable)
        # variableEscaped = variable
        # for char in charToEsc:
        #     variableEscaped = variableEscaped.replace(char, '\\' + char)

        # RESCAN THE VARIABLE EVERYTIME
        # this format: (\${RANDOM})\s*=([\s\S]*?)[;\n]; this will regrab the specific updated val of the var
        regexNewValue = "(" + variableEscaped + ")\\s*=([\\s\\S]*?)[;\\n]"
        # print("regexNewValue")
        # print(regexNewValue)
        # regex here
        regrabResult = re.search(regexNewValue, originalText, flags=re.I)
        value = regrabResult.group(2).strip()

        # regex to find all other variables except the original which initialized it; craft new regex for each
        # uses negative lookahead: ?!
        # \$Nnyjthcrzjoyv(?!\s*?=\s*?'937') example
        # \ must be escaped first vvv
        # ONLY BACKSLASH MUST BE ESCAPED FOR REPLACEMENT STRING: https://docs.python.org/3/library/re.html#re.escape
        valueEscaped = re.escape(value)
        valueEscapedRepl = value.replace("\\", r"\\")

        regex2 = variableEscaped + "(?!\\s*?=\\s*?" + valueEscaped + ")"

        # print("variableEscaped:")
        # print(variableEscaped)
        # print("Regex2:")
        # print(regex2)
        # print("valueEscaped:")
        # print(valueEscaped)
        # print("valueEscapedRepl:")
        # print(valueEscapedRepl)
        # print("\n---------------------------------------------------------------")

        originalText = re.sub(regex2, valueEscapedRepl, originalText, flags=re.I)  # i = ignorecase

    return originalText


# FORMAT STRINGS ("{0}{1}" -f v1, v2)
# CONCAT TO "V1V2"

def resolveFormatOperatorOLD(originalText):
    # group group 1: all brackets, group 3: all values; maybe need to trim
    regex = r'([\"\']\s*((\s*{\s*\d+\s*})+\s*)\s*[\"\']\s*-f\s*(([\"\'].*?[\"\']\s*,?\s*)+))'
    match = re.findall(regex, originalText, flags=re.IGNORECASE)

    for tuple in match:
        # grp 0: whole string to replace including ""
        # grp 1: {} pattern
        # grp 3: values

        # get the {} order as list; from grp 1
        regexOrder = r'\s*\{\s*(\d+)\s*\}\s*'
        matchOrder = re.findall(regexOrder, tuple[1])  # is list of order

        # split grp 3 to values
        regexValues = r'["\']+(.*?)["\']+'  # split by 'val1', 'val2'
        values = re.findall(regexValues, tuple[3])

        # reorder values with matchOrder
        reorderedValues = [values[int(i)] for i in matchOrder]
        reorderedString = ""
        reorderedString = reorderedString.join(reorderedValues)

        # replace old vals (grp1) with new rearranged val (reorderedString)
        originalText = originalText.replace(tuple[0], "\"" + reorderedString + "\"")

    return originalText


# improved resolve format operator
def resolveFormatOperator(originalText):
    regex = r'([\"\']([^\"\']+?)[\"\']\s*\)?\s*-f\s*((?:[\"\'].*?[\"\'],?)+))'
    match = re.findall(regex, originalText, flags=re.IGNORECASE)

    for tuple in match:
        fullStatement = tuple[0]
        toFormat = tuple[1]
        toSubIn = tuple[2]

        # split toSubIn into elements
        splittedToSubIn = toSubIn.split(",")

        # strip quotes
        for index, item in enumerate(splittedToSubIn):
            itemNew = item.strip("\'\"")
            splittedToSubIn[index] = itemNew

        # swap in?
        newFullStatement = toFormat  # to change
        regexBrackets = r'({\s*(\d+)\s*})'
        match2 = re.findall(regexBrackets, toFormat)
        for bracket in match2:
            subInIndex = int(bracket[1])
            newFullStatement = newFullStatement.replace(bracket[0], splittedToSubIn[subInIndex])

        # replace statement into original text
        originalText = originalText.replace(fullStatement, "\"" + newFullStatement + "\"")

    return originalText


# Remove useless brackets around initialised strings OLD
def removeBracketFromStringsOLD(originalText):
    regex = r'([=+(])(\s*\(\s*([\'\"][\s\S]*?[\'\"])\s*\))'
    originalText = re.sub(regex, "\\g<1>\\g<3>", originalText)

    return originalText


# Remove useless brackets around initialised strings
def removeBracketFromStrings(originalText):
    regex = r'\(([\'\"][^\'\"]+[\'\"])\)'
    originalText = re.sub(regex, "\\g<1>", originalText)

    return originalText


# resolve replace function( .replace() )
def resolveReplace(originalText):
    regex = r'([\'\"]([^\'\"]*?)[\'\"]\s*\)*\s*\.\s*[\'\"]?\s*replace\s*[\'\"]?\s*\(*\s*(?:\[string\])?\s*[\'\"]([^\'\"]*?)[\'\"]\)*\s*,\s*(?:\[string\])?\s*\(*\s*[\'\"]([^\'\"]*?)[\'\"]\)*\s*\))'
    match = re.findall(regex, originalText, flags=re.IGNORECASE)

    for tuple in match:
        fullStatement = tuple[0]
        originalString = tuple[1]
        searchString = tuple[2]
        replacementString = tuple[3]

        # replace
        newReplacedString = originalString.replace(searchString, replacementString)
        originalText = originalText.replace(fullStatement, "\"" + newReplacedString + "\"")

    return originalText


# resolve -CReplace operator
def resolveCReplace(originalText):
    regex = r'([\'\"]([^\'\"]*?)[\'\"]\s*-c?replace\s*\(?\s*[\'\"]([^\'\"]*)[\'\"]\s*,\s*[\'\"]([^\'\"]*)[\'\"])'
    match = re.findall(regex, originalText, flags=re.IGNORECASE)

    for tuple in match:
        fullStatement = tuple[0]
        originalString = tuple[1]
        searchString = tuple[2]
        replacementString = tuple[3]

        newReplacedString = originalString.replace(searchString, replacementString)
        originalText = originalText.replace(fullStatement, "\"" + newReplacedString + "\"")

    return originalText


def extractValueFromArray(originalText):
    regex = r'(\(\s*\[\s*array\s*\]\s*((?:[\'\"][^\'\"]*[\'\"]\s*,?\s*)+)\s*\)\s*\[\s*(\d+)\s*\])'
    match = re.findall(regex, originalText, flags=re.IGNORECASE)

    for tuple in match:
        fullStatement = tuple[0]
        arrayValues = tuple[1]
        arrayIndex = int(tuple[2])

        # get the value to sub in
        splittedArray = arrayValues.split(",")
        finalValue = splittedArray[arrayIndex]
        finalValue = finalValue.strip("\'\"")

        # replace the full statement
        originalText = originalText.replace(fullStatement, "\"" + finalValue + "\"")

    return originalText

# path = "C:\\Users\\user\\Desktop\\Deobsfuscation\\Test.txt"
# working = ""
# with open(path, "r") as myfile:
#     working = myfile.read()
#
# working = resolveSplit(working)
# print(working)
# print("---------------------------------------------------------\n")

# # ------------------------------------------------------------------------------------------------------------------------
# path = "C:\\Users\\user\\Desktop\\Deobsfuscation\\raw.txt"
# path = "C:\\Users\\user\\Desktop\\Deobsfuscation\\powershellToClean.txt"
# working = ""
# with open(path, "r") as myfile:
#     working = myfile.read()
#
# # working = decodeBase64(working)
# # working = removeUselessEscapeChars(working)
# # working = createNewLines(working)
# # working = resolveFormatOperator(working)
# # print(working)
#
#
# # path = "C:\\Users\\user\\Desktop\\Deobsfuscation\\2powershell.txt"
# # working2 = ""
# # with open(path, "r") as myfile:
# #     working2 = myfile.read()
# #
# # working2 = removeUselessEscapeChars(working2)
# # working2 = createNewLines(working2)
# # working2 = resolveFormatOperator(working2)
# # print(working2)
#
# print("Original:")
# print(working)
# print("-------------------------------------------------------------------------------------------------------\n")
#
# print("Decode Base64:")
# working = decodeBase64(working)
# print(working)
# print("-------------------------------------------------------------------------------------------------------\n")
#
# print("Add new lines:")
# working = createNewLines(working)
# print(working)
# print("-------------------------------------------------------------------------------------------------------\n")
#
# print("Remove unused variables:")
# working = removeUnusedVariables(working)
# print(working)
# print("-------------------------------------------------------------------------------------------------------\n")
#
# print("Remove useless escape characters:")
# working = removeUselessEscapeChars(working)
# print(working)
# print("-------------------------------------------------------------------------------------------------------\n")
#
# print("Concat broken strings:")
# working = concatBrokenStrings(working)
# print(working)
# print("-------------------------------------------------------------------------------------------------------\n")
#
# print("Replace variables with values:")
# working = replaceVariableWithValue(working)
# print(working)
# print("-------------------------------------------------------------------------------------------------------\n")
#
# print("Concat broken strings:")
# working = concatBrokenStrings(working)
# print(working)
# print("-------------------------------------------------------------------------------------------------------\n")
#
# print("Resolve char:")
# working = resolveChar(working)
# print(working)
# print("-------------------------------------------------------------------------------------------------------\n")
#
# print("Resolve Split:")
# working = resolveSplit(working)
# print(working)
# print("-------------------------------------------------------------------------------------------------------\n")
#
#
# print("Example 2:")
# path = "C:\\Users\\user\\Desktop\\Deobsfuscation\\2powershell.txt"
# working2 = ""
# with open(path, "r") as myfile:
#     working2 = myfile.read()
#
# print("Original:")
# print(working2)
# print("-------------------------------------------------------------------------------------------------------\n")
#
# print("Add new lines:")
# working2 = createNewLines(working2)
# print(working2)
# print("-------------------------------------------------------------------------------------------------------\n")
#
# print("Remove useless escape characters:")
# working2 = removeUselessEscapeChars(working2)
# print(working2)
# print("-------------------------------------------------------------------------------------------------------\n")
#
# print("Resolve format operators:")
# working2 = resolveFormatOperator(working2)
# print(working2)
# print("-------------------------------------------------------------------------------------------------------\n")


# working = decodeBase64(working)
# working = createNewLines(working)
# working = removeUnusedVariables(working)
# working = removeUselessEscapeChars(working)
# working = concatBrokenStrings(working)
# working = replaceVariableWithValue(working)
# working = concatBrokenStrings(working)
# working = resolveChar(working)
# working = resolveSplit(working)
# print(working)
