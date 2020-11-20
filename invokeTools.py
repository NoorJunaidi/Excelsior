from subprocess import Popen, PIPE
import os
import re


# get location of oledump
def getLocationOledump():
    ROOT_DIR = os.path.dirname(os.path.abspath(__file__))  # This is the Project Root
    OLEDUMP_PATH = os.path.join(ROOT_DIR, 'Imports\\DidierStevensTools\\oledump.py')  # requires `import os`
    return OLEDUMP_PATH


# get location of emldump
# def getLocationEmlDump():
#     ROOT_DIR = os.path.dirname(os.path.abspath(__file__))  # This is the Project Root
#     EMLDUMP_PATH = os.path.join(ROOT_DIR, 'Imports\\DidierStevensTools\\emldump.py') # requires `import os`
#     return EMLDUMP_PATH

# def callEmldumpD(filename):
#     EMLDUMP_PATH = getLocationEmlDump()
#
#
#     p = Popen(['python', EMLDUMP_PATH, "-d", filename], stdin=PIPE, stdout=PIPE, stderr=PIPE)
#     output, err = p.communicate()
#     rc = p.returncode  # optional?
#
#     print(output)
#     print(err.decode())
#
#     return output

############# OLE Metadata start #############


# call oledump with option -M
def oledumpMetadata(filename):
    # # run eml first; if warning: go to callDumpMetadata; if not: extract first then goto callDumpMeta
    # # call emldump
    # emlDumpOutput = callEmldumpD(filename)
    # print(emlDumpOutput)

    output = callDumpMetadata(filename)  # output is bytes not string

    metadataList = formatMetadataOutputToDict(output)

    # check for errors after changed to list (Not a valid file)
    # if not valid file will return "Error" first word; return -1
    if metadataList[0][0] == "Error":
        return -1
    return metadataList
    # print(output)


def callDumpMetadata(filename):
    OLEDUMP_PATH = getLocationOledump()
    p = Popen(['python', OLEDUMP_PATH, '-M',
               filename], stdin=PIPE, stdout=PIPE, stderr=PIPE)
    output, err = p.communicate()
    return output


def formatMetadataOutputToDict(output):
    # works; split to list
    result = output.splitlines()
    decodedResult = []
    # print(result)
    # remove 'b' prefix
    for item in result:
        decodedResult.append(item.decode().strip())
    del result

    # format to dictionary
    resultsList = []
    for item in decodedResult:
        temp = item.split(":")  # temp[0]=key, temp[1]]=value

        # save to list not dict; later in guiMain just display
        newLine = []
        newLine.append(temp[0])  # save attribute
        strippedString = temp[1].strip()
        if temp[1].startswith(" b\'"):  # remove the b' prefix on some values
            strippedString = strippedString.lstrip("b")
            strippedString = strippedString.lstrip("\'")
            strippedString = strippedString.rstrip("\'")
        newLine.append(strippedString)  # save value

        resultsList.append(newLine)

        # old code **********************************************************************************
        # if temp[1] == "":
        #     resultsDict[temp[0]] = "TITLE234"  # To identify headers in GUI; to be BOLDED
        # elif temp[1].startswith(" b\'"):  # remove the b' prefix on some values
        #     strippedString = temp[1].strip()
        #     strippedString = strippedString.lstrip("b")
        #     strippedString = strippedString.lstrip("\'")
        #     strippedString = strippedString.rstrip("\'")
        #     resultsDict[temp[0]] = strippedString
        # else:
        #     resultsDict[temp[0]] = temp[1].strip()
        # old code **********************************************************************************

    return resultsList


############# OLE Metadata end #############

############# OLE extract macros start #############
# main macro finding function
def oledumpMacro(filename):
    output = callDefaultOledump(filename)

    # returns [streamNo, macro type, size(bytes), stream name, decodedMacro]
    foundMacros = findStreamsWithMacros(output)

    # returns decoded macro as string. In a list
    foundMacros = decodeStreamsWithMacros(filename, foundMacros)

    return foundMacros


# calls the default oledump script for the first step to find all macros
def callDefaultOledump(filename):
    oledumpPath = getLocationOledump()
    p = Popen(['python', oledumpPath, filename], stdin=PIPE, stdout=PIPE, stderr=PIPE)
    output, err = p.communicate()

    return output


# Find streams with macros from the result of the default oledump call
def findStreamsWithMacros(output):
    # works; split to list
    result = output.splitlines()
    decodedResult = []
    # remove 'b' prefix
    for item in result:
        decodedResult.append(item.decode().strip())
    del result

    # list of list of found macros
    foundMacros = []
    for item in decodedResult:
        # find both m and M
        hasMacro = re.search("^\w*\d+: [Mm]", item)

        # Find streams with macros (M)
        if hasMacro:
            # Get the stream index
            temp = item.split()
            temp[0] = temp[0][:-1]  # remove colon at the end :
            temp[3] = temp[3][1:-1]  # remove extra quotes at start and end
            foundMacros.append(temp)
    return foundMacros


# returns list of decoded macros
def decodeStreamsWithMacros(filename, foundMacros):
    # return just 1 list
    # append decoded macro to each item in foundMacros

    for macro in foundMacros:
        result = calloledumpDecodeMacro(filename, str(macro[0]))  # will return 1 result; the decoded macro
        temp = result.decode()
        macro.append(temp)

    return foundMacros


# calls the decode macro oledump script
def calloledumpDecodeMacro(filename, index):
    oledumpPath = getLocationOledump()
    # s-stream number, v-decode
    p = Popen(['python', oledumpPath, '-s', index, "-v", filename],
              stdin=PIPE, stdout=PIPE, stderr=PIPE)
    output, err = p.communicate()
    return output

############# OLE extract macros end #############


# # TODO REMOVE TESTCODE LATER
# filename = r"C:\Users\user\Desktop\APT DANGEROUS\0_Malicious Document\a"
# resultList = oledumpMetadata(filename)
