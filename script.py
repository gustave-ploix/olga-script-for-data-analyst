import sys
import re

from pathlib import Path

import pandas as pd 


finalArray = []

readableFile = pd.read_excel("./marketing_score_may_22.xlsx")

# sys.setrecursionlimit(10000)

data = pd.DataFrame(readableFile, columns=[
    "Email ", "Marketing score", "Marketing score details"
])

dataInList = data.values.tolist()

split_lists = [dataInList[x:x+500] for x in range(0, len(dataInList), 500)]



def convertToXLSX(arrays) :
    listForDataFrame = []

    for i in range(0, len(arrays)) :
        listForDataFrame += arrays[i]

    frame = pd.DataFrame(data=listForDataFrame)


    frame.to_excel(str(Path.home() / "Downloads/data.xlsx"), index=False, header=True)


def removeOccurenceInMSD(data) :
    opens = []
    clicks = []
    softBounce = []
    hardBounce = []
    unsubscribed = []

    lastOpen = ""
    lastClick = ""
    lastSoftBounce = ""
    lastHardBounce = ""
    lastUnsubscribed = ""

    finalString = ""

    cleanData = []

    for j in range(0, len(data)) :
        if len(data[j]) > 1 :
            cleanData.append(data[j])

    for i in range(0, len(cleanData)) :
        if "open(s)" in cleanData[i] :
            opens.append(cleanData[i] + "; ")
            lastOpen = opens.pop()
            # finalString += lastOpen + "; "
        if "link click(s)" in cleanData[i] :
            clicks.append(cleanData[i] + "; ")
            lastClick = clicks.pop()
            # finalString += lastClick + "; "
        if "Soft Bounce" in cleanData[i] :
            softBounce.append(cleanData[i] + "; ")
            lastSoftBounce = softBounce.pop()
            # finalString += lastSoftBounce + "; "
        if "Hard Bounce" in cleanData[i] :
            hardBounce.append(cleanData[i] + "; ")
            lastHardBounce = hardBounce.pop()
            # finalString += lastHardBounce + "; "
        if "Unsubscribed" in cleanData[i] :
            unsubscribed.append(cleanData[i] + "; ")
            lastUnsubscribed = unsubscribed.pop()

    finalString = lastOpen + lastClick + lastSoftBounce + lastHardBounce + lastUnsubscribed

    return finalString







def removeOccurences(data) : 
    
    cleanData = []
    indexToDelete = []

    totalOpens = []
    totalClicks = []

    unsortedMarketingDetail = ""



    firstElem = {
        "email": data[0][0],
        "marketing score": 0,
        "marketing score detail": ""
    }

    for i in range(0, len(data)) : 
        if firstElem["email"] == data[i][0] :
            firstElem["marketing score"] += data[i][1]
            indexToDelete.append(i)

            if "Soft bounce;" in str(data[i][2]) :
                unsortedMarketingDetail += "Soft Bounce;"
            if "Hard bounce;" in str(data[i][2]) :
                unsortedMarketingDetail += "Hard Bounce;"
            if "open(s)" in str(data[i][2]):
                sum = 0
                opens = ""
                splittedList = str(data[i][2]).split(";")
                
                for j in range(0, len(splittedList)) :
                    if "open(s)" in splittedList[j] :
                        opens = splittedList[j]

                num = re.search(r'\d', opens).group()
                totalOpens.append(int(num))
                for k in range(0, len(totalOpens)) :
                    sum += totalOpens[k]

                finalString = str(sum) + " open(s);"
                unsortedMarketingDetail += finalString

            if "link click(s)" in str(data[i][2]):
                sum = 0
                clicks = ""
                splittedList = data[i][2].split(";")

                for j in range(0, len(splittedList)) :
                    if "link click(s)" in splittedList[j] :
                        clicks = splittedList[j]
                

                num = re.search(r'\d', clicks).group()
                totalClicks.append(int(num))
                for k in range(0, len(totalClicks)) :
                    sum += totalClicks[k]
                
                finalString = str(sum) + " link click(s);"
                
                unsortedMarketingDetail += finalString

            if "Unsubscribed" in str(data[i][2]):
                unsortedMarketingDetail += "Unsubscribed;"

    for j in reversed(indexToDelete) :
        del(data[j])

    tableOfMSD = unsortedMarketingDetail.split(";")
    firstElem["marketing score detail"] = removeOccurenceInMSD(tableOfMSD)

    cleanData.append(firstElem)

    # data.pop(0)

    if len(data) > 0 :
        cleanData = cleanData + removeOccurences(data)


    # if len(cleanData) > 10 :
    #     convertToXLSX(cleanData)

    return cleanData

for i in range(0, len(split_lists)) :
    # print(removeOccurences(split_lists[i]))
    finalArray.append(removeOccurences(split_lists[i]))

if len(finalArray) > 10 :
    convertToXLSX(finalArray)
