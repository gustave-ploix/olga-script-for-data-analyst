import sys
import pandas as pd 

readableFile = pd.read_excel("./marketing_score_may_22.xlsx")

# sys.setrecursionlimit(10000)

data = pd.DataFrame(readableFile, columns=[
    "Email ", "Marketing score", "Marketing score details"
])

dataInList = data.values.tolist()

split_lists = [dataInList[x:x+500] for x in range(0, len(dataInList), 500)]

print("input => " + str(len(split_lists[0])))

def removeOccurences(data) : 
    
    cleanData = []
    indexToDelete = []

    # print(data[0][0])

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
                firstElem["marketing score detail"] += "Soft Bounce;"


    for j in reversed(indexToDelete) :
        del(data[j])

    cleanData.append(firstElem)

    # data.pop(0)

    if len(data) > 0 :
        cleanData = cleanData + removeOccurences(data)


    return cleanData

half = round(len(split_lists) /2)
print(split_lists[half])

# print(split_lists[len(split_lists) - 1])

# for i in range (0, half + 1) :
#     print(removeOccurences(split_lists[i]))

for i in range(0, len(split_lists)) :
    print(removeOccurences(split_lists[i]))

    # print(i)
# removeOccurences(split_lists[0])
