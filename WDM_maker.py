'''

Author: Grant Pausanos
Date: Feb 19, 2019
Purpose: Autonomously creates an organized Weighted Decision Matrix after the user inputs required data to the table.
Version: 1.0

'''

import xlsxwriter


def addCriteria():
    print("\nAdding criteria")

    criteriaName = input("Enter criteria name: ")
    criteria.append(criteriaName)

    print("\nCurrent Criteria:")

    for counter in range(len(criteria)):
        print(criteria[counter])


def addWeight(givenWeights, userOption):
    if len(criteria) > 0:
        justificationON = None
        if len(givenWeights) != 0:
            userChoice = input("It seems that you have already inputted your weights. Would you like to overwrite the existing weights? (y/n): ")
            if userChoice == "y":
                print("Overwriting weights...")

                for counter in range(len(givenWeights)):
                    del givenWeights[0]
                    if len(weightJustifications) != 0:
                        del weightJustifications[0]
            else:
                print("Returning back to the menu...")
                return None
        print("\nAdding weight")
        weightSum = 0
        if userOption == "2":
            justificationON = justification()
        else:
            print("TEMPORARY WEIGHTS:")
        for counter in range(len(criteria)):
            userChoice = getProperPromptFloat("Enter a weight for " + str(criteria[counter]) + ": ")
            givenWeights.append(userChoice)
            weightSum += userChoice
            if justificationON == True:
                userChoice = input("Enter justification for weight: ")
                weightJustifications.append(userChoice)
        print("\nCurrent Weights: ")
        for counter in range(len(givenWeights)):
            if justificationON == True:
                print(str(criteria[counter]) + ": " + str(givenWeights[counter]) + "          Justification: " + str(weightJustifications[counter]))
            else:
                print(str(criteria[counter]) + ": " + str(givenWeights[counter]))
        if weightSum != 100.0:
            print("Error: Your weights do not add up to 100! You must add the proper weights")
    else:
        print("\nError: You did not type any criteria!")


def addConcepts():
    print("\nAdding concepts")
    ready = True
    tempConcept = []
    conceptName = input("Enter concept name: ")
    if concepts != 0:
        for counter in range(len(concepts)):
            if conceptName == concepts[counter]:
                print("It seems that you have already included this concept!")
                ready = False
    if ready == True:
        concepts.append(conceptName)
    print("\nCurrent Concepts:")
    for counter in range(len(concepts)):
        print(concepts[counter])


def inputRawScore():
    if len(criteria) > 0 and len(concepts) > 0:
        print("\nInputting raw scores")
        if len(rawScores) != 0:
            userChoice = input("It seems that you have already inputted your raw scores. Would you like to overwrite the existing raw scores? (y/n): ")
            if userChoice == "y":
                print("Overwriting raw scores and justifications...")
                for counter in range(len(weights)):
                    del rawScores[0]
                    del rawScoreJustification[0]
            else:
                print("Returning back to the menu...")
                return None
        justificationON = justification()
        for counter in range(len(concepts)):
            tempScore = []
            tempJustification = []
            for counterTwo in range(len(criteria)):
                userChoice = getProperPromptFloat("Enter a raw score for " + str(concepts[counter]) + " in " + str(criteria[counterTwo]) + ": ")
                while userChoice < 0 or userChoice > 10:
                    print("Error. You must type a raw score between 0 and 10")
                    userChoice = getProperPromptFloat("Enter a raw score for " + str(concepts[counter]) + " in " + str(criteria[counterTwo]) + ": ")
                tempScore.append(userChoice)
                if justificationON == True:
                    userChoice = input("Enter justification for weight: ")
                    tempJustification.append(userChoice)
            rawScoreJustification.append(tempJustification)
            rawScores.append(tempScore)

        for counter in range(len(concepts)):
            print("\nRaw scores for " + str(concepts[counter]) + ": ")
            for counterTwo in range(len(criteria)):
                if justificationON == True:
                    print(str(criteria[counterTwo]) + ": " + str(rawScores[counter][counterTwo]) + "          Justification: ", rawScoreJustification[counter][counterTwo])
                else:
                    print(str(criteria[counterTwo]) + ": " + str(rawScores[counter][counterTwo]))
            print("")
    else:
        print("\nError: You didn't type anything in the criteria or concepts!")


def deleteConcept():
    if len(concepts) != 0:
        print("\nDeleting concept")
        for counter in range(len(concepts)):
            print("Type " + str(counter) + " to delete " + str(concepts[counter]))
        userChoice = getProperPromptInt("\nSelect an option: ")
        print("Deleting " + str(concepts[userChoice]) + " and any scores inputted to " + str(concepts[userChoice]))
        deletedConcept = concepts[userChoice]
        del concepts[userChoice]

        if len(rawScores) != 0:
            del rawScores[userChoice]

        if len(rawScoreJustification) != 0:
            del rawScoreJustification[userChoice]

        if len(finalScoreForAll) != 0:
            del finalScoreForAll[userChoice]

        print("\nCurrent Concepts:")
        for counter in range(len(concepts)):
            print(concepts[counter])

        if len(finalRankings) != 0:
            deleteIndex = 0
            for counter in range(len(finalConceptRanking)):
                if finalConceptRanking[counter] == deletedConcept:
                    deleteIndex = counter
            del finalConceptRanking[deleteIndex]
            del finalRankings[deleteIndex]
    else:
        print("\nYou do not have any concepts available.")



def deleteCriteria():
    if len(criteria) != 0:
        print("\nDeleting criteria")
        for counter in range(len(criteria)):
            print("Type " + str(counter) + " to delete " + str(criteria[counter]))
        userChoice = getProperPromptInt("Select an option: ")
        print("Deleting " + str(criteria[userChoice]) + " and any relevant information of it.")
        del criteria[userChoice]

        if len(rawScores) != 0:
            for counter in range(len(rawScores)):
                if len(rawScores[counter]) != 0:
                    del rawScores[counter][userChoice]

        if len(rawScoreJustification) != 0:
            for counter in range(len(rawScoreJustification)):
                if len(rawScoreJustification[counter]) != 0:
                    del rawScoreJustification[counter][userChoice]

        if len(finalScoreForAll) != 0:
            for counter in range(len(finalScoreForAll)):
                if len(finalScoreForAll[counter]) != 0:
                    del finalScoreForAll[counter][userChoice]

        if len(weights) != 0:
            del weights[userChoice]

        if len(weightedScores) != 0:
            del weightedScores[userChoice]

        if len(weightJustifications) != 0:
            del weightJustifications[userChoice]

        print("\nCurrent Criteria:")
        for counter in range(len(criteria)):
            print(criteria[counter])
    else:
        print("\nYou do not have any criteria available.")


def advancedOptions():
    print("Type 0 to change the name of a criteria")
    print("Type 1 to change a weight justification")
    print("Type 2 to change the name of a concept")
    print("Type 3 to change a raw score")
    print("Type 4 to change a raw score justification")
    print("Press ENTER KEY to return to main menu")

    userChoice = input("\nSelect a choice: ")

    if userChoice == "0":
        changeCriteria()
    elif userChoice == "1":
        changeWeightJustification()
    elif userChoice == "2":
        changeConcept()
    elif userChoice == "3":
        changeRawScore()
    elif userChoice == "4":
        changeRawScoreJustification()


def changeCriteria():
    if len(criteria) != 0:
        for counter in range(len(criteria)):
            print("Type " + str(counter) + " to change the name of " + str(criteria[counter]))
        userChoice = getProperPromptInt("Select an option: ")
        criteria[userChoice] = input("\nChange the name of the criteria here: ")
        print("\nNew Criteria:")
        for counter in range(len(criteria)):
            print(criteria[counter])
    else:
        print("\nError: There is no criteria")


def changeWeightJustification():
    if len(weightJustifications) != 0:
        for counter in range(len(criteria)):
            print("Type " + str(counter) + " to change the justification of " + str(criteria[counter]))
        userChoice = getProperPromptInt("\nSelect an option: ")
        print("Current justification for " + str(criteria[userChoice]) + ": ", weightJustifications[userChoice])
        weightJustifications[userChoice] = input("\nChange the justification here: ")
        print("\nNew Weight Justifications:")
        for counter in range(len(criteria)):
            print(str(criteria[counter]) + ": " + str(weights[counter]) + "          Justification: " + str(weightJustifications[counter]))
    else:
        print("\nError: You do not have any justifications on your weights")


def changeConcept():
    if len(concepts) != 0:
        for counter in range(len(concepts)):
            print("Type " + str(counter) + " to change the name of " + str(concepts[counter]))
        userChoice = getProperPromptInt("\nSelect an option: ")
        changedName = concepts[userChoice]
        concepts[userChoice] = input("\nChange the name of the concept here: ")
        print("\nNew Concepts:")
        for counter in range(len(concepts)):
            print(concepts[counter])
        for counter in range(len(finalConceptRanking)):
            if finalConceptRanking[counter] == changedName:
                finalConceptRanking[counter] = concepts[userChoice]

    else:
        print("\nError: There are no concepts")


def changeRawScore():
    if len(rawScores) != 0:
        print("")
        for counter in range(len(concepts)):
            print("Type " + str(counter) + " to change the raw score of " + str(concepts[counter]))
        userChoice = getProperPromptInt("\nSelect an option: ")
        print("\nIn which criteria?")
        for counter in range(len(criteria)):
            print("Type " + str(counter) + " for " + str(criteria[counter]))
        userChoiceTwo = getProperPromptInt("\nSelect an option: ")
        print("")
        print("Current raw score for " + str(concepts[userChoice]) + "'s " + str(criteria[userChoiceTwo]) + ": ", rawScores[userChoice][userChoiceTwo])
        rawScores[userChoice][userChoiceTwo] = getProperPromptFloat("Change the raw score here: ")
        print("")
        print("New Raw Scores: ")
        if len(rawScores) != 0:
            for counter in range(len(concepts)):
                print("Raw scores for " + str(concepts[counter]) + ": ")
                for counterTwo in range(len(criteria)):
                    if len(rawScoreJustification[counter]) != 0:
                        print(str(criteria[counterTwo]) + ": " + str(rawScores[counter][counterTwo]) + "          Justification: ",rawScoreJustification[counter][counterTwo])
                    else:
                        print(str(criteria[counterTwo]) + ": " + str(rawScores[counter][counterTwo]))
    else:
        print("\nError: There are no raw scores")


def changeRawScoreJustification():
    if len(rawScoreJustification) != 0:
        for counter in range(len(concepts)):
            print("Type " + str(counter) + " to change the raw score justification of " + str(concepts[counter]))
        userChoice = getProperPromptInt("\nSelect an option: ")
        print("In which criteria?")
        for counter in range(len(criteria)):
            print("Type " + str(counter) + " for " + str(criteria[counter]))
        userChoiceTwo = getProperPromptInt("\nSelect an option: ")
        print("Current raw score justification for " + str(concepts[userChoice]) + " in " + str(criteria[userChoiceTwo]) + ": ", rawScoreJustification[userChoice][userChoiceTwo])
        rawScoreJustification[userChoice][userChoiceTwo] = input("Change the raw score justification here: ")
        for counter in range(len(concepts)):
            print("\nRaw scores for " + str(concepts[counter]) + ": ")
            for counterTwo in range(len(criteria)):
                print(str(criteria[counterTwo]) + ": " + str(rawScores[counter][counterTwo]) + "          Justification: ", rawScoreJustification[counter][counterTwo])
            print("")
    else:
        print("\nError: There are no raw score justifications.")


def computeScores(score, tempWeights, finalOrTemp):
    global finalScoreForAll
    global finalTotalScore
    global finalTotalScoreSorting
    global finalConceptSorting
    global finalConceptRanking
    global finalRankings

    if len(score) == 0 or len(tempWeights) == 0:
        print("\nYou do not have any scores or weights")
        return None

    print("\nComputing Scores")

    tempScoreForAll = []
    tempTotalScore = []
    tempTotalScoreSorting = []
    tempConceptSorting = []
    tempConceptRanking = []
    tempRankings = []

    for counter in range(len(score)):
        tempScores = []
        tempSum = 0
        tempConceptSorting.append(concepts[counter])
        for counterTwo in range(len(criteria)):
            tempScores.append(score[counter][counterTwo] * (tempWeights[counterTwo] / 100))
            tempSum += score[counter][counterTwo] * (tempWeights[counterTwo] / 100)
        tempScoreForAll.append(tempScores)
        tempTotalScore.append(tempSum)
        tempTotalScoreSorting.append(tempSum)
    while len(tempRankings) != len(concepts):
        greatestScore = 0
        greatestScoreIndex = 0
        for counter in range(len(tempTotalScoreSorting)):
            if greatestScore < tempTotalScoreSorting[counter]:
                greatestScore = tempTotalScoreSorting[counter]
                greatestScoreIndex = counter
        tempRankings.append(greatestScore)
        tempConceptRanking.append(tempConceptSorting[greatestScoreIndex])
        del tempTotalScoreSorting[greatestScoreIndex]
        del tempConceptSorting[greatestScoreIndex]

    for counter in range(len(concepts)):
        print("\nWeighted scores for " + str(concepts[counter]))
        for counterTwo in range(len(criteria)):
            print(str(criteria[counterTwo]) + ": " + str(tempScoreForAll[counter][counterTwo]))

    print("\nRankings in order:")
    for counter in range(len(tempRankings)):
        print(str(counter + 1) + ". " + str(tempConceptRanking[counter]) + " with a score of ", str(round(tempRankings[counter], 2)))

    if finalOrTemp == "8":
        finalScoreForAll = tempScoreForAll
        finalTotalScore = tempTotalScore
        finalTotalScoreSorting = tempTotalScoreSorting
        finalConceptSorting = tempConceptSorting
        finalConceptRanking = tempConceptRanking
        finalRankings = tempRankings



def justification():
    result = None
    print("Type y to add justification for weights")
    print("Type n to not add justification for weights\n")
    userChoice = input("Your option: ")
    print("")
    if userChoice == 'y':
        result = True
    elif userChoice == 'n':
        result = False
    return result


def displayInfo():
    weightSum = 0
    for counter in range(len(weights)):
        weightSum += weights[counter]

    if weightSum != 100:
        print("\nWeights do not add up to 100. Change the weights and compute the scores again to show accurate information.")
        print("")
        return None

    if len(criteria) > len(finalConceptRanking):
        print("\n You forgot to input raw scores or computes scores for newly added concept(s)")
        print("")
        return None

    print("\nCriteria and Weights:")
    if len(criteria) != 0 or len(weights) != 0:
        if len(weightJustifications) == 0:
            for counter in range(len(weights)):
                print(str(criteria[counter]) + ": " + str(weights[counter]))
        else:
            for counter in range(len(weights)):
                print(str(criteria[counter]) + ": " + str(weights[counter])+ "          Justification: " + str(weightJustifications[counter]))
    else:
        print("There are no criteria or weights.")

    print("\nConcepts:")
    if len(concepts) != 0:
        for counter in range(len(concepts)):
            print(concepts[counter])
    else:
        print("There are no concepts.")
    print("")

    print("Raw Scores:")
    if len(rawScores) != 0:
        for counter in range(len(concepts)):
            print("\nRaw scores for " + str(concepts[counter]) + ": ")
            for counterTwo in range(len(criteria)):
                if len(rawScoreJustification[counter]) != 0:
                    print(str(criteria[counterTwo]) + ": " + str(rawScores[counter][counterTwo]) + "          Justification: ",rawScoreJustification[counter][counterTwo])
                else:
                    print(str(criteria[counterTwo]) + ": " + str(rawScores[counter][counterTwo]))
    else:
        print("There are no raw scores.")
    print("")

    print("Weighted Scores:")
    if len(finalScoreForAll) != 0:
        for counter in range(len(concepts)):
            print("\nWeighted scores for " + str(concepts[counter]))
            for counterTwo in range(len(criteria)):
                print(str(criteria[counterTwo]) + ": " + str(finalScoreForAll[counter][counterTwo]))
    else:
        print("Weighted scores have not yet been calculated.")
    print("")

    print("\nRankings in order:")
    if len(finalRankings) != 0:
        for counter in range(len(finalRankings)):
            print(str(counter + 1) + ". " + str(finalConceptRanking[counter]) + " with a score of ", str(round(finalRankings[counter], 2)))
    else:
        print("Rankings have not yet been calculated.")

def quit():
    weightSum = 0
    for counter in range(len(weights)):
        weightSum += weights[counter]

    if weightSum != 100 or len(criteria) == 0 or len(weights) == 0 or len(concepts) == 0 or len(rawScores) == 0 or len(finalRankings) == 0:
        print("\nPlease ensure that you have all the appropriate information to complete the WDM")
        print("Type 0 to continue progress")
        print("Type 1 to quit progress (this will delete any existing data)")

        userChoice = getProperPromptInt("\nSelect a choice: ")

        if userChoice == 0:
            return True
        else:
            return False

    print("\nquittng...")
    workbook = xlsxwriter.Workbook(fileName)
    wdmMain = workbook.add_worksheet()

    weightJustificationsColumn = 0
    rawScoreJustificationColumn = 0
    furthestColumn = 0
    rawScoreJ = False

    mergeFormatHeadings = workbook.add_format({'bold': 1, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    mergeFormatBasic = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})

    for counter in range(len(concepts)):
        if len(rawScoreJustification[counter]) > 0:
            rawScoreJustificationColumn = 1
            rawScoreJ = True
    if len(weightJustifications) > 0:
        weightJustificationsColumn = 1

    for counter in range(len(criteria)):
        wdmMain.write(counter + 5, 1, criteria[counter], mergeFormatBasic)
        wdmMain.set_column(1, 1, 15)

    for counter in range(len(weights)):
        wdmMain.write(counter + 5, 2, str(weights[counter]) + "%", mergeFormatBasic)
        wdmMain.write(5 + len(weights), 2, "Sum", mergeFormatHeadings)
        wdmMain.write(6 + len(weights), 2, "Ranking", mergeFormatHeadings)

    if len(weightJustifications) > 0:
        for counter in range(len(weightJustifications)):
            wdmMain.write(counter + 5, 3, weightJustifications[counter], mergeFormatBasic)
            wdmMain.set_column(3, 3, 20)
            wdmMain.merge_range(1, 3, 4, 3, "Weight\nJustification", mergeFormatHeadings)

    for counter in range(len(concepts)):
        for counterTwo in range(len(criteria)):
            wdmMain.write(counterTwo + 5, findColumn(weightJustificationsColumn, rawScoreJustificationColumn, counter + 1), rawScores[counter][counterTwo], mergeFormatBasic)
            wdmMain.merge_range(3, findColumn(weightJustificationsColumn, rawScoreJustificationColumn, counter + 1), 4, findColumn(weightJustificationsColumn, rawScoreJustificationColumn, counter + 1), "Raw\nScore", mergeFormatBasic)
            wdmMain.merge_range(2, findColumn(weightJustificationsColumn, rawScoreJustificationColumn, counter + 1), 2, findColumn(weightJustificationsColumn, rawScoreJustificationColumn, counter + 1) + 1+ rawScoreJustificationColumn, concepts[counter], mergeFormatBasic)
            if rawScoreJ == True:
                wdmMain.write(counterTwo + 5, findColumn(weightJustificationsColumn, rawScoreJustificationColumn, counter + 1) + 1, rawScoreJustification[counter][counterTwo], mergeFormatBasic)
                wdmMain.merge_range(3, findColumn(weightJustificationsColumn, rawScoreJustificationColumn, counter + 1) + 1, 4, findColumn(weightJustificationsColumn, rawScoreJustificationColumn, counter + 1) + 1,  "Justification", mergeFormatBasic)
                wdmMain.set_column(findColumn(weightJustificationsColumn, rawScoreJustificationColumn, counter + 1) + 1, findColumn(weightJustificationsColumn, rawScoreJustificationColumn, counter + 1) + 1, 20)
            wdmMain.write(counterTwo + 5, findColumn(weightJustificationsColumn, rawScoreJustificationColumn, counter + 1) + 1 + rawScoreJustificationColumn, finalScoreForAll[counter][counterTwo], mergeFormatBasic)
        wdmMain.write(len(criteria) + 5, findColumn(weightJustificationsColumn, rawScoreJustificationColumn, counter + 1) + 1 + rawScoreJustificationColumn, finalTotalScore[counter], mergeFormatBasic)
        wdmMain.merge_range(3, findColumn(weightJustificationsColumn, rawScoreJustificationColumn, counter + 1) + 1 + rawScoreJustificationColumn, 4, findColumn(weightJustificationsColumn, rawScoreJustificationColumn, counter + 1) + 1 + rawScoreJustificationColumn, "Weighted\nScore", mergeFormatBasic)
        wdmMain.set_column(findColumn(weightJustificationsColumn, rawScoreJustificationColumn, counter + 1) + 1 + rawScoreJustificationColumn, findColumn(weightJustificationsColumn, rawScoreJustificationColumn, counter + 1) + 1 + rawScoreJustificationColumn, 12)
        if furthestColumn < findColumn(weightJustificationsColumn, rawScoreJustificationColumn, counter + 1) + 1 + rawScoreJustificationColumn:
            furthestColumn = findColumn(weightJustificationsColumn, rawScoreJustificationColumn, counter + 1) + 1 + rawScoreJustificationColumn
        for counterThree in range(len(finalConceptRanking)):
            if finalConceptRanking[counterThree] == concepts[counter]:
                wdmMain.write(len(criteria) + 6, findColumn(weightJustificationsColumn, rawScoreJustificationColumn, counter + 1) + 1 + rawScoreJustificationColumn, counterThree + 1, mergeFormatBasic)


    wdmMain.merge_range(1, findColumn(weightJustificationsColumn, rawScoreJustificationColumn, 1), 1, furthestColumn, "Concepts", mergeFormatHeadings)
    wdmMain.merge_range(1, 1, 4, 1, "Criteria", mergeFormatHeadings)
    wdmMain.merge_range(1, 2, 4, 2, "Weights", mergeFormatHeadings)

    workbook.close()
    return False


def findColumn(weightJColumn, rawScoreJColumn, conceptIndex):
    columnConstant = 3 + weightJColumn
    commonDifference = 2 + rawScoreJColumn
    desiredColumn = columnConstant + commonDifference * (conceptIndex - 1)
    return desiredColumn

def getProperPromptInt(userInput):
    try:
        value = int(input(userInput))
    except ValueError:
        print("Please type a number, not a letter.")
        return getProperPromptInt(userInput)
    return value

def getProperPromptFloat(userInput):
    try:
        value = float(input(userInput))
    except ValueError:
        print("Please type a number, not a letter.")
        return getProperPromptFloat(userInput)
    return value


running = True
criteria = []
concepts = []
rawScores = []
weightedScores = []
weights = []
weightJustifications = []
rawScoreJustification = []
sumScores = []
rankedConcepts = []
weightsRobustness = []

finalScoreForAll = []
finalTotalScore = []
finalTotalScoreSorting = []
finalConceptSorting = []
finalConceptRanking = []
finalRankings = []

print("Welcome to the WDM Maker! \n")

fileName = input("Enter the name of the WDM file: ")
fileName = fileName + ".xlsx"
print("File name is", fileName)

while running:
    print("\nType 1 to add criteria")
    print("Type 2 to add or change weights")
    print("Type 3 to add concepts")
    print("Type 4 to input raw scores")
    print("Type 5 to delete concept")
    print("Type 6 to delete criteria")
    print("Type 7 to see advanced options")
    print("Type 8 to compute scores")
    print("Type 9 to check robustness of highest ranked concept")
    print("Type 0 to display all information")
    print("Type 00 to quit/finish \n")

    userChoice = input("Select a choice: ")

    if userChoice == "1":
        addCriteria()
    elif userChoice == "2":
        addWeight(weights, userChoice)
    elif userChoice == "3":
        addConcepts()
    elif userChoice == "4":
        inputRawScore()
    elif userChoice == "5":
        deleteConcept()
    elif userChoice == "6":
        deleteCriteria()
    elif userChoice == "7":
        print("\n***ADVANCED OPTIONS WILL OVERWRITE ANY DESIRED DATA")
        advancedOptions()
    elif userChoice == "8":
        computeScores(rawScores, weights, userChoice)
    elif userChoice == "9":
        addWeight(weightsRobustness, userChoice)
        computeScores(rawScores, weightsRobustness, userChoice)
    elif userChoice == "0":
        displayInfo()
    elif userChoice == "00":
        running = quit()
