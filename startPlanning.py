import xlrd
import sys
# import pdb

from collections import defaultdict
sys.argv[1] = "C:\\Users\\theco\\Documents\\Hero Wars\\AssignmentScript\\FMEVsGermanSoldiers3-23.xlsx"
sys.argv[2] = True

class AssignmentInfo:
	def __init__(self, assignmentType, rivalName, rivalPower, rivalLocation):
		self.assignmentType = assignmentType
		self.rivalName = rivalName
		self.rivalPower = rivalPower
		self.rivalLocation = rivalLocation
		self.attackNote = ""

	def __init__(self, assignmentType, rivalName, rivalPower, rivalLocation, attackNote):
		self.assignmentType = assignmentType
		self.rivalName = rivalName
		self.rivalPower = rivalPower
		self.rivalLocation = rivalLocation
		self.attackNote = attackNote

	def __repr__(self):
		return self.__str__()
	def __str__(self):
		return "AssignmentInfo:{ " + self.assignmentType + "," +self.rivalName + ","+ str(self.rivalPower) + ","+ str(self.rivalLocation) + ","+ self.attackNote+"}"

class RivalChampionInfo:
	def __init__(self, heroPower, heroLocation, titanPower, titanLocation):
		self.titanPower = titanPower
		self.heroPower = heroPower
		self.titanCleared = False
		self.heroCleared = False
		self.heroLocation = heroLocation
		self.titanLocation = titanLocation

	def __repr__(self):
		return self.__str__()
	def __str__(self):
		return "RivalChampionInfo:{" + str(self.titanPower)+ "," + self.titanLocation + "," + str(self.titanCleared)+ "," + str(self.heroPower)+ "," + self.heroLocation+ "," + str(self.heroCleared)+ "}"

class ChampionInfo:
	def __init__(self, heroPower, titanPower):
		self.titanPower = titanPower
		self.heroPower = heroPower
		self.attacksRemaining = 2
		self.assignments = []
	
	def assign(self, assignmentType, rivalName, rivalPower, rivalLocation, attackNote):
		assignment = AssignmentInfo(assignmentType, rivalName, rivalPower, rivalLocation, attackNote)
		self.assignments.append(assignment)
		self.attacksRemaining = self.attacksRemaining - 1


	def assignAttacks(self, assignmentType, rivalName, rivalPower, rivalLocation, numberOfAttacks, attackNote):
		assignment = AssignmentInfo(assignmentType, rivalName, rivalPower, rivalLocation, attackNote)
		self.assignments.append(assignment)
		self.attacksRemaining = self.attacksRemaining - numberOfAttacks

	def printAssignments(self, champName):
		for assignment in self.assignments:
			# print(assignment)
			if assignment.assignmentType == "Hero":
				print(champName+"(H:"+str(self.heroPower)+ ") "+assignment.rivalName+ " "+ str(assignment.rivalPower)+ " "+ assignment.rivalLocation + assignment.attackNote)
			elif assignment.assignmentType == "Titan":
				print(champName+"(T:"+str(self.titanPower)+ ") "+assignment.rivalName+ " "+ str(assignment.rivalPower)+ " "+ assignment.rivalLocation + assignment.attackNote)
	
	def printAssignmentCategory(self, champName, assignmentType):
		for assignment in self.assignments:
			if assignment.assignmentType == assignmentType:
				if assignment.assignmentType == "Hero":
					print(champName+"(H:"+str(self.heroPower)+ ") "+assignment.rivalName+ "("+ str(assignment.rivalPower)+ ") "+ assignment.rivalLocation + assignment.attackNote)
				elif assignment.assignmentType == "Titan":
					print(champName+"(T:"+str(self.titanPower)+ ") "+assignment.rivalName+ "("+ str(assignment.rivalPower)+ ") "+ assignment.rivalLocation + assignment.attackNote)

	def clearAttacks(self):
		self.assignments.clear()
		self.attacksRemaining = 2

	def __repr__(self):
		return self.__str__()
	def __str__(self):
		return "ChampionInfo:{" + str(self.titanPower)+ "," + str(self.heroPower)+ "," + str(self.attacksRemaining)+ "," + str(self.assignments) + "}"

if not len(sys.argv) >= 2:
	print(sys.argv)
	print("usage: startPlanning.py scriptTemplate.xslx <CategorizeAssignments = True|False>")
	exit()

# =================================== Start Global Variable Setup ===================================
loc = sys.argv[1]
if(len(sys.argv) == 2):
	categorizeAssignments = False	
else:
	categorizeAssignments = sys.argv[2]
wb = xlrd.open_workbook(loc)
FMEChampions = wb.sheet_by_index(0)
FMEOpponentDefense = wb.sheet_by_index(1)

#Setup Values for Tracking
FMEChampionsInfo = {}
RivalChampionsInfo = {}
IgnoreBuildings = []
DifficultyScore = {}
Buildings = []

for i in range (FMEChampions.nrows):
	# print(FMEChampions.row_values(i))
	if(i==0):
		continue
	ChampName = FMEChampions.cell_value(i,0)
	FMEChampionsInfo[ChampName] = ChampionInfo(FMEChampions.cell_value(i,1), FMEChampions.cell_value(i,2))

for x in range (FMEOpponentDefense.nrows):
	if(x==0):
		continue
	RivalChamp = FMEOpponentDefense.cell_value(x,0)
	RivalChampionsInfo[RivalChamp]=RivalChampionInfo(FMEOpponentDefense.cell_value(x,1),FMEOpponentDefense.cell_value(x,2),FMEOpponentDefense.cell_value(x,3),FMEOpponentDefense.cell_value(x,4))


#Sort Our Heroes and Titan Powers for comparisons
sortedFMEChampHeroes = sorted(FMEChampionsInfo.keys(), reverse=True, key=lambda x:FMEChampionsInfo[x].heroPower)
sortedFMEChampTitans = sorted(FMEChampionsInfo.keys(), reverse=True, key=lambda x:FMEChampionsInfo[x].titanPower)
sortedRivalChampHeroesAndLocations = sorted(RivalChampionsInfo.keys(), reverse=True, key=lambda x:RivalChampionsInfo[x].heroPower) 
sortedRivalChampTitansAndLocations = sorted(RivalChampionsInfo.keys(), reverse=True, key=lambda x:RivalChampionsInfo[x].titanPower)

#Setup Some Info we May Use Later
for rival in RivalChampionsInfo.keys():
	# Initialize Building Names
	if RivalChampionsInfo[rival].titanLocation not in Buildings:
		Buildings.append(RivalChampionsInfo[rival].titanLocation)
	if RivalChampionsInfo[rival].heroLocation not in Buildings:
		Buildings.append(RivalChampionsInfo[rival].heroLocation)

	# Determine Building Difficulty
	# score = (x for x, y in enumerate(sortedRivalChampHeroesAndLocations) if y == rival)
	score = sortedRivalChampHeroesAndLocations.index(rival)
	if RivalChampionsInfo[rival].heroLocation in DifficultyScore:
		DifficultyScore[RivalChampionsInfo[rival].heroLocation] = DifficultyScore[RivalChampionsInfo[rival].heroLocation] + score
	else:
		DifficultyScore[RivalChampionsInfo[rival].heroLocation] = score

	# score = (x for x, y in enumerate(sortedRivalChampTitansAndLocations) if y == rival)
	score = sortedRivalChampTitansAndLocations.index(rival)
	if RivalChampionsInfo[rival].titanLocation in DifficultyScore:
		DifficultyScore[RivalChampionsInfo[rival].titanLocation] = DifficultyScore[RivalChampionsInfo[rival].titanLocation] + score
	else:
		DifficultyScore[RivalChampionsInfo[rival].titanLocation] = score

# print(DifficultyScore)
# exit()

# =================================== Stop Global Variable Setup. ===================================
# print(sortedFMEChampHeroes)
# print(sortedFMEChampTitans)
# exit()

# Determine Optimal Titan matchups
def calculateTitanAttacks(matchAgainstTop):
	# print("\n\nTitan Attacks:")
	for matchup in sortedRivalChampTitansAndLocations:
		optimalChamp = ""
		attackNote = ""
		attacksNeeded = 1
		positionCleared = True
		OptimizingAttack = False
		hardAttack = 0
		for FMEchamp in sortedFMEChampTitans:
			if (FMEChampionsInfo[FMEchamp].titanPower>RivalChampionsInfo[matchup].titanPower and 
				FMEChampionsInfo[FMEchamp].attacksRemaining>0 and
				RivalChampionsInfo[matchup].titanLocation not in IgnoreBuildings and
				RivalChampionsInfo[matchup].titanCleared == False):
				# don't use titan attacks of our top 15 heroes unless its bridge and there is no other option
				if ((x for x, y in enumerate(sortedFMEChampHeroes) if y == FMEchamp <= 5 and matchAgainstTop) or RivalChampionsInfo[matchup].titanLocation == "Bridge"):
					if optimalChamp == "":
						optimalChamp = FMEchamp
					elif (FMEChampionsInfo[FMEchamp].titanPower < FMEChampionsInfo[optimalChamp].titanPower and 
						FMEChampionsInfo[FMEchamp].heroPower < FMEChampionsInfo[optimalChamp].heroPower):
						optimalChamp = FMEchamp
					OptimizingAttack = True
					
					# Break and move on if we're matching against top guys
					if(matchAgainstTop):
						break


			elif (OptimizingAttack == False and FMEChampionsInfo[FMEchamp].attacksRemaining>0 and 
				RivalChampionsInfo[matchup].titanCleared == False and 
				(RivalChampionsInfo[matchup].titanLocation == "Bridge")):
				# print("Considering: "+FMEchamp[FMEChampionName] + " vs. " +matchup[RivalChampionName]+ " "+ str(matchup[1][RivalTitanPower]))
				if (FMEChampionsInfo[FMEchamp].titanPower + 10000 > RivalChampionsInfo[matchup].titanPower):
					optimalChamp = FMEchamp
					attackNote = " - Should be close, someone may need to cleanup"
					break
				elif (FMEChampionsInfo[FMEchamp].titanPower + 20000 > RivalChampionsInfo[matchup].titanPower):
					optimalChamp = FMEchamp
					attackNote = " - someone may need to clean this up"
					positionCleared = False
					break
				elif (FMEChampionsInfo[FMEchamp].titanPower + 40000 > RivalChampionsInfo[matchup].titanPower or FMEChampionsInfo[FMEchamp].titanPower + 50000 > RivalChampionsInfo[matchup].titanPower):
					optimalChamp = FMEchamp
					attacksNeeded = 2
					attackNote = " - attack 2x"
					positionCleared = True
					break
				elif (FMEChampionsInfo[FMEchamp].titanPower + 50000 < RivalChampionsInfo[matchup].titanPower):
					optimalChamp = FMEchamp
					attacksNeeded = 2
					attackNote = " - cleanup if target is down your attacks can be used elsewhere"
					positionCleared = False
					hardAttack = hardAttack + 1;
					
					# print(optimalChamp+"(T:"+str(FMEChampionsInfo[optimalChamp].titanPower)+ ") "+matchup+ " "+ str(RivalChampionsInfo[matchup].titanPower)+ " "+ RivalChampionsInfo[matchup].titanLocation + attackNote)
					FMEChampionsInfo[optimalChamp].assignAttacks("Titan", matchup, RivalChampionsInfo[matchup].titanPower, RivalChampionsInfo[matchup].titanLocation, attacksNeeded, attackNote)
					FMEChampionsInfo[optimalChamp].attackNote=attackNote

					if (hardAttack >= 2):
						positionCleared = True
					break
		if(optimalChamp != ""):
			# print(optimalChamp+"(T:"+str(FMEChampionsInfo[optimalChamp].titanPower)+ ") "+matchup+ " "+ str(RivalChampionsInfo[matchup].titanPower)+ " "+ RivalChampionsInfo[matchup].titanLocation + attackNote)
			FMEChampionsInfo[optimalChamp].assignAttacks("Titan", matchup, RivalChampionsInfo[matchup].titanPower, RivalChampionsInfo[matchup].titanLocation, attacksNeeded, attackNote)
			RivalChampionsInfo[matchup].titanCleared = positionCleared

def calculateHeroAttacks(matchAgainstTop):
	# print("\n\nHero Attacks:")
	# for matchup in sortedRivalChampHeroesAndLocations:
	for matchup in sortedRivalChampHeroesAndLocations:
		optimalChamp = ""
		attackNote = ""
		positionCleared = True
		OptimizingAttack = False
		# print("\n")
		for FMEchamp in sortedFMEChampHeroes:
			# print("Considering "+FMEchamp+" vs " + matchup)
			if ((FMEChampionsInfo[FMEchamp].heroPower > RivalChampionsInfo[matchup].heroPower or 
				FMEChampionsInfo[FMEchamp].heroPower + 1000 > RivalChampionsInfo[matchup].heroPower) and 
				RivalChampionsInfo[matchup].heroLocation not in IgnoreBuildings and
				FMEChampionsInfo[FMEchamp].attacksRemaining > 0 and RivalChampionsInfo[matchup].heroCleared == False):
				# Check to see if there is a clear advantage number wise
				# print("Considering "+FMEchamp+" vs " + matchup)
				

				OptimizingAttack = True
				if optimalChamp == "":
					optimalChamp = FMEchamp
					# print("Optimal Champ: "+ optimalChamp)
				elif (FMEChampionsInfo[FMEchamp].heroPower < FMEChampionsInfo[optimalChamp].heroPower):
					optimalChamp = FMEchamp

				# Break and move on if we're matching against top guys
				if(matchAgainstTop):
					break

			elif (not OptimizingAttack and FMEChampionsInfo[FMEchamp].attacksRemaining > 0 and 
				RivalChampionsInfo[matchup].heroLocation not in IgnoreBuildings and
				RivalChampionsInfo[matchup].heroCleared == False):
				if ((FMEChampionsInfo[FMEchamp].heroPower+5000)>RivalChampionsInfo[matchup].heroPower and FMEChampionsInfo[FMEchamp].attacksRemaining > 0 and 
					RivalChampionsInfo[matchup].heroCleared == False):
					attackNote = " - Should be close, may require cleanup"
					optimalChamp = FMEchamp
				elif ((FMEChampionsInfo[FMEchamp].heroPower+10000)>RivalChampionsInfo[matchup].heroPower and FMEChampionsInfo[FMEchamp].attacksRemaining > 0 and 
					RivalChampionsInfo[matchup].heroCleared == False):
					# Times are tough, check to see if there is someone within 10k power we can fight
					attackNote = " - Unless you get lucky this attack may require cleanup"
					optimalChamp = FMEchamp

		if(optimalChamp != "" and FMEChampionsInfo[optimalChamp].attacksRemaining > 0):
			# print(optimalChamp+"(H:"+str(FMEChampionsInfo[optimalChamp].heroPower)+ ") "+matchup+ " "+ str(RivalChampionsInfo[matchup].heroPower)+ " "+ RivalChampionsInfo[matchup].heroLocation + attackNote)
			FMEChampionsInfo[optimalChamp].assign("Hero", matchup, RivalChampionsInfo[matchup].heroPower, RivalChampionsInfo[matchup].heroLocation, attackNote)
			RivalChampionsInfo[matchup].heroCleared = positionCleared

# aType = Titan | Hero
def resetAssignmentCategory(aType):
	for champ in FMEChampionsInfo.keys():
		if FMEChampionsInfo[champ].attacksRemaining < 2:
			#
			tempAssignments = list(FMEChampionsInfo[champ].assignments)
			for assignment in tempAssignments:
				if(assignment.assignmentType == aType and aType == "Titan"):
					# if( 30000 > (FMEChampionsInfo[champ].titanPower - assignment.rivalPower) or True):
					# print(champ +" : "+ str(len(FMEChampionsInfo[champ].assignments)))
					# print(assignment.rivalName)
					RivalChampionsInfo[assignment.rivalName].titanCleared = False
					FMEChampionsInfo[champ].assignments.remove(assignment)
					FMEChampionsInfo[champ].attacksRemaining = 2 - len(FMEChampionsInfo[champ].assignments)
				elif(assignment.assignmentType == aType and aType == "Hero"):
					RivalChampionsInfo[assignment.rivalName].heroCleared = False
					FMEChampionsInfo[champ].assignments.remove(assignment)
					FMEChampionsInfo[champ].attacksRemaining = 2 - len(FMEChampionsInfo[champ].assignments)					

def printAssignedTargetBasedOnName(rivalChampName, assignmentType):
	for champ in FMEChampionsInfo.keys():
		for assignment in FMEChampionsInfo[champ].assignments:
			if assignment.rivalName == rivalChampName and assignment.assignmentType == assignmentType:
				if assignment.assignmentType == "Hero":
					print(champ+"(H:"+str(FMEChampionsInfo[champ].heroPower)+ ") "+assignment.rivalName+ " "+ str(assignment.rivalPower)+ assignment.attackNote)
				elif assignment.assignmentType == "Titan":
					print(champ+"(T:"+str(FMEChampionsInfo[champ].titanPower)+ ") "+assignment.rivalName+ " "+ str(assignment.rivalPower)+ assignment.attackNote)


def main():
	calculateTitanAttacks(False)
	calculateHeroAttacks(False)
	# resetAssignmentCategory("Titan")
	# calculateTitanAttacks(True)
	# exit()
	# Determine if any buildings should be ignored
	# pdb.set_trace()
	for target in RivalChampionsInfo.keys():
		if RivalChampionsInfo[target].heroCleared == False :
			if(RivalChampionsInfo[target].heroLocation not in IgnoreBuildings):
				IgnoreBuildings.append(RivalChampionsInfo[target].heroLocation)
		if RivalChampionsInfo[target].titanCleared == False :
			if(RivalChampionsInfo[target].titanLocation not in IgnoreBuildings):
				IgnoreBuildings.append(RivalChampionsInfo[target].titanLocation)

	if(len(IgnoreBuildings)>0):
		if(len(IgnoreBuildings)>2):
			easiestBuilding = ""
			for building in IgnoreBuildings:
				if easiestBuilding == "":
					easiestBuilding = building
				elif DifficultyScore[building] < DifficultyScore[easiestBuilding]:
					easiestBuilding = building
			print(easiestBuilding)
			IgnoreBuildings.remove(easiestBuilding)
			resetAssignmentCategory("Titan")
			resetAssignmentCategory("Hero")
			calculateTitanAttacks(False)
			calculateHeroAttacks(False)
		else:
			resetAssignmentCategory("Titan")
			resetAssignmentCategory("Hero")
			calculateTitanAttacks(False)
			calculateHeroAttacks(False)
			# resetAssignmentCategory("Titan")
			# calculateTitanAttacks(True)



	# --------------Print out Assignments--------------
	if(categorizeAssignments):
		categories = {"Titan", "Hero"}
		# buildings = {"Bridge","Fire","Nature","Ice","Spring","Lighthouse","Barracks","Mage", "Citadel","Foundry"}
		for category in categories:
			if category == "Titan":
				buildings = {"Bridge","Fire","Nature","Ice","Spring"}
			elif category == "Hero":
				buildings = {"Lighthouse","Barracks","Mage", "Citadel","Foundry"}
			print("\n\n__**"+ category +" Attacks:**__")
			for building in buildings:
				print("\n**"+ building +":**")
				for target in RivalChampionsInfo.keys():
					if RivalChampionsInfo[target].heroLocation == building or RivalChampionsInfo[target].titanLocation == building:
						if RivalChampionsInfo[target].heroCleared == False:
							print("??? - "+target+"(H:"+str(RivalChampionsInfo[target].heroPower)+")")
						elif RivalChampionsInfo[target].titanCleared == False:
							print("??? - "+target+"(T:"+str(RivalChampionsInfo[target].titanPower)+")")
						elif RivalChampionsInfo[target].heroCleared == True:
							printAssignedTargetBasedOnName(target,"Hero")
						elif RivalChampionsInfo[target].titanCleared == True:
							printAssignedTargetBasedOnName(target,"Titan")


				# for champ in FMEChampionsInfo.keys():
				# 	if FMEChampionsInfo[champ].attacksRemaining < 2:
				# 		FMEChampionsInfo[champ].printAssignmentCategory(champ, category)

	else:	
		print("\n\nAssignments:")
		for champ in FMEChampionsInfo.keys():
			if FMEChampionsInfo[champ].attacksRemaining < 2:
				FMEChampionsInfo[champ].printAssignments(champ)
	
	if(not categorizeAssignments):
		print("\n\nRemaining Targets:")
		for target in RivalChampionsInfo.keys():
			if RivalChampionsInfo[target].heroCleared == False :
				print(target+" H:"+str(RivalChampionsInfo[target].heroPower)+" at "+ RivalChampionsInfo[target].heroLocation)
			if RivalChampionsInfo[target].titanCleared == False :
				print(target+" T:"+str(RivalChampionsInfo[target].titanPower)+" at "+ RivalChampionsInfo[target].titanLocation)

	print("\n\nChampion Pool:")	
	# for hero in FMEChampAttacks.keys():
	for hero in FMEChampionsInfo.keys():
		if FMEChampionsInfo[hero].attacksRemaining > 0:
			print(hero, "(H:"+str(FMEChampionsInfo[hero].heroPower)+",T:"+str(FMEChampionsInfo[hero].titanPower)+") has "+ str(FMEChampionsInfo[hero].attacksRemaining)+" attacks remaining.")
	for hero in FMEChampionsInfo.keys():
		if FMEChampionsInfo[hero].attacksRemaining == 0:
			print(hero, "(H:"+str(FMEChampionsInfo[hero].heroPower)+",T:"+str(FMEChampionsInfo[hero].titanPower)+") has "+ str(FMEChampionsInfo[hero].attacksRemaining)+" attacks remaining.")


	print("\n\nIgnoring Buildings:"+str(IgnoreBuildings))
	print("Remaining Targets:")
	for target in RivalChampionsInfo.keys():
		if RivalChampionsInfo[target].heroCleared == False :
			print(target+" H:"+str(RivalChampionsInfo[target].heroPower)+" at "+ RivalChampionsInfo[target].heroLocation)
		if RivalChampionsInfo[target].titanCleared == False :
			print(target+" T:"+str(RivalChampionsInfo[target].titanPower)+" at "+ RivalChampionsInfo[target].titanLocation)


main()
input('\nPress Enter to exit')
