import xlrd
import sys
from collections import defaultdict

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
					print(champName+"(H:"+str(self.heroPower)+ ") "+assignment.rivalName+ " "+ str(assignment.rivalPower)+ " "+ assignment.rivalLocation + assignment.attackNote)
				elif assignment.assignmentType == "Titan":
					print(champName+"(T:"+str(self.titanPower)+ ") "+assignment.rivalName+ " "+ str(assignment.rivalPower)+ " "+ assignment.rivalLocation + assignment.attackNote)

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

# print(FMEChampionsInfo)
# print(RivalChampionsInfo)
# exit()

#Sort Our Heroes and Titan Powers for comparisons
sortedFMEChampHeroes = sorted(FMEChampionsInfo.keys(), reverse=True, key=lambda x:FMEChampionsInfo[x].heroPower)
sortedFMEChampTitans = sorted(FMEChampionsInfo.keys(), reverse=True, key=lambda x:FMEChampionsInfo[x].titanPower)
sortedRivalChampHeroesAndLocations = sorted(RivalChampionsInfo.keys(), reverse=True, key=lambda x:RivalChampionsInfo[x].heroPower) 
sortedRivalChampTitansAndLocations = sorted(RivalChampionsInfo.keys(), reverse=True, key=lambda x:RivalChampionsInfo[x].titanPower)


# =================================== Stop Global Variable Setup. ===================================
# print(sortedFMEChampHeroes)
# print(sortedFMEChampTitans)
# exit()

# Determine Optimal Titan matchups
def calculateTitanAttacks(byPassBridgeRestriction):
	print("\n\nTitan Attacks:")
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
				RivalChampionsInfo[matchup].titanCleared == False):
				# don't use titan attacks of our top 15 heroes unless its bridge and there is no other option
				if ((x for x, y in enumerate(sortedFMEChampHeroes) if y == FMEchamp <= 16) or RivalChampionsInfo[matchup].titanLocation == "Bridge"):
					if optimalChamp == "":
						optimalChamp = FMEchamp
					elif (FMEChampionsInfo[FMEchamp].titanPower < FMEChampionsInfo[optimalChamp].titanPower and 
						FMEChampionsInfo[FMEchamp].heroPower < FMEChampionsInfo[optimalChamp].heroPower):
						optimalChamp = FMEchamp
					OptimizingAttack = True
			elif (OptimizingAttack == False and FMEChampionsInfo[FMEchamp].attacksRemaining>0 and 
				RivalChampionsInfo[matchup].titanCleared == False and 
				(RivalChampionsInfo[matchup].titanLocation == "Bridge" or byPassBridgeRestriction == True)):
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
					
					print(optimalChamp+"(T:"+str(FMEChampionsInfo[optimalChamp].titanPower)+ ") "+matchup+ " "+ str(RivalChampionsInfo[matchup].titanPower)+ " "+ RivalChampionsInfo[matchup].titanLocation + attackNote)
					FMEChampionsInfo[optimalChamp].assignAttacks("Titan", matchup, RivalChampionsInfo[matchup].titanPower, RivalChampionsInfo[matchup].titanLocation, attacksNeeded, attackNote)
					FMEChampionsInfo[optimalChamp].attackNote=attackNote

					if (hardAttack >= 2):
						positionCleared = True
					break
		if(optimalChamp != ""):
			print(optimalChamp+"(T:"+str(FMEChampionsInfo[optimalChamp].titanPower)+ ") "+matchup+ " "+ str(RivalChampionsInfo[matchup].titanPower)+ " "+ RivalChampionsInfo[matchup].titanLocation + attackNote)
			FMEChampionsInfo[optimalChamp].assignAttacks("Titan", matchup, RivalChampionsInfo[matchup].titanPower, RivalChampionsInfo[matchup].titanLocation, attacksNeeded, attackNote)
			RivalChampionsInfo[matchup].titanCleared = positionCleared

def calculateHeroAttacks():
	print("\n\nHero Attacks:")
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
				FMEChampionsInfo[FMEchamp].attacksRemaining > 0 and RivalChampionsInfo[matchup].heroCleared == False):
				# Check to see if there is a clear advantage number wise
				# print("Considering "+FMEchamp+" vs " + matchup)
				OptimizingAttack = True
				if optimalChamp == "":
					optimalChamp = FMEchamp
					# print("Optimal Champ: "+ optimalChamp)
				elif (FMEChampionsInfo[FMEchamp].heroPower < FMEChampionsInfo[optimalChamp].heroPower):
					optimalChamp = FMEchamp

			elif (not OptimizingAttack and FMEChampionsInfo[FMEchamp].attacksRemaining > 0 and RivalChampionsInfo[matchup].heroCleared == False):
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
			print(optimalChamp+"(H:"+str(FMEChampionsInfo[optimalChamp].heroPower)+ ") "+matchup+ " "+ str(RivalChampionsInfo[matchup].heroPower)+ " "+ RivalChampionsInfo[matchup].heroLocation + attackNote)
			FMEChampionsInfo[optimalChamp].assign("Hero", matchup, RivalChampionsInfo[matchup].heroPower, RivalChampionsInfo[matchup].heroLocation, attackNote)
			RivalChampionsInfo[matchup].heroCleared = positionCleared
def main():
	calculateTitanAttacks(False)
	calculateHeroAttacks()
	if(categorizeAssignments):
		categories = {"Titan", "Hero"}
		for category in categories:
			print("\n\n"+ category +" Attacks:")
			for champ in FMEChampionsInfo.keys():
				if FMEChampionsInfo[champ].attacksRemaining < 2:
					FMEChampionsInfo[champ].printAssignmentCategory(champ, category)

	else:	
		print("\n\nAssignments:")
		for champ in FMEChampionsInfo.keys():
			if FMEChampionsInfo[champ].attacksRemaining < 2:
				FMEChampionsInfo[champ].printAssignments(champ)
	


	print("\n\nRemaining Attackers:")	
	# for hero in FMEChampAttacks.keys():
	for hero in FMEChampionsInfo.keys():
		if FMEChampionsInfo[hero].attacksRemaining > 0:
			print(hero, "(H:"+str(FMEChampionsInfo[hero].heroPower)+",T:"+str(FMEChampionsInfo[hero].titanPower)+") has "+ str(FMEChampionsInfo[hero].attacksRemaining)+" attacks remaining.")

	print("\n\nRemaining Targets:")
	# for target in RivalChampHeroesAndLocations:
	for target in RivalChampionsInfo.keys():
		# print(target)
		if RivalChampionsInfo[target].heroCleared == False :
			# print(RivalChampHeroesAndLocations[target])
			print(target+" H:"+str(RivalChampionsInfo[target].heroPower)+" at "+ RivalChampionsInfo[target].heroLocation)
		# if RivalChampTitansAndLocations[target][RivalTitanLocationNeedsClearing] == True :
		if RivalChampionsInfo[target].titanCleared == False :
			# print(RivalChampTitansAndLocations[target])
			print(target+" T:"+str(RivalChampionsInfo[target].titanPower)+" at "+ RivalChampionsInfo[target].titanLocation)


main()
input('\nPress Enter to exit')