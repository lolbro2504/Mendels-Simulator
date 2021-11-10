from openpyxl import*
import random

HomozygousDominantParent = ["D", "D"]
HomozygousRecessiveParent = ["d", "d"]

workbook = load_workbook(filename="Presentation.xlsx")
sheet = workbook.active

attempts = 30000

f1Children = []

#recreates the breeding process
def createOffsprings(Parent1, Parent2,isf1):
    child = []
    childAllele1 = Parent1[random.randrange(0,2)]
    childAllele2 = Parent2[random.randrange(0,2)]
    child.append(childAllele1)
    child.append(childAllele2)
    if isf1:
        f1Children.append(child)
    phenotype = CheckPhenotype(child)
    return phenotype

#checks if there is any dominant alleles
def CheckPhenotype(Children):
    for i in Children:
        if i == "D":
            return True
    return False


row = 4
for i in range(attempts):
    #reset all variables
    f1Children = []
    
    DominantsF1 = 0
    RecessivesF1 = 0
    
    DominantsF2 = 0
    RecessivesF2 = 0

    #F1
    repetitions = range(random.randrange(60,90))
    for i in repetitions:
        f1 = True
        offspringphenotype = createOffsprings(HomozygousDominantParent, HomozygousRecessiveParent, f1)
        if offspringphenotype:
            DominantsF1 += 1
        else:
            RecessivesF1 +=1


    #F2
    repetitions = range(random.randrange(100,200))
    for i in repetitions:
        f1 = False
        f2parent1 = f1Children[random.randrange(0,len(f1Children)-1)]
        f2parent2 = f1Children[random.randrange(0,len(f1Children)-1)]
        offspringphenotype = createOffsprings(f2parent1, f2parent2, f1)
        if offspringphenotype:
            DominantsF2 += 1
        else:
            RecessivesF2 +=1


    #Write on the Excel
    sheet["C"+str(row)] = str(DominantsF1)
    sheet["E"+str(row)] = str(RecessivesF1)
    sheet["G"+str(row)] = str(DominantsF2)
    sheet["I"+str(row)] = str(RecessivesF2)
    print(DominantsF1, RecessivesF1, DominantsF2, RecessivesF2)
    row +=1

workbook.save(filename="Presentation.xlsx")