

import DocxManager as DocMan
import time

unitscreated = 0
imagecount = 0
inputfolder = 'PutPicturesHere!'
safety = [False, False]


print("**Launching PicScribe**")
time.sleep(.5)
print(".")
time.sleep(.5)
print(".")
time.sleep(.5)
print(".")
time.sleep(.5)
print("Welcome, Teacher. I'm the PicScribe. You give me pictures, and I make '.docx' assignment papers out of them.")
time.sleep(3)


while safety[0] is not True:
    print("How many *units* do you need made?")
    unitsneeded = input()
    unitsneeded = int(unitsneeded)
    if unitsneeded is not str:
        imagecount = (unitsneeded) + (unitsneeded / 2)
        imagecount = round(imagecount)
        safety[0] = True
    else:
        print("That's not a number... You sure you should be an English teacher?")
        print(".")
        time.sleep(.5)
        print(".")
        time.sleep(.5)
        print("try again.")
print(".")
time.sleep(1)

while safety[1] is not True:
    print("What is the Unit-Number (Unit 1,2 3...etc) of the *First* unit you need made?")
    iterationx = int(input())
    iterationy = iterationx + 1

    if iterationx and iterationy is not str:
        safety[1] = True
    else:
        print("That's not a number... You sure you should be an English teacher?")
        print(".")
        time.sleep(.5)
        print(".")
        time.sleep(.5)
        print("try again.")

print("You need *" + str(imagecount) + "* pictures in the " + inputfolder + ". Make sure they're in there or I won't finish." )
time.sleep(1)
print(".")
time.sleep(1)
print(".")
print("Pre-check:")
time.sleep(2)
print("1. Are the correct number in the folder? and are they named correctly? (read the Instructions file!!!)")
time.sleep(2)
print("2. Is Microsoft Word closed? ( PicScribe will crash if one of the documents being changed is open.)")
time.sleep(2)
input(".*.*.*.*.*Press any Key When Ready.*.*.*.*.*")
print("Beginning Operation.")

i1 =1
i2 =2
i3 =3

while unitscreated < unitsneeded:
    DocMan.piccheck(i1, i2, i3)
    DocMan.docxmaker(iterationx, iterationy, i1, i2, i3)
    unitscreated += 2
    print("Units Completed: " + str(unitscreated) + "/" + str(unitsneeded))
    iterationx += 2
    iterationy += 2
    i1 += 3
    i2 += 3
    i3 += 3

print("Asssignments complete. They aughtta be in the CreatedAssignments folder now!")
print("Have a good'n. It's now safe to close this window.")
print("Bye! :)")
input(".*.*.*.*.*.Press Enter to close this Window.*.*.*.*.*.")