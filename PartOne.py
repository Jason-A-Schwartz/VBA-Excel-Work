# Jason Schwartz
# CINF 308 Midterm Project Part One

# Description: 
# This program prints out the correct syntax for VBA and the correct data values for the first set of 25
# I decided to break it up into 6 programs to make it easier on the debbuger and because the cell constants changing after every 25
# However, each program needs 2 parts to run the way I wanted it to in VBA

# Part One:
# Prints out the first half of the code which is joined by the other 6 parts
# This code sets a specific section of cells equal to another specific section

# This creates and keeps track of the 'V' cell value which starts at 11 so I started it at 9 so when it adds two it will display the correct value from the first iteration and so on
Vcount = 9

# This creates and keeps track of the 'EF', 'EG', 'FA1', 'FO1' cell values which starts at 75 so I started it at 73 so when it adds two it will display the correct value from the first iteration and so on
EFGAOcount = 73

# This creates and keeps track of the 'FA2' and 'FO2' cell values which starts at 76 so I started it at 4 so when it adds two it will display the correct value from the first iteration and so on
OFFcountone = 74

# This creates and keeps track of the 'H1', 'I1','I2', 'I4' cell values which starts at 10 so I started it at 8 so when it adds two it will display the correct value from the first iteration and so on
HIcount = 8

# This creates and keeps track of the 'I3' and 'I5' cell values which starts at 11 so I started it at 9 so when it adds two it will display the correct value from the first iteration and so on
OFFcounttwo = 9

# This creates and keeps track of the count for the program
Countercount = 0

# This starts the wile loop based off the EFGAOcount to make sure it only prints 25 times
while (EFGAOcount < 123):

    # This adds one to the Countercount variable
    Countercount = Countercount + 1

    # This adds two to the Vcount variable
    Vcount = Vcount + 2

    # This adds two to the EFGAOcount variable
    EFGAOcount = EFGAOcount + 2

    # This adds two to the OFFcountone variable
    OFFcountone = OFFcountone + 2

    # This adds two to the HIcount variable
    HIcount = HIcount + 2

    # This adds two to the Offcounttwo variable
    OFFcounttwo = OFFcounttwo + 2

#This is the begining of the formatting needed for the VBA code

    # This spaces out the section of code from the rest
    print()

    # This prints out the Countercount variable with the correct VBA syntax
    print("'" + str(Countercount))

    # This prints out the Vcount variable with the correct VBA syntax
    print('If Range("V' + str(Vcount) + '") = "" Then')

    # This creates a space for the correct VBA syntax
    print()

    # This prints out the EFGAOcount variable and the HIcount variable with the correct VBA syntax
    print('Range("EF' + str(EFGAOcount) + '").Value = Sheets("Kalamein Shop").Range("H' + str(HIcount) + '")')

    # This prints out the EFGAOcount variable and the HIcount variable with the correct VBA syntax
    print('Range("EG' + str(EFGAOcount) +'").Value = Sheets("Kalamein Shop").Range("I' + str(HIcount) + '")')

    # This prints out the EFGAOcount variable and the HIcount variable with the correct VBA syntax
    print('Range("FA' + str(EFGAOcount) +'").Value = Sheets("Assembly").Range("I' + str(HIcount) + '")')

    # This prints out the OFFcountone variable and the OFFcounttwo variable with the correct VBA syntax
    print('Range("FA' + str(OFFcountone) +'").Value = Sheets("Assembly").Range("I' + str(OFFcounttwo) + '")')

    # This prints out the EFGAOcount variable and the HIcount variable with the correct VBA syntax
    print('Range("FO' + str(EFGAOcount) +'").Value = Sheets("Prehang").Range("I' + str(HIcount) + '")')

    # This prints out the OFFcountone variable and the OFFcounttwo variable with the correct VBA syntax
    print('Range("FO' + str(OFFcountone) +'").Value = Sheets("Prehang").Range("I' + str(OFFcounttwo) + '")')

    # This creates a space for the correct VBA syntax
    print()

    # This prints out the Vcount variable with the correct VBA syntax
    print('ElseIf Range("V' + str(Vcount) + '") <> "" Then')

    # This creates a space for the correct VBA syntax
    print()

    # This prints out the EFGAOcount variable with the correct VBA syntax
    print('Range("EF' + str(EFGAOcount) + '").Value = ""')

    # This prints out the EFGAOcount variable with the correct VBA syntax
    print('Range("EG' + str(EFGAOcount) + '").Value = ""')

    # This prints out the EFGAOcount variable with the correct VBA syntax
    print('Range("FA' + str(EFGAOcount) + '").Value = ""')

    # This prints out the OFFcountone variable with the correct VBA syntax
    print('Range("FA' + str(OFFcountone) + '").Value = ""')

    # This prints out the EFGAOcount variable with the correct VBA syntax
    print('Range("FO' + str(EFGAOcount) + '").Value = ""')

    # This prints out the OFFcountone variable with the correct VBA syntax
    print('Range("FO' + str(OFFcountone) + '").Value = ""')

    # This creates a space for the correct VBA syntax
    print()

    # This ends the VBA code with the correct syntax
    print('End If')

    # This creates a space for the correct VBA syntax
    print()

# Part Two:
# Prints out the second half of the code which is applied to the bottom of the VBA program after all the other part ones have been completed
# This code sets a specific section of cells equal to another specific section 

# This keeps track of the 'EF', 'EG', 'FA1', 'FO1' cell values which starts at 75 so I started it at 73 so when it adds two it will display the correct value from the first iteration and so on
EFGAOcount = 73

# This creates and keeps track of the 'FA2' and 'FO2' cell values which starts at 76 so I started it at 4 so when it adds two it will display the correct value from the first iteration and so on
OFFcountone = 74

# This creates and keeps track of the 'H1', 'I1','I2', 'I4' cell values which starts at 10 so I started it at 8 so when it adds two it will display the correct value from the first iteration and so on
HIcount = 8

# This creates and keeps track of the 'I3' and 'I5' cell values which starts at 11 so I started it at 9 so when it adds two it will display the correct value from the first iteration and so on
OFFcounttwo = 9

# This creates and keeps track of the count for the program
Countercount = 0

# This starts the wile loop based off the EFGAOcount to make sure it only prints 25 times
while (EFGAOcount < 123):

    # This adds one to the Countercount variable
    Countercount = Countercount + 1

    # This adds two to the EFGAOcount variable
    EFGAOcount = EFGAOcount + 2

    # This adds two to the OFFcountone variable
    OFFcountone = OFFcountone + 2

    # This adds two to the HIcount variable
    HIcount = HIcount + 2

    # This adds two to the Offcounttwo variable
    OFFcounttwo = OFFcounttwo + 2

    # This creates a space for the correct VBA syntax
    print()

    # This prints out the Countercount variable with the correct VBA syntax
    print("'" + str(Countercount))

    # This prints out the HIcount variable and the EFGAOcount variable with the correct VBA syntax
    print('Sheets("Kalamein Shop").Range("H' + str(HIcount) + '").Value = Range("EF' + str(EFGAOcount) + '")')

    # This prints out the HIcount variable and the EFGAOcount variable with the correct VBA syntax
    print('Sheets("Kalamein Shop").Range("I' + str(HIcount) + '").Value = Range("EG' + str(EFGAOcount) + '")')

    # This prints out the HIcount variable and the EFGAOcount variable with the correct VBA syntax
    print('Sheets("Assembly").Range("I' + str(HIcount) + '").Value = Range("FA' + str(EFGAOcount) + '")')

    # This prints out the OFFcounttwo variable and the OFFcountone variable with the correct VBA syntax
    print('Sheets("Assembly").Range("I' + str(OFFcounttwo) + '").Value = Range("FA' + str(OFFcountone) + '")')

    # This prints out the HIcount variable and the EFGAOcount variable with the correct VBA syntax
    print('Sheets("Prehang").Range("I' + str(HIcount) + '").Value = Range("FO' + str(EFGAOcount) + '")')

    # This prints out the OFFcounttwo variable and the OFFcountone variable with the correct VBA syntax
    print('Sheets("Prehang").Range("I' + str(OFFcounttwo) + '").Value = Range("FO' + str(OFFcountone) + '")')