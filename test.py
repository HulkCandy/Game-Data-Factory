
EGM_TEMPLATE = r'C:\Users\ALiu\Desktop\Andy\16. Projects\Application\EGM-TEST\EGM\test all\Data Templates - EGM - WK 34.xlsx'
player_temp = r'C:\Users\ALiu\Desktop\Andy\16. Projects\Application\EGM-TEST\EGM\test all\Player_Temp.xlsx'
player_temp_txt = r'C:\Users\ALiu\Desktop\Andy\16. Projects\Application\EGM-TEST\EGM\test all\Player_temp.txt'
ooo = r'C:\Users\ALiu\Desktop\Andy\16. Projects\Application\EGM-TEST\EGM\test all\111.txt'
ttt = r'C:\Users\ALiu\Desktop\Andy\16. Projects\Application\EGM-TEST\EGM\test all\222.txt'
machine_data = r'C:\Users\ALiu\Desktop\Andy\16. Projects\Application\EGM-TEST\EGM\test all\EGM - Loyalty Data.txt'

inputFile = ooo
# Opening the given file in read-only mode.
readFile = open(inputFile, "r")

# output text file path
outputFile =ttt
# Opening the output file in write mode.
outFile = open(outputFile, "r")

# Read the above read file lines using readlines()
ReadFileLines = readFile.readlines()
writeFile = open(outputFile, "a")
writeFile.write("\n")

# Traverse in each line of the read text file
for i in range(0, len(ReadFileLines)):
    writeFile.write(ReadFileLines[i])
    print(ReadFileLines[i])

# Closing the write file
writeFile.close()

# Closing the read file
readFile.close()