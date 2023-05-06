''' GRAPPLE D  - michelle's auto grab of daily's from a job. Organizational app to reduce the clicks it takes to organize the submitted Dailies from the field.
install "python3 -m pip install pywin32"
'''
intro = "(c)2023 kaff1n8t3d\n\
      )  (\n\
     (   ) )\n\
      ) ( (\n\
    _______)_\n\
 .-|-(c)-----|  \n\
( C|  kaff1  |\n\
 '-|  n8t3d  |\n\
   '_________'\n\
    '-------'"

welcomeMsg = "   ______________________________\n\
 / \                             \.\n\
| O |                            |.\n\
 \_ |         GRAPPLE D.         |.\n\
    |                            |.\n\
    |    Grab a hold of daily    |.\n\
    |    reports and copy them   |.\n\
    |   into a designated folder.|.\n\
    |   With Optional CSV file   |.\n\
    |   export option.           |.\n\
    |                            |.\n\
    |   Select a menu option to  |.\n\
    |   begin.                   |.\n\
    |                            |.\n\
    |                            |.\n\
    |   _________________________|___\n\
    |  /                            /.\n\
    \_/____________________________/.\n"

''' Menu ''' #returns defaultOptionChoices(recurse(bool),outputStyl(bool),ext(str),csvOption(bool)),search_path(str),search_term(str),save_path(str)
menuMain = "\n ------------MAIN MENU------------\n 1.Default Settings\n 2.Choose Search Directory\n 3.Choose Search Term\n 4.Choose Save Directory\n\n 9.Exit GrappleD\n ---------------------------------\n"
menuMain_S = "\n ------------MAIN MENU------------\n 1.Default Settings\n 2.Choose Search Directory\n 3.Choose Search Term\n 4.Choose Save Directory\n\n 5.Search Now\n\n 9.Exit GrappleD\n ---------------------------------\n"
menuDefaults = "\n ------------Default Settings MENU------------\n 1.Recursive Search Setting?\n 2.File Extension\n 3.Output CSV?\n\n 9.Back to Main Menu\n ---------------------------------------------\n"# 2.Output Style\n 
recurse = 1 #recursive search 0 single folder search only; 1 Recursive search starting at TOP level folder
outputStyl = 0 #ouput style - 0 copies the file; 1 creates a shortcut to the file.
fileExt = 'pdf'
ext = f'{fileExt}'
csvOption = 0 # Creates a CSV of Found Files w/links to thier location - 0 Does not create; 1 Creates CSV
search_path, search_term, save_path = "", "", "" 

def checkBackSlash(strVariable):
    bkslsh = '\\'
    if strVariable[len(strVariable)-1] != bkslsh:
        strVariable = f'{strVariable}{bkslsh}'
    return strVariable
    
def MENUmain(menuMain,menuDefaults,recurse,outputStyl,ext,csvOption,search_path,search_term,save_path):
    Mainloop = 777
    menuLoop = 1
    defaultOptionChoices = [recurse,outputStyl,ext,csvOption]
    while menuLoop == 1:
        ''' MainMenu choice options -- Display's Search when ready. '''
        if search_path != "" and search_term != "" and save_path !="":
            menuMainChoice =int(input(menuMain_S))# MainMenu ready to search Choice
        else:
            menuMainChoice =int(input(menuMain))# MainMenu Choice
        
        if menuMainChoice == 9: #Exit App
            print('\n---Exiting Application---\\')
            menuLoop = 0
            Mainloop = 'exit'
            return Mainloop
            
        elif menuMainChoice == 1: #Default Menu
            defaultOptionChoices = MENUdefaults(menuDefaults,recurse,outputStyl,ext,csvOption)
            
        elif menuMainChoice == 2: #input Search Directory
            search_path_loop = 1
            while search_path_loop == 1:
                search_path = str(input('\n Enter the file path to the top level search directory (i.e. which folder to begin searching in).\n'))
                # check for trailing backslash
                search_path = checkBackSlash(search_path)
                
                pathCheckresponse = input(f'\n You have entered "{search_path}", is this correct?\n (Y for yes, or N for no.)\n')
                if pathCheckresponse == "n" or pathCheckresponse == "N" or pathCheckresponse == "no" or pathCheckresponse == "No" or pathCheckresponse == "NO":
                    print(' Okay, let\'s try again.')
                    search_path_loop = 1
                elif pathCheckresponse == "y" or pathCheckresponse == "Y" or pathCheckresponse == "yes" or pathCheckresponse == "YES" or pathCheckresponse == "Yes":
                    search_path_loop = 0
        
        elif menuMainChoice == 3: #input Search Term
            search_term_loop = 1
            while search_term_loop == 1:
                search_term = str(input('\n Enter the term to search for (i.e. "daily", or "notes").\n'))
                termCheckresponse = input(f'\n You have entered "{search_term}", is this correct?\n (Y for yes, or N for no.)\n')
                if termCheckresponse == "n" or termCheckresponse == "N" or termCheckresponse == "no" or termCheckresponse == "No" or termCheckresponse == "NO":
                    print(' Okay, let\'s try again.')
                    search_term_loop = 1
                elif termCheckresponse == "y" or termCheckresponse == "Y" or termCheckresponse == "yes" or termCheckresponse == "YES" or termCheckresponse == "Yes":
                    search_term_loop = 0
                    
        elif menuMainChoice == 4: #input Save Directory
            save_path_loop = 1
            while save_path_loop == 1:
                save_path = str(input('\n Enter the file path to the save directory (i.e. which folder to save the files/shortcuts in.).\n'))
                # check for trailing backslash
                save_path = checkBackSlash(save_path)
                
                pathChkresponse = input(f'\n You have entered "{save_path}", is this correct?\n (Y for yes, or N for no.)\n')
                if pathChkresponse == "n" or pathChkresponse == "N" or pathChkresponse == "no" or pathChkresponse == "No" or pathChkresponse == "NO":
                    print(' Okay, let\'s try again.')
                    save_path_loop = 1
                elif pathChkresponse == "y" or pathChkresponse == "Y" or pathChkresponse == "yes" or pathChkresponse == "YES" or pathChkresponse == "Yes":
                    save_path_loop = 0
                    
        elif menuMainChoice == 5: #Search Now
            Mainloop = 777
            return Mainloop,defaultOptionChoices,search_path,search_term,save_path
        else:
            menuLoop = 0
            Mainloop = 1
    
def MENUdefaults(menuDefaults,recurse,outputStyl,ext,csvOption):
    
    menueloop = 1
    while menueloop == 1:
        menuDefaultsChoice = int(input(menuDefaults))
        if menuDefaultsChoice == 9:#Cancel
            print('\n---returning to main menu---\n')
            menueloop = 0
            break
            
        elif menuDefaultsChoice == 1:#recursive option
            if recurse == 1:#convert bool to text
                r = 'ON'
            elif recurse == 0:
                r = 'OFF'
            print(f'\n The Default recursive search pattern is {r}.')
            recurse = int(input('\n...to modify, choose 0 for OFF or 1 for ON.\n Or press 9 to Continue.\n '))
            if recurse == 9:
                if r == 'OFF':
                    recurse = 0
                elif r == 'ON':
                    recurse = 1
                continue
            elif recurse == 1:
                print(' Recursive option set to ON.\n')
            elif recurse == 0:
                print(' Recursive option set to Off.\n')
            elif recurse != 0 or recurse != 1:
                recurse = 1
                print(' Recursive option set to ON.\n')
            menueloop = 1 #Stay in defaults menu
        
        # elif menuDefaultsChoice == 999:#Output style -- Shortcut or copied file
            # if outputStyl == 1:#convert bool to text
                # o = 'to make a Shortcut Link'
            # elif outputStyl == 0:
                # o = 'to make a Copy of the File'
            # print(f'\n The Default Output Style is set {o}.')
            # outputStyl = int(input('\n...to modify, choose 0 to Copy the File to the new directory, or choose 1 to create a Shortcut Link in the new directory.\n Or press 9 to continue.\n '))
            # while outputStyl == 9:#convert text back to bool
                # if o == 'to make a Copy of the File':
                    # outputStyl = 0
                # elif o == 'to make a Shortcut Link':
                    # outputStyl = 1
            # if outputStyl == 1:
                # print(' The Output Style is set to Create a Shortcut Link to the file.\n')
            # elif outputStyl == 0:
                # print(' The Output Style is set to Copy the File to your Destination.\n')
            # elif outputStyl != 0 or recurse != 1:
                # outputStyl = 1
                # print(' The Output Style is set to Create a Shortcut Link to the file.\n')
            
            # menueloop = 1 #Stay in defaults menu
            
        elif menuDefaultsChoice == 2: #File Extension
            EXTloop = 1
            while EXTloop == 1:
                print(f'\n The Default File Extension to Search For is set to {ext}.')
                Fext = str(input('\n...to modify, type in a new file extension,\n or enter 9 to continue.\n'))
                if Fext == "9":
                    EXTloop = 0
                    continue
                else:
                    Fext = Fext.replace(".","")
                    res = input(f' You entered {Fext}, is this correct?\n (Y for yes, N for no.)\n')
                    if res == "Y" or res == "y":
                        ext =f'{Fext}'
                        print(f' The Default File Extension to Search For is set to {ext}.')
                        EXTloop = 0
                    elif res == "N" or res == "n":
                        print('\n Okay, let\'s start again.')
                        EXTloop = 1
            menueloop = 1 #Stay in defaults menu
            
        elif menuDefaultsChoice == 3:#csv option
            if csvOption == 1:#convert bool to text
                csvO = 'ON'
            elif csvOption == 0:
                csvO = 'OFF'
            print(f'\n The Default CSV Option is set to {csvO}.')
            csvOption = int(input('\n...to create a CSV file with links to each file, choose 1 for ON or choose 0 for OFF.\n press 9 to Continue without any changes.\n '))
            if csvOption == 9:
                continue
            elif csvOption == 1:
                print(' CSV file option set to ON.\n')
            elif csvOption == 0:
                print(' CSV file option set to Off.\n')
            elif csvOption != 0 or csvOption != 1:
                csvOption = 1
                print(' CSV file option set to ON.\n')   
            menueloop = 1
            
        else:
            menueloop = 1
    return recurse,outputStyl,ext,csvOption

def writeCSV(resultLine,save_path,search_term):
    from datetime import datetime
    now = datetime.now()
    now = now.strftime("%Y-%m-%d %H.%M")#add for .%S later update when file is created.
    CSVfile = f'{save_path}{search_term}_SearchResults_{now}.csv'
    resLineSplit = resultLine.split("\\")
    LineFileName = resLineSplit[len(resLineSplit)-1]
    f = open(CSVfile, "a+t")
    f.write(f'{LineFileName},"{resultLine}",\n')
    f.close()

def copyFiles(file,save_path):
    import shutil
    shutil.copy2(file, save_path)
    #print(' Files Copied')

def createShortcuts(file,save_path):#in testing
#https://stackoverflow.com/questions/26986470/create-shortcut-files-in-windows-7-using-python
    import win32com.client
    import pythoncom
    import os
    #file = r"C:\Users\kaff1n8t3d\Desktop\test\465107_invoice.pdf"
    fileSplit = file.split("\\")
    print(fileSplit)
    fileName = fileSplit[len(fileSplit)-1]
    print(fileName)
    fileSplit = fileName.split(".")
    print(fileSplit)
    fileName = f'{fileSplit[0]}.lnk'
    print(fileName)
    # pythoncom.CoInitialize() # remove the '#' at the beginning of the line if running in a thread.
    path = os.path.join(save_path, fileName)
    print(path)
    targetShortcut = file #make this the total file path

    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = file
    shortcut.WindowStyle = 1 # 7 - Minimized, 3 - Maximized, 1 - Normal
    shortcut.save()

''' Main Func '''
def get_file_paths(SEARCHVALUES):#(recurse(bool),outputStyl(bool),ext(str),csvOption(bool)),search_path(str),search_term(str),save_path(str)
    recurse = SEARCHVALUES[1][0]#recurse(bool)
    outputStyl = SEARCHVALUES[1][1]#outputStyl(bool)
    ext = SEARCHVALUES[1][2]#ext(str)
    csvOption = SEARCHVALUES[1][3]#csvOption(bool)
    search_path = SEARCHVALUES[2]#search_path(str)
    search_term = SEARCHVALUES[3]#search_term(str)
    save_path = SEARCHVALUES[4]#save_path(str)
    import glob
    #print('Get File Paths')
    # search all files inside a specific folder
    if recurse == 0:
        #print('Not In recursive')
        target = f'{search_path}*{search_term}*.{ext}'
        #print(target)
        res = glob.glob(target)
        for r in res:
            print(r)
            if csvOption == 1:#Write CSV
                writeCSV(r,save_path,search_term)

            if outputStyl == 0:#move file
                copyFiles(r,save_path)
                #print('in copyFiles')
            else:
                continue
    # Search all files/folders inside a specific folder
    if recurse == 1:
        #print('In recursive')
        target = f'{search_path}**\\*{search_term}*.{ext}'
        for file in glob.glob(target, recursive=True):
            print(file)
            if file.find('DailyReportPkg') > 0:
                print('File Not included - package submit daily.')
            else:
                if csvOption == 1:#Write CSV
                    writeCSV(file,save_path,search_term)
                    
                if outputStyl == 0:#move file
                    copyFiles(file,save_path)      
                else:
                    continue
                
            

if __name__ == "__main__":  
    print(intro)
    print(welcomeMsg)
    Mainloop = 777
    while Mainloop == 777:
        SEARCHVALUES = MENUmain(menuMain,menuDefaults,recurse,outputStyl,ext,csvOption,search_path,search_term,save_path)
        if SEARCHVALUES != 'exit':
            get_file_paths(SEARCHVALUES)
            input("\n Files copied press any key to continue...")
        else:
            Mainloop = 0
    