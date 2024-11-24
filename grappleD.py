''' GRAPPLE D  - michelle's auto grab of daily's from a job. Organizational app to reduce the clicks it takes to organize the submitted Dailies from the field.
install "python3 -m pip install pywin32"
'''
import time
intro = "\n\n\n(c)2023 kaff1n8t3d\n\
      )  (\n\
     (   ) )\n\
      ) ( (\n\
    _______)_\n\
 .-|-(c)-----|  \n\
( C|  kaff1  |\n\
 '-|  n8t3d  |\n\
   '_________'\n\
    '-------'\n\n"

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
menuDefaults = "\n ------------Default Settings MENU------------\n 1.Recursive Search Setting?\n 2.File Extension\n 3.Output CSV?\n 4.Exclusion?\n\n 9.Back to Main Menu\n ---------------------------------------------\n"# 2.Output Style\n 
recurse = 1 #recursive search 0 single folder search only; 1 Recursive search starting at TOP level folder
outputStyl = 0 #ouput style - 0 copies the file; 1 creates a shortcut to the file.*not working yet
fileExt = 'pdf'
ext = f'{fileExt}'
csvOption = 0 # Creates a CSV of Found Files w/links to thier location - 0 Does not create; 1 Creates CSV
search_path, search_term, save_path = "", "", "" 
exclusion = 'DailyReportPkg' #string in file to exclude file if found

def checkBackSlash(strVariable):
    bkslsh = '\\'
    if strVariable[len(strVariable)-1] != bkslsh:
        strVariable = f'{strVariable}{bkslsh}'
    return strVariable
    
def MENUmain(menuMain,menuDefaults,recurse,outputStyl,ext,csvOption,search_path,search_term,save_path,exclusion):
    Mainloop = 777
    menuLoop = 1
    defaultOptionChoices = [recurse,outputStyl,ext,csvOption,exclusion]
    while menuLoop == 1:
        ''' MainMenu choice options -- Display's Search when ready. '''
        if search_path != "" and search_term != "" and save_path !="":
        
            loopIt = 1#check for key error. DONT CRASH
            while loopIt == 1:
                try:
                    menuMainChoice =int(input(menuMain_S))# MainMenu ready to search Choice
                    loopIt = 0
                except ValueError:  
                    print(f'\n Incorrect input, let\'s try again.')
                    loopIt = 1
                    time.sleep(1)
                except KeyboardInterrupt:
                    print(f'\n Incorrect input, let\'s try again.')
                    loopIt = 1
                    time.sleep(1)
                finally:
                    if menuMainChoice != 9 and menuMainChoice != 1 and menuMainChoice != 2 and menuMainChoice != 3 and menuMainChoice != 4 and menuMainChoice != 5:
                        print(f'\n Incorrect input, let\'s try again.')
                        time.sleep(1)
                        loopIt = 1
            del loopIt          
        else:
            loopIt = 1#check for key error. DONT CRASH
            while loopIt == 1:
                try:
                    menuMainChoice = int(input(menuMain))# MainMenu Choice
                    loopIt = 0
                except ValueError:  
                    print(f'\n Incorrect input, let\'s try again.')
                    loopIt = 1
                    menuMainChoice = ""
                    time.sleep(1)
                except KeyboardInterrupt:
                    print(f'\n Incorrect input, let\'s try again.')
                    loopIt = 1
                    menuMainChoice = ""
                    time.sleep(1)
                finally:
                    if menuMainChoice != 9 and menuMainChoice != 1 and menuMainChoice != 2 and menuMainChoice != 3 and menuMainChoice != 4:
                        if loopIt != 1:
                            print(f'\n Incorrect input, let\'s try again.')
                            time.sleep(1)
                            loopIt = 1
            del loopIt   
        
        if menuMainChoice == 9: #Exit App
            print('\n/---Exiting Application---\\')
            time.sleep(1)
            menuLoop = 0
            Mainloop = 'exit'
            return Mainloop
            
        elif menuMainChoice == 1: #Default Menu
            defaultOptionChoices = MENUdefaults(menuDefaults,defaultOptionChoices[0],defaultOptionChoices[1],defaultOptionChoices[2],defaultOptionChoices[3],defaultOptionChoices[4])

        elif menuMainChoice == 2: #input Search Directory
            search_path_loop = 1
            while search_path_loop == 1:
            
                loopIt = 1#check for key error. DONT CRASH
                while loopIt == 1:
                    if search_path != "":#set message
                        search_path_msg = f'\nThe current search path is set to {search_path}.\n\n To change it: Enter the file path to the top level search directory (i.e. which folder to begin searching in).\n...Or press 9 to erase and exit.\n'
                    else:
                        search_path_msg = '\n Enter the file path to the top level search directory (i.e. which folder to begin searching in).\n...Or press 9 to erase and exit.\n'
                        
                    try:
                        search_path = str(input(search_path_msg))
                        loopIt = 0
                    except ValueError:  
                        print(f'\n Incorrect input, let\'s try again.')
                        loopIt = 1
                        time.sleep(1)
                    except KeyboardInterrupt:
                        print(f'\n Incorrect input, let\'s try again.')
                        loopIt = 1
                        time.sleep(1)
                    finally:
                        if search_path == " " or search_path == "":
                            if loopIt != 1:
                                print(f'\n Incorrect input, let\'s try again.')
                                time.sleep(1)
                            loopIt = 1
                del loopIt

                if search_path == "9":
                    print(f'\n---Exiting---\nSearch Directory Reset.\n')
                    time.sleep(1)
                    search_path  = ""
                    break
                # check for trailing backslash
                if search_path !="":
                    search_path = checkBackSlash(search_path)
                
                loopIt = 1#check for key error. DONT CRASH
                while loopIt == 1:
                    try:
                        pathCheckresponse = input(f'\n You have entered "{search_path}", is this correct?\n (Y for yes, or N for no.)\n')
                        loopIt = 0
                    except ValueError:  
                        print(f'\n Incorrect input, let\'s try again.')
                        time.sleep(1)
                    except KeyboardInterrupt:
                        print(f'\n Incorrect input, let\'s try again.')
                        time.sleep(1)
                del loopIt
                
                if pathCheckresponse == "n" or pathCheckresponse == "N" or pathCheckresponse == "no" or pathCheckresponse == "No" or pathCheckresponse == "NO":
                    print(' Okay, let\'s try again.')
                    search_path  = ""
                    time.sleep(1)
                    search_path_loop = 1
                elif pathCheckresponse == "y" or pathCheckresponse == "Y" or pathCheckresponse == "yes" or pathCheckresponse == "YES" or pathCheckresponse == "Yes":
                    search_path_loop = 0
                    print(f'\nSearch Directory is set to {search_path}.')
                    time.sleep(1)
        
        elif menuMainChoice == 3: #input Search Term
            search_term_loop = 1
            while search_term_loop == 1:
                loopIt = 1#check for key error. DONT CRASH
                while loopIt == 1:
                    if search_term != "":#set message
                        search_term_msg = f'\nThe current search term is set to {search_term}.\n\n To change it: Enter the term to search for (i.e. "daily", or "notes").\n...Or press 9 to erase and exit.\n'
                    else:
                        search_term_msg = '\n Enter the term to search for (i.e. "daily", or "notes").\n...Or press 9 to erase and exit.\n'
                        
                    try:
                        search_term = str(input(search_term_msg))
                        loopIt = 0
                    except ValueError:  
                        print(f'\n Incorrect input, let\'s try again.')
                        time.sleep(1)
                    except KeyboardInterrupt:
                        print(f'\n Incorrect input, let\'s try again.')
                        time.sleep(1)
                    finally:
                        if search_term == " " or search_term == "":
                            print(f'\n Incorrect input, let\'s try again.')
                            time.sleep(1)
                            loopIt = 1
                del loopIt
                if search_term == str(9):
                    search_term = ""
                    print(f'\n---Exiting---\nThe Search Term has been reset.\n')
                    time.sleep(1)
                    break
                loopIt = 1#check for key error. DONT CRASH
                while loopIt == 1:
                    try:
                        termCheckresponse = input(f'\n You have entered "{search_term}", is this correct?\n (Y for yes, or N for no.)\n')
                        loopIt = 0
                    except ValueError:  
                        print(f'\n Incorrect input, let\'s try again.')
                        time.sleep(1)
                    except KeyboardInterrupt:
                        print(f'\n Incorrect input, let\'s try again.')
                        time.sleep(1)
                del loopIt
                
                if termCheckresponse == "n" or termCheckresponse == "N" or termCheckresponse == "no" or termCheckresponse == "No" or termCheckresponse == "NO":
                    print(' Okay, let\'s try again.')
                    search_term = ""
                    time.sleep(1)
                    search_term_loop = 1
                elif termCheckresponse == "y" or termCheckresponse == "Y" or termCheckresponse == "yes" or termCheckresponse == "YES" or termCheckresponse == "Yes":
                    search_term_loop = 0
                    print(f'\nSearch Term is set to {search_term}.')
                    time.sleep(1)
                    
        elif menuMainChoice == 4: #input Save Directory
            save_path_loop = 1
            while save_path_loop == 1:
                loopIt = 1#check for key error. DONT CRASH
                while loopIt == 1:
                    if save_path != "":#set message
                        print(f'\nThe current file path to save in is set to {save_path}.\n\n To change it:')

                    try:
                        save_path = str(input('\n Enter the file path to the save directory (i.e. which folder to save the files/shortcuts in.).\n...Or press 9 to erase and exit.\n'))
                        loopIt = 0
                    except ValueError:  
                        print(f'\n Incorrect input, let\'s try again.')
                        time.sleep(1)
                        loopIt = 1
                    except KeyboardInterrupt:
                        print(f'\n Incorrect input, let\'s try again.')
                        time.sleep(1)
                        loopIt = 1
                    finally:
                        if save_path == " " or save_path == "":
                            if loopIt != 1:
                                print(f'\n Incorrect input, let\'s try again.')
                                time.sleep(1)
                            loopIt = 1
                del loopIt
                if save_path == str(9):
                    save_path = ""
                    print(f'\n---Exiting---\nThe Save Path has been reset.\n')
                    time.sleep(1)
                    break
                    
                # check for trailing backslash
                save_path = checkBackSlash(save_path)
                
                loopIt = 1#check for key error. DONT CRASH
                while loopIt == 1:
                    try:
                        pathChkresponse = input(f'\n You have entered "{save_path}", is this correct?\n (Y for yes, or N for no.)\n')
                        loopIt = 0
                    except ValueError:  
                        print(f'\n Incorrect input, let\'s try again.')
                        time.sleep(1)
                    except KeyboardInterrupt:
                        print(f'\n Incorrect input, let\'s try again.')
                        time.sleep(1)
                del loopIt
               
                if pathChkresponse == "n" or pathChkresponse == "N" or pathChkresponse == "no" or pathChkresponse == "No" or pathChkresponse == "NO":
                    print(' Okay, let\'s try again.')
                    save_path_loop = 1
                elif pathChkresponse == "y" or pathChkresponse == "Y" or pathChkresponse == "yes" or pathChkresponse == "YES" or pathChkresponse == "Yes":
                    save_path_loop = 0
                    print(f'\nSave Directory is set to {save_path}.')
                    time.sleep(1)
                    
        elif menuMainChoice == 5: #Search Now
            Mainloop = 777
            return Mainloop,defaultOptionChoices,search_path,search_term,save_path
        else:
            menuLoop = 0
            Mainloop = 1
    
def MENUdefaults(menuDefaults,recurse,outputStyl,ext,csvOption,exclusion):
    
    menueloop = 1
    while menueloop == 1:
        loopIt = 1#check for key error. DONT CRASH
        while loopIt == 1:
            try:
                menuDefaultsChoice = int(input(menuDefaults))
                loopIt = 0
            except ValueError:  
                print(f'\n Incorrect input, let\'s try again.')
                time.sleep(1)
            except KeyboardInterrupt:
                print(f'\n Incorrect input, let\'s try again.')
                time.sleep(1)
        del loopIt
        
        if menuDefaultsChoice == 9:#Cancel
            print('\n---returning to main menu---\n')
            time.sleep(1)
            menueloop = 0
            break
            
        elif menuDefaultsChoice == 1:#recursive option
            if recurse == 1:#convert bool to text
                r = 'ON'
            elif recurse == 0:
                r = 'OFF'
            
            loopIt = 1#check for key error. DONT CRASH
            while loopIt == 1:
                try:
                    print(f'\nThe Default recursive search pattern is {r}.')
                    recurse = int(input('\n...to modify, choose 0 for OFF or 1 for ON.\n Or press 9 to Continue.\n '))
                    loopIt = 0
                except ValueError:  
                    print(f'\nIncorrect input, let\'s try again.')
                    loopIt = 1
                    time.sleep(1)
                except KeyboardInterrupt:
                    print(f'\n Incorrect input, let\'s try again.')
                    loopIt = 1
                    time.sleep(1)
                finally:
                    if recurse != 0 and recurse != 1 and recurse != 9:
                        if loopIt != 1:
                            print(f'\nIncorrect input, let\'s try again.')
                        loopIt = 1
                        time.sleep(1)
            del loopIt
            
            if recurse == 9:
                if r == 'OFF':
                    recurse = 0
                    print('\n---Exiting---\nRecursive option set to Off.\nRecursive Option Unchanged.\n')
                    time.sleep(1)
                elif r == 'ON':
                    recurse = 1
                    print('\n---Exiting---\nRecursive option set to ON.\nRecursive Option Unchanged.\n')
                    time.sleep(1)
                continue
            elif recurse == 1:
                print('Recursive option set to ON.\n')
                time.sleep(1)
            elif recurse == 0:
                print('Recursive option set to Off.\n')
                time.sleep(1)
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
                
                loopIt = 1#check for key error. DONT CRASH
                while loopIt == 1:
                    try:
                        print(f'\nThe Default File Extension to Search For is set to {ext}.')
                        Fext = str(input('\n...to modify, type in a new file extension,\n or enter 9 to continue.\n'))
                        loopIt = 0
                    except ValueError:  
                        print(f'\nIncorrect input, let\'s try again.')
                        loopIt = 1
                        time.sleep(1)
                    except KeyboardInterrupt:
                        print(f'\nIncorrect input, let\'s try again.')
                        loopIt = 1
                        time.sleep(1)
                    finally:
                        if Fext == " ":
                            print(f'\nIncorrect input, let\'s try again.')
                            time.sleep(1)
                            loopIt = 1
                del loopIt
                
                if Fext == "9" or len(Fext) == 0:
                    print(f'\n---Exiting---\nFile Extension Option Unchanged.\nThe Default File Extension to Search For is set to {ext}.')
                    time.sleep(1)
                    EXTloop = 0
                    continue
                else:
                    Fext = Fext.replace(".","")
                    res = input(f' You entered {Fext}, is this correct?\n (Y for yes, N for no.)\n')
                    if res == "Y" or res == "y":
                        ext =f'{Fext}'
                        print(f' The Default File Extension to Search For is set to {ext}.')
                        time.sleep(1)
                        EXTloop = 0
                    elif res == "N" or res == "n":
                        print('\n Okay, let\'s start again.')
                        time.sleep(1)
                        EXTloop = 1
            menueloop = 1 #Stay in defaults menu
            
        elif menuDefaultsChoice == 3:#csv option
            if csvOption == 1:#convert bool to text
                csvO = 'ON'
            elif csvOption == 0:
                csvO = 'OFF'
            
            loopIt = 1#check for key error. DONT CRASH
            while loopIt == 1:
                try:
                    print(f'\nThe Default CSV Option is set to {csvO}.')
                    csvOption = int(input('\n...to create a CSV file with links to each file, choose 1 for ON or choose 0 for OFF.\n press 9 to Continue without any changes.\n '))
                    loopIt = 0
                except ValueError:  
                    print(f'\nIncorrect input, let\'s try again.')
                    loopIt = 1
                    time.sleep(1)
                except KeyboardInterrupt:
                    print(f'\nIncorrect input, let\'s try again.')
                    loopIt = 1
                    time.sleep(1)
                finally:
                    if csvOption != 0 and csvOption!= 1 and csvOption!= 9:
                        print(f'\nIncorrect input, let\'s try again.')
                        loopIt = 1
                        time.sleep(1)
            del loopIt
            
            if csvOption == 9:
                print(f'\n---Exiting---\nCSV Option Unchanged.\nThe Default CSV Option is set to {csvO}.')
                time.sleep(1)
                continue
            elif csvOption == 1:
                print('CSV file option set to ON.\n')
                time.sleep(1)
            elif csvOption == 0:
                print('CSV file option set to Off.\n')
                time.sleep(1)
            menueloop = 1

        elif menuDefaultsChoice == 4: #File Exclusion
            EXTloop = 1
            while EXTloop == 1:
                try:
                    exclusion
                except NameError:
                    exclusion = int(0)
                    
                loopIt = 1#check for key error. DONT CRASH
                while loopIt == 1:
                    try:
                        if exclusion != int(0):#alter message if Exclusion is set.
                            print(f'\n Exclude a file with specific text in the name? Exclusion currently set to "{exclusion}"')
                        else:
                            print(f'\n Exclude a file with specific text in the name? There is not an exclusion set.')
                        Fexclusion = str(input('\n...to add/edit an exclusion, type in text to exclude from the search, and it will be excluded from results.,\n or enter 9 to continue.\n'))
                        loopIt = 0
                    except ValueError:  
                        print(f'\n Incorrect input, let\'s try again.')
                        loopIt = 1
                        time.sleep(1)
                    except KeyboardInterrupt:
                        print(f'\n Incorrect input, let\'s try again.')
                        loopIt = 1
                        time.sleep(1)
                    finally:
                        if Fexclusion ==" ":
                            if loopIt != 1:
                                print(f'\n Incorrect input, let\'s try again.')
                                loopIt = 1
                                time.sleep(1)
                del loopIt
                if Fexclusion == "9":
                    print(f'\n---Exiting---\nExclusion Option Unchanged.\nExclusion currently set to "{exclusion}"')
                    time.sleep(1)
                    EXTloop = 0
                    continue
                else:
                    res = input(f' You entered "{Fexclusion}", is this correct?\n (Y for yes, N for no.)\n')
                    if res == "Y" or res == "y":
                        exclusion =f'{Fexclusion}'
                        print(f' The File to exclude from search is set to {Fexclusion}.')
                        time.sleep(1)
                        EXTloop = 0
                    elif res == "N" or res == "n":
                        print('\n Okay, let\'s start again.')
                        time.sleep(1)
                        EXTloop = 1
            menueloop = 1 #Stay in defaults menu
            
        else:
            menueloop = 1
    return recurse,outputStyl,ext,csvOption,exclusion

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
    #print(SEARCHVALUES)
    recurse = SEARCHVALUES[1][0]#recurse(bool)
    outputStyl = SEARCHVALUES[1][1]#outputStyl(bool)
    ext = SEARCHVALUES[1][2]#ext(str)
    csvOption = SEARCHVALUES[1][3]#csvOption(bool)
    exclusion = SEARCHVALUES[1][4] #exclusion term
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
            if exclusion == int(0):
                continue
            else:
                if file.find(exclusion) > 0:
                    print(f'File Not included, because the name includes {exclusion}.')
                else:
                    if csvOption == 1:#Write CSV
                        writeCSV(file,save_path,search_term)
                        
                    if outputStyl == 0:#move file
                        copyFiles(file,save_path)      
                    else:
                        continue
                
            

if __name__ == "__main__":  
    print(intro)
    time.sleep(1)
    print(welcomeMsg)
    Mainloop = 777
    while Mainloop == 777:
        SEARCHVALUES = MENUmain(menuMain,menuDefaults,recurse,outputStyl,ext,csvOption,search_path,search_term,save_path,exclusion)
        if SEARCHVALUES != 'exit':
            get_file_paths(SEARCHVALUES)
            loopIt = 1#check for key error. DONT CRASH
            while loopIt == 1:
                try:
                    input("\n Files copied press any key to continue...")
                    loopIt = 0
                except ValueError:  
                    print(f'\n Incorrect input, let\'s try again.')
                    loopIt = 0
                    time.sleep(1)
                except KeyboardInterrupt:
                    print(f'\n Incorrect input, let\'s try again.')
                    loopIt = 0
                    time.sleep(1)
            del loopIt
        else:
            Mainloop = 0
    