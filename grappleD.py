''' GRAPPLE D  - michelle's auto grab of daily's from a job. Organizational app to reduce the clicks it takes to organize the submitted Dailies from the field.'''

def menuoptions(intro,welcomeMsg):
    print(intro)
    #print(welcomeMsg)
    
    if choice == 1:
        print(1)
    elif choice == 2:
        print(2)
    elif choice == 3:
        print(3)
    else:
        print('other')
    
    return chioceAction

def menu():
    loop = 1
    while lo op == 1:

    
''' Main Func '''
def get_file_paths(search_path,ext,recursive):
    import glob
    # search all files inside a specific folder
    if recursive == 0:
        target = f'{search_path}*AMAZON*{ext}'
        #print(target)
        res = glob.glob(target)
        for r in res:
            print(r)
    # Search all files/folders inside a specific folder
    if recursive == 1:
        target = f'{search_path}**\\*invoice*{ext}'
        for file in glob.glob(target, recursive=True):
            print(file)
ext = ".pdf"


if __name__ == "__main__":
    get_file_paths(r'C:\Users\cleeh\Documents\Google Drive\\', ext, 1)