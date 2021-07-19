import pyautogui
import time
import pandas as pd


# SOME INITIAL TESTING AND NOTES FOR FUTURE TESTING
#print(pyautogui.size())
#screenWidth=1920
#screenHeight=1080  #(0,0) is top left corner of screen
#pyautogui.  movet0(x,y,duration=)    moverel(x,y,duration=)
#ie: pyautogui.moveRel(100,100,duration=.5)
#pyautogui.  position() returns pos of mouse   click(x,y)   scroll(#pixels2goUp)
# typewrite("string")  typewrite(["a","left","ctrlleft"]) <-for indep press.   hotkey('ctrl','a')<-for combos
#ie: pyautogui.hotkey('ctrl', 'shift', 'esc') opens task mngr

#read the excel doc containing the constituent list and corresponding ratings, convert to dataframe, selecting only the first headerrange entries(to segment/etc)
df = pd.read_excel('C://Users//XXXXXXX//Desktop//XXXXXXX.xls', sheet_name='Sheet2')
headerrange=54
dfheader=df.head(headerrange)
#print(dfheader)


def addratcheck():
    #checks if the add rating button is on screen
    addratloc = pyautogui.locateOnScreen("C:/Users/XXXXXXX/Desktop/XXXXXXX/addbutscsht.PNG",
                                           region=(0, 100, 250, 950), grayscale=True)
    if isinstance(addratloc,type(None)):
        return "not found"
    else:
        return "found"
        
def ratingddownclk():
    #clicks the rating button to un-minimize the details of the field
    ratddownloc = pyautogui.locateOnScreen("C:/Users/XXXXXX/Desktop/XXXXXXX/ratingbutscsht.PNG",
                                           region=(0, 100, 250, 700))
    ratdownptx = ratddownloc[0] + 10
    ratdownpty = ratddownloc[1] + 15
    ratddownxy = ratdownptx, ratdownpty
    #print(ratddownxy)
    pyautogui.click(ratddownxy)
    
def addratclick(presence):
    #if add rating button not found, click the Ratings dropdown, then click the add rating button
    if presence == "not found":
        time.sleep(1.5)
        ratingddownclk()
        time.sleep(4.5)
        addratloc = pyautogui.locateOnScreen("C:/Users/XXXXXX/Desktop/XXXXXXX/addbutscsht.PNG",
                                             region=(0, 100, 250, 950), grayscale=True)
        # FOR RED ABOVE: what if the add rating is still not found? had one error with that...was a nonconstituent
        addratptx = addratloc[0] + 10
        addratpty = addratloc[1] + 15
        addratxy = addratptx, addratpty
        #print(addratxy)
        pyautogui.click(addratxy)

    else:
        #if add rating button found, click it
        addratloc = pyautogui.locateOnScreen("C:/Users/XXXXXXX/Desktop/XXXXXXX/addbutscsht.PNG",
                                             region=(0, 100, 250, 950))
        addratptx = addratloc[0] + 10
        addratpty = addratloc[1] + 15

        addratxy = addratptx, addratpty
        #print(addratxy)
        time.sleep(.5)
        pyautogui.click(addratxy,duration=.5)

def searchiconclick():
    #find and click the search icon (button)
    searchloc = pyautogui.locateOnScreen("C:/Users/XXXXXXX/Desktop/XXXXXXX/searchicon.PNG",
                                           region=(550, 120, 1500, 190))
    searchptx = searchloc[0] + 5
    searchpty = searchloc[1] + 5
    searchxy = searchptx, searchpty
    #print(searchxy)
    pyautogui.click(searchxy,duration=.5)

def ratingxcheck():
    #check if the x button on the rating entry popup is on screen
    xloc = pyautogui.locateOnScreen("C:/Users/XXXXXXX/Desktop/XXXXXXX/ratingx.PNG",
                                         region=(310,120, 1600, 200))
    if isinstance(xloc, type(None)):
        return "not found"
    else:
        return "found"
        
def ratingxclick(presence):
    #if rating popup x is found, click it to close
    if presence == "found":
        time.sleep(.5)
        xloc = pyautogui.locateOnScreen("C:/Users/XXXXXXXX/Desktop/XXXXXX/ratingx.PNG",
                                        region=(310, 120, 1600, 200))
        xptx = xloc[0] + 5
        xpty = xloc[1] + 5
        ratingxxy = xptx, xpty
        #print(ratingxxy)
        pyautogui.click(ratingxxy, duration=.5)
    else:
        print("noexitnec")

def nxtsetup():
    # checks mouse movement and clicks into search bar for initial constituent entry
    pyautogui.moveTo(1,1,duration=.5)
    print("setup 1/3 -mouse calibrated")
    searchiconclick()
    print("setup 2/3 -searchIcon selected")
    time.sleep(.2)
    print("setup 3/3 -click function set")
    #pyautogui.typewrite("test123!@#") testing only, DontUse
    print("setup completed successfully, proceed with iteration.")
    
def setupfn():
    #gives you a sec to open tabs, get ready before running nxtsetup
    print("initializing...")
    time.sleep(2.69)  # nice
    nxtsetup()

def entrylooper():
    #loops through entries in excel list and adds the rating for each constituent, using time.sleep because I'm lazy and checking if the website was loaded sounded complicated. likely buggy with bag wifi lol.
    #time.sleeps are largely arbitrary and based on my guesses. they fail occasionally with poor connection.
    for index, row in dfheader.iterrows():
        print("starting: "+str(index), ':', row['NAME'], ':', row['PORCIENTO'], '%')
        nombre=row['NAME']
        probability=row['PORCIENTO']
        #print(nombre, probability) #yup it works
        pyautogui.typewrite(str(nombre))
        time.sleep(6.69)
        pyautogui.press('enter')
        time.sleep(8)

        addratclick(addratcheck())

        time.sleep(8)
        pyautogui.typewrite("i")
        pyautogui.press('tab')
        time.sleep(.45)
        pyautogui.typewrite("p")
        pyautogui.press('tab',presses=3,interval=.5)
        time.sleep(1)
        pyautogui.typewrite(str(probability))
        time.sleep(.45)
        pyautogui.press('tab',presses=2,interval=.4)
        pyautogui.press('enter')
        time.sleep(3)

        #ratingxclick(ratingxcheck())
        time.sleep(4)
        searchiconclick()

        print(str(nombre)+str(probability)+"(index# "+str(index)+" completed)")
        #put nombre in search bar, enter, wait, move mouse, click add rating,
        #wait, type i, tab, type p, tab tab tab, type probability, tab tab, enter
        #move mouse to searchMagGlass, click, LOOP

#calling, runs setup, then loops entries to selected headerrange--see line 20 ish to manually change how many entries are entered
setupfn()
entrylooper()

#End prints for manual list entry validation and cool vibes
print("entryIndices to ExcelRow "+str(headerrange+1)+"(inclusive) entered successfully :)")







