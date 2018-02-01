print(' MAIN CONTROL CENTRE. Receives GPS data, sends data to nearby users')
import smtplib, imapclient, pyzmail, time, subprocess, pprint, random, imaplib, shelve, threading, pyperclip, requests, webbrowser
imaplib._MAXLINE = 10000000
 
def getInstructionEmails():
    global emailFrom
    instructions = []
    print('\n-- Scanning for new emails --')
    print('Checking UIDs')
    imapCli = imapclient.IMAPClient(imapServer, ssl = True)
    imapCli.login(botEmail, botPassword)
    imapCli.select_folder('INBOX', readonly = True)
    UIDs = imapCli.search(['SUBJECT Task', 'UNSEEN'])
    print(UIDs)
    try:
        UIDlist = list(UIDs)
        selectedUID = UIDlist[0]
        print('\nSelected the first UID number: ' + str(selectedUID))
    except IndexError as err:
        print('No new emails')
        time.sleep(1)
        getInstructionEmails()
        
    print('\n ----- CHECKING TO SEE IF THE EMAIL NAMES MATCH -----')
    rawMessages = imapCli.fetch([selectedUID], [b'BODY[]', 'FLAGS'])
    message = pyzmail.PyzMessage.factory(rawMessages[selectedUID][b'BODY[]'])
    emailFrom = (message.get_address('from'))  
    print('Email FROM: ' + str((emailFrom)))

    shelfFile = shelve.open('C:\\Users\\250gb NoSteam\\Desktop\\main\\membersList')
    dictionary = shelfFile['storage']
    shelfFile.close()
    usernames = list(dictionary.keys())
    print(usernames)    
    if emailFrom[1] in usernames:
        print('\nMATCHING NAMES FOUND! ' + emailFrom[1])
    else:
        print('-- UNKNOWN EMAIL SENDER! ---''\n''Someone has sent a request in the subject header, but are not in the mailing list.')
        imapCli.logout()
        print('\nINFORM ADMIN! \n''\n'' ------ L O G --------')
        time.sleep(30)
        mainStart()
    print('\nLogging out from SMTP / IMAP...')            
    for UID in rawMessages.keys():
        message = pyzmail.PyzMessage.factory(rawMessages[UID][b'BODY[]'])
        if message.html_part != None:
            body = message.html_part.get_payload().decode(message.html_part.charset)
        if message.text_part != None:
            body = message.text_part.get_payload().decode(message.text_part.charset)
        instructions.append(body)
    print('Instructions PREVIEW:')
    print(instructions)
    #imapCli.delete_messages(UIDs) #imapCli.expunge() #print('Expunged')
    imapCli.logout()
    return instructions

def parseInstructionEmail(instruction):
    global PSK, Long, Latt, ZIPcode, emailFrom
    shelfFile = shelve.open('C:\\Users\\250gb NoSteam\\Desktop\\main\\membersList')
    dictionary = shelfFile['storage']
    
    print('\n_______________________________''\n''\n --- START TO PARSE EMAIL INSTRUCTIONS ---\n')    
    emailAddress = emailFrom[1]
    shelvePSK = dictionary.get(emailAddress, 0)
    print('FROM:  ' + str(emailFrom[1]) + '\nSHELVED PSK: ' + str(shelvePSK))
    responseBody = 'Subject: Confirmation.\nInstruction received and completed.\nResponse:\n'
    lines = instruction.split('\n')
    for line in lines:
        print('\nPrinting line.')
        print(line)
        if line.startswith(str(shelvePSK)):
            print('\nFound password match!''\n' + str(line) + '\n' + str(shelvePSK))
            characters = '!"#$%&\'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\] ^_`abcdefghijklmnopqrstuvwxyz{|}~'
            pskLength = 42
            PSK = ""
            for i in range(pskLength):
                newPSK = random.randrange(len(characters))
                PSK += characters[newPSK]
            print('\nNew password: ' + PSK)
            responseBody += '\nPassword:\n' + PSK
            p=input('PAUSED')
            print('\n''________''\n''. . .''\n')
            
            mailServer = smtplib.SMTP('smtp-mail.outlook.com' , 587)
            mailServer.starttls()
            mailServer.login(botEmail, botPassword)
            mailServer.sendmail(botEmail, emailAddress, responseBody)
            mailServer.quit()
            print('\nNew password sent!''\n')

            print('\nShelving the new password...')  
            try:
                membersList = shelfFile['storage']                
            except KeyError as err:
                print('KeyError, probably from using pop')             
            pprint.pprint(membersList)
            print('\n''\n''\n')
            ele = membersList.pop(emailFrom[1])
            pprint.pprint(membersList)
            print('\n''\n')
            membersList[emailFrom[1]] = PSK
            pprint.pprint(membersList)
            shelfFile['storage'] = membersList
            shelfFile.close()

            p=input('PAUSE TO CHECK SHELVE')

            print('Checking for GPS co-ordinates...')
            for line in lines:
                if 'Latitude' in line:
                    print('\nLatitude in lines!!''\n')     
                    from ast import literal_eval
                    a = literal_eval(line)
                    print('PRINTING A')
                    print(a)
                    Latt = (a.get('Latitude', 0))
                    Long = (a.get('Longitude', 0))
                    print('Showing long and latt')
                    print(Latt)
                    print(Long)
                    address = str(Latt) + ' ' + str(Long)
                    print('Opening browser')
                    #webbrowser.open('https://www.google.com/maps/place/' + address)
                    time.sleep(5)
                    print('Should be open now')
                    #getGoogleMap()
                    p=input('END OF PROGRAM{}')
        else:
            print('INCORRECT PASSWORD!!''Known members email address, but with an incorrect password.''\n' 'INFORM ADMIN')
            time.sleep(10)
            print('---------------------------L O G ------------------------')
            startMain()    
        
def getGoogleMap():
    print('Cheat to get data')
    pyautogui.moveTo(341, 388)
    pyautogui.click()
    time.sleep(2)
    pyautogui.moveTo(613, 327)
    pyautogui.click()
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'c')
#    pyautogui.keyDown('ctrl'); pyautogui.press('w'); pyautogui.keyUp('ctrl')
#    pyautogui.keyDown('alt'); pyautogui.press('tab'); pyautogui.keyUp('alt')
    temp = pyperclip.paste()
    print(temp)

def newMember():
    global membersList
    print('Enter a name: (blank to quit)')
    name = input()
    if name in membersList:
        print(name + ' already exists')
    else:
        print(name + ' is available')
        print('Enter password')
        PASSKEY = '123456789012'
        membersList[name] = PASSKEY
        print('Database updated.')
        print(membersList)
        p=input('PAUSED. Needs work. - newMember()')    

Long = ''
Latt = ''
ZIPcode = ''

emailFrom = ''
botEmail = input('Enter email address: ')
botPassword = input('\nEnter email password: ')

imapServer = 'imap-mail.outlook.com'
smtpServer = 'imap-mail.outlook.com'
SMTP_PORT = 465

print('\n''\n----COMMAND CENTRE----\n''-----------------------------''\n''\n')

print('Program started')
def startMain():
    while True:
        instructions = getInstructionEmails()
        for instruction in instructions:
            parseInstructionEmail(instruction)
        time.sleep(1)
        
startMain()    
