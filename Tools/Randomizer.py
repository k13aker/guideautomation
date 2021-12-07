import random
import os

inputs = ['Mouse & Keyboard', 'Touch', 'Remote']

web = ['google.com', 'youtube.com', 'amazon.co.uk', 'bbc.co.uk', 'ebay.co.uk', 'twitter.com', 'wikipedia.org',
       'dailymail.co.uk', 'live.com', 'instagram.com', 'pornhub.com', 'reddit.com', 'rightmove.co.uk', 'paypal.com'
       'office.com', 'linkedin.com', 'express.co.uk', 'johnlewis.com', 'imdb.com', 'gov.uk',
       'indeed.co.uk', 'thesun.co.uk', 'amazon.com', 'asda.co.uk', 'tesco.com', 'tripadvisor.co.uk',
       'mirror.co.uk', 'autotrader.co.uk', 'currys.co.uk', 'asos.com', 'sky.com']

imap = ['U: dolphintest1@live.co.uk \n P: testsystem01',
        'U: guide.test1@gmail.com \n P: test01',
        'U: dolphin_matt@yahoo.com \n P: d0lphint3std0lphint3st',
        'U: DolphinTestAol@aol.com \n P: d0lphinT3st',
        'U: mattpw@btinternet.com \n P: d0lphint3st']

pop = ['U: testsystem1@guidemail.co.uk \n P: test01',
       'U: testsystem3@guidemail.co.uk \n P: test03']

random.shuffle(inputs)
random.shuffle(pop)
random.shuffle(web)
random.shuffle(imap)


def randomize():
    print("*********************************")
    print("      REGRESSION RANDOMIZER      ")
    print("*********************************")
    print("What would you like to Randomize?")
    print("---------------------------------")
    print("           E = Emails            ")
    print("          W = Websites           ")
    print("           I = Inputs            ")
    print("---------------------------------")
    answer = input(">").lower()
    os.system('cls')

    if "e" in answer:
        print("*********************************")
        print("             EMAILS              ")
        print("*********************************")
        print("IMAP \n", random.choice(imap))
        print("---------------------------------")
        print("POP \n", random.choice(pop))
        print("---------------------------------")
        print("Would you like to randomize more: Y/N ")
        answer = input(">").lower()
        if 'y'.lower() in answer:
            os.system('cls')
            randomize()

        else:
            exit()

    elif "w" in answer:
        print("*********************************")
        print("            WEBSITES             ")
        print("*********************************")
        print("1:", random.choice(web))
        print("2:", random.choice(web))
        print("3:", random.choice(web))
        print("---------------------------------")
        print("Would you like to randomize more: Y/N ")
        answer = input(">").lower()
        if 'y'.lower() in answer:
            os.system('cls')
            randomize()

        else:
            exit()

    elif "i" in answer:
        print("*********************************")
        print("             INPUTS              ")
        print("*********************************")
        print("Emails:", random.choice(inputs))
        print("Documents:", random.choice(inputs))
        print("Websites:", random.choice(inputs))
        print("Scanner & Camera:", random.choice(inputs))
        print("Bookshelf:", random.choice(inputs))
        print("Address Book & Calendar:", random.choice(inputs))
        print("Entertainment:", random.choice(inputs))
        print("Notes:", random.choice(inputs))
        print("Guide Community:", random.choice(inputs))
        print("Tools:", random.choice(inputs))
        print("Settings:", random.choice(inputs))
        print("Exit:", random.choice(inputs))
        print("---------------------------------")
        print("Would you like to randomize more: Y/N ")
        answer = input(">").lower()
        if 'y'.lower() in answer:
            os.system('cls')
            randomize()

        else:
            exit()

    else:
        os.system('cls')
        print("***************************")
        print("PLEASE CHOOSE A VALID INPUT")
        print("***************************")
        print("        R = Restart        ")
        print("     Any Input = Exit      ")
        print("---------------------------")
        answer = input(">")
        if 'r'.lower() in answer:
            os.system('cls')
            randomize()

        else:
            exit()


randomize()
