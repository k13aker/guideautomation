import win32con, win32api, win32gui, ctypes, ctypes.wintypes, threading, logging
from queue import Queue
from Misc.MyFunctions import *
from Suites.Regression import Regression

queue = Queue(1)
start = 0
timeout = 300


def thread(thread1, q):
    global data
    data = []

    class COPYDATASTRUCT(ctypes.Structure):
        _fields_ = [
            ('dwData', ctypes.wintypes.LPARAM),
            ('cbData', ctypes.wintypes.DWORD),
            ('lpData', ctypes.c_void_p)
        ]

    PCOPYDATASTRUCT = ctypes.POINTER(COPYDATASTRUCT)

    class Listener:
        def __init__(self):
            message_map = {
                win32con.WM_COPYDATA: self.OnCopyData
            }
            wc = win32gui.WNDCLASS()
            wc.lpfnWndProc = message_map
            wc.lpszClassName = 'MyPythonWindowClass'
            hinst = wc.hInstance = win32api.GetModuleHandle(None)
            classAtom = win32gui.RegisterClass(wc)
            self.hwnd = win32gui.CreateWindow(
                classAtom,
                "win32gui test",
                0,
                0,
                0,
                win32con.CW_USEDEFAULT,
                win32con.CW_USEDEFAULT,
                0,
                0,
                hinst,
                None
            )
            print(self.hwnd)

        def OnCopyData(self, hwnd, msg, wparam, lparam):
            pCDS = ctypes.cast(lparam, PCOPYDATASTRUCT)
            data = ctypes.wstring_at(pCDS.contents.lpData), wparam
            q.queue.clear()
            q.put(data, False)

    l = Listener()
    win32gui.PumpMessages()


api = threading.Thread(target=thread, args=("Thread-1", queue))
api.start()
launch()
data = readFromQueue(queue)


while 'Emails' not in data:
    data = readFromQueue(queue)
    if start < timeout:
        pag.sleep(10)
        start += 10
        if start == timeout:
            print(f"{data}: Timed out after {timeout} Seconds")
            returnhome()
            break

    if "GMENU_GETTING_STARTED" in data:
        gstart(queue)
        break

init()                                 # Awaits whilst GC Initialises
testdata()                             # Adds Test Data DIR to System Files
contacts()                             # Adds contacts.db to GC Contacts DIR
drive_locate()                         # Locates any USB Devices present

# Module: Emails
# Regression.email_send1(queue)         # Logs in and Sends email
# Regression.email_send2(queue)         # Checks sent folder for sent emails
# Regression.email_send3(queue)         # Logs into additional email account
# Regression.email_send4(queue)         # Open a received email
# Regression.email_send5(queue)         # Sends an email using address contact
# Regression.email_send6(queue)         # Adds/removes recipients from an email (from both manual & contacts)
# Regression.email_send7(queue)         # Changes email subject line
# Regression.email_send8(queue)         # Performs an email spell check
# Regression.email_send9(queue)         # Saves email to drafts (using F2 menu)
# Regression.email_send10(queue)        # Saves email to drafts (using ESC)
# Regression.email_send11(queue)        # Adds/removes attachments from an email
# Regression.email_send12(queue)        # Adds and sends all test date in an email
# Regression.email_receive1(queue)      # Manually checks for new email
# Regression.email_receive2(queue)      # Replies to a single recipient email
# Regression.email_receive3(queue)      # Replies to multiple recipient email
# Regression.email_receive4(queue)      # Forwards an email
# Regression.email_receive5(queue)      # Adds email contact to guide address book
# Regression.email_receive6(queue)      # Moves email to a different folder
# Regression.email_receive7(queue)      # Opens all viewers inside an email (IMAGE, DOC, PDF)
# Regression.email_receive8(queue)      # Adds a email contact to blocked senders (using F2 menu)
# Regression.email_receive9(queue)      # Views details of a received email
# Regression.email_receive10(queue)     # Marks an email as Read
# Regression.email_receive11(queue)     # Searches inbox for an email
# Regression.email_receive12(queue)     # Deletes an email
# Regression.email_receive13(queue)     # Saves attachments and opens them using my saved attachments
# Regression.email_attachments1(queue)  # Sorts documents
# Regression.email_attachments2(queue)  # Searches email attachments for certain attachment
# Regression.email_attachments3(queue)  # renames saved email attachment
# Regression.email_attachments4(queue)  # deletes saved email attachment
# Regression.email_folders1(queue)      # Adds custom email folder
# Regression.email_folders2(queue)      # Deletes custom email folder
# Regression.email_blocked1(queue)      # Edits details of an existing blocked contact
# Regression.email_blocked2(queue)      # Removes a blocked contact
# Regression.email_blocked3(queue)      # Adds a blocked sender
# Regression.email_blocked4(queue)      # Checks emails from blocked senders inputted in correct folder
# Regression.emailClean1(queue)         # Cleans the sent folder on Guideautomation1@outlook.com
# Regression.emailClean2(queue)         # Cleans all email folders on Guideautomation2@outlook.com

# Module: Letters & Documents
Regression.Docs1(queue)               # Saves document using ESC and views in recent documents
Regression.Docs2(queue)               # Saves document using F2 and views in my documents
# Regression.Docs3(queue)               # Spell checks document
# Regression.Docs4(queue)               # Looks up dictionary words
# Regression.Docs5(queue)               # Searches for a document from my documents
# Regression.Docs6(queue)               # Renames a document
# Regression.Docs7(queue)               # Deletes a document
# Regression.Docs8(queue)               # Opens a DOC file
# Regression.Docs9(queue)               # Opens a PDF file
# Regression.Docs10(queue)              # Opens a RTF file
# Regression.Docs11(queue)              # Opens a DOCX file
# Regression.Docs12(queue)              # Opens a DOC file from a USB
# Regression.Docs13(queue)              # Opens a PDF file from a USB
# Regression.Docs14(queue)              # Opens a RTF file from a USB
# Regression.Docs15(queue)              # Opens a DOCX file from a USB
# Regression.letter1(queue)             # Enters sender address and composes letter to address book contact
# Regression.letter2(queue)             # Composes letter with manual contact input
# Regression.letter3(queue)             # Changes sender address using F2
# Regression.letter4(queue)             # Changes recipient address using F2

# Module: Scanner & Camera
# Regression.scanner_text1(queue)       # Selects scanner to use
# Regression.scanner_text2(queue)       # Scans 1 page for text
# Regression.scanner_text3(queue)       # Scans 2 pages for text
# Regression.scanner_text4(queue)       # Re-scans a single document after multi scan
# Regression.scanner_text5(queue)       # Saves a scanned document
# Regression.scan_image1(queue)         # Scans for an image
# Regression.scan_image2(queue)         # Saves a scanned image
# Regression.scan_image3(queue)         # Scans for text inside of scanned image
# Regression.scanner9(queue)            # Renames a scanned document
# Regression.scanner10(queue)           # Deletes a scanned document
# Regression.camera1(queue)             # Captures image, saves and views it
# Regression.camera2(queue)             # Renames a saved captured image
# Regression.camera3(queue)             # Deletes a saved captured image

# Module: Bookshelf
# Regression.books1(queue)              # Checks all UK providers are present
# Regression.books2(queue)              # Downloads book from Project Gutenberg
# Regression.books3(queue)              # Continues reading a book
# Regression.books4(queue)              # Checks downloaded book is present in my books
# Regression.books5(queue)              # Views book information from my books
# Regression.books6(queue)              # Copies a book to a USB from my books
# Regression.books11(queue)             # Deletes a book from a USB device
# Regression.books7(queue)              # Deletes a book from my books
# Regression.books8(queue)              # Views details of a book library
# Regression.books9(queue)              # Views details of a book from find a new book
# Regression.books10(queue)             # Open and views a book from USB
# Regression.newspapers1(queue)         # Logs in and downloads newspaper
# Regression.newspapers2(queue)         # Continues playing a newspaper
# Regression.newspapers3(queue)         # Plays an edition from my newspapers
# Regression.newspapers4(queue)         # Unsubscribes from my newspapers
# Regression.newspapers5(queue)         # Views details of newspaper provider
# Regression.newspapers6(queue)         # Logs out of newspaper provider

# Module: Address Book & Calendar
Regression.Address1(queue)            # Adds contact to address book
# Regression.Address2(queue)            # Edits details of an existing contact
# Regression.Address3(queue)            # Sends an email to address book contact
# Regression.Address4(queue)            # Sends invite to video calling email to contact
# Regression.Address5(queue)            # Composes a letter to address book contact
# Regression.Address6(queue)            # Sends VCF as email attachment, deletes initial instance and imports received
# Regression.Address7(queue)            # Edit imported contacts details
# Regression.Address8(queue)            # Sort contacts by last name
# Regression.Address9(queue)            # Deletes multiple address book contacts
# Regression.calendar1(queue)           # Add new event to your calendar
# Regression.calendar2(queue)           # Edits details of an existing event
# Regression.calendar3(queue)           # Sends calendar attachment via email (.ics)
# Regression.calendar4(queue)           # Deletes a calendar event

# Module: Entertainment
# Regression.radio1(queue)              # Play a radio station from favourites
# Regression.radio2(queue)              # Continue listening to a station
# Regression.radio3(queue)              # Adds station to favourites from Favourites
# Regression.radio4(queue)              # Adds a custom radio station
# Regression.radio5(queue)              # Removes a radio station from favourites
# Regression.radio6(queue)              # View details of a favourites radio station
# Regression.radio7(queue)              # Remove a favourite station from play a new station
# Regression.radio8(queue)              # Add a station to favourites from play a new station
# Regression.radio9(queue)              # Searches play a new station
# Regression.radio10(queue)             # Changes station language list to Czech
# Regression.radio11(queue)             # View details of a station from play a new station
# Regression.radio12(queue)             # Iterates through active stations list vs expected
# Regression.Podcast1(queue)            # Play a new podcast from find a podcast
# Regression.Podcast2(queue)            # Continue playing a podcast
# Regression.Podcast3(queue)            # Add a podcast to favourites
# Regression.Podcast4(queue)            # Remove a podcast from favourites
# Regression.Podcast5(queue)            # View details of a favourite podcast
# Regression.Podcast6(queue)            # Add a custom podcast from favourites
# Regression.Podcast7(queue)            # Searches favourite podcasts list
# Regression.Podcast8(queue)            # Searches play a new podcast list
# Regression.Podcast9(queue)            # View details of a podcast from play a new podcast
# Regression.Podcast10(queue)           # Changes language of the podcast list
# Regression.Podcast11(queue)           # Searches and plays a podcast from play a new podcast
# Regression.Podcast12(queue)           # View details of a podcast episode from play a new podcast list
# Regression.Podcast13(queue)           # Add a podcast to favourites from play a new podcast
# Regression.Podcast14(queue)           # Download podcast and play from My downloaded podcasts
# Regression.Podcast15(queue)           # View details of a downloaded podcast
# Regression.Podcast16(queue)           # Copy a downloaded podcast to a USB
# Regression.Podcast18(queue)           # Delete a podcast from a USB device
# Regression.Podcast17(queue)           # Deletes a downloaded podcast
# Regression.music_1(queue)             # Import Music from computer
# Regression.music_2(queue)             # Album view: Play a track from album
# Regression.music_3(queue)             # Album view: Continue listening to a track
# Regression.music_4(queue)             # Album view: View and change album details
# Regression.music_5(queue)             # Album view: Rename an album
# Regression.music_6(queue)             # Album view: Add new album in details
# Regression.music_7(queue)             # Album view: Perform an album search
# Regression.music_8(queue)             # Album view (Track list): Views and changes all details
# Regression.music_9(queue)             # Album view (Track list): Delete a track
# Regression.music_10(queue)            # Album view: Search album list for a track
# Regression.music_11(queue)            # Album view: Remove albums then imports albums using F2
# Regression.music_12(queue)            # Album view: Change view
# Regression.music_13(queue)            # Artist view: Play a track
# Regression.music_14(queue)            # Artist view: Add and change the artist
# Regression.music_15(queue)            # Artist view: Rename the artist
# Regression.music_16(queue)            # Artist view: Remove an artist
# Regression.music_17(queue)            # Artist view: Perform a search
# Regression.music_18(queue)            # Artist view: Change view
# Regression.music_19(queue)            # Folder view: Search folders to play a track
# Regression.Hangman(queue)             # Launches Hangman
# Regression.Sudoku(queue)              # Launches Sudoku
# Regression.Blackjack(queue)           # Launches Blackjack

# Module: Notes
# Regression.tnote1(queue)              # Checks for No Notes present prompt
# Regression.tnote2(queue)              # Creates new text note
# Regression.tnote3(queue)              # Renames existing text note
# Regression.tnote4(queue)              # Edits a text note
# Regression.tnote5(queue)              # Sends text note via email
# Regression.tnote6(queue)              # Views details and deletes note
# Regression.tnote7(queue)              # Creates a note without saving
# Regression.anote1(queue)              # Creates new audio note
# Regression.anote2(queue)              # Sends audio note via email

# Module: Tools
# Regression.about(queue)               # Checks premium plan is active from about menu
# Regression.calc(queue)                # Using calculator to add, minus, multiply and divide
Regression.dict_valid(queue)          # Searches for a valid dictionary term
# Regression.dict_invalid(queue)        # Searches for an invalid dictionary term
# Regression.gstart_vids(queue)         # Plays the 4 Getting started videos
# Regression.learn_keyboard(queue)      # Inputs all keyboard presses on learn the keyboard
# Regression.shortcut(queue)            # Opens shortcut documentation
# Regression.backup(queue)              # Backs up guide connect
# Regression.restore(queue)             # Restores guide connect
# Regression.reset_settings(queue)      # Resets settings only
# Regression.factory(queue)             # Resets to factory settings
# Regression.file_1(queue)              # Locates DOCX and opens it
# Regression.file_2(queue)              # Opens an RTF file
# Regression.file_3(queue)              # Opens a TXT file
# Regression.file_4(queue)              # Opens a JPG file
# Regression.file_5(queue)              # Opens a PNG file
# Regression.file_6(queue)              # Opens a DOC file
# Regression.file_7(queue)              # Opens a DOCX file
# Regression.file_8(queue)              # Moves a selected file
# Regression.file_9(queue)              # Renames a file
# Regression.file_10(queue)             # Copies a file to a different location/
# Regression.file_11(queue)             # Deletes files and locates them in Deleted items
# Regression.file_12(queue)             # Restore multiple files from Deleted items
# Regression.file_13(queue)             # Delete multiple files from Deleted items
# Regression.file_14(queue)             # Remove all items from Deleted items

# Module: Settings
