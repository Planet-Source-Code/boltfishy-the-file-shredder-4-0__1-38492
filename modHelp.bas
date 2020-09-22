Attribute VB_Name = "HELP"
Global NumberOfTimes As Long 'number of times we should overwrite
                             'each character in the file

Global Setting As String 'ultra quick, paranoid etc

Global Rename As Boolean 'rename the file?

Global FileTemp As String

Global MyChars As Variant 'chars to be used to overwrite files

'----------------------------------------------------------
' App Name                : The File Shredder
' App Version             : 4.00
' Purpose                 : Securely delete files
' Written by              : Mischa Balen
' Contact E-mail          : boltfish@eml.cc
' Website                 : www26.brinkster.com/boltfish/
' Date Created            : July / August 2002
' Last Updated            : 30th August 2002
'----------------------------------------------------------

'What's new in Version 4?
'------------------------

'OK, well I have formatted the code nicely so it's easier to read,
'tweaked a few GUI issues and added in a neat function which enables
'you to choose which characters are used to overwrite your file with.
'E.G. "Hello" could be scrambled with "10" i.e. binary, as "10110" or
'it could be scrambled with "98A" as "89A8A".

'In version 4.5 or whatever, I will add a feature to unprotect read only
'files and delete them too.

'Bugs
'----
'   - No support for dragging and dropping folders
'   - Can't drag and drop multiple files from the desktop :-(
'   - Can't always remove the context menu


'CONTENTS:

'   1. Foreword
'   2. Global Variables
'   3. Application Model and Functions
'   4. Delete Clicked
'   5. Shred File Function
'   6. Binary Function
'   7. Adding Context Menus
'   8. Custom Deletion Methods


'           Brief Foreword and Introduction
'      ----------------------------------------

'Hello all! Thanks for downloading my code. If you have
'downloaded it from Planet Source Code, then I'm not
'going to beg for 5 globes (it doesn't deserve them) but
'I would appreciate it greatly if you would offer some
'feedback, either on PSC or by email. Thank you. That way,
'I can make this a better programme, which would be great.

'This code is copyrighted and you may not resuse it in any
'way without my permission. Please respect that; I have
'been kind enough to share it with you in the first place.

'I recommend that you read this article through before you
'do anything else; just to get a feel for how the app
'works, if anything at all.

'Keep on coding! Oh, and keep it open source ;)

'Mischa Balen aka ~boltfish~


'              GLOBAL VARIABLES - Useage
'      ----------------------------------------

'OK, so firstly lets deal with the global variables,
'which are:

' ---------------
'| NumberOfTimes |
'| Binary        |
'| MyChars       |
'| FileTemp      |
' ---------------

'1. NumberOfTimes =
'This is a global variable which the user can set from
'frmOptions.frm. It stores the data telling us how many
'times we should overwrite the file. Upon startup, its
'value is automatically 1000. If the user changes it,
'then it is updated and called by the ShredFile function.
'The maximum value is 999,999,999.

'2. Binary =

'The user can also specify whether or not to overwrite
'the files' contents with random character data.
'It is accessed via frmOptions2.frm
'It is therefore a Boolean and is called by the ShredFile
'Routine.

'3. MyChars =

'Characters which are used to overwrite the file IF binary
'is true. Specified by the user under frmChars.frm.

'4. FileTemp =
'When the user clicks 'delete file' it goes into a loop
'until all the files have been removed by the ShredFile
'routine. Before the file is deleted, its path is written
'to FileTemp. Then, we use the GetFileName sub to return
'the file name. This is so we can add the file name to
'the status bar panel even if it has just been deleted.


'          APPLICATION MODEL AND FUNCTIONS:
'      ----------------------------------------

'Now that we have got those sorted out, let's take a look
'at what happens when the use clicks the delete button.
'Here is the basic model process for the whole app:

'Delete Clicked -> File Name Read -> File Encrypted ->
'File OverWritten -> File Replaced with "" -> File Deleted

'    DELETE CLICKED - We declare the following:
'    ------------------------------------------

    ' ------------------------
    '| Dim i As Integer       |
    '| Dim b As Integer       |
    '| Dim File2Del As String |
    '| Dim msg As String      |
    ' ------------------------

'1. I / B = Counter. B is labelled as the number of files
'in the listbox - i.e. the number of files to be deleted.
'I can be thought of as the current file (which is being
'deleted).

'Using I, we progress in steps of 1.

'In every stage, we:
'
'   1)Set the display panel to "Deleting" i "of" b
'   2)Set the other panel to the file name of the current
'     file which is being deleted
'   3)Use ShredFile Function

'This loop runs until I = B, i.e. when the current file
'being deleted = the total number of files. Therefore we
'must have finished, so we take appropriate action.


'     SHRED FILE FUNCTION - This is explained:
'     -----------------------------------------

'This is the primary and most important feature in the
'programme. It makes the file safe before finally deleting
'it. Again, look at the model below:

'   1. Sets up the main loop 1 until NumberOfTimes
'   2. Sets up the second loop, 1 until number of characters
'      in the file.
'   3. For each character, it generates a random character
'      from the MyChars list and replaces the original
'      character with it.
'   4. Keeps doing this until every character in the file
'      has been overwritten however many times the user
'      wants (NumberOfTimes).
'   5. Exits loops
'   6. If Rename is true it renames the files
'   7. Then it deletes them


'         BINARY FUNCTION - Explained below:
'     -----------------------------------------

'How this works:

'OK, it works by working in two loops:
'1. for x = 1 to NumberOfTimes
'2. for i = 1 to lof(1)

'In the second loop, it opens the file and overwrites
'each character with user specified chars (MyChars). This
'is faster than the previous method of simply opening the
'file and generating a new string of length equal to the
'number of characters in the file.


'      ADDING CONTEXT MENUS - Explained below:
'     -----------------------------------------

'The File Shredder can add 'context menus' to files. The
'context menu is the menu that you get when you right
'click a file. A new option will be created called 'Delete
'with TFS' and when clicked, loads up the main programme
'with the file's path added to the main listbox.

'The user has the option to remove this function through
'the options menu. If the app path changes then they should
'choose to add the context menu from the options screen so
'that the registry is updated with the new value.

'When the app loads (form_load) it tries automatically to
'replace the context menu. This is in case the app has been
'moved or whatever, so it automatically updates the registry
'with its new path.


'      CUSTOM DELETION METHODS - Explained below:
'     -------------------------------------------

'From the options menu, the user can select one of five
'methods of deletion -

'1. Ultra Quick
'2. Quick
'3. Normal (default)
'4. Paranoid
'5. Custom

'When one of these is checked, the following global values
'are changed in accordance:

'Rename ----------- true / false - should we rename files?
'Binary ----------- true / false - should we use overwrite the file using random chars?
'NumberOfTimes ---- number of times to overwrite the file using the hex corrupt
'Setting ---------- name e.g. normal/paranoid/custom

'When the user loads up the custom form to edit the details,
'the controls are filled in with the global values already
'specified by the current setting.
