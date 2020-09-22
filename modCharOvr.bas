Attribute VB_Name = "modCharOvr"
'Please refer to HELP (modHelp.bas) for documentation.
'-----------------------------------------------------
 

Public Function RandomChar(Length As Long) As String
'generate random data (length of data)

'E.G. RandomChar(100) would produce random data
'of length 100

    Dim Position, StringLen As Long
    Dim rndString, Chars As String

    Chars = MyChars 'take characters from the user's input (default is 0 or 1 - binary)
    StringLen = 0

    Randomize

    Do Until StringLen = Length
        Position = Int((Len(Chars) * Rnd) + 1)
            rndString = rndString & Mid(Chars, Position, 1)
        StringLen = StringLen + 1
    Loop

    RandomChar = rndString

End Function
