Attribute VB_Name = "KiKz_Bas"
Option Explicit
'       wasup?
'i know a few subs on this bas dont quite work
'but dont complain...cause if u complain about my work u can go on ahead and write your OWN code
'remember, this is the first beta release of my bas
'
'all code in this bas was written by ME unless stated otherwise
'
'and all code in this bas is Legally Copyrighted to ME not U
'in otherwords...
'
'WRITE YOUR OWN MOTHERFUCKING CODE!
'BITCH
'
'Greetz-
'
'Kaos
'chad
'lax
'haaj
'KnK
'solja
'har0
'crzy
'kuso
'
'lata
'   -kik
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function CharUpper Lib "user32" Alias "CharUpperA" (ByVal lpsz As String) As String
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long


Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const WM_GETTEXT = &HD
Private Const SPI_SCREENSAVERRUNNING = 97
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_SETTEXT = &HC
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_ENABLE = &HA
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const WM_MOVE = &H3
Public Const WM_SYSCOMMAND = &H112
Public Const HIDE_WINDOW = 0
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_SELECTSTRING = &H18C

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5

Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_UP = &H26

Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_LBUTTONDBLCLK = &H203
Public Const VK_SPACE = &H20


Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Function BLWindow() As Long
'this could help u if u r writing your own bas...
'i dont care if u take this code
'it just finds the bl window
Dim Window As Long
Window& = FindWindow("_Oscar_BuddyListWin", vbNullString)
BLWindow& = Window&
End Function

Function IM_Sn() As String
'gets the screenname of the person u r talking to
    Dim IMWin As Long, GetIt As String, Clear As String

    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    GetIt$ = Get_Caption(IMWin&)
    Clear$ = ReplaceString(GetIt$, " - Instant Message", "")
    IM_Sn = Clear$
End Function

Function Dis_Ctrl_Alt_Del()
'this disables Ctrl+Alt+Delete
'make sure u enable it before your prog. ends
Dim ret As Integer
Dim pOld As Boolean
     ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Function
Function En_Ctrl_Alt_del()
'this enables Ctrl+Alt+Delete
Dim ret As Integer
Dim pOld As Boolean
     ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Function

Function Aim_Flash()
'this will drive sombody crazy...
'make a timer set the interval to whatever u want
'in it put Aim_Flash
'to stop it just disable the timer AND use...
'Aim_StopFlash
Dim BL As Long
BL& = BLWindow
Call FlashWindow(BL&, True)
End Function

Function Aim_StopFlash()
'stops Aim_Flash
Dim BL As Long
BL& = BLWindow
Call FlashWindow(BL&, False)
End Function

Function Flash_Form(Frm As Form)
'this will drive sombody crazy...
'make a timer set the interval to whatever u want
'in it put Flash_Form Me
'to stop it just disable the timer AND use...
'Flash_Stop
    Call FlashWindow(Frm.hwnd, True)
End Function

Function Flash_Stop(Frm As Form)
'stops Flash_Form
Call FlashWindow(Frm.hwnd, False)
End Function

Function Set_AimCaption(NwCaption As String)
'sets the caption of your BL window
'i dont reccomend u do this
'cause the get SN wont work
Dim BL As Long
BL& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Call SetWindowText(BL&, NwCaption$)
End Function

Public Function Aim_LastLine() As String
'my pride and joy...i wrote ALL this code it was not stolen
    On Error GoTo ErrHandler

    Dim ChatText As String
    ChatText$ = Aim_GetChatText
    If Len(ChatText$) > 500 Then
        ChatText$ = Right$(ChatText$, 250)
    End If
    If InStr(ChatText$, ")--></B></FONT><FONT COLOR=""#") <> 0 Then
        ChatText$ = Mid$(ChatText$, LastInStr(ChatText$, ")--></B></FONT><FONT COLOR=""#") + 38)
        ChatText$ = RemoveHTML(ChatText$)
        ChatText$ = Trim$(ChatText$)
        Aim_LastLine = Keep(ChatText$, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890)(*&^%$#@!~`"";:'?/\.,][}{+=_-|» ")
    Else
        Aim_LastLine = ""
    End If
    Exit Function
    
ErrHandler:
    Aim_LastLine = ""

End Function

Function Aim_CloseChat()
'closes the chat window
Dim ChtWnd As Long
ChtWnd& = FindWindow("AIM_ChatWnd", vbNullString)
Call SendMessage(ChtWnd&, WM_CLOSE, 0&, 0&)
End Function

Function Aim_CloseIM()
'closes the IM window
Dim IMWnd As Long
IMWnd& = FindWindow("AIM_IMessage", vbNullString)
Call SendMessage(IMWnd&, WM_CLOSE, 0&, 0&)
End Function

Function Aim_SignOff()
'Sign off aim
Dim BLWnd As Long
BLWnd& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Call SendMessage(BLWnd&, WM_CLOSE, 0&, 0&)
End Function

Function Chat_RoomName()
'gets the chat room name
    Dim ChatWin As Long, GetIt As String, Clear As String

    ChatWin& = FindWindow("AIM_ChatWnd", vbNullString)
    GetIt$ = Get_Caption(ChatWin&)
    Clear$ = ReplaceString(GetIt$, "Chat Room: ", "")
    Chat_RoomName = Clear$
End Function

Function LB_Search(Search As String, LB As ListBox)
'This is the fastest ListBox search there is
Call SendMessageByString(LB.hwnd, LB_SELECTSTRING, 0&, Search$)
End Function

Sub Chat_Child(Frm As Form)
'makes the aim chat room your programs child
Dim Cht As Long
Cht& = FindWindow("AIM_ChatWnd", vbNullString)
Call SetParent(Cht&, Frm.hwnd)
End Sub

Public Function Aim_LastSender() As String
'another one of my good functions...
    On Error GoTo ErrHandler
    
    Dim ChatText As String
    ChatText$ = Aim_GetChatText()
    If InStr(ChatText$, "<BODY BGCOLOR=""#") <> 0 Then
        If Len(ChatText$) > 500 Then
            ChatText$ = Right$(ChatText$, 250)
        End If
        ChatText$ = Mid$(ChatText$, LastInStr(ChatText$, "<BODY BGCOLOR=""#"))
        If InStr(ChatText$, "<!-- (") <> 0 Then
            ChatText$ = Left$(ChatText$, InStr(ChatText$, "<!-- (") - 1)
            ChatText$ = Mid$(ChatText$, LastInStr(ChatText$, ">") + 1)
            Aim_LastSender = ChatText$
        Else
            Aim_LastSender = ""
        End If
    Else
        Aim_LastSender = ""
    End If
    Exit Function
    
ErrHandler:
    Aim_LastSender = ""

End Function

Function Crzy_Mouse()
Do
    ShowCursor (False)
Pause 2#
    ShowCursor (True)
Loop
End Function

Function IM_GetAllText() As String
'gets all the text from the IM window
Dim IMWindow As Long, IMTextBox As Long, IMTextBoxLen As Long, Buffer As String
IMWindow& = FindWindow("AIM_IMessage", vbNullString)
IMTextBox = FindWindowEx(IMWindow, 0&, "WndAte32Class", "AteWindow")
IMTextBoxLen& = SendMessage(IMTextBox&, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = String(IMTextBox&, 0&)
Call SendMessageByString(IMTextBox&, WM_GETTEXT, IMTextBoxLen + 1, Buffer$)
IM_GetAllText = Buffer$
End Function

Function IM_GetAllNoHTML() As String
'gets all the text from the IM window and removes the html
Dim IMWindow As Long, IMTextBox As Long, IMTextBoxLen As Long, Buffer As String
IMWindow& = FindWindow("AIM_IMessage", vbNullString)
IMTextBox = FindWindowEx(IMWindow, 0&, "WndAte32Class", "AteWindow")
IMTextBoxLen& = SendMessage(IMTextBox&, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = String(IMTextBox&, 0&)
Call SendMessageByString(IMTextBox&, WM_GETTEXT, IMTextBoxLen + 1, Buffer$)
IM_GetAllNoHTML = RemoveHTML(Buffer$)
End Function

Public Function LastInStr(String1 As String, WhatToFind As String)
'finds the last occurence of a string within a nother string
    Dim CurrLoc As Long, I As Long
    For I = 1 To Len(String1$) - Len(WhatToFind$) + 1

        If Mid$(String1$, I, Len(WhatToFind$)) = WhatToFind$ Then CurrLoc& = I
    
    Next I

    LastInStr = CurrLoc&

End Function

Public Function Collection_ItemLoc(TheColl As Collection, Item As String)

    Dim I As Integer
    Collection_ItemLoc = 0
    For I = 1 To TheColl.Count
        If TheColl.Item(I) = Item$ Then
            Collection_ItemLoc = I
            Exit Function
        End If
    Next I
    
End Function

Sub Pause(interval)
'Pauses for a given time
    Dim Current
    
    Current = Timer
    Do While Timer - Current < Val(interval)
        DoEvents
    Loop
End Sub

Sub FormDrag(TheForm As Form)
'let you drag a form that doesnt have a border...example:
'Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Call ReleaseCapture
'    Call SendMessage(TheForm.hWnd, WM_SYSCOMMAND, WM_MOVE, 0&)
'End Sub
    Call ReleaseCapture
    Call SendMessage(TheForm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0&)
End Sub

Sub Save_ListBox(Path As String, Lst As ListBox)
'Ex: Call Save_ListBox("c:\windows\desktop\list.lst", list1)

    Dim Listz As Long
    On Error Resume Next

    Open Path$ For Output As #1
    For Listz& = 0 To Lst.ListCount - 1
        Print #1, Lst.List(Listz&)
        Next Listz&
    Close #1
End Sub

Sub Save_ComboBox(Path As String, Combo As ComboBox)
'Ex: Call Save_ComboBox("c:\windows\desktop\combo.cmb", combo1)

    Dim Saves As Long
    On Error Resume Next

    Open Path$ For Output As #1
    For Saves& = 0 To Combo.ListCount - 1
        Print #1, Combo.List(Saves&)
    Next Saves&
    Close #1
End Sub

Sub Load_ComboBox(Path As String, Combo As ComboBox)
'Call Load_ComboBox("c:\windows\desktop\combo.cmb", Combo1)

    Dim What As String
    On Error Resume Next
    Open Path$ For Input As #1
    While Not EOF(1)
        Input #1, What$
        DoEvents
        Combo.AddItem What$
    Wend
    Close #1
End Sub

Sub Load_ListBox(Path As String, Lst As ListBox)
'Ex: Call Load_ListBox("c:\windows\desktop\list.lst", list1)

    Dim What As String
    On Error Resume Next

    Open Path$ For Input As #1
    While Not EOF(1)
        Input #1, What$
        DoEvents
        Lst.AddItem What$
    Wend
    Close #1
End Sub

Function Lst_Extract(LstBox As ListBox, Txtbox As textbox)
'will add all the items in a listbox into a textbox and seperate them with ","
Dim a As Long
Dim b As String
    
        For a = 1 To LstBox.ListCount - LstBox.ListCount
            LstBox.AddItem ", " & a
        Next

        For a = 0 To (LstBox.ListCount - 1)
            b = b & LstBox.List(a) & ", "
            
    Next


        Txtbox.Text = Mid(b, 1, Len(b) - 2)
        
End Function


Sub Load_Text(Txt As textbox, FilePath As String)
'Ex: Call load_Text(list1,"c:\windows\desktop\text.txt")

    Dim mystr As String, FilePath2 As String, textz As String, a As String
    
    Open FilePath2$ For Input As #1
    Do While Not EOF(1)
    Line Input #1, a$
        textz$ = textz$ + a$ + Chr$(13) + Chr$(10)
        Loop
        Txt = textz$
    Close #1
End Sub

Sub Save_Text(Txt As textbox, FilePath As String)
'Ex: Call Save_Text(list1,"c:\windows\desktop\text.txt")
    Dim FilePath3 As String
    
    Open FilePath3$ For Output As #1
        Print #1, Txt
    Close 1
End Sub

Public Function Fade(obj As Object)
'this is tight as shit...
'make a timer set the interval to anything (the smaller the interval the faster the fade)
'in the timer put Fade object
'for object put the object
'this will fade the text in any control except a rich textbox
Dim x As Variant
Dim c(2) As Byte
For x = 0 To 2
Randomize
c(x) = Int((255 - 0 + 1) * Rnd + 0)
Next x
On Error GoTo 200
obj.ForeColor = RGB(c(0), c(1), c(2))
200
    Exit Function
End Function

Public Function Split(String1 As String, SplitBy As String)
'Splits a string into a given number of pieces
    Dim Word As String, ctSplits As Integer, ctSplitBys As Integer, I As Integer
    
    For I = 1 To Len(String1$)
        If Mid$(String1$, I, 1) = SplitBy$ Then ctSplitBys = ctSplitBys + 1
    Next I
    ReDim Splits(ctSplitBys)
    
    Do Until InStr(String1$, SplitBy) = 0
        Word$ = Left$(String1$, InStr(String1$, SplitBy$) - 1)
        String1$ = Mid$(String1$, InStr(String1$, SplitBy$) + 1)
        Splits(ctSplits) = Word$
        ctSplits = ctSplits + 1
    Loop
    Splits(ctSplits) = String1$
    
    Split = Splits

End Function



Public Function Win_WordPadText()
'Gets the text from wordpad
    Dim WordPad As Long, textbox As Long, TextLength As Long, Buffer As String
WordPad& = FindWindow("WordPadClass", vbNullString)
textbox& = FindWindowEx(WordPad&, 0&, "RichEdit20A", vbNullString)
TextLength& = SendMessage(textbox&, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = String(TextLength, 0&)
Call SendMessageByString(textbox&, WM_GETTEXT, TextLength + 1, Buffer$)
Win_WordPadText = Buffer$
End Function

Public Function Win_NotePadText()
'Gets the text from notepad
    Dim Window As Long, textbox As Long, TextLength As Long, Buffer As String
    Window& = FindWindow("Notepad", vbNullString)
    textbox = FindWindowEx(Window&, 0&, "Edit", vbNullString)
    TextLength& = SendMessage(textbox&, WM_GETTEXTLENGTH, 0&, 0&)
    Buffer$ = String(TextLength, 0&)
    Call SendMessageByString(textbox&, WM_GETTEXT, TextLength& + 1, Buffer$)
    Win_NotePadText = Buffer$

End Function

Public Function Aim_GetChatText()
'Gets all the text from the aim 2.1+ chat textbox
    Dim Window As Long, Window1 As Long, ChatTB As Long, ChatTBLength As Long, Buffer As String
    Window& = FindWindow("AIM_ChatWnd", vbNullString)
    Window1& = FindWindowEx(Window&, 0&, "WndAte32Class", "AteWindow")
    ChatTB& = FindWindowEx(Window1&, 0&, "Ate32Class", vbNullString)
    ChatTBLength& = SendMessage(ChatTB&, WM_GETTEXTLENGTH, 0&, 0&)
    Buffer$ = String(ChatTBLength&, 0&)
    Call SendMessageByString(ChatTB&, WM_GETTEXT, ChatTBLength& + 1, Buffer$)
    Aim_GetChatText = Buffer

End Function

Function Get_Caption(TheWin)
'gets the caption of a window
    Dim WindowLngth As Integer, WindowTtle As String, Moo As String
    
    WindowLngth% = GetWindowTextLength(TheWin)
    WindowTtle$ = String$(WindowLngth%, 0)
    Moo$ = GetWindowText(TheWin, WindowTtle$, (WindowLngth% + 1))
    Get_Caption = WindowTtle$
End Function

Function Aim_UserSN() As String
'gets the users SN
    Dim BuddyList As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
      Aim_UserSN = "[ not online ]"
      Exit Function
    End If

Start:
    Dim GetIt As String, Clear As String
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    GetIt$ = Get_Caption(BuddyList&)
    Clear$ = ReplaceString(GetIt$, "'s Buddy List Window", "")
     Aim_UserSN = Clear$
End Function

Function ReplaceString(MyString As String, ToFind As String, ReplaceWith As String) As String
'now...i didnt write this, Dos did , so dont tell me a put somthing of Dos's in my bas and said it was mine
'Dos gets FULL credit for this
  Dim Spot As Long, NewSpot As Long, LeftString As String
    Dim RightString As String, NewString As String
    Spot& = InStr(LCase(MyString$), LCase(ToFind))
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = Right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
            Else
                RightString = ""
            End If
            NewString$ = LeftString$ & ReplaceWith$ & RightString$
            MyString$ = NewString$
        Else
            NewString$ = MyString$
        End If
        Spot& = NewSpot& + Len(ReplaceWith$)
        If Spot& > 0 Then
            NewSpot& = InStr(Spot&, LCase(MyString$), LCase(ToFind$))
        End If
    Loop Until NewSpot& < 1
    ReplaceString$ = NewString$
End Function

Public Sub Aim_ChatSend(Message As String)
'Sends chat to the aim 2.1+ textbox
    Dim Window As Long, Window1 As Long, ChatTB As Long, SendIcon As Long
    Window& = FindWindow("AIM_ChatWnd", vbNullString)
    Window1& = FindWindowEx(Window&, 0&, "WndAte32Class", "AteWindow")
    Window1& = FindWindowEx(Window&, Window1&, "WndAte32Class", "AteWindow")
    ChatTB& = FindWindowEx(Window1&, 0&, "Ate32Class", vbNullString)
    SendIcon& = FindWindowEx(Window&, 0&, "_Oscar_IconBtn", vbNullString)
    SendIcon& = FindWindowEx(Window&, SendIcon&, "_Oscar_IconBtn", vbNullString)
    SendIcon& = FindWindowEx(Window&, SendIcon&, "_Oscar_IconBtn", vbNullString)
    SendIcon& = FindWindowEx(Window&, SendIcon&, "_Oscar_IconBtn", vbNullString)
    Call SendMessageByString(ChatTB&, WM_SETTEXT, 0&, Message$)
    Call SendMessage(SendIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendIcon&, WM_LBUTTONUP, 0&, 0&)

End Sub

Public Sub Aim_BoldChat(Message As String)
'makes your text bold
Aim_ChatSend ("<FONT><B>" & Message & "</B></FONT>")
End Sub

Public Sub Aim_ItalicChat(Message As String)
'makes your text italic
Aim_ChatSend ("<FONT><I>" & Message & "</I></FONT>")
End Sub

Public Sub Aim_UnderlineChat(Message As String)
'underlines your text
Aim_ChatSend ("<FONT><U>" & Message & "</U></FONT>")
End Sub

Public Sub Aim_ChatColor(Message As String, ColorOfText As ColorConstants)
'make u text a differnt color
Dim ColorTS As String
ColorTS = ColorOfText
ColorTS = "#" & ColorTS
Aim_ChatSend ("<FONT COLOR=" & """ & ColorTS & """ & Message & "</FONT>")
End Sub

Public Sub Aim_Popup(SendName As String)
'makes a blank im popup on sombodies screen...use it to piss ppl off
    Dim BuddyList As Long
    Dim TabWin As Long, IMbuttin As Long, IMWin As Long
    Dim ComboBox As Long, TextEditBox As Long, TextSet As Long
    Dim EditThing As Long, TextSet2 As Long, SendButtin As Long, Click As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    TabWin& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
    IMbuttin& = FindWindowEx(TabWin&, 0, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(SendButtin&, WM_ENABLE, 1&, 1&)
    Click& = SendMessage(IMbuttin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(IMbuttin&, WM_LBUTTONUP, 0, 0&)
  
    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    ComboBox& = FindWindowEx(IMWin&, 0, "_Oscar_PersistantCombo", vbNullString)
    TextEditBox& = FindWindowEx(ComboBox&, 0, "Edit", vbNullString)
    TextSet& = SendMessageByString(TextEditBox&, WM_SETTEXT, 0, SendName$)
 
    EditThing& = FindWindowEx(IMWin&, 0, "WndAte32Class", vbNullString)
    EditThing& = GetWindow(EditThing&, 2)
    SendButtin& = FindWindowEx(IMWin&, 0, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(SendButtin&, WM_ENABLE, 3&, 3&)

    Call PostMessage(SendButtin&, WM_KEYDOWN, VK_SPACE, 0&)
    Call PostMessage(SendButtin&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Public Sub LB_Item_Remove(LB As ListBox, Value As String)
'removes an item from a listbox with out knowing the index...all u do is type the name of the item to remove for value
    Dim I As Integer
    For I = 0 To LB.ListCount - 1
        If LCase$(LB.List(I)) = LCase$(Value$) Then
            LB.RemoveItem (I)
        End If
    Next I

End Sub

Public Function Keep(String1 As String, LettersToKeep)
'allows u to tell it a string and it will take out all the characters that u dont tell it to keep
    Dim String2 As String, I As Integer, Letter As String, InString As Integer
    For I = 1 To Len(String1$)
        Letter$ = Mid(String1$, I, 1)
        InString = InStr(LettersToKeep, Letter$)
        If InString <> 0 Then
            String2$ = String2$ & Letter$
        End If
    Next I
    Keep = String2$
    
End Function

Public Function RemoveHTML(String1 As String)
'removes all the html characters form a string
    Dim FH As String, LH As String, LocOfLT As Long, LocOfGT As Long
    Do Until InStr(String1$, "<") = 0 Or InStr(String1$, ">") = 0
        If InStr(String1$, "<") > InStr(String1$, ">") Then
            Exit Do
        End If
        LocOfLT& = InStr(String1$, "<")
        LocOfGT& = InStr(String1$, ">")
        FH$ = Left$(String1$, LocOfLT& - 1)
        LH$ = Mid$(String1$, LocOfGT& + 1)
        String1$ = FH$ & LH$
    Loop
    RemoveHTML = String1$

End Function

Sub Form_OnTop(Form As Form)
'sets a form on top of all other forms
    Call SetWindowPos(Form.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub

Sub Form_NotOnTop(Form As Form)
'sets a form back to regular after u use Form_OnTop
    Call SetWindowPos(Form.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub

Sub IM_Send(SendName As String, SayWhat As String)
' Example: Call IM_Send("ThereSn","Sup man")
    Dim BuddyList As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If

Start:
 
    Dim TabWin As Long, IMbuttin As Long, IMWin As Long
    Dim ComboBox As Long, TextEditBox As Long, TextSet As Long
    Dim EditThing As Long, TextSet2 As Long, SendButtin As Long, Click As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    TabWin& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
    IMbuttin& = FindWindowEx(TabWin&, 0, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(IMbuttin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(IMbuttin&, WM_LBUTTONUP, 0, 0&)
  
  
    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    ComboBox& = FindWindowEx(IMWin&, 0, "_Oscar_PersistantCombo", vbNullString)
    TextEditBox& = FindWindowEx(ComboBox&, 0, "Edit", vbNullString)
    TextSet& = SendMessageByString(TextEditBox&, WM_SETTEXT, 0, SendName$)
 
    EditThing& = FindWindowEx(IMWin&, 0, "WndAte32Class", vbNullString)
    EditThing& = GetWindow(EditThing&, 2)
    TextSet2& = SendMessageByString(EditThing&, WM_SETTEXT, 0, SayWhat$)
    SendButtin& = FindWindowEx(IMWin&, 0, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(SendButtin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(SendButtin&, WM_LBUTTONUP, 0, 0&)
    
End Sub

Sub AiM_Attention(Text As String)
'umm u know what this is
Aim_ChatSend ("(-----Attention-----)")
Pause 1#
Aim_ChatSend (Text)
Pause 1#
Aim_ChatSend ("(-----Attention-----)")
End Sub

Public Function OpenURL(ByVal URL As String) As Long
'alot of ppl always ask me how to do this so i had to add it...
'it opens the default browser to the url specified... EXAMPLE:
'Sub Label1_Click()
'OpenURL "http://acidfux.virtualave.net"
'End Sub
    OpenURL = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function
