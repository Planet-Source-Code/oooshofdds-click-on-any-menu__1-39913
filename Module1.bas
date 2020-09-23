Attribute VB_Name = "Module1"
'module by Stacked_shit@yahoo.com
           'Stacked_Shit@ hotmail.com
           'Webmaster@ immortal - hackers.com
           'dont forget to check my site out
           'http://www.immortal-hackers.com
           'good luck
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long
Public Declare Function GetMenuContextHelpId Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal fByPos As Long, ByVal gmdiFlags As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const WM_COMMAND = &H111
'on top constants
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_NOTOPMOST = -2

Sub RunMenubystring(Window, mnuCap)
    'Gimme credits if you use it in your projects
    Dim ToSearch        As Long
    Dim MenuCount       As Integer
    Dim FindString
    Dim ToSearchSub     As Long
    Dim MenuItemCount   As Integer
    Dim GetString
    Dim SubCount        As Long
    Dim MenuString      As String
    Dim GetStringMenu   As Integer
    Dim MenuItem        As Long
    Dim RunTheMenu      As Integer
    
    ToSearch& = GetMenu(Window)
    MenuCount% = GetMenuItemCount(ToSearch&)
    
    For FindString = 0 To MenuCount% - 1
        ToSearchSub& = GetSubMenu(ToSearch&, FindString)
        MenuItemCount% = GetMenuItemCount(ToSearchSub&)
        For GetString = 0 To MenuItemCount% - 1
            SubCount& = GetMenuItemID(ToSearchSub&, GetString)
            MenuString$ = String$(100, " ")
            GetStringMenu% = GetMenuString(ToSearchSub&, SubCount&, MenuString$, 100, 1)
            If InStr(UCase(MenuString$), UCase(mnuCap)) Then
                MenuItem& = SubCount&
                GoTo MatchString
            End If
    Next GetString
    Next FindString
MatchString:
    RunTheMenu% = SendMessage(Window, WM_COMMAND, MenuItem&, 0)
End Sub
Public Sub StayOnTop(Frm As Form)
    Call SetWindowPos(Frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub
Public Sub dontstayontop(Frm As Form)
    Call SetWindowPos(Frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
End Sub
Sub Pause(interval)
Dim Current
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub
