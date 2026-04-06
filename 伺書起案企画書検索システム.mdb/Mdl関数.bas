Attribute VB_Name = "Mdl関数"
Option Compare Database

'========== IMEの制御用 Win32API ===================================================================================================
'クラス名とウィンドウ名により指定されたウィンドウハンドルを
'取得する関数の宣言
Public Declare Function FindWindow Lib "User32" _
    Alias "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

'ウィンドウに関連付けされた入力コンテキストを取得する関数の宣言
Public Declare Function ImmGetContext Lib "imm32.dll" _
    (ByVal hWnd As Long) As Long

'IMEのオープン状態を設定する関数の宣言
Public Declare Function ImmSetOpenStatus Lib "imm32.dll" _
    (ByVal himc As Long, ByVal b As Long) As Long

'ウィンドウに関連付けされた入力コンテキストを開放する関数の宣言
Public Declare Function ImmReleaseContext Lib "imm32.dll" _
    (ByVal hWnd As Long, ByVal himc As Long) As Long

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long

'API functions to be used
Private Declare Function CallNextHookEx Lib "User32" (ByVal hHook As Long, _
                                                      ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Declare Function SetWindowsHookEx Lib "User32" Alias "SetWindowsHookExA" _
                                          (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, _
                                          ByVal dwThreadId As Long) As Long

Private Declare Function UnhookWindowsHookEx Lib "User32" (ByVal hHook As Long) As Long

Private Declare Function SendDlgItemMessage Lib "User32" Alias "SendDlgItemMessageA" _
                                            (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, _
                                            ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hWnd As Long, _
                                                                          ByVal lpClassName As String, _
                                                                          ByVal nMaxCount As Long) As Long

Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

'Constants to be used in our API functions
Private Const EM_SETPASSWORDCHAR = &HCC
Private Const WH_CBT = 5
Private Const HCBT_ACTIVATE = 5
Private Const HC_ACTION = 0

Private hHook As Long


Public Function InputBoxDK(Prompt, Optional Title, Optional Default, Optional XPos, _
                        Optional YPos, Optional HelpFile, Optional Context) As String
    Dim lngModHwnd As Long, lngThreadID As Long

    lngThreadID = GetCurrentThreadId
    lngModHwnd = GetModuleHandle(vbNullString)

    hHook = SetWindowsHookEx(WH_CBT, AddressOf NewProc, lngModHwnd, lngThreadID)

    InputBoxDK = InputBox(Prompt, Title, Default, XPos, YPos, HelpFile, Context)
    UnhookWindowsHookEx hHook

End Function

Public Function NewProc(ByVal lngCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim RetVal
    Dim strClassName As String, lngBuffer As Long

    If lngCode < HC_ACTION Then
        NewProc = CallNextHookEx(hHook, lngCode, wParam, lParam)
        Exit Function
    End If

    strClassName = String$(256, " ")
    lngBuffer = 255

    If lngCode = HCBT_ACTIVATE Then    'A window has been activated

        RetVal = GetClassName(wParam, strClassName, lngBuffer)

        If Left$(strClassName, RetVal) = "#32770" Then  'Class name of the Inputbox

            'This changes the edit control so that it display the password character *.
            'You can change the Asc("*") as you please.
            SendDlgItemMessage wParam, &H1324, EM_SETPASSWORDCHAR, Asc("*"), &H0
        End If

    End If

End Function

Function GetMyComputerName() As String

'自分のパソコンのコンピュータ名を返します。
Dim strCmptrNameBuff As String * 21

'API関数によってコンピューター名を取得します。コンピュータ名は変数strCmptrNameBuffに返されます。
GetComputerName strCmptrNameBuff, Len(strCmptrNameBuff)

'後続のNullを取り除いて返り値を設定します。
GetMyComputerName = Left$(strCmptrNameBuff, InStr(strCmptrNameBuff, vbNullChar) - 1)

End Function
