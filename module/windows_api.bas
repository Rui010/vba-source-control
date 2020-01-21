Attribute VB_Name = "windowsApiModule"

'/////////////////////////////////
'// Windows API�錾
'/////////////////////////////////

'Sleep�@�\�𗘗p����
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
        
Declare PtrSafe Function FindWindowEx Lib "user32.dll" _
    Alias "FindWindowExA" ( _
    ByVal hWndParent As Long, _
    ByVal hwndChildAfter As Long, _
    ByVal lpszClass As String, _
    ByVal lpszWindow As String) As Long
    
'Window�擾���C�u�����𗘗p����
Declare PtrSafe Function FindWindow Lib "USER32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
     ByVal lpWindowName As String) As Long
        
'�L�[���색�C�u�����𗘗p����
Declare PtrSafe Function SendMessage Lib "user32.dll" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal msg As Long, _
     ByVal wParam As Long, ByVal lParam As Any) As Long
