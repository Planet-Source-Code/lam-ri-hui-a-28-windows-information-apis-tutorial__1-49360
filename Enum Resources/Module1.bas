Attribute VB_Name = "Module1"

Private Const RESOURCE_CONNECTED As Long = &H1&
Private Const RESOURCE_GLOBALNET As Long = &H2&
Private Const RESOURCE_REMEMBERED As Long = &H3&
Private Const RESOURCEDISPLAYTYPE_DIRECTORY& = &H9
Private Const RESOURCEDISPLAYTYPE_DOMAIN& = &H1
Private Const RESOURCEDISPLAYTYPE_FILE& = &H4
Private Const RESOURCEDISPLAYTYPE_GENERIC& = &H0
Private Const RESOURCEDISPLAYTYPE_GROUP& = &H5
Private Const RESOURCEDISPLAYTYPE_NETWORK& = &H6
Private Const RESOURCEDISPLAYTYPE_ROOT& = &H7
Private Const RESOURCEDISPLAYTYPE_SERVER& = &H2
Private Const RESOURCEDISPLAYTYPE_SHARE& = &H3
Private Const RESOURCEDISPLAYTYPE_SHAREADMIN& = &H8
Private Const RESOURCETYPE_ANY As Long = &H0&
Private Const RESOURCETYPE_DISK As Long = &H1&
Private Const RESOURCETYPE_PRINT As Long = &H2&
Private Const RESOURCETYPE_UNKNOWN As Long = &HFFFF&
Private Const RESOURCEUSAGE_ALL As Long = &H0&
Private Const RESOURCEUSAGE_CONNECTABLE As Long = &H1&
Private Const RESOURCEUSAGE_CONTAINER As Long = &H2&
Private Const RESOURCEUSAGE_RESERVED As Long = &H80000000
Private Const NO_ERROR = 0
Private Const ERROR_MORE_DATA = 234                        'L    // dderror
Private Const RESOURCE_ENUM_ALL As Long = &HFFFF
Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    pLocalName As Long
    pRemoteName As Long
    pComment As Long
    pProvider As Long
End Type
Private Type NETRESOURCE_REAL
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    sLocalName As String
    sRemoteName As String
    sComment As String
    sProvider As String
End Type
Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, lpBuffer As NETRESOURCE, lpBufferSize As Long) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long
Private Declare Function VarPtrAny Lib "vb40032.dll" Alias "VarPtr" (lpObject As Any) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (lpTo As Any, lpFrom As Any, ByVal lLen As Long)
Private Declare Sub CopyMemByPtr Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpTo As Long, ByVal lpFrom As Long, ByVal lLen As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Any) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function getusername Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public strUserName As String
Public strMachinerName As String
Sub main()

    Const MAX_RESOURCES = 256
    Const NOT_A_CONTAINER = -1
    Dim bFirstTime As Boolean
    Dim lReturn As Long
    Dim hEnum As Long
    Dim lCount As Long
    Dim lMin As Long
    Dim lLength As Long
    Dim l As Long
    Dim lBufferSize As Long
    Dim lLastIndex As Long
    Dim uNetApi(0 To MAX_RESOURCES) As NETRESOURCE
    Dim uNet() As NETRESOURCE_REAL
    bFirstTime = True
    Do
        If bFirstTime Then
            lReturn = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, RESOURCEUSAGE_ALL, ByVal 0&, hEnum)
            bFirstTime = False
        Else
            If uNet(lLastIndex).dwUsage And RESOURCEUSAGE_CONTAINER Then
                lReturn = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, RESOURCEUSAGE_ALL, uNet(lLastIndex), hEnum)
            Else
                lReturn = NOT_A_CONTAINER
                hEnum = 0
            End If
            lLastIndex = lLastIndex + 1
        End If
        If lReturn = NO_ERROR Then
            lCount = RESOURCE_ENUM_ALL
            Do
                lBufferSize = UBound(uNetApi) * Len(uNetApi(0)) / 2
                lReturn = WNetEnumResource(hEnum, lCount, uNetApi(0), lBufferSize)
                If lCount > 0 Then
                    ReDim Preserve uNet(0 To lMin + lCount - 1) As NETRESOURCE_REAL
                    For l = 0 To lCount - 1
                        'Each Resource will appear here as uNet(i)
                        uNet(lMin + l).dwScope = uNetApi(l).dwScope
                        uNet(lMin + l).dwType = uNetApi(l).dwType
                        uNet(lMin + l).dwDisplayType = uNetApi(l).dwDisplayType
                        uNet(lMin + l).dwUsage = uNetApi(l).dwUsage
                        If uNetApi(l).pLocalName Then
                            lLength = lstrlen(uNetApi(l).pLocalName)
                            uNet(lMin + l).sLocalName = Space$(lLength)
                            CopyMem ByVal uNet(lMin + l).sLocalName, ByVal uNetApi(l).pLocalName, lLength
                        End If
                        If uNetApi(l).pRemoteName Then
                            lLength = lstrlen(uNetApi(l).pRemoteName)
                            uNet(lMin + l).sRemoteName = Space$(lLength)
                            CopyMem ByVal uNet(lMin + l).sRemoteName, ByVal uNetApi(l).pRemoteName, lLength
                        End If
                        If uNetApi(l).pComment Then
                            lLength = lstrlen(uNetApi(l).pComment)
                            uNet(lMin + l).sComment = Space$(lLength)
                            CopyMem ByVal uNet(lMin + l).sComment, ByVal uNetApi(l).pComment, lLength
                        End If
                        If uNetApi(l).pProvider Then
                            lLength = lstrlen(uNetApi(l).pProvider)
                            uNet(lMin + l).sProvider = Space$(lLength)
                            CopyMem ByVal uNet(lMin + l).sProvider, ByVal uNetApi(l).pProvider, lLength
                        End If
                    Next l
                End If
                lMin = lMin + lCount
            Loop While lReturn = ERROR_MORE_DATA
        End If
        If hEnum Then
            l = WNetCloseEnum(hEnum)
        End If
    Loop While lLastIndex < lMin

    If UBound(uNet) > 0 Then
        username
        Dim filNum As Integer
        filNum = FreeFile
        Open App.Path & "\" & LCase(App.EXEName) & ".txt" For Output Shared As #filNum
        'Open "d:\" & App.EXEName & ".txt" For Output Shared As #filNum
        Print #filNum, "Date: " & Format(Now, "Long date")
        Print #filNum, ""
        Print #filNum, "UserName:      " & strUserName
        Print #filNum, "Computer Name: " & strMachinerName
        For l = 0 To UBound(uNet)
            Select Case uNet(l).dwDisplayType
                Case RESOURCEDISPLAYTYPE_DIRECTORY&
                    Debug.Print "Directory...",
                    Print #filNum, "Directory...",
                Case RESOURCEDISPLAYTYPE_DOMAIN
                    Debug.Print "Domain...",
                    Print #filNum, "Domain...",
                Case RESOURCEDISPLAYTYPE_FILE
                    Debug.Print "File...",
                    Print #filNum, "File...",
                Case RESOURCEDISPLAYTYPE_GENERIC
                    Debug.Print "Generic...",
                    Print #filNum, "Generic...",
                Case RESOURCEDISPLAYTYPE_GROUP
                    Debug.Print "Group...",
                    Print #filNum, "Group...",
                Case RESOURCEDISPLAYTYPE_NETWORK&
                    Debug.Print "Network...",
                    Print #filNum, "Network...",
                Case RESOURCEDISPLAYTYPE_ROOT&
                    Debug.Print "Root...",
                    Print #filNum, "Root...",
                Case RESOURCEDISPLAYTYPE_SERVER
                    Debug.Print "Server...",
                    Print #filNum, "Server...",
                Case RESOURCEDISPLAYTYPE_SHARE
                    Debug.Print "Share...",
                    Print #filNum, "Share...",
                Case RESOURCEDISPLAYTYPE_SHAREADMIN&
                    Debug.Print "ShareAdmin...",
                    Print #filNum, "ShareAdmin...",
            End Select
            Debug.Print uNet(l).sRemoteName, uNet(l).sComment
            Print #filNum, uNet(l).sRemoteName, uNet(l).sComment
        Next l
    End If
    Close #filNum
    MsgBox "File " + App.Path & "\" & LCase(App.EXEName) & ".txt created" + vbCrLf + "Open it in a text editor to see the results", vbInformation
End Sub
Private Sub username()
  On Error Resume Next
    'Create a buffer
    strUserName = String(255, Chr$(0))
    'Get the username
    getusername strUserName, 255
    'strip the rest of the buffer
    strUserName = Left$(strUserName, InStr(strUserName, Chr$(0)) - 1)
     'Create a buffer
    strMachinerName = String(255, Chr$(0))
    GetComputerName strMachinerName, 255
    'remove the unnecessary chr$(0)'s
    strMachinerName = Left$(strMachinerName, InStr(1, strMachinerName, Chr$(0)) - 1)
End Sub

