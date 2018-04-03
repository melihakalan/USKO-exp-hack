Attribute VB_Name = "Module1"
Option Explicit
Public SndFnc(1 To 10) As Long
Public Keyboard As Long
Public Const DIK_0 As Long = 11
Public Const DIK_1 As Long = 2
Public Const DIK_2 As Long = 3
Public Const DIK_3 As Long = 4
Public Const DIK_4 As Long = 5
Public Const DIK_5 As Long = 6
Public Const DIK_6 As Long = 7
Public Const DIK_7 As Long = 8
Public Const DIK_8 As Long = 9
Public Const DIK_9 As Long = 10
Public Const DIK_F1 As Long = &H3B
Public Const DIK_F2 As Long = 60
Public Const DIK_F3 As Long = &H3D
Public Const DIK_F4 As Long = &H3E
Public Const DIK_F5 As Long = &H3F
Public Const DIK_F6 As Long = &H40
Public Const DIK_F7 As Long = &H41
Public Const DIK_F8 As Long = &H42
Public Const DIK_Z As Long = &H2C
Public Const DIK_C As Long = &H2E
Public Const DIK_B As Long = &H30
Public Const DIK_R As Long = &H13
Public Const DIK_TAB As Long = 15
Public Const DIK_E As Long = &H12
Public Const KeybPtr As Long = &HB6FC5C

Private BytesAddr As Long
Private FuncPtr As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Const MEM_COMMIT = &H1000
Private Const MEM_RELEASE = &H8000&
Private Const PAGE_READWRITE = &H4&
Private Const INFINITE = &HFFFF
Public Const MAILSLOT_NO_MESSAGE  As Long = (-1)
Public KOPHandle As Long, KO_CHRDMA As Long, KO_WLF As Long, KoPath As String, KOPid As Long
Public Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long 'apidir bunu ekle modulde yukarýya
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function ReadProcessMem Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function WriteProcessMem Lib "kernel32" Alias "WriteProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Long) As Long
Private Declare Function CreateMailslot Lib "kernel32" Alias "CreateMailslotA" (ByVal lpName As String, ByVal nMaxMessageSize As Long, ByVal lReadTimeout As Long, lpSecurityAttributes As Any) As Long
Private Declare Function GetMailslotInfo Lib "kernel32" (ByVal hMailSlot As Long, lpMaxMessageSize As Long, lpNextSize As Long, lpMessageCount As Long, lpReadTimeout As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetModuleInformation Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, lpmodinfo As MODULEINFO, ByVal cb As Long) As Long

Public Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
'KOJD WILL DECLARE SOMETHING
Public CancelWalk As Boolean
Public OffsetKO_OFF_X As Long
Public OffsetKO_OFF_Y As Long
Public OffsetKO_OFF_MOVTYPE As Long
Public OffsetKO_OFF_MVCHRTYP As Long
'KOJD HAS DECLARED SOMETHING
Public hexword As String
Public Rbugg As Long
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const LWA_COLORKEY = 1
Public Const LWA_ALPHA = 2
Public ko As Long
Public Const LWA_BOTH = 3
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = -20
Public Timer As Date
Public Diff As Long
Public UseOtocanPot
Public Hook
Public CurrentNation As String
Public Class As String
Public UseTimed As Boolean
Public PotionTimer As Date
Public Char As Long
Public TimedTimer As Date
Public AttackTimer As Long
Public AttackDiff As Long
Public TargetTimer As Date
Public TargetDiff As Long
Public TimedDiff As Long
Public PotionDiff As Long
Public PublicWolf As Long
Public UseAutoWolf
Public UseOtomanapot
Public UseAutoAttack
Public UseManaSave
Public pointer As Long
Public X As Long
Public Y As Long
Public Z As Long
Public x_cursor As Long
Public y_cursor As Long
Public Started
Public HealPercent As Long
Public ManaPercent As Long
Public HealType As Long
Public ManaType As Long
Public AttackALL As String
Public mob As Long
Public SellAll As Long
Public BonusFilter As Long
Public MSName
Public MSHandle
Public UseSitAutoAttack
Public UseWallHack
Public durab As Long
Public UseLupineEyes
Public UseAutoSwift
Public UseAutoSell
Public UseAutoLoot
Public CurrentMobHP As Long
Public BoxID
Public itemID
Public AttackNow
Public SecondID
Public ThirdID
Public FourthID
Public FifthID
Public SixthID
Public BoxOpened
Public Exp As Long
Public outack
Public Looting
Public ChatTipi As String
Public RepairID As String
Public ItemSlot As String
Public RecvID As String
Public LastBoxID As Long
Public OpenNextBox As Boolean
Public LastRepair As Date
Public RepairDiff As Long
Public LootBuffer As String
'start declaring autoattack log
Public LastID
Public TargetID As Long
Public MaxExp As Long
'stop declaring autoattack log
Public ChatTipi1 As Long
Public KO_RCVFNC As Long
Public KO_TITLE As String
Public KO_HANDLE As Long
Public KO_PID As Long
Public KO_PTR_CHR As Long
Public KO_PTR_CHR1 As Long
Public KO_PTR_DLG As Long
Public KO_PTR_PKT As Long
Public KO_SND_FNC As Long
Public KO_RECVHK As Long
Public KO_RCVHKB As Long
Public KO_ADR_CHR As Long
Public KO_ADR_DLG As Long
Public KO_OFF_RTM As Long
Public KO_OFF_NT As Long
Public KO_OFF_CLASS As Long
Public KO_OFF_SWIFT As Long
Public KO_OFF_HP As Long
Public KO_OFF_MAXHP As Long
Public KO_OFF_MP As Long
Public KO_OFF_MAXMP As Long
Public KO_OFF_MOB As Long
Public KO_OFF_WH As Long
Public KO_OFF_SIT As Long
Public KO_OFF_LUP As Long
Public KO_OFF_LUP2 As Long
Public KO_OFF_Y As Long
Public KO_OFF_ID As Long
Public KO_OFF_HD As Long
Public KO_OFF_X As Long
Public KO_OFF_SEL As Long
Public KO_OFF_LUPINE As Long
Public KO_OFF_EXP As Long
Public KO_OFF_MAXEXP As Long
Public KO_OFF_AP As Long
Public KO_OFF_GOLD As Long
Public KO_DLGBMA As Long
Public KO_KEYBPTR As Long
Public nation As Long
Public KO_TITLE1 As String
Public DINPUT_Handle As Long
Public DINPUT_lpBaseOfDLL As Long
Public DINPUT_SizeOfImage As Long
Public DINPUT_EntryPoint As Long
Public DINPUT_KEYDMA As Long
Public DINPUT_K_1 As Long
Public DINPUT_K_2 As Long
Public DINPUT_K_3 As Long
Public DINPUT_K_4 As Long
Public DINPUT_K_5 As Long
Public DINPUT_K_6 As Long
Public DINPUT_K_7 As Long
Public DINPUT_K_8 As Long
Public DINPUT_K_Z As Long
Public DINPUT_K_C As Long
Public DINPUT_K_S As Long
Public DINPUT_K_R As Long
Public KO_CHR As Long
Public MP1 As Long
Public Type MODULEINFO
lpBaseOfDLL As Long
SizeOfImage As Long
EntryPoint As Long
End Type
Public Function Txt2Code(txtCode As String)
Dim Kodlar(0 To 35) As String
Dim Ayirma
Dim I, ii, ab As Integer
Dim Yazi
Dim YaziSekli As String
Kodlar(0) = "q": Kodlar(1) = "w": Kodlar(2) = "e": Kodlar(3) = "r": Kodlar(4) = "t"
Kodlar(5) = "y": Kodlar(6) = "u": Kodlar(7) = "ý": Kodlar(8) = "o": Kodlar(9) = "p": Kodlar(10) = "ð"
Kodlar(11) = "ü": Kodlar(12) = "a": Kodlar(13) = "s": Kodlar(14) = "d": Kodlar(15) = "f": Kodlar(16) = "g"
Kodlar(17) = "h": Kodlar(18) = "j": Kodlar(19) = "k": Kodlar(20) = "l": Kodlar(21) = "þ": Kodlar(22) = "i"
Kodlar(23) = "z": Kodlar(24) = "x": Kodlar(25) = "c": Kodlar(26) = "v": Kodlar(27) = "b": Kodlar(27) = "n"
Kodlar(28) = "m": Kodlar(29) = "ö": Kodlar(30) = "ç"
ii = Len(txtCode) * 2
For I = 1 To ii
If Mid(txtCode, I, 1) = "#" Then
YaziSekli = "Büyük"
GoTo gecbunu
End If
If Mid(txtCode, I, 1) = "," Then
ab = val(Ayirma)
Ayirma = ""
If YaziSekli = "Büyük" Then
YaziSekli = ""
Yazi = Yazi & UCase(Kodlar(ab))
Else
Yazi = Yazi & Kodlar(ab)
End If
Else
Ayirma = Ayirma & Mid(txtCode, I, 1)
gecbunu:
End If
Next
Txt2Code = LTrim(RTrim(Yazi))
End Function

'Kullanýþý .... yazdir = Txt2Code("#13,2,#19,12,")
Public Function ko1()
KO_TITLE1 = "Knight OnLine Client"
    GetWindowThreadProcessId FindWindow(vbNullString, KO_TITLE1), KO_PID
    KO_HANDLE = OpenProcess(&H1F0FFF, False, KO_PID)
KO_PTR_CHR1 = ReadLong(&HB6E3BC)
MP1 = 2364
End Function
' dll inject komutlarý
Public Function HookDI8() As Boolean
Dim ret As Long
Dim lmodinfo As MODULEINFO
DINPUT_Handle = 0

DINPUT_Handle = FindModuleHandle("dinput8.dll")


ret = GetModuleInformation(KO_HANDLE, DINPUT_Handle, lmodinfo, Len(lmodinfo))
If ret <> 0 Then
With lmodinfo
DINPUT_EntryPoint = .EntryPoint
DINPUT_lpBaseOfDLL = .lpBaseOfDLL
DINPUT_SizeOfImage = .SizeOfImage
End With
Else
Exit Function
End If
SetupDInput
HookDI8 = True
End Function

Public Function FindModuleHandle(ModuleName As String) As Long
Dim hModules(1 To 256) As Long
Dim BytesReturned As Long
Dim ModuleNumber As Byte
Dim TotalModules As Byte
Dim FileName As String * 128
Dim ModName As String
EnumProcessModules KO_HANDLE, hModules(1), 1024, BytesReturned
TotalModules = BytesReturned / 4
For ModuleNumber = 1 To TotalModules
GetModuleFileNameExA KO_HANDLE, hModules(ModuleNumber), FileName, 128
ModName = Left(FileName, InStr(FileName, Chr(0)) - 1)
If UCase(Right(ModName, Len(ModuleName))) = UCase(ModuleName) Then
FindModuleHandle = hModules(ModuleNumber)
End If
Next
End Function

Sub SetupDInput()
DINPUT_KEYDMA = FindDInputKeyPtr
If DINPUT_KEYDMA <> 0 Then
DINPUT_K_1 = DINPUT_KEYDMA + 2
DINPUT_K_2 = DINPUT_KEYDMA + 3
DINPUT_K_3 = DINPUT_KEYDMA + 4
DINPUT_K_4 = DINPUT_KEYDMA + 5
DINPUT_K_5 = DINPUT_KEYDMA + 6
DINPUT_K_6 = DINPUT_KEYDMA + 7
DINPUT_K_7 = DINPUT_KEYDMA + 8
DINPUT_K_8 = DINPUT_KEYDMA + 9
DINPUT_K_Z = DINPUT_KEYDMA + 44
DINPUT_K_C = DINPUT_KEYDMA + 46
DINPUT_K_S = DINPUT_KEYDMA + 31
DINPUT_K_R = DINPUT_KEYDMA + 19
End If
End Sub

Function FindDInputKeyPtr() As Long
Dim pBytes() As Byte
Dim pSize As Long
Dim X As Long
pSize = DINPUT_SizeOfImage
ReDim pBytes(1 To pSize)
ReadByteArray DINPUT_lpBaseOfDLL, pBytes, pSize
For X = 1 To pSize - 10
If pBytes(X) = &H57 And pBytes(X + 1) = &H6A And pBytes(X + 2) = &H40 And pBytes(X + 3) = &H33 And pBytes(X + 4) = &HC0 And pBytes(X + 5) = &H59 And pBytes(X + 6) = &HBF Then
FindDInputKeyPtr = val("&H" & IIf(Len(Hex(pBytes(X + 10))) = 1, "0" & Hex(pBytes(X + 10)), Hex(pBytes(X + 10))) & IIf(Len(Hex(pBytes(X + 9))) = 1, "0" & Hex(pBytes(X + 9)), Hex(pBytes(X + 9))) & IIf(Len(Hex(pBytes(X + 8))) = 1, "0" & Hex(pBytes(X + 8)), Hex(pBytes(X + 8))) & IIf(Len(Hex(pBytes(X + 7))) = 1, "0" & Hex(pBytes(X + 7)), Hex(pBytes(X + 7))))
Exit For
End If
Next
End Function
' Buraya ben yolla yazdým sizde istediðinizi yaza bilir siniz.
'ama prejedeki Bütün Yolla yazan yerleri deðiþtirmelisiniz.
Function yolla(pKey As String) As Long
pKey = Strings.UCase(pKey)
Select Case pKey
Case "S"
yolla = DINPUT_K_S
Case "Z"
yolla = DINPUT_K_Z
Case "1"
yolla = DINPUT_K_1
Case "2"
yolla = DINPUT_K_2
Case "3"
yolla = DINPUT_K_3
Case "4"
yolla = DINPUT_K_4
Case "5"
yolla = DINPUT_K_5
Case "6"
yolla = DINPUT_K_6
Case "7"
yolla = DINPUT_K_7
Case "8"
yolla = DINPUT_K_8
Case "C"
yolla = DINPUT_K_C
Case "R"
yolla = DINPUT_K_R
End Select
End Function
Sub WriteByte(Addr As Long, pVal As Byte)
Dim pbw As Long
WriteProcessMem KO_HANDLE, Addr, pVal, 1, pbw
End Sub

Sub ReadByteArray(Addr As Long, pmem() As Byte, pSize As Long)
Dim Value As Byte
ReDim pmem(1 To pSize) As Byte
ReadProcessMem KO_HANDLE, Addr, pmem(1), pSize, 0&
End Sub
' Buraya ben TUS yazdým sizde istediðinizi yaza bilir siniz.
'ama prejedeki Bütün TUS yazan yerleri deðiþtirmelisiniz.
Sub Tuþ(pKey As Long, Optional pTimeMS As Long = 50)
WriteByte pKey, 128
f_Sleep pTimeMS, True
WriteByte pKey, 0
End Sub

Sub f_Sleep(pMS As Long, Optional pDoevents As Boolean = False)
Dim pTime As Long
pTime = GetTickCount
Do While pMS + pTime > GetTickCount
If pDoevents = True Then DoEvents
Loop
End Sub




Function ReadByte(pAddy As Long, Optional pHandle As Long) As Byte
    Dim Value As Byte
    
    If pHandle <> 0 Then
        ReadProcessMem pHandle, pAddy, Value, 1, 0&
    Else
        ReadProcessMem KOPHandle, pAddy, Value, 1, 0&
    End If
    ReadByte = Value

End Function
Public Function ReadLong(Addr As Long) As Long 'read a 4 byte value
    Dim Value As Long
    ReadProcessMem KO_HANDLE, Addr, Value, 4, 0&
    ReadLong = Value
End Function

Public Function ReadFloat(Addr As Long) As Long 'read a float value
    Dim Value As Single
    ReadProcessMem KO_HANDLE, Addr, Value, 4, 0&
    ReadFloat = Value
End Function

Public Function WriteFloat(Addr As Long, val As Single) 'write a float value
    WriteProcessMem KO_HANDLE, Addr, val, 4, 0&
End Function

Public Function WriteLong(Addr As Long, val As Long) ' write a 4 byte value
    WriteProcessMem KO_HANDLE, Addr, val, 4, 0&
End Function

Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) As Long
    If Topmost = True Then 'Make the window topmost
        SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
        SetTopMostWindow = False
    End If
End Function

Public Function ReadIni(FileName As String, Section As String, Key As String) As String
Dim RetVal As String * 255, v As Long
v = GetPrivateProfileString(Section, Key, "", RetVal, 255, FileName)
ReadIni = Left(RetVal, v)
End Function

'writes an Ini string
Public Function WriteIni(FileName As String, Section As String, Key As String, Value As String)
WritePrivateProfileString Section, Key, Value, FileName
End Function
Public Function ConvHEX2ByteArray(pStr As String, pbyte() As Byte)
On Error Resume Next
Dim I As Long
Dim j As Long
ReDim pbyte(1 To Len(pStr) / 2)

j = LBound(pbyte) - 1
For I = 1 To Len(pStr) Step 2
    j = j + 1
    pbyte(j) = CByte("&H" & Mid(pStr, I, 2))
Next
End Function
Public Function WriteByteArray(pA As Long, pS() As Byte, pSize As Long)
    WriteProcessMem KO_HANDLE, pA, pS(LBound(pS)), pSize, 0&
End Function
'start packet handling
Function ExecuteRemoteCode(pCode() As Byte, Optional WaitExecution As Boolean = False) As Long
Dim hThread As Long, ThreadID As Long, ret As Long
Dim SE As SECURITY_ATTRIBUTES

SE.nLength = Len(SE)
SE.bInheritHandle = False

ExecuteRemoteCode = 0
If FuncPtr = 0 Then
FuncPtr = VirtualAllocEx(KO_HANDLE, 0, 1024, MEM_COMMIT, PAGE_READWRITE)
End If
If FuncPtr <> 0 Then
    WriteByteArray FuncPtr, pCode, UBound(pCode) - LBound(pCode) + 1
   
   hThread = CreateRemoteThread(ByVal KO_HANDLE, SE, 0, ByVal FuncPtr, 0&, 0&, ThreadID)
   If hThread Then
      ret = WaitForSingleObject(hThread, INFINITE)
      ExecuteRemoteCode = ThreadID
   End If
   CloseHandle hThread
   ret = VirtualFreeEx(KO_HANDLE, FuncPtr, 0, MEM_RELEASE)
    
End If

End Function

Function SendPackets(pPacket() As Byte)
Dim pSize As Long
Dim pCode() As Byte

pSize = UBound(pPacket) - LBound(pPacket) + 1
If BytesAddr = 0 Then
BytesAddr = VirtualAllocEx(KO_HANDLE, 0, 1024, MEM_COMMIT, PAGE_READWRITE)
End If
If BytesAddr <> 0 Then
    WriteByteArray BytesAddr, pPacket, pSize
    ConvHEX2ByteArray "60A1" & AlignDWORD(KO_PTR_PKT) & "8B0D" & AlignDWORD(KO_PTR_PKT) & "68" & AlignDWORD(pSize) & "68" & AlignDWORD(BytesAddr) & "BF" & AlignDWORD(KO_SND_FNC) & "FFD761C3", pCode
    ExecuteRemoteCode pCode, True
End If
VirtualFreeEx KO_HANDLE, BytesAddr, 0, MEM_RELEASE&
End Function
':D alttaymýs zate:d
Function AlignDWORD(pParam As Long) As String
Dim HiW As Integer
Dim LoW As Integer

Dim HiBHiW As Byte
Dim HiBLoW As Byte

Dim LoBHiW As Byte
Dim LoBLoW As Byte

HiW = HiWord(pParam)
LoW = LoWord(pParam)

HiBHiW = HiByte(HiW)
HiBLoW = HiByte(LoW)

LoBHiW = LoByte(HiW)
LoBLoW = LoByte(LoW)

AlignDWORD = IIf(Len(Hex(LoBLoW)) = 1, "0" & Hex(LoBLoW), Hex(LoBLoW)) & _
         IIf(Len(Hex(HiBLoW)) = 1, "0" & Hex(HiBLoW), Hex(HiBLoW)) & _
         IIf(Len(Hex(LoBHiW)) = 1, "0" & Hex(LoBHiW), Hex(LoBHiW)) & _
         IIf(Len(Hex(HiBHiW)) = 1, "0" & Hex(HiBHiW), Hex(HiBHiW))
End Function 'sansa bak 8 miþ zate:D1:D
Function AlignDWORD8(pParam As Long) As String
Dim HiW As Integer
Dim LoW As Integer

Dim HiBHiW As Byte
Dim HiBLoW As Byte

Dim LoBHiW As Byte
Dim LoBLoW As Byte

HiW = HiWord(pParam)
LoW = LoWord(pParam)

HiBHiW = HiByte(HiW)
HiBLoW = HiByte(LoW)

LoBHiW = LoByte(HiW)
LoBLoW = LoByte(LoW)

AlignDWORD8 = IIf(Len(Hex(LoBLoW)) = 1, "0" & Hex(LoBLoW), Hex(LoBLoW)) & _
         IIf(Len(Hex(HiBLoW)) = 1, "0" & Hex(HiBLoW), Hex(HiBLoW)) & _
         IIf(Len(Hex(LoBHiW)) = 1, "0" & Hex(LoBHiW), Hex(LoBHiW)) & _
         IIf(Len(Hex(HiBHiW)) = 1, "0" & Hex(HiBHiW), Hex(HiBHiW)) & _
         IIf(Len(Hex(LoBLoW)) = 1, "0" & Hex(LoBLoW), Hex(LoBLoW)) & _
         IIf(Len(Hex(HiBLoW)) = 1, "0" & Hex(HiBLoW), Hex(HiBLoW)) & _
         IIf(Len(Hex(LoBHiW)) = 1, "0" & Hex(LoBHiW), Hex(LoBHiW)) & _
         IIf(Len(Hex(HiBHiW)) = 1, "0" & Hex(HiBHiW), Hex(HiBHiW))
End Function

Public Function HiByte(ByVal wParam As Integer) As Byte

    HiByte = (wParam And &HFF00&) \ (&H100)

End Function

Public Function LoByte(ByVal wParam As Integer) As Byte

LoByte = wParam And &HFF&

End Function

Function LoWord(DWord As Long) As Integer
   If DWord And &H8000& Then '
      LoWord = DWord Or &HFFFF0000
   Else
      LoWord = DWord And &HFFFF&
   End If
End Function

Function HiWord(DWord As Long) As Integer
   HiWord = (DWord And &HFFFF0000) \ &H10000
End Function

Public Function LogFile(sStr As String)
Open CurDir & "\Log.txt" For Append As #1
Print #1, ";==> " & Time & " " & sStr & " ;<=="
Close #1
End Function

Public Function AttachKO() As Boolean
If FindWindow(vbNullString, KO_TITLE) Then
'MSName = "\\.\mailslot\KOHACK0x" & Hex(GetTickCount)
'MShandle = EstablishMailSlot(MSName)
MSName = "\\.\mailslot\KOHack" & Hex(GetTickCount)
GetWindowThreadProcessId FindWindow(vbNullString, KO_TITLE), KO_PID
KO_HANDLE = OpenProcess(PROCESS_ALL_ACCESS, False, KO_PID)
If KO_HANDLE = 0 Then
MsgBox ("Cannot get handle from KO(" & KO_PID & ").")
AttachKO = False
End If
MSHandle = EstablishMailSlot(MSName)
If MSHandle = 0 Then End
'If KO_PID = 0 Then End
'HookRecvPackets3
HookDI8
AttachKO = True
Else
MsgBox ("can not handle ' " & KO_TITLE & " '"), vbCritical, "Hata"
End If
End Function

Public Function EstablishMailSlot(ByVal MailSlotName As String, Optional MaxMessageSize As Long = 0, Optional ReadTimeOut As Long = 50) As Long
EstablishMailSlot = CreateMailslot(MailSlotName, MaxMessageSize, ReadTimeOut, ByVal 0&)
End Function

Function LoadOffsets()
KO_TITLE = Form1.Text1.Text

KO_PTR_CHR = &HC05EE8
KO_PTR_DLG = &HC061D4
KO_PTR_PKT = &HC061A0
KO_SND_FNC = &H474E20
KO_RECVHK = &HB90CAC
KO_RCVHKB = &HB93470
KO_KEYBPTR = &HB84788

OffsetKO_OFF_X = &HCD4
OffsetKO_OFF_Y = &HCDC
OffsetKO_OFF_MOVTYPE = &HCC8
OffsetKO_OFF_MVCHRTYP = &H390

KO_OFF_SWIFT = 1582
KO_OFF_CLASS = 1416
KO_OFF_NT = 1412
KO_OFF_HP = &H5E4
KO_OFF_MAXHP = &H5E0
KO_OFF_MP = &H9A4
KO_OFF_MAXMP = &H9A8
KO_OFF_SIT = 2790
KO_OFF_MOB = &H580
KO_OFF_WH = 1432
KO_OFF_Y = 184
KO_OFF_X = 176
KO_OFF_LUP = 3204
KO_OFF_LUP2 = 2800
KO_OFF_EXP = 2072
KO_OFF_MAXEXP = 2068
KO_OFF_AP = 2136
KO_OFF_GOLD = 2064
KO_OFF_ID = &H5B4
KO_OFF_HD = 1300
nation = &H584
ko = ReadLong(KO_PTR_CHR)
Rbugg = &H59C
X = &HB4
Y = &HBC
Z = &H3E8
x_cursor = &HCD4
y_cursor = &HCDC
'//////ANVIL\\\\\\\\\\\\\\
 '\\\\
 '\\\\\\\\
 '\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Function

Function FindDLLFunc(pDLLName As String, pFuncName As String) As Long
Dim LoadAddr As Long
Dim ProcAddr As Long
Dim offset As Long
Dim RemoteAddr As Long

LoadAddr = LoadLibrary(pDLLName)
If LoadAddr = 0 Then End
ProcAddr = GetProcAddress(LoadAddr, pFuncName)
offset = ProcAddr - LoadAddr
FreeLibrary LoadAddr

RemoteAddr = FindModuleHandle(pDLLName)
Do While RemoteAddr = 0
    RemoteAddr = FindModuleHandle(pDLLName)
    DoEvents
Loop
FindDLLFunc = RemoteAddr + offset
End Function
Function HookRecvPackets3()

Dim CreateFileAADDR As Long
Dim WriteFileADDR As Long
Dim CloseHandleADDR As Long
Dim pBytesMSName() As Byte
Dim pBytes() As Byte
Dim pStr As String
Dim pStrKO_RCVFNC As String
Dim MSName As String

MSName = "\\.\mailslot\SuperPriest" & Hex(GetTickCount)
MSHandle = EstablishMailSlot(MSName)

CreateFileAADDR = FindDLLFunc("kernel32.dll", "CreateFileA")
WriteFileADDR = FindDLLFunc("kernel32.dll", "WriteFile")
CloseHandleADDR = FindDLLFunc("kernel32.dll", "CloseHandle")

KO_RCVFNC = VirtualAllocEx(KO_HANDLE, 0, 1024, MEM_COMMIT, PAGE_READWRITE)

'1
pBytesMSName = StrConv(MSName, vbFromUnicode)
WriteByteArray KO_RCVFNC + &H400, pBytesMSName, UBound(pBytesMSName) - LBound(pBytesMSName) + 1
'2
'pStr = AlignDWORD(&H76E58CA4)
pStr = AlignDWORD(CreateFileAADDR)
ConvHEX2ByteArray pStr, pBytes
WriteByteArray KO_RCVFNC + &H32A, pBytes, UBound(pBytes) - LBound(pBytes) + 1
'3
pStr = AlignDWORD(WriteFileADDR)
ConvHEX2ByteArray pStr, pBytes
WriteByteArray KO_RCVFNC + &H334, pBytes, UBound(pBytes) - LBound(pBytes) + 1
'4
pStr = AlignDWORD(CloseHandleADDR)
ConvHEX2ByteArray pStr, pBytes
WriteByteArray KO_RCVFNC + &H33E, pBytes, UBound(pBytes) - LBound(pBytes) + 1
'5
pStr = AlignDWORD(KO_RCVHKB)
ConvHEX2ByteArray pStr, pBytes
WriteByteArray KO_RCVFNC + &H208, pBytes, UBound(pBytes) - LBound(pBytes) + 1


pStr = AlignDWORD(KO_RECVHK)
ConvHEX2ByteArray pStr, pBytes
WriteByteArray KO_RCVFNC + &H212, pBytes, UBound(pBytes) - LBound(pBytes) + 1

pStr = AlignDWORD(KO_RCVFNC)
ConvHEX2ByteArray pStr, pBytes
WriteByteArray KO_RCVFNC + &H21C, pBytes, UBound(pBytes) - LBound(pBytes) + 1


                       
pStr = "52" + "890D" + AlignDWORD(KO_RCVFNC + &H320) + "8905" + AlignDWORD(KO_RCVFNC + &H3B6) + "8B4E04890d" + AlignDWORD(KO_RCVFNC + &H1F4) + "8B56088915" + AlignDWORD(KO_RCVFNC + &H1FE) + "81F9001000007D3E5068800000006A036A006A01680000004068" + AlignDWORD(KO_RCVFNC + &H400) + "FF15" + AlignDWORD(KO_RCVFNC + &H32A) + "83F8FF741D506A0054FF35" + AlignDWORD(KO_RCVFNC + &H1F4) + "ff35" + AlignDWORD(KO_RCVFNC + &H1FE) + "50ff15" + AlignDWORD(KO_RCVFNC + &H334) + "ff15" + AlignDWORD(KO_RCVFNC + &H33E) + "8b0d" + AlignDWORD(KO_RCVFNC + &H320) + "8b05" + AlignDWORD(KO_RCVFNC + &H3B6) + "5aff25" + AlignDWORD(KO_RCVFNC + &H208)
ConvHEX2ByteArray pStr, pBytes
WriteByteArray KO_RCVFNC, pBytes, UBound(pBytes) - LBound(pBytes) + 1


pStrKO_RCVFNC = AlignDWORD(KO_RCVFNC)
ConvHEX2ByteArray pStrKO_RCVFNC, pBytes
WriteByteArray KO_RECVHK, pBytes, UBound(pBytes) - LBound(pBytes) + 1

End Function


Sub HookRecvPackets2()

Dim CreateFileAADDR As Long
Dim WriteFileADDR As Long
Dim CloseHandleADDR As Long
Dim pBytesMSName() As Byte
Dim pBytes() As Byte
Dim pStr As String
Dim pStrKO_RCVFNC As String


CreateFileAADDR = FindDLLFunc("kernel32.dll", "CreateFileA")
WriteFileADDR = FindDLLFunc("kernel32.dll", "WriteFile")
CloseHandleADDR = FindDLLFunc("kernel32.dll", "CloseHandle")

KO_RCVFNC = VirtualAllocEx(KO_HANDLE, 0, 1024, MEM_COMMIT, PAGE_READWRITE)

'1
pBytesMSName = StrConv(MSName, vbFromUnicode)
WriteByteArray KO_RCVFNC + 1024, pBytesMSName, UBound(pBytesMSName) - LBound(pBytesMSName) + 1
'2
pStr = AlignDWORD(CreateFileAADDR)
ConvHEX2ByteArray pStr, pBytes
WriteByteArray KO_RCVFNC + 810, pBytes, UBound(pBytes) - LBound(pBytes) + 1
'3
pStr = AlignDWORD(WriteFileADDR)
ConvHEX2ByteArray pStr, pBytes
WriteByteArray KO_RCVFNC + 820, pBytes, UBound(pBytes) - LBound(pBytes) + 1
'4
pStr = AlignDWORD(CloseHandleADDR)
ConvHEX2ByteArray pStr, pBytes
WriteByteArray KO_RCVFNC + 830, pBytes, UBound(pBytes) - LBound(pBytes) + 1
'5
pStr = AlignDWORD(KO_RCVHKB)
ConvHEX2ByteArray pStr, pBytes
WriteByteArray KO_RCVFNC + 520, pBytes, UBound(pBytes) - LBound(pBytes) + 1

pStr = AlignDWORD(KO_RECVHK)
ConvHEX2ByteArray pStr, pBytes
WriteByteArray KO_RCVFNC + 530, pBytes, UBound(pBytes) - LBound(pBytes) + 1

pStr = AlignDWORD(KO_RCVFNC)
ConvHEX2ByteArray pStr, pBytes
WriteByteArray KO_RCVFNC + 540, pBytes, UBound(pBytes) - LBound(pBytes) + 1


                       
pStr = "50890D" + AlignDWORD(KO_RCVFNC + 800) + "8905" + AlignDWORD(KO_RCVFNC + 950) + "890D" + AlignDWORD(KO_RCVFNC + 500) + "8B57088915" + AlignDWORD(KO_RCVFNC + 510) + "68800000006A036A006A01680000004068" + AlignDWORD(KO_RCVFNC + 1024) + "FF15" + AlignDWORD(KO_RCVFNC + 810) + "83f8ff741d" + "506a0054ff35" + AlignDWORD(KO_RCVFNC + 500) + "ff35" + AlignDWORD(KO_RCVFNC + 510) + "50ff15" + AlignDWORD(KO_RCVFNC + 820) + "ff15" + AlignDWORD(KO_RCVFNC + 830) + "8b0d" + AlignDWORD(KO_RCVFNC + 900) + "8b05" + AlignDWORD(KO_RCVFNC + 950) + "8b57088a4402ff" + "ff25" + AlignDWORD(KO_RCVFNC + 520)

ConvHEX2ByteArray pStr, pBytes
WriteByteArray KO_RCVFNC, pBytes, UBound(pBytes) - LBound(pBytes) + 1

pStrKO_RCVFNC = "FF25" + AlignDWORD(KO_RCVFNC + 540) + "90"

ConvHEX2ByteArray pStrKO_RCVFNC, pBytes
WriteByteArray KO_RECVHK, pBytes, UBound(pBytes) - LBound(pBytes) + 1

End Sub


Sub HookRecvPackets()
Dim CreateFileAADDR As Long
Dim WriteFileADDR As Long
Dim CloseHandleADDR As Long
Dim pBytesMSName() As Byte
Dim pBytes() As Byte
Dim pStr As String
Dim pStrKO_RCVFNC As String


CreateFileAADDR = FindDLLFunc("kernel32.dll", "CreateFileA")
WriteFileADDR = FindDLLFunc("kernel32.dll", "WriteFile")
CloseHandleADDR = FindDLLFunc("kernel32.dll", "CloseHandle")



'Find to Addy for each funciton (prepare to write the next routine)
CreateFileAADDR = FindDLLFunc("kernel32.dll", "CreateFileA")
WriteFileADDR = FindDLLFunc("kernel32.dll", "WriteFile")
CloseHandleADDR = FindDLLFunc("kernel32.dll", "CloseHandle")

KO_RCVFNC = VirtualAllocEx(KO_HANDLE, 0, 1024, MEM_COMMIT, PAGE_READWRITE)


pBytesMSName = StrConv(MSName, vbFromUnicode)
WriteByteArray KO_RCVFNC + 1024, pBytesMSName, UBound(pBytesMSName) - LBound(pBytesMSName) + 1

pStr = AlignDWORD(CreateFileAADDR)
ConvHEX2ByteArray pStr, pBytes
WriteByteArray KO_RCVFNC + 810, pBytes, UBound(pBytes) - LBound(pBytes) + 1

pStr = AlignDWORD(WriteFileADDR)
ConvHEX2ByteArray pStr, pBytes
WriteByteArray KO_RCVFNC + 820, pBytes, UBound(pBytes) - LBound(pBytes) + 1

pStr = AlignDWORD(CloseHandleADDR)
ConvHEX2ByteArray pStr, pBytes
WriteByteArray KO_RCVFNC + 830, pBytes, UBound(pBytes) - LBound(pBytes) + 1

pStr = AlignDWORD(KO_RCVHKB)
ConvHEX2ByteArray pStr, pBytes
WriteByteArray KO_RCVFNC + 520, pBytes, UBound(pBytes) - LBound(pBytes) + 1

pStr = AlignDWORD(KO_RECVHK)
ConvHEX2ByteArray pStr, pBytes
WriteByteArray KO_RCVFNC + 530, pBytes, UBound(pBytes) - LBound(pBytes) + 1

pStr = AlignDWORD(KO_RCVFNC)
ConvHEX2ByteArray pStr, pBytes
WriteByteArray KO_RCVFNC + 540, pBytes, UBound(pBytes) - LBound(pBytes) + 1

                     
pStr = "50890D" + AlignDWORD(KO_RCVFNC + 800) + "8905" + AlignDWORD(KO_RCVFNC + 950) + "890D" + AlignDWORD(KO_RCVFNC + 500) + "8B57088915" + AlignDWORD(KO_RCVFNC + 510) + "68800000006A036A006A01680000004068" + AlignDWORD(KO_RCVFNC + 1024) + "FF15" + AlignDWORD(KO_RCVFNC + 810) + "83f8ff741d" + "506a0054ff35" + AlignDWORD(KO_RCVFNC + 500) + "ff35" + AlignDWORD(KO_RCVFNC + 510) + "50ff15" + AlignDWORD(KO_RCVFNC + 820) + "ff15" + AlignDWORD(KO_RCVFNC + 830) + "8b0d" + AlignDWORD(KO_RCVFNC + 900) + "8b05" + AlignDWORD(KO_RCVFNC + 950) + "8b57088a4402ff" + "ff25" + AlignDWORD(KO_RCVFNC + 520)
                                     
                    
ConvHEX2ByteArray pStr, pBytes
WriteByteArray KO_RCVFNC, pBytes, UBound(pBytes) - LBound(pBytes) + 1


pStrKO_RCVFNC = "FF25" + AlignDWORD(KO_RCVFNC + 540) + "90"

ConvHEX2ByteArray pStrKO_RCVFNC, pBytes
WriteByteArray KO_RECVHK, pBytes, UBound(pBytes) - LBound(pBytes) + 1

End Sub
Public Function HexVal(pStrHex As String) As Long
Dim TmpStr As String
Dim TmpHex As String
Dim hexcode As String
Dim I As Long
TmpStr = ""
For I = Len(pStrHex) To 1 Step -1
    TmpHex = Hex(Asc(Mid(pStrHex, I, 1)))
    If Len(TmpHex) = 1 Then TmpHex = "0" & TmpHex
    TmpStr = TmpStr & TmpHex
Next
hexcode = TmpStr

End Function

Sub Skill(skillid As String)
On Error Resume Next
Dim ID As String
Dim ID1 As String
Dim ID2 As String
Dim IDs As Long
Dim skillid1
Dim skillid2
Dim skillid3
Dim skillhex As String
Dim skillid4 As String

skillhex = Hex(skillid)
skillid1 = Mid$(skillhex, 4, 6)
skillid2 = Mid$(skillhex, 2, 2)
skillid3 = Mid$(skillhex, 1, 1)

IDs = ReadLong(KO_ADR_CHR + KO_OFF_ID)
ID = AlignDWORD(IDs)
ID1 = Strings.Mid(ID, 3, 2)
ID2 = Strings.Mid(ID, 5, 2)
Dim pBytes(1 To 18) As Byte
pBytes(1) = &H31
pBytes(2) = &H3
pBytes(3) = "&H" & skillid1
pBytes(4) = "&H" & skillid2
pBytes(5) = "&H" & skillid3
pBytes(6) = &H0
pBytes(7) = "&H" & ID1
pBytes(8) = "&H" & ID2
pBytes(9) = "&H" & ID1
pBytes(10) = "&H" & ID2
pBytes(11) = &H0
pBytes(12) = &H0
pBytes(13) = &H0
pBytes(14) = &H0
pBytes(15) = &H0
pBytes(16) = &H0
pBytes(17) = &H0
pBytes(18) = &H0
SendPackets pBytes
PotionTimer = Now
End Sub



Private Function ReadMessage(MailMessage As String, MessagesLeft As Long)
Dim lBytesRead As Long
Dim lNextMsgSize As Long
Dim lpBuffer     As String
ReadMessage = False
Call GetMailslotInfo(MSHandle, ByVal 0&, lNextMsgSize, MessagesLeft, ByVal 0&)
If MessagesLeft > 0 And lNextMsgSize <> MAILSLOT_NO_MESSAGE Then
    lBytesRead = 0
    lpBuffer = String$(lNextMsgSize, Chr$(0))
    Call ReadFile(MSHandle, ByVal lpBuffer, Len(lpBuffer), lBytesRead, ByVal 0&)
    If lBytesRead <> 0 Then
        MailMessage = Left(lpBuffer, lBytesRead)
        ReadMessage = True
        Call GetMailslotInfo(MSHandle, ByVal 0&, lNextMsgSize, MessagesLeft, ByVal 0&)
    End If
End If
End Function

Private Function CheckForMessages(MessageCount As Long)
Dim lMsgCount    As Long
Dim lNextMsgSize As Long
CheckForMessages = False
GetMailslotInfo MSHandle, ByVal 0&, lNextMsgSize, lMsgCount, ByVal 0&
MessageCount = lMsgCount
CheckForMessages = True
End Function
Public Function HexString(EvalString As String) As String
Dim intStrLen As Integer
Dim intLoop As Integer
Dim strHex As String

EvalString = Trim(EvalString)
intStrLen = Len(EvalString)
For intLoop = 1 To intStrLen
strHex = strHex & Hex(Asc(Mid(EvalString, intLoop, 1)))
Next
hexword = strHex
End Function


Public Function Hex2Val(pStrHex As String) As Long
Dim TmpStr As String
Dim TmpHex As String
Dim I As Long
TmpStr = ""
For I = Len(pStrHex) To 1 Step -1
    TmpHex = Hex(Asc(Mid(pStrHex, I, 1)))
    If Len(TmpHex) = 1 Then TmpHex = "0" & TmpHex
    TmpStr = TmpStr & TmpHex
Next
Hex2Val = CLng("&H" & TmpStr)
End Function
Public Function ReadLong1(Addr As Long) As Long '1 byte lýk deðer okur
    Dim Value As Long
    ReadProcessMem KO_HANDLE, Addr, Value, 1, 0&
    ReadLong1 = Value
End Function

'Public Sub MaxMana() 'The manasave
'Dim pPtr As Long
'pPtr = ReadLong(KO_CHRDMA)
'WriteLong pPtr + 2360, totalMP
'End Sub
'
Public Sub SetupKeyboard()
Keyboard = ReadLong(KeybPtr + 0)
Keyboard = Keyboard + 620
End Sub

Public Sub SendK(ByVal Key As Long)
WriteLong Keyboard + Key * 4, 1
Sleep (50)
WriteLong Keyboard + Key * 4, 0
End Sub

'Burda aslýnda KnightOnline ustemi deðilmi onu kontrol etmek lazým
'Eðer deðilse SendK yý kullansýn, Ama aktifse Knight o zaman Yolla Tuþ("Z") olabilir.
'ben size yolun nasýl gideceðini söylerim sizlerde nasýl yapýlýr onu araþtýrýp uygularsýnýz :D
Public Sub SelectMOB()
SendK ("DIK_Z")
End Sub
Function PM(ID As String)
Dim pStr As String
Dim pBytes() As Byte
pStr = "35010" & "7" & "00" & hexword
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes
End Function
Public Sub SendKeys(ByVal Key As Long)
WriteLong Keyboard + Key * 4, 1
Sleep (50)
WriteLong Keyboard + Key * 4, 0
End Sub
Sub DInputSendKeys(pKey As Long, Optional pTimeMS As Long = 50)
WriteByte pKey, 128
f_Sleep pTimeMS, True
WriteByte pKey, 0
End Sub

Public Function TersCevir(KCode As String)
Dim msgBuffer As String
msgBuffer = KCode
Dim I As Integer
Dim Sonuc As String
Dim TekCift As String
TekCift = Len(msgBuffer) / 2
Dim a As Integer
a = InStr(1, TekCift, ",")

If a = 1 Then
    Dim TekEkle As String
    TekEkle = Mid(msgBuffer, 1, 1)
    msgBuffer = Mid(msgBuffer, 2)
End If
For I = (Len(msgBuffer) - 1) To 1 Step -2
    Sonuc = Sonuc & Mid(msgBuffer, I, 2)
Next I
TersCevir = Sonuc & TekEkle
End Function

Sub DispatchMailSlot()
Dim MsgCount As Long
Dim rc As Long
Dim MessageBuffer As String
Dim pVal As Long
Dim fullcode
Dim code
Dim sKey
MsgCount = 1
Do While MsgCount <> 0
rc = CheckForMessages(MsgCount)
If CBool(rc) And MsgCount > 0 Then
If ReadMessage(MessageBuffer, MsgCount) Then
Call HexVal(MessageBuffer)
code = MessageBuffer
On Error Resume Next
Debug.Print Asc(Left(MessageBuffer, 1))
Select Case Asc(Left(MessageBuffer, 1))
Case 34 ' HOOK TARGET HP (22) / Kojd => Snoxd.net

End Select
End If
End If
Loop
End Sub


Function Notice(NoticeYazi As String)

Dim pStr As String

Dim pBytes() As Byte

HexString NoticeYazi

 

pStr = "1013FF00" & hexword

ConvHEX2ByteArray pStr, pBytes

SendPackets pBytes

End Function

Public Function HexToStr(ByVal Data As String) As String
Dim Buffer As String
Dim I As Integer
If Len(Data) Mod 2 <> 0 Then
HexToStr = vbNullString
Else
For I = 1 To Len(Data) - 1 Step 2
Buffer = Buffer & Chr("&H" & Mid(Data, I, 2))
Next I
HexToStr = Buffer
End If
End Function


