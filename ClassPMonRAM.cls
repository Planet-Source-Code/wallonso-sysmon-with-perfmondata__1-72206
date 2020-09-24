VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPMonRam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_MORE_DATA = 234

Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long


Private Declare Sub Memcopy Lib "KERNEL32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
'//*************************************************************
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SystemTime
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Type PERF_INSTANCE_DEFINITION
    ByteLength As Long
    ParentObjectTitleIndex As Long
    ParentObjectInstance As Long
    UniqueID As Long
    NameOffset As Long
    NameLength As Long
End Type

Private Type PERF_COUNTER_BLOCK
    ByteLength As Long
End Type

Private Type PERF_DATA_BLOCK
    Signature As String * 4
    LittleEndian As Long
    Version As Long
    Revision As Long
    TotalByteLength As Long
    HeaderLength As Long
    NumObjectTypes As Long
    DefaultObject As Long
    SystemTime As SystemTime
    PerfTime As LARGE_INTEGER
    PerfFreq As LARGE_INTEGER
    PerTime100nSec As LARGE_INTEGER
    SystemNameLength As Long
    SystemNameOffset As Long
End Type

Private Type PERF_OBJECT_TYPE
    TotalByteLength As Long
    DefinitionLength As Long
    HeaderLength As Long
    ObjectNameTitleIndex As Long
    ObjectNameTitle As Long
    ObjectHelpTitleIndex As Long
    ObjectHelpTitle As Long
    DetailLevel As Long
    NumCounters As Long
    DefaultCounter As Long
    NumInstances As Long
    CodePage As Long
    PerfTime As LARGE_INTEGER
    PerfFreq As LARGE_INTEGER
End Type

Private Type PERF_COUNTER_DEFINITION
    ByteLength As Long
    CounterNameTitleIndex As Long
    CounterNameTitle As Long
    CounterHelpTitleIndex As Long
    CounterHelpTitle As Long
    DefaultScale As Long
    DetailLevel As Long
    CounterType As Long
    CounterSize As Long
    CounterOffset As Long
End Type
'#define DEFAULT_BUFFER_SIZE 40960L
Private Const DEFAULT_BUFFER_SIZE = 4096
'//*************************************************************
Private mvar_RemoteServer As String
Private m_bLocalServer As Boolean
Public Property Let RemoteServer(ByVal NewData As String)
    If Len(NewData) Then
        If Left(NewData, 2) = "\\" Then
            mvar_RemoteServer = NewData
        Else
            mvar_RemoteServer = "\\" & NewData
        End If
        m_bLocalServer = True
    Else
        m_bLocalServer = False
    End If
End Property
'//*************************************************************


Public Function GetPerfMonValue(sInstanceValue As String, sDataValue As String)
Dim lKeyRegistry As Long, lResReg As Long, ret As Long
Dim lBufferSize As Long
Dim lAllocSz As Long
Dim perfDataBlock As PERF_DATA_BLOCK
Dim perfObjectType As PERF_OBJECT_TYPE
Dim perfCounterDefinition As PERF_COUNTER_DEFINITION
Dim perfInstanceDefinition As PERF_INSTANCE_DEFINITION
Dim perfCounterBlock As PERF_COUNTER_BLOCK
Dim lDest As Long, lSrc As Long
Dim i As Long
Dim ptrPOT As Long, ptrPCB As Long, lCPU As Long
Dim sInstanceName As String
Dim ValCounter As Long


    If m_bLocalServer Then
        lResReg = RegConnectRegistry(mvar_RemoteServer, HKEY_PERFORMANCE_DATA, lKeyRegistry)
        If lResReg <> 0 Then
            GetPerfMonValue = -1
        Exit Function
        End If
    Else
        lKeyRegistry = HKEY_PERFORMANCE_DATA
    End If
        
    lBufferSize = DEFAULT_BUFFER_SIZE
    lAllocSz = DEFAULT_BUFFER_SIZE
    
    
    ReDim aBuf(1 To lAllocSz) As Byte
    Do
        ret = RegQueryValueEx(lKeyRegistry, sInstanceValue, 0, 0, aBuf(1), lBufferSize)
        If ret = ERROR_MORE_DATA Then
        lAllocSz = lAllocSz + DEFAULT_BUFFER_SIZE
        ReDim aBuf(1 To lAllocSz) As Byte
        lBufferSize = lAllocSz
        ElseIf ret = 0 Then
            Exit Do
        Else
            Exit Function
        End If
    Loop While 1
    
    
    
    lDest = VarPtr(perfDataBlock)
    lSrc = VarPtr(aBuf(1))
    Memcopy ByVal lDest, ByVal lSrc, LenB(perfDataBlock)
    lBufferSize = lAllocSz
    If perfDataBlock.Signature <> "PERF" Then
        Debug.Print "No valid PerfData"
        
    End If
    lDest = VarPtr(perfObjectType)
    lSrc = VarPtr(aBuf(1)) + perfDataBlock.HeaderLength
    For i = 1 To perfDataBlock.NumObjectTypes
        Memcopy ByVal lDest, ByVal lSrc, LenB(perfObjectType)
        ptrPOT = lSrc
        If perfObjectType.ObjectNameTitleIndex = sInstanceValue Then Exit For
        lSrc = lSrc + perfObjectType.TotalByteLength
    Next i
    
    If perfObjectType.ObjectNameTitleIndex <> sInstanceValue Then
        'Err.Raise Err_CantFindProcessorPerfMon, ClassName & ".CollectData", "Unable to locate the 'Processor' performance object"
        Debug.Print "No Instance found"
    End If
        
    lDest = VarPtr(perfCounterDefinition)
    lSrc = lSrc + perfObjectType.HeaderLength
    For i = 1 To perfObjectType.NumCounters
        Memcopy ByVal lDest, ByVal lSrc, LenB(perfCounterDefinition)
        If perfCounterDefinition.CounterNameTitleIndex = sDataValue Then Exit For
        lSrc = lSrc + perfCounterDefinition.ByteLength
    Next i
    If perfCounterDefinition.CounterNameTitleIndex <> sDataValue Then
        Debug.Print "Have no Counter"
    End If
    lSrc = ptrPOT + perfObjectType.DefinitionLength
    If perfObjectType.NumInstances > 0 Then
        For i = 1 To perfObjectType.NumInstances
            lDest = VarPtr(perfInstanceDefinition)
            Memcopy ByVal lDest, ByVal lSrc, LenB(perfInstanceDefinition)
            sInstanceName = Space(perfInstanceDefinition.NameLength - 2)
            Memcopy ByVal sInstanceName, ByVal lSrc + perfInstanceDefinition.NameOffset, perfInstanceDefinition.NameLength - 2
            
            sInstanceName = StrConv(sInstanceName, vbFromUnicode)
    
            lSrc = lSrc + perfInstanceDefinition.ByteLength
            lDest = VarPtr(perfCounterBlock)
            Memcopy ByVal lDest, ByVal lSrc, LenB(perfCounterBlock)
            ptrPCB = lSrc
            If IsNumeric(sInstanceName) Then
                lCPU = CLng(sInstanceName)
                'm_aPrevCountersCPU(lCPU) = m_aCountersCPU(lCPU)
                Memcopy ByVal VarPtr(ValCounter), ByVal ptrPCB + perfCounterDefinition.CounterOffset, LenB(ValCounter)
            End If
    
            lSrc = lSrc + perfCounterBlock.ByteLength
        Next i
    Else
        '//So what now ?
            lDest = VarPtr(perfInstanceDefinition)
            ptrPCB = lSrc
            Memcopy ByVal VarPtr(ValCounter), ByVal ptrPCB + perfCounterDefinition.CounterOffset, LenB(ValCounter)
    End If
    RegCloseKey lKeyRegistry    '//Important to close the key after processing
    GetPerfMonValue = ValCounter
End Function

Public Function GetRamAmount() As Currency
Dim res As Long, lResultkey As Long
Dim lReserve As Long, sClass As String, lClass As Long, ft As FILETIME, lName As Long, sName As String
Dim nCnt As Long, lOpenResult
Dim sKey As String, sTemp As String
Dim xByte() As Byte
Dim isX64 As Boolean, xStep As Integer
Dim xSum As Double
'HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor
    If m_bLocalServer Then
        res = RegConnectRegistry(mvar_RemoteServer, HKEY_LOCAL_MACHINE, lResultkey)
        If res <> 0 Then
            GetRamAmount = -1
            Exit Function
        End If
    Else
        lResultkey = HKEY_LOCAL_MACHINE
    End If
    
    sKey = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
    sTemp = GetRegistryValueB(lResultkey, sKey, "PROCESSOR_ARCHITECTURE")
    
    isX64 = (sTemp = "AMD64")   '//Both intel an AMD set this Value for 64 bit architecture
    

    
    sKey = "HARDWARE\RESOURCEMAP\System Resources\Physical Memory" & vbNullString
    
    xByte = GetRegistryValueB(lResultkey, sKey, ".Translated")
    xStep = IIf(isX64, 20, 16)
    xSum = 0
    For nCnt = 32 To UBound(xByte) Step xStep
        If isX64 Then
            xSum = xSum + xByte(nCnt + 6) * 281474976710656#
            xSum = xSum + xByte(nCnt + 5) * 1099511627776#
            xSum = xSum + xByte(nCnt + 4) * 4294967296#
            xSum = xSum + xByte(nCnt + 3) * 16777216#
            xSum = xSum + xByte(nCnt + 2) * 65536#
            xSum = xSum + xByte(nCnt + 1) * 256#
            xSum = xSum + xByte(nCnt)
        Else
            xSum = xSum + xByte(nCnt + 3) * 16777216#
            xSum = xSum + xByte(nCnt + 2) * 65536#
            xSum = xSum + xByte(nCnt + 1) * 256#
            xSum = xSum + xByte(nCnt)
        End If

    Next
    RegCloseKey lResultkey
    GetRamAmount = xSum
    
End Function

Private Function GetRegistryValueB(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Variant
Dim lResult As Long
Dim lValueType As Long
Dim lBuf As Long
Dim lDataBufSize As Long
Dim r As Long
Dim Keyhand As Long
Dim xBuf() As Byte
Dim sResult As String
Dim lValue As Long

r = RegOpenKey(hKey, strPath, Keyhand)

 ' Get length/data type
lDataBufSize = 0
'//Default = DWORD
lResult = RegQueryValueEx(Keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)
If lResult = 234 Then
    ReDim xBuf(lDataBufSize)
    lResult = RegQueryValueEx(Keyhand, strValueName, 0&, lValueType, xBuf(0), lDataBufSize)
Else
    ReDim xBuf(lDataBufSize)
    lResult = RegQueryValueEx(Keyhand, strValueName, 0&, lValueType, xBuf(0), lDataBufSize)
    
End If

If lResult = ERROR_SUCCESS Then
    Select Case lValueType
        Case 1 '//String
            sResult = Left(StrConv(xBuf, vbUnicode), lDataBufSize - 1)
            GetRegistryValueB = sResult
        Case 8  '//Ressource
            GetRegistryValueB = xBuf

    End Select
End If

r = RegCloseKey(Keyhand)
    
End Function

