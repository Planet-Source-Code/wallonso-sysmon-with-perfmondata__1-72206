VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPMonCPU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ClassName As String = "CPULoad"

Private Const Err_Initialize As Long = vbObjectError + 8001
Private Const Err_UnableToStartPerfmon As Long = vbObjectError + 8002
Private Const Err_CPUIndexOOB As Long = vbObjectError + 8003
Private Const Err_CantFindProcessorPerfMon As Long = vbObjectError + 8004
Private Const Err_CantFindCPUUsagePerfMon As Long = vbObjectError + 8005
Private Const Err_UnableToReadPDB As Long = vbObjectError + 8006




Private Declare Sub Memcopy Lib "KERNEL32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Private Declare Function SystemTimeToFileTime Lib "KERNEL32" (lpSystemTime As SystemTime, lpFileTime As Currency) As Long
Private Declare Function GetVersionEx Lib "KERNEL32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long

Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long


Private Declare Sub GetSystemInfo Lib "KERNEL32" (lpSystemInfo As SYSTEM_INFO)



Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type

Private Const HKEY_LOCAL_MACHINE = &H80000002

Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const REG_DWORD = 4
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_MORE_DATA = 234

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128
End Type

Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32s = 0

Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_ALL = &H1F0000

Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or _
                                KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or _
                                KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or _
                                KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or _
                        KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

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

Private Const Processor_IDX_Str As String = "238"
Private Const Processor_IDX  As Long = 238
Private Const CPUUsageIDX As Long = 6


Private Const RAM_IDX_str = "4"
Private Const RAM_IDX = 4
Private Const RamRemain_IDX = 24


Private m_lProcessorsCount As Long
Private m_lBufferSizeCPU As Long
Private m_lBufferSizeRAM As Long
Private m_bIsWinNT As Boolean



Private PDB_CPU As PERF_DATA_BLOCK
Private POT_CPU As PERF_OBJECT_TYPE
Private PCD_CPU As PERF_COUNTER_DEFINITION
Private PID_CPU As PERF_INSTANCE_DEFINITION
Private PCB_CPU As PERF_COUNTER_BLOCK

'Private PDB_RAM As PERF_DATA_BLOCK
'Private POT_RAM As PERF_OBJECT_TYPE
'Private PCD_RAM As PERF_COUNTER_DEFINITION
'Private PID_RAM As PERF_INSTANCE_DEFINITION
'Private PCB_RAM As PERF_COUNTER_BLOCK
'
'Private VI As OSVERSIONINFO

Private SysTime As Currency
Private PrevSysTime As Currency
Private m_aCountersCPU() As Currency
Private m_aPrevCountersCPU() As Currency
Private m_aCountersRAM As Currency
Private m_aPrevCountersRAM As Currency

Private Const BYTEIncrement As Long = 4096

Private m_bConnected As Boolean
Private m_regHandleRemote As Long
Private m_sMachine As String

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


Private Sub Class_Initialize()


    m_lProcessorsCount = -1
    m_lBufferSizeCPU = BYTEIncrement
    m_lBufferSizeRAM = BYTEIncrement
End Sub

Private Sub Class_Terminate()
    '
End Sub

Public Function CollectCPUData() As Boolean
Dim lKeyRegistry As Long, lResReg As Long
Dim H As Long, r As Long
Dim aBuf() As Byte, lAllocSz As Long
Dim lSrc As Long, lDest As Long
Dim ptrPOT As Long, ptrPCB As Long
Dim i As Long, lCPU As Long
Dim ST As Currency
Dim sInstanceName As String
        If m_bLocalServer Then
            lResReg = RegConnectRegistry(mvar_RemoteServer, HKEY_PERFORMANCE_DATA, lKeyRegistry)
            If lResReg <> 0 Then
                CollectCPUData = -1
                Exit Function
            End If
        Else
            lKeyRegistry = HKEY_PERFORMANCE_DATA
        End If
    
        lAllocSz = m_lBufferSizeCPU
        ReDim aBuf(1 To lAllocSz) As Byte
        'While RegQueryValueEx(HKEY_PERFORMANCE_DATA, Processor_IDX_Str, 0, 0, aBuf(1), m_lBufferSizeCPU) = ERROR_MORE_DATA
        While RegQueryValueEx(lKeyRegistry, Processor_IDX_Str, 0, 0, aBuf(1), m_lBufferSizeCPU) = ERROR_MORE_DATA
            lAllocSz = lAllocSz + BYTEIncrement
            ReDim aBuf(1 To lAllocSz) As Byte
            m_lBufferSizeCPU = lAllocSz
        Wend

        lDest = VarPtr(PDB_CPU)
        lSrc = VarPtr(aBuf(1))
        Memcopy ByVal lDest, ByVal lSrc, LenB(PDB_CPU)
        m_lBufferSizeCPU = lAllocSz
        If PDB_CPU.Signature <> "PERF" Then
            Err.Raise Err_UnableToReadPDB, ClassName & ".CollectCPUData()", "Unable to read performance data"
        End If
        lDest = VarPtr(POT_CPU)
        lSrc = VarPtr(aBuf(1)) + PDB_CPU.HeaderLength
        For i = 1 To PDB_CPU.NumObjectTypes
            Memcopy ByVal lDest, ByVal lSrc, LenB(POT_CPU)
            ptrPOT = lSrc
            If POT_CPU.ObjectNameTitleIndex = Processor_IDX Then Exit For
            lSrc = lSrc + POT_CPU.TotalByteLength
        Next i
        
        
        If POT_CPU.ObjectNameTitleIndex <> Processor_IDX Then
            Err.Raise Err_CantFindProcessorPerfMon, ClassName & ".CollectData", "Unable to locate the 'Processor' performance object"
        End If
        If m_lProcessorsCount < 1 Then
            m_lProcessorsCount = GetCPUCount()
        End If
        lDest = VarPtr(PCD_CPU)
        lSrc = lSrc + POT_CPU.HeaderLength
        For i = 1 To POT_CPU.NumCounters
            Memcopy ByVal lDest, ByVal lSrc, LenB(PCD_CPU)
            If PCD_CPU.CounterNameTitleIndex = CPUUsageIDX Then Exit For
            lSrc = lSrc + PCD_CPU.ByteLength
        Next i
        If PCD_CPU.CounterNameTitleIndex <> CPUUsageIDX Then
            Err.Raise Err_CantFindCPUUsagePerfMon, ClassName & ".CollectData", "Unable to locate the '% of CPU usage' performance counter"
        End If
        
        
        lSrc = ptrPOT + POT_CPU.DefinitionLength
        For i = 1 To POT_CPU.NumInstances
            lDest = VarPtr(PID_CPU)
            Memcopy ByVal lDest, ByVal lSrc, LenB(PID_CPU)
            sInstanceName = Space(PID_CPU.NameLength - 2)
            Memcopy ByVal sInstanceName, ByVal lSrc + PID_CPU.NameOffset, PID_CPU.NameLength - 2
            
            sInstanceName = StrConv(sInstanceName, vbFromUnicode)

            lSrc = lSrc + PID_CPU.ByteLength
            lDest = VarPtr(PCB_CPU)
            Memcopy ByVal lDest, ByVal lSrc, LenB(PCB_CPU)
            ptrPCB = lSrc
            If IsNumeric(sInstanceName) Then
                lCPU = CLng(sInstanceName)
                m_aPrevCountersCPU(lCPU) = m_aCountersCPU(lCPU)
                Memcopy ByVal VarPtr(m_aCountersCPU(lCPU)), ByVal ptrPCB + PCD_CPU.CounterOffset, LenB(m_aCountersCPU(lCPU))
            End If

            lSrc = lSrc + PCB_CPU.ByteLength
        Next i
        PrevSysTime = SysTime
        SystemTimeToFileTime PDB_CPU.SystemTime, ST
        SysTime = ST
        
    RegCloseKey lKeyRegistry
End Function

Public Function GetPerfDataValue(sInstance As String, sValue As String) As Currency
Dim H As Long, r As Long
Dim aBuf() As Byte, lAllocSz As Long, lBufSize As Long
Dim lSrc As Long, lDest As Long
Dim ptrPOT As Long, ptrPCB As Long
Dim i As Long, j As Long, k As Long
Dim ST As Currency
Dim perfDataBlock As PERF_DATA_BLOCK
Dim perfObjectType As PERF_OBJECT_TYPE
Dim perfCtrDef As PERF_COUNTER_DEFINITION
Dim perfInstDef As PERF_INSTANCE_DEFINITION
Dim PerfCounterDef As PERF_COUNTER_DEFINITION

        lBufSize = BYTEIncrement
        lAllocSz = lBufSize
        ReDim aBuf(1 To lAllocSz) As Byte
        '//BaseObject
        While RegQueryValueEx(m_regHandleRemote, sInstance, 0, 0, aBuf(1), lBufSize) = ERROR_MORE_DATA
            lAllocSz = lAllocSz + BYTEIncrement
            ReDim aBuf(1 To lAllocSz) As Byte
            lBufSize = lAllocSz
        Wend
        
        lDest = VarPtr(perfDataBlock)
        lSrc = VarPtr(aBuf(1))
        Memcopy ByVal lDest, ByVal lSrc, LenB(perfDataBlock)
        'lBufSize = lAllocSz
        
        
        '// check for success and valid perf data block signature
        If perfDataBlock.Signature <> "PERF" Then
            Debug.Print "Cannot Get Datablock"
        End If
        
        lDest = VarPtr(perfObjectType)
        lSrc = VarPtr(aBuf(1)) + perfDataBlock.HeaderLength
        '//Verify Baseobject
        For i = 1 To perfDataBlock.NumObjectTypes
            Memcopy ByVal lDest, ByVal lSrc, LenB(perfObjectType)
            ptrPOT = lSrc
            If perfObjectType.ObjectNameTitleIndex = sInstance Then Exit For
            
            lSrc = lSrc + perfObjectType.TotalByteLength
        Next i
        
        
        If perfObjectType.ObjectNameTitleIndex <> sInstance Then
            'Err.Raise Err_CantFindProcessorPerfMon, ClassName & ".CollectData", "Unable to locate the 'RAM' performance object"
            Debug.Print "No result"
        End If
        '//Here we are, we found our Subcounter
        lSrc = lSrc + perfObjectType.HeaderLength
        lDest = VarPtr(PerfCounterDef)
        For i = 1 To perfObjectType.NumCounters
            
            Memcopy lDest, lSrc, LenB(PerfCounterDef)
            '//We found or Value of Interrest
            If PerfCounterDef.CounterNameTitleIndex = sValue Then
                Debug.Print
                Exit For
            Else
                lSrc = lSrc + PerfCounterDef.ByteLength
            End If
        Next
        
End Function


'  PERF_DATA_BLOCK
'    PERF_OBJECT_TYPE (1)
'      PERF_COUNTER_DEFINITION (1)
'      PERF_COUNTER_DEFINITION (2)
'      <...>
'      PERF_INSTANCE_DEFINITION (1)
'        PERF_COUNTER_BLOCK
'          counterData1
'          counterData2
'          counterData3
'          <...>
'      PERF_INSTANCE_DEFINITION (2)
'        PERF_COUNTER_BLOCK
'          counterData1
'          counterData2
'          counterData3
'          <...>
'    PERF_OBJECT_TYPE (2)
'    <...>


'Public Function GetCPUCountLocal() As Long
'Dim SI As SYSTEM_INFO
'
'    If m_lProcessorsCount < 1 Then
'        GetSystemInfo SI
'        GetCPUCountLocal = SI.dwNumberOrfProcessors
'        m_lProcessorsCount = SI.dwNumberOrfProcessors
'        'm_lProcessorsCount = 1
'        ReDim m_aPrevCountersCPU(0 To m_lProcessorsCount - 1) As Currency
'        ReDim m_aCountersCPU(0 To m_lProcessorsCount - 1) As Currency
'    Else
'        GetCPUCountLocal = m_lProcessorsCount
'    End If
'
'End Function



Public Function GetCPUCount() As Long
Dim res As Long, lResultkey As Long
Dim lReserve As Long, sClass As String, lClass As Long, ft As FILETIME, lName As Long, sName As String
Dim nCnt As Long, lOpenResult
Dim sKey As String
    If m_bLocalServer Then
        res = RegConnectRegistry(mvar_RemoteServer, HKEY_LOCAL_MACHINE, lResultkey)
        If res <> 0 Then
            GetCPUCount = -1
            Exit Function
        End If
    Else
        lResultkey = HKEY_LOCAL_MACHINE
    End If
    sKey = "HARDWARE\DESCRIPTION\System\CentralProcessor" & vbNullString
    res = RegOpenKey(lResultkey, sKey, lOpenResult)
    If res = 0 Then
        sClass = Space(255)
        sName = Space(255)
        lClass = 255
        lName = 255
        nCnt = 0
        Do While 1
            sClass = Space(255)
            sName = Space(255)
            lClass = 255
            lName = 255
            res = RegEnumKeyEx(lOpenResult, nCnt, sName, lName, lReserve, sClass, lClass, ft)
            If res <> 0 Then Exit Do
            nCnt = nCnt + 1
        Loop
        
    End If
    m_lProcessorsCount = nCnt
    ReDim m_aPrevCountersCPU(0 To m_lProcessorsCount - 1) As Currency
    ReDim m_aCountersCPU(0 To m_lProcessorsCount - 1) As Currency
    GetCPUCount = nCnt
    RegCloseKey lOpenResult
    RegCloseKey lResultkey
End Function

Public Function GetCPUUsage(Optional ByVal CPU_Index As Long = 1) As Long
Dim cpuIndex As Long
    cpuIndex = CPU_Index - 1

        If m_lProcessorsCount < 0 Then CollectCPUData

        If (cpuIndex >= m_lProcessorsCount) Or (cpuIndex < 0) Then
            Err.Raise Err_CPUIndexOOB, ClassName & ".GetCPUUsageLocal()", "CPU index out of bounds"
        End If

        If PrevSysTime = SysTime Then
            GetCPUUsage = 0
        Else
            GetCPUUsage = CLng(100 * (1 - (m_aCountersCPU(cpuIndex) - m_aPrevCountersCPU(cpuIndex)) / (SysTime - PrevSysTime)))
        End If


End Function
'Public Function GetCPUUsageRemote(Optional ByVal CPU_Index As Long = 1) As Long
'Dim cpuIndex As Long
'    cpuIndex = CPU_Index - 1
'
'        If m_lProcessorsCount < 0 Then CollectCPUData
'
'        If (cpuIndex >= m_lProcessorsCount) Or (cpuIndex < 0) Then
'            Err.Raise Err_CPUIndexOOB, ClassName & ".GetCPUUsageLocal()", "CPU index out of bounds"
'        End If
'
'        If PrevSysTime = SysTime Then
'            GetCPUUsageRemote = 0
'        Else
'            GetCPUUsageRemote = CLng(100 * (1 - (m_aCountersCPU(cpuIndex) - m_aPrevCountersCPU(cpuIndex)) / (SysTime - PrevSysTime)))
'        End If
'
'
'End Function


