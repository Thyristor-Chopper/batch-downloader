Attribute VB_Name = "RunMemory"
'https://www.vbforums.com/showthread.php?652830-How-to-include-EXE-in-VB6
' jmsrickland
' Anything I post is an example only and is not intended to be the only solution, the total solution nor the final solution to your request nor do I claim that it is. If you find it useful, it is entirely up to you to make whatever changes necessary that you feel are adequate for your purposes.

Private Const SIZE_OF_80387_REGISTERS = 80

Private Type FLOATING_SAVE_AREA
    ControlWord As Long
    StatusWord As Long
    TagWord As Long
    ErrorOffset As Long
    ErrorSelector As Long
    DataOffset As Long
    DataSelector As Long
    RegisterArea(1 To SIZE_OF_80387_REGISTERS) As Byte
    Cr0NpxState As Long
End Type

Private Type CONTEXT86
    ContextFlags As Long
    'These are selected by CONTEXT_DEBUG_REGISTERS
    Dr0 As Long
    Dr1 As Long
    Dr2 As Long
    Dr3 As Long
    Dr6 As Long
    Dr7 As Long
    'These are selected by CONTEXT_FLOATING_POINT
    FloatSave As FLOATING_SAVE_AREA
    'These are selected by CONTEXT_SEGMENTS
    SegGs As Long
    SegFs As Long
    SegEs As Long
    SegDs As Long
    'These are selected by CONTEXT_INTEGER
    Edi As Long
    Esi As Long
    Ebx As Long
    Edx As Long
    Ecx As Long
    Eax As Long
    'These are selected by CONTEXT_CONTROL
    Ebp As Long
    Eip As Long
    SegCs As Long
    EFlags As Long
    Esp As Long
    SegSs As Long
End Type

Private Declare Function GetThreadContext Lib "kernel32" (ByVal hThread As Long, lpContext As CONTEXT86) As Long
Private Declare Function SetThreadContext Lib "kernel32" (ByVal hThread As Long, lpContext As CONTEXT86) As Long
Private Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long

'
' Process creation and memory access stuff
'
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpAppName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESSINFO) As Long
Private Declare Function NtUnmapViewOfSection Lib "ntdll.dll" (ByVal hProcess As Long, ByVal lpBaseAddress As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const CONTEXT_X86 = &H10000
Private Const CONTEXT86_CONTROL = (CONTEXT_X86 Or &H1)
Private Const CONTEXT86_INTEGER = (CONTEXT_X86 Or &H2)
Private Const CONTEXT86_SEGMENTS = (CONTEXT_X86 Or &H4)
Private Const CONTEXT86_FLOATING_POINT = (CONTEXT_X86 Or &H8)
Private Const CONTEXT86_DEBUG_REGISTERS = (CONTEXT_X86 Or &H10)
Private Const CONTEXT86_FULL = (CONTEXT86_CONTROL Or CONTEXT86_INTEGER Or CONTEXT86_SEGMENTS)

Private Const CREATE_SUSPENDED = &H4
Private Const MEM_COMMIT As Long = &H1000&
Private Const MEM_RESERVE As Long = &H2000&
Private Const PAGE_NOCACHE As Long = &H200
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Const PAGE_EXECUTE_WRITECOPY As Long = &H80
Private Const PAGE_EXECUTE_READ As Long = &H20
Private Const PAGE_EXECUTE As Long = &H10
Private Const PAGE_READONLY As Long = &H2
Private Const PAGE_WRITECOPY As Long = &H8
Private Const PAGE_NOACCESS As Long = &H1
Private Const PAGE_READWRITE As Long = &H4

'Private Const CREATE_UNICODE_ENVIRONMENT As Long = &H400&

'
' Main stuff for any API code
'
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal L As Long)

Private Enum ImageSignatureTypes
    IMAGE_DOS_SIGNATURE = &H5A4D     ''\\ MZ
    IMAGE_OS2_SIGNATURE = &H454E     ''\\ NE
    IMAGE_OS2_SIGNATURE_LE = &H454C  ''\\ LE
    IMAGE_VXD_SIGNATURE = &H454C     ''\\ LE
    IMAGE_NT_SIGNATURE = &H4550      ''\\ PE00
End Enum

Private Type IMAGE_DOS_HEADER
    e_magic As Integer        ' Magic number
    e_cblp As Integer         ' Bytes on last page of file
    e_cp As Integer           ' Pages in file
    e_crlc As Integer         ' Relocations
    e_cparhdr As Integer      ' Size of header in paragraphs
    e_minalloc As Integer     ' Minimum extra paragraphs needed
    e_maxalloc As Integer     ' Maximum extra paragraphs needed
    e_ss As Integer           ' Initial (relative) SS value
    e_sp As Integer           ' Initial SP value
    e_csum As Integer         ' Checksum
    e_ip As Integer           ' Initial IP value
    e_cs As Integer           ' Initial (relative) CS value
    e_lfarlc As Integer       ' File address of relocation table
    e_ovno As Integer         ' Overlay number
    e_res(0 To 3) As Integer  ' Reserved words
    e_oemid As Integer        ' OEM identifier (for e_oeminfo)
    e_oeminfo As Integer      ' OEM information; e_oemid specific
    e_res2(0 To 9) As Integer ' Reserved words
    e_lfanew As Long          ' File address of new exe header
End Type

'
' MSDOS File header
'
Private Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    characteristics As Integer
End Type

'
' Directory format.
'
Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long
    Size As Long
End Type

'
' Optional header format.
'
Private Const IMAGE_NUMBEROF_DIRECTORY_ENTRIES = 16

Private Type IMAGE_OPTIONAL_HEADER
    ' Standard fields.
    Magic As Integer
    MajorLinkerVersion As Byte
    MinorLinkerVersion As Byte
    SizeOfCode As Long
    SizeOfInitializedData As Long
    SizeOfUnitializedData As Long
    AddressOfEntryPoint As Long
    BaseOfCode As Long
    BaseOfData As Long
    ' NT additional fields.
    ImageBase As Long
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    W32VersionValue As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    Subsystem As Integer
    DllCharacteristics As Integer
    SizeOfStackReserve As Long
    SizeOfStackCommit As Long
    SizeOfHeapReserve As Long
    SizeOfHeapCommit As Long
    LoaderFlags As Long
    NumberOfRvaAndSizes As Long
    DataDirectory(0 To IMAGE_NUMBEROF_DIRECTORY_ENTRIES - 1) As IMAGE_DATA_DIRECTORY
End Type

Private Type IMAGE_NT_HEADERS
    Signature As Long
    FileHeader As IMAGE_FILE_HEADER
    OptionalHeader As IMAGE_OPTIONAL_HEADER
End Type

'
' Section header
'
Private Const IMAGE_SIZEOF_SHORT_NAME = 8

Private Type IMAGE_SECTION_HEADER
    SecName As String * IMAGE_SIZEOF_SHORT_NAME
    VirtualSize As Long
    VirtualAddress  As Long
    SizeOfRawData As Long
    PointerToRawData As Long
    PointerToRelocations As Long
    PointerToLinenumbers As Long
    NumberOfRelocations As Integer
    NumberOfLinenumbers As Integer
    characteristics  As Long
End Type

Private Const OFFSET_4 = 4294967296#

Function RunFromMemory(abExeFile() As Byte, si As STARTUPINFO, pi As PROCESSINFO, Optional ByVal Arguments As String, Optional InheritHandles As Long = 0&, Optional CreationFlags As Long = 0&, Optional EnvironmentVariables As Long = 0&, Optional CurrentDir As String = vbNullString) As Long
    Dim idh As IMAGE_DOS_HEADER
    Dim inh As IMAGE_NT_HEADERS
    Dim ish As IMAGE_SECTION_HEADER
    Dim context As CONTEXT86
    Dim ImageBase As Long, ret As Long, i As Long
    Dim Addr As Long, lOffset As Long

    CopyMemory idh, abExeFile(0), Len(idh)
    If idh.e_magic <> IMAGE_DOS_SIGNATURE Then
        RunFromMemory = 0&
        Exit Function
    End If

    CopyMemory inh, abExeFile(idh.e_lfanew), Len(inh)
    If inh.Signature <> IMAGE_NT_SIGNATURE Then
        RunFromMemory = 0&
        Exit Function
    End If

    If CreateProcess(vbNullString, "cmd.exe " & Arguments, 0, 0, InheritHandles, CREATE_SUSPENDED Or CreationFlags, EnvironmentVariables, CurrentDir, si, pi) = 0 Then
        RunFromMemory = 0&
        Exit Function
    End If

    context.ContextFlags = CONTEXT86_FULL
    If GetThreadContext(pi.hThread, context) = 0 Then GoTo ClearProcess

    ReadProcessMemory pi.hProcess, ByVal context.Ebx + 8, Addr, 4, 0
    If Addr = inh.OptionalHeader.ImageBase Then
        NtUnmapViewOfSection pi.hProcess, Addr
    End If

    ImageBase = VirtualAllocEx(pi.hProcess, inh.OptionalHeader.ImageBase, inh.OptionalHeader.SizeOfImage, MEM_RESERVE Or MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If ImageBase = 0 Then
        ImageBase = VirtualAllocEx(pi.hProcess, 0&, inh.OptionalHeader.SizeOfImage, MEM_RESERVE Or MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        If ImageBase = 0 Then GoTo ClearProcess
    End If

    WriteProcessMemory pi.hProcess, ByVal ImageBase, abExeFile(0), inh.OptionalHeader.SizeOfHeaders, 0&

    lOffset = idh.e_lfanew + Len(inh)
    For i = 0 To inh.FileHeader.NumberOfSections - 1
        CopyMemory ish, abExeFile(lOffset + i * Len(ish)), Len(ish)
        WriteProcessMemory pi.hProcess, ByVal ImageBase + ish.VirtualAddress, abExeFile(ish.PointerToRawData), ish.SizeOfRawData, 0&
    Next i

    context.Eax = ImageBase + inh.OptionalHeader.AddressOfEntryPoint
    WriteProcessMemory pi.hProcess, ByVal context.Ebx + 8, ImageBase, 4, 0&

    SetThreadContext pi.hThread, context
    ResumeThread pi.hThread

    RunFromMemory = 1&
    Exit Function

ClearProcess:
    TerminateProcess pi.hProcess, 0&
    CloseHandle pi.hThread
    CloseHandle pi.hProcess
    RunFromMemory = 0&
End Function

