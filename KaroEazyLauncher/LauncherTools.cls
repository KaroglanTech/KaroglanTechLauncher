VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LauncherTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''��ģ����Ϣ'''''''''''''''''''''''
'����:JayshonYves
'����:LauncherTools
'����:
'       ActiveMovie control type library
'����ǰ��:
'       mmd.bas
'����:
'       FindJava() Ѱ��Java·�� �ҵ�����·�� δ�ҵ����ؿ�
'       PlayMusic(String)
'           ���������ļ�λ�� ���سɹ�����1 ���򷵻�0
'           �������0ʹ��PlayMusicOC�������
'       PlayMusicOC(Boolean)
'           �������ֿ��� ������ʹ��playMusic�����ļ�
'       RunMem() ���� ����ڴ� ��λMB
'       bakpic() ʹ�ñ���ͼƬģʽ t����ģʽ
'       bakmusic()������ű�������
'       bakpicReprint() ���»�������
'       CheckExeIsRun(String) �������Ƿ����
'       GetVerList(Obj,String) ��ȡ�汾�б���combobox/Listbox �������д��ַ��Ĭ��.minecraft\versions
'       OpenURL(Form,String) ʹ��Ĭ��������򿪵�ַ
'''''''''''''''''''''''''''''''''''''''''''''''''''
'������Ϣ:

Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type
Private Type MEMORYSTATUSEX
    dwLength As Long
    dwMemoryLoad As Long
    ullTotalPhys As LARGE_INTEGER
    ullAvailPhys As LARGE_INTEGER
    ullTotalPageFile As LARGE_INTEGER
    ullAvailPageFile As LARGE_INTEGER
    ullTotalVirtual As LARGE_INTEGER
    ullAvailVirtual As LARGE_INTEGER
    ullAvailExtendedVirtual As LARGE_INTEGER
End Type
Private Declare Function GlobalMemoryStatusEx Lib "kernel32.dll" (ByRef lpBuffer As MEMORYSTATUSEX) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function Findjava() As String
    On Error GoTo err
    Dim checkjava As Object
    Set checkjava = CreateObject("wscript.shell")
    Dim a As String
    a = checkjava.regread("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\javaws.exe\Path")
    Findjava = a + "\javaw.exe"
    Exit Function
err:
    Findjava = ""
End Function

Public Function CheckExeIsRun(exeName As String) As Boolean
    On Error GoTo err
    Dim WMI
    Dim Obj
    Dim Objs
    CheckExeIsRun = False
    Set WMI = GetObject("WinMgmts:")
    Set Objs = WMI.InstancesOf("Win32_Process")
    For Each Obj In Objs
        If (InStr(UCase(exeName), UCase(Obj.Description)) <> 0) Then
            CheckExeIsRun = True
            If Not Objs Is Nothing Then Set Objs = Nothing
            If Not WMI Is Nothing Then Set WMI = Nothing
            Exit Function
        End If
    Next
    If Not Objs Is Nothing Then Set Objs = Nothing
    If Not WMI Is Nothing Then Set WMI = Nothing
    Exit Function
err:
    If Not Objs Is Nothing Then Set Objs = Nothing
    If Not WMI Is Nothing Then Set WMI = Nothing
End Function

Public Function RunMem() As Integer
    If Dir(Environ("Windir") & "\SysWOW64", vbDirectory) <> "" Then
        RunMem = DisplayMemory / 3 * 2
    Else
        RunMem = 1024
    End If
End Function
Public Sub OpenURL(frm As Object, oURL As String)
    ShellExecute frm.hwnd, "Open", oURL, 0, 0, 0
End Sub

'''''''''''''''''''''''''''''''''�ڴ�'''''''''''''''''''
Private Function DisplayMemory() As Integer
    Dim MemStat As MEMORYSTATUSEX
    
    Dim TotalPhys As Currency
    Dim AvailPhys As Currency
    
    Dim POPSize As Currency
    
    Dim TotalPageFile As Currency
    Dim AvailPageFile As Currency
    
    Dim TotalVirtual As Currency
    Dim AvailVirtual As Currency
    With MemStat
        MemStat.dwLength = Len(MemStat)
        GlobalMemoryStatusEx MemStat
        POPSize = .dwLength
        
        TotalPhys = Int(LargeIntToCurrency(.ullTotalPhys) / 1024 ^ 2)
        AvailPhys = Int(LargeIntToCurrency(.ullAvailPhys) / 1024 ^ 2)
        
        TotalPageFile = Int(LargeIntToCurrency(.ullTotalPageFile) / 1024 ^ 2)
        AvailPageFile = Int(LargeIntToCurrency(.ullAvailPageFile) / 1024 ^ 2)
        
        TotalVirtual = Int(LargeIntToCurrency(.ullTotalVirtual) / 1024 ^ 2)
        AvailVirtual = Int(LargeIntToCurrency(.ullAvailVirtual) / 1024 ^ 2)
        
        'Label1.Caption = "��ջ��С:" & POPSize & " Bytes" & vbCrLf & _
"�����ڴ��С:" & TotalPhys & "MB" & vbCrLf & _
"���������ڴ�:" & TotalPhys - AvailPhys & "MB" & vbCrLf & _
"���������ڴ�:" & Int((AvailPhys / TotalPhys) * 100) & "%" & vbCrLf & _
"���������ڴ�:" & AvailPhys & "MB" & vbCrLf & vbCrLf & _
"�����ڴ��С:" & TotalVirtual & "MB" & vbCrLf & _
"���������ڴ�:" & TotalVirtual - AvailVirtual & "MB" & vbCrLf & _
"���������ڴ�:" & Int((AvailVirtual / TotalVirtual) * 100) & "%" & vbCrLf & _
"���������ڴ�:" & AvailVirtual & "MB" & vbCrLf & vbCrLf & _
        " �ܼ�:" & TotalPageFile & "MB" & vbCrLf & _
        " ����:" & TotalPageFile - AvailPageFile & "MB" & vbCrLf & _
        " ����:" & AvailPageFile & "MB" & vbCrLf & vbCrLf & _
        "-------------------------" & vbCrLf & _
"�Ѿ�ʹ����: " & .dwMemoryLoad & "%"
    End With
    DisplayMemory = AvailPhys
End Function

Private Function LargeIntToCurrency(liInput As LARGE_INTEGER) As Currency
    CopyMemory LargeIntToCurrency, liInput, LenB(liInput)
    LargeIntToCurrency = LargeIntToCurrency * 10000
End Function
