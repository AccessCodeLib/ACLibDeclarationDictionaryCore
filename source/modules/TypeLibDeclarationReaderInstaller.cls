VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TypeLibDeclarationReaderInstaller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal Hwnd As LongPtr, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Private Const SW_HIDE = 0

Private Const ProgId As String = "ACLibDeclarationReader.TypeLibDeclarationReader"
Private Const ClassGuid As String = "{50065C31-0430-4AF3-9CBA-AC20B9357A9D}"

Public Sub ExportDllFileToFolder(ByVal TargetFolderPath As String)

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsAttachment As DAO.Recordset2
    Dim fldAttachment As DAO.Field2
    Dim FileName As String
    Dim FilePath As String

    If Not Right(TargetFolderPath, 1) = "\" Then
      TargetFolderPath = TargetFolderPath & "\"
    End If

    Set db = CodeDb
    Set rs = db.OpenRecordset("SELECT * FROM USysAppFiles WHERE FileName = 'AccessCodeLib.DeclarationReader.dll'", dbOpenDynaset)

    If Not rs.EOF Then
        Set rsAttachment = rs.Fields("File").Value
        Do While Not rsAttachment.EOF
            FileName = rsAttachment.Fields("FileName")
            FilePath = TargetFolderPath & FileName
            rsAttachment.Fields("FileData").SaveToFile FilePath
            rsAttachment.MoveNext
        Loop
        rsAttachment.Close
        Set rsAttachment = Nothing
    Else
        Err.Raise vbObjectError, "", "File data missing in USysAppFiles"
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing

End Sub


Public Sub RegisterNetComDll(ByVal DllPath As String)

   Const FullClassName As String = "AccessCodeLib.DeclarationReader.TypeLibDeclarationReader"
   Const AssemblyName As String = "AccessCodeLib.DeclarationReader"
   Const TargetRuntimeVersion As String = "v4.0.30319"
   Const AssemblyFile As String = AssemblyName & ".dll"

#If Win64 Then
   Const Bitness As Long = 64
#Else
   Const Bitness As Long = 32
#End If

   Dim AssemblyVersion As String
   AssemblyVersion = "1.0.0"

   If Right(DllPath, 1) <> "\" Then
      DllPath = DllPath & "\"
   End If

   AddRegItems FullClassName, ClassGuid, ProgId, _
                     AssemblyName, DllPath & AssemblyFile, AssemblyVersion, Bitness, _
                     TargetRuntimeVersion

End Sub

Private Sub AddRegItems(ByVal FullClassName As String, ByVal ClassGuid As String, ByVal ProgId As String, _
                        ByVal AssemblyName As String, ByVal AssemblyPath As String, ByVal AssemblyVersion As String, ByVal Bitness As Long, _
                        ByVal TargetRuntimeVersion As String)

   AddRegItem "HKCU\Software\Classes\" & ProgId, vbNullString, "REG_SZ", FullClassName
   AddRegItem "HKCU\Software\Classes\" & ProgId & "\CLSID", vbNullString, "REG_SZ", ClassGuid
   AddRegItem "HKCU\Software\Classes\CLSID\" & ClassGuid, vbNullString, "REG_SZ", FullClassName, Bitness
   AddRegItem "HKCU\Software\Classes\CLSID\" & ClassGuid & "\ProgId", vbNullString, "REG_SZ", ProgId, Bitness
   AddRegItem "HKCU\Software\Classes\CLSID\" & ClassGuid & "\InprocServer32", vbNullString, "REG_SZ", "mscoree.dll", Bitness
   AddRegItem "HKCU\Software\Classes\CLSID\" & ClassGuid & "\InprocServer32", "ThreadingModel", "REG_SZ", "Both", Bitness
   AddRegItem "HKCU\Software\Classes\CLSID\" & ClassGuid & "\InprocServer32", "Class", "REG_SZ", FullClassName, Bitness
   AddRegItem "HKCU\Software\Classes\CLSID\" & ClassGuid & "\InprocServer32", "Assembly", "REG_SZ", AssemblyName & ", Version=" & AssemblyVersion & ", Culture=neutral, PublicKeyToken=null", Bitness
   AddRegItem "HKCU\Software\Classes\CLSID\" & ClassGuid & "\InprocServer32", "RuntimeVersion", "REG_SZ", TargetRuntimeVersion, Bitness
   AddRegItem "HKCU\Software\Classes\CLSID\" & ClassGuid & "\InprocServer32", "CodeBase", "REG_SZ", "file:///" & Replace(AssemblyPath, "\", "/"), Bitness

End Sub

Private Sub AddRegItem(ByVal RegPath As String, _
                       ByVal RegValue As String, ByVal RegDataType As String, ByVal RegData As String, _
              Optional ByVal Bitness As Long = 0)

   Const ParamStringTemplate As String = "ADD ""{RegPath}"" /v ""{RegValue}"" /t {RegDataType} /d ""{RegData}"" /reg:{Bitness} /f"

   Dim ParamString As String

   ParamString = ParamStringTemplate
   ParamString = Replace(ParamString, "{RegPath}", RegPath)

   If Len(RegValue) = 0 Then
      ParamString = Replace(ParamString, "/v ""{RegValue}""", "/ve")
   Else
      ParamString = Replace(ParamString, "{RegValue}", RegValue)
   End If

   ParamString = Replace(ParamString, "{RegDataType}", RegDataType)
   ParamString = Replace(ParamString, "{RegData}", RegData)

   If Bitness = 0 Then
      ParamString = Replace(ParamString, "/reg:{Bitness} ", vbNullString)
   Else
      ParamString = Replace(ParamString, "{Bitness}", Bitness)
   End If

   RunRegExe ParamString

End Sub

Private Sub RunRegExe(ByVal Params As String)

   Dim RegExeFilePath As String
   RegExeFilePath = Environ("windir") & "\System32\REG.exe"
   ShellExecute Application.hWndAccessApp, "open", RegExeFilePath, Params, vbNullString, SW_HIDE

End Sub

Public Sub UnRegisterNetComDll()

#If Win64 Then
   Const Bitness As Long = 64
#Else
   Const Bitness As Long = 32
#End If

   UnRegisterComDll ClassGuid, ProgId, Bitness

End Sub

Private Sub UnRegisterComDll(ByVal ClassGuid As String, ByVal ProgId As String, Bitness As Long)

   DelRegItem "HKCU\Software\Classes\" & ProgId
   DelRegItem "HKCU\Software\Classes\CLSID\" & ClassGuid, Bitness

End Sub

Private Sub DelRegItem(ByVal RegPath As String, _
              Optional ByVal Bitness As Long = 0)

   Const ParamStringTemplate As String = "DELETE ""{RegPath}"" /reg:{Bitness} /f"

   Dim ParamString As String

   ParamString = ParamStringTemplate
   ParamString = Replace(ParamString, "{RegPath}", RegPath)

   If Bitness = 0 Then
      ParamString = Replace(ParamString, "/reg:{Bitness} ", vbNullString)
   Else
      ParamString = Replace(ParamString, "{Bitness}", Bitness)
   End If

   RunRegExe ParamString

End Sub
