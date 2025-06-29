VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TypeLibDeclarationReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Private Const m_ComDllFileName As String = "AccessCodeLib.DeclarationReader.dll"
Private Const ProgId As String = "ACLibDeclarationReader.TypeLibDeclarationReader"

Private m_DeclDict As DeclarationDict
Private m_TypeLibReader As Object

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub ImportVBProject(ByVal VBProject2Import As VBProject, ByVal DeclDict As DeclarationDict)

   Dim ref As Object

   For Each ref In VBProject2Import.References
      ImportTypeLib ref.FullPath, DeclDict
   Next

End Sub

Public Sub ImportTypeLib(ByVal TypeLibFile As String, ByVal DeclDict As DeclarationDict)

   Dim Words() As String
   Dim i As Long

   Set m_DeclDict = DeclDict

   If m_TypeLibReader Is Nothing Then
      Set m_TypeLibReader = GetNewTypeLibDeclarationReaderInstance()
   End If

   Words = m_TypeLibReader.ReadTypeLib(TypeLibFile)

   For i = LBound(Words) To UBound(Words)
      m_DeclDict.AddWord Words(i)
   Next

End Sub

Private Function GetNewTypeLibDeclarationReaderInstance() As Object

   Dim ComDllRegistered As Boolean

On Error GoTo HandleCreateObjectErr
   Set GetNewTypeLibDeclarationReaderInstance = CreateObject(ProgId)

ExitHere:
   Exit Function

HandleCreateObjectErr:
   If Err.Number = 429 And Not ComDllRegistered Then ' Err 429 = ActiveX component can't create object
      ' try register com dll
      RegisterNetComDll
      ComDllRegistered = True
      Resume
   End If

   Err.Raise Err.Number, Err.Source, Err.Description
   Resume ExitHere

End Function

Private Sub RegisterNetComDll()

   With New TypeLibDeclarationReaderInstaller
      If Len(VBA.Dir(ComDllFilePath)) = 0 Then
         .ExportDllFileToFolder ComDllFolderPath
      End If
      Sleep 250
      .RegisterNetComDll ComDllFolderPath
      Sleep 250
   End With

End Sub

Private Function ComDllFilePath() As String
   ComDllFilePath = ComDllFolderPath & "\" & m_ComDllFileName
End Function

Private Function ComDllFolderPath() As String
   ComDllFolderPath = CodeProject.Path
End Function
