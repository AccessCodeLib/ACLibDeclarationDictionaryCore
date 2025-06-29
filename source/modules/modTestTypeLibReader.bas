Attribute VB_Name = "modTestTypeLibReader"
Option Compare Database
Option Explicit

Private Sub TestRegisteredComDll()

   Dim Dict As DeclarationDict
   Set Dict = New DeclarationDict

   With New TypeLibDeclarationReader
      .ImportVBProject VBE.ActiveVBProject, Dict
   End With

   Debug.Print Dict.Count

   Dict.ExportToFile CurrentProject.Path & "\dict.libs.txt"

End Sub

Private Sub UninstallTlbReaderComDll()

   With New TypeLibDeclarationReaderInstaller
      .UnRegisterNetComDll
   End With

End Sub

Private Sub TestNetComDomain()
' required: net 3.5 assembly
   Dim NCD As NetComDomain
   Dim obj As Object

   Set NCD = New NetComDomain

   Set obj = NCD.CreateObject("ACLibDeclarationReader.TypeLibDeclarationReader", "AccessCodeLib.DeclarationReader", CodeProject.Path & "\AccessCodeLib.DeclarationReader.dll")

   Dim Words As Variant

   Words = obj.ReadTypeLib("C:\Program Files\Common Files\Microsoft Shared\OFFICE16\ACEDAO.DLL")
   Stop

End Sub
