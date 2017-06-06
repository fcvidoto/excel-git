Attribute VB_Name = "git"
Option Explicit

'Referencias
'Microsoft Visual Basic for Applications Extensibility 5.3

'-------------------------------------------------------
'No evento dos botoes da Ribbon
Public Sub exportarModulos_onAction()
  Call ExportSourceFiles
  MsgBox "Modulos Exportados", vbInformation
End Sub
Public Sub importarModulos_onAction()
  If Not CurrentProject.FullName Like "*\*" Then
    MsgBox "Operacao cancelada. Salve a planilha primeiro", vbInformation
    Exit Sub
  End If
  If MsgBox("Importar modulos?") = vbYes Then
    Call ImportaModulos
    MsgBox "Modulos Importados", vbInformation
  End If
End Sub
Public Sub excluirModulos_onAction()
  If MsgBox("Deseja excluir todos os modulos VBA? Operacao irreversivel", vbYesNo) = vbYes Then
    Call ExcluirModulos
    MsgBox "Modulos Excluidos", vbInformation
  End If
End Sub
'-------------------------------------------------------
'Exporta todos os modulos
Public Sub ExportSourceFiles()
  
  Dim destPath As String: destPath = CurrentProject.path
  Dim component As VBComponent
  
  For Each component In Application.VBE.ActiveVBProject.VBComponents
    If (component.Type = vbext_ct_ClassModule Or _
       component.Type = vbext_ct_StdModule Or _
       component.Type = vbext_ct_MSForm Or _
       component.Type = vbext_ct_Document) And _
       Not component.name = "git" Then
         component.Export destPath & "\" & component.name & ToFileExtension(component.Type)
    End If
  Next
End Sub
 
'Extencao do arquivo a ser exportado
Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String

  Select Case vbeComponentType
  Case vbext_ComponentType.vbext_ct_ClassModule
    ToFileExtension = ".cls"
  Case vbext_ComponentType.vbext_ct_StdModule
    ToFileExtension = ".bas"
  Case vbext_ComponentType.vbext_ct_MSForm
    ToFileExtension = ".frm"
  Case vbext_ComponentType.vbext_ct_ActiveXDesigner
  Case vbext_ComponentType.vbext_ct_Document
    ToFileExtension = ".cls"
  Case Else
    ToFileExtension = vbNullString
  End Select
 
End Function

'Importa todos os arquivos para o projeto
Public Sub ImportaModulos()

  Dim currentPath As String: currentPath = Left(CurrentProject.FullName, InStr(1, CurrentProject.FullName, CurrentProject.name) - 1)
  Dim fso As FileSystemObject: Set fso = New FileSystemObject
  Dim pasta As folder: Set pasta = fso.GetFolder(currentPath)
  Dim arquivo As File
  Dim modulo As VBComponent
  Dim nomeModulo As String
  Dim moduloEncontrado As Boolean
  
  'Itera por todos os modulos da pasta
  For Each arquivo In pasta.Files
  
    If Not arquivo.name Like "*.xls" And _
       Not arquivo.name Like "*.xlsm" And _
       Not arquivo.name Like "*.xlam" And _
       Not arquivo.name Like "*git.bas" And _
       (Not arquivo.name Like "*Sheet##" And Not arquivo.name Like "*Sheet#") And _
       Not arquivo.name Like "*ThisWorkbook*" And _
       Not arquivo.name Like "*.log" And _
       Not arquivo.name Like "*.frx" Then
       
        nomeModulo = Mid(arquivo.name, Len(CurrentProject.name) + 1, 300)
        nomeModulo = Mid(nomeModulo, 1, InStr(1, nomeModulo, ".") - 1)
        
        'Itera pelo modulos do projeto ** Exclui o modulo !
        For Each modulo In Application.VBE.ActiveVBProject.VBComponents
          If (nomeModulo = modulo.name) And _
              Not nomeModulo Like "*.frx" And Not nomeModulo Like "*.log" Then
            moduloEncontrado = True
            
            'Exclui o modulo
            Call Application.VBE.ActiveVBProject.VBComponents.Remove(modulo)
            
            'Importa o modulo
            Call Application.VBE.ActiveVBProject.VBComponents.Import(CurrentProject.path & "\" & arquivo.name)
          End If
        Next modulo
               
        'Importa o modulo novo ao projeto
        If moduloEncontrado = False And _
           Not nomeModulo Like "*.frx" And Not nomeModulo Like "*.log" Then
            'Importa o modulo
            Call Application.VBE.ActiveVBProject.VBComponents.Import(CurrentProject.path & "\" & arquivo.name)
        End If
        
        'Valor default da variavel
        moduloEncontrado = False
    End If
  Next arquivo

End Sub

'Exclui todos os modulos
Public Sub ExcluirModulos()

  Dim component As VBComponent
  
  For Each component In Application.VBE.ActiveVBProject.VBComponents
    If (component.Type = vbext_ct_ClassModule Or _
       component.Type = vbext_ct_StdModule Or _
       component.Type = vbext_ct_MSForm) And _
       Not component.name = "git" Then
        'component.Export destPath & component.Name & ToFileExtension(component.Type)
        Call Application.VBE.ActiveVBProject.VBComponents.Remove(component)
    End If
  Next
  
End Sub


