VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VBA Export (re-import)"
   ClientHeight    =   3540
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton ImportCB 
      Caption         =   "Re-Import"
      Height          =   375
      Left            =   5640
      TabIndex        =   12
      ToolTipText     =   "Delete & re-import the project"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.PictureBox FloodPB 
      Height          =   330
      Left            =   105
      ScaleHeight     =   270
      ScaleWidth      =   6675
      TabIndex        =   11
      Top             =   2640
      Width           =   6735
   End
   Begin VB.CheckBox FormsCB 
      Caption         =   "Forms"
      Height          =   330
      Left            =   5775
      TabIndex        =   10
      Top             =   1155
      Width           =   1065
   End
   Begin VB.CheckBox CompOnlyCB 
      Caption         =   "Component Only"
      Height          =   435
      Left            =   5775
      TabIndex        =   9
      Top             =   690
      Width           =   1170
   End
   Begin VB.TextBox CompTB 
      Height          =   330
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   742
      Width           =   4530
   End
   Begin VB.TextBox ExportFolderTB 
      Height          =   855
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1260
      Width           =   4530
   End
   Begin VB.ComboBox ProjectCB 
      Height          =   315
      Left            =   1050
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   218
      Width           =   4530
   End
   Begin VB.CheckBox ProjectLockedCB 
      Caption         =   "Locked"
      Enabled         =   0   'False
      Height          =   330
      Left            =   5775
      TabIndex        =   4
      Top             =   210
      Width           =   960
   End
   Begin VB.CommandButton ExportButton 
      Caption         =   "Export"
      Height          =   375
      Left            =   5670
      TabIndex        =   0
      ToolTipText     =   "Export the project"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Done"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Component"
      Height          =   195
      Left            =   105
      TabIndex        =   7
      Top             =   810
      Width           =   810
   End
   Begin VB.Label Label2 
      Caption         =   "Export Folder"
      Height          =   510
      Left            =   105
      TabIndex        =   3
      Top             =   1260
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Project"
      Height          =   195
      Left            =   105
      TabIndex        =   2
      Top             =   285
      Width           =   495
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public VBInstance As VBIDE.VBE
Public Connect As Connect

Private myFloodPos As Long

Private mySelectedProject As VBProject

Public Sub showForm()

    Dim pVBAProject As VBProject
    
    CompOnlyCB.Value = vbUnchecked
    FormsCB.Value = vbChecked
    FormsCB.Enabled = True
    
    With VBInstance
    
        ProjectCB.Clear
        
        For Each pVBAProject In .VBProjects
        
            Call ProjectCB.AddItem(pVBAProject.Name)
            
        Next
        
        With .ActiveVBProject
    
            ProjectCB.Text = .Name
                    
            Call ProjectCB_Click
                    
            If Not VBInstance.SelectedVBComponent Is Nothing Then
                CompOnlyCB.Value = vbUnchecked
                CompOnlyCB.Enabled = True
                CompTB.Text = VBInstance.SelectedVBComponent.Name
            End If
    
        End With
        
    End With
    
    Call Me.Show
    
End Sub

Private Sub CancelButton_Click()

    Connect.Hide
    
End Sub

Private Sub CompOnlyCB_Click()

    If CompOnlyCB.Value = vbChecked Then
        FormsCB.Enabled = False
    Else
        FormsCB.Enabled = True
    End If
    
    Call FloodPB.Cls
    
End Sub

Private Sub Form_Load()

   'set the flood's initial attributes
   'white text (trust me, I know it says backcolor !)
    FloodPB.BackColor = &HFFFFFF
    FloodPB.ForeColor = &H7F0000
    FloodPB.DrawMode = 10
       
   'solid fill
    FloodPB.FillStyle = 0
    FloodPB.AutoRedraw = True 'required to prevent flicker!

    myFloodPos = 1

End Sub

Private Sub FormsCB_Click()

    Call FloodPB.Cls

End Sub

Private Sub ImportCB_Click()

    Dim oFSO As Scripting.FileSystemObject
    Dim pVBAProject As VBProject
    Dim vbComp As VBComponent  'VBA module, form, etc...
    Dim vbTempComp As VBComponent
    Dim strImportPath As String
    Dim oFile As Scripting.File
    Dim CodeMod As VBIDE.CodeModule
    
    Set oFSO = New Scripting.FileSystemObject
    
    Set pVBAProject = mySelectedProject ' VBInstance.ActiveVBProject
    
    strImportPath = ExportFolderTB.Text
    
    If oFSO.FolderExists(strImportPath) Then
         
        If MsgBox("Delete & Import Project from folder: " & strImportPath, vbYesNo, "Import") = vbYes Then
                    
            For Each vbComp In pVBAProject.VBComponents
            
                If vbComp.Type = vbext_ct_Document Then
                
                    Set CodeMod = vbComp.CodeModule
                    With CodeMod
                        .DeleteLines 1, .CountOfLines
                    End With
                    
                Else
                
                    Call pVBAProject.VBComponents.Remove(vbComp)
                
                End If
                
            Next
        
            For Each oFile In oFSO.GetFolder(strImportPath).Files
            
                Select Case oFSO.GetExtensionName(oFile.Name)
                
                    Case "cls"
                    
                        On Error Resume Next
                        
                        Set vbComp = pVBAProject.VBComponents(oFSO.GetBaseName(oFile.Name))
                        
                        On Error GoTo 0
                        
                        Set vbTempComp = pVBAProject.VBComponents.Import(oFile.Path)
                        
                        If Not vbComp Is Nothing Then
                        
                        If vbComp.Type = vbext_ct_Document Then
                        
                            Call vbComp.CodeModule.InsertLines(1, vbTempComp.CodeModule.Lines(1, vbTempComp.CodeModule.CountOfLines))
                            
                            Call pVBAProject.VBComponents.Remove(vbTempComp)
                            
                        End If
                        
                        End If
                        
                    Case "bas"
                    
                        Set vbTempComp = pVBAProject.VBComponents.Import(oFile.Path)
                    
                    Case "frm"
                        
                        Set vbTempComp = pVBAProject.VBComponents.Import(oFile.Path)
                        
                        If vbTempComp.CodeModule.CountOfLines > 0 Then
                            If Len(Trim$(vbTempComp.CodeModule.Lines(1, 1))) = 0 Then
                                Call vbTempComp.CodeModule.DeleteLines(1, 1)
                            End If
                        End If
                        
                End Select
                
            Next
            
        End If
        
    End If

End Sub

Private Sub ExportButton_Click()
    
    Dim oFSO As Scripting.FileSystemObject
    Dim pVBAProject As VBProject
    Dim vbComp As VBComponent  'VBA module, form, etc...
    Dim vbRef As Reference
    Dim strSavePath As String
    Dim i As Integer
    Dim ts As TextStream
    
    ' Get the VBA project
    ' If you want to export code for Normal instead, paste this macro into
    ' ThisDocument in the Normal VBA project and change the following line to:
  
    Set oFSO = New Scripting.FileSystemObject
    
    Set pVBAProject = mySelectedProject ' VBInstance.ActiveVBProject
  
    strSavePath = ExportFolderTB.Text
    
    If Not oFSO.FolderExists(strSavePath) Then
         If MsgBox("Create folder: " & strSavePath, vbYesNo, "Folder does not exist") = vbYes Then
            Call oFSO.CreateFolder(strSavePath)
        Else
            Exit Sub
        End If
    End If
    
    Call FloodPB.Cls
    
    If CompOnlyCB.Value Then
    
        FloodPB.Visible = True
        
        Set vbComp = VBInstance.SelectedVBComponent
        
        Select Case vbComp.Type

            Case vbext_ct_StdModule
                vbComp.Export strSavePath & "\" & vbComp.Name & ".bas"

            Case vbext_ct_Document, vbext_ct_ClassModule
                ' ThisDocument and class modules
                Call vbComp.Export(strSavePath & "\" & vbComp.Name & ".cls")

            Case vbext_ct_MSForm
                vbComp.Export strSavePath & "\" & vbComp.Name & ".frm"

            Case Else
                vbComp.Export strSavePath & "\" & vbComp.Name

        End Select
        
        Call floodUpdateTextPC(1, 1, "Export Complete")
                
    Else
    
        'Loop through all the components (modules, forms, etc) in the VBA project
          
        i = 0
        
        For Each vbComp In pVBAProject.VBComponents
      
           i = i + 1
            
           Select Case vbComp.Type
      
                Case vbext_ct_StdModule
                    vbComp.Export strSavePath & "\" & vbComp.Name & ".bas"
                    Call floodUpdateTextPC(pVBAProject.VBComponents.Count, i, vbComp.Name)
           
      
                Case vbext_ct_Document, vbext_ct_ClassModule
                    ' ThisDocument and class modules
                    Call vbComp.Export(strSavePath & "\" & vbComp.Name & ".cls")
                    Call floodUpdateTextPC(pVBAProject.VBComponents.Count, i, vbComp.Name)
           
      
                Case vbext_ct_MSForm
                
                    If FormsCB.Value = vbChecked Then
                        vbComp.Export strSavePath & "\" & vbComp.Name & ".frm"
                        Call floodUpdateTextPC(pVBAProject.VBComponents.Count, i, vbComp.Name)
           
                    End If
                    
                Case Else
                    vbComp.Export strSavePath & "\" & vbComp.Name
                    Call floodUpdateTextPC(pVBAProject.VBComponents.Count, i, vbComp.Name)
      
            End Select
            
        Next
  
        Set ts = oFSO.CreateTextFile(strSavePath & "\References.dat", True)
        
        For Each vbRef In pVBAProject.References
        
            With vbRef
            
                If Not .BuiltIn Then
                
                    On Error Resume Next
                    
                    Call ts.Write(.Name)
                    Call ts.Write("|")
                    Call ts.Write(.Description)
                    Call ts.Write("|")
                    Call ts.Write(.Guid)
                    Call ts.Write("|")
                    Call ts.Write(.Major)
                    Call ts.Write("|")
                    Call ts.Write(.Minor)
                    Call ts.Write("|")
                    Call ts.Write(.FullPath)
                    Call ts.WriteLine
                    
                    On Error GoTo 0
                    
                End If
            
            End With
            
        Next
        
        Call ts.Close
        
        Set ts = Nothing
        
        Call floodUpdateTextPC(pVBAProject.VBComponents.Count, i, "Export complete")
  
    End If
    
    Call CancelButton.SetFocus
    
End Sub

Private Sub ProjectCB_Click()

    Dim pVBAProject As VBProject
    Dim strToken() As String
    Dim oFSO As Scripting.FileSystemObject
    Dim oFolder As Scripting.Folder
    Dim oSubFolder As Scripting.Folder
    
    Set oFSO = New Scripting.FileSystemObject
    
    With VBInstance

        For Each pVBAProject In .VBProjects
    
            With pVBAProject
            
                If .Name = ProjectCB.Text Then
                
                    Set mySelectedProject = pVBAProject
                    
                    ProjectLockedCB.Value = IIf(.Protection = vbext_pp_none, vbUnchecked, vbChecked)
        
                    ExportButton.Enabled = .Protection = vbext_pp_none
                    ImportCB.Enabled = .Protection = vbext_pp_none
                    
                    strToken = Split(.FileName, "\")
                    
                    ExportFolderTB.Text = strToken(0) & "\" & strToken(1) & "\Sources\" & .Name
                    
                    Set oFolder = oFSO.GetFolder(strToken(0) & "\" & strToken(1))
                    
                    For Each oSubFolder In oFolder.SubFolders
                    
                        If InStr(1, oSubFolder.Name, "Sources") > 0 Then
                        
                            ExportFolderTB.Text = oSubFolder.Path & "\" & .Name
                        
                            Exit For
                        
                        End If
                
                    Next
                
                End If
                
            End With
            
        Next
        
        CompTB.Text = vbNullString
        CompOnlyCB.Value = vbUnchecked
        CompOnlyCB.Enabled = False
        
        FormsCB.Value = vbChecked
        FormsCB.Enabled = True
        
        Call FloodPB.Cls
        
    End With
    
End Sub

Public Sub floodUpdatePercent( _
    ByVal upperLimit As Long, _
    ByVal Progress As Long)

    Dim msg As String
    
   'make sure that the flood display hasn't already hit 100%
    If Progress <= upperLimit Then

     'error trap in case the code attempts
     'to set the scalewidth greater than
     'the max allowable
      If Progress > FloodPB.ScaleWidth Then
         Progress = FloodPB.ScaleWidth
      End If
            
     'erase the flood
      FloodPB.Cls
                  
     'set the ScaleWidth equal to the upper limit of the items to count
      FloodPB.ScaleWidth = upperLimit
      
     'format the progress into a percentage string to display
      msg = Format$(CLng((Progress / FloodPB.ScaleWidth) * 100)) + "%"
       
     'calculate the string's X & Y coordinates
     'in the PictureBox ... here, centered
      FloodPB.CurrentX = (FloodPB.ScaleWidth - FloodPB.TextWidth(msg)) \ 2
      FloodPB.CurrentY = (FloodPB.ScaleHeight - FloodPB.TextHeight(msg)) \ 2
         
     'print the percentage string in the text colour
      FloodPB.Print msg
        
     'print the flood bar to the new progress length in the line colour
      FloodPB.Line (0, 0)-(Progress, FloodPB.ScaleHeight), FloodPB.ForeColor, BF
       
     'allow the flood to complete drawing
      DoEvents
    
    End If

End Sub

Public Sub floodUpdateTextPC( _
    ByVal upperLimit As Long, _
    ByVal Progress As Long, _
    ByRef msg As String)

    Dim pc As String
    
    If Progress <= upperLimit Then

      If Progress > FloodPB.ScaleWidth Then
         Progress = FloodPB.ScaleWidth
      End If
           
      FloodPB.Cls
      FloodPB.ScaleWidth = upperLimit
          
     'format the progress into a percentage string to display
      pc = msg & " " & Format$(CLng((Progress / FloodPB.ScaleWidth) * 100)) + "%"
           
     'calculate the string's X & Y coordinates
     'in the PictureBox ... here, left justified and offset slightly
      FloodPB.CurrentX = 2
      FloodPB.CurrentY = (FloodPB.ScaleHeight - FloodPB.TextHeight(msg)) \ 2
           
     'calculate the string's X & Y coordinates
     'in the PictureBox based on the floodPos set
      Select Case myFloodPos
        Case 0  'left
                 FloodPB.CurrentX = 2
                 FloodPB.CurrentY = (FloodPB.ScaleHeight - FloodPB.TextHeight(pc)) \ 2
        
        Case 1  'centered
                 FloodPB.CurrentX = (FloodPB.ScaleWidth - FloodPB.TextWidth(pc)) \ 2
                 FloodPB.CurrentY = (FloodPB.ScaleHeight - FloodPB.TextHeight(pc)) \ 2
                  
        Case 2  'right
                 FloodPB.CurrentX = (FloodPB.ScaleWidth - FloodPB.TextWidth(pc)) - 3
                 FloodPB.CurrentY = (FloodPB.ScaleHeight - FloodPB.TextHeight(pc)) \ 2
      End Select
     
     'print the percentage string in the text colour
      FloodPB.Print pc
          
     'print the flood bar to the new progress length in the line colour
      FloodPB.Line (0, 0)-(Progress, FloodPB.ScaleHeight), FloodPB.ForeColor, BF
           
      DoEvents
    
    End If

End Sub

Public Sub floodUpdateText( _
    ByVal upperLimit As Long, _
    ByVal Progress As Long, _
    ByRef msg As String)

    If Progress <= upperLimit Then

      If Progress > FloodPB.ScaleWidth Then
         Progress = FloodPB.ScaleWidth
      End If
           
      FloodPB.Cls
      FloodPB.ScaleWidth = upperLimit
          
     'calculate the string's X & Y coordinates
     'in the PictureBox based on the floodPos set
      Select Case myFloodPos
        Case 0  'left
                 FloodPB.CurrentX = 2
                 FloodPB.CurrentY = (FloodPB.ScaleHeight - FloodPB.TextHeight(msg)) \ 2
        
        Case 1  'centered
                 FloodPB.CurrentX = (FloodPB.ScaleWidth - FloodPB.TextWidth(msg)) \ 2
                 FloodPB.CurrentY = (FloodPB.ScaleHeight - FloodPB.TextHeight(msg)) \ 2
                  
        Case 2  'right
                 FloodPB.CurrentX = (FloodPB.ScaleWidth - FloodPB.TextWidth(msg)) - 3
                 FloodPB.CurrentY = (FloodPB.ScaleHeight - FloodPB.TextHeight(msg)) \ 2
      End Select
           
     'print the string in the
     'at the position set above
      FloodPB.Print msg
          
     'print the flood bar to the new
     'progress length in the line colour
      FloodPB.Line (0, 0)-(Progress, FloodPB.ScaleHeight), FloodPB.ForeColor, BF
    
      DoEvents
    
    End If

End Sub

