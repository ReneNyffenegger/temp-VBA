 
' ThisWorkbook.VBProject.references.addFromGuid  GUID:="{0002E157-0000-0000-C000-000000000046}",  Major:=5, minor:=3
 
Option Explicit
Sub MakeUserForm()
    
    Dim MyUserForm As VBComponent
   Dim NewOptionButton As Msforms.OptionButton
    Dim NewCommandButton1 As Msforms.CommandButton
    Dim NewCommandButton2 As Msforms.CommandButton
    Dim MyComboBox As Msforms.ComboBox
    Dim N, X As Integer, MaxWidth As Long
    
     '//First, check the form doesn't already exist
    For N = 1 To ActiveWorkbook.VBProject.VBComponents.Count
        If ActiveWorkbook.VBProject.VBComponents(N).Name = "NewForm" Then
            ShowForm
            Exit Sub
        Else
        End If
    Next N
    
     '//Make a userform
    Set MyUserForm = ActiveWorkbook.VBProject _
    .VBComponents.Add(vbext_ct_MSForm)
    With MyUserForm
        .Properties("Height") = 100
        .Properties("Width") = 200
' tq84   On Error Resume Next
        .Name = "NewForm"
        .Properties("Caption") = "Here is your user form"
    End With
   
 
   
   'Debug.Print TypeName(MyUserForm.Designer)
    Dim frm As UserForm
    Set frm = MyUserForm.Designer
    
     '//Add a Cancel button to the form
    Set NewCommandButton1 = frm.Controls.Add("forms.CommandButton.1")
    With NewCommandButton1
        .Caption = "Cancel"
        .Height = 18
        .Width = 44
        .Left = MaxWidth + 147
        .Top = 6
    End With
    
     '//Add an OK button to the form
    Set NewCommandButton2 = frm.Controls.Add("forms.CommandButton.1")
    With NewCommandButton2
        .Caption = "OK"
        .Height = 18
        .Width = 44
        .Left = MaxWidth + 147
        .Top = 28
    End With
    
     '//Add code on the form for the CommandButtons
    With MyUserForm.CodeModule
        X = .CountOfLines
        .InsertLines X + 1, "Sub CommandButton1_Click()"
        .InsertLines X + 2, "    Unload Me"
        .InsertLines X + 3, "End Sub"
        .InsertLines X + 4, ""
        .InsertLines X + 5, "Sub CommandButton2_Click()"
        .InsertLines X + 6, "    Unload Me"
        .InsertLines X + 7, "End Sub"
    End With
    
     '//Add a combo box on the form
    Set MyComboBox = frm.Controls.Add("Forms.ComboBox.1")
    With MyComboBox
        .Name = "Combo1"
        .Left = 10
        .Top = 10
        .Height = 16
        .Width = 100
    End With
    
    ShowForm
End Sub
Sub ShowForm()
    NewForm.Show
End Sub

 
