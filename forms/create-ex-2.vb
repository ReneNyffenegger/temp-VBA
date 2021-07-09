 
' thisWorkbook.VBProject.references.addFromGuid "{0002E157-0000-0000-C000-000000000046}",  5, 3 ' VBE
' thisWorkbook.VBProject.references.addFromGuid "{0D452EE1-E08F-101A-852E-02608C4D0BB4}",  2, 0 ' MS Forms
' 
 
option explicit

sub main()
    
    dim frmVbComp        as vbComponent
    dim frm              as msForms.userForm

dim oleObj as excel.oleObject

    dim NewOptionButton   As Msforms.OptionButton
    dim NewCommandButton1 As Msforms.CommandButton
    dim NewCommandButton2 As Msforms.CommandButton
    dim MyComboBox        As Msforms.ComboBox
    dim N, X As Integer, MaxWidth As Long

    
  ' First, check the form doesn't already exist

    for N = 1 To ActiveWorkbook.VBProject.VBComponents.Count
        If ActiveWorkbook.VBProject.VBComponents(N).Name = "tq84Form" Then
            ShowForm
            Exit Sub
        Else
        End If
    next N
    
 '  Make a userform
    set frmVbComp = activeWorkbook.VBProject.vbComponents.add(vbext_ct_MSForm)

    set frm =  frmVbComp.designer
    set oleObj = frm

    debug.print "typename(frmVbComp.designer  ) = " & typename(frmVbComp.designer  )
    debug.print "typename(frmVbComp.designerId) = " & typename(frmVbComp.designerId)


    with frmVbComp ' {
        .properties("Height") = 100
        .properties("Width" ) = 200
' tq84   on Error Resume Next
        .name = "tq84Form"
        .properties("Caption") = "Here is your user form"
    end with ' }
   
    dim p as vbide.property
    set p = frmVbComp.properties(1)

  ' dim i as long
  ' for i = 1 to frmVbComp.properties.count
  '     debug.print(frmVbComp.properties(i).name & ": " & frmVbComp.properties(i).value)
  ' next i
 
   
   'Debug.Print TypeName(frmVbComp.Designer)
    
  ' Add a Cancel button to the form
    Set NewCommandButton1 = frm.Controls.Add("forms.CommandButton.1")
    With NewCommandButton1
        .Caption  = "Cancel"
        .Height   = 18
        .Width    = 44
        .Left     = MaxWidth + 147
        .Top      = 6
    End With
    
     '//Add an OK button to the form
    Set NewCommandButton2 = frm.Controls.Add("forms.CommandButton.1")
    With NewCommandButton2
        .Caption = "OK"
        .Height  = 18
        .Width   = 44
        .Left    = MaxWidth + 147
        .Top     = 28
    End With
    
     '//Add code on the form for the CommandButtons
    With frmVbComp.CodeModule
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

    With myComboBox
        .Name   = "Combo1"
        .Left   = 10
        .Top    = 10
        .Height = 16
        .Width = 100
    End With
    
 '
 '  Call ShowForm which calls
 '  tq84Form.show
 '  because otherwise, tq84Form would not
 '  be defined.
 '
    ShowForm
'   tq84Form.show

end sub

sub showForm()
    tq84Form.Show
end sub

 
