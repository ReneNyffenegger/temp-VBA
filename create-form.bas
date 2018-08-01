
'
'  https://stackoverflow.com/a/11541912/180275
'
sub CreateUserForm()

  dim myForm    as object ' msForms.userForm
  dim NewFrame  as MSForms.Frame

  dim NewButton as MSForms.CommandButton
 'dim NewComboBox as MSForms.ComboBox
  dim NewListBox as MSForms.ListBox
 'dim NewTextBox as MSForms.TextBox
 'dim NewLabel as MSForms.Label
 'dim NewOptionButton as MSForms.OptionButton
 'dim NewCheckBox as MSForms.CheckBox
  dim X    as integer
' dim line as integer
  
  'This is to stop screen flashing while creating form
  application.VBE.MainWindow.Visible = False
  
  set myForm = ThisWorkbook.VBProject.VBComponents.add(vbext_ct_MSForm)
  
  'Create the User Form
  with myForm
      .properties("Caption") = "New Form"
      .properties("Width") = 300
      .properties("Height") = 270
  end with
  
  'Create ListBox
  set NewListBox = myForm.Designer.Controls.add("Forms.listbox.1")
  with NewListBox
      .Name = "lst_1"
      .Top = 10
      .Left = 10
      .Width = 150
      .Height = 230
      .Font.Size = 8
      .Font.Name = "Tahoma"
      .BorderStyle = fmBorderStyleOpaque
      .SpecialEffect = fmSpecialEffectSunken
  end with
  
  'Create CommandButton Create
  set NewButton = myForm.Designer.Controls.add("Forms.commandbutton.1")
  with NewButton
      .Name = "cmd_1"
      .Caption = "clickMe"
      .Accelerator = "M"

      .Top       = 10
      .Left      = 200
      .Width     = 66
      .Height    = 20
      .Font.Size = 8
      .Font.Name = "Tahoma"
      .BackStyle = fmBackStyleOpaque
  end with
  
  'add code for listBox
  lstBoxData = "Data 1,Data 2,Data 3,Data 4"
  myForm.codeModule.InsertLines  1, "Private sub UserForm_Initialize()"
  myForm.codeModule.InsertLines  2, "   me.lst_1.addItem ""Data 1"" "
  myForm.codeModule.InsertLines  3, "   me.lst_1.addItem ""Data 2"" "
  myForm.codeModule.InsertLines  4, "   me.lst_1.addItem ""Data 3"" "
  myForm.codeModule.InsertLines  5, "end sub"
  
  'add code for Comand Button
  myForm.codeModule.InsertLines  6, "Private sub cmd_1_Click()"
  myForm.codeModule.InsertLines  7, "   If me.lst_1.text <>"""" Then"
  myForm.codeModule.InsertLines  8, "      msgbox (""You selected item: "" & me.lst_1.text )"
  myForm.codeModule.InsertLines  9, "   end If"
  myForm.codeModule.InsertLines 10, "end sub"
  'Show the form
  VBA.UserForms.add(myForm.Name).Show
  
  'Delete the form (Optional)
  'ThisWorkbook.VBProject.VBComponents.Remove myForm
end sub

