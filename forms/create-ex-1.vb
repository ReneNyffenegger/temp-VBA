'
'  https://stackoverflow.com/a/11541912/180275
'
option explicit

sub main() ' {
 
 '
' We don't wont no flashing screent while creating the form:
'
  application.VBE.mainWindow.visible = false
  
  dim frmVbComp     as VBIDE.vbComponent
  dim frmDesign     as msForms.userForm

  set frmVbComp = thisWorkbook.vbProject.VBComponents.add(vbext_ct_MSForm)
  set frmDesign = frmVbComp.designer
  
  with frmVbComp
      .properties("caption") = "The dynamically created form"
      .properties("width"  ) =  400
      .properties("height" ) =  300
  end with

' dim p as variant
' for each p in frmVbComp.properties ' {
'     on error resume next
'     debug.print(p.name & ": " & typename(p.value))
'     on error goto 0
' next p ' }


  dim fram          as msForms.frame

  dim btn           as msForms.commandButton

 'dim NewComboBox as MSForms.ComboBox
  dim NewListBox as MSForms.ListBox
 'dim NewTextBox as MSForms.TextBox
 'dim NewLabel as MSForms.Label
 'dim NewOptionButton as MSForms.OptionButton
 'dim NewCheckBox as MSForms.CheckBox
  dim X    as integer
' dim line as integer
  
  
' Create ListBox

' set NewListBox = frmVbComp.Designer.Controls.add("Forms.listbox.1")
  set NewListBox = frmDesign.Controls.add("Forms.listbox.1")
  with newListBox
      .name          ="lst_1"
      .top           =  10
      .left          =  10
      .width         = 150
      .height        = 230
      .font.Size     =   8
      .font.Name     ="Tahoma"
'     .borderStyle   = fmBorderStyleOpaque
      .specialEffect = fmSpecialEffectSunken
  end with

  frmVbComp.activate
  newListBox.addItem "first"
  newListBox.addItem "second"
  newListBox.addItem "third"
  
  'Create CommandButton Create

  set btn = frmDesign.Controls.add("Forms.commandbutton.1")
  with btn
      .name         = "btn"
      .caption      = "clickMe"
      .accelerator  = "M"
      .top          =  10
      .left         = 200
      .width        =  66
      .height       =  20
      .font.size    =   8
    ' .font.name    ="Tahoma"
      .backStyle    = fmBackStyleOpaque
  end  with

  with frmVbComp.codeModule ' {

      .insertLines .countOfLines, "option explicit"

      .insertLines .countOfLines,  _
         "private sub userForm_initialize()"   & chr(10) & _
         "   me.lst_1.addItem ""Item one"""    & chr(10) & _
         "   me.lst_1.addItem ""Item two"""    & chr(10) & _
         "   me.lst_1.addItem ""Item three"""  & chr(10) & _
         "end sub"

      .insertLines .countOfLines,  _
         "private sub btn_click()"                       & chr(10) & _
         "   if me.lst_1.text <> """" then"              & chr(10) & _
         "      msgBox ""selected: "" &  me.lst_1.txt"   & chr(10) & _
         "   me.lst_1.addItem ""Item three"""            & chr(10) & _
         "end sub"

  end with ' }
  
 ' add code for listBox
' lstBoxData = "Data 1,Data 2,Data 3,Data 4"
' frmVbComp.codeModule.InsertLines  1, "private sub userForm_Initialize()"
' frmVbComp.codeModule.InsertLines  2, "   me.lst_1.addItem ""Item one"""
' frmVbComp.codeModule.InsertLines  3, "   me.lst_1.addItem ""Item two"""
' frmVbComp.codeModule.InsertLines  4, "   me.lst_1.addItem ""Item three"""
' frmVbComp.codeModule.InsertLines  5, "end sub"
' 
' 'add code for Comand Button
' frmVbComp.codeModule.InsertLines  6, "Private sub btn_Click()"
' frmVbComp.codeModule.InsertLines  7, "   If me.lst_1.text <>"""" Then"
' frmVbComp.codeModule.InsertLines  8, "      msgbox (""You selected item: "" & me.lst_1.text )"
' frmVbComp.codeModule.InsertLines  9, "   end If"
' frmVbComp.codeModule.InsertLines 10, "end sub"

 'Show the form
  VBA.UserForms.add(frmVbComp.Name).Show
  
  'Delete the form (Optional)
  'ThisWorkbook.VBProject.VBComponents.Remove frmVbComp
end sub ' }

