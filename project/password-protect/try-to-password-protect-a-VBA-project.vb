'
'   Trying to password protect a VBA project.
'   Does not work correctly: Excel sheet is not really saved (althouth somehow it seems to be)
'
option explicit

sub main()

   dim password as string : password = "foo"

   dim xls as workbook
   set xls = workbooks.add


  
   dim vbComp as vbComponent


   set application.vbe.activeVBProject = xls.vbProject


   Set vbComp = Application.VBE.ActiveVBProject.VBComponents.Add(1)
   vbComp.Name = "aModule"
   vbComp.CodeModule.InsertLines 1, "' comment"
    

   


    application.vbe.commandBars(1).findControl(1, 2578, "", true, true).execute ' 1 = msoControlButton, 2578: the ID of the controk 

  '
  ' With the 'Project Properties' message box now open, we use sendKeys to
  ' protect and enter the password
  '
    application.sendKeys "^{tab}"      ' ctrl-tab: go to 'Protection' tab
    application.sendKeys  "{tab}"      ' advance to 'Lock project for viewing' check box
    application.sendKeys  " "          ' check checkbox (code ASSUMES the checkbox is not checked)
    application.sendKeys "{tab}{tab}"  ' Advance to first password field (apparantly, two tabs are needed) ....
    application.sendKeys password      ' ... and enter password
    application.sendKeys "{tab}"       ' Advance to 'confirm password field ...
    application.sendKeys password      ' ... and re-enter password
    application.sendKeys "{tab}"
    application.sendKeys "{enter}"
'   application.sendKeys " "

'   application.wait(now + timeValue("0:00:01")) 

'   application.vbe.commandBars(1).findControl(1, 2578, "", true, true).execute ' 1 = msoControlButton, 2578: the ID of the controk 
'   application.sendKeys "^{tab}"      ' ctrl-tab: go to 'Protection' tab

  dim guid_internal_use_only as string
      guid_internal_use_only = "9108d454-5c13-4905-93be-12ec8059c842"

    xls.customDocumentProperties.add "MSIP_Label_" & guid_internal_use_only & "_Enabled", false, msoPropertyTypeString, "true"

    if dir(environ("temp") & "\foo.xlsm") <> "" then kill environ("temp")& "\foo.xlsm"

    xls.saveAs environ("temp") & "\foo.xlsm", 52
    debug.print "xls.name                 = " & xls.name
    debug.print "xls.vbProject.name       = " & xls.vbProject.name
    debug.print "xls.vbProject.protection = " & xls.vbProject.protection

    debug.print "saved = " & xls.vbProject.saved
    debug.print xls.vbProject.saved
    xls.save
    debug.print "xls.vbProject.saved      = " & xls.vbProject.saved
    debug.print "xls.name                 = " & xls.name
    debug.print "xls.vbProject.name       = " & xls.vbProject.name
    debug.print "xls.vbProject.protection = " & xls.vbProject.protection
'  xls.save
   xls.close true

'   xls.save
'   xls.save


end sub 
