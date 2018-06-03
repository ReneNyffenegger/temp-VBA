' Add
'  Microsoft Word 14.0 Object libary

sub main()

    dim wordApp as word.application
    dim wordDoc as word.document

    set wordApp  = new word.application

    wordApp.visible = true

    set wordDoc = wordApp.documents.add


    wordApp.selection.style = wordDoc.styles("Heading 1")
    wordApp.selection.typeText "This is a heading"
    wordApp.selection.typeParagraph

    wordApp.selection.typeText "And here is some text"
    wordApp.selection.typeParagraph
    wordApp.selection.typeText "And here is more text"
    wordApp.selection.typeParagraph

    wordApp.selection.style = wordDoc.styles("Heading 1")
    wordApp.selection.typeText "This is another heading"
    wordApp.selection.typeParagraph

    wordApp.selection.typeText "Text that belongs to another heading"
    wordApp.selection.typeParagraph
    wordApp.selection.typeText "Again text that belongs to another heading"
    wordApp.selection.typeParagraph


    wordDoc.saveAs2 environ("TEMP") & "\wordFromExcel.docx"
    wordDoc.close

end sub
