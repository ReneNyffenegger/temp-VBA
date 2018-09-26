
Option Explicit

Type GUID
   data1 As Long
   data2 As Integer
   data3 As Integer
   data4(7) As Byte
End Type

' function addRef(g as GUID, x as longPtr) as longPtr ' 
function addRef(this as long) as longPtr ' {

'   msgBox "addRef: data1 = " & hex(g.data1)
'   msgBox "addRef: data2 = " & hex(g.data2)
'   msgBox "addRef: data3 = " & hex(g.data3)
    debug.print "addRef: this = " & this

'
'   ..... And .... crash!
'

'   debug.print "addRef"
    addRef = 9

end function ' }

sub main()

    dim o1 as obj
    dim o2 as obj
    
    set o1 = New obj
    set o2 = New obj
  
    dim objPtr1 as long
    dim objPtr2 as long
    
    dim varPtr1 as long
    
    varPtr1 = VarPtr(o1)
    objPtr1 = ObjPtr(o1)
    objPtr2 = ObjPtr(o2)
    
    debug.print "varPtr1            = " & varPtr1
    debug.print "objPtr1            = " & objPtr1
    debug.print "objPtr2            = " & objPtr2

    debug.print "objPtr1(+   0)     = " & GetMem4_(objPtr1+0)
    debug.print "objPtr2(+   0)     = " & GetMem4_(objPtr2+0)

    debug.print "0:                 = " & GetMem4_(GetMem4_(objPtr2+0) + 0)
    debug.print "4:                 = " & GetMem4_(GetMem4_(objPtr2+0) + 4)
    debug.print "8:                 = " & GetMem4_(GetMem4_(objPtr2+0) + 8)

'   PutMem4 GetMem4_(objPtr2+0) + 0, vba.cLng(addressof addRef)
    PutMem4 GetMem4_(objPtr2+0) + 4, vba.cLng(addressof addRef)
    debug.print "0:                 = " & GetMem4_(GetMem4_(objPtr2+0) + 0)

    set o1 = o2


  '
  ' Approx. size of object:
  '
  ' debug.print "objPtr1 - objPtr2  = " &(objPtr1-objPtr2)

  '
  ' The varPtr points at the objPtr
  '
    debug.print "GetMem4(varPtr1) = " & GetMem4_(varPtr1)
    

end sub


