'
'  vim: ft=vb
'
option explicit

declare sub ChangeVTable lib "C:\temp\temp-VBA\interna\object\c\vTable.dll" (byVal addrObj as any)

sub main() ' {

   dim o1 as new obj
   dim o2 as     obj

   if typeOf o1 is IUnknown then
      debug.print "Yes, iUnknown"
   end if

'  if typeOf o1 is IDispatch then
'     debug.print "Yes, IDispatch"
'  end if

   debug.print "objPtr(o1) = " & hex(objPtr(o1))

   ChangeVTable o1

   set o2 = o1

   debug.print o1.F_1

   set o1 = nothing
   set o2 = nothing

end sub ' }
