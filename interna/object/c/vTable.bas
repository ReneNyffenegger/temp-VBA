'
'  vim: ft=vb
'
option explicit

declare sub ChangeVTable lib "C:\temp\temp-VBA\interna\object\c\vTable.dll" (byVal addrObj as any)

sub main() ' {

   dim o1 as new obj
   dim o2 as     obj

   debug.print "objPtr(o1) = " & hex(objPtr(o1))

   ChangeVTable o1

   set o2 = o1

   set o1 = nothing
   set o2 = nothing

end sub ' }
