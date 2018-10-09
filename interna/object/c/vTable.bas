'
'  vim: ft=vb
'
option explicit

declare sub ChangeVTable lib "C:\temp\temp-VBA\interna\object\c\vTable.dll" (byVal addrObj as any)
declare sub dbg          lib "C:\temp\temp-VBA\interna\object\c\vTable.dll" (byVal txt as string       )

sub main() ' {

   dim o1 as new obj
   dim o2 as     obj

   if typeOf o1 is IUnknown then
      debug.print "Yes, iUnknown"
   end if

'  if typeOf o1 is IDispatch then
'     debug.print "Yes, IDispatch"
'  end if

   dbg "objPtr(o1) = " & hex(objPtr(o1)) & ", going to call ChangeVTable"

   ChangeVTable o1

   dbg "set o2 = o1"
   set o2 = o1

   dbg "Calling o1.F_1"
   debug.print o1.F_1

   dbg "set o1 = nothing"
   set o1 = nothing

   dbg "set o2 = nothing"
   set o2 = nothing

end sub ' }
