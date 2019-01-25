option explicit

sub main() ' {

'  debug.print vba.int(addressOf main)

   dim o1a as obj_1
   dim o1b as obj_1
   dim o1c as obj_1

   dim o2a as obj_2
   dim o2b as obj_2
   dim o2c as obj_2

   debug.print "varPtr(o1a) = " &  varPtr(o1a)
   debug.print "o1b : " & (varPtr(o1a) - varPtr(o1b)) ' 4
   debug.print "o1c : " & (varPtr(o1b) - varPtr(o1c)) ' 4
   debug.print "o2a : " & (varPtr(o1c) - varPtr(o2a)) ' 4
   debug.print "o2b : " & (varPtr(o2a) - varPtr(o2b)) ' 4
   debug.print "o2c : " & (varPtr(o2b) - varPtr(o2c)) ' 4

   set o1a = new obj_1
   set o1b = new obj_1
   set o1c = new obj_1
   set o2a = new obj_2
   set o2b = new obj_2
   set o2c = new obj_2

   debug.print "varPtr(o1a) = " &  varPtr(o1a)
   debug.print "o1b : " & (varPtr(o1a) - varPtr(o1b)) ' 4
   debug.print "o1c : " & (varPtr(o1b) - varPtr(o1c)) ' 4
   debug.print "o2a : " & (varPtr(o1c) - varPtr(o2a)) ' 4
   debug.print "o2b : " & (varPtr(o2a) - varPtr(o2b)) ' 4
   debug.print "o2c : " & (varPtr(o2b) - varPtr(o2c)) ' 4


   debug.print "objPtr(o1a) = " &  objPtr(o1a)
   debug.print "o1b : " & (objPtr(o1a) - objPtr(o1b)) ' 4
   debug.print "o1c : " & (objPtr(o1b) - objPtr(o1c)) ' 4
   debug.print "o2a : " & (objPtr(o1c) - objPtr(o2a)) ' 4
   debug.print "o2b : " & (objPtr(o2a) - objPtr(o2b)) ' 4
   debug.print "o2c : " & (objPtr(o2b) - objPtr(o2c)) ' 4

'  debug.print "lenB: " & lenB(o2b)

'  debug.print "varPtr(o1a) = " & varPtr(o1a)
'  debug.print "varPtr(o1b) = " & varPtr(o1b)
'  debug.print "varPtr(o2a) = " & varPtr(o2a)
'  debug.print "varPtr(o2b) = " & varPtr(o2b)

'  debug.print "objPtr(o1a) = " & objPtr(o1a)
'  debug.print "objPtr(o1b) = " & objPtr(o1b)
'  debug.print "objPtr(o1c) = " & objPtr(o1c)
'  debug.print "objPtr(o2a) = " & objPtr(o2a)
'  debug.print "objPtr(o2b) = " & objPtr(o2b)
'  debug.print "objPtr(o2c) = " & objPtr(o2c)

   cells(1,1) = objPtr(o1a)
   cells(2,1) = objPtr(o1b)
   cells(3,1) = objPtr(o2a)
   cells(4,1) = objPtr(o2b)


end sub ' }
