version 1.0 class
begin
  MultiUse            = -1  ' True
  Persistable         =  0  ' NotPersistable
  DataBindingBehavior =  0  ' vbNone
  DataSourceBehavior  =  0  ' vbNone
  MTSTransactionMode  =  0  ' NotAnMTSObject
end
attribute vb_Name            ="variantIterator"
attribute vb_GlobalNameSpace = false
attribute vb_Creatable       = true
attribute vb_PredeclaredId   = false
attribute vb_Exposed         = false

option explicit

private values_() as variant ' As Double

public function values(paramArray v()) As variantIterator

 ' If UBound(Values) < 0 Then Exit Function
 
 '  m_real = values
    
 '  Dim List()
 '  If UBound(Values) = 0 And IsArray(Values(0)) Then
 '      List = Values(0)
 '  Else
 '      List = Values
 '  End If
 '  
   'reDim m_Real(0 To uBound(List))
    values_ = v
 '   reDim m_Real(0 To uBound(values))
 '   dim i as long
 '   for i = 0 to uBound(m_Real)
 '       m_Real(i) = values(i)
 '   Next
    
    set values = me
    
end function

public function item(byVal i As Long) as variant ' {
       attribute item.vb_userMemId = 0
     '
     ' The attribute VB_UserMemId determines the method's DispatchID.
     ' The method whose DispatchId is 0 (zero) is the class's
     ' default member.
     
     
     
    if i < 0 or i > uBound(values_) then err.raise 9
    
    item = values_(i)
end function ' }

' Public Property Get item(byVal i As Long) As variant
'        attribute item.vb_userMemId = 0
' '
' ' The attribute VB_UserMemId determines the method's DispatchID.
' ' The method whose DispatchId is 0 (zero) is the class's
' ' default member.
'     if i < 0 or i > uBound(values_) then err.raise 9
'     
'     item = values_(i)
' end property

' Public Property Let Item(byVal i as long, byVal Value As variant)
'     If i < 0 Or i> UBound(values_) then Err.Raise 9
'     values_(i) = value
' End Property

' Public Property Get Length() As Long
'     Length = UBound(values_) + 1
' End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IEnumerable
public function GetEnumerator() as IEnumVARIANT ' {
       attribute       GetEnumerator.vb_userMemId   = -4
       
     ' attribute       GetEnumerator.vb_memberFlags = "40"
     '
     ' The attribute VB_UserMemId determines the method's DispatchID.
     ' A DispatchID with a negative number has a special meaning.
     ' In this case, -4 corresponds to DISPID_NEWENUM
     ' See also http://www.papwalker.com/ref101/ccol.html
     '

     '
     ' The ProcedureID must be set to -4 (DISPID_NEWENUM)
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
   ' set GetEnumerator = NewEnumerator(me, Length - 1, 0)
    set GetEnumerator = get_IEnumVARIANT_vTbl_etc(me, uBound(values_), 0)
end function ' }

' vim: ft=vb
