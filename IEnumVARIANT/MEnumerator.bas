attribute vb_name = "MEnumerator"
'
' Using some ideas of Dexter Freivald which were really helpful.
'
' MEnumerator.bas
'
' Implementation of IEnumVARIANT to support For Each in VB6
'
option explicit

type IUnknown_vtbl_T
     QueryInterface  as longPtr
     AddRef          as longPtr
     Release         as longPtr
end  type

type GUID__
     b_00 as byte
     b_01 as byte
     b_02 as byte
     b_03 as byte
     b_04 as byte
     b_05 as byte
     b_06 as byte
     b_07 as byte
     b_08 as byte
     b_09 as byte
     b_0a as byte
     b_0b as byte
     b_0c as byte
     b_0d as byte
     b_0e as byte
     b_0f as byte
end type

type IEnumVARIANT_vtbl_T
   '
   '  IEnumVARIANT inherits from IUnknown, thus the first element
   '  of IEnumVARIANT_vtbl_T needs to be an IUnknown_vtbl_T:
   '
      IUnknown_vtbl  as IUnknown_vtbl_T
   '
   '  The four function pointers to the specific methods of IEnumVARIANT:
   '
      Next           as longPtr
      Skip           as longPtr
      Reset          as longPtr
      Clone          as longPtr
end  type



' private vTbl(6) as long
' private vTbl(6) as longPtr
  
private IEnumVariant_vtbl as IEnumVariant_vtbl_T

private type TENUMERATOR
  ' VTablePtr          As Long
    VTablePtr          as longPtr
 '  VTablePtr_2        as longPtr
 '  vtbl_IEnumVariant  as IEnumVARIANT_vtbl
    refCount           as long
    Enumerable         As Object
    Index              As Long
    Upper              As Long
    Lower              As Long
end type

const NULL_ = 0




'
' Apparently, the declaration of GetMem4 in msvbvm60.bas is not compatible with the following declaration:
'
' private declare function GetMem4_local lib "msvbvm60" alias "GetMem4" (      src  as any    , dest   as any) as long
  private declare sub      GetMem4_local lib "msvbvm60" alias "GetMem4" (byRef src  as any    , dest   as any)
' private declare sub      GetMem4_l     lib "msvbvm60" alias "GetMem4" (byVal src  as LongPtr, dest   as any)
'         declare sub      GetMem4       Lib "msvbvm60.dll"             (ByVal addr As LongPtr, retVal As Long)

'
' It is unclear what the relationship between the function VarPtr in msvbvm60 and the function vba.varPtr()
' really is.
'
' private declare function funcPtr       lib "msvbvm60" Alias "VarPtr" (ByVal FunctionAddress As Long) As Long


public function get_IEnumVARIANT_vTbl_etc (     _
                    byRef Enumerable as object, _
                    byVal Upper      as long  , _
           optional byVal Lower      as long    _
                              ) as IEnumVARIANT
    
 '  Static VTable(6) As Long
 
     
 
     dim this As TENUMERATOR

   'if  vTbl(0) = NULL_ then
    if IEnumVariant_vtbl.IUnknown_vtbl.QueryInterface = NULL_ then
    
        IEnumVariant_vtbl.IUnknown_vtbl.QueryInterface = vba.cLng( addressOf IUnknown_QueryInterface )
        IEnumVariant_vtbl.IUnknown_vtbl.AddRef         = vba.cLng( addressOf IUnknown_AddRef         )
        IEnumVariant_vtbl.IUnknown_vtbl.Release        = vba.cLng( addressOf IUnknown_Release        )
      ' ----------------------------------------------------------------------------------------------
        IEnumVariant_vtbl.Next                         = vba.cLng( addressOf IEnumVARIANT_Next       )
        IEnumVariant_vtbl.Skip                         = vba.cLng( addressOf IEnumVARIANT_Skip       )
        IEnumVariant_vtbl.Reset                        = vba.cLng( addressOf IEnumVARIANT_Reset      )
        IEnumVariant_vtbl.Clone                        = vba.cLng( addressOf IEnumVARIANT_Clone      )
    
      '   vTbl(0) = vba.cLng( addressOf IUnknown_QueryInterface )
      '   vTbl(1) = vba.cLng( addressOf IUnknown_AddRef         )
      '   vTbl(2) = vba.cLng( addressOf IUnknown_Release        )
      '   vTbl(3) = vba.cLng( addressOf IEnumVARIANT_Next       )
      '   vTbl(4) = vba.cLng( addressOf IEnumVARIANT_Skip       )
      '   vTbl(5) = vba.cLng( addressOf IEnumVARIANT_Reset      )
      '   vTbl(6) = vba.cLng( addressOf IEnumVARIANT_Clone      )
        
    end if
    

   'with This
      ' this.VTablePtr    = vba.varPtr(vTbl(0))
        this.vTablePtr    = vba.varPtr(IEnumVariant_vtbl)
        this.Lower        = Lower
        this.Index        = Lower
        this.Upper        = Upper
        this.refCount     = 1
    set this.enumerable   = enumerable
   'End With
    
    dim pThis as longPtr
    
  '
  ' Allocate "COM" memory
  '
    pThis = CoTaskMemAlloc(lenB(This))
    
    
   ' Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" ( _
   '      ByRef dest As Any, _
   '      ByRef source As Any, _
   '      ByVal size As LongPtr)
   '      
   '      Declare Function vbaCopyBytesZero Lib "msvbvm60.dll" Alias "__vbaCopyBytesZero" (ByVal length As Long, dst As Any, src As Any) As Long
    
   'vbaCopyBytesZero lenB(this), byVal pThis, this
   
  '
  '
  '
    RtlMoveMemory                byVal pThis, this, lenB(this)
  ' debug.print "After Move memory"
  '
  ' This is sort of unbelievable, but "this" must be zeroed out.
  '
  ' Don Box states the reason for this (Advanced Visual Basic 6, p. 149):
  '    VB thinks the data in Struct needs to be  freed when the function goes out of scope VB has no
  '    way of knowing that ownership of the structure has moved elsewhere. If the
  '    structure contains object or variable-size String or array types, VB will
  '    kindly free them for you when the object goes out of scope. But you are still
  '    using the structure, so this is an unwanted favor. To prevent VB from freeing
  '    referenced memory in the stack object, simply ZeroMemory the structure. When
  '    you apply the CopyMemory call's ANSI/UNICODE precautions to ZeroMemory, you
  '    get the transfer code seen in the listing.
  '
  ' Apparently, the combination of RtlMoveMemory and RtlZeroMemory could be
  ' achieved in one go with vbaCopyBytesZero (declared in msvbvm60.dll).
  '  
    RtlZeroMemory                             this, lenB(this)
  ' debug.print "After Zero memory"
    
   'debug.print "lenB(get_IEnumVARIANT_vTbl_etc) = " & lenB(get_IEnumVARIANT_vTbl_etc)
   
   
      
   dim x as long
   
   x = GetMem4_(pThis)
   debug.print "*pThis = " & x
   debug.print "objPtr: " & objPtr(get_IEnumVARIANT_vTbl_etc)
   debug.print "varPtr: " & varPtr(get_IEnumVARIANT_vTbl_etc)
   debug.print "typeName: " & typeName(get_IEnumVARIANT_vTbl_etc)
   
    
   x = GetMem4_(varPtr(get_IEnumVARIANT_vTbl_etc))
   debug.print "*varPtr(get_IEnumVARIANT_vTbl_etc) = " & x
   
       
  ' GetMem4_local          pThis ,        get_IEnumVARIANT_vTbl_etc
  ' GetMem4_l       varPtr(pThis),        get_IEnumVARIANT_vTbl_etc
  ' GetMem4         varPtr(pThis), byRef objPtr(get_IEnumVARIANT_vTbl_etc)
  
    dim addr_IEnumVariant as longPtr
    dim addr_pThis        as longptr
    
    addr_IEnumVariant = varPtr(get_IEnumVARIANT_vTbl_etc)
    addr_pThis        = varPtr(pThis)
    
    RtlMoveMemory byVal addr_IEnumVariant, byVal addr_pThis, lenB(pThis)
    debug.print "varPtr: " & varPtr(get_IEnumVARIANT_vTbl_etc)
    debug.print "objPtr: " & objPtr(get_IEnumVARIANT_vTbl_etc)
    
    x = GetMem4_(varPtr(get_IEnumVARIANT_vTbl_etc))
    
    debug.print "*varPtr(get_IEnumVARIANT_vTbl_etc) = " & x
    debug.print "typeName: " & typeName(get_IEnumVARIANT_vTbl_etc)
    
    x = GetMem4_(objPtr(get_IEnumVARIANT_vTbl_etc))
    debug.print "x = " & x
    
  ' GetMem4        varPtr(pThis), varPtr(get_IEnumVARIANT_vTbl_etc)
    
   ' debug.print "***** GetMem4_local used in get_IEnumVARIANT_vTbl_etc"
    
      
end function

' Private Function IID$(byVal ansiString as longPtr)
'     StrRef(IID$) = SysAllocStringByteLen(ansiString, 16&)
' End Function
' 
' Private Function IID_IUnknown() As String
'     Static IID As String
'     
'     If StrPtr(IID) = NULL_ Then
'         IID = String$(8, vbNullChar)
'         IIDFromString StrPtr("{00000000-0000-0000-C000-000000000046}"), StrPtr(IID)
'     End If
' 
'     IID_IUnknown = IID
' End Function
' 
' Private Function IID_IEnumVARIANT() As String
'     Static IID As String
'     If StrPtr(IID) = NULL_ Then
'         IID = String$(8, vbNullChar)
'         IIDFromString StrPtr("{00020404-0000-0000-C000-000000000046}"), StrPtr(IID)
'     End If
'     IID_IEnumVARIANT = IID
' End Function

'    byVal riid      as longPtr    , 
private function IUnknown_QueryInterface(   _
            byRef this      as TENUMERATOR, _
                  g         as GUID__     , _
            byVal ppvObject as longPtr      _
                                         ) as long
  ' debug.print "IUnknown_QueryInterface"                              
                                         
    If  ppvObject = NULL_ Then
        IUnknown_QueryInterface = E_POINTER
        exit function
    End If

    
        
    dim siid as string
   'siid = IID$(g)
    debug.print "siid = " & siid
    
    debug.print hex(g.b_00)
    debug.print hex(g.b_01)
    debug.print hex(g.b_02)
    debug.print hex(g.b_03)
    debug.print "--"
    debug.print hex(g.b_04)
    debug.print hex(g.b_05)
    debug.print "--"
    debug.print hex(g.b_06)
    debug.print hex(g.b_07)
    debug.print "--"
    debug.print hex(g.b_08)
    debug.print hex(g.b_09)
    debug.print "--"
    
    
    
    debug.print varPtr(g.b_00)
    debug.print varPtr(g.b_01)
    debug.print varPtr(g.b_02)
    debug.print varPtr(g.b_03)
    
    
    
    if ( g.b_00 = &h00 and g.b_01 = &h00 and g.b_02 = &h00 and g.b_03 = &h00 and g.b_04 = &h00 and g.b_05 = &h00 and g.b_06 = &h00 and g.b_07 = &h00 and g.b_08 = &hc0 and g.b_09 = &h00 and g.b_0a = &h00 and g.b_0b = &h00 and g.b_0c = &h00 and g.b_0d = &h00 and g.b_0e = &h00 and g.b_0f = &h46 ) or _ 
       ( g.b_00 = &h04 and g.b_01 = &h04 and g.b_02 = &h02 and g.b_03 = &h00 and g.b_04 = &h00 and g.b_05 = &h00 and g.b_06 = &h00 and g.b_07 = &h00 and g.b_08 = &hc0 and g.b_09 = &h00 and g.b_0a = &h00 and g.b_0b = &h00 and g.b_0c = &h00 and g.b_0d = &h00 and g.b_0e = &h00 and g.b_0f = &h46 ) then
    
       debug.print "Yes yes"

   'if siid = IID_IUnknown Or siid = IID_IEnumVARIANT Then
  
        dim addr_this as longPtr
        addr_this = vba.varPtr(this)
        
        dim addr_addr_this as longPtr
        addr_addr_this  = vba.varPtr(addr_this)
        
        dim addr_ppvObject as longPtr
        addr_ppvObject = vba.varPtr(ppvObject)        
    
        debug.print "ppvObject = " &            ppvObject
        debug.print "this      = " & vba.varPtr(this)
        debug.print "lenB(ppvObject) = " & lenB(ppvObject)
      '  DeRef(ppvObject) = vba.varPtr(This)
      ' PutMem4             ppvObject, varPtr(this)
      ' RtlMoveMemory       byVal ppvObject,  byVal addr_addr_this, lenB(ppvObject)
        RtlMoveMemory       byVal ppvObject,             addr_this, lenB(ppvObject)
        
    ' Private Property Let DeRef(ByVal Address As Long, ByVal Value As Long)
     '  GetMem4_local       Value, ByVal Address        
     '   GetMem4_local    vba.varPtr(this), varPtr(ppvObject)
      ' GetMem4       varPtr(this), ppvObject
      
        IUnknown_AddRef           this
        IUnknown_QueryInterface = S_OK
        
    else
        IUnknown_QueryInterface = E_NOINTERFACE
    end If
    
End Function

private function IUnknown_AddRef(byRef this as TENUMERATOR) as long ' {

    this.refCount   = this.refCount + 1
    IUnknown_AddRef = this.refCount
    
  ' debug.print "IUnknown_AddRef, refCount = " & this.refCount
     
 '   With This
 '       .References = .References + 1
 '       IUnknown_AddRef = .References
 '   End With
end function ' }

private function IUnknown_Release(byRef this As TENUMERATOR) as long ' {

   this.refCount = this.refCount - 1
   
 ' debug.print "IUnknown_Release, refCount = " & this.refCount

   if this.refCount = 0 then ' {
      set this.enumerable = nothing
      CoTaskMemFree vba.varPtr(this)
   end if ' }

   ' With This
   '     .References = .References - 1
   '     IUnknown_Release = .References
   '     If .References = 0 Then
   '         Set .Enumerable = Nothing
   '         CoTaskMemFree vba.varPtr(this)
   '     End If
   ' End With
end function ' }

Private Function IEnumVARIANT_Next(           _
           byRef this         as TENUMERATOR, _
           byVal celt         as long       , _
           byVal rgVar        as long       , _
           byVal pceltFetched as long        _
                                   ) as long
                                   
                                   
    If rgVar = NULL_ Then
        IEnumVARIANT_Next = E_POINTER
        Exit Function
    End If
    
    Dim Fetched As Long
    With This
        Do Until .Index > .Upper
          ' VariantCopyToPtr rgVar, .Enumerable(.Index)
            VariantCopy      rgVar, .enumerable(.index)
           .Index = .Index + 1&
            Fetched = Fetched + 1&
            If Fetched = celt Then Exit Do
            rgVar = PtrAdd(rgVar, 16&)
        Loop
    End With
    
    If pceltFetched   Then 
        DLng(pceltFetched) = Fetched
        
        '  DeRef(ppvObject) = vba.varPtr(This)
      ' PutMem4             ppvObject, varPtr(this)
      ' RtlMoveMemory       byVal ppvObject,  byVal addr_addr_this, lenB(ppvObject)
      ' RtlMoveMemory       byVal ppvObject,             addr_this, lenB(ppvObject)
        
    end if
    
    If Fetched < celt Then IEnumVARIANT_Next = S_FALSE
    
End Function

Private Function IEnumVARIANT_Skip(ByRef This As TENUMERATOR, ByVal celt As Long) As Long
    IEnumVARIANT_Skip = E_NOTIMPL
End Function

Private Function IEnumVARIANT_Reset(ByRef This As TENUMERATOR) As Long
    IEnumVARIANT_Reset = E_NOTIMPL
End Function

Private Function IEnumVARIANT_Clone(ByRef This As TENUMERATOR, ByVal ppEnum As Long) As Long
    IEnumVARIANT_Clone = E_NOTIMPL
End Function

Private Function PtrAdd(ByVal Pointer As Long, ByVal Offset As Long) As Long
    PtrAdd = (Pointer Xor &H80000000) + Offset Xor &H80000000
End Function

' Private Property Let DeRef(ByVal Address As Long, ByVal Value As Long)
' 
'     debug.print "DeRef: address = " & address
'     debug.print "DeRef: value   = " & value
' 
'     GetMem4_local       Value, ByVal Address
'   ' GetMem4       byRef Value, ByVal Address
'   
' end property

Private Property Let DLng(ByVal Address As Long, ByVal Value As Long)
    GetMem4_local Value, ByVal Address
End Property

' Private Property Let StrRef(ByRef Str As String, ByVal Value As Long)
'     GetMem4_local Value, byVal vba.varPtr(str)
' End Property

