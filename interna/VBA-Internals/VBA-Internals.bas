option explicit

'       lib "C:\Users\r.nyffenegger\personal\VBA-Internals\VBA-Internals.dll" (  _
'       lib "c:\temp\temp-VBA\interna\VBA-Internals\VBA-Internals.dll" (  _

declare sub addDllFunctionBreakpoint                                             _
           lib "VBA-Internals.dll" (  _
           byVal module   as string,                                                   _
           byVal funcName as string                                                    _
        )

'       lib "C:\Users\r.nyffenegger\personal\VBA-Internals\VBA-Internals.dll" (  
declare sub VBAInternalsInit                                                     _
           lib "VBA-Internals.dll" (  _
           byVal addrCallBack as long                                            _
        )

sub init() ' {
    VBAInternalsInit(vba.int(addressOf callBack))

    call addDllFunctionBreakpoint("VBE7.dll", "rtcMsgBox")
    call addDllFunctionBreakpoint("VBE7.dll", "rtcRound" )
end sub ' }

sub main() '
    dim i as integer
    debug.print "Does next line implicitely call rtcRound?"
    i = 42.5

    debug.print "Next line explicititely calls rtcRound (via round)"
    i = round(42.7)

'   call addDllFunctionBreakpoint("VBE7.dll", "rtcRound" )

    debug.print "Does next line call rtcMsgBox?"
    MsgBox "i = " & i

    debug.print "Calling round again"
    i = round(42.7)
end sub ' }

sub callBack(byVal txt as string) ' {
    debug.print("CallBack, txt = " & txt)
end sub ' }
