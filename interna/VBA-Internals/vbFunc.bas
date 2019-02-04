option explicit

sub vbSub() ' {

end sub ' }

sub tryToDebugCallTovbSub()

    debug.print(vba.int(addressOf vbSub))

    addBreakpoint vba.int(addressOf vbSub), "vbSub"

end sub
