option explicit

sub f(b1 as boolean, b2 as boolean, b3 as boolean) ' {

    debug.print b1 & chr(9) & b2 & chr(9) & b3
end sub ' }


sub g(n as long) ' {

    f n and 1, n and 2, n and 4

end sub ' }

sub main() ' {

    dim i as long

    for i = 0 to 10
        g i
    next i

end sub ' }
