vTable.dll: vTable.s
	gcc -shared $< -o $@  -Wl,--add-stdcall-alias -lOle32

vTable.s: vTable.i
	gcc -S $< -o $@ 

vTable.i: vTable.c
	gcc -E -CC $< -o $@
