vTable.dll: vTable.s
	gcc -shared $< -o $@  -Wl,--add-stdcall-alias

vTable.s: vTable.c
	gcc -S $< -o $@ 
