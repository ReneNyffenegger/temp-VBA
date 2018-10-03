all: vTable.dll vTable.objdump

vTable.objdump: vTable.o
	objdump -d -S $< > $@

vTable.dll: vTable.o
	gcc    -shared $< -o $@  -Wl,--add-stdcall-alias -lOle32

vTable.o: vTable.s
	gcc -g -c $< -o $@

vTable.s: vTable.i
	gcc -O0 -S $< -o $@ 

vTable.i: vTable.c Makefile
	gcc -E -CC $< -o $@
