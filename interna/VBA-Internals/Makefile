
VBA-Internals.dll: VBA-Internals.o mhook.o disasm.o disasm_x86.o cpu.o 
	gcc -shared VBA-Internals.o mhook.o disasm.o disasm_x86.o cpu.o -lOle32 -lOleAut32 -o VBA-Internals.dll -Wl,--add-stdcall-alias

VBA-Internals.o: VBA-Internals.c
	gcc -c VBA-Internals.c

mhook.o: \temp\mhook\mhook-lib\mhook.cpp
	gcc -c \temp\mhook\mhook-lib\mhook.cpp

disasm.o: disasm/disasm.c
	gcc -c disasm/disasm.c

disasm_x86.o: disasm\disasm_x86.c
	gcc -c disasm/disasm_x86.c

cpu.o: \temp\mhook\disasm-lib\cpu.c
	gcc -c \temp\mhook\disasm-lib\cpu.c
