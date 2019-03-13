.SILENT:

prog.exe: main.obj
	link main.obj /nodefaultlib /subsystem:console /entry:run kernel32.lib user32.lib /out:prog.exe

main.obj: main.c
	cl /DUNICODE /c /W4 main.c /Fomain.obj
