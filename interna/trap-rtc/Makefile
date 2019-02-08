all: trap-rtc.dll

trap-rtc.dll: trap-rtc.o msg.o funcsInDll.o error.o
	gcc -shared trap-rtc.o msg.o funcsInDll.o error.o -lOleAut32 -lImagehlp -o trap-rtc.dll -Wl,--add-stdcall-alias

trap-rtc.o: trap-rtc.c
	gcc -c trap-rtc.c -o trap-rtc.o

msg.o: msg.c
	gcc -c msg.c -o msg.o

error.o: error.c
	gcc -c error.c -o error.o

funcsInDll.o: funcsInDll.c
	gcc -c funcsInDll.c -o funcsInDll.o
