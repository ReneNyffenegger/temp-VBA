all: trap-rtc.dll

#
#  -e to specify entry point (because otherwise, error message would be: »cannot find entry symbol _DllMainCRTStartup@12«)
#
trap-rtc.dll: trap-rtc.o msg.o funcsInDll.o error.o txt.o int3.o tq84-search.o
	rem gcc -shared trap-rtc.o msg.o funcsInDll.o error.o txt.o int3.o tq84-search.o -nostdlib -lKernel32 -lUser32 -lOleAut32 -lImagehlp  -e _DllMain -o trap-rtc.dll -Wl,--add-stdcall-alias
	    gcc -shared trap-rtc.o msg.o funcsInDll.o error.o txt.o int3.o tq84-search.o           -lKernel32 -lUser32 -lOleAut32 -lImagehlp  -e _DllMain -o trap-rtc.dll -Wl,--add-stdcall-alias

trap-rtc.o: trap-rtc.c
	gcc -c trap-rtc.c -o trap-rtc.o

msg.o: msg.c
	gcc -c msg.c -o msg.o

error.o: error.c
	gcc -c error.c -o error.o

funcsInDll.o: funcsInDll.c
	gcc -c funcsInDll.c -o funcsInDll.o

txt.o: txt.c
	gcc -c txt.c -o txt.o

int3.o: int3.c
	gcc -c int3.c -o int3.o

tq84-search.o: /about/libc/search/t/tq84-tsearch.c
	gcc -c $< -o $@
