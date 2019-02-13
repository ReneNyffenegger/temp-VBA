#include <windows.h>

#include <stdlib.h>
#include <string.h>

#include "int3.h"
#include "/about/libc/search/t/tq84-tsearch.h"

#include "msg.h"
#include "txt.h"

//
// https://stackoverflow.com/a/48829052
#define free_const(x) free((void*)/*(intptr_t)*/(x))

void *breakpoint_tree_root = 0;


static char replace_byte(int3_addr addr, char new_byte) {

    static HANDLE curProcess = 0;

    if (! curProcess) {
       curProcess = GetCurrentProcess();
    }

    DWORD oldProtection;
    VirtualProtect(addr, 1, PAGE_EXECUTE_READWRITE, &oldProtection);

    char   old_byte;
    SIZE_T nofBytes;
    if (! ReadProcessMemory(curProcess, addr, &old_byte, 1, &nofBytes)) {
      MessageBox(0, "Could not ReadProcessMemory", 0, 0);
    }

    if (! WriteProcessMemory(curProcess, addr, &new_byte, 1, &nofBytes)) {
      MessageBox(0, "Could not ReadProcessMemory", 0, 0);
    }

    if (! FlushInstructionCache(curProcess, addr, 1)) {
      MessageBox(0, "Could not FlushInstructionCache", 0, 0);
    }

    writeToFile("replace_byte, returning old_byte\n");
    return old_byte;

}

void free_breakpoint(breakpoint *bp) {

    replace_byte(bp->addr, bp->orig_byte);

    free_const(bp->name);
    free      (bp      );
}

int compare_breakpoints(const void *const item_left, const void *const item_right) {

  return ((breakpoint*)item_left )->addr -
         ((breakpoint*)item_right)->addr;

}

void add_breakpoint(void *addr, const char *name) {


    breakpoint *bp = malloc(sizeof(breakpoint));
    bp->addr = addr;
//  bp->name = strdup(name);

    breakpoint *bp_in_tree = *( (breakpoint**) tq84_tsearch(bp, &breakpoint_tree_root, compare_breakpoints) );

    if (bp_in_tree != bp) {

     // Already in tree -> overwriting

        free_const(bp_in_tree->name);

        bp_in_tree->name = strdup(name);
     
        free_breakpoint(bp);
     
    }
    else {
        bp->name      = strdup(name);
        bp->orig_byte = replace_byte(addr, 0xcc); // 0xcc = INT3
    }

}

breakpoint *addr_to_breakpoint(int3_addr addr) {

    
    breakpoint **node_in_tree = tq84_tfind(&addr, &breakpoint_tree_root, compare_breakpoints);
    
    if (node_in_tree) {
       return *node_in_tree;
    }

    return NULL;

}


void set_int3(breakpoint *bp) {
    replace_byte(bp->addr, 0xcc); // 0xcc = INT3
}

void int3_to_orig(breakpoint *bp) {
    replace_byte(bp->addr, bp->orig_byte);
}

void rm_all_int3() {

    tq84_tdestroy(breakpoint_tree_root, (tq84__free_fn_t) free_breakpoint);
}
