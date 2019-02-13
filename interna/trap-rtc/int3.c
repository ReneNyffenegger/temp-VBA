#include <stdlib.h>
#include <string.h>

#include "int3.h"
#include "/about/libc/search/t/tq84-tsearch.h"


//
// https://stackoverflow.com/a/48829052
#define free_const(x) free((void*)/*(intptr_t)*/(x))

void *breakpoint_tree = 0;

typedef struct
{
    int3_addr   addr;
    const char *name;
}
breakpoint;


void free_breakpoint(breakpoint *bp) {

    free_const(bp->name);
    free      (bp      );
}

int compare_breakpoints(const void *const item_left, const void *const item_right) {

  return ((breakpoint*)item_left )->addr -
         ((breakpoint*)item_right)->addr;

}

void add_int3(void *addr, const char *name) {


    breakpoint *bp = malloc(sizeof(breakpoint));
    bp->addr = addr;
//  bp->name = strdup(name);

    breakpoint *bp_in_tree = *( (breakpoint**) tq84_tsearch(bp, &breakpoint_tree, compare_breakpoints) );

    if (bp_in_tree != bp) {

     // Already in tree -> overwriting

      free_const(bp_in_tree->name);

      bp_in_tree->name = strdup(name);
     
      free_breakpoint(bp);
     
    }
    else {
      bp->name = strdup(name);
    }

}
