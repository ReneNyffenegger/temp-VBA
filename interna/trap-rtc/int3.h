
typedef void *int3_addr;

typedef struct
{
    int3_addr   addr;
    const char *name;
    char        orig_byte;
}
breakpoint;


void add_breakpoint(int3_addr addr, const char *name);

// void rm_int3(int3_addr addr);

breakpoint *addr_to_breakpoint(int3_addr addr);
// void int3_exception(int3_addr addr);

//void single_step(int3_addr addr);

void set_int3(breakpoint*);
void int3_to_orig(breakpoint*);

void rm_all_int3();
