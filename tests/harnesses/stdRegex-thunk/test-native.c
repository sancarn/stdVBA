/* Execute the extracted blob through the same Windows API used by VBA.
 * Build/run as both x86 and x64; x86 runs under WoW64 on an x64 host.
 * This checks ABI/layout, stack cleanup, result sign extension and semantics.
 */
#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include "../../../Tools/stdRegex-thunk/matcher.c"
_Static_assert(sizeof(I)==4, "VBA Long is 32 bits");
_Static_assert(sizeof(Args)==(sizeof(void*)==8?72:52), "VBA NativeArgs layout");
static I canon_table[65536], workspace[3*4096], captures[20];
static int checks;
static void check(WNDPROC entry, I *bc, I bc_length, const U16 *text,
                  I length, I limit, I multiline, I dotall, I capacity,
                  I expected, I capture_start, I capture_end) {
    Args a={bc,canon_table,text,workspace,captures,length,0,limit,multiline,dotall,1,capacity,bc_length};
    LRESULT value=CallWindowProcW(entry,(HWND)&a,0,0,0);
    if(value!=expected || (expected>=0 && (captures[0]!=capture_start || captures[1]!=capture_end))) {
        fprintf(stderr,"case %d: result %lld expected %d; captures %d,%d expected %d,%d\n",
                checks,(long long)value,expected,captures[0],captures[1],capture_start,capture_end);
        exit(1);
    }
    checks++;
}
int main(int argc,char **argv) {
    FILE *f; long size; unsigned char *code; DWORD old; int iteration;
    I literal[]={1,0,0,11,0,2,97,11,1,1};
    I insensitive[]={1,0,0,11,0,514,65,11,1,1};
    I dot[]={1,0,0,11,0,3,11,1,1};
    I range[]={1,0,0,11,0,4,1,97,122,11,1,1};
    I inverse[]={1,0,0,11,0,5,1,97,122,11,1,1};
    I start[]={1,0,0,11,0,16,2,97,11,1,1};
    I end[]={1,0,0,11,0,2,97,17,11,1,1};
    I boundary[]={1,0,0,11,0,18,2,97,11,1,1};
    I backref[]={3,0,0,11,0,11,2,2,97,11,3,15,1,11,1,1};
    I split[]={1,0,0,11,0,7,4,2,98,6,2,2,97,11,1,1};
    I invalid[]={1,0,0,0};
    U16 a[]={97}, aa[]={97,97}, ba[]={98,97}, newline[]={10}, multiline_text[]={10,97}, digit[]={49};
    if(argc!=2) return 2;
    f=fopen(argv[1],"rb"); if(!f) return 2;
    fseek(f,0,SEEK_END);size=ftell(f);rewind(f);
    code=VirtualAlloc(NULL,size,MEM_COMMIT|MEM_RESERVE,PAGE_READWRITE);
    if(!code || fread(code,1,size,f)!=(size_t)size) return 2;
    fclose(f);
    if(!VirtualProtect(code,size,PAGE_EXECUTE_READ,&old)||!FlushInstructionCache(GetCurrentProcess(),code,size)) return 2;
    canon_table[97]=-32;
#define RUN(b,t,l,expected,s,e) check((WNDPROC)code,b,sizeof(b)/sizeof(I),t,l,10000,0,0,4096,expected,s,e)
    for(iteration=0;iteration<1000;iteration++) {
        RUN(literal,a,1,1,0,1); RUN(literal,ba,2,2,1,2);
        RUN(literal,digit,1,-1,0,0); RUN(literal,a,0,-1,0,0);
        RUN(insensitive,a,1,1,0,1); RUN(dot,newline,1,-1,0,0);
        check((WNDPROC)code,dot,9,newline,1,10000,0,1,4096,1,0,1);
        RUN(range,a,1,1,0,1); RUN(inverse,digit,1,1,0,1);
        RUN(start,multiline_text,2,-1,0,0);
        check((WNDPROC)code,start,11,multiline_text,2,10000,1,0,4096,2,1,2);
        RUN(end,aa,2,2,1,2); RUN(boundary,ba,2,-1,0,0);
        RUN(backref,aa,2,2,0,2); RUN(split,a,1,1,0,1);
        check((WNDPROC)code,literal,10,a,1,1,0,0,4096,-2,0,0);
        check((WNDPROC)code,literal,10,a,1,10000,0,0,1,-3,0,0);
        RUN(invalid,a,1,-4,0,0);
    }
    if(!VirtualFree(code,0,MEM_RELEASE)) return 2;
    printf("%u-bit: %d checks passed via CallWindowProcW\n",(unsigned)(8*sizeof(void*)),checks);
    return 0;
}
