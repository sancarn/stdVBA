/* Experimental native port of sihlfall/vba-regex's DFS interpreter.
 * Original MIT license is retained in the generated class and LICENSE.engine.
 * No CRT, imports, globals, callbacks or relocations: one self-contained entry.
 * The compiler and Unicode canonicalization table remain the original VBA ones.
 */
typedef int I;
typedef unsigned short U16;
typedef struct {
    const I *bc, *canon;
    const U16 *text;
    I *work, *out;
    I length, start, limit, multiline, dotall, raise_limit, capacity, bc_length;
} Args;
typedef struct { I master, clen, cow, qlen, pc, sp, landmark, delta, q, top; } Frame;
#define INLINE static __inline__ __attribute__((always_inline))
INLINE I peek(Args *a, I p) { return p < 0 || p >= a->length ? -1 : a->text[p]; }
INLINE I get(Args *a, I *p, I d) { I c=peek(a,*p); if(c>=0) *p+=d; return c; }
INLINE I canon(Args *a,I c) { return (unsigned)c>65535 ? c : (c+a->canon[c])&65535; }
INLINE I line(I c) { return c==10 || c==13 || c==0x2028 || c==0x2029; }
INLINE I word(I c) { return (c>=65&&c<=90)||(c>=97&&c<=122)||(c>=48&&c<=57)||c==95; }

/* -1 no match; -2 step limit; -3 workspace exhausted (retry larger);
 * -4 invalid internal bytecode. Nonnegative return is the end position.
 * Each of three workspace regions has capacity integers, owned by VBA.
 */
#ifdef _WIN64
typedef long long Result;
#else
typedef long Result;
#endif
Result __attribute__((stdcall)) regex_entry(Args *a, unsigned ignored, void *unused1, void *unused2) {
    I cap=a->capacity, np=a->bc[0]+1, nc=np+a->bc[1];
    I *cs=a->work, *qs=cs+cap;
    Frame *frames=(Frame*)(qs+cap);
    I maxframes=cap/(sizeof(Frame)/sizeof(I));
    I candidate, i, pc, sp, landmark, delta, q, master, clen, cow, qlen, top, flen;
    I op,n,t,c1,c2,aa,bb,mm,idx,off,r1,r2,qmax,qq,steps,oldclen;
    Frame f;
    (void)ignored; (void)unused1; (void)unused2;
    if(nc<2 || nc>cap || maxframes<1) return -3;
#define BC() (pc>=0 && pc<a->bc_length ? a->bc[pc++] : 0)
#define PUSH(P,S,D,Q) do { \
    if(flen>=maxframes) return -3; \
    f.master=master; f.clen=clen; f.cow=cow; f.qlen=qlen; \
    f.pc=(P); f.sp=(S); f.landmark=landmark; f.delta=(D); f.q=(Q); f.top=top; \
    frames[flen++]=f; cow=1; \
} while(0)
#define RESTORE() do { \
    master=f.master; clen=f.clen; cow=f.cow; qlen=f.qlen; \
    pc=f.pc; sp=f.sp; landmark=f.landmark; delta=f.delta; q=f.q; top=f.top; \
} while(0)
#define SETCAP(K,V) do { \
    if(cow) { if(clen>cap-nc) return -3; \
        for(i=0;i<nc;i++) cs[clen+i]=cs[clen-nc+i]; clen+=nc; cow=0; } \
    cs[clen-nc+(K)]=(V); \
} while(0)
#define QPUSH(V) do { if(qlen>cap-3) return -3; \
    qs[qlen]=top; qs[qlen+1]=flen; qs[qlen+2]=(V); qlen+=3; top=qlen-1; } while(0)
#define QPOP(V) do { if(!top) (V)=-2; else { \
    (V)=qs[top]; if(qs[top-1]==flen) qlen-=3; top=qs[top-2]; } } while(0)
    for(candidate=a->start; candidate<=a->length; candidate++) {
        clen=nc; for(i=0;i<nc;i++) cs[i]=-1;
        qlen=top=flen=cow=0; master=-1;
        pc=3+3*a->bc[1]; sp=candidate; landmark=-1; delta=1; q=-2; steps=0;
        goto success;
fail:
        if(!flen) continue;
        f=frames[--flen]; RESTORE();
success:
        if(++steps>=a->limit) { if(a->raise_limit) return -2; continue; }
        op=BC();
        switch(op&255) {
        case 1:
            for(i=0;i<nc;i++) a->out[i]=cs[clen-nc+i];
            return sp;
        case 31: /* positive lookaround succeeds, preserve its captures */
            if(master<0 || master>=flen) return -4;
            oldclen=clen; flen=master; f=frames[flen]; RESTORE();
            if(oldclen!=clen) {
                if(cow) { clen+=nc; cow=0; }
                for(i=1;i<=nc;i++) cs[clen-i]=cs[oldclen-i];
            }
            pc++; n=BC(); pc+=n; goto success;
        case 32:
            if(master<0 || master>=flen) return -4;
            flen=master; f=frames[flen]; RESTORE(); goto fail;
        case 2:
            landmark=pc-1; c1=BC(); c2=get(a,&sp,delta);
            if(op&512) c2=canon(a,c2);
            if(c1!=c2) goto fail; goto success;
        case 3:
            landmark=pc-1; c1=get(a,&sp,delta);
            if(c1<0) goto fail;
            if((op-(a->dotall?4096:0))&8192) goto success;
            if(line(c1)) goto fail; goto success;
        case 4: case 5:
            landmark=pc-1; n=BC(); c1=get(a,&sp,delta);
            if(c1<0) goto fail;
            if(op&512) c1=canon(a,c1);
            aa=pc-1; pc+=2*n; bb=pc+1;
            if(n<1 || pc>a->bc_length) return -4;
            for(;;) { mm=aa+2*((bb-aa)/4);
                if(a->bc[mm]>=c1) bb=mm; else aa=mm;
                if(bb-aa==2) break;
            }
            t=bb<pc && a->bc[bb-1]<=c1;
            if(((op&255)==4)!=t) goto fail; goto success;
        case 16:
            if(sp+(1-delta)/2<=0) goto success;
            if(!((op-(a->multiline?1024:0))&2048)) goto fail;
            if(line(peek(a,sp-(delta+1)/2))) goto success; goto fail;
        case 17:
            c1=peek(a,sp-(delta-1)/2); if(c1==-1) goto success;
            if(!((op-(a->multiline?1024:0))&2048)) goto fail;
            if(line(c1)) goto success; goto fail;
        case 18: case 19:
            aa=sp+(1-delta)/2;
            c1=aa>0 && word(peek(a,aa-1)); c2=aa<a->length && word(peek(a,aa));
            if(((op&255)==18)==(c1==c2)) goto fail; goto success;
        case 6:
            n=BC(); if(n>0) {pc+=n; goto success;}
            t=pc+n; if(landmark<=t) goto success;
            pc=landmark=t; goto success;
        case 7:
            n=BC(); PUSH(pc+n,sp,delta,q); goto success;
        case 8:
            n=BC(); PUSH(pc,sp,delta,q); pc+=n; goto success;
        case 20:
            QPUSH(q); q=0; goto success;
        case 21: pc+=2; goto success;
        case 22:
            qmax=BC(); n=BC(); q++;
            if(q<qmax) { t=pc-n-3; if(landmark>t) landmark=t; pc=t; }
            else { QPOP(q); } goto success;
        case 23: case 26:
            QPUSH(q); q=-1; goto success;
        case 24:
            qmax=BC(); n=BC(); q++;
            if(q<qmax) { PUSH(pc,sp,delta,q); }
            QPOP(q); pc+=n; goto success;
        case 25:
            pc++; n=BC(); t=pc-n-3; if(landmark<=t) goto fail;
            pc=landmark=t; goto success;
        case 27:
            qmax=BC(); n=BC(); q++;
            if(q<qmax) { QPOP(qq); PUSH(pc+n,sp,delta,qq); QPUSH(qq); }
            else { pc+=n; QPOP(q); } goto success;
        case 28:
            pc++; n=BC(); t=pc-n-3; if(landmark<=t) goto success;
            pc=landmark=t; goto success;
        case 11:
            idx=BC(); if(idx<0||idx>=nc) return -4;
            SETCAP(idx,sp); goto success;
        case 12:
            r1=BC(); r2=BC(); if(r1<0||np+r1>=nc) return -4;
            SETCAP(np+r1,r2); goto success;
        case 29: case 30:
            PUSH(pc,sp,delta,q); master=flen-1; pc+=2;
            if((op&255)==29) { sp+=(1-delta)/2; delta=1; }
            else { sp-=(delta+1)/2; delta=-1; } goto success;
        case 13: n=BC(); pc+=n; goto fail;
        case 14: n=BC(); pc+=n; goto success;
        case 15:
            landmark=pc-1; idx=2*BC(); if(idx<2||idx+1>=nc) return -4;
            aa=cs[clen-nc+idx]; bb=cs[clen-nc+idx+1];
            if(aa>=0&&bb>=0) {
                off=delta==1?aa:bb-1;
                while(delta==1?off<bb:off>=aa) {
                    c1=get(a,&off,delta); c2=get(a,&sp,delta);
                    if(c1!=c2 && (!(op&512)||canon(a,c1)!=canon(a,c2))) goto fail;
                }
            } goto success;
        default: return -4;
        }
    }
    return -1;
}
