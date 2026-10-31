; SUBCLASSPROC -> stdICallable.Run(ParamArray), Win64 ABI.
; Four arguments are boxed on the stack, with no Automation heap allocation.
; 0x1122334455667788 is patched to the shared active-call counter.
; 0x2233445566778899 is patched to DefSubclassProc.
; 0x33445566778899AA is patched to VBE EbMode.
; 0x445566778899AABB is patched to RemoveWindowSubclass.
; The matching unwind record describes this seven-byte stack allocation.
sub rsp, 216
mov [rsp+120], rcx
mov [rsp+144], edx
mov [rsp+168], r8
mov [rsp+192], r9
xor eax, eax
mov [rsp+32], rax
mov [rsp+40], rax
mov [rsp+48], rax
mov [rsp+64], rax
mov [rsp+72], rax
mov [rsp+80], rax
mov [rsp+88], rax
mov [rsp+112], rax
mov [rsp+128], rax
mov [rsp+136], rax
mov [rsp+152], rax
mov [rsp+160], rax
mov [rsp+176], rax
mov [rsp+184], rax
mov [rsp+200], rax
mov word ptr [rsp+112], 20
mov word ptr [rsp+136], 3
mov word ptr [rsp+160], 20
mov word ptr [rsp+184], 20
mov dword ptr [rsp+64], 0x08120001
mov dword ptr [rsp+68], 24
mov dword ptr [rsp+72], 1
lea rax, [rsp+112]
mov [rsp+80], rax
mov dword ptr [rsp+88], 4
lea rax, [rsp+64]
mov [rsp+96], rax
mov r10, 0x1122334455667788
inc dword ptr [r10]
; The low dwRefData bit opts out of IDE mode checks for this registration.
test byte ptr [rsp+264], 1
jnz invoke
; Check the IDE before touching the VBA interface, including AddRef/Release.
mov rax, 0x33445566778899AA
call rax
cmp eax, 1
je invoke
cmp eax, 2
jne detach
; Break mode forwards, but a dying window must still detach natively.
cmp dword ptr [rsp+144], 0x82
jne forward
detach:
mov rcx, [rsp+120]
mov rdx, 0x1122334455667788
sub rdx, 4096
mov r8, [rsp+256]
mov rax, 0x445566778899AABB
call rax
test eax, eax
jz forward
mov r10, 0x1122334455667788
dec dword ptr [r10+4]
jmp forward
invoke:
mov rcx, [rsp+264]
and rcx, -2
mov [rsp+104], rcx
mov rax, [rcx]
call qword ptr [rax+8]
mov rcx, [rsp+104]
mov rax, [rcx]
lea rdx, [rsp+96]
lea r8, [rsp+32]
call qword ptr [rax+56]
mov [rsp+208], eax
mov rcx, [rsp+104]
mov rax, [rcx]
call qword ptr [rax+16]
mov eax, [rsp+208]
test eax, eax
js forward
mov rax, [rsp+40]
jmp done
forward:
mov rcx, [rsp+120]
mov edx, [rsp+144]
mov r8, [rsp+168]
mov r9, [rsp+192]
mov rax, 0x2233445566778899
call rax
done:
mov r10, 0x1122334455667788
dec dword ptr [r10]
add rsp, 216
ret

