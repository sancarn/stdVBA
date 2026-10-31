; SUBCLASSPROC -> stdICallable.Run(ParamArray), x86 stdcall ABI.
; Stack-only SAFEARRAY and VARIANTs; only EBP is nonvolatile.
; Counter and DefSubclassProc placeholders are patched before execution.
push ebp
mov ebp, esp
sub esp, 128
xor eax, eax
mov [ebp-16], eax
mov [ebp-12], eax
mov [ebp-8], eax
mov [ebp-4], eax
mov [ebp-40], eax
mov [ebp-36], eax
mov [ebp-32], eax
mov [ebp-28], eax
mov [ebp-24], eax
mov [ebp-20], eax
mov [ebp-108], eax
mov [ebp-104], eax
mov [ebp-100], eax
mov [ebp-96], eax
mov [ebp-92], eax
mov [ebp-88], eax
mov [ebp-84], eax
mov [ebp-80], eax
mov [ebp-76], eax
mov [ebp-72], eax
mov [ebp-68], eax
mov [ebp-64], eax
mov [ebp-60], eax
mov [ebp-56], eax
mov [ebp-52], eax
mov [ebp-48], eax
mov dword ptr [ebp-108], 3
mov dword ptr [ebp-92], 3
mov dword ptr [ebp-76], 3
mov dword ptr [ebp-60], 3
mov eax, [ebp+8]
mov [ebp-100], eax
mov eax, [ebp+12]
mov [ebp-84], eax
mov eax, [ebp+16]
mov [ebp-68], eax
mov eax, [ebp+20]
mov [ebp-52], eax
mov dword ptr [ebp-40], 0x08120001
mov dword ptr [ebp-36], 16
mov dword ptr [ebp-32], 1
lea eax, [ebp-108]
mov [ebp-28], eax
mov dword ptr [ebp-24], 4
lea eax, [ebp-40]
mov [ebp-44], eax
mov eax, 0x11223344
inc dword ptr [eax]
test byte ptr [ebp+28], 1
jnz invoke
; EbMode: Run=1, Break=2. Never inspect the COM pointer in other modes.
mov eax, 0x33445566
call eax
cmp eax, 1
je invoke
cmp eax, 2
jne detach
cmp dword ptr [ebp+12], 0x82
jne forward
detach:
push dword ptr [ebp+24]
mov eax, 0x11223344
sub eax, 4096
push eax
push dword ptr [ebp+8]
mov eax, 0x44556677
call eax
test eax, eax
jz forward
mov eax, 0x11223344
dec dword ptr [eax+4]
jmp forward
invoke:
mov eax, [ebp+28]
and eax, -2
push eax
mov eax, [eax]
call dword ptr [eax+4]
lea eax, [ebp-16]
push eax
lea eax, [ebp-44]
push eax
mov eax, [ebp+28]
and eax, -2
push eax
mov eax, [eax]
call dword ptr [eax+28]
mov [ebp-116], eax
mov eax, [ebp+28]
and eax, -2
push eax
mov eax, [eax]
call dword ptr [eax+8]
mov eax, [ebp-116]
test eax, eax
js forward
mov eax, [ebp-8]
jmp done
forward:
push dword ptr [ebp+20]
push dword ptr [ebp+16]
push dword ptr [ebp+12]
push dword ptr [ebp+8]
mov eax, 0x22334455
call eax
done:
mov edx, 0x11223344
dec dword ptr [edx]
mov esp, ebp
pop ebp
ret 24

