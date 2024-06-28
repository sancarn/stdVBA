;This is the basis for the thunk
; The expectation is we will prepare the `method_name`, `dispparams` struct and other required items in VBA
; and then supply a pointer to thunk which will call a public member by name.
; It'll be a little slow but hopefully very stable.

section .data
    method_name dw 'MethodName', 0  ; Replace 'MethodName' with the actual method name
    disp_params dd 2                ; Placeholder for DISPPARAMS structure (adjust size as needed)
    arg_count dd 1                  ; Example argument count
    pDisp dq 0                      ; Placeholder for the object pointer (set this to the actual object pointer)
    iid_null db 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00
    localeID dd 0x0409              ; Locale ID (e.g., 0x0409 for English - United States)

section .bss
    dispid resd 1
    hresult resd 1
    refcount resd 1

section .text
    global call_dispatch_methods
    call_dispatch_methods:
        ; Save registers that we will use
        push rbx
        push rsi
        push rdi

        ; Load the method name and DISPPARAMS from the data section
        lea rcx, [method_name]        ; Pointer to the method name (wchar_t*)
        lea rdx, [disp_params]        ; Pointer to the DISPPARAMS struct
        mov r8d, [arg_count]          ; Count of arguments
        mov r9, qword [pDisp]         ; Pointer to the object to call on

        ; 1. Get function pointers from vtable
        mov rdi, r9                    ; pDisp (the object pointer)
        mov rax, [rdi]                 ; Load vtable
        mov rbx, [rax + 0x38]          ; GetIDsOfNames is the 7th function (offset 0x38)
        mov rsi, [rax + 0x40]          ; Invoke is the 9th function (offset 0x40)

        ; 2. Check reference count
        mov eax, [rdi + 0x18]           ; Load reference count (example offset, may vary)
        mov [refcount], eax
        test eax, eax
        jz zero_refcount_failed        ; Jump if reference count is zero

        ; 3. Call IDispatch::GetIDsOfNames
        mov rdi, r9                    ; pDisp (the object pointer)
        lea r8, [localeID]             ; lcid
        lea rdx, [dispid]              ; rgDispId
        lea rcx, [iid_null]            ; riid (IID_NULL)
        mov r9d, 1                     ; cNames (only one name)
        call rbx                       ; Call GetIDsOfNames

        ; Check HRESULT
        mov [hresult], eax
        test eax, eax
        js get_ids_of_names_failed     ; Jump if failed (negative HRESULT)

        ; 4. Call IDispatch::Invoke
        ; DISPID dispid obtained from GetIDsOfNames is stored in [dispid]
        mov rdi, r9                    ; pDisp (the object pointer)
        mov esi, [dispid]              ; dispIdMember
        lea rcx, [iid_null]            ; riid (IID_NULL)
        lea r8, [localeID]             ; lcid
        mov rdx, rdx                   ; pVarResult (assuming NULL)
        mov r9, rdx                    ; pExcepInfo (assuming NULL)
        xor r9d, r9d                   ; puArgErr (assuming NULL)
        mov r9, rdx                    ; pDispParams (DISPPARAMS struct pointer)
        call rsi                       ; Call Invoke

        ; Check HRESULT
        mov [hresult], eax

        ; Restore registers and return
    get_ids_of_names_failed:
    zero_refcount_failed:
        pop rdi
        pop rsi
        pop rbx
        ret

