Private Const CL_FALSE                          As Long = 0
Private Const CL_TRUE                           As Long = 1

Private Const CL_MEM_READ_WRITE               As Currency = 1 / 10000@
Private Const CL_MEM_WRITE_ONLY               As Currency = 2 / 10000@
Private Const CL_MEM_READ_ONLY                As Currency = 4 / 10000@
Private Const CL_MEM_USE_HOST_PTR             As Currency = 8 / 10000@
Private Const CL_MEM_ALLOC_HOST_PTR           As Currency = 16 / 10000@
Private Const CL_MEM_COPY_HOST_PTR            As Currency = 32 / 10000@
Private Const CL_MEM_HOST_WRITE_ONLY          As Currency = 64 / 10000@
Private Const CL_MEM_HOST_READ_ONLY           As Currency = 128 / 10000@
Private Const CL_MEM_HOST_NO_ACCESS           As Currency = 256 / 10000@
Private Const CL_MEM_SVM_FINE_GRAIN_BUFFER    As Currency = 512 / 10000@
Private Const CL_MEM_SVM_ATOMICS              As Currency = 1024 / 10000@
Private Const CL_MEM_KERNEL_READ_AND_WRITE    As Currency = 2048 / 10000@

Private Const CL_MAP_READ                     As Currency = 1 / 10000@
Private Const CL_MAP_WRITE                    As Currency = 2 / 10000@
Private Const CL_MAP_WRITE_INVALIDATE_REGION  As Currency = 4 / 10000@

Private Const CL_DEVICE_TYPE_DEFAULT          As Currency = 1 / 10000@
Private Const CL_DEVICE_TYPE_CPU              As Currency = 2 / 10000@
Private Const CL_DEVICE_TYPE_GPU              As Currency = 4 / 10000@
Private Const CL_DEVICE_TYPE_ACCELERATOR      As Currency = 8 / 10000@
Private Const CL_DEVICE_TYPE_CUSTOM           As Currency = 16 / 10000@
Private Const CL_DEVICE_TYPE_ALL              As Currency = &HFFFFFFFF


#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Function clGetPlatformIDs Lib "opencl" (ByVal num_entries As Long, cl_platform_id As LongPtr, num_platforms As Any) As Long
    Private Declare PtrSafe Function clGetDeviceIDs Lib "opencl" (ByVal platform As LongPtr, ByVal device_type As Currency, ByVal num_entries As Long, devices As LongPtr, num_devices As Any) As Long
    Private Declare PtrSafe Function clReleaseDevice Lib "opencl" (ByVal Device As LongPtr) As Long
    Private Declare PtrSafe Function clCreateContext Lib "opencl" (ByVal properties As Long, ByVal num_devices As Long, devices As LongPtr, ByVal pfn_notify As Long, ByVal user_data As Long, errcode_ret As Any) As LongPtr
    Private Declare PtrSafe Function clReleaseContext Lib "opencl" (ByVal context As LongPtr) As Long
    Private Declare PtrSafe Function clCreateCommandQueue Lib "opencl" (ByVal context As LongPtr, ByVal Device As LongPtr, ByVal properties As Currency, errcode_ret As Any) As LongPtr
    Private Declare PtrSafe Function clReleaseCommandQueue Lib "opencl" (ByVal command_queue As LongPtr) As Long
    Private Declare PtrSafe Function clCreateProgramWithSource Lib "opencl" (ByVal context As LongPtr, ByVal count As Long, strings As Any, lengths As Any, errcode_ret As Any) As LongPtr
    Private Declare PtrSafe Function clBuildProgram Lib "opencl" (ByVal program As LongPtr, ByVal num_devices As Long, device_list As Any, ByVal options As String, ByVal pfn_notify As Long, ByVal user_data As Long) As Long
    Private Declare PtrSafe Function clCreateKernel Lib "opencl" (ByVal program As LongPtr, ByVal kernel_name As String, errcode_ret As Any) As LongPtr
    Private Declare PtrSafe Function clReleaseKernel Lib "opencl" (ByVal kernel As LongPtr) As Long
    Private Declare PtrSafe Function clCreateBuffer Lib "opencl" (ByVal context As LongPtr, ByVal flags As Currency, ByVal size As Long, ByVal host_ptr As LongPtr, errcode_ret As Any) As LongPtr
    Private Declare PtrSafe Function clReleaseMemObject Lib "opencl" (ByVal memobj As LongPtr) As Long
    Private Declare PtrSafe Function clSetKernelArg Lib "opencl" (ByVal kernel As LongPtr, ByVal arg_index As Long, ByVal arg_size As Long, arg_value As Any) As Long
    Private Declare PtrSafe Function clEnqueueNDRangeKernel Lib "opencl" (ByVal command_queue As LongPtr, ByVal kernel As LongPtr, ByVal work_dim As Long, global_work_offset As Any, global_work_size As Any, local_work_size As Any, ByVal num_events_in_wait_list As Long, event_wait_list As Any, event_ As Any) As Long
    Private Declare PtrSafe Function clFinish Lib "opencl" (ByVal command_queue As LongPtr) As Long
    Private Declare PtrSafe Function clEnqueueMapBuffer Lib "opencl" (ByVal command_queue As LongPtr, ByVal buffer As LongPtr, ByVal blocking_map As Long, ByVal map_flags As Currency, ByVal offset As LongPtr, ByVal size As LongPtr, ByVal num_events_in_wait_list As Long, event_wait_list As Any, event_ As Any, errcode_ret As Any) As LongPtr
    Private Declare PtrSafe Function clReleaseEvent Lib "opencl" (ByVal event_ As LongPtr) As Long
#Else
    Private Enum LongPtr
        [_]
    End Enum
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Private Declare Function clGetPlatformIDs Lib "opencl" (ByVal num_entries As Long, cl_platform_id As LongPtr, num_platforms As Any) As Long
    Private Declare Function clGetDeviceIDs Lib "opencl" (ByVal platform As LongPtr, ByVal device_type As currency, ByVal num_entries As Long, devices As LongPtr, num_devices As Any) As Long
    Private Declare Function clReleaseDevice Lib "opencl" (ByVal device As LongPtr) As Long
    Private Declare Function clCreateContext Lib "opencl" (ByVal properties As Long, ByVal num_devices As Long, devices As LongPtr, ByVal pfn_notify As Long, ByVal user_data As Long, errcode_ret As Any) As LongPtr
    Private Declare Function clReleaseContext Lib "opencl" (ByVal context As LongPtr) As Long
    Private Declare Function clCreateCommandQueue Lib "opencl" (ByVal context As LongPtr, ByVal device As LongPtr, ByVal properties As currency, errcode_ret As Any) As LongPtr
    Private Declare Function clReleaseCommandQueue Lib "opencl" (ByVal command_queue As LongPtr) As Long
    Private Declare Function clCreateProgramWithSource Lib "opencl" (ByVal context As LongPtr, ByVal count As Long, strings As Any, lengths As Any, errcode_ret As Any) As LongPtr
    Private Declare Function clBuildProgram Lib "opencl" (ByVal program As LongPtr, ByVal num_devices As Long, device_list As Any, ByVal options As String, ByVal pfn_notify As Long, ByVal user_data As Long) As Long
    Private Declare Function clCreateKernel Lib "opencl" (ByVal program As LongPtr, ByVal kernel_name As String, errcode_ret As Any) As LongPtr
    Private Declare Function clReleaseKernel Lib "opencl" (ByVal kernel As LongPtr) As Long
    Private Declare Function clCreateBuffer Lib "opencl" (ByVal context As LongPtr, ByVal flags As currency, ByVal size As Long, ByVal host_ptr As LongPtr, errcode_ret As Any) As LongPtr
    Private Declare Function clReleaseMemObject Lib "opencl" (ByVal memobj As LongPtr) As Long
    Private Declare Function clSetKernelArg Lib "opencl" (ByVal kernel As LongPtr, ByVal arg_index As Long, ByVal arg_size As Long, arg_value As Any) As Long
    Private Declare Function clEnqueueNDRangeKernel Lib "opencl" (ByVal command_queue As LongPtr, ByVal kernel As LongPtr, ByVal work_dim As Long, global_work_offset As Any, global_work_size As Any, local_work_size As Any, ByVal num_events_in_wait_list As Long, event_wait_list As Any, event_ As Any) As Long
    Private Declare Function clFinish Lib "opencl" (ByVal command_queue As LongPtr) As Long
    Private Declare Function clEnqueueMapBuffer Lib "opencl" (ByVal command_queue As LongPtr, ByVal buffer As LongPtr, ByVal blocking_map As Long, ByVal map_flags As currency, ByVal offset As LongPtr, ByVal size As LongPtr, ByVal num_events_in_wait_list As Long, event_wait_list As Any, event_ As Any, errcode_ret As Any) As Long
    Private Declare Function clReleaseEvent Lib "opencl" (ByVal event_ As LongPtr) As Long
#End If




Private Type TThis
    SourceCode as string
    hPlatform as LongPtr
    hDevice as LongPtr
    hContext as LongPtr
    hQueu as LongPtr
    hProgram as LongPtr
    hKernel as LongPtr
    hParams() as LongPtr
End Type
Private This as TThis

Public Function Create(ByVal sEntryPointName as string, ByVal sSrc as string) as stdOpenCL
    Set Create = new stdOpenCL
    Call Create.protInit(sEntryPointName, sSrc)
End Function

Public Sub protInit(ByRef sEntryPointName as string, ByRef sSrc as string)
    This.SourceCode = sSrc
    Dim source() as Byte: source = StrConv(sSrc & vbNullChar, vbFromUnicode)

    Dim lError as Long
    On Error GoTo ErrorHandler

    'Get platform handle
    lError = clGetPlatformIDs(1, This.hPlatform, ByVal 0)
    Call ErrorGuard(lError, "clGetPlatformIDs")

    'Get device handle for platform
    lError = clGetDeviceIDs(This.hPlatform, CL_DEVICE_TYPE_GPU, 1, This.hDevice, ByVal 0)
    Call ErrorGuard(lError, "clGetDeviceIDs")

    'Create context
    This.hContext = clCreateContext(0, 1, This.hDevice, 0, 0, lError)
    Call ErrorGuard(lError, "clCreateContext")

    'Create command queue
    This.hQueue = clCreateCommandQueue(This.hContext, This.hDevice, ByVal 0, lError)
    Call ErrorGuard(lError, "clCreateCommandQueue")

    'Inject source code
    This.hProgram = clCreateProgramWithSource(This.hContext, 1, VarPtr(source(0)), ByVal 0, lError)
    Call ErrorGuard(lError, "clCreateProgramWithSource")

    'Compile and link source to machine code
    lError = clBuildProgram(This.hProgram, 1, This.hDevice, vbNullString, 0, 0)
    Call ErrorGuard(lError, "clBuildProgram")

    'Create kernel for function in code body
    hKernel = clCreateKernel(This.hProgram, sEntryPointName, lError)
    Call ErrorGuard(lError, "clCreateKernel")
    Exit Sub
ErrorHandler:
    Err.raise Err.number, err.source, err.description
End Sub

Public Function Run() as variant

End Function

Public Function Bind(ByRef data as variant, ByVal cArgType as vbVarType) as stdOpenCl

End Function




Private Sub ErrorGuard(ByVal lError as long, ByVal sName as string)
    If lError <> 0 Then
        Err.Raise vbObjectError, sSource, "Error " & lError
    End If
End Sub