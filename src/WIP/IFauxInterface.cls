VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IFauxInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = true

'SRC: http://www.vbforums.com/showthread.php?785939-VB6-Creating-Using-COM-Interfaces-without-TLBs

' --------------------------------------------------------------------------------------------------------
' This class is a replication of COM objects and the object returned from this class can
'   be passed to DLLs/COM classes that expect a valid COM interface handle/pointer.
' --------------------------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------------------------
'   How to set up this class...
'    1. Create a new class, copy & paste this code into that new class. Close this class; template only.
'    2. In the Class_Initialize event, define the interface(s) this class will create
'    3. At the bottom of this class, create the interface methods. Pay attention to ByRef/ByVal for parameters
'    4. If this class will contain multiple interfaces, then:
'        In the InitThunks() function, add code as needed, to determine which interfaces will be created
'        In the GetInterfacePointer() & GetInterface() functions, identify which interface is returned
'    5. Add any custom declarations and functions as needed. However, strongly recommend you do not
'        declare any as public. If to be called from outside this class, use Friend vs. Public.
'         CAUTION! If, at any time, you use the keyword Implements in this class, VB appends the methods of
'        whatever you are implementing. VB usually pastes them at the end of the class. Those must be moved
'        from the bottom, to anywhere above the pvIUnknown_QueryInterface() function. See Class_Initialize.
'    6. The pvCallFunction_COM() function is a courtesy. If you will not be calling any methods from interface
'        pointers that this class creates or receives via events, then you can remove it, if desired.
' --------------------------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------------------------
'   How to use this class after set up... In comments below, swap IFauxInterface with new class' name
'   Notes: ALL public functions/subs must be implemented in code objects (form, class, usercontrol, etc)
'       that use the keyword: Implements. Friend functions/subs are not implemented, VB ignores these for
'       Implements. Public Subs can have the same name as a Public Event. Events can't be functions. Add a
'       parameter to act as a function return value, i.e., Event Unloading(ByRef Cancel As Boolean)

'    1) Call the InitThunks() function. This creates the interfaces. Ensure function returns True.

'    2) When an interface is needed, call the GetInterfacePointer() or GetInterface() function

'    3) If this class will be accessed from other code objects via the WithEvents or Implements keywords...
'       For WithEvents:
'            Add all the Public Event EventName(parameters) declarations as needed
'            Ensure you raise those events via RaiseEvent calls in the methods you created at bottom of
'                class, i.e., RaiseEvent EventName(parameters)
'            The code object wanting events will include, in its declarations section, a line of code like:
'                Dim m_MyInterface As IFauxInterface
'            During code object's Load/Initialize event or as needed: Set m_MyInterface = New IFauxInterface
'            Call InitThunks() & remember to set m_MyInterface to nothing when no longer needed
'       For Implements:
'            Add a new variable to this class' declarations section: Private m_Owner As IFauxInterface
'            Add a new Friend property, in this class, to receive the object that is using this class:
'                Friend Property Set Owner(theOwner As IFauxInterface)
'                    Set m_Owner = theOwner
'                End Property
'            Set the Owner property before calling InitThunks()
'            Remember to set m_Owner to Nothing on Class_Terminate
'            Add Public sub/function stubs that will be implemented in those code objects. Stubs have no code.
'            Ensure you call those stubs in the methods you created at bottom of the class, i.e,
'                If Not m_Owner Is Nothing Then Call m_Owner.StubName(parameters)
'            The object being passed to the Owner() property must have declared: Implements IFauxInterface
'               Instantiate an instance of this class and pass it the Owner, i.e.,
'                Set objThisClass.Owner = CodeObjectImplementingThisClass & call InitThunks()
'                If being called from within the owner, then:
'                    In declarations section: Private m_MyInterface As IFauxInterface
'                    Then when needed: Set m_MyInterface.Owner = Me & call InitThunks()
'                    Remember to set instance to nothing when no longer needed

'    4) It is possible to create a hybrid that can be used as a stand-alone instance, WithEvents,
'        and/or Implements or any combination of those three. To do this...
'            Add Public Event declarations as described above
'            Add Public stubs that are named exactly same as the Public Events, exact same parameters
'            Add code in the stubs if this class can be used as a stand-alone instance; else don't add code
'            Create the Friend Property to identify the owner.
'       You will need to devise a way (maybe a property/sub) to know the mode this class is being used for.
'       The code in each of the methods you created at the bottom of this class would have logic like:
'             if stubs have any executable code, prevent its execution if the mode is not As_StandAlone
'            Select Case m_Mode
'                Case As_StandAloneInstance: Call Me.StubName(params)    ' the stub must have executable code
'                Case As_WithEvents: RaiseEvent StubName(params)
'                Case As_Implements: Call m_Owner.StubName(params)
'            End Select
' --------------------------------------------------------------------------------------------------------

Option Explicit
' --------------------------------------------------------------------------------------------------------
'//// Your personal/custom variables, enums, declarations, public events, Implements, etc, go here:

' --------------------------------------------------------------------------------------------------------

' --------------------------------------------------------------------------------------------------------
'//// Optional declarations used for pvCallFunction_COM only. That private function is not
'     required for this class. If you wish to remove it, you can and then you can remove these also:
Private Const IUnk_AddRef As Long = &H4
Private Const IUnk_Release As Long = &H8
Private Declare Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
Private Declare Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As Long, ByVal offsetinVft As Long, _
                                                    ByVal CallConv As Long, ByVal retTYP As Integer, _
                                                    ByVal paCNT As Long, ByRef paTypes As Integer, _
                                                    ByRef paValues As Long, ByRef retVAR As Variant) As Long
' --------------------------------------------------------------------------------------------------------
'//// Required: API declarations needed for this base class
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function CoTaskMemAlloc Lib "ole32.dll" (ByVal cb As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function StringFromGUID2 Lib "ole32.dll" (ByVal rguid As Long, ByVal lpsz As Long, ByVal cchMax As Long) As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long
Private Const E_NOINTERFACE As Long = &H80004002        ' possible return value of QueryInterface
Private Const E_NOTIMPL As Long = &H80004001            ' possible return value of any COM method
Private Const E_POINTER As Long = &H80004003            ' possible return value of any COM method
Private Const S_FALSE As Long = &H1
Private Const E_ABORT As Long = &H80004004
Private Const S_OK As Long = &H0
Private Const FUNK_LEN As Long = &H30&                  ' size of thunk, per method: 48 bytes
Private Const IID_IUnknown As String = "{00000000-0000-0000-C000-000000000046}"
'//// variables required for this class
Private m_VTable As Long, m_vSize As Long               ' memory location of VTable & thunks
Private m_IIDs As Collection                            ' class IID/GUIDs this class replicates
Private m_MethodCounts As Collection                    ' number of interface methods; used once & cleared
Private m_ParamCounts As Collection                     ' method param counts; used once & cleared
Private m_KeyIIDs As Collection                         ' m_IID keys; used once & cleared
Private m_Interface As Collection                       ' created interface pointers
Private m_ActiveInterfaces As Long                      ' active interfaces; default=all
' --------------------------------------------------------------------------------------------------------


' --------------------------------------------------------------------------------------------------------
'//////////////////////////////////////////////////////////////////////////////////////////////////////////
' Start of class-related code follows. None should be modified unless specified elsewhere
' You call this routine to create the needed thunks. The sub has no parameters, but if this class will
' support mulitple interface types and you want user to select which will be created, you should create
' parameters to allow the user to choose. By default, all interfaces defined in Class_Initialize are created
Friend Function InitThunks() As Boolean
    If m_VTable = 0& Then
        ' if you want to choose which to create, then the m_ActiveInterfaces variable must be set.
        ' Use powers of two (bit values) and OR it to the variable for each interface. Example:
        ' Here are the bit values for the first 10 interfaces:
        '    1, 2, 4, 8, 16, 32, 64, 128, 256, 512
        ' Let's say you have 3 interfaces and you want #1 & #3 active, then
        '    m_ActiveInterfaces = 1 Or 4 (1st & 3rd bit values above)

        ' Note: the m_ActiveInterfaces variable can be used in your interface methods, if desired.
        ' Example: If (ActiveInterfaces And 2) = 0 Then have method return E_NOIMPL
        InitThunks = pvCreateVTable()
    End If
End Function
Friend Function GetInterfacePointer() As Long
    ' This function returns just the pointer of a created interface.
    ' The code below returns the first created interface.
    ' If you are supporting multiple interfaces in this class, you will want to modify the
    '    the function to have parameters to choose which interface is to be returned and
    '     then modify the code below to return the correct m_Interface instance.
    '    Tip: the m_Interface collection is keyed with what you supplied to pvAddInterface()
    Dim p As Long
    If m_Interface Is Nothing Then Exit Function
    For p = 1& To m_Interface.Count
        If Not m_Interface.Item(p) = 0& Then
            GetInterfacePointer = m_Interface.Item(p)
            Exit For
        End If
    Next
End Function
Friend Function GetInterface() As Object
    ' This function returns a VB object of the requested interface.
    ' Note: These interfaces likely don't support IDispatch, so IsObject(GetInterface()) will
    '        be false and any methods cannot be directly called by VB on these objects.
    ' If you added GetInterfacePointer parameters, do the same here & include those
    '    parameters in the GetInterfacePointer call below:
    Dim ptr As Long, tUnk As IUnknown
    ptr = GetInterfacePointer()
    If Not ptr = 0& Then
        CopyMemory tUnk, CLng(m_Interface.Item(ptr)), 4&
           Set GetInterface = tUnk
           CopyMemory tUnk, 0&, 4&
    End If
End Function

Private Function pvCreateVTable() As Boolean
    '/// This method must not be modified.  Method is only called once
    '    The thunks for callbacks to this class are created here.
    ' The structure of memory created and used looks like the following:
    '     4 byte entries that become the VTable. There is one entry for every interface
    '        method, including IUnknown methods. And this structure is repeated for
    '         each interface created (if multiple interfaces are supported)
    '     Immediately following this VTable stack is one entry for every thunk.
    '     A thunk is created for every interface method you added to the bottom
    '     of this class, if that interface is created. Each thunk uses 48 bytes.
    '     Each VTable entry points to one of these thunks.
    ' An interface pointer is a memory address that contains just 8 bytes.
    '    The 1st 4 bytes point to the appropriate VTable entry of the interface.
    '    The 2nd 4 bytes is a reference count. See pvIUnknown_AddRef/Release

    Const CB_PAGE_RWX    As Long = &H40         ' Allocate executable memory
    Const CB_MEM_COMMIT  As Long = &H1000       ' Commit allocated memory
   
    Dim lLastMethod As Long, pIndex As Long
    Dim offsetCode As Long, offsetVTable As Long, offsetThunk As Long
    Dim cb() As Long, p As Long, m As Long
   
    If m_IIDs Is Nothing Then Exit Function
    offsetCode = pvFindLastMethod() - 12&    ' get pointer to last private method
    If offsetCode = 0& Then
        Debug.Assert False
        ' if code stopped here, critical error
        ' do you have Public/Friend related code towards bottom of this class?
        End
    End If
   
    For p = 1& To m_MethodCounts.Count
        ' move from bottom to top so we are on the pvIUnknown_QueryInterface method
        offsetCode = offsetCode - m_MethodCounts(p) * 4&
    Next
    If m_ActiveInterfaces = 0& Then m_ActiveInterfaces = -1& ' default: all active
    For p = 0& To m_IIDs.Count - 1&
        If (m_ActiveInterfaces And (2& ^ p)) Then       ' is this interface in use?
            m_vSize = m_vSize + m_MethodCounts(p + 1&)  ' if so, how many methods
            m = m + 1                                    ' number interfaces in use
        End If
    Next
    offsetThunk = m_vSize * 4& + m * 12&                ' add 12 bytes for IUnknown methods
    m_vSize = (m_vSize + 3&) * (FUNK_LEN + 4&)           ' calc size of vtable
    m_VTable = VirtualAlloc(0&, m_vSize, CB_MEM_COMMIT, CB_PAGE_RWX)
   
    ReDim cb(0 To FUNK_LEN \ 4& - 1&)           ' Allocate executable memory
    cb(0) = ObjPtr(Me)                          ' thunks call back here
    cb(2) = &HBB60E089: cb(4) = &HE883C589: cb(5) = &HB9509004: cb(7) = &H74FF06E3
    cb(8) = &HFAE2008D: cb(9) = &H53FF33FF: cb(10) = &HC2906104
   
    Set m_Interface = New Collection
    offsetThunk = offsetThunk + m_VTable
    offsetVTable = m_VTable
    
    ' let's build the 3 IUnknown thunks first & add them to the VTable. All active interfaces share these
    For p = 0& To 2&
        CopyMemory cb(1), ByVal offsetCode, 4&  ' location of callback function in this class
        cb(3) = offsetThunk                     ' location of thunk code
        cb(6) = IIf(p = 0&, 3&, 1&)             ' number parameters for this method
        cb(11) = cb(6) * 4&                     ' Bytes to release on return
        CopyMemory ByVal offsetVTable + p * 4&, offsetThunk + 8&, 4&
        CopyMemory ByVal offsetThunk, cb(0), FUNK_LEN ' copy this thunk
        offsetThunk = offsetThunk + FUNK_LEN
        offsetCode = offsetCode + 4&
    Next
    ' now let's build the methods for active interfaces
    For p = 1& To m_IIDs.Count
        If (m_ActiveInterfaces And (2& ^ (p - 1&))) Then    ' active interface
            m = CoTaskMemAlloc(8&)
            CopyMemory ByVal m, offsetVTable, 4&
            CopyMemory ByVal m + 4&, 0&, 4&
            m_Interface.Add m, m_KeyIIDs(p)

            ' copy IUnknown entries for 2nd and subsequent interfaces
            If offsetVTable > m_VTable Then CopyMemory ByVal offsetVTable, ByVal m_VTable, 12&
            offsetVTable = offsetVTable + 12&

            ' setup the current interface's methods
            For m = 1& To m_MethodCounts(p)
                pIndex = pIndex + 1&                   ' which m_ParamCounts.Item we are on
                CopyMemory cb(1), ByVal offsetCode, 4&  ' location of callback function in this class
                cb(3) = offsetThunk                     ' location of thunk code
                cb(6) = m_ParamCounts(pIndex) + 1&
                cb(11) = cb(6) * 4&                     ' Bytes to release on return
                CopyMemory ByVal offsetVTable, offsetThunk + 8&, 4&
                CopyMemory ByVal offsetThunk, cb(0), FUNK_LEN      ' copy this thunk

                offsetThunk = offsetThunk + FUNK_LEN
                offsetCode = offsetCode + 4&
                offsetVTable = offsetVTable + 4&
            Next
        Else                                            ' skipping these
            m_Interface.Add 0&, m_KeyIIDs(p)
            offsetCode = offsetCode + 4& * m_MethodCounts(p)
            pIndex = pIndex + m_MethodCounts(p)
        End If
    Next
    Set m_MethodCounts = Nothing
    Set m_ParamCounts = Nothing
    Set m_KeyIIDs = Nothing
    pvCreateVTable = True
End Function

Private Function pvFindLastMethod() As Long
    '/// This method must not be modified. Method is only called once
    ' Return the address of the specified ordinal private method, 1 = last private method, 2 = second last private method, etc
  Dim bSub  As Byte                         ' Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                         ' Address of the vTable
  Dim i     As Long                         ' Loop index
  Dim j     As Long                         ' Loop limit
  Dim tblLoc As Long
    CopyMemory nAddr, ByVal ObjPtr(Me), 4&  ' Get the address of this object instance
    If pvProbeClassMethods(nAddr + &H1C, i, bSub) = 0 Then Exit Function ' Probe for a Class method
    i = i + 4&                              ' Bump to the next entry
    j = i + 1024&                           ' Set a reasonable limit, scan 256 vTable entries
    Do While i < j
      CopyMemory nAddr, ByVal i, 4&         ' Get the address stored in this vTable entry
      If IsBadCodePtr(nAddr) Then           ' Is the entry an invalid code address?
        tblLoc = i                          ' Cache the vTable end-point
        GoTo Found                          ' Bad method signature, quit loop
      End If
      CopyMemory bVal, ByVal nAddr, 1&      ' Get the byte pointed to by the vTable entry
      If bVal <> bSub Then                  ' If the byte doesn't match the expected value...
        tblLoc = i                          ' Cache the vTable end-point
        GoTo Found                          ' Bad method signature, quit loop
      End If
      i = i + 4&                            ' Next vTable entry
    Loop
    Exit Function                           ' Final method not found
Found:                                      ' Return the specified vTable entry address
  pvFindLastMethod = tblLoc
End Function

Private Function pvProbeClassMethods(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  '/// This method must not be modified.  Method is only called once
  'Probe at the specified start address for a method signature
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  nAddr = nStart                            ' Start address
  nLimit = nAddr + 32&                      ' Probe eight entries
  Do While nAddr < nLimit                   ' While we've not reached our probe depth
    CopyMemory nEntry, ByVal nAddr, 4&      ' Get the vTable entry
    If nEntry <> 0 Then                     ' If not an implemented interface
      CopyMemory bVal, ByVal nEntry, 1&     ' Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then    ' Check for a native or pcode method signature
        nMethod = nAddr                     ' Store the vTable entry
        bSub = bVal                         ' Store the found method signature
        pvProbeClassMethods = True          ' Indicate success
        Exit Function                       ' Return
      End If
    End If
    nAddr = nAddr + 4&                      ' Next vTable entry
  Loop
End Function

Private Function pvCallFunction_COM(ByVal InterfacePointer As Long, ByVal VTableByteOffset As Long, _
                    ParamArray FunctionParameters() As Variant) As Variant
' Used to call active-x or COM objects, not standard dlls
' Provided as a courtesy. This method is not required for the class

' Return value. Will be a variant containing a value of FunctionReturnType
'   If this method fails, the return value will always be Empty. This can be verified by checking
'       the Err.LastDLLError value. It will be non-zero if the function failed else zero.
'   If the method succeeds, there is no guarantee that the Interface function you called succeeded. The
'       success/failure of that function would be indicated by this method's return value.
'       Typically, success is returned as S_OK (zero) and any other value is an error code.
'   If calling a sub vs function & this method succeeds, the return value will be zero.
'   Summarizing: if method fails to execute, Err.LastDLLError value will be non-zero
'       If method executes ok, if the return value is zero, method succeeded else return is error code

' Parameters:
'   InterfacePointer. A pointer to an object/class, i.e., ObjPtr(IPicture)
'       Passing invalid pointers likely to result in crashes
'   VTableOffset. The offset from the passed InterfacePointer where the virtual function exists.
'       The value is in bytes. These offsets are generally in multiples of 4. Value cannot be negative.
'       Example: to call IUnknown:Release (3rd method), CallFunction_COM InterfacePointer, 8&

    '// minimal sanity check for these 4 parameters:
    If VTableByteOffset < 0& Or InterfacePointer = 0& Then Exit Function

    Const CC_STDCALL As Long = 4&
    Dim pIndex As Long, pCount As Long
    Dim vParamPtr() As Long, vParamType() As Integer
    Dim vRtn As Variant, vParams() As Variant
   
    vParams() = FunctionParameters()                    ' copy passed parameters, if any
    pCount = Abs(UBound(vParams) - LBound(vParams) + 1&)
    If pCount = 0& Then                                 ' no return value (sub vs function)
        ReDim vParamPtr(0 To 0)
        ReDim vParamType(0 To 0)
    Else
        ReDim vParamPtr(0 To pCount - 1&)               ' need matching array of parameter types
        ReDim vParamType(0 To pCount - 1&)              ' and pointers to the parameters
        For pIndex = 0& To pCount - 1&
            vParamPtr(pIndex) = VarPtr(vParams(pIndex))
            vParamType(pIndex) = VarType(vParams(pIndex))
        Next
    End If
    pIndex = DispCallFunc(InterfacePointer, VTableByteOffset, CC_STDCALL, vbLong, _
                          pCount, vParamType(0), vParamPtr(0), vRtn)
    If pIndex = 0& Then                                 ' 0 = S_OK
        pvCallFunction_COM = vRtn                       ' return result
    Else
        SetLastError pIndex                             ' set error & return Empty
    End If
End Function

Private Function pvPointerToObject(ByVal oPtr As Long) As Object
' Courtesy function should you want to convert an interface pointer to an object.
' As an object, when you set it to nothing, VB will call the Release on the pointer
    Dim tUnk As Object
    If Not oPtr = 0& Then
        CopyMemory tUnk, oPtr, 4&
        Set pvPointerToObject = tUnk
        CopyMemory tUnk, 0&, 4&
    End If
End Function

Private Sub pvAddInterface(Name As String, IIDs As String, NumberOfMethods As Long, _
                            ParamArray ParamCountOfEachMethod() As Variant)
    On Error GoTo ExitRoutine
    If Name = vbNullString Then Err.Raise 5     '  can't be blank
    If Len(IIDs) < 38& Then Err.Raise 5         '  not correct
    If NumberOfMethods < 1& Then Err.Raise 5    ' can't be zero
    If IsMissing(ParamCountOfEachMethod) Then Err.Raise 5   ' must be provided
    If Not (UBound(ParamCountOfEachMethod) + 1& = NumberOfMethods) Then Err.Raise 5 ' mismatch
    If m_KeyIIDs Is Nothing Then
        Set m_IIDs = New Collection             ' used in pvIUnknown_QueryInterface
        Set m_KeyIIDs = New Collection          ' destroyed after pvCreateVTable called
        Set m_MethodCounts = New Collection     ' destroyed after pvCreateVTable called
        Set m_ParamCounts = New Collection      ' destroyed after pvCreateVTable called
    ElseIf m_KeyIIDs.Count = 30 Then
         Err.Raise 5                            ' max number of interfaces reached
    End If
    Dim p As Long
    m_KeyIIDs.Add Name
    For p = 0& To NumberOfMethods - 1&
        m_ParamCounts.Add CLng(ParamCountOfEachMethod(p))
    Next
    m_MethodCounts.Add p
    m_IIDs.Add UCase$(IIDs)
ExitRoutine:
    If Err Then
        Debug.Assert False    '  fix the problem now before damage done.  the problem can result in a crash
        Set m_IIDs = Nothing
        Set m_KeyIIDs = Nothing
        Set m_MethodCounts = Nothing
        Set m_ParamCounts = Nothing
    End If
End Sub

Private Sub Class_Terminate()
    ' Append whatever additional class cleanup code you require,
    ' but do not modify the code between the /// markers
    '///
    If Not m_VTable = 0& Then
        Const CB_MEM_DECOMMIT  As Long = &H4000&    ' Decommit allocated memory flag
        Const CB_MEM_RELEASE   As Long = &H8000&    ' Release allocated memory flag
        If Not VirtualFree(m_VTable, m_vSize, CB_MEM_DECOMMIT) = 0& Then
          If Not VirtualFree(m_VTable, 0&, CB_MEM_RELEASE) = 0& Then m_VTable = 0&
        End If
        For m_vSize = m_Interface.Count To 1& Step -1
            ' free up any 8-byte interface objects we created
            If Not m_Interface(m_vSize) = 0& Then CoTaskMemFree m_Interface(m_vSize)
        Next
    End If
    '/// End of IFauxInterface required code
End Sub

Private Sub Class_Initialize()
    '/// Required.

    ' The sample pvAddInterface call below is an example of a valid call to pvAddInterface.
    '    Here are the parameter requirements:
    '    1st param: Interface name. This becomes a Key for each m_Interface collection item; must be unique
    '    2nd param: IID of the interface. Not case-sensitive. Brackets required. Must be 38 chars.
    '        If interface inherits from multiple interfaces, include each IID, example: {IID1}{IID2}{IID3}
    '        Do not include the IUnknown IID. All interfaces support it, so it is handled already in this class
    '        Do not use the same GUID more than once. Doing so can return the wrong ObjPtr when GUID
    '           is requested. If this is a problem, use more than 1 IFauxInterface class and add the
    '           other interface(s) to that other class.
    '    3rd param: number of methods in entire interface, but excluding the 3 IUnknown methods
    '    4th-n params: array of method parameters. For each of the methods in the interface, the number of
    '        parameters the method has. Do not include the mandatory pUnk parameter this class requires.
    '   Notes: All methods must be accounted for, not just ones you plan on using. Interfaces that do not
    '       support IDispatch are called VTable-only interfaces. In order for Windows & other code to call a
    '       function on them, they need to know where the VTable starts. The offset from the VTable for a
    '       specific function is known in advance; hence, you must account for all methods, none skipped.
    ' --------------------------------------------------------------------------------------------------------

    ' pvAddInterface "IFileDialogEvents", "{973510DB-7D7F-452B-8975-74A85828D354}", 7, 1,2,1,1,3,1,3
    '' fyi: 7 methods: 1st has 1 parameter, 2nd has 2, 3rd/4th have 1, 5th has 3, 6th has 1, 7th has 3

    ' --------------------------------------------------------------------------------------------------------
    '    As a real-world example, I have a project that uses 3 interfaces.
    '    Each interface only inherits from IUnknown, so no multiple IIDs used in pvAddInterface.
    '    The code in that class would look like this:
    '    In the Class_Initialize  event...
    '        Call pvAddInterface("IShellItemFilter","{2659B475-EEB8-48B7-8F07-B378810F48CF}",2,1,2)
    '        Call pvAddInterface("IFileDialogControlEvents","{36116642-D713-4b97-9B83-7484A9D00433}",4,2,2,3,2)
    '        Call pvAddInterface("IFileDialogEvents","{973510DB-7D7F-452B-8975-74A85828D354}",7,1,2,1,1,3,1,3)
    '    The interfaces' methods look like this...
    '    Private Function pvIUnknown_QueryInterface(ByVal pUnk As Long, ByVal rIID As Long, ByRef pOut As Long) As Long
    '    Private Function pvIUnknown_AddRef(ByVal pUnk As Long) As Long
    '    Private Function pvIUnknown_Release(ByVal pUnk As Long) As Long
    '    '///// IShellItemFilter - don't include IUnknown
    '    Private Function pvIShellItemFilter_IncludeItem(ByVal pUnk As Long, ...) As Long
    '    Private Function pvIShellItemFilter_GetEnumFlagsForItem(ByVal pUnk As Long, ...) As Long
    '    '///// IFileDialogControlEvents - don't include IUnknown
    '    Private Function pvIFileDialogControlEvents_OnItemSelected(ByVal pUnk As Long, ...) As Long
    '    Private Function pvIFileDialogControlEvents_OnButtonClicked(ByVal pUnk As Long, ...) As Long
    '    Private Function pvIFileDialogControlEvents_OnCheckButtonToggled(ByVal pUnk As Long, ...) As Long
    '    Private Function pvIFileDialogControlEvents_OnControlActivating(ByVal pUnk As Long, ...) As Long
    '    '///// IFileDialogEvents - don't include IUnknown
    '    Private Function pvIFileDialogEvents_OnFileOk(ByVal pUnk As Long, ...) As Long
    '     '    and then the remaining 6 methods
End Function

' ////////////////////////////////////////////////////////////////////////////////////////////////////
'   Interface methods begin here, must all be in VTable order.
'   No additional code can be added after the last interface method you create below.
'    No additional code can be inserted between pvIUnknown_QueryInterface & your last created method.
'     The 1st 3 methods are provided and should not require modifications. You must provide each
'        of the methods, used in each interface, you added via the pvAddInterface() function.
'    Do not create additional IUnknown events. All interfaces will re-route to these three.
' ////////////////////////////////////////////////////////////////////////////////////////////////////
' --------------------------------------------------------------------------------------------------------
Private Function pvIUnknown_QueryInterface(ByVal pUnk As Long, ByVal rIID As Long, ByRef pOut As Long) As Long
    If rIID = 0& Then                           ' required if null pointer received
        pOut = 0&: pvIUnknown_QueryInterface = E_POINTER
        Exit Function
    End If
    Dim sIID As String                          ' get GUID from passed parameter
    sIID = Space$(38)                           ' and convert to String for comparisons
    StringFromGUID2 rIID, StrPtr(sIID), 39&
    If StrComp(sIID, IID_IUnknown, vbTextCompare) = 0& Then
        For pOut = m_IIDs.Count To 1& Step -1&  ' all interfaces implement IUnknown
            pOut = m_Interface(pOut): Exit For  ' but validate passed pUnk is ours else hacking attempt?
        Next
        If pOut Then pOut = m_VTable            ' when IUnknown GUID queried, must return same value everytime
    Else
        For pOut = m_IIDs.Count To 1& Step -1&  ' see if we implement that IID
            If InStr(1, m_IIDs(pOut), sIID, vbTextCompare) > 0& Then
                pOut = m_Interface(pOut): Exit For
            End If
        Next
    End If
    If pOut = 0& Then                           ' if we don't support IID, indicate so
        pvIUnknown_QueryInterface = E_NOINTERFACE
    Else
        Call pvIUnknown_AddRef(pUnk)            ' must AddRef to anything we return
    End If
End Function
Private Function pvIUnknown_AddRef(ByVal pUnk As Long) As Long
  '/// This method must not be modified. Each active interface has its own ref counter
  ' If supporting multiple interfaces, pUnk will be one of those that were created
    CopyMemory pvIUnknown_AddRef, ByVal pUnk + 4&, 4&
    pvIUnknown_AddRef = pvIUnknown_AddRef + 1&
    CopyMemory ByVal pUnk + 4&, pvIUnknown_AddRef, 4&
End Function
Private Function pvIUnknown_Release(ByVal pUnk As Long) As Long
  '/// Each active interface has its own ref counter
    CopyMemory pvIUnknown_Release, ByVal pUnk + 4&, 4&
    pvIUnknown_Release = pvIUnknown_Release - 1&
    CopyMemory ByVal pUnk + 4&, pvIUnknown_Release, 4&
  ' If needed, you can check when the ref count reaches zero & which interface pUnk is.
  ' If pvIUnknown_Release is 0, then all references to that interface have been released.
  ' To determine which interface pUnk refers to, loop thru the m_Interface collection &
  '     compare its value to pUnk
End Function
' ////////////////////////////////////////////////////////////////////////////////////////////////////
' The remaining methods of the interface MUST be added below, in VTable order of the interface and
'    each interface added in same order you used during pvAddInterface() calls in Class_Initialize
'   Each method MUST:
'       - be a function that returns Long. Never a Sub. Always returns Long
'       - be declared as Private
'       - have 1st parameter exactly as: ByVal pUnk As Long
'       - has a function return value appropriate for the method. See MSDN documentation
'             typical return values are S_OK, S_FALSE, E_NOIMPL, E_FAIL
'   For each method you want to forward, be sure to include that in the method, before it returns.
'    Tip: if passed any interface pointers that you must add or remove a reference, you can do that
'        with the courtesy function provided:
'            pvCallFunction_COM thePointer, IUnk_AddRef
'            pvCallFunction_COM thePointer, IUnk_Release
' ////////////////////////////////////////////////////////////////////////////////////////////////////

' ////////////////////////////////////////////////////////////////////////////////////////////////////
'   ABSOLUTELY NO EXECUTABLE CODE, WHATSOEVER, CAN BE PLACED BELOW YOUR LAST INTERFACE METHOD
' ////////////////////////////////////////////////////////////////////////////////////////////////////



