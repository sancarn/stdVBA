Attribute VB_Name = "modMyClassDef"
Option Explicit

Private Type tMyCOMcompatibleVTable
  'Space for the 3 Function-Pointers of the IUnknown-Interface
  QueryInterface As Long
  AddRef         As Long
  Release        As Long
  'followed by Space for the single Function-Pointer of our concrete Method
  AddTwoLongs    As Long
End Type

Private mVTable As tMyCOMcompatibleVTable 'preallocated (static, non-Heap) Space for the VTable

Public Function VTablePtr() As Long 'the only Public Function here (later called from modMyClassFactory)
  If mVTable.QueryInterface = 0 Then InitVTable 'initializes only, when not already done
  VTablePtr = VarPtr(mVTable) 'just hand out the Pointer to the statically defined mVTable-Variable
End Function

Private Sub InitVTable() 'this method will be called only once
  mVTable.QueryInterface = FuncPtr(AddressOf modMyClassFactory.QueryInterface)
  mVTable.AddRef = FuncPtr(AddressOf modMyClassFactory.AddRef)
  mVTable.Release = FuncPtr(AddressOf modMyClassFactory.Release)
  
  mVTable.AddTwoLongs = FuncPtr(AddressOf modMyClassFactory.AddTwoLongs)
End Sub
