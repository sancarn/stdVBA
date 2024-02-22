VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdUIAutomationElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@inspiration https://www.mrexcel.com/board/threads/custom-accessibility-class.1253735

#If VBA7 = 0 Then
    Enum LongPtr
        [_]
    End Enum
#End If

'@docs https://github.com/tpn/winsdk-10/blob/master/Include/10.0.14393.0/um/UIAutomationClient.idl#L2239-L2494
Private Enum EUIAutomation
  QueryInterface
  AddRef
  Release
  CompareElements
  CompareRuntimeIds 
  GetRootElement 
  ElementFromHandle
  ElementFromPoint
  GetFocusedElement
  GetRootElementBuildCache
  ElementFromHandleBuildCache
  ElementFromPointBuildCache
  GetFocusedElementBuildCache
  CreateTreeWalker
  ControlViewWalker
  ContentViewWalker
  RawViewWalker
  RawViewCondition
  ControlViewCondition
  ContentViewCondition
  CreateCacheRequest
  CreateTrueCondition
  CreateFalseCondition
  CreatePropertyCondition
  CreatePropertyConditionEx
  CreateAndCondition
  CreateAndConditionFromArray
  CreateAndConditionFromNativeArray
  CreateOrCondition
  CreateOrConditionFromArray
  CreateOrConditionFromNativeArray
  CreateNotCondition
  AddAutomationEventHandler
  RemoveAutomationEventHandler
  AddPropertyChangedEventHandlerNativeArray
  AddPropertyChangedEventHandler 
  RemovePropertyChangedEventHandler 
  AddStructureChangedEventHandler 
  RemoveStructureChangedEventHandler 
  AddFocusChangedEventHandler 
  RemoveFocusChangedEventHandler 
  RemoveAllEventHandlers 
  IntNativeArrayToSafeArray 
  IntSafeArrayToNativeArray 
  RectToVariant 
  VariantToRect 
  SafeArrayToRectNativeArray 
  CreateProxyFactoryEntry
  ProxyFactoryMapping
  GetPropertyProgrammaticName
  GetPatternProgrammaticName
  PollForPotentialSupportedPatterns
  PollForPotentialSupportedProperties
  CheckNotSupported
  ReservedNotSupportedValue
  ReservedMixedAttributeValue
  ElementFromIAccessible 
  ElementFromIAccessibleBuildCache 
  'IUIAutomation2 (34723aff-0c9d-49d0-9896-7ab52df8cd8a)
  AutoSetFocus_get
  AutoSetFocus_let
  ConnectionTimeout_get
  ConnectionTimeout_let
  TransactionTimeout_get
  TransactionTimeout_let
  'IUIAutomation3 (73D768DA-9B51-4B89-936E-C209290973E7)
  AddTextEditTextChangedEventHandler
  RemoveTextEditTextChangedEventHandler
  'IUIAutomation4 (1189C02A-05F8-4319-8E21-E817E3DB2860)
  AddChangesEventHandler
  RemoveChangesEventHandler
End Enum
Private Enum EUIAutomationTreeWalker
  QueryInterface
  AddRef
  Release
  GetParentElement
  GetFirstChildElement
  GetLastChildElement
  GetNextSiblingElement
  GetPreviousSiblingElement
  NormalizeElement
  GetParentElementBuildCache
  GetFirstChildElementBuildCache
  GetLastChildElementBuildCache
  GetNextSiblingElementBuildCache
  GetPreviousSiblingElementBuildCache
  NormalizeElementBuildCache
  Condition
End Enum

Private Enum EUIAutomationElement
  QueryInterface
  AddRef
  Release
  SetFocus
  GetRuntimeId
  FindFirst
  FindAll
  FindFirstBuildCache
  FindAllBuildCache
  BuildUpdatedCache
  GetCurrentPropertyValue
  GetCurrentPropertyValueEx
  GetCachedPropertyValue
  GetCachedPropertyValueEx
  GetCurrentPatternAs
  GetCachedPatternAs
  GetCurrentPattern
  GetCachedPattern
  GetCachedParent
  GetCachedChildren
  CurrentProcessId
  CurrentControlType
  CurrentLocalizedControlType
  CurrentName
  CurrentAcceleratorKey
  CurrentAccessKey
  CurrentHasKeyboardFocus
  CurrentIsKeyboardFocusable
  CurrentIsEnabled
  CurrentAutomationId
  CurrentClassName
  CurrentHelpText
  CurrentCulture
  CurrentIsControlElement
  CurrentIsContentElement
  CurrentIsPassword
  CurrentNativeWindowHandle
  CurrentItemType
  CurrentIsOffscreen
  CurrentOrientation
  CurrentFrameworkId
  CurrentIsRequiredForForm
  CurrentItemStatus
  CurrentBoundingRectangle
  CurrentLabeledBy
  CurrentAriaRole
  CurrentAriaProperties
  CurrentIsDataValidForForm
  CurrentControllerFor
  CurrentDescribedBy
  CurrentFlowsTo
  CurrentProviderDescription
  CachedProcessId
  CachedControlType
  CachedLocalizedControlType
  CachedName
  CachedAcceleratorKey
  CachedAccessKey
  CachedHasKeyboardFocus
  CachedIsKeyboardFocusable
  CachedIsEnabled
  CachedAutomationId
  CachedClassName
  CachedHelpText
  CachedCulture
  CachedIsControlElement
  CachedIsContentElement
  CachedIsPassword
  CachedNativeWindowHandle
  CachedItemType
  CachedIsOffscreen
  CachedOrientation
  CachedFrameworkId
  CachedIsRequiredForForm
  CachedItemStatus
  CachedBoundingRectangle
  CachedLabeledBy
  CachedAriaRole
  CachedAriaProperties
  CachedIsDataValidForForm
  CachedControllerFor
  CachedDescribedBy
  CachedFlowsTo
  CachedProviderDescription
  GetClickablePoint
  'IUIAutomationElement2 (6749c683-f70d-4487-a698-5f79d55290d6)
  CurrentOptimizeForVisualContent
  CachedOptimizeForVisualContent
  CurrentLiveSetting
  CachedLiveSetting
  CurrentFlowsFrom
  CachedFlowsFrom
  'IUIAutomationElement3 (8471DF34-AEE0-4A01-A7DE-7DB9AF12C296)
  ShowContextMenu 
  CurrentIsPeripheral 
  CachedIsPeripheral 
  'IUIAutomationElement4 (3B6E233C-52FB-4063-A4C9-77C075C2A06B)
  CurrentPositionInSet
  CurrentSizeOfSet
  CurrentLevel
  CurrentAnnotationTypes
  CurrentAnnotationObjects
  CachedPositionInSet
  CachedSizeOfSet
  CachedLevel
  CachedAnnotationTypes
  CachedAnnotationObjects
  'IUIAutomationElement5 (98141C1D-0D0E-4175-BBE2-6BFF455842A7)
  CurrentLandmarkType
  CurrentLocalizedLandmarkType
  CachedLandmarkType
  CachedLocalizedLandmarkType
  'IUIAutomationElement6 (4780d450-8bca-4977-afa5-a4a517f555e3)
  CurrentFullDescription
  CachedFullDescription
End Enum

'@docs https://github.com/tpn/winsdk-10/blob/master/Include/10.0.14393.0/um/UIAutomationClient.idl#L2116-L2124
Private Enum EUIAutomationElementArray
  QueryInterface
  AddRef
  Release
  Length
  GetElement
End Enum

Private Type TThis
    hElement As LongPtr
End Type
Private This as TThis

Public Function CreateFromHwnd() As stdUIAutomationElement
    Set CreateFromHwnd = New stdUIAutomationElement

End Function




Private vTableCall