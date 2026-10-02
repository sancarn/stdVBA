VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdQuadTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@module
'@description A strict quadtree implementation in VBA. Use this class to store points in a 2D space
'and query them efficiently.
'@example A simple quadtree example:
'```vb
'Dim q As stdQuadTree: set q = stdQuadTree.Create(0, 0, 100)
'Call q.insertPoint(1, 1, "Point 1")
'Call q.insertPoint(2, 2, "Point 2")
'Call q.insertPoint(3, 3, "Point 3")
'Call q.insertPoint(4, 4, "Point 4")
'Dim points As Collection: Set points = q.queryRangeBounds(0, 0, 8, 8)
'```
'@devNote [Grid instead of quadtree](https://softwareengineering.stackexchange.com/a/364179)
'@devNote [Loose quadtrees](https://stackoverflow.com/a/48384354/6302131)




Private Type XY
  x As Double
  y As Double
  meta As Variant
End Type
Private Type AABB
  center As XY
  halfDimension As Double
End Type

Private Enum BranchIndex
  branchNorthWest
  branchNorthEast
  branchSouthWest
  branchSouthEast
End Enum
Private Type BranchInfo
  children(BranchIndex.branchNorthWest To BranchIndex.branchSouthEast) As stdQuadTree
End Type
Private Type LeafInfo
  points() As XY
  index As Long
End Type
Private Enum QuadTreeType
    QuadTreeBranch
    QuadTreeLeaf
End Enum


'Arbitrary constant to indicate how many elements can be stored in this quad tree node
Private Type TThis
  Type As QuadTreeType
  Branch As BranchInfo
  Leaf As LeafInfo

  'Capacity before subdivision
  capacityBeforeSubdivision As Long
  
  'Min boundary size
  minimumHalfDimension As Double

  'Axis-aligned bounding box stored as a center with half-dimensions to represent the boundaries of this quad tree
  boundary As AABB

  'Parent of this quadtree
  parent As stdQuadTree
End Type
Private This As TThis

'Create quadtree from center coords and halfDimension / radius
'@constructor
'@param cx - The center x coord
'@param cy - The center y coord
'@param halfDimension - the "radius" of the cell
'@param capacityBeforeSubdivision - The number of nodes a quadtree can hold before it subdivides. This can also be seen as the number of duplicate points allowed. If there are duplicates beyond this number you will get a stack overflow error.
'@param minimumHalfDimension - Minimum boundary size when points stop causing subdivisions. Helps handle duplicate points
'@returns - Quadtree instance
'@remark
Public Function Create(ByVal cx As Double, ByVal cy As Double, ByVal halfDimension As Double, Optional ByVal capacityBeforeSubdivision As Long = 4, Optional minimumHalfDimension As Double = 0.00001) As stdQuadTree
  With stdError.getSentry("stdQuadTree#Create", "cx", cx, "cy", cy, "halfDimension", halfDimension, "capacityBeforeSubdivision", capacityBeforeSubdivision, "minimumHalfDimension", minimumHalfDimension)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Set Create = New stdQuadTree
      Call Create.protInit(cx, cy, halfDimension, capacityBeforeSubdivision, minimumHalfDimension)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Creates a quadtree from min/max x/y
'@constructor
'@param minX - Minimum x coord
'@param minY - Minimum y coord
'@param maxX - Maximum x coord
'@param maxY - Maximum y coord
'@param capacityBeforeSubdivision - The number of nodes a quadtree can hold before it subdivides. This can also be seen as the number of duplicate points allowed. If there are duplicates beyond this number you will get a stack overflow error.
'@param minimumHalfDimension - Minimum boundary size when points stop causing subdivisions. Helps handle duplicate points
'@returns - Quadtree instance
Public Function CreateFromBounds(ByVal minX As Double, ByVal minY As Double, ByVal maxX As Double, ByVal maxY As Double, Optional ByVal capacityBeforeSubdivision As Long = 4, Optional minimumHalfDimension As Double = 0.00001) As stdQuadTree
  With stdError.getSentry("stdQuadTree#CreateFromBounds", "minX", minX, "minY", minY, "maxX", maxX, "maxY", maxY, "capacityBeforeSubdivision", capacityBeforeSubdivision, "minimumHalfDimension", minimumHalfDimension)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim rx As Double: rx = (maxX - minX) / 2
      Dim ry As Double: ry = (maxY - minY) / 2
      Dim r_ As Double: r_ = IIf(ry > rx, ry, rx)
      Dim cx As Double: cx = minX + rx
      Dim cy As Double: cy = minY + ry
      
      Set CreateFromBounds = New stdQuadTree
      Call CreateFromBounds.protInit(cx, cy, r_, capacityBeforeSubdivision, minimumHalfDimension)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Creates a quadtree from min/max x/y
'@constructor
'@param minX - Minimum x coord
'@param minY - Minimum y coord
'@param maxX - Maximum x coord
'@param maxY - Maximum y coord
'@param capacityBeforeSubdivision - The number of nodes a quadtree can hold before it subdivides. This can also be seen as the number of duplicate points allowed. If there are duplicates beyond this number you will get a stack overflow error.
'@param minimumHalfDimension - Minimum boundary size when points stop causing subdivisions. Helps handle duplicate points
'@returns - Quadtree instance with parent
Private Function CreateChild(ByVal cx As Double, ByVal cy As Double, ByVal halfDimension As Double, Optional ByVal capacityBeforeSubdivision As Long = 4, Optional minimumHalfDimension As Double = 0.00001) As stdQuadTree
  With stdError.getSentry("stdQuadTree#CreateChild", "cx", cx, "cy", cy, "halfDimension", halfDimension, "capacityBeforeSubdivision", capacityBeforeSubdivision, "minimumHalfDimension", minimumHalfDimension)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Set CreateChild = New stdQuadTree
      Call CreateChild.protInit(cx, cy, halfDimension, capacityBeforeSubdivision, minimumHalfDimension, Me)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Initialise quadtree
'@protected
'@param cx - Center X coord
'@param cy - Center Y coord
'@param halfDimension - radius of square
'@param capacityBeforeSubdivision - The number of nodes a quadtree can hold before it subdivides. This can also be seen as the number of duplicate points allowed. If there are duplicates beyond this number you will get a stack overflow error.
'@param minimumHalfDimension - Minimum boundary size when points stop causing subdivisions. Helps handle duplicate points
'@param parent - The parent of this quadtree element
Public Sub protInit(ByVal cx As Double, ByVal cy As Double, ByVal halfDimension As Double, capacityBeforeSubdivision As Long, minimumHalfDimension As Double, Optional ByVal parent As stdQuadTree = Nothing)
  With stdError.getSentry("stdQuadTree#protInit", "cx", cx, "cy", cy, "halfDimension", halfDimension, "capacityBeforeSubdivision", capacityBeforeSubdivision, "minimumHalfDimension", minimumHalfDimension, "parent", parent)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      With This
        .Type = QuadTreeLeaf
        .capacityBeforeSubdivision = capacityBeforeSubdivision
        .minimumHalfDimension = minimumHalfDimension
        Set .parent = parent
        With .Leaf
          ReDim .points(1 To capacityBeforeSubdivision)
          .index = 0
        End With
        With .boundary
          .center.x = cx
          .center.y = cy
          .halfDimension = halfDimension
        End With
      End With
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub



'@param x - X coord of point
'@param y - Y coord of point
'@param meta - Metadata object i.e. holding additional geospatial information
'@returns - True if inserted into the tree, false if error e.g. out of bounds.
Public Function insertPoint(ByVal x As Double, ByVal y As Double, Optional ByVal meta As Variant) As Boolean
  With stdError.getSentry("stdQuadTree#insertPoint", "x", x, "y", y, "meta", meta)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      'If not within bounds then object cannot be added
      If Not AABB_containsXY(This.boundary, x, y) Then Exit Function
      
      Select Case This.Type
        Case QuadTreeLeaf
          With This.Leaf
            'If there is space within this quadtree and no subdivisions, add point
            If .index < This.capacityBeforeSubdivision Or This.boundary.halfDimension < This.minimumHalfDimension Then
              If This.boundary.halfDimension < This.minimumHalfDimension Then ReDim Preserve .points(1 To .index + 1)
              .index = .index + 1
              With .points(.index)
                .x = x
                .y = y
                Call CopyVariant(.meta, meta)
              End With
              insertPoint = True
              Exit Function
            Else
              Call SubDivide
              GoTo insertQuadTreeBranch
            End If
          End With
        Case QuadTreeBranch
    insertQuadTreeBranch:
          'We have to add the points/data contained in this quad array to the new quads if we only want the last node to hold the data
          Dim i as long
          For i = branchNorthWest To branchSouthEast
            If This.Branch.children(i).insertPoint(x, y, meta) Then
              insertPoint = True
              Exit Function
            End If
          Next
        Case Else
          'Otherwise point cannot be inserted for some unknown reason (this should never happen)
          Debug.Print "stdQuadTree Error: Couldn't add point (" & x & "," & y & ")"
      End Select
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Query for points within a range depicted by center and radius (half dimension)
'@param cx - Center x coordinate
'@param cy - Center y coordinate
'@param radius - minimumDimension / radius of range to search
'@returns Collection<Variant<Array1D<xcoord, ycoord, metadata>>> - The points found within the range
Public Function queryRange(ByVal cx As Double, ByVal cy As Double, ByVal radius As Double) As Collection
  With stdError.getSentry("stdQuadTree#queryRange", "cx", cx, "cy", cy, "radius", radius)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim rect As AABB: rect = AABB_Create(cx, cy, radius)
      Set queryRange = pQueryRange(rect)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Query for points within a rect
'@param minX - Minimum x coord of rect
'@param minY - Minimum y coord of rect
'@param maxX - Maximum x coord of rect 
'@param maxY - Maximum y coord of rect
'@returns Collection<Variant<Array1D<xcoord, ycoord, metadata>>> - The points found within the bounds
Public Function queryRangeBounds(ByVal minX As Double, ByVal minY As Double, ByVal maxX As Double, ByVal maxY As Double) As Collection
  With stdError.getSentry("stdQuadTree#queryRangeBounds", "minX", minX, "minY", minY, "maxX", maxX, "maxY", maxY)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim rx As Double: rx = (maxX - minX) / 2
      Dim ry As Double: ry = (maxY - minY) / 2
      Dim r_ As Double: r_ = IIf(ry > rx, ry, rx)
      Dim cx As Double: cx = minX + rx
      Dim cy As Double: cy = minY + ry
      Dim rect As AABB: rect = AABB_Create(cx, cy, r_)
      Set queryRangeBounds = pQueryRange(rect)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function






'Query for points within some rect
'@param rect - Rect to identify points within
'@returns - Collection of meta objects
Private Function pQueryRange(ByRef rect As AABB) As Collection
  With stdError.getSentry("stdQuadTree#pQueryRange")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Set pQueryRange = New Collection
      If Not AABB_intersects(This.boundary, rect) Then Exit Function
      Select Case This.Type
        Case QuadTreeLeaf
          Dim i As Long: For i = 1 To This.Leaf.index
            If AABB_containsPoint(rect, This.Leaf.points(i)) Then
              With This.Leaf.points(i)
                Call pQueryRange.Add(Array(.x, .y, .meta))
              End With
            End If
          Next
        Case QuadTreeBranch
          For i = branchNorthWest To branchSouthEast
            Call inplaceConcat(pQueryRange, This.Branch.children(i).queryRange(rect.center.x, rect.center.y, rect.halfDimension))
          Next
      End Select
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Create four children that fully divide this quad into four quads of equal area, and move all points to these children
Private Sub SubDivide()
  With stdError.getSentry("stdQuadTree#SubDivide")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      This.Type = QuadTreeBranch
    
      'Create the new quadrants
      Dim r_ As Double: r_ = This.boundary.halfDimension / 2
      With This.boundary.center
        Set This.Branch.children(branchNorthWest) = CreateChild(.x - r_, .y + r_, r_, This.capacityBeforeSubdivision, This.minimumHalfDimension)
        Set This.Branch.children(branchNorthEast) = CreateChild(.x + r_, .y + r_, r_, This.capacityBeforeSubdivision, This.minimumHalfDimension)
        Set This.Branch.children(branchSouthWest) = CreateChild(.x - r_, .y - r_, r_, This.capacityBeforeSubdivision, This.minimumHalfDimension)
        Set This.Branch.children(branchSouthEast) = CreateChild(.x + r_, .y - r_, r_, This.capacityBeforeSubdivision, This.minimumHalfDimension)
      End With
      
      'Insert points into children
      Dim pt As XY, inserted As Boolean
      Dim i As Long
      For i = 1 To This.Leaf.index
        pt = This.Leaf.points(i)
        inserted = False
        Dim j As Long
        For j = branchNorthWest To branchSouthEast
          Dim qt As stdQuadTree: Set qt = This.Branch.children(j)
          'Note: This may recurse down to a lower level
          If qt.insertPoint(pt.x, pt.y, pt.meta) Then
            inserted = True
            Exit For
          End If
        Next
        
        'This should never happen
        If Not inserted Then Err_Raise 1, "stdQuadTree#Subdivide", "QuadTree Error: Could not be inserted?"
      Next
      
      'Free memory
      Erase This.Leaf.points
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub


'Create a point
'@param x - The x point (typically horizontal axis)
'@param y - THe y poiont (typically vertical axis)
'@returns - The point UDT
'@remark it's assumed that x increases from bottom to top (south to north), left to right (west to east), but this shouldn't impact calculations
Private Function XY_Create(ByVal x As Double, ByVal y As Double, Optional ByVal meta As Variant) As XY
  With stdError.getSentry("stdQuadTree#XY_Create", "x", x, "y", y, "meta", meta)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      XY_Create.x = x
      XY_Create.y = y
      Call CopyVariant(XY_Create.meta, meta)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Creates a box
'@param x - center x coord
'@param y - center y coord
'@param halfDimension - radius of box
'@returns - The box UDT
Private Function AABB_Create(ByVal x As Double, ByVal y As Double, ByVal halfDimension As Double) As AABB
  With stdError.getSentry("stdQuadTree#AABB_Create", "x", x, "y", y, "halfDimension", halfDimension)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      AABB_Create.center.x = x
      AABB_Create.center.y = y
      AABB_Create.halfDimension = halfDimension
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Creates a box
'@param center - center point
'@param halfDimension - radius of box
'@returns - The box UDT
Private Function AABB_CreateEx(ByRef center As XY, ByVal halfDimension As Double) As AABB
  With stdError.getSentry("stdQuadTree#AABB_CreateEx", "halfDimension", halfDimension)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      AABB_CreateEx.center.x = center.x
      AABB_CreateEx.center.y = center.y
      AABB_CreateEx.halfDimension = halfDimension
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Checks if a box contains a point
'@param box - The box to check
'@param point - The point to check
'@returns - True if the box contains the point, false otherwise.
Private Function AABB_containsPoint(ByRef box As AABB, ByRef point As XY) As Boolean
  With stdError.getSentry("stdQuadTree#AABB_containsPoint")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      AABB_containsPoint = True
      AABB_containsPoint = AABB_containsPoint And ((box.center.x - box.halfDimension) <= point.x)
      AABB_containsPoint = AABB_containsPoint And ((box.center.x + box.halfDimension) >= point.x)
      AABB_containsPoint = AABB_containsPoint And ((box.center.y - box.halfDimension) <= point.y)
      AABB_containsPoint = AABB_containsPoint And ((box.center.y + box.halfDimension) >= point.y)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Checks if a box contains a XY
'@param box - The box to check
'@param x - The x coord of a point to check
'@param y - The y coord of a point to check
'@returns - True if the box contains the point, false otherwise.
Private Function AABB_containsXY(ByRef box As AABB, ByRef x As Double, ByVal y As Double) As Boolean
  With stdError.getSentry("stdQuadTree#AABB_containsXY", "x", x, "y", y)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      AABB_containsXY = True
      AABB_containsXY = AABB_containsXY And ((box.center.x - box.halfDimension) <= x)
      AABB_containsXY = AABB_containsXY And ((box.center.x + box.halfDimension) >= x)
      AABB_containsXY = AABB_containsXY And ((box.center.y - box.halfDimension) <= y)
      AABB_containsXY = AABB_containsXY And ((box.center.y + box.halfDimension) >= y)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function


'Checks whether a box intersects with another box
'@param box - The box to check
'@param boxCheck - The other potentially intersecting box
'@returns - True if the boxes intersect, false otherwise.
Private Function AABB_intersects(ByRef box As AABB, ByRef boxCheck As AABB) As Boolean
  With stdError.getSentry("stdQuadTree#AABB_intersects")
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim TL As XY: TL = XY_Create(box.center.x - box.halfDimension, box.center.y - box.halfDimension)
      Dim BL As XY: BL = XY_Create(box.center.x + box.halfDimension, box.center.y - box.halfDimension)
      Dim TR As XY: TR = XY_Create(box.center.x - box.halfDimension, box.center.y + box.halfDimension)
      Dim BR As XY: BR = XY_Create(box.center.x + box.halfDimension, box.center.y + box.halfDimension)
      AABB_intersects = True
      AABB_intersects = AABB_intersects And AABB_containsPoint(box, TL)
      AABB_intersects = AABB_intersects And AABB_containsPoint(box, BL)
      AABB_intersects = AABB_intersects And AABB_containsPoint(box, TR)
      AABB_intersects = AABB_intersects And AABB_containsPoint(box, BR)
    
    Exit Function
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Function

'Concatenates values from cToConcat into cOut
'@param cOut - Collection to append items to
'@param cToConcat - Collection to get items from
Private Sub inplaceConcat(ByRef cOut As Collection, ByVal cToConcat As Collection)
  With stdError.getSentry("stdQuadTree#inplaceConcat", "cOut", cOut, "cToConcat", cToConcat)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      Dim v: For Each v In cToConcat
        Call cOut.Add(v)
      Next
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub

'Copy variant from vIn to vOut. Sets objects, lets values
'@param vOut - destination to copy variant to
'@param vIn - source of data to copy
Private Sub CopyVariant(ByRef vOut As Variant, ByVal vIn As Variant)
  With stdError.getSentry("stdQuadTree#CopyVariant", "vOut", vOut, "vIn", vIn)
    On Error GoTo stdErrorWrapper_ErrorOccurred
    
      If IsObject(vIn) Then
        Set vOut = vIn
      Else
        Let vOut = vIn
      End If
    
    Exit Sub
    stdErrorWrapper_ErrorOccurred:
      Call Err_Raise(Err.Number, Err.Source, Err.Description)
  End With
End Sub




Private Sub Err_Raise(ByVal number as Long, Optional ByVal source as string = "", Optional ByVal description as string = "")
  Call stdError.Raise(description)
End Sub
