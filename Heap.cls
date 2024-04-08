VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Heap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Data type to hold values and  priorities.
Private Type ValueInfo
    value As String
    priority As String
End Type

' The item array.
Private Items() As ValueInfo
Private NumItems As Integer
Private NumAllocated As Integer

' Push the item in position min down Into the
' tree until its children are larger than it is.
Private Sub HeapPushDown(ByVal min As Integer, _
                                ByVal max As Integer)
  Dim tmp As ValueInfo
  Dim j As Integer
      
   Do
     ' Examine the children of min.
    j = 2 * min
     ' See if we      have dropped off the tree.
     If j <= max Then
       '      We have not. Make j point to the child
       '      with the higher priority.
      
      If j < max Then
         'If Items(j + 1).priority > Items(j).priority Then
         If Items(j + 1).priority < Items(j).priority Then
                 ' The other child has a higher priority.
          j = j + 1
         End If
      End If
       'Compare the min and j entries.
      
      'If Items(j).priority <= Items(min).priority Then
      If Items(j).priority >= Items(min).priority Then
         ' The min item is already larger than
         ' its children. We're done.
         Exit Do
      Else
         ' The min item is smaller than the j item.
        ' Swap them.
        tmp = Items(min)
        Items(min) = Items(j)
        Items(j) = tmp
         ' Continue pushing from this position.
        min = j
      End If
    Else
        Exit Do
     End If
   Loop
      
End Sub

' Push the last item up through the heap
' until it is larger than its children.
Private Sub HeapPushUp(ByVal max As Integer)
  Dim tmp As ValueInfo
  Dim j As Integer
      
   Do
     ' Find the    parent of max.
    j = max \ 2
     ' See if we      have reached the top.
     If j < 1 Then Exit Do
     ' Compare the j      entry to the max entry.
     'If Items(j).priority < Items(max).priority Then
     If Items(j).priority > Items(max).priority Then
       '      The parent is smaller. Swap them.
      tmp = Items(j)
      Items(j) = Items(max)
      Items(max) = tmp
       '      Continue pushing from this position.
      max = j
    Else
        Exit Do
     End If
   Loop
      
End Sub

' Add a value to the heap.
Public Sub Enter(ByVal value As String, _
                        ByVal priority As String)
    ' Make sure there is room      for the new item.
  NumItems = NumItems + 1
   If NumItems > NumAllocated Then
    NumAllocated = NumAllocated + 10
     ReDim Preserve Items(1 To NumAllocated)
   End If
      
   ' Add the new item.
   With Items(NumItems)
    .priority = priority
    .value = value
   End With
 
   ' Make the array a heap      again.
  HeapPushUp NumItems
End Sub

' Remove the highest priority item from the queue.
Public Function Leave(ByRef value As String, _
                        ByRef priority As String) As Boolean
    If NumItems = 0 Then
        Leave = False
        Exit Function
    End If
        
   ' Return the top item.
   With Items(1)
    value = .value
    priority = .priority
   End With
      
   ' Move the last item to the      top.
  Items(1) = Items(NumItems)
  NumItems = NumItems - 1
   ' Push this item down into      the tree until
   ' we have a heap again.
      
  HeapPushDown 1, NumItems
  Leave = True
End Function









