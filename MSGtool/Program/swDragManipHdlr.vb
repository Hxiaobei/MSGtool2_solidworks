Imports SolidWorks.Interop.swpublished
Imports SolidWorks.Interop.swconst
Imports System.Diagnostics

<System.Runtime.InteropServices.ComVisibleAttribute(True)>
Public Class swDragManipHdlr
    Implements SwManipulatorHandler2


    Private Function SwManipulatorHandler2_OnDelete(ByVal pManipulator As Object) As Boolean Implements SwManipulatorHandler2.OnDelete
        Debug.Print("Manipulator deleted")
    End Function

    Private Sub SwManipulatorHandler2_OnDirectionFlipped(ByVal pManipulator As Object) Implements SwManipulatorHandler2.OnDirectionFlipped
        Debug.Print("Direction flipped")
    End Sub

    Private Function SwManipulatorHandler2_OnDoubleValueChanged(ByVal pManipulator As Object, ByVal handleIndex As Integer, ByRef Value As Double) As Boolean Implements SwManipulatorHandler2.OnDoubleValueChanged
        Debug.Print("Double value changed")
        Debug.Print("  Value = " & Value)
    End Function

    Private Sub SwManipulatorHandler2_OnEndNoDrag(ByVal pManipulator As Object, ByVal handleIndex As Integer) Implements SwManipulatorHandler2.OnEndNoDrag
        Debug.Print("Mouse button released")
    End Sub

    Private Sub SwManipulatorHandler2_OnEndDrag(ByVal pManipulator As Object, ByVal handleIndex As Integer) Implements SwManipulatorHandler2.OnEndDrag
        Debug.Print("Mouse button released after dragging a manipulator handle")
    End Sub

    Private Sub SwManipulatorHandler2_OnHandleRmbSelected(ByVal pManipulator As Object, ByVal handleIndex As Integer) Implements SwManipulatorHandler2.OnHandleRmbSelected
        Debug.Print("Right-mouse button clicked")
        Debug.Print("  HandleIndex = " + handleIndex)
    End Sub

    Private Sub SwManipulatorHandler2_OnHandleSelected(ByVal pManipulator As Object, ByVal handleIndex As Integer) Implements SwManipulatorHandler2.OnHandleSelected
        Debug.Print("Manipulator handle selected")
        'Debug.Print("  HandleIndex = " + handleIndex)
    End Sub

    Private Sub SwManipulatorHandler2_OnItemSetFocus(ByVal pManipulator As Object, ByVal Id As Integer) Implements SwManipulatorHandler2.OnItemSetFocus
        Debug.Print("Focus set on item")
        Debug.Print("  Item ID = " & Id)
    End Sub

    Private Function SwManipulatorHandler2_OnHandleLmbSelected(ByVal pManipulator As Object) As Boolean Implements SwManipulatorHandler2.OnHandleLmbSelected
        Debug.Print("Left-mouse button clicked")
    End Function

    Private Function SwManipulatorHandler2_OnStringValueChanged(ByVal pManipulator As Object, ByVal handleIndex As Integer, ByRef Value As String) As Boolean Implements SwManipulatorHandler2.OnStringValueChanged
        Debug.Print("String value changed")
        Debug.Print("  String value  = " & Value)
    End Function

    Private Sub SwManipulatorHandler2_OnUpdateDrag(ByVal pManipulator As Object, ByVal handleIndex As Integer, ByVal newPosMathPt As Object) Implements SwManipulatorHandler2.OnUpdateDrag
        Debug.Print("Manipulator handle moved while left- or right-mouse button depressed")
        Debug.Print("  HandleIndex = " & handleIndex)
    End Sub

End Class

