Option Explicit
Dim traverseLevel As Long


Sub main()
    Dim swApp As SldWorks= Application.SldWorks
    Dim myModel AsModelDoc2= swApp.ActiveDoc
    Dim featureMgr AsFeatureManager= myModel.FeatureManager
    Dim rootNode AsTreeControlItem = featureMgr.GetFeatureTreeRootItem()
    If Not IsNothing(rootNode)  Then
        Debug.Print("")
        traverseLevel = 0
        traverse_node(rootNode) 
    End If
End Sub


Private Sub traverse_node(node As TreeControlItem)
    Dim childNode AsTreeControlItem
    Dim featureNode AsFeature
    Dim componentNode AsComponent2
    
    
    Dim restOfString As String
    Dim indent As String
    Dim i As Long
    
    Dim compName As String
    Dim suppr As Long, supprString As String
    Dim vis As Long, visString As String
    Dim fixed As Boolean, fixedString As String
    Dim componentDoc As Object, docString As String
    Dim refConfigName As String


    Dim displayNodeInfo As Boolean = False
    Dim nodeObjectType As Long = node.ObjectType
    Dim nodeObject As Object = node.Object
    
    Select Case nodeObjectType
    Case SwConst.swTreeControlItemType_e.swFeatureManagerItem_Feature:
        displayNodeInfo = True
        If Not nodeObject Is Nothing Then
             featureNode = nodeObject
            restOfString = "[FEATURE: " & featureNode.Name & "]"
        Else
            restOfString = "[FEATURE: object null?]"
        End If
    Case SwConst.swTreeControlItemType_e.swFeatureManagerItem_Component:
        displayNodeInfo = True
        If Not nodeObject Is Nothing Then
             componentNode = nodeObject
            compName = componentNode.Name2
            If (compName = "") Then
                compName = "?"
            End If
            suppr = componentNode.GetSuppression()
            Select Case (suppr)
            Case SwConst.swComponentSuppressionState_e.swComponentFullyResolved
                supprString = "Resolved"
            Case SwConst.swComponentSuppressionState_e.swComponentLightweight
                supprString = "Lightweight"
            Case SwConst.swComponentSuppressionState_e.swComponentSuppressed
                supprString = "Suppressed"
            End Select
            vis = componentNode.Visible
            Select Case (vis)
            Case SwConst.swComponentVisibilityState_e.swComponentHidden
                visString = "Hidden"
            Case SwConst.swComponentVisibilityState_e.swComponentVisible
                visString = "Visible"
            End Select
            fixed = componentNode.IsFixed
            If fixed = 0 Then
                fixedString = "Floating"
            Else
                fixedString = "Fixed"
            End If
             componentDoc = componentNode.GetModelDoc
            If componentDoc Is Nothing Then
                docString = "Not loaded"
            Else
                docString = "Loaded"
            End If
            refConfigName = componentNode.ReferencedConfiguration
            If (refConfigName = "") Then
                refConfigName = "?"
            End If
            restOfString = "[COMPONENT: " & compName & " " & docString & " " & supprString & " " & visString & " " & refConfigName & "]"
        Else
            restOfString = "[COMPONENT: object null?]"
        End If
    Case Else:
        displayNodeInfo = True
        If Not nodeObject Is Nothing Then
            restOfString = "[object type not handled]"
        Else
            restOfString = "[object null?]"
        End If
    End Select
    For i = 1 To traverseLevel
        indent = indent & "  "
    Next i
    If (displayNodeInfo) Then
        Debug.Print indent & node.Text & " : " & restOfString
    End If
    traverseLevel = traverseLevel + 1
     childNode = node.GetFirstChild()
    While Not childNode Is Nothing
        traverse_node childNode
         childNode = childNode.GetNext
    Wend
    traverseLevel = traverseLevel - 1
End Sub



'-----------------------------------------
' Preconditions:
' 1. Open public_documents\samples\tutorial\api\partequations.sldprt.
' 2. Open the Immediate window.
'
' Postconditions:
' 1. Gets each equation's value and index and whether the 
'    equation is a global variable. 
' 2. Examine the Immediate window.
'------------------------------------------
Option Explicit
Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swEqnMgr As SldWorks.EquationMgr
    Dim i As Long
    Dim nCount As Long

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swEqnMgr = swModel.GetEquationMgr
    Debug.Print "File = " & swModel.GetPathName
    nCount = swEqnMgr.GetCount
    For i = 0 To nCount - 1
        Debug.Print "  Equation(" & i & ")  = " & swEqnMgr.Equation(i)
        Debug.Print "    Value = " & swEqnMgr.Value(i)
        Debug.Print "    Index = " & swEqnMgr.Status
        Debug.Print "    Global variable? " & swEqnMgr.GlobalVariable(i)
    Next i
End Sub


This example shows how to get all of the mates (IMate2 and IMateInPlace objects) in an assembly document. 

'----------------------------------------
' Preconditions:
' 1. Open public_documents\samples\tutorial\advdrawings\bladed shaft.sldasm.
' 2. Open the Immediate window.
'
' Postconditions:
' 1. Gets the components and mates.
' 2. Examine the Immediate window.
'-----------------------------------------
Option Explicit

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swComponent As SldWorks.Component2
Dim swAssembly As SldWorks.AssemblyDoc
Dim Components As Variant
Dim SingleComponent As Variant
Dim Mates As Variant
Dim SingleMate As Variant
Dim swMate As SldWorks.Mate2
Dim swMateInPlace As SldWorks.MateInPlace
Dim numMateEntities As Long
Dim typeOfMate As Long
Dim i As Long
Sub main()
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swAssembly = swModel    
    Components = swAssembly.GetComponents(False)
    For Each SingleComponent In Components
        Set swComponent = SingleComponent
        Debug.Print "Name of component: " & swComponent.Name2
        Mates = swComponent.GetMates()
        If (Not IsEmpty(Mates)) Then
            For Each SingleMate In Mates
                If TypeOf SingleMate Is SldWorks.Mate2 Then
                    Set swMate = SingleMate
                    typeOfMate = swMate.Type
                    Select Case typeOfMate
                        Case 0
                            Debug.Print "  Mate type: Coincident"
                        Case 1
                            Debug.Print "  Mate type: Concentric"
                        Case 2
                            Debug.Print "  Mate type: Perpendicular"
                        Case 3
                            Debug.Print "  Mate type: Parallel"
                        Case 4
                            Debug.Print "  Mate type: Tangent"
                        Case 5
                            Debug.Print "  Mate type: Distance"
                        Case 6
                            Debug.Print "  Mate type: Angle"
                        Case 7
                            Debug.Print "  Mate type: Unknown"
                        Case 8
                            Debug.Print "  Mate type: Symmetric"
                        Case 9
                            Debug.Print "  Mate type: CAM follower"
                        Case 10
                            Debug.Print "  Mate type: Gear"
                        Case 11
                            Debug.Print "  Mate type: Width"
                        Case 12
                            Debug.Print "  Mate type: Lock to sketch"
                        Case 13
                            Debug.Print "  Mate type: Rack pinion"
                        Case 14
                            Debug.Print "  Mate type: Max mates"
                        Case 15
                            Debug.Print "  Mate type: Path"
                        Case 16
                            Debug.Print "  Mate type: Lock"
                        Case 17
                            Debug.Print "  Mate type: Screw"
                        Case 18
                            Debug.Print "  Mate type: Linear coupler"
                        Case 19
                            Debug.Print "  Mate type: Universal joint"
                        Case 20
                            Debug.Print "  Mate type: Coordinate"
                        Case 21
                            Debug.Print "  Mate type: Slot"
                        Case 22
                            Debug.Print "  Mate type: Hinge"
                        ' Add new mate types introduced after SOLIDWORKS 2010 FCS here
                    End Select
                End If
                If TypeOf SingleMate Is SldWorks.MateInPlace Then
                    Set swMateInPlace = SingleMate
                    numMateEntities = swMateInPlace.GetMateEntityCount
                    For i = 0 To numMateEntities - 1
                        Debug.Print "  Mate component name: " & swMateInPlace.MateComponentName(i)
                        Debug.Print "    Type of Inplace mate entity: " & swMateInPlace.MateEntityType(i)
                    Next i
                End If
            Next
        End If
        Debug.Print ""
    Next
End Sub
