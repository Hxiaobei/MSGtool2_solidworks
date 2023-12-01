Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports System.Runtime.InteropServices
Imports System

Module CavityTriadManipulatorHandler

    Dim swDragHdlr As swDragManipHdlr

    Sub CTriadManipulator()

        swDragHdlr = New swDragManipHdlr

        Dim swMathUtil As MathUtility = MySwApp.GetMathUtility
        Dim swModel As ModelDoc2 = MySwApp.ActiveDoc
        Dim swSelMgr As SelectionMgr = swModel.SelectionManager
        Dim swFace As Face2 = swSelMgr.GetSelectedObject6(1, -1)

        Dim vPickPt As Object = swSelMgr.GetSelectionPoint2(1, -1)
        Dim swPickPt As MathPoint = swMathUtil.CreatePoint((vPickPt))
        Dim swModViewMgr As ModelViewManager = swModel.ModelViewManager

        Dim swManip As Manipulator = swModViewMgr.CreateManipulator(swManipulatorType_e.swTriadManipulator, swDragHdlr)
        Dim Triad As TriadManipulator = swManip.GetSpecificManipulator()

        Triad.Origin = swPickPt
        'Triad.DoNotShow = (swTriadManipulatorDoNotShow_e.swTriadManipulatorDoNotShowOrigin +'暂时不开启拖动点
        'swTriadManipulatorDoNotShow_e.swTriadManipulatorDoNotShowZAxis +
        'swTriadManipulatorDoNotShow_e.swTriadManipulatorDoNotShowXYRING +
        'swTriadManipulatorDoNotShow_e.swTriadManipulatorDoNotShowYZRING +
        'swTriadManipulatorDoNotShow_e.swTriadManipulatorDoNotShowZXRING +
        'swTriadManipulatorDoNotShow_e.swTriadManipulatorDoNotShowXYPlane +
        'swTriadManipulatorDoNotShow_e.swTriadManipulatorDoNotShowYZPlane +
        'swTriadManipulatorDoNotShow_e.swTriadManipulatorDoNotShowZXPlane)
        Triad.UpdatePosition()
        swManip.Show(swModel)

    End Sub
End Module
