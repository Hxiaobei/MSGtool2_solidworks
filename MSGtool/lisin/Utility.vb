Imports SolidWorks.Interop.sldworks

Public Class Utility

    ' Get 3D sketch point in profile sketch
    ' Sketch plane was defined by this point and an edge
    Public Function Get3DSktPoint(ByVal sketch As Object) As Object
        Dim pnt(2) As Double
        Dim SktPnt As SketchPoint = sketch

        pnt(0) = SktPnt.X
        pnt(1) = SktPnt.Y
        pnt(2) = SktPnt.Z

        Get3DSktPoint = pnt
    End Function

    Public Function GetProfileBody(skt_profile_sel As Object) As Body2

        Dim xformv As Object = skt_profile_sel.ModelToSketchXform
        Dim xform As MathTransform = MathUtil.CreateTransform(xformv)
        xform = xform.Inverse
        xformv = xform.ArrayData
        Dim edge As Object = skt_profile_sel.GetContourEdges((xformv))

        If UBound(edge) < 0 Then
            GetProfileBody = Nothing
            Exit Function
        End If

        GetProfileBody = edge(0).GetBody
    End Function
    ' Get 3D sketch point
    ' 3D sketch is a subfeature of wizard hole
    Public Function Get3DSktPnt(CurFeature As Feature) As Object

        Dim sk As Object = CurFeature.GetFirstSubFeature()
        Dim sk2 As Object = sk.GetSpecificFeature
        Dim segs As Object = sk2.GetSketchPoints

        Return segs
    End Function

    ' Transform point from world coordinate system into sketch coordinate system
    Public Sub Trans_to_SktPoint(pnt As Object, plnpara As Object, coord As Object)


        Dim xArray(2), zArray(2) As Double

        For i = 0 To 2
            xArray(i) = plnpara(i + 3)
            zArray(i) = plnpara(i + 6)
        Next i

        Dim xAxis, zAxis, yAxis As MathVector
        xAxis = MathUtil.CreateVector(xArray)
        zAxis = MathUtil.CreateVector(zArray)
        yAxis = zAxis.Cross(xAxis)

        Dim xData, yData, zData As Object
        xData = xAxis.ArrayData
        yData = yAxis.ArrayData
        zData = zAxis.ArrayData

        Dim xformv(15) As Double
        For i = 0 To 2
            xformv(i) = xData(i)
            xformv(i + 3) = yData(i)
            xformv(i + 6) = zData(i)
        Next

        xformv(12) = 1
        xformv(13) = 0
        xformv(14) = 0
        xformv(15) = 0

        Dim xform As MathTransform = MathUtil.CreateTransform(xformv)
        xform = xform.Inverse
        Dim invData As Object = xform.ArrayData
        For i = 0 To 2
            invData(i + 9) = plnpara(i)
        Next
        xform = MathUtil.CreateTransform((invData))


        Dim temp_pt(2) As Double
        temp_pt(0) = pnt(3)
        temp_pt(1) = pnt(4)
        temp_pt(2) = pnt(5)
        Dim mpt As MathPoint = MathUtil.CreatePoint(temp_pt)
        mpt = mpt.MultiplyTransform(xform)
        Dim cpt As Object = mpt.ArrayData
        coord(0) = cpt(0)
        coord(1) = cpt(1)
        coord(2) = cpt(2)

        temp_pt(0) = pnt(6)
        temp_pt(1) = pnt(7)
        temp_pt(2) = pnt(8)
        mpt = MathUtil.CreatePoint(temp_pt)
        mpt = mpt.MultiplyTransform(xform)
        cpt = mpt.ArrayData
        coord(3) = cpt(0)
        coord(4) = cpt(1)
        coord(5) = cpt(2)

        temp_pt(0) = pnt(9)
        temp_pt(1) = pnt(10)
        temp_pt(2) = pnt(11)
        mpt = MathUtil.CreatePoint(temp_pt)
        mpt = mpt.MultiplyTransform(xform)
        cpt = mpt.ArrayData
        coord(6) = cpt(0)
        coord(7) = cpt(1)
        coord(8) = cpt(2)

        temp_pt(0) = pnt(12)
        temp_pt(1) = pnt(13)
        temp_pt(2) = pnt(14)
        mpt = MathUtil.CreatePoint(temp_pt)
        mpt = mpt.MultiplyTransform(xform)
        cpt = mpt.ArrayData

        coord(9) = cpt(0)
        coord(10) = cpt(1)
        coord(11) = cpt(2)

    End Sub

    Public Function GetMathPointFromSketchPoint(SktPnt As SketchPoint) As MathPoint
        Dim pnt(2) As Double
        pnt(0) = SktPnt.X
        pnt(1) = SktPnt.Y
        pnt(2) = SktPnt.Z
        Return MathUtil.CreatePoint((pnt))
    End Function


End Class
