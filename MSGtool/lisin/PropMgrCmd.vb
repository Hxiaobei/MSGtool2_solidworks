Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst.swMacroFeatureParamType_e
Imports SolidWorks.Interop.swconst.swSpringType_e
Imports SolidWorks.Interop.swconst.swSpringDefineType_e
Imports SolidWorks.Interop.swconst.swSpringProfileType_e
Imports SolidWorks.Interop.swconst.swSketchCheckFeatureProfileUsage_e
Public Class PropMgrCmd

    Private m_Part As ModelDoc2
    Private m_feature As Feature
    Private m_featData As MacroFeatureData
    Private m_modelComp As Component2
    Private m_cmdState As swPageCmdState
    Private m_macroPath As String
    Public m_SpringType As Long
    '0 - compression spring
    '1 - extension spring
    '2 - torsion spring
    '3 - spiral
    Public m_DefineType As Long
    Public m_ProfileType As Long
    '0 - circle profile
    '1 - rectangle profile
    '2 - trapezoid profile
    Private m_Height As Double       'spring height
    Private m_Pitch As Double        'pitch
    Private m_sPitch As Double
    Private m_Revolution As Double   'number of revolutions
    Private m_secparam1 As Double    'used for defining section profile
    Private m_secparam2 As Double
    Private m_secparam3 As Double
    Private m_StartLength As Double  'length of start straight section for torsion spring
    Private m_EndLength As Double    'length of end straight section for torsion spring
    Private m_StartRevolution As Double  'revolution of start section for compression spring
    Private m_EndRevolution As Double    'revolution of end section for compression spring
    Private m_StartPitch As Double   'pitch of start section for compression spring
    Private m_EndPitch As Double     'pitch of end section for compression spring
    Private m_Angle As Double        'taper angle
    Public m_TaperOutward As Long   'for check
    Public m_Direction As Long      'for check
    Public m_RightHand As Long      'for check
    Public m_Ground As Long         'for check
    Public m_ShowTempBody As Long   'for check
    Public m_EndTypeS As Long
    '0-hook, 1-straight, 2-off, start section of torsion spring
    '0-full loop, 1-half loop, 2-user define, for extension spring
    Public m_EndTypeE As Long
    Public m_sketchSel As Object
    Private m_EntSels(13) As Object
    'm_EntSels(0) is base sketch
    'm_EntSels(1) is profile sketch
    'm_EntSels(2) is 3D point used to create sketch plane
    'm_EntSels(3) is sketch plane
    'm_EntSels(4)~(13) are for end path sketch
    Private m_tobe_DelEnt(13) As Feature
    'features to delete if inserting macro feature fails
    'm_tobe_DelEnt(0) is 3D point used to create sketch plane
    'm_tobe_DelEnt(1) is sketch plane
    'm_tobe_DelEnt(2) is profile sketch
    Private m_EntSelMarks(13) As Long
    Public m_SelTypes As Object
    Public m_SelItems_num As Long
    Public m_EditDefinition As Long
    Dim m_Spring As Spring
    Dim m_Utility As Utility
    Dim TempBody(8) As Body2
    Public EnableShowTempBody As Boolean

    Public Sub Init(Part As ModelDoc2, feature As Feature, cmdState As Byte, macroPath As String)
        m_Part = Part
        m_Utility = New Utility
        m_macroPath = macroPath
        If Not IsNothing(feature) Then m_feature = feature

        Dim swModeler As Modeler = Module1.MySwApp.GetModeler
        m_Spring = swModeler.CreateSpring

        If cmdState = swPageCmdState.swCmdEdit Then ' On Edit Definition
            m_SelItems_num = 0

            m_featData = m_feature.GetDefinition
            m_modelComp = m_feature.GetComponent
            m_cmdState = cmdState

            Dim ret As Boolean = m_featData.AccessSelections(m_Part, m_modelComp)

            Dim sels, types, selmarks As Object
            m_featData.GetSelections(sels, types, selmarks)

            m_SelTypes = types
            If Not IsNothing(sels) Then

                For i = 0 To UBound(sels)
                    m_EntSels(i) = sels(i)
                    m_EntSelMarks(i) = selmarks(i)
                    If Not m_EntSels(i) Is Nothing And i = 0 Then m_EntSels(i).Select2(True, selmarks(i))
                    m_SelItems_num = i + 1
                Next i
            End If

            Call m_featData.GetDoubleByName("Height", m_Height)
            Call m_featData.GetDoubleByName("Pitch", m_Pitch)
            Call m_featData.GetDoubleByName("Revolution", m_Revolution)
            Call m_featData.GetDoubleByName("Angle", m_Angle)
            Call m_featData.GetDoubleByName("Secparam1", m_secparam1)
            Call m_featData.GetDoubleByName("Secparam2", m_secparam2)
            Call m_featData.GetDoubleByName("Secparam3", m_secparam3)
            Call m_featData.GetIntegerByName("SpringType", m_SpringType)
            Call m_featData.GetIntegerByName("DefineType", m_DefineType)
            Call m_featData.GetIntegerByName("ProfileType", m_ProfileType)
            Call m_featData.GetIntegerByName("CHKdirection", m_Direction)
            Call m_featData.GetIntegerByName("CHKrighthand", m_RightHand)
            Call m_featData.GetIntegerByName("CHKground", m_Ground)
            Call m_featData.GetIntegerByName("CHKtaperoutward", m_TaperOutward)
            Call m_featData.GetIntegerByName("CHKpreview", m_ShowTempBody)
            Call m_featData.GetDoubleByName("StartLength", m_StartLength)
            Call m_featData.GetDoubleByName("EndLength", m_EndLength)
            Call m_featData.GetDoubleByName("StartRevolution", m_StartRevolution)
            Call m_featData.GetDoubleByName("EndRevolution", m_EndRevolution)
            Call m_featData.GetDoubleByName("StartPitch", m_StartPitch)
            Call m_featData.GetDoubleByName("EndPitch", m_EndPitch)
            Call m_featData.GetIntegerByName("EndTypeS", m_EndTypeS)
            Call m_featData.GetIntegerByName("EndTypeE", m_EndTypeE)

            m_EditDefinition = 1
            EnableShowTempBody = True
            If m_ShowTempBody = 1 Then ShowDifferentParam()

        Else ' On Insert Feature
            m_Part.ClearSelection()
            m_SpringType = swSpringType_Compression
            m_DefineType = swSpringDefineType_PitchAndRevolution
            m_ProfileType = 0
            m_EndTypeS = 0
            m_EndTypeE = 0
            m_Height = 0.03
            m_Pitch = 0.006
            m_Revolution = 5
            m_secparam1 = 0.003
            m_secparam2 = 0.003
            m_secparam3 = 0.002
            m_StartLength = 0.01
            m_EndLength = 0.01
            m_StartPitch = 0.0033
            m_EndPitch = 0.0033
            m_StartRevolution = 1.5
            m_EndRevolution = 1.5
            m_Angle = 0#
            m_RightHand = 0
            m_Direction = 0
            m_TaperOutward = 0
            m_Ground = 0
            m_ShowTempBody = 1
            m_EditDefinition = 0
            EnableShowTempBody = False
        End If
    End Sub

    Public Sub OnOk()
        If m_cmdState = swPageCmdState.swCmdEdit Then ' On Edit Definition
            HideBody()
            Dim sels, types, selmarks As Object
            Call m_featData.GetSelections(sels, types, selmarks)

            Dim newSels(13) As Object
            Dim newSelMarks(13) As Long

            For i = 0 To 13
                newSels(i) = m_EntSels(i)
                newSelMarks(i) = m_EntSelMarks(i)
            Next i

            sels = newSels
            selmarks = newSelMarks

            Call m_featData.Selections((sels), (selmarks))
            Call m_featData.DoubleByName("Height", m_Height)
            Call m_featData.DoubleByName("Pitch", m_Pitch)
            Call m_featData.DoubleByName("Revolution", m_Revolution)
            Call m_featData.DoubleByName("Angle", m_Angle)
            Call m_featData.DoubleByName("Secparam1", m_secparam1)
            Call m_featData.DoubleByName("Secparam2", m_secparam2)
            Call m_featData.DoubleByName("Secparam3", m_secparam3)
            Call m_featData.IntegerByName("SpringType", m_SpringType)
            Call m_featData.IntegerByName("DefineType", m_DefineType)
            Call m_featData.IntegerByName("ProfileType", m_ProfileType)
            Call m_featData.IntegerByName("CHKdirection", m_Direction)
            Call m_featData.IntegerByName("CHKrighthand", m_RightHand)
            Call m_featData.IntegerByName("CHKground", m_Ground)
            Call m_featData.IntegerByName("CHKtaperoutward", m_TaperOutward)
            Call m_featData.IntegerByName("CHKpreview", m_ShowTempBody)
            Call m_featData.DoubleByName("StartLength", m_StartLength)
            Call m_featData.DoubleByName("EndLength", m_EndLength)
            Call m_featData.DoubleByName("StartRevolution", m_StartRevolution)
            Call m_featData.DoubleByName("EndRevolution", m_EndRevolution)
            Call m_featData.DoubleByName("StartPitch", m_StartPitch)
            Call m_featData.DoubleByName("EndPitch", m_EndPitch)
            Call m_featData.IntegerByName("EndTypeS", m_EndTypeS)
            Call m_featData.IntegerByName("EndTypeE", m_EndTypeE)

            Call m_feature.ModifyDefinition(m_featData, m_Part, m_modelComp)

        Else ' On Insert feature
            HideBody()

            Dim erro As Long
            If m_EntSels(1) Is Nothing Then erro = InsertSketch() 'm_EntSel(1) is profile tag

            If erro <> 0 Then GoTo errorend ' Reselect base sketch and profile sketch
            m_Part.ClearSelection()

            For i = 0 To 13
                If Not m_EntSels(i) Is Nothing Then m_EntSels(i).Select2(True, m_EntSelMarks(i))
            Next i

            Dim paramNameArray(22) As String
            Dim paramTypeArray(22) As Long
            Dim paramValueArray(22) As String
            paramNameArray(0) = "Height"
            paramTypeArray(0) = swMacroFeatureParamTypeDouble
            paramValueArray(0) = Str(m_Height)
            paramNameArray(1) = "Pitch"
            paramTypeArray(1) = swMacroFeatureParamTypeDouble
            paramValueArray(1) = Str(m_Pitch)
            paramNameArray(2) = "Revolution"
            paramTypeArray(2) = swMacroFeatureParamTypeDouble
            paramValueArray(2) = Str(m_Revolution)
            paramNameArray(3) = "Angle"
            paramTypeArray(3) = swMacroFeatureParamTypeDouble
            paramValueArray(3) = Str(m_Angle)
            paramNameArray(4) = "Secparam1"
            paramTypeArray(4) = swMacroFeatureParamTypeDouble
            paramValueArray(4) = Str(m_secparam1)
            paramNameArray(5) = "Secparam2"
            paramTypeArray(5) = swMacroFeatureParamTypeDouble
            paramValueArray(5) = Str(m_secparam2)
            paramNameArray(6) = "Secparam3"
            paramTypeArray(6) = swMacroFeatureParamTypeDouble
            paramValueArray(6) = Str(m_secparam3)

            paramNameArray(7) = "SpringType"
            paramTypeArray(7) = swMacroFeatureParamTypeInteger
            paramValueArray(7) = Str(m_SpringType)
            paramNameArray(8) = "DefineType"
            paramTypeArray(8) = swMacroFeatureParamTypeInteger
            paramValueArray(8) = Str(m_DefineType)
            paramNameArray(9) = "ProfileType"
            paramTypeArray(9) = swMacroFeatureParamTypeInteger
            paramValueArray(9) = Str(m_ProfileType)

            paramNameArray(10) = "CHKdirection"
            paramTypeArray(10) = swMacroFeatureParamTypeInteger
            paramValueArray(10) = Str(m_Direction)
            paramNameArray(11) = "CHKrighthand"
            paramTypeArray(11) = swMacroFeatureParamTypeInteger
            paramValueArray(11) = Str(m_RightHand)
            paramNameArray(12) = "CHKground"
            paramTypeArray(12) = swMacroFeatureParamTypeInteger
            paramValueArray(12) = Str(m_Ground)
            paramNameArray(13) = "CHKtaperoutward"
            paramTypeArray(13) = swMacroFeatureParamTypeInteger
            paramValueArray(13) = Str(m_TaperOutward)
            paramNameArray(14) = "CHKpreview"
            paramTypeArray(14) = swMacroFeatureParamTypeInteger
            paramValueArray(14) = Str(m_ShowTempBody)

            paramNameArray(15) = "StartLength"
            paramTypeArray(15) = swMacroFeatureParamTypeDouble
            paramValueArray(15) = Str(m_StartLength)
            paramNameArray(16) = "EndLength"
            paramTypeArray(16) = swMacroFeatureParamTypeDouble
            paramValueArray(16) = Str(m_EndLength)
            paramNameArray(17) = "StartRevolution"
            paramTypeArray(17) = swMacroFeatureParamTypeDouble
            paramValueArray(17) = Str(m_StartRevolution)
            paramNameArray(18) = "EndRevolution"
            paramTypeArray(18) = swMacroFeatureParamTypeDouble
            paramValueArray(18) = Str(m_EndRevolution)
            paramNameArray(19) = "StartPitch"
            paramTypeArray(19) = swMacroFeatureParamTypeDouble
            paramValueArray(19) = Str(m_StartPitch)
            paramNameArray(20) = "EndPitch"
            paramTypeArray(20) = swMacroFeatureParamTypeDouble
            paramValueArray(20) = Str(m_EndPitch)
            paramNameArray(21) = "EndTypeS"
            paramTypeArray(21) = swMacroFeatureParamTypeInteger
            paramValueArray(21) = Str(m_EndTypeS)
            paramNameArray(22) = "EndTypeE"
            paramTypeArray(22) = swMacroFeatureParamTypeInteger
            paramValueArray(22) = Str(m_EndTypeE)

            Dim methodsArray() As String = {m_macroPath, "Macros", "swmMain", m_macroPath, "Macros", "swmPM", "", "", ""}

            Dim featMgr As FeatureManager = m_Part.FeatureManager
            Dim feat As Feature = featMgr.InsertMacroFeature3("Spring", "", methodsArray, paramNameArray, paramTypeArray, paramValueArray, Nothing, Nothing, Nothing, "", 0)

            If feat Is Nothing Then ' Delete 3D point, sketch plane, and profile sketch
                For i = 0 To 13
                    If (i <> 3) Then
                        m_tobe_DelEnt(i).Select2(True, 0)
                    End If
                Next i
                m_Part.DeleteSelection(False)
            Else ' Add 3D point, sketch plane, profile sketch, and end path sketch as subfeature of spring
                m_Part.ClearSelection()
                Dim ret As Boolean
                For i = 0 To 13
                    If Not m_tobe_DelEnt(13 - i) Is Nothing Then
                        ret = feat.MakeSubFeature(m_tobe_DelEnt(13 - i))
                        m_tobe_DelEnt(13 - i).Select2(True, 0)
                        m_Part.BlankRefGeom()
                        m_tobe_DelEnt(13 - i).DeSelect()
                    End If
                Next i
            End If
errorend: End If

    End Sub

    Public Sub OnCancel()
        If m_cmdState = swPageCmdState.swCmdEdit Then m_featData.ReleaseSelectionAccess()

        HideBody()
    End Sub

    ' Insert profile sketch (include a 3D point shetch, a sketch plane,
    ' and a profile sketch) and path sketch
    Public Function InsertSketch() As Long
        InsertSketch = 0
        'Dim sels, types, selmarks As Object

        'Dim skt_profile_sel As Object

        Dim SketchSel As Object = m_EntSels(0)
        m_tobe_DelEnt(3) = SketchSel
        Dim status, opencount, closecount As Long

        SketchSel = SketchSel.GetSpecificFeature()
        status = SketchSel.CheckFeatureUse(swSketchCheckFeature_UNSET, opencount, closecount)
        If opencount <> 0 Then Exit Function
        If closecount <> 1 Then Exit Function

        Dim ArcCount As Long = SketchSel.GetArcCount()
        If ArcCount <> 1 Then Exit Function


        Dim segs As Object = SketchSel.GetSketchSegments
        Dim arc As SketchArc = segs(0)


        Dim finished As Long

        Dim circleBody As Body2 = m_Utility.GetProfileBody(SketchSel)

        m_Spring.BaseProfile = circleBody
        Dim params() As Double = {m_secparam1, m_secparam2, m_secparam3}

        m_Spring.ProfileParameters = ((params))
        m_Spring.ProfileType = m_ProfileType

        Dim pts As Object = m_Spring.GetProfilePoints

        Dim pnt(14) As Double
        For i = 0 To UBound(pts)
            Dim pntVal As Object = pts(i).ArrayData
            pnt(i * 3) = pntVal(0)
            pnt(i * 3 + 1) = pntVal(1)
            pnt(i * 3 + 2) = pntVal(2)
        Next
        Dim ProfilePt As Object = pnt

        Dim selm, sk, sk2 As Object
        selm = m_Part.SelectionManager
        ' Create 3D point
        Dim isAddtoDbOld As Boolean = m_Part.GetAddToDB
        m_Part.AddToDB(True)
        m_Part.ClearSelection()
        m_Part.Insert3DSketch()
        m_Part.CreatePoint2(pnt(0), pnt(1), pnt(2)) 'after created the point is selected by default
        arc.Select2(True, 0) ' select circle sketch segment as constraint condition
        m_Part.SketchAddConstraints("sgCOINCIDENT")
        m_Part.Insert3DSketch()
        m_tobe_DelEnt(0) = selm.GetSelectedObject3(1) ' save feature for deletion (if creation of spring fails)
        ' Create sketch plane
        sk = selm.GetSelectedObject3(1) ' 3D point sketch is only selected entity after creation of spring
        sk2 = sk.GetSpecificFeature
        Dim segments As Object = sk2.GetSketchPoints
        arc.Select2(False, 0) ' reselect the edge
        segments(0).Select2(True, 0) ' select 3D point on sketch
        m_EntSels(2) = segments(0) ' add 3D point into select list array
        m_EntSelMarks(2) = ID_SELECTION_MARK_3
        Dim pln As RefPlane = m_Part.CreatePlanePerCurveAndPassPoint3(1, 1)
        Dim plnpar As Object = pln.GetRefPlaneParams()
        Dim planeFeat As Feature = sk.GetNextFeature
        planeFeat.Select2(False, 0)

        m_tobe_DelEnt(1) = planeFeat ' save feature for delete (if creation of spring fails)
        m_EntSels(3) = planeFeat ' add sketch plane to select list array
        m_EntSelMarks(3) = ID_SELECTION_MARK_4
        ' Create profile
        m_Part.InsertSketch()

        If m_ProfileType = swSpringProfileType_Circle Then
            If CirProfile() = False Then
                InsertSketch = 1
                m_Part.ClearSelection()
                m_Part.InsertSketch()
                Exit Function
            End If
        End If

        If m_ProfileType = swSpringProfileType_Rectangle Then
            If RectProfile(ProfilePt, plnpar) = False Then
                InsertSketch = 1
                m_Part.ClearSelection()
                m_Part.InsertSketch()
                Exit Function
            End If
        End If

        If m_ProfileType = swSpringProfileType_Trapezoid Then
            If TaperProfile(ProfilePt, plnpar) = False Then
                InsertSketch = 1
                m_Part.ClearSelection()
                m_Part.InsertSketch()
                Exit Function
            End If
        End If
        m_Part.ClearSelection()
        m_Part.InsertSketch()

        If m_SpringType <> swSpringType_Extension Then
            m_Part.AddToDB(isAddtoDbOld)
        End If
        '  select mark for profile
        m_EntSels(1) = selm.GetSelectedObject3(1)
        m_EntSelMarks(1) = ID_SELECTION_MARK_2
        m_tobe_DelEnt(2) = selm.GetSelectedObject3(1) ' save feature for deletion (if creation of spring fails)

        ''''''''''''''''''''
        InsertSketch = 0    ' Successful
    End Function

    Public Function CirProfile() As Boolean
        If m_secparam1 > m_Pitch Then
            CirProfile = False
            Exit Function
        End If
        m_Part.CreateCircleByRadius2(0#, 0#, 0#, m_secparam1 / 2.0#)
        CirProfile = True
    End Function

    Public Function RectProfile(pnt As Object, plnpara As Object) As Boolean
        Dim coord(11) As Double
        If m_secparam1 = 0 Or m_secparam2 = 0 Then
            RectProfile = False
            Exit Function
        End If
        If m_SpringType = swSpringType_Spiral Then
            If m_Pitch < m_secparam2 Then
                RectProfile = False
                Exit Function
            End If
        ElseIf m_SpringType = swSpringType_Compression Then
            If m_Pitch < m_secparam1 Then
                RectProfile = False
                Exit Function
            End If
        End If

        m_Utility.Trans_to_SktPoint(pnt, plnpara, coord)

        Dim seg(3) As Object
        seg(0) = m_Part.CreateLine2(coord(0), coord(1), coord(2), coord(3), coord(4), coord(5))
        seg(1) = m_Part.CreateLine2(coord(3), coord(4), coord(5), coord(6), coord(7), coord(8))
        seg(2) = m_Part.CreateLine2(coord(6), coord(7), coord(8), coord(9), coord(10), coord(11))
        seg(3) = m_Part.CreateLine2(coord(9), coord(10), coord(11), coord(0), coord(1), coord(2))
        seg(0).Select2(False, 0)
        seg(2).Select2(True, 0)
        m_Part.SketchAddConstraints("sgPARALLEL")
        seg(1).Select2(False, 0)
        seg(3).Select2(True, 0)
        m_Part.SketchAddConstraints("sgPARALLEL")
        seg(0).Select2(False, 0)
        seg(1).Select2(True, 0)
        m_Part.SketchAddConstraints("sgPERPENDICULAR")
        seg(1).Select2(False, 0)
        seg(2).Select2(True, 0)
        m_Part.SketchAddConstraints("sgPERPENDICULAR")
        RectProfile = True
    End Function

    Public Function TaperProfile(pnt As Object, plnpara As Object) As Boolean
        Dim coord(11) As Double
        If m_secparam3 = 0 And m_secparam2 = 0 Then
            TaperProfile = False
            Exit Function
        End If
        If m_SpringType = swSpringType_Spiral Then
            If m_Pitch < m_secparam1 Then
                TaperProfile = False
                Exit Function
            End If
        ElseIf m_SpringType = swSpringType_Compression Then
            If m_Pitch < m_secparam2 Or m_Pitch < m_secparam3 Then
                TaperProfile = False
                Exit Function
            End If
        End If

        m_Utility.Trans_to_SktPoint(pnt, plnpara, coord)

        Dim seg(3) As Object
        If coord(0) <> coord(3) Or coord(1) <> coord(4) Or coord(2) <> coord(5) Then
            seg(0) = m_Part.CreateLine2(coord(0), coord(1), coord(2), coord(3), coord(4), coord(5))
        End If
        If coord(3) <> coord(6) Or coord(4) <> coord(7) Or coord(5) <> coord(8) Then
            seg(1) = m_Part.CreateLine2(coord(3), coord(4), coord(5), coord(6), coord(7), coord(8))
        End If
        If coord(6) <> coord(9) Or coord(7) <> coord(10) Or coord(8) <> coord(11) Then
            seg(2) = m_Part.CreateLine2(coord(6), coord(7), coord(8), coord(9), coord(10), coord(11))
        End If
        If coord(0) <> coord(9) Or coord(1) <> coord(10) Or coord(2) <> coord(11) Then
            seg(3) = m_Part.CreateLine2(coord(9), coord(10), coord(11), coord(0), coord(1), coord(2))
        End If
        If m_secparam2 <> 0 And m_secparam3 <> 0 Then 'if m_secparam2 or m_secparam3 equal 0,this profile is triangular
            seg(0).Select2(False, 0)
            seg(2).Select2(True, 0)
            m_Part.SketchAddConstraints("sgPARALLEL")
            seg(1).Select2(False, 0)
            seg(3).Select2(True, 0)
            m_Part.SketchAddConstraints("sgSAMELENGTH")
        End If
        TaperProfile = True
    End Function


    Public Function getHeight() As Double
        getHeight = m_Height
    End Function

    Public Sub Height(height As Double)
        m_Height = height
        If m_SpringType = swSpringType_Compression Or m_SpringType = swSpringType_Extension Then
            If m_DefineType = swSpringDefineType_HeightAndRevolution Then
                m_Pitch = m_Height / m_Revolution
            End If
            If m_DefineType = swSpringDefineType_HeightAndPitch Then
                m_Revolution = m_Height / m_Pitch
            End If
        End If
        If m_SpringType = swSpringType_Torsion Then
            If m_DefineType = swSpringDefineType_PitchAndRevolution Then
                If m_ProfileType = swSpringProfileType_Circle Then
                    m_Revolution = m_Height / m_secparam1
                End If
            End If
        End If
    End Sub

    Public Function getPitch() As Double
        getPitch = m_Pitch
    End Function

    Public Sub Pitch(Pitch As Double)
        m_Pitch = Pitch
        If m_SpringType = swSpringType_Compression Or m_SpringType = swSpringType_Extension Then
            If m_DefineType = swSpringDefineType_PitchAndRevolution Then
                m_Height = m_Pitch * m_Revolution
            End If
            If m_DefineType = swSpringDefineType_HeightAndPitch Then
                m_Revolution = m_Height / m_Pitch
            End If
        End If
    End Sub

    Public Function getRevolution() As Double
        getRevolution = m_Revolution
    End Function

    Public Sub Revolution(Revolution As Double)
        m_Revolution = Revolution
        If m_SpringType = swSpringType_Compression Or m_SpringType = swSpringType_Extension Then
            If m_DefineType = swSpringDefineType_PitchAndRevolution Then
                m_Height = m_Pitch * m_Revolution
            End If
            If m_DefineType = swSpringDefineType_HeightAndRevolution Then
                m_Pitch = m_Height / m_Revolution
            End If
        End If
        If m_SpringType = swSpringType_Torsion Then
            If m_DefineType = swSpringDefineType_HeightAndRevolution Then
                If m_ProfileType = swSpringProfileType_Circle Then
                    m_Height = m_Revolution * m_secparam1
                End If
            End If
        End If
    End Sub

    Public Function getStartPitch() As Double
        getStartPitch = m_StartPitch
    End Function

    Public Sub StartPitch(StartPitch As Double)
        m_StartPitch = StartPitch
    End Sub

    Public Function getEndPitch() As Double
        getEndPitch = m_EndPitch
    End Function

    Public Sub EndPitch(EndPitch As Double)
        m_EndPitch = EndPitch
    End Sub

    Public Function getStartRevolution() As Double
        getStartRevolution = m_StartRevolution
    End Function

    Public Sub StartRevolution(StartRevolution As Double)
        m_StartRevolution = StartRevolution
    End Sub

    Public Function getEndRevolution() As Double
        getEndRevolution = m_EndRevolution
    End Function

    Public Sub EndRevolution(EndRevolution As Double)
        m_EndRevolution = EndRevolution
    End Sub

    Public Function getStartLength() As Double
        getStartLength = m_StartLength
    End Function

    Public Sub StartLength(StartLength As Double)
        m_StartLength = StartLength
    End Sub

    Public Function getEndLength() As Double
        getEndLength = m_EndLength
    End Function

    Public Sub EndLength(EndLength As Double)
        m_EndLength = EndLength
    End Sub

    Public Function getAngle() As Double
        getAngle = m_Angle
    End Function

    Public Sub Angle(angle As Double)
        m_Angle = angle
    End Sub

    Public Function getSecparam1() As Double
        getSecparam1 = m_secparam1
    End Function

    Public Sub Secparam1(Secparam1 As Double)
        m_secparam1 = Secparam1
        If m_SpringType = swSpringType_Torsion Then
            If m_ProfileType = swSpringProfileType_Circle Or m_ProfileType = swSpringProfileType_Rectangle Then
                m_Pitch = m_secparam1
                If m_DefineType = swSpringDefineType_PitchAndRevolution Then
                    m_Revolution = m_Height / m_secparam1
                Else
                    m_Height = m_Revolution * m_secparam1
                End If
            End If
        End If
        If m_SpringType = swSpringType_Compression Or m_SpringType = swSpringType_Extension Then
            If m_ProfileType = swSpringProfileType_Circle Or m_ProfileType = swSpringProfileType_Rectangle Then
                If (m_Pitch < m_secparam1) Then
                    m_Pitch = m_secparam1
                    If m_DefineType = swSpringDefineType_PitchAndRevolution Then
                        m_Revolution = m_Height / m_secparam1
                    Else
                        m_Height = m_Revolution * m_secparam1
                    End If
                End If
            End If
        End If
        If m_SpringType = swSpringType_Compression Then
            If (m_StartPitch < m_secparam1) Then m_StartPitch = m_secparam1
            If (m_EndPitch < m_secparam1) Then m_EndPitch = m_secparam1
        End If
    End Sub

    Public Function getSecparam2() As Double
        getSecparam2 = m_secparam2
    End Function

    Public Sub Secparam2(Secparam2 As Double)
        m_secparam2 = Secparam2
        If m_SpringType = swSpringType_Torsion Then
            If m_ProfileType = swSpringProfileType_Trapezoid Then
                m_Pitch = m_secparam2
                If (m_secparam2 < m_secparam3) Then
                    m_Pitch = m_secparam3
                End If
                If m_DefineType = swSpringDefineType_PitchAndRevolution Then
                    m_Revolution = m_Height / m_Pitch
                Else
                    m_Height = m_Revolution * m_Pitch
                End If
            End If
        End If
        If m_SpringType = swSpringType_Compression Or m_SpringType = 1 Then
            If m_ProfileType = swSpringProfileType_Trapezoid Then
                If (m_Pitch < m_secparam2) Then
                    m_Pitch = m_secparam2
                    If (m_Pitch < m_secparam3) Then m_Pitch = m_secparam3
                    If m_DefineType = swSpringDefineType_PitchAndRevolution Then
                        m_Revolution = m_Height / m_Pitch
                    Else
                        m_Height = m_Revolution * m_Pitch
                    End If
                End If
            End If
        End If
        If m_SpringType = swSpringType_Compression Then
            If (m_StartPitch < m_secparam2) Then
                m_StartPitch = m_secparam2
                If (m_StartPitch < m_secparam3) Then m_StartPitch = m_secparam3
            End If
            If (m_EndPitch < m_secparam2) Then
                m_EndPitch = m_secparam2
                If (m_EndPitch < m_secparam3) Then m_EndPitch = m_secparam3
            End If
        End If
    End Sub

    Public Function getSecparam3() As Double
        getSecparam3 = m_secparam3
    End Function

    Public Sub Secparam3(Secparam3 As Double)
        m_secparam3 = Secparam3
        If m_SpringType = swSpringType_Torsion Then
            If m_ProfileType = swSpringProfileType_Trapezoid Then
                m_Pitch = m_secparam2
                If (m_secparam2 < m_secparam3) Then
                    m_Pitch = m_secparam3
                End If
                If m_DefineType = swSpringDefineType_PitchAndRevolution Then
                    m_Revolution = m_Height / m_Pitch
                Else
                    m_Height = m_Revolution * m_Pitch
                End If
            End If
        End If
        If m_SpringType = swSpringType_Compression Or m_SpringType = 1 Then
            If m_ProfileType = swSpringProfileType_Trapezoid Then
                If (m_Pitch < m_secparam3) Then
                    m_Pitch = m_secparam3
                    If (m_Pitch < m_secparam2) Then m_Pitch = m_secparam2
                    If m_DefineType = swSpringDefineType_PitchAndRevolution Then
                        m_Revolution = m_Height / m_Pitch
                    Else
                        m_Height = m_Revolution * m_Pitch
                    End If
                End If
            End If
        End If
        If m_SpringType = swSpringType_Compression Then
            If (m_StartPitch < m_secparam3) Then
                m_StartPitch = m_secparam3
                If (m_StartPitch < m_secparam2) Then m_StartPitch = m_secparam2
            End If
            If (m_EndPitch < m_secparam3) Then
                m_EndPitch = m_secparam3
                If (m_EndPitch < m_secparam2) Then m_EndPitch = m_secparam2
            End If
        End If
    End Sub

    Public Sub OnSketchSelect(ByVal Id As Long)
        If m_EditDefinition = 1 Then Exit Sub

        Dim selm As Object = m_Part.SelectionManager
        Dim SelId As Integer, SelMark As Long
        If Id = ID_SELECTION_SKETCH Then
            SelId = 0
            SelMark = ID_SELECTION_MARK_1
        End If

        '    Dim j As Integer
        '    Dim k As Integer

        For i = 1 To selm.GetSelectedObjectCount
            If SelMark = selm.GetSelectedObjectMark(i) Then
                Dim selObj As Object = selm.GetSelectedObject3(i)
                If SelMark = ID_SELECTION_MARK_1 Then
                    m_EntSels(SelId) = selObj
                    m_EntSelMarks(SelId) = SelMark
                End If
            End If
        Next

        If (m_EntSels(0) IsNot Nothing) Then
            EnableShowTempBody = True
            ShowDifferentParam()
        Else
            EnableShowTempBody = False
        End If

    End Sub
    'called by PropMgrHdlr function such as NumberBoxChange... and selection change
    Public Function ShowDifferentParam()

        If EnableShowTempBody Then
            HideBody()

            If m_ShowTempBody = 1 Then
                DisplayBody()
            End If
        End If

    End Function
    'hide temporary body when Cancel,OK or profile parameters were changed
    Public Function HideBody()

        For i = 0 To 8
            If Not TempBody(i) Is Nothing Then
                TempBody(i).Hide(m_Part)
                TempBody(i) = Nothing
            End If
        Next

    End Function
    'create and display temporary body
    Public Function DisplayBody()
        Dim skt_base_Sel, skt_profile_sel As Object
        'Dim Spring_Height, Spring_Angle As Double
        'Dim Spring_TaperOutward As Long
        Dim params(2) As Double

        If m_EditDefinition = 1 Then 'this is edit definition case
            'get spring parameters from variables defined in this module, these variable was assigned in Init()
            skt_base_Sel = m_EntSels(0)
            skt_base_Sel = skt_base_Sel.GetSpecificFeature()
            m_Spring.BaseProfile = m_Utility.GetProfileBody(skt_base_Sel)
            skt_profile_sel = m_EntSels(1)
            skt_profile_sel = skt_profile_sel.GetSpecificFeature()
            m_Spring.SectionProfile = m_Utility.GetProfileBody(skt_profile_sel)

            m_Spring.Height = m_Height
            m_Spring.TaperAngle = m_Angle
            m_Spring.TaperOutward = m_TaperOutward
            m_Spring.Tolerance = 0.0000001
            m_Spring.Clockwise = m_RightHand
            m_Spring.ReverseDirection = m_Direction
            m_Spring.GroundType = m_Ground
            m_Spring.DefineType = m_DefineType
            m_Spring.Pitch = m_Pitch
            m_Spring.Revolution = m_Revolution
            m_Spring.EndingEndType = m_EndTypeE
            m_Spring.StartingEndType = m_EndTypeS
            m_Spring.StartingEndLength = m_StartLength
            m_Spring.EndingEndLength = m_EndLength
            m_Spring.StartingRevolution = m_StartRevolution
            m_Spring.EndingRevolution = m_EndRevolution
            m_Spring.StartingPitch = m_StartPitch
            m_Spring.EndingPitch = m_EndPitch
            params(0) = m_secparam1
            params(1) = m_secparam2
            params(2) = m_secparam3
            m_Spring.ProfileParameters = (params)
            m_Spring.ProfileType = m_ProfileType
            m_Spring.SpringType = m_SpringType

            'get center point of profile sketch
            Dim SktPnt As SketchPoint
            SktPnt = m_EntSels(2)
            m_Spring.SectionProfileCenter = m_Utility.GetMathPointFromSketchPoint(SktPnt)
            'create temporary spring body
            TempBody(0) = m_Spring.GetSpringBody
            'display
            TempBody(0).Display(m_Part, 255)

        End If

        If m_EditDefinition = 0 Then 'this new feature was created
            'get spring parameters from variables defined in this module, these variable was assigned in Init()
            skt_base_Sel = m_EntSels(0)
            skt_base_Sel = skt_base_Sel.GetSpecificFeature()
            m_Spring.BaseProfile = m_Utility.GetProfileBody(skt_base_Sel)
            m_Spring.TaperAngle = m_Angle
            m_Spring.TaperOutward = m_TaperOutward
            m_Spring.Tolerance = 0.0000001
            m_Spring.Clockwise = m_RightHand
            m_Spring.ReverseDirection = m_Direction
            m_Spring.GroundType = m_Ground
            m_Spring.DefineType = m_DefineType
            m_Spring.Height = m_Height
            m_Spring.Pitch = m_Pitch
            m_Spring.Revolution = m_Revolution
            m_Spring.EndingEndType = m_EndTypeE
            m_Spring.StartingEndType = m_EndTypeS
            m_Spring.StartingEndLength = m_StartLength
            m_Spring.EndingEndLength = m_EndLength
            m_Spring.StartingRevolution = m_StartRevolution
            m_Spring.EndingRevolution = m_EndRevolution
            m_Spring.StartingPitch = m_StartPitch
            m_Spring.EndingPitch = m_EndPitch
            params(0) = m_secparam1
            params(1) = m_secparam2
            params(2) = m_secparam3
            m_Spring.ProfileParameters = (params)
            m_Spring.ProfileType = m_ProfileType
            m_Spring.SpringType = m_SpringType

            ' Create temporary spring body
            TempBody(0) = m_Spring.GetSpringBody
            ' Display
            If Not TempBody(0) Is Nothing Then
                TempBody(0).Display(m_Part, 255)
            End If
        End If
    End Function


End Class
