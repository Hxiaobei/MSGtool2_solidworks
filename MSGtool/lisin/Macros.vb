Imports SolidWorks.Interop.sldworks
Module Macros
#Region "ss"
    Public Enum swPageCmdState
        swCmdCreate = 1
        swCmdEdit = 2
    End Enum

    Public Const ID_GROUP_BASIC_PARAMETER = 1
    Public Const ID_SELECTION_SKETCH = 2
    Public Const ID_LABEL_TYPE = 3
    Public Const ID_COMBO_TYPE = 4
    Public Const ID_LABEL_PARAM_METHOD = 5
    Public Const ID_COMBO_DEFINE_METHOD = 6
    Public Const ID_HEIGHT = 7
    Public Const ID_PITCH = 8
    Public Const ID_REVOLUTION = 9

    Public Const ID_GROUP_SECTION_PROFILE = 10
    Public Const ID_COMBO_SECTION_PROFILE = 11
    Public Const ID_SECTION_PARAMETER1 = 12
    Public Const ID_SECTION_PARAMETER2 = 13
    Public Const ID_SECTION_PARAMETER3 = 14

    Public Const ID_GROUP_END_SPECIFICATION = 15
    Public Const ID_COMBO_END_TYPE = 16
    Public Const ID_LABEL_START_SECTION = 17
    Public Const ID_COMBO_START_ENDTYPE = 18
    Public Const ID_COMPRESSION_START_PITCH = 19
    Public Const ID_COMPRESSION_START_REVOLUTION = 20
    Public Const ID_TORSION_START_LENGTH = 21
    Public Const ID_LABEL_END_SECTION = 22
    Public Const ID_COMBO_END_ENDTYPE = 23
    Public Const ID_COMPRESSION_END_PITCH = 24
    Public Const ID_COMPRESSION_END_REVOLUTION = 25
    Public Const ID_TORSION_END_LENGTH = 26
    Public Const ID_CHECK_GROUND = 27

    Public Const ID_GROUP_ADDITIONAL_OPTION = 28
    Public Const ID_CHECK_DIRECTION = 29
    Public Const ID_CHECK_RIGHT_SPIN = 30
    Public Const ID_CHECK_TEMP_BODY = 31
    Public Const ID_TAPER_ANGLE = 32
    Public Const ID_CHECK_TAPER_OUTWARD = 33

    Public Const ID_SELECTION_MARK_1 = 1
    Public Const ID_SELECTION_MARK_2 = 2
    Public Const ID_SELECTION_MARK_3 = 4
    Public Const ID_SELECTION_MARK_4 = 8
    Public Const ID_SELECTION_MARK_5 = 16
    Public Const ID_SELECTION_MARK_6 = 32
    Public Const ID_SELECTION_MARK_7 = 64
    Public Const ID_SELECTION_MARK_8 = 128
    Public Const ID_SELECTION_MARK_9 = 256
    Public Const ID_SELECTION_MARK_10 = 512
    Public Const ID_SELECTION_MARK_11 = 1024
    Public Const ID_SELECTION_MARK_12 = 2048
    Public Const ID_SELECTION_MARK_13 = 4096
    Public Const ID_SELECTION_MARK_14 = 8192
#End Region
    ' Handle to macro feature regeneration
    Public Function swmMain(featureIn As Feature) As Object

        Dim swModeler As Modeler = MySwApp.GetModeler
        Dim m_Spring As Spring = swModeler.CreateSpring
        Dim featdata As Object = featureIn.GetDefinition
        Dim Utility As New Utility
        Dim sels, types, selmarks As Object

        featdata.GetSelections(sels, types, selmarks)

        If IsNothing(sels) Then Return "Sketch has not been selected!"

        ' Get body from selected circular sketch
        Dim skt_base_Sel As Object = sels(0)
        skt_base_Sel = skt_base_Sel.GetSpecificFeature()
        Dim segs As Object = skt_base_Sel.GetSketchSegments
        If UBound(segs) <> 0 Then Return "Sketch has too many segments!"

        Dim ArcCount As Long = skt_base_Sel.GetArcCount()
        If ArcCount <> 1 Then Return "Sketch has no circular segment!"


        m_Spring.BaseProfile = Utility.GetProfileBody(skt_base_Sel)
        ' Get body from selected profile
        Dim skt_profile_sel As Object = sels(1)
        skt_profile_sel = skt_profile_sel.GetSpecificFeature()
        m_Spring.SectionProfile = Utility.GetProfileBody(skt_profile_sel)
        ' Get spring parameters
        m_Spring.Height = featdata.GetDoubleByName("Height")
        m_Spring.TaperAngle = featdata.GetDoubleByName("Angle")
        m_Spring.TaperOutward = featdata.GetIntegerByName("CHKtaperoutward")
        m_Spring.Tolerance = 0.0000001
        m_Spring.Clockwise = featdata.GetIntegerByName("CHKrighthand")
        m_Spring.ReverseDirection = featdata.GetIntegerByName("CHKdirection")
        m_Spring.GroundType = featdata.GetIntegerByName("CHKground")
        m_Spring.Pitch = featdata.GetDoubleByName("Pitch")
        m_Spring.Revolution = featdata.GetDoubleByName("Revolution")
        m_Spring.EndingEndType = featdata.GetIntegerByName("EndTypeE")
        m_Spring.StartingEndType = featdata.GetIntegerByName("EndTypeS")
        m_Spring.StartingEndLength = featdata.GetDoubleByName("StartLength")
        m_Spring.EndingEndLength = featdata.GetDoubleByName("EndLength")
        m_Spring.StartingRevolution = featdata.GetDoubleByName("StartRevolution")
        m_Spring.EndingRevolution = featdata.GetDoubleByName("EndRevolution")
        m_Spring.StartingPitch = featdata.GetDoubleByName("StartPitch")
        m_Spring.EndingPitch = featdata.GetDoubleByName("EndPitch")
        Dim params() As Double = {featdata.GetDoubleByName("Secparam1"), featdata.GetDoubleByName("Secparam2"), featdata.GetDoubleByName("Secparam3")}
        m_Spring.ProfileParameters = (params)
        m_Spring.DefineType = featdata.GetIntegerByName("DefineType")
        m_Spring.SpringType = featdata.GetIntegerByName("SpringType")
        m_Spring.ProfileType = featdata.GetIntegerByName("ProfileType")

        ' Get center point of profile sketch
        Dim SktPnt As SketchPoint = sels(2)
        m_Spring.SectionProfileCenter = Utility.GetMathPointFromSketchPoint(SktPnt)

        ' Create spring
        Dim springBody As Body2 = m_Spring.GetSpringBody
        If springBody Is Nothing Then Return "An error occured during the creation of the spring's body"
        Return springBody
    End Function

    ' Handle to macro feature edit definition
    Sub swmPM(partIn As ModelDoc2, featureIn As Feature)
        Dim swPage As New PropMgr()
        swPage.Init(partIn, featureIn, swPageCmdState.swCmdEdit, MySwApp.GetCurrentMacroPathName)
        swPage.Show()
    End Sub

    ' Run this procedure to insert macro feature with custom PropertyManager page
    Public Sub swmInsertThreadMacroFeature()
        Dim swPage As New PropMgr()
        Dim Part As ModelDoc2 = MySwApp.ActiveDoc
        swPage.Init(Part, Nothing, swPageCmdState.swCmdCreate, MySwApp.GetCurrentMacroPathName)
        swPage.Show()
    End Sub


End Module
