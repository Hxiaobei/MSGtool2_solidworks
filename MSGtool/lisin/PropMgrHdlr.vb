
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst.swPropertyManagerPageCloseReasons_e
Imports SolidWorks.Interop.swconst.swSpringProfileType_e
Imports SolidWorks.Interop.swconst.swSpringType_e
Imports SolidWorks.Interop.swconst.swSpringDefineType_e

Public Class PropMgrHdlr

    Dim m_pageObj As PropMgr
    Dim checked As Integer
    Dim m_isOk As Boolean

    Public Sub Init(pageObj As PropMgr)
        m_pageObj = pageObj
    End Sub

    Private Function PropertyManagerPage2Handler2_ConnectToSW(ByVal ThisSW As Object, ByVal Cookie As Long) As Boolean

    End Function

    Private Sub PropertyManagerPage2Handler2_OnButtonPress(ByVal Id As Long)

    End Sub

    Private Sub PropertyManagerPage2Handler2_OnClose(ByVal reason As Long)
        m_isOk = False
        If reason = swPropertyManagerPageClose_Okay Then
            m_isOk = True
        ElseIf reason = swPropertyManagerPageClose_Cancel Then
            m_pageObj.GetCmd().OnCancel()
        End If
    End Sub

    Private Sub PropertyManagerPage2Handler2_OnCheckboxCheck(ByVal Id As Long, ByVal checked As Boolean)
        If Id = ID_CHECK_RIGHT_SPIN Then
            If (checked) Then
                m_pageObj.GetCmd().m_RightHand = 1
            Else : m_pageObj.GetCmd().m_RightHand = 0
            End If
        End If
        If Id = ID_CHECK_DIRECTION Then
            If (checked) Then
                m_pageObj.GetCmd().m_Direction = 1
            Else : m_pageObj.GetCmd().m_Direction = 0
            End If
        End If
        If Id = ID_CHECK_TEMP_BODY Then
            If (checked) Then
                m_pageObj.GetCmd().m_ShowTempBody = 1
                m_pageObj.GetCmd().DisplayBody()
            Else
                m_pageObj.GetCmd().m_ShowTempBody = 0
                m_pageObj.GetCmd().HideBody()
            End If
        End If
        If Id = ID_CHECK_TAPER_OUTWARD Then
            If (checked) Then
                m_pageObj.GetCmd().m_TaperOutward = 1
            Else : m_pageObj.GetCmd().m_TaperOutward = 0
            End If
        End If
        If Id = ID_CHECK_GROUND Then
            If (checked) Then
                m_pageObj.GetCmd().m_Ground = 1
            Else : m_pageObj.GetCmd().m_Ground = 0
            End If
        End If
        m_pageObj.GetCmd().ShowDifferentParam()
    End Sub

    Private Sub PropertyManagerPage2Handler2_OnComboboxSelectionChanged(ByVal Id As Long, ByVal Item As Long)
        If Id = ID_COMBO_TYPE Then
            If m_pageObj.m_Combo_Basic_Type.CurrentSelection = 0 Then
                m_pageObj.GetCmd().m_SpringType = swSpringType_Compression
                m_pageObj.m_Combo_Define_Method.Clear()
                m_pageObj.m_Combo_Define_Method.AddItems("Pitch and Revolution")
                m_pageObj.m_Combo_Define_Method.AddItems("Height and Revolution")
                m_pageObj.m_Combo_Define_Method.AddItems("Height and Pitch")
                m_pageObj.m_Combo_Define_Method.CurrentSelection = 0
                m_pageObj.GetCmd().m_DefineType = swSpringDefineType_PitchAndRevolution
                m_pageObj.m_Height_ctrl.enabled = False
                m_pageObj.m_Pitch_ctrl.enabled = True
                m_pageObj.m_Pitch_ctrl.tip = "Pitch"
                m_pageObj.m_Revolution_ctrl.enabled = True
                m_pageObj.m_Start_Length_ctrl.Visible = False
                m_pageObj.m_End_Length_ctrl.Visible = False
                m_pageObj.m_Start_Pitch_ctrl.Visible = True
                m_pageObj.m_Start_Revolution_ctrl.Visible = True
                m_pageObj.m_End_Pitch_ctrl.Visible = True
                m_pageObj.m_End_Revolution_ctrl.Visible = True
                m_pageObj.m_Group_End_Specification.Visible = True
                m_pageObj.m_Label_Start_Section.Visible = True
                m_pageObj.m_Label_End_Section.Visible = True

                m_pageObj.GetCmd().Height(0.03)
                m_pageObj.m_Height_ctrl.Value = m_pageObj.GetCmd().getHeight()
                m_pageObj.GetCmd().Pitch(0.006)
                m_pageObj.m_Pitch_ctrl.Value = m_pageObj.GetCmd().getPitch()
                m_pageObj.GetCmd().Revolution(5)
                m_pageObj.m_Revolution_ctrl.Value = m_pageObj.GetCmd().getRevolution()
                m_pageObj.GetCmd().Angle(0)
                m_pageObj.m_Angle_ctrl.Value = m_pageObj.GetCmd().getAngle()
                m_pageObj.m_Check_Ground.enabled = True
                m_pageObj.m_Combo_Start_EndType.Visible = False
                m_pageObj.m_Combo_End_EndType.Visible = False
                m_pageObj.m_Angle_ctrl.enabled = True
                If Math.Abs(m_pageObj.GetCmd().getAngle()) < 0.000001 Then
                    m_pageObj.m_Check_Taper_Outward.enabled = False
                Else
                    m_pageObj.m_Check_Taper_Outward.enabled = True
                End If
                m_pageObj.m_Check_Direction.enabled = True
                m_pageObj.m_Check_Ground.Visible = True
                m_pageObj.m_Angle_ctrl.Visible = True
                m_pageObj.m_Check_Taper_Outward.Visible = True
            ElseIf m_pageObj.m_Combo_Basic_Type.CurrentSelection = 1 Then
                m_pageObj.GetCmd().m_SpringType = swSpringType_Extension
                m_pageObj.m_Combo_Define_Method.Clear()
                m_pageObj.m_Combo_Define_Method.AddItems("Pitch and Revolution")
                m_pageObj.m_Combo_Define_Method.AddItems("Height and Revolution")
                m_pageObj.m_Combo_Define_Method.AddItems("Height and Pitch")
                m_pageObj.m_Combo_Define_Method.CurrentSelection = 0
                m_pageObj.GetCmd().m_DefineType = swSpringDefineType_PitchAndRevolution
                m_pageObj.m_Combo_Start_EndType.Clear()
                m_pageObj.m_Combo_Start_EndType.AddItems("Full loop")
                m_pageObj.m_Combo_Start_EndType.AddItems("Half loop")
                m_pageObj.m_Combo_Start_EndType.CurrentSelection = 0
                m_pageObj.GetCmd().m_EndTypeS = 0
                m_pageObj.m_Combo_End_EndType.Clear()
                m_pageObj.m_Combo_End_EndType.AddItems("Full loop")
                m_pageObj.m_Combo_End_EndType.AddItems("Half loop")
                m_pageObj.m_Combo_End_EndType.CurrentSelection = 0
                m_pageObj.GetCmd().m_EndTypeE = 0
                m_pageObj.m_Height_ctrl.enabled = False
                m_pageObj.m_Pitch_ctrl.enabled = True
                m_pageObj.m_Revolution_ctrl.enabled = True
                m_pageObj.m_Start_Length_ctrl.Visible = False
                m_pageObj.m_End_Length_ctrl.Visible = False
                m_pageObj.m_Start_Pitch_ctrl.Visible = False
                m_pageObj.m_Start_Revolution_ctrl.Visible = False
                m_pageObj.m_End_Pitch_ctrl.Visible = False
                m_pageObj.m_End_Revolution_ctrl.Visible = False
                m_pageObj.m_Group_End_Specification.Visible = True
                m_pageObj.m_Label_Start_Section.Visible = True
                m_pageObj.m_Label_End_Section.Visible = True

                m_pageObj.GetCmd().Height(0.03)
                m_pageObj.m_Height_ctrl.Value = m_pageObj.GetCmd().getHeight()
                m_pageObj.GetCmd().Pitch(0.003)
                m_pageObj.m_Pitch_ctrl.Value = m_pageObj.GetCmd().getPitch()
                m_pageObj.GetCmd().Revolution(10)
                m_pageObj.m_Revolution_ctrl.Value = m_pageObj.GetCmd().getRevolution()
                m_pageObj.GetCmd().Angle(0)
                m_pageObj.m_Check_Ground.enabled = False
                m_pageObj.m_Combo_Start_EndType.Visible = True
                m_pageObj.m_Combo_End_EndType.Visible = True
                m_pageObj.m_Angle_ctrl.enabled = True
                If Math.Abs(m_pageObj.GetCmd().getAngle()) < 0.000001 Then
                    m_pageObj.m_Check_Taper_Outward.enabled = False
                Else
                    m_pageObj.m_Check_Taper_Outward.enabled = True
                End If
                m_pageObj.m_Check_Direction.enabled = True
                m_pageObj.m_Check_Ground.Visible = False
                m_pageObj.m_Angle_ctrl.Visible = False
                m_pageObj.m_Check_Taper_Outward.Visible = False
            ElseIf m_pageObj.m_Combo_Basic_Type.CurrentSelection = 2 Then
                m_pageObj.GetCmd().m_SpringType = swSpringType_Torsion
                m_pageObj.m_Combo_Define_Method.Clear()
                m_pageObj.m_Combo_Define_Method.AddItems("Height")
                m_pageObj.m_Combo_Define_Method.AddItems("Revolution")
                m_pageObj.m_Combo_Define_Method.CurrentSelection = 1
                m_pageObj.GetCmd().m_DefineType = swSpringDefineType_HeightAndRevolution
                m_pageObj.m_Combo_Start_EndType.Clear()
                m_pageObj.m_Combo_Start_EndType.AddItems("Hook")
                m_pageObj.m_Combo_Start_EndType.AddItems("Straight")
                m_pageObj.m_Combo_Start_EndType.AddItems("Hinge")
                m_pageObj.m_Combo_Start_EndType.AddItems("Straight off")
                m_pageObj.m_Combo_Start_EndType.CurrentSelection = 0
                m_pageObj.GetCmd().m_EndTypeS = 0
                m_pageObj.m_Combo_End_EndType.Clear()
                m_pageObj.m_Combo_End_EndType.AddItems("Hook")
                m_pageObj.m_Combo_End_EndType.AddItems("Straight")
                m_pageObj.m_Combo_End_EndType.AddItems("Hinge")
                m_pageObj.m_Combo_End_EndType.AddItems("Straight off")
                m_pageObj.m_Combo_End_EndType.CurrentSelection = 0
                m_pageObj.GetCmd().m_EndTypeE = 0

                m_pageObj.m_Height_ctrl.enabled = False
                m_pageObj.m_Pitch_ctrl.enabled = False
                m_pageObj.m_Pitch_ctrl.tip = "For extension spring, Pitch depend on section profile"
                m_pageObj.m_Revolution_ctrl.enabled = True
                m_pageObj.m_Start_Length_ctrl.Visible = True
                m_pageObj.m_End_Length_ctrl.Visible = True
                m_pageObj.m_Start_Pitch_ctrl.Visible = False
                m_pageObj.m_Start_Revolution_ctrl.Visible = False
                m_pageObj.m_End_Pitch_ctrl.Visible = False
                m_pageObj.m_End_Revolution_ctrl.Visible = False
                m_pageObj.m_Group_End_Specification.Visible = True
                m_pageObj.m_Label_Start_Section.Visible = True
                m_pageObj.m_Label_End_Section.Visible = True

                m_pageObj.GetCmd().Height(0.015)
                m_pageObj.m_Height_ctrl.Value = m_pageObj.GetCmd().getHeight()
                m_pageObj.GetCmd().Pitch(0.003)
                m_pageObj.m_Pitch_ctrl.Value = m_pageObj.GetCmd().getPitch()
                m_pageObj.GetCmd().Revolution(5)
                m_pageObj.m_Revolution_ctrl.Value = m_pageObj.GetCmd().getRevolution()
                m_pageObj.GetCmd().Angle(0)
                m_pageObj.m_Check_Ground.enabled = False
                m_pageObj.m_Combo_Start_EndType.Visible = True
                m_pageObj.m_Combo_End_EndType.Visible = True
                m_pageObj.m_Angle_ctrl.enabled = True
                If Math.Abs(m_pageObj.GetCmd().getAngle()) < 0.000001 Then
                    m_pageObj.m_Check_Taper_Outward.enabled = False
                Else
                    m_pageObj.m_Check_Taper_Outward.enabled = True
                End If
                m_pageObj.m_Check_Direction.enabled = True
                m_pageObj.m_Check_Ground.Visible = False
                m_pageObj.m_Angle_ctrl.Visible = False
                m_pageObj.m_Check_Taper_Outward.Visible = False
            Else
                m_pageObj.GetCmd().m_SpringType = swSpringType_Spiral
                m_pageObj.m_Combo_Define_Method.Clear()
                m_pageObj.m_Combo_Define_Method.AddItems("Pitch and Revolution")
                m_pageObj.m_Combo_Define_Method.CurrentSelection = 0
                m_pageObj.GetCmd().m_DefineType = swSpringDefineType_PitchAndRevolution
                m_pageObj.m_Height_ctrl.enabled = False
                m_pageObj.m_Pitch_ctrl.enabled = True
                m_pageObj.m_Pitch_ctrl.tip = "Pitch"
                m_pageObj.m_Revolution_ctrl.enabled = True
                m_pageObj.m_Start_Length_ctrl.Visible = False
                m_pageObj.m_End_Length_ctrl.Visible = False
                m_pageObj.m_Start_Pitch_ctrl.Visible = False
                m_pageObj.m_Start_Revolution_ctrl.Visible = False
                m_pageObj.m_End_Pitch_ctrl.Visible = False
                m_pageObj.m_End_Revolution_ctrl.Visible = False
                m_pageObj.m_Group_End_Specification.Visible = False
                m_pageObj.m_Label_Start_Section.Visible = False
                m_pageObj.m_Label_End_Section.Visible = False

                m_pageObj.GetCmd().Height(0#)
                m_pageObj.m_Height_ctrl.Value = m_pageObj.GetCmd().getHeight()
                m_pageObj.GetCmd().Pitch(0.006)
                m_pageObj.m_Pitch_ctrl.Value = m_pageObj.GetCmd().getPitch()
                m_pageObj.GetCmd().Revolution(5)
                m_pageObj.m_Revolution_ctrl.Value = m_pageObj.GetCmd().getRevolution()
                m_pageObj.GetCmd().Angle(0)
                m_pageObj.m_Check_Ground.enabled = False
                m_pageObj.m_Combo_Start_EndType.Visible = False
                m_pageObj.m_Combo_End_EndType.Visible = False
                m_pageObj.m_Angle_ctrl.enabled = False
                m_pageObj.m_Check_Taper_Outward.enabled = False
                m_pageObj.m_Check_Direction.enabled = False
                m_pageObj.m_Check_Ground.Visible = False
                m_pageObj.m_Angle_ctrl.Visible = False
                m_pageObj.m_Check_Taper_Outward.Visible = False
            End If
            m_pageObj.GetCmd().ShowDifferentParam()
        End If
        If Id = ID_COMBO_DEFINE_METHOD Then
            If m_pageObj.GetCmd().m_SpringType = swSpringType_Compression Or m_pageObj.GetCmd().m_SpringType = swSpringType_Extension Then
                If m_pageObj.m_Combo_Define_Method.CurrentSelection = 0 Then
                    m_pageObj.GetCmd().m_DefineType = swSpringDefineType_PitchAndRevolution
                    m_pageObj.m_Height_ctrl.enabled = False
                    m_pageObj.m_Pitch_ctrl.enabled = True
                    m_pageObj.m_Revolution_ctrl.enabled = True
                ElseIf m_pageObj.m_Combo_Define_Method.CurrentSelection = 1 Then
                    m_pageObj.GetCmd().m_DefineType = swSpringDefineType_HeightAndRevolution
                    m_pageObj.m_Height_ctrl.enabled = True
                    m_pageObj.m_Pitch_ctrl.enabled = False
                    m_pageObj.m_Revolution_ctrl.enabled = True
                ElseIf m_pageObj.m_Combo_Define_Method.CurrentSelection = 2 Then
                    m_pageObj.GetCmd().m_DefineType = swSpringDefineType_HeightAndPitch
                    m_pageObj.m_Height_ctrl.enabled = True
                    m_pageObj.m_Pitch_ctrl.enabled = True
                    m_pageObj.m_Revolution_ctrl.enabled = False
                End If
            End If
            If m_pageObj.GetCmd().m_SpringType = swSpringType_Torsion Then
                If m_pageObj.m_Combo_Define_Method.CurrentSelection = 0 Then
                    m_pageObj.GetCmd().m_DefineType = swSpringDefineType_PitchAndRevolution
                    m_pageObj.m_Height_ctrl.enabled = True
                    m_pageObj.m_Pitch_ctrl.enabled = False
                    m_pageObj.m_Revolution_ctrl.enabled = False
                ElseIf m_pageObj.m_Combo_Define_Method.CurrentSelection = 1 Then
                    m_pageObj.GetCmd().m_DefineType = swSpringDefineType_HeightAndRevolution
                    m_pageObj.m_Height_ctrl.enabled = False
                    m_pageObj.m_Pitch_ctrl.enabled = False
                    m_pageObj.m_Revolution_ctrl.enabled = True
                End If
            End If
            m_pageObj.GetCmd().ShowDifferentParam()
        End If
        If Id = ID_COMBO_SECTION_PROFILE Then
            If m_pageObj.m_Combo_Section_Profile.CurrentSelection = 0 Then
                m_pageObj.GetCmd().m_ProfileType = swSpringProfileType_Circle
                m_pageObj.m_secparam1_ctrl.enabled = True
                m_pageObj.m_secparam1_ctrl.tip = "Diameter"
                m_pageObj.m_secparam2_ctrl.enabled = False
                m_pageObj.m_secparam2_ctrl.tip = ""
                m_pageObj.m_secparam3_ctrl.enabled = False
                m_pageObj.m_secparam3_ctrl.tip = ""
            ElseIf m_pageObj.m_Combo_Section_Profile.CurrentSelection = 1 Then
                m_pageObj.GetCmd().m_ProfileType = swSpringProfileType_Rectangle
                m_pageObj.m_secparam1_ctrl.enabled = True
                m_pageObj.m_secparam1_ctrl.tip = "Width"
                m_pageObj.m_secparam2_ctrl.enabled = True
                m_pageObj.m_secparam2_ctrl.tip = "Height"
                m_pageObj.m_secparam3_ctrl.enabled = False
                m_pageObj.m_secparam3_ctrl.tip = ""
            ElseIf m_pageObj.m_Combo_Section_Profile.CurrentSelection = 2 Then
                m_pageObj.GetCmd().m_ProfileType = swSpringProfileType_Trapezoid
                m_pageObj.m_secparam1_ctrl.enabled = True
                m_pageObj.m_secparam1_ctrl.tip = "Height"
                m_pageObj.m_secparam2_ctrl.enabled = True
                m_pageObj.m_secparam2_ctrl.tip = "Bottom width"
                m_pageObj.m_secparam3_ctrl.enabled = True
                m_pageObj.m_secparam3_ctrl.tip = "Top width"
            End If
            m_pageObj.GetCmd().ShowDifferentParam()
        End If
        If Id = ID_COMBO_START_ENDTYPE Then
            If m_pageObj.m_Combo_Start_EndType.CurrentSelection = 0 Then
                m_pageObj.GetCmd().m_EndTypeS = 0
            ElseIf m_pageObj.m_Combo_Start_EndType.CurrentSelection = 1 Then
                m_pageObj.GetCmd().m_EndTypeS = 1
            ElseIf m_pageObj.m_Combo_Start_EndType.CurrentSelection = 2 Then
                m_pageObj.GetCmd().m_EndTypeS = 2
            ElseIf m_pageObj.m_Combo_Start_EndType.CurrentSelection = 3 Then
                m_pageObj.GetCmd().m_EndTypeS = 3
            End If
            m_pageObj.GetCmd().ShowDifferentParam()
        End If
        If Id = ID_COMBO_END_ENDTYPE Then
            If m_pageObj.m_Combo_End_EndType.CurrentSelection = 0 Then
                m_pageObj.GetCmd().m_EndTypeE = 0
            ElseIf m_pageObj.m_Combo_End_EndType.CurrentSelection = 1 Then
                m_pageObj.GetCmd().m_EndTypeE = 1
            ElseIf m_pageObj.m_Combo_End_EndType.CurrentSelection = 2 Then
                m_pageObj.GetCmd().m_EndTypeE = 2
            ElseIf m_pageObj.m_Combo_End_EndType.CurrentSelection = 3 Then
                m_pageObj.GetCmd().m_EndTypeE = 3
            End If
            m_pageObj.GetCmd().ShowDifferentParam()
        End If

    End Sub

    Private Sub PropertyManagerPage2Handler2_OnGroupCheck(ByVal Id As Long, ByVal checked As Boolean)

    End Sub

    Private Sub PropertyManagerPage2Handler2_OnGroupExpand(ByVal Id As Long, ByVal Expanded As Boolean)

    End Sub

    Private Function PropertyManagerPage2Handler2_OnHelp() As Boolean

    End Function

    Private Sub PropertyManagerPage2Handler2_OnListboxSelectionChanged(ByVal Id As Long, ByVal Item As Long)

    End Sub

    Private Sub PropertyManagerPage2Handler2_OnNumberboxChanged(ByVal Id As Long, ByVal Value As Double)
        If Id = ID_HEIGHT Then
            m_pageObj.GetCmd().Height(Value)
            m_pageObj.m_Pitch_ctrl.Value = m_pageObj.GetCmd().getPitch()
            m_pageObj.m_Height_ctrl.Value = m_pageObj.GetCmd().getHeight()
            m_pageObj.m_Revolution_ctrl.Value = m_pageObj.GetCmd().getRevolution()
        ElseIf Id = ID_PITCH Then
            m_pageObj.GetCmd().Pitch(Value)
            m_pageObj.m_Pitch_ctrl.Value = m_pageObj.GetCmd().getPitch()
            m_pageObj.m_Height_ctrl.Value = m_pageObj.GetCmd().getHeight()
            m_pageObj.m_Revolution_ctrl.Value = m_pageObj.GetCmd().getRevolution()
        ElseIf Id = ID_REVOLUTION Then
            m_pageObj.GetCmd().Revolution(Value)
            m_pageObj.m_Pitch_ctrl.Value = m_pageObj.GetCmd().getPitch()
            m_pageObj.m_Height_ctrl.Value = m_pageObj.GetCmd().getHeight()
            m_pageObj.m_Revolution_ctrl.Value = m_pageObj.GetCmd().getRevolution()
        ElseIf Id = ID_TORSION_START_LENGTH Then
            m_pageObj.GetCmd().StartLength(Value)
        ElseIf Id = ID_TORSION_END_LENGTH Then
            m_pageObj.GetCmd().EndLength(Value)
        ElseIf Id = ID_COMPRESSION_START_REVOLUTION Then
            m_pageObj.GetCmd().StartRevolution(Value)
        ElseIf Id = ID_COMPRESSION_END_REVOLUTION Then
            m_pageObj.GetCmd().EndRevolution(Value)
        ElseIf Id = ID_COMPRESSION_START_PITCH Then
            m_pageObj.GetCmd().StartPitch(Value)
        ElseIf Id = ID_COMPRESSION_END_PITCH Then
            m_pageObj.GetCmd().EndPitch(Value)
        ElseIf Id = ID_SECTION_PARAMETER1 Then
            m_pageObj.GetCmd().Secparam1(Value)
            m_pageObj.m_Pitch_ctrl.Value = m_pageObj.GetCmd().getPitch()
            m_pageObj.m_Height_ctrl.Value = m_pageObj.GetCmd().getHeight()
            m_pageObj.m_Revolution_ctrl.Value = m_pageObj.GetCmd().getRevolution()
            If (m_pageObj.GetCmd().m_SpringType = swSpringType_Compression) Then
                m_pageObj.m_Start_Pitch_ctrl.Value = m_pageObj.GetCmd().getStartPitch()
                m_pageObj.m_End_Pitch_ctrl.Value = m_pageObj.GetCmd().getEndPitch()
            End If
        ElseIf Id = ID_SECTION_PARAMETER2 Then
            If Value = 0# And m_pageObj.GetCmd().getSecparam3() = 0# Then
                Exit Sub
            End If
            m_pageObj.GetCmd().Secparam2(Value)
            m_pageObj.m_Pitch_ctrl.Value = m_pageObj.GetCmd().getPitch()
            m_pageObj.m_Height_ctrl.Value = m_pageObj.GetCmd().getHeight()
            m_pageObj.m_Revolution_ctrl.Value = m_pageObj.GetCmd().getRevolution()
            If (m_pageObj.GetCmd().m_SpringType = swSpringType_Compression) Then
                m_pageObj.m_Start_Pitch_ctrl.Value = m_pageObj.GetCmd().getStartPitch()
                m_pageObj.m_End_Pitch_ctrl.Value = m_pageObj.GetCmd().getEndPitch()
            End If
        ElseIf Id = ID_SECTION_PARAMETER3 Then
            If Value = 0# And m_pageObj.GetCmd().getSecparam2() = 0# Then
                Exit Sub
            End If
            m_pageObj.GetCmd().Secparam3(Value)
            m_pageObj.m_Pitch_ctrl.Value = m_pageObj.GetCmd().getPitch()
            m_pageObj.m_Height_ctrl.Value = m_pageObj.GetCmd().getHeight()
            m_pageObj.m_Revolution_ctrl.Value = m_pageObj.GetCmd().getRevolution()
            If (m_pageObj.GetCmd().m_SpringType = swSpringType_Compression) Then
                m_pageObj.m_Start_Pitch_ctrl.Value = m_pageObj.GetCmd().getStartPitch()
                m_pageObj.m_End_Pitch_ctrl.Value = m_pageObj.GetCmd().getEndPitch()
            End If
        ElseIf Id = ID_TAPER_ANGLE Then
            m_pageObj.GetCmd().Angle(Value)
            If Math.Abs(m_pageObj.GetCmd().getAngle()) < 0.000001 Then
                m_pageObj.m_Check_Taper_Outward.enabled = False
            Else
                m_pageObj.m_Check_Taper_Outward.enabled = True
            End If
        End If
        m_pageObj.GetCmd().ShowDifferentParam()
    End Sub

    Private Sub PropertyManagerPage2Handler2_AfterClose()
        If m_isOk Then
            m_pageObj.GetCmd().OnOk()
        End If
        m_pageObj = Nothing
    End Sub

    Private Sub PropertyManagerPage2Handler2_OnOptionCheck(ByVal Id As Long)

    End Sub

    Private Sub PropertyManagerPage2Handler2_OnSelectionboxFocusChanged(ByVal Id As Long)

    End Sub

    Private Sub PropertyManagerPage2Handler2_OnTextboxChanged(ByVal Id As Long, ByVal Text As String)

    End Sub

    Private Sub PropertyManagerPage2Handler2_OnSelectionBoxListChanged(ByVal Id As Long, ByVal Text As Long)
        m_pageObj.GetCmd().OnSketchSelect(Id)
    End Sub

    Private Sub PropertyManagerPage2Handler2_OnSelectionboxCalloutCreated(ByVal Id As Long)

    End Sub

    Private Sub PropertyManagerPage2Handler2_OnSelectionboxCalloutDestroyed(ByVal Id As Long)

    End Sub

    Private Sub PropertyManagerPage2Handler2_OnComboboxEditChanged(ByVal Id As Long, ByVal Text As String)

    End Sub

    Private Function PropertyManagerPage2Handler2_OnActiveXControlCreated(ByVal Id As Long, ByVal status As Boolean) As Long

    End Function

    Private Function PropertyManagerPage2Handler2_OnPreviousPage() As Boolean

    End Function

    Private Function PropertyManagerPage2Handler2_OnNextPage() As Boolean

    End Function



End Class
