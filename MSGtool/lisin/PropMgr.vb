Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst.swPropertyManagerPageOptions_e
Imports SolidWorks.Interop.swconst.swAddGroupBoxOptions_e
Imports SolidWorks.Interop.swconst.swPropertyManagerPageControlType_e
Imports SolidWorks.Interop.swconst.swPropertyManagerPageControlLeftAlign_e
Imports SolidWorks.Interop.swconst.swAddControlOptions_e
Imports SolidWorks.Interop.swconst.swSelectType_e
Imports SolidWorks.Interop.swconst.swControlBitmapLabelType_e
Imports SolidWorks.Interop.swconst.swPropMgrPageLabelStyle_e
Imports SolidWorks.Interop.swconst.swSpringType_e
Imports SolidWorks.Interop.swconst.swSpringDefineType_e
Imports SolidWorks.Interop.swconst.swNumberboxUnitType_e
Imports SolidWorks.Interop.swconst.swSpringProfileType_e
'Imports SolidWorks.Interop.swconst.
Public Class PropMgr

    Private m_Part As ModelDoc2
    Private m_feature As Feature
    Private m_Page As PropertyManagerPage2
    Private m_swPageCmd As New PropMgrCmd
    Private m_cmdState As swPageCmdState
    Private m_pageHdlr As New PropMgrHdlr
    ' Group 1
    Private m_Group_Basic_Parameter As PropertyManagerPageGroup
    Private m_Selection_Sketch As PropertyManagerPageSelectionbox
    Private m_Label_Type As PropertyManagerPageLabel
    Public m_Combo_Basic_Type As PropertyManagerPageCombobox
    Private m_Label_DefinedBy As PropertyManagerPageLabel
    Public m_Combo_Define_Method As PropertyManagerPageCombobox
    Public m_Height_ctrl As PropertyManagerPageNumberbox
    Public m_Pitch_ctrl As PropertyManagerPageNumberbox
    Public m_Revolution_ctrl As PropertyManagerPageNumberbox
    ' Group 2
    Private m_Group_Section_Profile As PropertyManagerPageGroup
    Public m_Combo_Section_Profile As PropertyManagerPageCombobox
    Public m_secparam1_ctrl As PropertyManagerPageNumberbox
    Public m_secparam2_ctrl As PropertyManagerPageNumberbox
    Public m_secparam3_ctrl As PropertyManagerPageNumberbox
    ' Group 3
    Public m_Group_End_Specification As PropertyManagerPageGroup
    Public m_Label_Start_Section As PropertyManagerPageLabel
    Public m_Combo_Start_EndType As PropertyManagerPageCombobox
    Public m_Start_Pitch_ctrl As PropertyManagerPageNumberbox
    Public m_Start_Revolution_ctrl As PropertyManagerPageNumberbox
    Public m_Start_Length_ctrl As PropertyManagerPageNumberbox
    Public m_Label_End_Section As PropertyManagerPageLabel
    Public m_Combo_End_EndType As PropertyManagerPageCombobox
    Public m_End_Pitch_ctrl As PropertyManagerPageNumberbox
    Public m_End_Revolution_ctrl As PropertyManagerPageNumberbox
    Public m_End_Length_ctrl As PropertyManagerPageNumberbox
    ' Group 4
    Private m_Group_Additional_Option As PropertyManagerPageGroup
    Public m_Angle_ctrl As PropertyManagerPageNumberbox
    Public m_Check_Taper_Outward As PropertyManagerPageCheckbox
    Public m_Check_Direction As PropertyManagerPageCheckbox
    Public m_Check_Right_Spin As PropertyManagerPageCheckbox
    Public m_Check_Ground As PropertyManagerPageCheckbox
    Public m_Check_DisplayTempBody As PropertyManagerPageCheckbox


    Private Sub Layout()
        Dim swPage As PropertyManagerPage2
        Dim swControl As PropertyManagerPageControl
        Dim title As String
        Dim buttonTypes As Long
        Dim message As String
        Dim Id As Long
        Dim controlType As Integer
        Dim caption As String
        Dim alignment As Integer
        Dim options As Long
        Dim tip As String
        Dim filterArray(0) As Long

        m_pageHdlr.Init(Me)

        If m_cmdState = swPageCmdState.swCmdCreate Then
            title = "Spring"
        Else
            title = m_feature.Name
        End If
        buttonTypes = swPropertyManagerOptions_OkayButton + swPropertyManagerOptions_CancelButton + swPropertyManagerOptions_LockedPage
        Dim ErrorH As Long
        m_Page = myswapp.CreatePropertyManagerPage(title, buttonTypes, m_pageHdlr, ErrorH)
        If Not m_Page Is Nothing Then

            '''''''' GROUP BOX 1(Basic Parameter)-----------------------------------------
            Id = ID_GROUP_BASIC_PARAMETER
            caption = "Basic parameters"
            options = swGroupBoxOptions_Visible + swGroupBoxOptions_Expanded
            m_Group_Basic_Parameter = m_Page.AddGroupBox(Id, caption, options)
            If Not m_Group_Basic_Parameter Is Nothing Then

                ' CONTROL Selection box  ------------------------------------------
                Id = ID_SELECTION_SKETCH
                controlType = swControlType_Selectionbox
                caption = "Sample selection box"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = "Select circular sketch"
                swControl = m_Group_Basic_Parameter.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_Selection_Sketch = swControl
                    filterArray(0) = swSelSKETCHES
                    m_Selection_Sketch.SingleEntityOnly = True
                    m_Selection_Sketch.Mark = ID_SELECTION_MARK_1
                    m_Selection_Sketch.Height = 14
                    m_Selection_Sketch.SetSelectionFilters(filterArray)
                    m_Selection_Sketch.setStandardPictureLabel(swBitmapLabel_SelectProfile)
                    If m_swPageCmd.m_EditDefinition = 1 Then
                        m_Selection_Sketch.enabled = False
                    End If
                End If

                ' CONTROL Label ---------------------------------------------------
                Id = ID_LABEL_TYPE
                controlType = swControlType_Label
                caption = "Spring type: "
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = ""
                swControl = m_Group_Basic_Parameter.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_Label_Type = swControl
                    m_Label_Type.Left = 6
                    m_Label_Type.Style = swPropMgrPageLabelStyle_LeftText
                End If

                ' CONTROL Combo box ------------------------------------------------------------------
                Id = ID_COMBO_TYPE
                controlType = swControlType_Combobox
                caption = "Basic Type"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = "Basic Type"
                swControl = m_Group_Basic_Parameter.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_Combo_Basic_Type = swControl
                    m_Combo_Basic_Type.Left = 6
                    m_Combo_Basic_Type.top = 60
                    m_Combo_Basic_Type.AddItems("Compression")
                    m_Combo_Basic_Type.AddItems("Extension")
                    m_Combo_Basic_Type.AddItems("Torsion")
                    m_Combo_Basic_Type.AddItems("Spiral")
                    If m_swPageCmd.m_SpringType = swSpringType_Compression Then
                        m_Combo_Basic_Type.CurrentSelection = 0
                    ElseIf m_swPageCmd.m_SpringType = swSpringType_Extension Then
                        m_Combo_Basic_Type.CurrentSelection = 1
                    ElseIf m_swPageCmd.m_SpringType = swSpringType_Torsion Then
                        m_Combo_Basic_Type.CurrentSelection = 2
                    Else
                        m_Combo_Basic_Type.CurrentSelection = 3
                    End If
                    If m_swPageCmd.m_EditDefinition = 1 Then
                        m_Combo_Basic_Type.enabled = False
                    End If
                End If

                ' CONTROL Label ---------------------------------------------------
                Id = ID_LABEL_PARAM_METHOD
                controlType = swControlType_Label
                caption = "Defined by: "
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = ""
                swControl = m_Group_Basic_Parameter.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_Label_DefinedBy = swControl
                    m_Label_DefinedBy.top = 80
                    m_Label_DefinedBy.Left = 6
                    m_Label_DefinedBy.Style = swPropMgrPageLabelStyle_LeftText
                End If

                ' CONTROL Combo box  ----------------------------------------------
                Id = ID_COMBO_DEFINE_METHOD
                controlType = swControlType_Combobox
                caption = "Defined By"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = "Select one group parameter to define"
                swControl = m_Group_Basic_Parameter.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_Combo_Define_Method = swControl
                    m_Combo_Define_Method.Left = 6
                    m_Combo_Define_Method.top = 90
                    If m_swPageCmd.m_SpringType = swSpringType_Compression Or m_swPageCmd.m_SpringType = swSpringType_Extension Then
                        m_Combo_Define_Method.AddItems("Pitch and Revolution")
                        m_Combo_Define_Method.AddItems("Height and Revolution")
                        m_Combo_Define_Method.AddItems("Height and Pitch")
                        If m_swPageCmd.m_DefineType = swSpringDefineType_PitchAndRevolution Then
                            m_Combo_Define_Method.CurrentSelection = 0
                        ElseIf m_swPageCmd.m_DefineType = swSpringDefineType_HeightAndRevolution Then
                            m_Combo_Define_Method.CurrentSelection = 1
                        Else
                            m_Combo_Define_Method.CurrentSelection = 2
                        End If
                    ElseIf m_swPageCmd.m_SpringType = swSpringType_Torsion Then
                        m_Combo_Define_Method.AddItems("Height")
                        m_Combo_Define_Method.AddItems("Revolution")
                        If m_swPageCmd.m_DefineType = swSpringDefineType_PitchAndRevolution Then
                            m_Combo_Define_Method.CurrentSelection = 0
                        Else
                            m_Combo_Define_Method.CurrentSelection = 1
                        End If
                    Else
                        m_Combo_Define_Method.AddItems("Pitch and Revolution")
                    End If
                End If

                ' CONTROL Number box  ---------------------------------------------
                Id = ID_HEIGHT
                controlType = swControlType_Numberbox
                caption = "Height"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = "Height"
                swControl = m_Group_Basic_Parameter.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_Height_ctrl = swControl
                    If m_swPageCmd.m_SpringType = swSpringType_Compression Or m_swPageCmd.m_SpringType = swSpringType_Extension Then
                        If m_swPageCmd.m_DefineType = swSpringDefineType_PitchAndRevolution Then
                            m_Height_ctrl.enabled = False
                        Else
                            m_Height_ctrl.enabled = True
                        End If
                    ElseIf m_swPageCmd.m_SpringType = swSpringType_Torsion Then
                        If m_swPageCmd.m_DefineType = swSpringDefineType_PitchAndRevolution Then
                            m_Height_ctrl.enabled = True
                        Else
                            m_Height_ctrl.enabled = False
                        End If
                    Else
                        m_Height_ctrl.enabled = False
                    End If
                    m_Height_ctrl.SetRange(swNumberBox_Length, 0#, 10.0#, 0.01, True)
                    m_Height_ctrl.Value = m_swPageCmd.getHeight()
                    m_Height_ctrl.setStandardPictureLabel(swBitmapLabel_LinearDistance)
                End If

                Id = ID_PITCH
                controlType = swControlType_Numberbox
                caption = "Pitch"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = "Pitch"
                swControl = m_Group_Basic_Parameter.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_Pitch_ctrl = swControl
                    If m_swPageCmd.m_SpringType = swSpringType_Compression Or m_swPageCmd.m_SpringType = swSpringType_Extension Then
                        If m_swPageCmd.m_DefineType = swSpringDefineType_HeightAndRevolution Then
                            m_Pitch_ctrl.enabled = False
                        Else
                            m_Pitch_ctrl.enabled = True
                        End If
                    ElseIf m_swPageCmd.m_SpringType = swSpringType_Torsion Then
                        m_Pitch_ctrl.enabled = False
                    Else
                        m_Pitch_ctrl.enabled = True
                    End If
                    m_Pitch_ctrl.SetRange(swNumberBox_Length, 0#, 1.0#, 0.001, True)
                    m_Pitch_ctrl.Value = m_swPageCmd.getPitch()
                    m_Pitch_ctrl.setStandardPictureLabel(swBitmapLabel_LinearDistance)
                End If

                Id = ID_REVOLUTION
                controlType = swControlType_Numberbox
                caption = "Revolution"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = "Number of revolutions"
                swControl = m_Group_Basic_Parameter.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_Revolution_ctrl = swControl
                    If m_swPageCmd.m_SpringType = swSpringType_Compression Or m_swPageCmd.m_SpringType = swSpringType_Extension Then
                        If m_swPageCmd.m_DefineType = 2 Then
                            m_Revolution_ctrl.enabled = False
                        Else
                            m_Revolution_ctrl.enabled = True
                        End If
                    ElseIf m_swPageCmd.m_SpringType = swSpringType_Torsion Then
                        If m_swPageCmd.m_DefineType = swSpringDefineType_PitchAndRevolution Then
                            m_Revolution_ctrl.enabled = False
                        Else
                            m_Revolution_ctrl.enabled = True
                        End If
                    Else
                        m_Revolution_ctrl.enabled = True
                    End If
                    m_Revolution_ctrl.SetRange(swNumberBox_UnitlessDouble, 0.01, 400.0#, 1, True)
                    m_Revolution_ctrl.Value = m_swPageCmd.getRevolution()
                    m_Revolution_ctrl.setStandardPictureLabel(swBitmapLabel_LinearDistance)
                End If

            End If

            '''''''' GROUP BOX 2 (Profile for Section Plane)------------------------------
            Id = ID_GROUP_SECTION_PROFILE
            caption = "Section profile"
            If m_swPageCmd.m_EditDefinition = 1 Then
                options = 0
            Else
                options = swGroupBoxOptions_Visible '+ swGroupBoxOptions_Checkbox
            End If
            m_Group_Section_Profile = m_Page.AddGroupBox(Id, caption, options)
            If Not m_Group_Section_Profile Is Nothing Then

                ' CONTROL Combo box  ----------------------------------------------
                Id = ID_COMBO_SECTION_PROFILE
                controlType = swControlType_Combobox
                caption = "Profile"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = "Select one section profile type"
                swControl = m_Group_Section_Profile.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_Combo_Section_Profile = swControl
                    m_Combo_Section_Profile.Left = 6
                    m_Combo_Section_Profile.AddItems("Circle")
                    m_Combo_Section_Profile.AddItems("Rectangle")
                    m_Combo_Section_Profile.AddItems("Trapezoid")
                    If m_swPageCmd.m_ProfileType = swSpringProfileType_Circle Then
                        m_Combo_Section_Profile.CurrentSelection = 0
                    End If
                    If m_swPageCmd.m_ProfileType = swSpringProfileType_Rectangle Then
                        m_Combo_Section_Profile.CurrentSelection = 1
                    End If
                    If m_swPageCmd.m_ProfileType = swSpringProfileType_Trapezoid Then
                        m_Combo_Section_Profile.CurrentSelection = 2
                    End If
                    If m_swPageCmd.m_EditDefinition = 1 Then
                        m_Combo_Section_Profile.enabled = False
                    End If
                End If

                ' CONTROL Number box ----------------------------------------------
                Id = ID_SECTION_PARAMETER1
                controlType = swControlType_Numberbox
                caption = "Param1"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = ""
                swControl = m_Group_Section_Profile.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_secparam1_ctrl = swControl
                    m_secparam1_ctrl.tip = "Diameter"
                    m_secparam1_ctrl.SetRange(swNumberBox_Length, 0.00001, 1.0#, 0.00001, True)
                    m_secparam1_ctrl.Value = m_swPageCmd.getSecparam1()
                    m_secparam1_ctrl.setStandardPictureLabel(swBitmapLabel_LinearDistance)
                    If m_swPageCmd.m_EditDefinition = 1 Then
                        m_secparam1_ctrl.enabled = False
                    End If
                End If

                Id = ID_SECTION_PARAMETER2
                controlType = swControlType_Numberbox
                caption = "Param2"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = ""
                swControl = m_Group_Section_Profile.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_secparam2_ctrl = swControl
                    If m_swPageCmd.m_ProfileType = swSpringProfileType_Circle Then
                        m_secparam2_ctrl.enabled = False
                    End If
                    m_secparam2_ctrl.SetRange(swNumberBox_Length, 0#, 1.0#, 0.00001, True)
                    m_secparam2_ctrl.Value = m_swPageCmd.getSecparam2()
                    m_secparam2_ctrl.setStandardPictureLabel(swBitmapLabel_LinearDistance)
                    If m_swPageCmd.m_EditDefinition = 1 Then
                        m_secparam2_ctrl.enabled = False
                    End If
                End If

                Id = ID_SECTION_PARAMETER3
                controlType = swControlType_Numberbox
                caption = "Param3"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = ""
                swControl = m_Group_Section_Profile.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_secparam3_ctrl = swControl
                    If m_swPageCmd.m_ProfileType = swSpringProfileType_Circle Or m_swPageCmd.m_ProfileType = swSpringProfileType_Rectangle Then
                        m_secparam3_ctrl.enabled = False
                    End If
                    m_secparam3_ctrl.SetRange(swNumberBox_Length, 0#, 1.0#, 0.00001, True)
                    m_secparam3_ctrl.Value = m_swPageCmd.getSecparam3()
                    m_secparam3_ctrl.setStandardPictureLabel(swBitmapLabel_LinearDistance)
                    If m_swPageCmd.m_EditDefinition = 1 Then
                        m_secparam3_ctrl.enabled = False
                    End If
                End If
            End If

            '''''''' GROUP BOX 3 (End specification)-------------------------------------------
            Id = ID_GROUP_END_SPECIFICATION
            caption = "End specification"
            options = swGroupBoxOptions_Visible '+ swGroupBoxOptions_Checkbox
            m_Group_End_Specification = m_Page.AddGroupBox(Id, caption, options)
            If m_swPageCmd.m_SpringType = swSpringType_Spiral Then
                m_Group_End_Specification.Visible = False
            End If
            If Not m_Group_End_Specification Is Nothing Then

                ' CONTROL Label ---------------------------------------------------
                Id = ID_LABEL_START_SECTION
                controlType = swControlType_Label
                caption = "Start section: "
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = ""
                swControl = m_Group_End_Specification.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_Label_Start_Section = swControl
                    m_Label_Start_Section.Left = 6
                    m_Label_Start_Section.Style = swPropMgrPageLabelStyle_LeftText
                End If

                ' CONTROL Combo box (End type for torsion spring) ----------------------------------------------
                Id = ID_COMBO_START_ENDTYPE
                controlType = swControlType_Combobox
                caption = "End tpye"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled + swControlOptions_SmallGapAbove
                tip = "Select one end type"
                swControl = m_Group_End_Specification.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_Combo_Start_EndType = swControl
                    If m_swPageCmd.m_SpringType = swSpringType_Extension Then
                        m_Combo_Start_EndType.AddItems("Full loop")
                        m_Combo_Start_EndType.AddItems("Half loop")
                    End If
                    If m_swPageCmd.m_SpringType = swSpringType_Torsion Then
                        m_Combo_Start_EndType.AddItems("Hook")
                        m_Combo_Start_EndType.AddItems("Straight")
                        m_Combo_Start_EndType.AddItems("Hinge")
                        m_Combo_Start_EndType.AddItems("Straight off")
                    End If
                    If m_swPageCmd.m_SpringType = swSpringType_Extension Or m_swPageCmd.m_SpringType = swSpringType_Torsion Then
                        m_Combo_Start_EndType.Visible = True
                    Else
                        m_Combo_Start_EndType.Visible = False
                    End If
                    If m_swPageCmd.m_EndTypeS = 0 Then
                        m_Combo_Start_EndType.CurrentSelection = 0
                    ElseIf m_swPageCmd.m_EndTypeS = 1 Then
                        m_Combo_Start_EndType.CurrentSelection = 1
                    ElseIf m_swPageCmd.m_EndTypeS = 2 Then
                        m_Combo_Start_EndType.CurrentSelection = 2
                    Else
                        m_Combo_Start_EndType.CurrentSelection = 3
                    End If
                End If


                ' CONTROL Number box  ---------------------------------------------
                Id = ID_COMPRESSION_START_PITCH
                controlType = swControlType_Numberbox
                caption = "the pitch of start section"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled + swControlOptions_SmallGapAbove
                tip = "pitch"
                swControl = m_Group_End_Specification.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_Start_Pitch_ctrl = swControl
                    If m_swPageCmd.m_SpringType = swSpringType_Compression Then
                        m_Start_Pitch_ctrl.Visible = True
                    Else
                        m_Start_Pitch_ctrl.Visible = False
                    End If
                    m_Start_Pitch_ctrl.SetRange(swNumberBox_Length, 0#, 1.0#, 0.01, True)
                    m_Start_Pitch_ctrl.Value = m_swPageCmd.getStartPitch()
                    m_Start_Pitch_ctrl.setStandardPictureLabel(swBitmapLabel_LinearDistance)
                End If

                ' CONTROL Number box  ---------------------------------------------
                Id = ID_COMPRESSION_START_REVOLUTION
                controlType = swControlType_Numberbox
                caption = "the revolution of start section"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = "revolution"
                swControl = m_Group_End_Specification.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_Start_Revolution_ctrl = swControl
                    If m_swPageCmd.m_SpringType = swSpringType_Compression Then
                        m_Start_Revolution_ctrl.Visible = True
                    Else
                        m_Start_Revolution_ctrl.Visible = False
                    End If
                    m_Start_Revolution_ctrl.SetRange(swNumberBox_UnitlessDouble, 0#, 30.0#, 0.25, True)
                    m_Start_Revolution_ctrl.Value = m_swPageCmd.getStartRevolution()
                    m_Start_Revolution_ctrl.setStandardPictureLabel(swBitmapLabel_LinearDistance)
                End If

                ' CONTROL Number box  ---------------------------------------------
                Id = ID_TORSION_START_LENGTH
                controlType = swControlType_Numberbox
                caption = "the length of start straight section for torsion spring"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = "the length of start straight section for torsion spring"
                swControl = m_Group_End_Specification.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_Start_Length_ctrl = swControl
                    If m_swPageCmd.m_SpringType = swSpringType_Torsion Then
                        m_Start_Length_ctrl.Visible = True
                    Else
                        m_Start_Length_ctrl.Visible = False
                    End If
                    m_Start_Length_ctrl.SetRange(swNumberBox_Length, 0#, 10.0#, 0.01, True)
                    m_Start_Length_ctrl.Value = m_swPageCmd.getStartLength()
                    m_Start_Length_ctrl.setStandardPictureLabel(swBitmapLabel_LinearDistance)
                End If


                ' CONTROL Label ---------------------------------------------------
                Id = ID_LABEL_END_SECTION
                controlType = swControlType_Label
                caption = "End section: "
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = ""
                swControl = m_Group_End_Specification.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_Label_End_Section = swControl
                    m_Label_End_Section.Left = 6
                    If m_swPageCmd.m_SpringType = swSpringType_Spiral Then
                        m_Label_End_Section.Visible = False
                    Else
                        m_Label_End_Section.Visible = True
                    End If
                    m_Label_End_Section.Style = swPropMgrPageLabelStyle_LeftText
                End If

                ' CONTROL Combo box (End type for torsion spring) ----------------------------------------------
                Id = ID_COMBO_END_ENDTYPE
                controlType = swControlType_Combobox
                caption = "End tpye"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled + swControlOptions_SmallGapAbove
                tip = "Select one end type"
                swControl = m_Group_End_Specification.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_Combo_End_EndType = swControl
                    If m_swPageCmd.m_SpringType = swSpringType_Extension Then
                        m_Combo_End_EndType.AddItems("Full loop")
                        m_Combo_End_EndType.AddItems("Half loop")
                    End If
                    If m_swPageCmd.m_SpringType = swSpringType_Torsion Then
                        m_Combo_End_EndType.AddItems("Hook")
                        m_Combo_End_EndType.AddItems("Straight")
                        m_Combo_End_EndType.AddItems("Hinge")
                        m_Combo_End_EndType.AddItems("Straight off")
                    End If
                    If m_swPageCmd.m_SpringType = swSpringType_Extension Or m_swPageCmd.m_SpringType = swSpringType_Torsion Then
                        m_Combo_End_EndType.Visible = True
                    Else
                        m_Combo_End_EndType.Visible = False
                    End If
                    If m_swPageCmd.m_EndTypeE = 0 Then
                        m_Combo_End_EndType.CurrentSelection = 0
                    ElseIf m_swPageCmd.m_EndTypeE = 1 Then
                        m_Combo_End_EndType.CurrentSelection = 1
                    ElseIf m_swPageCmd.m_EndTypeE = 2 Then
                        m_Combo_End_EndType.CurrentSelection = 2
                    Else
                        m_Combo_End_EndType.CurrentSelection = 3
                    End If
                End If

                ' CONTROL Number box  ---------------------------------------------
                Id = ID_COMPRESSION_END_PITCH
                controlType = swControlType_Numberbox
                caption = "the pitch of end section"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled + swControlOptions_SmallGapAbove
                tip = "pitch"
                swControl = m_Group_End_Specification.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_End_Pitch_ctrl = swControl
                    If m_swPageCmd.m_SpringType = swSpringType_Compression Then
                        m_End_Pitch_ctrl.Visible = True
                    Else
                        m_End_Pitch_ctrl.Visible = False
                    End If
                    m_End_Pitch_ctrl.SetRange(swNumberBox_Length, 0#, 1.0#, 0.01, True)
                    m_End_Pitch_ctrl.Value = m_swPageCmd.getEndPitch()
                    m_End_Pitch_ctrl.setStandardPictureLabel(swBitmapLabel_LinearDistance)
                End If

                ' CONTROL Number box  ---------------------------------------------
                Id = ID_COMPRESSION_END_REVOLUTION
                controlType = swControlType_Numberbox
                caption = "the revolution of end section"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = "revolution"
                swControl = m_Group_End_Specification.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_End_Revolution_ctrl = swControl
                    If m_swPageCmd.m_SpringType = swSpringType_Compression Then
                        m_End_Revolution_ctrl.Visible = True
                    Else
                        m_End_Revolution_ctrl.Visible = False
                    End If
                    m_End_Revolution_ctrl.SetRange(swNumberBox_UnitlessDouble, 0#, 30.0#, 0.25, True)
                    m_End_Revolution_ctrl.Value = m_swPageCmd.getEndRevolution()
                    m_End_Revolution_ctrl.setStandardPictureLabel(swBitmapLabel_LinearDistance)
                End If

                ' CONTROL Number box  ---------------------------------------------
                Id = ID_TORSION_END_LENGTH
                controlType = swControlType_Numberbox
                caption = "the length of end straight section for torsion spring"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = "the length of end straight section for torsion spring"
                swControl = m_Group_End_Specification.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_End_Length_ctrl = swControl
                    If m_swPageCmd.m_SpringType = swSpringType_Torsion Then
                        m_End_Length_ctrl.Visible = True
                    Else
                        m_End_Length_ctrl.Visible = False
                    End If
                    m_End_Length_ctrl.SetRange(swNumberBox_Length, 0#, 10.0#, 0.01, True)
                    m_End_Length_ctrl.Value = m_swPageCmd.getEndLength()
                    m_End_Length_ctrl.setStandardPictureLabel(swBitmapLabel_LinearDistance)
                End If

                ' Check Box(Grinding end)------------------------------------------
                Id = ID_CHECK_GROUND
                controlType = swControlType_Checkbox
                caption = "Ground two ends"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = ""
                swControl = m_Group_End_Specification.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_Check_Ground = swControl
                    If m_swPageCmd.m_SpringType = swSpringType_Compression Then
                        m_Check_Ground.Visible = True
                    Else
                        m_Check_Ground.Visible = False
                    End If
                    If m_swPageCmd.m_Ground = 1 Then
                        m_Check_Ground.Checked = True
                    Else
                        m_Check_Ground.Checked = False
                    End If
                End If
            End If

            '''''''' GROUP BOX 4 (Additional option)-------------------------------------------
            Id = ID_GROUP_ADDITIONAL_OPTION
            caption = "Additional option"
            options = swGroupBoxOptions_Visible ' + swGroupBoxOptions_Expanded
            m_Group_Additional_Option = m_Page.AddGroupBox(Id, caption, options)
            If Not m_Group_Additional_Option Is Nothing Then

                ' Check Box (Reverse direction)-------------------------------------
                Id = ID_CHECK_DIRECTION
                controlType = swControlType_Checkbox
                caption = "Reverse direction"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = ""
                swControl = m_Group_Additional_Option.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_Check_Direction = swControl
                    If m_swPageCmd.m_SpringType = swSpringType_Spiral Then
                        m_Check_Direction.enabled = False
                    End If
                    If m_swPageCmd.m_Direction = 1 Then
                        m_Check_Direction.Checked = True
                    Else
                        m_Check_Direction.Checked = False
                    End If
                End If

                ' Check Box(Clockwise or counter-clockwise)------------------------
                Id = ID_CHECK_RIGHT_SPIN
                controlType = swControlType_Checkbox
                caption = "Right-hand spin"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = "Change spin direction"
                swControl = m_Group_Additional_Option.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_Check_Right_Spin = swControl
                    If m_swPageCmd.m_RightHand = 1 Then
                        m_Check_Right_Spin.Checked = True
                    Else
                        m_Check_Right_Spin.Checked = False
                    End If
                End If

                ' Check Box(Show temporary body)----------------------------------------------
                Id = ID_CHECK_TEMP_BODY
                controlType = swControlType_Checkbox
                caption = "Show preview"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = "display temporary body"
                swControl = m_Group_Additional_Option.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_Check_DisplayTempBody = swControl
                    If m_swPageCmd.m_ShowTempBody = 1 Then
                        m_Check_DisplayTempBody.Checked = True
                    Else
                        m_Check_DisplayTempBody.Checked = False
                    End If
                End If

                ' CONTROL Number box  ---------------------------------------------
                Id = ID_TAPER_ANGLE
                controlType = swControlType_Numberbox
                caption = "Taper angle"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = "Taper angle"
                swControl = m_Group_Additional_Option.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_Angle_ctrl = swControl
                    If m_swPageCmd.m_SpringType = swSpringType_Compression Then
                        m_Angle_ctrl.Visible = True
                    Else
                        m_Angle_ctrl.Visible = False
                    End If
                    m_Angle_ctrl.SetRange(swNumberBox_Angle, 0#, 1.570796325, 0.01745329252, True)
                    m_Angle_ctrl.Value = m_swPageCmd.getAngle()
                    m_Angle_ctrl.setStandardPictureLabel(swBitmapLabel_AngularDistance)
                End If

                ' Check Box(Taper outward)-----------------------------------------
                Id = ID_CHECK_TAPER_OUTWARD
                controlType = swControlType_Checkbox
                caption = "Taper outward"
                alignment = swControlAlign_Indent
                options = swControlOptions_Visible + swControlOptions_Enabled
                tip = "Taper outward"
                swControl = m_Group_Additional_Option.AddControl(Id, controlType, caption, alignment, options, tip)
                If Not swControl Is Nothing Then
                    m_Check_Taper_Outward = swControl
                    If m_swPageCmd.m_SpringType = swSpringType_Compression Then
                        m_Check_Taper_Outward.Visible = True
                        If Math.Abs(m_swPageCmd.getAngle()) < 0.000001 Then
                            m_Check_Taper_Outward.enabled = False
                        Else
                            m_Check_Taper_Outward.enabled = True
                        End If
                    Else
                        m_Check_Taper_Outward.Visible = False
                    End If
                    If m_swPageCmd.m_TaperOutward = 1 Then
                        m_Check_Taper_Outward.Checked = True
                    Else
                        m_Check_Taper_Outward.Checked = False
                    End If
                End If

            End If

        End If
    End Sub


    Public Sub Show()
        m_Page.Show()
    End Sub


    Sub Init(Part As ModelDoc2, feature As Feature, cmdState As Byte, macroPath As String)
        m_Part = Part
        If Not IsNothing(feature) Then
            m_feature = feature
        End If
        m_cmdState = cmdState
        If m_cmdState = swPageCmdState.swCmdCreate Then
            m_Part.ClearSelection()
        End If
        m_swPageCmd.Init(Part, feature, cmdState, macroPath)
        Layout()
    End Sub


    Public Function GetCmd() As PropMgrCmd
        GetCmd = m_swPageCmd
    End Function



End Class
