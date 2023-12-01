Imports System
Imports System.Collections
Imports System.Reflection
Imports System.Runtime.InteropServices

Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports SolidWorks.Interop.swpublished
Imports SolidWorksTools
Imports SolidWorksTools.File

Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.ComponentModel


<Guid("7001ad42-9e78-48aa-8843-cddca338021e")>
<ComVisible(True)>'�Ƿ��Զ����ز��
<SwAddin(Description:="��ϵ����17477451180", Title:="MSG��ľ", LoadAtStartup:=True)>
Public Class SwAddin
    Implements SolidWorks.Interop.swpublished.SwAddin
    Implements SwComFeature  'MIK

#Region "SolidWorks Registration ע�ᣬж��"

    <ComRegisterFunction()> Public Shared Sub RegisterFunction(ByVal t As Type)

        ' Get Custom Attribute: SwAddinAttribute
        Dim SWattr As SwAddinAttribute = Nothing
        Dim attributes() As Object = System.Attribute.GetCustomAttributes(GetType(SwAddin), GetType(SwAddinAttribute))

        If attributes.Length > 0 Then SWattr = DirectCast(attributes(0), SwAddinAttribute)

        Try
            Dim hklm As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine
            Dim hkcu As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser

            Dim keyname As String = "SOFTWARE\SolidWorks\Addins\{" + t.GUID.ToString() + "}"
            Dim addinkey As Microsoft.Win32.RegistryKey = hklm.CreateSubKey(keyname)

            addinkey.SetValue(Nothing, 0)
            addinkey.SetValue("Description", SWattr.Description)
            addinkey.SetValue("Title", SWattr.Title)

            keyname = "Software\SolidWorks\AddInsStartup\{" + t.GUID.ToString() + "}"
            addinkey = hkcu.CreateSubKey(keyname)

            addinkey.SetValue(Nothing, SWattr.LoadAtStartup, Microsoft.Win32.RegistryValueKind.DWord)
        Catch nl As System.NullReferenceException
            System.Windows.Forms.MessageBox.Show("ע����dllʱ�������⣺SWattrΪ�ա�\ n" & nl.Message)
        Catch e As System.Exception
            System.Windows.Forms.MessageBox.Show("ע���dllʱ�������⣺ " & e.Message)
        End Try
    End Sub

    <ComUnregisterFunction()> Public Shared Sub UnregisterFunction(ByVal t As Type)
        Try
            Dim hklm As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine
            Dim hkcu As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser

            Dim keyname As String = "SOFTWARE\SolidWorks\Addins\{" + t.GUID.ToString() + "}"
            hklm.DeleteSubKey(keyname)

            keyname = "Software\SolidWorks\AddInsStartup\{" + t.GUID.ToString() + "}"
            hkcu.DeleteSubKey(keyname)

        Catch Ex1 As NullReferenceException
            Windows.Forms.MessageBox.Show("ע����dllʱ�������⣺SWattrΪ�ա�\ n" & Ex1.Message)
        Catch Ex2 As Exception
            Windows.Forms.MessageBox.Show("ȡ��ע���dllʱ��������: " & Ex2.Message)
        End Try

    End Sub

#End Region

#Region "��������"

    Dim WithEvents iSwApp As SldWorks
    Dim iCmdMgr As ICommandManager
    Dim openDocs As Hashtable
    Dim SwEventPtr As SldWorks

    Dim ppage As UserPMPage
    Dim WithEvents WinHook As WinHooks.LocalWindowsHook

    Dim iBmp As BitmapHandler
    Dim TaskPanWinFormControl As Pakg

    Public Const mainCmdGroupID As Integer = 157
    Public Const mainItemID1 As Integer = 0, mainItemID2 As Integer = 1
    Public Const flyoutGroupID As Integer = 91

    Public PackGoTask As TaskpaneView
    Public bRet, bRet2 As Boolean

    ReadOnly Property SwdApp() As SldWorks
        Get
            Return iSwApp
        End Get
    End Property

    ReadOnly Property CmdMgr() As ICommandManager
        Get
            Return iCmdMgr
        End Get
    End Property

    ReadOnly Property OpenDocumentsTable() As Hashtable
        Get
            Return openDocs
        End Get
    End Property
#End Region

#Region "ISwAddin Implementation"
    '��װ������ô˷��� ʵ�ֲ����sw�����ӶϿ����Լ������Ƴ��������
    Function ConnectToSW(ByVal ThisSW As Object, ByVal Cookie As Integer) As Boolean Implements SolidWorks.Interop.swpublished.SwAddin.ConnectToSW
        iSwApp = ThisSW ' SldWorks ���õĳ������  ��SW������ϵ
        Module1.MySwApp = iSwApp '����ȫ��
        Module1.MathUtil = iSwApp.GetMathUtility
        If dllPath = "" Then
            dllPath = Assembly.GetExecutingAssembly().Location
            dllPath = Strings.Left(dllPath, Strings.InStrRev(dllPath, "\"))
            Fpth = dllPath & "SZM.data"
        End If
        'addinID = 'Cookie ' ����� id�� ��sw�õ�cookie
        iSwApp.SetAddinCallbackInfo(0, Me, Cookie) '���ò���ص���Ϣ
        iCmdMgr = iSwApp.GetCommandManager(Cookie) '��ȡָ����ӳ����CommandManager��

        AddCommandMgr() '��ӹ�����
        AddMSGtool()

        SwEventPtr = iSwApp '�����¼��������
        openDocs = New Hashtable

        AttachEventHandlers()

        'Setup Sample Property Manager
        AddPMP()
        AddHooks()
        Return True '���سɹ�����TRUE,����FALSE
    End Function
    'ж�ز�����ô˷���
    Function DisconnectFromSW() As Boolean Implements SolidWorks.Interop.swpublished.SwAddin.DisconnectFromSW

        RemoveCommandMgr()
        DetachEventHandlers()

        System.Runtime.InteropServices.Marshal.ReleaseComObject(iCmdMgr)
        iCmdMgr = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(iSwApp)
        iSwApp = Nothing
        'The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers 
        GC.Collect()
        GC.WaitForPendingFinalizers()

        GC.Collect()
        GC.WaitForPendingFinalizers()

        Return True '�ɹ�����true ���� ���� FALSE
    End Function
#End Region

#Region "ISwComFeature Implementation"'MIK

    Function Edit(ByVal app As Object, ByVal modelDoc As Object, ByVal feature As Object) As Object Implements SwComFeature.Edit
        MsgBox("Macro feature edit")
        Return Nothing
    End Function
    '�깦���ؽ�
    Function Regenerate(ByVal app As Object, ByVal modelDoc As Object, ByVal feature As Object) As Object Implements SwComFeature.Regenerate
        Dim OutputBodies As New Collection

        Dim swBodies() As Body2
        Dim swMacroFeatData As MacroFeatureData = feature.GetDefinition
        swMacroFeatData.EnableMultiBodyConsume = True

        Dim swModeler As Modeler = app.GetModeler
        Dim dblData(8) As Double
        dblData(0) = 0 : dblData(1) = 0 : dblData(2) = 0
        dblData(3) = 1 : dblData(4) = 0 : dblData(5) = 0
        dblData(6) = 0.1 : dblData(7) = 0.1 : dblData(8) = 0.1

        'Output body 1
        Dim swBody As Body2 = swModeler.CreateBodyFromBox3(dblData)
        OutputBodies.Add(swBody)

        'Output body 2
        dblData(1) = 0.15
        swBody = swModeler.CreateBodyFromBox3(dblData)
        OutputBodies.Add(swBody)

        ReDim swBodies(OutputBodies.Count - 1)
        For i = 1 To OutputBodies.Count
            swBody = OutputBodies.Item(i)
            Dim vEdges As Object = swBody.GetEdges
            Dim vFaces As Object = swBody.GetFaces

            For j = 0 To UBound(vEdges)
                swMacroFeatData.SetEdgeUserId(vEdges(j), j, 0)
            Next j
            For j = 0 To UBound(vFaces)
                swMacroFeatData.SetFaceUserId(vFaces(j), j, 0)
            Next j

            swBodies(i - 1) = OutputBodies.Item(i)
        Next i

        Regenerate = swBodies

    End Function

    Function Security(ByVal app As Object, ByVal modelDoc As Object, ByVal feature As Object) As Object Implements SwComFeature.Security
        'MsgBox("Macro feature security")
        'Return Nothing
    End Function

#End Region

#Region "UI Methods - ���淽��"

    Public Sub AddCommandMgr()

        If iBmp Is Nothing Then iBmp = New BitmapHandler()

        Dim Title As String = "MSG��ľ", ToolTip As String = "����"

        Dim ThisAssembly As Assembly = Assembly.GetAssembly(Me.GetType())

        Dim CmdGroupErr As Integer = 0
        Dim IgnorePrevious As Boolean = False
        Dim RegistryIDs As Object = Nothing
        Dim GetDataResult As Boolean = iCmdMgr.GetGroupDataFromRegistry(mainCmdGroupID, RegistryIDs)
        Dim KnownIDs As Integer() = New Integer(1) {mainItemID1, mainItemID2}

        If GetDataResult Then If Not CompareIDs(RegistryIDs, KnownIDs) Then IgnorePrevious = True '���ID��ƥ�䣬������CommandGroup

        Dim CmdGroup As ICommandGroup = iCmdMgr.CreateCommandGroup2(mainCmdGroupID, Title, ToolTip, "��ӭʹ��MSG", 0, IgnorePrevious, CmdGroupErr)

        If CmdGroup Is Nothing Or ThisAssembly Is Nothing Then Throw New NullReferenceException()

        CmdGroup.LargeIconList = iBmp.CreateFileFromResourceBitmap("MSGtool.aa_L.png", ThisAssembly)
        CmdGroup.SmallIconList = iBmp.CreateFileFromResourceBitmap("MSGtool.aa_S.png", ThisAssembly)
        CmdGroup.LargeMainIcon = iBmp.CreateFileFromResourceBitmap("MSGtool.Bf_L.bmp", ThisAssembly)
        CmdGroup.SmallMainIcon = iBmp.CreateFileFromResourceBitmap("MSGtool.Bf_S.bmp", ThisAssembly)


        '����λ������
        Dim cmdIndex(24) As Integer
        Dim menuToolbarOption As Integer = 1 Or 2 Or 3 Or 4

        cmdIndex(0) = CmdGroup.AddCommandItem2("MGS����", -1, "MSG��ľ������ʹ�ó�ʼ����", "��ʼ����", 0, "MSG����", "", 0, menuToolbarOption)
        cmdIndex(1) = CmdGroup.AddCommandItem2("�̶�����ɫ", -1, "Ϊ�ӽ�̶�����ɫ", "�̶�����ɫ", 1, "�̶�����ɫ", "", 1, menuToolbarOption)
        cmdIndex(2) = CmdGroup.AddCommandItem2("����ƽ��ʽ", -1, "�����ӽ�ƽ��ʽ", "����ƽ��ʽ", 2, "����ƽ��ʽ", "Enable", 2, menuToolbarOption)
        cmdIndex(3) = CmdGroup.AddCommandItem2("��������ͼ", -1, "�����������������ͼ������ͼ����һ�ӽǣ�������", "��������ͼ", 3, "��������ͼ", "", 3, menuToolbarOption)
        cmdIndex(4) = CmdGroup.AddCommandItem2("�ӽ��ע", -1, "Ϊ�ӽ𹤳�ͼ��ע�����ߴ�", "�ӽ��ע", 4, "�ӽ��ע", "Enable", 4, menuToolbarOption)
        cmdIndex(5) = CmdGroup.AddCommandItem2("��עǨ��", -1, "�����޸Ĺ���ͼ��ͼ����"�� "��ͼ����", 5�� "��ͼ����", "", 5, menuToolbarOption)
        cmdIndex(6) = CmdGroup.AddCommandItem2("��ͼ��ת", -1, "��ͼ��ת90"�� "��ͼ��ת", 6�� "��ͼ��ת", "", 6, menuToolbarOption)
        cmdIndex(7) = CmdGroup.AddCommandItem2("ȫ��ϵ������", -1, "���ӽ���������������ϵ������"�� "ȫ��ϵ������", 7�� "ȫ��ϵ������", "Enable", 7, menuToolbarOption)
        cmdIndex(8) = CmdGroup.AddCommandItem2("����ϵ������", -1, "�����ӽ�����ϵ������"�� "����ϵ������", 8�� "����ϵ������", "Enable", 8, menuToolbarOption)
        cmdIndex(9) = CmdGroup.AddCommandItem2("������ͼ����", -1, "��ʾ/����������ͼ"�� "������ͼ����", 9�� "������ͼ����", "Enable", 9, menuToolbarOption)
        cmdIndex(10) = CmdGroup.AddCommandItem2("���Ա༭", -1, ""�� "����ת�����Ϊ�ӽ�", 10�� "���Ա༭", "Enable", 10, menuToolbarOption)
        cmdIndex(11) = CmdGroup.AddCommandItem2("����ѡ��", -1, "��ǿѡ����"�� "����ѡ��", 11�� "�������ͬ������ͼ", "", 11, menuToolbarOption)
        cmdIndex(12) = CmdGroup.AddCommandItem2("չ���۵�", -1, "�ӽ�չ���۵������ÿ�ݼ���"�� "չ���۵�", 12�� "չ���۵�", "", 12, menuToolbarOption)
        cmdIndex(13) = CmdGroup.AddCommandItem2("ѡ��ģʽ", -1, "��ѡ��ģʽ���л�"�� "ѡ��ģʽ", 13�� "ѡ��ģʽ", "", 13, menuToolbarOption)
        cmdIndex(14) = CmdGroup.AddCommandItem2("����������", -1, "������ѡ�������cad"�� "����������", 14�� "����������", "Enable", 14, menuToolbarOption)
        cmdIndex(15) = CmdGroup.AddCommandItem2("������", -1, "������"�� "����ѡ�����", 15�� "������", "", 15, menuToolbarOption)
        cmdIndex(16) = CmdGroup.AddCommandItem2("������", -1, "������"�� "����ѡ�����", 16�� "������", "Enable", 16, menuToolbarOption)
        cmdIndex(17) = CmdGroup.AddCommandItem2("�Ƴ���", -1, "�Ƴ���"�� "ǿ��ɾ��������", 16�� "�Ƴ���", "Enable", 16, menuToolbarOption)
        cmdIndex(18) = CmdGroup.AddCommandItem2("������ѡȡ", -1, "ѡȡ��Ŀ�����һ��������������"�� "������ѡȡ", 17�� "������ѡȡ", "Enable", 17, menuToolbarOption)
        cmdIndex(19) = CmdGroup.AddCommandItem2("������", -1, "�ò�����༭װ����λ��"�� "������", 18�� "������", "", 18, menuToolbarOption)
        cmdIndex(20) = CmdGroup.AddCommandItem2("���Ĺ̶���", -1, "��ѡ��������Ϊ���ӽ�Ĺ̶���"�� "���Ĺ̶���", 19�� "���Ĺ̶���", "", 19, menuToolbarOption)
        cmdIndex(21) = CmdGroup.AddCommandItem2("cadģʽ", -1, "�ఴ���������"�� "cadģʽ", 20�� "cadģʽ", "", 20, 2)
        cmdIndex(22) = CmdGroup.AddCommandItem2("�湤��ͼ���", -1, "�湤��ͼ��渱��"�� "�湤��ͼ���", 21�� "�湤��ͼ���", "", 21, 2)
        cmdIndex(23) = CmdGroup.AddCommandItem2("��ʾ����", -1, "��ʾ������ѡ"�� "��ʾ����", 21�� "��ʾ����", "", 21, 2)
        cmdIndex(24) = CmdGroup.AddCommandItem2("�Ե���ʾ����", -1, "�Ե���ʾ����"�� "�Ե���ʾ����", 21�� "�Ե���ʾ����", "", 21, 2)

        '�˵�������    �� ��ʾ     ������������ ,ִ�к���
        '��ͼcadģʽ����   cad����  �����༭��� ͬ����ģģʽ ��ʾ������ͼ ������������ͼ
        '���ƣ��ƶ������죬���꣬�޼�������ƫ�ƣ���ϣ��ϲ����޸�ͬ�ȶ���

        CmdGroup.HasToolbar = True : CmdGroup.HasMenu = True : CmdGroup.Activate()

        Dim DocTypes() As Integer = {2, 3, 1}
        For Each docType As Integer In DocTypes
            Dim cmdTab As ICommandTab = iCmdMgr.GetCommandTab(docType, Title)

            If Not IsNothing(cmdTab) And Not GetDataResult Or IgnorePrevious Then '���ѡ����ڣ������Ǻ�����ע�����Ϣ�������´���ѡ� 
                Dim res As Boolean = iCmdMgr.RemoveCommandTab(cmdTab)
                cmdTab = Nothing
            End If

            If IsNothing(cmdTab) Then
                cmdTab = iCmdMgr.AddCommandTab(docType, Title)

                Dim cmdBox As CommandTabBox = cmdTab.AddCommandTabBox
                Dim cmdIDs(24), TextType(24) As Integer
                'TextType ��2��ֱ ��4ˮƽ ��1ֻ��ʾ�ı�
                cmdIDs(0) = CmdGroup.CommandID(cmdIndex(0)) : TextType(0) = 2
                cmdIDs(1) = CmdGroup.CommandID(cmdIndex(1)) : TextType(1) = 4
                cmdIDs(2) = CmdGroup.CommandID(cmdIndex(2)) : TextType(2) = 4
                cmdIDs(3) = CmdGroup.CommandID(cmdIndex(3)) : TextType(3) = 4
                cmdIDs(4) = CmdGroup.CommandID(cmdIndex(4)) : TextType(4) = 4
                cmdIDs(5) = CmdGroup.CommandID(cmdIndex(5)) : TextType(5) = 4
                cmdIDs(6) = CmdGroup.CommandID(cmdIndex(6)) : TextType(6) = 4
                cmdIDs(7) = CmdGroup.CommandID(cmdIndex(7)) : TextType(7) = 4
                cmdIDs(8) = CmdGroup.CommandID(cmdIndex(8)) : TextType(8) = 4
                cmdIDs(9) = CmdGroup.CommandID(cmdIndex(9)) : TextType(9) = 2
                cmdIDs(10) = CmdGroup.CommandID(cmdIndex(10)) : TextType(10) = 2
                cmdIDs(11) = CmdGroup.CommandID(cmdIndex(11)) : TextType(11) = 2
                cmdIDs(12) = CmdGroup.CommandID(cmdIndex(12)) : TextType(12) = 2
                cmdIDs(13) = CmdGroup.CommandID(cmdIndex(13)) : TextType(13) = 2
                cmdIDs(14) = CmdGroup.CommandID(cmdIndex(14)) : TextType(14) = 2
                cmdIDs(15) = CmdGroup.CommandID(cmdIndex(15)) : TextType(15) = 4
                cmdIDs(16) = CmdGroup.CommandID(cmdIndex(16)) : TextType(16) = 4
                cmdIDs(17) = CmdGroup.CommandID(cmdIndex(17)) : TextType(17) = 4
                cmdIDs(18) = CmdGroup.CommandID(cmdIndex(18)) : TextType(18) = 2
                cmdIDs(19) = CmdGroup.CommandID(cmdIndex(19)) : TextType(19) = 2
                cmdIDs(20) = CmdGroup.CommandID(cmdIndex(20)) : TextType(20) = 2
                cmdIDs(21) = CmdGroup.CommandID(cmdIndex(21)) : TextType(21) = 2
                cmdIDs(22) = CmdGroup.CommandID(cmdIndex(22)) : TextType(22) = 4
                cmdIDs(23) = CmdGroup.CommandID(cmdIndex(23)) : TextType(23) = 4
                cmdIDs(24) = CmdGroup.CommandID(cmdIndex(24)) : TextType(24) = 4

                cmdBox.AddCommands(cmdIDs, TextType)

                'Dim cmdBox1 As CommandTabBox = cmdTab.AddCommandTabBox
                'ReDim cmdIDs(2), TextType(2)
                'cmdBox1.AddCommands(cmdIDs, TextType) '��ӵ���������
                'cmdTab.AddSeparator(cmdBox1, cmdIDs(1)) '��ӷָ�

            End If
        Next

        ThisAssembly = Nothing

    End Sub

    Public Sub RemoveCommandMgr()
        Try
            iBmp.Dispose()
            iCmdMgr.RemoveCommandGroup(mainCmdGroupID)
        Catch e As Exception
        End Try
    End Sub

    Public Sub AddMSGtool()
        If Not Regis() Then Exit Sub
        Dim thisAssembly As Assembly = Assembly.GetAssembly(Me.GetType())
        If iBmp Is Nothing Then iBmp = New BitmapHandler()
        Dim taskBmpPath As String = iBmp.CreateFileFromResourceBitmap("MSGtool.Bf_S.bmp", thisAssembly)
        PackGoTask = iSwApp.CreateTaskpaneView2(taskBmpPath, "MSGtool������")

        Dim MsgTool As Pakg = New Pakg()
        PackGoTask.DisplayWindowFromHandlex64(MsgTool.Handle.ToInt64())
        MsgTool.Show()
        'bRet2 = Not bRet2
    End Sub

    'AddMacroFeature
    Function SMFeat() As Boolean

        Dim moddoc As ModelDoc2 = iSwApp.ActiveDoc
        Dim FeatMgr As FeatureManager = moddoc.FeatureManager
        'Collect input bodies
        Dim vBodies As Object = moddoc.GetBodies2(-1, False)
        Dim icopth() As String = {"MSGtool.Bf_S.bmp", "MSGtool.Bf_L.bmp", "MSGtool.Bf_S.bmp"}
        'Create the macro feature
        Dim MacroFeature As Feature = FeatMgr.InsertMacroFeature3("SFeature", "MSGtool.SwAddin", Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, vBodies, icopth, 0)

    End Function

    '��֪��ʲô����
    Function CompareIDs(ByVal storedIDs() As Integer, ByVal addinIDs() As Integer) As Boolean

        Dim storeList As New List(Of Integer)(storedIDs)
        Dim addinList As New List(Of Integer)(addinIDs)

        addinList.Sort()
        storeList.Sort()

        If Not addinList.Count = storeList.Count Then

            Return False
        Else

            For i As Integer = 0 To addinList.Count - 1
                If Not addinList(i) = storeList(i) Then

                    Return False
                End If
            Next
        End If

        Return True
    End Function

    Function AddPMP() As Boolean
        ppage = New UserPMPage
        ppage.Init(iSwApp, Me)
    End Function

    'Function RemovePMP() As Boolean
    '    ppage = Nothing
    'End Function


    Function AddHooks() As Boolean
        WinHook = New WinHooks.LocalWindowsHook(WinHooks.HookType.WH_KEYBOARD_LL)
        'WinHook.Install()
    End Function

    'Function RemoveHooks() As Boolean
    '    WinHook.Uninstall()
    'End Function
#End Region

#Region "Event Methods �¼�����"

    Sub AttachEventHandlers()
        AttachSWEvents()

        'Listen for events on all currently open docs
        AttachEventsToAllDocuments()
    End Sub

    Sub DetachEventHandlers()
        DetachSWEvents()

        'Close events on all currently open docs
        Dim docHandler As DocumentEventHandler
        Dim key As ModelDoc2
        Dim numKeys As Integer
        numKeys = openDocs.Count
        If numKeys > 0 Then
            Dim keys() As Object = New Object(numKeys - 1) {}

            'Remove all document event handlers
            openDocs.Keys.CopyTo(keys, 0)
            For Each key In keys
                docHandler = openDocs.Item(key)
                docHandler.DetachEventHandlers() 'This also removes the pair from the hash
                docHandler = Nothing
                key = Nothing
            Next
        End If
    End Sub

    Sub AttachSWEvents()
        Try
            AddHandler iSwApp.ActiveDocChangeNotify, AddressOf Me.SldWorks_ActiveDocChangeNotify
            AddHandler iSwApp.DocumentLoadNotify2, AddressOf Me.SldWorks_DocumentLoadNotify2
            AddHandler iSwApp.FileNewNotify2, AddressOf Me.SldWorks_FileNewNotify2
            AddHandler iSwApp.ActiveModelDocChangeNotify, AddressOf Me.SldWorks_ActiveModelDocChangeNotify
            AddHandler iSwApp.FileOpenPostNotify, AddressOf Me.SldWorks_FileOpenPostNotify
        Catch e As Exception
            Console.WriteLine(e.Message)
        End Try
    End Sub

    Sub DetachSWEvents()
        Try
            RemoveHandler iSwApp.ActiveDocChangeNotify, AddressOf Me.SldWorks_ActiveDocChangeNotify
            RemoveHandler iSwApp.DocumentLoadNotify2, AddressOf Me.SldWorks_DocumentLoadNotify2
            RemoveHandler iSwApp.FileNewNotify2, AddressOf Me.SldWorks_FileNewNotify2
            RemoveHandler iSwApp.ActiveModelDocChangeNotify, AddressOf Me.SldWorks_ActiveModelDocChangeNotify
            RemoveHandler iSwApp.FileOpenPostNotify, AddressOf Me.SldWorks_FileOpenPostNotify
        Catch e As Exception
            Console.WriteLine(e.Message)
        End Try
    End Sub

    Sub AttachEventsToAllDocuments()
        Dim modDoc As ModelDoc2
        modDoc = iSwApp.GetFirstDocument()
        While Not modDoc Is Nothing
            If Not openDocs.Contains(modDoc) Then
                AttachModelDocEventHandler(modDoc)
            End If
            modDoc = modDoc.GetNext()
        End While
    End Sub

    Function AttachModelDocEventHandler(ByVal modDoc As ModelDoc2) As Boolean
        If modDoc Is Nothing Then Return False

        Dim docHandler As DocumentEventHandler = Nothing

        If Not openDocs.Contains(modDoc) Then
            Select Case modDoc.GetType
                Case swDocumentTypes_e.swDocPART
                    docHandler = New PartEventHandler()
                Case swDocumentTypes_e.swDocASSEMBLY
                    docHandler = New AssemblyEventHandler()
                Case swDocumentTypes_e.swDocDRAWING
                    docHandler = New DrawingEventHandler()
            End Select

            docHandler.Init(iSwApp, Me, modDoc)
            docHandler.AttachEventHandlers()
            openDocs.Add(modDoc, docHandler)
        End If
    End Function

    Sub DetachModelEventHandler(ByVal modDoc As ModelDoc2)
        Dim docHandler As DocumentEventHandler
        docHandler = openDocs.Item(modDoc)
        openDocs.Remove(modDoc)
        modDoc = Nothing
        docHandler = Nothing
    End Sub

#End Region

#Region "Event Handlers �¼�����" '����ʱ��ľ��
    Function SldWorks_ActiveDocChangeNotify() As Integer
        'TODO: Add your implementation here
        'MIK
        'Dim bRet As Boolean = SMFeat()
    End Function

    Function SldWorks_DocumentLoadNotify2(ByVal docTitle As String, ByVal docPath As String) As Integer

    End Function

    Function SldWorks_FileNewNotify2(ByVal newDoc As Object, ByVal doctype As Integer, ByVal templateName As String) As Integer
        AttachEventsToAllDocuments()
    End Function

    Function SldWorks_ActiveModelDocChangeNotify() As Integer
        'TODO: Add your implementation here
    End Function

    Function SldWorks_FileOpenPostNotify(ByVal FileName As String) As Integer
        AttachEventsToAllDocuments()
    End Function
#End Region

#Region "UI Callbacks - ���淴��"

    Sub MSG����()
        Dim SZM As New SZM
        'SZM.Show()
        SZM.ShowDialog()
        SZM.Dispose()
    End Sub

    Sub �̶�����ɫ()
        Dim Model As ModelDoc2 = iSwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        Select Case Model.GetType
            Case 1
                FixedFaceColour(Model)
            Case 2
                Dim Mgr As SelectionMgr = Model.SelectionManager
                Dim Count As Integer = Mgr.GetSelectedObjectCount2(-1)
                If Count = 0 Then Exit Sub
                For i = 1 To Count
                    Dim Comp As Component2 = Mgr.GetSelectedObjectsComponent4(i, -1)
                    ' Dim CName As String = Comp.GetPathName
                    Dim Part As PartDoc = Comp.GetModelDoc2
                    FixedFaceColour(Part)
                Next
            Case Else
                Exit Sub
        End Select
        'FixedFaceColour(Model)
    End Sub
    Sub ����ƽ��ʽ()
        Dim bRet(1) As Boolean
        Dim Model As ModelDoc2 = iSwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        If Model.GetType <> 1 Then Exit Sub
        bRet(0) = MySwApp.GetUserPreferenceToggle(8)
        bRet(1) = MySwApp.GetUserPreferenceToggle(86)
        MySwApp.SetUserPreferenceToggle(8, False)
        MySwApp.SetUserPreferenceToggle(86, False)
        Dim Eor As Object = swToosl.ExportToCAD(Model, Nothing, Nothing)
        MySwApp.SetUserPreferenceToggle(8, bRet(0))
        MySwApp.SetUserPreferenceToggle(86, bRet(1))
        If Eor(4) <> "" Then MsgBox(Eor(4))
        If Eor(5) <> "" Then MsgBox(Eor(5))
    End Sub


    Sub ��������ͼ()

        Dim Model As ModelDoc2 = MySwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        If Model.GetType <> 1 Then MsgBox("���������ִ��") : Exit Sub
        CreateDrw(Model)

    End Sub
    Sub �ӽ��ע()
        Dim Model As ModelDoc2 = iSwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        If Model.GetType <> 3 Then MsgBox("���ڹ���ͼ��ִ��") : Exit Sub
        Dim Drawing As DrawingDoc = Model
        BendMark(Drawing)
    End Sub
    Sub ��ͼ����()
        Dim Model As ModelDoc2 = iSwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        If Model.GetType <> 1 Then Exit Sub
    End Sub
    Sub ��ͼ��ת()
        Dim Model As ModelDoc2 = iSwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        If Model.GetType <> 3 Then Exit Sub
        ViewRotate(Model)
    End Sub

    Sub ȫ��ϵ������()
        Dim Model As ModelDoc2 = iSwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        If Model.GetType <> 1 Then Exit Sub
        GlobalCorrection(Model)
    End Sub
    Sub ����ϵ������()
        Dim Model As ModelDoc2 = iSwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        If Model.GetType <> 1 Then Exit Sub
        LocalCorrection(Model)
    End Sub

    Sub ��������()
        Dim Model As ModelDoc2 = MySwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        If Model.GetType <> 1 Then MsgBox("���������ִ��") : Exit Sub
        CreateDrw(Model)
    End Sub

    Sub ����ת���ӽ�()

    End Sub

    Sub ����ѡ��()
        Dim Model As ModelDoc2 = iSwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub

    End Sub

    Sub չ���۵�()
        Dim Model As ModelDoc2 = iSwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        If Model.GetType <> 1 Then Exit Sub
        swToosl.չ���۵�(Model)
    End Sub

    Sub ѡ��ģʽ()
        Dim bRet As Boolean = MySwApp.GetUserPreferenceToggle(swUserPreferenceToggle_e.swEdgesHiddenEdgeSelectionInHLR)
        MySwApp.SetUserPreferenceToggle(swUserPreferenceToggle_e.swEdgesHiddenEdgeSelectionInHLR, Not bRet)
    End Sub

    Sub ����������()
        Dim Model As ModelDoc2 = iSwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        If Model.GetType <> 1 Then Exit Sub

        Try
            Dim swPart As PartDoc = Model
            Dim swSelMgr As SelectionMgr = Model.SelectionManager
            Dim Face As Face2 = swSelMgr.GetSelectedObject6(1, -1)
            If Face Is Nothing Then
                MsgBox("��ѡ����")
                Exit Sub
            End If
            Dim path As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop)
            Dim sPathName As String = path & "\SW������.dwg"
            swPart.ExportToDWG2(sPathName, Model.GetPathName, 2, False, Nothing, False, False, Nothing, 2)

        Catch ex As Exception
            Exit Sub
        End Try

    End Sub

    Sub ������()
        Dim Model As ModelDoc2 = iSwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub

        Dim SelMgr As SelectionMgr = Model.SelectionManager
        Dim Count As Integer = SelMgr.GetSelectedObjectCount2(-1)

        Try
            Select Case Model.GetType
                Case 1
                    Dim Part As PartDoc = Model
                    For i = 1 To Count
                        Dim Face As Face2 = SelMgr.GetSelectedObject6(i, -1)
                        Dim BodyFace As Body2 = Face.CreateSheetBody()
                        Part.CreateFeatureFromBody3(BodyFace, False, 2)
                    Next

                Case 2
                    Dim Assem As AssemblyDoc = Model
                    Dim EditComp As Component2 = Assem.GetEditTargetComponent
                    Dim Part As PartDoc = EditComp.GetModelDoc2
                    Dim EditXform As MathTransform = EditComp.Transform2
                    EditXform = EditXform.Inverse
                    If Count = 0 Then Exit Sub
                    For i = 1 To Count
                        Dim Face As Face2 = SelMgr.GetSelectedObject6(i, -1)
                        Dim Comp As Component2 = SelMgr.GetSelectedObjectsComponent4(i, -1)
                        Dim Xform As MathTransform = Comp.Transform2
                        Xform = Xform.Multiply(EditXform)
                        Dim BodyFace As Body2 = Face.CreateSheetBody()
                        BodyFace.ApplyTransform(Xform)

                        FixedFaceColour(Part)
                        Part.CreateFeatureFromBody3(BodyFace, False, 2)
                    Next
                Case Else
                    Exit Sub
            End Select
        Catch ex As Exception
            MsgBox("ѡ�����")
        End Try

        'FixedFaceColour(Model)
    End Sub

    Sub ����ȱ��()

    End Sub

    Sub ������ѡȡ()
        Dim Model As ModelDoc2 = iSwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        If Model.GetType <> 2 Then Exit Sub
        Assem.RangeSelection(Model)

    End Sub

    Sub �����༭()
        Dim Model As ModelDoc2 = iSwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        If Model.GetType <> 2 Then Exit Sub
        Dim AssyDoc As AssemblyDoc = Model
        Dim EditModel As ModelDoc2 = AssyDoc.GetEditTarget

        Dim Entityf As Entity = EditModel
        Entityf.Select4(False, Nothing)
        Dim Comp1 As Component2 = Entityf.GetComponent
        If Entityf Is Nothing Then MsgBox("�����������ѡȡһ�����") : Exit Sub

    End Sub

    Sub ���Ĺ̶���()
        Dim Model As ModelDoc2 = iSwApp.ActiveDoc
        If IsNothing(Model) Then Exit Sub
        setFixFace(Model)
    End Sub
    'װ�����и���
    Sub ���ƹ���1()
        'If Not ppage Is Nothing Then
        '    ppage.Show()
        'End If

        WinHook.Install()
        'MsgBox("���ܶ�������ϵ����qq983134237")
    End Sub
    Sub ���ƹ���2()
        'If Not ppage Is Nothing Then
        '    ppage.Show()
        'End If

        WinHook.Install()
        'MsgBox("���ܶ�������ϵ����qq983134237")
    End Sub

#End Region

    Function Enable() As Integer

        If Not Regis() Then Return 1 '0

        Return 1
    End Function

#Region "����"
    'Private Sub hk_HookEvenHookInvoked() Handles WinHook.HookInvoked
    '    Dim BB As WinHooks.HookEventArgs
    'End Sub
    Sub aac() '״̬����ʾ
        Dim Fra As Frame = iSwApp.Frame
        Dim StaBar As StatusBarPane = Fra.GetStatusBarPane
        StaBar.Text = "ccccc"
    End Sub
#End Region


End Class

'�����༭���
'�ӹ���ͼ�д���
'pdf����
'���ٴ��
'��ȡʵ��