Imports System.Security.Principal
Imports System.Xml
Imports Sympraxis.Utilities.INormalPlugin
Imports Sympraxis.Utilities.LayoutManager
Imports Sympraxis.Common
Imports System.Windows.Controls.Ribbon
Imports Sympraxis.Plugins.CaseInfo

Imports System.Windows.Forms
Imports Sympraxis.Common.Qualifier
Imports Sympraxis.Utilities.IInputParser
Imports Sympraxis.Utilities.IExitParser
Imports Sympraxis.Common.IOManager
Imports Sympraxis.Utilities.KeyboardHook

Imports System.IO
Imports System.Windows.Threading
Imports System.Media
Imports Sympraxis.Utilities.Exception
Imports System.IO.Packaging
Imports System.Windows.Resources

Imports Sympraxis.Utilities.InstanceManager
Imports System.Runtime.Remoting
Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Remoting.Lifetime
Imports System.Runtime.Remoting.Channels.Tcp
Imports System.Threading
Imports System.Configuration
Imports System.Net.NetworkInformation
Imports Sympraxis.Utilities.WinZip

Class InteractiveHost
    Implements IPluginhost

    Implements IPluginhostExtented

    Private _CurrentlayoutIdx As Integer = 0

    Private CompletedObjectCount As Integer = 0
    Private WorkedObjectCount As Integer = 0

    Private MyDictCount As New Dictionary(Of String, String)
    Private OfflineWPPath As String
    Private _Isautoopen As Boolean = False
    

#Region "Gobalvariable"
    Private Shared mutex As Mutex
#Region "ConstantVariable"

#End Region

#Region "ReadOnlyVariable"

#End Region

#Region "PrivateVariable"

    'Class Variable
    Private isExitflag As Boolean = False

    Private objFrmSplash As SplashForm
    Private WithEvents ObjFrmException As SympraxisException
    Private appHostConfigXml As AppHostConfig
    Private overrideAppHostConfig As OverrideAppHostConfig
    Private loadPluginsConfigXml As IntHostConfig
    Private WIConfig As WIConfig
    Private defaultPluginsConfigXml As IntHostConfig
    Private objWCF As IWorkflowClient
    Private configMgr As Common.ConfigurationManager
    Private configBannersettings As Banner = Nothing
    Private IOManager As Sympraxis.Common.IOManager.IOManager
    Private PluginDetail As PluginDetails
    Private mpanelcontainer As LayoutSetting.panelcontainer
    Private WindowsFormsHost As New System.Windows.Forms.Integration.WindowsFormsHost

    Private objSInstruction As New Instructions
    Private objDInstruction As New Instructions
    Private objCInstruction As New Instructions

    Private IHErrException As System.Exception
    Private Errortype As SympraxisException.ErrorTypes
    Private IHErrormsg As String
    Private IHErrormsgDetails As String
    Private IHErrormsgHeader As String
    Private IHErrorCode As String
    Private IHErrorFooter As String

    Dim OfflinePath As String
    'Dim tempappstart As String

    'Timer

    Private DITimer As Timers.Timer
    Private DICloseTimer As Timers.Timer
    Private CITimer As Timers.Timer
    Private CICloseTimer As Timers.Timer
    Private SITimer As Timers.Timer
    Private SICloseTimer As Timers.Timer


    'WithEvents
    Private WithEvents mLayoutmanager As New Sympraxis.Utilities.LayoutManager.LayoutManager
    Private WithEvents KeyboardHook As Sympraxis.Utilities.KeyboardHook.KeyboardHook
    Private WithEvents DAXObject As Sympraxis.Utilities.DataXChange.DataXchange



    'String Variable
    Private intrHostProcessName As String
    Private intrHostWorkFlowName As String
    Private delegatorUser As String
    Private processType As String
    Private getworkType() As String
    Private workingMode As String
    Private appStartBy As String
    Private cmdprofile As String
    Private userloginID As String
    Private NotifyMessageObjectId As String = ""
    Private lastLogger As String = ""
    Private _SystemStTime As String = ""

    Private dtInterval As DateTime = New DateTime(1, 1, 1, 0, 0, 0)
    Private loginHourTimer As New System.Timers.Timer
    Private IdleHourTimer As New System.Timers.Timer
    Private WorkHourTimer As New System.Timers.Timer

    'Integer Variable

    Private logininterval As Integer = 0
    Private idleinterval As Integer = 0
    Private CurrIdleInterval As Integer = 0
    Private openworkinterval As Integer = 0
    Dim ActiveWImodeIndex As Integer
    Dim ActiveWImodeplugin As Object


    Private waitCursorSec As Integer = 1
    Private goToBack As Integer = 0

    'Double Variable
    Private DIopacity As Double = 1

    'Boolean Variable
    Dim ActiveWIMode As Boolean = False

    Private BlnClosework As Boolean = True
    Private blnidlestart As Boolean = False
    Private BlnOpenAftClose As Boolean = False
    Private multiWindow As Boolean = False
    Private bIdleTime As Boolean = False
    Private IsAbortValidation As Boolean = False
    Private IsCaseAttended As Boolean = False
    Private ShowDI As Boolean = False
    Private FocusDI As Boolean = False
    Private SIopacity As Double = 1
    Private ShowSI As Boolean = False
    Private FocusSI As Boolean = False
    Private CIopacity As Double = 1
    Private ShowCI As Boolean = False
    Private FocusCI As Boolean = False
    Private DIMinimize As Boolean = False
    Private SIMinimize As Boolean = False
    Private CIMinimize As Boolean = False
    Private DIPin As Boolean = False
    Private SIPin As Boolean = False
    Private CIPin As Boolean = False
    Private ActiveWI As Boolean = False
    Private LValue As Boolean = False
    Private IsObjectautoclose As Boolean = False
    'HashTables

    Private hshTempDp As Hashtable
    Private hshDataPlugins As New Hashtable
    Private hshGetPluginTitle As New Hashtable
    Private hshEnableDefault As New Hashtable
    Private HiddenPlugin As New Hashtable
    Private HshGetDefaultsWOTPA As New Hashtable
    Private Hshtemp As Hashtable
    Private SubHshtemp As Hashtable
    Private hshDataParts As Hashtable
    Private HshPluginObj As New Hashtable

    Private hshDataPartsName As New Hashtable
    'timer

    Private ManualIdleTimer As System.Timers.Timer
    Private IdleWaitTimer As System.Timers.Timer
    Private HideInfobarTimer As System.Timers.Timer
    Private AppAutoCloseTimer As System.Timers.Timer
    Private AutoCloseIdleTimer As New System.Timers.Timer
    Dim AutoCloseIdleStarttime, AutoCloseIdleEndtime As DateTime

    'Datetime
    Private AutoCloseAHIdleStarttime As DateTime
    Private ManualIdleStarttime As DateTime


    'List
    Private defaultLayouts As List(Of Integer), EnablePlugin As New List(Of Integer)
    Private lstPluginOrder As List(Of Object)
    Private lstStartupPluginOrder As New List(Of Object)
    Private PluginMenu As List(Of Object)

    'Dictionary Variable 
    Private DicKeytip As Dictionary(Of String, String)
    Private DicSubMenuKeytip As Dictionary(Of String, String)
    Private PluginMenuEnableDisable As New Dictionary(Of String, String)
    Private fullDP As Dictionary(Of String, String)

    Dim ChildFullDPValue As Dictionary(Of String, Dictionary(Of String, String))
    Private SetWIModefullDP As Dictionary(Of String, String)
    'Delegates
    Private Delegate Sub ApplyLayoutDelegate(index As Integer)

    Private Delegate Sub DelEnableNotificationPanel(input As System.Xml.XmlDocument, size As Integer)
    Public Delegate Sub DelShowNotificationbar(message As String, notify As System.Xml.XmlDocument, ByVal size As Integer)
    Public Delegate Sub DelCloseWork(ByRef iserror As Boolean)
    Public Delegate Sub DelExiting(ByRef iserror As Boolean)
    Public Delegate Sub DelOpenWork(Values As String, GetObjectTypeby As [Enum], ByRef irc As Integer, ByRef Exception As String)
    Public Delegate Sub DelEnbwaitcur(val As Boolean)

    Public Delegate Sub DelOpenWorkshorcut()

    Private objWinZip As WinZip

    Dim MethodBase As System.Reflection.MethodBase
#End Region

#Region "PublicVariable"
    Public WithEvents Logger As Sympraxis.Utilities.logger.Logstatus
#End Region

#End Region

#Region "Constructor"
    Public Sub New()

        Dim ErrorMsg As String = String.Empty

        Try
            InitializeComponent()
            System.Windows.Forms.Application.EnableVisualStyles()
            Me.Topmost = True
            DAXObject = Sympraxis.Utilities.DataXChange.DataXchange.GetInstance()
            Logger = Sympraxis.Utilities.logger.Logstatus.GetInstance()
            IOManager = Common.IOManager.IOManager.GetInstance
            Dim ExpManager As ExceptionManager = ExceptionManager.GetInstance()

            AddHandler NetworkChange.NetworkAvailabilityChanged, AddressOf AvailabilityChanged

            _SystemStTime = Now.ToUniversalTime

            AddHandler ExpManager.HandledFormMessage, AddressOf HandleFormMessage
            AddHandler ExpManager.HandledFormMessages, AddressOf HandleFormMessages
            ObjFrmException = SympraxisException.GetInstance()
            Me.WindowStartupLocation = WindowStartupLocation.CenterScreen
            Me.WindowState = WindowState.Maximized

            ToolstripError.Visibility = System.Windows.Visibility.Collapsed

            'Throw (New Exception("Zghfgh"))
            'Throw (New SympraxisException.SettingsException("Zero Height found"))

            AddHandler ObjFrmException.HandledException, AddressOf SetupErrorMsg
            AddHandler WorkHourTimer.Elapsed, AddressOf WorkHourTimer_Tick
            AddHandler IdleHourTimer.Elapsed, AddressOf IdleHourTimer_Tick
            AddHandler loginHourTimer.Elapsed, AddressOf loginHourTimer_Tick
            KeyboardHook = Sympraxis.Utilities.KeyboardHook.KeyboardHook.GetInstance()
            If KeyboardHook.KeyCodetable.Contains(System.Windows.Forms.Keys.Escape) = False Then
                KeyboardHook.KeyCodetable.Add(System.Windows.Forms.Keys.Escape, Me)
            End If

            logininterval = -1
            idleinterval = -1
            openworkinterval = -1

            If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.LGTimer) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.LGTimer.ToUpper = "TRUE" Then
                logininterval = 0
            End If
            If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.IDTimer) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.IDTimer.ToUpper = "TRUE" Then
                idleinterval = 0
            End If
            If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.WPTimer) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.WPTimer.ToUpper = "TRUE" Then
                openworkinterval = 0
            End If

            Logger.IslogEnabled = False
            Logger.WriteToLogFile("IHost", "Load Application Start")
            LoadApp()


            Me.BringIntoView()


            Logger.WriteToLogFile("IHost", "Load Application Completed")
            pnlStadInstruction.IsOpen = False
            pnlDynamicInstruction.IsOpen = False
            pnlContextInstruction.IsOpen = False


            Me.Activate()
            Me.Focus()
            '   Me.Topmost = False

        Catch ex As SympraxisException.SettingsException


            ShowException(ex)

            Environment.Exit(-1)
        Catch ex As Exception
            ShowException(ex)
            Environment.Exit(-1)



        Finally

        End Try
    End Sub
#End Region

#Region "UIEventHandle"
    Dim Prvs As Boolean = False
    'Ribbon Menu Button Click Event
    Private Sub ToggleClickEvent(sender As Object, e As System.EventArgs)
        Dim Irc As Integer = 0
        Dim Errmsg As String = ""
        Dim Obj As Object = Nothing

        If TypeOf (sender) Is RibbonToggleButton Then

            Obj = DirectCast(sender, RibbonToggleButton).Tag
            If Not IsNothing(Obj) Then
                If Obj.ToString().ToLower() = "offline" Then
                    If CheckNetworkisUp() = True Then
                        workingMode = "normal"

                        InstanceOnline()

                        OffileAutoOpen()

                        configMgr.AppbaseConfig.EnvironmentVariables.IsOnline = True
                        ToggleChange(True)
                        BtnOffline.Visibility = System.Windows.Visibility.Collapsed
                        DAXObject.SetValue("SYS_REFRESH_REFERRAL", True)
                    Else
                        ShowErrorToolstrip("NetWork is not connected,We cannot Go Online", SympraxisException.ErrorTypes.Information, False)
                    End If

                ElseIf Obj.ToString().ToLower() = "online" Then




                    If configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True Then
                        configMgr.AppbaseConfig.EnvironmentVariables.IsOnline = False
                        workingMode = "Offline"
                        DAXObject.SetValue("SYSOFFLINE", "True")
                        RefreshOfflineGrid()
                    ElseIf configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = False Then
                        If CheckOfflineDataAvailable() = True Then
                            RefreshOfflineGrid()
                        Else
                            ShowErrorToolstrip("No Offline Data, Cannot go Offline ", SympraxisException.ErrorTypes.Information, False)
                        End If
                    End If
                End If


            End If
        End If
    End Sub

    Private Sub CallClickEvent(sender As Object, e As System.EventArgs)
        Dim Irc As Integer = 0

        Dim Obj As Object = Nothing

        Me.Topmost = False
        Try
            If TypeOf (sender) Is RibbonButton Then

                Obj = DirectCast(sender, RibbonButton).Tag

                If Not IsNothing(Obj) Then
                    If Obj.ToString().ToLower() = "open" Then
                        AHMenuOpenWork()
                        Menuhandle()
                    ElseIf Obj.ToString().ToLower() = "close" Then
                        AHMenuCloseWork(False)
                        '   Menuhandle()
                    ElseIf Obj.ToString().ToLower() = "next" Then
                        AHMenuNextWork(False)
                        Menuhandle()
                    ElseIf Obj.ToString().ToLower() = "previous" Then
                        AHMenuPreviousWork(False)
                        Prvs = False
                        Menuhandle()
                    ElseIf Obj.ToString().ToLower() = "exit" Then
                        Dim bhandle As Boolean = False

                        AHMenuExit(bhandle)
                    ElseIf Obj.ToString().ToLower() = "about" Then
                        Aboutbox()

                    ElseIf Obj.ToString().ToLower() = "shortcuts" Then
                        SetHelpfile()
                    ElseIf Obj.ToString().ToLower() = "workinstruction" Then
                        If Not IsNothing(configMgr) AndAlso Not IsNothing(configMgr.AppbaseConfig) AndAlso Not IsNothing(configMgr.AppbaseConfig.EnvironmentVariables) AndAlso Not IsNothing(configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState) AndAlso configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True Then

                            LoadWorkInstructions()
                        End If

                    Else
                        If Not IsNothing(PluginMenu) Then
                            Dim Getmenustrip As ToolStripMenuItem
                            If PluginMenu.Count > 0 Then
                                Getmenustrip = PluginMenu.Cast(Of ToolStripMenuItem).Where(Function(x) x.Text.ToLower() = Obj.ToString().ToLower()).FirstOrDefault()
                                If Not IsNothing(Getmenustrip) Then
                                    Getmenustrip.PerformClick()
                                Else
                                    Irc = 1
                                    Throw New Exception("not added PluginMenu list so cannot perform Click")

                                End If
                            End If
                        Else
                            Irc = 1
                            Throw New Exception("PluginMenu list is empty cannot perform Click ")

                        End If
                    End If
                Else
                    Throw New Exception(String.Concat("RibbonMenu", Convert.ToString(Obj), " Tag is empty cannot perform Click "))

                End If
            End If


        Catch ex As Exception
            SetupErrorMsg(ex)
        Finally
            Prvs = False
            ShowWaitCursor(False)

        End Try
    End Sub
    Dim version As String
    Dim ObjAbout As Sympraxis.Utilities.AboutBox.frmAboutBox
    Private Sub Aboutbox()


        Try

            ObjAbout = New Sympraxis.Utilities.AboutBox.frmAboutBox
            ObjAbout.Text = My.Application.Info.AssemblyName
            ' ObjAbout.lblVerLogo.Text = "v6.1"

            Dim assembly As Reflection.Assembly = Reflection.Assembly.GetEntryAssembly()
            Dim fvi As FileVersionInfo = FileVersionInfo.GetVersionInfo(assembly.Location)
            version = fvi.FileVersion
            aboutTool(ObjAbout)
            Dim ReleaseDate As String = ""
            Dim yearstring As String = ""
            Dim Monthstring As String = ""
            Dim DayString As String = ""
            Dim lastindexofdot As Integer = -1

            ''     version = Replace(version, My.Application.Info.Version.Major.ToString & "." & My.Application.Info.Version.Minor.ToString & ".", "")
            version = Replace(version, fvi.FileMajorPart.ToString & "." & fvi.FileMinorPart.ToString & ".", "")
            ' About toolstrip Menu item 

            ' About toolstrip Menu item 
            lastindexofdot = version.LastIndexOf(".")
            If lastindexofdot > -1 Then
                version = version.Substring(0, lastindexofdot)
                If version.Trim.Length = 8 Then
                    yearstring = version.Substring(0, 4)
                    Monthstring = version.Substring(4, 2)
                    DayString = version.Substring(6, 2)
                    ReleaseDate = (DayString & "/" & Monthstring & "/" & yearstring)
                    'ObjAbout.SympraxisAbout1.txtLastUpdate.Text = ReleaseDate

                End If
            End If

            Dim Plugindetails As PluginDefinition

            Dim strLoadPluginName As String
            ObjAbout.Plglist.Rows.Clear()
            For i As Integer = 0 To loadPluginsConfigXml.ApplicationHost.Plugins.Count - 1
                PluginDetail = New PluginDetails
                If IsNothing(loadPluginsConfigXml.ApplicationHost.Plugins(i).Name) Or (Not IsNothing(loadPluginsConfigXml.ApplicationHost.Plugins(i).Name) AndAlso loadPluginsConfigXml.ApplicationHost.Plugins(i).Name.Trim() = String.Empty) Then
                    Throw New SympraxisException.SettingsException("Missing plugin configuration details (Plugin name).")
                Else
                    strLoadPluginName = loadPluginsConfigXml.ApplicationHost.Plugins(i).Name

                End If



                Plugindetails = appHostConfigXml.PluginDefinitions.PluginDefinition.Where(Function(x) x.Name = strLoadPluginName).FirstOrDefault()
                If IsNothing(Plugindetails) Then
                    Throw New SympraxisException.SettingsException(String.Concat("Unable to find the plugin defined in the defintion for process.[", intrHostProcessName, "] [", strLoadPluginName, "]"))
                End If

                If Not IsNothing(Plugindetails) Then
                    ObjAbout.Plglist.Rows.Add(Plugindetails.Name.ToString(), Plugindetails.Assembly.ToString(), version, ReleaseDate, Plugindetails.ClassName.ToString())
                End If
            Next




            ObjAbout.ShowDialog()
        Catch ex As Exception

        Finally

        End Try
    End Sub



    Private Sub aboutTool(ByVal sender As Sympraxis.Utilities.AboutBox.frmAboutBox)
        Try
            ObjAbout.lblVerLogo.Text = "6.0 R2"
            ObjAbout.ViWF.Text = configMgr.AppbaseConfig.EnvironmentVariables.WFName
            ObjAbout.ViTrans.Text = configMgr.AppbaseConfig.TransDBURL
            If Not configMgr.AppbaseConfig.InstMgr Is Nothing Then
                ObjAbout.ViInsmgr.Text = configMgr.AppbaseConfig.InstMgr.IMURL
            End If
            ObjAbout.ViUsr.Text = configMgr.AppbaseConfig.EnvironmentVariables.sUserId
            ObjAbout.ViProcess.Text = configMgr.AppbaseConfig.EnvironmentVariables.CurrentProcessName
            ObjAbout.ViMc.Text = System.Net.Dns.GetHostName
            ObjAbout.ViUsr.Text = configMgr.AppbaseConfig.EnvironmentVariables.UserName
            ObjAbout.ViDt.Text = _SystemStTime
            ' ObjAbout.ABPWD = _ABPWD
            If Not configMgr.AppbaseConfig.EnvironmentVariables.myLoginRsp Is Nothing Then
                If configMgr.AppbaseConfig.EnvironmentVariables.myLoginRsp.stUserSession.stProcessingUser.sUserId = configMgr.AppbaseConfig.EnvironmentVariables.myLoginRsp.stUserSession.sUserId Then
                    ObjAbout.ViUsr.Text = configMgr.AppbaseConfig.EnvironmentVariables.myLoginRsp.stUserSession.sUserId & " (" & configMgr.AppbaseConfig.EnvironmentVariables.myLoginRsp.stUserSession.sUserName & " )"
                Else
                    ObjAbout.ViUsr.Text = configMgr.AppbaseConfig.EnvironmentVariables.myLoginRsp.stUserSession.stProcessingUser.sUserId
                End If

                If Not configMgr.AppbaseConfig.EnvironmentVariables.myLoginRsp.stUserSession Is Nothing Then
                    ObjAbout.lblVwSession.Text = configMgr.AppbaseConfig.EnvironmentVariables.myLoginRsp.stUserSession.iSessionId
                End If
                'If Not configMgr.AppbaseConfig.EnvironmentVariables.myLoginRsp.stWorkSession Is Nothing Then
                '    ObjAbout.lblVwSession.Text = configMgr.AppbaseConfig.EnvironmentVariables.myLoginRsp.stUserSession.iSessionId & "(U)/(W)" & configMgr.AppbaseConfig.EnvironmentVariables.myLoginRsp.stWorkSession(0).lSessionId
                'End If
            End If
            ObjAbout.ViMode.Text = "Direct"
            'ObjAbout.ViDt.Text = Now.ToUniversalTime
            ObjAbout.ViDt.Text = DateTime.UtcNow.ToString("MM/dd/yyyy HH:mm:ss") & " (UTC)"
            'ObjAbout.NewSympraxisAbout1.Txtmodule.Text = version
            Dim Configuration As String
            Configuration = "V R2"
            Dim strg As Process
            strg = Process.GetCurrentProcess
            Dim str As String
            str = String.Format(System.Math.Round(My.Computer.Info.AvailablePhysicalMemory / (1024 * 1024)), 2).ToString

            Dim str1 As String
            str1 = String.Format(System.Math.Round(My.Computer.Info.TotalPhysicalMemory / (1024 * 1024)), 2).ToString
            Dim str3 As String
            str3 = "Total" & Space(2) & str & "/" & str1
            ObjAbout.ViCPU.Text = GetMemoryUsage(strg.ProcessName) & Space(2) & str3
            Dim AppCPU As New PerformanceCounter("Process", "% Processor Time", strg.ProcessName, True)
            ObjAbout.ViMry.Text = Math.Round(AppCPU.NextValue())
            System.Threading.Thread.Sleep(100)
            ObjAbout.vicpuMry.Text = Math.Round(AppCPU.NextValue()) & " %"


        Catch ex As Exception
            'AppBaseHelper.ExpManager.ProcessException(ex, ErrorMessage)
        End Try
    End Sub
    Public Shared Function GetMemoryUsage(ByVal ProcessName As String) As String
        Dim _Process As Process = Nothing

        Dim _Return As String = ""

        For Each _Process In Process.GetProcessesByName(ProcessName)

            If _Process.ToString.Remove(0, 27).ToLower = "(" & ProcessName.ToLower & ")" Then

                _Return = (_Process.WorkingSet64 / 1024).ToString("0,000") & " K"
            End If
        Next

        If Not _Process Is Nothing Then

            _Process.Dispose()

            _Process = Nothing

        End If

        Return _Return

    End Function
    'Error toolstrip close button click
    Private Sub ButtonToolstripError_Click(sender As Object, e As RoutedEventArgs)
        Try
            ToolstripError.Visibility = System.Windows.Visibility.Collapsed

        Catch ex As Exception

            SetupErrorMsg(ex)
        End Try
    End Sub

    Private Function CheckOfflineDataAvailable() As Boolean
        Dim Rt As Boolean = False
        Dim Xdoc As XmlDocument = Nothing

        If IOManager.CheckFileExists(appHostConfigXml.OfflinePath & "\OfflineIndex.SXML") Then
            Xdoc = New XmlDocument
            Xdoc.Load(appHostConfigXml.OfflinePath & "\OfflineIndex.SXML")
            If Not IsNothing(Xdoc.SelectSingleNode("//RL")) Then
                Rt = True
            End If
        End If
        Return Rt
    End Function
    Private Sub RefreshOfflineGrid()
        Dim Xdoc As XmlDocument = Nothing

        If IOManager.CheckFileExists(appHostConfigXml.OfflinePath & "\OfflineIndex.SXML") Then
            Xdoc = New XmlDocument
            Xdoc.Load(appHostConfigXml.OfflinePath & "\OfflineIndex.SXML")
            If Not IsNothing(Xdoc.SelectSingleNode("//RL")) Then

                DAXObject.SetValue("SYSREFDATA", Xdoc)
                Me.Dispatcher.Invoke(Sub() ToggleChange(False))
                BtnOffline.Visibility = System.Windows.Visibility.Visible
            Else
                ShowErrorToolstrip("No Offline Data", SympraxisException.ErrorTypes.Information, False)
            End If
        Else
            ShowErrorToolstrip("Offline Data Path not Exist in" & appHostConfigXml.OfflinePath & "\OfflineIndex.SXML", SympraxisException.ErrorTypes.Information, False)
        End If
    End Sub

    Private Sub AvailabilityChanged(ByVal sender As Object, ByVal e As NetworkAvailabilityEventArgs)
        If e.IsAvailable Then

            MessageBox.Show("Network connected!")
        Else


            configMgr.AppbaseConfig.EnvironmentVariables.IsOnline = False
            workingMode = "offline"
            Me.Dispatcher.Invoke(Sub() RefreshOfflineGrid())
        End If
    End Sub

    'Keyboard hook keydown event
    Private Sub KeyboardHook_KeyDown(Key As Keys, IsControlPressed As Boolean, IsAltPressed As Boolean, IsShiftPressed As Boolean) Handles KeyboardHook.KeyDown

        Try
            Dim keycode As Integer = Key
            If Not IsNothing(IdleWaitTimer) Then
                IdleWaitTimer.Enabled = False

            End If
            If Not IsNothing(IdleHourTimer) Then
                IdleHourTimer.Enabled = False
                IdleHourTimer.Stop()
            End If


            AutoCloseAHIdleStarttime = DateTime.Now
            AutoCloseIdleStarttime = DateTime.Now
            ManualIdleStarttime = DateTime.Now
            If bIdleTime = True Then

                RaiseEvent IdleTimeEnd(DateTime.UtcNow)
                bIdleTime = False
            End If
            If Not IsNothing(IdleWaitTimer) Then
                IdleWaitTimer.Enabled = True

            End If
            If Key = Keys.Escape Then
                If FocusDI = True Then
                    btnDIClose_Click(Nothing, Nothing)
                End If
                If FocusSI = True Then
                    btnSIClose_Click(Nothing, Nothing)
                End If
                If FocusCI = True Then
                    btnCIClose_Click(Nothing, Nothing)
                End If
                pnlNotifyHeader.Visibility = System.Windows.Visibility.Collapsed
                ToolstripError.Visibility = System.Windows.Visibility.Collapsed
            Else
                If WIConfig IsNot Nothing AndAlso WIConfig.DynamicInstruction IsNot Nothing AndAlso WIConfig.DynamicInstruction.iSKey <> String.Empty Then
                    Dim str As String = String.Empty
                    Dim ChValue1 As Integer
                    str = splitstring(WIConfig.DynamicInstruction.iSKey)
                    Dim KeyConverter As New System.Windows.Forms.KeysConverter
                    Dim S As String = KeyConverter.ConvertToString(str)
                    Dim O As System.Windows.Forms.Keys = KeyConverter.ConvertFrom(S)
                    ChValue1 = CType(O, Integer)
                    If ChValue1 = Key Then
                        Dim convertedkeycode As String = keycodestring(Key, IsControlPressed, IsShiftPressed, IsAltPressed)
                        If convertedkeycode = WIConfig.DynamicInstruction.iSKey Then
                            LoadDynamicInstruction()
                            If DIDataGrid.Items.Count > 0 Then
                                FocusDI = True
                                pnlDynamicInstruction.Focus()
                                DIDataGrid.Focus()
                                DIDataGrid.SelectedItem = DIDataGrid.Items(0)
                                DIDataGrid.CurrentCell = New DataGridCellInfo(DIDataGrid.SelectedItem, DIDataGrid.Columns(0))
                            End If
                        End If
                    End If
                End If
                If WIConfig IsNot Nothing AndAlso WIConfig.StandardInstruction IsNot Nothing AndAlso WIConfig.StandardInstruction.iSKey <> String.Empty Then
                    Dim str As String = String.Empty
                    Dim ChValue1 As Integer
                    str = splitstring(WIConfig.StandardInstruction.iSKey)
                    Dim KeyConverter As New System.Windows.Forms.KeysConverter
                    Dim S As String = KeyConverter.ConvertToString(str)
                    Dim O As System.Windows.Forms.Keys = KeyConverter.ConvertFrom(S)
                    ChValue1 = CType(O, Integer)
                    If ChValue1 = Key Then
                        Dim convertedkeycode As String = keycodestring(Key, IsControlPressed, IsShiftPressed, IsAltPressed)
                        If convertedkeycode = WIConfig.StandardInstruction.iSKey Then
                            LoadStandardInstrction()
                            If SIDataGrid.Items.Count > 0 Then
                                FocusSI = True
                                pnlStadInstruction.Focus()
                                SIDataGrid.Focus()
                                SIDataGrid.SelectedItem = SIDataGrid.Items(0)
                                SIDataGrid.CurrentCell = New DataGridCellInfo(SIDataGrid.SelectedItem, SIDataGrid.Columns(0))
                            End If
                        End If
                    End If
                End If
                If WIConfig IsNot Nothing AndAlso WIConfig.ContextInstruction IsNot Nothing AndAlso WIConfig.ContextInstruction.iSKey <> String.Empty Then
                    Dim str As String = String.Empty
                    Dim ChValue1 As Integer
                    str = splitstring(WIConfig.ContextInstruction.iSKey)
                    Dim KeyConverter As New System.Windows.Forms.KeysConverter
                    Dim S As String = KeyConverter.ConvertToString(str)
                    Dim O As System.Windows.Forms.Keys = KeyConverter.ConvertFrom(S)
                    ChValue1 = CType(O, Integer)
                    If ChValue1 = Key Then
                        Dim convertedkeycode As String = keycodestring(Key, IsControlPressed, IsShiftPressed, IsAltPressed)
                        If convertedkeycode = WIConfig.ContextInstruction.iSKey Then
                            LoadContextInstrction()
                            If CIDataGrid.Items.Count > 0 Then
                                FocusCI = True
                                pnlContextInstruction.Focus()
                                CIDataGrid.Focus()
                                CIDataGrid.SelectedItem = CIDataGrid.Items(0)
                                CIDataGrid.CurrentCell = New DataGridCellInfo(CIDataGrid.SelectedItem, CIDataGrid.Columns(0))
                            End If
                        End If
                    End If
                End If
            End If

            If IMClientPort <> "" Then
                If IsControlPressed = True And (keycode = 48 Or keycode = 96) Then

                    ObjClientInstanceMgr.SetFocustoWD()
                End If
            End If

        Catch ex As Exception


            SetupErrorMsg(ex)
        End Try
    End Sub

    'Public HostActive As Boolean = False

    Private Sub BtnNotificationClose_Click(sender As Object, e As RoutedEventArgs)
        Try

            pnlNotifyHeader.Visibility = System.Windows.Visibility.Collapsed

        Catch ex As Exception


            SetupErrorMsg(ex)
        End Try
    End Sub

    'Timer  autoclose for Unattentedpin timmer and attendedPin timmer
    Private Sub ButtonPin_Click(sender As Object, e As MouseButtonEventArgs)
        Try
            If txtACStatus.Text = "OFF" Then
                IsCaseAttended = True
                txtACStatus.Text = "ON"
                txtAutoClose.Foreground = New SolidColorBrush(Colors.Red)
                txtACStatus.Foreground = New SolidColorBrush(Colors.Red)
            Else
                IsCaseAttended = False
                txtACStatus.Text = "OFF"
                txtAutoClose.Foreground = New SolidColorBrush(Colors.DarkBlue)
                txtACStatus.Foreground = New SolidColorBrush(Colors.DarkBlue)
            End If
            ManualIdleStarttime = DateTime.Now
        Catch ex As Exception
            SetupErrorMsg(ex)
        End Try
    End Sub


    Private Sub ButtonExport_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs)
        Try
            If Not IsNothing(IHErrException) Then
                ShowException(IHErrException)
            End If
        Catch ex As Exception

            SetupErrorMsg(ex)
        End Try
    End Sub

    Private Sub btnOpenNext_Click(sender As Object, e As MouseButtonEventArgs)
        Try


            If configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True Then
                AHMenuCloseWork(False)
                If configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = False Then
                    AHMenuOpenWork()
                End If
            Else

                AHMenuOpenWork()

            End If



        Catch ex As Exception

            SetupErrorMsg(ex)
        End Try
    End Sub


    Private Sub CIStyle_MouseEnter(sender As Object, e As Input.MouseEventArgs)
        Try
            If CIPin = False AndAlso CITimer IsNot Nothing Then
                If Not CICloseTimer Is Nothing Then
                    If CICloseTimer.Enabled = True Then
                        CICloseTimer.Enabled = False
                        CICloseTimer.Stop()
                        CIStyle.Opacity = 1
                        CIopacity = 1
                    End If
                End If
                If CITimer.Enabled = True Then
                    CITimer.Enabled = False
                End If
            End If
        Catch ex As Exception
            SetupErrorMsg(ex)

        End Try
    End Sub

    Private Sub CIStyle_MouseLeave(sender As Object, e As Input.MouseEventArgs)
        Try
            If CIPin = False AndAlso CITimer IsNot Nothing Then
                If CITimer.Enabled = False Then
                    CITimer.Enabled = True
                End If
            End If
        Catch ex As Exception
            SetupErrorMsg(ex)

        End Try
    End Sub

    Private Sub CIStyle_LostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        Try


            If CIPin = False AndAlso CITimer IsNot Nothing Then
                If CITimer.Enabled = False Then
                    CITimer.Enabled = True
                End If
            End If
            FocusCI = False
        Catch ex As Exception
            SetupErrorMsg(ex)

        End Try
    End Sub

    Private Sub CIStyle_GotKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        Try

            If CIPin = False AndAlso CITimer IsNot Nothing Then
                If Not CICloseTimer Is Nothing Then
                    If CICloseTimer.Enabled = True Then
                        CICloseTimer.Enabled = False
                        CICloseTimer.Stop()
                        CIStyle.Opacity = 1
                        CIopacity = 1
                    End If
                End If
                If CITimer.Enabled = True Then
                    CITimer.Enabled = False
                End If
            End If
            FocusCI = True
        Catch ex As Exception
            SetupErrorMsg(ex)

        End Try
    End Sub

    Private Sub btnCIClose_Click(sender As Object, e As RoutedEventArgs)
        Try
            If CITimer IsNot Nothing Then
                If CITimer IsNot Nothing Then
                    CITimer.Enabled = False
                    CITimer.Stop()
                End If
                If CICloseTimer IsNot Nothing Then
                    CICloseTimer.Enabled = False
                    CICloseTimer.Stop()
                End If
            End If
            pnlContextInstruction.IsOpen = False
            ShowCI = False
            FocusCI = False
            SetInstructionLocation(Nothing)
        Catch ex As Exception

            SetupErrorMsg(ex)

        End Try
    End Sub

    Private Sub btnCIMinimize_Click(sender As Object, e As RoutedEventArgs)
        Try
            CIMinimize = True
            btnCIClose_Click(Nothing, Nothing)

        Catch ex As Exception

            SetupErrorMsg(ex)

        End Try
    End Sub

    Private Sub CIndicate_Click(sender As Object, e As RoutedEventArgs)
        Dim irc As Integer = 0


        Try
            LoadContextInstrction()
            If CIDataGrid.Items.Count > 0 Then
                FocusCI = True

                pnlContextInstruction.Focus()
                CIDataGrid.Focus()
                CIDataGrid.SelectedItem = CIDataGrid.Items(0)
                CIDataGrid.CurrentCell = New DataGridCellInfo(CIDataGrid.SelectedItem, CIDataGrid.Columns(0))
            End If
        Catch ex As Exception

            SetupErrorMsg(ex)

        Finally

        End Try
    End Sub

    Private Sub pinCI_Click(sender As Object, e As RoutedEventArgs)
        Try
            If CIPin = False Then
                CIPin = True
                pinCI.Visibility = System.Windows.Visibility.Visible
                unpinCI.Visibility = System.Windows.Visibility.Collapsed
                If CITimer IsNot Nothing Then
                    CITimer.Enabled = False
                    CITimer.Stop()
                End If
                If CICloseTimer IsNot Nothing Then
                    CICloseTimer.Enabled = False
                    CICloseTimer.Stop()
                End If
                CIStyle.Opacity = 1
            Else
                CIPin = False
                pinCI.Visibility = System.Windows.Visibility.Collapsed
                unpinCI.Visibility = System.Windows.Visibility.Visible
                If CITimer IsNot Nothing Then
                    CITimer.Enabled = True
                End If
            End If
        Catch ex As Exception
            SetupErrorMsg(ex)


        End Try
    End Sub



    Private Sub SIStyle_MouseEnter(sender As Object, e As Input.MouseEventArgs)
        Try


            If SIPin = False AndAlso SITimer IsNot Nothing Then
                If Not SICloseTimer Is Nothing Then
                    If SICloseTimer.Enabled = True Then
                        SICloseTimer.Enabled = False
                        SICloseTimer.Stop()
                        SIStyle.Opacity = 1
                        SIopacity = 1
                    End If
                End If
                If SITimer.Enabled = True Then
                    SITimer.Enabled = False
                End If
            End If
        Catch ex As Exception
            SetupErrorMsg(ex)

        End Try
    End Sub

    Private Sub SIStyle_MouseLeave(sender As Object, e As Input.MouseEventArgs)
        Try



            If SIPin = False AndAlso SITimer IsNot Nothing Then
                If SITimer.Enabled = False Then
                    SITimer.Enabled = True
                End If
            End If
        Catch ex As Exception

            SetupErrorMsg(ex)

        End Try

    End Sub

    Private Sub SIStyle_LostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        Try


            If SIPin = False AndAlso SITimer IsNot Nothing Then
                If SITimer.Enabled = False Then
                    SITimer.Enabled = True
                End If
            End If
            FocusSI = False
        Catch ex As Exception

            SetupErrorMsg(ex)
        End Try
    End Sub

    Private Sub SIStyle_GotKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        Try


            If SIPin = False AndAlso SITimer IsNot Nothing Then
                If Not SICloseTimer Is Nothing Then
                    If SICloseTimer.Enabled = True Then
                        SICloseTimer.Enabled = False
                        SICloseTimer.Stop()
                        SIStyle.Opacity = 1
                        SIopacity = 1
                    End If
                End If
                If SITimer.Enabled = True Then
                    SITimer.Enabled = False
                End If
            End If
            FocusSI = True
        Catch ex As Exception
            SetupErrorMsg(ex)

        End Try
    End Sub

    Private Sub btnSIClose_Click(sender As Object, e As RoutedEventArgs)
        Try
            If SITimer IsNot Nothing Then
                If SITimer IsNot Nothing Then
                    SITimer.Enabled = False
                    SITimer.Stop()
                End If
                If SICloseTimer IsNot Nothing Then
                    SICloseTimer.Enabled = False
                    SICloseTimer.Stop()
                End If
            End If
            pnlStadInstruction.IsOpen = False

            ShowSI = False
            FocusSI = False
            SetInstructionLocation(Nothing)
        Catch ex As Exception

            SetupErrorMsg(ex)


        End Try
    End Sub

    Private Sub btnSIMinimize_Click(sender As Object, e As RoutedEventArgs)
        Try
            SIMinimize = True
            btnSIClose_Click(Nothing, Nothing)
        Catch ex As Exception
            SetupErrorMsg(ex)

        End Try
    End Sub

    Private Sub SIndicate_Click(sender As Object, e As RoutedEventArgs)
        Dim irc As Integer = 0


        Try
            LoadStandardInstrction()

            If irc <> 0 Then Exit Sub

            If SIDataGrid.Items.Count > 0 Then
                FocusSI = True
                pnlStadInstruction.Focus()
                SIDataGrid.Focus()
                SIDataGrid.SelectedItem = SIDataGrid.Items(0)
                SIDataGrid.CurrentCell = New DataGridCellInfo(SIDataGrid.SelectedItem, SIDataGrid.Columns(0))
            End If
        Catch ex As Exception

            SetupErrorMsg(ex)



        End Try
    End Sub

    Private Sub pinSI_Click(sender As Object, e As RoutedEventArgs)
        Try
            If SIPin = False Then
                SIPin = True
                pinSI.Visibility = System.Windows.Visibility.Visible
                unpinSI.Visibility = System.Windows.Visibility.Collapsed
                If SITimer IsNot Nothing Then
                    SITimer.Enabled = False
                    SITimer.Stop()
                End If
                If SICloseTimer IsNot Nothing Then
                    SICloseTimer.Enabled = False
                    SICloseTimer.Stop()
                End If

                SIStyle.Opacity = 1

            Else
                SIPin = False
                pinSI.Visibility = System.Windows.Visibility.Collapsed
                unpinSI.Visibility = System.Windows.Visibility.Visible
                If SITimer IsNot Nothing Then
                    SITimer.Enabled = True
                End If
            End If
        Catch ex As Exception
            SetupErrorMsg(ex)


        End Try
    End Sub


    Private Sub DICloseTimer_Elapsed(sender As Object, e As System.Timers.ElapsedEventArgs)
        Try
            If DIPin = False AndAlso Not DICloseTimer Is Nothing Then
                If DICloseTimer.Enabled = True Then
                    DIopacity = DIopacity - 0.05
                    DIStyle.Dispatcher.Invoke(Sub() DIStyle.Opacity = DIopacity)
                    If DIopacity = 0 Or DIopacity < 0 Then
                        ShowDI = False
                        DICloseTimer.Enabled = False
                        DICloseTimer.Stop()
                        DIopacity = 1
                        DITimer.Enabled = False
                        DITimer.Stop()
                        Me.Dispatcher.Invoke(Sub() SetInstructionLocation(Nothing))
                    End If
                End If
            End If
        Catch ex As Exception

            SetupErrorMsg(ex)


        End Try
    End Sub

    Private Sub DIStyle_MouseEnter(sender As Object, e As Input.MouseEventArgs)
        Try


            If DIPin = False AndAlso DITimer IsNot Nothing Then
                If Not DICloseTimer Is Nothing Then
                    If DICloseTimer.Enabled = True Then
                        DICloseTimer.Enabled = False
                        DICloseTimer.Stop()
                        DIStyle.Opacity = 1
                        DIopacity = 1
                    End If
                End If
                If DITimer.Enabled = True Then
                    DITimer.Enabled = False
                End If
            End If
        Catch ex As Exception
            SetupErrorMsg(ex)

        End Try
    End Sub

    Private Sub DIStyle_MouseLeave(sender As Object, e As Input.MouseEventArgs)
        Try


            If DIPin = False AndAlso DITimer IsNot Nothing Then
                If DITimer.Enabled = False Then
                    DITimer.Enabled = True
                End If
            End If
        Catch ex As Exception
            SetupErrorMsg(ex)

        End Try
    End Sub

    Private Sub DIStyle_LostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        Try


            If DIPin = False AndAlso DITimer IsNot Nothing Then
                If DITimer.Enabled = False Then
                    DITimer.Enabled = True
                End If
            End If
            FocusDI = False
        Catch ex As Exception
            SetupErrorMsg(ex)
        End Try
    End Sub

    Private Sub DIStyle_GotKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs)
        Try

            If DIPin = False AndAlso DITimer IsNot Nothing Then
                If Not DICloseTimer Is Nothing Then
                    If DICloseTimer.Enabled = True Then
                        DICloseTimer.Enabled = False
                        DICloseTimer.Stop()
                        DIStyle.Opacity = 1
                        DIopacity = 1
                    End If
                End If
                If DITimer.Enabled = True Then
                    DITimer.Enabled = False
                End If
            End If
            FocusDI = True

        Catch ex As Exception
            SetupErrorMsg(ex)

        End Try
    End Sub

    Private Sub btnDIClose_Click(sender As Object, e As RoutedEventArgs)
        Try
            If DITimer IsNot Nothing Then
                If DITimer IsNot Nothing Then
                    DITimer.Enabled = False
                    DITimer.Stop()
                End If
                If DICloseTimer IsNot Nothing Then
                    DICloseTimer.Enabled = False
                    DICloseTimer.Stop()
                End If
            End If
            pnlDynamicInstruction.IsOpen = False

            ShowDI = False
            FocusDI = False
            SetInstructionLocation(Nothing)
        Catch ex As Exception
            SetupErrorMsg(ex)

        End Try
    End Sub

    Private Sub btnDIMinimize_Click(sender As Object, e As RoutedEventArgs)
        Try
            DIMinimize = True
            btnDIClose_Click(Nothing, Nothing)
        Catch ex As Exception
            SetupErrorMsg(ex)
        End Try
    End Sub

    Private Sub DIndicate_Click(sender As Object, e As RoutedEventArgs)
        Dim irc As Integer = 0


        Try
            LoadDynamicInstruction()

            If irc <> 0 Then Exit Sub

            If DIDataGrid.Items.Count > 0 Then
                FocusDI = True
                pnlDynamicInstruction.Focus()
                DIDataGrid.Focus()
                DIDataGrid.SelectedItem = DIDataGrid.Items(0)
                DIDataGrid.CurrentCell = New DataGridCellInfo(DIDataGrid.SelectedItem, DIDataGrid.Columns(0))
            End If
        Catch ex As Exception
            SetupErrorMsg(ex)


        Finally

        End Try
    End Sub

    Private Sub pinDI_Click(sender As Object, e As RoutedEventArgs)
        Try
            If DIPin = False Then
                DIPin = True
                pinDI.Visibility = System.Windows.Visibility.Visible
                unpinDI.Visibility = System.Windows.Visibility.Collapsed
                If DITimer IsNot Nothing Then
                    DITimer.Enabled = False
                    DITimer.Stop()
                End If
                If DICloseTimer IsNot Nothing Then
                    DICloseTimer.Enabled = False
                    DICloseTimer.Stop()
                End If
                DIStyle.Opacity = 1
            Else
                DIPin = False
                pinDI.Visibility = System.Windows.Visibility.Collapsed
                unpinDI.Visibility = System.Windows.Visibility.Visible
                If DITimer IsNot Nothing Then
                    DITimer.Enabled = True
                End If
            End If
        Catch ex As Exception
            SetupErrorMsg(ex)

        End Try
    End Sub



    Private Sub BtnNotificationOpen_Click(sender As Object, e As RoutedEventArgs)
        Dim irc As Integer = 0
        Dim Errmsg As String = String.Empty
        pnlNotifyHeader.Visibility = System.Windows.Visibility.Collapsed

        Try
            If BtnNotificationOpen.Content = "Open" Then
                BlnOpenAftClose = False
                If configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True Then
                    DialogResult = MessageBox.Show("Already opened Workpacket will close,Are you sure want to proceed?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)
                    If DialogResult = True Then
                        AHMenuCloseWork(False)
                        If irc <> 0 Then Exit Sub
                    Else
                        Exit Sub
                    End If
                End If
                If Not IsNothing(WorkActivity) AndAlso WorkActivity.ToUpper() <> "WORK" Then


                    Throw New SympraxisException.InformationException("Please select valid activity(work) in timesheet")

                End If
                OpenWorkAcrossProcessAndWF(NotifyMessageObjectId, Sympraxis.Utilities.INormalPlugin.IPluginhost.GetObjectType.ObjectId, irc, Errmsg)
            ElseIf BtnNotificationOpen.Content = "Open Next" Then
                BlnOpenAftClose = True
            End If

        Catch ex As Exception
            SetupErrorMsg(ex)
        Finally

        End Try
    End Sub







#End Region

#Region "WindowEventHandle"
    Public Sub MouseMoveLM() Handles mLayoutmanager.LMMouseMove
        IdleWaitTimer.Enabled = False
        IdleHourTimer.Enabled = False
        IdleHourTimer.Stop()
        AutoCloseAHIdleStarttime = DateTime.Now
        IdleWaitTimer.Enabled = True
        ManualIdleStarttime = DateTime.Now
        AutoCloseIdleStarttime = DateTime.Now
        If bIdleTime = True Then

            RaiseEvent IdleTimeEnd(DateTime.UtcNow)
            bIdleTime = False
        End If
    End Sub
    Private Sub RibbonWindow_KeyDown(sender As Object, e As Input.KeyEventArgs)
        Try

            Me.Topmost = False
            IdleWaitTimer.Enabled = False
            IdleHourTimer.Enabled = False
            IdleHourTimer.Stop()
            AutoCloseAHIdleStarttime = DateTime.Now
            AutoCloseIdleStarttime = DateTime.Now
            IdleWaitTimer.Enabled = True
            ManualIdleStarttime = DateTime.Now
            If bIdleTime = True Then

                RaiseEvent IdleTimeEnd(DateTime.UtcNow)
                bIdleTime = False
            End If

        Catch ex As Exception
            SetupErrorMsg(ex)

        End Try
    End Sub
    Dim isformactive As Boolean = False

    Private Sub EnbWaitcur(ByVal flag As Boolean)
        If flag = True Then
            If iswaitcursoron = True Then
                Waitcursor.IsOpen = True
                Waitcursor.Placement = Controls.Primitives.PlacementMode.Center
            End If
        Else

            Waitcursor.IsOpen = False
        End If
       
    End Sub
    Private Sub RibbonWindow_Activated(sender As Object, e As EventArgs)
        Try
            isformactive = True
            Try
                Dispatcher.Invoke(CType(AddressOf EnbWaitcur, DelEnbwaitcur), True)

            Catch ex As Exception

            End Try

            KeyboardHook.KeyHookEnable = True
            DAXObject.SetValue("SYSACTIVATED", "TRUE")
            ShowInstruction(True)

        Catch ex As Exception


            SetupErrorMsg(ex)
        End Try
    End Sub

    Private Sub RibbonWindow_Deactivated(sender As Object, e As EventArgs)
        Try
            isformactive = False
            Try
                Dispatcher.Invoke(CType(AddressOf EnbWaitcur, DelEnbwaitcur), False)
            Catch ex As Exception

            End Try


            KeyboardHook.KeyHookEnable = False

            ShowInstruction(False)
            DAXObject.SetValue("SYSACTIVATED", "FALSE")
        Catch ex As Exception


            SetupErrorMsg(ex)

        End Try
    End Sub

    Private Sub RibbonWindow_StateChanged(sender As Object, e As EventArgs)

    End Sub


    Private Sub RibbonWindow_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs)



        Try
            Dim bhandle As Boolean = True
            AHMenuExit(bhandle)
            If bhandle = True Then
                e.Cancel = True
            End If
        Catch ex As Exception
            SetupErrorMsg(ex)
            e.Cancel = True

        End Try
    End Sub


    Private Function keycodestring(ByRef Key As System.Windows.Forms.Keys, ByRef IsControlPressed As Boolean, ByRef IsShiftPressed As Boolean, ByRef IsAltPressed As Boolean) As String
        Dim r As String = Nothing

        If IsControlPressed Then
            r = "^" '^ Indicates  Control Key
        End If
        If IsAltPressed Then
            r += "%" '% Indicates Alt Key
        End If
        If IsShiftPressed Then
            r += "_" '_ Indicates Shift Key
        End If
        r += Key.ToString

        Return r
    End Function


    Private Sub RibbonWindow_Loaded(sender As Object, e As RoutedEventArgs)
        Dim Irc As Integer = 0

        Try

            SendKeys("%{N}")
            IdleTimerStart()
            AppAutoCloseTimerStart()
            If logininterval <> -1 Then
                loginHourTimer.Interval = 1000
                loginHourTimer.Enabled = True

            End If
            WorkActivity = "WORK"
            RaiseEvent AppReady()



            '  Me.Activate()
        Catch ex As Exception


            SetupErrorMsg(ex)

        End Try

    End Sub

    Public Sub Sendkeys(text As Object, Optional wait As Boolean = False)
        Try
            Dim WshShell As Object
            WshShell = CreateObject("wscript.shell")
            WshShell.Sendkeys(CStr(text), wait)
            WshShell = Nothing
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Window_SizeChanged(sender As Object, e As SizeChangedEventArgs)
        Try
            If (Me.WindowState = WindowState.Minimized) Or (Me.WindowState = WindowState.Maximized) Or (Me.WindowState = WindowState.Normal) Then
                If ActiveWI = True Then SetInstructionLocation(Nothing)
            End If
            If Me.WindowState = FormWindowState.Minimized Then
                KeyboardHook.KeyHookEnable = False
                DAXObject.SetValue("SYSACTIVATED", "FALSE")
            End If

        Catch ex As Exception


            SetupErrorMsg(ex)
        End Try
    End Sub


    Private Sub Logger_LogUpdated(Logmsg As String) Handles Logger.LogUpdated
        Try
            Dim str As String() = Logmsg.Split("|")
            If str.Count > 0 Then
                lastLogger = str(str.Count - 1)

            End If
        Catch ex As Exception
            SetupErrorMsg(ex)
        End Try

    End Sub
#End Region

#Region "PrivateMethod"

    Private Sub HandleFormMessage(ByVal exp As ExceptionManager.HandledExceptionMessage, ByVal FormIndex As Integer, ByVal messageType As ExceptionManager.MessageTypes)

        Try
            ShowWaitCursor(False)
        Catch ex As Exception

        End Try

        If Not exp.InnerExceptionMessage = Nothing AndAlso exp.InnerExceptionMessage <> "" Then
            IHErrException = New System.Exception(exp.Message, New Exception(exp.InnerExceptionMessage))
        Else
            IHErrException = New System.Exception(exp.Message)
        End If

        Logger.WriteToLogFile("InHostShowWaitCursor", "AppBaseHelper_AppbaseError1  False End")


        If messageType = ExceptionManager.MessageTypes.AppError Then
            If exp.InnerExceptionMessage = "" Then
                ShowErrorToolstrip(exp.Message & IIf((exp.InnerExceptionMessage = Nothing And Not exp.Message.Contains(exp.String)), exp.String, ""), SympraxisException.ErrorTypes.AppError, True)
            Else
                ShowErrorToolstrip(exp.Message & IIf((exp.InnerExceptionMessage = Nothing And Not exp.Message.Contains(exp.String)), exp.String, ":" & exp.InnerExceptionMessage), SympraxisException.ErrorTypes.AppError, True)
            End If


        ElseIf messageType = Common.ExceptionManager.MessageTypes.Information Then
            If exp.InnerExceptionMessage = "" Then
                ShowErrorToolstrip(exp.Message & IIf((exp.InnerExceptionMessage = Nothing And Not exp.Message.Contains(exp.String)), exp.String, ""), SympraxisException.ErrorTypes.Information, True)
            Else
                ShowErrorToolstrip(exp.Message & IIf((exp.InnerExceptionMessage = Nothing And Not exp.Message.Contains(exp.String)), exp.String, ":" & exp.InnerExceptionMessage), SympraxisException.ErrorTypes.Information, True)
            End If

        ElseIf messageType = ExceptionManager.MessageTypes.Warning Then
            If exp.InnerExceptionMessage = "" Then
                ShowErrorToolstrip(exp.Message & IIf((exp.InnerExceptionMessage = Nothing And Not exp.Message.Contains(exp.String)), exp.String, ""), SympraxisException.ErrorTypes.Warning, True)
            Else
                ShowErrorToolstrip(exp.Message & IIf((exp.InnerExceptionMessage = Nothing And Not exp.Message.Contains(exp.String)), exp.String, ":" & exp.InnerExceptionMessage), SympraxisException.ErrorTypes.Warning, True)
            End If

        End If


    End Sub
    Private Sub HandleFormMessages(ByVal exp As ExceptionManager.HandledExceptionMessage, ByVal FormIndex As Integer, ByVal messageType As ExceptionManager.MessageTypes, Optional ByVal ErrorCode As ExceptionManager.ExceptionTypes = ExceptionManager.ExceptionTypes.None)

        Try
            ShowWaitCursor(False)
        Catch ex As Exception

        End Try
        If Not exp.InnerExceptionMessage = Nothing AndAlso exp.InnerExceptionMessage <> "" Then
            IHErrException = New System.Exception(exp.Message, New Exception(exp.InnerExceptionMessage))
        Else
            IHErrException = New System.Exception(exp.Message)
        End If
        '  IHErrException = New System.Exception(exp.Message & IIf((exp.InnerExceptionMessage = Nothing And Not exp.Message.Contains(exp.String)), exp.String, ":" & exp.InnerExceptionMessage))
        Logger.WriteToLogFile("InHostShowWaitCursor", "AppBaseHelper_AppbaseError1  False End")


        If messageType = ExceptionManager.MessageTypes.AppError Then
            If exp.InnerExceptionMessage = "" Then
                ShowErrorToolstrip(exp.Message & IIf((exp.InnerExceptionMessage = Nothing And Not exp.Message.Contains(exp.String)), exp.String, ""), SympraxisException.ErrorTypes.AppError, True)
            Else
                ShowErrorToolstrip(exp.Message & IIf((exp.InnerExceptionMessage = Nothing And Not exp.Message.Contains(exp.String)), exp.String, ":" & exp.InnerExceptionMessage), SympraxisException.ErrorTypes.AppError, True)
            End If


        ElseIf messageType = Common.ExceptionManager.MessageTypes.Information Then
            If exp.InnerExceptionMessage = "" Then
                ShowErrorToolstrip(exp.Message & IIf((exp.InnerExceptionMessage = Nothing And Not exp.Message.Contains(exp.String)), exp.String, ""), SympraxisException.ErrorTypes.Information, True)
            Else
                ShowErrorToolstrip(exp.Message & IIf((exp.InnerExceptionMessage = Nothing And Not exp.Message.Contains(exp.String)), exp.String, ":" & exp.InnerExceptionMessage), SympraxisException.ErrorTypes.Information, True)
            End If

        ElseIf messageType = ExceptionManager.MessageTypes.Warning Then
            If exp.InnerExceptionMessage = "" Then
                ShowErrorToolstrip(exp.Message & IIf((exp.InnerExceptionMessage = Nothing And Not exp.Message.Contains(exp.String)), exp.String, ""), SympraxisException.ErrorTypes.Warning, True)
            Else
                ShowErrorToolstrip(exp.Message & IIf((exp.InnerExceptionMessage = Nothing And Not exp.Message.Contains(exp.String)), exp.String, ":" & exp.InnerExceptionMessage), SympraxisException.ErrorTypes.Warning, True)
            End If

        End If


    End Sub

    Private Sub AHMenuOpenWithNewItems()




        ShowWaitCursor(True, "Opening...", "WorkItem Open..")
        WaitCursorProgressvalue(40)
        DAXObject.SetValue("SYS_OP_ObjectId", objWCF.Objectid)

        WaitCursortext("WorkItem Opening..", "Opening...")
        WaitCursorProgressvalue(60)



        CreateIOmanagerDirectory()



        'AHWorkitemCount = WCF.WorkitemCount


        SetWorkitemEnvironmentvariable()


        If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.GotoBack) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.GotoBack.ToUpper <> "" Then
            If objWCF.WorkitemCount > 0 Then
                goToBack = loadPluginsConfigXml.ApplicationHost.Settings.GotoBack
            End If



        End If


        CheckInputParser()

        WaitCursorProgressvalue(60)



        If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.ParallelWorkItem) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.ParallelWorkItem.ToUpper = "TRUE" Then
            Logger.WriteToLogFile("IHost", "Load Parallel Workitem Starts")
            LoadParallelWorkitem()
            Logger.WriteToLogFile("IHost", "Load Parallel Workitem Ends")
        Else
            Logger.WriteToLogFile("IHost", "Read Datapart Starts")
        ReadDatapart(fullDP)

            Logger.WriteToLogFile("IHost", "Read Datapart Ends")

            Logger.WriteToLogFile("IHost", "Load Workitem Starts")
            LoadWorkitem(fullDP)        '  HshGetDP = HshFullDP(WCF.Objectid)

            Logger.WriteToLogFile("IHost", "Load Workitem Ends")

        End If




        WaitCursorProgressvalue(80)

        SetApplyLayout()

        SetMatchBanner()

        AHEnableDisableMenu("Help", "WorkInstruction", True)
        LoadWorkInstructions()


        Logger.WriteToLogFile("IHost", "Show lastEvent Starts")
        ShowlastEvent()
        Logger.WriteToLogFile("IHost", "Show lastEvent Ends")


        DAXObject.SetValue("SYS_CURRENTFIELD", "DLN")

        configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True

        SetApplicationTitle()

        Menuhandle()


        WaitCursorProgressvalue(100)




    End Sub

    Private Sub ShowInstruction(ByVal active As Boolean)

        If active = True Then
            If ShowDI = True Then
                grdDI.Opacity = 1
            End If
            If ShowSI = True Then
                grdSI.Opacity = 1
            End If
            If ShowCI = True Then
                grdCI.Opacity = 1
            End If

            If pnlIndicator.IsOpen = True Then grdIndicator.Opacity = 1

        Else
            If ShowDI = True Then
                grdDI.Opacity = 0
            End If
            If ShowSI = True Then
                grdSI.Opacity = 0
            End If
            If ShowCI = True Then
                grdCI.Opacity = 0
            End If

            If pnlIndicator.IsOpen = True Then grdIndicator.Opacity = 0

        End If

    End Sub

    Private Sub ResetPerformanceLogs() Handles Me.ResetPerformanceLog
        '' Reset Performancelog 
    End Sub

    Private Sub Exiting(ByRef Auto As Boolean)
        Dim isCancel As Boolean = False
        Dim Irc As Integer
        Dim Errmsg As String = String.Empty

        ShowWaitCursor(True, "Exiting...", "Application Closing..")
        WaitCursortext("Application Closing..", "Exiting...")
        Logger.WriteToLogFile("Inhost", "Exiting   Plugin starts")
        If Not IsNothing(lstPluginOrder) Then
            For Each element As Object In lstPluginOrder
                DirectCast(element, INormalPlugin).Exiting(isCancel)
                If isCancel Then
                    If Auto = False Then
                        Throw New SympraxisException.EmptyException
                    End If
                End If

            Next
        End If
        If Not IsNothing(lstStartupPluginOrder) Then
            For Each element As Object In lstStartupPluginOrder
                DirectCast(element, INormalPlugin).Exiting(isCancel)
                If isCancel Then
                    If Auto = False Then
                        Throw New SympraxisException.EmptyException
                    End If
                End If

            Next
        End If




        Logger.WriteToLogFile("Inhost", "Exiting  Plugin Ends")
        '/Err
        Logger.WriteToLogFile("Inhost", "Exiting logoff starts")
        objWCF.Logoff(Irc, Errmsg)

        Logger.WriteToLogFile("Inhost", "Exiting logoff ends")
        '/Err

        Logger.WriteToLogFile("Inhost", "Exiting DeleteWorkDirectory starts")
        DeleteWorkDirectory()
        Logger.WriteToLogFile("Inhost", "Exiting DeleteWorkDirectory ends")
        mLayoutmanager.ClearInterval(True, True, True)
        Environment.Exit(-1)

    End Sub


    Private Sub Closingplugin(ByRef DPValue As Dictionary(Of String, String), ByRef IsAborted As Boolean, ByRef BNotify As Boolean, ByRef explugin As Object)

        Dim strDpvalue As String()
        Dim strDpValues As String
        Dim Dp As Dictionary(Of String, String) = Nothing


        Try


            For index As Integer = 0 To lstPluginOrder.Count - 1
                If Not IsNothing(explugin) Then
                    If DirectCast(explugin, Sympraxis.Utilities.INormalPlugin.INormalPlugin).ToString().Trim() = DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.INormalPlugin).ToString().Trim() Then
                        Continue For
                    End If
                End If
                strDpValues = Convert.ToString(DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.INormalPlugin).DataPart)
                If Not IsNothing(strDpValues) AndAlso strDpValues.Count > 0 Then
                    Dp = New Dictionary(Of String, String)
                    strDpvalue = strDpValues.Split("|")
                    For i = 0 To strDpvalue.Length - 1
                        If strDpvalue(i) <> String.Empty AndAlso strDpvalue(i).Substring(0, 1) = "@" Then
                            BuildDP4IndexnRev(strDpvalue(i), Dp, DPValue)
                        Else
                            If DPValue.ContainsKey(strDpvalue(i)) Then
                                Dp.Add(strDpvalue(i), DPValue(strDpvalue(i)))
                            End If
                        End If
                    Next
                ElseIf IsNothing(strDpValues) Or (Not IsNothing(strDpValues) AndAlso strDpValues = String.Empty) Then
                    Continue For
                End If

                Logger.WriteToLogFile("IHost", "Closing Plugin starts -- >" & lstPluginOrder(index).ToString())


                'to do 

                DirectCast(lstPluginOrder(index), INormalPlugin).Closing(Dp, IsAborted)
                If IsAborted Then

                    Throw New SympraxisException.EmptyException()
                End If

                If DirectCast(lstPluginOrder(index), INormalPlugin).ToString.Trim = "Sympraxis.Plugins.Notification.Notification" Or DirectCast(lstPluginOrder(index), INormalPlugin).ToString.Trim = "Sympraxis.Plugins.SendTo.Notification" Then
                    Dim lOwnerId As String = ""
                    Dim lOwnerType As String = ""
                    Dim OutPutString As String = ""
                    Dim y As New Xml.XmlDocument
                    Dim GetXnode As XmlNode
                    If Dp.Count > 0 Then
                        If Not Dp.Item("XNotify") Is Nothing Then
                            y.LoadXml(Dp.Item("XNotify"))
                        End If
                    End If


                    GetXnode = y.FirstChild
                    'GetXnode = DPValue(0)
                    If Not IsNothing(GetXnode) Then
                        If DAXObject.GetValue("SYS_NOTIFY_VALID") = "TRUE" OrElse DAXObject.GetValue("SYS_NOTIFY_VALID") = "FAlSE" Then
                            lOwnerId = GetXnode.LastChild.LastChild.SelectSingleNode("SendUID").InnerText.ToString.Trim
                            lOwnerType = GetXnode.LastChild.LastChild.SelectSingleNode("Type").InnerText.ToString.Trim
                            Dim irc As Integer
                            Dim errmsg As String
                            If lOwnerId <> 0 AndAlso lOwnerId <> "" Then
                                objWCF.SetNextOwnerandOwnerClass(lOwnerId, lOwnerType, irc, errmsg)
                                If irc <> 0 Then

                                    Throw New SympraxisException.WorkitemException(errmsg)
                                End If
                                Logger.WriteToLogFile("configMgr.AppbaseConfig.", "Notify Owner Changed")
                            Else
                                objWCF.SetNextOwnerandOwnerClass(configMgr.AppbaseConfig.EnvironmentVariables.lWorkGroupId, "W", irc, errmsg)
                                If irc <> 0 Then

                                    Throw New SympraxisException.WorkitemException(errmsg)
                                End If

                                Logger.WriteToLogFile("InterHost", "Notify Owner Changed")
                            End If

                            If DAXObject.GetValue("SYS_NOTIFY_VALID") = "TRUE" Then
                                BNotify = True
                            End If
                        End If
                    End If
                End If




                If Not IsNothing(TryCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                    If DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrCode <> 0 Then
                        Throw New SympraxisException.WorkitemException(DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrMsg)
                    End If
                End If
                Logger.WriteToLogFile("IHost", "Closing Plugin Ends")

                If Not IsNothing(Dp) Then
                    For Each eDp In Dp
                        If strDpValues.Substring(0, 1) = "@" Then
                            DPValue("@" + eDp.Key) = eDp.Value
                        Else
                            DPValue(eDp.Key) = eDp.Value
                        End If
                    Next

                    Dp = Nothing
                End If
            Next




        Finally
            strDpvalue = Nothing
            strDpValues = Nothing
            Dp = Nothing
        End Try

    End Sub

    Private Sub CloseWork(ByRef IsClose As Integer)

        Dim Irc As Integer
        Dim ErrMsg As String = String.Empty

        Dim isErrormached As Boolean = False
        Dim IsAborted As Boolean = False
        Dim Breject As Boolean = False
        Dim BReWork As Boolean = False
        Dim BNotify As Boolean = False

        If Not IsNothing(fullDP) Then
            Logger.WriteToLogFile("IHost", "Close plugin Starts")

            Closingplugin(fullDP, IsAborted, BNotify, Nothing)

            If IsAborted Then
                IsClose = 1
                Exit Sub
            End If
            Logger.WriteToLogFile("IHost", "Close plugin Ends")


            If fullDP.ContainsKey("bRej") AndAlso fullDP.Item("bRej") <> "" Then

                If fullDP.Item("bRej") = "1" Or fullDP.Item("bRej") = "True" Then
                    Breject = True
                End If



            End If
            If fullDP.ContainsKey("bReWk") AndAlso fullDP.Item("bReWk") <> "" Then

                If fullDP.Item("bReWk") = "1" Or fullDP.Item("bReWk") = "True" Then
                    BReWork = True
                End If

            End If

            Logger.WriteToLogFile("IHost", "Write Datapart Starts")
            Irc = objWCF.Write(fullDP, ErrMsg)

            If Irc <> 0 Then
                Throw New SympraxisException.WorkitemException(ErrMsg)

            End If
            Logger.WriteToLogFile("IHost", "Write Datapart Ends")


            Logger.WriteToLogFile("IHost", "UpdateSaveWFStatus Flush Data starts")
            objWCF.UpdateSaveWFStatus("FLUSH", Irc, ErrMsg)
            Logger.WriteToLogFile("IHost", "UpdateSaveWFStatus Flush Data Ends")

            If Irc <> 0 Then
                Throw New Exception()
            End If




            If Breject = True Or BReWork = True Then
                objWCF.UpdateSaveWFStatus("REJECT", Irc, ErrMsg)
            ElseIf BNotify = True Then
                Logger.WriteToLogFile("InteractiveHost", "Parser skipped due to Notification")
            Else
                ClosingRules(isErrormached, "")


                If isErrormached = True Then
                    IsClose = 1
                    Exit Sub
                End If
            End If
            fullDP = Nothing
       
        End If

        ResetInstruction()
    End Sub



    Private Sub CloseParallelWork(ByRef IsClose As Integer)

        Dim Irc As Integer
        Dim ErrMsg As String = String.Empty
        Dim ParallelDp As Dictionary(Of String, Dictionary(Of String, String)) = Nothing
        Dim ParallelChildDp As Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, String))) = Nothing
        Dim childDatapart As Dictionary(Of String, Dictionary(Of String, String)) = Nothing
        Dim objparalleldp As clsParallelDp = Nothing
        Dim isErrormached As Boolean = False
        Dim IsAborted As Boolean = False
        Dim Breject As Boolean = False
        Dim BReWork As Boolean = False
        Dim BNotify As Boolean = False


        Try

       
        If ActiveWIMode = True Then
            DeActiveWorkitemMode(Irc, ErrMsg)


            If Irc <> 0 Then
                Throw New Exception(ErrMsg)
            End If

        End If


        If Not IsNothing(listParallelDP) Then
            Logger.WriteToLogFile("IHost", "Close plugin Starts")

            For index As Integer = 0 To lstPluginOrder.Count - 1
                If Not IsNothing(TryCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then

                        objparalleldp = listParallelDP.Where(Function(x) x.pluginName = lstPluginOrder(index).ToString()).FirstOrDefault()
                    If Not IsNothing(objparalleldp) Then
                        ParallelDp = objparalleldp.ParallelDP
                        ParallelChildDp = objparalleldp.ParallelChildDp

                        DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).Closing(ParallelDp, ParallelChildDp, IsAborted)
                        If IsAborted Then

                            Throw New SympraxisException.EmptyException()
                        End If

                    End If
                End If
            Next




                If Not IsNothing(ParallelDp) Then


                    For i As Integer = 0 To ParallelDp.Count - 1
                        Breject = False
                        BReWork = False
                        BNotify = False
                        objWCF.SwitchWorkitembyObjectid(ParallelDp.ElementAt(i).Key, Irc, ErrMsg)
                        '' Fail 
                        If Irc <> 0 Then
                            Throw New SympraxisException.WorkitemException(ErrMsg)
                        End If
                        Dim DPValue = New Dictionary(Of String, String)
                        DPValue = ParallelDp.ElementAt(i).Value


                        If DPValue.Count > 0 Then
                            Logger.WriteToLogFile("IHost", "Close plugin Ends")
                            If DPValue.ContainsKey("bRej") AndAlso DPValue.Item("bRej") <> "" Then
                                If DPValue.Item("bRej") = "1" Or DPValue.Item("bRej") = "True" Then
                                    Breject = True
                                End If
                            End If
                            If DPValue.ContainsKey("bReWk") AndAlso DPValue.Item("bReWk") <> "" Then
                                If DPValue.Item("bReWk") = "1" Or DPValue.Item("bReWk") = "True" Then
                                    BReWork = True
                                End If
                            End If

                            Logger.WriteToLogFile("IHost", "Write Datapart Starts")
                            Irc = objWCF.Write(DPValue, ErrMsg)

                            If Irc <> 0 Then
                                Throw New SympraxisException.WorkitemException(ErrMsg)

                            End If
                            Logger.WriteToLogFile("IHost", "Write Datapart Ends")

                            If ParallelChildDp.Count > 0 Then
                                childDatapart = ParallelChildDp(ParallelDp.ElementAt(i).Key)

                                If Not IsNothing(childDatapart) AndAlso childDatapart.Count > 0 Then
                                    objWCF.WriteChildDataPart(childDatapart, Irc, ErrMsg)

                                    If Irc <> 0 Then
                                        Throw New SympraxisException.WorkitemException(ErrMsg)

                                    End If
                                End If
                            End If



                            Logger.WriteToLogFile("IHost", "UpdateSaveWFStatus Flush Data starts")
                            objWCF.UpdateSaveWFStatus("FLUSH", Irc, ErrMsg)
                            Logger.WriteToLogFile("IHost", "UpdateSaveWFStatus Flush Data Ends")

                            If Irc <> 0 Then
                                Throw New Exception()
                            End If
                            If Breject = True Or BReWork = True Then
                                objWCF.UpdateSaveWFStatus("REJECT", Irc, ErrMsg)
                                If Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.ChildWFsts) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.ChildWFsts.ToUpper() = "TRUE" Then
                                Else
                                    objWCF.UpdateChildSaveWFStatus("REJECT", "ALL", Irc, ErrMsg)

                                End If
                            ElseIf BNotify = True Then
                                Logger.WriteToLogFile("InteractiveHost", "Parser skipped due to Notification")
                            Else
                                ClosingRules(isErrormached, "")


                                If isErrormached = True Then
                                    IsClose = 1
                                    Exit Sub
                                End If
                            End If


                        End If
                        If Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.ChildWFsts) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.ChildWFsts.ToUpper() = "TRUE" Then


                            Dim Brejectchd As Boolean = False
                            Dim BReWorkchd As Boolean = False

                            If Not IsNothing(childDatapart) AndAlso childDatapart.Count > 0 Then
                                For icnt As Integer = 0 To childDatapart.Count - 1
                                    Dim closedp As New Dictionary(Of String, String)(childDatapart.ElementAt(icnt).Value)

                                    If closedp.ContainsKey("bRej") AndAlso closedp.Item("bRej") <> "" Then

                                        If closedp.Item("bRej") = "1" Or closedp.Item("bRej") = "True" Then
                                            Brejectchd = True
                                        End If

                                    End If
                                    If closedp.ContainsKey("bReWk") AndAlso closedp.Item("bReWk") <> "" Then

                                        If closedp.Item("bReWk") = "1" Or closedp.Item("bReWk") = "True" Then
                                            BReWorkchd = True
                                        End If

                                    End If
                                    If Brejectchd = True Or BReWorkchd = True Then
                                        objWCF.UpdateChildSaveWFStatus("REJECT", childDatapart.ElementAt(icnt).Key, Irc, ErrMsg)

                                    Else
                                        Dim isErrormachedchd As Boolean = False
                                        ClosingRules(isErrormachedchd, childDatapart.ElementAt(icnt).Key)


                                        If isErrormachedchd = True Then

                                            Exit Sub
                                        End If
                                    End If
                                    closedp = Nothing
                                Next
                            End If
                        End If
                        childDatapart = Nothing
                    Next
                End If


            End If

            ResetInstruction()
            listParallelDP = Nothing
        Finally

            ParallelDp = Nothing
            ParallelChildDp = Nothing
            childDatapart = Nothing
            objparalleldp = Nothing

        End Try
    End Sub

    Private Sub SaveErrorWp()
        If Not IsNothing(appHostConfigXml.ErrorWPFolderPath) AndAlso appHostConfigXml.ErrorWPFolderPath <> "" Then
            Dim Irc As Integer = 0
            Dim ErrMsg As String = String.Empty
            Dim Savefilepath As String = appHostConfigXml.ErrorWPFolderPath
            Dim lastchar As Char = Savefilepath(Savefilepath.Length - 1)
            If lastchar <> "\" Then
                Savefilepath = Savefilepath & "\"
            ElseIf IOManager.CheckDirectoryExists(Savefilepath) = False Then
                IOManager.CreateDirectory(Savefilepath)
            End If
            Savefilepath = Savefilepath & intrHostProcessName & "_" & DateTime.Now.ToString("ddmmyyyyHHmm")
            Dim FinalOutPut As String = ""
            objWCF.GetFinalWp(FinalOutPut, Irc, ErrMsg)

            If Irc = 0 AndAlso FinalOutPut <> "" Then
                Dim file As System.IO.StreamWriter
                file = My.Computer.FileSystem.OpenTextFileWriter(Savefilepath, True)
                file.WriteLine(FinalOutPut)
                file.Close()
            End If
        End If

    End Sub
    Private Sub OfflineCloseWork()
        Dim IsAborted As Boolean = False
        Dim BNotify As Boolean = False
        If Not IsNothing(fullDP) Then
            Closingplugin(fullDP, IsAborted, BNotify, Nothing)
        End If



        For Each element As Object In lstPluginOrder
            If Not IsNothing(TryCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                DirectCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrCode = 0
            End If
            DirectCast(element, INormalPlugin).Closed()
            If Not IsNothing(TryCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                If DirectCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrCode <> 0 Then
                    Throw New SympraxisException.AppExceptionSaveWP(DirectCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrMsg)
                End If
            End If
        Next



        If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.SetWD) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.SetWD.ToUpper() = "TRUE" Then
            mLayoutmanager.SETWD()

            For Each MenuElement In PluginMenuEnableDisable
                AHEnableDisableMenu(MenuElement.Key, "", False)
            Next
        Else
            AHEnableMenu("Home", "Open|Exit")
        End If




        BannerPanel.Children.Clear()
        BannerPanel.Visibility = System.Windows.Visibility.Collapsed
        ButtonPin.Visibility = System.Windows.Visibility.Collapsed
        DAXObject.SetValue("SYSBannerHeight", 0)
        fullDP = Nothing
        configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = False

    End Sub
    Private Sub AHMenuCloseWork(ByRef IsError As Boolean)

        Dim IsClose As Integer = 0
        Dim Irc As Integer = 0
        Dim ErrMsg As String = String.Empty

        Try

            If configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True Then

                If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.ParallelWorkItem) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.ParallelWorkItem.ToUpper = "FALSE" Then


                If IsAbortValidation = False AndAlso objWCF.CurrentWorkitemIdx <> objWCF.WorkitemCount Then

                    If BlnClosework Then

                        ShowErrorToolstrip("UNCOMPLETED workitems present, Press Ctrl+W to push.", SympraxisException.ErrorTypes.Information, False)


                        BlnClosework = False

                        Exit Sub
                    Else

                        BlnClosework = True

                    End If

                End If
                End If

                ShowWaitCursor(True, "Closing..", "Closing...")

                If IsError = False AndAlso configMgr.AppbaseConfig.EnvironmentVariables.IsReadonlyWP = False Then
                    isExitflag = True

                    WaitCursorProgressvalue(10)

                    Logger.WriteToLogFile("IHost", "Close Work starts")

                    If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.ParallelWorkItem) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.ParallelWorkItem.ToUpper = "TRUE" Then
                        CloseParallelWork(IsClose)
                    Else

                    CloseWork(IsClose)

                        CloseChildworkitem()

                    End If

                    If IsClose <> 0 Then Exit Sub
                    Logger.WriteToLogFile("IHost", "Close Work ends")



                    Logger.WriteToLogFile("IHost", "Setpredicate starts")
                    objWCF.Runsetpredicate(Irc, ErrMsg)
                    If Irc <> 0 Then
                        Throw New SympraxisException.AppExceptionSaveWP(ErrMsg)
                    End If
                    Logger.WriteToLogFile("IHost", "Setpredicate ends")

                End If



                WorkHourTimer.Enabled = False
                WorkHourTimer.Stop()
                If openworkinterval <> -1 Then
                    openworkinterval = 0
                    'mLayoutmanager.ClearInterval(False, True, False)
                End If

                If workingMode.ToLower() = "normal" Then


                    If configMgr.AppbaseConfig.EnvironmentVariables.IsReadonlyWP = False Then


                        Logger.WriteToLogFile("IHost", "Update Workitem starts")
                        If objWCF.OfflineWorkOpen = True Then

                            GoOnline()
                        End If
                        WaitCursorProgressvalue(20)
                        objWCF.PutWork(Irc, ErrMsg)
                        DeleteWIPWorkDirectory()
                        If Irc <> 0 Then
                            Throw New Exception(ErrMsg)
                        End If
                        Logger.WriteToLogFile("IHost", "Update Workitem ends")
                    Else
                        configMgr.AppbaseConfig.EnvironmentVariables.IsReadonlyWP = False
                        DeleteWIPWorkDirectory()
                    End If



                Else

                    Dim FinalOutPut As String = ""
                    objWCF.GetFinalWp(FinalOutPut, Irc, ErrMsg)

                    If Irc = 0 AndAlso FinalOutPut <> "" Then

                        If workingMode.ToLower() <> "offline" Then

                            Dim dlg As New Microsoft.Win32.SaveFileDialog()

                            ' Display OpenFileDialog by calling ShowDialog method 
                            Dim result As Nullable(Of Boolean) = dlg.ShowDialog()

                            If objWCF.WorkitemCount > objWCF.CurrentWorkitemIdx Then

                            Else


                                If result = True Then


                                    Dim file As System.IO.StreamWriter
                                    file = My.Computer.FileSystem.OpenTextFileWriter(dlg.FileName, True)
                                    file.WriteLine(FinalOutPut)
                                    file.Close()

                                End If
                            End If
                        Else
                            If IsNothing(objWinZip) Then
                                objWinZip = New WinZip
                            End If

                            Dim unZipPath As String = appHostConfigXml.OfflinePath & "\" & intrHostWorkFlowName & "\" & configMgr.AppbaseConfig.EnvironmentVariables.sUserId & "\" & intrHostProcessName & "\" & objWCF.Workpacketid & ".zip"

                            Dim parent As String = System.IO.Path.GetDirectoryName(unZipPath)

                            Dim LocalPath As String = appHostConfigXml.OfflinePath & "\" & intrHostWorkFlowName & "\" & configMgr.AppbaseConfig.EnvironmentVariables.sUserId & "\" & intrHostProcessName & "\" & objWCF.Workpacketid & "\" & objWCF.Workpacketid & ".txt"
                            If IOManager.CheckFileExists(LocalPath) Then
                                IOManager.DeleteFile(LocalPath)

                                Dim file As System.IO.StreamWriter
                                file = My.Computer.FileSystem.OpenTextFileWriter(LocalPath, True)
                                file.WriteLine(FinalOutPut)
                                file.Close()
                            Else
                                If IOManager.CheckFileExists(unZipPath) Then
                                    objWinZip.UnZIP(unZipPath, parent & "/" & objWCF.Workpacketid, parent & "/" & objWCF.Workpacketid, Irc, ErrMsg)
                                    IOManager.DeleteFile(unZipPath)
                                    If IOManager.CheckFileExists(LocalPath) Then
                                        IOManager.DeleteFile(LocalPath)
                                    End If

                                    Dim file As System.IO.StreamWriter
                                    file = My.Computer.FileSystem.OpenTextFileWriter(LocalPath, True)
                                    file.WriteLine(FinalOutPut)
                                    file.Close()
                                    objWinZip.ZIP(parent & "/" & objWCF.Workpacketid, parent & "/" & objWCF.Workpacketid, parent & "/" & objWCF.Workpacketid, Irc, ErrMsg)
                                    IOManager.DeleteDirectory(parent & "/" & objWCF.Workpacketid)
                                End If
                            End If




                            RefreshlocOfflineGrid()
                            OfflineWPPath = ""
                        End If



                        AHEnableMenu("Home", "Open|Exit")
                    Else
                        Throw New SympraxisException.InformationException(ErrMsg)
                    End If
                End If
                configMgr.AppbaseConfig.EnvironmentVariables.IsReadonlyWP = False
                'ToolstripReadOnly.Visibility = System.Windows.Visibility.Collapsed
                btnReadnly.Visibility = System.Windows.Visibility.Collapsed

                Logger.WriteToLogFile("IHost", "Closed plugin starts")
                For Each element As Object In lstPluginOrder
                    If Not IsNothing(TryCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                        DirectCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrCode = 0
                    End If
                    DirectCast(element, INormalPlugin).Closed()
                    If Not IsNothing(TryCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                        If DirectCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrCode <> 0 Then
                            Throw New SympraxisException.AppExceptionSaveWP(DirectCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrMsg)
                        End If
                    End If
                Next


                If Not ObjClientInstanceMgr Is Nothing Then
                    Logger.WriteToLogFile("InteractiveHost", "ObjClientInstanceMgr.CloseWorkpacket St")
                    ObjClientInstanceMgr.CloseWorkpacket(intrHostProcessName)
                    Logger.WriteToLogFile("InteractiveHost", "ObjClientInstanceMgr.CloseWorkpacket End")
                    If Not ObjClientInstanceMgr.MyDAXObject Is Nothing Then
                        ObjClientInstanceMgr.MyDAXObject.ClearAll()
                    End If

                End If



                Try
                    If Not ObjManagerInstanceMgr Is Nothing Then
                        ObjManagerInstanceMgr.CloseWorkpacket(intrHostProcessName)
                        Logger.WriteToLogFile("InteractiveHost", "ObjManagerInstanceMgr.CloseWorkpacket End")


                        Logger.WriteToLogFile("InteractiveHost", "ObjManagerInstanceMgr.CloseWorkpacket St")
                        'If Not IsNothing(ObjManagerInstanceMgr) AndAlso Not IsNothing(ObjManagerInstanceMgr.MyDAXObject) Then
                        '    ObjManagerInstanceMgr.MyDAXObject.ClearAll()
                        'End If



                    End If
                Catch ex As Exception

                End Try
                If IsError = False Then
                    ShowErrorToolstrip("WorkPacket Saved Successfully", SympraxisException.ErrorTypes.Information, False)

                End If


                Logger.WriteToLogFile("IHost", "Closed plugin ends")
                If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.SetWD) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.SetWD.ToUpper() = "TRUE" Then

                    mLayoutmanager.SETWD()

                    For Each MenuElement In PluginMenuEnableDisable
                        AHEnableDisableMenu(MenuElement.Key, "", False)
                    Next
                    If workingMode.Contains("debug") Then
                        AHEnableMenu("Home", "Open|Exit")
                    Else
                        AHEnableMenu("Home", "Exit")
                    End If

                Else
                    AHEnableMenu("Home", "Open|Exit")
                End If


                BannerPanel.Children.Clear()
                BannerPanel.Visibility = System.Windows.Visibility.Collapsed
                ButtonPin.Visibility = System.Windows.Visibility.Collapsed
                configBannersettings = Nothing
                DAXObject.SetValue("SYSBannerHeight", 0)
                IsCaseAttended = False
                IsAbortValidation = False
                IsObjectautoclose = False
                If Not ManualIdleTimer Is Nothing Then
                    ManualIdleTimer.Stop()
                    ManualIdleTimer.Enabled = False
                    ManualIdleTimer = Nothing
                End If

                DAXObject.SetValue("SYS_OP_ObjectId", 0)

                goToBack = 0

                Logger.WriteToLogFile("IHost", "Disable Menu starts")
                For Each MenuElement In PluginMenuEnableDisable
                    AHEnableDisableMenu(MenuElement.Key, "", False)
                Next
                Logger.WriteToLogFile("IHost", "Disable Menu Ends")
                If Irc = 0 Then
                    fullDP = Nothing
                    configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = False
                End If

                If BlnOpenAftClose = True Then
                    OpenWorkAcrossProcessAndWF(NotifyMessageObjectId, Sympraxis.Utilities.INormalPlugin.IPluginhost.GetObjectType.ObjectId, Irc, "")
                    BlnOpenAftClose = False
                End If
               
            Else

                For Each element As Object In lstPluginOrder
                    If Not IsNothing(TryCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                        DirectCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrCode = 0
                    End If
                    DirectCast(element, INormalPlugin).Closed()
                    If Not IsNothing(TryCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                        If DirectCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrCode <> 0 Then
                            Throw New SympraxisException.AppExceptionSaveWP(DirectCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrMsg)
                        End If
                    End If
                Next



                If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.SetWD) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.SetWD.ToUpper() = "TRUE" Then
                    mLayoutmanager.SETWD()
                    For Each MenuElement In PluginMenuEnableDisable
                        AHEnableDisableMenu(MenuElement.Key, "", False)
                    Next
                    AHEnableMenu("Home", "Exit")
                Else
                    AHEnableMenu("Home", "Open|Exit")
                End If




                BannerPanel.Children.Clear()
                BannerPanel.Visibility = System.Windows.Visibility.Collapsed
                ButtonPin.Visibility = System.Windows.Visibility.Collapsed
                DAXObject.SetValue("SYSBannerHeight", 0)
                If Irc = 0 Then
                    fullDP = Nothing
                    configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = False
                End If
            End If

            If Not ObjClientInstanceMgr Is Nothing Then
                Logger.WriteToLogFile("InHost", "ObjClientInstanceMgr.CloseWorkpacket St")
                ObjClientInstanceMgr.CloseWorkpacket(intrHostProcessName)
                Logger.WriteToLogFile("InHost", "ObjClientInstanceMgr.CloseWorkpacket End")

            End If
            pnlNotifyHeader.Visibility = System.Windows.Visibility.Collapsed


        Finally
            isExitflag = False
            ShowWaitCursor(False)
            configMgr.AppbaseConfig.EnvironmentVariables.IsAutoClose = False
            configMgr.AppbaseConfig.EnvironmentVariables.IsAutoOpen = False
            _Isautoopen = False

        End Try
    End Sub


    Private Sub AHMenuPreviousWork(ByRef isError As Boolean)
        Try

            ShowWaitCursor(True, "Opening Previous..", "Previous...")

            Dim Irc As Integer = 0
            Dim Errmsg As String = String.Empty
            Dim IsClose As Integer = 0
            Prvs = True
            If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.ParallelWorkItem) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.ParallelWorkItem.ToUpper = "FALSE" Then


                If objWCF.WorkitemCount > 0 Then
                    If goToBack <= 0 Then
                        ShowErrorToolstrip(String.Concat("You can only move upto ", loadPluginsConfigXml.ApplicationHost.Settings.GotoBack, "  previous workitems from your current workitem."), SympraxisException.ErrorTypes.Information, False)
                        Exit Sub
                    Else
                        If isError = False Then
                            goToBack = goToBack - 1

                        End If
                    End If
                End If

                'DFor Testing


                If isError = False Then
                    Logger.WriteToLogFile("Inhost", "PreviousWork:CloseWork starts")
                    CloseWork(IsClose)
                    If IsClose <> 0 Then Exit Sub
                    Logger.WriteToLogFile("Inhost", "PreviousWork:CloseWork Ends")
                Else

                    objWCF.UpdateSaveWFStatus("Error", Irc, Errmsg)
                    If Irc <> 0 Then
                        Throw New Exception()
                    End If
                End If




                Logger.WriteToLogFile("Inhost", "PreviousWork:Closed plugin starts")
                For Each element As Object In lstPluginOrder
                    If Not IsNothing(TryCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                        DirectCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrCode = 0
                    End If
                    DirectCast(element, INormalPlugin).Closed()
                    If Not IsNothing(TryCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                        If DirectCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrCode <> 0 Then
                            Throw New SympraxisException.AppExceptionSaveWP(DirectCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrMsg)
                        End If
                    End If
                Next
                Logger.WriteToLogFile("Inhost", "PreviousWork:Closed plugin Ends")

                ' why ?
                '  mLayoutmanager.SETWD()


                Logger.WriteToLogFile("Inhost", "PreviousWork:Move to Previous WI WCF starts")
                objWCF.MovetoPreviousWI(Irc, Errmsg)

                If Irc <> 0 Then

                    Throw New SympraxisException.WorkitemException(Errmsg)
                End If

                ObjectQualifyingAdaptor = objWCF.ObjectQualifyingAdaptor

                Logger.WriteToLogFile("Inhost", "PreviousWork:Move to Previous WI WCF Ends")

                '  CreateIOmanagerDirectory(irc, Errmsg)





                Logger.WriteToLogFile("Inhost", "PreviousWork:Read Datapart starts")
                ReadDatapart(fullDP)
                If Irc <> 0 Then Exit Sub
                Logger.WriteToLogFile("Inhost", "PrevousWork:Read Datapart Ends")

                Logger.WriteToLogFile("Inhost", "PrevousWork:Load Workitem starts")
                LoadWorkitem(fullDP)
                If Irc <> 0 Then Exit Sub
                Logger.WriteToLogFile("Inhost", "PrevousWork:Load Workitem Ends")


                Menuhandle()

                If Not IsNothing(configBannersettings) Then

                    Logger.WriteToLogFile("Inhost", "PrevousWork:Apply Banner starts")
                    ApplyBanner(configBannersettings)
                    Logger.WriteToLogFile("Inhost", "PrevousWork:Apply Banner Ends")

                End If
            End If

        Finally
            ShowWaitCursor(False)
        End Try



    End Sub


    'Private Sub CheckandOpenNextWorkItem(ByVal CurrentWIstst As String, ByRef irc As Integer, ByRef Expdetails As ExceptionForm.ExceptionDetails)
    '    Try
    '        irc = 0
    '        If CurrentWIstst <> "" Then
    '            objWCF.UpdateSaveWFStatus(CurrentWIstst, irc, Expdetails.ErrDetails)
    '        End If

    '        If irc = 0 And objWCF.CurrentWorkitemIdx < objWCF.WorkitemCount Then
    '            objWCF.MovetoNextWI(irc, Expdetails.ErrDetails)

    '            If irc <> 0 Then Exit Sub

    '            GetandLoadWI(irc, Expdetails)

    '            If irc <> 0 Then Exit Sub
    '        ElseIf objWCF.CurrentWorkitemIdx = objWCF.WorkitemCount Then
    '            objWCF.PutWork(irc, Expdetails.ErrDetails)
    '            If Not IsNothing(lstPluginOrder) AndAlso lstPluginOrder.Count > 0 Then
    '                For Each element As Object In lstPluginOrder
    '                    DirectCast(element, INormalPlugin).Closed()
    '                Next
    '            End If
    '            irc = 0
    '        End If

    '    Catch ex As Exception
    '        irc = 1
    '        Expdetails.ErrMsg = String.Concat("CheckOpenWorkItem :", ex.Message)
    '        IHErrException = ex


    '    End Try
    'End Sub


    Private Sub AHMenuNextWork(ByRef isError As Boolean)
        Try


            Dim Irc As Integer = 0
            Dim Errmsg As String = String.Empty
            Dim IsClose As Integer = 0
            ShowWaitCursor(True, "Opening NextWork..", "NextWork...")
            If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.ParallelWorkItem) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.ParallelWorkItem.ToUpper = "FALSE" Then

                If isError = False Then
                    Logger.WriteToLogFile("IHost", "Close Work Start")
                    CloseWork(IsClose)
                    If IsClose <> 0 Then Exit Sub
                    Logger.WriteToLogFile("IHost", "Close Work Ends")
                Else

                    objWCF.UpdateSaveWFStatus("Error", Irc, Errmsg)
                    If Irc <> 0 Then
                        Throw New Exception()
                    End If
                End If


                Logger.WriteToLogFile("IHost", "Close Work Start")
                For Each element As Object In lstPluginOrder
                    If Not IsNothing(TryCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                        DirectCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrCode = 0
                    End If
                    DirectCast(element, INormalPlugin).Closed()
                    If Not IsNothing(TryCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                        If DirectCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrCode <> 0 Then
                            Throw New SympraxisException.AppExceptionSaveWP(DirectCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrMsg)
                        End If
                    End If
                Next
                Logger.WriteToLogFile("IHost", "Close Work Ends")

                ' mLayoutmanager.SETWD()



                Logger.WriteToLogFile("IHost", "Moveto NextWI Start")
                objWCF.MovetoNextWI(Irc, Errmsg)

                If Irc <> 0 Then
                    Throw New SympraxisException.WorkitemException(Errmsg)
                End If
                ObjectQualifyingAdaptor = objWCF.ObjectQualifyingAdaptor

                Logger.WriteToLogFile("IHost", "Moveto NextWI Ends")
                Logger.WriteToLogFile("IHost", "Moveto NextWI Start")

                GetandLoadWI()

                Logger.WriteToLogFile("IHost", "Moveto NextWI Start")



            Else
                '' Have to Implement parallel workitem 

            End If

        Finally

            ShowWaitCursor(False)

        End Try


    End Sub

    Private Sub GetandLoadWI()

        ' CreateIOmanagerDirectory(Irc, Errmsg)




        '  AHCurrenWorkItem = AHCurrenWorkItem + 1
        ' HshGetDP = HshFullDP(WCF.Objectid)
        'ResetDp()
        Logger.WriteToLogFile("IHost", "Read Datapart Start")
        ReadDatapart(fullDP)

        Logger.WriteToLogFile("IHost", "Read Datapart Ends")
        '  CheckandOpenNextWorkItem("Error", Irc, Expdetails)



        Logger.WriteToLogFile("IHost", "Load Workitem Start")
        LoadWorkitem(fullDP)

        Logger.WriteToLogFile("IHost", "Load Workitem Ends")
        LoadChildworkitem()

        ' If Irc = 0 Then

        Menuhandle()
        If Not IsNothing(configBannersettings) Then

            Logger.WriteToLogFile("IHost", "Apply Banner Start")
            ApplyBanner(configBannersettings)
            Logger.WriteToLogFile("IHost", "Apply Banner Ends")

        End If
        'Else
        'AHEnableMenu("Work", "Close")
        'End If

    End Sub
    Private Sub AHMenuExit(ByRef bhandle As Boolean)
        Dim IMApplicationCount As Integer

        If workingMode <> "\debug" Then
            If ObjClientInstanceMgr IsNot Nothing Then
                ObjClientInstanceMgr.RemoveApp(intrHostProcessName & "^" & intrHostWorkFlowName)
            End If
            If ObjManagerInstanceMgr IsNot Nothing Then
                IMApplicationCount = ObjManagerInstanceMgr.GetApplicationCount()
                If IMApplicationCount > 0 Then

                    Throw New SympraxisException.InformationException("Please close the processes(" & IMApplicationCount & ") before exit the ARMS Application")
                    bhandle = True
                    Exit Sub
                Else
                    ObjManagerInstanceMgr = Nothing
                End If
            End If
        End If
        If Not IsNothing(appHostConfigXml.AppExitConfirm) AndAlso appHostConfigXml.AppExitConfirm.ToUpper() = "TRUE" Then
            If MsgBox("Do you want to Exit?", MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then
                bhandle = True
                Exit Sub

            End If
        End If

        If configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = False Then
            Logger.WriteToLogFile("Inhost", "Exiting logoff starts")
            Exiting(False)
            Logger.WriteToLogFile("Inhost", "Exiting logoff starts")
            'Note
        Else

            Throw New SympraxisException.InformationException("Work Already Opened")

            Exit Sub
        End If

    End Sub


    Private Sub AHMenuOpenWork()
        Dim Irc As Integer = 0
        Dim ErrMsg As String = String.Empty


        If Not ObjClientInstanceMgr Is Nothing Then
            Logger.WriteToLogFile("InHost", "ObjClientInstanceMgr.IsWorkpacketOpened Start")
            Dim _CurrentlyActiveApp As String = ""

            If ObjClientInstanceMgr.IsWorkpacketOpened(intrHostProcessName, _CurrentlyActiveApp) Then
                Throw New SympraxisException.InformationException("Workpacket opened in " & _CurrentlyActiveApp & " application.")

            End If
            Logger.WriteToLogFile("InHost", "ObjClientInstanceMgr.IsWorkpacketOpened End")
        End If

        If Not ObjManagerInstanceMgr Is Nothing Then
            Logger.WriteToLogFile("InHost", "ObjManagerInstanceMgr.IsWorkpacketOpened Start")
            Dim _CurrentlyActiveApp As String = ""

            If ObjManagerInstanceMgr.IsWorkpacketOpened(intrHostProcessName, _CurrentlyActiveApp) Then
                Throw New SympraxisException.InformationException("Workpacket opened in " & _CurrentlyActiveApp & " application.")

            End If
            Logger.WriteToLogFile("InHost", "ObjManagerInstanceMgr.IsWorkpacketOpened End")
        End If
        ShowWaitCursor(True, "Opening...", "WorkItem Open..")
        WaitCursorProgressvalue(40)
        Logger.WriteToLogFile("IHost", "GetWork Starts")
        If workingMode <> "\debug" Then

            Irc = objWCF.GetWork(ErrMsg)
            ObjectQualifyingAdaptor = objWCF.ObjectQualifyingAdaptor

        Else
            Dim dlg As New Microsoft.Win32.OpenFileDialog()

            ' Display OpenFileDialog by calling ShowDialog method 
            Dim result As Nullable(Of Boolean) = dlg.ShowDialog()
            If result = True Then


                objWCF.GetWorkPackets4Debug(dlg.FileName.Trim, Irc, ErrMsg)

                If Irc <> 0 Then
                    Throw New SympraxisException.AppException(ErrMsg)
                End If

                ObjectQualifyingAdaptor = objWCF.ObjectQualifyingAdaptor
                objWCF.SetPredicateOnOpenWorkPacket(Irc, ErrMsg)

            Else
                Exit Sub
            End If
            dlg = Nothing
            result = Nothing
        End If
        If Irc <> 0 Then
            If ErrMsg.Contains("There are no Objects to Work in Process ") Then
                Throw New SympraxisException.InformationException(ErrMsg)
            Else
                Throw New SympraxisException.WorkitemException(ErrMsg)
            End If


        End If





        If openworkinterval <> -1 Then
            WorkHourTimer.Interval = 1000
            WorkHourTimer.Enabled = True
            'WorkHourTimer.Start()
        End If
        Logger.WriteToLogFile("IHost", "GetWork Ends")

        DAXObject.SetValue("SYS_OP_ObjectId", objWCF.Objectid)

        WaitCursortext("WorkItem Opening..", "Opening...")
        WaitCursorProgressvalue(60)


        Logger.WriteToLogFile("IHost", "Create IOmanager Directory Starts")
        CreateIOmanagerDirectory()
        If Irc <> 0 Then Exit Sub
        Logger.WriteToLogFile("IHost", "Create IOmanager Directory  Ends")
        'AHWorkitemCount = WCF.WorkitemCount


        SetWorkitemEnvironmentvariable()

        If Irc = 0 Then
            If objWCF.WorkitemCount > 1 Then
                If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not String.IsNullOrEmpty(loadPluginsConfigXml.ApplicationHost.Settings.GotoBack) Then
                    goToBack = loadPluginsConfigXml.ApplicationHost.Settings.GotoBack
                End If
            End If

        End If

        Logger.WriteToLogFile("IHost", "InputParser Starts")
        CheckInputParser()
        WaitCursorProgressvalue(60)

        Logger.WriteToLogFile("IHost", "InputParser Ends")


        If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.ParallelWorkItem) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.ParallelWorkItem.ToUpper = "TRUE" Then
            Logger.WriteToLogFile("IHost", "Load Parallel Workitem Starts")
            LoadParallelWorkitem()
            Logger.WriteToLogFile("IHost", "Load Parallel Workitem Ends")
        Else
        Logger.WriteToLogFile("IHost", "Read Datapart Starts")
        ReadDatapart(fullDP)

        Logger.WriteToLogFile("IHost", "Read Datapart Ends")

            Logger.WriteToLogFile("IHost", "Load Workitem Starts")
            LoadWorkitem(fullDP)        '  HshGetDP = HshFullDP(WCF.Objectid)

            Logger.WriteToLogFile("IHost", "Load Workitem Ends")

            LoadChildworkitem()

        End If

       

        Logger.WriteToLogFile("IHost", "Set Match Banner Starts")
        SetMatchBanner()

        Logger.WriteToLogFile("IHost", "Set Match Banner Ends")


        WaitCursorProgressvalue(80)

        SetApplyLayout()



        Logger.WriteToLogFile("IHost", "Load Work Instructions Starts")
        AHEnableDisableMenu("Help", "WorkInstruction", True)
        LoadWorkInstructions()

        Logger.WriteToLogFile("IHost", "Load Work Instructions Ends")

        Logger.WriteToLogFile("IHost", "Show lastEvent Starts")
        ShowlastEvent()
        Logger.WriteToLogFile("IHost", "Show lastEvent Ends")

        If Not IsNothing(loadPluginsConfigXml) Then
            If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.Applylayout) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.Applylayout.Trim() <> String.Empty Then
                ApplyLayout(loadPluginsConfigXml.ApplicationHost.Settings.Applylayout)
            End If
        End If

        configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True



        SetApplicationTitle()

        Menuhandle()


        WaitCursorProgressvalue(100)

        ShowWaitCursor(False)

        If workingMode <> "\debug" Then
            If Not ObjClientInstanceMgr Is Nothing Then
                Logger.WriteToLogFile("InHost", "ObjClientInstanceMgr.OpenWorkpacket St")
                ObjClientInstanceMgr.OpenWorkpacket(intrHostProcessName)
                Logger.WriteToLogFile("InHost", "ObjClientInstanceMgr.OpenWorkpacket Ed")
            End If
            If Not ObjManagerInstanceMgr Is Nothing Then
                Logger.WriteToLogFile("InHost", "ObjManagerInstanceMgr.OpenWorkpacket St")
                ObjManagerInstanceMgr.OpenWorkpacket(intrHostProcessName)
                Logger.WriteToLogFile("InHost", "ObjManagerInstanceMgr.OpenWorkpacket Ed")
            End If
        End If

        If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.sShowStatusBar) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.sShowStatusBar.ToUpper = "TRUE" Then


            If Not MyDictCount.ContainsKey(intrHostProcessName) Then
                WorkedObjectCount = 0
                CompletedObjectCount = 0
                MyDictCount.Add(intrHostProcessName, CompletedObjectCount.ToString & "|" & WorkedObjectCount.ToString)
            Else
                Dim myval As String = MyDictCount(intrHostProcessName)
                WorkedObjectCount = CInt(myval.Split("|")(1).ToString)
                CompletedObjectCount = CInt(myval.Split("|")(0).ToString)
                MyDictCount.Remove(intrHostProcessName)
                MyDictCount.Add(intrHostProcessName, CompletedObjectCount.ToString & "|" & WorkedObjectCount.ToString)
            End If

            mLayoutmanager.AddingStatusBarControls(True, CompletedObjectCount.ToString, WorkedObjectCount.ToString)
        End If
    End Sub

    Private Sub ToolstripStyle(ByVal Message As String)


        If Not IsNothing(appHostConfigXml.ToolStripStyle) AndAlso appHostConfigXml.ToolStripStyle.Count > 0 Then
            For i = 0 To appHostConfigXml.ToolStripStyle.Count - 1
                If appHostConfigXml.ToolStripStyle(i).Type.ToString.ToUpper = Message.ToUpper() Then

                    Dim BgColor As Color
                    BgColor = ColorConverter.ConvertFromString(appHostConfigXml.ToolStripStyle(i).BGColor)
                    Dim ForeColor As Color
                    ForeColor = ColorConverter.ConvertFromString(appHostConfigXml.ToolStripStyle(i).ForeColor)

                    lblToolstripError.Background = New System.Windows.Media.SolidColorBrush(BgColor)
                    lblToolstripError.Foreground = New System.Windows.Media.SolidColorBrush(ForeColor)
                    ToolstripError.Background = New System.Windows.Media.SolidColorBrush(BgColor)
                    lblToolstripError.FontSize = appHostConfigXml.ToolStripStyle(i).Size

                    Exit Sub
                End If
            Next
        Else
            Dim BgColor As Color
            BgColor = ColorConverter.ConvertFromString("#ffffcc")
            lblToolstripError.Background = New System.Windows.Media.SolidColorBrush(BgColor)
            lblToolstripError.Foreground = Brushes.Black
            ToolstripError.Background = New System.Windows.Media.SolidColorBrush(BgColor)
            lblToolstripError.FontSize = 12
        End If


    End Sub
    Private Sub ShowErrorToolstrip(ByVal Errmsg As String, ByRef ErrorCodetype As SympraxisException.ErrorTypes, ByRef isErrexception As Boolean)
        Try
            If isErrexception = False Then
                IHErrException = Nothing
            End If
            ButtonExport.Visibility = System.Windows.Visibility.Visible

            Dim resourceUri As System.Uri = Nothing
            If ErrorCodetype = SympraxisException.ErrorTypes.AppError Then
                resourceUri = New Uri("Resources/ExpError.png", UriKind.Relative)
                ToolstripStyle("ERROR")
            ElseIf ErrorCodetype = SympraxisException.ErrorTypes.Warning Then
                resourceUri = New Uri("Resources/ExpWarning.png", UriKind.Relative)
                ToolstripStyle("WARNING")
            ElseIf ErrorCodetype = SympraxisException.ErrorTypes.Information Then
                resourceUri = New Uri("Resources/ExpInfo.png", UriKind.Relative)
                ToolstripStyle("INFORMATION")
            End If

            Dim streamInfo As StreamResourceInfo = Application.GetResourceStream(resourceUri)
            Dim temp As BitmapFrame = BitmapFrame.Create(streamInfo.Stream)
            Dim brush = New ImageBrush()

            brush.ImageSource = temp
            lblerrIcon.Background = brush

            If Errmsg.Trim <> "" Then
                If IsNothing(IHErrException) Or (Not IsNothing(IHErrException) AndAlso IHErrException.Message.Trim() = "") Then
                    ButtonExport.Visibility = System.Windows.Visibility.Collapsed
                End If


                ToolstripError.Visibility = System.Windows.Visibility.Visible
                lblToolstripError.Text = Errmsg

                If Errmsg = "WorkPacket Saved Successfully" Then
                    HideInfobarTimerStart(3000)
                Else
                    HideInfobarTimerStart(30000)
                End If

                Logger.WriteToLogFile("Ihost Error:", Errmsg)
            End If
            Errortype = ErrorCodetype
        Catch ex As Exception
            ToolstripError.Visibility = System.Windows.Visibility.Collapsed


        End Try
    End Sub
    Private Sub chkclosingcomplete(ByRef isErrormached As Boolean)
        If IsAbortValidation = False Then
            If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.ParallelWorkItem) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.ParallelWorkItem.ToUpper = "TRUE" Then
                If objWCF.CurrentWorkitemIdx < objWCF.WorkitemCount Then
                    Exit Sub
                End If

            End If
            If isExitflag = True Then
                ShowWaitCursor(False)
                If MessageBox.Show("Current workitem not complete.Do you want to push?", "Alert Message", MessageBoxButtons.YesNo) = Forms.DialogResult.No Then
                    isErrormached = True

                    ShowWaitCursor(True, "Closing..", "Closing...")

                    WaitCursorProgressvalue(15)
                    Exit Sub
                Else
                    ShowWaitCursor(True, "Closing..", "Closing...")
                    WaitCursorProgressvalue(15)
                End If
            Else
                ShowErrorToolstrip("Current workitem is Incomplete, cannot move to Next workitem", SympraxisException.ErrorTypes.Information, False)
                isErrormached = True
                Exit Sub
            End If
        End If
    End Sub
    Private Sub ClosingRules(ByRef isErrormached As Boolean, ByVal ChildID As String)
        Dim listMatchedRulesval As List(Of SetClosingRuleManifest) = Nothing
        isErrormached = False
        Dim Irc As Integer = 0
        Dim Errmsg As String = String.Empty
        Dim iscompleted As Boolean = False
        Try
            If Not IsNothing(loadPluginsConfigXml.ApplicationHost.ClosingRules) Then
                ' For Error message
                Logger.WriteToLogFile("IHost", "Closing Rules Starts")
                If Not IsNothing(loadPluginsConfigXml.ApplicationHost.ClosingRules.Errors) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.ClosingRules.Errors.SetRulesval) Then

                    MatchClosingRules(listMatchedRulesval, loadPluginsConfigXml.ApplicationHost.ClosingRules.Errors.SetRulesval, ChildID)


                    If Not IsNothing(listMatchedRulesval) AndAlso listMatchedRulesval.Count > 0 Then
                        If listMatchedRulesval(0).MSG.ToString() <> "" Then

                            ShowErrorToolstrip(listMatchedRulesval(0).MSG.ToString(), SympraxisException.ErrorTypes.Information, False)

                            isErrormached = True
                            If listMatchedRulesval(0).ErrorCode <> "" Then
                                DAXObject.SetValue("SYS_VALIDATIONERROR", listMatchedRulesval(0).ErrorCode)
                            End If


                            Exit Sub
                        End If
                    End If
                End If
                listMatchedRulesval = Nothing
                ' For Alert
                If Not IsNothing(loadPluginsConfigXml.ApplicationHost.ClosingRules.Alerts) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.ClosingRules.Alerts.SetRulesval) Then
                    MatchClosingRules(listMatchedRulesval, loadPluginsConfigXml.ApplicationHost.ClosingRules.Alerts.SetRulesval, ChildID)



                    If Not IsNothing(listMatchedRulesval) AndAlso listMatchedRulesval.Count > 0 Then
                        For i As Integer = 0 To listMatchedRulesval.Count - 1
                            If IsAbortValidation = True Then
                                If listMatchedRulesval(i).ErrorCode <> "" Then
                                    DAXObject.SetValue("SYS_VALIDATIONERROR", listMatchedRulesval(i).ErrorCode)
                                End If
                            Else

                                ShowWaitCursor(False)
                                If MessageBox.Show(listMatchedRulesval(i).MSG, "Alert Message", MessageBoxButtons.YesNo) = Forms.DialogResult.No Then
                                    isErrormached = True
                                    If listMatchedRulesval(i).ErrorCode <> "" Then
                                        DAXObject.SetValue("SYS_VALIDATIONERROR", listMatchedRulesval(i).ErrorCode)
                                    End If
                                    ShowWaitCursor(True, "Closing..", "Closing...")

                                    WaitCursorProgressvalue(15)
                                    Exit Sub
                                Else
                                    ShowWaitCursor(True, "Closing..", "Closing...")
                                    WaitCursorProgressvalue(15)
                                End If
                            End If

                        Next
                    End If
                End If
                listMatchedRulesval = Nothing
                Dim ObjSaveStatus As String = String.Empty

                'Exit Parser
                Logger.WriteToLogFile("IHost", "Exit Parser Starts")
                ExitParser(ObjSaveStatus, ChildID)
                Logger.WriteToLogFile("IHost", "Exit Parser Ends")

                If ObjSaveStatus.ToUpper() = "ERROR" Then
                    If IsAbortValidation = False Then
                        chkclosingcomplete(isErrormached)
                        If isErrormached Then
                            Exit Sub
                        End If

                    End If
                    objWCF.UpdateSaveWFStatus("Error", Irc, Errmsg)
                    If Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.ChildWFsts) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.ChildWFsts.ToUpper() = "TRUE" Then
                        objWCF.UpdateChildSaveWFStatus("Error", ChildID, Irc, Errmsg)
                    Else
                        objWCF.UpdateChildSaveWFStatus("Error", "ALL", Irc, Errmsg)

                    End If
                    If Irc <> 0 Then
                        Throw New Exception()
                    End If

                ElseIf ObjSaveStatus.ToUpper() = "COMPLETE" Then
                    objWCF.UpdateSaveWFStatus("Complete", Irc, Errmsg)
                    If Irc <> 0 Then
                        Throw New Exception()
                    End If
                    If Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.ChildWFsts) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.ChildWFsts.ToUpper() = "TRUE" Then
                        objWCF.UpdateChildSaveWFStatus("Complete", ChildID, Irc, Errmsg)
                    Else
                        objWCF.UpdateChildSaveWFStatus("Complete", "ALL", Irc, Errmsg)

                    End If
                    iscompleted = True
                End If

                listMatchedRulesval = Nothing
                ' For Completed
                If ObjSaveStatus = String.Empty Then

                    If Not IsNothing(loadPluginsConfigXml.ApplicationHost.ClosingRules.Completion) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.ClosingRules.Completion.SetRulesval) Then
                        MatchClosingRules(listMatchedRulesval, loadPluginsConfigXml.ApplicationHost.ClosingRules.Completion.SetRulesval, ChildID)

                        If Not IsNothing(listMatchedRulesval) AndAlso listMatchedRulesval.Count > 0 Then
                            objWCF.UpdateSaveWFStatus("Complete", Irc, Errmsg)
                            If Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.ChildWFsts) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.ChildWFsts.ToUpper() = "TRUE" Then
                                objWCF.UpdateChildSaveWFStatus("Complete", ChildID, Irc, Errmsg)
                            Else
                                objWCF.UpdateChildSaveWFStatus("Complete", "ALL", Irc, Errmsg)

                            End If
                            If Irc <> 0 Then
                                Throw New Exception()
                            End If
                            iscompleted = True
                        Else
                            chkclosingcomplete(isErrormached)
                            If isErrormached Then
                                Exit Sub
                        End If
                    End If
                    Else
                        chkclosingcomplete(isErrormached)
                        If isErrormached Then
                            Exit Sub
                        End If
                    End If

                End If
                listMatchedRulesval = Nothing
            End If

            Logger.WriteToLogFile("IHost", "Closing Rules Ends")


            If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.sShowStatusBar) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.sShowStatusBar.ToUpper = "TRUE" Then



                If MyDictCount.ContainsKey(intrHostProcessName) Then
                    WorkedObjectCount = CInt(MyDictCount(intrHostProcessName).Split("|")(1).ToString()) + 1
                    If iscompleted = True Then
                        CompletedObjectCount = CInt(MyDictCount(intrHostProcessName).Split("|")(0).ToString()) + 1
                    Else
                        CompletedObjectCount = CInt(MyDictCount(intrHostProcessName).Split("|")(0).ToString())
                    End If
                    MyDictCount.Remove(intrHostProcessName)
                    MyDictCount.Add(intrHostProcessName, CompletedObjectCount.ToString & "|" & WorkedObjectCount.ToString)
                End If

                mLayoutmanager.AddingStatusBarControls(True, CompletedObjectCount.ToString, WorkedObjectCount.ToString)
            End If

        Finally

            iscompleted = False
            listMatchedRulesval = Nothing

        End Try
    End Sub

    Private Sub MatchClosingRules(ByRef ReturnMatchedRulesvallist As List(Of SetClosingRuleManifest), ByRef SetRulesvalManifest As List(Of SetClosingRuleManifest), ByRef childId As String)

        Dim MatchedRule As List(Of Rule) = Nothing
        Dim CrRule As List(Of Rule) = Nothing
        Dim LstRulelist As List(Of Rule) = Nothing
        Dim listSetRulesval As List(Of SetClosingRuleManifest) = SetRulesvalManifest

        Dim Irc As Integer = 0
        Dim Errmsg As String = String.Empty


        Try



            LstRulelist = loadPluginsConfigXml.ApplicationHost.ClosingRules.Rules.Ruleslist.Where(Function(x) listSetRulesval.Any(Function(y) y.RULEID = x.Name)).ToList()


            GetCrRulelist(CrRule, LstRulelist, loadPluginsConfigXml.ApplicationHost.ClosingRules.Rules.Ruleslist)

            Logger.WriteToLogFile("IHost", "GetMatchRule Starts")

            If childId = "" Then
            objWCF.GetMatchRule(MatchedRule, LstRulelist, CrRule, Irc, Errmsg)

            Else
                objWCF.GetChildMatchRule(MatchedRule, LstRulelist, CrRule, childId, Irc, Errmsg)

            End If




            If Irc <> 0 Then
                Throw New SympraxisException.AppException(Errmsg)
            End If
            Logger.WriteToLogFile("IHost", "GetMatchRule Ends")

            If Not IsNothing(MatchedRule) Then
                If MatchedRule.Count > 0 Then
                    ReturnMatchedRulesvallist = listSetRulesval.Where(Function(x) MatchedRule.Any(Function(y) y.Name = x.RULEID)).ToList()
                End If
            End If
            listSetRulesval = Nothing

        Finally
            MatchedRule = Nothing
            CrRule = Nothing
            LstRulelist = Nothing
        End Try
    End Sub

    Private Sub GetCrRulelist(ByRef ReturnCrRule As List(Of Rule), ByRef Rulelist As List(Of Rule), ByRef ConfigRuleslist As List(Of Rule))

        If Not IsNothing(Rulelist) AndAlso Not IsNothing(ConfigRuleslist) Then


            Dim CrSinglematch As List(Of String) = Rulelist.Where(Function(x) Convert.ToString(x.ConditionReference) <> "" AndAlso x.ConditionReference.Contains("|") = False AndAlso x.ConditionReference.Contains("+") = False).Select(Function(x) x.ConditionReference).ToList()
            Dim OrRulelist As List(Of CrRuleList) = Rulelist.Where(Function(x) Convert.ToString(x.ConditionReference) <> "" AndAlso x.ConditionReference.Contains("|")).Select(Function(x) New CrRuleList With {.Crmatch = x.ConditionReference.Split("|").ToList()}).ToList()
            Dim AndRulelist As List(Of CrRuleList) = Rulelist.Where(Function(x) Convert.ToString(x.ConditionReference) <> "" AndAlso x.ConditionReference.Contains("+")).Select(Function(x) New CrRuleList With {.Crmatch = x.ConditionReference.Split("+").ToList()}).ToList()
            OrRulelist.AddRange(AndRulelist)
            ReturnCrRule = ConfigRuleslist.Where(Function(x) OrRulelist.Any(Function(y) y.Crmatch.Any(Function(k) k = x.Name)) Or CrSinglematch.Any(Function(s) s = x.Name)).Select(Function(x) x).ToList()

        End If
    End Sub

    Private Sub Pathvalidation(ByRef Path As String)
        If Path.Length = 0 Then
            Throw New Exception("Path is Empty")

        End If

        Dim lastchar As Char = Path(Path.Length - 1)
        If lastchar <> "\" Then
            Path = Path & "\"
        ElseIf IOManager.CheckDirectoryExists(Path) = False Then
            Throw New SympraxisException.SettingsException(" Path :" & Path & " is not Exist")

        End If
    End Sub



    Private Sub LoadStickers()

        If Not IsNothing(appHostConfigXml.Sticker) Then

            If Not IsNothing(appHostConfigXml.Sticker.StickerFolder) AndAlso appHostConfigXml.Sticker.StickerFolder <> "" Then

                Pathvalidation(appHostConfigXml.Sticker.StickerFolder)




                If Not IsNothing(appHostConfigXml.Sticker.DefaultStickerName) AndAlso appHostConfigXml.Sticker.DefaultStickerName <> "" Then
                    If IOManager.CheckFileExists(String.Concat(appHostConfigXml.Sticker.StickerFolder, appHostConfigXml.Sticker.DefaultStickerName)) Then
                        If Not IsNothing(appHostConfigXml.Sticker.StickerWidth) AndAlso appHostConfigXml.Sticker.StickerWidth <> "" Then
                            objFrmSplash.MySplashControl.LogoWidth = appHostConfigXml.Sticker.StickerWidth
                        Else
                            objFrmSplash.MySplashControl.LogoWidth = 80
                        End If
                        objFrmSplash.MySplashControl.LogoPath = String.Concat(appHostConfigXml.Sticker.StickerFolder, appHostConfigXml.Sticker.DefaultStickerName)
                        'Else
                        '    Throw New Exception("Unable to locate/access the project DefaultSticker File")

                    End If
                End If

                If Not IsNothing(appHostConfigXml.Sticker.AppLogo) AndAlso appHostConfigXml.Sticker.AppLogo <> "" Then
                    If IOManager.CheckFileExists(String.Concat(appHostConfigXml.Sticker.StickerFolder, appHostConfigXml.Sticker.AppLogo)) Then
                        objFrmSplash.MySplashControl.AppLogoPath = String.Concat(appHostConfigXml.Sticker.StickerFolder, appHostConfigXml.Sticker.AppLogo)
                    Else
                        Throw New Exception("Unable to locate/access the project logo File")


                    End If
                    'Else
                    '    Throw New Exception("Missing configuraion details for Project Logo Environment")

                End If

                If Not IsNothing(appHostConfigXml.Sticker.Poweredby) AndAlso appHostConfigXml.Sticker.Poweredby <> "" Then
                    objFrmSplash.MySplashControl._Poweredbyflg = appHostConfigXml.Sticker.Poweredby
                End If
            End If
        Else
            ' Throw New Exception("Missing configuraion details for Project Logo Environment ")


            Exit Sub
        End If






    End Sub

    Private Sub loadWorkGroupStickers(ByVal Workgroup As String)

        If Not IsNothing(appHostConfigXml.Sticker) AndAlso Not IsNothing(appHostConfigXml.Sticker.Stickers) AndAlso appHostConfigXml.Sticker.Stickers.Count > 0 AndAlso Not IsNothing(appHostConfigXml.Sticker.StickerFolder) AndAlso appHostConfigXml.Sticker.StickerFolder <> "" Then

            Dim WGStk As WGStrickers = appHostConfigXml.Sticker.Stickers.Where(Function(x) x.WG = Workgroup).FirstOrDefault()
            If Not IsNothing(WGStk) Then

                If Not String.IsNullOrEmpty(WGStk.StickerName) Then

                    If IOManager.CheckFileExists(String.Concat(appHostConfigXml.Sticker.StickerFolder, WGStk.StickerName)) Then
                        If Not IsNothing(WGStk.stickerWidth) AndAlso WGStk.stickerWidth <> "" Then
                            objFrmSplash.MySplashControl.LogoWidth = WGStk.stickerWidth
                        Else
                            objFrmSplash.MySplashControl.LogoWidth = 80
                        End If
                        objFrmSplash.MySplashControl.LogoPath = String.Concat(appHostConfigXml.Sticker.StickerFolder, WGStk.StickerName)
                    Else

                        Throw New Exception("Unable to locate/access the project WorkGroupSticker File")

                    End If

                End If
            End If

        End If



    End Sub




    Private Sub AHEnableMenu(ByVal Headername As String, ByVal Menuname As String)

        Dim Disablemenu As String() = Menuname.Split("|")
        For Each RibTab As RibbonTab In RibbonWin.Items

            If Not IsNothing(RibTab) Then
                If RibTab.Header.ToString().ToLower() = Headername.ToLower() Then
                    For Each sRibGroup As RibbonGroup In RibTab.Items

                        For Each RibButton As Object In sRibGroup.Items
                            Dim aRibButton As RibbonButton = TryCast(RibButton, RibbonButton)
                            If Not IsNothing(aRibButton) Then
                                If Disablemenu.Contains(RibButton.Label) Then
                                    If RibButton.Label.ToString().ToUpper() = "CLOSE" Then
                                        BtncloseWork.Visibility = System.Windows.Visibility.Visible
                                    End If
                                    RibButton.Visibility = System.Windows.Visibility.Visible
                                    If Headername.ToLower() = "home" Then
                                        CreatshortCutforMenu(RibButton.Label, True)
                                    End If
                                Else
                                    RibButton.Visibility = System.Windows.Visibility.Collapsed
                                    If RibButton.Label.ToString().ToUpper() = "CLOSE" Then
                                        BtncloseWork.Visibility = System.Windows.Visibility.Collapsed
                                    End If
                                    If Headername.ToLower() = "home" Then
                                        CreatshortCutforMenu(RibButton.Label, False)
                                    End If
                                End If
                            End If

                        Next
                    Next
                End If
            End If

        Next


    End Sub
    Private Sub ReorderHelpTab()

            Dim deletetab As RibbonTab = Nothing
            For Each RibTab As RibbonTab In RibbonWin.Items
            If RibTab.Header.ToString().ToLower() = "help" Then
                deletetab = RibTab
                Exit For
            End If
            Next
            If Not IsNothing(deletetab) Then
                RibbonWin.Items.Remove(deletetab)
            RibbonWin.Items.Add(deletetab)
            deletetab.Visibility = System.Windows.Visibility.Visible
            RibbonWin.UpdateLayout()
            Me.UpdateLayout()
            End If


    End Sub

  
    Private Sub AHEnableDisableMenu(ByVal Headername As String, ByVal Menuname As String, ByVal Enable As Boolean)


        If Headername <> "" AndAlso Headername.Substring(0, 1) = "&" Then
            Headername = Headername.Remove(0, 1)
        End If
        If Menuname <> "" Then
            Dim Disablemenu As String() = Menuname.Split("|")
            For Each RibTab As RibbonTab In RibbonWin.Items
                If RibTab.Header.ToString().ToLower() = Headername.ToLower() Then
                    If Enable = True Then

                        RibTab.Visibility = System.Windows.Visibility.Visible
                    End If

                    For Each sRibGroup As RibbonGroup In RibTab.Items
                        For Each RibButton In sRibGroup.Items
                            If Disablemenu.Contains(RibButton.Label) Then
                                If Enable = True Then
                                    RibButton.IsEnabled = True
                                    RibButton.Visibility = System.Windows.Visibility.Visible

                                Else
                                    RibButton.IsEnabled = False
                                    RibButton.Visibility = System.Windows.Visibility.Collapsed
                                End If


                            End If

                        Next
                    Next
                End If
            Next
        Else
            For Each RibTab As RibbonTab In RibbonWin.Items
                If RibTab.Header.ToString().ToLower() = Headername.ToLower() Then
                    For Each sRibGroup As RibbonGroup In RibTab.Items
                        For Each RibButton In sRibGroup.Items

                            If Enable = True Then
                                RibButton.IsEnabled = True
                                RibButton.Visibility = System.Windows.Visibility.Visible

                            Else
                                RibButton.IsEnabled = False
                                RibButton.Visibility = System.Windows.Visibility.Collapsed
                            End If




                        Next
                    Next
                    If Enable = True Then

                        RibTab.Visibility = System.Windows.Visibility.Visible

                    Else

                        RibTab.Visibility = System.Windows.Visibility.Collapsed
                    End If


                End If
            Next
        End If


    End Sub

    Private Sub ShowException(ByRef Ex As Exception)
        If Not IsNothing(ObjFrmException) Then
            ShowWaitCursor(False, "", "")
            ObjFrmException.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
            ObjFrmException.ShowError(Ex)
            ObjFrmException.ShowDialog()
            lastLogger = String.Empty
        End If
    End Sub

    'Private Sub ShowException(ByRef DetailsError As String, ByRef strError As String, ByRef strHeader As String, ByRef strFooter As String, ByRef ErrorType As SympraxisException.ErrorTypes)

    '    ShowWaitCursor(False, "", "")
    '    ObjFrmException.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
    '    ObjFrmException.ShowError(DetailsError, strError, strHeader, strFooter, ErrorType)
    '    ObjFrmException.ShowDialog()

    'End Sub

    'Private Sub ShowException(ByRef ErrorDetails As SympraxisException.ExceptionDetails)

    '    If Not IsNothing(ErrorDetails) Then
    '        ShowWaitCursor(False, "", "")

    '        If (IsNothing(ErrorDetails.ErrFooter)) Or (Not IsNothing(ErrorDetails.ErrFooter) AndAlso ErrorDetails.ErrFooter = String.Empty) Then
    '            If Not IsNothing(ErrorDetails.ErrCode) AndAlso ErrorDetails.ErrCode <> "" Then
    '                Dim Errcode As String = ErrorDetails.ErrCode
    '                ErrorDetails.ErrFooter = appHostConfigXml.ExceptionInfo.Where(Function(x) If(x.ErrorCode.Contains("|"), x.ErrorCode.Split("|").Any(Function(y) y = Errcode), x.ErrorCode = Errcode)).Select(Function(x) x.ErrorMsg).FirstOrDefault()
    '            End If

    '        End If

    '        ObjFrmException.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
    '        ObjFrmException.ShowError(ErrorDetails)
    '    End If



    'End Sub

    Private Sub ButtonpinVisible()
        txtACStatus.Text = "OFF"
        txtAutoClose.Foreground = New SolidColorBrush(Colors.DarkBlue)
        txtACStatus.Foreground = New SolidColorBrush(Colors.DarkBlue)
        ManualIdleStarttime = DateTime.Now
        ButtonPin.Visibility = System.Windows.Visibility.Visible
    End Sub

    Private Sub ShowNotificationbar(ByVal message As String, notify As System.Xml.XmlDocument, ByVal size As Integer)

        pnlNotifyHeader.Visibility = System.Windows.Visibility.Visible
        If Not IsNothing(notify) Then
            Dispatcher.Invoke(CType(AddressOf EnableNotificationPanel, DelEnableNotificationPanel), notify, size)
        End If


    End Sub

    Private Sub EnableNotificationPanel(input As System.Xml.XmlDocument, size As Integer)
        Dim rc As Int32 = 0
        Dim Documents As XmlDocument
        Dim PanelCaseInfo As New Sympraxis.Plugins.CaseInfo.CaseInfoControl

        If Not IsNothing(overrideAppHostConfig.PanelNotify) Then

            pnlNotify.Visibility = System.Windows.Visibility.Visible
            If overrideAppHostConfig.PanelNotify.Height <> 0 Then
                If size = 0 Then
                    DisplayText.Height = overrideAppHostConfig.PanelNotify.Height
                    pnlNotify.Height = overrideAppHostConfig.PanelNotify.Height
                    BtnNotificationOpen.Height = overrideAppHostConfig.PanelNotify.Height
                    BtnNotificationClose.Height = overrideAppHostConfig.PanelNotify.Height
                Else
                    pnlNotify.Height = 30
                    DisplayText.Height = 30

                    BtnNotificationOpen.Height = 30
                    BtnNotificationClose.Height = 30

                End If
            End If
            If IsNothing(PanelCaseInfo) Then
                PanelCaseInfo = New Plugins.CaseInfo.CaseInfoControl
            End If

            pnlNotify.Children.Clear()
            pnlNotify.Children.Add(PanelCaseInfo)
            'PanelElementHost.Child = Nothing
            'PanelElementHost.Child = PanelCaseInfo
            rc = PanelCaseInfo.LoadXaml(overrideAppHostConfig.PanelNotify.Path)
            If rc = 0 Then
                Documents = New Xml.XmlDocument
                Dim provider As New XmlDataProvider
                provider = PanelCaseInfo.GetXmlDoc()
                ''provider.Document = input
                Documents.LoadXml(provider.Document.InnerXml)

                Documents.LoadXml(input.OuterXml)
                PanelCaseInfo.UpdateXmlDoc(Documents)
                Dim res As System.Xml.XmlNode
                res = input.DocumentElement()
                ''PanelCaseInfo.LoadDocument(input)

            Else
                pnlNotify.Height = 0
            End If

        Else
            pnlNotify.Visibility = System.Windows.Visibility.Collapsed
        End If

    End Sub


    Private Sub PlayNotificationSoundAlert()
        Dim soundAlert As SoundPlayer = Nothing
        If Not IsNothing(overrideAppHostConfig.PanelNotify.sNotificationSound) AndAlso overrideAppHostConfig.PanelNotify.sNotificationSound <> String.Empty _
            AndAlso overrideAppHostConfig.PanelNotify.sNotificationSound.ToUpper() = "TRUE" Then

            If Not IsNothing(overrideAppHostConfig.PanelNotify.sSoundAlertPath) AndAlso overrideAppHostConfig.PanelNotify.sSoundAlertPath <> String.Empty Then
                If IOManager.CheckFileExists(overrideAppHostConfig.PanelNotify.sSoundAlertPath) Then
                    Dim alertpath As String = overrideAppHostConfig.PanelNotify.sSoundAlertPath
                    If IsNothing(soundAlert) Then soundAlert = New SoundPlayer(alertpath)
                    soundAlert.Play()
                Else
                    System.Media.SystemSounds.Asterisk.Play()
                End If
            Else
                System.Media.SystemSounds.Asterisk.Play()

            End If

            soundAlert = Nothing
        End If
    End Sub

    Private Sub ShowNotification(ByVal NotifyCode As String, ByVal Showmessage As String, ByRef notifydoc As XmlDocument)
        BlnOpenAftClose = False
        If NotifyCode.ToUpper = Sympraxis.Utilities.INormalPlugin.IPluginhost.NotifyCode.NewWorkAssignedToUser.ToString.ToUpper Then '' New Work Assigned To User
            DisplayText.Visibility = System.Windows.Visibility.Visible
            DisplayText.Content = "New"

            If configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = False Then
                BtnNotificationOpen.Content = "Open"
                BtnNotificationOpen.Width = 69
                BtnNotificationOpen.Focus()
            ElseIf configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True Then
                If Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.AutoOpenNextNotify) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.AutoOpenNextNotify.ToUpper() = "TRUE" Then
                    BlnOpenAftClose = True
                End If
                BtnNotificationOpen.Content = "Open Next"
                BtnNotificationOpen.Width = 69
            End If

        ElseIf NotifyCode.ToUpper = Sympraxis.Utilities.INormalPlugin.IPluginhost.NotifyCode.RequestedWorkAssigned.ToString.ToUpper Then ''Requested Work Assigned
            DisplayText.Visibility = System.Windows.Visibility.Visible
            DisplayText.Content = "Ready"

            If configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = False Then
                BtnNotificationOpen.Content = "Open"
                BtnNotificationOpen.Width = 69
                BtnNotificationOpen.Focus()
            ElseIf configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True Then
                If Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.AutoOpenNextNotify) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.AutoOpenNextNotify.ToUpper() = "TRUE" Then
                    BlnOpenAftClose = True
                End If
                BtnNotificationOpen.Content = "Open Next"
                BtnNotificationOpen.Width = 69
            End If
        ElseIf NotifyCode.ToUpper = Sympraxis.Utilities.INormalPlugin.IPluginhost.NotifyCode.WorkReassignedToUser.ToString.ToUpper Then 'Work Reassigned To User
            DisplayText.Visibility = System.Windows.Visibility.Collapsed
            BtnNotificationOpen.Visibility = System.Windows.Visibility.Collapsed
        End If

        Logger.WriteToLogFile("Ihost", "ShowNotificationbar Start")
        
            Dispatcher.BeginInvoke(CType(AddressOf ShowNotificationbar, DelShowNotificationbar), Showmessage, notifydoc, 0)

        Logger.WriteToLogFile("Ihost", "ShowNotificationbar Ends")

        Logger.WriteToLogFile("Ihost", "PlayNotificationSoundAlert Start")
        PlayNotificationSoundAlert()

        Logger.WriteToLogFile("Ihost", "PlayNotificationSoundAlert Ends")

    End Sub

    ' Unwanted Code
#Region " UnwantedCode"


    Public Class testerr1
        Public ant1 As Integer
    End Class

    Private Class testerr
        Inherits testerr1

        Public ant As Integer
        Public Sub a()
            Dim obj1 As New testerr1
            obj1.ant1 = "vxc"
        End Sub

    End Class
#End Region

    Private Function LoadApp()

        Try

            Dim Irc As Integer = 0





            objFrmSplash = New SplashForm
            'Dim a As Integer = 1
            'Dim b As Integer = 0
            'a = a / b
            'Dim a As String = "sdfsdfsd"
            'a.Substring(55)

            SplashForm.InitializeSplashScreen(objFrmSplash)

            SplashForm.ShowForm(objFrmSplash)
            objFrmSplash.MySplashControl.lblStatus.Text = "Application is Loading... Please Wait..."
            objFrmSplash.MySplashControl.lblStatus.Text = "Reading Profile....."
            objFrmSplash.MySplashControl.Refresh()

            LoadConfig(appHostConfigXml, loadPluginsConfigXml, True)

            If IMClientPort = "" Then
                RemoveDuplicateApplication()
            End If


            objFrmSplash.MySplashControl.lblAppTitle.Text = intrHostProcessName

            Logger.WriteToLogFile("IHost", "Load Stickers Start")
            LoadStickers()


            Logger.WriteToLogFile("IHost", "Load Stickers Completed ")

            objFrmSplash.MySplashControl.lblStatus.Text = "Loading  WCF....."
            objFrmSplash.MySplashControl.Refresh()

            Logger.WriteToLogFile("IHost", "Load WCF Start")
            LoadWCF(overrideAppHostConfig.WFClientFrameWork)


            Logger.WriteToLogFile("IHost", "Load WCF Completed")




            Dim workGroup As String = String.Empty
            If Irc = 0 AndAlso workingMode <> "\debug" AndAlso workingMode.ToUpper() <> "OFFLINE" AndAlso (processType.ToUpper() = "SAWUA" Or processType.ToUpper() = "WORKFLOW") Then
                objFrmSplash.MySplashControl.lblStatus.Text = "Logging In....."
                objFrmSplash.MySplashControl.Refresh()

                Logger.WriteToLogFile("IHost", "LoginUser Start")
                LoginUser()


                Logger.WriteToLogFile("IHost", "LoginUser Completed")
                workGroup = objWCF.sWorkGroup
                If workGroup <> String.Empty Then
                    Logger.WriteToLogFile("IHost", "load WorkGroup Stickers Start")
                    loadWorkGroupStickers(workGroup)

                    Logger.WriteToLogFile("IHost", "loadWorkGroup Stickers Completed")
                End If
            End If
            Logger.WriteToLogFile("IHost", "Assign EnvironmentVariable Start")
            EnvironmentVariable()

            If workingMode.ToUpper() <> "OFFLINE" Then
                configMgr.AppbaseConfig.EnvironmentVariables.IsOnline = True
            Else
                configMgr.AppbaseConfig.EnvironmentVariables.IsOnline = False
            End If

            Logger.WriteToLogFile("IHost", "Assign EnvironmentVariable Start")

            If Irc = 0 AndAlso appStartBy.ToUpper() <> "STARTUP" AndAlso workingMode.ToUpper() <> "OFFLINE" AndAlso workingMode <> "\debug" AndAlso processType.ToUpper() = "WORKFLOW" Then


                objFrmSplash.MySplashControl.lblStatus.Text = String.Concat("Attaching Process ", intrHostProcessName, " .....")
                objFrmSplash.MySplashControl.Refresh()
                Logger.WriteToLogFile("IHost", "Attach Process Start")
                AttachProcess(True)


                Logger.WriteToLogFile("Completed", "Start Attach Process")
            End If



            If Irc = 0 AndAlso workingMode <> "\debug" AndAlso workingMode.ToUpper() <> "OFFLINE" AndAlso appStartBy.ToUpper() <> "STARTUP" AndAlso processType.ToUpper() = "WORKFLOW" Then
                objFrmSplash.MySplashControl.lblStatus.Text = "Registering WorkFlow....."
                objFrmSplash.MySplashControl.Refresh()
                objFrmSplash.MySplashControl.lblStatus.Text = String.Concat("Registering workflow ", intrHostWorkFlowName, " .....")
                objFrmSplash.MySplashControl.Refresh()
                Logger.WriteToLogFile("IHost", "Register WorkFlow Start")
                RegisterWF(True)

                Logger.WriteToLogFile("IHost", "Complete Register WorkFlow")
            End If

            If Not IsNothing(appHostConfigXml.InstanceMgr) Then
                If IMClientPort = "" Then

                    InitIMServercommunication()
                Else
                    ConnectingInstanceMethod()
                    If Not ObjClientInstanceMgr Is Nothing Then
                        ObjClientInstanceMgr.InsMgrUpdateDAX()
                    End If
                End If


            End If

            Logger.WriteToLogFile("IHost", "Update EnvironmentVariable Start")
            UpdateEnvironmentvariable()

            Logger.WriteToLogFile("IHost", "Update EnvironmentVariable Completed")
            If workingMode <> "\debug" AndAlso workingMode.ToUpper() <> "OFFLINE" AndAlso (processType.ToUpper() = "SAWUA" Or processType.ToUpper() = "SAWOUA") Then
                CreateMenu("Home", "Exit")
            Else
                CreateMenu("Home", "Open|Previous|Next|Close|Exit")
            End If

            CreateMenu("Help", "WorkInstruction|ShortCuts|About")


            objFrmSplash.MySplashControl.lblStatus.Text = "Loading Plugins....."
            objFrmSplash.MySplashControl.Refresh()

            Logger.WriteToLogFile("IHost", "Load Plugins Start")
            LoadPlugins(appHostConfigXml.PluginDefinitions, loadPluginsConfigXml.ApplicationHost.Plugins, loadPluginsConfigXml.ApplicationHost.LayoutSettings)

            Logger.WriteToLogFile("IHost", "Load Plugins Ends")
            'Temp Code To create Menu

            'If workingMode <> "\debug" AndAlso workingMode.ToUpper() <> "OFFLINE" AndAlso (processType.ToUpper() = "WORKFLOW" Or processType.ToUpper() = "SAWUA" Or processType.ToUpper() = "SAWOUA") Then
            '    AHEnableMenu("Work", "Exit")

            '    RibbonWin.IsMinimized = True
            'ElseIf workingMode = "\debug" Or appStartBy.ToUpper() <> "STARTUP" Then
            '    AHEnableMenu("Work", "Open|Exit")

            '    RibbonWin.IsMinimized = True
            'Else
            '    AHEnableMenu("Work", "Exit")
            'End If

            Menuhandle()
            RibbonWin.IsMinimized = True

            '' remove

            If configMgr.AppbaseConfig.EnvironmentVariables.IsOnline = False Then
                System.Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(Sub() RefreshOfflineGrid())
            End If


            If Irc = 0 AndAlso workingMode <> "\debug" AndAlso workingMode.ToUpper() <> "OFFLINE" Then
                OffileAutoOpen()
            End If
            AHEnableDisableMenu("Help", "WorkInstruction", False)
            SetApplicationTitle()
            CreateRibbonHelpContentMenu()
        Finally

            '   CreateMenu("Work", " Reset ( ALT+F10)")
            '  ShowErrorToolstrip("WorkPacket Saved Successfully", SympraxisException.ErrorTypes.Information, False)
            SplashForm.CloseForm(objFrmSplash)
            objFrmSplash = Nothing


            'If Irc <> 0 Then
            '    System.Windows.MessageBox.Show("Error in loadApp :" & ErrMsg)
            'End If
        End Try

    End Function

    Private Sub DoEvents()
        If isformactive = True Then


        Dim frame As New DispatcherFrame()

        Dim disp As New DispatcherOperationCallback(AddressOf ExitFrame)
        Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background, disp, frame)
        Dispatcher.PushFrame(frame)

        frame = Nothing
        disp = Nothing
        End If
    End Sub


    Private Sub dispatcherTimer_Tick(ByVal sender As Object, ByVal e As EventArgs)

        'waitCursorSec = waitCursorSec + 1
        'Dim ts As TimeSpan = TimeSpan.FromSeconds(waitCursorSec)
        'LblTimer.Content = String.Format("{0}", New DateTime(ts.Ticks).ToString("mm:ss"))

        'CommandManager.InvalidateRequerySuggested()

        'DoEvents()

    End Sub

    Private Function ExitFrame(ByVal f As Object) As Object
        CType(f, DispatcherFrame).Continue = False
        Return Nothing
    End Function
    Private Sub LoadPluginsConfig(c)

        ' LOAD ARRANGE PLUGIN CONFIG WHILE EACH PROCESS SWITCH


        If appStartBy.ToUpper() = "STARTUP" Then
            defaultPluginsConfigXml = configMgr.GetConfig("Startup", GetType(IntHostConfig))
            If configMgr.errorMessage <> "" Then
                Throw New SympraxisException.SettingsException(configMgr.errorMessage)
                configMgr.errorMessage = ""
            End If
            loadPluginsConfigxml = configMgr.GetConfig("Startup", GetType(IntHostConfig))
            If configMgr.errorMessage <> "" Then
                Throw New SympraxisException.SettingsException(configMgr.errorMessage)
                configMgr.errorMessage = ""
            End If
        Else

            loadPluginsConfigxml = configMgr.GetConfig(intrHostProcessName, GetType(IntHostConfig))
            If configMgr.errorMessage <> "" Then
                Throw New SympraxisException.SettingsException(configMgr.errorMessage)
                configMgr.errorMessage = ""
            End If
        End If


    End Sub


    Private Sub ValidateConfigandAssignGlobalVariables(ByRef AppHostConfigxml As AppHostConfig)

        'VALIDATING CONFIG AND ASSIGNING VARIABLES



        '' Need write Validate Config
        IOManager.LocalFolder = AppHostConfigxml.LocalFolder

        If String.IsNullOrEmpty(AppHostConfigxml.ProcessType) Then
            Throw New SympraxisException.SettingsException("Invalid implementation on Configuration ProcessType is Empty")


        Else
            processType = AppHostConfigxml.ProcessType
            If processType <> "WORKFLOW" Then
                workingMode = "Normal"
            End If
        End If
        If Not String.IsNullOrEmpty(AppHostConfigxml.MultiWindow) AndAlso AppHostConfigxml.MultiWindow.ToUpper() = "TRUE" Then
            multiWindow = True
        End If
        If String.IsNullOrEmpty(intrHostProcessName) Then
            If String.IsNullOrEmpty(AppHostConfigxml.Process) Then
                appStartBy = "STARTUP"
                intrHostProcessName = "Startup"
                'tempappstart = "STARTUP"
            Else
                appStartBy = AppHostConfigxml.Process
                intrHostProcessName = AppHostConfigxml.Process
            End If
        Else
            appStartBy = intrHostProcessName
            'tempappstart = "STARTUP"
        End If
        If String.IsNullOrEmpty(intrHostWorkFlowName) Then
            If Not String.IsNullOrEmpty(AppHostConfigxml.Workflow) Then
                intrHostWorkFlowName = AppHostConfigxml.Workflow
            End If
        End If


        getworkType = AppHostConfigxml.getWorkType.Split("|")



    End Sub


    Private Function CheckNetworkisUp() As Boolean
        Dim sReturn As Boolean = False
        For Each networkCard As NetworkInterface In NetworkInterface.GetAllNetworkInterfaces

            For Each gatewayAddr As GatewayIPAddressInformation In networkCard.GetIPProperties.GatewayAddresses
                ' if gateway address is NOT 0.0.0.0 and the network card status is UP then we've found the main network card  
                If gatewayAddr.Address.ToString <> "0.0.0.0" And networkCard.OperationalStatus.ToString() = "Up" Then
                    sReturn = True
                End If
            Next
        Next
        Return sReturn
    End Function

    Private Sub LoadConfig(ByRef AppHostConfigxml As AppHostConfig, ByRef loadPluginsConfigxml As IntHostConfig, ByVal initial As Boolean)

        '  LOAD  APPLICATION HOST CONFIG 

        Dim Args() As String = Environment.GetCommandLineArgs()

        Dim sULWFKey As String = ""
        Dim username As String = ""
        Dim workflow As String = ""

        Dim workflowKey As String = ""
        Dim sCmdLineArgds As String = ""

        Dim ProfilePath As String = String.Empty


        Logger.WriteToLogFile("InHost", "LoadConfig Starts")

        If Not IsNothing(Args) AndAlso (Args.Length = 2 Or Args.Length = 3 Or Args.Length = 4 Or Args.Length = 5 Or Args.Length = 6) Then
            If CheckNetworkisUp() = False Then


                If IsNothing(workingMode) AndAlso workingMode = "" Then workingMode = "Offline"
                BtnOffline.Visibility = System.Windows.Visibility.Visible

            End If
            BtnOffline.Visibility = System.Windows.Visibility.Collapsed


            If Args(1).Split(":")(0) = "\in" Then
                cmdprofile = Args(1).Split(":")(1)
            End If
            If Args(1).Contains("\debug") Then
                If Args(1).Contains("u-") Or Args(1).Contains("wf-") Then
                    sULWFKey = Args(1).ToString().ToLower().Split(":"c)(1)
                Else
                    sULWFKey = Args(1)
                End If
                If sULWFKey.Contains(";") Then
                    If sULWFKey.Contains("u-") Then
                        username = sULWFKey.ToString().ToLower().Split(";"c)(0)
                    End If
                    If sULWFKey.Contains("wf-") Then
                        workflow = sULWFKey.ToString().ToLower().Split(";"c)(1)
                    End If
                Else
                    If sULWFKey.Contains("u-") Then
                        username = sULWFKey.ToString().ToLower().Split(";"c)(0)
                    End If
                    If sULWFKey.Contains("wf-") Then
                        workflow = sULWFKey.ToString().ToLower().Split(";"c)(0)
                    End If
                End If

                If workflow <> "" Then
                    workflowKey = workflow.ToString().ToLower().Split("-")(1)
                    If username <> "" Then
                        userloginID = username.ToString().ToLower().Split("-")(1)
                    End If
                    If workflowKey = 0 Then
                        '' sCmdLineArgds = Args(1)
                        Args(1) = "\debug"
                        workingMode = "\debug"
                    ElseIf workflowKey = 1 Then
                        sCmdLineArgds = Args(1)
                        workingMode = "Normal"
                        Args = Args.Skip(1).ToArray()
                    End If
                End If
            End If

            If Args(1).ToLower = "\debug" Then
                If Args.Length = 6 Then
                    If Args(5).ToLower().Contains("pt-") Then
                        IMClientPort = Args(5).ToString().Split("-")(1)
                    Else
                        IMClientPort = Args(5)
                    End If


                    If Args(3).ToLower().Contains("wf-") Then
                        intrHostWorkFlowName = Args(3).ToString().Split("-")(1)
                    Else
                        intrHostWorkFlowName = Args(3)
                    End If

                    If Args(4).ToLower().Contains("pn-") Then
                        intrHostProcessName = Args(4).ToString().Split("-")(1)
                    Else
                        intrHostProcessName = Args(4)
                    End If
                    ProfilePath = Args(2)
                    workingMode = Args(1)

                ElseIf Args.Length = 5 Then
                    If Args(3).ToLower().Contains("pt-") Then
                        IMClientPort = Args(3).ToString().Split("-")(1)

                    ElseIf Args(3).ToLower().Contains("wf-") Then
                        intrHostWorkFlowName = Args(3).ToString().Split("-")(1)

                    ElseIf Args(3).ToLower().Contains("pn-") Then
                        intrHostWorkFlowName = Args(3).ToString().Split("-")(1)
                    Else
                        Throw New SympraxisException.SettingsException("Cannot Identify PortNumber or Workflowname or ProcessName, Please  add valid Prefix (pt-,wf-,pn-)")
                    End If


                    If Args(4).ToLower().Contains("pn-") Then
                        intrHostProcessName = Args(4).ToString().Split("-")(1)
                    ElseIf Args(4).ToLower().Contains("wf-") Then
                        intrHostProcessName = Args(4).ToString().Split("-")(1)
                    ElseIf Args(4).ToLower().Contains("pt-") Then
                        IMClientPort = Args(4).ToString().Split("-")(1)
                    Else
                        Throw New SympraxisException.SettingsException("Cannot Identify PortNumber or Workflowname or ProcessName, Please  add valid Prefix (pt-,wf-,pn-)")
                    End If

                    ProfilePath = Args(2)
                    workingMode = Args(1)
                ElseIf Args.Length = 4 Then
                    If Args(3).ToLower().Contains("wf-") Then
                        intrHostWorkFlowName = Args(3).ToString().Split("-")(1)
                    ElseIf Args(3).ToLower().Contains("pn-") Then
                        intrHostProcessName = Args(3).ToString().Split("-")(1)
                    ElseIf Args(3).ToLower().Contains("pt-") Then
                        IMClientPort = Args(3).ToString().Split("-")(1)
                    Else
                        Throw New SympraxisException.SettingsException("Cannot Identify PortNumber or Workflowname or ProcessName, Please  add valid Prefix (pt-,wf-,pn-)")


                    End If

                    ProfilePath = Args(2)
                    workingMode = Args(1)
                ElseIf Args.Length = 3 Then
                    ProfilePath = Args(2)
                    workingMode = Args(1)
                End If
            ElseIf Args(1).ToLower = "\offline" Then
                ProfilePath = Args(2)
                If IsNothing(workingMode) AndAlso workingMode = "" Then workingMode = "Offline"
            Else
                If Args.Length = 5 Then

                    If Args(4).ToLower().Contains("pt-") Then
                        IMClientPort = Args(4).ToString().Split("-")(1)
                    Else
                        IMClientPort = Args(4)
                    End If
                    If Args(3).ToLower().Contains("pn-") Then
                        intrHostProcessName = Args(3).ToString().Split("-")(1)
                    Else
                        intrHostProcessName = Args(3)
                    End If

                    If Args(2).ToLower().Contains("wf-") Then
                        intrHostWorkFlowName = Args(2).ToString().Split("-")(1)
                    Else
                        intrHostWorkFlowName = Args(2)
                    End If

                    ProfilePath = Args(1)

                    If IsNothing(workingMode) AndAlso workingMode = "" Then workingMode = "Normal"
                ElseIf Args.Length = 4 Then

                    If Args(3).ToLower().Contains("pn-") Then
                        intrHostProcessName = Args(3).ToString().Split("-")(1)
                    ElseIf Args(3).ToLower().Contains("wf-") Then
                        intrHostWorkFlowName = Args(3).ToString().Split("-")(1)
                    ElseIf Args(3).ToLower().Contains("pt-") Then
                        IMClientPort = Args(3).ToString().Split("-")(1)
                    Else
                        Throw New SympraxisException.SettingsException("Cannot Identify PortNumber or Workflowname or ProcessName, Please  add valid Prefix (pt-,wf-,pn-)")

                    End If

                    If Args(2).ToLower().Contains("pn-") Then
                        intrHostProcessName = Args(2).ToString().Split("-")(1)
                    ElseIf Args(2).ToLower().Contains("wf-") Then
                        intrHostWorkFlowName = Args(2).ToString().Split("-")(1)
                    ElseIf Args(2).ToLower().Contains("pt-") Then
                        IMClientPort = Args(2).ToString().Split("-")(1)
                    Else
                        Throw New SympraxisException.SettingsException("Cannot Identify PortNumber or Workflowname or ProcessName, Please  add valid Prefix (pt-,wf-,pn-)")

                    End If

                    ProfilePath = Args(1)

                    If IsNothing(workingMode) AndAlso workingMode = "" Then workingMode = "Normal"
                ElseIf Args.Length = 3 Then
                    If Args(2).ToLower().Contains("wf-") Then
                        intrHostWorkFlowName = Args(2).ToString().Split("-")(1)
                    ElseIf Args(2).ToLower().Contains("pn-") Then
                        intrHostProcessName = Args(2).ToString().Split("-")(1)
                    ElseIf Args(2).ToLower().Contains("pt-") Then
                        IMClientPort = Args(2).ToString().Split("-")(1)

                    Else


                        Throw New SympraxisException.SettingsException("Pls check whether the arguments are available in below order <ConfigurationFile>  Optional[wf-<WorkFlow>] or  Optional[Pn-<Process>]")

                    End If


                    ProfilePath = Args(1)
                    If IsNothing(workingMode) AndAlso workingMode = "" Then workingMode = "Normal"
                Else
                    ProfilePath = Args(1)
                    If IsNothing(workingMode) AndAlso workingMode = "" Then workingMode = "Normal"
                End If
            End If

            If ProfilePath = String.Empty AndAlso Not IOManager.CheckFileExists(ProfilePath) Then
                Throw New SympraxisException.SettingsException(String.Concat("Profile is not exists in ", ProfilePath))


            End If

            If IsNothing(configMgr) Then
                configMgr = New Common.ConfigurationManager(ProfilePath)
                If Not IsNothing(configMgr) AndAlso configMgr.errorMessage <> "" Then
                    Throw New SympraxisException.SettingsException(configMgr.errorMessage)
                    configMgr.errorMessage = ""
                End If
            Else
                configMgr.errorMessage = ""
                configMgr.LoadConfig(ProfilePath)
                If configMgr.errorMessage <> "" Then
                    Throw New SympraxisException.SettingsException(configMgr.errorMessage)
                    configMgr.errorMessage = ""
                End If
            End If
            If Not IsNothing(configMgr) Then

                If initial = True Then
                    AppHostConfigxml = configMgr.GetConfig("ApplicationHost", GetType(AppHostConfig))
                    If configMgr.errorMessage <> "" Then

                        Throw New SympraxisException.SettingsException(configMgr.errorMessage)
                        configMgr.errorMessage = ""
                    End If
                End If

                'Changing Profile path for workflow


                If Not IsNothing(AppHostConfigxml) Then
                    ValidateConfigandAssignGlobalVariables(AppHostConfigxml)


                Else
                    Throw New Exception("AppHostConfigxml is Nothing ")

                End If
                If Not IsNothing(objWCF) AndAlso Not IsNothing(objWCF.workFlowName) AndAlso objWCF.workFlowName <> "" Then
                    intrHostWorkFlowName = objWCF.workFlowName
                End If
                If Not IsNothing(objWCF) AndAlso Not IsNothing(objWCF.processName) AndAlso objWCF.processName <> "" Then
                    intrHostProcessName = objWCF.processName
                End If
                If Not IsNothing(intrHostWorkFlowName) AndAlso intrHostWorkFlowName <> "" Then
                    Dim strarry As String() = ProfilePath.Split("\")
                    ProfilePath = ProfilePath.Replace(strarry(strarry.Length - 1).ToString(), intrHostWorkFlowName.ToUpper() + ".SXML")
                    If ProfilePath = String.Empty AndAlso Not IOManager.CheckFileExists(ProfilePath) Then
                        Throw New SympraxisException.SettingsException(String.Concat("Profile is not exists in ", ProfilePath))

                        Exit Sub
                    End If
                    If IsNothing(configMgr) Then
                        configMgr = New Common.ConfigurationManager(ProfilePath)
                        If configMgr.errorMessage <> "" Then

                            Throw New SympraxisException.SettingsException(configMgr.errorMessage)
                            configMgr.errorMessage = ""
                        End If
                    Else
                        configMgr.errorMessage = ""
                        configMgr.LoadConfig(ProfilePath)
                        If configMgr.errorMessage <> "" Then

                            Throw New SympraxisException.SettingsException(configMgr.errorMessage)
                            configMgr.errorMessage = ""
                        End If
                    End If

                End If
                overrideAppHostConfig = configMgr.GetConfig("ApplicationHost", GetType(OverrideAppHostConfig))
                If configMgr.errorMessage <> "" Then

                    Throw New SympraxisException.SettingsException(configMgr.errorMessage)
                    configMgr.errorMessage = ""
                End If
                WIConfig = configMgr.GetConfig("WorkInstruction", GetType(WIConfig))
                If configMgr.errorMessage <> "" Then

                    Throw New SympraxisException.SettingsException(configMgr.errorMessage)
                    configMgr.errorMessage = ""
                End If
                Logger.WriteToLogFile("InHost", "Load Plugin Config Starts")
                LoadPluginsConfig(loadPluginsConfigxml)


                Logger.WriteToLogFile("InHost", "Load Plugin Config Ends")
                AddWIHotKeys()


                If Not AppHostConfigxml.AppLogger Is Nothing AndAlso AppHostConfigxml.AppLogger.Value.ToString.ToUpper = "TRUE" Then

                    ''IsLogin = True

                    Logger.IslogEnabled = True
                    Logger.AllowedProcess = AppHostConfigxml.AppLogger.AllowedProcess
                    Logger.XcludeProcess = AppHostConfigxml.AppLogger.XcludeProcess
                    Logger.CurrentProcess = intrHostProcessName
                    Logger.LoggerKey = AppHostConfigxml.AppLogger.LoggerKey
                    Logger.xLoggerKey = AppHostConfigxml.AppLogger.xLoggerKey
                    CreateLog()
                End If
                logininterval = -1
                idleinterval = -1
                openworkinterval = -1

                If Not IsNothing(loadPluginsConfigxml) AndAlso Not IsNothing(loadPluginsConfigxml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigxml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigxml.ApplicationHost.Settings.LGTimer) AndAlso loadPluginsConfigxml.ApplicationHost.Settings.LGTimer.ToUpper = "TRUE" Then
                    logininterval = 0
                End If
                If Not IsNothing(loadPluginsConfigxml) AndAlso Not IsNothing(loadPluginsConfigxml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigxml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigxml.ApplicationHost.Settings.IDTimer) AndAlso loadPluginsConfigxml.ApplicationHost.Settings.IDTimer.ToUpper = "TRUE" Then
                    idleinterval = 0
                End If
                If Not IsNothing(loadPluginsConfigxml) AndAlso Not IsNothing(loadPluginsConfigxml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigxml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigxml.ApplicationHost.Settings.WPTimer) AndAlso loadPluginsConfigxml.ApplicationHost.Settings.WPTimer.ToUpper = "TRUE" Then
                    openworkinterval = 0
                End If
            Else
                Throw New SympraxisException.SettingsException("Pls check whether the arguments are available in below order Optional[<Mode>] <ConfigurationFile>  Optional[<WorkFlow>]  Optional[<Process>]")

            End If

        Else
            Throw New SympraxisException.SettingsException("Pls check whether the arguments are available in below order Optional[<Mode>] <ConfigurationFile>  Optional[<WorkFlow>]  Optional[<Process>]")

        End If







    End Sub
    Private Sub CreateLog()

        Dim irc As Int32 = 0
        Dim Lpath As String = String.Empty

        If Not Logger.AllowedProcess Is Nothing AndAlso Logger.AllowedProcess <> "" AndAlso Logger.AllowedProcess.Split("|").Contains(Logger.CurrentProcess) = False Then
            Exit Sub
        End If


        If Logger.Filepath <> "" Then
            Logger.WriteToLogFile("IHost", "LOG_COMPLETED|LogCompleted")
            Logger.CloseFile()
        End If

        Lpath = appHostConfigXml.AppLogger.Path.ToString() & "\" & Now.ToString("MMddyyyy")
        IOManager.CreateDirectory(Lpath)


        Logger.Filepath = Lpath & "\" & intrHostProcessName & "_" & Environment.UserName & "_" & Environment.MachineName & "_" & Now.ToString("MMddyyyyHHmmssFFFtt") & ".Slog"

        If Lpath <> "" Then
            Dim Filepath As String = Lpath & "\" & intrHostProcessName & "_" & Environment.UserName & "_" & Environment.MachineName & "_" & Now.ToString("MMddyyyyHHmmssFFFtt") & ".Slog"
            IOManager.CreateFile2Stream(Filepath)
            Logger.Filepath = Filepath
            Logger.openfile(Filepath)
            Logger.WriteToLogFile("IHost", "LogStarted")
            Logger.WriteToLogFile("IHost", "UserName:" & Environment.UserName)
            Logger.WriteToLogFile("IHost", "MachineName:" & Environment.MachineName)
        End If
    End Sub

    Private Sub CreateLog(ByVal WorkpacketId As String)

        Dim irc As Int32 = 0
        Dim Lpath As String = String.Empty

        If Not Logger.AllowedProcess Is Nothing AndAlso Logger.AllowedProcess <> "" AndAlso Logger.AllowedProcess.Split("|").Contains(Logger.CurrentProcess) = False Then
            Exit Sub
        End If


        If Logger.Filepath <> "" Then
            Logger.WriteToLogFile("IHost", "LOG_COMPLETED|LogCompleted")
            Logger.CloseFile()
        End If

        Lpath = appHostConfigXml.AppLogger.Path.ToString() & "\" & Now.ToString("MMddyyyy")
        IOManager.CreateDirectory(Lpath)


        Logger.Filepath = Lpath & "\" & intrHostProcessName & "_" & Environment.UserName & "_" & Environment.MachineName & "_" & Now.ToString("MMddyyyyHHmmssFFFtt") & ".Slog"

        If Lpath <> "" Then
            Dim Filepath As String = Lpath & "\" & intrHostProcessName & "_" & Environment.UserName & "_" & Environment.MachineName & "_" & WorkpacketId & "_" & Now.ToString("MMddyyyyHHmmssFFFtt") & ".Slog"
            IOManager.CreateFile2Stream(Filepath)
            Logger.Filepath = Filepath
            Logger.openfile(Filepath)
            Logger.WriteToLogFile("IHost", "LogStarted")
            Logger.WriteToLogFile("IHost", "UserName:" & Environment.UserName)
            Logger.WriteToLogFile("IHost", "MachineName:" & Environment.MachineName)
        End If
    End Sub


    Private Sub IdleTimerStart()
        If Not IsNothing(appHostConfigXml.IdleTime) AndAlso Not IsNothing(appHostConfigXml.IdleTime.Interval) AndAlso appHostConfigXml.IdleTime.Interval <> 0 Then

            Dim DblTime As Double = 1000 * appHostConfigXml.IdleTime.Interval
            IdleWaitTimer = New System.Timers.Timer(DblTime)
            AddHandler IdleWaitTimer.Elapsed, AddressOf IdleWaitTimer_Elapsed
            IdleWaitTimer.Enabled = True

        End If
    End Sub


    Private Sub HideInfobarTimerStart(ByVal interval As Integer)

        HideInfobarTimer = New System.Timers.Timer
        HideInfobarTimer.Interval = interval

        AddHandler HideInfobarTimer.Elapsed, AddressOf HideInfobarTimer_Elapsed
        HideInfobarTimer.Enabled = True
    End Sub

    Private Sub HideErrorToolstrip()
        ToolstripError.Visibility = System.Windows.Visibility.Collapsed
    End Sub
    Private Sub AppAutoCloseTimerStart()
        If Not IsNothing(overrideAppHostConfig) AndAlso Not IsNothing(overrideAppHostConfig.AHSettings) AndAlso Not IsNothing(overrideAppHostConfig.AHSettings.AppHostAutocloseInterval) AndAlso overrideAppHostConfig.AHSettings.AppHostAutocloseInterval <> 0 Then

            Dim DblTime As Double = 1000 * overrideAppHostConfig.AHSettings.AppHostAutocloseInterval
            AppAutoCloseTimer = New System.Timers.Timer(DblTime)
            AddHandler AppAutoCloseTimer.Elapsed, AddressOf AppAutoCloseTimer_Elapsed
            AppAutoCloseTimer.Enabled = True

                    AutoCloseAHIdleStarttime = DateTime.Now
        End If
    End Sub


    Private Sub SetIndicator()

        Dim pnlIndWidth As Integer = 92, pnlpoistion As Integer = 100, layoutwidth As Integer = 0
        mLayoutmanager.FindPanelWidth(layoutwidth)
        pnlpoistion = layoutwidth + pnlpoistion
        Dim DIdata As Boolean = False, SIdata As Boolean = False, CIdata As Boolean = False
        CIWidth.Width = New GridLength(30)
        SIWidth.Width = New GridLength(30)
        DIWidth.Width = New GridLength(32)
        If IsNothing(objDInstruction.INS) = False AndAlso objDInstruction.INS.Count > 0 Then
            DIdata = True
        End If
        If IsNothing(objSInstruction.INS) = False AndAlso objSInstruction.INS.Count > 0 Then
            SIdata = True
        End If
        If IsNothing(objCInstruction.INS) = False AndAlso objCInstruction.INS.Count > 0 Then
            CIdata = True
        End If
        If DIdata = True And SIdata = True And CIdata = True Then
            pnlIndWidth = 92
        ElseIf DIdata = True And SIdata = True Then
            pnlIndWidth = 62
            CIWidth.Width = New GridLength(0)
        ElseIf DIdata = True And CIdata = True Then
            pnlIndWidth = 62
            SIWidth.Width = New GridLength(0)
        ElseIf SIdata = True And CIdata = True Then
            pnlIndWidth = 62
            DIWidth.Width = New GridLength(0)
        Else
            If CIdata = True Then
                pnlIndWidth = 32
                SIWidth.Width = New GridLength(0)
                DIWidth.Width = New GridLength(0)
            End If
            If DIdata = True Then
                pnlIndWidth = 32
                CIWidth.Width = New GridLength(0)
                SIWidth.Width = New GridLength(0)
            End If
            If SIdata = True Then
                pnlIndWidth = 32
                CIWidth.Width = New GridLength(0)
                DIWidth.Width = New GridLength(0)
            End If
        End If
        If CIdata = False And DIdata = False And SIdata = False Then
            pnlIndicator.Width = pnlIndWidth
            pnlIndicator.PlacementRectangle = New Rect(Me.ActualWidth - pnlpoistion, Me.ActualHeight - 50, pnlIndWidth, 22)
            pnlIndicator.IsOpen = False
        Else
            pnlIndicator.Width = pnlIndWidth
            pnlIndicator.PlacementRectangle = New Rect(Me.ActualWidth - pnlpoistion, Me.ActualHeight - 50, pnlIndWidth, 22)
            pnlIndicator.IsOpen = True
        End If

    End Sub


    Private Sub mLayoutmanager_KeyDown(sender As Object, e As Input.KeyEventArgs) Handles mLayoutmanager.KeyDown

    End Sub

    Private Sub Window_MouseMove(sender As Object, e As Input.MouseEventArgs)
        Me.Topmost = False
    End Sub

    Private Sub IdleWaitTimer_Elapsed(sender As Object, e As System.Timers.ElapsedEventArgs)
        If IdleWaitTimer.Enabled = True Then
            bIdleTime = True
            IdleWaitTimer.Enabled = False
            If idleinterval <> -1 Then
                CurrIdleInterval = 0
                IdleHourTimer.Interval = 1000
                IdleHourTimer.Enabled = True
            End If
            RaiseEvent IdleTimeStart(DateTime.UtcNow)
        Else
            IdleHourTimer.Enabled = False
            IdleHourTimer.Stop()
        End If
    End Sub
    Private Sub AddWIHotKeys()
        If WIConfig IsNot Nothing AndAlso WIConfig.DynamicInstruction IsNot Nothing AndAlso WIConfig.DynamicInstruction.iSKey <> String.Empty Then
            Dim str As String = String.Empty
            str = splitstring(WIConfig.DynamicInstruction.iSKey)
            Dim KeyConverter As New System.Windows.Forms.KeysConverter
            Dim S As String = KeyConverter.ConvertToString(str)
            Dim O As System.Windows.Forms.Keys = KeyConverter.ConvertFrom(S)
            If KeyboardHook.KeyCodetable.Contains(O) = False Then
                KeyboardHook.KeyCodetable.Add(O, Me)
            End If
        End If
        If WIConfig IsNot Nothing AndAlso WIConfig.StandardInstruction IsNot Nothing AndAlso WIConfig.StandardInstruction.iSKey <> String.Empty Then
            Dim str As String = String.Empty
            str = splitstring(WIConfig.StandardInstruction.iSKey)
            Dim KeyConverter As New System.Windows.Forms.KeysConverter
            Dim S As String = KeyConverter.ConvertToString(str)
            Dim O As System.Windows.Forms.Keys = KeyConverter.ConvertFrom(S)
            If KeyboardHook.KeyCodetable.Contains(O) = False Then
                KeyboardHook.KeyCodetable.Add(O, Me)
            End If
        End If
        If WIConfig IsNot Nothing AndAlso WIConfig.ContextInstruction IsNot Nothing AndAlso WIConfig.ContextInstruction.iSKey <> String.Empty Then
            Dim str As String = String.Empty
            str = splitstring(WIConfig.ContextInstruction.iSKey)
            Dim KeyConverter As New System.Windows.Forms.KeysConverter
            Dim S As String = KeyConverter.ConvertToString(str)
            Dim O As System.Windows.Forms.Keys = KeyConverter.ConvertFrom(S)
            If KeyboardHook.KeyCodetable.Contains(O) = False Then
                KeyboardHook.KeyCodetable.Add(O, Me)
            End If
        End If
    End Sub


    Private Function splitstring(ByVal stringtocheck As String)
        splitstring = Nothing
        If stringtocheck.Trim() <> "" Then
            Dim substrings() As String = stringtocheck.Split(New [Char]() {" "c, "_"c, "%"c, "^"c, CChar(vbTab)})
            If Not IsNothing(substrings) AndAlso substrings.Count > 0 Then
                For Each substring In substrings
                    splitstring = substring
                Next
            End If
        End If

        Return splitstring
    End Function

    Private Sub LoadWCF(ByRef WFClientFrameWork As WCFSettings)
        ' Initiate WorkFlow Client Framework 
        Dim Irc As Integer = 0
        Dim Errmsg As String = String.Empty


        If WFClientFrameWork Is Nothing Then
            Throw New Exception("Missing configuration details for WCF")

        End If

        If Not IsNothing(WFClientFrameWork.Assembly) AndAlso WFClientFrameWork.Assembly.Trim() <> "" Then
            If IOManager.CheckFileExists(WFClientFrameWork.Assembly) = False Then
                Throw New Exception("Unable to locate/access the WCF file provided in configuratation")

            End If
        Else
            Throw New Exception("Unable to locate/access the WCF file provided in configuratation")
        End If



        If Not IsNothing(WFClientFrameWork.ClassName) AndAlso WFClientFrameWork.ClassName <> "" Then
            Try
                objWCF = DirectCast(getAssemblyObject(WFClientFrameWork.Assembly, WFClientFrameWork.ClassName), IWorkflowClient)
            Catch ex As Exception
                Throw ex
            End Try

            If Not IsNothing(objWCF) Then

                Dim AppKey As String = ""

                If intrHostProcessName = "" Then
                    AppKey = "Startup"
                Else
                    AppKey = intrHostProcessName
                End If

                objWCF.processName = intrHostProcessName
                Logger.WriteToLogFile("IHost", "WCF Init Start")
                Irc = objWCF.Init(configMgr, AppKey, Errmsg)

                If Irc <> 0 Then
                    Throw New Exception(Errmsg)
                End If
                Logger.WriteToLogFile("IHost", "WCF Init Completed")
            Else
                Throw New Exception("Unable to load the WCF file provided in configuration")

            End If

        Else

            Throw New Exception("Unable to load the WCF file provided in configuration. Missing class name.")


        End If



    End Sub



    Private Sub LoginUser()
        Dim Irc As Integer = 0
        Dim Errmsg As String = String.Empty
        Dim LoginType As String = String.Empty

        If Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.LoginType) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.LoginType <> "" Then
            LoginType = loadPluginsConfigXml.ApplicationHost.Settings.LoginType
        Else
            Throw New SympraxisException.SettingsException("LoginType settings is not Implemented")
        End If

        If Not IsNothing(userloginID) AndAlso userloginID <> "" Then
            Logger.WriteToLogFile("IHost", userloginID & " --> Start login")
            Irc = objWCF.Login(userloginID, LoginType, Errmsg)

            If Irc <> 0 Then
                Throw New Exception(Errmsg)
            End If

        ElseIf Not IsNothing(overrideAppHostConfig) AndAlso Not IsNothing(overrideAppHostConfig.AHSettings) AndAlso Not IsNothing(overrideAppHostConfig.AHSettings.AddDomain) AndAlso overrideAppHostConfig.AHSettings.AddDomain.ToString().ToUpper() = "TRUE" Then
            userloginID = WindowsIdentity.GetCurrent().Name.ToString().ToLower().Split("\")(1) + "@" + System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName
            Logger.WriteToLogFile("IHost", userloginID & " --> Start login")
            Irc = objWCF.Login(userloginID, LoginType, Errmsg)
            If Irc <> 0 Then
                Throw New Exception(Errmsg)
            End If
        Else
            userloginID = WindowsIdentity.GetCurrent().Name.ToString().ToLower().Split("\")(1)
            Logger.WriteToLogFile("IHost", userloginID & " --> Start login")
            'userloginID = "mp1037675"
            'userloginID = "aa1034183"
            '  userloginID = "krishna"
            ' userloginID = "al2271"
            'userloginID = "appservadmintest"
            Irc = objWCF.Login(userloginID, LoginType, Errmsg)
            If Irc <> 0 Then
                Throw New Exception(Errmsg)
            End If
        End If

        Logger.WriteToLogFile("IHost", userloginID & " --> Sucessfully logined")


    End Sub

    Private Sub EnvironmentVariable()

        If IsNothing(configMgr.AppbaseConfig) Then
            configMgr.AppbaseConfig = New AppbaseConfig
        End If
        If Not IsNothing(cmdprofile) Then
            configMgr.AppbaseConfig.EnvironmentVariables.profilecmd = cmdprofile
        End If
        If IsNothing(configMgr.AppbaseConfig.EnvironmentVariables) Then
            configMgr.AppbaseConfig.EnvironmentVariables = New EnvironmentVariables
        End If

        If Not IsNothing(objWCF.myLoginRsp) Then
            configMgr.AppbaseConfig.EnvironmentVariables.myLoginRsp = objWCF.myLoginRsp
        End If

        configMgr.AppbaseConfig.EnvironmentVariables.CurrentProcessName = intrHostProcessName
        configMgr.AppbaseConfig.EnvironmentVariables.WFName = intrHostWorkFlowName
        configMgr.AppbaseConfig.EnvironmentVariables.UserDomainName = Environment.UserDomainName
        configMgr.AppbaseConfig.EnvironmentVariables.UserName = Environment.UserName

        'configMgr.AppbaseConfig.EnvironmentVariables.sUserId = Environment.UserName

        'OverrideAppHostConfig.WORKPACKETID = WORKPACKETID
        configMgr.AppbaseConfig.EnvironmentVariables.CurrentDate = Now.ToLongDateString
        configMgr.AppbaseConfig.EnvironmentVariables.SystemName = System.Net.Dns.GetHostName()


        configMgr.AppbaseConfig.EnvironmentVariables.WCF = objWCF



        If Not IsNothing(objWCF.ConnectionMode) Then
            configMgr.AppbaseConfig.EnvironmentVariables.ConnectionMode = objWCF.ConnectionMode
        End If

        If Not IsNothing(objWCF.WebServiceConnection) Then
            configMgr.AppbaseConfig.EnvironmentVariables.WebServiceConnection = objWCF.WebServiceConnection
        End If
        If Not IsNothing(overrideAppHostConfig.TransDBURL) Then
            configMgr.AppbaseConfig.TransDBURL = overrideAppHostConfig.TransDBURL
        End If
        If Not IsNothing(overrideAppHostConfig.TransDBWebURL) Then
            configMgr.AppbaseConfig.TransDBWebURL = overrideAppHostConfig.TransDBWebURL
        End If
        If Not IsNothing(overrideAppHostConfig.ARMSURL) Then
            configMgr.AppbaseConfig.ArmsURL = overrideAppHostConfig.ARMSURL
        End If

    End Sub


    Private Sub LoadPlugins(ByRef PluginDefinitions As PluginDefinitions, ByRef ArrangePlugins As List(Of PluginRoute), ByRef LayoutSettings As List(Of LayoutSetting))
        ' configMgr.AppbaseConfig = New AppbaseConfig

        mpanelcontainer = Nothing
        mpanelcontainer = New LayoutSetting.panelcontainer

        Dim irc As Integer = 0
        Dim Errmsg As String = String.Empty

        Dim _configkey As String = ""
        Try

            fullDP = Nothing
            hshTempDp = Nothing
            hshTempDp = New Hashtable

            'Validation 
            If (IsNothing(PluginDefinitions)) Or (Not IsNothing(PluginDefinitions) AndAlso IsNothing(PluginDefinitions.PluginDefinition)) Or (Not IsNothing(PluginDefinitions) AndAlso PluginDefinitions.PluginDefinition.Count = 0) Then

                Throw New SympraxisException.SettingsException("Missing plugin configuration details (AHS).")

                Exit Sub
            End If

            If IsNothing(LayoutSettings) Or (Not IsNothing(LayoutSettings) AndAlso LayoutSettings.Count = 0) Then
                Throw New SympraxisException.SettingsException("Missing layout configuration details for the process ")

            End If


            If IsNothing(ArrangePlugins) Then
                Throw New SympraxisException.SettingsException("Missing plugin configuration details  in " & intrHostProcessName)
            End If
            If ArrangePlugins.Count = 0 Then

                Throw New SympraxisException.SettingsException("Unable to load the plugin configured for the process " & intrHostProcessName)


            End If

            lstPluginOrder = Nothing
            lstPluginOrder = New List(Of Object)

            Dim strLoadPluginName As String = String.Empty
            Dim pluginAssemblypath As String = String.Empty
            Dim isWinForm As String = String.Empty
            Dim PluginObj As Object

            Dim Plugindetails As PluginDefinition

            If HshPluginObj.Count > 0 AndAlso HshGetDefaultsWOTPA.Count > 0 Then
                Dim tempHash As Hashtable = HshPluginObj.Clone()
                For Each a In tempHash
                    If Not HshGetDefaultsWOTPA.Contains(a.Key) Then
                        HshPluginObj.Remove(a.Key)
                    End If
                Next
                tempHash = Nothing
            End If

            For i As Integer = 0 To ArrangePlugins.Count - 1
                PluginDetail = New PluginDetails
                If IsNothing(ArrangePlugins(i).Name) Or (Not IsNothing(ArrangePlugins(i).Name) AndAlso ArrangePlugins(i).Name.Trim() = String.Empty) Then
                    Throw New SympraxisException.SettingsException("Missing plugin configuration details (Plugin name).")
                Else
                    strLoadPluginName = ArrangePlugins(i).Name

                End If
                If Not hshDataPlugins.Contains(ArrangePlugins(i).Name) Then


                    Plugindetails = PluginDefinitions.PluginDefinition.Where(Function(x) x.Name = strLoadPluginName).FirstOrDefault()
                    If IsNothing(Plugindetails) Then
                        Throw New SympraxisException.SettingsException(String.Concat("Unable to find the plugin defined in the defintion for process.[", intrHostProcessName, "] [", strLoadPluginName, "]"))
                    End If

                    isWinForm = Plugindetails.IsWinControl

                    If IsNothing(Plugindetails.PathFolder) Or Plugindetails.PathFolder.Trim() = "" Then

                        If Not IsNothing(PluginDefinitions.PluginFolder) Or PluginDefinitions.PluginFolder.Trim() <> "" Then
                            pluginAssemblypath = PluginDefinitions.PluginFolder
                        Else
                            Throw New SympraxisException.SettingsException(String.Concat("Missing plugin folder path in configuration . [", strLoadPluginName, "]"))

                        End If

                    Else
                        pluginAssemblypath = Plugindetails.PathFolder
                    End If

                    Pathvalidation(pluginAssemblypath)

                    If IOManager.CheckFileExists(String.Concat(pluginAssemblypath, Plugindetails.Assembly)) Then


                        If IsNothing(Plugindetails.ClassName) Or Plugindetails.ClassName = "" Then
                            Throw New SympraxisException.SettingsException(String.Concat("Unable to locate/access the plugin file. Missing class Name", strLoadPluginName))

                        End If


                        PluginObj = getAssemblyObject(String.Concat(pluginAssemblypath, Plugindetails.Assembly), Plugindetails.ClassName)

                    Else

                        Throw New SympraxisException.SettingsException(String.Concat("Unable to locate/access the plugin file. . [", strLoadPluginName, "]"))

                    End If

                Else
                    PluginObj = hshDataPlugins(strLoadPluginName)
                    Plugindetails = PluginDefinitions.PluginDefinition.Where(Function(x) x.Name = strLoadPluginName).FirstOrDefault()
                    If IsNothing(Plugindetails) Then
                        Throw New SympraxisException.SettingsException(String.Concat("Unable to find the plugin defined in the defintion for process.[", intrHostProcessName, "] [", strLoadPluginName, "]"))
                    End If

                    isWinForm = Plugindetails.IsWinControl
                End If




                'Dim PluginObj = DirectCast(TempObj, INormalPlugin)
                'Dim PluginObj = DirectCast(TempObj, INormalPlugin)

                If ArrangePlugins(i).ConfigKey <> "" Then
                    _configkey = intrHostProcessName + "/" + ArrangePlugins(i).ConfigKey
                Else
                    _configkey = intrHostProcessName
                End If

                If Not hshDataPlugins.Contains(strLoadPluginName) Then
                    hshDataPlugins.Add(strLoadPluginName, PluginObj)
                End If

                If Not HshGetDefaultsWOTPA.ContainsKey(ArrangePlugins(i).Title) Then
                    Logger.WriteToLogFile("IHost", "Init Plugin-->" & strLoadPluginName & " Starts")
                    DirectCast(PluginObj, INormalPlugin).Init(configMgr, _configkey, Me, irc, Errmsg)
                    If irc <> 0 Then
                        Throw New SympraxisException.SettingsException(Errmsg)
                    End If

                    'If hshDataPartsName.ContainsKey(strLoadPluginName) = False Then

                    '    If Not IsNothing(DirectCast(PluginObj, INormalPlugin).DataPart) AndAlso DirectCast(PluginObj, INormalPlugin).DataPart.ToString() <> "" Then
                    '        hshDataPartsName.Add(strLoadPluginName, DirectCast(PluginObj, INormalPlugin).DataPart)
                    '    End If
                    'Else
                    '    DirectCast(PluginObj, INormalPlugin).DataPart = ""
                    '    If IsNothing(DirectCast(PluginObj, INormalPlugin).DataPart) Or (Not IsNothing(DirectCast(PluginObj, INormalPlugin).DataPart) AndAlso DirectCast(PluginObj, INormalPlugin).DataPart.ToString() = "") Then
                    '        DirectCast(PluginObj, INormalPlugin).DataPart = hshDataPartsName(strLoadPluginName)
                    '    End If
                    'End If
                End If

                Logger.WriteToLogFile("IHost", "Initialized Plugin -->" & strLoadPluginName & " Ends")

                'If Not FullDP.ContainsKey(PluginObj) Then
                '    FullDP.Add(PluginObj, Nothing)
                '    HshtempDp.Add(PluginObj, Nothing)
                'End If

                If appStartBy.ToUpper() = "STARTUP" Then
                    If Not HshGetDefaultsWOTPA.ContainsKey(ArrangePlugins(i).Title) Then
                        HshGetDefaultsWOTPA.Add(ArrangePlugins(i).Title, i)
                    End If
                    If lstStartupPluginOrder.Contains(PluginObj) = False Then
                        lstStartupPluginOrder.Add(PluginObj)
                    End If

                End If

                If Not hshGetPluginTitle.ContainsKey(ArrangePlugins(i).Title) And Not HshGetDefaultsWOTPA.Contains(ArrangePlugins(i).Title) Then
                    hshGetPluginTitle.Add(ArrangePlugins(i).Title, Nothing)
                End If

                lstPluginOrder.Add(PluginObj)

                PluginDetail.sName = ArrangePlugins(i).Name
                PluginDetail.sKey = ArrangePlugins(i).SKey
                PluginDetail.height = ArrangePlugins(i).Height
                PluginDetail.Width = ArrangePlugins(i).Width
                PluginDetail.IsWinForm = isWinForm
                PluginDetail.PluginObj = PluginObj
                If Not HshPluginObj.ContainsKey(ArrangePlugins(i).Title) Then
                    HshPluginObj.Add(ArrangePlugins(i).Title, PluginDetail)

                End If

                'If appStartBy.ToUpper() = "STARTUP" Then
                '    lstPluginOrder.Add(PluginObj)

                '    PluginDetail.sName = ArrangePlugins(i).Name
                '    PluginDetail.sKey = ArrangePlugins(i).SKey
                '    PluginDetail.height = ArrangePlugins(i).Height
                '    PluginDetail.Width = ArrangePlugins(i).Width
                '    PluginDetail.IsWinForm = isWinForm
                '    PluginDetail.PluginObj = PluginObj
                '    If Not HshPluginObj.ContainsKey(ArrangePlugins(i).Title) Then
                '        HshPluginObj.Add(ArrangePlugins(i).Title, PluginDetail)

                '    End If

                'ElseIf Not HshGetDefaultsWOTPA.Contains(ArrangePlugins(i).Title) Then

                '    lstPluginOrder.Add(PluginObj)


                '    PluginDetail.sName = ArrangePlugins(i).Name
                '    PluginDetail.sKey = ArrangePlugins(i).SKey
                '    PluginDetail.height = ArrangePlugins(i).Height
                '    PluginDetail.Width = ArrangePlugins(i).Width
                '    PluginDetail.IsWinForm = isWinForm
                '    PluginDetail.PluginObj = PluginObj
                '    If Not HshPluginObj.ContainsKey(ArrangePlugins(i).Title) Then
                '        HshPluginObj.Add(ArrangePlugins(i).Title, PluginDetail)
                '    End If
                'End If



                'If isWinForm <> String.Empty AndAlso isWinForm.ToUpper = "FALSE" Then
                '    mLayoutmanager.AddWPFcontrol(PluginObj, mpanelcontainer, ArrangePlugins(i).Panel, ArrangePlugins(i).Title, ArrangePlugins(i).SKey, ArrangePlugins(i).IconSource, ArrangePlugins(i).Width, ArrangePlugins(i).Height)
                'Else

                '    WindowsFormsHost = New Integration.WindowsFormsHost
                '    WindowsFormsHost.Child = PluginObj
                '    WindowsFormsHost.Name = ArrangePlugins(i).Name

                '    WindowsFormsHost.Uid = ArrangePlugins(i).Name
                '    mLayoutmanager.AddWindowscontrol(WindowsFormsHost, mpanelcontainer, ArrangePlugins(i).Panel, ArrangePlugins(i).Title, ArrangePlugins(i).SKey, ArrangePlugins(i).IconSource, ArrangePlugins(i).Width, ArrangePlugins(i).Height)
                '    WindowsFormsHost.BringIntoView()
                'End If


                'ElementHost.BringToFront(PluginObj)


            Next


            Logger.WriteToLogFile("IHost", "Add Controls To Panel Starts")
            AddControlsToPanel(0, mpanelcontainer, WindowsFormsHost, LayoutSettings)
            Logger.WriteToLogFile("IHost", "Add Controls To Panel Ends")

            strLoadPluginName = String.Empty
            pluginAssemblypath = String.Empty
            'Timesheet show in process panel
            mLayoutmanager.Tflag = False
            mLayoutmanager.clear()
            mLayoutmanager.CollapseModeFalse()

            mLayoutmanager.GetPlugin = New Hashtable()


            If appStartBy.ToUpper() <> "STARTUP" Then
                If workingMode = "\debug" Then
                    mLayoutmanager.Getdefaults = New Hashtable()
                End If
                For i As Integer = 0 To hshGetPluginTitle.Count - 1
                    mLayoutmanager.GetPlugin.Add(i, hshGetPluginTitle.Keys(i))
                Next

                mLayoutmanager._CurrLayouts = LayoutSettings(0).Layouts(0).DisplayText
                mLayoutmanager.Init(LayoutSettings(0), mpanelcontainer)
                'DelegateApplylayout(0)
                ''mLayoutmanager.ApplyLayout(0)
                'mLayoutmanager.Tflag = True

            Else
                Logger.WriteToLogFile("IHost", "Layout panel settings 4 WD Starts")
                mLayoutmanager.Getdefaults = New Hashtable()

                For i As Integer = 0 To HshGetDefaultsWOTPA.Count - 1
                    mLayoutmanager.Getdefaults.Add(i, HshGetDefaultsWOTPA.Keys(i))
                Next


                Dim _Layouts As String = ""

                '   WCF.GetWorkList(Irc, ErrMsg)
                For i = 0 To LayoutSettings(0).Layouts.Count - 1
                    Dim Display As String
                    If LayoutSettings(0).Layouts(i).DisplayText = "" Then
                        Dim send As String
                        send = LayoutSettings(0).Layouts(i).DisplayImage

                        Dim Dimg() As String
                        Dimg = send.Split("\")
                        Dim DLimg As String
                        DLimg = Dimg(Dimg.Length - 1).ToString()
                        DLimg = DLimg.Substring(0, DLimg.LastIndexOf("."))
                        Display = DLimg

                        send = String.Empty
                        Dimg = Nothing

                    Else
                        Display = LayoutSettings(0).Layouts(i).DisplayText
                    End If
                    _Layouts = String.Concat(_Layouts, "|", Display)
                    If Not hshEnableDefault.Contains(Display) Then
                        hshEnableDefault.Add(Display, False)
                    End If
                Next

                'GetdefaultSettings()

                'changeDefaults()

                Dim index As Integer = 0

                mLayoutmanager.Init(defaultPluginsConfigXml.ApplicationHost.LayoutSettings(0), mpanelcontainer)
                mLayoutmanager._DefaultLayouts = _Layouts
                mLayoutmanager.Tflag = False
                mLayoutmanager.clear()
                mLayoutmanager.CollapseModeFalse()
                mLayoutmanager.IsEscapePressed = True
                mLayoutmanager.UpdateLayout()
                Dim DefaultLayoutSettingLayout = New LayoutSetting.Layout


                DefaultLayoutSettingLayout = defaultPluginsConfigXml.ApplicationHost.LayoutSettings(0).Layouts(index)

                If hshEnableDefault.ContainsKey(returnimagename(defaultPluginsConfigXml.ApplicationHost.LayoutSettings(0).Layouts(index).DisplayImage)) Then
                    hshEnableDefault.Remove(returnimagename(defaultPluginsConfigXml.ApplicationHost.LayoutSettings(0).Layouts(index).DisplayImage))
                    hshEnableDefault.Add(returnimagename(defaultPluginsConfigXml.ApplicationHost.LayoutSettings(0).Layouts(index).DisplayImage), True)
                End If
                mLayoutmanager.EnableDefaults = hshEnableDefault


                mLayoutmanager.ApplyLayout(0)
                'mLayoutmanager.Tflag = True

                Logger.WriteToLogFile("IHost", "Layout panel settings 4 WD Ends")
            End If



            ' mLayoutmanager.ApplyLayout(0)

            mLayoutmanager.UpdateLayout()
            mLayoutmanager.BringIntoView()
            mLayoutmanager.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch
            mLayoutmanager.VerticalAlignment = System.Windows.VerticalAlignment.Stretch


            MasterLayout.Children.Clear()


            MasterLayout.Children.Add(mLayoutmanager)







        Finally
            mpanelcontainer = Nothing
            _configkey = String.Empty
        End Try
    End Sub

    Private Sub AddControlsToPanel(index As Integer, mpanelcontainer As LayoutSetting.panelcontainer, WindowsFormsHost As System.Windows.Forms.Integration.WindowsFormsHost, LayoutSettings As List(Of LayoutSetting))
        Dim PluginTitle As String = ""
        Dim Panel As String = ""
        If IsNothing(LayoutSettings(0).Layouts(index).ToolsSetting.Value) = False AndAlso LayoutSettings(0).Layouts(index).ToolsSetting.Value.ToString <> String.Empty Then
            PluginTitle = LayoutSettings(0).Layouts(index).ToolsSetting.Value.ToString
            If LayoutSettings(0).Layouts(index).ToolsSetting.RightSize > 0 Then
                Panel = "TPB"
            Else
                Panel = "TPA"
            End If

            If PluginTitle.Contains("|") Then
                For Each Plugin In PluginTitle.Split("|")
                    If HshPluginObj.ContainsKey(Plugin) Then
                        If HshPluginObj(Plugin).IsWinForm <> String.Empty AndAlso HshPluginObj(Plugin).IsWinForm.ToUpper = "FALSE" Then
                            mLayoutmanager.AddWPFcontrol(HshPluginObj(Plugin).PluginObj, mpanelcontainer, Panel, Plugin, HshPluginObj(Plugin).sKey, HshPluginObj(Plugin).Width, HshPluginObj(Plugin).height)
                        Else
                            WindowsFormsHost = New Integration.WindowsFormsHost
                            WindowsFormsHost.Child = HshPluginObj(Plugin).PluginObj
                            WindowsFormsHost.Name = HshPluginObj(Plugin).sName
                            WindowsFormsHost.Uid = HshPluginObj(Plugin).sName
                            mLayoutmanager.AddWindowscontrol(WindowsFormsHost, mpanelcontainer, Panel, HshPluginObj(Plugin).sName, HshPluginObj(Plugin).sKey, HshPluginObj(Plugin).Width, HshPluginObj(Plugin).height)
                            WindowsFormsHost.BringIntoView()
                        End If
                    End If
                Next
            Else
                If HshPluginObj.ContainsKey(PluginTitle) Then
                    If HshPluginObj(PluginTitle).IsWinForm <> String.Empty AndAlso HshPluginObj(PluginTitle).IsWinForm.ToUpper = "FALSE" Then
                        mLayoutmanager.AddWPFcontrol(HshPluginObj(PluginTitle).PluginObj, mpanelcontainer, Panel, PluginTitle, HshPluginObj(PluginTitle).sKey, HshPluginObj(PluginTitle).Width, HshPluginObj(PluginTitle).height)
                    Else
                        WindowsFormsHost = New Integration.WindowsFormsHost
                        WindowsFormsHost.Child = HshPluginObj(PluginTitle).PluginObj
                        WindowsFormsHost.Name = HshPluginObj(PluginTitle).sName
                        WindowsFormsHost.Uid = HshPluginObj(PluginTitle).sName
                        mLayoutmanager.AddWindowscontrol(WindowsFormsHost, mpanelcontainer, Panel, HshPluginObj(PluginTitle).sName, HshPluginObj(PluginTitle).sKey, HshPluginObj(PluginTitle).Width, HshPluginObj(PluginTitle).height)
                        WindowsFormsHost.BringIntoView()
                    End If
                End If
            End If
        End If
        If IsNothing(LayoutSettings(0).Layouts(index).TopPanelSetting.TopPanelSettingA.Value) = False AndAlso LayoutSettings(0).Layouts(index).TopPanelSetting.TopPanelSettingA.Value.ToString <> String.Empty Then
            PluginTitle = LayoutSettings(0).Layouts(index).TopPanelSetting.TopPanelSettingA.Value.ToString
            Panel = "TWPA"
            If PluginTitle.Contains("|") Then
                For Each Plugin In PluginTitle.Split("|")
                    If HshPluginObj.ContainsKey(Plugin) Then
                        If HshPluginObj(Plugin).IsWinForm <> String.Empty AndAlso HshPluginObj(Plugin).IsWinForm.ToUpper = "FALSE" Then
                            mLayoutmanager.AddWPFcontrol(HshPluginObj(Plugin).PluginObj, mpanelcontainer, Panel, Plugin, HshPluginObj(Plugin).sKey, HshPluginObj(Plugin).Width, HshPluginObj(Plugin).height)
                        Else
                            WindowsFormsHost = New Integration.WindowsFormsHost
                            WindowsFormsHost.Child = HshPluginObj(Plugin).PluginObj
                            WindowsFormsHost.Name = HshPluginObj(Plugin).sName
                            WindowsFormsHost.Uid = HshPluginObj(Plugin).sName
                            mLayoutmanager.AddWindowscontrol(WindowsFormsHost, mpanelcontainer, Panel, HshPluginObj(Plugin).sName, HshPluginObj(Plugin).sKey, HshPluginObj(Plugin).Width, HshPluginObj(Plugin).height)
                            WindowsFormsHost.BringIntoView()
                        End If
                    End If
                Next
            Else
                If HshPluginObj.ContainsKey(PluginTitle) Then
                    If HshPluginObj(PluginTitle).IsWinForm <> String.Empty AndAlso HshPluginObj(PluginTitle).IsWinForm.ToUpper = "FALSE" Then
                        mLayoutmanager.AddWPFcontrol(HshPluginObj(PluginTitle).PluginObj, mpanelcontainer, Panel, PluginTitle, HshPluginObj(PluginTitle).sKey, HshPluginObj(PluginTitle).Width, HshPluginObj(PluginTitle).height)
                    Else
                        WindowsFormsHost = New Integration.WindowsFormsHost
                        WindowsFormsHost.Child = HshPluginObj(PluginTitle).PluginObj
                        WindowsFormsHost.Name = PluginTitle.ToString
                        WindowsFormsHost.Uid = PluginTitle
                        mLayoutmanager.AddWindowscontrol(WindowsFormsHost, mpanelcontainer, Panel, HshPluginObj(PluginTitle).sName, HshPluginObj(PluginTitle).sKey, HshPluginObj(PluginTitle).Width, HshPluginObj(PluginTitle).height)
                        WindowsFormsHost.BringIntoView()
                    End If
                End If
            End If
        End If
        If IsNothing(LayoutSettings(0).Layouts(index).TopPanelSetting.TopPanelSettingB.Value) = False AndAlso LayoutSettings(0).Layouts(index).TopPanelSetting.TopPanelSettingB.Value.ToString <> String.Empty Then
            PluginTitle = LayoutSettings(0).Layouts(index).TopPanelSetting.TopPanelSettingB.Value.ToString
            Panel = "TWPB"
            If PluginTitle.Contains("|") Then
                For Each Plugin In PluginTitle.Split("|")
                    If HshPluginObj.ContainsKey(Plugin) Then
                        If HshPluginObj(Plugin).IsWinForm <> String.Empty AndAlso HshPluginObj(Plugin).IsWinForm.ToUpper = "FALSE" Then
                            mLayoutmanager.AddWPFcontrol(HshPluginObj(Plugin).PluginObj, mpanelcontainer, Panel, Plugin, HshPluginObj(Plugin).sKey, HshPluginObj(Plugin).Width, HshPluginObj(Plugin).height)
                        Else
                            WindowsFormsHost = New Integration.WindowsFormsHost
                            WindowsFormsHost.Child = HshPluginObj(Plugin).PluginObj
                            WindowsFormsHost.Name = HshPluginObj(Plugin).sName
                            WindowsFormsHost.Uid = HshPluginObj(Plugin).sName
                            mLayoutmanager.AddWindowscontrol(WindowsFormsHost, mpanelcontainer, Panel, HshPluginObj(Plugin).sName, HshPluginObj(Plugin).sKey, HshPluginObj(Plugin).Width, HshPluginObj(Plugin).height)
                            WindowsFormsHost.BringIntoView()
                        End If
                    End If
                Next
            Else
                If HshPluginObj.ContainsKey(PluginTitle) Then
                    If HshPluginObj(PluginTitle).IsWinForm <> String.Empty AndAlso HshPluginObj(PluginTitle).IsWinForm.ToUpper = "FALSE" Then
                        mLayoutmanager.AddWPFcontrol(HshPluginObj(PluginTitle).PluginObj, mpanelcontainer, Panel, PluginTitle, HshPluginObj(PluginTitle).sKey, HshPluginObj(PluginTitle).Width, HshPluginObj(PluginTitle).height)
                    Else
                        WindowsFormsHost = New Integration.WindowsFormsHost
                        WindowsFormsHost.Child = HshPluginObj(PluginTitle).PluginObj
                        WindowsFormsHost.Name = HshPluginObj(PluginTitle).sName
                        WindowsFormsHost.Uid = HshPluginObj(PluginTitle).sName
                        mLayoutmanager.AddWindowscontrol(WindowsFormsHost, mpanelcontainer, Panel, HshPluginObj(PluginTitle).sName, HshPluginObj(PluginTitle).sKey, HshPluginObj(PluginTitle).Width, HshPluginObj(PluginTitle).height)
                        WindowsFormsHost.BringIntoView()
                    End If
                End If
            End If
        End If
        If IsNothing(LayoutSettings(0).Layouts(index).ContentPanelSetting.WorkPanel.WorkPanelSettingA.Value) = False AndAlso LayoutSettings(0).Layouts(index).ContentPanelSetting.WorkPanel.WorkPanelSettingA.Value.ToString <> String.Empty Then
            PluginTitle = LayoutSettings(0).Layouts(index).ContentPanelSetting.WorkPanel.WorkPanelSettingA.Value.ToString
            Panel = "WPA"
            If PluginTitle.Contains("|") Then
                For Each Plugin In PluginTitle.Split("|")
                    If HshPluginObj.ContainsKey(Plugin) Then
                        If HshPluginObj(Plugin).IsWinForm <> String.Empty AndAlso HshPluginObj(Plugin).IsWinForm.ToUpper = "FALSE" Then
                            mLayoutmanager.AddWPFcontrol(HshPluginObj(Plugin).PluginObj, mpanelcontainer, Panel, Plugin, HshPluginObj(Plugin).sKey, HshPluginObj(Plugin).Width, HshPluginObj(Plugin).height)
                        Else
                            WindowsFormsHost = New Integration.WindowsFormsHost
                            WindowsFormsHost.Child = HshPluginObj(Plugin).PluginObj
                            WindowsFormsHost.Name = HshPluginObj(Plugin).sName
                            WindowsFormsHost.Uid = HshPluginObj(Plugin).sName
                            mLayoutmanager.AddWindowscontrol(WindowsFormsHost, mpanelcontainer, Panel, HshPluginObj(Plugin).sName, HshPluginObj(Plugin).sKey, HshPluginObj(Plugin).Width, HshPluginObj(Plugin).height)
                            WindowsFormsHost.BringIntoView()
                        End If
                    End If
                Next
            Else
                If HshPluginObj.ContainsKey(PluginTitle) Then
                    If HshPluginObj(PluginTitle).IsWinForm <> String.Empty AndAlso HshPluginObj(PluginTitle).IsWinForm.ToUpper = "FALSE" Then
                        mLayoutmanager.AddWPFcontrol(HshPluginObj(PluginTitle).PluginObj, mpanelcontainer, Panel, PluginTitle, HshPluginObj(PluginTitle).sKey, HshPluginObj(PluginTitle).Width, HshPluginObj(PluginTitle).height)
                    Else
                        WindowsFormsHost = New Integration.WindowsFormsHost
                        WindowsFormsHost.Child = HshPluginObj(PluginTitle).PluginObj
                        WindowsFormsHost.Name = HshPluginObj(PluginTitle).sName
                        WindowsFormsHost.Uid = HshPluginObj(PluginTitle).sName
                        mLayoutmanager.AddWindowscontrol(WindowsFormsHost, mpanelcontainer, Panel, HshPluginObj(PluginTitle).sName, HshPluginObj(PluginTitle).sKey, HshPluginObj(PluginTitle).Width, HshPluginObj(PluginTitle).height)
                        WindowsFormsHost.BringIntoView()
                    End If
                End If
            End If
        End If
        If IsNothing(LayoutSettings(0).Layouts(index).ContentPanelSetting.WorkPanel.WorkPanelSettingB.Value) = False AndAlso LayoutSettings(0).Layouts(index).ContentPanelSetting.WorkPanel.WorkPanelSettingB.Value.ToString <> String.Empty Then
            PluginTitle = LayoutSettings(0).Layouts(index).ContentPanelSetting.WorkPanel.WorkPanelSettingB.Value.ToString
            Panel = "WPB"
            If PluginTitle.Contains("|") Then
                For Each Plugin In PluginTitle.Split("|")
                    If HshPluginObj.ContainsKey(Plugin) Then
                        If HshPluginObj(Plugin).IsWinForm <> String.Empty AndAlso HshPluginObj(Plugin).IsWinForm.ToUpper = "FALSE" Then
                            mLayoutmanager.AddWPFcontrol(HshPluginObj(Plugin).PluginObj, mpanelcontainer, Panel, Plugin, HshPluginObj(Plugin).sKey, HshPluginObj(Plugin).Width, HshPluginObj(Plugin).height)
                        Else
                            WindowsFormsHost = New Integration.WindowsFormsHost
                            WindowsFormsHost.Child = HshPluginObj(Plugin).PluginObj
                            WindowsFormsHost.Name = HshPluginObj(Plugin).sName
                            WindowsFormsHost.Uid = HshPluginObj(Plugin).sName
                            mLayoutmanager.AddWindowscontrol(WindowsFormsHost, mpanelcontainer, Panel, HshPluginObj(Plugin).sName, HshPluginObj(Plugin).sKey, HshPluginObj(Plugin).Width, HshPluginObj(Plugin).height)
                            WindowsFormsHost.BringIntoView()
                        End If
                    End If
                Next
            Else
                If HshPluginObj.ContainsKey(PluginTitle) Then
                    If HshPluginObj(PluginTitle).IsWinForm <> String.Empty AndAlso HshPluginObj(PluginTitle).IsWinForm.ToUpper = "FALSE" Then
                        mLayoutmanager.AddWPFcontrol(HshPluginObj(PluginTitle).PluginObj, mpanelcontainer, Panel, PluginTitle, HshPluginObj(PluginTitle).sKey, HshPluginObj(PluginTitle).Width, HshPluginObj(PluginTitle).height)
                    Else
                        WindowsFormsHost = New Integration.WindowsFormsHost
                        WindowsFormsHost.Child = HshPluginObj(PluginTitle).PluginObj
                        WindowsFormsHost.Name = HshPluginObj(PluginTitle).sName
                        WindowsFormsHost.Uid = HshPluginObj(PluginTitle).sName
                        mLayoutmanager.AddWindowscontrol(WindowsFormsHost, mpanelcontainer, Panel, HshPluginObj(PluginTitle).sName, HshPluginObj(PluginTitle).sKey, HshPluginObj(PluginTitle).Width, HshPluginObj(PluginTitle).height)
                        WindowsFormsHost.BringIntoView()
                    End If
                End If
            End If
        End If
        If IsNothing(LayoutSettings(0).Layouts(index).ContentPanelSetting.ReferencePanel.RefPanelSettingA.Value) = False AndAlso LayoutSettings(0).Layouts(index).ContentPanelSetting.ReferencePanel.RefPanelSettingA.Value.ToString <> String.Empty Then
            PluginTitle = LayoutSettings(0).Layouts(index).ContentPanelSetting.ReferencePanel.RefPanelSettingA.Value.ToString
            Panel = "RPA"
            If PluginTitle.Contains("|") Then
                For Each Plugin In PluginTitle.Split("|")
                    If HshPluginObj.ContainsKey(Plugin) Then
                        If HshPluginObj(Plugin).IsWinForm <> String.Empty AndAlso HshPluginObj(Plugin).IsWinForm.ToUpper = "FALSE" Then
                            mLayoutmanager.AddWPFcontrol(HshPluginObj(Plugin).PluginObj, mpanelcontainer, Panel, Plugin, HshPluginObj(Plugin).sKey, HshPluginObj(Plugin).Width, HshPluginObj(Plugin).height)
                        Else
                            WindowsFormsHost = New Integration.WindowsFormsHost
                            WindowsFormsHost.Child = HshPluginObj(Plugin).PluginObj
                            WindowsFormsHost.Name = HshPluginObj(Plugin).sName
                            WindowsFormsHost.Uid = HshPluginObj(Plugin).sName
                            mLayoutmanager.AddWindowscontrol(WindowsFormsHost, mpanelcontainer, Panel, HshPluginObj(Plugin).sName, HshPluginObj(Plugin).sKey, HshPluginObj(Plugin).Width, HshPluginObj(Plugin).height)
                            WindowsFormsHost.BringIntoView()
                        End If
                    End If
                Next
            Else
                If HshPluginObj.ContainsKey(PluginTitle) Then
                    If HshPluginObj(PluginTitle).IsWinForm <> String.Empty AndAlso HshPluginObj(PluginTitle).IsWinForm.ToUpper = "FALSE" Then
                        mLayoutmanager.AddWPFcontrol(HshPluginObj(PluginTitle).PluginObj, mpanelcontainer, Panel, PluginTitle, HshPluginObj(PluginTitle).sKey, HshPluginObj(PluginTitle).Width, HshPluginObj(PluginTitle).height)
                    Else
                        WindowsFormsHost = New Integration.WindowsFormsHost
                        WindowsFormsHost.Child = HshPluginObj(PluginTitle).PluginObj
                        WindowsFormsHost.Name = HshPluginObj(PluginTitle).sName
                        WindowsFormsHost.Uid = HshPluginObj(PluginTitle).sName
                        mLayoutmanager.AddWindowscontrol(WindowsFormsHost, mpanelcontainer, Panel, HshPluginObj(PluginTitle).sName, HshPluginObj(PluginTitle).sKey, HshPluginObj(PluginTitle).Width, HshPluginObj(PluginTitle).height)
                        WindowsFormsHost.BringIntoView()
                    End If
                End If
            End If
        End If
        If IsNothing(LayoutSettings(0).Layouts(index).ContentPanelSetting.ReferencePanel.RefPanelSettingB.Value) = False AndAlso LayoutSettings(0).Layouts(index).ContentPanelSetting.ReferencePanel.RefPanelSettingB.Value.ToString <> String.Empty Then
            PluginTitle = LayoutSettings(0).Layouts(index).ContentPanelSetting.ReferencePanel.RefPanelSettingB.Value.ToString
            Panel = "RPB"
            If PluginTitle.Contains("|") Then
                For Each Plugin In PluginTitle.Split("|")
                    If HshPluginObj.ContainsKey(Plugin) Then
                        If HshPluginObj(Plugin).IsWinForm <> String.Empty AndAlso HshPluginObj(Plugin).IsWinForm.ToUpper = "FALSE" Then
                            mLayoutmanager.AddWPFcontrol(HshPluginObj(Plugin).PluginObj, mpanelcontainer, Panel, PluginTitle, HshPluginObj(Plugin).sKey, HshPluginObj(Plugin).Width, HshPluginObj(Plugin).height)
                        Else
                            WindowsFormsHost = New Integration.WindowsFormsHost
                            WindowsFormsHost.Child = HshPluginObj(Plugin).PluginObj
                            WindowsFormsHost.Name = HshPluginObj(Plugin).sName
                            WindowsFormsHost.Uid = HshPluginObj(Plugin).sName
                            mLayoutmanager.AddWindowscontrol(WindowsFormsHost, mpanelcontainer, Panel, HshPluginObj(Plugin).sName, HshPluginObj(Plugin).sKey, HshPluginObj(Plugin).Width, HshPluginObj(Plugin).height)
                            WindowsFormsHost.BringIntoView()
                        End If
                    End If
                Next
            Else
                If HshPluginObj.ContainsKey(PluginTitle) Then
                    If HshPluginObj(PluginTitle).IsWinForm <> String.Empty AndAlso HshPluginObj(PluginTitle).IsWinForm.ToUpper = "FALSE" Then
                        mLayoutmanager.AddWPFcontrol(HshPluginObj(PluginTitle).PluginObj, mpanelcontainer, Panel, PluginTitle, HshPluginObj(PluginTitle).sKey, HshPluginObj(PluginTitle).Width, HshPluginObj(PluginTitle).height)
                    Else
                        WindowsFormsHost = New Integration.WindowsFormsHost
                        WindowsFormsHost.Child = HshPluginObj(PluginTitle).PluginObj
                        WindowsFormsHost.Name = HshPluginObj(PluginTitle).sName
                        WindowsFormsHost.Uid = HshPluginObj(PluginTitle).sName
                        mLayoutmanager.AddWindowscontrol(WindowsFormsHost, mpanelcontainer, Panel, HshPluginObj(PluginTitle).sName, HshPluginObj(PluginTitle).sKey, HshPluginObj(PluginTitle).Width, HshPluginObj(PluginTitle).height)
                        WindowsFormsHost.BringIntoView()
                    End If
                End If
            End If
        End If
        If IsNothing(LayoutSettings(0).Layouts(index).BottomPanelSetting.BottomPanelSettingA.Value) = False AndAlso LayoutSettings(0).Layouts(index).BottomPanelSetting.BottomPanelSettingA.Value.ToString <> String.Empty Then
            PluginTitle = LayoutSettings(0).Layouts(index).BottomPanelSetting.BottomPanelSettingA.Value.ToString
            Panel = "BWPA"
            If PluginTitle.Contains("|") Then
                For Each Plugin In PluginTitle.Split("|")
                    If HshPluginObj.ContainsKey(Plugin) Then
                        If HshPluginObj(Plugin).IsWinForm <> String.Empty AndAlso HshPluginObj(Plugin).IsWinForm.ToUpper = "FALSE" Then
                            mLayoutmanager.AddWPFcontrol(HshPluginObj(Plugin).PluginObj, mpanelcontainer, Panel, Plugin, HshPluginObj(Plugin).sKey, HshPluginObj(Plugin).Width, HshPluginObj(Plugin).height)
                        Else
                            WindowsFormsHost = New Integration.WindowsFormsHost
                            WindowsFormsHost.Child = HshPluginObj(Plugin).PluginObj
                            WindowsFormsHost.Name = HshPluginObj(Plugin).sName
                            WindowsFormsHost.Uid = HshPluginObj(Plugin).sName
                            mLayoutmanager.AddWindowscontrol(WindowsFormsHost, mpanelcontainer, Panel, HshPluginObj(Plugin).sName, HshPluginObj(Plugin).sKey, HshPluginObj(Plugin).Width, HshPluginObj(Plugin).height)
                            WindowsFormsHost.BringIntoView()
                        End If
                    End If
                Next
            Else
                If HshPluginObj.ContainsKey(PluginTitle) Then
                    If HshPluginObj(PluginTitle).IsWinForm <> String.Empty AndAlso HshPluginObj(PluginTitle).IsWinForm.ToUpper = "FALSE" Then
                        mLayoutmanager.AddWPFcontrol(HshPluginObj(PluginTitle).PluginObj, mpanelcontainer, Panel, PluginTitle, HshPluginObj(PluginTitle).sKey, HshPluginObj(PluginTitle).Width, HshPluginObj(PluginTitle).height)
                    Else
                        WindowsFormsHost = New Integration.WindowsFormsHost
                        WindowsFormsHost.Child = HshPluginObj(PluginTitle).PluginObj
                        WindowsFormsHost.Name = HshPluginObj(PluginTitle).sName
                        WindowsFormsHost.Uid = HshPluginObj(PluginTitle).sName
                        mLayoutmanager.AddWindowscontrol(WindowsFormsHost, mpanelcontainer, Panel, HshPluginObj(PluginTitle).sName, HshPluginObj(PluginTitle).sKey, HshPluginObj(PluginTitle).Width, HshPluginObj(PluginTitle).height)
                        WindowsFormsHost.BringIntoView()
                    End If
                End If
            End If
        End If
        If IsNothing(LayoutSettings(0).Layouts(index).BottomPanelSetting.BottomPanelSettingB.Value) = False AndAlso LayoutSettings(0).Layouts(index).BottomPanelSetting.BottomPanelSettingB.Value.ToString <> String.Empty Then
            PluginTitle = LayoutSettings(0).Layouts(index).BottomPanelSetting.BottomPanelSettingB.Value.ToString
            Panel = "BWPB"
            If PluginTitle.Contains("|") Then
                For Each Plugin In PluginTitle.Split("|")
                    If HshPluginObj.ContainsKey(Plugin) Then
                        If HshPluginObj(PluginTitle).IsWinForm <> String.Empty AndAlso HshPluginObj(PluginTitle).IsWinForm.ToUpper = "FALSE" Then
                            mLayoutmanager.AddWPFcontrol(HshPluginObj(Plugin), mpanelcontainer, Panel, Plugin, HshPluginObj(Plugin).sKey, HshPluginObj(Plugin).Width, HshPluginObj(Plugin).height)
                        Else
                            WindowsFormsHost = New Integration.WindowsFormsHost
                            WindowsFormsHost.Child = HshPluginObj(Plugin).PluginObj
                            WindowsFormsHost.Name = HshPluginObj(Plugin).sName
                            WindowsFormsHost.Uid = HshPluginObj(Plugin).sName
                            mLayoutmanager.AddWindowscontrol(WindowsFormsHost, mpanelcontainer, Panel, HshPluginObj(Plugin).sName, HshPluginObj(Plugin).sKey, HshPluginObj(Plugin).Width, HshPluginObj(Plugin).height)
                            WindowsFormsHost.BringIntoView()
                        End If
                    End If
                Next
            Else
                If HshPluginObj.ContainsKey(PluginTitle) Then
                    If HshPluginObj(PluginTitle).IsWinForm <> String.Empty AndAlso HshPluginObj(PluginTitle).IsWinForm.ToUpper = "FALSE" Then
                        mLayoutmanager.AddWPFcontrol(HshPluginObj(PluginTitle).PluginObj, mpanelcontainer, Panel, PluginTitle, HshPluginObj(PluginTitle).sKey, HshPluginObj(PluginTitle).Width, HshPluginObj(PluginTitle).height)
                    Else
                        WindowsFormsHost = New Integration.WindowsFormsHost
                        WindowsFormsHost.Child = HshPluginObj(PluginTitle).PluginObj
                        WindowsFormsHost.Name = HshPluginObj(PluginTitle).sName
                        WindowsFormsHost.Uid = HshPluginObj(PluginTitle).sName
                        mLayoutmanager.AddWindowscontrol(WindowsFormsHost, mpanelcontainer, Panel, HshPluginObj(PluginTitle).sName, HshPluginObj(PluginTitle).sKey, HshPluginObj(PluginTitle).Width, HshPluginObj(PluginTitle).height)
                        WindowsFormsHost.BringIntoView()
                    End If
                End If
            End If
        End If
    End Sub





    Private Function returnimagename(ByVal url As String)
        Dim send As String
        send = url
        Dim Dimg() As String
        Dimg = send.Split("\")
        Dim DLimg As String
        DLimg = Dimg(Dimg.Length - 1).ToString()
        DLimg = DLimg.Substring(0, DLimg.LastIndexOf("."))
        Return DLimg
    End Function

    Private Function getAssemblyObject(ByVal Assemblyname As String, ByVal Classname As String) As Object
        Dim IType As Type
        Dim tempobj As Object
        Dim DynAssmbly As System.Reflection.Assembly = System.Reflection.Assembly.UnsafeLoadFrom(Assemblyname)
        Dim DynAsmTypes() As Type = DynAssmbly.GetTypes()
        IType = DynAssmbly.GetType(Classname)
        tempobj = New Object
        tempobj = Activator.CreateInstance(IType)
        Return tempobj
    End Function



    Private Sub AttachProcess(ByVal ProcessFlag As Boolean)
        Dim irc As Integer = 0
        Dim Errmsg As String = String.Empty

        irc = objWCF.AttachProcess(intrHostProcessName, intrHostWorkFlowName, Errmsg, ProcessFlag)
        If irc <> 0 Then
            Throw New SympraxisException.SettingsException("Unable to attach the porcess to the User [" & intrHostProcessName & "] [" & userloginID & "] ")
        End If

    End Sub

    Private Sub RegisterWF(ByVal ProcessFlag As Boolean)
        Dim irc As Integer = 0
        Dim Errmsg As String = String.Empty

        irc = objWCF.RegisterWorkFlow(intrHostProcessName, intrHostWorkFlowName, Errmsg, ProcessFlag)
        If irc <> 0 Then
            Throw New SympraxisException.SettingsException("Unable to register the workflow to the User [" & intrHostProcessName & "] [" & userloginID & "] [" & intrHostWorkFlowName & "] ")
        End If
    End Sub

    Private Sub SetApplicationTitle()



        If Not IsNothing(loadPluginsConfigXml) Then
            If Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.ApplicationTitle) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.ApplicationTitle <> String.Empty Then
                If Not IsNothing(configMgr.AppbaseConfig.EnvironmentVariables.sFilterCriteria) AndAlso configMgr.AppbaseConfig.EnvironmentVariables.sFilterCriteria <> String.Empty Then
                    If delegatorUser <> "" Then
                        Me.Title = String.Concat(delegatorUser, " : ", loadPluginsConfigXml.ApplicationHost.Settings.ApplicationTitle.ToString, " (", configMgr.AppbaseConfig.EnvironmentVariables.sFilterCriteria, ") ")

                    Else
                        Me.Title = String.Concat(loadPluginsConfigXml.ApplicationHost.Settings.ApplicationTitle.ToString, " (", configMgr.AppbaseConfig.EnvironmentVariables.sFilterCriteria, ") ")

                    End If
                Else
                    If delegatorUser <> "" Then
                        Me.Title = String.Concat(delegatorUser, " : ", loadPluginsConfigXml.ApplicationHost.Settings.ApplicationTitle.ToString)

                    Else
                        Me.Title = loadPluginsConfigXml.ApplicationHost.Settings.ApplicationTitle.ToString

                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Menuhandle()

        If configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True Then
            If Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.ParallelWorkItem) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.ParallelWorkItem.ToUpper = "TRUE" Then
                AHEnableMenu("Home", "Close|Exit")

            ElseIf objWCF.WorkitemCount > 1 Then
                If objWCF.CurrentWorkitemIdx = 1 Then

                    AHEnableMenu("Home", "Next|Close|Exit")

                ElseIf objWCF.CurrentWorkitemIdx = objWCF.WorkitemCount Then

                    AHEnableMenu("Home", "Previous|Close|Exit")

                Else
                    AHEnableMenu("Home", "Previous|Next|Close|Exit")

                End If
            Else
                AHEnableMenu("Home", "Close|Exit")

            End If
        Else
            If workingMode = "\debug" Or appStartBy.ToUpper() <> "STARTUP" Then
                AHEnableMenu("Home", "Open|Exit")
            Else
                AHEnableMenu("Home", "Exit")
            End If

        End If


    End Sub


    'Private Function FindVisualChildren1(Of T As DependencyObject)(depObj As DependencyObject) As List(Of T)
    '    Dim list As New List(Of T)()
    '    If depObj IsNot Nothing Then
    '        For i As Integer = 0 To VisualTreeHelper.GetChildrenCount(depObj) - 1
    '            Dim child As DependencyObject = VisualTreeHelper.GetChild(depObj, i)
    '            If child IsNot Nothing AndAlso TypeOf child Is T Then
    '                list.Add(DirectCast(child, T))
    '            End If

    '            Dim childItems As List(Of T) = FindVisualChildren1(Of T)(child)
    '            If childItems IsNot Nothing AndAlso childItems.Count() > 0 Then
    '                For Each item As Object In childItems
    '                    list.Add(item)
    '                Next
    '            End If
    '        Next
    '    End If
    '    Return list
    'End Function


    Private Sub GetPluginObj(ByVal PluginName As String, ByRef PluginObj As Object)
        Dim Plugindetails As PluginDefinition
        Dim pluginAssemblypath As String = String.Empty
        Dim irc As Integer = 0
        Dim Errmsg As String = String.Empty

        Plugindetails = appHostConfigXml.PluginDefinitions.PluginDefinition.Where(Function(x) x.Name = PluginName).FirstOrDefault()
        If Not IsNothing(Plugindetails) Then

            If IsNothing(Plugindetails.PathFolder) Or Plugindetails.PathFolder.Trim() = "" Then
                pluginAssemblypath = appHostConfigXml.PluginDefinitions.PluginFolder
            ElseIf Not IsNothing(Plugindetails.PathFolder) AndAlso Plugindetails.PathFolder.Trim() <> "" Then
                pluginAssemblypath = Plugindetails.PathFolder

            End If


            Pathvalidation(pluginAssemblypath)

            If IOManager.CheckFileExists(String.Concat(pluginAssemblypath, Plugindetails.Assembly)) Then

                PluginObj = getAssemblyObject(String.Concat(pluginAssemblypath, Plugindetails.Assembly), Plugindetails.ClassName)


            Else
                Throw New SympraxisException.SettingsException(String.Concat("Invalid Implementation, Cannot found plugin Path ", pluginAssemblypath, Plugindetails.Assembly))

            End If


        Else

            Throw New SympraxisException.SettingsException(String.Concat("Invalid Implementation, Cannot find plugin name ", PluginName, " in pluginDefinition"))
        End If

    End Sub


    Private Sub ExitParser(ByRef ObjSaveStatus As String, ByVal ChdObjectid As String)

        Dim PluginObj As Object = Nothing
        Dim DPValue As Dictionary(Of String, String) = Nothing
        Dim strDpValues As String
        Dim strDpvalue As String()
        Dim Irc As Integer
        Dim ErrMsg As String = String.Empty
        Try

            If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Parser) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Parser.ExitParser) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Parser.ExitParser.Name) AndAlso loadPluginsConfigXml.ApplicationHost.Parser.ExitParser.Name <> "" Then
                Logger.WriteToLogFile("IHost", "ExitParser  Starts")

                Dim ExitparserPlugin As String() = loadPluginsConfigXml.ApplicationHost.Parser.ExitParser.Name.Split("|")
                If Not IsNothing(ExitparserPlugin) AndAlso ExitparserPlugin.Count > 0 Then
                    For i As Integer = 0 To ExitparserPlugin.Count - 1
                        If Not IsNothing(hshDataPlugins) AndAlso hshDataPlugins.ContainsKey(ExitparserPlugin(i)) Then
                            PluginObj = hshDataPlugins(ExitparserPlugin(i))
                        Else
                            GetPluginObj(ExitparserPlugin(i), PluginObj)
                        End If






                        If Not IsNothing(PluginObj) Then
                            Dim _configkey As String
                            If Not IsNothing(loadPluginsConfigXml.ApplicationHost.Parser.ExitParser.ConfigKey) AndAlso loadPluginsConfigXml.ApplicationHost.Parser.ExitParser.ConfigKey <> "" Then
                                _configkey = intrHostProcessName + "/" + loadPluginsConfigXml.ApplicationHost.Parser.ExitParser.ConfigKey
                            Else
                                _configkey = intrHostProcessName
                            End If


                            DirectCast(PluginObj, IExitParser).Init(configMgr, _configkey, Me, Irc, ErrMsg)

                            If Irc <> 0 Then
                                Throw New SympraxisException.SettingsException(ErrMsg)
                            End If

                            ' WCF.GetParserDatapart(PluginObj, DPValue, irc, Errmsg)
                            DPValue = New Dictionary(Of String, String)
                            strDpValues = Convert.ToString(DirectCast(PluginObj, IExitParser).DataPart)
                            If Not IsNothing(strDpValues) Then
                                strDpvalue = strDpValues.Split("|")
                                For K = 0 To strDpvalue.Length - 1
                                    DPValue.Add(strDpvalue(K), Nothing)


                                Next

                            End If
                            hshDataParts = New Hashtable(DPValue)
                            Logger.WriteToLogFile("IHost", "ExitParser  Read DP Starts")
                            If ChdObjectid = "" Then
                            Irc = objWCF.Read(DPValue, ErrMsg) ' - vl1056306

                            Else
                                objWCF.GetChildDataPartByChdObjectid(DPValue, ChdObjectid, Irc, ErrMsg) ' - vl1056306

                            End If
                            ' irc = WCF.Read(hshDataParts, Errmsg)
                            If Irc <> 0 Then
                                Throw New SympraxisException.WorkitemException(ErrMsg)
                            End If
                            Logger.WriteToLogFile("IHost", "ExitParser Read DP Ends")

                            If Irc = 0 AndAlso Not IsNothing(DPValue) AndAlso DPValue.Count > 0 Then
                                Logger.WriteToLogFile("IHost", "ExitParser  LoadWorkitem Starts")
                                DirectCast(PluginObj, IExitParser).ExitParserLoadWorkItem(DPValue, ObjSaveStatus, Irc, ErrMsg)
                                If Irc <> 0 Then
                                    Throw New SympraxisException.WorkitemException(ErrMsg)
                                End If
                                Logger.WriteToLogFile("IHost", "ExitParser  LoadWorkitem Ends")
                            End If

                        End If

                    Next

                End If
                Logger.WriteToLogFile("IHost", "ExitParser  Ends")
            End If


        Finally
            PluginObj = Nothing
            DPValue = Nothing
            strDpValues = Nothing
            strDpvalue = Nothing
        End Try
    End Sub

    Private Sub CheckInputParser()

        Dim DicInputParser As Dictionary(Of String, Object)
        Dim PluginObj As Object = Nothing
        Dim DPValue As Dictionary(Of String, String)

        Dim strDpValues As String
        Dim strDpvalue As String()

        Dim liststr As New List(Of Integer)

        Logger.WriteToLogFile("IHost", "InputParser starts")
        Dim iRC As Integer = 0
        Dim ErrMsg As String = String.Empty
        Try


            DicInputParser = New Dictionary(Of String, Object)
            If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Parser) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Parser.InputParser) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Parser.InputParser.Name) AndAlso loadPluginsConfigXml.ApplicationHost.Parser.InputParser.Name <> "" Then

                Dim InputparserPlugin As String() = loadPluginsConfigXml.ApplicationHost.Parser.InputParser.Name.Split("|")
                If Not IsNothing(InputparserPlugin) AndAlso InputparserPlugin.Count > 0 Then


                    For i As Integer = 0 To InputparserPlugin.Count - 1
                        Logger.WriteToLogFile("IHost", "InputParser get Plugin Object starts")
                        If Not IsNothing(hshDataPlugins) AndAlso hshDataPlugins.ContainsKey(InputparserPlugin(i)) Then
                            PluginObj = hshDataPlugins(InputparserPlugin(i))
                        Else
                            GetPluginObj(InputparserPlugin(i), PluginObj)
                        End If




                        Logger.WriteToLogFile("IHost", "InputParser get Plugin Object Ends")


                        If Not IsNothing(PluginObj) Then
                            Dim _configkey As String
                            If Not IsNothing(loadPluginsConfigXml.ApplicationHost.Parser.InputParser.ConfigKey) AndAlso loadPluginsConfigXml.ApplicationHost.Parser.InputParser.ConfigKey <> "" Then
                                _configkey = intrHostProcessName + "/" + loadPluginsConfigXml.ApplicationHost.Parser.InputParser.ConfigKey
                            Else
                                _configkey = intrHostProcessName
                            End If

                            Logger.WriteToLogFile("IHost", "InputParser Init Starts")
                            DirectCast(PluginObj, IInputParser).Init(configMgr, _configkey, Me, iRC, ErrMsg)

                            If iRC <> 0 Then

                                Throw New SympraxisException.SettingsException(ErrMsg)

                            End If
                            Logger.WriteToLogFile("IHost", "InputParser Init Ends")

                            DicInputParser.Add(InputparserPlugin(i), PluginObj)
                        End If

                    Next



                    If iRC = 0 AndAlso DicInputParser.Count > 0 Then

                        For Each Inparser In DicInputParser

                            For i As Integer = 1 To objWCF.WorkitemCount
                                If i <> 1 Then
                                    objWCF.SwitchWorkitem(i, iRC, ErrMsg)

                                    '' Fail 
                                    If iRC <> 0 Then
                                        Continue For
                                    End If

                                End If

                                DPValue = New Dictionary(Of String, String)
                                ' WCF.GetParserDatapart(Inparser.Value, DPValue, irc, Errmsg)
                                strDpValues = Convert.ToString(DirectCast(Inparser.Value, IInputParser).DataPart)
                                If Not IsNothing(strDpValues) Then
                                    strDpvalue = strDpValues.Split("|")
                                    For K = 0 To strDpvalue.Length - 1
                                        DPValue.Add(strDpvalue(K), Nothing)
                                    Next
                                End If
                                hshDataParts = New Hashtable(DPValue)
                                Logger.WriteToLogFile("IHost", "InputParser Read Datapart Starts")
                                iRC = objWCF.Read(DPValue, ErrMsg)
                                'irc = WCF.Read(hshDataParts, Errmsg)
                                If iRC <> 0 Then
                                    objWCF.UpdateSaveWFStatus("Error", iRC, ErrMsg)
                                    If iRC <> 0 Then
                                        Throw New Exception()
                                    End If

                                    Continue For
                                End If
                                Logger.WriteToLogFile("IHost", "InputParser Read Datapart Starts")

                                If iRC = 0 AndAlso Not IsNothing(DPValue) AndAlso DPValue.Count > 0 Then
                                    DirectCast(Inparser.Value, IInputParser).InputParserLoadWorkItem(DPValue, iRC, ErrMsg)
                                    If iRC <> 0 Then

                                        objWCF.UpdateSaveWFStatus("Error", iRC, ErrMsg)
                                        If iRC <> 0 Then
                                            Throw New Exception()
                                        End If

                                        '   AHWorkitemCount = AHWorkitemCount - 1
                                        If liststr.Contains(i) Then
                                            liststr.Remove(i)
                                        End If
                                    Else
                                        If Not liststr.Contains(i) Then
                                            liststr.Add(i)
                                        End If
                                    End If

                                    iRC = 0
                                End If

                                DPValue = Nothing
                            Next
                        Next

                        ''19/12/19 To Do HAndle SwitchWorkitem exception
                        If liststr.Count > 0 Then
                            liststr.Sort()
                            Dim IsSwitchWorkitemapply As Boolean = False
                            For x As Integer = 0 To liststr.Count - 1
                                objWCF.SwitchWorkitem(liststr(x), iRC, ErrMsg)
                                If iRC = 0 Then
                                    IsSwitchWorkitemapply = True
                                    Exit For

                                End If

                            Next

                            If IsSwitchWorkitemapply = False Then
                                Throw New SympraxisException.WorkitemException(ErrMsg)
                            End If

                        Else
                            Throw New SympraxisException.WorkitemException("No Workitem to load,All Workitem moved to error by Inputparser")



                        End If


                        Logger.WriteToLogFile("IHost", "InputParser Ends")

                    End If
                End If
            End If


        Finally
            DicInputParser = Nothing

            PluginObj = Nothing
            DPValue = Nothing
            liststr = Nothing
            strDpvalue = Nothing
        End Try


    End Sub

    Private Sub SetWorkitemEnvironmentvariable()
        configMgr.AppbaseConfig.EnvironmentVariables.lObjectId = objWCF.Objectid
        configMgr.AppbaseConfig.EnvironmentVariables.sObjectName = objWCF.Objectname
        configMgr.AppbaseConfig.EnvironmentVariables.WFName = objWCF.workFlowName
        configMgr.AppbaseConfig.EnvironmentVariables.WORKPACKETID = objWCF.Workpacketid
    End Sub

    Private Sub BuildEmptyDP4IndexnRev(ByRef strDpvalue As String, ByRef FullDPValue As Dictionary(Of String, String))

        If strDpvalue.Substring(0, 4) = "@CIC" Then
            If strDpvalue.Substring(0, 4) = "@CIC" Then
                If strDpvalue.Contains(":") Then
                    Dim strDefaultCI As String() = strDpvalue.Split(":")
                    If strDefaultCI.Count > 0 AndAlso Not IsNothing(strDefaultCI(1)) AndAlso strDefaultCI(1).ToString <> String.Empty Then
                        If Not FullDPValue.ContainsKey("@CI:" + strDefaultCI(1).ToString) Then
                            FullDPValue.Add("@CI:" + strDefaultCI(1).ToString, Nothing)
                        End If
                    End If
                Else
                    If Not FullDPValue.ContainsKey("CI") Then
                        FullDPValue.Add("@CI", Nothing)
                    End If
                End If
            End If

            If Not FullDPValue.ContainsKey("IM") Then
                FullDPValue.Add("@IM", Nothing)
            End If
            If Not FullDPValue.ContainsKey("ID") Then
                FullDPValue.Add("@ID", Nothing)
            End If

        ElseIf strDpvalue.Substring(0, 5) = "@CRQN" Then


            If strDpvalue.Contains(":") Then
                Dim StrCR As String() = strDpvalue.Split(":")
                If StrCR.Count > 0 AndAlso Not IsNothing(StrCR(1)) AndAlso StrCR(1).ToString <> String.Empty Then
                    If Not FullDPValue.ContainsKey("@CR:" + StrCR(1).ToString) Then
                        FullDPValue.Add("@CR:" + StrCR(1).ToString, Nothing)

                    End If
                End If
            Else
                If Not FullDPValue.ContainsKey("CR") Then
                    FullDPValue.Add("@CR", Nothing)
                End If

            End If

            If Not FullDPValue.ContainsKey("RM") Then
                FullDPValue.Add("@RM", Nothing)
            End If
            If Not FullDPValue.ContainsKey("RD") Then
                FullDPValue.Add("@RD", Nothing)
            End If
        End If
    End Sub


    Private Sub ReadPluginDatapart(ByRef FullDPValue As Dictionary(Of String, String), ByRef pluginobj As Object)
        Dim strDpvalue As String()
        Dim strDpValues As String
        Dim Irc As Integer = 0
        Dim Errmsg As String = String.Empty
        Try

            'Creating  FullDatapart collection Dictionary(Of String, String)

            FullDPValue = New Dictionary(Of String, String)

            strDpValues = Convert.ToString(DirectCast(pluginobj, Sympraxis.Utilities.INormalPlugin.INormalPlugin).DataPart)
            If IsNothing(strDpValues) Or (Not IsNothing(strDpValues) AndAlso strDpValues.Length = 0) Then
                Exit Sub
            End If

            strDpvalue = strDpValues.Split("|")

            For i = 0 To strDpvalue.Length - 1
                If strDpvalue(i).ToString <> String.Empty AndAlso strDpvalue(i).Substring(0, 1) = "@" Then
                    Logger.WriteToLogFile("IHost", "BuildEmpty DP IndexnRev Start")
                    BuildEmptyDP4IndexnRev(strDpvalue(i), FullDPValue)
                    Logger.WriteToLogFile("IHost", "BuildEmpty DP IndexnRev Ends")
                Else
                    If strDpvalue(i).ToString <> String.Empty AndAlso Not FullDPValue.ContainsKey(strDpvalue(i)) Then
                        FullDPValue.Add(strDpvalue(i), Nothing)
                    End If
                End If

            Next



            Irc = objWCF.Read(FullDPValue, Errmsg)
            If Irc <> 0 Then
                Throw New SympraxisException.WorkitemException(Errmsg)
        End If

            ' irc = WCF.Read(hshDataParts, ErrMsg)


        Finally
            strDpvalue = Nothing
            strDpValues = Nothing
        End Try
    End Sub

    Private Sub ReadDatapart(ByRef FullDPValue As Dictionary(Of String, String))
        Dim strDpvalue As String()
        Dim strDpValues As String
        Dim Irc As Integer = 0
        Dim Errmsg As String = String.Empty
        Try

            'Creating  FullDatapart collection Dictionary(Of String, String)

            FullDPValue = New Dictionary(Of String, String)
            For index As Integer = 0 To lstPluginOrder.Count - 1
                strDpValues = Convert.ToString(DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.INormalPlugin).DataPart)
                If IsNothing(strDpValues) Or (Not IsNothing(strDpValues) AndAlso strDpValues.Length = 0) Then Continue For
                strDpvalue = strDpValues.Split("|")

                For i = 0 To strDpvalue.Length - 1
                    If strDpvalue(i).ToString <> String.Empty AndAlso strDpvalue(i).Substring(0, 1) = "@" Then
                        Logger.WriteToLogFile("IHost", "BuildEmpty DP IndexnRev Start")
                        BuildEmptyDP4IndexnRev(strDpvalue(i), FullDPValue)
                        Logger.WriteToLogFile("IHost", "BuildEmpty DP IndexnRev Ends")
                    Else
                        If strDpvalue(i).ToString <> String.Empty AndAlso Not FullDPValue.ContainsKey(strDpvalue(i)) Then
                            FullDPValue.Add(strDpvalue(i), Nothing)
                        End If
                    End If

                Next
            Next


            Irc = objWCF.Read(FullDPValue, Errmsg)
            If Irc <> 0 Then
                Throw New SympraxisException.WorkitemException(Errmsg)
            End If

            ' irc = WCF.Read(hshDataParts, ErrMsg)


        Finally
            strDpvalue = Nothing
            strDpValues = Nothing
        End Try
    End Sub

    Private Sub LoadChildworkitem()
        ChildFullDPValue = Nothing
        ChildFullDPValue = New Dictionary(Of String, Dictionary(Of String, String))
        Dim CurrntChildDPValue = New Dictionary(Of String, Dictionary(Of String, String))

        Dim ParallelChildDp = New Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, String)))

        ReadChildDatapart(ChildFullDPValue)


        Dim strDpValues As String = String.Empty


        Dim Dp As Dictionary(Of String, String)
        Dim strDpvalue As String()
        If ChildFullDPValue.Count > 0 Then
            For index As Integer = 0 To lstPluginOrder.Count - 1
                If Not IsNothing(TryCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                    strDpValues = Convert.ToString(DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).ChildDataPart)
                    If Not IsNothing(strDpValues) AndAlso strDpValues <> String.Empty Then

                        For Each element In ChildFullDPValue
                            Dp = New Dictionary(Of String, String)

                            strDpvalue = strDpValues.Split("|")

                            For i = 0 To strDpvalue.Length - 1
                                If strDpvalue(i).Substring(0, 1) = "@" Then

                                    BuildDP4IndexnRev(strDpvalue(i), Dp, element.Value)
                                Else

                                    Dp.Add(strDpvalue(i), element.Value(strDpvalue(i)))

                                End If

                            Next
                            CurrntChildDPValue.Add(element.Key, Dp)
                            Dp = Nothing
                        Next


                    ElseIf (IsNothing(strDpValues) Or (Not IsNothing(strDpValues) AndAlso strDpValues = String.Empty)) Then
                        Continue For
                    End If
                    If CurrntChildDPValue.Count > 0 Then
                        ParallelChildDp.Add(objWCF.Objectid, CurrntChildDPValue)
                        CurrntChildDPValue = Nothing
                        DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).LoadWorkItem(Nothing, ParallelChildDp)
                    End If



                End If
            Next
        End If







    End Sub


    Private Sub CloseChildworkitem()
        If Not IsNothing(ChildFullDPValue) AndAlso ChildFullDPValue.Count > 0 Then
            Dim CurrntChildDPValue = New Dictionary(Of String, Dictionary(Of String, String))

            Dim ParallelChildDp = New Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, String)))
            Dim strDpValues As String = String.Empty
            Dim Irc As Integer = 0
            Dim Errmsg As String = String.Empty

            Dim Dp As Dictionary(Of String, String)
            Dim strDpvalue As String()

            For index As Integer = 0 To lstPluginOrder.Count - 1
                If Not IsNothing(TryCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                    strDpValues = Convert.ToString(DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).ChildDataPart)
                    If Not IsNothing(strDpValues) AndAlso strDpValues <> String.Empty Then

                        For Each element In ChildFullDPValue
                            Dp = New Dictionary(Of String, String)

                            strDpvalue = strDpValues.Split("|")

                            For i = 0 To strDpvalue.Length - 1
                                If strDpvalue(i).Substring(0, 1) = "@" Then

                                    BuildDP4IndexnRev(strDpvalue(i), Dp, element.Value)
                                Else

                                    Dp.Add(strDpvalue(i), element.Value(strDpvalue(i)))

                                End If

                            Next
                            CurrntChildDPValue.Add(element.Key, Dp)
                            Dp = Nothing
                        Next


                    ElseIf (IsNothing(strDpValues) Or (Not IsNothing(strDpValues) AndAlso strDpValues = String.Empty)) Then
                        Continue For
                    End If
                    If CurrntChildDPValue.Count > 0 Then
                        ParallelChildDp.Add(objWCF.Objectid, CurrntChildDPValue)
                        CurrntChildDPValue = Nothing
                        Dim IsAborted As Boolean = False
                        DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).Closing(Nothing, ParallelChildDp, IsAborted)
                        If IsAborted Then

                            Throw New SympraxisException.EmptyException()
                        End If
                        IsAborted = Nothing
                        CurrntChildDPValue = ParallelChildDp.ElementAt(0).Value
                        If Not IsNothing(CurrntChildDPValue) Then
                            Dim subdp As Dictionary(Of String, String)
                            Dim finaldp As Dictionary(Of String, String)
                            For Each eDp In CurrntChildDPValue
                                subdp = New Dictionary(Of String, String)(eDp.Value)

                                If Not IsNothing(subdp) Then
                                    finaldp = ChildFullDPValue(eDp.Key)
                                    For Each subdpeDp In subdp
                                        If strDpValues.Substring(0, 1) = "@" Then
                                            If finaldp.ContainsKey("@" + subdpeDp.Key) Then
                                                finaldp("@" + subdpeDp.Key) = subdpeDp.Value

                                            Else
                                                finaldp(subdpeDp.Key) = subdpeDp.Value

                                            End If
                                        Else
                                            finaldp(subdpeDp.Key) = subdpeDp.Value
                                        End If
                                    Next
                                    ChildFullDPValue(eDp.Key) = finaldp

                                End If

                                subdp = Nothing
                            Next

                            finaldp = Nothing
                        End If



                    End If



                End If
            Next

            If Not IsNothing(ChildFullDPValue) AndAlso ChildFullDPValue.Count > 0 Then
                objWCF.WriteChildDataPart(ChildFullDPValue, Irc, Errmsg)

                If Irc <> 0 Then
                    Throw New SympraxisException.WorkitemException(Errmsg)

                End If
                Logger.WriteToLogFile("IHost", "UpdateSaveWFStatus Flush Data starts")
                objWCF.UpdateSaveWFStatus("FLUSH", Irc, Errmsg)
                Logger.WriteToLogFile("IHost", "UpdateSaveWFStatus Flush Data Ends")

                If Irc <> 0 Then
                    Throw New Exception()
                End If
                If Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.ChildWFsts) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.ChildWFsts.ToUpper() = "TRUE" Then


                    Dim Breject As Boolean = False
                    Dim BReWork As Boolean = False

                    If ChildFullDPValue.Count > 0 Then
                        For icnt As Integer = 0 To ChildFullDPValue.Count - 1
                            Dim closedp As New Dictionary(Of String, String)(ChildFullDPValue.ElementAt(icnt).Value)

                            If closedp.ContainsKey("bRej") AndAlso closedp.Item("bRej") <> "" Then

                                If closedp.Item("bRej") = "1" Or closedp.Item("bRej") = "True" Then
                                    Breject = True
                                End If

                            End If
                            If closedp.ContainsKey("bReWk") AndAlso closedp.Item("bReWk") <> "" Then

                                If closedp.Item("bReWk") = "1" Or closedp.Item("bReWk") = "True" Then
                                    BReWork = True
                                End If

                            End If
                            If Breject = True Or BReWork = True Then
                                objWCF.UpdateChildSaveWFStatus("REJECT", ChildFullDPValue.ElementAt(icnt).Key, Irc, Errmsg)

                            Else
                                Dim isErrormached As Boolean = False
                                ClosingRules(isErrormached, ChildFullDPValue.ElementAt(icnt).Key)


                                If isErrormached = True Then

                                    Exit Sub
                                End If
                            End If
                            closedp = Nothing
                        Next
                    End If
                End If
            End If


        End If





    End Sub


    Private Sub ReadPluginChildDatapart(ByRef FullChildDPValue As Dictionary(Of String, Dictionary(Of String, String)), ByRef plugin As Object)
        Dim strDpvalue As String()
        Dim strDpValues As String
        Dim Irc As Integer = 0
        Dim Errmsg As String = String.Empty
        Try

            'Creating  FullDatapart collection Dictionary(Of String, String)

            Dim DPValue As New Dictionary(Of String, String)

            If Not IsNothing(TryCast(plugin, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then


                strDpValues = Convert.ToString(DirectCast(plugin, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).ChildDataPart)
                If IsNothing(strDpValues) Or (Not IsNothing(strDpValues) AndAlso strDpValues.Length = 0) Then
                    Exit Sub
                End If

                strDpvalue = strDpValues.Split("|")

                For i = 0 To strDpvalue.Length - 1
                    If strDpvalue(i).ToString <> String.Empty AndAlso strDpvalue(i).Substring(0, 1) = "@" Then
                        Logger.WriteToLogFile("IHost", "BuildEmpty DP IndexnRev Start")
                        BuildEmptyDP4IndexnRev(strDpvalue(i), DPValue)
                        Logger.WriteToLogFile("IHost", "BuildEmpty DP IndexnRev Ends")
                    Else
                        If strDpvalue(i).ToString <> String.Empty AndAlso Not DPValue.ContainsKey(strDpvalue(i)) Then
                            DPValue.Add(strDpvalue(i), Nothing)
                        End If
                    End If

                Next
            End If

            Dim lstchild As New List(Of String)
            objWCF.GetchildLinkange(lstchild, Irc, Errmsg)

            If Irc <> 0 Then
                Throw New SympraxisException.WorkitemException(Errmsg)
            End If
            If lstchild.Count > 0 Then
                For i As Integer = 0 To lstchild.Count - 1
                    FullChildDPValue.Add(lstchild(i), DPValue)
            Next


            End If

            objWCF.GetChildDataPart(FullChildDPValue, Irc, Errmsg)
            If Irc <> 0 Then
                Throw New SympraxisException.WorkitemException(Errmsg)
            End If

            ' irc = WCF.Read(hshDataParts, ErrMsg)


        Finally
            strDpvalue = Nothing
            strDpValues = Nothing
        End Try
    End Sub



    Private Sub ReadChildDatapart(ByRef FullChildDPValue As Dictionary(Of String, Dictionary(Of String, String)))
        Dim strDpvalue As String()
        Dim strDpValues As String
        Dim Irc As Integer = 0
        Dim Errmsg As String = String.Empty
        Try

            'Creating  FullDatapart collection Dictionary(Of String, String)

            Dim DPValue As New Dictionary(Of String, String)
            For index As Integer = 0 To lstPluginOrder.Count - 1
                If Not IsNothing(TryCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then


                    strDpValues = Convert.ToString(DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).ChildDataPart)
                    If IsNothing(strDpValues) Or (Not IsNothing(strDpValues) AndAlso strDpValues.Length = 0) Then Continue For
                    strDpvalue = strDpValues.Split("|")

                    For i = 0 To strDpvalue.Length - 1
                        If strDpvalue(i).ToString <> String.Empty AndAlso strDpvalue(i).Substring(0, 1) = "@" Then
                            Logger.WriteToLogFile("IHost", "BuildEmpty DP IndexnRev Start")
                            BuildEmptyDP4IndexnRev(strDpvalue(i), DPValue)
                            Logger.WriteToLogFile("IHost", "BuildEmpty DP IndexnRev Ends")
                        Else
                            If strDpvalue(i).ToString <> String.Empty AndAlso Not DPValue.ContainsKey(strDpvalue(i)) Then
                                DPValue.Add(strDpvalue(i), Nothing)
                            End If
                        End If

                    Next
                End If
            Next
            Dim lstchild As New List(Of String)
            objWCF.GetchildLinkange(lstchild, Irc, Errmsg)

            If Irc <> 0 Then
                Throw New SympraxisException.WorkitemException(Errmsg)
            End If
            If lstchild.Count > 0 Then
                For i As Integer = 0 To lstchild.Count - 1
                    FullChildDPValue.Add(lstchild(i), DPValue)
                Next


            End If

            objWCF.GetChildDataPart(FullChildDPValue, Irc, Errmsg)
            If Irc <> 0 Then
                Throw New SympraxisException.WorkitemException(Errmsg)
            End If

            ' irc = WCF.Read(hshDataParts, ErrMsg)


        Finally
            strDpvalue = Nothing
            strDpValues = Nothing
        End Try
    End Sub



    Private Sub ReadPluginChildDatapart(ByRef FullChildDPValue As Dictionary(Of String, Dictionary(Of String, String)), ByVal pluginindex As Integer)
        Dim strDpvalue As String()
        Dim strDpValues As String
        Dim Irc As Integer = 0
        Dim Errmsg As String = String.Empty
        Try

            'Creating  FullDatapart collection Dictionary(Of String, String)



            If Not IsNothing(TryCast(lstPluginOrder(pluginindex), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                Dim DPValue As New Dictionary(Of String, String)

                strDpValues = Convert.ToString(DirectCast(lstPluginOrder(pluginindex), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).ChildDataPart)
                If IsNothing(strDpValues) Or (Not IsNothing(strDpValues) AndAlso strDpValues.Length = 0) Then
                    Exit Sub
                End If

                strDpvalue = strDpValues.Split("|")

                For i = 0 To strDpvalue.Length - 1
                    If strDpvalue(i).ToString <> String.Empty AndAlso strDpvalue(i).Substring(0, 1) = "@" Then
                        Logger.WriteToLogFile("IHost", "BuildEmpty DP IndexnRev Start")
                        BuildEmptyDP4IndexnRev(strDpvalue(i), DPValue)
                        Logger.WriteToLogFile("IHost", "BuildEmpty DP IndexnRev Ends")
                    Else
                        If strDpvalue(i).ToString <> String.Empty AndAlso Not DPValue.ContainsKey(strDpvalue(i)) Then
                            DPValue.Add(strDpvalue(i), Nothing)
                        End If
                    End If

                Next


                Dim lstchild As New List(Of String)
                objWCF.GetchildLinkange(lstchild, Irc, Errmsg)

                If Irc <> 0 Then
                    Throw New SympraxisException.WorkitemException(Errmsg)
                End If
                If lstchild.Count > 0 Then
                    For i As Integer = 0 To lstchild.Count - 1
                        FullChildDPValue.Add(lstchild(i), DPValue)
                    Next


                End If

                objWCF.GetChildDataPart(FullChildDPValue, Irc, Errmsg)
                If Irc <> 0 Then
                    Throw New SympraxisException.WorkitemException(Errmsg)
                End If
            End If



        Finally
            strDpvalue = Nothing
            strDpValues = Nothing
        End Try
    End Sub


    Private Sub BuildDP4IndexnRev(ByRef strDpvalue As String, ByRef ReturnDp As Dictionary(Of String, String), ByRef DPValue As Dictionary(Of String, String))
        If strDpvalue.Substring(0, 4) = "@CIC" Then
            If strDpvalue = "@CIC" Then
                ReturnDp.Add("CI", DPValue("@CI"))
            ElseIf strDpvalue = "@CIC:" Then
                Dim strDefaultCI As String() = strDpvalue.Split(":")
                If strDefaultCI.Count > 0 AndAlso Not IsNothing(strDefaultCI(1)) AndAlso strDefaultCI(1).ToString <> String.Empty Then

                    ReturnDp.Add("CI", DPValue("@CI:" + strDefaultCI(1)))
                End If
            End If
            ReturnDp.Add("IM", DPValue("@IM"))
            ReturnDp.Add("ID", DPValue("@ID"))

        ElseIf strDpvalue.Substring(0, 5) = "@CRQN" Then
            If strDpvalue = "@CRQN" Then
                ReturnDp.Add("CR", DPValue("@CR"))
            ElseIf strDpvalue.Substring(0, 6) = "@CRQN:" Then
                Dim strDefaultCI As String() = strDpvalue.Split(":")
                If strDefaultCI.Count > 0 AndAlso Not IsNothing(strDefaultCI(1)) AndAlso strDefaultCI(1).ToString <> String.Empty Then

                    ReturnDp.Add("CR", DPValue("@CR:" + strDefaultCI(1)))
                End If

            End If

            ReturnDp.Add("RM", DPValue("@RM"))
            ReturnDp.Add("RD", DPValue("@RD"))
        End If

    End Sub

    Private Sub LoadWorkitem(ByRef DPValue As Dictionary(Of String, String))
        Dim strDpvalue As String()
        Dim strDpValues As String


        Dim Dp As Dictionary(Of String, String) = Nothing
        Try



            For index As Integer = 0 To lstPluginOrder.Count - 1

                strDpValues = Convert.ToString(DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.INormalPlugin).DataPart)
                If Not IsNothing(strDpValues) AndAlso strDpValues <> String.Empty Then
                    Dp = New Dictionary(Of String, String)

                    strDpvalue = strDpValues.Split("|")

                    For i = 0 To strDpvalue.Length - 1
                        If strDpvalue(i).Substring(0, 1) = "@" Then

                            BuildDP4IndexnRev(strDpvalue(i), Dp, DPValue)
                        Else

                            Dp.Add(strDpvalue(i), DPValue(strDpvalue(i)))

                        End If

                    Next
                ElseIf (IsNothing(strDpValues) Or (Not IsNothing(strDpValues) AndAlso strDpValues = String.Empty)) Then
                    Continue For
                End If

                Try

                    Logger.WriteToLogFile("IHost", "Plugin LoadWorkItem -->" & lstPluginOrder(index).ToString() & " Starts")
                    If Not IsNothing(TryCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                        DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrCode = 0
                    End If

                    DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.INormalPlugin).LoadWorkItem(Dp)

                    If Not IsNothing(TryCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                        If DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrCode <> 0 Then
                            Throw New SympraxisException.WorkitemException(DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrMsg)
                        End If
                    End If

                Catch ex As Exception
                    Throw New SympraxisException.WorkitemException(ex.Message)
                End Try
                Logger.WriteToLogFile("IHost", "Plugin LoadWorkItem -->" & lstPluginOrder(index).ToString() & " Ends")
                Dp = Nothing
            Next

            PluginWorkLoaded()


            '  WorkAutoCloseTimerStart()

            AttenandUnattentedTimeStartEvent()




        Finally
            strDpvalue = Nothing
            strDpValues = Nothing
            Dp = Nothing
        End Try
    End Sub


    Dim listParallelDP As List(Of clsParallelDp)
    Private Sub LoadParallelWorkitem()
        Dim ParallelDp As Dictionary(Of String, Dictionary(Of String, String))
        Dim ParallelChildDp As Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, String)))
        Dim DPValue As Dictionary(Of String, String)
        Dim ChildDPValue As Dictionary(Of String, Dictionary(Of String, String))

        Try

            If chkIExtenedPluginImplemented() Then

                listParallelDP = New List(Of clsParallelDp)
                Dim Irc As Integer = 0
                Dim ErrMsg As String = String.Empty
                ParallelDp = New Dictionary(Of String, Dictionary(Of String, String))
                ParallelChildDp = New Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, String)))

                Dim ChkerrorWI As Boolean = False
                For index As Integer = 0 To lstPluginOrder.Count - 1

                    If Not IsNothing(TryCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                        For i As Integer = 1 To objWCF.WorkitemCount

                            objWCF.SwitchWorkitem(i, Irc, ErrMsg)
                            If Irc <> 0 Then
                                Throw New SympraxisException.WorkitemException(ErrMsg)
                            End If
                            ChkerrorWI = False
                            objWCF.ChkcurrentWorkitemErrorstatus(ChkerrorWI, Irc, ErrMsg)

                            If Irc <> 0 Then
                                Throw New SympraxisException.WorkitemException(ErrMsg)
                            End If
                            '' Fail 
                            If ChkerrorWI = False Then

                                DPValue = New Dictionary(Of String, String)
                                ChildDPValue = New Dictionary(Of String, Dictionary(Of String, String))

                                ReadPluginDatapart(DPValue, lstPluginOrder(index))

                                If DPValue.Count > 0 Then
                                    ParallelDp.Add(objWCF.Objectid, DPValue)
                                End If

                                ReadPluginChildDatapart(ChildDPValue, lstPluginOrder(index))

                                If ChildDPValue.Count > 0 Then
                                    ParallelChildDp.Add(objWCF.Objectid, ChildDPValue)
                                End If
                                DPValue = Nothing
                                ChildDPValue = Nothing

                            End If
                        Next

                        If ParallelDp.Count > 0 Then
                            listParallelDP.Add(New clsParallelDp With {.pluginName = lstPluginOrder(index).ToString(), .ParallelDP = ParallelDp, .ParallelChildDp = ParallelChildDp})

                            DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).LoadWorkItem(ParallelDp, ParallelChildDp)
                        End If

                    End If
                Next

            Else
                Throw New SympraxisException.SettingsException("IExtented")
            End If
        Finally
            DPValue = Nothing
            ChildDPValue = Nothing
            ParallelDp = Nothing
            ParallelChildDp = Nothing
        End Try
    End Sub
    Public Class clsParallelDp
        Property pluginName As String
        Property ParallelDP As Dictionary(Of String, Dictionary(Of String, String))
        Property ParallelChildDp As Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, String)))


    End Class

    Private Function chkIExtenedPluginImplemented()
        Dim BReturn As Boolean = False
        For index As Integer = 0 To lstPluginOrder.Count - 1
            If Not IsNothing(TryCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                BReturn = True
                Exit For
            End If
        Next
        Return BReturn
    End Function

    Private Sub PluginWorkLoaded()

        Dim irc As Integer = 0
        Dim Errmsg As String = String.Empty

        If Not IsNothing(objWCF) AndAlso Not IsNothing(objWCF.Objectid) Then

            For index As Integer = 0 To lstPluginOrder.Count - 1


                If Not IsNothing(TryCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                    Logger.WriteToLogFile("IHost", "Plugin WorkLoaded -->" & lstPluginOrder(index).ToString() & " Starts")
                    DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).WorkLoaded(objWCF.Objectid, irc, Errmsg)
                    If irc <> 0 Then
                        Throw New SympraxisException.WorkitemException(Errmsg)
                    End If
                    Logger.WriteToLogFile("IHost", "Plugin WorkLoaded -->" & lstPluginOrder(index).ToString() & " Ends")
                End If
            Next
            For index As Integer = 0 To lstStartupPluginOrder.Count - 1


                If Not IsNothing(TryCast(lstStartupPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                    Logger.WriteToLogFile("IHost", "Plugin WorkLoaded -->" & lstStartupPluginOrder(index).ToString() & " Starts")
                    DirectCast(lstStartupPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).WorkLoaded(objWCF.Objectid, irc, Errmsg)
                    If irc <> 0 Then
                        Throw New SympraxisException.WorkitemException(Errmsg)
                    End If
                    Logger.WriteToLogFile("IHost", "Plugin WorkLoaded -->" & lstStartupPluginOrder(index).ToString() & " Ends")
                End If
            Next


        End If

    End Sub

    Private Sub UpdateEnvironmentvariable()

        'WCF.EnvironmentVariables(irc, ErrMsg)




        configMgr.AppbaseConfig.EnvironmentVariables.AppExchange = objWCF.ObjAppExchange
        configMgr.AppbaseConfig.EnvironmentVariables.OrganizationName = objWCF.OrganizationName
        configMgr.AppbaseConfig.EnvironmentVariables.OrganizationSName = objWCF.OrganizationSName
        configMgr.AppbaseConfig.EnvironmentVariables.WorkCenterName = objWCF.WorkCenterName
        configMgr.AppbaseConfig.EnvironmentVariables.WorkCenterSName = objWCF.WorkCenterSName
        configMgr.AppbaseConfig.EnvironmentVariables.sWorkGroup = objWCF.sWorkGroup
        configMgr.AppbaseConfig.EnvironmentVariables.sWorkGroupDesc = objWCF.sWorkGroupDesc
        configMgr.AppbaseConfig.EnvironmentVariables.iOrgId = objWCF.iOrgId
        configMgr.AppbaseConfig.EnvironmentVariables.iWorkCenterId = objWCF.iWorkCenterId
        configMgr.AppbaseConfig.EnvironmentVariables.lWorkGroupId = objWCF.lWorkGroupId
        configMgr.AppbaseConfig.EnvironmentVariables.lUserId = objWCF.lUserId

        If Not IsNothing(objWCF.sUserId) Then
            configMgr.AppbaseConfig.EnvironmentVariables.sUserId = objWCF.sUserId
        Else
            configMgr.AppbaseConfig.EnvironmentVariables.sUserId = Environment.UserName
        End If

        If Not IsNothing(objWCF.myLoginRsp) Then
            configMgr.AppbaseConfig.EnvironmentVariables.myLoginRsp = objWCF.myLoginRsp
            If Not IsNothing(objWCF.myLoginRsp.stWFInfo) AndAlso Not IsNothing(objWCF.myLoginRsp.stWFInfo.aUser) Then
                configMgr.AppbaseConfig.EnvironmentVariables.UserName = String.Concat(configMgr.AppbaseConfig.EnvironmentVariables.myLoginRsp.stWFInfo.aUser(0).sFirstName, " ", configMgr.AppbaseConfig.EnvironmentVariables.myLoginRsp.stWFInfo.aUser(0).sLastName)
            End If
        End If
        If Not IsNothing(objWCF.WFResourceRequest) Then
            configMgr.AppbaseConfig.EnvironmentVariables.WFResourceRequest = objWCF.WFResourceRequest
        End If


        configMgr.AppbaseConfig.EnvironmentVariables.lProcessId = objWCF.lProcessId
        'Dim aaerr As New testerr
        'aaerr.ant1 = "cv"
        'aaerr.ant = "dsdfdsf"
        If Not IsNothing(overrideAppHostConfig.TransDBURL) Then
            configMgr.AppbaseConfig.TransDBURL = overrideAppHostConfig.TransDBURL
        End If
        If Not IsNothing(overrideAppHostConfig.TransDBWebURL) Then
            configMgr.AppbaseConfig.TransDBWebURL = overrideAppHostConfig.TransDBWebURL
        End If
        If Not IsNothing(overrideAppHostConfig.ARMSURL) Then
            configMgr.AppbaseConfig.ArmsURL = overrideAppHostConfig.ARMSURL
        End If
    End Sub


    Private Sub ChangeWorkFlow(ByRef ProcessFlag As Boolean)
        Dim Irc As Integer
        Dim Errmsg As String = String.Empty



        For index As Integer = lstPluginOrder.Count - 1 To 0 Step -1

            If Not IsNothing(TryCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                Logger.WriteToLogFile("IHost", "Plugin Dismantle -->" & lstPluginOrder(index).ToString() & " Starts")
                DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).Dismantle(Irc, Errmsg)
                If Irc <> 0 Then

                    Throw New SympraxisException.WorkitemException(Errmsg)
                End If
                Logger.WriteToLogFile("IHost", "Plugin Dismantle Finished -->" & lstPluginOrder(index).ToString() & " Ends")
            End If

        Next


        If Not IsNothing(PluginMenuEnableDisable) Then
            For Each MenuElement In PluginMenuEnableDisable
                AHEnableDisableMenu(MenuElement.Key, MenuElement.Value, False)
            Next
        End If

        intrHostWorkFlowName = objWCF.workFlowName
        intrHostProcessName = objWCF.processName
        appStartBy = objWCF.processName


        Logger.WriteToLogFile("IHost", "Load Config Started")
        LoadConfig(appHostConfigXml, loadPluginsConfigXml, False)
        If Irc <> 0 Then Exit Sub
        Logger.WriteToLogFile("IHost", "Load Config Ends")

        Logger.WriteToLogFile("IHost", "Init WCF Started")
        Irc = objWCF.Init(configMgr, intrHostProcessName, Errmsg)
        If Irc <> 0 Then
            Throw New SympraxisException.SettingsException(Errmsg)
        End If
        Logger.WriteToLogFile("IHost", "Init WCF Completed")


        Logger.WriteToLogFile("IHost", "Assigning EnvironmentVariable Started")
        EnvironmentVariable()
        Logger.WriteToLogFile("IHost", "Assigning EnvironmentVariable Completed")

        'ApplyLayout(0)
        'mLayoutmanager.Tflag = False
        'mLayoutmanager.clear()
        'mLayoutmanager.CleartheLayoutdocumentvalues()
        ''mLayoutmanager = New LayoutManager

        'mLayoutmanager.ClearCollection()


        'mLayoutmanager.UpdateLayout()


        If appStartBy.ToUpper() <> "STARTUP" AndAlso workingMode <> "\debug" AndAlso processType.ToUpper() = "WORKFLOW" Then

            Logger.WriteToLogFile("IHost", "Attach Process Started")
            AttachProcess(ProcessFlag)
            If Irc <> 0 Then Exit Sub
            Logger.WriteToLogFile("IHost", "Attach Process Ends")

            Logger.WriteToLogFile("IHost", "Register WF Starts")
            RegisterWF(ProcessFlag)
            If Irc <> 0 Then Exit Sub
            Logger.WriteToLogFile("IHost", "Register WF Ends")
        End If









        'If Not IsNothing(defaultPluginsConfigXml.ApplicationHost.Plugins) Then
        '    For A As Integer = 0 To defaultPluginsConfigXml.ApplicationHost.Plugins.Count - 1
        '        If Not loadPluginsConfigXml.ApplicationHost.Plugins.Contains(defaultPluginsConfigXml.ApplicationHost.Plugins(A)) Then
        '            loadPluginsConfigXml.ApplicationHost.Plugins.Add(defaultPluginsConfigXml.ApplicationHost.Plugins(A))
        '            ''HshDataPlugins.Remove(ApplicationFrameWorkConfig.Plugins(0).Plugins(0))
        '        End If
        '    Next
        'End If

        If Not IsNothing(defaultPluginsConfigXml) AndAlso Not IsNothing(defaultPluginsConfigXml.ApplicationHost.LayoutSettings) Then
            For b As Integer = 0 To defaultPluginsConfigXml.ApplicationHost.LayoutSettings(0).Layouts.Count - 1
                If Not loadPluginsConfigXml.ApplicationHost.LayoutSettings(0).Layouts.Contains(defaultPluginsConfigXml.ApplicationHost.LayoutSettings(0).Layouts(b)) Then
                    loadPluginsConfigXml.ApplicationHost.LayoutSettings(0).Layouts.Add(defaultPluginsConfigXml.ApplicationHost.LayoutSettings(0).Layouts(b))
                End If
            Next
        End If
        'If appStartBy.ToUpper() <> "STARTUP" AndAlso tempappstart.ToUpper() <> "STARTUP" Then
        '    Dispatcher.BeginInvoke(Sub() DelegateApplylayout(0))
        'End If
        'If appStartBy.ToUpper() <> "STARTUP" Then
        '    Dispatcher.BeginInvoke(Sub() DelegateApplylayout(0))
        'End If

    End Sub

    Private Sub SetApplyLayout()
        Dim LayoutManifestlist As New List(Of SetLayoutManifest)
        Dim LstRulelist As New List(Of Rule)
        Dim MatchedRule As New List(Of Rule)
        Dim CrRule As New List(Of Rule)
        Dim Irc As Integer = 0
        Dim Errmsg As String = ""

        If Not IsNothing(overrideAppHostConfig.ApplyLayoutConfig) AndAlso Not IsNothing(overrideAppHostConfig.ApplyLayoutConfig.LayoutManifest) Then
            LayoutManifestlist = overrideAppHostConfig.ApplyLayoutConfig.LayoutManifest.Where(Function(x) If(x.ProcessName.Contains("|"), x.ProcessName.Split("|").Any(Function(y) y = intrHostProcessName), x.ProcessName = intrHostProcessName Or x.ProcessName = "")).ToList()
            If Not IsNothing(LayoutManifestlist) AndAlso LayoutManifestlist.Count > 0 Then

                If LayoutManifestlist.Any(Function(x) x.MatchID <> "") AndAlso Not IsNothing(overrideAppHostConfig.ApplyLayoutConfig.Rules) AndAlso Not IsNothing(overrideAppHostConfig.ApplyLayoutConfig.Rules.Ruleslist) Then
                    LstRulelist = overrideAppHostConfig.ApplyLayoutConfig.Rules.Ruleslist.Where(Function(x) LayoutManifestlist.Any(Function(y) y.MatchID = x.Name)).ToList()

                    If Not IsNothing(LstRulelist) AndAlso LstRulelist.Count = 0 Then
                        GetCrRulelist(CrRule, LstRulelist, overrideAppHostConfig.BannerSettings.Rules.Ruleslist)
                        objWCF.GetMatchRule(MatchedRule, LstRulelist, CrRule, Irc, Errmsg)
                        If Not IsNothing(MatchedRule) AndAlso MatchedRule.Count > 0 Then
                            Dim layoutnum As String = overrideAppHostConfig.ApplyLayoutConfig.LayoutManifest.Where(Function(x) MatchedRule.Any(Function(y) y.Name = x.MatchID)).Select(Function(x) x.LayoutNumber).FirstOrDefault()
                            If Not String.IsNullOrEmpty(layoutnum) Then
                                ApplyLayout(layoutnum)
                            End If
                        End If
                    Else
                        Dim layoutnum As String = LayoutManifestlist(0).LayoutNumber
                        If Not String.IsNullOrEmpty(layoutnum) Then
                            ApplyLayout(layoutnum)
                        End If
                    End If
                Else
                    Dim layoutnum As String = LayoutManifestlist(0).LayoutNumber
                    If Not String.IsNullOrEmpty(layoutnum) Then
                        ApplyLayout(layoutnum)
                    End If
                End If

            End If
        End If
    End Sub


    Private Sub SetHelpfile()

        Dim HelpFileManifest As New List(Of SetHelpFileManifest)
        

        Dim irc As Integer = 0
        Dim Errmsg As String = String.Empty

        Try

            If Not IsNothing(overrideAppHostConfig.HelpFileSettings) AndAlso Not IsNothing(overrideAppHostConfig.HelpFileSettings.SetHelpFileManifest) Then




                Logger.WriteToLogFile("Ihost", "SetHelpfile Starts")

                If Not String.IsNullOrEmpty(intrHostProcessName) Then
                    HelpFileManifest = overrideAppHostConfig.HelpFileSettings.SetHelpFileManifest.Where(Function(x) If(x.ProcessName.Contains("|"), x.ProcessName.Split("|").Any(Function(y) y = intrHostProcessName), x.ProcessName = intrHostProcessName)).ToList()



                    If Not IsNothing(HelpFileManifest) AndAlso HelpFileManifest.Count > 0 Then
                        Dim xpath As String = HelpFileManifest(0).HTMLFileName

                        If File.Exists(xpath) = False Then

                            ShowErrorToolstrip("Help File not exists", SympraxisException.ErrorTypes.Information, False)

                            Exit Sub
                        End If
                        Dim help As New Help()
                        help.HelpUrl = xpath
                        help.WindowState = System.Windows.WindowState.Maximized
                        help.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
                        help.ShowDialog()


                        Logger.WriteToLogFile("Ihost", "SetMatchBanner Ends")
                    End If
                End If
            Else
                ShowErrorToolstrip("Helpfile  is  not configured  for " & intrHostProcessName & " Process", SympraxisException.ErrorTypes.Information, False)
                Logger.WriteToLogFile("Ihost", "Helpfile  is  not configured  for " & intrHostProcessName & " Process")
                Exit Sub
                'To do 19/12/19 Logg the not match banner
            End If

        Finally

            '    BannerManifestlist = Nothing


         
        End Try
    End Sub



    Private Sub SetMatchBanner()

        Dim BannerManifestlist As New List(Of SetBannerManifest)
        Dim LstRulelist As New List(Of Rule)
        Dim MatchedRule As New List(Of Rule)
        Dim CrRule As New List(Of Rule)

        Dim irc As Integer = 0
        Dim Errmsg As String = String.Empty

        Try

            If Not IsNothing(overrideAppHostConfig.BannerSettings) AndAlso Not IsNothing(overrideAppHostConfig.BannerSettings.BannerManifest) AndAlso Not IsNothing(overrideAppHostConfig.BannerSettings.Rules) AndAlso Not IsNothing(overrideAppHostConfig.BannerSettings.Rules.Ruleslist) AndAlso Not IsNothing(overrideAppHostConfig.BannerSettings.BannerMeta) Then




                Logger.WriteToLogFile("Ihost", "SetMatchBanner Starts")

                If Not String.IsNullOrEmpty(intrHostProcessName) Then
                    BannerManifestlist = overrideAppHostConfig.BannerSettings.BannerManifest.Where(Function(x) If(x.ProcessName.Contains("|"), x.ProcessName.Split("|").Any(Function(y) y = intrHostProcessName), x.ProcessName = intrHostProcessName Or x.ProcessName = "")).OrderBy(Function(x) x.CaptureSequence).ToList()
                End If


                If Not IsNothing(BannerManifestlist) AndAlso BannerManifestlist.Count > 0 Then


                    LstRulelist = overrideAppHostConfig.BannerSettings.Rules.Ruleslist.Where(Function(x) BannerManifestlist.Any(Function(y) y.MatchID = x.Name)).ToList()

                    If Not IsNothing(LstRulelist) AndAlso LstRulelist.Count = 0 Then
                        configBannersettings = overrideAppHostConfig.BannerSettings.BannerMeta.Where(Function(x) x.Name = BannerManifestlist(0).BannerName).FirstOrDefault()
                        If Not IsNothing(configBannersettings) Then


                            Logger.WriteToLogFile("Ihost", "ApplyBanner Starts")
                            ApplyBanner(configBannersettings)

                            Logger.WriteToLogFile("Ihost", "ApplyBanner Ends")


                        End If
                    Else
                        GetCrRulelist(CrRule, LstRulelist, overrideAppHostConfig.BannerSettings.Rules.Ruleslist)

                        '  MessageBox.Show(CrRule.Item(0).Name)
                        Logger.WriteToLogFile("Ihost", "GetMatchRule Starts")

                        objWCF.GetMatchRule(MatchedRule, LstRulelist, CrRule, irc, Errmsg)


                        If irc <> 0 Then
                            Throw New SympraxisException.AppException(Errmsg)

                        End If
                        Logger.WriteToLogFile("Ihost", "GetMatchRule Ends")

                        If Not IsNothing(MatchedRule) AndAlso MatchedRule.Count > 0 Then

                            Dim bannername As String = BannerManifestlist.Where(Function(x) MatchedRule.Any(Function(y) y.Name = x.MatchID)).Select(Function(x) x.BannerName).FirstOrDefault()
                            If Not String.IsNullOrEmpty(bannername) Then
                                configBannersettings = overrideAppHostConfig.BannerSettings.BannerMeta.Where(Function(x) x.Name = bannername).FirstOrDefault()
                                If Not IsNothing(configBannersettings) Then
                                    Logger.WriteToLogFile("Ihost", "ApplyBanner Starts")
                                    ApplyBanner(configBannersettings)
                                    If irc <> 0 Then Exit Sub
                                    Logger.WriteToLogFile("Ihost", "ApplyBanner Ends")
                                Else
                                    'To do 19/12/19 Logg the not match banner
                                    Throw New Exception("Banner match setings is nothing")
                                    Logger.WriteToLogFile("Ihost", "Banner match setings is not configured in " & intrHostProcessName & "Process")



                                End If
                            End If
                        End If

                    End If
                    Logger.WriteToLogFile("Ihost", "SetMatchBanner Ends")
                End If
            Else

                Logger.WriteToLogFile("Ihost", "Banner  is  not configured  for this" & intrHostProcessName & "Process")
                Exit Sub
                'To do 19/12/19 Logg the not match banner
            End If
            Dim bannerheight As Integer = 0
            If BannerPanel.Visibility = System.Windows.Visibility.Visible Then
                If configBannersettings IsNot Nothing AndAlso Not String.IsNullOrEmpty(configBannersettings.Height) AndAlso IsNumeric(configBannersettings.Height) Then
                    bannerheight = configBannersettings.Height
                End If
            End If
            DAXObject.SetValue("SYSBannerHeight", bannerheight)
        Finally

            BannerManifestlist = Nothing


            LstRulelist = Nothing
            MatchedRule = Nothing
            CrRule = Nothing
        End Try
    End Sub
        Dim BannerDocuments As XmlDocument
    Dim CaseInfo As Object
    Private Sub ApplyBanner(ByRef Bannersettings As Banner)
        Dim irc As Integer = 0
        Dim Errmsg As String = String.Empty

        CaseInfo = Nothing
        If IsNothing(CaseInfo) Then
            CaseInfo = New Plugins.CaseInfo.CaseInfoControl
        End If

        BannerPanel.Children.Clear()
        BannerPanel.Children.Add(CaseInfo)
        BannerPanel.Visibility = System.Windows.Visibility.Visible
        If Not String.IsNullOrEmpty(Bannersettings.Height) Then
            BannerPanel.Height = Bannersettings.Height
        End If

        irc = CaseInfo.LoadXaml(Bannersettings.XamlPath)



        If irc = 0 Then
            Logger.WriteToLogFile("Ihost", "ApplyBanner BindSource XpathValues St")
            BannerDocuments = Nothing
            BannerDocuments = New Xml.XmlDocument
            Dim provider As New XmlDataProvider
            provider = CaseInfo.GetXmlDoc()
            If Not IsNothing(provider.Document) AndAlso Not String.IsNullOrEmpty(provider.Document.InnerXml) Then
                BannerDocuments.LoadXml(provider.Document.InnerXml)
                Dim SorceXpath As String = ""
                Dim Targetvalue As String = ""
                If Not IsNothing(Bannersettings.F) Then


                    For i = 0 To Bannersettings.F.Count - 1
                        Targetvalue = String.Empty
                        If Not IsNothing(Bannersettings.F(i).Type) Then
                            If Bannersettings.F(i).Type.ToString.ToUpper = "WI" Or Bannersettings.F(i).Type.ToString.ToUpper = "WP" Then

                                If Not Bannersettings.F(i).SXP.ToString.StartsWith("{") Then
                                    If Not String.IsNullOrEmpty(Bannersettings.F(i).SXP) Then


                                        objWCF.GetDataElementValue(Bannersettings.F(i).SXP, Targetvalue, irc, Errmsg)
                                        If irc <> 0 Then
                                            'To do 19/12/19 Logg the handle the GetData Element Exception
                                            Throw New SympraxisException.AppException(Errmsg)
                                            Logger.WriteToLogFile("Ihost", "Error Occured while Get Data Element Value " & Bannersettings.F(i).SXP & "in apply banner")
                                        End If


                                        If Not String.IsNullOrEmpty(Targetvalue) Then
                                            If Not String.IsNullOrEmpty(Bannersettings.F(i).TXP) Then
                                                BannerDocuments.SelectSingleNode(Bannersettings.F(i).TXP).InnerXml = Targetvalue
                                                CaseInfo.UpdateXmlDoc(BannerDocuments)
                                            End If

                                        End If
                                    Else
                                        Throw New SympraxisException.SettingsException("Source Xapth should be empty")
                                    End If
                                End If
                            ElseIf Bannersettings.F(i).Type.ToString.ToUpper() = "DAX" Then
                                If Bannersettings.F(i).SXP = "{SYSCURRENTOBJIDX}" Then
                                    DAXObject.SetValue("SYSCURRENTOBJIDX", objWCF.CurrentWorkitemIdx)
                                End If
                                If Bannersettings.F(i).SXP = "{SYSWORKITEMCOUNT}" Then
                                    DAXObject.SetValue("SYSWORKITEMCOUNT", objWCF.WorkitemCount)
                                End If

                                If Bannersettings.F(i).SXP.ToString.StartsWith("{") Then
                                    Targetvalue = DAXObject.GetValue(Bannersettings.F(i).SXP.Replace("{", "").Replace("}", ""))
                                    If Not String.IsNullOrEmpty(Targetvalue) Then
                                        If Not String.IsNullOrEmpty(Bannersettings.F(i).TXP) Then
                                            BannerDocuments.SelectSingleNode(Bannersettings.F(i).TXP).InnerXml = Targetvalue
                                            CaseInfo.UpdateXmlDoc(BannerDocuments)
                                        End If
                                    End If
                                End If


                            End If

                        End If

                    Next

                End If

            End If
            Logger.WriteToLogFile("Ihost", "ApplyBanner BindSource XpathValues Ends")
        Else
            Throw New SympraxisException.AppException("Error occured while load xaml in caseinfo" & Bannersettings.XamlPath)
        End If

    End Sub
    Private Sub CreateIOmanagerDirectory()

        Dim IRC As Integer = 0

        If Not IsNothing(objWCF.WIPPath) Then
            IOManager.WIPPath = objWCF.WIPPath
            IRC = IOManager.CreateDirectory(IOManager.WIPPath)
            If IRC <> 0 Then

                Throw New SympraxisException.AppException("Cannot Create path" & IOManager.WIPPath)
            End If
        Else
            IOManager.WIPPath = Nothing
        End If

        If Not IsNothing(objWCF.Workpacketid) Then
            IOManager.WorkpacketId = objWCF.Workpacketid
        End If

        CreateLog(objWCF.Workpacketid)


        If Not IsNothing(appHostConfigXml.LocalFolder) Then
            If Not IsNothing(objWCF.myLoginRsp) AndAlso Not IsNothing(objWCF.myLoginRsp.stUserSession) Then
                IOManager.LocalFolder = String.Concat(appHostConfigXml.LocalFolder, "\", objWCF.myLoginRsp.stUserSession.iSessionId, "\", intrHostProcessName)
            Else
                IOManager.LocalFolder = String.Concat(appHostConfigXml.LocalFolder, "\", intrHostProcessName)
            End If
            IRC = IOManager.CreateDirectory()
            If IRC <> 0 Then
                Throw New SympraxisException.SettingsException("Cannot Create path" & appHostConfigXml.LocalFolder & "\" & intrHostProcessName)
            End If
        End If

    End Sub
    Private Sub DeleteWorkDirectory()
        Try

       
        If Not IsNothing(IOManager.LocalFolder) Then
            If Directory.Exists(IOManager.LocalFolder) Then
                If Not IsNothing(objWCF.myLoginRsp) AndAlso Not IsNothing(objWCF.myLoginRsp.stUserSession) Then
                    IOManager.DeleteDirectory(String.Concat(appHostConfigXml.LocalFolder, "\", objWCF.myLoginRsp.stUserSession.iSessionId))
                Else
                    IOManager.DeleteDirectory(String.Concat(appHostConfigXml.LocalFolder, "\", intrHostProcessName))
                End If
            End If
        End If
        Catch ex As Exception

        End Try

    End Sub
    Private Sub DeleteWIPWorkDirectory()
        Try
            If Not IsNothing(IOManager.LocalFolder) Then
                If Directory.Exists(IOManager.LocalFolder) Then
                    If Not IsNothing(objWCF.myLoginRsp) AndAlso Not IsNothing(objWCF.myLoginRsp.stUserSession) Then
                        IOManager.DeleteDirectory(String.Concat(appHostConfigXml.LocalFolder, "\", objWCF.myLoginRsp.stUserSession.iSessionId) & "\" & intrHostProcessName & "\" & objWCF.Workpacketid)
                    Else
                        IOManager.DeleteDirectory(String.Concat(appHostConfigXml.LocalFolder, "\", intrHostProcessName, "\", objWCF.Workpacketid))
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub FormMenuHeader(sender As Object, e As EventArgs)

        Dim Menuname As String = DirectCast(sender, System.Windows.Forms.ToolStripMenuItem).Text.ToString()
        '  Dim Headername As String = DirectCast(DirectCast(sender, System.Windows.Forms.ToolStripMenuItem).OwnerItem, System.Windows.Forms.ToolStripMenuItem).ToString()
        AHEnableDisableMenu(Menuname, "", DirectCast(sender, System.Windows.Forms.ToolStripMenuItem).Enabled)
    End Sub
    Private Sub FormMenuenable(sender As Object, e As EventArgs)

        Dim Menuname As String = DirectCast(sender, System.Windows.Forms.ToolStripMenuItem).Name.ToString()
        Dim Headername As String = DirectCast(DirectCast(sender, System.Windows.Forms.ToolStripMenuItem).OwnerItem, System.Windows.Forms.ToolStripMenuItem).ToString()
        AHEnableDisableMenu(Headername, Menuname, DirectCast(sender, System.Windows.Forms.ToolStripMenuItem).Enabled)
    End Sub



    Private Sub SetupErrorMsg(ByVal ex As Exception)
        Try


            Dim Irc As Integer = 0
            Dim Errmsg As String = String.Empty

            HostErrmsg = ex.Message

            If ex.GetType().Name = "SettingsException" Then

                lblToolstripError.Background = Brushes.Red
                lblToolstripError.Foreground = Brushes.White
                ToolstripError.Background = Brushes.Red
                ShowErrorToolstrip("Sympraxis Configuration Error : Error Occurred while setting up the application. Please close the application and contact support.", SympraxisException.ErrorTypes.AppError, True)
                If configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True Then

                    BuildnUpdateErrorInWP(ex)
                    AHMenuCloseWork(True)
                    configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = False

                End If
                AHEnableDisableMenu("Home", "Open|Previous|Next|Close", False)
            ElseIf ex.GetType().Name = "AppException" Then

                ShowErrorToolstrip("Sympraxis Application Exception Error Occured . Please close the application and contact support.", SympraxisException.ErrorTypes.AppError, True)
                If configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True Then
                    objWCF.UpdateSaveWFStatus("Error", Irc, Errmsg)
                    If Irc <> 0 Then
                        Throw New Exception()
                    End If
                    BuildnUpdateErrorInWP(ex)
                    AHMenuCloseWork(True)

                    configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = False

                End If
                AHEnableDisableMenu("Home", "Open|Previous|Next|Close", False)
            ElseIf ex.GetType().Name = "AppExceptionSaveWP" Then
                ShowErrorToolstrip("Sympraxis Application Exception Error Occured . Please close the application and contact support.", SympraxisException.ErrorTypes.AppError, True)
                If configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True Then
                    objWCF.UpdateSaveWFStatus("Error", Irc, Errmsg)
                    If Irc <> 0 Then
                        Throw New Exception()
                    End If
                    BuildnUpdateErrorInWP(ex)
                    AHMenuCloseWork(True)
                    SaveErrorWp()
                    configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = False
                End If
                AHEnableDisableMenu("Home", "Open|Previous|Next|Close", False)
                objWCF.ApplogEntry(ex.Message, Common.ExceptionManager.ExceptionTypes.None, Irc, Errmsg)

            ElseIf ex.GetType().Name = "WorkitemException" Then
                Dim Color As Color
                Color = ColorConverter.ConvertFromString("#ffffcc")
                lblToolstripError.Background = New System.Windows.Media.SolidColorBrush(Color)
                lblToolstripError.Foreground = Brushes.Black
                ToolstripError.Background = New System.Windows.Media.SolidColorBrush(Color)

                If configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True Then
                    ShowWaitCursor(False)
                    BuildnUpdateErrorInWP(ex)
                    If Prvs = False Then
                        If objWCF.CurrentWorkitemIdx < objWCF.WorkitemCount Then
                            If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.ParallelWorkItem) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.ParallelWorkItem.ToUpper = "TRUE" Then
                                Dim DialogResult As DialogResult = MessageBox.Show("Sympraxis Error : Could not able to work on this item, Press OK to close the work packet", "Error", MessageBoxButtons.OK)
                                Logger.WriteToLogFile("IHost", "Sympraxis Error : Could not able to work on this item, Press OK to close the work packet")
                                If DialogResult = DialogResult.OK Then
                                    AHMenuCloseWork(True)
                                End If
                            Else
                            Dim DialogResult As DialogResult = MessageBox.Show("Sympraxis Error : Could not able to work on this item.Press OK to move to next item,Press Cancel to close the work packet", "Error", MessageBoxButtons.OKCancel)
                            Logger.WriteToLogFile("IHost", "Sympraxis Error : Could not able to work on this item.Press OK to move to next item,Press Cancel to close the work packet")

                            objWCF.UpdateSaveWFStatus("Error", Irc, Errmsg)
                            If Irc <> 0 Then
                                Throw New Exception()
                            End If

                            If DialogResult = DialogResult.OK Then

                                AHMenuNextWork(True)
                            Else
                                AHMenuCloseWork(True)
                            End If
                            End If



                        ElseIf objWCF.CurrentWorkitemIdx = objWCF.WorkitemCount Then
                            Dim DialogResult As DialogResult = MessageBox.Show("Sympraxis Error : Could not able to work on this item, Press OK to close the work packet", "Error", MessageBoxButtons.OK)
                            Logger.WriteToLogFile("IHost", "Sympraxis Error : Could not able to work on this item, Press OK to close the work packet")
                            If DialogResult = DialogResult.OK Then
                                AHMenuCloseWork(True)
                            End If


                        End If
                    ElseIf Prvs = True Then
                        If objWCF.CurrentWorkitemIdx > 1 AndAlso objWCF.CurrentWorkitemIdx < objWCF.WorkitemCount Then
                            If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.ParallelWorkItem) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.ParallelWorkItem.ToUpper = "TRUE" Then
                                Dim DialogResult As DialogResult = MessageBox.Show("Sympraxis Error : Could not able to work on this item, Press OK to close the work packet", "Error", MessageBoxButtons.OK)
                                Logger.WriteToLogFile("IHost", "Sympraxis Error : Could not able to work on this item, Press OK to close the work packet")
                                If DialogResult = DialogResult.OK Then
                                    AHMenuCloseWork(True)
                                End If
                            Else
                            Dim DialogResult As DialogResult = MessageBox.Show("Sympraxis Error : Could not able to work on this item.Press OK to move to Previous item,Press Cancel to close the work packet", "Error", MessageBoxButtons.OKCancel)
                            Logger.WriteToLogFile("IHost", "Sympraxis Error : Could not able to work on this item.Press OK to move to Previous item,Press Cancel to close the work packet")
                            objWCF.UpdateSaveWFStatus("Error", Irc, Errmsg)
                            If Irc <> 0 Then
                                Throw New Exception()
                            End If

                            If DialogResult = DialogResult.OK Then

                                AHMenuPreviousWork(True)
                                Prvs = False
                            Else
                                AHMenuCloseWork(True)
                            End If
                            End If
                        ElseIf objWCF.CurrentWorkitemIdx = 1 Then
                            Dim DialogResult As DialogResult = MessageBox.Show("Sympraxis Error : Could not able to work on this item, Press OK to close the work packet", "Error", MessageBoxButtons.OK)
                            Logger.WriteToLogFile("IHost", "Sympraxis Error : Could not able to work on this item, Press OK to close the work packet")
                            If DialogResult = DialogResult.OK Then
                                AHMenuCloseWork(True)
                            End If


                        End If
                    End If
                Else

                    ShowErrorToolstrip(ex.Message, SympraxisException.ErrorTypes.AppError, True)

                End If
            ElseIf ex.GetType().Name = "InformationException" Then
                ShowErrorToolstrip(ex.Message, SympraxisException.ErrorTypes.Information, True)
            ElseIf ex.GetType().Name = "EmptyException" Then

            Else
                Dim Color As Color
                Color = ColorConverter.ConvertFromString("#ffffcc")
                lblToolstripError.Background = New System.Windows.Media.SolidColorBrush(Color)
                lblToolstripError.Foreground = Brushes.Black
                ToolstripError.Background = New System.Windows.Media.SolidColorBrush(Color)
                ShowErrorToolstrip(ex.Message, SympraxisException.ErrorTypes.AppError, True)
            End If

            IHErrException = ex



        Catch ex1 As Exception
            SetupErrorMsg(ex1)
        End Try

    End Sub
    Private Sub BuildnUpdateErrorInWP(ByRef Ex As Exception)
        Dim irc As Integer = 0
        Dim Errmsg As String = String.Empty
        Dim ErrorMessage As String = ""
        ObjFrmException.BuildErrorString(ErrorMessage, Ex)
        objWCF.UpdatexmlErrors(ErrorMessage, irc, Errmsg)
    End Sub

    'Private Sub SetupErrorMsg(ByVal ex As Exception, ByVal errmsg As String, ByVal errmsghdr As String)
    '    HostIrc = 1
    '    If errmsg.Trim() <> "" Then
    '        HostErrmsg = errmsg
    '    Else
    '        If Not IsNothing(ex) Then
    '            HostErrmsg = ex.Message
    '        End If
    '    End If

    '    IHErrormsgHeader = errmsghdr
    '    IHErrException = ex
    '    IHErrormsg = HostErrmsg
    '    ShowErrorToolstrip(IHErrormsg)
    'End Sub

    Private Sub DelegateApplylayout(index As Integer)
        mLayoutmanager.PaneRetainchanges()
        mLayoutmanager.RetainTitle()
        mLayoutmanager.clear()
        mLayoutmanager.CleartheLayoutdocumentvalues()

        LMLayoutChanged(index)

        mLayoutmanager.ApplyLayout(index)


        mLayoutmanager.UpdateLayout()
        mLayoutmanager.AFlag = True
        mLayoutmanager.A = 1


    End Sub
    'Private Sub DelegateChangeWFApplylayout(index As Integer)
    '    Try

    '        mLayoutmanager.PaneRetainchanges()
    '        mLayoutmanager.RetainTitle()
    '        mLayoutmanager.clear()
    '        mLayoutmanager.CleartheLayoutdocumentvalues()

    '        'LMLayoutChanged(index)
    '        mLayoutmanager.clear()

    '        mpanelcontainer = Nothing
    '        mpanelcontainer = New LayoutSetting.panelcontainer

    '        AddControlsToPanel(index, mpanelcontainer, WindowsFormsHost, loadPluginsConfigXml.ApplicationHost.LayoutSettings)

    '        mLayoutmanager.PanelContent = mpanelcontainer
    '        _CurrentlayoutIdx = index

    '        mLayoutmanager.ApplyLayout(index)


    '        mLayoutmanager.UpdateLayout()
    '        mLayoutmanager.AFlag = True
    '        mLayoutmanager.A = 1
    '    Catch ex As Exception
    '        SetupErrorMsg(ex)
    '    End Try

    'End Sub
    Private Sub Submenu(ByRef Rib As RibbonMenuButton, ByVal SubMenuItem As ToolStripMenuItem)
        Dim ImgPath As String = String.Empty
        If SubMenuItem.DropDownItems.Count > 0 Then
            Dim RibMenuBtn As RibbonButton
            Dim image As BitmapImage
            Dim Irc As Integer = 0
            Dim Keytip As String = ""

            AddSubKeyTip(SubMenuItem)

            For i As Integer = 0 To SubMenuItem.DropDownItems.Count - 1

                If Not PluginMenu.Contains(SubMenuItem.DropDownItems(i)) Then
                    PluginMenu.Add(SubMenuItem.DropDownItems(i))

                End If
                RibMenuBtn = New RibbonButton
                image = Nothing
                image = New BitmapImage
                If Not String.IsNullOrEmpty(overrideAppHostConfig.OverrideValue.ResourceFolder.MenuIcon.iconPath) Then
                    ImgPath = overrideAppHostConfig.OverrideValue.ResourceFolder.MenuIcon.iconPath + "\" + SubMenuItem.DropDownItems.Item(i).Text + ".png"
                    If System.IO.File.Exists(ImgPath) = False Then
                        ImgPath = overrideAppHostConfig.OverrideValue.ResourceFolder.MenuIcon.iconPath + "\" + "Default.png"
                        If System.IO.File.Exists(ImgPath) = False Then
                            ImgPath = ""
                        End If
                    End If

                    If ImgPath <> "" Then
                        image.BeginInit()
                        image.UriSource = New Uri(ImgPath)
                        image.EndInit()
                    End If

                    'If IOManager.CheckFileExists(ImgPath) Then

                    '    image.BeginInit()
                    '    image.UriSource = New Uri(ImgPath)
                    '    image.EndInit()
                    'End If
                End If
                If DicSubMenuKeytip.ContainsKey(SubMenuItem.DropDownItems.Item(i).Text) Then
                    Keytip = DicSubMenuKeytip(SubMenuItem.DropDownItems.Item(i).Text)
                Else
                    Keytip = ""
                End If
                With RibMenuBtn
                    .SmallImageSource = image
                    .Label = SubMenuItem.DropDownItems.Item(i).Text
                    .KeyTip = Keytip
                    .Tag = SubMenuItem.DropDownItems.Item(i).Text
                End With
                If DirectCast(SubMenuItem.DropDownItems.Item(i), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys <> Keys.None Then
                    CreateRibbtnShortCut(SubMenuItem.DropDownItems.Item(i), DirectCast(SubMenuItem.DropDownItems.Item(i), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys)

                End If
                Rib.Items.Add(RibMenuBtn)
                AddHandler RibMenuBtn.Click, AddressOf CallClickEvent
            Next

        End If
    End Sub


    Private Sub CreateRibbtnShortCut(ByRef Menu As ToolStripMenuItem, ByRef SrtKey As System.Windows.Forms.Keys)
        Try

       
        Dim fkey As New RoutedCommand()
        Dim a As Integer = CInt(SrtKey)

        Dim keyCode As Key = CType((SrtKey And Keys.KeyCode), Key)
        Dim modifiers As Keys = CType((SrtKey And Keys.Modifiers), Keys)

        Dim wpfKey = KeyInterop.KeyFromVirtualKey(CInt(keyCode))

        If modifiers = Keys.Control Or Keys.Alt Then
            fkey.InputGestures.Add(New KeyGesture(wpfKey, ModifierKeys.Control Or ModifierKeys.Alt))
        End If

        If modifiers = Keys.Control Or Keys.Shift Then
            fkey.InputGestures.Add(New KeyGesture(wpfKey, ModifierKeys.Control Or ModifierKeys.Shift))
        End If
        If modifiers = Keys.Alt Or Keys.Shift Then
            fkey.InputGestures.Add(New KeyGesture(wpfKey, ModifierKeys.Alt Or ModifierKeys.Shift))
        End If

        If modifiers = Keys.Control Then
            fkey.InputGestures.Add(New KeyGesture(wpfKey, ModifierKeys.Control))
        End If
        'If modifiers = Keys.Control And Keys.Control Then
        '    fkey.InputGestures.Add(New KeyGesture(wpfKey, ModifierKeys.Control))
        'End If


        If modifiers = Keys.Alt Then
            fkey.InputGestures.Add(New KeyGesture(wpfKey, ModifierKeys.Alt))
        ElseIf modifiers = Keys.Shift Then
            fkey.InputGestures.Add(New KeyGesture(wpfKey, ModifierKeys.Shift))
        ElseIf modifiers = Keys.LWin Or modifiers = Keys.RWin Then
            fkey.InputGestures.Add(New KeyGesture(wpfKey, ModifierKeys.Windows))

        End If
        Dim acmdbind As CommandBinding = New CommandBinding(fkey, AddressOf Menu.PerformClick)
        If CommandBindings.Contains(acmdbind) = False Then
            CommandBindings.Add(acmdbind)
        End If
        Catch ex As Exception
            SetupErrorMsg(ex)
        End Try

    End Sub


    Private Sub CreatshortCutforMenu(ByVal Menu As String, ByVal bvisible As Boolean)
        Dim fkey As New RoutedCommand()

        Dim keyCode As Key
        Dim wpfKey
        Dim acmdbind As CommandBinding


        keyCode = Nothing
        wpfKey = Nothing

        If Menu.ToLower() = "close" Then
            Dim Keystring As List(Of MenuShortcutConfig) = overrideAppHostConfig.MenuShortcuts.Where(Function(x) Menu.Any(Function(y) "close" = x.MenuName.ToLower)).Select(Function(x) x).ToList()
            If Not Keystring Is Nothing And Keystring.Count > 0 AndAlso Keystring(0).key.Length = 2 Then
                Dim Ky As Object = KeyConvertion(Keystring(0).key.Chars(0))
                Dim Ky2 As Object = KeyConvertion(Keystring(0).key.Chars(1).ToString)

                keyCode = CType((Ky2 And Keys.KeyCode), Key)

                wpfKey = KeyInterop.KeyFromVirtualKey(CInt(keyCode))
                fkey.InputGestures.Add(New KeyGesture(wpfKey, Ky))

            Else
                keyCode = CType((Keys.W And Keys.KeyCode), Key)
                wpfKey = KeyInterop.KeyFromVirtualKey(CInt(keyCode))
                fkey.InputGestures.Add(New KeyGesture(wpfKey, ModifierKeys.Control))
            End If


            acmdbind = New CommandBinding(fkey, AddressOf sAHMenuCloseWork)

            If bvisible = True Then

                If CommandBindings.Contains(acmdbind) = False Then
                    CommandBindings.Add(acmdbind)
                End If
                Dim Obj = CommandBindings.Cast(Of CommandBinding).Select(Function(ic) ic.Command).Cast(Of RoutedCommand).Select(Function(r) r.InputGestures).ToList()
            Else
                If CommandBindings.Contains(acmdbind) Then
                    CommandBindings.Remove(acmdbind)
                End If
            End If


        ElseIf Menu.ToLower() = "open" Then

            Dim Keystring As List(Of MenuShortcutConfig) = overrideAppHostConfig.MenuShortcuts.Where(Function(x) Menu.Any(Function(y) "open" = x.MenuName.ToLower)).Select(Function(x) x).ToList()
            If Not Keystring Is Nothing And Keystring.Count > 0 AndAlso Keystring(0).key.Length = 2 Then
                Dim Ky As Object = KeyConvertion(Keystring(0).key.Chars(0))
                Dim Ky2 As Object = KeyConvertion(Keystring(0).key.Chars(1).ToString)

                keyCode = CType((Ky2 And Keys.KeyCode), Key)

                wpfKey = KeyInterop.KeyFromVirtualKey(CInt(keyCode))
                fkey.InputGestures.Add(New KeyGesture(wpfKey, Ky))

            Else
                keyCode = CType((Keys.O And Keys.KeyCode), Key)
                wpfKey = KeyInterop.KeyFromVirtualKey(CInt(keyCode))
                fkey.InputGestures.Add(New KeyGesture(wpfKey, ModifierKeys.Control))
            End If

            acmdbind = New CommandBinding(fkey, AddressOf sAHMenuOpenWork)

            If bvisible = True Then
                If CommandBindings.Contains(acmdbind) = False Then
                    CommandBindings.Add(acmdbind)
                End If
            Else
                If CommandBindings.Contains(acmdbind) Then
                    CommandBindings.Remove(acmdbind)
                End If
            End If

        ElseIf Menu.ToLower() = "next" Then

            Dim Keystring As List(Of MenuShortcutConfig) = overrideAppHostConfig.MenuShortcuts.Where(Function(x) Menu.Any(Function(y) "next" = x.MenuName.ToLower)).Select(Function(x) x).ToList()
            If Not Keystring Is Nothing And Keystring.Count > 0 AndAlso Keystring(0).key.Length = 2 Then
                Dim Ky As Object = KeyConvertion(Keystring(0).key.Chars(0))
                Dim Ky2 As Object = KeyConvertion(Keystring(0).key.Chars(1).ToString)

                keyCode = CType((Ky2 And Keys.KeyCode), Key)

                wpfKey = KeyInterop.KeyFromVirtualKey(CInt(keyCode))
                fkey.InputGestures.Add(New KeyGesture(wpfKey, Ky))

            Else
                keyCode = CType((Keys.N And Keys.KeyCode), Key)
                wpfKey = KeyInterop.KeyFromVirtualKey(CInt(keyCode))
                fkey.InputGestures.Add(New KeyGesture(wpfKey, ModifierKeys.Control))
            End If

            acmdbind = New CommandBinding(fkey, AddressOf sAHMenuNextWork)

            If bvisible = True Then
                If CommandBindings.Contains(acmdbind) = False Then
                    CommandBindings.Add(acmdbind)
                End If
            Else
                If CommandBindings.Contains(acmdbind) Then
                    CommandBindings.Remove(acmdbind)
                End If
            End If



        ElseIf Menu.ToLower() = "previous" Then


            Dim Keystring As List(Of MenuShortcutConfig) = overrideAppHostConfig.MenuShortcuts.Where(Function(x) Menu.Any(Function(y) "previous" = x.MenuName.ToLower)).Select(Function(x) x).ToList()
            If Not Keystring Is Nothing And Keystring.Count > 0 AndAlso Keystring(0).key.Length = 2 Then
                Dim Ky As Object = KeyConvertion(Keystring(0).key.Chars(0))
                Dim Ky2 As Object = KeyConvertion(Keystring(0).key.Chars(1).ToString)

                keyCode = CType((Ky2 And Keys.KeyCode), Key)

                wpfKey = KeyInterop.KeyFromVirtualKey(CInt(keyCode))
                fkey.InputGestures.Add(New KeyGesture(wpfKey, Ky))

            Else
                keyCode = CType((Keys.P And Keys.KeyCode), Key)
                wpfKey = KeyInterop.KeyFromVirtualKey(CInt(keyCode))
                fkey.InputGestures.Add(New KeyGesture(wpfKey, ModifierKeys.Control))
            End If

            acmdbind = New CommandBinding(fkey, AddressOf sAHMenuPreviousWork)

            If bvisible = True Then
                If CommandBindings.Contains(acmdbind) = False Then
                    CommandBindings.Add(acmdbind)
                End If
            Else
                If CommandBindings.Contains(acmdbind) Then
                    CommandBindings.Remove(acmdbind)
                End If
            End If

        ElseIf Menu.ToLower() = "exit" Then

            Dim Keystring As List(Of MenuShortcutConfig) = overrideAppHostConfig.MenuShortcuts.Where(Function(x) Menu.Any(Function(y) "exit" = x.MenuName.ToLower)).Select(Function(x) x).ToList()
            If Not Keystring Is Nothing And Keystring.Count > 0 AndAlso Keystring(0).key.Length = 2 Then
                Dim Ky As Object = KeyConvertion(Keystring(0).key.Chars(0))
                Dim Ky2 As Object = KeyConvertion(Keystring(0).key.Chars(1).ToString)

                keyCode = CType((Ky2 And Keys.KeyCode), Key)

                wpfKey = KeyInterop.KeyFromVirtualKey(CInt(keyCode))
                fkey.InputGestures.Add(New KeyGesture(wpfKey, Ky))

            Else
                keyCode = CType((Keys.X And Keys.KeyCode), Key)
                wpfKey = KeyInterop.KeyFromVirtualKey(CInt(keyCode))
                fkey.InputGestures.Add(New KeyGesture(wpfKey, ModifierKeys.Control))
            End If


            acmdbind = New CommandBinding(fkey, AddressOf sAHMenuExit)

            If bvisible = True Then
                If CommandBindings.Contains(acmdbind) = False Then
                    CommandBindings.Add(acmdbind)
                End If
            Else
                If CommandBindings.Contains(acmdbind) Then
                    CommandBindings.Remove(acmdbind)
                End If
            End If
        End If

        fkey = Nothing

        keyCode = Nothing
        wpfKey = Nothing
        acmdbind = Nothing
    End Sub

    Function KeyConvertion(ByVal StringToCheck As String) As Object

        Dim Keyvalues As System.Windows.Forms.Keys
        Dim sKey As System.Windows.Forms.Keys = Keys.None
        Try

            Dim chars As Char() = StringToCheck.ToCharArray()
            Dim KeyConverter As New System.Windows.Forms.KeysConverter
            For c = 0 To chars.Count - 1
                If chars(c) = "^" Then
                    Keyvalues = 2
                ElseIf chars(c) = "%" Then
                    Keyvalues = 1
                ElseIf chars(c) = "_" Then
                    Keyvalues = 4
                Else
                    sKey = KeyConverter.ConvertFrom(chars(c).ToString)
                    If KeyboardHook.KeyCodetable.Contains(sKey) = False Then
                        KeyboardHook.KeyCodetable.Add(sKey, Me)
                    End If
                End If
                If sKey <> Keys.None Then
                    Keyvalues = CType((Keyvalues Or sKey), System.Windows.Forms.Keys)
                End If
            Next


        Catch ex As Exception
            SetupErrorMsg(ex)
        Finally
            KeyConvertion = Keyvalues
        End Try

    End Function

    Public Sub sAHMenuOpenWork()
        Try
            ' Dispatcher.Invoke(CType(AddressOf AHMenuOpenWork, DelOpenWorkshorcut))
            AHMenuOpenWork()
            Menuhandle()
        Catch ex As Exception
            SetupErrorMsg(ex)
        Finally
            Prvs = False
            ShowWaitCursor(False)
        End Try
        '   AHMenuCloseWork(False)
    End Sub
    Public Sub sAHMenuCloseWork()
        Try

      
        ' Dispatcher.Invoke(CType(AddressOf AHMenuCloseWork, DelOpenWork), NotifyMessageObjectId, Sympraxis.Utilities.INormalPlugin.IPluginhost.GetObjectType.ObjectId, 0, "")
        AHMenuCloseWork(False)
         Catch ex As Exception
            SetupErrorMsg(ex)
        Finally
            Prvs = False
            ShowWaitCursor(False)
        End Try
    End Sub

    Public Sub sAHMenuNextWork()
        Try

        
        If objWCF.WorkitemCount > 1 Then

            If objWCF.CurrentWorkitemIdx < objWCF.WorkitemCount Then

                AHMenuNextWork(False)

                    Menuhandle()
                ElseIf objWCF.CurrentWorkitemIdx = objWCF.WorkitemCount Then
                    Throw New SympraxisException.InformationException("You are at the last work item")


            End If
            End If
        Catch ex As Exception
            SetupErrorMsg(ex)
        Finally
            Prvs = False
            ShowWaitCursor(False)
        End Try
    End Sub

    Public Sub sAHMenuPreviousWork()
        Try

        
        If objWCF.WorkitemCount > 1 Then
            If objWCF.CurrentWorkitemIdx <> 1 Then
                    AHMenuPreviousWork(False)
                    Menuhandle()
                Else
                    Throw New SympraxisException.InformationException("You are at the first work item")
            End If
            End If
        Catch ex As Exception
            SetupErrorMsg(ex)
        Finally
            Prvs = False
            ShowWaitCursor(False)
        End Try
    End Sub
    Public Sub sAHMenuExit()
        AHMenuExit(False)
    End Sub
    Private Function ChkKeyHash(ByVal Keytip As String, ByRef KeytipDic As Dictionary(Of String, String), ByRef HshtempVal As Hashtable) As Boolean
        Dim Rtn As Boolean = False



        If Not IsNothing(KeytipDic) AndAlso Keytip.Trim() <> String.Empty Then

            If HshtempVal.ContainsKey(Keytip.ToUpper()) = True Then
                Rtn = False
                ChkKeyHash = Rtn
                Exit Function
            End If

            If KeytipDic.ContainsValue(Keytip.ToUpper()) = True Then


                If Keytip.Length = 1 Then
                    HshtempVal.Add(Keytip.ToUpper(), Keytip.ToUpper())
                    Dim Key As String = KeytipDic.Where(Function(x) x.Value = Keytip).Select(Function(x) x.Key).FirstOrDefault()
                    Dim Keyval As String = Key
                    Dim Keyarray As String() = Keyval.Split(" ")
                    If Keyarray.Count > 1 Then
                        If Keyarray(1).Length > 0 AndAlso Keyarray(0).Length > 0 Then
                            KeytipDic(Key) = Keyarray(0).Substring(0, 1).ToUpper() + Keyarray(1).Substring(0, 1).ToUpper()
                        Else
                            KeytipDic(Key) = Keyarray(0).Substring(0, 2).ToUpper()
                        End If
                    Else
                        KeytipDic(Key) = Keyarray(0).Substring(0, 2).ToUpper()
                    End If

                End If


                Rtn = False
            Else
                Rtn = True
            End If

        End If


        ChkKeyHash = Rtn


    End Function

    Private Function CheckMainKeytip(ByVal FullKeytip As String, ByVal Keytipval As Integer) As Integer


        'Need 2 Check all Validation
        Dim irc As Integer
        Dim Keytip As String = ""

        If FullKeytip.Length >= Keytipval Then
            If Keytipval > 1 Then
                Dim Keyarray As String() = FullKeytip.Split(" ")
                If Keyarray.Count > 1 AndAlso Keytipval = 2 Then
                    Keytip = Keyarray(0).Substring(0, 1) + Keyarray(1).Substring(0, 1)
                Else
                    Dim Trimtxt As String = FullKeytip.Trim()

                    Keytip = String.Concat(Trimtxt.Substring(0, 1), Trimtxt.Substring(Keytipval - 1, 1))


                End If
            Else
                Keytip = FullKeytip.Substring(0, Keytipval)
            End If

            If DicKeytip.Count = 0 Then
                DicKeytip.Add(FullKeytip, Keytip.ToUpper())
            Else
                If ChkKeyHash(Keytip, DicKeytip, Hshtemp) Then

                    DicKeytip.Add(FullKeytip, Keytip.ToUpper())
                    Exit Function
                Else

                    Keytipval = Keytipval + 1
                    CheckMainKeytip(FullKeytip, Keytipval)
                End If
            End If
        End If




    End Function


    Private Function CheckSubKeytip(ByVal FullKeytip As String, ByVal Keytipval As Integer) As Integer
        'Need 2 Check all Validation
        Dim irc As Integer
        Dim Keytip As String = ""
        Try
            If FullKeytip.Length >= Keytipval Then
                If Keytipval > 1 Then
                    Dim Keyarray As String() = FullKeytip.Split(" ")
                    If Keyarray.Count > 1 AndAlso Keytipval = 2 Then
                        Keytip = Keyarray(0).Substring(0, 1) + Keyarray(1).Substring(0, 1)
                    Else
                        Dim Trimtxt As String = FullKeytip.Trim()

                        Keytip = String.Concat(Trimtxt.Substring(0, 1), Trimtxt.Substring(Keytipval - 1, 1))

                    End If
                Else
                    Keytip = FullKeytip.Substring(0, Keytipval)
                End If

                If DicKeytip.Count = 0 Then
                    DicSubMenuKeytip.Add(FullKeytip, Keytip.ToUpper())
                Else
                    If ChkKeyHash(Keytip, DicSubMenuKeytip, SubHshtemp) Then
                        DicSubMenuKeytip.Add(FullKeytip, Keytip.ToUpper())
                        Exit Function
                    Else

                        Keytipval = Keytipval + 1
                        CheckSubKeytip(FullKeytip, Keytipval)
                    End If
                End If
            End If


        Catch ex As Exception
            irc = 1
            IHErrException = ex

        Finally
            CheckSubKeytip = irc
        End Try
    End Function

    Private Sub AddKeyTip(ByRef Menu As Forms.ToolStripMenuItem)
        DicKeytip = New Dictionary(Of String, String)
        Hshtemp = New Hashtable
        For i As Integer = 0 To Menu.DropDownItems.Count - 1
            CheckMainKeytip(Menu.DropDownItems.Item(i).Text.TrimStart(), 1)
        Next
        Hshtemp = Nothing
    End Sub

    Private Sub AddSubKeyTip(ByRef Menu As Forms.ToolStripMenuItem)
        DicSubMenuKeytip = New Dictionary(Of String, String)
        SubHshtemp = New Hashtable
        For i As Integer = 0 To Menu.DropDownItems.Count - 1
            CheckSubKeytip(Menu.DropDownItems.Item(i).Text.TrimStart(), 1)
        Next
        SubHshtemp = Nothing
    End Sub

    Private Sub CreateMenu(ByRef Menu As Forms.ToolStripMenuItem)
        Dim RibnBtn As RibbonButton
        Dim RibGrp As RibbonGroup = Nothing
        Dim RibTab As RibbonTab
        Dim ImgPath As String = String.Empty
        Dim RibnMenuBtn As RibbonMenuButton
        Dim HeaderTabname As String = ""





        If Not String.IsNullOrEmpty(Menu.Text) AndAlso Menu.Text.Substring(0, 1) = "&" Then
            HeaderTabname = Menu.Text.Remove(0, 1)
        End If
        RibTab = New RibbonTab
        AddHandler Menu.EnabledChanged, AddressOf FormMenuHeader
        RibTab.Header = HeaderTabname

        RibGrp = Nothing
        RibGrp = New RibbonGroup
        Dim image As BitmapImage

        Dim Keytip As String = ""

        AddKeyTip(Menu)
        For i As Integer = 0 To Menu.DropDownItems.Count - 1
            'Dim Rib As New RibbonMenuButton
            'Rib.Items.Add(RibnBtn)
            AddHandler Menu.DropDownItems(i).EnabledChanged, AddressOf FormMenuenable
            If Not PluginMenu.Contains(Menu.DropDownItems(i)) Then
                PluginMenu.Add(Menu.DropDownItems(i))
            End If

            image = Nothing
            image = New BitmapImage
            If Not String.IsNullOrEmpty(overrideAppHostConfig.OverrideValue.ResourceFolder.MenuIcon.iconPath) Then
                ImgPath = overrideAppHostConfig.OverrideValue.ResourceFolder.MenuIcon.iconPath + "\" + Menu.DropDownItems.Item(i).Text.TrimStart() + ".png"
                If System.IO.File.Exists(ImgPath) = False Then
                    ImgPath = overrideAppHostConfig.OverrideValue.ResourceFolder.MenuIcon.iconPath + "\" + "Default.png"
                    If System.IO.File.Exists(ImgPath) = False Then
                        ImgPath = ""
                    End If
                End If

                If ImgPath <> "" Then
                    image.BeginInit()
                    image.UriSource = New Uri(ImgPath)
                    image.EndInit()
                End If
            End If
            If DirectCast(Menu.DropDownItems.Item(i), System.Windows.Forms.ToolStripMenuItem).DropDown.Items.Count = 0 Then
                RibnBtn = New RibbonButton
                If DicKeytip.ContainsKey(Menu.DropDownItems.Item(i).Text) Then
                    Keytip = DicKeytip(Menu.DropDownItems.Item(i).Text)
                Else
                    Keytip = ""
                End If
                With RibnBtn
                    .LargeImageSource = image
                    .Label = Menu.DropDownItems.Item(i).Text
                    .KeyTip = Keytip
                    .Tag = Menu.DropDownItems.Item(i).Text
                End With
                If DirectCast(Menu.DropDownItems.Item(i), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys <> Keys.None Then
                    CreateRibbtnShortCut(Menu.DropDownItems.Item(i), DirectCast(Menu.DropDownItems.Item(i), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys)

                End If
                RibGrp.Items.Add(RibnBtn)
                AddHandler RibnBtn.Click, AddressOf CallClickEvent




            ElseIf DirectCast(Menu.DropDownItems.Item(i), System.Windows.Forms.ToolStripMenuItem).DropDown.Items.Count > 0 Then
                RibnMenuBtn = New RibbonMenuButton
                If DicKeytip.ContainsKey(Menu.DropDownItems.Item(i).Text) Then
                    Keytip = DicKeytip(Menu.DropDownItems.Item(i).Text)
                Else
                    Keytip = ""
                End If
                With RibnMenuBtn
                    .LargeImageSource = image
                    .Label = Menu.DropDownItems.Item(i).Text
                    .KeyTip = Keytip
                    .Tag = Menu.DropDownItems.Item(i).Text
                End With
                If DirectCast(Menu.DropDownItems.Item(i), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys <> Keys.None Then
                    CreateRibbtnShortCut(Menu.DropDownItems.Item(i), DirectCast(Menu.DropDownItems.Item(i), System.Windows.Forms.ToolStripMenuItem).ShortcutKeys)

                End If
                Submenu(RibnMenuBtn, Menu.DropDownItems.Item(i))
                RibGrp.Items.Add(RibnMenuBtn)
            End If

            'If Menu.DropDownItems.Item(i) Then

            'End If





        Next
        RibTab.Items.Add(RibGrp)
        RibTab.KeyTip = HeaderTabname.Substring(0, 1)
        RibbonWin.Items.Add(RibTab)
        RibbonWin.UpdateLayout()
        Me.UpdateLayout()







    End Sub
    Private Sub ToggleChange(ByVal isonline As Boolean)
        Dim image As BitmapImage
        Dim ImgPath As String = String.Empty

        For Each RibTab As RibbonTab In RibbonWin.Items

            If Not IsNothing(RibTab) Then
                If RibTab.Header.ToString().ToLower() = "home" Then
                    For Each sRibGroup As RibbonGroup In RibTab.Items

                        For Each ToggleButton As Object In sRibGroup.Items
                            Dim aToggleButton As RibbonToggleButton = TryCast(ToggleButton, RibbonToggleButton)
                            If Not IsNothing(aToggleButton) Then
                                image = Nothing
                                image = New BitmapImage
                                If Not String.IsNullOrEmpty(overrideAppHostConfig.OverrideValue.ResourceFolder.MenuIcon.iconPath) Then
                                    If isonline = True Then
                                        ImgPath = overrideAppHostConfig.OverrideValue.ResourceFolder.MenuIcon.iconPath + "\Online.png"

                                    Else
                                        ImgPath = overrideAppHostConfig.OverrideValue.ResourceFolder.MenuIcon.iconPath + "\Offline.png"

                                    End If
                                    If System.IO.File.Exists(ImgPath) = False Then
                                        ImgPath = overrideAppHostConfig.OverrideValue.ResourceFolder.MenuIcon.iconPath + "\" + "Default.png"
                                        If System.IO.File.Exists(ImgPath) = False Then
                                            ImgPath = ""
                                        End If
                                    End If

                                    If ImgPath <> "" Then
                                        image.BeginInit()
                                        image.UriSource = New Uri(ImgPath)
                                        image.EndInit()
                                    End If

                                End If
                                If isonline = True Then
                                    With aToggleButton
                                        .LargeImageSource = image
                                        .Label = "Online"
                                        .Tag = "Online"
                                    End With
                                Else

                                    With aToggleButton
                                        .LargeImageSource = image
                                        .Label = "Offline"
                                        .Tag = "Offline"
                                    End With

                                End If

                            End If
                        Next
                    Next
                End If

            End If

        Next

    End Sub
    Private Sub CreateMenu(ByRef HeaderTabname As String, ByRef Menuname As String)

        Dim RibnBtn As RibbonButton
        Dim RibGrp As RibbonGroup = Nothing
        Dim RibTab As RibbonTab
        Dim ImgPath As String = String.Empty

        If Not String.IsNullOrEmpty(HeaderTabname) AndAlso HeaderTabname.Substring(0, 1) = "&" Then
            HeaderTabname = HeaderTabname.Remove(0, 1)
        End If
        RibTab = New RibbonTab

        RibTab.Header = HeaderTabname



        Dim Createmenu As String() = Menuname.Split("|")

        RibGrp = Nothing
        RibGrp = New RibbonGroup
        Dim image As BitmapImage
        For i As Integer = 0 To Createmenu.Count - 1
            'Dim Rib As New RibbonMenuButton
            'Rib.Items.Add(RibnBtn)
            image = Nothing
            image = New BitmapImage
            If Not String.IsNullOrEmpty(overrideAppHostConfig.OverrideValue.ResourceFolder.MenuIcon.iconPath) Then

                If Createmenu(i) = "Offline" Then
                    If configMgr.AppbaseConfig.EnvironmentVariables.IsOnline = True Then
                        ImgPath = overrideAppHostConfig.OverrideValue.ResourceFolder.MenuIcon.iconPath + "\Online.png"
                    Else
                        ImgPath = overrideAppHostConfig.OverrideValue.ResourceFolder.MenuIcon.iconPath + "\Offline.png"
                    End If
                Else
                    ImgPath = overrideAppHostConfig.OverrideValue.ResourceFolder.MenuIcon.iconPath + "\" + Createmenu(i) + ".png"

                End If
                If System.IO.File.Exists(ImgPath) = False Then
                    ImgPath = overrideAppHostConfig.OverrideValue.ResourceFolder.MenuIcon.iconPath + "\" + "Default.png"
                    If System.IO.File.Exists(ImgPath) = False Then
                        ImgPath = ""
                    End If
                End If

                If ImgPath <> "" Then
                    image.BeginInit()
                    image.UriSource = New Uri(ImgPath)
                    image.EndInit()
                End If
            End If
            If Createmenu(i) = "Offline" Then
                Dim toggleButton As New RibbonToggleButton
                If configMgr.AppbaseConfig.EnvironmentVariables.IsOnline = False Then

                    With toggleButton
                        .LargeImageSource = image
                        .Label = "Offline"
                        .Tag = "Offline"
                    End With
                Else

                    With toggleButton
                        .LargeImageSource = image
                        .Label = "Online"
                        .Tag = "Online"
                    End With
                End If


                AddHandler toggleButton.Click, AddressOf ToggleClickEvent
                RibGrp.Items.Add(toggleButton)

            Else
                RibnBtn = New RibbonButton

                With RibnBtn
                    .LargeImageSource = image
                    .Label = Createmenu(i)
                    .KeyTip = Createmenu(i).TrimStart().Substring(0, 1)
                    .Tag = Createmenu(i)
                End With

                AddHandler RibnBtn.Click, AddressOf CallClickEvent
                RibGrp.Items.Add(RibnBtn)
            End If



        Next
        RibTab.Items.Add(RibGrp)
        RibTab.KeyTip = HeaderTabname.Substring(0, 1)
        RibbonWin.Items.Add(RibTab)
        RibbonWin.UpdateLayout()
        Me.UpdateLayout()



    End Sub
    Private Sub CreateRibbonHelpContentMenu()
        If overrideAppHostConfig.MenuConfigManifest.Count > 0 Then
            For Each item In overrideAppHostConfig.MenuConfigManifest
                Dim ProcessName() As String
                ProcessName = item.ProcessName.ToString.Split("|")
                If ProcessName.Count > 0 Then
                    For Each PItem In ProcessName
                        If PItem.ToString = intrHostProcessName Then
                            If Not IsNothing(item.Menu) Then


                            Dim Menu() As String
                            Menu = item.Menu.ToString.Split("|")
                            If Menu.Count > 0 Then
                                For iMenu = 0 To Menu.Count - 1
                                    If Menu(iMenu).ToString.ToUpper = "OpenNext".ToUpper Then
                                        btnOpenNext.Visibility = System.Windows.Visibility.Visible
                                    End If
                                    If Menu(iMenu).ToString.ToUpper = "Readonly".ToUpper Then
                                        btnReadnly.Visibility = System.Windows.Visibility.Visible
                                    End If
                                Next
                            End If
                        End If
                            If Not IsNothing(item.HMenu) AndAlso item.HMenu.Trim.Length > 0 Then
                                AHEnableDisableMenu("Home", item.HMenu, False)
                            End If


                        End If
                    Next
                End If

            Next
        End If
    End Sub


   
    Private Sub ShowlastEvent()
        objWCF.ShowLastEvents()
    End Sub
    Private isIntructionShown As Boolean = False
    Private Sub LoadWorkInstructions()

        If Not configMgr Is Nothing Then
            isIntructionShown = False
            Logger.WriteToLogFile("Ihost", "ResetInstruction Starts")
            ResetInstruction()

            Logger.WriteToLogFile("Ihost", "ResetInstruction Ends")

            If WIConfig IsNot Nothing AndAlso WIConfig.DynamicInstruction IsNot Nothing AndAlso WIConfig.DynamicInstruction.DWIns.Count > 0 Then
                Logger.WriteToLogFile("Ihost", "SetDynamicInstruction Starts")
                SetDynamicInstruction(WIConfig.DynamicInstruction)

                Logger.WriteToLogFile("Ihost", "SetDynamicInstruction Ends")

                Logger.WriteToLogFile("Ihost", "LoadDynamicInstruction Starts")
                LoadDynamicInstruction()

                Logger.WriteToLogFile("Ihost", "LoadDynamicInstruction Ends")
            End If

            If WIConfig IsNot Nothing AndAlso WIConfig.StandardInstruction IsNot Nothing Then
                Logger.WriteToLogFile("Ihost", "SetStandardInstrction Starts")
                SetStandardInstrction()

                Logger.WriteToLogFile("Ihost", "SetStandardInstrction Ends")

                Logger.WriteToLogFile("Ihost", "LoadStandardInstrction Starts")
                LoadStandardInstrction()

                Logger.WriteToLogFile("Ihost", "LoadStandardInstrction Ends")
            End If






            If WIConfig IsNot Nothing AndAlso WIConfig.ContextInstruction IsNot Nothing AndAlso WIConfig.ContextInstruction.CWIns.Count > 0 Then
                isIntructionShown = True
                CIndicate.ToolTip = String.Concat("Context (", WIConfig.ContextInstruction.iSKey, ")")
                CIStyle.ToolTip = CIndicate.ToolTip.ToString
                CIBorder.Background = New SolidColorBrush(ColorConverter.ConvertFromString(WIConfig.ContextInstruction.iBackColor))
                'SetContextInstruction("SYS_CURRENTFIELD", "DLN", irc, Errmsg)
            End If
            If isIntructionShown = False Then
                AHEnableDisableMenu("Help", "WorkInstruction", False)

            Else
                AHEnableDisableMenu("Help", "WorkInstruction", True)
            End If
        End If


    End Sub

    Private Sub ResetInstruction()

        objDInstruction = New Instructions
        If objDInstruction.INS Is Nothing Then
            objDInstruction.INS = New List(Of INS)
        End If
        objSInstruction = New Instructions
        If objSInstruction.INS Is Nothing Then
            objSInstruction.INS = New List(Of INS)
        End If
        objCInstruction = New Instructions
        If objCInstruction.INS Is Nothing Then
            objCInstruction.INS = New List(Of INS)
        End If
        DIPin = False
        SIPin = False
        CIPin = False
        pinDI.Visibility = System.Windows.Visibility.Collapsed
        unpinDI.Visibility = System.Windows.Visibility.Collapsed
        pinSI.Visibility = System.Windows.Visibility.Collapsed
        unpinSI.Visibility = System.Windows.Visibility.Collapsed
        pinCI.Visibility = System.Windows.Visibility.Collapsed
        unpinCI.Visibility = System.Windows.Visibility.Collapsed
        pnlIndicator.IsOpen = False
        pnlContextInstruction.IsOpen = False
        pnlDynamicInstruction.IsOpen = False
        pnlStadInstruction.IsOpen = False
        If CITimer IsNot Nothing Then
            If CITimer IsNot Nothing Then
                CITimer.Enabled = False
                CITimer.Stop()
            End If
            If CICloseTimer IsNot Nothing Then
                CICloseTimer.Enabled = False
                CICloseTimer.Stop()
            End If
        End If
        If SITimer IsNot Nothing Then
            If SITimer IsNot Nothing Then
                SITimer.Enabled = False
                SITimer.Stop()
            End If
            If SICloseTimer IsNot Nothing Then
                SICloseTimer.Enabled = False
                SICloseTimer.Stop()
            End If
        End If
        If DITimer IsNot Nothing Then
            If DITimer IsNot Nothing Then
                DITimer.Enabled = False
                DITimer.Stop()
            End If
            If DICloseTimer IsNot Nothing Then
                DICloseTimer.Enabled = False
                DICloseTimer.Stop()
            End If
        End If
        ActiveWI = False


    End Sub

    Private Sub SetContextInstruction(ByVal Daxkey As String, ByVal DaxValue As String)
        Dim Irc As Integer = 0
        Dim Errmsg As String = String.Empty
        If WIConfig IsNot Nothing AndAlso WIConfig.ContextInstruction IsNot Nothing AndAlso WIConfig.ContextInstruction.CWIns.Count > 0 Then
            Dim cins As ContextSetting = WIConfig.ContextInstruction
            Dim LstRulelist As List(Of Rule) = Nothing
            Dim MatchedRule As List(Of Rule) = Nothing
            Dim CrRule As List(Of Rule) = Nothing
            Dim lstDWI As New List(Of ContextWorkIns)
            CIndicate.ToolTip = String.Concat("Context (", WIConfig.ContextInstruction.iSKey, ")")
            CIStyle.ToolTip = CIndicate.ToolTip.ToString
            CIBorder.Background = New SolidColorBrush(ColorConverter.ConvertFromString(WIConfig.ContextInstruction.iBackColor))
            If Not cins Is Nothing AndAlso cins.CWIns.Count > 0 Then
                objCInstruction.INS = New List(Of INS)
                LstRulelist = cins.Rules.Ruleslist.ToList()
                If Not LstRulelist Is Nothing AndAlso LstRulelist.Count = 0 Then
                    lstDWI = cins.CWIns.ToList()
                    If Not lstDWI Is Nothing Then
                        AddContextInstruction(lstDWI)
                    End If
                Else
                    GetCrRulelist(CrRule, LstRulelist, LstRulelist)
                    'Dim name = CrRule.Item(0).Name
                    objWCF.GetMatchRule(MatchedRule, LstRulelist, CrRule, Irc, Errmsg)
                    If Irc <> 0 Then

                        Throw New SympraxisException.AppException(Errmsg)
                    End If
                    If Not MatchedRule Is Nothing Then
                        If MatchedRule.Count > 0 Then
                            Dim matchlist As List(Of ContextWorkIns) = cins.CWIns.Where(Function(x) MatchedRule.Any(Function(y) y.Name = x.cName AndAlso x.cKey = Daxkey AndAlso x.Context = DaxValue)).Select(Function(x) x).ToList()
                            Dim UnNamelist As List(Of ContextWorkIns) = cins.CWIns.Where(Function(x) x.cName = String.Empty AndAlso x.cKey = Daxkey AndAlso x.Context = DaxValue).Select(Function(x) x).ToList()
                            If matchlist IsNot Nothing Then
                                If UnNamelist IsNot Nothing AndAlso UnNamelist.Count > 0 Then
                                    For Each unlist In UnNamelist
                                        matchlist.Add(unlist)
                                    Next
                                End If
                                If matchlist.Count > 0 Then
                                    'For Each listitem In matchlist
                                    AddContextInstruction(matchlist)
                                    'Next
                                End If
                            End If
                        End If
                    End If
                End If
                If objCInstruction.INS.Count > 0 Then
                    LoadContextInstrction()
                End If
            End If
        End If

    End Sub

    Private Sub SetInstructionLocation(ByRef pnlinst As Controls.Primitives.Popup)

        ActiveWI = True
        'If HostActive = True Then
        SetIndicator()
        If ShowCI = True And ShowDI = True And ShowSI = True Then
            pnlContextInstruction.PlacementRectangle = New Rect(Me.ActualWidth - 1110, Me.ActualHeight - 80, 360, 170)
            pnlStadInstruction.PlacementRectangle = New Rect(Me.ActualWidth - 740, Me.ActualHeight - 80, 360, 170)
            pnlDynamicInstruction.PlacementRectangle = New Rect(Me.ActualWidth - 370, Me.ActualHeight - 80, 360, 170)
            pnlContextInstruction.IsOpen = True
            pnlStadInstruction.IsOpen = True
            pnlDynamicInstruction.IsOpen = True
        ElseIf ShowDI = True And ShowSI = True Then
            pnlStadInstruction.PlacementRectangle = New Rect(Me.ActualWidth - 740, Me.ActualHeight - 80, 360, 170)
            pnlDynamicInstruction.PlacementRectangle = New Rect(Me.ActualWidth - 370, Me.ActualHeight - 80, 360, 170)
            pnlStadInstruction.IsOpen = True
            pnlDynamicInstruction.IsOpen = True
        ElseIf ShowDI = True And ShowCI = True Then
            pnlContextInstruction.PlacementRectangle = New Rect(Me.ActualWidth - 740, Me.ActualHeight - 80, 360, 170)
            pnlDynamicInstruction.PlacementRectangle = New Rect(Me.ActualWidth - 370, Me.ActualHeight - 80, 360, 170)
            pnlContextInstruction.IsOpen = True
            pnlDynamicInstruction.IsOpen = True
        ElseIf ShowSI = True And ShowCI = True Then
            pnlContextInstruction.PlacementRectangle = New Rect(Me.ActualWidth - 740, Me.ActualHeight - 80, 360, 170)
            pnlStadInstruction.PlacementRectangle = New Rect(Me.ActualWidth - 370, Me.ActualHeight - 80, 360, 170)
            pnlContextInstruction.IsOpen = True
            pnlStadInstruction.IsOpen = True
        Else
            If pnlinst IsNot Nothing Then
                pnlinst.PlacementRectangle = New Rect(Me.ActualWidth - 370, Me.ActualHeight - 80, 360, 170)
                pnlinst.IsOpen = True
            Else
                If ShowCI = True Then
                    pnlContextInstruction.PlacementRectangle = New Rect(Me.ActualWidth - 370, Me.ActualHeight - 80, 360, 170)
                    pnlContextInstruction.IsOpen = True
                End If
                If ShowDI = True Then
                    pnlDynamicInstruction.PlacementRectangle = New Rect(Me.ActualWidth - 370, Me.ActualHeight - 80, 360, 170)
                    pnlDynamicInstruction.IsOpen = True
                End If
                If ShowSI = True Then
                    pnlStadInstruction.PlacementRectangle = New Rect(Me.ActualWidth - 370, Me.ActualHeight - 80, 360, 170)
                    pnlStadInstruction.IsOpen = True
                End If
            End If
        End If
        'End If

    End Sub
    Private Sub LoadContextInstrction()

        If objCInstruction.INS IsNot Nothing AndAlso objCInstruction.INS.Count > 0 Then
            If CITimer IsNot Nothing Then
                CITimer.Enabled = False
                CITimer.Stop()
            End If
            If CICloseTimer IsNot Nothing Then
                CICloseTimer.Enabled = False
                CICloseTimer.Stop()
            End If
            CIStyle.Opacity = 1
            ShowCI = True
            CIMinimize = False
            SetInstructionLocation(pnlContextInstruction)
            pnlContextInstruction.UpdateLayout()

            If WIConfig.ContextInstruction.iBackColor <> String.Empty Then
                CIStyle.Background = New SolidColorBrush(ColorConverter.ConvertFromString(WIConfig.ContextInstruction.iBackColor))
            Else
                CIStyle.Background = New SolidColorBrush(ColorConverter.ConvertFromString("#FFF5EB80"))
            End If

            CIDataGrid.ItemsSource = objCInstruction.INS
            CIDataGrid.UpdateLayout()

            If CIPin = False AndAlso WIConfig.ContextInstruction.iInterval > 0 Then
                unpinCI.Visibility = System.Windows.Visibility.Visible
                Dim WIInterval As Integer = 1000 * WIConfig.ContextInstruction.iInterval
                CITimer = New Timers.Timer(WIInterval)
                AddHandler CITimer.Elapsed, AddressOf CITimer_Elapsed
                CICloseTimer = New Timers.Timer(200)
                AddHandler CICloseTimer.Elapsed, AddressOf CICloseTimer_Elapsed
                CITimer.Enabled = True
            End If


        End If

    End Sub

    Private Sub AddContextInstruction(ByVal dins As List(Of ContextWorkIns))

        For Each ln In dins
            Dim objHeaderIns As New INS
            Dim objBodyIns As New INS
            objHeaderIns.LN = ln.cHeader
            objHeaderIns.Fcolor = WIConfig.ContextInstruction.HeaderColor
            objHeaderIns.BColor = WIConfig.ContextInstruction.iBackColor
            objHeaderIns.FSize = 16

            objBodyIns.LN = ln.cBody
            objBodyIns.Fcolor = WIConfig.ContextInstruction.iForeColor
            objBodyIns.BColor = WIConfig.ContextInstruction.iBackColor
            objBodyIns.FSize = 14
            objCInstruction.INS.Add(objHeaderIns)
            objCInstruction.INS.Add(objBodyIns)
        Next


    End Sub


    Private Sub SetStandardInstrction()
        Dim Irc As Integer = 1
        Dim Errmsg As String = String.Empty
        Dim XMLWIDP = New Dictionary(Of String, String), InsNode As String = "XINSTRU"
        XMLWIDP.Add(InsNode, Nothing)
        Irc = objWCF.Read(XMLWIDP, Errmsg)
        If Irc <> 0 Then
            Throw New SympraxisException.WorkitemException(Errmsg)
        End If
        If XMLWIDP(InsNode) IsNot Nothing AndAlso XMLWIDP(InsNode) <> "" Then
            Dim icnt As Integer = 0
            Dim xInsdoc As New XmlDocument
            xInsdoc.LoadXml(XMLWIDP(InsNode))
            SIndicate.ToolTip = String.Concat("Standard (", WIConfig.StandardInstruction.iSKey, ")")
            SIStyle.ToolTip = SIndicate.ToolTip.ToString
            SIBorder.Background = New SolidColorBrush(ColorConverter.ConvertFromString(WIConfig.StandardInstruction.iBackColor))
            Dim Tnoxe As Xml.XmlNode = xInsdoc.SelectSingleNode("//Instructions")
            If Tnoxe IsNot Nothing Then
                If Tnoxe.SelectSingleNode("//IN") IsNot Nothing Then
                    'DgXmlIns.Rows.Clear()
                    Dim xDoc As XDocument = New XDocument()
                    xDoc = XDocument.Load(New XmlNodeReader(Tnoxe))
                    Dim nodeList = xDoc.Descendants("IN").OrderByDescending(Function(x) x.Attribute("ID").Value).ToList()
                    If nodeList.Count > 0 Then
                        For Each xnode In nodeList
                            If xnode.Attribute("TP").Value IsNot String.Empty Then
                                Dim targetprocess() As String = xnode.Attribute("TP").Value.Split("|")

                                For Each tp As String In targetprocess
                                    If tp = intrHostProcessName Or tp = "ALL" Then
                                        icnt = icnt + 1
                                        Dim header As String = xnode.Attribute("HD").Value
                                        Dim sval As String = xnode.Attribute("TX").Value
                                        Dim DT As String = xnode.Attribute("DT").Value

                                        Dim objHeaderIns As New INS
                                        Dim objBodyIns As New INS
                                        objHeaderIns.LN = xnode.Attribute("HD").Value
                                        objHeaderIns.Fcolor = WIConfig.StandardInstruction.HeaderColor
                                        objHeaderIns.BColor = WIConfig.StandardInstruction.iBackColor
                                        objHeaderIns.FSize = 16
                                        'objBodyIns.LN = DT & " " & sval.Replace("\n", vbCrLf).Replace("\t", vbTab) & vbCrLf
                                        objBodyIns.LN = String.Concat(DT, " ", sval.Replace("\n", vbCrLf).Replace("\t", vbTab))
                                        objBodyIns.Fcolor = WIConfig.StandardInstruction.iForeColor
                                        objBodyIns.BColor = WIConfig.StandardInstruction.iBackColor
                                        objBodyIns.FSize = 14
                                        If objSInstruction.INS Is Nothing Then
                                            objSInstruction.INS = New List(Of INS)
                                        End If
                                        objSInstruction.INS.Add(objHeaderIns)
                                        objSInstruction.INS.Add(objBodyIns)
                                    End If
                                Next
                            End If
                        Next
                    End If
                End If
            End If
        End If

    End Sub


    Private Sub LoadStandardInstrction()

        If objSInstruction.INS IsNot Nothing AndAlso objSInstruction.INS.Count > 0 Then
            isIntructionShown = True
            If SITimer IsNot Nothing Then
                SITimer.Enabled = False
                SITimer.Stop()
            End If
            If SICloseTimer IsNot Nothing Then
                SICloseTimer.Enabled = False
                SICloseTimer.Stop()
            End If
            SIStyle.Opacity = 1
            ShowSI = True
            SIMinimize = False
            SetInstructionLocation(pnlStadInstruction)
            'pnlStadInstruction.IsOpen = True
            'pnlStadInstruction.Visibility = System.Windows.Visibility.Visible
            pnlStadInstruction.UpdateLayout()

            If WIConfig.StandardInstruction.iBackColor <> String.Empty Then
                SIStyle.Background = New SolidColorBrush(ColorConverter.ConvertFromString(WIConfig.StandardInstruction.iBackColor))
            Else
                SIStyle.Background = New SolidColorBrush(ColorConverter.ConvertFromString("#FFF5EB80"))
            End If
            SIDataGrid.ItemsSource = objSInstruction.INS
            SIDataGrid.UpdateLayout()
            If SIPin = False AndAlso WIConfig.StandardInstruction.iInterval > 0 Then
                unpinSI.Visibility = System.Windows.Visibility.Visible
                Dim WIInterval As Integer = 1000 * WIConfig.StandardInstruction.iInterval
                SITimer = New Timers.Timer(WIInterval)
                AddHandler SITimer.Elapsed, AddressOf SITimer_Elapsed
                SICloseTimer = New Timers.Timer(200)
                AddHandler SICloseTimer.Elapsed, AddressOf SICloseTimer_Elapsed
                SITimer.Enabled = True
            End If
        End If

    End Sub


    Private Sub AddDynamicInstruction(ByVal dins As List(Of DyanamicWorkIns))

        For Each ln In dins
            Dim objHeaderIns As New INS
            Dim objBodyIns As New INS
            objHeaderIns.LN = ln.dHeader
            objHeaderIns.Fcolor = WIConfig.DynamicInstruction.HeaderColor
            objHeaderIns.BColor = WIConfig.DynamicInstruction.iBackColor
            objHeaderIns.FSize = 16

            objBodyIns.LN = ln.dBody
            objBodyIns.Fcolor = WIConfig.DynamicInstruction.iForeColor
            objBodyIns.BColor = WIConfig.DynamicInstruction.iBackColor
            objBodyIns.FSize = 14
            If objDInstruction.INS Is Nothing Then
                objDInstruction.INS = New List(Of INS)
            End If
            objDInstruction.INS.Add(objHeaderIns)
            objDInstruction.INS.Add(objBodyIns)
        Next

    End Sub

    Private Sub SetDynamicInstruction(ByVal dins As DynamicSetting)

        Dim LstRulelist As List(Of Rule) = Nothing
        Dim MatchedRule As List(Of Rule) = Nothing
        Dim CrRule As List(Of Rule) = Nothing
        Dim lstDWI As New List(Of DyanamicWorkIns)
        Dim Irc As Integer = 0
        Dim Errmsg As String = String.Empty


        DIndicate.ToolTip = String.Concat("Dynamic (", WIConfig.DynamicInstruction.iSKey, ")")
        DIStyle.ToolTip = DIndicate.ToolTip.ToString
        DIBorder.Background = New SolidColorBrush(ColorConverter.ConvertFromString(WIConfig.DynamicInstruction.iBackColor))
        If Not dins Is Nothing AndAlso dins.DWIns.Count > 0 Then
            LstRulelist = dins.Rules.Ruleslist.ToList()
            If Not LstRulelist Is Nothing AndAlso LstRulelist.Count = 0 Then
                lstDWI = dins.DWIns.ToList()
                If Not lstDWI Is Nothing Then
                    AddDynamicInstruction(lstDWI)
                End If
            Else
                GetCrRulelist(CrRule, LstRulelist, LstRulelist)
                'Dim name = CrRule.Item(0).Name
                objWCF.GetMatchRule(MatchedRule, LstRulelist, CrRule, Irc, Errmsg)
                If Irc <> 0 Then
                    Throw New SympraxisException.AppException(Errmsg)
                End If
                'To do 19/12/19 objWCF.GetMatchRule habe to handle

                If Not MatchedRule Is Nothing Then
                    If MatchedRule.Count > 0 Then
                        Dim matchlist As List(Of DyanamicWorkIns) = dins.DWIns.Where(Function(x) MatchedRule.Any(Function(y) y.Name = x.dName)).Select(Function(x) x).ToList()
                        If matchlist IsNot Nothing AndAlso matchlist.Count > 0 Then
                            'For Each listitem In matchlist
                            AddDynamicInstruction(matchlist)
                            'Next
                        End If
                    End If

                End If
            End If
        End If

    End Sub

    Private Sub LoadDynamicInstruction()

        If objDInstruction.INS IsNot Nothing AndAlso objDInstruction.INS.Count > 0 Then
            isIntructionShown = True
            If DITimer IsNot Nothing Then
                DITimer.Enabled = False
                DITimer.Stop()
            End If
            If DICloseTimer IsNot Nothing Then
                DICloseTimer.Enabled = False
                DICloseTimer.Stop()
            End If
            DIStyle.Opacity = 1
            ShowDI = True
            DIMinimize = False
            SetInstructionLocation(pnlDynamicInstruction)
            pnlDynamicInstruction.UpdateLayout()
            If WIConfig.DynamicInstruction.iBackColor <> String.Empty Then
                DIStyle.Background = New SolidColorBrush(ColorConverter.ConvertFromString(WIConfig.DynamicInstruction.iBackColor))
            Else
                DIStyle.Background = New SolidColorBrush(ColorConverter.ConvertFromString("#FFF5EB80"))
            End If
            DIDataGrid.ItemsSource = objDInstruction.INS
            DIDataGrid.UpdateLayout()
            If DIPin = False AndAlso WIConfig.DynamicInstruction.iInterval > 0 Then
                unpinDI.Visibility = System.Windows.Visibility.Visible
                Dim WIInterval As Integer = 1000 * WIConfig.DynamicInstruction.iInterval
                'If DITimer Is Nothing Then
                DITimer = New Timers.Timer(WIInterval)
                AddHandler DITimer.Elapsed, AddressOf DITimer_Elapsed
                'End If
                'If DICloseTimer Is Nothing Then
                DICloseTimer = New Timers.Timer(200)
                AddHandler DICloseTimer.Elapsed, AddressOf DICloseTimer_Elapsed
                'End If
                DITimer.Enabled = True
            End If
        End If

    End Sub


#Region "ManagerInstance"
    Dim channel As TcpServerChannel
    Dim GetIMPort As Integer = 9999
    Dim bIsPortExist As Boolean = False
    Dim lease As ILease
    Dim WithEvents ObjManagerInstanceMgr As InstanceManager = Nothing
    Dim ListAddAppKeys As New List(Of InstanceManager.AddAppKeys)
    Dim IMURL As String = ""
    Private Sub InitIMServercommunication()
        Try
            Dim bIsPortExist As Boolean = False
            Dim IncPort As Integer = 0
            Dim MinPort As Integer = 0
            Dim Maxport As Integer = 0
            Dim IMErrorMsg As String = ""
            Dim IMStackTraceMsg As String = ""
            Dim iRc As Integer = 0
            If Not IsNothing(appHostConfigXml.InstanceMgr) AndAlso Not IsNothing(appHostConfigXml.InstanceMgr.StartPort) AndAlso appHostConfigXml.InstanceMgr.StartPort <> "" AndAlso Not IsNothing(appHostConfigXml.InstanceMgr.EndPort) AndAlso appHostConfigXml.InstanceMgr.EndPort <> "" AndAlso Not IsNothing(appHostConfigXml.InstanceMgr.ExternalInstance) AndAlso appHostConfigXml.InstanceMgr.ExternalInstance <> "" Then
                MinPort = appHostConfigXml.InstanceMgr.StartPort
                Maxport = appHostConfigXml.InstanceMgr.EndPort
                Logger.WriteToLogFile("InHost:", "InitIMServercommunication st")
                For i As Integer = MinPort To Maxport
                    IMErrorMsg = ""
                    IMStackTraceMsg = ""
                    Try
                        If Not bIsPortExist Then
                            GetIMPort = MinPort + IncPort
                            IMURL = "tcp://localhost:IMPort/InstanceManager"
                            IMURL = Replace(IMURL, "IMPort", GetIMPort.ToString)
                            IncPort = IncPort + 1
                            ' MsgBox("Port:" & myurl)
                            If channel IsNot Nothing Then
                                channel.StopListening(Nothing)
                            End If
                            Dim obj_FormatProvider As New BinaryServerFormatterSinkProvider()
                            obj_FormatProvider.TypeFilterLevel = Runtime.Serialization.Formatters.TypeFilterLevel.Full
                            channel = New TcpServerChannel("Server", GetIMPort, obj_FormatProvider)
                            ChannelServices.RegisterChannel(channel, False)
                            RemotingConfiguration.RegisterWellKnownServiceType(GetType(InstanceManager), "InstanceManager", WellKnownObjectMode.Singleton)
                            bIsPortExist = True
                            Exit For
                        End If
                    Catch ex As Exception
                        IMErrorMsg = ex.Message.ToString
                        IMStackTraceMsg = ex.StackTrace.ToString
                        bIsPortExist = False
                        GetIMPort = 0
                        '  ExceptionManager.ProcessException(ex, Common.ExceptionManager.MessageTypes.AppError, Common.ExceptionManager.MessageTypes.AppError)
                    End Try
                Next

                If Not IMURL Is Nothing And bIsPortExist Then
                    ObjManagerInstanceMgr = Activator.GetObject(GetType(InstanceManager), IMURL)
                    Dim lease As ILease
                    lease = RemotingServices.GetLifetimeService(ObjManagerInstanceMgr)
                    lease.Renew(TimeSpan.FromHours(18))
                    Dim sProcessId As Integer
                    sProcessId = System.Diagnostics.Process.GetCurrentProcess.Id
                    ObjManagerInstanceMgr.mgrWinprocessid = sProcessId
                    AddProcessKeys(ListAddAppKeys)
                    If Not IsNothing(ListAddAppKeys) Then
                        ObjManagerInstanceMgr.ListOfpKeys = ListAddAppKeys
                    End If
                    Logger.WriteToLogFile("InHost:", IMURL)
                ElseIf bIsPortExist = False AndAlso IMErrorMsg <> "" Then
                    '  gErrMsg = SetErrMsg("LoadMainPage", IMErrorMsg)
                ElseIf bIsPortExist = False AndAlso GetIMPort = 0 Then
                    MsgBox("Port not registered.Contact support...!: ", "LoadMainPage")
                End If
            Else
                Throw New SympraxisException.SettingsException("Cannot init InstanceManger StartPort or EndPort ")
            End If
            Logger.WriteToLogFile("InHost:", "InitIMServercommunication Ed")
        Catch ex As Exception
            ''    ExceptionManager.ProcessException(ex, Common.ExceptionManager.MessageTypes.AppError, Common.ExceptionManager.MessageTypes.AppError)
        Finally

        End Try


    End Sub


    Public Sub AddProcessKeys(ByRef lsAddAppKeys As List(Of InstanceManager.AddAppKeys))
        Dim Irc As Int32 = 0
        Dim Errmsg As String = ""
        Dim ObjectAddAppKeys As InstanceManager.AddAppKeys
        '   Dim objIM As New InstanceManager
        Dim lsthsh As New Hashtable
        Dim templstprocess As New Hashtable
        Try
            ' myIMProcessItems = Nothing
            lsAddAppKeys = New List(Of InstanceManager.AddAppKeys)
            'myReturnStatus = GetIMProcessItems(myIMProcessItems, GetIMPort)
            '  Dim sConfigPath As String = "\\inch-cmtest01\SY54_LBG_BAT\SympraxisClientGUI_Binaries\Configuration"
            Dim myIMProcessItems = objWCF.GetProcessItem(Irc, Errmsg)

            If Irc = 0 AndAlso Not IsNothing(myIMProcessItems) AndAlso myIMProcessItems.Count > 0 Then
                For i As Integer = 0 To myIMProcessItems.Count - 1
                    ObjectAddAppKeys = New InstanceManager.AddAppKeys
                    ObjectAddAppKeys.lProcessId = myIMProcessItems(i).lProcessId
                    ObjectAddAppKeys.sProcessName = myIMProcessItems(i).sProcess
                    ObjectAddAppKeys.sWorkflow = myIMProcessItems(i).sWorkflow
                    ObjectAddAppKeys.sExePath = myIMProcessItems(i).sExePath

                    'lsthsh.Add(myIMProcessItems(i).sProcess, ObjectAddAppKeys.sExePath)
                    '  objIM.hshApps.Add(ObjectAddAppKeys.sProcessName, ObjectAddAppKeys.sExePath)
                    ' templstprocess.Add(ObjectAddAppKeys.sProcessName, ObjectAddAppKeys.sWorkflow & " " & ObjectAddAppKeys.sProcessName & " " & myClientConfig.InstallBase & " " & ObjectAddAppKeys.sExePath)
                    'templstprocess.Add(ObjectAddAppKeys.sProcessName, myClientConfig.InstallBase & ObjectAddAppKeys.sExePath & " " & ObjectAddAppKeys.sWorkflow & " " & ObjectAddAppKeys.sProcessName & " " & myClientConfig.InstallBase & " " & GetIMPort)
                    'templstprocess.Add(ObjectAddAppKeys.sProcessName & "^" & ObjectAddAppKeys.sWorkflow, ObjectAddAppKeys.sExePath & " " & myIMProcessItems(i).sArgument & " " & ObjectAddAppKeys.sWorkflow & " " & ObjectAddAppKeys.sProcessName & " " & GetIMPort)
                    templstprocess.Add(ObjectAddAppKeys.sProcessName & "^" & ObjectAddAppKeys.sWorkflow, ObjectAddAppKeys.sExePath & " " & myIMProcessItems(i).sArgument & " " & ObjectAddAppKeys.sWorkflow & " " & ObjectAddAppKeys.sProcessName & " " & GetIMPort)
                    ' lsthsh.Add()
                    lsAddAppKeys.Add(ObjectAddAppKeys)
                Next
                If Not IsNothing(lsAddAppKeys) AndAlso lsAddAppKeys.Count > 0 Then
                    lsAddAppKeys = lsAddAppKeys
                Else
                    lsAddAppKeys = Nothing
                End If
                ObjManagerInstanceMgr.lstprocess = templstprocess
            End If
        Catch ex As Exception
            'm  ExceptionManager.ProcessException(ex, Common.ExceptionManager.MessageTypes.AppError, Common.ExceptionManager.MessageTypes.AppError)
        End Try


    End Sub
#End Region



#Region "Client Instance"
    Dim WithEvents ObjClientInstanceMgr As InstanceManager = Nothing
    Dim WithEvents IMClientDAX As Sympraxis.Utilities.DataXChange.DataXchange
    Dim WithEvents IMClient As InstanceManagerWFClient
    Dim IMClientPort As String = ""
    Public Class InstanceManagerWFClient
        Inherits Utilities.InstanceManager.InstanceManagerClient
        Dim channel As TcpServerChannel
        Dim myurl As String
        Public Sub New()
            MyBase.New()
            Try

                Dim obj_FormatProvider As New BinaryServerFormatterSinkProvider()
                obj_FormatProvider.TypeFilterLevel = Runtime.Serialization.Formatters.TypeFilterLevel.Full
                Dim WinProcId As String = System.Diagnostics.Process.GetCurrentProcess.Id
                channel = New TcpServerChannel("Server", System.Int32.Parse(WinProcId), obj_FormatProvider)
                ChannelServices.RegisterChannel(channel, False)
                RemotingConfiguration.RegisterWellKnownServiceType(GetType(Utilities.InstanceManager.InstanceManagerClient), WinProcId, WellKnownObjectMode.Singleton)
                myurl = "tcp://localhost:" & WinProcId & "/" & WinProcId
                Dim imc As InstanceManagerWFClient = Activator.GetObject(GetType(Utilities.InstanceManager.InstanceManagerClient), myurl)
                Dim l As ILease = RemotingServices.GetLifetimeService(imc)
                l.Renew(TimeSpan.FromHours(25))



            Catch ex As Exception
                'MessageBox.Show("InstanceManagerWFClient : " & ex.Message)

            End Try
        End Sub
    End Class

    Private Sub ConnectingInstanceMethod()

        If IMClientPort <> "" Then
            'IMClientURL = "tcp://localhost:" & WDPort & "/InstanceManager"
            IMURL = "tcp://localhost:IMPort/InstanceManager"
            IMURL = Replace(IMURL, "IMPort", IMClientPort)
            ObjClientInstanceMgr = Activator.GetObject(GetType(Utilities.InstanceManager.InstanceManager), IMURL)
            Me.Activate()
            If ObjClientInstanceMgr IsNot Nothing Then
                Dim sProcessId As Integer
                sProcessId = System.Diagnostics.Process.GetCurrentProcess.Id
                If ObjClientInstanceMgr IsNot Nothing Then
                    ' intrHostWorkFlowName = "UNET"
                    If Not ObjClientInstanceMgr.AddApp(intrHostProcessName & "^" & intrHostWorkFlowName, sProcessId) Then
                        System.Diagnostics.Process.GetCurrentProcess.Kill()
                    End If
                    'IMClientDAX = ObjClientInstanceMgr.MyDAXObject
                    IMClient = New InstanceManagerWFClient()
                End If
            End If
        End If

    End Sub
#End Region


#End Region

#Region "PublicMethod"




    Public Sub GetReassignablelist(lManagerId As String, lProcessName As String, ByRef ReAssignableInfo As Object, ByRef Irc As Integer, ByRef Errmsg As String) Implements IPluginhostExtented.GetReassignablelist
        Try
            If (workingMode <> "\debug") Then

                MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
                Logger.WriteToLogFile("IHost", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)
                objWCF.GetReassignablelist(lManagerId, lProcessName, ReAssignableInfo, Irc, Errmsg)
                If Irc <> 0 Then
                    Throw New Exception
                End If

            End If
        Catch ex As Exception
            SetupErrorMsg(ex)
        End Try
    End Sub


    Public Sub LMLayoutChanged(ByVal index As Integer) Handles mLayoutmanager.LayoutChanged
        Try

            mLayoutmanager.clear()

            mpanelcontainer = Nothing
            mpanelcontainer = New LayoutSetting.panelcontainer

            AddControlsToPanel(index, mpanelcontainer, WindowsFormsHost, loadPluginsConfigXml.ApplicationHost.LayoutSettings)

            mLayoutmanager.PanelContent = mpanelcontainer
            _CurrentlayoutIdx = index
            RaiseEvent LayoutChanged(index)

        Catch ex As Exception
            SetupErrorMsg(ex)
        End Try
    End Sub



    Public Sub ApplyLayout(index As Integer) Implements IPluginhost.ApplyLayout

        Try
            If _CurrentlayoutIdx = index Then
                mLayoutmanager.HightLightedPanelFocus()
                Exit Sub
            End If
            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)
            Dispatcher.Invoke(CType(AddressOf DelegateApplylayout, ApplyLayoutDelegate), index)
            _CurrentlayoutIdx = index
        Catch ex As Exception

            SetupErrorMsg(ex)

        End Try
    End Sub

    Public Sub RemoveDuplicateApplication()

        If Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.ApplicationTitle) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.ApplicationTitle <> "" Then




            Dim message As String = loadPluginsConfigXml.ApplicationHost.Settings.ApplicationTitle + " already running...."
            Dim mutexCreated As Boolean
            'mutex = New Mutex(True, "Global\" + loadPluginsConfigXml.ApplicationHost.Settings.ApplicationTitle, mutexCreated)
            mutex = New Mutex(True, loadPluginsConfigXml.ApplicationHost.Settings.ApplicationTitle + Environment.UserName + Environment.MachineName, mutexCreated)
            If mutexCreated Then
                mutex.ReleaseMutex()
            End If
            If Not mutexCreated Then
                Dim result As DialogResult
                ' Displays the MessageBox.
                result = MessageBox.Show(New Form With {.TopMost = True}, message, "Sympraxis.Application.InteractiveHost", MessageBoxButtons.OK)
                If (result = Forms.DialogResult.OK) Then
                    ' Closes the parent form.
                    SplashForm.CloseForm(objFrmSplash)
                    Environment.Exit(-1)
                End If
            Else
                Dim sb As New List(Of String)
                Dim p1() As Process = System.Diagnostics.Process.GetProcessesByName("Sympraxis.Application.InteractiveHost")
                If p1.Count > 0 Then
                    For i As Integer = 0 To p1.Count() - 1
                        If p1(i).MainWindowTitle.ToString().Split("-")(0).ToString.Replace(" ", "").ToString() = loadPluginsConfigXml.ApplicationHost.Settings.ApplicationTitle.ToString().Replace(" ", "") Then
                            ''If p1(i).MainWindowTitle.ToString() = ApplicationFrameWorkConfig.Setting.ApplicationTitle.ToString() Then

                            Dim buttons As MessageBoxButtons = MessageBoxButtons.OK
                            Dim kProcessId As Integer = System.Diagnostics.Process.GetCurrentProcess.Id
                            Dim result As DialogResult
                            ' Displays the MessageBox.
                            result = MessageBox.Show(New Form With {.TopMost = True}, message, "Sympraxis.Application.InteractiveHost", MessageBoxButtons.OK)
                            If (result = Forms.DialogResult.OK) Then
                                ' Closes the parent form.
                                SplashForm.CloseForm(objFrmSplash)
                                Environment.Exit(-1)
                            End If
                        End If
                    Next
                End If
            End If
        Else
            Throw New Exception("Application Title is Empty")
        End If


    End Sub


    Public Sub ChangeWorkflowRequest(WorkflowName As String, ProcessName As String) Implements IPluginhost.ChangeWorkflowRequest
        Dim irc As Integer = 0
        ' Dim Errmsg As String = String.Empty


        Try

            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("IHost", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)

            Logger.WriteToLogFile("IHost", "Change Workflow Request Start")
            ' WCF.SwitchworkflowProcess(WorkflowName, ProcessName)
            If configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True Then
                Throw New SympraxisException.InformationException("Workpacket is already open.Please close")
                Logger.WriteToLogFile("IHost", "Workpacket is already open.Please close")
                Exit Sub
            End If
            RibbonWin.IsMinimized = True

            Dim sfiltercriteria As String = configMgr.AppbaseConfig.EnvironmentVariables.sFilterCriteria
            Dim sWDWorkgroupby As String = configMgr.AppbaseConfig.EnvironmentVariables.sWDWorkgroupby
            Dim sWDlOwner As String = configMgr.AppbaseConfig.EnvironmentVariables.sWDlOwner
            Dim sWDOwnerClass As String = configMgr.AppbaseConfig.EnvironmentVariables.sWDOwnerClass

            Dim isOpenSelectBy As Boolean = configMgr.AppbaseConfig.EnvironmentVariables.isOpenSelectBy
            Dim SelectByXml As String = configMgr.AppbaseConfig.EnvironmentVariables.SelectByXml

            objWCF.workFlowName = WorkflowName
            objWCF.processName = ProcessName

            If Not IsNothing(appHostConfigXml.InstanceMgr) AndAlso Not IsNothing(appHostConfigXml.InstanceMgr.ExternalInstance) AndAlso appHostConfigXml.InstanceMgr.ExternalInstance <> "" Then
                Dim Processlist As String() = appHostConfigXml.InstanceMgr.ExternalInstance.Split("|")

                If Processlist.Contains(ProcessName) = True Then
                    If Not ObjManagerInstanceMgr Is Nothing Then
                        Logger.WriteToLogFile("IHost:", "StartProcess :" & ProcessName & "-" & ProcessName)
                        ObjManagerInstanceMgr.StartProcess(ProcessName, WorkflowName)

                        Logger.WriteToLogFile("IHost:", "StartProcess :" & ProcessName & "-" & ProcessName)
                        Exit Sub
                    End If
                End If
            Else
                Logger.WriteToLogFile("IHost:", "Instance Mgr not configured")
            End If

            If intrHostWorkFlowName <> objWCF.workFlowName Or intrHostProcessName <> objWCF.processName Then
                Logger.WriteToLogFile("IHost", "Change WorkFlow Started")
                ChangeWorkFlow(True)
                configMgr.AppbaseConfig.EnvironmentVariables.WFName = objWCF.workFlowName
                If irc <> 0 Then Exit Sub
                Logger.WriteToLogFile("IHost", "Change WorkFlow Ends")
            End If


            configMgr.AppbaseConfig.EnvironmentVariables.sFilterCriteria = sfiltercriteria
            configMgr.AppbaseConfig.EnvironmentVariables.sWDWorkgroupby = sWDWorkgroupby
            configMgr.AppbaseConfig.EnvironmentVariables.sWDlOwner = sWDlOwner
            configMgr.AppbaseConfig.EnvironmentVariables.sWDOwnerClass = sWDOwnerClass

            configMgr.AppbaseConfig.EnvironmentVariables.isOpenSelectBy = isOpenSelectBy
            configMgr.AppbaseConfig.EnvironmentVariables.SelectByXml = SelectByXml

            Logger.WriteToLogFile("IHost", "Update Environmentvariable Started")
            UpdateEnvironmentvariable()
            If irc <> 0 Then Exit Sub
            Logger.WriteToLogFile("IHost", "Update Environmentvariable Ends")

            If irc = 0 Then

                Logger.WriteToLogFile("IHost", "Load Plugins Started")
                LoadPlugins(appHostConfigXml.PluginDefinitions, loadPluginsConfigXml.ApplicationHost.Plugins, loadPluginsConfigXml.ApplicationHost.LayoutSettings)
                If irc <> 0 Then Exit Sub
                Logger.WriteToLogFile("IHost", "Load Plugins Ends")

            End If


            'AA
            If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.sShowStatusBar) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.sShowStatusBar.ToUpper = "TRUE" Then
                If Not MyDictCount.ContainsKey(intrHostProcessName) Then
                    WorkedObjectCount = 0
                    CompletedObjectCount = 0
                    MyDictCount.Add(intrHostProcessName, CompletedObjectCount.ToString & "|" & WorkedObjectCount.ToString)
                Else
                    Dim myval As String = MyDictCount(intrHostProcessName)
                    WorkedObjectCount = CInt(myval.Split("|")(1).ToString)
                    CompletedObjectCount = CInt(myval.Split("|")(0).ToString)
                    MyDictCount.Remove(intrHostProcessName)
                    MyDictCount.Add(intrHostProcessName, CompletedObjectCount.ToString & "|" & WorkedObjectCount.ToString)
                End If





                mLayoutmanager.AddingStatusBarControls(True, CompletedObjectCount.ToString, WorkedObjectCount.ToString)

            End If

            ' DPValue = Nothing

            If irc = 0 Then

                Menuhandle()
                SetApplicationTitle()
                If appStartBy.ToUpper() <> "STARTUP" Then
                    Dispatcher.BeginInvoke(Sub() DelegateApplylayout(0))
                End If
            End If

            Logger.WriteToLogFile("IHost", "Change Workflow Request Ends")

        Catch ex As Exception

            SetupErrorMsg(ex)

        Finally
            If irc <> 0 Then
                '    SetupErrorMsg(IHErrException) Expdetails, "ChangeWork flow Request Error", ErrorCode)
            End If

        End Try



    End Sub





    Public Sub GetDatapart(Name As String, ByRef Value As Object) Implements IPluginhost.GetDatapart

        Dim Irc As Integer = 0
        Dim Errmsg As String = String.Empty

        ' Dim Errmsg As String = String.Empty

        Dim dp As New Dictionary(Of String, String)

        Try
            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("IHost", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)
            If Not String.IsNullOrEmpty(Name) Then
                dp.Add(Name, Nothing)

                Irc = objWCF.Read(dp, Errmsg)

                If Irc = 0 Then
                    If Not IsNothing(dp) AndAlso dp.ContainsKey(Name) Then
                        Value = dp(Name)
                    Else
                        Value = ""
                    End If
                End If
            End If

        Catch ex As Exception
            SetupErrorMsg(ex)





        End Try
    End Sub

    Public Sub GetMatchingQualifier(FieldMap As Array, ByRef MatchField As Object) Implements IPluginhost.GetMatchingQualifier

        Dim irc As Integer = 0
        Dim Errmsg As String = String.Empty


        Try

            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("IHost", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)

            objWCF.GetMatchingQualifier(FieldMap, MatchField, irc, Errmsg)
            If irc <> 0 Then
                Throw New Exception(Errmsg)
            End If
        Catch ex As Exception
            SetupErrorMsg(ex)



        End Try

    End Sub



    Public Sub GetPlugin(Name As String, ByRef Plugin As Object) Implements IPluginhost.GetPlugin

        Dim irc As Integer = 0


        Try
            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("IHost", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)

            If Not String.IsNullOrEmpty(Name) Then
                If Not IsNothing(hshDataPlugins) AndAlso hshDataPlugins.ContainsKey(Name) Then

                    Plugin = hshDataPlugins(Name)
                Else
                    '  Throw New Exception(String.Concat(Name, "- Plugin not loaded Initially"))
                End If
            Else
                Throw New Exception("PluginName is empty")
            End If
        Catch ex As Exception
            SetupErrorMsg(ex)


        Finally

        End Try
    End Sub


    Public Sub GetMatchingQualifiers(FieldMap As Array, ByRef MatchField As Object) Implements IPluginhost.GetMatchingQualifiers

        Dim irc As Integer = 0
        Dim Errmsg As String = String.Empty


        Try
            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("IHost", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)

            objWCF.GetMatchingQualifiers(FieldMap, MatchField, irc, Errmsg)
            If irc <> 0 Then
                Throw New Exception(Errmsg)
            End If
        Catch ex As Exception
            SetupErrorMsg(ex)

        End Try
    End Sub

    Public Sub EnableLayout(LayoutList As List(Of Integer), EnableNdisable As Boolean) Implements IPluginhost.EnableLayout
        Dim irc As Integer = 0


        Try
            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("IHost", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)

            If LayoutList.Count = 0 Then

                Throw New Exception("No Layout in  LayoutList  to Enable or Disable")
            End If

            Dim i As Integer

            For i = 0 To LayoutList.Count - 1
                If LayoutList.Item(i) <= loadPluginsConfigXml.ApplicationHost.LayoutSettings(0).Layouts.Count - 1 Then
                    mLayoutmanager.LayoutSettingEnable(LayoutList.Item(i), EnableNdisable)
                End If
            Next
        Catch ex As Exception

            SetupErrorMsg(ex)
        Finally

        End Try
    End Sub




    Public Sub AddWorkWithNewItems(WorkItemList As List(Of Workflow.API.Workpacket.WorkPacket.WFObject), ByRef irc As Integer) Implements IPluginhost.AddWorkWithNewItems

        Dim Errmsg As String = String.Empty

        Try
            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("IHost", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)

            irc = 0
            objWCF.AddWorkWithNewItems(WorkItemList, irc, Errmsg)
            If irc <> 0 Then

                Throw New Exception(Errmsg)
            End If
            ObjectQualifyingAdaptor = objWCF.ObjectQualifyingAdaptor

            If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.CSP) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.CSP.ToUpper = "TRUE" Then
                AHMenuOpenWithNewItems()
            Else
                AHMenuCloseWork(False)
            End If

        Catch ex As Exception
            SetupErrorMsg(ex)
        Finally
            If irc <> 0 Then

            End If
        End Try
    End Sub

    Public Sub ClearMsg() Implements IPluginhost.ClearMsg

    End Sub

    Public Sub CloseWork() Implements IPluginhost.CloseWork
        Dim irc As Integer = 0


        '' Parallel workitem not allowed


        Try
            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("IHost", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)

            If configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True Then
                If objWCF.CurrentWorkitemIdx < objWCF.WorkitemCount Then
                    AHMenuNextWork(False)
                ElseIf objWCF.CurrentWorkitemIdx = objWCF.WorkitemCount Then

                    Dispatcher.Invoke(CType(AddressOf AHMenuCloseWork, DelCloseWork), False)
                End If


            End If

        Catch ex As Exception

            SetupErrorMsg(ex)



        Finally

        End Try
    End Sub

    Public Sub CloseWork(AbortValidation As Boolean) Implements IPluginhost.CloseWork

    End Sub


    Public Sub GetReadonly(Values As String, GetObjectTypeby As [Enum], ByRef irc As Integer, ByRef Exception As String, ByRef ReadonlyObject As String) Implements IPluginhost.GetReadonly

    End Sub


    Public Sub GetWITemplate(ObjectType As [Enum], CurrView As [Enum], ByRef Work As Workflow.API.Workpacket.WorkPacket.WFObject) Implements IPluginhost.GetWITemplate
        Dim rc As Int32 = 0
        'Dim myEn As Sympraxis.Utilities.INormalPlugin.IPluginhost.ObjectType
        'Dim sTemplatePath As String = ""
        Dim gErrmsg As String

        Try
            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)

            If (workingMode <> "\debug") Then
                rc = objWCF.ObjAppExchange.GetView4Process(intrHostProcessName, intrHostWorkFlowName, ObjectType.ToString, CurrView)
                If rc = 0 AndAlso Not IsNothing(objWCF.ObjAppExchange.myLoginRsp.stProcessView) Then
                    Work = objWCF.ObjAppExchange.myLoginRsp.stProcessView
                Else
                    gErrmsg = objWCF.ObjAppExchange.gErrMsg
                    If (gErrmsg = "") Then
                        rc = "Failed to get the Template for Workflow=(" & intrHostWorkFlowName & ") and Process=(" & intrHostProcessName & ")"
                    End If
                End If
            End If

        Catch ex As Exception
            rc = 1
            SetupErrorMsg(ex)
        Finally

        End Try
    End Sub

    Public Sub HideToolstrip() Implements IPluginhost.HideToolstrip

        Try
            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("IHost", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)

            ToolstripError.Visibility = System.Windows.Visibility.Collapsed
            pnlNotifyHeader.Visibility = System.Windows.Visibility.Collapsed

        Catch ex As Exception


            SetupErrorMsg(ex)

        End Try
    End Sub

    Public flag As Boolean = True
    Public Sub InsMgrWDOwnerUpdate(sWDlOwner As String, sWDOwnerClass As String, sFilterCriteria As String, ByRef errmsg As String, ByRef Irc As Integer) Implements IPluginhost.InsMgrWDOwnerUpdate

        MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
        Logger.WriteToLogFile("IHost", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)

        If flag = True Then

            DAXObject.SetValue("SYS_IM_sWDlOwner", sWDlOwner)
            DAXObject.SetValue("SYS_IM_sWDOwnerClass", sWDOwnerClass)
            DAXObject.SetValue("SYS_IM_sFilterCriteria", sFilterCriteria)
        ElseIf flag = False Then
            If Not ObjManagerInstanceMgr Is Nothing Then
                ObjManagerInstanceMgr.UpdateDaX("SYS_IM_sWDlOwner", sWDlOwner)
                ObjManagerInstanceMgr.UpdateDaX("SYS_IM_sWDOwnerClass", sWDOwnerClass)
                ObjManagerInstanceMgr.UpdateDaX("SYS_IM_sFilterCriteria", sFilterCriteria)
            End If
        End If
        flag = False
    End Sub



    Public Sub LoginDelegateUser(userid As String) Implements IPluginhost.LoginDelegateUser
        Dim irc As Integer = 0
        Dim Errmsg As String = String.Empty
        HostIrc = 0

        Try
            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)

            delegatorUser = objWCF.GetUserName(userid, irc, Errmsg)

            If irc <> 0 Then
                Throw New Exception()
            End If


            Me.Title = delegatorUser + " : " + Me.Title
            objWCF.LoginDelegateUser(userid, irc, Errmsg)

            If irc <> 0 Then
                Throw New Exception()
            End If
        Catch ex As Exception

            SetupErrorMsg(ex)
        Finally

        End Try
    End Sub



    Private Sub DAXObject_OnDaxUpdated(key As String) Handles DAXObject.OnDaxUpdated
        Try

            Dim Irc As Integer = 0
            'Work Instruction config
            If Not key Is Nothing Then
                Dim dvalue As Object
                If WIConfig IsNot Nothing AndAlso WIConfig.ContextInstruction IsNot Nothing AndAlso WIConfig.ContextInstruction.CWIns.Count > 0 Then
                    dvalue = DAXObject.GetObject(key)
                    If Not dvalue Is Nothing Then
                        Dim Keylist = WIConfig.ContextInstruction.CWIns.Where(Function(a) a.cKey = key AndAlso a.Context = dvalue)
                        If Keylist.Count > 0 Then
                            Logger.WriteToLogFile("Ihost", "SetContextInstruction Starts")
                            SetContextInstruction(key, dvalue)

                            Logger.WriteToLogFile("Ihost", "SetContextInstruction Ends")
                        End If
                    End If
                End If
                If key.ToString.ToUpper = "SYSTITLE" Then

                    dvalue = DAXObject.GetObject("SYSTITLE")

                    If configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True Then
                        If Not dvalue Is Nothing Then
                            mLayoutmanager.DyamictitleChange(dvalue)
                            mLayoutmanager.PaneTitlechanges(mpanelcontainer, dvalue)
                            mLayoutmanager.DyamictitleChange(Nothing)
                        End If
                    End If
                End If
                If key.ToString = "SYS_IM_sWDlOwner" Then

                    configMgr.AppbaseConfig.EnvironmentVariables.sWDlOwner = DAXObject.GetValue(key).ToString

                ElseIf key.ToString = "SYS_IM_sWDOwnerClass" Then
                    configMgr.AppbaseConfig.EnvironmentVariables.sWDOwnerClass = DAXObject.GetValue(key).ToString

                ElseIf key.ToString = "SYS_IM_sFilterCriteria" Then
                    configMgr.AppbaseConfig.EnvironmentVariables.sFilterCriteria = DAXObject.GetValue(key).ToString

                End If

                Dim obj As New Object
                obj = key
                Dim str As String
                If Not obj Is Nothing Then

                    If Not configBannersettings Is Nothing AndAlso Not IsNothing(CaseInfo) AndAlso Not IsNothing(BannerDocuments) Then
                        If configBannersettings.F.Count > 0 Then
                            For i = 0 To configBannersettings.F.Count - 1

                                If configBannersettings.F(i).Type = "DAX" AndAlso configBannersettings.F(i).SXP.ToString.StartsWith("{") AndAlso configBannersettings.F(i).SXP.ToString.Contains(obj) Then
                                    str = ""
                                    If configBannersettings.F(i).SXP <> "" Then
                                        str = configBannersettings.F(i).SXP
            End If
                                    If str <> "" Then
                                        Dim s1 As String
                                        Dim Targetvalue As String
                                        s1 = (str.Replace("{", "")).Replace("}", "")
                                        If s1 <> "" Then
                                            Targetvalue = DAXObject.GetValue(s1)
                                            If Targetvalue <> "" Then
                                                If configBannersettings.F(i).TXP <> "" Then
                                                    BannerDocuments.SelectSingleNode(configBannersettings.F(i).TXP).InnerXml = Targetvalue
                                                    CaseInfo.UpdateXmlDoc(BannerDocuments)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If



                            Next
                        End If
                    End If

                End If
            End If
        Catch ex As Exception

            SetupErrorMsg(ex)
        End Try
    End Sub

    Private Sub FormattedErrmsg(ByRef Errmsg As String, ByRef outErrmsg As String)
        If Errmsg.ToLower.IndexOfAny("'objectId' are locked for processing by another user") = True Then
            Dim i As Integer = Errmsg.IndexOf("[")
            Dim Userid As String = String.Empty
            If i <> -1 Then
                Userid = Errmsg.Substring(i + 1, Errmsg.IndexOf("]", i + 1) - i - 1)

            End If
            If Userid <> String.Empty Then
                outErrmsg = "unable to open the case. Case is locked by" & Userid

            Else
                outErrmsg = "unable to open the case. Case is locked "

            End If
        ElseIf Errmsg.ToLower.IndexOf("Does not have access to the Process") > 0 Then
            outErrmsg = Errmsg
        ElseIf Errmsg.ToLower.IndexOf("not have access for the given") > 0 Then
            outErrmsg = "Access to the case failed. Case could have moved out of accessible queues."
        Else
            outErrmsg = Errmsg
        End If


    End Sub
    Public Sub OpenWorkAcrossProcessAndWF(Values As String, GetObjectTypeby As [Enum], ByRef irc As Integer, ByRef Errmsg As String, Optional SetBGProcess As Boolean = False) Implements IPluginhost.OpenWorkAcrossProcessAndWF


        Try
            HideErrorToolstrip()
            ShowWaitCursor(True, "Opening...", "WorkItem Open..")
            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("IHost", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)
            RibbonWin.IsMinimized = True
            If Not ObjClientInstanceMgr Is Nothing Then
                Dim _CurrentlyActiveApp As String = ""
                If ObjClientInstanceMgr.IsWorkpacketOpened(intrHostProcessName, _CurrentlyActiveApp) Then
                    Throw New SympraxisException.InformationException("Workpacket opened in " & _CurrentlyActiveApp & " application.")
                    Exit Sub
                End If
            End If
            If configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True Then
                Throw New SympraxisException.InformationException("Workpacket is already open.Please close")
                Exit Sub
            End If

            If Not IsNothing(WorkActivity) AndAlso WorkActivity.ToUpper() <> "WORK" Then
                Throw New SympraxisException.InformationException("Please select valid activity(work) in timesheet")
                Exit Sub
            End If

            Logger.WriteToLogFile("IHost", "Disable Menu starts")


            If configMgr.AppbaseConfig.EnvironmentVariables.IsOnline = False Then


                DAXObject.SetValue("SYSOFFLINEOPEN", Values)

            Else


                ' temp code
                '  Values = "546"

                Logger.WriteToLogFile("SAFe", "Host.OpenWork :  " & Values.ToString)
                Logger.WriteToLogFile("Appbase", "GET_WORK | SELECTIVE")
                objWCF.OpenWorkbyObjectId(Values, GetObjectTypeby, False, irc, Errmsg, SetBGProcess)



                If irc <> 0 Then

                    If Errmsg.Contains("'ObjectId' are locked for processing by another user") = True Then
                        Dim frmMessagebox As New FrmMessage
                        Dim _Result As String = ""
                        frmMessagebox.lblMessage = "Work Object is locked for processing" & vbCrLf & "Open 'Read-Only or click 'Notify to receive notification when the work Object is no longer in use''"
                        frmMessagebox.ShowDialog()
                        _Result = frmMessagebox._Result
                        Logger.WriteToLogFile("Ihost", frmMessagebox.lblMessage)
                        If _Result = "btnReadonly" Then
                            PeekObject(Values, GetObjectTypeby, irc, Errmsg, "")
                        ElseIf _Result = "btnNotify" Then
                            objWCF.OpenWorkbyObjectId(Values, GetObjectTypeby, True, irc, Errmsg, SetBGProcess)
                            Logger.WriteToLogFile("Ihost", "Notified to OtherUser")
                        Else
                            Dim Outstring As String = String.Empty
                            FormattedErrmsg(Errmsg, Outstring)

                            Throw New SympraxisException.InformationException(Outstring)

                        End If
                    Else
                        Dim Outstring As String = String.Empty
                        FormattedErrmsg(Errmsg, Outstring)
                        Throw New SympraxisException.InformationException(Outstring)
                    End If

                End If
                Logger.WriteToLogFile("IHost", "Disable Menu starts")

                ObjectQualifyingAdaptor = objWCF.ObjectQualifyingAdaptor

                DAXObject.SetValue("SYS_OP_ObjectId", objWCF.Objectid)

                SetWorkitemEnvironmentvariable()
                WaitCursortext("Opening Workpacket..", "Opening...")

                If intrHostWorkFlowName <> objWCF.workFlowName Or intrHostProcessName <> objWCF.processName Then
                    Logger.WriteToLogFile("IHost", "Disable Menu starts")
                    WaitCursortext("Loading Process", "Loading " & objWCF.workFlowName & "," & objWCF.processName)
                    ChangeWorkFlow(False)
                    configMgr.AppbaseConfig.EnvironmentVariables.WFName = objWCF.workFlowName

                    Logger.WriteToLogFile("IHost", "Disable Menu starts")
                    If irc = 0 Then
                        WaitCursortext("Opening Workitem..", "Init Plugins...")
                        LoadPlugins(appHostConfigXml.PluginDefinitions, loadPluginsConfigXml.ApplicationHost.Plugins, loadPluginsConfigXml.ApplicationHost.LayoutSettings)
                    End If
                End If

                objWCF.SetPredicateOnOpenWorkPacket(irc, Errmsg)
                If irc <> 0 Then
                    Throw New SympraxisException.AppException(Errmsg)
                End If

                UpdateEnvironmentvariable()

                CreateIOmanagerDirectory()

                CheckInputParser()



                ReadDatapart(fullDP)

                WaitCursortext("Opening Workitem..", "LoadWorkitem...")

                LoadWorkitem(fullDP)
                'HshGetDP = HshFullDP(WCF.Objectid)

                ' DPValue = Nothing
                LoadChildworkitem()

                AHEnableMenu("Home", "Close|Exit")



                SetApplyLayout()

                SetMatchBanner()

                AHEnableDisableMenu("Help", "WorkInstruction", True)
                LoadWorkInstructions()


                Logger.WriteToLogFile("IHost", "Show lastEvent Starts")
                ShowlastEvent()
                Logger.WriteToLogFile("IHost", "Show lastEvent Ends")

                If Not IsNothing(loadPluginsConfigXml) Then
                    If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.Applylayout) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.Applylayout.Trim() <> String.Empty Then
                        ApplyLayout(loadPluginsConfigXml.ApplicationHost.Settings.Applylayout)
                    End If
                End If

                SetApplicationTitle()
                configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True
                If _Isautoopen = True Then
                    configMgr.AppbaseConfig.EnvironmentVariables.IsAutoOpen = True
                    _Isautoopen = False
                End If

                If Not IsNothing(loadPluginsConfigXml) Then
                    If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.Applylayout) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.Applylayout.Trim() <> String.Empty Then
                        ApplyLayout(loadPluginsConfigXml.ApplicationHost.Settings.Applylayout)
                    End If
                End If

                'AA

                If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.sShowStatusBar) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.sShowStatusBar.ToUpper = "TRUE" Then
                    If Not MyDictCount.ContainsKey(intrHostProcessName) Then
                        WorkedObjectCount = 0
                        CompletedObjectCount = 0
                        MyDictCount.Add(intrHostProcessName, CompletedObjectCount.ToString & "|" & WorkedObjectCount.ToString)
                    Else
                        Dim myval As String = MyDictCount(intrHostProcessName)
                        WorkedObjectCount = CInt(myval.Split("|")(1).ToString)
                        CompletedObjectCount = CInt(myval.Split("|")(0).ToString)
                        MyDictCount.Remove(intrHostProcessName)
                        MyDictCount.Add(intrHostProcessName, CompletedObjectCount.ToString & "|" & WorkedObjectCount.ToString)
                    End If

                    mLayoutmanager.AddingStatusBarControls(True, CompletedObjectCount.ToString, WorkedObjectCount.ToString)

                End If

            End If
        Catch ex As Exception
            configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = False
            SetupErrorMsg(ex)
            Errmsg = ex.Message
        Finally
            ShowWaitCursor(False)
        End Try
    End Sub


    Public Sub SetMenu(ByRef Menu As Forms.ToolStripMenuItem) Implements IPluginhost.SetMenu

        Dim irc As Integer = 0


        HostIrc = 0
        HostErrmsg = String.Empty

        MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
        Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)

        Try
            If Not IsNothing(Menu) AndAlso Menu.DropDownItems.Count > 0 Then
                Dim Headername As String = Menu.Text
                Dim EnableMenuNameList As String = ""
                If IsNothing(PluginMenu) Then
                    PluginMenu = New List(Of Object)
                End If




                For i As Integer = 0 To Menu.DropDownItems.Count - 1

                    If EnableMenuNameList = "" Then
                        EnableMenuNameList = Menu.DropDownItems(i).Text
                    Else
                        EnableMenuNameList = EnableMenuNameList + "|" + Menu.DropDownItems(i).Text
                    End If
                Next
                If Not PluginMenuEnableDisable.ContainsKey(Headername) Then
                    PluginMenuEnableDisable.Add(Headername, EnableMenuNameList)

                    CreateMenu(Menu)

                Else
                    PluginMenuEnableDisable(Headername) = EnableMenuNameList
                    AHEnableDisableMenu(Headername, EnableMenuNameList, True)

                End If
                ReorderHelpTab()

                'Dim objMenu As ToolStripMenuItem = New ToolStripMenuItem
                'objMenu.Name = "TEST"
                'objMenu.ShortcutKeys = Keys.B + Keys.Control
                'objMenu.ShortcutKeyDisplayString = "Ctrl+B"
                'Menu.DropDownItems.Add(objMenu)
                '    If Not PluginMenuEnableDisable.ContainsKey(Headername) Then
                '        PluginMenuEnableDisable.Add(Headername, EnableMenuNameList)
                '        If EnableMenuNameList <> "" Then
                '            CreateMenu(Headername, EnableMenuNameList)
                '        End If
                '    Else
                '        PluginMenuEnableDisable(Headername) = EnableMenuNameList
                '        AHEnableDisableMenu(Headername, EnableMenuNameList, True)

                '    End If

                'ElseIf Not Menu Is Nothing AndAlso Menu.DropDownItems.Count = 0 Then
                '    Dim Headername As String = Menu.Text
                '    PluginMenu.Add(Menu)
                '    CreateMenu(Headername, Headername)
            End If



        Catch ex As Exception
            SetupErrorMsg(ex)
        Finally
            'If irc <> 0 Then
            '    SetupErrorMsg(IHErrException) "", "SetMenu Error", ErrorCode)
            'End If
        End Try

        'AddHandler RibnBtn.Click, AddressOf CallClickEvent
    End Sub

    Dim iswaitcursoron As Boolean = False
    Dim dt As DispatcherTimer = Nothing
    Public Sub ShowWaitCursor(Status As Boolean, Optional HeaderValue As String = "", Optional Subheaderval As String = "") Implements IPluginhost.ShowWaitCursor

        'If IsNothing(dt) Then
        '    dt = New DispatcherTimer()
        'End If
        MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
        Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)
        Try
            If Status = True Then
                If isformactive = True Then
                Waitcursor.IsOpen = True
                    Waitcursor.Placement = Controls.Primitives.PlacementMode.Center
                End If


                WaitCursortext(Subheaderval, HeaderValue)

                Dim ts As TimeSpan = TimeSpan.FromSeconds(0)
                LblTimer.Content = String.Format("{0}", New DateTime(ts.Ticks).ToString("mm:ss"))
                DoEvents()
                '  AddHandler dt.Tick, AddressOf dispatcherTimer_Tick

                'dt.Interval = New TimeSpan(0, 0, 1)

                'dt.Start()
                iswaitcursoron = True
            Else
                Waitcursor.IsOpen = False
                iswaitcursoron = False
                'dt.Stop()
                'dt = Nothing
            End If

        Catch ex As Exception




        End Try
    End Sub


    Public Sub PluginPrompt(Plugin As Object) Implements IPluginhost.PluginPrompt

        Try
            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)
            If Not IsNothing(mLayoutmanager) Then


                mLayoutmanager.INPGotFocus(Plugin)
                mLayoutmanager.ToggleAutohidens(Plugin)

                mLayoutmanager.UpdateLayout()


            End If

        Catch ex As Exception


            SetupErrorMsg(ex)

        End Try
    End Sub

    Public Sub RemoveMenu(ByRef Menu As Forms.ToolStripMenuItem) Implements IPluginhost.RemoveMenu


        Try
            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)

            Dim EnableMenuNameList As String = ""
            Dim HeaderTabname As String = ""
            Dim Header As String = ""
            If Not String.IsNullOrEmpty(Menu.Text) AndAlso Menu.Text.Substring(0, 1) = "&" Then
                HeaderTabname = Menu.Text.Remove(0, 1)
            End If
            If Not IsNothing(Menu) AndAlso Menu.DropDownItems.Count > 0 Then
                For i As Integer = 0 To Menu.DropDownItems.Count - 1
                    If EnableMenuNameList = "" Then
                        EnableMenuNameList = Menu.DropDownItems(i).Text
                    Else
                        EnableMenuNameList = EnableMenuNameList + "|" + Menu.DropDownItems(i).Text
                    End If
                Next
            End If
            AHEnableDisableMenu(HeaderTabname, EnableMenuNameList, False)
        Catch ex As Exception

            SetupErrorMsg(ex)

        End Try
    End Sub

    Public Sub Unit(units As Common.XmlUnits.Units) Implements IPluginhost.Unit
        Try
            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)

            Dim Irc As Integer
            Dim Errmsg As String = String.Empty
            objWCF.Unit(units, Irc, Errmsg)
            If Irc <> 0 Then
                Throw New Exception
            End If
        Catch ex As Exception
            SetupErrorMsg(ex)
        End Try
    End Sub

    Public Sub Units(Units As Common.XmlUnits.Units, WFObjectIndex As Integer) Implements IPluginhost.Units
        Try
            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)
            Dim Irc As Integer
            Dim Errmsg As String = String.Empty

            objWCF.Units(Units, WFObjectIndex, Irc, Errmsg)
            If Irc <> 0 Then
                Throw New Exception
            End If
        Catch ex As Exception
            SetupErrorMsg(ex)
        End Try
    End Sub

    Public Sub UpdateDataPartValues(DP As Dictionary(Of String, String)) Implements IPluginhost.UpdateDataPartValues

        Try
            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)
            Logger.WriteToLogFile("Ihost", "Write datapart starts")
            objWCF.Write(DP, HostErrmsg)
            Logger.WriteToLogFile("Ihost", "Write datapart Ends")

        Catch ex As Exception
            SetupErrorMsg(ex)
        End Try
    End Sub

    Public Sub UpdatePerformancelog(ObjectIndex As Integer, TimeSpan As TimeSpan) Implements IPluginhost.UpdatePerformancelog

    End Sub

    Public Sub WaitCursorProgressvalue(InputTime As Double) Implements IPluginhost.WaitCursorProgressvalue

        Try
            Prgbar.Value = InputTime
            DoEvents()
        Catch ex As Exception


            SetupErrorMsg(ex)
        End Try

    End Sub

    Public Sub WaitCursortext(SubHeaderValue As String, HeaderValue As String) Implements IPluginhost.WaitCursortext

        Try
            LblLoading.Content = HeaderValue
            LblShowtext.Content = SubHeaderValue
            DoEvents()
        Catch ex As Exception


            SetupErrorMsg(ex)


        End Try
    End Sub


    Public Sub GetMatchRule(ByRef MatchedRule As Object, ByRef LstRulelist As Object, ByRef CrRule As Object, ByRef Irc As Integer, ByRef Errmsg As String) Implements IPluginhostExtented.GetMatchRule
        Try
            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)
            objWCF.GetMatchRule(MatchedRule, LstRulelist, CrRule, Irc, Errmsg)
            If Irc <> 0 Then
                Throw New SympraxisException.AppException(Errmsg)
            End If
        Catch ex As Exception
            SetupErrorMsg(ex)

        End Try
    End Sub


    Public Sub NotificationMsg(Message As String, NotifyCode As String, ByRef Irc As Integer, ByRef ErrMsg As String) Implements IPluginhost.NotificationMsg


        Dim notifydoc As System.Xml.XmlDocument
        Dim Showmessage As String = ""

        MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
        Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)

        Try
            If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.bNotify) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.bNotify <> String.Empty _
                 AndAlso loadPluginsConfigXml.ApplicationHost.Settings.bNotify.ToUpper() = "TRUE" Then

                If Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.AutoOpenNotify) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.AutoOpenNotify <> String.Empty Then
                    notifydoc = New XmlDocument

                    notifydoc.LoadXml(Message)
                    Showmessage = notifydoc.DocumentElement.GetAttribute("Msg")
                    If String.IsNullOrEmpty(Showmessage.Trim()) Then
                        Irc = 1
                        ErrMsg = "Notification Message is Empty"
                        Exit Sub
                    End If


                    If NotifyCode.ToUpper <> Sympraxis.Utilities.INormalPlugin.IPluginhost.NotifyCode.CurrentWorkRequested.ToString.ToUpper Then
                        Dim objId As String = notifydoc.DocumentElement.GetAttribute("lObjectId")
                        If objId <> String.Empty Then
                            NotifyMessageObjectId = objId
                        End If
                    End If

                    If loadPluginsConfigXml.ApplicationHost.Settings.AutoOpenNotify.ToUpper() = "FALSE" Then
                        Logger.WriteToLogFile("Ihost", "ShowNotification starts")
                        ShowNotification(NotifyCode, Showmessage, notifydoc)

                        Logger.WriteToLogFile("Ihost", "ShowNotification Ends")
                    ElseIf loadPluginsConfigXml.ApplicationHost.Settings.AutoOpenNotify.ToUpper() = "TRUE" Then
                        Dim processname As String = notifydoc.DocumentElement.GetAttribute("PN")
                        Dim DBProcesstype As String = ""
                        If Not String.IsNullOrEmpty(processname.Trim()) Then

                            DBProcesstype = objWCF.GetProcessType(processname, Irc, ErrMsg)
                        End If

                        If DBProcesstype.ToUpper() = "PUSH2USER" AndAlso configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = False AndAlso Not IsNothing(NotifyMessageObjectId) AndAlso NotifyMessageObjectId <> "0" Then
                            _Isautoopen = True
                            If Dispatcher.CheckAccess() Then
                                Dispatcher.Invoke(CType(AddressOf OpenWorkAcrossProcessAndWF, DelOpenWork), NotifyMessageObjectId, Sympraxis.Utilities.INormalPlugin.IPluginhost.GetObjectType.ObjectId, 0, "")
                            Else
                                Dispatcher.BeginInvoke(CType(AddressOf OpenWorkAcrossProcessAndWF, DelOpenWork), NotifyMessageObjectId, Sympraxis.Utilities.INormalPlugin.IPluginhost.GetObjectType.ObjectId, 0, "")
                            End If
                            If configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True Then
                                configMgr.AppbaseConfig.EnvironmentVariables.IsAutoOpen = True
                            End If
                        Else

                            Logger.WriteToLogFile("Ihost", "ShowNotification starts")
                            ShowNotification(NotifyCode, Showmessage, notifydoc)

                            Logger.WriteToLogFile("Ihost", "ShowNotification Ends")

                        End If
                    End If
                Else
                    Irc = 1
                    ErrMsg = "No Configuration Setting for AutoOpen"

                End If





            End If
        Catch ex As Exception
            SetupErrorMsg(ex)
        Finally
            notifydoc = Nothing

        End Try
    End Sub


    Public Sub SetRibbonMenu(ByRef RibTab As Object, ByRef Irc As Integer, ByRef Errmsg As String) Implements IPluginhostExtented.SetRibbonMenu
        Try

            If Not IsNothing(RibTab) Then
                MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
                Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)
                Dim isTabAvl As Boolean = False
                'For Each RibTablp As RibbonTab In RibbonWin.Items
                '    If RibTablp.Header.ToString().ToLower() = RibTab.Header.ToLower() Then
                '        isTabAvl = True
                '        Exit For
                '    End If
                'Next
               
                
                If isTabAvl Then

                    For Each RibTablp As RibbonTab In RibbonWin.Items
                        If RibTablp.Header.ToString().ToLower() = RibTab.Headername.ToLower() Then
                            For Each sRibGroup As RibbonGroup In RibTablp.Items
                                For Each Group As RibbonGroup In RibTab.Groups
                                    For Each Btn As RibbonButton In Group.Items
                                        sRibGroup.Items.Add(Btn)
                                    Next
                                Next
                            Next
                        End If
                    Next
                Else

                    RibbonWin.Items.Add(DirectCast(RibTab, RibbonTab))
                    RibbonWin.UpdateLayout()
                Me.UpdateLayout()

            End If
                ReorderHelpTab()
            End If
        Catch ex As Exception

            SetupErrorMsg(ex)
        End Try

    End Sub



    Public Sub LMActiveContentChanged(ByVal obj As Object) Handles mLayoutmanager.LMActiveContentChanged
        Try
            Dim rc As Int32 = 0
            Logger.WriteToLogFile("IHost", "LMActiveContentChanged Starts")
            For Each element As Object In lstPluginOrder
                If TypeOf (obj) Is System.Windows.Forms.Integration.WindowsFormsHost Then
                    Dim str As New System.Windows.Forms.Integration.WindowsFormsHost
                    str = obj
                    If element.Equals(str.Child) Then
                        If Not IsNothing(element) Then
                            Logger.WriteToLogFile("IHost", "Plugin SetFocus -->" & element.ToString() & " Starts")
                            DirectCast(element, INormalPlugin).SetFocus()
                            Logger.WriteToLogFile("IHost", "Plugin SetFocus -->" & element.ToString() & " Ends")

                        End If

                    End If
                Else
                    If element.Equals(obj) Then
                        Logger.WriteToLogFile("IHost", "Plugin SetFocus -->" & element.ToString() & " Starts")
                        DirectCast(element, INormalPlugin).SetFocus()
                        Logger.WriteToLogFile("IHost", "Plugin SetFocus -->" & element.ToString() & " Ends")
                    End If
                End If
            Next
            For Each element As Object In lstStartupPluginOrder
                If TypeOf (obj) Is System.Windows.Forms.Integration.WindowsFormsHost Then
                    Dim str As New System.Windows.Forms.Integration.WindowsFormsHost
                    str = obj
                    If element.Equals(str.Child) Then
                        If Not IsNothing(element) Then
                            Logger.WriteToLogFile("IHost", "Plugin SetFocus -->" & element.ToString() & " Starts")
                            DirectCast(element, INormalPlugin).SetFocus()
                            Logger.WriteToLogFile("IHost", "Plugin SetFocus -->" & element.ToString() & " Ends")

                        End If

                    End If
                Else
                    If element.Equals(obj) Then
                        Logger.WriteToLogFile("IHost", "Plugin SetFocus -->" & element.ToString() & " Starts")
                        DirectCast(element, INormalPlugin).SetFocus()
                        Logger.WriteToLogFile("IHost", "Plugin SetFocus -->" & element.ToString() & " Ends")
                    End If
                End If

            Next
        Catch ex As Exception

            SetupErrorMsg(ex)
        End Try
    End Sub

    Public Sub PeekObject(Values As String, GetObjectTypeby As [Enum], ByRef irc As Integer, ByRef Exception As String, ByRef PeekObject As String) Implements IPluginhost.PeekObject

        Try
            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)

            RibbonWin.IsMinimized = True

            Logger.WriteToLogFile("IHost", "PeekObject Starts")

            objWCF.PeekObject(Values, GetObjectTypeby, irc, Exception, PeekObject)

            If irc <> 0 Then
                Throw New Exception(Exception)
            End If
            ObjectQualifyingAdaptor = objWCF.ObjectQualifyingAdaptor

            configMgr.AppbaseConfig.EnvironmentVariables.IsReadonlyWP = True
            'ToolstripReadOnly.Visibility = System.Windows.Visibility.Visible
            btnReadnly.Visibility = System.Windows.Visibility.Visible
            Logger.WriteToLogFile("IHost", "PeekObject Ends")


            ' DAXObject.SetValue("SYS_OP_ObjectId", objWCF.Objectid)

            SetWorkitemEnvironmentvariable()


            If intrHostWorkFlowName <> objWCF.workFlowName Or intrHostProcessName <> objWCF.processName Then
                Logger.WriteToLogFile("IHost", "Disable Menu starts")
                ChangeWorkFlow(True)
                configMgr.AppbaseConfig.EnvironmentVariables.WFName = objWCF.workFlowName
                Logger.WriteToLogFile("IHost", "Disable Menu starts")
                If irc = 0 Then
                    LoadPlugins(appHostConfigXml.PluginDefinitions, loadPluginsConfigXml.ApplicationHost.Plugins, loadPluginsConfigXml.ApplicationHost.LayoutSettings)
                End If

            End If

            CreateIOmanagerDirectory()


            CheckInputParser()



            ReadDatapart(fullDP)



            LoadWorkitem(fullDP)
            'HshGetDP = HshFullDP(WCF.Objectid)

            ' DPValue = Nothing
            LoadChildworkitem()

            AHEnableMenu("Home", "Close|Exit")

            SetApplyLayout()

            SetMatchBanner()
            AHEnableDisableMenu("Help", "WorkInstruction", True)

            LoadWorkInstructions()

            Logger.WriteToLogFile("IHost", "Show lastEvent Starts")
            ShowlastEvent()
            Logger.WriteToLogFile("IHost", "Show lastEvent Ends")

            If Not IsNothing(loadPluginsConfigXml) Then
                If Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.Applylayout) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.Applylayout.Trim() <> String.Empty Then
                    ApplyLayout(loadPluginsConfigXml.ApplicationHost.Settings.Applylayout)
                End If
            End If

            SetApplicationTitle()
            configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True



            'AA

            If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.sShowStatusBar) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.sShowStatusBar.ToUpper = "TRUE" Then
                If Not MyDictCount.ContainsKey(intrHostProcessName) Then
                    WorkedObjectCount = 0
                    CompletedObjectCount = 0
                    MyDictCount.Add(intrHostProcessName, CompletedObjectCount.ToString & "|" & WorkedObjectCount.ToString)
                Else
                    Dim myval As String = MyDictCount(intrHostProcessName)
                    WorkedObjectCount = CInt(myval.Split("|")(1).ToString)
                    CompletedObjectCount = CInt(myval.Split("|")(0).ToString)
                    MyDictCount.Remove(intrHostProcessName)
                    MyDictCount.Add(intrHostProcessName, CompletedObjectCount.ToString & "|" & WorkedObjectCount.ToString)
                End If

                mLayoutmanager.AddingStatusBarControls(True, CompletedObjectCount.ToString, WorkedObjectCount.ToString)

            End If

        Catch ex As Exception
            SetupErrorMsg(ex)





        End Try
    End Sub
    ''//
    Public Sub EnableMenus(Menuname As String, State As Boolean) Implements IPluginhost.EnableMenus
        MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
        Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)
    End Sub

    ''//
    Public Sub BussinessObjectUpdate() Implements IPluginhost.BussinessObjectUpdate
        MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
        Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)
    End Sub
    ''//
    Public Sub CloseWorkWithNewItems() Implements IPluginhost.CloseWorkWithNewItems
        MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
        Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)
    End Sub

    ''//
    Public Property CurrentWorkitemIdx As Integer Implements IPluginhost.CurrentWorkitemIdx

    ''//
    Public Property ObjectQualifyingAdaptor As Object Implements IPluginhost.ObjectQualifyingAdaptor

    ''//
    Public Sub OpenWork() Implements IPluginhost.OpenWork
        Dim irc As Integer = 0

        Try
            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)
            If Not ObjClientInstanceMgr Is Nothing Then
                Dim _CurrentlyActiveApp As String = ""
                If ObjClientInstanceMgr.IsWorkpacketOpened(intrHostProcessName, _CurrentlyActiveApp) Then
                    Throw New SympraxisException.InformationException("Workpacket opened in " & _CurrentlyActiveApp & " application.")
                    Exit Sub
                End If
            End If

            If configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True Then
                If objWCF.CurrentWorkitemIdx < objWCF.WorkitemCount Then
                    AHMenuNextWork(False)
                ElseIf objWCF.CurrentWorkitemIdx = objWCF.WorkitemCount Then

                    Throw New SympraxisException.InformationException("WorkItem is already Open...")

                End If

            ElseIf configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = False Then
                AHMenuOpenWork()
            End If
        Catch ex As Exception

            SetupErrorMsg(ex)
        Finally

        End Try
    End Sub
    ''//
    Public Sub OpenWork(Values As String, GetObjectTypeby As [Enum], ByRef irc As Integer, ByRef Exception As String) Implements IPluginhost.OpenWork
        MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
        Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)
    End Sub

    ''//
    Public Sub StandAloneCloseWork(ByRef Xnode As String) Implements IPluginhost.StandAloneCloseWork

    End Sub
    ''//
    Public Sub StandAloneOPenWork(Xnode As String) Implements IPluginhost.StandAloneOPenWork

    End Sub
    ''//
    Public Sub StartApp() Implements IPluginhost.StartApp
        MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
        Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)
    End Sub
    ''//
    Public Sub StopApp() Implements IPluginhost.StopApp

    End Sub
    ''//
    Public Sub OpenWorkAndSwapOwner(Values As String, GetObjectTypeby As [Enum], ByRef irc As Integer, ByRef Exception As String) Implements IPluginhost.OpenWorkAndSwapOwner
        MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
        Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)
    End Sub
    ''//
    Public Sub OpenWorkByObjectsFixSize(Values As String, GetObjectTypeby As [Enum], iFixCnt As Integer, ByRef irc As Integer, ByRef Exception As String) Implements IPluginhost.OpenWorkByObjectsFixSize
        MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
        Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)
    End Sub



    ''//
    Public Sub UpdateObjects(WFObject As Workflow.API.Workpacket.WorkPacket.WFObject, WFObjectIndex As Integer) Implements IPluginhost.UpdateObjects
        MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
        Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)
    End Sub


#End Region





#Region "Timer"

    Private Sub loginHourTimer_Tick(sender As Object, e As System.Timers.ElapsedEventArgs)
        Try
            If loginHourTimer.Enabled = True Then
                logininterval = logininterval + 1
                Dim loginhours = dtInterval.AddSeconds(logininterval).ToString("HH:mm:ss")
                loginhours = loginhours
                Dispatcher.Invoke(Sub() mLayoutmanager.LoginHours(loginhours))
            End If
        Catch ex As Exception
            SetupErrorMsg(ex)
        End Try
    End Sub

    Private Sub IdleHourTimer_Tick(sender As Object, e As System.Timers.ElapsedEventArgs)
        Try
            If IdleHourTimer.Enabled = True Then
                idleinterval = idleinterval + 1
                Dim curridlehours As String = Nothing
                'If blnidlestart = True Then
                CurrIdleInterval = CurrIdleInterval + 1
                curridlehours = dtInterval.AddSeconds(CurrIdleInterval).ToString("HH:mm:ss")
                'Else
                '    CurrIdleInterval = 0
                'End If
                Dim idlehours = dtInterval.AddSeconds(idleinterval).ToString("HH:mm:ss")
                If curridlehours IsNot Nothing AndAlso curridlehours IsNot String.Empty Then
                    idlehours = " " & curridlehours & " / " & idlehours
                End If
                'idlehours = " " & idlehours
                Dispatcher.Invoke(Sub() mLayoutmanager.IdleHours(idlehours))
            End If
        Catch ex As Exception
            SetupErrorMsg(ex)
        End Try
    End Sub

    Private Sub TotalIdleHour()
        Try
            If IdleHourTimer.Enabled = False Then
                If idleinterval > 0 Then
                    Dim idlehours = dtInterval.AddSeconds(idleinterval).ToString("HH:mm:ss")
                    idlehours = " " & idlehours
                    Dispatcher.Invoke(Sub() mLayoutmanager.IdleHours(idlehours))
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub WorkHourTimer_Tick(sender As Object, e As System.Timers.ElapsedEventArgs)
        Try
            If WorkHourTimer.Enabled = True Then
                openworkinterval = openworkinterval + 1
                Dim workhours = dtInterval.AddSeconds(openworkinterval).ToString("HH:mm:ss")
                workhours = " " & workhours
                Dispatcher.Invoke(Sub() mLayoutmanager.OpenWorkHours(workhours))
            End If
        Catch ex As Exception
            SetupErrorMsg(ex)
        End Try
    End Sub


    Private Sub HideInfobarTimer_Elapsed(sender As Object, e As EventArgs)
        Try



            If Not HideInfobarTimer Is Nothing Then
                HideInfobarTimer.Enabled = False
                HideInfobarTimer.Stop()
                HideInfobarTimer = Nothing

                Me.Dispatcher.Invoke(Sub() ToolstripError.Visibility = System.Windows.Visibility.Collapsed)


            End If
        Catch ex As Exception

        End Try
    End Sub


    Private Sub AppAutoCloseTimer_Elapsed(sender As Object, e As System.Timers.ElapsedEventArgs)
        Dim Irc As Integer = 0

        Try


            Dim diff As TimeSpan
            diff = DateTime.Now - AutoCloseAHIdleStarttime

            If Not IsNothing(overrideAppHostConfig) AndAlso Not IsNothing(overrideAppHostConfig.AHSettings) AndAlso Not IsNothing(overrideAppHostConfig.AHSettings.AppHostAutocloseInterval) AndAlso overrideAppHostConfig.AHSettings.AppHostAutocloseInterval <> 0 Then

                If CType(diff.TotalMilliseconds, Integer) >= overrideAppHostConfig.AHSettings.AppHostAutocloseInterval * 60 * 1000 Then

                    AppAutoCloseTimer.Stop()
                    AutoCloseAHIdleStarttime = DateTime.Now
                    AppAutoCloseTimer = Nothing
                    configMgr.AppbaseConfig.EnvironmentVariables.AutocloseIdleTimeSt = AutoCloseAHIdleStarttime

                    If configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = False Then
                        Dispatcher.Invoke(CType(AddressOf Exiting, DelExiting), True)

                    Else

                        IsAbortValidation = True

                        configMgr.AppbaseConfig.EnvironmentVariables.IsAutoClose = True
                        Dispatcher.Invoke(CType(AddressOf AHMenuCloseWork, DelCloseWork), False)

                        Dispatcher.Invoke(CType(AddressOf Exiting, DelExiting), True)




                    End If
                End If


            End If
        Catch ex As Exception
            SetupErrorMsg(ex)
        End Try
    End Sub


    Private Sub AttenandUnattentedTimeStartEvent()
        Dim Irc As Integer = 0



        If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not _
       IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso (Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.AttendedInterval) AndAlso _
       loadPluginsConfigXml.ApplicationHost.Settings.AttendedInterval <> 0) Or (Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.UnAttendedCloseInterval) AndAlso _
       loadPluginsConfigXml.ApplicationHost.Settings.UnAttendedCloseInterval <> 0) Then

            If Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.AttendedInterval) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.AttendedInterval > 0 Then
                If IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.UnAttendedCloseInterval) Or (Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.UnAttendedCloseInterval) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.UnAttendedCloseInterval = 0) Then
                    Throw New SympraxisException.AppException(" Configuration Error:Please Config UnAttended CloseInterval ")
                Else
                    ButtonpinVisible()
                End If
            End If



            If IsNothing(ManualIdleTimer) Then

                ManualIdleStarttime = DateTime.Now
                Dim DblTime As Double = 1000
                ManualIdleTimer = New System.Timers.Timer(DblTime)
                AddHandler ManualIdleTimer.Elapsed, AddressOf AttendtedTimer_Elapsed
                ManualIdleTimer.Enabled = True

                'OpenPin.Visible = True
                'OpenPin.Image.Tag = "UnAttentedPin"
            End If

        End If

        If Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.AutocloseInterval) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.AutocloseInterval > 0 Then

            IsObjectautoclose = True
            AutocloseIdleTimerStart()

        End If

    End Sub

    Private Sub AutocloseIdleTimerStart()
        Try
            If Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.AutocloseInterval) Then
                If Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.AutocloseInterval) Then
                    If IsNothing(AutoCloseIdleTimer) Then
                        If Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.AutocloseInterval) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.AutocloseInterval > 0 And IsObjectautoclose = True Then
                            Dim DblTime As Double = loadPluginsConfigXml.ApplicationHost.Settings.AutocloseInterval * 60 * 100
                            AutoCloseIdleTimer = New System.Timers.Timer(DblTime)
                            AddHandler AutoCloseIdleTimer.Elapsed, AddressOf AutoCloseIdleTimer_Elapsed
                            AutoCloseIdleTimer.Enabled = True
                            AutoCloseIdleStarttime = DateTime.Now
                        End If
                    ElseIf Not IsNothing(AutoCloseIdleTimer) AndAlso AutoCloseIdleTimer.Enabled = False Then
                        If Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.AutocloseInterval) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.AutocloseInterval > 0 And IsObjectautoclose = True Then
                            Dim DblTime As Double = loadPluginsConfigXml.ApplicationHost.Settings.AutocloseInterval * 60 * 100
                            AutoCloseIdleTimer = New System.Timers.Timer(DblTime)
                            AddHandler AutoCloseIdleTimer.Elapsed, AddressOf AutoCloseIdleTimer_Elapsed
                            AutoCloseIdleTimer.Enabled = True
                            AutoCloseIdleStarttime = DateTime.Now
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub AutoCloseIdleTimer_Elapsed(sender As Object, e As System.Timers.ElapsedEventArgs)

        Dim diff As TimeSpan
        diff = DateTime.Now - AutoCloseIdleStarttime
        If CType(diff.TotalMilliseconds, Integer) >= loadPluginsConfigXml.ApplicationHost.Settings.AutocloseInterval * 60 * 1000 Then
            If IsObjectautoclose = True Then
                AutoCloseIdleTimer.Stop()
                AutoCloseIdleStarttime = DateTime.Now
                If configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True Then

                    IsAbortValidation = True
                    objWCF.UnitsAutoClose = True
                    If Not IsNothing(AutoCloseIdleTimer) Then
                        AutoCloseIdleTimer.Enabled = False
                        AutoCloseIdleTimer.Stop()
                        AutoCloseIdleTimer = Nothing
                    End If
                    configMgr.AppbaseConfig.EnvironmentVariables.IsAutoClose = True
                    Dispatcher.Invoke(CType(AddressOf AHMenuCloseWork, DelCloseWork), False)

                End If

            End If

        End If
    End Sub

    Private Sub AttendtedTimer_Elapsed(sender As Object, e As System.Timers.ElapsedEventArgs)
        Dim diff As TimeSpan
        Dim Irc As Integer = 0


        diff = DateTime.Now - ManualIdleStarttime
        Dim CaseTime As String = String.Empty
        If IsCaseAttended Then
            CaseTime = loadPluginsConfigXml.ApplicationHost.Settings.AttendedInterval
        Else
            CaseTime = loadPluginsConfigXml.ApplicationHost.Settings.UnAttendedCloseInterval
        End If
        If CType(diff.TotalMilliseconds, Double) >= CDbl(CaseTime) * 60 * 1000 Then

            If configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True Then

                IsAbortValidation = True
                objWCF.UnitsAutoClose = True
                If Not IsNothing(ManualIdleTimer) Then
                    ManualIdleTimer.Enabled = False
                    ManualIdleTimer.Stop()
                    ManualIdleTimer = Nothing
                End If
                configMgr.AppbaseConfig.EnvironmentVariables.IsAutoClose = True
                Dispatcher.Invoke(CType(AddressOf AHMenuCloseWork, DelCloseWork), False)

            End If
        End If
    End Sub




    Private Sub SITimer_Elapsed(sender As Object, e As System.Timers.ElapsedEventArgs)

        If SIPin = False AndAlso SICloseTimer.Enabled = False Then
            SICloseTimer.Enabled = True
            SITimer.Enabled = False
            SITimer.Stop()
        End If

    End Sub

    Private Sub SICloseTimer_Elapsed(sender As Object, e As System.Timers.ElapsedEventArgs)

        If SIPin = False AndAlso Not SICloseTimer Is Nothing Then
            If SICloseTimer.Enabled = True Then
                SIopacity = SIopacity - 0.05
                SIStyle.Dispatcher.Invoke(Sub() SIStyle.Opacity = SIopacity)
                If SIopacity = 0 Or SIopacity < 0 Then
                    ShowSI = False
                    SICloseTimer.Enabled = False
                    SICloseTimer.Stop()
                    SIopacity = 1
                    pnlStadInstruction.Dispatcher.Invoke(Sub() pnlStadInstruction.IsOpen = False)
                    'pnlStadInstruction.Visibility = System.Windows.Visibility.Collapsed
                    SITimer.Enabled = False
                    SITimer.Stop()
                    Me.Dispatcher.Invoke(Sub() SetInstructionLocation(Nothing))
                End If
            End If
        End If

    End Sub

    Private Sub DITimer_Elapsed(sender As Object, e As System.Timers.ElapsedEventArgs)

        If DIPin = False AndAlso DICloseTimer.Enabled = False Then
            DICloseTimer.Enabled = True
            DITimer.Enabled = False
            DITimer.Stop()
        End If

    End Sub


    Private Sub CITimer_Elapsed(sender As Object, e As System.Timers.ElapsedEventArgs)

        If CIPin = False AndAlso CICloseTimer.Enabled = False Then
            CICloseTimer.Enabled = True
            CITimer.Enabled = False
            CITimer.Stop()
        End If

    End Sub

    Private Sub CICloseTimer_Elapsed(sender As Object, e As System.Timers.ElapsedEventArgs)

        If CIPin = False AndAlso Not CICloseTimer Is Nothing Then
            If CICloseTimer.Enabled = True Then
                CIopacity = CIopacity - 0.05
                CIStyle.Dispatcher.Invoke(Sub() CIStyle.Opacity = CIopacity)
                If CIopacity = 0 Or CIopacity < 0 Then
                    ShowCI = False
                    CICloseTimer.Enabled = False
                    CICloseTimer.Stop()
                    CIopacity = 1
                    pnlContextInstruction.Dispatcher.Invoke(Sub() pnlContextInstruction.IsOpen = False)
                    'pnlContextInstruction.Visibility = System.Windows.Visibility.Collapsed
                    CITimer.Enabled = False
                    CITimer.Stop()
                    Me.Dispatcher.Invoke(Sub() SetInstructionLocation(Nothing))
                End If

            End If
        End If

    End Sub

#End Region




#Region "IPluginhost Property"

    Public Property HostIrc As Integer Implements IPluginhostExtented.iErrorCode
    Public Property HostErrmsg As String Implements IPluginhostExtented.sErrorMsg
    Public Property CurrentOfflineWP As String Implements IPluginhost.CurrentOfflineWP

    Private _WorkActivity As String
    Public Property WorkActivity() As String Implements IPluginhost.WorkActivity
        Get
            Return _WorkActivity
        End Get
        Set(value As String)
            _WorkActivity = value
            DAXObject.SetValue("WorkActivity", _WorkActivity)
        End Set
    End Property


#End Region

#Region "IPluginhos Event"
    Public Event IdleTimeEnd(sTime As Date) Implements IPluginhostExtented.IdleTimeEnd
    Public Event IdleTimeStart(sTime As Date) Implements IPluginhostExtented.IdleTimeStart


    Public Event AppReady() Implements IPluginhostExtented.AppReady
    Public Event ResetPerformanceLog() Implements IPluginhostExtented.ResetPerformanceLog


    Public Event LayoutChanged(index As Integer) Implements IPluginhost.LayoutChanged
    Public Event BussinessObjectUpdated() Implements IPluginhost.BussinessObjectUpdated


#End Region







    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub mLayoutmanager_Hyperlink() Handles mLayoutmanager.Hyperlink
        Dim dt As New System.Data.DataTable("dt")

        Dim keyColumn As System.Data.DataColumn = New System.Data.DataColumn("ProcessName")
        keyColumn.DataType = System.Type.GetType("System.String")
        Dim valueColumn As System.Data.DataColumn = New System.Data.DataColumn("Completed")
        valueColumn.DataType = System.Type.GetType("System.String")

        dt.Columns.Add(keyColumn)
        dt.Columns.Add(valueColumn)

        For Each item As KeyValuePair(Of String, String) In MyDictCount
            dt.Rows.Add(item.Key, item.Value)
        Next

        DGObjCount.ItemsSource = dt.DefaultView


        pnlCompleteObjectCount.IsOpen = True
    End Sub

    Private Sub btnObjClose_Click(sender As Object, e As RoutedEventArgs)
        pnlCompleteObjectCount.IsOpen = False
    End Sub

    Private Sub InstanceOnline()


        Logger.WriteToLogFile("IHost", "LoginUser Start")

        LoginUser()

        Logger.WriteToLogFile("IHost", "Assign EnvironmentVariable Start")

        EnvironmentVariable()
        If appStartBy.ToUpper() <> "STARTUP" Then
            Logger.WriteToLogFile("IHost", "Attach Process Start")

            AttachProcess(True)

            Logger.WriteToLogFile("Completed", "Start Attach Process")

            RegisterWF(True)
        End If
       

        UpdateEnvironmentvariable()

    End Sub

    Public Sub GoOffline(ByRef Path As String, ByRef WIPPath As String, ByRef Irc As Integer, ByRef ErrMsg As String) Implements IPluginhost.GoOffline
        Try
            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)

            OfflinePath = Path
            objWCF.GoOffline(Path, WIPPath, Irc, ErrMsg)

            '  OfflineCloseWork()
        Catch ex As Exception
            SetupErrorMsg(ex)
        End Try
    End Sub

    Private Sub OfflineLoadConfig(ByRef AppHostConfigxml As AppHostConfig, ByRef loadPluginsConfigxml As IntHostConfig, ByRef ProfilePath As String)

        '  LOAD  APPLICATION HOST CONFIG 


        Dim sULWFKey As String = ""
        Dim username As String = ""
        Dim workflow As String = ""

        Dim workflowKey As String = ""
        Dim sCmdLineArgds As String = ""




        ' Logger.WriteToLogFile("InHost", "LoadConfig Starts")





        If ProfilePath = String.Empty AndAlso Not IOManager.CheckFileExists(ProfilePath) Then
            Throw New SympraxisException.SettingsException(String.Concat("Profile is not exists in ", ProfilePath))


        End If

        If IsNothing(configMgr) Then
            configMgr = New Common.ConfigurationManager(ProfilePath)
            If Not IsNothing(configMgr) AndAlso configMgr.errorMessage <> "" Then
                Throw New SympraxisException.SettingsException(configMgr.errorMessage)
                configMgr.errorMessage = ""
            End If
        Else
            configMgr.errorMessage = ""
            configMgr.LoadConfig(ProfilePath)
            If configMgr.errorMessage <> "" Then
                Throw New SympraxisException.SettingsException(configMgr.errorMessage)
                configMgr.errorMessage = ""
            End If
        End If
        If Not IsNothing(configMgr) Then

            AppHostConfigxml = configMgr.GetConfig("ApplicationHost", GetType(AppHostConfig))
            If configMgr.errorMessage <> "" Then

                Throw New SympraxisException.SettingsException(configMgr.errorMessage)
                configMgr.errorMessage = ""
            End If

            If Not AppHostConfigxml.AppLogger Is Nothing AndAlso AppHostConfigxml.AppLogger.Value.ToString.ToUpper = "TRUE" Then

                ''IsLogin = True

                Logger.IslogEnabled = True
                Logger.AllowedProcess = AppHostConfigxml.AppLogger.AllowedProcess
                Logger.XcludeProcess = AppHostConfigxml.AppLogger.XcludeProcess
                Logger.CurrentProcess = intrHostProcessName
                Logger.LoggerKey = AppHostConfigxml.AppLogger.LoggerKey
                Logger.xLoggerKey = AppHostConfigxml.AppLogger.xLoggerKey
                CreateLog()
            End If

            'Changing Profile path for workflow


            If Not IsNothing(AppHostConfigxml) Then
                ValidateConfigandAssignGlobalVariables(AppHostConfigxml)


            Else
                Throw New Exception("AppHostConfigxml is Nothing ")

            End If

            If Not IsNothing(intrHostWorkFlowName) AndAlso intrHostWorkFlowName <> "" Then
                Dim strarry As String() = ProfilePath.Split("\")
                ProfilePath = ProfilePath.Replace(strarry(strarry.Length - 1).ToString(), intrHostWorkFlowName.ToUpper() + ".SXML")
                If ProfilePath = String.Empty AndAlso Not IOManager.CheckFileExists(ProfilePath) Then
                    Throw New SympraxisException.SettingsException(String.Concat("Profile is not exists in ", ProfilePath))

                    Exit Sub
                End If
                configMgr = Nothing
                If IsNothing(configMgr) Then
                    configMgr = New Common.ConfigurationManager(ProfilePath)
                    If configMgr.errorMessage <> "" Then

                        Throw New SympraxisException.SettingsException(configMgr.errorMessage)
                        configMgr.errorMessage = ""
                    End If
                Else
                    configMgr.errorMessage = ""
                    configMgr.LoadConfig(ProfilePath)
                    If configMgr.errorMessage <> "" Then

                        Throw New SympraxisException.SettingsException(configMgr.errorMessage)
                        configMgr.errorMessage = ""
                    End If
                End If

            End If
            overrideAppHostConfig = configMgr.GetConfig("ApplicationHost", GetType(OverrideAppHostConfig))
            If configMgr.errorMessage <> "" Then

                Throw New SympraxisException.SettingsException(configMgr.errorMessage)
                configMgr.errorMessage = ""
            End If
            WIConfig = configMgr.GetConfig("WorkInstruction", GetType(WIConfig))
            If configMgr.errorMessage <> "" Then

                Throw New SympraxisException.SettingsException(configMgr.errorMessage)
                configMgr.errorMessage = ""
            End If
            Logger.WriteToLogFile("InHost", "Load Plugin Config Starts")
            LoadPluginsConfig(loadPluginsConfigxml)


            Logger.WriteToLogFile("InHost", "Load Plugin Config Ends")
            AddWIHotKeys()


            Logger.WriteToLogFile("InHost", "Load Plugin Config Starts")
        Else
            Throw New SympraxisException.SettingsException("Pls check whether the arguments are available in below order Optional[<Mode>] <ConfigurationFile>  Optional[<WorkFlow>]  Optional[<Process>]")

        End If








    End Sub
    Private Sub RefreshlocOfflineGrid()
        Dim Xdoc As XmlDocument = Nothing
        If IOManager.CheckFileExists(appHostConfigXml.OfflinePath & "\OfflineIndex.SXML") Then
            Xdoc = New XmlDocument
            Xdoc.Load(appHostConfigXml.OfflinePath & "\OfflineIndex.SXML")


            Xdoc.SelectSingleNode("//REFERRAL_DATA/RL[@WPID='" & objWCF.Workpacketid & "']/@ST").InnerXml = "Completed"
            Xdoc.Save(appHostConfigXml.OfflinePath & "\OfflineIndex.SXML")
            DAXObject.SetValue("SYSREFDATA", Xdoc)
        End If


    End Sub
    Private Sub OffileAutoOpen()


        Dim Xdoc As XmlDocument = Nothing
        Dim workFlowName As String = String.Empty
        Dim processName As String = String.Empty
        Dim LocalPath As String = String.Empty
        Dim unZipPath As String = String.Empty
        Dim Errmsg As String = String.Empty
        If Not IsNothing(appHostConfigXml.OfflinePath) AndAlso appHostConfigXml.OfflinePath <> "" Then

            If IOManager.CheckFileExists(appHostConfigXml.OfflinePath & "\OfflineIndex.SXML") = True Then
                Xdoc = New XmlDocument



                Xdoc.Load(appHostConfigXml.OfflinePath & "\OfflineIndex.SXML")
                Dim nodelist As XmlNodeList
                nodelist = Xdoc.SelectNodes("//RL")
                If nodelist.Count > 0 Then

                    Dim WPIDid As String = String.Empty
                    For Each node As XmlElement In nodelist
                        If node.Attributes("ST").Value.ToLower() = "completed" Then
                            WPIDid = node.Attributes("WPID").Value
                            workFlowName = node.Attributes("WF").Value
                            processName = node.Attributes("PN").Value
                            LocalPath = appHostConfigXml.OfflinePath & "\" & workFlowName & "\" & configMgr.AppbaseConfig.EnvironmentVariables.sUserId & "\" & processName & "\" & WPIDid & "\" & WPIDid & ".txt"
                            unZipPath = appHostConfigXml.OfflinePath & "\" & workFlowName & "\" & configMgr.AppbaseConfig.EnvironmentVariables.sUserId & "\" & processName & "\" & WPIDid & ".zip"
                            Exit For
                        End If
                    Next

                    If IsNothing(objWinZip) Then
                        objWinZip = New WinZip
                    End If



                    Dim parent As String = System.IO.Path.GetDirectoryName(unZipPath)

                    If IOManager.CheckFileExists(parent & "/" & WPIDid) Then
                        IOManager.DeleteFile(parent & "/" & WPIDid)
                    End If

                    If IOManager.CheckFileExists(unZipPath) Then
                        objWinZip.UnZIP(unZipPath, parent & "/" & WPIDid, parent & "/" & WPIDid, 0, Errmsg)
                        IOManager.DeleteFile(unZipPath)
                        

                    Else
                        Exit Sub


                    End If

                    If workFlowName <> "" AndAlso processName <> "" AndAlso LocalPath <> "" Then

                        intrHostProcessName = processName
                        intrHostWorkFlowName = workFlowName
                        appStartBy = processName
                        Dim Irc As Integer = 0
                        '  Dim ErrMsg As String




                        workingMode = "normal"


                        LoadConfig(appHostConfigXml, loadPluginsConfigXml, True)

                        EnvironmentVariable()



                        Irc = objWCF.Init(configMgr, intrHostProcessName, Errmsg)

                        objWCF.GetWorkPackets4Debug(LocalPath, Irc, Errmsg)


                        If Irc <> 0 Then
                            Throw New SympraxisException.SettingsException(Errmsg)
                        End If
                        Logger.WriteToLogFile("IHost", "Init WCF Completed")


                        SetWorkitemEnvironmentvariable()


                        'If Not IsNothing(defaultPluginsConfigXml.ApplicationHost.Plugins) Then
                        '    For A As Integer = 0 To defaultPluginsConfigXml.ApplicationHost.Plugins.Count - 1
                        '        If Not loadPluginsConfigXml.ApplicationHost.Plugins.Contains(defaultPluginsConfigXml.ApplicationHost.Plugins(A)) Then
                        '            loadPluginsConfigXml.ApplicationHost.Plugins.Add(defaultPluginsConfigXml.ApplicationHost.Plugins(A))
                        '            ''HshDataPlugins.Remove(ApplicationFrameWorkConfig.Plugins(0).Plugins(0))
                        '        End If
                        '    Next
                        'End If

                        If Not IsNothing(defaultPluginsConfigXml.ApplicationHost.LayoutSettings) Then
                            For b As Integer = 0 To defaultPluginsConfigXml.ApplicationHost.LayoutSettings(0).Layouts.Count - 1
                                If Not loadPluginsConfigXml.ApplicationHost.LayoutSettings(0).Layouts.Contains(defaultPluginsConfigXml.ApplicationHost.LayoutSettings(0).Layouts(b)) Then
                                    loadPluginsConfigXml.ApplicationHost.LayoutSettings(0).Layouts.Add(defaultPluginsConfigXml.ApplicationHost.LayoutSettings(0).Layouts(b))
                                End If
                            Next
                        End If





                        CreateIOmanagerDirectory()



                        If Irc = 0 Then
                            LoadPlugins(appHostConfigXml.PluginDefinitions, loadPluginsConfigXml.ApplicationHost.Plugins, loadPluginsConfigXml.ApplicationHost.LayoutSettings)
                        End If




                        CheckInputParser()



                        ReadDatapart(fullDP)



                        LoadWorkitem(fullDP)
                        'HshGetDP = HshFullDP(WCF.Objectid)

                        ' DPValue = Nothing
                        LoadChildworkitem()

                        AHEnableMenu("Home", "Close|Exit")



                        SetApplyLayout()

                        SetMatchBanner()

                        AHEnableDisableMenu("Help", "WorkInstruction", True)
                        LoadWorkInstructions()


                        Logger.WriteToLogFile("IHost", "Show lastEvent Starts")
                        ShowlastEvent()
                        Logger.WriteToLogFile("IHost", "Show lastEvent Ends")

                        SetApplicationTitle()



                        If Not IsNothing(loadPluginsConfigXml) Then
                            If Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.Applylayout) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.Applylayout.Trim() <> String.Empty Then
                                ApplyLayout(loadPluginsConfigXml.ApplicationHost.Settings.Applylayout)
                            End If
                        Else
                            ApplyLayout(0)


                        End If

                        'AAtest

                        If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.sShowStatusBar) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.sShowStatusBar.ToUpper = "TRUE" Then
                            If Not MyDictCount.ContainsKey(intrHostProcessName) Then
                                WorkedObjectCount = 0
                                CompletedObjectCount = 0
                                MyDictCount.Add(intrHostProcessName, CompletedObjectCount.ToString & "|" & WorkedObjectCount.ToString)
                            Else
                                Dim myval As String = MyDictCount(intrHostProcessName)
                                WorkedObjectCount = CInt(myval.Split("|")(1).ToString)
                                CompletedObjectCount = CInt(myval.Split("|")(0).ToString)
                                MyDictCount.Remove(intrHostProcessName)
                                MyDictCount.Add(intrHostProcessName, CompletedObjectCount.ToString & "|" & WorkedObjectCount.ToString)
                            End If

                            mLayoutmanager.AddingStatusBarControls(True, CompletedObjectCount.ToString, WorkedObjectCount.ToString)

                        End If

                        configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True

                        RibbonWin.IsMinimized = True
                        Menuhandle()

                    End If
                End If
            End If
        End If
    End Sub


    Private Sub GoOnline()
        Dim Irc As Integer
        Dim ErrMsg As String = String.Empty

        Dim Xdoc As New XmlDocument
        Xdoc.Load(appHostConfigXml.OfflinePath & "\OfflineIndex.SXML")
        Dim WIPPath As String = Xdoc.SelectSingleNode("//RL[@WPID='" & objWCF.Workpacketid & "']/@WIPP").InnerXml.ToString()
        Dim workFlowName As String = Xdoc.SelectSingleNode("//RL[@WPID='" & objWCF.Workpacketid & "']/@WF").Value
        Dim processName As String = Xdoc.SelectSingleNode("//RL[@WPID='" & objWCF.Workpacketid & "']/@PN").Value

        GoOnline(WIPPath, Irc, ErrMsg)


        Dim Path As String = appHostConfigXml.OfflinePath & "\" & workFlowName & "\" & configMgr.AppbaseConfig.EnvironmentVariables.sUserId & "\" & processName & "\" & objWCF.Workpacketid
        If Irc = 0 Then

            IOManager.DeleteDirectory(WIPPath)

            IOManager.CopyDirectory(Path & "\WIP", WIPPath)

            Dim Rmovenode As XmlNode = Xdoc.SelectSingleNode("//RL[@WPID='" & objWCF.Workpacketid & "']")
            Xdoc.DocumentElement.RemoveChild(Rmovenode)
            Xdoc.Save(appHostConfigXml.OfflinePath & "\OfflineIndex.SXML")

            IOManager.DeleteFile(Path & ".zip")

            IOManager.DeleteDirectory(Path)
            RefreshOfflineGrid()
        End If


    End Sub


    Public Sub OpenOfflineWp(ByVal ProfilePath As String, ByVal sOfflinePath As String, ByVal LocalPath As String, ByVal sWorkFlowName As String, ByVal sProcessname As String, ByRef Irc As Integer, ByRef ErrMsg As String) Implements IPluginhost.OpenOfflineWp
        Try
            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)
            configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = False
            Dim Profile As String = String.Empty

            OfflineWPPath = LocalPath

            OfflinePath = sOfflinePath
            If ProfilePath <> String.Empty Then
                Profile = ProfilePath + "Profile\AHS.SXML"


                appStartBy = sProcessname
                intrHostProcessName = sProcessname
                intrHostWorkFlowName = sWorkFlowName


                workingMode = "Offline"

                'BtnOffline.visible = Visibility.Collapsed

                OfflineLoadConfig(appHostConfigXml, loadPluginsConfigXml, Profile)

                EnvironmentVariable()

                Irc = objWCF.Init(configMgr, intrHostProcessName, ErrMsg)

                objWCF.GetWorkPackets4Debug(LocalPath, Irc, ErrMsg)

                UpdateEnvironmentvariable()

                If Irc <> 0 Then
                    Throw New SympraxisException.SettingsException(ErrMsg)
                End If
                Logger.WriteToLogFile("IHost", "Init WCF Completed")





                'If Not IsNothing(defaultPluginsConfigXml.ApplicationHost.Plugins) Then
                '    For A As Integer = 0 To defaultPluginsConfigXml.ApplicationHost.Plugins.Count - 1
                '        If Not loadPluginsConfigXml.ApplicationHost.Plugins.Contains(defaultPluginsConfigXml.ApplicationHost.Plugins(A)) Then
                '            loadPluginsConfigXml.ApplicationHost.Plugins.Add(defaultPluginsConfigXml.ApplicationHost.Plugins(A))
                '            ''HshDataPlugins.Remove(ApplicationFrameWorkConfig.Plugins(0).Plugins(0))
                '        End If
                '    Next
                'End If

                If Not IsNothing(defaultPluginsConfigXml.ApplicationHost.LayoutSettings) Then
                    For b As Integer = 0 To defaultPluginsConfigXml.ApplicationHost.LayoutSettings(0).Layouts.Count - 1
                        If Not loadPluginsConfigXml.ApplicationHost.LayoutSettings(0).Layouts.Contains(defaultPluginsConfigXml.ApplicationHost.LayoutSettings(0).Layouts(b)) Then
                            loadPluginsConfigXml.ApplicationHost.LayoutSettings(0).Layouts.Add(defaultPluginsConfigXml.ApplicationHost.LayoutSettings(0).Layouts(b))
                        End If
                    Next
                End If


                objWCF.WIPPath = Path.GetDirectoryName(LocalPath)

                CreateIOmanagerDirectory()


                If Irc = 0 Then
                    LoadPlugins(appHostConfigXml.PluginDefinitions, loadPluginsConfigXml.ApplicationHost.Plugins, loadPluginsConfigXml.ApplicationHost.LayoutSettings)
                End If




                CheckInputParser()



                ReadDatapart(fullDP)



                LoadWorkitem(fullDP)
                LoadChildworkitem()
                'HshGetDP = HshFullDP(WCF.Objectid)

                ' DPValue = Nothing

                AHEnableMenu("Home", "Close|Exit")



                SetApplyLayout()

                SetMatchBanner()

                AHEnableDisableMenu("Help", "WorkInstruction", True)

                LoadWorkInstructions()


                Logger.WriteToLogFile("IHost", "Show lastEvent Starts")
                ShowlastEvent()
                Logger.WriteToLogFile("IHost", "Show lastEvent Ends")

                SetApplicationTitle()



                If Not IsNothing(loadPluginsConfigXml) Then
                    If Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.Applylayout) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.Applylayout.Trim() <> String.Empty Then
                        ApplyLayout(loadPluginsConfigXml.ApplicationHost.Settings.Applylayout)
                    End If
                Else
                    ApplyLayout(0)
                End If

                'AA

                If Not IsNothing(loadPluginsConfigXml) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings) AndAlso Not IsNothing(loadPluginsConfigXml.ApplicationHost.Settings.sShowStatusBar) AndAlso loadPluginsConfigXml.ApplicationHost.Settings.sShowStatusBar.ToUpper = "TRUE" Then
                    If Not MyDictCount.ContainsKey(intrHostProcessName) Then
                        WorkedObjectCount = 0
                        CompletedObjectCount = 0
                        MyDictCount.Add(intrHostProcessName, CompletedObjectCount.ToString & "|" & WorkedObjectCount.ToString)
                    Else
                        Dim myval As String = MyDictCount(intrHostProcessName)
                        WorkedObjectCount = CInt(myval.Split("|")(1).ToString)
                        CompletedObjectCount = CInt(myval.Split("|")(0).ToString)
                        MyDictCount.Remove(intrHostProcessName)
                        MyDictCount.Add(intrHostProcessName, CompletedObjectCount.ToString & "|" & WorkedObjectCount.ToString)
                    End If

                    mLayoutmanager.AddingStatusBarControls(True, CompletedObjectCount.ToString, WorkedObjectCount.ToString)

                End If

                configMgr.AppbaseConfig.EnvironmentVariables.WPOpenState = True

                RibbonWin.IsMinimized = True
                Menuhandle()
            End If
        Catch ex As Exception
            SetupErrorMsg(ex)
        End Try

    End Sub



    Public Sub GoOnline(ByRef WIPPath As String, ByRef Irc As Integer, ByRef Errmsg As String) Implements IPluginhost.GoOnline
       
        Try
            MethodBase = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod()
            Logger.WriteToLogFile("SAFe", "Calling MethodBase :" & MethodBase.DeclaringType.Name.ToString & "." & MethodBase.Name.ToString)
            objWCF.GoOnline(WIPPath, Irc, Errmsg)
            CloseWork()
        Catch ex As Exception
            Irc = 1
            SetupErrorMsg(ex)
        End Try
    End Sub
    Public Sub DisplayMessage(enMTG As ExceptionManager.MessageTypes, sMsg As String) Implements IPluginhost.DisplayMessage

    End Sub

    Public Function SaveWPAckMessage() As Integer Implements IPluginhost.SaveWPAckMessage

    End Function

    Public Sub WorkCenterNameUpdated(sWDlOwner As String, sWDOwnerClass As String, sFilterCriteria As String, ByRef errmsg As String, ByRef Irc As Integer) Implements IPluginhost.WorkCenterNameUpdated

    End Sub
    Public Sub GetDataElementValue(DataElementName As String, ByRef TargetValue As String, ByRef Irc As Integer, ByRef ErrMsg As String) Implements IPluginhost.GetDataElementValue
        objWCF.GetDataElementValue(DataElementName, TargetValue, Irc, ErrMsg)
    End Sub
    Public Sub UpdateDataElementValue(DataElementName As String, TargetValue As String, ByRef Irc As Integer, ByRef ErrMsg As String) Implements IPluginhost.UpdateDataElementValue
        objWCF.UpdateDataElementValue(DataElementName, TargetValue, Irc, ErrMsg)
    End Sub
    Public Sub SetEvent(EventName As String, ByRef Irc As Integer, ByRef Errmsg As String, Optional EventType As Integer = 3, Optional CustomDate As String = "") Implements IPluginhost.SetEvent
        objWCF.SetEvent(EventName, Irc, Errmsg, EventType, CustomDate)
    End Sub
    Public Sub BuildnUpdateErrorInWorkpacket(Exception As String, ByRef irc As Integer, ByRef ErrMsg As String) Implements IPluginhost.BuildnUpdateErrorInWorkpacket
        objWCF.UpdatexmlErrors(Exception, irc, ErrMsg)
    End Sub
    Public Sub CloseEvent(EventName As String, ByRef Irc As Integer, ByRef Errmsg As String) Implements IPluginhost.CloseEvent
    End Sub


    Public Sub AddChildWorkItems(WorkItemList As List(Of Workflow.API.Workpacket.WorkPacket.WFObject), ByRef IsChildworkitem As Boolean, ByRef irc As Integer) Implements IPluginhostExtented.AddChildWorkItems
        Try
            objWCF.AddChildWorkItems(WorkItemList, IsChildworkitem, irc)
        Catch ex As Exception
            irc = 1
            SetupErrorMsg(ex)
        End Try
    End Sub

    Public Sub UpdateSaveWFStatus(ByRef Status As String, ByRef Irc As String, ByRef ErrMsg As String) Implements IPluginhost.UpdateSaveWFStatus
        Try
            objWCF.UpdateSaveWFStatus(Status, Irc, ErrMsg)
        Catch ex As Exception
            Irc = 1
            SetupErrorMsg(ex)
        End Try
    End Sub





    Public Sub SwitchWorkitem(Index As Integer, ByRef Irc As Integer, ByRef Errmsg As String) Implements IPluginhostExtented.SwitchWorkitem
        Try
            objWCF.SwitchWorkitem(Index, Irc, Errmsg)

        Catch ex As Exception
            Irc = 1
            SetupErrorMsg(ex)
        End Try
    End Sub

    Public Sub GetMutantObject(ByRef MutantObject As List(Of Workflow.API.Workpacket.WorkPacket.WFObject), ByRef Irc As Integer, ByRef Errmsg As String) Implements IPluginhostExtented.GetMutantObject
        objWCF.GetMutantObject(MutantObject, Irc, Errmsg)
    End Sub

    Public Sub SetMutantObject(ByRef MutantObject As List(Of Workflow.API.Workpacket.WorkPacket.WFObject), ByRef Irc As Integer, ByRef Errmsg As String) Implements IPluginhostExtented.SetMutantObject
        objWCF.SetMutantObject(MutantObject, Irc, Errmsg)
    End Sub

    Public Sub GetCurrentWIClone(ByRef CurrentWIClone As Workflow.API.Workpacket.WorkPacket.WFObject, ByRef Irc As Integer, ByRef Errmsg As String) Implements IPluginhostExtented.GetCurrentWIClone
        objWCF.GetCurrentWIClone(CurrentWIClone, Irc, Errmsg)
    End Sub

    Public Sub GetMutantStatus(ByRef Mustatus As String, ByRef Irc As Integer, ByRef Errmsg As String) Implements IPluginhostExtented.GetMutantStatus
        objWCF.GetMutantStatus(Mustatus, Irc, Errmsg)
    End Sub

    Public Sub SetMutantStatus(ByRef Mustatus As String, ByRef Irc As Integer, ByRef Errmsg As String) Implements IPluginhostExtented.SetMutantStatus
        objWCF.SetMutantStatus(Mustatus, Irc, Errmsg)
    End Sub

    Public Sub GetChildDataPart(ByRef Dp As Dictionary(Of String, Dictionary(Of String, String)), ByRef Irc As Integer, ByRef Errmsg As String) Implements IPluginhostExtented.GetChildDataPart
        objWCF.GetChildDataPart(Dp, Irc, Errmsg)
    End Sub

   

    Public Sub ActiveWorkitemMode(ByRef Curridx As Integer, ByRef Plugin As Object, ByRef Irc As Integer, ByRef Errmsg As String) Implements IPluginhostExtented.ActiveWorkitemMode
        Dim strDpvalue As String()
        Dim strDpValues As String
        Dim Dp As Dictionary(Of String, String) = Nothing
        SetWIModefullDP = New Dictionary(Of String, String)
        Try
            ActiveWImodeIndex = Curridx
            ActiveWImodeplugin = Plugin
            objWCF.SwitchWorkitem(Curridx, Irc, Errmsg)
            If Irc = 0 AndAlso Not IsNothing(Plugin) Then
                ReadDatapart(SetWIModefullDP)


                For index As Integer = 0 To lstPluginOrder.Count - 1
                    If DirectCast(Plugin, INormalPlugin).ToString.Trim <> DirectCast(lstPluginOrder(index), INormalPlugin).ToString.Trim Then
                        strDpValues = Convert.ToString(DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.INormalPlugin).DataPart)
                        If Not IsNothing(strDpValues) AndAlso strDpValues <> String.Empty Then
                            Dp = New Dictionary(Of String, String)

                            strDpvalue = strDpValues.Split("|")

                            For i = 0 To strDpvalue.Length - 1
                                If strDpvalue(i).Substring(0, 1) = "@" Then

                                    BuildDP4IndexnRev(strDpvalue(i), Dp, SetWIModefullDP)
                                Else

                                    Dp.Add(strDpvalue(i), SetWIModefullDP(strDpvalue(i)))

                                End If

                            Next
                        ElseIf IsNothing(strDpValues) Or (Not IsNothing(strDpValues) AndAlso strDpValues = String.Empty) Then
                            Continue For
                        End If

                        Try

                            Logger.WriteToLogFile("IHost", "Plugin LoadWorkItem -->" & lstPluginOrder(index).ToString() & " Starts")
                            If Not IsNothing(TryCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                                DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrCode = 0
                            End If

                            DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.INormalPlugin).LoadWorkItem(Dp)

                            If Not IsNothing(TryCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                                If DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrCode <> 0 Then
                                    Throw New SympraxisException.WorkitemException(DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrMsg)
                                End If
                            End If
                            If Not IsNothing(TryCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                                Logger.WriteToLogFile("IHost", "Plugin WorkLoaded -->" & lstPluginOrder(index).ToString() & " Starts")
                                DirectCast(lstPluginOrder(index), Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).WorkLoaded(objWCF.Objectid, Irc, Errmsg)
                                If Irc <> 0 Then
                                    Throw New SympraxisException.WorkitemException(Errmsg)
                                End If
                                Logger.WriteToLogFile("IHost", "Plugin WorkLoaded -->" & lstPluginOrder(index).ToString() & " Ends")
                            End If
                        Catch ex As Exception
                            Throw New SympraxisException.WorkitemException(ex.Message)
                        End Try
                        Logger.WriteToLogFile("IHost", "Plugin LoadWorkItem -->" & lstPluginOrder(index).ToString() & " Ends")
                        Dp = Nothing
                    End If
                Next

                Logger.WriteToLogFile("IHost", "Set Match Banner Starts")
                SetMatchBanner()

                Logger.WriteToLogFile("IHost", "Set Match Banner Ends")

                ActiveWIMode = True

            End If
        Catch ex As Exception
            ActiveWIMode = False
            Irc = 1
            SetupErrorMsg(ex)


        Finally
            strDpvalue = Nothing
            strDpValues = Nothing
            Dp = Nothing
        End Try
    End Sub



    Public Sub DeActiveWorkitemMode(ByRef Irc As Integer, ByRef Errmsg As String) Implements IPluginhostExtented.DeActiveWorkitemMode
        Try
            Dim IsAborted As Boolean = False
            Dim Breject As Boolean = False
            Dim BReWork As Boolean = False
            Dim BNotify As Boolean = False
            If ActiveWIMode = True AndAlso Not IsNothing(ActiveWImodeplugin) Then
                If Not IsNothing(SetWIModefullDP) Then
                    objWCF.SwitchWorkitem(ActiveWImodeIndex, Irc, Errmsg)

                    If Irc <> 0 Then
                        Throw New Exception(Errmsg)
                    End If

                    Closingplugin(SetWIModefullDP, IsAborted, BNotify, ActiveWImodeplugin)

                    Irc = objWCF.Write(SetWIModefullDP, Errmsg)

                    If Irc <> 0 Then
                        Throw New Exception(Errmsg)
                    End If

                    Logger.WriteToLogFile("IHost", "UpdateSaveWFStatus Flush Data starts")
                    objWCF.UpdateSaveWFStatus("FLUSH", Irc, Errmsg)
                    Logger.WriteToLogFile("IHost", "UpdateSaveWFStatus Flush Data Ends")

                    Logger.WriteToLogFile("Inhost", "PreviousWork:Closed plugin starts")


                    For Each element As Object In lstPluginOrder
                        If DirectCast(ActiveWImodeplugin, Sympraxis.Utilities.INormalPlugin.INormalPlugin).ToString().Trim() = DirectCast(element, Sympraxis.Utilities.INormalPlugin.INormalPlugin).ToString().Trim() Then
                            Continue For
                        End If
                        Dim strDpValues As String = String.Empty
                        strDpValues = Convert.ToString(DirectCast(element, Sympraxis.Utilities.INormalPlugin.INormalPlugin).DataPart)
                        If Not IsNothing(strDpValues) AndAlso strDpValues <> String.Empty Then
                            If Not IsNothing(TryCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                                DirectCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrCode = 0
                            End If
                            DirectCast(element, INormalPlugin).Closed()
                            If Not IsNothing(TryCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin)) Then
                                If DirectCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrCode <> 0 Then
                                    Throw New SympraxisException.AppExceptionSaveWP(DirectCast(element, Sympraxis.Utilities.INormalPlugin.IExtentedPlugin).iPluginErrMsg)
                                End If
                            End If
                        End If

                    Next
                    Logger.WriteToLogFile("Inhost", "PreviousWork:Closed plugin Ends")

                End If
            End If


        Catch ex As Exception
            ActiveWIMode = False
            Irc = 1
            SetupErrorMsg(ex)
        Finally
            ActiveWIMode = False
            ActiveWImodeplugin = Nothing
            SetWIModefullDP = Nothing
        End Try
    End Sub

    Public Sub RunPredicate(ByRef clsWP As Workflow.API.Workpacket.WorkPacket, ByRef iRc As Integer, ByRef sErrMsg As String) Implements IPluginhost.RunPredicate

    End Sub

    Public Sub RemoveRibbonMenu(ByRef RibTab As Object, ByRef Irc As Integer, ByRef Errmsg As String) Implements IPluginhostExtented.RemoveRibbonMenu
        Try

            If Not IsNothing(RibTab) Then
                RibbonWin.Items.Remove(RibTab)
            End If
        Catch ex As Exception
            Irc = 1
            Errmsg = ex.Message
        End Try
    End Sub
    Private Sub BtncloseWork_PreviewMouseDown(sender As Object, e As MouseButtonEventArgs)
        Try
            AHMenuCloseWork(False)
        Catch ex As Exception
            SetupErrorMsg(ex)
        End Try
    End Sub

End Class

Public Class CrRuleList

    Public Property Crmatch As List(Of String)

End Class
Public Class PluginDetails
    Public PluginObj As Object
    Public IsWinForm As String
    Public Width As String
    Public height As String
    Public sKey As String
    Public sName As String
End Class