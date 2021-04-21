Imports System.Xml.Serialization
Imports Sympraxis.Utilities.LayoutManager
Imports Sympraxis.Common.Qualifier

Public Class AppHostConfig

    <XmlElement("ProcessType")> Public ProcessType As String
    <XmlElement("MultiWindow")> Public MultiWindow As String


    <XmlElement("PluginDefinitions")> Public PluginDefinitions As PluginDefinitions
    <XmlElement("GetWork")> Public getWorkType As String

    <XmlElement("Workflow")> Public Workflow As String
    <XmlElement("Process")> Public Process As String
    <XmlElement("IM")> Public InstanceMgr As InstanceMgr
    <XmlElement("Sticker")> Public Sticker As Sticker

    <XmlElement("WIPPath")> Public WIPPath As String

    '<XmlElement("AppSettings")> Public AppSettings As AppSettings

    <XmlElement("LocalFolder")> Public LocalFolder As String

    <XmlElement("ExitConfirm")> Public AppExitConfirm As String

    <XmlElement("ErrorWPFolder")> Public ErrorWPFolderPath As String


    <XmlElement("IdleTime")> Public IdleTime As IdleTime

    <XmlArray("ExceptionInfo")> <XmlArrayItem("E", GetType(ErrCodeVal))> Public Property ExceptionInfo As List(Of ErrCodeVal)
    <XmlElement("AppLogger")> Public AppLogger As AppLog
    '<XmlElement("UAGG")> Public UAGG As UnitAggregation
    <XmlElement("OfflinePath")> Public OfflinePath As String

    <XmlArray("ToolStripStyle"), XmlArrayItem("Style")> Public ToolStripStyle As List(Of StyleList)
    '<XmlArray("Plugins")> <XmlArrayItem("Plugin", GetType(PluginRoute))> Public Property Plugins As List(Of PluginRoute)

End Class


Public Class AppSettings
    <XmlAttribute("LocalFolder")> Public LocalFolder As String
    <XmlAttribute("ExitConfirm")> Public AppExitConfirm As String
    <XmlAttribute("ErrorWPFolder")> Public ErrorWPFolderPath As String
    <XmlAttribute("IdleTime")> Public IdleTime As Integer
End Class



Public Class ApplyLayout
    <XmlArray("SetLayout")> <XmlArrayItem("S", GetType(SetLayoutManifest))> Public Property LayoutManifest As List(Of SetLayoutManifest)
    <XmlElement("Rules")> Public Property Rules As Rules

End Class

Public Class SetLayoutManifest
    <XmlAttribute("RID")> Public Property MatchID As String
    <XmlAttribute("PN")> Public Property ProcessName As String
    <XmlAttribute("LayoutNumber")> Public Property LayoutNumber As String
End Class

Public Class InstanceMgr
    <XmlAttribute("StPort")> Public StartPort As String
    <XmlAttribute("EPort")> Public EndPort As String
    <XmlAttribute("EI")> Public ExternalInstance As String

End Class
Public Class StyleList
    <XmlAttribute("Type")> Public Type As String
    <XmlAttribute("ForeColor")> Public ForeColor As String
    <XmlAttribute("BGColor")> Public BGColor As String
    <XmlAttribute("FontSize")> Public Size As String

End Class
Public Class UnitAggregation
    <XmlAttribute("SUM")> Public Sum As String
    <XmlAttribute("MAX")> Public Max As String
End Class
Public Class AppLog
    <XmlAttribute("logPath")> Public Path As String
    <XmlAttribute("Value")> Public Value As String
    <XmlAttribute("AP")> Public AllowedProcess As String
    <XmlAttribute("LK")> Public LoggerKey As String
    <XmlAttribute("XP")> Public XcludeProcess As String
    <XmlAttribute("XLK")> Public xLoggerKey As String
End Class
Public Class ErrCodeVal
    <XmlAttribute("ErrCode")> Public ErrorCode As String
    <XmlAttribute("ErrMsg")> Public ErrorMsg As String
End Class
Public Class IdleTime
    <XmlAttribute("IT")> Public Interval As Integer
End Class
Public Class OverrideAppHostConfig

    <XmlElement("WCF")> Public WFClientFrameWork As WCFSettings
    <XmlElement("BannerSettings")> Public BannerSettings As BannerSetting
    <XmlElement("HelpFileSettings")> Public HelpFileSettings As HelpFileSetting

    <XmlElement("Overrides")> Public OverrideValue As OverrideValue
    <XmlElement("Notify")> Public PanelNotify As Notify
    <XmlElement("AHSettings")> Public AHSettings As AHSettings
    <XmlElement("TransDBURL")> Public TransDBURL As String
    <XmlElement("ARMSURL")> Public ARMSURL As String
    <XmlElement("TransDBWebURL")> Public TransDBWebURL As String
    <XmlArray("MenuConfig")> <XmlArrayItem("M", GetType(MenuConfigManifest))> Public Property MenuConfigManifest As List(Of MenuConfigManifest)
    <XmlArray("MenuShortcut")> <XmlArrayItem("Shortcut", GetType(MenuShortcutConfig))> Public Property MenuShortcuts As List(Of MenuShortcutConfig)
    <XmlElement("ApplyLayout")> Public ApplyLayoutConfig As ApplyLayout

End Class
Public Class AHSettings

    <XmlAttribute("AppAutocloseInterval")> Public AppHostAutocloseInterval As String
    <XmlAttribute("AddDomain")> Public AddDomain As String
End Class

Public Class Notify
    <XmlAttribute("XPath")> Public Path As String
    <XmlAttribute("Height")> Public Height As String
    <XmlAttribute("NotificationSound")> Public sNotificationSound As String
    <XmlAttribute("SoundAlertPath")> Public sSoundAlertPath As String

End Class

Public Class WCFSettings
    <XmlAttribute("Name")> Public Property Name As String
    <XmlAttribute("Assembly")> Public Property Assembly As String
    <XmlAttribute("PF")> Public Property PathFolder As String
    <XmlAttribute("Class")> Public Property ClassName As String
    <XmlAttribute("TY")> Public Property Type As String


    <XmlElement("ServiceConfig")> Public ServiceConfig As String

    <XmlElement("DataElementPath")> Public Property DataElementPath As String
    <XmlElement("QualifierPath")> Public Property QualifierPath As String
    <XmlElement("ValueSetPath")> Public Property ValueSetPath As String

    <XmlElement("SetIndexPath")> Public Property SetIndexPath As String
    <XmlElement("SetReviewPath")> Public Property SetReviewPath As String
    <XmlElement("SetInstructionPath")> Public Property SetInstructionPath As String
    <XmlElement("SetDataPath")> Public Property SetDataPath As String
    <XmlElement("SetActionPath")> Public Property SetActionPath As String
End Class

Public Class BannerSetting

    <XmlArray("SetBanner")> <XmlArrayItem("R", GetType(SetBannerManifest))> Public Property BannerManifest As List(Of SetBannerManifest)

    <XmlElement("Rules")> Public Property Rules As Rules

    <XmlArray("Banners")> <XmlArrayItem("Banner", GetType(Banner))> Public Property BannerMeta As List(Of Banner)

End Class


Public Class HelpFileSetting
    <XmlArray("SetHelpfile")> <XmlArrayItem("R", GetType(SetHelpFileManifest))> Public Property SetHelpFileManifest As List(Of SetHelpFileManifest)

End Class



Public Class SetBannerManifest
    <XmlAttribute("RID")> Public Property MatchID As String
    <XmlAttribute("PN")> Public Property ProcessName As String
    <XmlAttribute("NM")> Public Property BannerName As String
    <XmlAttribute("CS")> Public Property CaptureSequence As Integer


End Class

Public Class SetHelpFileManifest

    <XmlAttribute("PN")> Public Property ProcessName As String
    <XmlAttribute("HTMLFileName")> Public Property HTMLFileName As String



    Public Sub New()

    End Sub
End Class

Public Class MenuConfigManifest
    <XmlAttribute("PN")> Public Property ProcessName As String
    <XmlAttribute("Menu")> Public Property Menu As String
    <XmlAttribute("HideMenu")> Public Property HMenu As String
End Class
Public Class MenuShortcutConfig
    <XmlAttribute("MenuName")> Public MenuName As String
    <XmlAttribute("Key")> Public key As String
End Class


Public Class Banner
    <XmlElement("F")> Public F As List(Of Fields)
    Private _EnableMode As String
    <XmlAttribute("EN")> _
    Public Property EnableMode() As String
        Get
            Return _EnableMode
        End Get
        Set(ByVal value As String)
            _EnableMode = value
        End Set
    End Property
    Private _Name As String
    <XmlAttribute("Name")> _
    Public Property Name() As String
        Get
            Return _Name
        End Get
        Set(ByVal value As String)
            _Name = value
        End Set
    End Property
    Private _XamlPath As String
    <XmlAttribute("XAML")> _
    Public Property XamlPath() As String
        Get
            Return _XamlPath
        End Get
        Set(ByVal value As String)
            _XamlPath = value
        End Set
    End Property
    Private _Height As String
    <XmlAttribute("Height")> _
    Public Property Height() As String
        Get
            Return _Height
        End Get
        Set(ByVal value As String)
            _Height = value
        End Set
    End Property
End Class

Public Class Fields
    Private _SXP As String
    <XmlAttribute("SXP")> _
    Public Property SXP() As String
        Get
            Return _SXP
        End Get
        Set(ByVal value As String)
            _SXP = value
        End Set
    End Property
    Private _TXP As String
    <XmlAttribute("TXP")> _
    Public Property TXP() As String
        Get
            Return _TXP
        End Get
        Set(ByVal value As String)
            _TXP = value
        End Set
    End Property
    Private _Type As String
    <XmlAttribute("Type")> _
    Public Property Type() As String
        Get
            Return _Type
        End Get
        Set(ByVal value As String)
            _Type = value
        End Set
    End Property
End Class

'Public Class ApplicationHost

'    <XmlElement("Implement")> Public Implement As String
'    <XmlElement("WCF")> Public Property WFClientFrameWork As PluginDefinition
'    <XmlElement("PluginDefinitions")> Public Plugins As PluginDefinitions
'    <XmlElement("GetWork")> Public Property GetWorkType As String

'    <XmlElement("Workflow")> Public Workflow As String
'    <XmlElement("Process")> Public Process As String

'    <XmlElement("Overrides")> Public OverrideValue As OverrideValue

'End Class

Public Class Sticker
    <XmlAttribute("StickerFolder")> Public StickerFolder As String
    <XmlAttribute("DefaultStickerName")> Public DefaultStickerName As String

    <XmlAttribute("Poweredby")> Public Poweredby As String
    <XmlAttribute("AppLogo")> Public AppLogo As String
    <XmlAttribute("StickerWidth")> Public StickerWidth As String
    <XmlElement("S", GetType(WGStrickers))> Public Stickers As List(Of WGStrickers)
End Class
Public Class WGStrickers
    <XmlAttribute("WG")> Public WG As String
    <XmlAttribute("StickerName")> Public StickerName As String
    <XmlAttribute("StickerWidth")> Public stickerWidth As String
End Class

Public Class OverrideValue
    <XmlElement("ResourceFolder")> Public ResourceFolder As ResourceFolder
End Class

Public Class ResourceFolder
    <XmlElement("MenuIcon")> Public MenuIcon As MenuIcon
End Class
Public Class MenuIcon
    <XmlAttribute("Path")> Public Property iconPath As String
End Class
Public Class PluginDefinitions
    <XmlElement("PD", GetType(PluginDefinition))> Public Property PluginDefinition As List(Of PluginDefinition)

    <XmlAttribute("PluginFolder")> Public PluginFolder As String


End Class

Public Class PluginDefinition
    <XmlAttribute("Name")> Public Property Name As String
    <XmlAttribute("Assembly")> Public Property Assembly As String
    <XmlAttribute("PF")> Public Property PathFolder As String
    <XmlAttribute("Class")> Public Property ClassName As String
    <XmlAttribute("TY")> Public Property Type As String
    <XmlAttribute("IsWinControl")> Public Property IsWinControl As String



End Class
