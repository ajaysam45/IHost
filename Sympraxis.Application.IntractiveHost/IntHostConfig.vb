Imports System.Xml.Serialization
Imports Sympraxis.Utilities.LayoutManager
Imports Sympraxis.Common.Qualifier


Public Class IntHostConfig

    <XmlElement("AH")> Public Property ApplicationHost As ConfigApplicationHost
    
End Class
 
 
Public Class ConfigApplicationHost

    <XmlElement("ClosingRules")> Public ClosingRules As ClosingRules
    <XmlElement("Parser")> Public Parser As ParserClass
    <XmlElement("Settings")> Public Settings As Setting
    <XmlArray("ArrangePlugins")> <XmlArrayItem("Plugin", GetType(PluginRoute))> Public Property Plugins As List(Of PluginRoute)
    <XmlElement("LayoutSetting")> Public LayoutSettings As List(Of LayoutSetting)
End Class

Public Class Setting

    <XmlAttribute("ParallelWorkItem")> Public ParallelWorkItem As String
    <XmlAttribute("SetWD")> Public SetWD As String
    <XmlAttribute("ApplicationTitle")> Public ApplicationTitle As String
    <XmlAttribute("WorkOffline")> Public WorkOffline As String
    <XmlAttribute("Applylayout")> Public Applylayout As String
    <XmlAttribute("AttendedCloseInterval")> Public AttendedInterval As String
    <XmlAttribute("UnAttendedCloseInterval")> Public UnAttendedCloseInterval As String
    <XmlAttribute("GTB")> Public GotoBack As String
    Private _AutocloseInterval As String
    <XmlAttribute("AutocloseInterval")> _
    Public Property AutocloseInterval() As String
        Get
            Return _AutocloseInterval
        End Get
        Set(ByVal value As String)
            _AutocloseInterval = value
        End Set
    End Property


    <XmlAttribute("bNotify")> Public bNotify As String
    <XmlAttribute("LoginType")> Public LoginType As String
    <XmlAttribute("AutoOpenNotify")> Public AutoOpenNotify As String
    <XmlAttribute("AutoNextNotify")> Public AutoOpenNextNotify As String
    <XmlAttribute("CSP")> Public CSP As String
    <XmlAttribute("ShowStatusBar")> Public sShowStatusBar As String

    <XmlAttribute("LGTimer")> Public LGTimer As String
    <XmlAttribute("IDTimer")> Public IDTimer As String
    <XmlAttribute("WPTimer")> Public WPTimer As String
    <XmlAttribute("ChildWFsts")> Public ChildWFsts As String
End Class

Public Class ParserClass
    <XmlElement("InputParser")> Public InputParser As Parser
    <XmlElement("ExitParser")> Public ExitParser As Parser
End Class
Public Class Parser
    <XmlAttribute("Name")> Public Name As String
    <XmlAttribute("ConfigKey")> Public ConfigKey As String
End Class
Public Class ClosingRules
    <XmlElement("Errors")> Public Errors As SetClosingRules
    <XmlElement("Alerts")> Public Alerts As SetClosingRules
    <XmlElement("Completion")> Public Completion As SetClosingRules
    <XmlElement("Rules")> Public Property Rules As Rules
End Class


Public Class SetClosingRules
    <XmlArray("SetRules")> <XmlArrayItem("R", GetType(SetClosingRuleManifest))> Public Property SetRulesval As List(Of SetClosingRuleManifest)

End Class


Public Class SetClosingRuleManifest
    <XmlAttribute("RULEID")> Public RULEID As String
    <XmlAttribute("MSG")> Public MSG As String
    <XmlAttribute("ErrorCode")> Public ErrorCode As String
End Class


Public Class PluginRoute
    <XmlAttribute("Name")> Public Property Name As String
    <XmlAttribute("ConfigKey")> Public Property ConfigKey As String
    <XmlAttribute("PNL")> Public Property Panel As String
    <XmlAttribute("Title")> Public Property Title As String
    <XmlAttribute("SKey")> Public Property SKey As String
    <XmlAttribute("IconSource")> Public Property IconSource As String
    <XmlAttribute("Width")> Public Property Width As String
    <XmlAttribute("Height")> Public Property Height As String
End Class


