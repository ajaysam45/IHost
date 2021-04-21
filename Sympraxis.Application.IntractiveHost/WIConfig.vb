Imports System.Xml.Serialization
Imports Sympraxis.Common.Qualifier

<Serializable()> _
<XmlRoot("WorkInstruction")> _
Public Class WIConfig
    <XmlElement("StandardInstruction")> Public Property StandardInstruction As InstructionProperty
    <XmlElement("DynamicInstruction")> Public Property DynamicInstruction As DynamicSetting
    <XmlElement("ContextInstruction")> Public Property ContextInstruction As ContextSetting
End Class

<Serializable()> _
Public Class DynamicSetting
    Inherits InstructionProperty
    <XmlElement("Rules")> Public Property Rules As Rules
    <XmlArray("WorkIns"), XmlArrayItem("WI", GetType(DyanamicWorkIns))> _
    Public DWIns As List(Of DyanamicWorkIns)
End Class

<Serializable()> _
Public Class ContextSetting
    Inherits InstructionProperty
    <XmlElement("Rules")> Public Property Rules As Rules
    <XmlArray("WorkIns"), XmlArrayItem("WI", GetType(ContextWorkIns))> _
    Public CWIns As List(Of ContextWorkIns)
End Class


Public Class StandardSetting
    Private _SWIns As List(Of DyanamicWorkIns)
    Public Property SWIns() As List(Of DyanamicWorkIns)
        Get
            Return _SWIns
        End Get
        Set(ByVal value As List(Of DyanamicWorkIns))
            _SWIns = value
        End Set
    End Property
End Class




<Serializable()> _
Public Class DyanamicWorkIns

    Private _Name As String
    <XmlAttribute("NM")> _
    Public Property dName() As String
        Get
            Return _Name
        End Get
        Set(ByVal value As String)
            _Name = value
        End Set
    End Property

    Private _Header As String
    <XmlAttribute("HDR")> _
    Public Property dHeader() As String
        Get
            Return _Header
        End Get
        Set(ByVal value As String)
            _Header = value
        End Set
    End Property

    Private _Body As String
    <XmlAttribute("BY")> _
    Public Property dBody() As String
        Get
            Return _Body
        End Get
        Set(ByVal value As String)
            _Body = value
        End Set
    End Property

End Class

<Serializable()> _
Public Class ContextWorkIns

    Private _Name As String
    <XmlAttribute("NM")> _
    Public Property cName() As String
        Get
            Return _Name
        End Get
        Set(ByVal value As String)
            _Name = value
        End Set
    End Property

    Private _KY As String
    <XmlAttribute("KY")> _
    Public Property cKey() As String
        Get
            Return _KY
        End Get
        Set(ByVal value As String)
            _KY = value
        End Set
    End Property

    Private _CTX As String
    <XmlAttribute("CTX")> _
    Public Property Context() As String
        Get
            Return _CTX
        End Get
        Set(ByVal value As String)
            _CTX = value
        End Set
    End Property

    Private _Header As String
    <XmlAttribute("HDR")> _
    Public Property cHeader() As String
        Get
            Return _Header
        End Get
        Set(ByVal value As String)
            _Header = value
        End Set
    End Property

    Private _Body As String
    <XmlAttribute("BY")> _
    Public Property cBody() As String
        Get
            Return _Body
        End Get
        Set(ByVal value As String)
            _Body = value
        End Set
    End Property

End Class

<Serializable()> _
Public Class InstructionProperty

    Private _Title As String
    <XmlAttribute("Title")> _
    Public Property iTitle() As String
        Get
            Return _Title
        End Get
        Set(ByVal value As String)
            _Title = value
        End Set
    End Property

    Private _BackColor As String = "#FFF5EB80"
    <XmlAttribute("BackColor")> _
    Public Property iBackColor() As String
        Get
            Return _BackColor
        End Get
        Set(ByVal value As String)
            _BackColor = value
        End Set
    End Property

    Private _ForeColor As String = "black"
    <XmlAttribute("ForeColor")> _
    Public Property iForeColor() As String
        Get
            Return _ForeColor
        End Get
        Set(ByVal value As String)
            _ForeColor = value
        End Set
    End Property

    Private _HeaderColor As String = "black"
    <XmlAttribute("HeaderColor")> _
    Public Property HeaderColor() As String
        Get
            Return _HeaderColor
        End Get
        Set(ByVal value As String)
            _HeaderColor = value
        End Set
    End Property

    Private _SKey As String
    <XmlAttribute("SKey")> _
    Public Property iSKey() As String
        Get
            Return _SKey
        End Get
        Set(ByVal value As String)
            _SKey = value
        End Set
    End Property

    Private _AWI As String
    <XmlAttribute("AWI")> _
    Public Property iAWI() As String
        Get
            Return _AWI
        End Get
        Set(ByVal value As String)
            _AWI = value
        End Set
    End Property

    Private _Interval As Integer
    <XmlAttribute("Interval")> _
    Public Property iInterval() As Integer
        Get
            Return _Interval
        End Get
        Set(ByVal value As Integer)
            _Interval = value
        End Set
    End Property
End Class


<Serializable(), XmlRoot("WI")> _
Public Class Instructions
    Private _INS As List(Of INS)
    <XmlElement("INS")> _
    Public Property INS() As List(Of INS)
        Get
            Return _INS
        End Get
        Set(ByVal value As List(Of INS))
            _INS = value
        End Set
    End Property


End Class

Public Class INS
    Private _LN As String
    <XmlAttribute("LN")> _
    Public Property LN() As String
        Get
            Return _LN
        End Get
        Set(ByVal value As String)
            _LN = value
        End Set
    End Property
    Private _Fcolor As String
    <XmlAttribute("Fcolor")> _
    Public Property Fcolor() As String
        Get
            Return _Fcolor
        End Get
        Set(ByVal value As String)
            _Fcolor = value
        End Set
    End Property
    Private _FSize As String
    <XmlAttribute("FSize")> _
    Public Property FSize() As String
        Get
            Return _FSize
        End Get
        Set(ByVal value As String)
            _FSize = value
        End Set
    End Property
    'Private _FFamily As String
    '<XmlAttribute("FFamily")> _
    'Public Property FFamily() As String
    '    Get
    '        Return _FFamily
    '    End Get
    '    Set(ByVal value As String)
    '        _FFamily = value
    '    End Set
    'End Property
    Private _BColor As String
    <XmlAttribute("BColor")> _
    Public Property BColor() As String
        Get
            Return _BColor
        End Get
        Set(ByVal value As String)
            _BColor = value
        End Set
    End Property
End Class
