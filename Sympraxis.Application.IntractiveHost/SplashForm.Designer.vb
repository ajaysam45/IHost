Imports System.Windows.Forms
Imports Sympraxis.Utilities.Splash

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SplashForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.MySplashControl = New Sympraxis.Utilities.Splash.SplashDrawingArea()
        Me.SuspendLayout()
        '
        'MySplashControl
        '
        Me.MySplashControl._Poweredbyflg = False
        Me.MySplashControl.AppLogoPath = Nothing
        Me.MySplashControl.BackColor = System.Drawing.Color.Black
        Me.MySplashControl.Dock = System.Windows.Forms.DockStyle.Fill
        Me.MySplashControl.Font = New System.Drawing.Font("Calibri", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MySplashControl.lblCopyRight = Nothing
        Me.MySplashControl.lblRegisterTradeMark = Nothing
        Me.MySplashControl.lblSympraxisTitle = Nothing
        Me.MySplashControl.lblWorkingTogether = Nothing
        Me.MySplashControl.Location = New System.Drawing.Point(0, 0)
        Me.MySplashControl.LogoPath = Nothing
        Me.MySplashControl.LogoWidth = 0
        Me.MySplashControl.Name = "MySplashControl"
        Me.MySplashControl.Size = New System.Drawing.Size(543, 371)
        Me.MySplashControl.TabIndex = 2
        '
        'SplashForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoValidate = System.Windows.Forms.AutoValidate.Disable
        Me.ClientSize = New System.Drawing.Size(543, 371)
        Me.ControlBox = False
        Me.Controls.Add(Me.MySplashControl)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SplashForm"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.TransparencyKey = System.Drawing.Color.Black
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents MySplashControl As SplashDrawingArea

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        'Me.SetStyle(ControlStyles.DoubleBuffer, True)
        'Me.SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        'Me.SetStyle(ControlStyles.UserPaint, True)
        ' Add any initialization after the InitializeComponent() call.

    End Sub
End Class
