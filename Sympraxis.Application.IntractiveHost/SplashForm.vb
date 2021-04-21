
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Filename     : frmSplash.vb
'
' Author       : 
'
' Company      : Firstsource Solutions Ltd.
'
' Product      : Sympraxis V5.1
'
' Solution     :
'
' Date Created : 
'
' Description  : Structured Exception Handling
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Revision History
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   File last modified
'     by $Author: $ Mahesh/KavithaP
'     on $Date:  $ 29/06/11
'     VSS revision number $Revision: $
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Revision User Log
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  $Log: $
'
'
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''





Imports System.Threading
Imports System.Windows.Forms

Public Class SplashForm

    'Shared m_SplahsScreen As frmSplash
    Shared SplashThread As Thread

    Private FormLoaded As Boolean = False

    'Public Shared Property SplashForm() As frmSplash
    '    Get
    '        Return m_SplahsScreen
    '    End Get
    '    Set(ByVal value As frmSplash)
    '        m_SplahsScreen = value
    '    End Set
    'End Property
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="splashForm"></param>
    ''' <remarks></remarks>
    Public Shared Sub InitializeSplashScreen(ByVal splashForm As SplashForm)
        SplashThread = New Thread(AddressOf InitializeForm)
        SplashThread.IsBackground = True
        SplashThread.SetApartmentState(ApartmentState.MTA)
        SplashThread.Start(splashForm)
        While splashForm.FormLoaded = False
            System.Threading.Thread.Sleep(1000)
        End While
    End Sub


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="splashForm"></param>
    ''' <remarks></remarks>
    Private Shared Sub InitializeForm(ByVal splashForm As Object)
        CType(splashForm, SplashForm).Visible = False
        Control.CheckForIllegalCrossThreadCalls = False
        CType(splashForm, SplashForm).Hide()
        CType(splashForm, SplashForm).FormLoaded = True
        System.Windows.Forms.Application.Run(CType(splashForm, SplashForm))
    End Sub



    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="splashForm"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowForm(ByRef splashForm As SplashForm)
        splashForm.Visible = True
        splashForm.Show()
    End Sub


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="splashForm"></param>
    ''' <remarks></remarks>
    Public Shared Sub CloseForm(ByRef splashForm As SplashForm)
        splashForm.Close()
        SplashThread = Nothing
        splashForm = Nothing
    End Sub

End Class