Imports System.Web.Services

<System.Web.Services.WebService(Namespace:="http://tempuri.org/WebDevelopmentProjectWS/WebDevelopmentProjectWebService")> _
Public Class WebDevelopmentProjectWebService
   Inherits System.Web.Services.WebService

#Region " Web Services Designer Generated Code "

   Public Sub New()
      MyBase.New()

      'This call is required by the Web Services Designer.
      InitializeComponent()

      'Add your own initialization code after the InitializeComponent() call

   End Sub

   'Required by the Web Services Designer
   Private components As System.ComponentModel.IContainer

   'NOTE: The following procedure is required by the Web Services Designer
   'It can be modified using the Web Services Designer.  
   'Do not modify it using the code editor.
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      components = New System.ComponentModel.Container
   End Sub

   Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
      'CODEGEN: This procedure is required by the Web Services Designer
      'Do not modify it using the code editor.
      If disposing Then
         If Not (components Is Nothing) Then
            components.Dispose()
         End If
      End If
      MyBase.Dispose(disposing)
   End Sub

#End Region

   <WebMethod()> _
   Public Function GetInitTime() As DateTime
      Dim comObject As New COMTest.VB6COMClass

      Return comObject.GetInitTime
      comObject = Nothing
   End Function
End Class