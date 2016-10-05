Imports System.Globalization
Imports System.Threading
Imports System.Resources

Public Class WebForm1
   Inherits System.Web.UI.Page

   Protected m_pageResourceManager As ResourceManager

#Region " Web Form Designer Generated Code "

   'This call is required by the Web Form Designer.
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

   End Sub
   Protected WithEvents clickMeButton As System.Web.UI.WebControls.Button
   Protected WithEvents initAssemblyTimeLabel As System.Web.UI.WebControls.Label
   Protected WithEvents staticImage As System.Web.UI.WebControls.Image
   Protected WithEvents dynamicImage As System.Web.UI.WebControls.Image
   Protected WithEvents invokeWebServiceTimeLabel As System.Web.UI.WebControls.Label
   Protected WithEvents daLangAnchor As System.Web.UI.HtmlControls.HtmlAnchor
   Protected WithEvents enLangAnchor As System.Web.UI.HtmlControls.HtmlAnchor
   Protected WithEvents languageInput As System.Web.UI.HtmlControls.HtmlInputHidden

   'NOTE: The following placeholder declaration is required by the Web Form Designer.
   'Do not delete or move it.
   Private designerPlaceholderDeclaration As System.Object

   Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
      'CODEGEN: This method call is required by the Web Form Designer
      'Do not modify it using the code editor.
      InitializeComponent()
   End Sub

#End Region

   Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      m_pageResourceManager = New ResourceManager("WebDeploymentProject.Strings", _
         System.Reflection.Assembly.GetExecutingAssembly)

      If IsPostBack Then
         ChangePageCulture(languageInput.Value)
      Else
         ChangePageCulture("en")
         languageInput.Value = "en"
      End If

      clickMeButton.Text = m_pageResourceManager.GetString("CLICKMEBUTTON")
      clickMeButton.ToolTip = m_pageResourceManager.GetString("CLICKMEBUTTON_TOOLTIP")
      initAssemblyTimeLabel.Text = m_pageResourceManager.GetString("INITASSEMBLYTIMELABEL")
      invokeWebServiceTimeLabel.Text = m_pageResourceManager.GetString("INVOKEWEBSERVICETIMELABEL")
   End Sub

   Private Sub clickMeButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clickMeButton.Click
      Dim wsObject As New WebDevelopmentProjectWebService.WebDevelopmentProjectWebService
      Dim assemblyObject As New TestAssembly1.AssemblyTest

      ' Load and display dynamic picture
      dynamicImage.ImageUrl = "Author.jpg"

      ' Update labels
      initAssemblyTimeLabel.Text = m_pageResourceManager.GetString("INITASSEMBLYTIMELABEL") & _
         assemblyObject.GetInitTime.ToShortTimeString
      invokeWebServiceTimeLabel.Text = m_pageResourceManager.GetString("INVOKEWEBSERVICETIMELABEL") & _
         wsObject.GetInitTime.ToShortTimeString
   End Sub

   Protected Overridable Sub daLangAnchor_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles daLangAnchor.ServerClick
      ' Change to Danish culture
      ChangePageCulture("da")
   End Sub

   Protected Overridable Sub enLangAnchor_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles enLangAnchor.ServerClick
      ' Change to English culture
      ChangePageCulture("en")
   End Sub

   Private Sub ChangePageCulture(ByVal cultureName As String)
      Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture(cultureName)
      Thread.CurrentThread.CurrentUICulture = New CultureInfo(cultureName)
   End Sub
End Class