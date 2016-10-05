Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
   Friend WithEvents testButton As System.Windows.Forms.Button
   Friend WithEvents testLabel As System.Windows.Forms.Label
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Me.testButton = New System.Windows.Forms.Button
      Me.testLabel = New System.Windows.Forms.Label
      Me.SuspendLayout()
      '
      'testButton
      '
      Me.testButton.Location = New System.Drawing.Point(139, 197)
      Me.testButton.Name = "testButton"
      Me.testButton.TabIndex = 0
      Me.testButton.Text = "Test"
      '
      'testLabel
      '
      Me.testLabel.Location = New System.Drawing.Point(7, 7)
      Me.testLabel.Name = "testLabel"
      Me.testLabel.Size = New System.Drawing.Size(339, 100)
      Me.testLabel.TabIndex = 1
      Me.testLabel.Text = "Click the Test button to test the application."
      Me.testLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
      '
      'Form1
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
      Me.ClientSize = New System.Drawing.Size(353, 260)
      Me.Controls.Add(Me.testLabel)
      Me.Controls.Add(Me.testButton)
      Me.Name = "Form1"
      Me.Text = "AutoDeploymentApplication"
      Me.ResumeLayout(False)

   End Sub

#End Region

   Private Sub testButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles testButton.Click
      testLabel.Text = "This Windows application was invoked from " & _
         AppDomain.CurrentDomain.FriendlyName & "." & Constants.vbCrLf & "Version: " & _
         Application.ProductVersion
   End Sub
End Class
