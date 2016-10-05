Imports System.Globalization
Imports System.Threading
Imports System.Collections.Specialized
Imports System.Resources
Imports System.Configuration

Public Class Form1
   Inherits System.Windows.Forms.Form

   Const enUS As String = "en-US"
   Const daDK As String = "da-DK"

   Private m_cultureStrings() As String = {enUS, daDK}

#Region " Windows Form Designer generated code "

   Public Sub New()
      MyBase.New()

      Randomize()
      Thread.CurrentThread.CurrentUICulture = New CultureInfo(m_cultureStrings(CInt(Rnd())))

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
   Friend WithEvents staticPicture As System.Windows.Forms.PictureBox
   Friend WithEvents dynamicPicture As System.Windows.Forms.PictureBox
   Friend WithEvents initCOMTimeLabel As System.Windows.Forms.Label
   Friend WithEvents initAssemblyTimeLabel As System.Windows.Forms.Label
   Friend WithEvents clickMeButton As System.Windows.Forms.Button
   Friend WithEvents buttonToolTip As System.Windows.Forms.ToolTip
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Me.components = New System.ComponentModel.Container
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
      Dim configurationAppSettings As System.Configuration.AppSettingsReader = New System.Configuration.AppSettingsReader
      Me.clickMeButton = New System.Windows.Forms.Button
      Me.staticPicture = New System.Windows.Forms.PictureBox
      Me.dynamicPicture = New System.Windows.Forms.PictureBox
      Me.initCOMTimeLabel = New System.Windows.Forms.Label
      Me.initAssemblyTimeLabel = New System.Windows.Forms.Label
      Me.buttonToolTip = New System.Windows.Forms.ToolTip(Me.components)
      Me.SuspendLayout()
      '
      'clickMeButton
      '
      Me.clickMeButton.AccessibleDescription = resources.GetString("clickMeButton.AccessibleDescription")
      Me.clickMeButton.AccessibleName = resources.GetString("clickMeButton.AccessibleName")
      Me.clickMeButton.Anchor = CType(resources.GetObject("clickMeButton.Anchor"), System.Windows.Forms.AnchorStyles)
      Me.clickMeButton.BackgroundImage = CType(resources.GetObject("clickMeButton.BackgroundImage"), System.Drawing.Image)
      Me.clickMeButton.Dock = CType(resources.GetObject("clickMeButton.Dock"), System.Windows.Forms.DockStyle)
      Me.clickMeButton.Enabled = CType(resources.GetObject("clickMeButton.Enabled"), Boolean)
      Me.clickMeButton.FlatStyle = CType(resources.GetObject("clickMeButton.FlatStyle"), System.Windows.Forms.FlatStyle)
      Me.clickMeButton.Font = CType(resources.GetObject("clickMeButton.Font"), System.Drawing.Font)
      Me.clickMeButton.Image = CType(resources.GetObject("clickMeButton.Image"), System.Drawing.Image)
      Me.clickMeButton.ImageAlign = CType(resources.GetObject("clickMeButton.ImageAlign"), System.Drawing.ContentAlignment)
      Me.clickMeButton.ImageIndex = CType(resources.GetObject("clickMeButton.ImageIndex"), Integer)
      Me.clickMeButton.ImeMode = CType(resources.GetObject("clickMeButton.ImeMode"), System.Windows.Forms.ImeMode)
      Me.clickMeButton.Location = CType(resources.GetObject("clickMeButton.Location"), System.Drawing.Point)
      Me.clickMeButton.Name = "clickMeButton"
      Me.clickMeButton.RightToLeft = CType(resources.GetObject("clickMeButton.RightToLeft"), System.Windows.Forms.RightToLeft)
      Me.clickMeButton.Size = CType(resources.GetObject("clickMeButton.Size"), System.Drawing.Size)
      Me.clickMeButton.TabIndex = CType(resources.GetObject("clickMeButton.TabIndex"), Integer)
      Me.clickMeButton.Text = resources.GetString("clickMeButton.Text")
      Me.clickMeButton.TextAlign = CType(resources.GetObject("clickMeButton.TextAlign"), System.Drawing.ContentAlignment)
      Me.buttonToolTip.SetToolTip(Me.clickMeButton, resources.GetString("clickMeButton.ToolTip"))
      Me.clickMeButton.Visible = CType(resources.GetObject("clickMeButton.Visible"), Boolean)
      '
      'staticPicture
      '
      Me.staticPicture.AccessibleDescription = resources.GetString("staticPicture.AccessibleDescription")
      Me.staticPicture.AccessibleName = resources.GetString("staticPicture.AccessibleName")
      Me.staticPicture.Anchor = CType(resources.GetObject("staticPicture.Anchor"), System.Windows.Forms.AnchorStyles)
      Me.staticPicture.BackgroundImage = CType(resources.GetObject("staticPicture.BackgroundImage"), System.Drawing.Image)
      Me.staticPicture.Dock = CType(resources.GetObject("staticPicture.Dock"), System.Windows.Forms.DockStyle)
      Me.staticPicture.Enabled = CType(resources.GetObject("staticPicture.Enabled"), Boolean)
      Me.staticPicture.Font = CType(resources.GetObject("staticPicture.Font"), System.Drawing.Font)
      Me.staticPicture.Image = CType(resources.GetObject("staticPicture.Image"), System.Drawing.Image)
      Me.staticPicture.ImeMode = CType(resources.GetObject("staticPicture.ImeMode"), System.Windows.Forms.ImeMode)
      Me.staticPicture.Location = CType(resources.GetObject("staticPicture.Location"), System.Drawing.Point)
      Me.staticPicture.Name = "staticPicture"
      Me.staticPicture.RightToLeft = CType(resources.GetObject("staticPicture.RightToLeft"), System.Windows.Forms.RightToLeft)
      Me.staticPicture.Size = CType(resources.GetObject("staticPicture.Size"), System.Drawing.Size)
      Me.staticPicture.SizeMode = CType(resources.GetObject("staticPicture.SizeMode"), System.Windows.Forms.PictureBoxSizeMode)
      Me.staticPicture.TabIndex = CType(resources.GetObject("staticPicture.TabIndex"), Integer)
      Me.staticPicture.TabStop = False
      Me.staticPicture.Text = resources.GetString("staticPicture.Text")
      Me.buttonToolTip.SetToolTip(Me.staticPicture, resources.GetString("staticPicture.ToolTip"))
      Me.staticPicture.Visible = CType(resources.GetObject("staticPicture.Visible"), Boolean)
      '
      'dynamicPicture
      '
      Me.dynamicPicture.AccessibleDescription = resources.GetString("dynamicPicture.AccessibleDescription")
      Me.dynamicPicture.AccessibleName = resources.GetString("dynamicPicture.AccessibleName")
      Me.dynamicPicture.Anchor = CType(resources.GetObject("dynamicPicture.Anchor"), System.Windows.Forms.AnchorStyles)
      Me.dynamicPicture.BackgroundImage = CType(resources.GetObject("dynamicPicture.BackgroundImage"), System.Drawing.Image)
      Me.dynamicPicture.Dock = CType(resources.GetObject("dynamicPicture.Dock"), System.Windows.Forms.DockStyle)
      Me.dynamicPicture.Enabled = CType(resources.GetObject("dynamicPicture.Enabled"), Boolean)
      Me.dynamicPicture.Font = CType(resources.GetObject("dynamicPicture.Font"), System.Drawing.Font)
      Me.dynamicPicture.Image = CType(resources.GetObject("dynamicPicture.Image"), System.Drawing.Image)
      Me.dynamicPicture.ImeMode = CType(resources.GetObject("dynamicPicture.ImeMode"), System.Windows.Forms.ImeMode)
      Me.dynamicPicture.Location = CType(resources.GetObject("dynamicPicture.Location"), System.Drawing.Point)
      Me.dynamicPicture.Name = "dynamicPicture"
      Me.dynamicPicture.RightToLeft = CType(resources.GetObject("dynamicPicture.RightToLeft"), System.Windows.Forms.RightToLeft)
      Me.dynamicPicture.Size = CType(resources.GetObject("dynamicPicture.Size"), System.Drawing.Size)
      Me.dynamicPicture.SizeMode = CType(resources.GetObject("dynamicPicture.SizeMode"), System.Windows.Forms.PictureBoxSizeMode)
      Me.dynamicPicture.TabIndex = CType(resources.GetObject("dynamicPicture.TabIndex"), Integer)
      Me.dynamicPicture.TabStop = False
      Me.dynamicPicture.Text = resources.GetString("dynamicPicture.Text")
      Me.buttonToolTip.SetToolTip(Me.dynamicPicture, resources.GetString("dynamicPicture.ToolTip"))
      Me.dynamicPicture.Visible = CType(resources.GetObject("dynamicPicture.Visible"), Boolean)
      '
      'initCOMTimeLabel
      '
      Me.initCOMTimeLabel.AccessibleDescription = resources.GetString("initCOMTimeLabel.AccessibleDescription")
      Me.initCOMTimeLabel.AccessibleName = resources.GetString("initCOMTimeLabel.AccessibleName")
      Me.initCOMTimeLabel.Anchor = CType(resources.GetObject("initCOMTimeLabel.Anchor"), System.Windows.Forms.AnchorStyles)
      Me.initCOMTimeLabel.AutoSize = CType(resources.GetObject("initCOMTimeLabel.AutoSize"), Boolean)
      Me.initCOMTimeLabel.Dock = CType(resources.GetObject("initCOMTimeLabel.Dock"), System.Windows.Forms.DockStyle)
      Me.initCOMTimeLabel.Enabled = CType(resources.GetObject("initCOMTimeLabel.Enabled"), Boolean)
      Me.initCOMTimeLabel.Font = CType(resources.GetObject("initCOMTimeLabel.Font"), System.Drawing.Font)
      Me.initCOMTimeLabel.Image = CType(resources.GetObject("initCOMTimeLabel.Image"), System.Drawing.Image)
      Me.initCOMTimeLabel.ImageAlign = CType(resources.GetObject("initCOMTimeLabel.ImageAlign"), System.Drawing.ContentAlignment)
      Me.initCOMTimeLabel.ImageIndex = CType(resources.GetObject("initCOMTimeLabel.ImageIndex"), Integer)
      Me.initCOMTimeLabel.ImeMode = CType(resources.GetObject("initCOMTimeLabel.ImeMode"), System.Windows.Forms.ImeMode)
      Me.initCOMTimeLabel.Location = CType(resources.GetObject("initCOMTimeLabel.Location"), System.Drawing.Point)
      Me.initCOMTimeLabel.Name = "initCOMTimeLabel"
      Me.initCOMTimeLabel.RightToLeft = CType(resources.GetObject("initCOMTimeLabel.RightToLeft"), System.Windows.Forms.RightToLeft)
      Me.initCOMTimeLabel.Size = CType(resources.GetObject("initCOMTimeLabel.Size"), System.Drawing.Size)
      Me.initCOMTimeLabel.TabIndex = CType(resources.GetObject("initCOMTimeLabel.TabIndex"), Integer)
      Me.initCOMTimeLabel.Text = CType(configurationAppSettings.GetValue("initCOMTimeLabel.Text", GetType(System.String)), String)
      Me.initCOMTimeLabel.TextAlign = CType(resources.GetObject("initCOMTimeLabel.TextAlign"), System.Drawing.ContentAlignment)
      Me.buttonToolTip.SetToolTip(Me.initCOMTimeLabel, resources.GetString("initCOMTimeLabel.ToolTip"))
      Me.initCOMTimeLabel.Visible = CType(resources.GetObject("initCOMTimeLabel.Visible"), Boolean)
      '
      'initAssemblyTimeLabel
      '
      Me.initAssemblyTimeLabel.AccessibleDescription = resources.GetString("initAssemblyTimeLabel.AccessibleDescription")
      Me.initAssemblyTimeLabel.AccessibleName = resources.GetString("initAssemblyTimeLabel.AccessibleName")
      Me.initAssemblyTimeLabel.Anchor = CType(resources.GetObject("initAssemblyTimeLabel.Anchor"), System.Windows.Forms.AnchorStyles)
      Me.initAssemblyTimeLabel.AutoSize = CType(resources.GetObject("initAssemblyTimeLabel.AutoSize"), Boolean)
      Me.initAssemblyTimeLabel.Dock = CType(resources.GetObject("initAssemblyTimeLabel.Dock"), System.Windows.Forms.DockStyle)
      Me.initAssemblyTimeLabel.Enabled = CType(resources.GetObject("initAssemblyTimeLabel.Enabled"), Boolean)
      Me.initAssemblyTimeLabel.Font = CType(resources.GetObject("initAssemblyTimeLabel.Font"), System.Drawing.Font)
      Me.initAssemblyTimeLabel.Image = CType(resources.GetObject("initAssemblyTimeLabel.Image"), System.Drawing.Image)
      Me.initAssemblyTimeLabel.ImageAlign = CType(resources.GetObject("initAssemblyTimeLabel.ImageAlign"), System.Drawing.ContentAlignment)
      Me.initAssemblyTimeLabel.ImageIndex = CType(resources.GetObject("initAssemblyTimeLabel.ImageIndex"), Integer)
      Me.initAssemblyTimeLabel.ImeMode = CType(resources.GetObject("initAssemblyTimeLabel.ImeMode"), System.Windows.Forms.ImeMode)
      Me.initAssemblyTimeLabel.Location = CType(resources.GetObject("initAssemblyTimeLabel.Location"), System.Drawing.Point)
      Me.initAssemblyTimeLabel.Name = "initAssemblyTimeLabel"
      Me.initAssemblyTimeLabel.RightToLeft = CType(resources.GetObject("initAssemblyTimeLabel.RightToLeft"), System.Windows.Forms.RightToLeft)
      Me.initAssemblyTimeLabel.Size = CType(resources.GetObject("initAssemblyTimeLabel.Size"), System.Drawing.Size)
      Me.initAssemblyTimeLabel.TabIndex = CType(resources.GetObject("initAssemblyTimeLabel.TabIndex"), Integer)
      Me.initAssemblyTimeLabel.Text = resources.GetString("initAssemblyTimeLabel.Text")
      Me.initAssemblyTimeLabel.TextAlign = CType(resources.GetObject("initAssemblyTimeLabel.TextAlign"), System.Drawing.ContentAlignment)
      Me.buttonToolTip.SetToolTip(Me.initAssemblyTimeLabel, resources.GetString("initAssemblyTimeLabel.ToolTip"))
      Me.initAssemblyTimeLabel.Visible = CType(resources.GetObject("initAssemblyTimeLabel.Visible"), Boolean)
      '
      'Form1
      '
      Me.AccessibleDescription = resources.GetString("$this.AccessibleDescription")
      Me.AccessibleName = resources.GetString("$this.AccessibleName")
      Me.AutoScaleBaseSize = CType(resources.GetObject("$this.AutoScaleBaseSize"), System.Drawing.Size)
      Me.AutoScroll = CType(resources.GetObject("$this.AutoScroll"), Boolean)
      Me.AutoScrollMargin = CType(resources.GetObject("$this.AutoScrollMargin"), System.Drawing.Size)
      Me.AutoScrollMinSize = CType(resources.GetObject("$this.AutoScrollMinSize"), System.Drawing.Size)
      Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
      Me.ClientSize = CType(resources.GetObject("$this.ClientSize"), System.Drawing.Size)
      Me.Controls.Add(Me.initAssemblyTimeLabel)
      Me.Controls.Add(Me.initCOMTimeLabel)
      Me.Controls.Add(Me.dynamicPicture)
      Me.Controls.Add(Me.staticPicture)
      Me.Controls.Add(Me.clickMeButton)
      Me.Enabled = CType(resources.GetObject("$this.Enabled"), Boolean)
      Me.Font = CType(resources.GetObject("$this.Font"), System.Drawing.Font)
      Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
      Me.ImeMode = CType(resources.GetObject("$this.ImeMode"), System.Windows.Forms.ImeMode)
      Me.Location = CType(resources.GetObject("$this.Location"), System.Drawing.Point)
      Me.MaximumSize = CType(resources.GetObject("$this.MaximumSize"), System.Drawing.Size)
      Me.MinimumSize = CType(resources.GetObject("$this.MinimumSize"), System.Drawing.Size)
      Me.Name = "Form1"
      Me.RightToLeft = CType(resources.GetObject("$this.RightToLeft"), System.Windows.Forms.RightToLeft)
      Me.StartPosition = CType(resources.GetObject("$this.StartPosition"), System.Windows.Forms.FormStartPosition)
      Me.Text = resources.GetString("$this.Text")
      Me.buttonToolTip.SetToolTip(Me, resources.GetString("$this.ToolTip"))
      Me.ResumeLayout(False)

   End Sub

#End Region

   Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clickMeButton.Click
      Dim comObject As New COMTest.VB6COMClass
      Dim assemblyObject As New TestAssembly1.AssemblyTest

      ' Load and display dynamic picture
      dynamicPicture.Image = Image.FromFile("Author.jpg")

      ' Update labels
      initCOMTimeLabel.Text += comObject.GetInitTime.ToShortTimeString
      initAssemblyTimeLabel.Text += assemblyObject.GetInitTime.ToShortTimeString
   End Sub
End Class