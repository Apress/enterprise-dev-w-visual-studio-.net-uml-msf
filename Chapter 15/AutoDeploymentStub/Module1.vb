Imports System.Windows.Forms
Imports System.Reflection

Module Module1
   Sub Main()
      Dim remoteAssembly As [Assembly]
      Dim remoteObject As Object
      Dim remoteForm As Form
      Dim classType As Type

      ' Load remote assembly using URL
      remoteAssembly = [Assembly].LoadFrom( _
         "http://localhost/AutoDeploymentApplication" & _
         "/AutoDeploymentApplication.exe")

      ' Retrieve the class type
      classType = _
         remoteAssembly.GetType("AutoDeploymentApplication.Form1")

      ' Create instance of, type cast, and show the form
      remoteObject = Activator.CreateInstance(classType)
      remoteForm = CType(remoteObject, Form)
      Application.Run(remoteForm)
   End Sub
End Module