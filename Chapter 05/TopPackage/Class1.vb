' Static Model


Public Class Class1
		 Implements ObjectControl, _
					_ImageShrink

	Public Sub Activate () Implements ObjectControl.Activate
		
	End Sub

	Public Sub Deactivate () Implements ObjectControl.Deactivate
		
	End Sub

	Public Sub CanBePooled () Implements ObjectControl.CanBePooled
		
	End Sub

	Public Function Shrink (ByVal BMPImage As Object, _
						 ByVal ShrinkPercentage As Integer) As Object Implements _ImageShrink.Shrink
		
	End Function

End Class ' END CLASS DEFINITION Class1
