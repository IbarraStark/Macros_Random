Sub Circle1()
    Dim i As Integer
    Dim shp As Shape
    Dim sld As Slide

    Set sld = Application.ActiveWindow.View.Slide
    Set shp = sld.Shapes.AddShape(msoShapeOval, 120, 20, 99, 32)
    shp.Line.ForeColor.RGB = RGB(255, 0, 0)
    shp.Fill.Visible = msoFalse
    
End Sub