Option Explicit

Dim now_x, now_y As Long


Sub makemap()

    Dim i As Long: For i = 0 To 6
    
        Dim j As Long: For j = 0 To 6
        
            Presentations("game_10").Slides(4).Shapes.AddShape Type:=msoShapeRectangle, _
                Left:=j * 60 + 90, Top:=i * 60 + 60, Width:=60, Height:=60
            
        Next
    
    Next

End Sub


Sub makemap2()

    Dim i As Long: For i = 0 To 13
    
        Dim j As Long: For j = 0 To 13
        
            If j Mod 2 = 1 Then
            
                Presentations("game_10").Slides(4).Shapes.AddShape Type:=msoShapeRectangle, _
                Left:=j * 30 + 90, Top:=i * 30 + 60, Width:=30, Height:=30
            
            Else
            
                Presentations("game_10").Slides(5).Shapes.AddShape Type:=msoShapeRectangle, _
                Left:=j * 30 + 90, Top:=i * 30 + 60, Width:=30, Height:=30

            End If
            
        Next
    
    Next

End Sub


Sub mappaint()

    Dim i As Long: For i = 1 To 14 ^ 2
    
        With Presentations("game_10").Slides(1).Shapes(i)
        
            .Name = i - 1
            
            .Line.Visible = msoFalse
        
        End With
        
        If (i \ 2) Mod 2 = 1 Xor ((i - 1) \ 14 \ 2) Mod 2 = 1 Then
        
            Presentations("game_10").Slides(1).Shapes(i).Visible = msoTrue
        
        Else
        
            Presentations("game_10").Slides(1).Shapes(i).Visible = msoTrue
        
        End If
    
    Next

End Sub


Sub mappaint2()
    
    Dim i As Long: For i = 0 To 6
    
    Dim j As Long: For j = 1 To 49
    
        With Presentations("game_10").Slides(9).Shapes(i * 49 + j)
        
            .Name = i * 100 + j - 1
            
            .Line.Visible = msoFalse
        
        End With
    
    Next
    
    Next

End Sub


Sub mappaint3()

    Dim shp As Shape: For Each shp In ActivePresentation.Slides(4).Shapes
    
        If IsNumeric(shp.Name) = True Then
        
            If shp.Name > 100 Then
            
                shp.Delete
            
            End If
        
        End If
    
    Next

End Sub


Sub mappaint4()
    
    Dim a As Long: a = 284
    MsgBox ActivePresentation.Slides(1).Shapes(a).Name
    MsgBox IsNumeric(ActivePresentation.Slides(1).Shapes(a).Name)

End Sub





Sub test()

    Dim i As Long: For i = 1 To 1
    
        'Presentations("game_10").Slides(1).Shapes.AddShape Type:=msoShapeRectangle, _
            Left:=0, Top:=0, Width:=60, Height:=60
    
        Presentations("game_10").Slides(1).Shapes.AddShape Type:=msoShapeOval, _
            Left:=90, Top:=60, Width:=420, Height:=420

    Next

End Sub


Sub icon()
    
        Presentations("game_10").Slides(1).Shapes.AddShape Type:=msoShapeRightArrow, _
        Left:=515, Top:=245, Width:=50, Height:=50
    

End Sub


Sub zyunnbann()

    Dim i As Long: For i = 0 To 195
    
        Dim shp As Shape: For Each shp In Presentations("game_10").Slides(4).Shapes
        
            If shp.Name = (i) Then
            
                shp.Copy
                Presentations("game_10").Slides(5).Shapes.Paste
                
                Exit For
            
            End If
    
        Next
    
    Next

End Sub



