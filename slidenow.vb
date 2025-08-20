Option Explicit

Dim map(50, 50) As Boolean
Dim mapitem(50, 50) As Long
Dim now_x, now_y As Long
Dim player_arrow As Long
Dim life, hammer, key As Long
Dim maplarge As Long
Dim isdecreaselife, ishammeruse As Boolean
Dim itemgetnum As Long

Dim time As Long

Dim remsg As VbMsgBoxResult

Dim howtoplatend As Long

Dim firstplaycheck As Boolean
Dim howtoplayend As Long
Dim howtoplayend_2_stage As Long

Const nm As String = "3-4-05 automaze_new.pptm"
Dim sld As Long


Sub a_newstart()

    firstplaycheck = True
    Presentations(nm).SlideShowWindow.View.GotoSlide 3

End Sub


Sub a_makemap(ByVal level As Long)
    
    firstplaycheck = False
    
    Select Case level
    
    Case 1
        sld = 8
        maplarge = 10
        life = 30
    
    Case 2
        sld = 13
        maplarge = 20
        life = 90
    
    Case 3
        sld = 18
        maplarge = 30
        life = 180
    
    End Select

    Presentations(nm).Slides(sld).Shapes("start").Visible = msoTrue

    Dim counti, countj As Long
    For counti = 0 To 50
    
        For countj = 0 To 50
                
            map(counti, countj) = True
            
        Next
    
    Next
    
    map(maplarge - 1, maplarge) = False
    
    '--------------<map>-------------------------------------------------------------
    
    
    Dim fin As Boolean: fin = False
    
    Dim makepoint_x As Long
    Dim makepoint_y As Long
    
    Dim roadconnect As Boolean
    Dim roadconnectdirection As Long
    
    Dim roadaroundcheck(1 To 4) As Boolean
    Dim roadaroundcheckcount As Long
    
    Dim roadNOTconnect_falsecheckcount As Long
    
    Randomize
    makepoint_x = Int(Rnd * maplarge / 2 + 1) * 2 - 1
    
    Randomize
    makepoint_y = Int(Rnd * maplarge / 2 + 1) * 2 - 1
            
    map(makepoint_x, makepoint_y) = False
    roadconnect = True

    
    Do Until fin = True
    
        If roadconnect = True Then '壁がなければ...
        
            roadaroundcheckcount = 0 '道を作れる個数
        
            '左のチェック
            
            If makepoint_x = 1 Then
                roadaroundcheck(1) = False
            Else
                roadaroundcheck(1) = map(makepoint_x - 2, makepoint_y) '左がかべかどうか
            End If
            If roadaroundcheck(1) = True Then 'もし左がかべならば...
                roadaroundcheckcount = roadaroundcheckcount + 1
            End If

            '上のチェック

            If makepoint_y = 1 Then
                roadaroundcheck(2) = False
            Else
                roadaroundcheck(2) = map(makepoint_x, makepoint_y - 2)
            End If
            If roadaroundcheck(2) = True Then
                roadaroundcheckcount = roadaroundcheckcount + 1
            End If

            '右のチェック
            
            If makepoint_x = maplarge - 1 Then
                roadaroundcheck(3) = False
            Else
                roadaroundcheck(3) = map(makepoint_x + 2, makepoint_y)
            End If
            If roadaroundcheck(3) = True Then
                roadaroundcheckcount = roadaroundcheckcount + 1
            End If

            '下のチェック
            
            If makepoint_y = maplarge - 1 Then
                roadaroundcheck(4) = False
            Else
                roadaroundcheck(4) = map(makepoint_x, makepoint_y + 2)
            End If
            If roadaroundcheck(4) = True Then
                roadaroundcheckcount = roadaroundcheckcount + 1
            End If

            
            '------------------------------------------
            
            
            If roadaroundcheckcount = 0 Then '行き止まりならば...
            
                roadconnect = False
            
            Else
                
                Do
                    
                    Randomize
                    roadconnectdirection = Int(Rnd * 4 + 1) 'ランダムに壁を伸ばす方向を決める
                
                Loop Until roadaroundcheck(roadconnectdirection) = True
                
                Dim countl As Long: For countl = 1 To 2 '2回続けて壁を伸ばす
                
                    Select Case roadconnectdirection
                
                        Case 1
                            makepoint_x = makepoint_x - 1
                            map(makepoint_x, makepoint_y) = False
                        
                        Case 2
                            makepoint_y = makepoint_y - 1
                            map(makepoint_x, makepoint_y) = False
                    
                        Case 3
                            makepoint_x = makepoint_x + 1
                            map(makepoint_x, makepoint_y) = False
                    
                        Case 4
                            makepoint_y = makepoint_y + 1
                            map(makepoint_x, makepoint_y) = False
                    
                    End Select
                    
                Next
            
            End If
            
        Else
        
            roadNOTconnect_falsecheckcount = 0
        
            For counti = 0 To maplarge
            
                For countj = 0 To maplarge
                
                    If counti Mod 2 = 1 And countj Mod 2 = 1 And map(counti, countj) = False Then
                    
                        roadNOTconnect_falsecheckcount = roadNOTconnect_falsecheckcount + 1
                    
                    End If
                
                Next
            
            Next
        
            If roadNOTconnect_falsecheckcount = (maplarge / 2) ^ 2 Then '全ての壁をくり抜いたら...
            
                fin = True
            
            Else
            
                Do
        
                    Randomize
                    makepoint_x = Int(Rnd * maplarge / 2 + 1) * 2 - 1
            
                    Randomize
                    makepoint_y = Int(Rnd * maplarge / 2 + 1) * 2 - 1
        
                Loop Until map(makepoint_x, makepoint_y) = False '次のスタート地点をテキトーに決める
                roadconnect = True
        
            End If
        
        End If
        
    Loop
    
    
    '-----------<item>----------------------------------------------------------------------
    
    
    For counti = 0 To maplarge
    
        For countj = 0 To maplarge
        
            mapitem(counti, countj) = 0
                
        Next
    
    Next
    
    mapitem(maplarge - 1, maplarge) = 1
    
    counti = 0
    Do
            
        Randomize
        makepoint_x = Int(Rnd * maplarge / 2 + 1) * 2 - 1
        
        Randomize
        makepoint_y = Int(Rnd * maplarge / 2 + 1) * 2 - 1
    
        If makepoint_x <> 1 And makepoint_y <> 1 And makepoint_x <> maplarge - 1 And makepoint_y <> maplarge - 1 And _
                mapitem(makepoint_x, makepoint_y) = 0 Then
            
            mapitem(makepoint_x, makepoint_y) = 2
            counti = counti + 1
            
        End If
    
    Loop Until counti = Int(maplarge / 5) '肉の生成

    counti = 0
    Do
            
        Randomize
        makepoint_x = Int(Rnd * maplarge / 2 + 1) * 2 - 1
        
        Randomize
        makepoint_y = Int(Rnd * maplarge / 2 + 1) * 2 - 1
    
        If makepoint_x <> 1 And makepoint_y <> 1 And makepoint_x <> maplarge - 1 And makepoint_y <> maplarge - 1 _
                And mapitem(makepoint_x, makepoint_y) = 0 Then
            
            mapitem(makepoint_x, makepoint_y) = 3
            counti = counti + 1
            
        End If

    Loop Until counti = Int(maplarge / 20) + 1 'ハンマーの生成
    
    counti = 0
    Do
            
        Randomize
        makepoint_x = Int(Rnd * maplarge / 2 + 1) * 2 - 1
        
        Randomize
        makepoint_y = Int(Rnd * maplarge / 2 + 1) * 2 - 1
    
        If makepoint_x <> 1 And makepoint_y <> 1 And makepoint_x <> maplarge - 1 And makepoint_y <> maplarge - 1 _
                And mapitem(makepoint_x, makepoint_y) = 0 Then
            
            mapitem(makepoint_x, makepoint_y) = 4
            counti = counti + 1
            
        End If
    
    Loop Until counti = Int(maplarge / 30) + 1 '鍵の生成
    
    now_x = 1
    now_y = 1
    player_arrow = 4
    isdecreaselife = False
    hammer = 0
    ishammeruse = False
    key = 0
    itemgetnum = 0
    time = timer() + 8
    
    Call a_drawmap
        
    Presentations(nm).SlideShowWindow.View.GotoSlide sld - 2
    
End Sub


Sub a_drawmap()
    
    Dim meat_get As Boolean: meat_get = False
    Dim hammer_get As Boolean: hammer_get = False
    Dim key_get As Boolean: key_get = False
    
    Presentations(nm).Slides(sld).Shapes("座標").TextFrame.TextRange.Text = "※現在のプレイヤーの座標：（" & now_x & "," & now_y & "）"
    
    If mapitem(now_x, now_y) = 2 Then
        meat_get = True
        mapitem(now_x, now_y) = 0
    
    ElseIf mapitem(now_x, now_y) = 3 Then
        hammer_get = True
        mapitem(now_x, now_y) = 0
    
    ElseIf mapitem(now_x, now_y) = 4 Then
        key_get = True
        mapitem(now_x, now_y) = 0
    
    End If
    
    Do Until Presentations(nm).Slides(sld).Shapes(1).Name = "0"
    
        Presentations(nm).Slides(sld).Shapes(1).Delete
    
    Loop

    Dim now_lox, now_loy As Long
    Dim isdraw As Boolean
    Dim grid_num As Long

    Dim counti As Long: For counti = 0 To 6
    
        Dim countj As Long: For countj = 0 To 6
            
            now_lox = countj + now_x - 3
            now_loy = counti + now_y - 3
            
            If 0 <= now_lox And now_lox <= maplarge And 0 <= now_loy And now_loy <= maplarge Then
            
                If map(now_lox, now_loy) = True Then
                
                    isdraw = True
                    
                Else
                
                    isdraw = False
                
                End If
            
            Else
                
                isdraw = False
                
            End If
            
            grid_num = counti * 7 + countj
            
            If isdraw = True Xor Presentations(nm).Slides(sld).Shapes(CStr(grid_num)).Visible = True Then
            
                If isdraw = True Then
                    
                    Presentations(nm).Slides(sld).Shapes(CStr(grid_num)).Visible = True
                    
                Else
                
                    Presentations(nm).Slides(sld).Shapes(CStr(grid_num)).Visible = False
                
                End If
            
            End If
            
            If 0 <= now_lox And now_lox <= maplarge And 0 <= now_loy And now_loy <= maplarge Then
            
                If mapitem(now_lox, now_loy) <> 0 Then

                    Select Case mapitem(now_lox, now_loy)
                
                        Case 1

                            Presentations(nm).Slides(23).Shapes("door").Copy
                        
                        Case 2
                        
                            Presentations(nm).Slides(23).Shapes("meat").Copy
                        
                        Case 3
                        
                            Presentations(nm).Slides(23).Shapes("hammer").Copy
                        
                        Case 4
                        
                            Presentations(nm).Slides(23).Shapes("key").Copy
                        
                    End Select
                
                    Presentations(nm).Slides(sld).Shapes.Paste
                
                    With Presentations(nm).Slides(sld).Shapes(Presentations(nm).Slides(sld).Shapes.Count)
                
                        .Left = countj * 60 + 90
                        .Top = counti * 60 + 60
                        .ZOrder (msoSendToBack)
                
                    End With
                    
                End If
            
            End If
            
        Next
    
    Next

    With Presentations(nm).Slides(sld)

        Select Case player_arrow
    
            Case 1
                .Shapes("a").Visible = msoTrue: .Shapes("b").Visible = msoFalse
                .Shapes("c").Visible = msoFalse: .Shapes("d").Visible = msoFalse
            
            Case 2
                .Shapes("b").Visible = msoTrue: .Shapes("a").Visible = msoFalse
                .Shapes("c").Visible = msoFalse: .Shapes("d").Visible = msoFalse

            Case 3
                .Shapes("c").Visible = msoTrue: .Shapes("b").Visible = msoFalse
                .Shapes("a").Visible = msoFalse: .Shapes("d").Visible = msoFalse

            Case 4
                .Shapes("d").Visible = msoTrue: .Shapes("b").Visible = msoFalse
                .Shapes("c").Visible = msoFalse: .Shapes("a").Visible = msoFalse

        End Select 'プレイヤーの向きを変える
    
    End With

    If isdecreaselife = True Then
        life = life - 1
    End If
    
    If meat_get = True Then
    
        life = life + 20
        Presentations(nm).Slides(sld).Shapes("アイテム値1").TextFrame.TextRange.Text = "残り" & life
        If life >= 20 Then
            Presentations(nm).Slides(sld).Shapes("アイテム値1").TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
        End If
        
        DoEvents
        
        MsgBox "アイテム「肉」を獲得し、体力ゲージが20増えた！", vbInformation + vbOKOnly, "Message from AUTOMAZE"
        itemgetnum = itemgetnum + 1
    
    End If
    
    Presentations(nm).Slides(sld).Shapes("アイテム値1").TextFrame.TextRange.Text = "残り" & life
    If life = 19 And isdecreaselife = True Then
        Presentations(nm).Slides(sld).Shapes("アイテム値1").TextFrame.TextRange.Font.Color.RGB = RGB(256, 0, 0)
        DoEvents
        MsgBox "体力が少なくなってきています！　至急、アイテムの「肉」をゲットし、体力を回復させてください！", vbExclamation + vbOKOnly, _
                "Message from AUTOMAZE"
    ElseIf life >= 20 Then
        Presentations(nm).Slides(sld).Shapes("アイテム値1").TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    End If
    
    If hammer_get = True Then
    
        hammer = hammer + 1
        Presentations(nm).Slides(sld).Shapes("アイテム値2").TextFrame.TextRange.Text = hammer & "個所持"
        Presentations(nm).Slides(sld).Shapes("アイテム値2").TextFrame.TextRange.Font.Color.RGB = RGB(0, 224, 0)
        DoEvents
        
        MsgBox "アイテム「ハンマー」を獲得した！　ハンマー1個につき1回、壁を壊すことができる。", vbInformation + vbOKOnly, _
                "Message from AUTOMAZE"
        itemgetnum = itemgetnum + 1
    
    ElseIf hammer = 0 Then
        
        Presentations(nm).Slides(sld).Shapes("アイテム値2").TextFrame.TextRange.Text = hammer & "個所持"
        Presentations(nm).Slides(sld).Shapes("アイテム値2").TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    
    End If
    
    If ishammeruse = True Then
    
        MsgBox "壁を壊すことに成功した！", vbInformation + vbOKOnly, "Message from AUTOMAZE"
    
    End If
    
    If key_get = True Then
        
        key = key + 1
        Presentations(nm).Slides(sld).Shapes("アイテム値3").TextFrame.TextRange.Text = key & "個所持"
        Presentations(nm).Slides(sld).Shapes("アイテム値3").TextFrame.TextRange.Font.Color.RGB = RGB(0, 256, 256)
            
        MsgBox "アイテム「鍵」を手に入れた！　この鍵をドアのところへ持っていけば、迷路から脱出できるようだ。" _
                , vbInformation + vbOKOnly, "Message from AUTOMAZE"
        itemgetnum = itemgetnum + 1

    ElseIf key = 0 Then
    
        Presentations(nm).Slides(sld).Shapes("アイテム値3").TextFrame.TextRange.Text = key & "個所持"
        Presentations(nm).Slides(sld).Shapes("アイテム値3").TextFrame.TextRange.Font.Color.RGB = RGB(256, 0, 0)
        
    Else
    
        Presentations(nm).Slides(sld).Shapes("アイテム値3").TextFrame.TextRange.Text = key & "個所持"
    
    End If
    
    If life = 0 Then
    
        MsgBox "もう体力がない...", vbCritical + vbOKOnly, "Message from AUTOMAZE"
        Presentations(nm).SlideShowWindow.View.GotoSlide 24
    
    End If
    
    If Presentations(nm).SlideShowWindow.View.CurrentShowPosition = sld Then
    
        Presentations(nm).Slides(sld).Shapes("start").Visible = msoFalse
    
    End If
    
End Sub


Sub a_brakewall()
    
    If hammer >= 1 Then
    
        Dim player_arrow_japanese As String
        Dim erasemap_x, erasemap_y As Long
        erasemap_x = now_x: erasemap_y = now_y
        
        Dim wallaroundnot As Boolean: wallaroundnot = True
    
        Select Case player_arrow
    
            Case 1
                player_arrow_japanese = "右"
                erasemap_x = erasemap_x + 1
                
                If erasemap_x = maplarge Then
                
                    wallaroundnot = False
                
                End If
            
            Case 2
                player_arrow_japanese = "左"
                erasemap_x = erasemap_x - 1
                
                If erasemap_x = 0 Then
                
                    wallaroundnot = False
                
                End If

            Case 3
                player_arrow_japanese = "上"
                erasemap_y = erasemap_y - 1
                
                If erasemap_y = 0 Then
                
                    wallaroundnot = False
                
                End If

            Case 4
                player_arrow_japanese = "下"
                erasemap_y = erasemap_y + 1
                
                If erasemap_y = maplarge Then
                
                    wallaroundnot = False
                
                End If

        End Select
    
    
        If wallaroundnot = True Then
    
            remsg = MsgBox("ハンマー1個を使って、プレイヤーの一つ" & player_arrow_japanese & "の壁を壊そうとしています。よろしいですか？", _
                vbExclamation + vbYesNo, "Message from AUTOMAZE")
            
            If remsg = vbYes Then
            
                isdecreaselife = False
                hammer = hammer - 1
            ishammeruse = True
            map(erasemap_x, erasemap_y) = False
            
            Call a_drawmap

            End If
        
        Else
        
            MsgBox "外枠の壁はハンマーで壊せません。ごめんなさい...", vbExclamation + vbOKOnly, "Message from AUTOMAZE"
            
        End If
    
    End If

End Sub


Sub a_right()

    If player_arrow <> 1 Or map(now_x + 1, now_y) = False Then
    
        ishammeruse = False
        player_arrow = 1
    
        If map(now_x + 1, now_y) = False Then
    
            now_x = now_x + 1
            isdecreaselife = True

        Else
        
            isdecreaselife = False
        
        End If
        
        Call a_drawmap
    
    Else
    
        Call a_brakewall
    
    End If

End Sub


Sub a_left()

    If player_arrow <> 2 Or map(now_x - 1, now_y) = False Then
        
        ishammeruse = False
        player_arrow = 2
    
        If map(now_x - 1, now_y) = False Then
    
            now_x = now_x - 1
            isdecreaselife = True

        Else
        
            isdecreaselife = False
        
        End If
        
        Call a_drawmap
        
    Else
    
        Call a_brakewall

    End If

End Sub


Sub a_up()
    
    If player_arrow <> 3 Or map(now_x, now_y - 1) = False Then
        
        ishammeruse = False
        player_arrow = 3

        If map(now_x, now_y - 1) = False Then
    
            now_y = now_y - 1
            isdecreaselife = True

        Else
        
            isdecreaselife = False
        
        End If
        
        Call a_drawmap
        
    Else
    
        Call a_brakewall

    End If

End Sub


Sub a_down()
    
    If now_x = maplarge - 1 And now_y = maplarge - 1 Then
    
        player_arrow = 4
        isdecreaselife = False
        Call a_drawmap
        
            If key = 0 Then
            
                MsgBox "このドアを開けるには、鍵が必要なようだ。", vbExclamation + vbOKOnly, "Message from AUTOMAZE"
            
            Else
            
                remsg = MsgBox("鍵を1個使って、ドアを開錠しようとしています。よろしいですか？", _
                        vbExclamation + vbYesNo, "Message from AUTOMAZE")
            
                If remsg = vbYes Then
                    
                    Presentations(nm).Slides(sld + 2).Shapes("データ").TextFrame.TextRange.Text = _
                            "かかった時間：" & Int((timer() - time) / 60) & "分" & Int((timer() - time) Mod 60) & "秒" & vbCrLf & _
                            "アイテム獲得数：" & itemgetnum & "/" & Int(maplarge / 5) + Int(maplarge / 20) + Int(maplarge / 30) + 2 & "個" & vbCrLf & _
                            "残りの体力：" & life
                    
                    Presentations(nm).SlideShowWindow.View.GotoSlide sld + 1
                
                End If

        End If
    
    ElseIf player_arrow <> 4 Or map(now_x, now_y + 1) = False Then
        
        ishammeruse = False
        player_arrow = 4

        If map(now_x, now_y + 1) = False Then
    
            now_y = now_y + 1
            isdecreaselife = True

        Else
        
            isdecreaselife = False
        
        End If
        
        Call a_drawmap
        
    Else
    
        Call a_brakewall

    End If

End Sub


Sub a_hint()

    MsgBox "hint"

End Sub



Sub a_gotoURL()

    remsg = MsgBox("外部サイトを開きます。よろしいですか？" & vbCrLf & "（なお、授業中であれば、移動先のサイトで遊ばないようにしてください。）", _
                        vbExclamation + vbYesNo, "Message from AUTOMAZE")
                        
    If remsg = vbYes Then
    
        VBA.Interaction.CreateObject("WScript.Shell").Run ("chrome.exe -url " & "https://scratch.mit.edu/users/41shell_1221/")
        
    End If

End Sub



Sub a_gobackmain()

    Presentations(nm).SlideShowWindow.View.GotoSlide 4

End Sub


Sub a_gotoselectmenw()

    Presentations(nm).SlideShowWindow.View.GotoSlide 5

End Sub


Sub a_choice_easy()
    
    If firstplaycheck = True Then
    
        Call a_firstplay_howtoplay
        howtoplayend_2_stage = 1
    
    Else
    
        Call a_makemap(1)
    
    End If

End Sub


Sub a_choice_normal()

    If firstplaycheck = True Then
    
        Call a_firstplay_howtoplay
        howtoplayend_2_stage = 1
    
    Else

        Call a_makemap(2)
    
    End If

End Sub


Sub a_choice_difficult()
    
    If firstplaycheck = True Then
    
        Call a_firstplay_howtoplay
        howtoplayend_2_stage = 1
    
    Else

        Call a_makemap(3)
        
    End If

End Sub


Sub a_menu_howtoplay()
    
    firstplaycheck = False
    howtoplayend = 1
    Presentations(nm).SlideShowWindow.View.GotoSlide 25

End Sub


Sub a_firstplay_howtoplay()

    howtoplayend = 2
    Presentations(nm).SlideShowWindow.View.GotoSlide 27

End Sub


Sub a_game_howtoplay()

    howtoplayend = 3
    Presentations(nm).SlideShowWindow.View.GotoSlide 25

End Sub


Sub a_howtoplay_ignore()

    Select Case howtoplayend_2_stage
    
        Case 1
            Call a_makemap(1)
        
        Case 2
            Call a_makemap(2)
        
        Case 3
            Call a_makemap(3)
    
    End Select

End Sub


Sub a_howtoplayend()

    Select Case howtoplayend
    
        Case 1
            Presentations(nm).SlideShowWindow.View.GotoSlide 3
        
        Case 2
            Call a_howtoplay_ignore
        
        Case 3
            Presentations(nm).SlideShowWindow.View.GotoSlide sld
    
    End Select


End Sub


Sub test1()

    life = 20
    hammer = 1
    
End Sub


Sub test2()

    Presentations(nm).Slides().Shapes.AddShape msoShapeRectangle, 0, 0, 540, 540

End Sub






