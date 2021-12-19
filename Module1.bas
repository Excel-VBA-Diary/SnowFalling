Attribute VB_Name = "Module1"
Option Explicit

'雪の画像はこのマクロブックを置くフォルダにサブホルダーsnowを作り複数個用意する。
'雪の画像のサイズは80x80～100x100ピクセルを推奨。
'雪の画像が大きい場合はSnow.AppearメソッドのReduceパラメータで縮小する

Const ScreenWidth As Double = 400           '雪を降らせる範囲の横幅
Const Screenheight As Double = 300          '雪を降らせる範囲の縦幅

Sub Main()
    Dim i As Long, j As Long
    Dim c As New Collection
    Dim s As Snow
    Dim picturePath As String
    
    Call ClearSnow
    
    For i = 1 To 500
        
        '新たな雪を表示する
        If c.Count <= 50 And Rnd() <= 0.15 Then
            picturePath = ThisWorkbook.Path & "\snow\snow" & CStr(Int(Rnd() * 6) + 1) & ".png"
            Set s = New Snow
            s.Appear picturePath, Rnd() * ScreenWidth, 30, Rnd() * 5 + 1, 0.4
            c.Add s
            Set s = Nothing
        End If
        
        '表示されている雪を順に処理する
        For j = c.Count To 1 Step -1
            Set s = c.Item(j)
            '消える少し手前でだんだん薄くする
            If s.Top > Screenheight - 50 Then
                s.Brightness = s.Brightness + 0.005
            End If
            If s.Top < Screenheight Then
                s.Fall
            Else
                c.Remove j
                s.Disappear
                Set s = Nothing
            End If
        Next
        
        Application.Wait [now()] + 50 / 86400000
    
    Next
    
End Sub

'アクティブシート上のPictueをすべて削除する
Private Sub ClearSnow()
    Dim sp As Shape
    For Each sp In ActiveSheet.Shapes
        If sp.Type = 11 Then sp.Delete
    Next
End Sub
