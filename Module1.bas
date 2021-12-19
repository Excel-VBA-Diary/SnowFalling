Attribute VB_Name = "Module1"
Option Explicit

'��̉摜�͂��̃}�N���u�b�N��u���t�H���_�ɃT�u�z���_�[snow����蕡���p�ӂ���B
'��̉摜�̃T�C�Y��80x80�`100x100�s�N�Z���𐄏��B
'��̉摜���傫���ꍇ��Snow.Appear���\�b�h��Reduce�p�����[�^�ŏk������

Const ScreenWidth As Double = 400           '����~�点��͈͂̉���
Const Screenheight As Double = 300          '����~�点��͈͂̏c��

Sub Main()
    Dim i As Long, j As Long
    Dim c As New Collection
    Dim s As Snow
    Dim picturePath As String
    
    Call ClearSnow
    
    For i = 1 To 500
        
        '�V���Ȑ��\������
        If c.Count <= 50 And Rnd() <= 0.15 Then
            picturePath = ThisWorkbook.Path & "\snow\snow" & CStr(Int(Rnd() * 6) + 1) & ".png"
            Set s = New Snow
            s.Appear picturePath, Rnd() * ScreenWidth, 30, Rnd() * 5 + 1, 0.4
            c.Add s
            Set s = Nothing
        End If
        
        '�\������Ă��������ɏ�������
        For j = c.Count To 1 Step -1
            Set s = c.Item(j)
            '�����鏭����O�ł��񂾂񔖂�����
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

'�A�N�e�B�u�V�[�g���Pictue�����ׂč폜����
Private Sub ClearSnow()
    Dim sp As Shape
    For Each sp In ActiveSheet.Shapes
        If sp.Type = 11 Then sp.Delete
    Next
End Sub
