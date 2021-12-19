VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Snow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private icon As Picture
Private fallSpeed As Single             '�Ⴊ�~�藎���Ă������x
Private rotationSpeed As Single         '��̉�]���x

Property Get Top() As Single
    Top = icon.Top
End Property

Property Let Brightness(Rate As Single)
    icon.ShapeRange.PictureFormat.Brightness = Rate
End Property

Property Get Brightness() As Single
    Brightness = icon.ShapeRange.PictureFormat.Brightness
End Property

'IconPath�F��̉摜�̃t�@�C���p�X
'X0, Y0�F�`��ʒu
'Reduce�F�摜�̏k����
'Distance�F�����i1�ȏ�j
Public Sub Appear(IconPath As String, X0 As Single, Y0 As Single, Distance As Single, Optional Reduce As Single = 1)
    If Distance < 1 Then Distance = 1
    Set icon = ActiveSheet.Pictures.Insert(IconPath)
    With icon
        .Left = X0
        .Top = Y0
        With .ShapeRange
            .PictureFormat.TransparentBackground = True
            .PictureFormat.TransparencyColor = rgbWhite
            .ScaleHeight Reduce / Distance, msoFalse, msoScaleFromTopLeft
            .Fill.PictureEffects.Insert(msoEffectBlur).EffectParameters(1).Value = Distance * 2
        End With
    End With
    fallSpeed = 1 / Distance + 0.5
    rotationSpeed = (Rnd() - 0.5) * 1.5
End Sub

Public Sub Fall()
    icon.Top = icon.Top + fallSpeed
    icon.ShapeRange.IncrementRotation rotationSpeed
End Sub

Public Sub Disappear()
    icon.Delete
End Sub