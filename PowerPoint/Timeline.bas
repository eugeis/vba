Attribute VB_Name = "Timeline"
Sub CopySmileyChangePath()
' NOTE: Change slide and shape index based on where/when you are executing this macro
' 3rd shape on slide 2 is smiley which has a custom path already defined
ActivePresentation.Slides(3).shapes(3).Copy

' Paste shape in slide 1
'With ActivePresentation.Slides(2).Shapes.Paste .Name = "Smiley-" &amp; ActivePresentation.Slides(2).Shapes.Count
'.Left = 200
'.Top = 200
'End With

' Note: Chnage main sequence index if needed
' Change shape path to triange
ActivePresentation.Slides(2).Timeline.MainSequence(2).Behaviors(1).MotionEffect.Path = "M 0 0 L 0.125 0.216 L -0.125 0.216 L 0 0 Z"

End Sub


Sub Resize()
    Dim sh As Shape
    Set sl = ActivePresentation.Slides(6)
    Dim shapes As New Collection
    Dim idToShape As New Dictionary
    Dim x As Integer
    Dim eff As Effect
    
    Dim left As Integer
    left = 0
    
    x = ActivePresentation.PageSetup.SlideWidth / 2
    
    For Each sh In sl.shapes
      If sh.Type = msoPicture And Not idToShape.Exists(sh.Id) Then
      
        Set idToShape(sh.Id) = sh
        With sh
         ' Set position:
         .left = left
         .Top = 400
          ' Set size:
         .Height = 100
        End With
        left = left + sh.Width
        shapes.Add sh
      End If
    Next
    
    For Each sh In shapes
        sh.Copy
        Set newShapeRange = sl.shapes.Paste
        Set newShape = newShapeRange.Item(1)
        
        With newShape
         .Height = 300
         .Top = 20
         .left = x - (.Width / 2)
        End With
        
        Set eff = sl.Timeline.MainSequence _
        .AddEffect(Shape:=newShape, effectId:=msoAnimEffectCustom)
        
        Set aniMotion = eff.Behaviors.Add(msoAnimTypeMotion)

        With aniMotion.MotionEffect
            .FromX = -30
            .FromY = -30
            .ToX = 30
            .ToY = 30
        End With
        
        Set eff = sl.Timeline.MainSequence.AddEffect(Shape:=newShape, _
        effectId:=msoAnimEffectZoom, Trigger:=msoAnimTriggerWithPrevious)
        
    Next
    
End Sub

Sub Animation(s1 As Slide)
' Add custom effect to the shape
        Set effNew = sl.Timeline.MainSequence _
        .AddEffect(Shape:=sh, effectId:=msoAnimEffectCustom, _
        Trigger:=msoAnimTriggerWithPrevious)
        
        ' Add Motion effect
        Set aniMotion = effNew.Behaviors.Add(msoAnimTypeMotion)
        effNew.Exit = msoFalse
        
        ' Set Motion Path as square path
        'aniMotion.MotionEffect.Path = "M 0 0 L 0.25 0 L 0.25 0.25 L 0 0.25 L 0 0 Z"
        aniMotion.MotionEffect.Path = "M 0 0 L 0.125 0.216 L -0.125 0.216 L 0 0 Z"
End Sub
