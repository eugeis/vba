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

Sub Timeline1()
    Timeline 425, 100, 10, 410
End Sub

Sub Timeline(timelineTop As Integer, timelineHeight As Integer, _
            zoomedTop As Integer, zoomedHeight As Integer)
            
    Dim sh As Shape
    Set sl = Application.ActiveWindow.View.Slide
    Dim shapes As New Collection
    Dim x, mx, sumWidth As Integer
    Dim eff As Effect
    
    Dim left As Integer
    
    x = ActivePresentation.PageSetup.SlideWidth / 2
    
    For Each sh In sl.shapes
      If sh.Type = msoPicture Then
        shapes.Add sh
      End If
    Next
    
    'set height and calculate start left
    sumWidth = 0
    For Each sh In shapes
        sh.Height = timelineHeight
        sumWidth = sumWidth + sh.Width
    Next
    left = (ActivePresentation.PageSetup.SlideWidth - sumWidth) / 2

    For Each sh In shapes
        'resize and move for timeline
        sh.left = left
        sh.Top = timelineTop
        left = left + sh.Width
    
    
        'copy for zoomed
        sh.Copy
        Set newShapeRange = sl.shapes.Paste
        Set newShape = newShapeRange.Item(1)
        
        With newShape
         .Height = zoomedHeight
         .Top = zoomedTop
         .left = x - (.Width / 2)
        End With
        
        Set eff = sl.Timeline.MainSequence _
        .AddEffect(Shape:=newShape, effectId:=msoAnimEffectCustom)
        
        eff.Timing.Duration = 1
        
        
        Set aniMotion = eff.Behaviors.Add(msoAnimTypeMotion)
        
        'calculate percent position
        mx = 50 - (sh.Width / 2 + sh.left) * 100 / ActivePresentation.PageSetup.SlideWidth
        
        With aniMotion.MotionEffect
            .FromX = -mx
            .FromY = 50
            .ToX = 0
            .ToY = 0
        End With
        
        Set eff = sl.Timeline.MainSequence.AddEffect(Shape:=newShape, _
        effectId:=msoAnimEffectZoom, Trigger:=msoAnimTriggerWithPrevious)
        
        eff.Timing.Duration = 1
    Next
    
End Sub
