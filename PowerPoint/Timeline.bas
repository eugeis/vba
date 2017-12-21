Attribute VB_Name = "Timeline"
Sub TimelineForm()
    TimelineGenerator.Show
End Sub

Sub Timeline()
    GenerateTimeline 425, 100, 10, 410
End Sub

Function FindPictures(sl As Slide) As Collection
    Dim shapes As New Collection
    Dim sh As Shape
    For Each sh In sl.shapes
      If sh.Type = msoPicture Then
        shapes.Add sh
      End If
    Next
    Set FindPictures = shapes
End Function

Sub GenerateTimeline(TimelineTop As Integer, TimelineHeight As Integer, _
            zoomedTop As Integer, zoomedHeight As Integer)
            
    Dim shapes As Collection
    Dim sh, newShape As Shape
    Dim newShapeRange As ShapeRange
    Dim sl As Slide
    Dim x, mx, sumWidth As Integer
    Dim eff As Effect
    Dim AniMotion As AnimationBehavior
    
    Dim left As Integer
    
    Set sl = Application.ActiveWindow.View.Slide

    Set shapes = FindPictures(sl)

    'set height and calculate start left
    sumWidth = 0
    For Each sh In shapes
        sh.Height = TimelineHeight
        sumWidth = sumWidth + sh.Width
    Next
    left = (ActivePresentation.PageSetup.SlideWidth - sumWidth) / 2

    x = ActivePresentation.PageSetup.SlideWidth / 2
    For Each sh In shapes
        'resize and move for timeline
        sh.left = left
        sh.Top = TimelineTop
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
        
        'first shape
        If eff Is Nothing Then
            Set eff = sl.Timeline.MainSequence _
            .AddEffect(Shape:=newShape, effectId:=msoAnimEffectCustom)
        Else
            Set eff = sl.Timeline.MainSequence _
            .AddEffect(Shape:=newShape, effectId:=msoAnimEffectCustom, Trigger:=msoAnimTriggerWithPrevious)
        End If
        
        eff.Timing.Duration = 1
        
        
        Set AniMotion = eff.Behaviors.Add(msoAnimTypeMotion)
        
        'calculate percent position
        mx = 50 - (sh.Width / 2 + sh.left) * 100 / ActivePresentation.PageSetup.SlideWidth
        
        With AniMotion.MotionEffect
            .FromX = -mx
            .FromY = 50
            .ToX = 0
            .ToY = 0
        End With
        
        Set eff = sl.Timeline.MainSequence.AddEffect(Shape:=newShape, _
        effectId:=msoAnimEffectZoom, Trigger:=msoAnimTriggerWithPrevious)
        
        eff.Timing.Duration = 1
        
        Set eff = sl.Timeline.MainSequence.AddEffect(Shape:=newShape, _
        effectId:=msoAnimEffectFade)
        eff.Exit = msoTrue
        
        eff.Timing.Duration = 0.7
    Next
    
End Sub
