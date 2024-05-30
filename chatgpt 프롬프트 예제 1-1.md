## 파워포인트 슬라이드 만들기 위한 VBA 스크립트 ##
```
I would like you to create VBA code for a PowerPoint presentation about the history of AI. 
There are no placeholders; you must fill in the entire text with your own understanding.
I need five slides. Add slide note for each slide as well. 
```
------------------
결과 값
```
Sub CreateAIHistoryPresentation()
    ' Create a new PowerPoint application
    Dim pptApp As Object
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    
    ' Add a new presentation
    Dim pptPres As Object
    Set pptPres = pptApp.Presentations.Add
    
    ' Define the slide layout constants
    Const ppLayoutTitle As Long = 1
    Const ppLayoutText As Long = 2
    
    ' Add the title slide
    Dim slide1 As Object
    Set slide1 = pptPres.Slides.Add(1, ppLayoutTitle)
    With slide1
        .Shapes.Title.TextFrame.TextRange.Text = "The History of Artificial Intelligence"
        .Shapes.Placeholders(2).TextFrame.TextRange.Text = "A Journey Through Time"
        .NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = "This slide introduces the topic: The History of Artificial Intelligence."
    End With
    
    ' Add the second slide: The Beginning of AI
    Dim slide2 As Object
    Set slide2 = pptPres.Slides.Add(2, ppLayoutText)
    With slide2
        .Shapes.Title.TextFrame.TextRange.Text = "The Beginning of AI"
        .Shapes.Placeholders(2).TextFrame.TextRange.Text = "The concept of artificial intelligence dates back to ancient myths and legends. However, the formal research on AI began in the mid-20th century. Alan Turing, with his famous Turing Test in 1950, laid the foundation for AI as a field of study."
        .NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = "This slide covers the early ideas and the inception of AI as a field, highlighting Alan Turing's contributions."
    End With
    
    ' Add the third slide: The Early Days of AI Research
    Dim slide3 As Object
    Set slide3 = pptPres.Slides.Add(3, ppLayoutText)
    With slide3
        .Shapes.Title.TextFrame.TextRange.Text = "The Early Days of AI Research"
        .Shapes.Placeholders(2).TextFrame.TextRange.Text = "During the 1950s and 1960s, AI research gained momentum with the development of early algorithms and programs. Notable milestones include the Logic Theorist by Allen Newell and Herbert A. Simon in 1955, and the General Problem Solver in 1957. These were some of the first attempts to mimic human problem-solving."
        .NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = "This slide discusses the early AI programs and their creators, emphasizing their impact on the field."
    End With
    
    ' Add the fourth slide: The AI Winter and Revival
    Dim slide4 As Object
    Set slide4 = pptPres.Slides.Add(4, ppLayoutText)
    With slide4
        .Shapes.Title.TextFrame.TextRange.Text = "The AI Winter and Revival"
        .Shapes.Placeholders(2).TextFrame.TextRange.Text = "In the 1970s and 1980s, AI research faced setbacks due to high expectations and limited computing power, leading to the 'AI Winter'. However, the field revived in the late 1980s and 1990s with advancements in machine learning, expert systems, and neural networks."
        .NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = "This slide explains the challenges faced by AI research during the AI Winter and its subsequent revival."
    End With
    
    ' Add the fifth slide: Modern AI and Future Prospects
    Dim slide5 As Object
    Set slide5 = pptPres.Slides.Add(5, ppLayoutText)
    With slide5
        .Shapes.Title.TextFrame.TextRange.Text = "Modern AI and Future Prospects"
        .Shapes.Placeholders(2).TextFrame.TextRange.Text = "Today, AI is ubiquitous in various fields such as healthcare, finance, and entertainment. Breakthroughs in deep learning and big data have propelled AI to new heights. The future holds promises of even more sophisticated AI, raising both exciting possibilities and ethical considerations."
        .NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = "This slide provides an overview of the current state of AI and its future potential, highlighting key advancements and ethical issues."
    End With
    
    ' Show the presentation
    pptApp.ActivePresentation.SlideShowSettings.Run
End Sub
```

1. 생성된 VBA 스크립트를 빈 파워포인트 슬라이드에서 Alt+F11 키를 눌러 Visual Basic 편집기에서 메크로 실행을 위해 F5 클릭.
2. 팝업창에서 메크로 이름을 demo로 입력하고 '만들기' 버튼 클릭
3. 기존 내용을 삭제 하고 해당 코드를 붙여넣고 실행 하면 파워포인트 슬라이드가 자동으로 생성됩니다.
