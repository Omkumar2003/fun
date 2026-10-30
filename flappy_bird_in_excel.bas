Option Explicit


Global gSheet As Object
Global gTimer As Object
Global gListener As Object
Global gTimerRunning As Boolean
Global gCallbackBusy As Boolean
Global gDrawInitialized As Boolean
Global gLastColors(1 To 24, 1 To 40) As Long

Global gBird As Integer
Global gVelocity As Integer
Global gPipe1 As Integer
Global gPipe2 As Integer
Global gGap1 As Integer
Global gGap2 As Integer
Global gScore As Integer
Global gPassed1 As Boolean
Global gPassed2 As Boolean
Global gPlaying As Boolean
Global gPaused As Boolean
Global gGameOver As Boolean

Const W = 40
Const H = 24
Const BX = 8
Const GAP = 7
Const FRAME_DELAY = 60

Sub StartFlappy()
    Dim doc As Object
    Dim r As Integer
    Dim c As Integer

    StopGameTimer

    doc = ThisComponent
    gSheet = doc.CurrentController.ActiveSheet

    On Error Resume Next
    doc.CurrentController.ViewSettings.ZoomValue = 100
    On Error GoTo 0

    For c = 0 To W - 1
        gSheet.Columns.getByIndex(c).Width = 650
    Next c

    For r = 0 To H - 1
        gSheet.Rows.getByIndex(r).Height = 650
    Next r

    gSheet.getCellRangeByName("A1:AN24").clearContents(1023)

    Randomize
    gBird = 11
    gVelocity = 0
    gPipe1 = 24
    gPipe2 = 39
    gGap1 = 7 + Int(Rnd() * 9)
    gGap2 = 7 + Int(Rnd() * 9)
    gScore = 0
    gPassed1 = False
    gPassed2 = False
    gPaused = False
    gGameOver = False
    gPlaying = True
    gCallbackBusy = False
    gDrawInitialized = False

    CreateGameButtons
    On Error Resume Next
    ThisComponent.CurrentController.FormDesignMode = False
    On Error GoTo 0
    DrawBoard
    UpdateHUD
    StartGameTimer
End Sub

Sub StartGameTimer()
    Dim stage As String
    Dim errNumber As Long
    Dim errMessage As String

    On Error GoTo TimerError
    gTimerRunning = False
    gCallbackBusy = False

    stage = "creating com.sun.star.awt.AsyncCallback service"
    gTimer = CreateUnoService("com.sun.star.awt.AsyncCallback")

    If IsNull(gTimer) Or IsEmpty(gTimer) Then
        MsgBox "AsyncCallback service create nahi hua." & Chr(10) & _
               "Is LibreOffice build mein callback service available nahi hai.", _
               16, "Flappy Bird - Timer"
        Exit Sub
    End If

    stage = "creating XCallback listener"
    gListener = CreateUnoListener("FlappyTimer_", "com.sun.star.awt.XCallback")

    If IsNull(gListener) Or IsEmpty(gListener) Then
        MsgBox "XCallback listener create nahi hua.", 16, "Flappy Bird - Timer"
        gTimer = Nothing
        Exit Sub
    End If

    gTimerRunning = True
    stage = "queueing first callback"
    gTimer.addCallback(gListener, 0)
    Exit Sub

TimerError:
    errNumber = Err
    errMessage = Error$
    On Error Resume Next
    gTimerRunning = False
    gCallbackBusy = False
    gTimer = Nothing
    gListener = Nothing
    On Error GoTo 0

    MsgBox "Game loop start nahi hua." & Chr(10) & _
           "Stage: " & stage & Chr(10) & _
           "Error " & errNumber & ": " & errMessage, _
           16, "Flappy Bird - Timer"
End Sub

Sub StopGameTimer()
    gTimerRunning = False
    gCallbackBusy = False
    On Error Resume Next
    gTimer = Nothing
    gListener = Nothing
    On Error GoTo 0
End Sub

Sub FlappyTimer_notify(aData)
    On Error GoTo CallbackError

    If Not gTimerRunning Then Exit Sub
    If gCallbackBusy Then Exit Sub

    gCallbackBusy = True

    If gPlaying And Not gPaused And Not gGameOver Then
        AdvanceGame
    End If

    If gTimerRunning And gPlaying And Not gGameOver Then
        Wait FRAME_DELAY
        If gTimerRunning And gPlaying And Not gGameOver Then
            gCallbackBusy = False
            gTimer.addCallback(gListener, 0)
            Exit Sub
        End If
    End If

    gCallbackBusy = False
    Exit Sub

CallbackError:
    Dim errNumber As Long
    Dim errMessage As String
    errNumber = Err
    errMessage = Error$
    gCallbackBusy = False
    gTimerRunning = False
    MsgBox "Game callback error " & errNumber & ": " & errMessage, _
           16, "Flappy Bird - Timer"
End Sub

Sub FlapBird()
    If Not gPlaying Or gPaused Or gGameOver Then Exit Sub
    gVelocity = -3
    AdvanceGame
End Sub

Sub PauseGame()
    If Not gPlaying Or gGameOver Then Exit Sub
    gPaused = Not gPaused
    UpdateHUD
End Sub

Sub RestartGame()
    StartFlappy
End Sub

Sub AdvanceGame()
    If Not gPlaying Or gPaused Or gGameOver Then Exit Sub

    gVelocity = gVelocity + 1
    If gVelocity > 2 Then gVelocity = 2
    gBird = gBird + gVelocity

    gPipe1 = gPipe1 - 1
    gPipe2 = gPipe2 - 1

    If gPipe1 < 1 Then
        gPipe1 = W - 1
        gGap1 = 4 + Int(Rnd() * 12)
        gPassed1 = False
    End If

    If gPipe2 < 1 Then
        gPipe2 = W - 1
        gGap2 = 4 + Int(Rnd() * 12)
        gPassed2 = False
    End If

    If Not gPassed1 And gPipe1 + 1 < BX Then
        gScore = gScore + 1
        gPassed1 = True
    End If

    If Not gPassed2 And gPipe2 + 1 < BX Then
        gScore = gScore + 1
        gPassed2 = True
    End If

    If gBird < 1 Or gBird >= H - 1 Then
        GameOver
        Exit Sub
    End If

    If BX >= gPipe1 And BX < gPipe1 + 2 Then
        If gBird < gGap1 Or gBird >= gGap1 + GAP Then
            GameOver
            Exit Sub
        End If
    End If

    If BX >= gPipe2 And BX < gPipe2 + 2 Then
        If gBird < gGap2 Or gBird >= gGap2 + GAP Then
            GameOver
            Exit Sub
        End If
    End If

    DrawBoard
    UpdateHUD
End Sub

Sub GameOver()
    gGameOver = True
    gPlaying = False
    StopGameTimer
    DrawBoard
    UpdateHUD
End Sub

Sub DrawBoard()
    Dim c As Integer
    Dim r As Integer
    Dim cell As Object
    Dim colorValue As Long

    On Error GoTo DrawError

    For r = 1 To H
        For c = 1 To W
            colorValue = RGB(135, 206, 235)

            If r >= H - 1 Then colorValue = RGB(205, 170, 95)

            If c >= gPipe1 And c < gPipe1 + 2 Then
                If r < gGap1 Or r >= gGap1 + GAP Then
                    colorValue = RGB(35, 170, 65)
                End If
            End If

            If c >= gPipe2 And c < gPipe2 + 2 Then
                If r < gGap2 Or r >= gGap2 + GAP Then
                    colorValue = RGB(35, 170, 65)
                End If
            End If

            If c = BX And r = gBird Then colorValue = RGB(255, 220, 20)

            ' Only write cells whose colour actually changed.
            If Not gDrawInitialized Then
                cell = gSheet.getCellByPosition(c - 1, r - 1)
                cell.CellBackColor = colorValue
            ElseIf gLastColors(r, c) <> colorValue Then
                cell = gSheet.getCellByPosition(c - 1, r - 1)
                cell.CellBackColor = colorValue
            End If

            gLastColors(r, c) = colorValue
        Next c
    Next r

    gDrawInitialized = True
    On Error GoTo 0
    Exit Sub

DrawError:
    Dim drawErr As Long
    Dim drawMsg As String
    drawErr = Err
    drawMsg = Error$
    On Error GoTo 0
    MsgBox "Board draw error " & drawErr & ": " & drawMsg, 16, "Flappy Bird"
End Sub

Sub UpdateHUD()
    On Error Resume Next

    gSheet.getCellRangeByName("AP1:AU1").merge(True)
    gSheet.getCellRangeByName("AP1").String = "FLAPPY BIRD"
    gSheet.getCellRangeByName("AP1").CellBackColor = RGB(38, 50, 56)
    gSheet.getCellRangeByName("AP1").CharColor = RGB(255, 255, 255)
    gSheet.getCellRangeByName("AP1").CharHeight = 16

    gSheet.getCellRangeByName("AP3").String = "SCORE: " & gScore
    gSheet.getCellRangeByName("AP3").CharHeight = 14

    If gGameOver Then
        gSheet.getCellRangeByName("AP4").String = "GAME OVER"
    ElseIf gPaused Then
        gSheet.getCellRangeByName("AP4").String = "PAUSED"
    Else
        gSheet.getCellRangeByName("AP4").String = "PLAYING"
    End If

    gSheet.getCellRangeByName("AP6").String = "Controls"
    gSheet.getCellRangeByName("AP7").String = "Use buttons below"
    gSheet.getCellRangeByName("AP8").String = "FLAP = jump"
    gSheet.getCellRangeByName("AP9").String = "PAUSE = pause/resume"
    gSheet.getCellRangeByName("AP10").String = "RESTART = new game"

    On Error GoTo 0
End Sub

Sub CreateGameButtons()
    Dim doc As Object
    Dim page As Object
    Dim forms As Object
    Dim form As Object
    Dim i As Long
    Dim shp As Object
    Dim ctlName As String

    doc = ThisComponent
    page = gSheet.DrawPage
    forms = page.Forms

    ' Remove the previous control shapes too (Restart creates the UI again).
    On Error Resume Next
    For i = page.getCount() - 1 To 0 Step -1
        shp = page.getByIndex(i)
        If shp.supportsService("com.sun.star.drawing.ControlShape") Then
            ctlName = ""
            ctlName = shp.Control.Name
            If ctlName = "btnFlap" Or ctlName = "btnPause" Or ctlName = "btnRestart" Then
                page.remove(shp)
            End If
        End If
    Next i

    If forms.hasByName("FlappyControls") Then forms.removeByName("FlappyControls")
    On Error GoTo ButtonError

    form = doc.createInstance("com.sun.star.form.component.Form")
    forms.insertByName("FlappyControls", form)

    AddButton form, page, "btnFlap", "FLAP", "FlapBird", 27000, 6500
    AddButton form, page, "btnPause", "PAUSE", "PauseGame", 27000, 8000
    AddButton form, page, "btnRestart", "RESTART", "RestartGame", 27000, 9500

    ' Controls do not respond to clicks while Form Design Mode is enabled.
    doc.CurrentController.FormDesignMode = False
    Exit Sub

ButtonError:
    MsgBox "Buttons create nahi hue." & Chr(10) & _
           "Error " & Err & ": " & Error$, 16, "Flappy Bird - Buttons"
End Sub

Sub AddButton(form As Object, page As Object, _
              controlName As String, labelText As String, _
              macroName As String, x As Long, y As Long)
    Dim model As Object
    Dim shape As Object
    Dim pos As Object
    Dim sz As Object
    Dim scriptEvent As Object
    Dim idx As Long

    model = ThisComponent.createInstance("com.sun.star.form.component.CommandButton")
    model.Name = controlName
    model.Label = labelText
    model.Enabled = True
    form.insertByName(controlName, model)

    idx = form.getCount() - 1
    scriptEvent = CreateUnoStruct("com.sun.star.script.ScriptEventDescriptor")
    scriptEvent.ListenerType = "com.sun.star.awt.XActionListener"
    scriptEvent.EventMethod = "actionPerformed"
    scriptEvent.AddListenerParam = ""
    scriptEvent.ScriptType = "Script"
    scriptEvent.ScriptCode = "vnd.sun.star.script:Standard.Module1." & macroName & "?language=Basic&location=document"
    form.registerScriptEvent(idx, scriptEvent)

    shape = ThisComponent.createInstance("com.sun.star.drawing.ControlShape")
    shape.Control = model

    pos = CreateUnoStruct("com.sun.star.awt.Point")
    sz = CreateUnoStruct("com.sun.star.awt.Size")
    pos.X = x
    pos.Y = y
    sz.Width = 3200
    sz.Height = 1100

    shape.Position = pos
    shape.Size = sz
    page.add(shape)
End Sub
