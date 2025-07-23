Option Explicit

' フィールドサイズ
Const FIELD_COLS As Integer = 6
Const FIELD_ROWS As Integer = 13
Const VISIBLE_ROWS As Integer = 12

' ぷよ色
Enum PuyoColor
    None = 0
    Red = 1
    Green = 2
    Blue = 3
    Yellow = 4
End Enum

' ぷよペアの構造体
Type PuyoPair
    MainColor As Integer
    SubColor As Integer
    MainRow As Integer
    MainCol As Integer
    SubRow As Integer
    SubCol As Integer
    Rotation As Integer
End Type

' ゲーム変数
Dim field(1 To 13, 1 To 6) As Integer
Dim currentPuyo As PuyoPair
Dim nextPuyo As PuyoPair
Dim gameScore As Long
Dim gameLevel As Integer
Dim chainCount As Integer
Dim maxChain As Integer
Dim dropTimer As Integer
Dim dropSpeed As Integer
Dim gameRunning As Boolean
Dim soundEnabled As Boolean
Dim puyoFixed As Boolean

' ボタン自動作成機能
Sub CreateGameButtons()
    On Error GoTo ErrHandler
    
    ' 既存のボタンを削除
    RemoveGameButtons
    
    ' ActiveXボタンを作成
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' ボタン作成のヘルパー
    CreateButton ws, "btnLeft", "←", Range("K10"), "KeyLeft"
    CreateButton ws, "btnRight", "→", Range("M10"), "KeyRight"
    CreateButton ws, "btnRotate", "↑回転", Range("L9"), "KeyUp"
    CreateButton ws, "btnDown", "↓", Range("L11"), "KeyDown"
    CreateButton ws, "btnDrop", "落下", Range("L12"), "KeySpace"
    CreateButton ws, "btnRestart", "リスタート", Range("K8"), "KeyRestart"
    CreateButton ws, "btnSound", "音ON/OFF", Range("M8"), "KeySound"
    CreateButton ws, "btnInit", "ゲーム開始", Range("K7"), "InitGame"
    
    MsgBox "操作ボタンを自動作成しました！"
    
    Exit Sub
ErrHandler:
    MsgBox "ボタン作成でエラー: " & Err.Description
End Sub

' 単一ボタン作成ヘルパー
Sub CreateButton(ws As Worksheet, btnName As String, caption As String, targetRange As Range, macroName As String)
    On Error Resume Next
    
    ' 既存ボタン削除（同名があれば）
    ws.OLEObjects(btnName).Delete
    
    On Error GoTo ErrHandler
    
    ' ボタン作成
    Dim btn As OLEObject
    Set btn = ws.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
                                Link:=False, _
                                DisplayAsIcon:=False, _
                                Left:=targetRange.Left, _
                                Top:=targetRange.Top, _
                                Width:=targetRange.Width, _
                                Height:=targetRange.Height)
    
    ' ボタン設定
    btn.Name = btnName
    btn.Object.Caption = caption
    btn.Object.Font.Size = 10
    btn.Object.Font.Bold = True
    
    ' マクロ割り当て
    btn.OnAction = macroName
    
    Exit Sub
ErrHandler:
    ' エラーは無視（既存ボタンなど）
End Sub

' 既存ボタン削除
Sub RemoveGameButtons()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' ゲーム用ボタン削除
    ws.OLEObjects("btnLeft").Delete
    ws.OLEObjects("btnRight").Delete
    ws.OLEObjects("btnRotate").Delete
    ws.OLEObjects("btnDown").Delete
    ws.OLEObjects("btnDrop").Delete
    ws.OLEObjects("btnRestart").Delete
    ws.OLEObjects("btnSound").Delete
    ws.OLEObjects("btnInit").Delete
End Sub

' 初期化（ボタン自動作成付き）
Sub InitGame()
    On Error GoTo ErrHandler
    
    ' ボタン自動作成
    CreateGameButtons
    
    ' フィールド初期化
    Dim i As Integer, j As Integer
    For i = 1 To FIELD_ROWS
        For j = 1 To FIELD_COLS
            field(i, j) = PuyoColor.None
        Next j
    Next i
    
    ' ゲーム変数初期化
    gameScore = 0
    gameLevel = 1
    chainCount = 0
    maxChain = 0
    dropTimer = 0
    dropSpeed = 60
    gameRunning = True
    soundEnabled = True
    puyoFixed = False
    
    ' UI初期化
    InitializeUI
    
    ' 最初のぷよペア生成
    GenerateNextPuyo
    SpawnNewPuyo
    
    ' 描画
    DrawField
    DrawUI
    
    MsgBox "ぷよぷよ開始！" & vbCrLf & _
           "ボタン操作またはキーボード操作が可能です"
    
    Exit Sub
ErrHandler:
    MsgBox "InitGameでエラー: " & Err.Description
End Sub

' UI初期化
Sub InitializeUI()
    On Error GoTo ErrHandler
    
    ' フィールドクリア（表示行のみ）
    Range("A1:F12").Interior.ColorIndex = xlNone
    Range("A1:F12").Borders.LineStyle = xlContinuous
    Range("A1:F12").Value = ""
    
    ' セル幅・高さ調整
    Range("A1:F12").ColumnWidth = 3
    Range("A1:F12").RowHeight = 20
    
    ' スコア表示エリア
    Range("H1").Value = "スコア:"
    Range("I1").Value = gameScore
    Range("H2").Value = "レベル:"
    Range("I2").Value = gameLevel
    Range("H3").Value = "連鎖:"
    Range("I3").Value = chainCount
    Range("H4").Value = "最大連鎖:"
    Range("I4").Value = maxChain
    
    ' 効果音状態表示
    Range("H5").Value = "効果音:"
    Range("I5").Value = IIf(soundEnabled, "ON", "OFF")
    
    ' 次のぷよ表示エリア
    Range("H15").Value = "次のぷよ:"
    Range("H16:I17").Interior.ColorIndex = xlNone
    Range("H16:I17").Borders.LineStyle = xlContinuous
    Range("H16:I17").Value = ""
    
    ' 操作説明
    Range("H19").Value = "■操作方法■"
    Range("H20").Value = "ボタンクリック"
    Range("H21").Value = "または"
    Range("H22").Value = "キーボード操作"
    
    Exit Sub
ErrHandler:
    MsgBox "InitializeUIでエラー: " & Err.Description
End Sub

' 効果音再生
Sub PlaySound(Optional ByVal soundType As String = "normal")
    On Error Resume Next
    
    If Not soundEnabled Then Exit Sub
    
    Select Case soundType
        Case "chain"
            Beep
        Case "levelup"
            Beep
            Application.Wait (Now + TimeValue("0:00:00.1"))
            Beep
        Case "gameover"
            Beep
            Application.Wait (Now + TimeValue("0:00:00.2"))
            Beep
        Case "erase"
            Beep
    End Select
End Sub

' 効果音ON/OFF切り替え
Sub ToggleSound()
    On Error GoTo ErrHandler
    
    soundEnabled = Not soundEnabled
    Range("I5").Value = IIf(soundEnabled, "ON", "OFF")
    
    If soundEnabled Then PlaySound "normal"
    
    Exit Sub
ErrHandler:
    MsgBox "ToggleSoundでエラー: " & Err.Description
End Sub

' 次のぷよ生成
Sub GenerateNextPuyo()
    On Error GoTo ErrHandler
    
    With nextPuyo
        .MainColor = Int(Rnd() * 4) + 1
        .SubColor = Int(Rnd() * 4) + 1
        .MainRow = 2
        .MainCol = 3
        .SubRow = 1
        .SubCol = 3
        .Rotation = 0
    End With
    
    Exit Sub
ErrHandler:
    MsgBox "GenerateNextPuyoでエラー: " & Err.Description
End Sub

' 新しいぷよペアをフィールドに配置
Sub SpawnNewPuyo()
    On Error GoTo ErrHandler
    
    currentPuyo = nextPuyo
    puyoFixed = False
    
    With currentPuyo
        .MainRow = 2
        .MainCol = 3
        .SubRow = 1
        .SubCol = 3
        .Rotation = 0
    End With
    
    If field(currentPuyo.MainRow, currentPuyo.MainCol) <> PuyoColor.None Or _
       field(currentPuyo.SubRow, currentPuyo.SubCol) <> PuyoColor.None Then
        GameOver
        Exit Sub
    End If
    
    field(currentPuyo.MainRow, currentPuyo.MainCol) = currentPuyo.MainColor
    field(currentPuyo.SubRow, currentPuyo.SubCol) = currentPuyo.SubColor
    
    GenerateNextPuyo
    
    Exit Sub
ErrHandler:
    MsgBox "SpawnNewPuyoでエラー: " & Err.Description
End Sub

' セルが範囲内かチェック
Function IsValidPosition(ByVal row As Integer, ByVal col As Integer) As Boolean
    IsValidPosition = (row >= 1 And row <= FIELD_ROWS And col >= 1 And col <= FIELD_COLS)
End Function

' 指定位置が空かチェック
Function IsEmptyPosition(ByVal row As Integer, ByVal col As Integer) As Boolean
    On Error GoTo ErrHandler
    
    If Not IsValidPosition(row, col) Then
        IsEmptyPosition = False
        Exit Function
    End If
    
    If (row = currentPuyo.MainRow And col = currentPuyo.MainCol) Or _
       (row = currentPuyo.SubRow And col = currentPuyo.SubCol) Then
        IsEmptyPosition = True
    Else
        IsEmptyPosition = (field(row, col) = PuyoColor.None)
    End If
    
    Exit Function
ErrHandler:
    IsEmptyPosition = False
End Function

' フィールド描画
Sub DrawField()
    On Error GoTo ErrHandler
    
    Dim i As Integer, j As Integer
    
    For i = 2 To FIELD_ROWS
        For j = 1 To FIELD_COLS
            Dim cellRow As Integer
            cellRow = i - 1
            
            Select Case field(i, j)
                Case PuyoColor.None
                    Cells(cellRow, j).Interior.ColorIndex = xlNone
                    Cells(cellRow, j).Value = ""
                Case PuyoColor.Red
                    Cells(cellRow, j).Interior.Color = RGB(255, 100, 100)
                    Cells(cellRow, j).Value = "●"
                Case PuyoColor.Green
                    Cells(cellRow, j).Interior.Color = RGB(100, 255, 100)
                    Cells(cellRow, j).Value = "●"
                Case PuyoColor.Blue
                    Cells(cellRow, j).Interior.Color = RGB(100, 100, 255)
                    Cells(cellRow, j).Value = "●"
                Case PuyoColor.Yellow
                    Cells(cellRow, j).Interior.Color = RGB(255, 255, 100)
                    Cells(cellRow, j).Value = "●"
            End Select
            
            With Cells(cellRow, j)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Size = 14
                .Font.Bold = True
            End With
        Next j
    Next i
    
    Exit Sub
ErrHandler:
    MsgBox "DrawFieldでエラー: " & Err.Description
End Sub

' UI描画
Sub DrawUI()
    On Error GoTo ErrHandler
    
    Range("I1").Value = gameScore
    Range("I2").Value = gameLevel
    Range("I3").Value = chainCount
    Range("I4").Value = maxChain
    Range("I5").Value = IIf(soundEnabled, "ON", "OFF")
    
    DrawNextPuyo
    
    Exit Sub
ErrHandler:
    MsgBox "DrawUIでエラー: " & Err.Description
End Sub

' 次のぷよ描画
Sub DrawNextPuyo()
    On Error GoTo ErrHandler
    
    Range("H16:I17").Interior.ColorIndex = xlNone
    Range("H16:I17").Value = ""
    
    With Range("I17")
        Select Case nextPuyo.MainColor
            Case PuyoColor.Red: .Interior.Color = RGB(255, 100, 100)
            Case PuyoColor.Green: .Interior.Color = RGB(100, 255, 100)
            Case PuyoColor.Blue: .Interior.Color = RGB(100, 100, 255)
            Case PuyoColor.Yellow: .Interior.Color = RGB(255, 255, 100)
        End Select
        .Value = "●"
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
    End With
    
    With Range("I16")
        Select Case nextPuyo.SubColor
            Case PuyoColor.Red: .Interior.Color = RGB(255, 100, 100)
            Case PuyoColor.Green: .Interior.Color = RGB(100, 255, 100)
            Case PuyoColor.Blue: .Interior.Color = RGB(100, 100, 255)
            Case PuyoColor.Yellow: .Interior.Color = RGB(255, 255, 100)
        End Select
        .Value = "●"
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
    End With
    
    Exit Sub
ErrHandler:
    MsgBox "DrawNextPuyoでエラー: " & Err.Description
End Sub

' 左移動
Sub MoveLeft()
    On Error GoTo ErrHandler
    
    If Not gameRunning Or puyoFixed Then Exit Sub
    
    Dim newMainCol As Integer, newSubCol As Integer
    newMainCol = currentPuyo.MainCol - 1
    newSubCol = currentPuyo.SubCol - 1
    
    If IsEmptyPosition(currentPuyo.MainRow, newMainCol) And _
       IsEmptyPosition(currentPuyo.SubRow, newSubCol) Then
        
        field(currentPuyo.MainRow, currentPuyo.MainCol) = PuyoColor.None
        field(currentPuyo.SubRow, currentPuyo.SubCol) = PuyoColor.None
        
        currentPuyo.MainCol = newMainCol
        currentPuyo.SubCol = newSubCol
        field(currentPuyo.MainRow, currentPuyo.MainCol) = currentPuyo.MainColor
        field(currentPuyo.SubRow, currentPuyo.SubCol) = currentPuyo.SubColor
        
        DrawField
    End If
    
    Exit Sub
ErrHandler:
    MsgBox "MoveLeftでエラー: " & Err.Description
End Sub

' 右移動
Sub MoveRight()
    On Error GoTo ErrHandler
    
    If Not gameRunning Or puyoFixed Then Exit Sub
    
    Dim newMainCol As Integer, newSubCol As Integer
    newMainCol = currentPuyo.MainCol + 1
    newSubCol = currentPuyo.SubCol + 1
    
    If IsEmptyPosition(currentPuyo.MainRow, newMainCol) And _
       IsEmptyPosition(currentPuyo.SubRow, newSubCol) Then
        
        field(currentPuyo.MainRow, currentPuyo.MainCol) = PuyoColor.None
        field(currentPuyo.SubRow, currentPuyo.SubCol) = PuyoColor.None
        
        currentPuyo.MainCol = newMainCol
        currentPuyo.SubCol = newSubCol
        field(currentPuyo.MainRow, currentPuyo.MainCol) = currentPuyo.MainColor
        field(currentPuyo.SubRow, currentPuyo.SubCol) = currentPuyo.SubColor
        
        DrawField
    End If
    
    Exit Sub
ErrHandler:
    MsgBox "MoveRightでエラー: " & Err.Description
End Sub

' 回転
Sub RotatePuyo()
    On Error GoTo ErrHandler
    
    If Not gameRunning Or puyoFixed Then Exit Sub
    
    Dim newSubRow As Integer, newSubCol As Integer
    Dim newRotation As Integer
    
    newRotation = (currentPuyo.Rotation + 1) Mod 4
    
    Select Case newRotation
        Case 0
            newSubRow = currentPuyo.MainRow - 1
            newSubCol = currentPuyo.MainCol
        Case 1
            newSubRow = currentPuyo.MainRow
            newSubCol = currentPuyo.MainCol + 1
        Case 2
            newSubRow = currentPuyo.MainRow + 1
            newSubCol = currentPuyo.MainCol
        Case 3
            newSubRow = currentPuyo.MainRow
            newSubCol = currentPuyo.MainCol - 1
    End Select
    
    If IsEmptyPosition(newSubRow, newSubCol) Then
        field(currentPuyo.SubRow, currentPuyo.SubCol) = PuyoColor.None
        
        currentPuyo.SubRow = newSubRow
        currentPuyo.SubCol = newSubCol
        currentPuyo.Rotation = newRotation
        field(currentPuyo.SubRow, currentPuyo.SubCol) = currentPuyo.SubColor
        
        DrawField
    End If
    
    Exit Sub
ErrHandler:
    MsgBox "RotatePuyoでエラー: " & Err.Description
End Sub

' 1段落下
Sub DropPuyoOneStep()
    On Error GoTo ErrHandler
    
    If Not gameRunning Or puyoFixed Then Exit Sub
    
    Dim newMainRow As Integer, newSubRow As Integer
    newMainRow = currentPuyo.MainRow + 1
    newSubRow = currentPuyo.SubRow + 1
    
    If IsEmptyPosition(newMainRow, currentPuyo.MainCol) And _
       IsEmptyPosition(newSubRow, currentPuyo.SubCol) Then
        
        field(currentPuyo.MainRow, currentPuyo.MainCol) = PuyoColor.None
        field(currentPuyo.SubRow, currentPuyo.SubCol) = PuyoColor.None
        
        currentPuyo.MainRow = newMainRow
        currentPuyo.SubRow = newSubRow
        field(currentPuyo.MainRow, currentPuyo.MainCol) = currentPuyo.MainColor
        field(currentPuyo.SubRow, currentPuyo.SubCol) = currentPuyo.SubColor
        
        DrawField
    Else
        FixPuyo
    End If
    
    Exit Sub
ErrHandler:
    MsgBox "DropPuyoOneStepでエラー: " & Err.Description
End Sub

' 最下段まで落下
Sub DropPuyoToBottom()
    On Error GoTo ErrHandler
    
    If Not gameRunning Or puyoFixed Then Exit Sub
    
    Do While True
        Dim newMainRow As Integer, newSubRow As Integer
        newMainRow = currentPuyo.MainRow + 1
        newSubRow = currentPuyo.SubRow + 1
        
        If IsEmptyPosition(newMainRow, currentPuyo.MainCol) And _
           IsEmptyPosition(newSubRow, currentPuyo.SubCol) Then
            
            field(currentPuyo.MainRow, currentPuyo.MainCol) = PuyoColor.None
            field(currentPuyo.SubRow, currentPuyo.SubCol) = PuyoColor.None
            
            currentPuyo.MainRow = newMainRow
            currentPuyo.SubRow = newSubRow
            field(currentPuyo.MainRow, currentPuyo.MainCol) = currentPuyo.MainColor
            field(currentPuyo.SubRow, currentPuyo.SubCol) = currentPuyo.SubColor
        Else
            Exit Do
        End If
    Loop
    
    DrawField
    FixPuyo
    
    Exit Sub
ErrHandler:
    MsgBox "DropPuyoToBottomでエラー: " & Err.Description
End Sub

' ぷよ固定
Sub FixPuyo()
    On Error GoTo ErrHandler
    
    puyoFixed = True
    chainCount = 0
    
    CheckAndEraseChain
    
    If chainCount > maxChain Then maxChain = chainCount
    
    ApplyGravity
    
    SpawnNewPuyo
    DrawField
    DrawUI
    
    Exit Sub
ErrHandler:
    MsgBox "FixPuyoでエラー: " & Err.Description
End Sub

' 連鎖チェック＆消去
Sub CheckAndEraseChain()
    On Error GoTo ErrHandler
    
    Dim hasErasure As Boolean
    
    Do
        hasErasure = False
        Dim totalErased As Integer
        totalErased = 0
        
        Dim targetColor As Integer
        For targetColor = 1 To 4
            Dim eraseResult As Integer
            eraseResult = CheckAndEraseColor(targetColor)
            
            If eraseResult > 0 Then
                hasErasure = True
                totalErased = totalErased + eraseResult
            End If
        Next targetColor
        
        If hasErasure Then
            chainCount = chainCount + 1
            PlaySound "erase"
            
            Dim chainBonus As Integer
            chainBonus = IIf(chainCount = 1, 1, chainCount * chainCount)
            gameScore = gameScore + totalErased * 10 * chainBonus
            
            If gameScore > gameLevel * 1000 Then
                gameLevel = gameLevel + 1
                dropSpeed = dropSpeed - 5
                If dropSpeed < 10 Then dropSpeed = 10
                PlaySound "levelup"
            End If
            
            If chainCount > 1 Then PlaySound "chain"
            
            DrawField
            DrawUI
            DoEvents
            Application.Wait (Now + TimeValue("0:00:00.3"))
            
            ApplyGravity
            DrawField
            DoEvents
            Application.Wait (Now + TimeValue("0:00:00.2"))
        End If
        
    Loop While hasErasure
    
    Exit Sub
ErrHandler:
    MsgBox "CheckAndEraseChainでエラー: " & Err.Description
End Sub

' 指定色の連結ぷよをチェック＆消去
Function CheckAndEraseColor(ByVal targetColor As Integer) As Integer
    On Error GoTo ErrHandler
    
    Dim erasedCount As Integer
    erasedCount = 0
    
    Dim checked(1 To 13, 1 To 6) As Boolean
    Dim i As Integer, j As Integer
    
    For i = 1 To FIELD_ROWS
        For j = 1 To FIELD_COLS
            checked(i, j) = False
        Next j
    Next i
    
    For i = 1 To FIELD_ROWS
        For j = 1 To FIELD_COLS
            If field(i, j) = targetColor And Not checked(i, j) Then
                Dim groupCount As Integer
                groupCount = CountConnectedGroup(i, j, targetColor, checked)
                
                If groupCount >= 4 Then
                    EraseConnectedGroup i, j, targetColor
                    erasedCount = erasedCount + groupCount
                End If
            End If
        Next j
    Next i
    
    CheckAndEraseColor = erasedCount
    Exit Function
    
ErrHandler:
    CheckAndEraseColor = 0
End Function

' 連結グループの数をカウント
Function CountConnectedGroup(ByVal startRow As Integer, ByVal startCol As Integer, ByVal targetColor As Integer, ByRef checked() As Boolean) As Integer
    On Error GoTo ErrHandler
    
    Dim count As Integer
    count = 0
    
    Dim stackRows(1 To 78) As Integer
    Dim stackCols(1 To 78) As Integer
    Dim stackTop As Integer
    stackTop = 0
    
    stackTop = stackTop + 1
    stackRows(stackTop) = startRow
    stackCols(stackTop) = startCol
    checked(startRow, startCol) = True
    
    Do While stackTop > 0
        Dim currentRow As Integer, currentCol As Integer
        currentRow = stackRows(stackTop)
        currentCol = stackCols(stackTop)
        stackTop = stackTop - 1
        count = count + 1
        
        Dim dr(1 To 4) As Integer, dc(1 To 4) As Integer
        dr(1) = -1: dc(1) = 0
        dr(2) = 1: dc(2) = 0
        dr(3) = 0: dc(3) = -1
        dr(4) = 0: dc(4) = 1
        
        Dim d As Integer
        For d = 1 To 4
            Dim newRow As Integer, newCol As Integer
            newRow = currentRow + dr(d)
            newCol = currentCol + dc(d)
            
            If IsValidPosition(newRow, newCol) Then
                If field(newRow, newCol) = targetColor And Not checked(newRow, newCol) Then
                    stackTop = stackTop + 1
                    stackRows(stackTop) = newRow
                    stackCols(stackTop) = newCol
                    checked(newRow, newCol) = True
                End If
            End If
        Next d
    Loop
    
    CountConnectedGroup = count
    Exit Function
    
ErrHandler:
    CountConnectedGroup = 0
End Function

' 連結グループを消去
Sub EraseConnectedGroup(ByVal startRow As Integer, ByVal startCol As Integer, ByVal targetColor As Integer)
    On Error GoTo ErrHandler
    
    Dim stackRows(1 To 78) As Integer
    Dim stackCols(1 To 78) As Integer
    Dim stackTop As Integer
    stackTop = 0
    
    stackTop = stackTop + 1
    stackRows(stackTop) = startRow
    stackCols(stackTop) = startCol
    
    Do While stackTop > 0
        Dim currentRow As Integer, currentCol As Integer
        currentRow = stackRows(stackTop)
        currentCol = stackCols(stackTop)
        stackTop = stackTop - 1
        
        If field(currentRow, currentCol) = targetColor Then
            field(currentRow, currentCol) = PuyoColor.None
            
            Dim dr(1 To 4) As Integer, dc(1 To 4) As Integer
            dr(1) = -1: dc(1) = 0
            dr(2) = 1: dc(2) = 0
            dr(3) = 0: dc(3) = -1
            dr(4) = 0: dc(4) = 1
            
            Dim d As Integer
            For d = 1 To 4
                Dim newRow As Integer, newCol As Integer
                newRow = currentRow + dr(d)
                newCol = currentCol + dc(d)
                
                If IsValidPosition(newRow, newCol) Then
                    If field(newRow, newCol) = targetColor Then
                        stackTop = stackTop + 1
                        stackRows(stackTop) = newRow
                        stackCols(stackTop) = newCol
                    End If
                End If
            Next d
        End If
    Loop
    
    Exit Sub
ErrHandler:
    MsgBox "EraseConnectedGroupでエラー: " & Err.Description
End Sub

' 重力適用
Sub ApplyGravity()
    On Error GoTo ErrHandler
    
    Dim col As Integer, row As Integer, r As Integer
    
    For col = 1 To FIELD_COLS
        For row = FIELD_ROWS To 2 Step -1
            If field(row, col) = PuyoColor.None Then
                For r = row - 1 To 1 Step -1
                    If field(r, col) <> PuyoColor.None Then
                        field(row, col) = field(r, col)
                        field(r, col) = PuyoColor.None
                        Exit For
                    End If
                Next r
            End If
        Next row
    Next col
    
    Exit Sub
ErrHandler:
    MsgBox "ApplyGravityでエラー: " & Err.Description
End Sub

' ゲームオーバー
Sub GameOver()
    On Error GoTo ErrHandler
    
    gameRunning = False
    PlaySound "gameover"
    
    MsgBox "ゲームオーバー！" & vbCrLf & _
           "最終スコア: " & gameScore & vbCrLf & _
           "最終レベル: " & gameLevel & vbCrLf & _
           "最大連鎖: " & maxChain & vbCrLf & _
           "効果音: " & IIf(soundEnabled, "ON", "OFF") & vbCrLf & _
           "リスタートボタンで再開可能"
    
    Exit Sub
ErrHandler:
    MsgBox "GameOverでエラー: " & Err.Description
End Sub

' キー操作用のプロシージャ
Sub KeyLeft()
    MoveLeft
End Sub

Sub KeyRight()
    MoveRight
End Sub

Sub KeyUp()
    RotatePuyo
End Sub

Sub KeyDown()
    DropPuyoOneStep
End Sub

Sub KeySpace()
    DropPuyoToBottom
End Sub

Sub KeyRestart()
    InitGame
End Sub

Sub KeySound()
    ToggleSound
End Sub