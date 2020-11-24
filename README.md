# VBA-Pen_Soccer
![Pen_Soccer](https://user-images.githubusercontent.com/66747535/100058194-3d9d5980-2e6c-11eb-82ea-22767803fa8a.gif)

엑셀에서 VBA 매크로를 통해 실행할 수 있는 펜 축구 게임이다.

## 적용법
1. VBA 편집창에 들어간다.
2. 모듈이 아니라 적용할 시트의 코드 창에 아래의 코드를 모두 넣는다.
3. 매크로 직접 실행으로 Format 실행

## 코드
<details>
    <summary>코드보기</summary>

```
'(1,1) 임시셀 (1,2) 현재위치행 (1,3) 현재위치열 (1,4) 현재차례 1p = 0. 2p = 1

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    clr_next = RGB(101, 255, 101)
    Cells(4, 20) = "" '메세지창 비우기
    
    Cells(1, 1) = "=rows(" + Selection.Address + ")" '(1,1)임시셸
    m = Cells(1, 1) '행 크기
    Cells(1, 1) = "=columns(" + Selection.Address + ")"
    n = Cells(1, 1) '열 크기
    
    If m > 1 Or n > 1 Then '여러셀 동시 선택
        Cells(4, 20) = "잘못 누름"
    ElseIf Selection.Address = Cells(2, 19).Address Then '게임시작
        Start
    ElseIf Selection.Address = Cells(2, 20).Address Then '무르기, 저장된 시점 불러옴
        Range(Cells(1, 21), Cells(21, 40)).Copy Range(Cells(1, 1), Cells(21, 20))
        Cells(4, 20) = "back"
        Where
    ElseIf Selection.Interior.Color = clr_next Then '연결할 셀 클릭
        Click
    Else '나머지 부분 클릭
        Cells(4, 20) = "잘못 누름"
    End If
    
    Cells(1, 1).Select
End Sub

Function Click() '연결할 셀 클릭
    clr_line = RGB(0, 0, 0)
    clr_1p = RGB(255, 203, 203)
    clr_2p = RGB(203, 203, 255)
    
    '--턴 확인---------------------------
    If Cells(1, 4) = 0 Then 'IP의 차례
        clr_p = clr_1p
    Else '2P의 차례
        clr_p = clr_2p
    End If
    
    '--선 긋기---------------------------
    a = Cells(1, 2)
    b = Cells(1, 3)
    x = Selection.Row
    y = Selection.Column
    
    Set 클릭 = Cells((a + x) / 2, (b + y) / 2)
    k = (a - x) * (b - y)
    
    If k = 0 Then '가로 혹은 세로
        Cells((a + x) / 2, (b + y) / 2).Interior.Color = clr_p
        Cells((a + x) / 2, (b + y) / 2).Font.Color = clr_p
        Cells((a + x) / 2, (b + y) / 2) = 1 '선이 그어진 곳으로 설정
    ElseIf k > 0 Then '대각선 아래
        With Cells((a + x) / 2, (b + y) / 2).Borders(xlDiagonalDown)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .Color = clr_p
        End With
        Cells((a + x) / 2, (b + y) / 2).Font.Color = Cells((a + x) / 2, (b + y) / 2).Interior.Color
        Cells((a + x) / 2, (b + y) / 2) = Cells((a + x) / 2, (b + y) / 2) + 2 '선이 그어진 곳으로 설정
    ElseIf k < 0 Then '대각선 위
        With Cells((a + x) / 2, (b + y) / 2).Borders(xlDiagonalUp)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .Color = clr_p
        End With
        Cells((a + x) / 2, (b + y) / 2).Font.Color = Cells((a + x) / 2, (b + y) / 2).Interior.Color
        Cells((a + x) / 2, (b + y) / 2) = Cells((a + x) / 2, (b + y) / 2) + 3 '선이 그어진 곳으로 설정
    End If
    
    Cells(1, 2) = x '현재위치 변경
    Cells(1, 3) = y
    
    '--이전의 점 제거------------------------------
    Call Re(a - 2, b - 2)
    Call Re(a - 2, b)
    Call Re(a - 2, b + 2)
    
    Call Re(a, b - 2)
    Cells(a, b).Interior.Color = clr_line
    Call Re(a, b + 2)
    
    Call Re(a + 2, b - 2)
    Call Re(a + 2, b)
    Call Re(a + 2, b + 2)

    '--승패확인-------------------------------------
    If x = 1 Then '2P WIN
        Cells(4, 19) = "2P WIN!"
        Exit Function
    ElseIf x = 21 Then '1P WIN
        Cells(4, 19) = "1P WIN!"
        Exit Function
    End If
    
    '--턴 변경 확인(충돌지점 확인, 설정)------------
    If Selection < 10 Then ' 충돌 없는 지점이었을 때
        Selection = Selection + 10 '충돌 지점으로 설정
        
        'Cells(1,4) = (Cells(1,4)+1) mod 2 ' 턴변경
        If Cells(1, 4) = 0 Then ' 턴변경
            Cells(1, 4) = 1
            Cells(4, 19) = "2P 차례"
            Cells(4, 19).Interior.Color = clr_2p
        Else
            Cells(1, 4) = 0
            Cells(4, 19) = "1P 차례"
            Cells(4, 19).Interior.Color = clr_1p
        End If
        '턴이 바뀔 때 시점 저장
        Range(Cells(1, 1), Cells(21, 20)).Copy Range(Cells(1, 21), Cells(21, 40))
    End If
     
    Where '이동할 수 있는 곳 표시
End Function

Function Re(x, y) '점 제거용
    clr_line = RGB(0, 0, 0)
    clr_next = RGB(101, 255, 101)
    If x < 1 Or y < 1 Then '시트 밖일 때
        Exit Function
    ElseIf Cells(x, y).Interior.Color = clr_next Then
        Cells(x, y).Interior.Color = clr_line
    End If
End Function

Function Where() '이동할 수 있는 곳 표시
    clr_now = RGB(255, 255, 0)
    clr_next = RGB(101, 255, 101)
    clr_1p = RGB(255, 203, 203)
    clr_2p = RGB(203, 203, 255)
    
    a = Cells(1, 2)
    b = Cells(1, 3)
    
    k = Pos(a, b, a - 2, b - 2) + Pos(a, b, a, b - 2) + Pos(a, b, a + 2, b - 2) + Pos(a, b, a - 2, b) + Pos(a, b, a + 2, b) + Pos(a, b, a - 2, b + 2) + Pos(a, b, a, b + 2) + Pos(a, b, a + 2, b + 2)
    
    Cells(a, b).Interior.Color = clr_now
    
    If k = 0 Then '이동할 곳이 없음
        If Cells(1, 4) = 1 Then '현재 2P의 차례
            Cells(4, 19) = "1P WIN!"
            Cells(4, 19).Interior.Color = clr_1p
            Exit Function
        Else '현재 1P의 차례
            Cells(4, 19) = "2P WIN!"
            Cells(4, 19).Interior.Color = clr_2p
            Exit Function
        End If
    End If
End Function
Function Pos(a, b, x, y) '해당 지점이 가능한지 확인 가능하면 1을 반환
    clr_next = RGB(101, 255, 101)
    
    k = (a - x) * (b - y)
    
    If x < 1 Or y < 1 Then '시트 밖일 때
        Exit Function
    ElseIf Cells(x, y).Interior.Color <> RGB(0, 0, 0) Then '게임판 위가 아닐 때
        Exit Function
    ElseIf k = 0 And Cells(a, b) * Cells(x, y) Mod 2 = 1 Then '가로 세로 // 벽타기일 때
        Exit Function
    ElseIf k = 0 And Cells((a + x) / 2, (b + y) / 2) = 1 Then '가로 세로 // 이미 선이 그어진 방향일 때
        Exit Function
'    ElseIf k <> 0 And Cells(a, b) * Cells(x, y) Mod 2 = 1 Then '대각선 // 벽타기
'        Exit Function
    ElseIf k > 0 And Cells((a + x) / 2, (b + y) / 2) / 3 <> Int(Cells((a + x) / 2, (b + y) / 2) / 3) Then '대각선 아래 // 이미
        Exit Function
    ElseIf k < 0 And Cells((a + x) / 2, (b + y) / 2) / 2 <> Int(Cells((a + x) / 2, (b + y) / 2) / 2) Then '대각선 위 // 이미
        Exit Function
    Else
        Cells(x, y).Interior.Color = clr_next
    End If
    
    Pos = 1
End Function

Sub Start() '시작
    '게임판 적용'
    clr_now = RGB(255, 255, 0)
    clr_next = RGB(101, 255, 101)
    clr_1p = RGB(255, 203, 203)
    clr_2p = RGB(203, 203, 255)
    clr_line = RGB(0, 0, 0)
        
    'Application.ScreenUpdating = False
    
    Range("A1:XFD1048576").EntireRow.Clear
    Range("A1:XFD1048576").EntireColumn.Clear
    Range("U22:XFD1048576").EntireRow.Hidden = True
    Range("U22:XFD1048576").EntireColumn.Hidden = True
    
    Range(Cells(1, 1), Cells(21, 20)).Interior.Color = RGB(255, 255, 255)
    For i = 0 To 9 '게임판 비율조정
        Rows(2 * i + 1).RowHeight = 7
        Rows(2 * i + 2).RowHeight = 50
        Columns(2 * i + 1).ColumnWidth = 0.9
        Columns(2 * i + 2).ColumnWidth = 8.11
    Next
    Rows(21).RowHeight = 7
    Columns(19).ColumnWidth = 8.11
    
    With Cells(2, 19) '시작버튼
        .Value = "시작" & vbCrLf & "[클릭]"
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(229, 229, 229)
    End With
    With Cells(2, 20) '무르기 버튼
        .Value = "무르기" & vbCrLf & "[클릭]"
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(229, 229, 229)
    End With
    With Cells(4, 19) '차례 표시
        .HorizontalAlignment = xlCenter
        .Value = "1P 차례"
        .Interior.Color = clr_1p
    End With
    With Cells(4, 20) '메세지 위치
        .HorizontalAlignment = xlCenter
    End With
    Range(Cells(1, 1), Cells(1, 3)).Font.Color = RGB(255, 255, 255)
    
    '테두리구분을 위한 1 + 충돌지점을 위한 10 = 11을 넣어줘야함
    '일단 판 전체에 적용한 뒤 내부는 되돌리는 식으로 처리
    With Range("A3:Q19,G1:K21")
        .Interior.Color = clr_line
        .Value = 11
    End With
            
    Range("B4:P18, H2:J20").Value = "" '내부는 다시 0으로
    
    Range("H2,J2").Interior.Color = clr_1p 'lP 자리
    Range("H20,J20").Interior.Color = clr_2p '2P 자리
    
    For i = 1 To 8
        For j = 1 To 8
            Cells(2 + 2 * i, 2 * j).Interior.Color = RGB(255, 255, 255)
        Next
    Next
    
    Cells(11, 9).Interior.Color = clr_now '(9,9) 시작위치
    Cells(11, 9) = 10 '시작위치는 자동으로 충돌지점
    
    Cells(1, 2) = 11 '현재위치 저장
    Cells(1, 3) = 9
    
    Range("G9,I9,K9,G11,K11,G13,I13,K13").Interior.Color = clr_next '이동할 수 있는 곳 표시
    
    Application.ScreenUpdating = True
End Sub
```

<details>
