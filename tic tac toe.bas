Attribute VB_Name = "Module1"
Option Explicit
'array which holds the info about each box
Private box(1 To 9) As block
'active player
'at determines if player is playing as X or 0
Private active%, at%(1 To 2)
'ai checks to see if the player is a computer player
Private ai(1 To 2) As Boolean
Private aif(1 To 2) As Boolean
'the user defiend type for each box to see if its
'x, 0, or empty
Type block
    state As Integer
    '0 - empty
    '1 - x
    '2 - 0
End Type
Function rnum%(upperbound, Optional lowerbound = 1)
'generate a random number
Randomize
1 rnum = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
If rnum > upperbound Or rnum < lowerbound Then GoTo 1
End Function
Function CheckEmpty(BoxIndex%) As Boolean
'check if specific box is empty
If box(BoxIndex).state = 0 Then
    CheckEmpty = True
Else
    CheckEmpty = False
End If
End Function
Sub CheckWin()
'there are 8 possible ways to win
'1 2 3
'4 5 6
'7 8 9
'if you picture box like that then the winning combinations are
'1,2,3  4,5,6   7,8,9   1,4,7   2,5,8
'3,6,9  1,5,9   3,5,7
'this function checks to see if all 3 of those blocks
'are the same. Checks if 3 are x or 0
'if yes then it calls playerwin function
'if no then it calls nextturn function
If box(1).state = box(2).state And box(2).state = box(3).state And box(1).state <> 0 And box(2).state <> 0 And box(3).state <> 0 Then
    Call PlayerWin
ElseIf box(4).state = box(5).state And box(5).state = box(6).state And box(4).state <> 0 And box(5).state <> 0 And box(6).state <> 0 Then
    Call PlayerWin
ElseIf box(7).state = box(8).state And box(8).state = box(9).state And box(7).state <> 0 And box(8).state <> 0 And box(9).state <> 0 Then
    Call PlayerWin
ElseIf box(1).state = box(5).state And box(5).state = box(9).state And box(1).state <> 0 And box(5).state <> 0 And box(9).state <> 0 Then
    Call PlayerWin
ElseIf box(3).state = box(5).state And box(5).state = box(7).state And box(3).state <> 0 And box(5).state <> 0 And box(7).state <> 0 Then
    Call PlayerWin
ElseIf box(1).state = box(4).state And box(4).state = box(7).state And box(1).state <> 0 And box(4).state <> 0 And box(7).state <> 0 Then
    Call PlayerWin
ElseIf box(2).state = box(5).state And box(5).state = box(8).state And box(2).state <> 0 And box(5).state <> 0 And box(8).state <> 0 Then
    Call PlayerWin
ElseIf box(3).state = box(6).state And box(6).state = box(9).state And box(3).state <> 0 And box(6).state <> 0 And box(9).state <> 0 Then
    Call PlayerWin
Else
    Call NextPlayer
End If
End Sub
Function CheckDanger(piece%) As Integer
'this function works in to ways
'it can be used to check if there is any place
'where you can win or if there is any place where you
'can loose
'for example:
'if you are playing as X (which internaly is 1)
'and you pass it 1 it will check if there is
'any row where there are 2 x's and 1 blank spot
'if you are playing as X and you want to see if there
'is any place where you can loose then you pass it
'2 (which is 0) to see if there is anywhere you
'need to go so that you dont loose
'if returns the winning/loosing block if such exists
'or -1
Dim barray(1 To 3), i%
Dim WinChance(1 To 8) As Boolean
CheckDanger = -1
'assign the 3 blocks that are going to be checked
barray(1) = 1: barray(2) = 2: barray(3) = 3
'check the 3 blocks
WinChance(1) = Check(piece, barray, CheckDanger)
'if winning/loosing block found then return it
'and exit the function
'note: rest of the function works the same way
If WinChance(1) = True Then Exit Function
barray(1) = 4: barray(2) = 5: barray(3) = 6
WinChance(2) = Check(piece, barray, CheckDanger)
If WinChance(2) = True Then Exit Function
barray(1) = 7: barray(2) = 8: barray(3) = 9
WinChance(3) = Check(piece, barray, CheckDanger)
If WinChance(3) = True Then Exit Function
barray(1) = 1: barray(2) = 4: barray(3) = 7
WinChance(4) = Check(piece, barray, CheckDanger)
If WinChance(4) = True Then Exit Function
barray(1) = 2: barray(2) = 5: barray(3) = 8
WinChance(5) = Check(piece, barray, CheckDanger)
If WinChance(5) = True Then Exit Function
barray(1) = 3: barray(2) = 6: barray(3) = 9
WinChance(6) = Check(piece, barray, CheckDanger)
If WinChance(6) = True Then Exit Function
barray(1) = 1: barray(2) = 5: barray(3) = 9
WinChance(7) = Check(piece, barray, CheckDanger)
If WinChance(7) = True Then Exit Function
barray(1) = 3: barray(2) = 5: barray(3) = 7
WinChance(8) = Check(piece, barray, CheckDanger)
End Function
Function GoodMove%(piece%)
'this function is similar in a way to check function
'lets say you are playing as X (internal value of 1)
'it will check if there is any line where there is
'one X and two empty blocks, if thats the case it will
'return that block, if no it will return -1
'note: if there is more than 1 move which is
'considered good it will generate a random move
Dim barray(1 To 3), i%, retval As Boolean
Dim GoodChance() As Integer, x%, temp%
ReDim GoodChance(0)
x = 0
GoodMove = -1
'assign the 3 blocks that are going to be checked
barray(1) = 1: barray(2) = 2: barray(3) = 3
'check the block
GoSub checker
barray(1) = 4: barray(2) = 5: barray(3) = 6
GoSub checker
barray(1) = 7: barray(2) = 8: barray(3) = 9
GoSub checker
barray(1) = 1: barray(2) = 4: barray(3) = 7
GoSub checker
barray(1) = 2: barray(2) = 5: barray(3) = 8
GoSub checker
barray(1) = 3: barray(2) = 6: barray(3) = 9
GoSub checker
barray(1) = 1: barray(2) = 5: barray(3) = 9
GoSub checker
barray(1) = 3: barray(2) = 5: barray(3) = 7
GoSub checker
'pick random block
GoodMove = GoodChance(rnum(UBound(GoodChance), LBound(GoodChance)))
Exit Function

'check subfunction
checker:
'check to see if line has
retval = Check2(piece, barray, temp%)
If retval = True Then
    ReDim Preserve GoodChance(x)
    GoodChance(x) = temp
    x = x + 1
End If
Return

End Function
Function Check(piece%, boxs As Variant, safe%) As Boolean
Dim i%
'array that holds number of empty blocks
'number of X blocks
'and number of 0 blocks
ReDim retval%(0 To 2)
'this loop checks how many empty blocks, x blocks
'and 0 blocks there are in a line
For i = 1 To 3
    retval(box(boxs(i)).state) = retval(box(boxs(i)).state) + 1
Next i
'if a line has 2 x's or 2 0's and an empty block
'then it is considered winning/loosing block in a line
If retval(0) = 1 And retval(piece) = 2 Then
    'check which block in a line is the
    'winning/loosing block and return it
    If box(boxs(1)).state = 0 Then
        safe = boxs(1)
    ElseIf box(boxs(2)).state = 0 Then
        safe = boxs(2)
    Else
        safe = boxs(3)
    End If
    Check = True
Else
    safe = -1
    Check = False
End If
End Function
Function Check2(piece%, boxs As Variant, safe%) As Boolean
Dim i%
ReDim retval%(0 To 2)
'works kinda like check, but its looking for a line
'with 2 empty blocks and 1 block of type passed
For i = 1 To 3
    retval(box(boxs(i)).state) = retval(box(boxs(i)).state) + 1
Next i
If retval(0) = 2 And retval(piece) = 1 Then
    If box(boxs(1)).state = 0 Then
        safe = boxs(1)
    ElseIf box(boxs(2)).state = 0 Then
        safe = boxs(2)
    Else
        safe = boxs(3)
    End If
    Check2 = True
Else
    safe = -1
    Check2 = False
End If
End Function
Sub GameOver()
'restart the game
Dim i%
'player one is the active player
active = 1
'decide who is gonna be X and who is gonna be 0
at(1) = rnum(2)
at(2) = 3 - at(1)
'dont touch these variables
aif(1) = False
aif(2) = False
'depending on the checkboxes decide the game mode
'player vs player
If Form1.Option1.Value = True And Form1.Option3.Value = True Then
    ai(1) = False
    ai(2) = False
'player vs cpu
ElseIf Form1.Option1.Value = True And Form1.Option4.Value = True Then
    ai(1) = False
    ai(2) = True
'cpu vs player
ElseIf Form1.Option2.Value = True And Form1.Option3.Value = True Then
    ai(1) = True
    ai(2) = False
'cpu vs cpu
ElseIf Form1.Option2.Value = True And Form1.Option4.Value = True Then
    ai(1) = True
    ai(2) = True
End If
'this checks if all the blocks have been used up
'meaning nobody won
For i = 1 To 9
    box(i).state = 0
    Form1.Picture1(i).Picture = LoadPicture("")
Next i
'if the next player is the computer then we call
'the function which will generate computers move
'if not then we wait for user to go
If ai(active) = True Then Call AIMove
End Sub
Sub AIMove()
Dim i%, j%
'check to see if there is any block
'where computer can go to win
1 i = CheckDanger(at(active))
If i <= 0 Then
    'if no such block then check to see if there
    'is any block that computer need to take
    'in order not to loose
    i = CheckDanger(at(3 - active))
    If i <= 0 Then
        'if no then try to take the middle block
        If CenterFree = True Then
            If aif(active) = False Then
                Call AIMove2(5)
                Call NextPlayer
                Exit Sub
                aif(active) = True
            End If
        Else
            'first it checks
            'if it will be the first turn ai will do
            'if yes it will check if middle piece is
            'taken
            'if yes then it will take a corner
            'if not
            'check for a good move
            'a good move is when in a line
            'u have 1 piece and 2 blocks
            'are empty
            If aif(active) = False And CenterFree = False Then
                If CornerDanger(at(3 - active)) = True Then
                    i = FreeSide
                    Call AIMove2(i)
                    Call NextPlayer
                    Exit Sub
                End If
                i = FreeCorner(at(active))
                If i > 0 Then
                    Call AIMove2(i)
                    Call NextPlayer
                    Exit Sub
                End If
                aif(active) = True
            End If
            i = GoodMove(at(active))
            If i > 0 Then
                Call AIMove2(i)
                Call NextPlayer
                Exit Sub
            End If
            'if no good moves then try to
            'take a corner
            i = FreeCorner(at(active))
            If i > 0 Then
                Call AIMove2(i)
                Call NextPlayer
                Exit Sub
            End If
        End If
        'if all taken then take any rnadom block
        Do: DoEvents
            i = rnum(9, 1)
            If box(i).state = 0 Then
                Call AIMove2(i)
                Call NextPlayer
                Exit Do
            End If
        Loop
    Else
        'take the block needed not to loose
        Call AIMove2(i)
        Call NextPlayer
        Exit Sub
    End If
Else
    'win the gmae
    Call AIMove2(i)
    Call PlayerWin
End If
End Sub
Sub AIMove2(i%)
'set the picture
Form1.Picture1(i).Picture = Form1.Picture2(at(active)).Picture
'update the array
box(i).state = at(active)
End Sub
Sub NextPlayer()
Dim i%
'check to see if all the blocks have been used up
'if no then check if its players or computers turn
'if players then wait for player to go
'if computers then generate a move
'if all blocks are tkaen show a message
'stating that game ended in a tie
For i = 1 To 9
    If box(i).state = 0 Then
        active = 3 - active
        If ai(active) = True Then Call AIMove
        Exit Sub
    End If
Next i
MsgBox "Its a tie!"
End Sub
Function CenterFree() As Boolean
'check to see if middle block is empty
    If box(5).state = 0 Then
        CenterFree = True
    Else
        CenterFree = False
    End If
End Function
Function FreeCorner(piece%) As Integer
'check all four corners
'and pick random empty corner
Dim corners(), x%
x = 0
ReDim corners(x)
If box(1).state = 0 And (box(9).state = 0 Or box(9).state = piece) Then
    ReDim Preserve corners(x)
    corners(x) = 1
    x = x + 1
End If
If box(3).state = 0 And (box(7).state = 0 Or box(7).state = piece) Then
    ReDim Preserve corners(x)
    corners(x) = 3
    x = x + 1
End If
If box(7).state = 0 And (box(3).state = 0 Or box(3).state = piece) Then
    ReDim Preserve corners(x)
    corners(x) = 7
    x = x + 1
End If
If box(9).state = 0 And (box(1).state = 0 Or box(1).state = piece) Then
    ReDim Preserve corners(x)
    corners(x) = 9
    x = x + 1
End If
FreeCorner = corners(rnum(UBound(corners), LBound(corners)))
End Function
Sub PlayerWin()
Dim temp$
'show a message that player/computer won
If ai(active) = True Then
    temp$ = "Computer "
Else
    temp$ = "Player "
End If
MsgBox temp$ & active & " wins!"
End Sub
Sub PlayerMove(index%)
If CheckEmpty(index) = True Then
    Form1.Picture1(index).Picture = Form1.Picture2(at(active)).Picture
    box(index).state = at(active)
    Call CheckWin
End If
End Sub
Function CornerDanger(piece%) As Boolean
'this checks if a player controls 2 corner blocks
'accross of each other
'1 and 9, 3 and 7
'if its computers second move and the 2 blocks
'are controlled then instead of taking a corner
'computer will attempt to make a winning situtation
'to prevemnt the other player from creating a fork
CornerDanger = False
If box(1).state = piece And box(9).state = piece Then
    CornerDanger = True
ElseIf box(3).state = piece And box(7).state = piece Then
    CornerDanger = True
End If
End Function
Function FreeSide() As Integer
'check all four corners
'and pick random empty corner
Dim corners(), x%
x = 0
ReDim corners(x)
If box(2).state = 0 Then
    ReDim Preserve corners(x)
    corners(x) = 2
    x = x + 1
End If
If box(4).state = 0 Then
    ReDim Preserve corners(x)
    corners(x) = 4
    x = x + 1
End If
If box(6).state = 0 Then
    ReDim Preserve corners(x)
    corners(x) = 6
    x = x + 1
End If
If box(8).state = 0 Then
    ReDim Preserve corners(x)
    corners(x) = 8
    x = x + 1
End If
FreeSide = corners(rnum(UBound(corners), LBound(corners)))
End Function
