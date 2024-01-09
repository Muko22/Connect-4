Public Class Form1
    Dim pics(5, 6) As PictureBox
    Dim currentPlayer As Integer = 0
    Dim playerTwo As String = "Player 2"
    Dim isComputer As Boolean = False

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        'This resets all game data
        resetGameData() 'clear picture box images and tags
        lblScore1.Text = 0
        lblScore2.Text = 0
        'Register new player two
        If MsgBox("Do you want to play against computer?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            playerTwo = "Computer"
            isComputer = True
            lblP2.Text = "Computer" : lblP1.Text = "You"
        Else
            isComputer = False
            lblP2.Text = "Player 2" : lblP1.Text = "Player 1"
        End If
    End Sub

    Private Sub resetGameData()
        'clear picture box images and tags
        For Each item In GroupBox1.Controls
            If TypeOf (item) Is PictureBox Then
                item.imagelocation = "" 'My.Resources.Resources.if_accepted_48_10249
                item.image = Nothing
                item.tag = ""
            End If
        Next

        'Prevent the last player from starting the new game.
        If currentPlayer <> 1 Then
            currentPlayer = 1
            picPlayer2.Enabled = False
            picPlayer1.Enabled = True
            picPlayer1.Focus()
        Else
            currentPlayer = 2
            picPlayer2.Enabled = True
            picPlayer1.Enabled = False
            picPlayer2.Focus()
            If isComputer Then
                comPlay()
            End If
        End If
        lblCurrentPlayer.Text = "Player " & currentPlayer
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Assign values to each of the pics array slot
        'row 1
        pics(0, 0) = PictureBox1    'Col 1
        pics(0, 1) = PictureBox8    'Col 2
        pics(0, 2) = PictureBox9    'Col 3
        pics(0, 3) = PictureBox10    'Col 4
        pics(0, 4) = PictureBox11    'Col 5
        pics(0, 5) = PictureBox12    'Col 6
        pics(0, 6) = PictureBox13    'Col 7
        'row 2
        pics(1, 0) = PictureBox3    'Col 1
        pics(1, 1) = PictureBox16   'Col 2
        pics(1, 2) = PictureBox18   'Col 3
        pics(1, 3) = PictureBox14    'Col 4
        pics(1, 4) = PictureBox17    'Col 5
        pics(1, 5) = PictureBox19    'Col 6
        pics(1, 6) = PictureBox15    'Col 7
        'row 3
        pics(2, 0) = PictureBox4
        pics(2, 1) = PictureBox22
        pics(2, 2) = PictureBox24
        pics(2, 3) = PictureBox20
        pics(2, 4) = PictureBox23
        pics(2, 5) = PictureBox25
        pics(2, 6) = PictureBox21
        'row 4
        pics(3, 0) = PictureBox5
        pics(3, 1) = PictureBox28
        pics(3, 2) = PictureBox30
        pics(3, 3) = PictureBox26
        pics(3, 4) = PictureBox29
        pics(3, 5) = PictureBox31
        pics(3, 6) = PictureBox27
        'row 5
        pics(4, 0) = PictureBox6
        pics(4, 1) = PictureBox34
        pics(4, 2) = PictureBox36
        pics(4, 3) = PictureBox32
        pics(4, 4) = PictureBox35
        pics(4, 5) = PictureBox37
        pics(4, 6) = PictureBox33
        'row 6
        pics(5, 0) = PictureBox7
        pics(5, 1) = PictureBox40
        pics(5, 2) = PictureBox42
        pics(5, 3) = PictureBox38
        pics(5, 4) = PictureBox41
        pics(5, 5) = PictureBox43
        pics(5, 6) = PictureBox39


        'Load previous game data from the text file.
        Dim fileContents As String
        fileContents = My.Computer.FileSystem.ReadAllText(Application.StartupPath & "\DB.txt")
        Dim str = fileContents.Split(",")
        If str.Count > 0 Then
            If str(0) <> "" Then
                resetGameData()
                'Load player 1 data
                lblP1.Text = str(0).Split(":")(0)
                lblScore1.Text = str(0).Split(":")(1)
                'Load player 2 data
                lblP2.Text = str(1).Split(":")(0)
                lblScore2.Text = str(1).Split(":")(1)
            Else
                btnReset.PerformClick()
            End If
        End If
    End Sub

    Private Sub ColSelected(sender As Object, e As EventArgs) Handles Button8.Click, Button7.Click, Button6.Click, Button5.Click, Button4.Click, Button3.Click, Button1.Click
        If currentPlayer = 0 Then
            MsgBox("Restart the Game and try again.") : Exit Sub
        End If
        'Get the column button that was clicked.
        Dim btn As Button = sender
        playGame(btn.Text)
    End Sub

    Private Sub playGame(index As Integer)
        Dim isColumnAvailable As Boolean = False
        'Loop through each row of this column and drop the ball(image) of the current player.
        For i = 0 To 5  'Loop through the rows
            If pics(i, index - 1).Image Is Nothing Then
                isColumnAvailable = True
                pics(i, index - 1).Tag = currentPlayer
                If currentPlayer = 2 Then
                    pics(i, index - 1).Image = My.Resources.Cancel_02 : Exit For
                Else
                    pics(i, index - 1).Image = My.Resources.if_accepted_48_10249 : Exit For
                End If
            End If
        Next

        If Not isColumnAvailable Then
            MsgBox("This row is filled!") : Exit Sub
        End If

        'Check if this player won
        If CheckWin() Then
            If currentPlayer = 2 Then
                lblScore2.Text = CInt(lblScore2.Text) + 1
            Else
                lblScore1.Text = CInt(lblScore1.Text) + 1
            End If
            saveScores()
            MsgBox("Player " & currentPlayer & " Won!")
            resetGameData()
            Exit Sub
        End If

        'Prevent the previous player from playing again.
        If currentPlayer = 2 Then
            currentPlayer = 1
            picPlayer2.Enabled = False
            picPlayer1.Enabled = True
            picPlayer1.Focus()
        Else
            currentPlayer = 2
            picPlayer2.Enabled = True
            picPlayer1.Enabled = False
            picPlayer2.Focus()
            If isComputer Then
                comPlay()   'Play for computer
            End If
        End If
        lblCurrentPlayer.Text = "Player " & currentPlayer
    End Sub

    Private Sub comPlay()
        Try
            'Play and check if the computer won.
            If comCheck() Then
                If currentPlayer = 2 Then
                    lblScore2.Text = CInt(lblScore2.Text) + 1
                Else
                    lblScore1.Text = CInt(lblScore1.Text) + 1
                End If
                saveScores()
                'Delay Processing
                System.Threading.Thread.Sleep(1000)
                MsgBox("Player " & currentPlayer & " Won!")
                resetGameData()
                Exit Sub
            Else
                currentPlayer = 1
                picPlayer1.Enabled = True
                picPlayer2.Enabled = False
                picPlayer1.Focus()
                lblCurrentPlayer.Text = "Player " & currentPlayer

            End If
        Catch ex As Exception
            If ex.Message = "Game Over" Then
                resetGameData()
            Else
                Throw New Exception(ex.Message)
            End If
        End Try
    End Sub

    Private Sub saveScores()
        'Save current user data
        My.Computer.FileSystem.WriteAllText(Application.StartupPath & "\DB.txt", lblP1.Text &
                                            ":" & lblScore1.Text & "," & lblP2.Text &
                                            ":" & lblScore2.Text, False)
    End Sub

    ''' <summary>
    ''' This method considers the best possible position to play
    ''' and also checks if computer won the game
    ''' </summary>
    Private Function comCheck()

        'Check computer's possible win positions
        For i = 0 To 5  'Loop through the rows

            For j = 0 To 6  'Loop through the columns
                'vertical Check
                If j <= 3 Then
                    If (pics(i, j).Tag.Equals(currentPlayer) And pics(i, j + 1).Tag.Equals(currentPlayer) And
                    pics(i, j + 2).Tag.Equals(currentPlayer) And pics(i, j + 3).Tag.Equals("")) Then
                        Return comHit(i, j + 3) 'Call a new method that sets the image to the picture box

                    End If
                End If
                'Horizontal check
                If i <= 2 Then
                    If pics(i, j).Tag.Equals(currentPlayer) And pics(i + 1, j).Tag.Equals(currentPlayer) And
                    pics(i + 2, j).Tag.Equals(currentPlayer) And pics(i + 3, j).Tag.Equals("") Then
                        Return comHit(i + 3, j) 'Call a new method that sets the image to the picture box
                    End If
                End If

            Next

            'Diagonal check
            If i <= 2 Then
                'left to right check
                If (pics(i, 0).Tag.Equals(currentPlayer) And pics(i + 1, 1).Tag.Equals(currentPlayer) And
                        pics(i + 2, 2).Tag.Equals(currentPlayer) And pics(i + 3, 3).Tag.Equals("")) Then
                    'Won

                    Return comHit(i + 3, 3)
                    'left to right check
                ElseIf (pics(i, 1).Tag.Equals(currentPlayer) And pics(i + 1, 2).Tag.Equals(currentPlayer) And
                        pics(i + 2, 3).Tag.Equals(currentPlayer) And pics(i + 3, 4).Tag.Equals("")) Then
                    'Won

                    Return comHit(i + 3, 4)
                    'left to right check
                ElseIf (pics(i, 2).Tag.Equals(currentPlayer) And pics(i + 1, 3).Tag.Equals(currentPlayer) And
                        pics(i + 2, 4).Tag.Equals(currentPlayer) And pics(i + 3, 5).Tag.Equals("")) Then
                    'Won

                    Return comHit(i + 3, 5)
                    'Right to left check
                ElseIf (pics(i, 6).Tag.Equals(currentPlayer) And pics(i + 1, 5).Tag.Equals(currentPlayer) And
                        pics(i + 2, 4).Tag.Equals(currentPlayer) And pics(i + 3, 3).Tag.Equals("")) Then
                    'Won
                    Return comHit(i + 3, 3)
                    'Right to left check
                ElseIf (pics(i, 5).Tag.Equals(currentPlayer) And pics(i + 1, 4).Tag.Equals(currentPlayer) And
                        pics(i + 2, 3).Tag.Equals(currentPlayer) And pics(i + 3, 2).Tag.Equals("")) Then
                    'Won

                    Return comHit(i + 3, 2)
                    'Right to left check
                ElseIf (pics(i, 4).Tag.Equals(currentPlayer) And pics(i + 1, 3).Tag.Equals(currentPlayer) And
                        pics(i + 2, 2).Tag.Equals(currentPlayer) And pics(i + 3, 1).Tag.Equals("")) Then
                    'Won

                    Return comHit(i + 3, 1)
                    'Center Check
                ElseIf (pics(i, 3).Tag.Equals(currentPlayer) And pics(i + 1, 4).Tag.Equals(currentPlayer) And
                        pics(i + 2, 5).Tag.Equals(currentPlayer) And pics(i + 3, 6).Tag.Equals("")) Then
                    'Won

                    Return comHit(i + 3, 6)
                ElseIf (pics(i, 3).Tag.Equals(currentPlayer) And pics(i + 1, 2).Tag.Equals(currentPlayer) And
                        pics(i + 2, 1).Tag.Equals(currentPlayer) And pics(i + 3, 0).Tag.Equals("")) Then
                    'Won

                    Return comHit(i + 3, 0)
                End If

            End If
        Next

        ' This method starts by checking user's possible win positions and blocking them(it),
        For i = 0 To 5  'Loop through the rows

            For j = 0 To 6  'Loop through the columns
                'vertical Check
                If j <= 3 Then
                    If (pics(i, j).Tag.Equals(1) And pics(i, j + 1).Tag.Equals(1) And
                    pics(i, j + 2).Tag.Equals(1) And pics(i, j + 3).Tag.Equals("")) Then
                        Return comHit(i, j + 3)
                    End If
                Else    'Check from right to left
                    If (pics(i, j).Tag.Equals(1) And pics(i, j - 1).Tag.Equals(1) And
                    pics(i, j - 2).Tag.Equals(1) And pics(i, j - 3).Tag.Equals("")) Then
                        Return comHit(i, j - 3)
                    End If
                End If
                'Horizontal check
                If i <= 2 Then
                    If pics(i, j).Tag.Equals(1) And pics(i + 1, j).Tag.Equals(1) And
                    pics(i + 2, j).Tag.Equals(1) And pics(i + 3, j).Tag.Equals("") Then
                        Return comHit(i + 3, j)
                    End If
                End If

            Next

            'Diagonal check
            If i <= 2 Then
                'left to right check
                If (pics(i, 0).Tag.Equals(1) And pics(i + 1, 1).Tag.Equals(1) And
                        pics(i + 2, 2).Tag.Equals(1) And pics(i + 3, 3).Tag.Equals("")) Then

                    Return comHit(i + 3, 3)
                    'left to right check
                ElseIf (pics(i, 1).Tag.Equals(1) And pics(i + 1, 2).Tag.Equals(1) And
                        pics(i + 2, 3).Tag.Equals(1) And pics(i + 3, 4).Tag.Equals("")) Then

                    Return comHit(i + 3, 4)
                    'left to right check
                ElseIf (pics(i, 2).Tag.Equals(1) And pics(i + 1, 3).Tag.Equals(1) And
                        pics(i + 2, 4).Tag.Equals(1) And pics(i + 3, 5).Tag.Equals("")) Then

                    Return comHit(i + 3, 5)
                    'Right to left check
                ElseIf (pics(i, 6).Tag.Equals(1) And pics(i + 1, 5).Tag.Equals(1) And
                        pics(i + 2, 4).Tag.Equals(1) And pics(i + 3, 3).Tag.Equals("")) Then

                    Return comHit(i + 3, 3)
                    'Right to left check
                ElseIf (pics(i, 5).Tag.Equals(1) And pics(i + 1, 4).Tag.Equals(1) And
                        pics(i + 2, 3).Tag.Equals(1) And pics(i + 3, 2).Tag.Equals("")) Then

                    Return comHit(i + 3, 2)
                    'Right to left check
                ElseIf (pics(i, 4).Tag.Equals(1) And pics(i + 1, 3).Tag.Equals(1) And
                        pics(i + 2, 2).Tag.Equals(1) And pics(i + 3, 1).Tag.Equals("")) Then

                    Return comHit(i + 3, 1)
                    'Center Check
                ElseIf (pics(i, 3).Tag.Equals(1) And pics(i + 1, 4).Tag.Equals(1) And
                        pics(i + 2, 5).Tag.Equals(1) And pics(i + 3, 6).Tag.Equals("")) Then

                    Return comHit(i + 3, 6)
                ElseIf (pics(i, 3).Tag.Equals(1) And pics(i + 1, 2).Tag.Equals(1) And
                        pics(i + 2, 1).Tag.Equals(1) And pics(i + 3, 0).Tag.Equals("")) Then

                    Return comHit(i + 3, 0) : Return False
                End If

            End If
        Next

        'Hit nearest Button to coms Positions

        'Check for double balls
        Dim k As Integer = 0
        While k <= 5

            For j = 0 To 6
                'vertical Check
                If j <= 3 Then
                    If (pics(k, j).Tag.Equals(currentPlayer) And pics(k, j + 1).Tag.Equals(currentPlayer) And
                    pics(k, j + 2).Tag.Equals("")) Then
                        Return comHit(k, j + 2)
                    End If
                End If
                'Horizontal check
                If k <= 2 Then
                    If pics(k, j).Tag.Equals(currentPlayer) And pics(k + 1, j).Tag.Equals(currentPlayer) And
                    pics(k + 2, j).Tag.Equals("") Then
                        Return comHit(k + 2, j)

                    End If
                End If

            Next

            'Diagonal check
            If k <= 2 Then
                'left to right check
                If (pics(k, 0).Tag.Equals(currentPlayer) And pics(k + 1, 1).Tag.Equals(currentPlayer) And
                        pics(k + 2, 2).Tag.Equals("")) Then
                    Return comHit(k + 2, 2)
                ElseIf (pics(k, 1).Tag.Equals(currentPlayer) And pics(k + 1, 2).Tag.Equals(currentPlayer) And
                        pics(k + 2, 3).Tag.Equals("")) Then
                    Return comHit(k + 2, 3)
                ElseIf (pics(k, 2).Tag.Equals(currentPlayer) And pics(k + 1, 3).Tag.Equals(currentPlayer) And
                        pics(k + 2, 4).Tag.Equals("")) Then
                    Return comHit(k + 2, 4)

                    'Right to left check
                ElseIf (pics(k, 6).Tag.Equals(currentPlayer) And pics(k + 1, 5).Tag.Equals(currentPlayer) And
                        pics(k + 2, 4).Tag.Equals("")) Then
                    Return comHit(k + 2, 4)

                ElseIf (pics(k, 5).Tag.Equals(currentPlayer) And pics(k + 1, 4).Tag.Equals(currentPlayer) And
                      pics(k + 2, 3).Tag.Equals("")) Then
                    Return comHit(k + 2, 3)

                ElseIf (pics(k, 4).Tag.Equals(currentPlayer) And pics(k + 1, 3).Tag.Equals(currentPlayer) And
                        pics(k + 2, 2).Tag.Equals("")) Then
                    Return comHit(k + 2, 2)

                    'Center Check
                ElseIf (pics(k, 3).Tag.Equals(currentPlayer) And pics(k + 1, 4).Tag.Equals(currentPlayer) And
                        pics(k + 2, 5).Tag.Equals("")) Then
                    Return comHit(k + 2, 5)

                ElseIf (pics(k, 3).Tag.Equals(currentPlayer) And pics(k + 1, 2).Tag.Equals(currentPlayer) And
                    pics(k + 2, 1).Tag.Equals("")) Then
                    Return comHit(k + 2, 1)
                End If

            End If
            k += 1
        End While

        'Check for single balls
        k = 0
        While k <= 5

            For j = 0 To 6
                'vertical Check
                If j <= 3 Then
                    If (pics(k, j).Tag.Equals(currentPlayer) And pics(k, j + 1).Tag.Equals("")) Then
                        Return comHit(k, j + 1)
                    End If
                End If
                'Horizontal check
                If k <= 2 Then
                    If pics(k, j).Tag.Equals(currentPlayer) And pics(k + 1, j).Tag.Equals("") Then
                        Return comHit(k + 1, j)

                    End If
                End If

            Next

            'Diagonal check
            If k <= 2 Then
                'left to right check
                If (pics(k, 0).Tag.Equals(currentPlayer) And pics(k + 1, 1).Tag.Equals("")) Then
                    Return comHit(k + 1, 1)
                    'shouldPlay = True : pos = (k + 2) & "," & 2 : Exit While
                ElseIf (pics(k, 1).Tag.Equals(currentPlayer) And pics(k + 1, 2).Tag.Equals("")) Then
                    Return comHit(k + 1, 2)
                    'shouldPlay = True : pos = (k + 2) & "," & 3 : Exit While
                ElseIf (pics(k, 2).Tag.Equals(currentPlayer) And pics(k + 1, 3).Tag.Equals("")) Then
                    Return comHit(k + 1, 3)

                    'Right to left check
                ElseIf (pics(k, 6).Tag.Equals(currentPlayer) And pics(k + 1, 5).Tag.Equals("")) Then
                    Return comHit(k + 1, 5)

                ElseIf (pics(k, 5).Tag.Equals(currentPlayer) And pics(k + 1, 4).Tag.Equals("")) Then
                    Return comHit(k + 1, 4)

                ElseIf (pics(k, 4).Tag.Equals(currentPlayer) And pics(k + 1, 3).Tag.Equals("")) Then
                    Return comHit(k + 1, 3)

                    'Center Check
                ElseIf (pics(k, 3).Tag.Equals(currentPlayer) And pics(k + 1, 4).Tag.Equals("")) Then
                    Return comHit(k + 1, 4)

                ElseIf (pics(k, 3).Tag.Equals(currentPlayer) And pics(k + 1, 2).Tag.Equals("")) Then
                    Return comHit(k + 1, 2)

                End If

            End If
            k += 1
        End While

        'At this point, the only option is to hit a single ball in the nearest position
        For i = 0 To 5
            For j = 0 To 6
                If pics(i, j).Tag.Equals("") Then
                    Return comHit(i, j)
                End If
            Next
        Next

        MsgBox("Game over! no winner for this match.")
        Throw New Exception("Game Over")
        'Return False
    End Function
    Private Function comHit(i As Integer, j As Integer)
        'This checks if there's another empty space below where the computer is intending to play
        If i > 0 Then
            Try
                While pics(i - 1, j).Tag.Equals("")
                    i -= 1
                End While
            Catch ex As Exception   'May run out of range

            End Try
        End If
        pics(i, j).Image = My.Resources.Cancel_02
        pics(i, j).Tag = 2
        Return CheckWin()
    End Function

    Private Function CheckWin() As Boolean
        'Follow certain pattern to check if there's any position with 
        '4 balls in a row eigther horizontally, vertically or diagonally
        For i = 0 To 5  'Loop through the rows
            For j = 0 To 6  'Loop through the columns

                'vertical Check
                If j <= 3 Then
                    If (pics(i, j).Tag.Equals(currentPlayer) And pics(i, j + 1).Tag.Equals(currentPlayer) And
                    pics(i, j + 2).Tag.Equals(currentPlayer) And pics(i, j + 3).Tag.Equals(currentPlayer)) Then
                        Return True
                    End If
                End If
                'Horizontal check
                If i <= 2 Then
                    If pics(i, j).Tag.Equals(currentPlayer) And pics(i + 1, j).Tag.Equals(currentPlayer) And
                    pics(i + 2, j).Tag.Equals(currentPlayer) And pics(i + 3, j).Tag.Equals(currentPlayer) Then
                        Return True
                    End If
                End If

            Next

            'Diagonal check
            If i <= 2 Then
                'left to right check
                If (pics(i, 0).Tag.Equals(currentPlayer) And pics(i + 1, 1).Tag.Equals(currentPlayer) And
                        pics(i + 2, 2).Tag.Equals(currentPlayer) And pics(i + 3, 3).Tag.Equals(currentPlayer)) Or
                        (pics(i, 1).Tag.Equals(currentPlayer) And pics(i + 1, 2).Tag.Equals(currentPlayer) And
                        pics(i + 2, 3).Tag.Equals(currentPlayer) And pics(i + 3, 4).Tag.Equals(currentPlayer)) Or
                        (pics(i, 2).Tag.Equals(currentPlayer) And pics(i + 1, 3).Tag.Equals(currentPlayer) And
                        pics(i + 2, 4).Tag.Equals(currentPlayer) And pics(i + 3, 5).Tag.Equals(currentPlayer)) Then
                    'Won

                    Return True
                    'Right to left check
                ElseIf (pics(i, 6).Tag.Equals(currentPlayer) And pics(i + 1, 5).Tag.Equals(currentPlayer) And
                        pics(i + 2, 4).Tag.Equals(currentPlayer) And pics(i + 3, 3).Tag.Equals(currentPlayer)) Or
                        (pics(i, 5).Tag.Equals(currentPlayer) And pics(i + 1, 4).Tag.Equals(currentPlayer) And
                        pics(i + 2, 3).Tag.Equals(currentPlayer) And pics(i + 3, 2).Tag.Equals(currentPlayer)) Or
                        (pics(i, 4).Tag.Equals(currentPlayer) And pics(i + 1, 3).Tag.Equals(currentPlayer) And
                        pics(i + 2, 2).Tag.Equals(currentPlayer) And pics(i + 3, 1).Tag.Equals(currentPlayer)) Then
                    'Won

                    Return True
                    'Center Check
                ElseIf (pics(i, 3).Tag.Equals(currentPlayer) And pics(i + 1, 4).Tag.Equals(currentPlayer) And
                        pics(i + 2, 5).Tag.Equals(currentPlayer) And pics(i + 3, 6).Tag.Equals(currentPlayer)) Or
                        (pics(i, 3).Tag.Equals(currentPlayer) And pics(i + 1, 2).Tag.Equals(currentPlayer) And
                        pics(i + 2, 1).Tag.Equals(currentPlayer) And pics(i + 3, 0).Tag.Equals(currentPlayer)) Then
                    'Won

                    Return True
                End If

            End If
        Next

        Return False
    End Function
End Class
