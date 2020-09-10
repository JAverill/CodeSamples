' Simple Calculator code sample
' Author: John Averill
' Version 1.0

' Form Class
Public Class Form1
    ' Form Load
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        sLeft = "" ' Declare each variable as blank
        sRight = "" ' Declare each variable as blank
        sOperator = "" ' Declare each variable as blank
        ResultsBox.Text = "0" ' Declare each variable as blank
        bLeft = True ' Set to left of equation
    End Sub

    ' Declare Global Variables
    Dim sLeft As String, sRight As String, sOperator As String ' Left and Right side of equation Strings, and Operator string
    Dim iLeft As Double, iRight As Double, iResult As Double ' Doubles for the value of the Left and Right side, and for the Result
    Dim bLeft As Boolean ' Boolean to determine if user is entering numbers for the Left or Right side of equation, will be set to true at beginning, and changed to false by the input of an Operator.

    ' Subroutine to handle the input of numbers on the calculator.

    Private Sub InputHandler(sNumber As String)
        If bLeft Then ' If left side of equation
            sLeft = sLeft + sNumber ' Set Left value
            ResultsBox.Text = sLeft ' Display Left value
            Refresh() ' Refresh Form
        Else ' If right side of equation
            sRight = sRight + sNumber ' Set Right value
            ResultsBox.Text = sRight ' Display Right value
            Refresh() ' Refresh Form
        End If
    End Sub

    ' Subroutine to handle the mathematical operator for an equation.
    Private Sub OperatorHandler(sNewOperator As String)
        If bLeft Then ' If left side of equation
            sOperator = sNewOperator ' Set Operator string
            bLeft = False ' Set to right side of equation
        Else
            If sLeft <> "" And sRight = "" And sOperator <> "" Then ' If left is not blank, but right is, and an operator has been entered
                sRight = sLeft ' Set Right to be equal to Left, so operations like 2* will function.
            End If
            If sLeft <> "" And sRight <> "" And sOperator <> "" Then ' If no variable is blank
                iLeft = sLeft ' Set Left double to Left String
                iRight = sRight ' Set Right double to Left String
                Select Case sOperator ' Case switch to handle actual math operations.
                    Case "+" ' Addition switch
                        iResult = iLeft + iRight
                    Case "-" ' Subtraction switch
                        iResult = iLeft - iRight
                    Case "/" ' Division switch
                        iResult = iLeft / iRight
                    Case "*" ' Multiplication switch
                        iResult = iLeft * iRight
                End Select
                ResultsBox.Text = iResult ' Display results
                Refresh() ' Refresh Form
                sLeft = iResult ' Set current result as left variable for continued operations on result
                sRight = "" ' Set Right to blank for same reason
                bLeft = True ' Return to Left side of equation to allow new operators
            End If
            sOperator = sNewOperator ' Handle new operator
            sRight = "" ' Set Right to blank
            bLeft = False ' Set to right side of equation
        End If
    End Sub

    ' Create Handlers for number buttons
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btn1.Click
        InputHandler("1") ' Call Input Handler
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles btn2.Click
        InputHandler("2") ' Call Input Handler
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles btn3.Click
        InputHandler("3") ' Call Input Handler
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles btn4.Click
        InputHandler("4") ' Call Input Handler
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles btn5.Click
        InputHandler("5") ' Call Input Handler
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles btn6.Click
        InputHandler("6") ' Call Input Handler
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles btn7.Click
        InputHandler("7") ' Call Input Handler
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles btn8.Click
        InputHandler("8") ' Call Input Handler
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles btn9.Click
        InputHandler("9") ' Call Input Handler
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles btn0.Click
        InputHandler("0") ' Call Input Handler
    End Sub

    ' Create handlers for operator buttons

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles btnDivide.Click
        OperatorHandler("/") ' Call Operator Handler
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles btnMultiply.Click
        OperatorHandler("*") ' Call Operator Handler
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles btnSubtract.Click
        OperatorHandler("-") ' Call Operator Handler
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        OperatorHandler("+") ' Call Operator Handler
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles btnDecimal.Click
        InputHandler(".") ' Call Operator Handler
    End Sub

    ' Enter button handler

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnEnter.Click
        ' Allow operator to apply to initial integer if no second integer applied
        If sLeft <> "" And sRight = "" And sOperator <> "" Then
            sRight = sLeft
        End If
        ' Perform operation if all variables are not empty
        If sLeft <> "" And sRight <> "" And sOperator <> "" Then
            iLeft = sLeft ' Set doubles from strings
            iRight = sRight ' Set doubles from strings
            Select Case sOperator ' Operator switch
                Case "+" ' Addition
                    iResult = iLeft + iRight
                Case "-" ' Subtraction
                    iResult = iLeft - iRight
                Case "/" ' Division
                    iResult = iLeft / iRight
                Case "*" ' Multiplication
                    iResult = iLeft * iRight
            End Select
            ResultsBox.Text = iResult ' Display result
            Refresh() ' Refresh Form
            sLeft = iResult ' Set result to left variable to allow additional operations
            sRight = "" ' Set right variable to blank
            bLeft = True ' Set left to true to allow input of operator
        End If
    End Sub

    ' Percentage button handler.

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles btnPercent.Click
        If Not bLeft Then ' Set to only function on second operator for obvious math purposes
            If sRight <> "" Then
                iRight = sRight ' Get right variable
            Else
                iRight = 0 ' If right variable is blank set to 0
            End If
            iRight = iRight * (iLeft / 100) ' Convert right variable to percentage
            ResultsBox.Text = iRight ' Display percentage
            Refresh() ' Refresh Form
            If iRight <> 0 Then
                sRight = iRight ' Set variable to percentage if it is not 0
            Else
                sRight = "" ' Set variable to blank if it is 0
            End If
        End If
    End Sub

    ' Handler for clearing the entire equation

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        sLeft = "" ' Declare each variable as blank
        sRight = "" ' Declare each variable as blank
        sOperator = "" ' Declare each variable as blank
        ResultsBox.Text = "0" ' Declare each variable as blank
        bLeft = True ' Set to left of equation
        Refresh() ' Refresh Form
    End Sub

    ' Handler for clearing only the current entry

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles btnCE.Click
        If bLeft Then ' If left variable
            sLeft = "" ' Set left to blank
            ResultsBox.Text = "0"
            Refresh() ' Refresh Form
        Else
            sRight = "" ' If right variable set right to blank
            ResultsBox.Text = "0"
            Refresh() ' Refresh Form
        End If
    End Sub
End Class
