' PRELIM '
Public Class Form1
    Dim QZ1_P, QZ2_P, QZ3_P, ATT_P, REC1_P, REC2_P, ACT1_P, ACT2_P, ACT3_P, EXAM_P As Double

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Box33_TextChanged(sender As Object, e As EventArgs) Handles Box33.TextChanged

    End Sub

    Private Sub Label33_Click(sender As Object, e As EventArgs) Handles Label33.Click

    End Sub

    Dim TQ1_P, TATT_P, TREC_P, TACT_P, TEXAM_P, P As Double


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        QZ1_P = Val(Box1.Text)
        QZ2_P = Val(Box2.Text)
        QZ3_P = Val(Box3.Text)
        ATT_P = Val(Box4.Text)
        REC1_P = Val(Box5.Text)
        REC2_P = Val(Box6.Text)
        ACT1_P = Val(Box7.Text)
        ACT2_P = Val(Box8.Text)
        ACT3_P = Val(Box9.Text)
        EXAM_P = Val(Box10.Text)

        TQ1_P = ((QZ1_P + QZ2_P + QZ3_P) / 3) * 0.25
        TATT_P = ATT_P * 0.1
        TREC_P = ((REC1_P + REC2_P) / 2) * 0.1
        TACT_P = ((ACT1_P + ACT2_P + ACT3_P) / 3) * 0.25
        TEXAM_P = EXAM_P * 0.3

        P = (TQ1_P + TATT_P + TREC_P + TACT_P + TEXAM_P) * 0.3

        Box31.Text = P

    End Sub


    ' MIDTERM '
    Dim QZ1_M, QZ2_M, QZ3_M, ATT_M, REC1_M, REC2_M, ACT1_M, ACT2_M, ACT3_M, EXAM_M As Double
    Dim TQ1_M, TATT_M, TREC_M, TACT_M, TEXAM_M, M As Double

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        QZ1_M = Val(Box11.Text)
        QZ2_M = Val(Box12.Text)
        QZ3_M = Val(Box13.Text)
        ATT_M = Val(Box14.Text)
        REC1_M = Val(Box15.Text)
        REC2_M = Val(Box16.Text)
        ACT1_M = Val(Box17.Text)
        ACT2_M = Val(Box18.Text)
        ACT3_M = Val(Box19.Text)
        EXAM_M = Val(Box20.Text)

        TQ1_M = ((QZ1_M + QZ2_M + QZ3_M) / 3) * 0.25
        TATT_M = ATT_M * 0.1
        TREC_M = ((REC1_M + REC2_M) / 2) * 0.1
        TACT_M = ((ACT1_M + ACT2_M + ACT3_M) / 3) * 0.25
        TEXAM_M = EXAM_M * 0.3

        M = (TQ1_M + TATT_M + TREC_M + TACT_M + TEXAM_M) * 0.3

        Box32.Text = M

    End Sub


    ' FINALS '
    Dim QZ1_F, QZ2_F, QZ3_F, ATT_F, REC1_F, REC2_F, ACT1_F, ACT2_F, ACT3_F, EXAM_F As Double
    Dim TQ1_F, TATT_F, TREC_F, TACT_F, TEXAM_F, F As Double

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        QZ1_F = Val(Box21.Text)
        QZ2_F = Val(Box22.Text)
        QZ3_F = Val(Box23.Text)
        ATT_F = Val(Box24.Text)
        REC1_F = Val(Box25.Text)
        REC2_F = Val(Box26.Text)
        ACT1_F = Val(Box27.Text)
        ACT2_F = Val(Box28.Text)
        ACT3_F = Val(Box29.Text)
        EXAM_F = Val(Box30.Text)

        TQ1_F = ((QZ1_F + QZ2_F + QZ3_F) / 3) * 0.25
        TATT_F = ATT_F * 0.1
        TREC_F = ((REC1_F + REC2_F) / 2) * 0.1
        TACT_F = ((ACT1_F + ACT2_F + ACT3_F) / 3) * 0.25
        TEXAM_F = EXAM_F * 0.3

        F = (TQ1_F + TATT_F + TREC_F + TACT_F + TEXAM_F) * 0.4

        Box33.Text = F

    End Sub


    ' TOTAL GRADE '
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Box34.Text = P + M + F


    End Sub

End Class
