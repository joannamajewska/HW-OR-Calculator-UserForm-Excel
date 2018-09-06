VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calculator 
   Caption         =   "UserForm1"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11250
   OleObjectBlob   =   "Calculator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Dim cHomo, homo, rHomo As Integer
    Dim sumAll As Integer
    Dim p, q As Double
    Dim p2, pq, q2 As Double
    Dim cHomoEN, homoEN, rHomoEN As Double
    Dim cHomoOF, homoOF, rHomoOF As Double
    Dim cHomoEF, homoEF, rHomoEF As Double
    Dim chi, pValue As Double
    Dim chControl, hControl, rhControl As Integer
    Dim chTreatment, hTreatment, rhTreatment As Integer
    Dim chSumRow, hSumRow, rhSumRow As Integer
    Dim sumControlAll, sumTreatmentAll, sumRowAll As Integer
    Dim chControlE, hControlE, rhControlE As Double
    Dim chTreatmentE, hTreatmentE, rhTreatmentE As Double
    Dim chi1squared, chi2squared, chi3squared As Double
    Dim chN As Control
    Dim hN As Control
    Dim rhN As Control
    Dim chNC As Control
    Dim chNT As Control
    Dim hNC As Control
    Dim hNT As Control
    Dim rhNC As Control
    Dim rhNT As Control

Private Sub calculate_Click()
    
    cHomo = Val(chNumber.Text)
    homo = Val(hNumber.Text)
    rHomo = Val(rhNumber.Text)
            
    On Error GoTo errorHandler
    
    sumAll = cHomo + homo + rHomo
    p = (2 * cHomo + homo) / (2 * sumAll)
    q = (2 * rHomo + homo) / (2 * sumAll)
        
    p2 = p ^ 2
    pq = 2 * p * q
    q2 = q ^ 2
    cHomoEN = p2 * sumAll
    homoEN = pq * sumAll
    rHomoEN = q2 * sumAll
    cHomoOF = cHomo / sumAll
    homoOF = homo / sumAll
    rHomoOF = rHomo / sumAll
    cHomoEF = cHomoEN / sumAll
    homoEF = homoEN / sumAll
    rHomoEF = rHomoEN / sumAll
    
    chi = (cHomo - cHomoEN) ^ 2 / cHomoEN + _
        (homo - homoEN) ^ 2 / homoEN + _
        (rHomo - rHomoEN) ^ 2 / rHomoEN
    df = 1
    pValue = Application.WorksheetFunction.ChiDist(chi, df)
    
    pFreq.Value = Round(p, 3)
    qFreq.Value = Round(q, 3)
    chObsN.Value = cHomo
    hObsN.Value = homo
    rhObsN.Value = rHomo
    chExpN.Value = Round(cHomoEN, 3)
    hExpN.Value = Round(homoEN, 3)
    rhExpN.Value = Round(rHomoEN, 3)
    chObsFreq.Value = Round(cHomoOF, 3)
    hObsFreq.Value = Round(homoOF, 3)
    rhObsFreq.Value = Round(rHomoOF, 3)
    chExpFreq.Value = Round(cHomoEF, 3)
    hExpFreq.Value = Round(homoEF, 3)
    rhExpFreq.Value = Round(rHomoEF, 3)

    sumObsN.Value = sumAll
    sumExpN.Value = cHomoEN + homoEN + rHomoEN
    sumObsFreq.Value = cHomoOF + homoOF + rHomoOF
    sumExpFreq.Value = cHomoEF + homoEF + rHomoEF
    
    df.Value = df
    chiSquare.Value = Round(chi, 3)
    pV.Value = Round(pValue, 3)
    
    If pValue > 0.05 Then
        MsgBox "The population is in Hardy - Weinberg proportions", vbInformation
    Else
        MsgBox "The population is not in Hardy - Weinberg proportions", vbInformation
    End If
    
    Exit Sub

errorHandler:
    MsgBox "It does not make sense! More than one observed number is equal to 0.", vbCritical

End Sub

Private Sub calculate2_Click()
    
    chControl = Val(chNumberControl.Text)
    hControl = Val(hNumberControl.Text)
    rhControl = Val(rhNumberControl.Text)
    chTreatment = Val(chNumberTreatment.Text)
    hTreatment = Val(hNumberTreatment.Text)
    rhTreatment = Val(rhNumberTreatment.Text)
    
    chSumRow = chControl + chTreatment
    hSumRow = hControl + hTreatment
    rhSumRow = rhControl + rhTreatment
    sumControlAll = chControl + hControl + rhControl
    sumTreatmentAll = chTreatment + hTreatment + rhTreatment
    sumRowAll = chSumRow + hSumRow + rhSumRow

    chObsControl.Value = chControl
    hObsControl.Value = hControl
    rhObsControl.Value = rhControl
    chObsTreatment.Value = chTreatment
    hObsTreatment.Value = hTreatment
    rhObsTreatment.Value = rhTreatment
    chExpControl.Value = Round(sumControlAll * chSumRow / sumRowAll, 3)
    hExpControl.Value = Round(sumControlAll * hSumRow / sumRowAll, 3)
    rhExpControl.Value = Round(rhControlE = sumControlAll * rhSumRow / sumRowAll, 3)
    chExpTreatment.Value = Round(sumTreatmentAll * chSumRow / sumRowAll, 3)
    hExpTreatment.Value = Round(sumTreatmentAll * hSumRow / sumRowAll, 3)
    rhExpTreatment.Value = Round(rhTreatmentE = sumTreatmentAll * rhSumRow / sumRowAll, 3)
    sumObsControl.Value = sumControlAll
    sumExpControl.Value = chControlE + hControlE + rhControlE
    sumExpTreatment.Value = chTreatmentE + hTreatmentE + rhTreatmentE
    sumObsTreatment.Value = sumTreatmentAll
    chSum.Value = chSumRow
    hSum.Value = hSumRow
    rhSum.Value = rhSumRow
    totalSum.Value = sumRowAll
    
    chi1squared = (chControl - (sumControlAll * chSumRow / sumRowAll)) ^ 2 / (sumControlAll * chSumRow / sumRowAll) + _
    (hControl - (sumControlAll * hSumRow / sumRowAll)) ^ 2 / (sumControlAll * hSumRow / sumRowAll) + _
    (rhControl - (sumControlAll * rhSumRow / sumRowAll)) ^ 2 / (sumControlAll * rhSumRow / sumRowAll) + _
    (chTreatment - (sumTreatmentAll * chSumRow / sumRowAll)) ^ 2 / (sumTreatmentAll * chSumRow / sumRowAll) + _
    (hTreatment - (sumTreatmentAll * hSumRow / sumRowAll)) ^ 2 / (sumTreatmentAll * hSumRow / sumRowAll) + _
    (rhTreatment - (sumTreatmentAll * rhSumRow / sumRowAll)) ^ 2 / (sumTreatmentAll * rhSumRow / sumRowAll)
    
    chhObsControl.Value = chControl + hControl
    chhObsTreatment.Value = chTreatment + hTreatment
    rhObsControl2.Value = rhControl
    rhObsTreatment2.Value = rhTreatment
    chhExpControl2.Value = Round(sumControlAll * (chControl + hControl + chTreatment + hTreatment) / sumRowAll, 3)
    rhExpControl2.Value = Round(sumControlAll * rhSumRow / sumRowAll, 3)
    chhExpTreatment2.Value = Round(sumTreatmentAll * (chControl + hControl + chTreatment + hTreatment) / sumRowAll, 3)
    rhExpTreatment2.Value = Round(sumTreatmentAll * rhSumRow / sumRowAll, 3)
    chSum2.Value = chControl + hControl + chTreatment + hTreatment
    rhSum2.Value = rhSumRow
    sumObsControl2.Value = sumControlAll
    sumObsTreatment2.Value = sumTreatmentAll
    sumExpControl2.Value = chControlE + hControlE + rhControlE
    sumExpTreatment2.Value = chTreatmentE + hTreatmentE + rhTreatmentE
    totalSum2.Value = sumRowAll
    
    chi2squared = ((chControl + hControl) - (sumControlAll * (chControl + hControl + chTreatment + hTreatment) / sumRowAll)) ^ 2 _
    / (sumControlAll * (chControl + hControl + chTreatment + hTreatment) / sumRowAll) + _
    (rhControl - (sumControlAll * rhSumRow / sumRowAll)) ^ 2 / (sumControlAll * rhSumRow / sumRowAll) + _
    ((chTreatment + hTreatment) - (sumTreatmentAll * (chControl + hControl + chTreatment + hTreatment) / sumRowAll)) ^ 2 _
    / (sumTreatmentAll * (chControl + hControl + chTreatment + hTreatment) / sumRowAll) + _
    (rhTreatment - (sumTreatmentAll * rhSumRow / sumRowAll)) ^ 2 / (sumTreatmentAll * rhSumRow / sumRowAll)

    chObsControl3.Value = chControl
    chObsTreatment3.Value = chTreatment
    hrhObsControl.Value = hControl + rhControl
    hrhObsTreatment.Value = hTreatment + rhTreatment
    chExpControl3.Value = Round(sumControlAll * chSumRow / sumRowAll, 3)
    hrhExpControl3.Value = Round(sumTreatmentAll * (hControl + rhControl + hTreatment + rhTreatment) / sumRowAll, 3)
    chExpTreatment3.Value = Round(sumTreatmentAll * chSumRow / sumRowAll, 3)
    hrhExpTreatment3.Value = Round(sumTreatmentAll * (hControl + rhControl + hTreatment + rhTreatment) / sumRowAll, 3)
    chSum3.Value = chSumRow
    rhSum3.Value = hControl + rhControl + hTreatment + rhTreatment
    sumObsControl3.Value = sumControlAll
    sumObsTreatment3.Value = sumTreatmentAll
    sumExpControl3.Value = chControlE + hControlE + rhControlE
    sumExpTreatment3.Value = chTreatmentE + hTreatmentE + rhTreatmentE
    totalSum3.Value = sumRowAll
    
    chi3squared = (chControl - (sumControlAll * chSumRow / sumRowAll)) ^ 2 / (sumControlAll * chSumRow / sumRowAll) + _
    ((hControl + rhControl) - (sumControlAll * (hControl + rhControl + hTreatment + rhTreatment) / sumRowAll)) ^ 2 _
    / (sumControlAll * (hControl + rhControl + hTreatment + rhTreatment) / sumRowAll) + _
    (chTreatment - (sumTreatmentAll * chSumRow / sumRowAll)) ^ 2 / (sumTreatmentAll * chSumRow / sumRowAll) + _
    ((hTreatment + rhTreatment) - (sumTreatmentAll * (hControl + rhControl + hTreatment + rhTreatment) / sumRowAll)) ^ 2 _
    / (sumTreatmentAll * (hControl + rhControl + hTreatment + rhTreatment) / sumRowAll)
    
    chi1.Value = Round(chi1squared, 3)
    chi2.Value = Round(chi2squared, 3)
    chi3.Value = Round(chi3squared, 3)
    df1.Value = 2
    df2.Value = 1
    df3.Value = 1
    p1v.Value = Round(Application.WorksheetFunction.ChiDist(chi1squared, 2), 3)
    p2v.Value = Round(Application.WorksheetFunction.ChiDist(chi2squared, 1), 3)
    p3v.Value = Round(Application.WorksheetFunction.ChiDist(chi3squared, 1), 3)
    or1.Value = "NA"
    or2.Value = Round(((chTreatment + hTreatment) / rhTreatment) / ((chControl + hControl) / rhControl), 3)
    or3.Value = Round((chTreatment / (hTreatment + rhTreatment)) / (chControl / (hControl + rhControl)), 3)
    
End Sub

Private Sub checkValues(ByRef tb As Control)

    If tb.Value < 0 Then
        tb.Value = Abs(tb.Value)
    ElseIf tb.Value = "" Then
        tb.Value = 1
    ElseIf Not IsNumeric(tb.Value) Then
        MsgBox "An incorrect value was provided! Enter a numerical value", vbCritical
        tb.Value = 1
    End If

End Sub

Private Sub chNumber_Change()

    Set chN = Me![chNumber]
    checkValues chN

End Sub

Private Sub hNumber_Change()

    Set hN = Me![hNumber]
    checkValues hN
    

End Sub

Private Sub rhNumber_Change()

    Set rhN = Me![rhNumber]
    checkValues rhN

End Sub

Private Sub checkValues2(ByRef tb As Control)

    If tb.Value < 0 Then
        tb.Value = Abs(tb.Value)
    ElseIf tb.Value = "" Then
        tb.Value = 1
    ElseIf tb.Value = 0 Then
        MsgBox "In case - control studies, the number of groups can not be 0. The default value is set to 1.", vbExclamation
        tb.Value = 1
    ElseIf Not IsNumeric(tb.Value) Then
        MsgBox "An incorrect value was provided! The default value is set to 1.", vbCritical
        tb.Value = 1
    End If

End Sub

Private Sub chNumberControl_Change()

    Set chNC = Me![chNumberControl]
    checkValues2 chNC
    
End Sub

Private Sub chNumberTreatment_Change()

    Set chNT = Me![chNumberTreatment]
    checkValues2 chNT
    
End Sub

Private Sub hNumberControl_Change()

    Set hNC = Me![hNumberControl]
    checkValues2 hNC

End Sub

Private Sub hNumberTreatment_Change()

    Set hNT = Me![hNumberTreatment]
    checkValues2 hNT

End Sub

Private Sub rhNumberControl_Change()

    Set rhNC = Me![rhNumberControl]
    checkValues2 rhNC

End Sub

Private Sub rhNumberTreatment_Change()

    Set rhNT = Me![rhNumberTreatment]
    checkValues2 rhNT

End Sub

Private Sub chChange_change()

    If chNumber.Value = "" Then
        chNumber.Value = 0 + chChange.Value
        chChange.Value = 0
    ElseIf chNumber.Value < 0 Then
        chNumber.Value = 0
    Else
        chNumber.Value = chNumber.Value + chChange.Value
        chChange.Value = 0
    End If
    
End Sub

Private Sub hChange_Change()

    If hNumber.Value = "" Then
        hNumber.Value = 0 + hChange.Value
        hChange.Value = 0
    ElseIf hNumber.Value < 0 Then
        hNumber.Value = 0
    Else
        hNumber.Value = hNumber.Value + hChange.Value
        hChange.Value = 0
    End If
End Sub

Private Sub rhChange_Change()

    If rhNumber.Value = "" Then
        rhNumber.Value = 0 + rhChange.Value
        rhChange.Value = 0
    ElseIf rhNumber.Value < 0 Then
        rhNumber.Value = 0
    Else
        rhNumber.Value = rhNumber.Value + rhChange.Value
        rhChange.Value = 0
    End If

End Sub

Private Sub hwCheck_Click()
    If hwCheck.Value = True Then
        labelCheck.Enabled = True
        controlCheck.Enabled = True
        treatmentCheck.Enabled = True
    Else
        labelCheck.Enabled = False
        controlCheck.Enabled = False
        treatmentCheck.Enabled = False
    End If
    
End Sub

Private Sub treatmentCheck_Click()

    If treatmentCheck.Value = True Then
        chNumber.Value = chNumberTreatment.Value
        hNumber.Value = hNumberTreatment.Value
        rhNumber.Value = rhNumberTreatment.Value
        calculate_Click
        MultiPage1.Value = 0
        
    End If
    
End Sub

Private Sub controlCheck_Click()

    If controlCheck.Value = True Then
        chNumber.Value = chNumberControl.Value
        hNumber.Value = hNumberControl.Value
        rhNumber.Value = rhNumberControl.Value
        calculate_Click
        MultiPage1.Value = 0
    End If
    
End Sub

Private Sub editVisible_Click()

    If editVisible.Value = True Then
        chNumber.Enabled = True
        hNumber.Enabled = True
        rhNumber.Enabled = True
        chChange.Enabled = True
        hChange.Enabled = True
        rhChange.Enabled = True
    Else
        chNumber.Enabled = False
        hNumber.Enabled = False
        rhNumber.Enabled = False
        chChange.Enabled = False
        hChange.Enabled = False
        rhChange.Enabled = False
    End If

End Sub

Private Sub editVisible2_Click()

    If editVisible2.Value = True Then
        chNumberControl.Enabled = True
        hNumberControl.Enabled = True
        rhNumberControl.Enabled = True
        chNumberTreatment.Enabled = True
        hNumberTreatment.Enabled = True
        rhNumberTreatment.Enabled = True
    Else
        chNumberControl.Enabled = False
        hNumberControl.Enabled = False
        rhNumberControl.Enabled = False
        chNumberTreatment.Enabled = False
        hNumberTreatment.Enabled = False
        rhNumberTreatment.Enabled = False
    End If

End Sub
