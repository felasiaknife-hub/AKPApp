Attribute VB_Name = "modCalcFunction"
Option Explicit

Public Function CalTransFeeFromDistance_Cust(fromAumID As String, toAumID As String, TruckTypeID As String, WorkDate As String) As BWGTransVendorChargeRate
Dim tmpCol As Collection
Dim tmpS As New BWGSearchManager
    Set tmpCol = tmpS.TransCustChargeRateSearch("CarTypeID='" & TruckTypeID & "' and FromAumphurID='" & fromAumID & "' And ToAumphurID='" & toAumID & "' AND '" & WorkDate & "' >= StartDate AND '" & WorkDate & "' <= EndDate")
    If tmpCol.Count > 0 Then
        Set CalTransFeeFromDistance_Cust = tmpCol(1)
    Else
        Set CalTransFeeFromDistance_Cust = Nothing
    End If
End Function
Public Function CalJobDataCarUsed(JobDataCarID As String) As BWGJobDataCarUsed
Dim tmpCol As Collection
Dim tmpS As New BWGSearchManager
    Set tmpCol = tmpS.JobDataCarSearch("JobDataCarID= '" & JobDataCarID & "' ")
    If tmpCol.Count > 0 Then
        Set CalJobDataCarUsed = tmpCol(1)
    Else
        Set CalJobDataCarUsed = Nothing
    End If
End Function

Public Function CalTransFeeFromDistance(fromAumID As String, toAumID As String, TruckTypeID As String, WorkDate As String) As BWGTransVendorChargeRate
Dim tmpCol As Collection
Dim tmpS As New BWGSearchManager
    Set tmpCol = tmpS.TransVendorChargeRateSearch("CarTypeID='" & TruckTypeID & "' and FromAumphurID='" & fromAumID & "' And ToAumphurID='" & toAumID & "' AND '" & WorkDate & "' >= StartDate AND '" & WorkDate & "' <= EndDate")
    If tmpCol.Count > 0 Then
        Set CalTransFeeFromDistance = tmpCol(1)
    Else
        Set CalTransFeeFromDistance = Nothing
    End If
End Function

Public Function CalCustTreatFee(selTripControl As BWGJobDataTimeTable) As Double
Dim x%, tmpItem As BWGMenifest, tmpList As ListItem
Dim tmpCol As Collection, tmpAmt As Double, tmpS As New BWGSearchManager
Dim tmpJobDet As BWGJobDataDetail, tmpCol2 As Collection, tmpAddWaste As BWGMenifestWasteAdded
Dim tripType As String, minType$, maxType$, tmpChargeRate As BWGTransVendorChargeRate
Dim curTreatUnitID$, curMfID$, tmpTotalAmt As Double, y%, tmpTranUnit$, nonGroupWeight As Double, maxNonTreatRate As Double
Dim maxNonUnit$, curGroupWeight As Double, tmpWeight As Double, tmpWasteType As String
Dim nonWeight As Double, hazWeight As Double, nblWeight As Double, hblWeight As Double, ColWasteWeight As Collection, tmpWaste As BWGWasteDataCR
Dim MaxWeight As Double, MaxWasteID As String, SumTotalWeight As Double, isUnderWeight As Boolean, colTrip As Collection, z%
Dim amtCol2 As String, amtCol4 As String, curUnitTimeTable$, minMenifestID$
    minType = "ZZZ"
    maxType = "AAA"
    
    '''''''''''''เริ่มการหาขั้นต่ำ
    'Set colTrip = tmpS.JobDetailWasteInTripSearch("JobDataCarID='" & selTripControl.JobDataCarID & "' And TripNo=" & selTripControl.tripNo, "TimeTableNo,DocNo,WorkDate,WasteType,TreatmentUnitID")
    Set colTrip = tmpS.JobDetailWasteInTripSearch("TimeTableNo='" & selTripControl.TimeTableNo & "'", "TimeTableNo,DocNo,WorkDate,WasteType,TreatmentUnitID")
    Set ColWasteWeight = New Collection
    For x = 1 To colTrip.Count
        Set tmpJobDet = colTrip(x)
        If LCase(tmpJobDet.WasteType) = "non" Then
            If isExist(ColWasteWeight, "'NON'") Then
                Set tmpWaste = ColWasteWeight("'NON'")
                If tmpJobDet.ChargeWeightType = "A" Then
                    tmpWaste.MaxQty = tmpWaste.MaxQty + tmpJobDet.DisposerWeight
                Else
                    tmpWaste.MaxQty = tmpWaste.MaxQty + tmpJobDet.EstWasteQty
                End If
                If tmpWaste.MinWeightPerCar = 0 Then tmpWaste.MinWeightPerCar = tmpJobDet.QuoMinWeightPerCar
                If tmpWaste.NonMinRate = 0 Then tmpWaste.NonMinRate = tmpJobDet.NonMinRate
                Set tmpWaste = Nothing
            Else
                Set tmpWaste = New BWGWasteDataCR
                tmpWaste.id = "NON"
                tmpWaste.WasteName = "NON"
                tmpWaste.WasteType = "Non"
                tmpWaste.menifestID = tmpJobDet.menifestID
                tmpWaste.TreatmentUnitID = tmpJobDet.QuoTreatmentUnitID
                tmpWaste.MinWeightPerCar = tmpJobDet.QuoMinWeightPerCar
                tmpWaste.NonMinRate = tmpJobDet.NonMinRate
                If tmpJobDet.ChargeWeightType = "A" Then
                    tmpWaste.MaxQty = tmpJobDet.DisposerWeight
                Else
                    tmpWaste.MaxQty = tmpJobDet.EstWasteQty
                End If
                tmpWaste.Kor = tmpJobDet.menifestID
                ColWasteWeight.Add tmpWaste, "'NON'"
                Set tmpWaste = Nothing
            End If
        ElseIf LCase(tmpJobDet.WasteType) = "nbl" Then
            If isExist(ColWasteWeight, "'NBL'") Then
                Set tmpWaste = ColWasteWeight("'NBL'")
                If tmpJobDet.ChargeWeightType = "A" Then
                    tmpWaste.MaxQty = tmpWaste.MaxQty + tmpJobDet.DisposerWeight
                Else
                    tmpWaste.MaxQty = tmpWaste.MaxQty + tmpJobDet.EstWasteQty
                End If
                If tmpWaste.MinWeightPerCar = 0 Then tmpWaste.MinWeightPerCar = tmpJobDet.QuoMinWeightPerCar
                If tmpWaste.NBLMinRate = 0 Then tmpWaste.NBLMinRate = tmpJobDet.NBLMinRate
                Set tmpWaste = Nothing
            Else
                Set tmpWaste = New BWGWasteDataCR
                tmpWaste.id = "NBL"
                tmpWaste.WasteName = "NBL"
                tmpWaste.WasteType = "NBL"
                tmpWaste.menifestID = tmpJobDet.menifestID
                tmpWaste.TreatmentUnitID = tmpJobDet.QuoTreatmentUnitID
                tmpWaste.MinWeightPerCar = tmpJobDet.QuoMinWeightPerCar
                tmpWaste.NBLMinRate = tmpJobDet.NBLMinRate
                If tmpJobDet.ChargeWeightType = "A" Then
                    tmpWaste.MaxQty = tmpJobDet.DisposerWeight
                Else
                    tmpWaste.MaxQty = tmpJobDet.EstWasteQty
                End If
                tmpWaste.Kor = tmpJobDet.menifestID
                ColWasteWeight.Add tmpWaste, "'NBL'"
                Set tmpWaste = Nothing
            End If
        Else
            If isExist(ColWasteWeight, "'" & tmpJobDet.wastedataID & "'") Then
                Set tmpWaste = ColWasteWeight("'" & tmpJobDet.wastedataID & "'")
                If tmpJobDet.ChargeWeightType = "A" Then
                    tmpWaste.MaxQty = tmpWaste.MaxQty + tmpJobDet.DisposerWeight
                Else
                    tmpWaste.MaxQty = tmpWaste.MaxQty + tmpJobDet.EstWasteQty
                End If
                If tmpWaste.MinWeightPerCar = 0 Then tmpWaste.MinWeightPerCar = tmpJobDet.QuoMinWeightPerCar
                Set tmpWaste = Nothing
            Else
                Set tmpWaste = New BWGWasteDataCR
                tmpWaste.id = tmpJobDet.wastedataID
                tmpWaste.WasteName = tmpJobDet.WasteName
                tmpWaste.WasteType = tmpJobDet.WasteType
                tmpWaste.menifestID = tmpJobDet.menifestID
                tmpWaste.TreatmentUnitID = tmpJobDet.QuoTreatmentUnitID
                tmpWaste.MinWeightPerCar = tmpJobDet.QuoMinWeightPerCar
                If LCase(tmpJobDet.WasteType) = "haz" Then
                    If tmpJobDet.HazMinRate > 0 Then
                        tmpWaste.HazMinRate = tmpJobDet.HazMinRate
                    Else
                        tmpWaste.HazMinRate = tmpJobDet.QuoTreatmentRate
                    End If
                ElseIf LCase(tmpJobDet.WasteType) = "hbl" Then
                    If tmpJobDet.HBLMinRate > 0 Then
                        tmpWaste.HBLMinRate = tmpJobDet.HBLMinRate
                    Else
                        tmpWaste.HBLMinRate = tmpJobDet.QuoTreatmentRate
                    End If
                ElseIf LCase(tmpJobDet.WasteType) = "nbl" Then
                    If tmpJobDet.NBLMinRate > 0 Then
                        tmpWaste.NBLMinRate = tmpJobDet.NBLMinRate
                    Else
                        tmpWaste.NBLMinRate = tmpJobDet.QuoTreatmentRate
                    End If
                End If
                If tmpJobDet.ChargeWeightType = "A" Then
                    tmpWaste.MaxQty = tmpJobDet.DisposerWeight
                Else
                    tmpWaste.MaxQty = tmpJobDet.EstWasteQty
                End If
                tmpWaste.Kor = tmpJobDet.menifestID
                ColWasteWeight.Add tmpWaste, "'" & tmpWaste.id & "'"
                Set tmpWaste = Nothing
            End If
        End If
        Set tmpJobDet = Nothing
    Next
    For x = 1 To ColWasteWeight.Count
        Set tmpWaste = ColWasteWeight(x)
        SumTotalWeight = SumTotalWeight + tmpWaste.MaxQty
        If MaxWeight < tmpWaste.MaxQty Then
            MaxWeight = tmpWaste.MaxQty
            MaxWasteID = tmpWaste.id
            minMenifestID = tmpWaste.Kor
        End If
        Set tmpWaste = Nothing
    Next
    If isExist(ColWasteWeight, "'" & MaxWasteID & "'") Then
        isUnderWeight = SumTotalWeight < ColWasteWeight("'" & MaxWasteID & "'").MinWeightPerCar
    Else
        isUnderWeight = False
    End If
    '''''''''''''่สิ้นสุดการหาขั้นต่ำ
        curTreatUnitID = ""
        'Set tmpCol2 = selTripControl.GetMenifests("T")  'Or "A"
        Set tmpCol2 = selTripControl.GetMenifests("A")
        
        For x = 1 To tmpCol2.Count
            Set tmpItem = tmpCol2(x)
            If Trim(tmpItem.id) <> Trim(curMfID) Then
                curMfID = Trim(tmpItem.id)
                If maxType < LCase(tmpItem.NonOrHz) Then maxType = LCase(tmpItem.NonOrHz)
                If minType > LCase(tmpItem.NonOrHz) Then minType = LCase(tmpItem.NonOrHz)
                '.TextMatrix(curRow, 0) = tmpItem.ID
                '.TextMatrix(curRow, 1) = "วันที่ " & FormatYMD_to_DMY(tmpItem.WorkDate, "/", "/") & " เลขที่ " & tmpItem.docNo
                '.TextMatrix(curRow, 2) = ""
                '.TextMatrix(curRow, 3) = ""
                '.TextMatrix(curRow, 4) = ""
                '.TextMatrix(curRow, 5) = ""
                '.TextMatrix(curRow, 6) = ""
                'curRow = curRow + 1
                'If .Rows - 1 = curRow Then .Rows = .Rows + 5
                If Trim(tmpItem.InOtherDesc) <> "" Then
                    '.TextMatrix(curRow, 0) = tmpItem.ID
                    '.TextMatrix(curRow, 1) = tmpItem.InOtherDesc
                    '.TextMatrix(curRow, 2) = "1"
                    '.TextMatrix(curRow, 3) = ""
                    '.TextMatrix(curRow, 4) = tmpItem.InOtherFee
                    '.TextMatrix(curRow, 5) = ""
                    '.TextMatrix(curRow, 6) = tmpItem.InOtherFee
                    tmpTotalAmt = tmpTotalAmt + tmpItem.InOtherFee
                    'curRow = curRow + 1
                    'If .Rows - 1 = curRow Then .Rows = .Rows + 5
                End If
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Set tmpCol = Nothing
            Set tmpCol = New Collection
            For z = 1 To colTrip.Count
                If Trim(colTrip(z).menifestID) = Trim(tmpItem.id) Then
                    tmpCol.Add colTrip(z), "'" & colTrip(z).id & "'"
                End If
            Next
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            For y = 1 To tmpCol.Count
                Set tmpJobDet = tmpCol(y)
                tmpTranUnit = tmpJobDet.QuoTransUnitID
                If Trim(tmpJobDet.FailType) = "" Then
                    '.TextMatrix(curRow, 0) = tmpItem.ID
                    '.TextMatrix(curRow, 1) = tmpJobDet.WasteName
                    If UCase(tmpJobDet.WasteType) = "NON" Or UCase(tmpJobDet.WasteType) = "NBL" Then
                        If tmpJobDet.ChargeWeightType = "A" Then
                            nonGroupWeight = nonGroupWeight + tmpJobDet.DisposerWeight
                        Else
                            nonGroupWeight = nonGroupWeight + tmpJobDet.EstWasteQty
                        End If
                        If maxNonTreatRate <= tmpJobDet.QuoTreatmentRate Then
                            maxNonTreatRate = tmpJobDet.QuoTreatmentRate
                            maxNonUnit = tmpJobDet.QuoTreatmentUnitID
                        End If
                    End If
                    If tmpJobDet.QuoTreatmentUnitID = "บาท/เที่ยว" Then
                        If Trim(curTreatUnitID) = "" And Trim(curUnitTimeTable) <> Trim(tmpJobDet.TimeTableNo) Then
                            '.TextMatrix(curRow - 1, 2) = "1"
                            '.TextMatrix(curRow - 1, 3) = "เที่ยว"
                            '.TextMatrix(curRow - 1, 4) = Format(tmpJobDet.QuoTreatmentRate, "#,##0.00")
                            '.TextMatrix(curRow - 1, 5) = tmpJobDet.QuoTreatmentUnitID
                            '.TextMatrix(curRow - 1, 6) = Format(tmpJobDet.QuoTreatmentRate, "#,##0.00")
                            tmpTotalAmt = tmpTotalAmt + tmpJobDet.QuoTreatmentRate
                            curUnitTimeTable = Trim(tmpJobDet.TimeTableNo)
                        End If
                        'If UCase(tmpItem.NonOrHz) = "NON" Or UCase(tmpItem.NonOrHz) = "NBL" Then
                            '.TextMatrix(curRow, 2) = ""
                            '.TextMatrix(curRow, 3) = ""
                            '.TextMatrix(curRow, 4) = ""
                            '.TextMatrix(curRow, 5) = ""
                            '.TextMatrix(curRow, 6) = ""
                        'Else
                            'If tmpJobDet.ChargeWeightType = "A" Then
                            '    .TextMatrix(curRow, 2) = Format(tmpJobDet.DisposerWeight, "#,##0.000")
                            'Else
                            '    .TextMatrix(curRow, 2) = Format(tmpJobDet.EstWasteQty, "#,##0.000")
                            'End If
                            '.TextMatrix(curRow, 3) = "ตัน"
                            '.TextMatrix(curRow, 4) = ""
                            '.TextMatrix(curRow, 5) = ""
                            '.TextMatrix(curRow, 6) = ""
                        'End If
                        If tmpJobDet.ChargeWeightType = "A" Then
                            curGroupWeight = curGroupWeight + tmpJobDet.DisposerWeight
                        Else
                            curGroupWeight = curGroupWeight + tmpJobDet.EstWasteQty
                        End If
                        curTreatUnitID = Trim(tmpJobDet.QuoTreatmentUnitID)
                    Else
                        If Trim(curTreatUnitID) <> "" Then
                            '.TextMatrix(curRow, 0) = tmpItem.ID
                            '.TextMatrix(curRow, 1) = "น้ำหนักรวม"
                            '.TextMatrix(curRow, 2) = Format(tmpJobDet.MenifestActQty, "#,##0.000")
                            '.TextMatrix(curRow, 3) = "ตัน"
                            '.TextMatrix(curRow, 4) = ""
                            '.TextMatrix(curRow, 5) = ""
                            '.TextMatrix(curRow, 6) = ""
                            'curRow = curRow + 1
                            'If .Rows - 1 = curRow Then .Rows = .Rows + 5
                        End If
                        curTreatUnitID = ""
                        curGroupWeight = 0
                        If UCase(tmpJobDet.WasteType) <> "NON" And UCase(tmpJobDet.WasteType) <> "NBL" Then
                            If tmpJobDet.ChargeWeightType = "A" Then
                                '.TextMatrix(curRow, 2) = Format(tmpJobDet.DisposerWeight, "#,##0.000")
                                '.TextMatrix(curRow, 3) = Replace(tmpJobDet.QuoTreatmentUnitID, "บาท/", "")
                                tmpWeight = tmpJobDet.DisposerWeight
                            Else
                                '.TextMatrix(curRow, 2) = Format(tmpJobDet.EstWasteQty, "#,##0.000")
                                '.TextMatrix(curRow, 3) = Replace(tmpJobDet.QuoTreatmentUnitID, "บาท/", "")
                                tmpWeight = tmpJobDet.EstWasteQty
                            End If
                            If isUnderWeight And LCase(Trim(MaxWasteID)) = LCase(Trim(tmpJobDet.wastedataID)) Then
                                If LCase(tmpJobDet.WasteType) = "haz" Then
                                    If tmpJobDet.HazMinRate > 0 Then
                                        '.TextMatrix(curRow, 4) = Format(tmpJobDet.HazMinRate, "#,##0.00")
                                        amtCol4 = Format(tmpJobDet.HazMinRate, "#,##0.00")
                                    Else
                                        '.TextMatrix(curRow, 4) = Format(tmpJobDet.QuoTreatmentRate, "#,##0.00")
                                        amtCol4 = Format(tmpJobDet.QuoTreatmentRate, "#,##0.00")
                                    End If
                                ElseIf LCase(tmpJobDet.WasteType) = "hbl" Then
                                    If tmpJobDet.HBLMinRate > 0 Then
                                        '.TextMatrix(curRow, 4) = Format(tmpJobDet.HBLMinRate, "#,##0.00")
                                        amtCol4 = Format(tmpJobDet.HBLMinRate, "#,##0.00")
                                    Else
                                        '.TextMatrix(curRow, 4) = Format(tmpJobDet.QuoTreatmentRate, "#,##0.00")
                                        amtCol4 = Format(tmpJobDet.QuoTreatmentRate, "#,##0.00")
                                    End If
                                End If
                                '.row = curRow
                                '.col = 4
                                '.CellForeColor = vbRed
                            Else
                                '.TextMatrix(curRow, 4) = Format(tmpJobDet.QuoTreatmentRate, "#,##0.00")
                                amtCol4 = Format(tmpJobDet.QuoTreatmentRate, "#,##0.00")
                            End If
                            '.TextMatrix(curRow, 5) = tmpJobDet.QuoTreatmentUnitID
                            If IsNumeric(amtCol4) Then
                                '.TextMatrix(curRow, 6) = Format(tmpWeight * CDbl(amtCol4), "#,##0.00")
                                tmpTotalAmt = tmpTotalAmt + CDbl(Format(tmpWeight * CDbl(amtCol4), "#,##0.00"))
                            End If
                        End If
                    End If
                    'curRow = curRow + 1
                    'If .Rows - 1 = curRow Then .Rows = .Rows + 5
                End If
                tmpWasteType = tmpJobDet.WasteType
                Set tmpJobDet = Nothing
            Next y
            If (UCase(tmpWasteType) = "NON" Or UCase(tmpWasteType) = "NBL") And InStr(1, maxNonUnit, "ตัน") > 0 Then
                If isUnderWeight And LCase(Trim(MaxWasteID)) = LCase(Trim(tmpWasteType)) And Trim(minMenifestID) = Trim(tmpItem.id) Then
                    '.TextMatrix(curRow, 0) = tmpItem.ID
                    '.TextMatrix(curRow, 1) = "น้ำหนักรวม"
                    '.TextMatrix(curRow, 2) = Format(nonGroupWeight, "#,##0.000")
                    '.TextMatrix(curRow, 3) = "ตัน"
                    '.TextMatrix(curRow, 4) = ""
                    '.TextMatrix(curRow, 5) = ""
                    '.TextMatrix(curRow, 6) = ""
                    'curRow = curRow + 1
                    'If .Rows - 1 = curRow Then .Rows = .Rows + 5
                    Set tmpWaste = ColWasteWeight("'" & MaxWasteID & "'")
                    '.TextMatrix(curRow, 0) = tmpItem.ID
                    '.TextMatrix(curRow, 1) = "น้ำหนักไม่ถึง " & tmpWaste.MinWeightPerCar & " ตัน คิดน้ำหนักที่"
                    '.TextMatrix(curRow, 2) = Format(tmpWaste.MinWeightPerCar, "#,##0.000")
                    amtCol2 = Format(tmpWaste.MinWeightPerCar, "#,##0.000")
                    '.TextMatrix(curRow, 3) = "ตัน"
                    If LCase(tmpWaste.WasteType) = "non" Then
                        If tmpWaste.NonMinRate > 0 Then
                            '.TextMatrix(curRow, 4) = Format(tmpWaste.NonMinRate, "#,##0.00")
                            amtCol4 = Format(tmpWaste.NonMinRate, "#,##0.00")
                        Else
                            '.TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                            amtCol4 = Format(tmpWaste.TreatmentFee, "#,##0.00")
                        End If
                    ElseIf LCase(tmpWaste.WasteType) = "nbl" Then
                        If tmpWaste.NBLMinRate > 0 Then
                            '.TextMatrix(curRow, 4) = Format(tmpWaste.NBLMinRate, "#,##0.00")
                            amtCol4 = Format(tmpWaste.NBLMinRate, "#,##0.00")
                        Else
                            '.TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                            amtCol4 = Format(tmpWaste.TreatmentFee, "#,##0.00")
                        End If
                    End If
                    '.TextMatrix(curRow, 5) = "บาท/ตัน"
                    '.TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                    tmpTotalAmt = tmpTotalAmt + CDbl(Format(CDbl(amtCol4) * CDbl(amtCol2), "#,##0.00"))
                    'curRow = curRow + 1
                    'If .Rows - 1 = curRow Then .Rows = .Rows + 5
                ElseIf isUnderWeight = False Then
                    '.TextMatrix(curRow, 0) = tmpItem.ID
                    '.TextMatrix(curRow, 1) = "น้ำหนักรวม"
                    '.TextMatrix(curRow, 2) = Format(nonGroupWeight, "#,##0.000")
                    '.TextMatrix(curRow, 3) = "ตัน"
                    '.TextMatrix(curRow, 4) = Format(maxNonTreatRate, "#,##0.00")
                    '.TextMatrix(curRow, 5) = maxNonUnit
                    '.TextMatrix(curRow, 6) = Format(nonGroupWeight * maxNonTreatRate, "#,##0.00")
                    tmpTotalAmt = tmpTotalAmt + Format(nonGroupWeight * maxNonTreatRate, "0.00") * 1
                    'curRow = curRow + 1
                    'If .Rows - 1 = curRow Then .Rows = .Rows + 5
                End If
                nonGroupWeight = 0
                maxNonTreatRate = 0
            ElseIf (UCase(tmpWasteType) = "HAZ" Or UCase(tmpWasteType) = "HBL") And InStr(1, Trim(curTreatUnitID), "เที่ยว") = 0 And Trim(MaxWasteID) <> "" Then
                Set tmpWaste = ColWasteWeight("'" & MaxWasteID & "'")
                If InStr(1, tmpWaste.TreatmentUnitID, "ตัน") > 0 Then
                    If isUnderWeight And LCase(Trim(tmpWaste.WasteType)) = LCase(Trim(tmpWasteType)) And Trim(minMenifestID) = Trim(tmpItem.id) Then
                        'If Trim(.TextMatrix(curRow, 1)) <> "" Then
                        '    curRow = curRow + 1
                        '    If .Rows - 1 = curRow Then .Rows = .Rows + 5
                        'End If
                        '.TextMatrix(curRow, 1) = "น้ำหนักไม่ถึง " & tmpWaste.MinWeightPerCar & " ตัน คิดน้ำหนักเพิ่มอีก"
                        '.TextMatrix(curRow, 2) = Format(tmpWaste.MinWeightPerCar - SumTotalWeight, "#,##0.000")
                        amtCol2 = Format(tmpWaste.MinWeightPerCar - SumTotalWeight, "#,##0.000")
                        '.TextMatrix(curRow, 3) = "ตัน"
                        If LCase(tmpWaste.WasteType) = "haz" Then
                            If tmpWaste.HazMinRate > 0 Then
                                '.TextMatrix(curRow, 4) = Format(tmpWaste.HazMinRate, "#,##0.00")
                                amtCol4 = Format(tmpWaste.HazMinRate, "#,##0.00")
                            Else
                                '.TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                amtCol4 = Format(tmpWaste.TreatmentFee, "#,##0.00")
                            End If
                        ElseIf LCase(tmpWaste.WasteType) = "hbl" Then
                            If tmpWaste.HBLMinRate > 0 Then
                                '.TextMatrix(curRow, 4) = Format(tmpWaste.HBLMinRate, "#,##0.00")
                                amtCol4 = Format(tmpWaste.HBLMinRate, "#,##0.00")
                            Else
                                '.TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                amtCol4 = Format(tmpWaste.TreatmentFee, "#,##0.00")
                            End If
                        End If
                        '.TextMatrix(curRow, 5) = "บาท/ตัน"
                        '.TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                        tmpTotalAmt = tmpTotalAmt + CDbl(Format(CDbl(amtCol4) * CDbl(amtCol2), "#,##0.00"))
                        'curRow = curRow + 1
                        'If .Rows - 1 = curRow Then .Rows = .Rows + 5
                    End If
                End If
                Set tmpWaste = Nothing
            ElseIf Trim(curTreatUnitID) <> "" Then
                '.TextMatrix(curRow, 1) = "น้ำหนักรวม"
                '.TextMatrix(curRow, 2) = Format(curGroupWeight, "#,##0.000")
                '.TextMatrix(curRow, 3) = "ตัน"
                '.TextMatrix(curRow, 4) = ""
                '.TextMatrix(curRow, 5) = ""
                '.TextMatrix(curRow, 6) = ""
                'curRow = curRow + 1
                'If .Rows - 1 = curRow Then .Rows = .Rows + 5
                curGroupWeight = 0
            End If
            Set tmpItem = Nothing
        Next x
        CalCustTreatFee = tmpTotalAmt
End Function

Public Sub CalCustTripFee(selTripControl As BWGJobDataTimeTable, ByRef maxWorker%, ByRef maxTransFee As Double, ByRef DiscountTrans As String)
Dim tmpS As New BWGSearchManager, tmpPrice As Double
Dim tmpCol As Collection ', tmpTransRate As BWGTransVendorChargeRate
Dim tmpJobDet As BWGJobDataDetail, x%, maxWorkerRate As Double, tmpAddWaste As BWGMenifestWasteAdded

    'Set tmpCol = tmpS.JobDetailInTripSearch("JobDataCarID='" & selTripControl.JobDataCarID & "' And TripNo=" & selTripControl.tripNo)
    Set tmpCol = tmpS.JobDetailInTripSearch("TimeTableNo='" & selTripControl.TimeTableNo & "' AND isMain='Y' AND isCancel ='N' AND isCanceled <> 'Y'")
    
    For x = 1 To tmpCol.Count
        Set tmpJobDet = tmpCol(x)
        If tmpJobDet.QuoLaborCount > maxWorker Then maxWorker = tmpJobDet.QuoLaborCount
        If tmpJobDet.QuoLaborCharge > maxWorkerRate Then maxWorkerRate = tmpJobDet.QuoLaborCharge
        If DiscountTrans = "" And tmpJobDet.IsDiscountPriceTrans Then
        DiscountTrans = "Y"
        End If
        If tmpJobDet.isPriceIncTrans = False Then
            If tmpJobDet.QuoTransUnitID = "บาท/เที่ยว" Then
                If tmpJobDet.QuoTransportFee > maxTransFee Then maxTransFee = tmpJobDet.QuoTransportFee
            ElseIf tmpJobDet.QuoTransUnitID = "บาท/ตัน" Then
                If tmpJobDet.ChargeWeightType = "A" Then 'ผู้รับกำจัด
                    If tmpJobDet.DisposerWeight < tmpJobDet.QuoMinWeightPerCar Then
                        'tmpPrice = tmpJobDet.QuoMinWeightPerCar * tmpJobDet.QuoTreatmentRate
                        tmpPrice = tmpJobDet.QuoTransportFee
                    Else
                        tmpPrice = tmpJobDet.QuoTransportFee
                    End If
                ElseIf tmpJobDet.ChargeWeightType = "B" Then 'ผู้ก่อกำเนิด
                    If tmpJobDet.EstWasteQty < tmpJobDet.QuoMinWeightPerCar Then
                        'tmpPrice = tmpJobDet.QuoMinWeightPerCar * tmpJobDet.QuoTreatmentRate
                        tmpPrice = tmpJobDet.QuoTransportFee
                    Else
                        'tmpPrice = tmpJobDet.EstWasteQty * tmpJobDet.QuoTreatmentRate
                        tmpPrice = tmpJobDet.QuoTransportFee
                    End If
                End If
                If tmpPrice > maxTransFee Then maxTransFee = tmpPrice
            End If
        End If
        Set tmpJobDet = Nothing
    Next
    
'    If tmpTransRate Is Nothing Then
'        maxTransFee = 0
'    Else
'        If selTripControl.isTrailer Then
'            maxTransFee = tmpTransRate.TrailerRateTrip
'        Else
'            maxTransFee = tmpTransRate.RatePerTrip
'        End If
'    End If
    
    Set tmpCol = tmpS.MenifestWasteAddSearch("JobDataCarID='" & selTripControl.JobDataCarID & "' And TripNo=" & selTripControl.TripNo)
    For x = 1 To tmpCol.Count
        Set tmpAddWaste = tmpCol(x)
        If tmpAddWaste.LaborCount > maxWorker Then maxWorker = tmpAddWaste.LaborCount
        If tmpAddWaste.ExtraLaborFee > maxWorkerRate Then maxWorkerRate = tmpAddWaste.ExtraLaborFee
        If DiscountTrans = "" And tmpAddWaste.IsDiscountPriceTrans Then
            DiscountTrans = "Y"
        End If
        If tmpAddWaste.isPriceIncTrans = False Then
            If tmpAddWaste.TransferUnitID = "บาท/เที่ยว" Then
                If tmpAddWaste.TransportFee > maxTransFee Then maxTransFee = tmpAddWaste.TransportFee
            ElseIf tmpAddWaste.TransferUnitID = "บาท/ตัน" Then
                If tmpAddWaste.ChargeWeightType = "A" Then 'ผู้รับกำจัด
                    If tmpAddWaste.DisposerWeight < tmpAddWaste.MinWeightPerCar Then
                        tmpPrice = tmpAddWaste.MinWeightPerCar * tmpAddWaste.TransportFee
                    Else
                        tmpPrice = tmpAddWaste.DisposerWeight * tmpAddWaste.TransportFee
                    End If
                ElseIf tmpAddWaste.ChargeWeightType = "B" Then 'ผู้ก่อกำเนิด
                    If tmpAddWaste.EstWasteQty < tmpAddWaste.MinWeightPerCar Then
                        tmpPrice = tmpAddWaste.MinWeightPerCar * tmpAddWaste.TransportFee
                    Else
                        tmpPrice = tmpAddWaste.EstWasteQty * tmpAddWaste.TransportFee
                    End If
                End If
                If tmpPrice > maxTransFee Then maxTransFee = tmpPrice
            End If
        End If
        Set tmpJobDet = Nothing
    Next
    
    If selTripControl.EditTypeTrip <> "M" Then
'    If selTripControl.TripTranFee = 0 Then
        selTripControl.TripTranFee = maxTransFee
'**** Remark Date 21/04/2011 : 09:05
'        If maxWorker < selTripControl.WorkerCount Then
'            selTripControl.TripMoreWorker = selTripControl.WorkerCount - maxWorker
'            selTripControl.TripMoreWorkerFee = (selTripControl.WorkerCount - maxWorker) * maxWorkerRate
'        End If
    End If
    
    If maxWorker < selTripControl.WorkerCount Then
        selTripControl.TripMoreWorker = selTripControl.WorkerCount - maxWorker
        selTripControl.TripMoreWorkerFee = (selTripControl.WorkerCount - maxWorker) * IIf(maxWorkerRate = 0, 500, maxWorkerRate)
    End If
    
    Set tmpS = Nothing
    Set tmpCol = Nothing

End Sub

Public Sub CalVendorTripFee(selTripControl As BWGJobDataTimeTable, FromAumphurID As String, ToAumphurID As String, TruckTypeID As String, WorkDate As String)
Dim tmpChargeRate As BWGTransVendorChargeRate
        Set tmpChargeRate = CalTransFeeFromDistance(FromAumphurID, ToAumphurID, TruckTypeID, WorkDate)
        If Not tmpChargeRate Is Nothing Then
            If selTripControl.isTrailer Then
                selTripControl.VDTranFee = tmpChargeRate.TrailerRateTrip
                selTripControl.Allowance = tmpChargeRate.Allowance
                selTripControl.Dream = tmpChargeRate.Dream
            Else
                selTripControl.VDTranFee = tmpChargeRate.RatePerTrip
                selTripControl.Allowance = tmpChargeRate.Allowance
                selTripControl.Dream = tmpChargeRate.Dream
            End If
        End If
End Sub

