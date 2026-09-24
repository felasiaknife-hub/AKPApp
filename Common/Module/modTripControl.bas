Attribute VB_Name = "modTripControl"
Option Explicit

Public Sub AddMinWeight(ByRef colMinWeightCheck As Collection, ByRef ColWasteWeight As Collection, ByVal TimeTableNo As String)
    Dim isUnderWeight As Boolean, isOverWeight As Boolean
    Dim SumTotalWeight As Double, MaxWeight As Double, MaxWasteID As String, selMID As String, MaxWasteType As String
    Dim tmpWaste As BWGWasteDataCR, tmpTripMinW As BWGTripMinWeight, x As Integer
    '*** เพิ่มข้อมูลลงใน colMinWeightCheck ในกรณีที่มีใบคุมใบเดียว หรือ ใบคุมใบสุดท้าย
    SumTotalWeight = 0
    For x = 1 To ColWasteWeight.Count
        Set tmpWaste = ColWasteWeight(x)
        If InStr(1, tmpWaste.TreatmentUnitID, "กิโลกรัม") Then
            SumTotalWeight = SumTotalWeight + (tmpWaste.MaxQty / 1000)
        Else
           SumTotalWeight = SumTotalWeight + tmpWaste.MaxQty
        End If
        'SumTotalWeight = SumTotalWeight + tmpWaste.MaxQty
        If MaxWeight < tmpWaste.MaxQty Then
            MaxWeight = tmpWaste.MaxQty
            MaxWasteID = tmpWaste.ID
            MaxWasteType = tmpWaste.WasteType
            selMID = tmpWaste.menifestID
        End If
        Set tmpWaste = Nothing
    Next
    
    '*** เช็คว่าน้ำหนักรวมของเวสต์ น้อยกว่าน้ำหนักขั้นต่ำหรือไม่
    If isExist(ColWasteWeight, "'" & MaxWasteID & "'") Then
         If ColWasteWeight("'" & MaxWasteID & "'").MinWeightPerCar > 0 Then
            isUnderWeight = SumTotalWeight < ColWasteWeight("'" & MaxWasteID & "'").MinWeightPerCar
        Else
            isUnderWeight = False
        End If

    End If
    
    '*** เช็คว่าน้ำหนักรวมของเวสต์มากกว่าน้ำหนักที่กำหนดไว้หรือไม่
    If isExist(ColWasteWeight, "'" & MaxWasteID & "'") Then
        If ColWasteWeight("'" & MaxWasteID & "'").MaxWeightPerCar > 0 Then
            isOverWeight = SumTotalWeight > ColWasteWeight("'" & MaxWasteID & "'").MaxWeightPerCar
        Else
            isOverWeight = False
        End If
    Else
        isOverWeight = False
    End If
    
    If isExist(colMinWeightCheck, "'" & TimeTableNo & "'") = False Then
        Set tmpTripMinW = New BWGTripMinWeight
        tmpTripMinW.isUnderWeight = isUnderWeight
        tmpTripMinW.isOverWeight = isOverWeight
        Set tmpTripMinW.ColWasteWeight = ColWasteWeight
        tmpTripMinW.MaxWasteID = MaxWasteID
        tmpTripMinW.MaxWeight = MaxWeight
        tmpTripMinW.MaxWasteType = MaxWasteType
        tmpTripMinW.TimeTableNo = TimeTableNo
        tmpTripMinW.TotalWeight = SumTotalWeight
        tmpTripMinW.menifestID = selMID
        colMinWeightCheck.Add tmpTripMinW, "'" & TimeTableNo & "'"
        Set tmpTripMinW = Nothing
        isUnderWeight = False
        Set ColWasteWeight = New Collection
        MaxWasteID = ""
        MaxWeight = 0
        SumTotalWeight = 0
    End If
End Sub

Public Sub CalMinMaxWeight(ByVal MenifestIDArray As String, ByRef colTrip As Collection, ByRef colMinWeightCheck As Collection)
    On Error GoTo ErrD
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim tmpJobDet As BWGJobDataDetail, tmpWaste As BWGWasteDataCR, tmpS As New BWGSearchManager
    Dim ColWasteWeight As Collection
    Dim curTimeTableNo As String, tmpTimeTableNo As String
    Dim selMID As String, curMenifest As String

    Dim x%, y%, z%, i%
    
    If Trim(MenifestIDArray) = "" Then Exit Sub
    
    '''''''''''''เริ่มการหาขั้นต่ำ
    Set colTrip = tmpS.JobDetailWasteInTripSearch("MenifestID in (" & MenifestIDArray & ")", "WorkDate,TimeTableNo,DocNo, WasteNo, WasteType,TreatmentUnitID")
    Set ColWasteWeight = New Collection
    
    ''''******************* วบเช็ค Waste วนตามใบคุม, ใบเมนิเฟส, ประเภทเวสต์ ********************
    For x = 1 To colTrip.Count
        
        Set tmpJobDet = colTrip(x)
        If curTimeTableNo <> tmpJobDet.TimeTableNo Then
            If Trim(curTimeTableNo) <> "" Then
                '*** เพิ่มข้อมูลลงใน colMinWeightCheck ในกรณีที่มีใบคุมหลายใบ
                Call AddMinWeight(colMinWeightCheck, ColWasteWeight, curTimeTableNo)
            End If
            curTimeTableNo = tmpJobDet.TimeTableNo
            curMenifest = tmpJobDet.menifestID
        End If
        
        '***** Calculate min rate with waste type ********************
        If LCase(tmpJobDet.WasteType) = "non" Then
            If isExist(ColWasteWeight, "'NON'") Then
                Set tmpWaste = ColWasteWeight("'NON'")
                If tmpJobDet.ChargeWeightType = "A" Then
                    tmpWaste.MaxQty = tmpWaste.MaxQty + tmpJobDet.DisposerWeight
                Else
                    tmpWaste.MaxQty = tmpWaste.MaxQty + tmpJobDet.EstWasteQty
                End If
                If tmpWaste.MinWeightPerCar = 0 Then tmpWaste.MinWeightPerCar = tmpJobDet.QuoMinWeightPerCar
                If tmpWaste.NonMinRate = 0 Then tmpWaste.NonMinRate = tmpJobDet.QuoTreatmentRate
                
                If tmpWaste.MaxWeightPerCar = 0 Then tmpWaste.MaxWeightPerCar = tmpJobDet.QuoMaxWeightPerCar
                If tmpWaste.MaxRate = 0 Then tmpWaste.MaxRate = tmpJobDet.QuoMaxRate
                If tmpWaste.MaxRateType = "" Then tmpWaste.MaxRateType = tmpJobDet.QuoMaxRateType
                Set tmpWaste = Nothing
            Else
                Set tmpWaste = New BWGWasteDataCR
                tmpWaste.ID = "NON"
                tmpWaste.WasteName = "NON"
                tmpWaste.WasteType = "Non"
                tmpWaste.menifestID = tmpJobDet.menifestID
                tmpWaste.TreatmentUnitID = tmpJobDet.QuoTreatmentUnitID
                tmpWaste.MinWeightPerCar = tmpJobDet.QuoMinWeightPerCar
                If tmpWaste.NonMinRate = 0 Then tmpWaste.NonMinRate = tmpJobDet.QuoTreatmentRate
                tmpWaste.MaxWeightPerCar = tmpJobDet.QuoMaxWeightPerCar
                tmpWaste.MaxRate = tmpJobDet.QuoMaxRate
                tmpWaste.MaxRateType = tmpJobDet.QuoMaxRateType
                If tmpJobDet.ChargeWeightType = "A" Then
                    tmpWaste.MaxQty = tmpJobDet.DisposerWeight
                Else
                    tmpWaste.MaxQty = tmpJobDet.EstWasteQty
                End If
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
                If tmpWaste.NBLMinRate = 0 Then tmpWaste.NBLMinRate = tmpJobDet.QuoTreatmentRate
                
                If tmpWaste.MaxWeightPerCar = 0 Then tmpWaste.MaxWeightPerCar = tmpJobDet.QuoMaxWeightPerCar
                If tmpWaste.MaxRate = 0 Then tmpWaste.MaxRate = tmpJobDet.QuoMaxRate
                If tmpWaste.MaxRateType = "" Then tmpWaste.MaxRateType = tmpJobDet.QuoMaxRateType
                Set tmpWaste = Nothing
            Else
                Set tmpWaste = New BWGWasteDataCR
                tmpWaste.ID = "NBL"
                tmpWaste.WasteName = "NBL"
                tmpWaste.WasteType = "NBL"
                tmpWaste.menifestID = tmpJobDet.menifestID
                tmpWaste.TreatmentUnitID = tmpJobDet.QuoTreatmentUnitID
                tmpWaste.MinWeightPerCar = tmpJobDet.QuoMinWeightPerCar
                If tmpWaste.NBLMinRate = 0 Then tmpWaste.NBLMinRate = tmpJobDet.QuoTreatmentRate
                tmpWaste.MaxWeightPerCar = tmpJobDet.QuoMaxWeightPerCar
                tmpWaste.MaxRate = tmpJobDet.QuoMaxRate
                tmpWaste.MaxRateType = tmpJobDet.QuoMaxRateType
                If tmpJobDet.ChargeWeightType = "A" Then
                    tmpWaste.MaxQty = tmpJobDet.DisposerWeight
                Else
                    tmpWaste.MaxQty = tmpJobDet.EstWasteQty
                End If
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
                If tmpWaste.HazMinRate = 0 Then tmpWaste.HazMinRate = tmpJobDet.QuoTreatmentRate
                If tmpWaste.HBLMinRate = 0 Then tmpWaste.HBLMinRate = tmpJobDet.QuoTreatmentRate
                If tmpWaste.MaxWeightPerCar = 0 Then tmpWaste.MaxWeightPerCar = tmpJobDet.QuoMaxWeightPerCar
                If tmpWaste.MaxRate = 0 Then tmpWaste.MaxRate = tmpJobDet.QuoMaxRate
                If tmpWaste.MaxRateType = "" Then tmpWaste.MaxRateType = tmpJobDet.QuoMaxRateType
                Set tmpWaste = Nothing
            Else
                Set tmpWaste = New BWGWasteDataCR
                tmpWaste.ID = tmpJobDet.wastedataID
                tmpWaste.WasteName = tmpJobDet.WasteName
                tmpWaste.WasteType = tmpJobDet.WasteType
                tmpWaste.menifestID = tmpJobDet.menifestID
                tmpWaste.TreatmentUnitID = tmpJobDet.QuoTreatmentUnitID
                tmpWaste.MaxWeightPerCar = tmpJobDet.QuoMaxWeightPerCar
                tmpWaste.MaxRate = tmpJobDet.QuoMaxRate
                tmpWaste.MaxRateType = tmpJobDet.QuoMaxRateType
                tmpWaste.MinWeightPerCar = tmpJobDet.QuoMinWeightPerCar
                If tmpWaste.HazMinRate = 0 Then tmpWaste.HazMinRate = tmpJobDet.QuoTreatmentRate
                If tmpWaste.HBLMinRate = 0 Then tmpWaste.HBLMinRate = tmpJobDet.QuoTreatmentRate

                If tmpJobDet.ChargeWeightType = "A" Then
                    tmpWaste.MaxQty = tmpJobDet.DisposerWeight
                Else
                    tmpWaste.MaxQty = tmpJobDet.EstWasteQty
                End If
                ColWasteWeight.Add tmpWaste, "'" & tmpWaste.ID & "'"
                Set tmpWaste = Nothing
            End If
        End If
        tmpTimeTableNo = tmpJobDet.TimeTableNo
        selMID = tmpJobDet.menifestID
        Set tmpJobDet = Nothing
    Next '******** วนเช็คน้ำหนัก Waste วนตามใบคุม, ใบเมนิเฟส, ประเภทเวสต์
    
    '*** เพิ่มข้อมูลลงใน colMinWeightCheck ในกรณีที่มีใบคุมใบเดียว หรือ ใบคุมใบสุดท้าย
    Call AddMinWeight(colMinWeightCheck, ColWasteWeight, tmpTimeTableNo)

    '************* End Calculate min rate with waste type ***********************

    Exit Sub
ErrD:
    Resume Next
'    MsgBox Err.Description & vbCrLf & "เกิดข้อผิดพลาดกรุณาลองใหม่ หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
'    Exit Sub
End Sub


