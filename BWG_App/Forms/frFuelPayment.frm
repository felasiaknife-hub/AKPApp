VERSION 5.00
Begin VB.Form frFuelPayment 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frFuelPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub ShowFuel()
'Dim tmpS As New BWGSearchManager
'Dim x%, tmpItem As BWGFuelBuyRecord, tmpList As ListItem
'    Set colFuel_ = tmpS.FuelBuyRecordSearch("VendorID='" & curVendor_.ID & "'")
'    Set tmpS = Nothing
'    lvFuel.ListItems.Clear
'    For x = 1 To colFuel_.Count
'        Set tmpItem = colFuel_(x)
'        Set tmpList = lvFuel.ListItems.Add(, "'" & tmpItem.ID & "'", FormatYMD_to_DMY(tmpItem.BuyDate, "/", "/"))
'        tmpList.SubItems(1) = Trim(tmpItem.CarRegisID)
'        tmpList.SubItems(2) = Trim(tmpItem.DriverName)
'        tmpList.SubItems(3) = Format(tmpItem.FuelQty * tmpItem.UnitPrice, "#,##0.00")
'        Set tmpItem = Nothing
'        Set tmpList = Nothing
'    Next
'End Sub
'
