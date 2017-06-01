Attribute VB_Name = "ResetWorkbook_DeveloperMode"
Option Explicit

Sub ResetWkbk()
Attribute ResetWkbk.VB_ProcData.VB_Invoke_Func = " \n14"
    Windows("Sep Sci Instrument Expiration dates.xls").Activate
    Sheets("2015").Select
    Range("S19:S25").Select
    Selection.ClearContents
    Sheets("2014").Select
    Range("S33,S38,S58,S66,S73").Select
    Selection.ClearContents
    Windows("PM-Manager v2.1.xlsm").Activate
    Sheets("Organize Data").Select
    Columns("A:AP").Select
    Selection.ClearContents
    Range("A1").Select
    Sheets("Sheet1").Select
End Sub
