Option Explicit
Option Base 1
Dim dsName As String
Dim calStruct As String, cal0 As String, cal1 As String, cal2 As String, cal3 As String, cal4 As String, cal5 As String, cal6 As String
Dim kfStruct As String, kf1 As String, kf2 As String, kf3 As String, kf4 As String, kf5 As String, kf6 As String


Public Sub Callback_AfterRedisplay()

Dim StartTime As Double
Dim SecondsElapsed As Double

'Remember time when macro starts
  StartTime = Timer

'*****************************
'Start of code
'*****************************

Call initQueryUID
Call setTabDisplay
Call setStyle
Call setFormula


'*****************************
'End of code
'*****************************


'Determine how many seconds code took to run
  SecondsElapsed = Round(Timer - StartTime, 2)

'Notify user in seconds
  MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation


End Sub

Public Sub initQueryUID()

'Datasource name
dsName = "DS_1"

'Calendar structure declaration
calStruct = "5B2JHK6U1ZDQ3S4CT6LQVIV5O"
cal0 = "5B1Y8KZ3ZC4IC1RDKQ9STXXKS"
cal1 = "5B2JIE9WBIEO22ANNXRU97U4C"
cal2 = "5B2JHKM73WL551794UQFFMSL8"
cal3 = "5B2JHKTVMV6UNNQPAOSRPORB0"
cal4 = "5B2JHL1K5TSK6AA5GIV3ZQQ0S"
cal5 = "5B2JHL98OSE9OWTLMCXG9SOQK"
cal6 = "5B2JHLGX7QZZ7JD1S6ZSJUNGC"

'Key figure structure declaration
kfStruct = "5B106DSEYTLDAIO11EWMD7WHO"
kf1 = "5B3B84KF3MEQMV1K32KATI9FG"
kf2 = "5B3B84S3ML0G5HL08WMN3K858"
kf3 = "5B3B9PBN7ATA9EQZ77V140RWC"
kf4 = "5B3B9PJBQ9EZS1AFD1XDE2QM4"
kf5 = "5B35BM5P2NF4EFDUO7XZX66V0"
kf6 = "5B35BN0F6HTYGXJNBK7D1E1Q4"


End Sub

Public Sub setTabDisplay()

Dim lResult As Long
Dim val As String, val1 As String, val2 As String, val3 As String, val4 As String, val5 As String
Dim member As String

member = calStruct
val1 = cal1 'janvier
val2 = cal2 'fevrier
val3 = cal3 'mars
val4 = cal4 'avril
val = val1 & ";" & val2 & ";" & val3 & ";" & val4 & ";" & val5

lResult = Application.Run("SAPSetFilter", dsName, member, val, "INPUT_STRING")


End Sub

Public Sub setFormula()

Dim cell As Range

For Each cell In Range("SAPCrosstab1")
    If cell.Style = "SAPExceptionLevel1" Then
        cell.FormulaR1C1 = "=SUM(R[-1]C+RC[-1])"
    End If
Next cell


End Sub



Public Sub setStyle()

Dim styleRules(10, 6) As Variant

'Style 1 : Editable fields
styleRules(1, 1) = "Format2"
styleRules(1, 2) = "Editable1"
styleRules(1, 3) = dsName
styleRules(1, 4) = "SAPExceptionLevel1"
styleRules(1, 5) = "MEMBER" & ";" & kfStruct & ";" & kf3
styleRules(1, 6) = "MEMBER" & ";" & calStruct & ";" & cal3

Call setFormatRules(styleRules)


End Sub

Public Sub setFormatRules(styleRules As Variant)

Dim i As Long
Dim v1 As String, v2 As String, v3 As String, v4 As String, v5 As String, v6 As String


For i = LBound(styleRules, 1) To UBound(styleRules, 1)
    v1 = styleRules(i, 1)
    v2 = styleRules(i, 2)
    v3 = styleRules(i, 3)
    v4 = styleRules(i, 4)
    v5 = styleRules(i, 5)
    v6 = styleRules(i, 6)
    
    Select Case v1
        Case "Format2" 'Type 2 members
            Call setFormat2(v2, v3, v4, v5, v6)
    
    End Select
Next i

End Sub



Public Sub setFormat2(ruleName As String, dataSource As String, styleName As String, Member1 As String, Member2 As String)

Dim lResult As Variant

lResult = Application.Run("SAPSetFormat", ruleName, dataSource, styleName, "TUPLE", Member1, "TUPLE", Member2)


End Sub
