Option Compare Database
Option Explicit

Sub DbaseName()

Dim sName As String
sName = "Order Tracking"
ShowDbaseName (sName)

End Sub

Sub ShowDbaseName(sDbase As String)

MsgBox sDbase

End Sub


Option Compare Database

Sub ShowFormName()

Dim myForm As Form
Set myForm = Forms("Products")
MsgBox myForm.Name

Set myForm = Forms("Categories")
MsgBox myForm.Name

Set myForm = Nothing

End Sub




Option Compare Database

Sub ForNextEx()

Dim iCaseCount As Integer, curCaseCost As Currency
curCaseCost = 47

For iCaseCount = 5 To 10 Step 2

    MsgBox (iCaseCount & " cases costs $ " & iCaseCount * curCaseCost)

Next 'iCaseCount

End Sub


Option Compare Database

Sub ForDemo()

Dim varShipTypes As Variant, varType As Variant, myF As Form
varShipTypes = Array("Overnight", "TwoDay", "ThreeDay", "Ground")

For Each varType In varShipTypes

    MsgBox (varType)
    
Next 'varType

For Each myF In Forms
    
    MsgBox myF.Name

Next 'myF

End Sub


Option Compare Database
Option Explicit

Sub DoDemo()

Dim lngCount As Long, myR As Recordset

'Set myR = CurrentDb.OpenRecordset("Orders")

'Do While lngCount < 25 And Not myR.EOF
    'lngCount = lngCount + myR![Quantity]
    'MsgBox lngCount
    'myR.MoveNext
'Loop

Set myR = CurrentDb.OpenRecordset("Categories")

Do Until myR.EOF

    MsgBox myR![CategoryName]
    myR.MoveNext
    
Loop

End Sub


Option Compare Database

Sub IfTest()

Dim curTotal As Currency, curComm As Currency
curTotal = InputBox("Please enter the sale amount.")

If curTotal >= 1000 Then

    curComm = curTotal * 0.08
    ElseIf curTotal >= 1000 Then curComm = curTotal * 0.06
    
    Else
        curComm = curTotal * 0.05
End If

MsgBox ("Commision is $" & curComm)


End Sub


Option Compare Database
Option Explicit

Sub CalcComm()

Dim curTotal As Currency, curComm As Currency
curTotal = InputBox("Enter the sale amount.")
Select Case curTotal

Case Is >= 10000
    curComm = curTotal * 0.08
    
Case Is >= 1000
    curComm = curTotal * 0.06
    
Case Is >= 500
    curComm = curTotal * 0.06
    
Case Else
    curComm = curTotal * 0.04
    
End Select

MsgBox ("The commision is $" & curComm)

End Sub


Option Compare Database

Sub CheckError()

Dim lngNumber As Long, lngResult As Long

'On Error GoTo 0
'On Error Resume Next
On Error GoTo Handler:

lngNumber = InputBox("Enter a number.")
lngResult = lngNumber + 3

MsgBox (lngResult)
Exit Sub

Handler:
MsgBox ("Not a number.")

End Sub


Option Compare Database

Sub DbaseName()

Dim sName As String
sName = "Order Tracking"

ShowDbaseName (sName)

End Sub

Sub ShowDbaseName(sDbase As String)

MsgBox (sDbase)

End Sub


Option Compare Database

Sub DbaseName()

Dim sName As String
sName = "Order Tracking"

ShowDbaseName (sName)
MsgBox ("End of the subroutine.")

End Sub

Sub ShowDbaseName(sDbase As String)

MsgBox (sDbase)

End Sub


Option Compare Database

Sub CheckError()

Dim lngNumber As Long, lngResult As Long
On Error GoTo Handler:

lngNumber = InputBox("Enter a number.")
Debug.Print lngNumber
lngResult = lngNumber + 3
Debug.Print lngResult

MsgBox (lngResult)

Exit Sub

Handler:
MsgBox ("Not a number.")

End Sub


Option Compare Database

Sub NameDbase()

'Created January 18, 2019

'MsgBox "This is the customer database" 'Add info about contents

MsgBox "This is the orders"

End Sub


Option Compare Database

Sub CountQuantity()

Dim lngCount As Long, myR As Recordset
Set myR = CurrentDb.OpenRecordset("Orders")

Do While lngCount < 25 And Not myR.EOF

    lngCount = lngCount + myR![Quantity]
    myR.MoveNext
    
Loop

End Sub


Option Compare Database

Sub PrintObject()

DoCmd.PrintOut acPrintAll, , , acHigh, 1, True

End Sub

Option Compare Database

Sub AddNewRecord()

Dim myR As Recordset

Set myR = CurrentDb.OpenRecordset("Customers")

myR.AddNew
myR![CustNum] = myR.RecordCount + 1
myR![CustFirstName] = "Curtis"
myR![CustLastName] = "Curtis"

myR.Update


End Sub


Option Compare Database

Function Multbypi(dblLength As Double)

Multbypi = dblLength * 3.14159

End Function

Sub GetCirc()

Dim dbleDiameter As Double
dbleDiameter = InputBox("Please enter the diameter of the circle.")
MsgBox (Multbypi(dbleDiameter))

End Sub


Option Compare Database
Option Explicit

Sub CalcComm()

Dim curValue As Currency, curComm As Currency

curValue = 1500
curComm = curValue * 0.05
MsgBox (curComm)

End Sub



Option Compare Database
Option Explicit
Dim dblRate As Double

Sub GetSale()

Dim curSale As Currency

curSale = InputBox("Enter value of sale.")
dblRate = 0.05
MsgBox ("$" & CalcComm(curSale))

End Sub

Function CalcComm(curValue As Currency)

CalcComm = curValue * dblRate

End Function

Option Compare Database
Const curDelCharge As Currency = 20

Sub DelCharge()

Dim curSaleValue As Currency
Static curSaleTotal As Currency

curSaleValue = InputBox("Value of the sale?")

curSaleTotal = curSaleTotal + curSaleValue + curDelCharge
MsgBox ("$" & curSaleTotal)


End Sub

Sub Del2()

curSaleValue = InputBox("Value of the sale?")

curSaleTotal = curSaleValue + curDelCharge

MsgBox ("$" & curSaleTotal)

End Sub


Option Compare Database
Option Explicit

Sub Explore()

Dim iTotal As Integer, iLeftover As Integer, iCases As Integer

iTotal = 28

iTotal = iTotal + 7
iTotal = iTotal - 14
iTotal = iTotal * 4

iCases = iTotal \ 10
iLeftover = iTotal - (iCases * 10)

MsgBox (iLeftover)

End Sub


Option Compare Database

Sub DisplayValue()

Dim sMessage As String
sMessage = "Welcome to the database."


sCount = "13"
sExplanation = "Users in the last hour."
MsgBox (sMessage & " " & sCount & " " & sExplanation)

End Sub


Option Compare Database

Sub Arrays()

Dim iRates(3, 1) As Currency

For iCounter1 = 0 To 3

    For iCounter2 = 0 To 1
    
        iRates(iCounter1, iCounter2) = (iCounter1 * 5) + (iCounter2 * 10)

    Next
    
Next

For iCounter1 = 0 To 3

    For iCounter2 = 0 To 1
    
        MsgBox (iRates(iCounter1, iCounter2))
    
    Next
    
Next

End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim Name As String
Name = "Mattias"
End Sub

DoCmd.RunSQL "INSERT INTO Table1 (FirstName, LastName) VALUES('" & Name & "', 'Lindgren');"

Public Sub sendMail()

    Dim myMail      As Outlook.MailItem
    Dim myOutlApp   As Outlook.Application

    ' Creating an Outlook-Instance and a new Mailitem
    Set myOutlApp = New Outlook.Application
    Set myMail = myOutlApp.CreateItem(olMailItem)

    With myMail
        ' defining the primary recipient
        .To = "recipient@somewhere.invalid"
        ' adding a CC-recipient
        .CC = "other.recipient@somewhere.else.invalid"
        ' defining a subject for the mail
        .Subject = "My first mail sent with Outlook-Automation"
        ' Addimg some body-text to the mail
        .Body = "Hello dear friend, " & vbCrLf & vbCrLf & _
                "This is my first mail produced and sent via Outlook-Automation." & vbCrLf & vbCrLf & _
                "And now I will try add an attachment."
        ' Adding an attachment from filesystem
        .Attachments.Add "c:\path\to\a\file.dat"

        ' sending the mail
        .Send
        ' You can as well display the generated mail by calling the Display-Method
        ' of the Mailitem and let the user send it manually later. 
    End With

    ' terminating the Outlook-Application instance
    myOutlApp.Quit

    ' Destroy the object variables and free the memory
    Set myMail = Nothing
    Set myOutlApp = Nothing

End Sub
