VERSION 5.00
Begin VB.Form frmExcel 
   Caption         =   "VB-to-Excel"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdCreateReport 
      Caption         =   "Create Report"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.Label lblMessage 
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCreateReport_Click()
    On Error GoTo handleError
    
    Dim conn As New ADODB.Connection
    Dim xlApp As New Excel.Application
    Dim xlwk As New Excel.Workbook
    
    Dim rs, SQL
    
    conn.ConnectionTimeout = 15
    conn.CommandTimeout = 30
    conn.Open "DSN=nwind;UID=;DATABASE=nwind;"
    
    SQL = "SELECT customerID, companyName, contactName, address, city, phone FROM customers"

    Set rs = conn.Execute(SQL)
    xlApp.Interactive = True

    Set xlwk = xlApp.Workbooks.Open(App.Path & "\excelReport.xls")
    
    ctr = 5 ' start data after headings
    Do While Not rs.EOF
        ctr = ctr + 1
        xlApp.Range("A" & Trim(Str(ctr))).Value = rs("customerID")
        xlApp.Range("B" & Trim(Str(ctr))).Value = rs("companyName")
        xlApp.Range("C" & Trim(Str(ctr))).Value = rs("contactName")
        xlApp.Range("D" & Trim(Str(ctr))).Value = rs("address")
        xlApp.Range("E" & Trim(Str(ctr))).Value = rs("city")
        xlApp.Range("F" & Trim(Str(ctr))).Value = rs("phone")
        rs.MoveNext
    Loop
    
    xlApp.Visible = True
    'xlwk.Save
    'xlApp.Quit
    
handleError:
    If Err.Number <> 0 Then
        MsgBox "Error #: " & Err.Number & vbCrLf & Err.Description, vbCritical, "Critical Error"
    End If

End Sub

Private Sub cmdExit_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()
    lblMessage.Caption = "This is a very simple example that shows you how to use the Excel Object to write database values " & _
                         "to an excel spreadsheet." & vbCrLf & vbCrLf & "I didn't bother spending a lot of time designing an interface or " & _
                         "making this too complex because I wanted to get my point across as easily as possible." & _
                         vbCrLf & vbCrLf & "If you have any questions/comments, you can email me at hovila@hotmail.com"

End Sub
       


