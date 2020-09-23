VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Easy Librarian!"
   ClientHeight    =   5535
   ClientLeft      =   3060
   ClientTop       =   3120
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   9060
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Book"
      Height          =   495
      Left            =   7320
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddBook 
      Caption         =   "Add Book"
      Default         =   -1  'True
      Height          =   495
      Left            =   7320
      TabIndex        =   14
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtTitleofBook 
      Height          =   285
      Left            =   720
      TabIndex        =   13
      Top             =   3240
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.TextBox txtAuthorLastName 
      Height          =   285
      Left            =   720
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtAuthorFirstName 
      Height          =   285
      Left            =   720
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   3615
   End
   Begin MSDataGridLib.DataGrid dbgBooks 
      Height          =   2175
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   3836
      _Version        =   393216
      BackColor       =   16777215
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         ScrollBars      =   2
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Start Search"
      Height          =   495
      Left            =   7320
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtLastName 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   1111
      ButtonWidth     =   2355
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add New Book"
            Description     =   "Add new book to list."
            Object.ToolTipText     =   "Add new book to list."
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit the Program"
            Description     =   "exit"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search by Author"
            Object.ToolTipText     =   "Search for book by Author"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search by Title"
            Object.ToolTipText     =   "Search for book by Title"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Show All Books"
            Object.ToolTipText     =   "Display all books in the database."
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print Book List"
            Object.ToolTipText     =   "Print the current displayed book list."
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5160
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "This is a test"
            TextSave        =   "This is a test"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTitleofBook 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Title of Book:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblAuthorLastName 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Author Last Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblAuthorFirstName 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Author First Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFF00&
      Caption         =   "Enter title of book:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblLastName 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Enter author's last name:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAdd 
         Caption         =   "A&dd Book to List"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print Book LIst"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuAuthor 
         Caption         =   "Search ByA&uthor"
      End
      Begin VB.Menu mnuTitle 
         Caption         =   "Search By &Title"
      End
      Begin VB.Menu mnuAll 
         Caption         =   "Show a&ll Books"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSearchType As String    'Are we searching by Title or Author?
Dim rst As ADODB.Recordset     'This variable holds the recordset
Dim RowValue As Integer        'Global variable to identify the Database entry
                               'the user has clicked on

Sub dbgBooks_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Get the value of the row that the user clicks on
    RowValue = dbgBooks.RowContaining(Y)
    
    'Make the delete command button visible
    cmdDelete.Visible = True
End Sub


Private Sub cmdAddBook_Click()
    'Declare variables
    Dim cnn As New ADODB.Connection
    Dim intCount As Integer    'number assigned to the database entry
    
    'error handler
    On Error GoTo ErrorHandler
    
    'For Access 97, use Microsoft.Jet.OLEDB.3.51;
    cnn.Open "Provider=Microsoft.Jet.OLEDB.3.51;" & _
        "Data Source=" & App.Path & "\library.mdb;"

    'Create recordset reference and set its properties.
    Set rst = New ADODB.Recordset
    rst.CursorType = adOpenStatic       'can move forward and backwards in the database
    rst.LockType = adLockPessimistic    'lock record when opened

    'Open the recordset.
    rst.Open "books", cnn, , , adCmdTable
    
    'Check textboxes for complete information and exit sub if not complete
    If txtAuthorFirstName.Text = "" Or txtAuthorLastName.Text = "" Or txtTitleofBook.Text = "" Then
        MsgBox "You must enter both the Title and the Author's full name.", vbOKOnly, "Oops"
        'set focus on first blank field
        If txtAuthorFirstName.Text = "" Then
            txtAuthorFirstName.SetFocus
        Else
            If txtAuthorLastName.Text = "" Then
                txtAuthorLastName.SetFocus
            Else
                txtTitleofBook.SetFocus
            End If
        End If
        Exit Sub
    End If

    'Everything is cool, so save the book information
    With rst
        If rst.RecordCount = 0 Then
            intCount = 1
        Else
            .MoveLast
            intCount = .Fields("BookID") + 1    'increment the book ID number
        End If
        .AddNew
        !Title = txtTitleofBook.Text
        !FirstName = txtAuthorFirstName.Text
        !LastName = txtAuthorLastName.Text
        !BookID = intCount
        .Update
    End With
    
    'Clear the text boxes
    txtTitleofBook.Text = ""
    txtAuthorFirstName.Text = ""
    txtAuthorLastName.Text = ""
    MsgBox "Book has been saved.", vbOKOnly, "Done!"
    
    'close the connection to the database
    rst.Close
    cnn.Close
    Exit Sub


ErrorHandler:                   ' Error-handling routine.
    Select Case Err.Number      ' Evaluate error number.
       
       Case Else
         
      Msg = "Unexpected error #" & Str(Err.Number)
      Msg = Msg & " occurred: " & Err.Description
      ' Display message box with Stop sign icon and
      ' OK button.
      MsgBox Msg, vbCritical
      
   End Select
   Resume Next  ' Resume execution at same line that caused the error.

End Sub

Private Sub cmdDelete_Click()
    'Declare variables
    Dim cnn As New ADODB.Connection
    Dim strSearch As String
    Dim response As String
    Dim intCount As Integer
    
    'error handler
    On Error GoTo ErrorHandler
    
    'Verify delete
    response = MsgBox("Are you sure you want to delete this book?", vbYesNo, "Delete Book?")
    If response = vbYes Then
    
        'For Access 97, use Microsoft.Jet.OLEDB.3.51;
        cnn.Open "Provider=Microsoft.Jet.OLEDB.3.51;" & _
            "Data Source=" & App.Path & "\library.mdb;"

        'Create recordset reference and set its properties.
        Set rst = New ADODB.Recordset
        rst.CursorType = adOpenStatic       'forward, reverse
        rst.LockType = adLockPessimistic    'lock record when opened
    
        'Open recordset
        rst.Open "Books", cnn

        'build search string
        strSearch = "BookID = " & (RowValue + 1)   'dbgrid control begins numbering rows at 0
    
        'Find selected database entry and delete it
        With rst
            .MoveFirst
            .Find strSearch
            .Delete
            .Update
        End With
        Set dbgBooks.DataSource = rst    'refresh display
        
        'Update the row numbers
        intCount = 1
        With rst
            .MoveFirst
            Do While Not .EOF
                .Fields("BookID") = intCount
                intCount = intCount + 1
                .MoveNext
            Loop
        End With
        
        cmdDelete.Visible = False    'hide the delete command button again
        
    End If
    
    Exit Sub
    
ErrorHandler:                   ' Error-handling routine.
    Select Case Err.Number      ' Evaluate error number.
       
       Case Else
         
      Msg = "Unexpected error #" & Str(Err.Number)
      Msg = Msg & " occurred: " & Err.Description
      ' Display message box with Stop sign icon and
      ' OK button.
      MsgBox Msg, vbCritical
      
   End Select
   Resume Next  ' Resume execution at same line that caused the error.

End Sub

Private Sub cmdSearch_Click()
    
    'Declare variables
    Dim cnn As New ADODB.Connection
    Dim strSearch As String
    
    'error handler
    On Error GoTo ErrorHandler
    
    'For Access 97, use Microsoft.Jet.OLEDB.3.51;
    cnn.Open "Provider=Microsoft.Jet.OLEDB.3.51;" & _
        "Data Source=" & App.Path & "\library.mdb;"

    'Create recordset reference and set its properties.
    Set rst = New ADODB.Recordset
    rst.CursorType = adOpenStatic       'forward, reverse
    rst.LockType = adLockPessimistic    'lock record when opened

    'Set search query
    If txtTitle.Visible = True Then    'search by title
        strSearch = "Select FirstName, LastName, Title from Books where Title = " & "'" & txtTitle.Text & "'"
        strSearchType = "Title"
    Else                                'search by author, last name only
        strSearch = "Select FirstName, LastName, Title from Books where LastName = " & "'" & txtLastName.Text & "'"
        strSearchType = "Author"
    End If
    
    'Open the recordset.
    rst.Open strSearch, cnn, adOpenStatic, adLockOptimistic
    
    'Display results
    dbgBooks.Visible = True
    Set dbgBooks.DataSource = rst
    dbgBooks.ScrollBars = dbgVertical   'vertical scrollbar only

    Exit Sub


ErrorHandler:                   ' Error-handling routine.
    Select Case Err.Number      ' Evaluate error number.
       
       Case Else
         
      Msg = "Unexpected error #" & Str(Err.Number)
      Msg = Msg & " occurred: " & Err.Description
      ' Display message box with Stop sign icon and
      ' OK button.
      MsgBox Msg, vbCritical
      
   End Select
   Resume Next  ' Resume execution at same line that caused the error.

End Sub


Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuAdd_Click()
    'Handle visiblity of objects
    lblAuthorFirstName.Visible = True
    lblAuthorLastName.Visible = True
    lblTitleofBook.Visible = True
    txtAuthorFirstName.Visible = True
    txtAuthorLastName.Visible = True
    txtTitleofBook.Visible = True
    lblTitle.Visible = False
    txtTitle.Visible = False
    lblLastName.Visible = False
    txtLastName.Visible = False
    cmdSearch.Visible = False
    cmdAddBook.Visible = True
    dbgBooks.Visible = False
    txtAuthorFirstName.SetFocus
    cmdAddBook.Default = True
    StatusBar1.SimpleText = "Click on Save button to save this book to the database."

End Sub

Private Sub mnuAll_Click()
    lblAuthorFirstName.Visible = False
    lblAuthorLastName.Visible = False
    lblTitleofBook.Visible = False
    txtAuthorFirstName.Visible = False
    txtAuthorLastName.Visible = False
    txtTitleofBook.Visible = False
    lblTitle.Visible = False
    txtTitle.Visible = False
    lblLastName.Visible = False
    txtLastName.Visible = False
    cmdSearch.Visible = False
    cmdAddBook.Visible = False
    dbgBooks.Visible = False
    cmdSearch.Default = False
    
        'Declare variables
    Dim cnn As New ADODB.Connection
    Dim strSearch As String
    
    'error handler
    On Error GoTo ErrorHandler
    
    'For Access 97, use Microsoft.Jet.OLEDB.3.51;
    cnn.Open "Provider=Microsoft.Jet.OLEDB.3.51;" & _
        "Data Source=" & App.Path & "\library.mdb;"

    'Create recordset reference and set its properties.
    Set rst = New ADODB.Recordset
    rst.CursorType = adOpenStatic       'forward, reverse
    rst.LockType = adLockPessimistic    'lock record when opened

    'Open the recordset.
    rst.Open "books", cnn, , , adCmdTable
    
    'Display results
    dbgBooks.Visible = True
    Set dbgBooks.DataSource = rst
    dbgBooks.ScrollBars = dbgVertical
    strSearchType = "Author"
    
    StatusBar1.SimpleText = "This shows all of the books in the database."

Exit Sub

ErrorHandler:                       ' Error-handling routine.
    Select Case Err.Number      ' Evaluate error number.
       
       Case Else
         
      Msg = "Unexpected error #" & Str(Err.Number)
      Msg = Msg & " occurred: " & Err.Description
      ' Display message box with Stop sign icon and
      ' OK button.
      MsgBox Msg, vbCritical
      
   End Select
   Resume Next  ' Resume execution at same line that caused the error.

End Sub

Private Sub mnuAuthor_Click()
    lblAuthorFirstName.Visible = False
    lblAuthorLastName.Visible = False
    lblTitleofBook.Visible = False
    txtAuthorFirstName.Visible = False
    txtAuthorLastName.Visible = False
    txtTitleofBook.Visible = False
    lblTitle.Visible = False
    txtTitle.Visible = False
    lblLastName.Visible = True
    txtLastName.Visible = True
    cmdSearch.Visible = True
    cmdAddBook.Visible = False
    dbgBooks.Visible = False
    txtLastName.SetFocus
    cmdSearch.Default = True
    StatusBar1.SimpleText = "Click on Start Search button to find the Book."
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuPrint_Click()

    'Declare variables
    Dim strPrint As String
    Dim intLineCounter As Integer
    Dim intRow As Integer
    
    'error handler
    On Error GoTo ErrorHandler
    
    'Make sure there are books listed in the grid control
    If Not dbgBooks.Visible Then
        MsgBox "You must search by Author or Title, or click Show All Books before printing.", vbOKOnly, "Ooops!"
        Exit Sub
    End If
    
    'Print the information. I don't know how to set margins, so I tab to simulate it
    If strSearchType = "Author" Then    'Print Author and Title information
    
        'Print Column Headings
        Printer.Print vbTab & "Author Name" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Title of book"
        Printer.Print vbTab & "____________________________________________________________________________________________"
        Printer.Print " "
        intLineCounter = 3
        
        'step through rows and print each one
        With rst
            .MoveFirst
            For intRow = 0 To .RecordCount - 1
                strPrint = vbTab & .Fields("LastName") & ", " & .Fields("FirstName") _
                    & vbTab & vbTab & vbTab & vbTab & vbTab & .Fields("Title")
                Printer.Print strPrint
                .MoveNext
                intLineCounter = intLineCounter + 1
                If intLineCounter > 90 Then
                    Printer.NewPage
                    Printer.Print vbTab & "Author Name" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Title of book"
                    Printer.Print vbTab & "____________________________________________________________________________________________"
                    Printer.Print " "
                    intLineCounter = 3
                End If
            Next intRow
        End With
        Printer.EndDoc
        
    Else        'Print only title information
    
        'Print Column Headings
        Printer.Print vbTab & "Title of book"
        Printer.Print vbTab & "____________________________________________________________________________________________"
        Printer.Print " "
        intLineCounter = 3
        
        'step through rows and print each one
        With rst
            .MoveFirst
            For intRow = 0 To .RecordCount - 1
                strPrint = vbTab & .Fields("Title")
                Printer.Print strPrint
                .MoveNext
                intLineCounter = intLineCounter + 1
                If intLineCounter > 90 Then
                    Printer.NewPage
                    Printer.Print vbTab & "Title of book"
                    Printer.Print vbTab & "____________________________________________________________________________________________"
                    Printer.Print " "
                    intLineCounter = 3
                End If
            Next intRow
         End With
        Printer.EndDoc          'Send it to the printer
    End If
    
    Exit Sub

ErrorHandler:                   ' Error-handling routine.
    Select Case Err.Number      ' Evaluate error number.
       
       Case Else
         
      Msg = "Unexpected error #" & Str(Err.Number)
      Msg = Msg & " occurred: " & Err.Description
      ' Display message box with Stop sign icon and
      ' OK button.
      MsgBox Msg, vbCritical
      
   End Select
   Resume Next  ' Resume execution at same line that caused the error.

End Sub

Private Sub mnuTitle_Click()
    lblAuthorFirstName.Visible = False
    lblAuthorLastName.Visible = False
    lblTitleofBook.Visible = False
    txtAuthorFirstName.Visible = False
    txtAuthorLastName.Visible = False
    txtTitleofBook.Visible = False
    lblTitle.Visible = True
    txtTitle.Visible = True
    lblLastName.Visible = False
    txtLastName.Visible = False
    cmdSearch.Visible = True
    cmdAddBook.Visible = False
    dbgBooks.Visible = False
    txtTitle.SetFocus
    cmdSearch.Default = True
    StatusBar1.SimpleText = "Click on Start Search button to find the Book."
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button
        Case "Add New Book"
            mnuAdd_Click
        Case "Exit the Program"
            mnuExit_Click
        Case "Search by Author"
            mnuAuthor_Click
        Case "Search by Title"
            mnuTitle_Click
        Case "Print Book List"
            mnuPrint_Click
        Case "Show All Books"
            mnuAll_Click
    End Select
End Sub

