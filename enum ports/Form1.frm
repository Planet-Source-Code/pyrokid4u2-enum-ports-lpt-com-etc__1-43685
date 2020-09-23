VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2370
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0002
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   2370
      ItemData        =   "Form1.frx":0004
      Left            =   1200
      List            =   "Form1.frx":0006
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      Height          =   2370
      ItemData        =   "Form1.frx":0008
      Left            =   2640
      List            =   "Form1.frx":000A
      TabIndex        =   4
      Top             =   1200
      Width           =   2775
   End
   Begin VB.ListBox List4 
      Appearance      =   0  'Flat
      Height          =   2370
      ItemData        =   "Form1.frx":000C
      Left            =   5400
      List            =   "Form1.frx":000E
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ListBox List5 
      Appearance      =   0  'Flat
      Height          =   2370
      ItemData        =   "Form1.frx":0010
      Left            =   6840
      List            =   "Form1.frx":0012
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Get Ports"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reserved"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   11
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Moniter Name"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Port Type"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Port Name"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enum Ports"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function enumports Lib "winspool.drv" Alias "EnumPortsA" (ByVal pName As String, ByVal Level As Long, ByVal lpbPorts As Long, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long

Private Type PORT_INFO_2
    pPortName As String
    pMonitorName As String
    pDescription As String
    fPortType As Long
    Reserved As Long
End Type
Private Type API_PORT_INFO_2
    pPortName As Long
    pMonitorName As Long
    pDescription As Long
    fPortType As Long
    Reserved As Long
End Type

Dim Ports(0 To 100) As PORT_INFO_2

Public Function CutString(strName As String) As String
    'Finds a null then trims the string
    Dim x As Integer
    x = InStr(strName, vbNullChar)
    If x > 0 Then CutString = Left(strName, x - 1) Else CutString = strName
End Function

Public Function LPSTRtoSTRING(ByVal lngPointer As Long) As String
    Dim lngLength As Long
    'number of characters
    lngLength = lstrlenW(lngPointer) * 2
    'Initialize the string
    LPSTRtoSTRING = String(lngLength, 0)
    'Copy the string
    CopyMem ByVal StrPtr(LPSTRtoSTRING), ByVal lngPointer, lngLength
    'Convert to Unicode
    LPSTRtoSTRING = CutString(StrConv(LPSTRtoSTRING, vbUnicode))
End Function

'You can specify a server name (example //WIN2KWKSTN) to get the ports of that machine
Public Function getports(ServerName As String) As Long
    Dim ret As Long
    Dim PortsStruct(0 To 100) As API_PORT_INFO_2
    Dim pcbNeeded As Long
    Dim pcReturned As Long
    Dim tmpbuffer As Long
    Dim i As Integer
    'determine amount of bytes needed
    ret = enumports(ServerName, 2, tmpbuffer, 0, pcbNeeded, pcReturned)
    'use api to allocate the buffer
    tmpbuffer = HeapAlloc(GetProcessHeap(), 0, pcbNeeded)
    ret = enumports(ServerName, 2, tmpbuffer, pcbNeeded, pcbNeeded, pcReturned)
    If ret Then
        'convert string pointer value to vb-readable value
        CopyMem PortsStruct(0), ByVal tmpbuffer, pcbNeeded
        For i = 0 To pcReturned - 1
            Ports(i).pDescription = LPSTRtoSTRING(PortsStruct(i).pDescription)
            Ports(i).pPortName = LPSTRtoSTRING(PortsStruct(i).pPortName)
            Ports(i).pMonitorName = LPSTRtoSTRING(PortsStruct(i).pMonitorName)
            Ports(i).fPortType = PortsStruct(i).fPortType
        Next
    End If
    getports = pcReturned
    If tmpbuffer Then HeapFree GetProcessHeap(), 0, tmpbuffer
End Function

Private Sub Command1_Click()
    Dim NumPorts As Long
    Dim i As Integer
    Dim item As ListItem
    NumPorts = getports("")
    For i = 0 To NumPorts - 1
        With Ports(i)
        List1.AddItem .pPortName
        List2.AddItem .fPortType
        List3.AddItem .pDescription
        List4.AddItem .pMonitorName
        List5.AddItem .Reserved
        End With
    Next
End Sub

