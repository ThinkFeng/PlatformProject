VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   10335
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   480
      TabIndex        =   13
      Top             =   5160
      Width           =   2295
   End
   Begin VB.CommandButton SearchPort 
      Caption         =   "�������д���"
      Height          =   300
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton OpenPort 
      Caption         =   "�򿪴���"
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   3840
      Width           =   1095
   End
   Begin VB.ComboBox Combo_stop 
      Height          =   300
      ItemData        =   "������.frx":0000
      Left            =   840
      List            =   "������.frx":000A
      TabIndex        =   10
      Text            =   "1"
      Top             =   3480
      Width           =   975
   End
   Begin VB.ComboBox Combo_data 
      Appearance      =   0  'Flat
      Height          =   300
      ItemData        =   "������.frx":0016
      Left            =   840
      List            =   "������.frx":0023
      TabIndex        =   9
      Text            =   "8"
      Top             =   3120
      Width           =   975
   End
   Begin VB.ComboBox Combo_check 
      Height          =   300
      ItemData        =   "������.frx":0033
      Left            =   840
      List            =   "������.frx":0040
      TabIndex        =   8
      Text            =   "NONE"
      Top             =   2760
      Width           =   975
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1200
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.ComboBox COM 
      Height          =   300
      ItemData        =   "������.frx":0054
      Left            =   840
      List            =   "������.frx":0056
      TabIndex        =   3
      Text            =   "COM"
      Top             =   840
      Width           =   975
   End
   Begin VB.Frame botelv 
      Caption         =   "�����ʣ�"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   1695
      Begin VB.OptionButton btl_l 
         Caption         =   "9600"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton btl_h 
         Caption         =   "115200"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Shape ShapeDisp 
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   240
      Shape           =   3  'Circle
      Top             =   3920
      Width           =   255
   End
   Begin VB.Label stop 
      AutoSize        =   -1  'True
      Caption         =   "ֹͣλ��"
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   720
   End
   Begin VB.Label data 
      AutoSize        =   -1  'True
      Caption         =   "����λ��"
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   720
   End
   Begin VB.Label check 
      AutoSize        =   -1  'True
      Caption         =   "У��λ��"
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   720
   End
   Begin VB.Label ck 
      AutoSize        =   -1  'True
      Caption         =   "���ڣ�"
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   915
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
        MSComm1.PortOpen = False
        ShapeDisp.FillColor = vbRed
        OpenPort.Caption = "�򿪴���"
End Sub

Private Sub Form_Load()
Call COM_Check
End Sub
Private Sub COM_Check()
  COM.Clear
  Dim i As Integer
  For i = 1 To 16
    MSComm1.CommPort = i
    On Error Resume Next
    '��̽�Ե�ȥ��
    MSComm1.PortOpen = True
     If Err.Number = 0 Then
         COM.AddItem "COM" & i
     End If
    MSComm1.PortOpen = False
    'ȷ��ÿһ����̽�Դ򿪺���رոô���
    Next i
    If COM.ListCount = 0 Then
      COM.Text = "No COM"
      COM.ForeColor = vbRed
    Else
      COM.ForeColor = vbBlack
      COM.Text = COM.List(0)
    End If
  End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ExitError
        MSComm1.PortOpen = False
ExitError:
End Sub

Private Sub OpenPort_Click()
Dim btl, DataValue, StopValue As Integer
Dim CheckString, MSCpro As String
On Error GoTo uerror
    If OpenPort.Caption = "�򿪴���" Then
        MSComm1.CommPort = Val(Right(COM.Text, 1))
        MSComm1.PortOpen = True '��Trueʱ�Ǵ�
        ShapeDisp.FillColor = vbGreen
        OpenPort.Caption = "�رմ���"
        If btl_h.Value = True Then btl = 115200
        If btl_l.Value = True Then btl = 9600
        MSCpro = Str(btl) & ",N" & Str(Combo_data.Text) & "," & Str(Combo_stop.Text)
        MSComm1.Settings = MSCpro
        MSComm1.RThreshold = 1
        '���ò���������
        COM.Enabled = False
        botelv.Enabled = False
        Combo_check.Enabled = False
        Combo_data.Enabled = False
        Combo_stop.Enabled = False
        SearchPort.Enabled = False
        '�������ϲ���������
    Else

        MSComm1.PortOpen = False
        ShapeDisp.FillColor = vbRed
        OpenPort.Caption = "�򿪴���"
        '���ò���������
        COM.Enabled = True
        botelv.Enabled = True
        Combo_check.Enabled = True
        Combo_data.Enabled = True
        Combo_stop.Enabled = True
        SearchPort.Enabled = True
        '�������ϲ���������
    End If
    Exit Sub
uerror:
       ShapeDisp.FillColor = vbRed
       OpenPort.Caption = "�򿪴���"
End Sub

Private Sub SearchPort_Click()
Call COM_Check
End Sub

