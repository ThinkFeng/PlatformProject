VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7335
   ClientLeft      =   6945
   ClientTop       =   2175
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   10335
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   1920
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   6615
      Left            =   3360
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   480
      Width           =   3135
   End
   Begin VB.CommandButton OpenPort 
      Caption         =   "打开串口"
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   3840
      Width           =   1200
   End
   Begin VB.ComboBox Combo_stop 
      Height          =   300
      ItemData        =   "主窗体.frx":0000
      Left            =   840
      List            =   "主窗体.frx":0002
      TabIndex        =   10
      Text            =   "1"
      Top             =   3480
      Width           =   975
   End
   Begin VB.ComboBox Combo_data 
      Appearance      =   0  'Flat
      Height          =   300
      ItemData        =   "主窗体.frx":0004
      Left            =   840
      List            =   "主窗体.frx":0006
      TabIndex        =   9
      Text            =   "8"
      Top             =   3120
      Width           =   975
   End
   Begin VB.ComboBox Combo_check 
      Height          =   300
      ItemData        =   "主窗体.frx":0008
      Left            =   840
      List            =   "主窗体.frx":000A
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
      ItemData        =   "主窗体.frx":000C
      Left            =   840
      List            =   "主窗体.frx":000E
      TabIndex        =   3
      Text            =   "COM"
      Top             =   840
      Width           =   975
   End
   Begin VB.Frame botelv 
      Caption         =   "波特率："
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   1695
      Begin VB.OptionButton btl_l 
         Caption         =   "9600"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton btl_h 
         Caption         =   "115200"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   2
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
      Caption         =   "停止位："
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   720
   End
   Begin VB.Label data 
      AutoSize        =   -1  'True
      Caption         =   "数据位："
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   720
   End
   Begin VB.Label check 
      AutoSize        =   -1  'True
      Caption         =   "校验位："
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   720
   End
   Begin VB.Label ck 
      AutoSize        =   -1  'True
      Caption         =   "串口："
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
Dim HData As String
Public Sub SysInfo1_DeviceArrival(ByVal DeviceType As Long, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long) '检测即插即用设备的插入
Call COM_Check
End Sub
Private Sub SysInfo1_DeviceRemoveComplete(ByVal DeviceType As Long, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long) '检测即插即用设备的拔出
Call COM_Check
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
    '试探性地去打开
    MSComm1.PortOpen = True
     If Err.Number = 0 Then
         COM.AddItem "COM" & i
     End If
    MSComm1.PortOpen = False
    '确保每一次试探性打开后均关闭该串口
    Next i
    If COM.ListCount = 0 Then
      COM.Text = "No COM"
      COM.ForeColor = vbRed
    Else
      COM.ForeColor = vbBlack
      COM.Text = COM.List(0)
    End If
  End Sub
Private Sub MSComm1_OnComm()
On Error Resume Next
Dim BytesReceived() As Byte
Dim buffer As String
Dim i As Integer
Select Case MSComm1.CommEvent
  Case comEvReceive '接收十六进制数据。并以十六进制显示
      'MSComm.InputLen = 0
      buffer = MSComm1.Input '接收数据至字符串中
      BytesReceived() = buffer '将数据转入Byte数组中
      For i = 0 To UBound(BytesReceived) '显示结果以十六进制显示
         If Len(Hex(BytesReceived(i))) = 1 Then
            HData = HData & "0" & Hex(BytesReceived(i))
         Else
            HData = HData & Hex(BytesReceived(i))
         End If
      Next
     Text1.Text = Text1.Text & HData
     HData = ""
            '最后将结果后入Text1中
     MSComm.OutBufferCount = 0 '清除发送缓冲区
     MSComm.InBufferCount = 0 '清除接收缓冲区
End Select
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
    If OpenPort.Caption = "打开串口" Then
        MSComm1.CommPort = Val(Right(COM.Text, 1))
        ShapeDisp.FillColor = vbGreen
        OpenPort.Caption = "关闭串口"
        If btl_h.Value = True Then btl = 115200
        If btl_l.Value = True Then btl = 9600
        MSCpro = Str(btl) & ",N" & Str(Combo_data.Text) & "," & Str(Combo_stop.Text)
        MSComm1.Settings = MSCpro
        MSComm1.RThreshold = 1
        MSComm1.InputMode = comInputModeBinary   '以二进制方式取回传入的数据
        'MSComm1.Handshaking = 0
        MSComm1.PortOpen = True '当True时是打开
        'MSComm1.InputLen = 0 '表示一次读取整个缓冲区,这样读取之后,缓冲区就自动被清空了
        '禁用参数设置项
        COM.Enabled = False
        botelv.Enabled = False
        Combo_check.Enabled = False
        Combo_data.Enabled = False
        Combo_stop.Enabled = False
        '禁用以上参数设置项
    Else

        MSComm1.PortOpen = False
        ShapeDisp.FillColor = vbRed
        OpenPort.Caption = "打开串口"
        '启用参数设置项
        COM.Enabled = True
        botelv.Enabled = True
        Combo_check.Enabled = True
        Combo_data.Enabled = True
        Combo_stop.Enabled = True
        '启用以上参数设置项
    End If
    Exit Sub
uerror:
       ShapeDisp.FillColor = vbRed
       OpenPort.Caption = "打开串口"
End Sub

