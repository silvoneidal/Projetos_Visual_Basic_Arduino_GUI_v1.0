VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desconectado"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7335
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPort 
      Caption         =   "Port:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   240
      Width           =   735
   End
   Begin VB.ComboBox cboPort 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox cboBoard 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   240
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSketch 
      Caption         =   "Sketch"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "Upload"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   2
      Top             =   1200
      Width           =   3495
   End
   Begin VB.CommandButton cmdCompile 
      Caption         =   "Compile"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox txtSketch 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   5655
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   240
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   7
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   2880
   End
   Begin VB.TextBox txtConsole 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   2040
      Width           =   7095
   End
   Begin VB.Label Label2 
      Caption         =   "Board:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   735
   End
   Begin VB.Menu mMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mSend 
         Caption         =   "Send"
      End
      Begin VB.Menu mClear 
         Caption         =   "Clear"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Shell
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Sleep
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Variável global
Dim Port As String
Dim Board As String
Dim scan As Boolean

Private Sub Form_Load()
   Me.Caption = App.Title & "_v" & App.Major & "." & App.Minor & " by Dalçóquio Automação"
      
   cboBoard.AddItem "Arduino Uno"
   cboBoard.AddItem "Arduino Nano"
   cboBoard.AddItem "Arduino Mega"
   cboBoard.AddItem "Digispark (Default-16.5mhz)"
   cboBoard.AddItem "Generic ESP8266 Module"
   cboBoard.AddItem "NodeMCU 0.9 (ESP-12 Module)"
   cboBoard.AddItem "NodeMCU 1.0 (ESP-12E Module)"
   cboBoard.Text = cboBoard.List(0)
   
   cmdCompile.Enabled = False
   cmdUpload.Enabled = False
   txtSketch.Locked = True
   
   Call cmdPort_Click
   
   ' Verifica se existe a pasta "C:\arduino-cli
   Dim folderPath As String
   folderPath = "C:\arduino-cli"
   If Not Dir(folderPath, vbDirectory) <> "" Then
       MsgBox "Pasta " & folderPath & " não foi encontrada !!!", vbExclamation, "DALÇOQUIO AUTOMAÇÃO"
       End ' FECHA O APLICATIVO
   End If
   
End Sub

Private Sub scanPort()
   cboPort.Clear
   Dim i As Integer
   For i = 1 To 16 'Procura portas COM de 1 a 16
      MSComm1.CommPort = i
      On Error Resume Next 'ignora o tratamento de erro
      MSComm1.PortOpen = True 'tenta abrir a porta
      If Err.Number = 0 Then 'a porta está disponível
         cboPort.AddItem "COM" & i
         cboPort.ListIndex = 1
         MSComm1.PortOpen = False 'fecha a porta
      End If
      On Error GoTo 0 'ativa o tratamento de erro novamente
   Next i
   
   If cboPort.List(0) <> Empty Then cboPort.Text = cboPort.List(0)
   
   writeConsole ("Scan finalizado...")
   Beep
   
End Sub

Private Sub cmdPort_Click()
   writeConsole ("Scanning...")
   Call scanPort
   
End Sub

Private Sub cboPort_Click()
   Port = cboPort.Text
   
End Sub

Private Sub cboBoard_Click()
   If cboBoard.ListIndex = 0 Then Board = "arduino:avr:uno" ' Arduino Uno
   If cboBoard.ListIndex = 1 Then Board = "arduino:avr:nano" ' Arduino Nano
   If cboBoard.ListIndex = 2 Then Board = "arduino:avr:mega" ' arduino Mega
   If cboBoard.ListIndex = 2 Then Board = "digistump:avr:digispark-tiny" ' Digispark (Default-16.5mhz)
   If cboBoard.ListIndex = 2 Then Board = "esp8266:esp8266:generic" ' Generic ESP8266 Module
   If cboBoard.ListIndex = 2 Then Board = "esp8266:esp8266:nodemcu" ' NodeMCU 0.9 (ESP-12 Module)
   If cboBoard.ListIndex = 2 Then Board = "esp8266:esp8266:nodemcuv2" ' NodeMCU 1.0 (ESP-12E Module)
   
End Sub

Private Sub writeConsole(mensagem As String)
   txtConsole.Text = txtConsole.Text & "> " & mensagem & vbCrLf
   txtConsole.SelStart = Len(txtConsole.Text)
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   writeConsole ("Fechando o sistema...")
   DoEvents
   Sleep (1000)
   End

End Sub

Private Sub cmdSketch_Click()
    ' Define o filtro para exibir apenas arquivos de texto
    CommonDialog1.Filter = "Arquivos de Texto (*.ino)|*.ino"
    
    ' Abre o diálogo de seleção de arquivo
    CommonDialog1.ShowOpen
    
    ' Obtém o caminho completo do arquivo selecionado
    Dim filePath As String
    filePath = CommonDialog1.FileName
    
    ' Exibe o caminho do arquivo no TextBox
    txtSketch.Text = filePath
    txtSketch.ToolTipText = txtSketch.Text
    
End Sub

Private Sub cmdCompile_Click()
    ' Define o caminho do arduino-cli
    Dim arduinoCliPath As String
    arduinoCliPath = "C:\arduino-cli.exe" ' Substitua pelo caminho correto do arduino-cli em seu sistema

    ' Define o caminho do arquivo .ino
    Dim sketchPath As String
    sketchPath = txtSketch.Text  ' Substitua "seu_arquivo.ino" pelo nome correto do seu arquivo .ino

    ' Define o comando para compilar o arquivo .ino
    Dim compileCmd As String
    'compileCmd = "arduino-cli compile " & Board & " -v " & sketchPath " ' sem arquivo binário, com detalhe processo
     compileCmd = "arduino-cli compile --fqbn " & Board & " -v -e " & sketchPath ' com arquivo binário, com detalhe processo

    ' Executa o comando de compilação e abre a janela do prompt de comando
    ShellExecute Me.hwnd, "open", "cmd.exe", "/k " & compileCmd, vbNullString, vbNormalFocus
    
    writeConsole (compileCmd)
    writeConsole ("Compile iniciado...")
    writeConsole ("Finalizando, feche o prompt de comando")
    
End Sub

Private Sub cmdUpload_Click()
    ' Define o caminho do arduino-cli
    Dim arduinoCliPath As String
    arduinoCliPath = "arduino-cli.exe" ' Substitua pelo caminho correto do arduino-cli em seu sistema

    ' Define o caminho do arquivo .ino
    Dim sketchPath As String
    sketchPath = txtSketch.Text ' Substitua "seu_arquivo.ino" pelo nome correto do seu arquivo .ino

    ' Define o comando para fazer o upload do arquivo compilado para o Arduino Uno
    Dim uploadCmd As String
    uploadCmd = "arduino-cli upload -p " & Port & " --fqbn " & Board & " -v " & sketchPath
    
    ' Executa o comando de upload e abre a janela do prompt de comando
    ShellExecute Me.hwnd, "open", "cmd.exe", "/k " & uploadCmd, vbNullString, vbNormalFocus
    
    writeConsole (uploadCmd)
    writeConsole ("Upload iniciado...")
    writeConsole ("Finalizando, feche o prompt de comando")
    
End Sub

Private Sub Timer1_Timer()
   If cboBoard.Text = Empty Or cboPort.Text = Empty Or txtSketch.Text = Empty Then
      cmdCompile.Enabled = False
      cmdUpload.Enabled = False
   Else
      cmdCompile.Enabled = True
      cmdUpload.Enabled = True
   End If
   
End Sub


