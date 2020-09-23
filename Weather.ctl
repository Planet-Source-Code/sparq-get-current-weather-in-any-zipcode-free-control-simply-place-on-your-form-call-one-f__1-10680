VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.UserControl JaysWeather 
   CanGetFocus     =   0   'False
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Weather.ctx":0000
   ScaleHeight     =   2535
   ScaleWidth      =   90
   ToolboxBitmap   =   "Weather.ctx":0674
   Begin InetCtlsObjects.Inet Inet 
      Left            =   1080
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "JaysWeather"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function GetTemp(ZipCode As String)
    Dim Text As String
    Dim Search As String
    Dim Spot As Integer
    Dim Spot2 As Integer
    
    Search = "<FONT FACE=""Arial, Helvetica, Chicago, Sans Serif"" SIZE=3><B>"
    
    Text = Inet.OpenURL("http://www.weather.com/weather/us/zips/" & ZipCode & ".html")
    Spot = InStr(1, Text, Search) + Len(Search)
    Spot2 = InStr(Spot, Text, "</B>")
    GetTemp = Mid$(Text, Spot, Spot2 - Spot)
    
End Function
Private Sub UserControl_Resize()
    Width = 945
    Height = 1065
End Sub
