VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cl_final"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private aTicker As String
Private aYearly_Change As Double
Private aPercent_Change As Double
Private aTotal_Volume As Double
Private aDate_Open As Long
Private aDate_Close As Long
Private aOpen_Stock As Double
Private aClose_Stock As Double
' Properties
Property Get Ticker() As String
    Ticker = aTicker
End Property

Property Let Ticker(dTicker As String)
    aTicker = dTicker
End Property

Property Get Yearly_Change() As Double
    Yearly_Change = aYearly_Change
End Property

Property Let Yearly_Change(dYearly_Change As Double)
    aYearly_Change = dYearly_Change
End Property

Property Get Percent_Change() As Double
    Percent_Change = aPercent_Change
End Property

Property Let Percent_Change(dPercent_Change As Double)
    aPercent_Change = dPercent_Change
End Property
Property Get Total_Volume() As Double
    Total_Volume = aTotal_Volume
End Property

Property Let Total_Volume(dTotal_Volume As Double)
    aTotal_Volume = dTotal_Volume
End Property

Property Get Date_Open() As Long
    Date_Open = aDate_Open
End Property

Property Let Date_Open(dDate_Open As Long)
    aDate_Open = dDate_Open
End Property

Property Get Date_Close() As Long
    Date_Close = aDate_Close
End Property

Property Let Date_Close(dDate_Close As Long)
    aDate_Close = dDate_Close
End Property

Property Get Open_Stock() As Double
    Open_Stock = aOpen_Stock
End Property

Property Let Open_Stock(dOpen_Stock As Double)
    aOpen_Stock = dOpen_Stock
End Property

Property Get Close_Stock() As Double
    Close_Stock = aClose_Stock
End Property

Property Let Close_Stock(dClose_Stock As Double)
    aClose_Stock = dClose_Stock
End Property
' Event - triggered when class created


' Methods
'

