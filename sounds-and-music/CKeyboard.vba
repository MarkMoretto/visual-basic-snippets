Option Explicit
''' Keyboard class

Private pKeyNumber As Long
Private pKeyName As String
Private pKeyAltName As String
Private pKeyColor As String

''' KeyNumber property
Public Property Get KeyNumber() As Long
    KeyNumber = pKeyNumber
End Property

Public Property Let KeyNumber(Value As Long)
    pKeyNumber = Value
End Property

''' KeyName property
Public Property Get KeyName() As String
    KeyName = pKeyName
End Property

Public Property Let KeyName(Value As String)
    pKeyName = Value
End Property

''' KeyAltName property
Public Property Get KeyAltName() As String
    KeyAltName = pKeyAltName
End Property

Public Property Let KeyAltName(Value As String)
    pKeyAltName = Value
End Property

''' KeyColor property
Public Property Get KeyColor() As String
    KeyColor = pKeyColor
End Property

Public Property Let KeyColor(Value As String)
    pKeyColor = Value
End Property

Public Function New_Keyboard()
' Allow for extensibility
    Set New_Keyboard = New CKeyboard
End Function
