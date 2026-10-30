Attribute VB_Name = "ControlInfo"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private m_Left As Single
Private m_Top As Single
Private m_Width As Single
Private m_Height As Single
Private m_FontSize As Single
Private m_IsMultiPage As Boolean
Private m_PageCount As Long

Public Property Get IsMultiPage() As Boolean
    IsMultiPage = m_IsMultiPage
End Property
Public Property Let IsMultiPage(Value As Boolean)
    m_IsMultiPage = Value
End Property

Public Property Get PageCount() As Long
    PageCount = m_PageCount
End Property
Public Property Let PageCount(Value As Long)
    m_PageCount = Value
End Property
Public Property Get Left() As Single
    Left = m_Left
End Property
Public Property Let Left(Value As Single)
    m_Left = Value
End Property

Public Property Get tOP() As Single
    tOP = m_Top
End Property
Public Property Let tOP(Value As Single)
    m_Top = Value
End Property

Public Property Get Width() As Single
    Width = m_Width
End Property
Public Property Let Width(Value As Single)
    m_Width = Value
End Property

Public Property Get Height() As Single
    Height = m_Height
End Property
Public Property Let Height(Value As Single)
    m_Height = Value
End Property

Public Property Get FontSize() As Single
    FontSize = m_FontSize
End Property
Public Property Let FontSize(Value As Single)
    m_FontSize = Value
End Property

Public Sub InitFromControl(ctrl As Control)
    m_Left = ctrl.Left
    m_Top = ctrl.tOP
    m_Width = ctrl.Width
    m_Height = ctrl.Height
    On Error Resume Next
    m_FontSize = ctrl.Font.Size
    On Error GoTo 0
End Sub
