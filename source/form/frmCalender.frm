VERSION 5.00
Object = "{82392BA0-C18D-11D2-B0EA-00A024695830}#1.0#0"; "ticaldr6.ocx"
Begin VB.Form frmCalender 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   Icon            =   "frmCalender.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "选  择"
      Height          =   420
      Left            =   720
      TabIndex        =   1
      Top             =   2970
      Width           =   1320
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关  闭"
      Height          =   420
      Left            =   2385
      TabIndex        =   0
      Top             =   2970
      Width           =   1320
   End
   Begin TDBCalendar6Ctl.TDBCalendar tdCld 
      Height          =   2115
      Left            =   360
      TabIndex        =   2
      Top             =   180
      Width           =   3495
      _Version        =   65536
      _ExtentX        =   6165
      _ExtentY        =   3731
      ShowContextMenu =   -1  'True
      Appearance      =   0
      AutoSize        =   0   'False
      BorderStyle     =   1
      BackColor       =   -2147483643
      StartOfMonth    =   1
      EmptyRows       =   0
      Enabled         =   -1  'True
      FirstMonth      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LineColors0     =   -2147483632
      LineStyles0     =   0
      LineColors1     =   -2147483632
      LineStyles1     =   0
      LineColors2     =   -2147483632
      LineStyles2     =   0
      LineColors3     =   -2147483632
      LineStyles3     =   0
      LineColors4     =   -2147483632
      LineStyles4     =   0
      LineColors5     =   -2147483632
      LineStyles5     =   0
      LineColors6     =   -2147483632
      LineStyles6     =   2
      MarginBottom    =   0
      MarginTitle     =   0
      MarginTop       =   0
      MarginLeft      =   0
      MarginRight     =   0
      MarginWidth     =   0
      MarginHeight    =   0
      MaxDate         =   5373484
      MinDate         =   2456659
      MousePointer    =   0
      YearType        =   1
      MonthRows       =   1
      MonthCols       =   1
      MultiSelect     =   0
      NavOrientation  =   3
      ScrollRate      =   1
      ScrollTipAlign  =   3
      SelEdgeWidth    =   8
      SelectStyle     =   1
      SelectWhat      =   0
      ShowMenu        =   -1  'True
      ShowNavigator   =   2
      ShowScrollTip   =   -1  'True
      ShowTrailing    =   -1  'True
      StartOfWeek     =   1
      Templates       =   0
      TipInterval     =   500
      TitleHeight     =   0
      TitleFormat     =   "yyyy - mm"
      ValueIsNull     =   0   'False
      Value           =   2456885
      OverrideTipText =   ""
      TopDate         =   2456871
      AttribStyles    =   "frmCalender.frx":000C
      StyleSets       =   "frmCalender.frx":00CC
      CtrlType        =   8
      CtrlValue       =   "CtrlStyle"
      DayType         =   8
      DayValue        =   "DayStyle"
      TitleType       =   8
      TitleValue      =   "TitleStyle"
      WeekType        =   8
      WeekValue       =   "WeekStyle"
      TrailType       =   8
      TrailValue      =   "TrailAttrib"
      SelType         =   8
      SelValue        =   "SelAttrib"
      WeekRests0      =   0
      WeekReflect0    =   0
      WeekCaption0    =   "日"
      WeekAttrib0Type =   8
      WeekAttrib0Value=   "SunAttrib"
      WeekRests1      =   0
      WeekReflect1    =   0
      WeekCaption1    =   "一"
      WeekAttrib1Type =   1
      WeekRests2      =   0
      WeekReflect2    =   0
      WeekCaption2    =   "二"
      WeekAttrib2Type =   1
      WeekRests3      =   0
      WeekReflect3    =   0
      WeekCaption3    =   "三"
      WeekAttrib3Type =   1
      WeekRests4      =   0
      WeekReflect4    =   0
      WeekCaption4    =   "四"
      WeekAttrib4Type =   1
      WeekRests5      =   0
      WeekReflect5    =   0
      WeekCaption5    =   "五"
      WeekAttrib5Type =   1
      WeekRests6      =   0
      WeekReflect6    =   0
      WeekCaption6    =   "六"
      WeekAttrib6Type =   8
      WeekAttrib6Value=   "SatAttrib"
      HolidayStyles   =   "frmCalender.frx":022C
      UserStyles      =   ""
      Key             =   "frmCalender.frx":0248
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
End
Attribute VB_Name = "frmCalender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SelectedDate As String


Public CallerControl As VB.Control
Public ValueReturnControl As VB.Control
Public MinDate As String
Public MaxDate As String
Public SelectDate As String
'直接用公共变量来做外部接口，似乎太不优雅了，但是也没办法，谁让这样快呢？

Public Property Get ReturnDate() As Variant
    ReturnDate = Me.tdCld.Value
End Property

Private Sub cmdClose_Click()
    Call Unload(Me)
    
End Sub

Private Sub cmdOK_Click()
    Call tdCld_DateDblClick
End Sub

Private Sub Form_Load()

    Me.ZOrder 0
    If IsDate(MinDate) Then
        Me.tdCld.MinDate = MinDate
    End If
    If IsDate(MaxDate) Then
        Me.tdCld.MaxDate = MaxDate
    End If
    
    If IsDate(SelectDate) Then
        Me.tdCld.Value = SelectDate
    End If
End Sub

Private Sub Form_LostFocus()
    Call Unload(Me)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.tdCld.Top = 0
    Me.tdCld.Left = 0
    Me.tdCld.Height = Me.Height - Me.cmdClose.Height - 100
    Me.tdCld.width = Me.width
End Sub


Private Sub tdCld_DateDblClick()
    SelectedDate = Me.tdCld.Value

    Select Case TypeName(ValueReturnControl)

        Case "CommandButton"
            ValueReturnControl.Caption = Me.tdCld.Value

        Case "TextBox"
            ValueReturnControl.Text = Me.tdCld.Value

    End Select

    Call Unload(Me)
End Sub
