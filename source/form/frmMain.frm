VERSION 5.00
Object = "{E5B0E85C-65F0-11D2-ACBA-0080ADA85544}#1.0#0"; "JSBBar16.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.3#0"; "Codejock.SuiteCtrls.Unicode.v10.3.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "LogisticERP_Caravan"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14190
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   14190
   StartUpPosition =   2  '屏幕中心
   Begin JSBtnBar16.ButtonBar BBLeft 
      Height          =   7395
      Left            =   0
      Top             =   360
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   13044
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GroupsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LargeImagesCount=   15
      LargeImageHeight=   32
      LargeImageWidth =   32
      LargeImagePicture1=   "frmMain.frx":0CCA
      LargeImageKey1  =   "OrderOutCompany"
      LargeImagePicture2=   "frmMain.frx":3D1C
      LargeImagePicture3=   "frmMain.frx":6D6E
      LargeImageKey3  =   "TransferCompany"
      LargeImagePicture4=   "frmMain.frx":9DC0
      LargeImagePicture5=   "frmMain.frx":CE12
      LargeImageKey5  =   "Employee"
      LargeImagePicture6=   "frmMain.frx":FE64
      LargeImageKey6  =   "OrderToCompany"
      LargeImagePicture7=   "frmMain.frx":12EB6
      LargeImagePicture8=   "frmMain.frx":15F08
      LargeImageKey8  =   "OrderInCompany"
      LargeImagePicture9=   "frmMain.frx":18F5A
      LargeImageKey9  =   "WarehouseOverView"
      LargeImagePicture10=   "frmMain.frx":1BFAC
      LargeImagePicture11=   "frmMain.frx":1EFFE
      LargeImageKey11 =   "OrderOutCompanyReceipt"
      LargeImagePicture12=   "frmMain.frx":22050
      LargeImageKey12 =   "Order"
      LargeImagePicture13=   "frmMain.frx":250A2
      LargeImageKey13 =   "PickupReceipt"
      LargeImagePicture14=   "frmMain.frx":280F4
      LargeImageKey14 =   "PackageDelivery"
      LargeImagePicture15=   "frmMain.frx":2B146
      LargeImageKey15 =   "Warehouse"
      SmallImageHeight=   16
      SmallImageWidth =   16
      GroupCount      =   2
      GroupCaption1   =   "功能"
      Group1ItemCount =   8
      Group1Item1Caption=   "取件单"
      Group1Item1Key  =   "PickUpReceipt"
      Group1Item1LargeIcon=   "PickUpReceipt"
      Group1Item2Caption=   "取件入库"
      Group1Item2Key  =   "OrdertoCompany"
      Group1Item2LargeIcon=   "OrdertoCompany"
      Group1Item3Caption=   "订单"
      Group1Item3Key  =   "Order"
      Group1Item3LargeIcon=   "Order"
      Group1Item4Caption=   "出库单"
      Group1Item4Key  =   "OutWarehouseReceipt"
      Group1Item4LargeIcon=   "OrderOutCompanyReceipt"
      Group1Item5Caption=   "快件出库"
      Group1Item5Key  =   "OutWarehouse"
      Group1Item5LargeIcon=   "OrderOutCompany"
      Group1Item6Caption=   "快件入库"
      Group1Item6Key  =   "InWarehouse"
      Group1Item6LargeIcon=   "OrderInCompany"
      Group1Item7Caption=   "快件派送"
      Group1Item7Key  =   "PackageDeliveryReceipt"
      Group1Item7LargeIcon=   "PackageDelivery"
      Group1Item8Caption=   "库存情况"
      Group1Item8Key  =   "WarehouseOverview"
      Group1Item8LargeIcon=   "WarehouseOverView"
      GroupCaption2   =   "基础信息"
      Group2ItemCount =   4
      Group2Item1Caption=   "员工信息"
      Group2Item1Key  =   "Employee"
      Group2Item1LargeIcon=   "Employee"
      Group2Item2Caption=   "客户信息"
      Group2Item2Key  =   "Cust"
      Group2Item2LargeIcon=   "Employee"
      Group2Item3Caption=   "物流公司"
      Group2Item3Key  =   "Transfer"
      Group2Item3LargeIcon=   "TransferCompany"
      Group2Item4Caption=   "仓库"
      Group2Item4Key  =   "Warehouse"
      Group2Item4LargeIcon=   "Warehouse"
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   3555
      Left            =   2460
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   2715
      _Version        =   655363
      _ExtentX        =   4789
      _ExtentY        =   6271
      _StockProps     =   64
      Appearance      =   6
      Color           =   16
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Sub BBLeft_ItemClick(ByVal Item As JSBtnBar16.JSGroupItem)
    
    Me.tcMain.Visible = True
    
    Select Case Item.key
    
        Case "Order"
            'Dim fOrder As frmOrder

            If Not isOpen("frmOrder") Then
            
                'Set fOrder = New frmOrder
                Load frmOrder
                
                Me.tcMain.InsertItem 0, "订单系统", frmOrder.hWnd, 0
                Call frmOrder.doSearch(0)
            Else
            
                Call SetActiveTab("订单系统")
            End If

        Case "PickUpReceipt"
            'Dim fOrder As frmOrder

            If Not isOpen("frmPickupReceipt") Then
            
                'Set fOrder = New frmOrder
                Load frmPickupReceipt
            
                Me.tcMain.InsertItem 0, "取件单", frmPickupReceipt.hWnd, 0
                Call frmPickupReceipt.doSearch(0)
            Else
            
                Call SetActiveTab("取件单")
            
            End If

        Case "Cust"
            'Dim fOrder As frmOrder

            If Not isOpen("frmCust") Then
            
                'Set fOrder = New frmOrder
                Load frmCust
            
                Me.tcMain.InsertItem 0, "客户信息", frmCust.hWnd, 0
                Call frmCust.doSearch(0)
                
            Else
                Call SetActiveTab("客户信息")
            End If

        Case "Transfer"
            'Dim fOrder As frmOrder

            If Not isOpen("frmTransfer") Then
            
                'Set fOrder = New frmOrder
                Load frmTransfer
            
                Me.tcMain.InsertItem 0, "物流公司", frmTransfer.hWnd, 0
                Call frmTransfer.doSearch(0)
            Else
                Call SetActiveTab("物流公司")
            End If
            
        Case "Warehouse"
            'Dim fOrder As frmOrder

            If Not isOpen("frmWarehouse") Then
            
                'Set fOrder = New frmOrder
                Load frmWarehouse
            
                Me.tcMain.InsertItem 0, "仓库信息", frmWarehouse.hWnd, 0
                Call frmWarehouse.doSearch(0)
            Else
                Call SetActiveTab("仓库信息")
            End If
        
        Case "Employee"
            'Dim fOrder As frmOrder

            If Not isOpen("frmEmployee") Then
            
                'Set fOrder = New frmOrder
                Load frmEmployee
            
                Me.tcMain.InsertItem 0, "员工信息", frmEmployee.hWnd, 0
                Call frmEmployee.doSearch(0)
            Else
                Call SetActiveTab("员工信息")
            End If

        Case "OrdertoCompany"
            'Dim fOrder As frmOrder

            If Not isOpen("frmOrdertoCompany") Then
            
                'Set fOrder = New frmOrder
                Load frmOrdertoCompany
            
                Me.tcMain.InsertItem 0, "取件入库", frmOrdertoCompany.hWnd, 0
                'Call frmOrderInWarehouse.doSearch(0)
            Else
                Call SetActiveTab("取件入库")
            End If
            
        Case "OutWarehouseReceipt"
            'Dim fOrder As frmOrder

            If Not isOpen("frmOutWarehouseReceipt") Then
            
                'Set fOrder = New frmOrder
                Load frmOutWarehouseReceipt
            
                Me.tcMain.InsertItem 0, "出库单", frmOutWarehouseReceipt.hWnd, 0
                Call frmOutWarehouseReceipt.doSearch(0)
            Else
                Call SetActiveTab("出库单")
            End If

        Case "OutWarehouse"
            'Dim fOrder As frmOrder

            If Not isOpen("frmOrderOutWarehouse") Then
            
                'Set fOrder = New frmOrder
                Load frmOrderOutWarehouse
                frmOrderOutWarehouse.Caption = ""
                
                Me.tcMain.InsertItem 0, "快件出库", frmOrderOutWarehouse.hWnd, 0
                'Call frmOrderInWarehouse.doSearch(0)
            Else
                Call SetActiveTab("快件出库")
            End If

        Case "InWarehouse"
            'Dim fOrder As frmOrder

            If Not isOpen("frmOrderInWarehouse") Then
            
                'Set fOrder = New frmOrder
                Load frmOrderInWarehouse
                
                Me.tcMain.InsertItem 0, "快件入库", frmOrderInWarehouse.hWnd, 0
                frmOrderInWarehouse.Caption = ""
                'Call frmOrderInWarehouse.doSearch(0)
            Else
                Call SetActiveTab("快件入库")
            End If
            
        Case "PackageDeliveryReceipt"
            'Dim fOrder As frmOrder

            If Not isOpen("frmPackageDeliveryReceipt") Then
            
                'Set fOrder = New frmOrder
                Load frmPackageDeliveryReceipt
                frmPackageDeliveryReceipt.Caption = ""
                Me.tcMain.InsertItem 0, "快件派送", frmPackageDeliveryReceipt.hWnd, 0
                Call frmPackageDeliveryReceipt.doSearch(0)
            Else
                Call SetActiveTab("快件派送")
            End If
            
        Case "WarehouseOverview"
            'Dim fOrder As frmOrder

            If Not isOpen("frmWarehouseOverview") Then
            
                'Set fOrder = New frmOrder
                Load frmWarehouseOverview
                frmWarehouseOverview.Caption = ""
                Me.tcMain.InsertItem 0, "库存概况", frmWarehouseOverview.hWnd, 0
                Call frmWarehouseOverview.doSearch(0)
            Else
                Call SetActiveTab("库存概况")
            End If
    End Select

End Sub

Public Sub SetActiveTab(ByVal strCaption As String)

    Dim i As Integer

    For i = 0 To Me.tcMain.ItemCount - 1

        If Me.tcMain.Item(i).Caption = strCaption Then
            Me.tcMain.SelectedItem = i
        End If

    Next

End Sub

Private Sub Form_Load()
    Call Form_Resize

    'Load frmNaviLeft
    'SetParent frmNaviLeft.hWnd, frmMain.hWnd
    'frmNaviLeft.Show
    'Call movefrmLeft(mWidth, mHeight, mLeft)
End Sub

'Private Sub movefrmLeft(mWidth, mHeight, mLeft)
'    frmNaviLeft.width = LEFT_WIDTH
'
'    frmNaviLeft.Height = mHeight - LEFT_MARGINTOP - LEFT_MARGINBOTTOM
'
'    frmNaviLeft.Left = -mLeft - SUBFORM_OFFSETLEFT + LEFT_MARGINLEFT
'    frmNaviLeft.Top = LEFT_MARGINTOP '+ SUBFORM_OFFSETTOP
'End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call CloseAllForms
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Height < 10500 Then
        Me.Height = 10500
    End If

    If Me.width < 18000 Then
        Me.width = 18000
    End If

    gHeight = Me.Height
    gWidth = Me.width
    gLeft = Me.Left
    gTop = Me.Top
    Me.BBLeft.Move 50, 50, 1000, Me.Height - 700
    Me.tcMain.Move 1200, 50, Me.width - 1500, Me.Height - 700
End Sub
Function isOpen(fName As String) As Boolean
    Dim F As Form

    For Each F In Forms

        If F.name = fName Then
            isOpen = True
            Exit For
        End If

    Next

End Function

Public Function CloseTab(ByVal strTabName As String) As Boolean
    Dim i As Integer
    On Error Resume Next

    For i = 0 To Me.tcMain.ItemCount - 1

        If Me.tcMain.Item(i).Caption = strTabName Then
            Call Me.tcMain.RemoveItem(i)
        End If

    Next

    If Me.tcMain.ItemCount = 0 Then
        Me.tcMain.Visible = False
    End If

End Function
