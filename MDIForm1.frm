VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   8505
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   14055
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu DASHBOARD 
      Caption         =   "DASHBOARD"
   End
   Begin VB.Menu ENTRY 
      Caption         =   "ENTRY"
      Begin VB.Menu FARMER_ENTRY 
         Caption         =   "FARMER ENTRY"
      End
      Begin VB.Menu CUSTOMER_ENTRY 
         Caption         =   "CUSTOMER ENTRY"
      End
      Begin VB.Menu PRODUCT_ENTRY 
         Caption         =   "PRODUCT ENTRY"
      End
      Begin VB.Menu CENTRE_ENTRY 
         Caption         =   "CENTRE ENTRY"
      End
   End
   Begin VB.Menu DAIRY 
      Caption         =   "DAIRY"
      Begin VB.Menu DAIRY_ENTRY 
         Caption         =   "DAIRY ENTRY"
      End
      Begin VB.Menu DAIRY_PRODUCT 
         Caption         =   "DAIRY PRODUCT"
      End
      Begin VB.Menu DAIRY_SALE 
         Caption         =   "DAIRY SALE"
      End
   End
   Begin VB.Menu COLLECTION 
      Caption         =   "COLLECTION"
      Begin VB.Menu MILK_COLLECTION 
         Caption         =   "MILK COLLECTION"
      End
      Begin VB.Menu MILK_COLLECTION_VIEW 
         Caption         =   "MILK COLLECTION VIEW"
      End
      Begin VB.Menu FAT_LIST 
         Caption         =   "FAT LIST"
      End
   End
   Begin VB.Menu ORDER 
      Caption         =   "ORDER"
      Begin VB.Menu PURCHASE_ORDER 
         Caption         =   "PURCHASE ORDER"
      End
      Begin VB.Menu CUSTOMER_ORDER 
         Caption         =   "CUSTOMER ORDER"
      End
   End
   Begin VB.Menu TRANSACTION 
      Caption         =   "TRANSACTION"
      Begin VB.Menu SALES 
         Caption         =   "SALES"
         Begin VB.Menu SALE_BILL 
            Caption         =   "SALE BILL"
         End
         Begin VB.Menu SALE_BILL_VIEW 
            Caption         =   "SALE BILL VIEW"
         End
      End
      Begin VB.Menu PURCHASE 
         Caption         =   "PURCHASE"
         Begin VB.Menu PURCHASE_BILL 
            Caption         =   "PURCHASE BILL"
         End
         Begin VB.Menu PURCHASE_VIEW 
            Caption         =   "PURCHASE VIEW"
         End
      End
      Begin VB.Menu RETURN 
         Caption         =   "RETURN"
         Begin VB.Menu SALES_RETURN 
            Caption         =   "SALES RETURN"
         End
         Begin VB.Menu SALES_RETURN_VIEW 
            Caption         =   "SALES RETURN VIEW"
         End
         Begin VB.Menu PURCHASE_RETURN 
            Caption         =   "PURCHASE RETURN"
         End
         Begin VB.Menu PURCHASE_RETURN_VIEW 
            Caption         =   "PURCHASE RETURN VIEW"
         End
      End
      Begin VB.Menu STOCK 
         Caption         =   "STOCK"
      End
   End
   Begin VB.Menu PAYMENT 
      Caption         =   "PAYMENT"
      Begin VB.Menu FARMER_PAYMENT 
         Caption         =   "FARMER PAYMENT"
      End
      Begin VB.Menu FARMER_PAYMENT_VIEW 
         Caption         =   "FARMER PAYMENT VIEW"
      End
      Begin VB.Menu DAIRY_PAYMENT 
         Caption         =   "DAIRY PAYMENT"
      End
      Begin VB.Menu DAIRY_PAYMENT_VIEW 
         Caption         =   "DAIRY PAYMENT VIEW"
      End
   End
   Begin VB.Menu DUES 
      Caption         =   "DUES"
      Begin VB.Menu CUSTOMER_DUES 
         Caption         =   "CUSTOMER DUES"
      End
      Begin VB.Menu CENTRE_DUES 
         Caption         =   "CENTRE DUES"
      End
   End
   Begin VB.Menu REPORT 
      Caption         =   "REPORT"
      Begin VB.Menu PRODUCT_REPORT 
         Caption         =   "PRODUCT REPORT"
      End
      Begin VB.Menu CUSTOMER_REPORT 
         Caption         =   "CUSTOMER REPORT"
      End
      Begin VB.Menu FARMER_REPORT 
         Caption         =   "FARMER REPORT"
      End
      Begin VB.Menu STOCK_REPORT 
         Caption         =   "STOCK REPORT"
      End
   End
   Begin VB.Menu LOGOUT 
      Caption         =   "LOGOUT"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CENTRE_DUES_Click()
CentreDues.Show
End Sub

Private Sub CENTRE_ENTRY_Click()
CollectionCentre.Show
End Sub

Private Sub CUSTOMER_DUES_Click()
CustomerDues.Show
End Sub

Private Sub CUSTOMER_ENTRY_Click()
CustomerEntry.Show
End Sub

Private Sub CUSTOMER_ORDER_Click()
CustomerOrder.Show
End Sub

Private Sub CUSTOMER_REPORT_Click()
DataReport1.Show
End Sub

Private Sub DAIRY_ENTRY_Click()
DairyEntry.Show
End Sub

Private Sub DAIRY_PAYMENT_Click()
DairyPayment.Show
End Sub

Private Sub DAIRY_PAYMENT_VIEW_Click()
DairyPaymentView.Show
End Sub

Private Sub DAIRY_PRODUCT_Click()
DairyProduct.Show
End Sub

Private Sub DAIRY_SALE_Click()
DairySale.Show
End Sub

Private Sub DASHBOARD_Click()
home.Show
End Sub

Private Sub FARMER_ENTRY_Click()
FarmerEntry.Show
End Sub

Private Sub FARMER_PAYMENT_Click()
FarmerPayment.Show
End Sub

Private Sub FARMER_PAYMENT_VIEW_Click()
FarmerPaymentView.Show
End Sub

Private Sub FARMER_REPORT_Click()
DataReport3.Show
End Sub

Private Sub FAT_LIST_Click()
FatList.Show
End Sub

Private Sub LOGOUT_Click()
End
LoginForm.Show
End Sub

Private Sub MILK_COLLECTION_Click()
MilkCollection.Show
End Sub

Private Sub MILK_COLLECTION_VIEW_Click()
MilkCollectionView.Show
End Sub

Private Sub PRODUCT_ENTRY_Click()
Product.Show
End Sub

Private Sub PRODUCT_REPORT_Click()
ProductReport1.Show
End Sub

Private Sub PURCHASE_BILL_Click()
PurchaseBill.Show
End Sub

Private Sub PURCHASE_ORDER_Click()
PurchaseOrder.Show
End Sub

Private Sub PURCHASE_RETURN_Click()
PurchaseReturn.Show
End Sub

Private Sub PURCHASE_RETURN_VIEW_Click()
PurchaseReturnView.Show
End Sub

Private Sub PURCHASE_VIEW_Click()
Centrepurchaseview.Show
End Sub



Private Sub SALE_BILL_Click()
Salebill.Show
End Sub

Private Sub SALE_BILL_VIEW_Click()
SalesView.Show
End Sub

Private Sub SALES_RETURN_Click()
SalesReturn.Show
End Sub

Private Sub SALES_RETURN_VIEW_Click()
SalesReturnView.Show
End Sub

Private Sub STOCK_Click()
StockView.Show
End Sub

Private Sub STOCK_REPORT_Click()
StockReport4.Show
End Sub
