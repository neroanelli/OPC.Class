VERSION 5.00
Begin VB.Form fmMrep 
   Caption         =   "月报表"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   5160
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command2 
      Caption         =   "导出……"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "月表"
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      Begin VB.CommandButton Command1 
         Caption         =   "预览"
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "MonthReport.frx":0000
         Left            =   1200
         List            =   "MonthReport.frx":0028
         TabIndex        =   1
         Text            =   "1月"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "月份："
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
   End
End
Attribute VB_Name = "fmMrep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
