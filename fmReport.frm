VERSION 5.00
Object = "{5C8CED40-8909-11D0-9483-00A0C91110ED}#1.0#0"; "MSDATREP.OCX"
Begin VB.Form fmReport 
   Caption         =   "报表"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   10875
   StartUpPosition =   3  '窗口缺省
   Begin MSDataRepeaterLib.DataRepeater DataRepeater1 
      Height          =   6000
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   10583
      _StreamID       =   -1412567295
      _Version        =   393216
      Caption         =   "DataRepeater1"
      BeginProperty RepeatedControlName {21FC0FC0-1E5C-11D1-A327-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
      EndProperty
   End
End
Attribute VB_Name = "fmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
