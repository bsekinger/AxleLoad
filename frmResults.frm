VERSION 5.00
Begin VB.Form frmResults 
   Caption         =   "Results"
   ClientHeight    =   5568
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   ScaleHeight     =   5568
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResults 
      Height          =   5232
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6732
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub ShowResults(sText As String)
    
    txtResults.Text = sText
    Me.Show

End Sub


