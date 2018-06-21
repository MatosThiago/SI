VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSeparar 
   Caption         =   "Separar planilha conforme coluna específica"
   ClientHeight    =   1950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5730
   OleObjectBlob   =   "frmSeparar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSeparar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGerar_Click()
    Dim lRefColunas As Range
    Dim lRefBase    As Range
    
    Set lRefColunas = Range(frmSeparar.refEditColunas.Text)
    Set lRefBase = Range(frmSeparar.refEditBase.Text)
    
    lPasta = SelectFolder
    lsPreparar lRefColunas, lRefBase
    lsSeparar lRefColunas, lRefBase
    
    Unload Me
End Sub
