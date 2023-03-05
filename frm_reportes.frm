VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_reportes 
   Caption         =   "UserForm1"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9345.001
   OleObjectBlob   =   "frm_reportes.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frm_reportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_genreporte_Click()
    Call Auxiliar.generar_reporte
End Sub

Private Sub UserForm_Initialize()
    Me.BackColor = RGB(255, 153, 0)
    
    lbl_titulo.BackColor = RGB(255, 153, 0)
    lbl_pais.BackColor = RGB(255, 153, 0)
    lbl_ncliente.BackColor = RGB(255, 153, 0)
    
    lbl_titulo.ForeColor = RGB(255, 255, 255)
    lbl_pais.ForeColor = RGB(255, 255, 255)
    lbl_ncliente.ForeColor = RGB(255, 255, 255)
    
    'Llenar combos
    
    Call Auxiliar.llena_combo(cbo_pais, "Seleccionar país", 1, Hoja3)
    Call Auxiliar.llena_combo(cbo_ncliente, "Seleccionar nombre del cliente", 2, Hoja3)

End Sub

