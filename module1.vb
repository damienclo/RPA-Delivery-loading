Sub test()

    Dim mainDir As String
    mainDir = Application.ActiveWorkbook.Path & "\"
    
    Main mainDir & "\ODOO.XLS" _
    , mainDir & "SAP.XLSX" _
    , mainDir & "test.xlsx" _
    , mainDir & "report.xlsx"

End Sub

Sub Main(odooPath As String, sapPath As String, exportPath As String, reportPath As String)
    Application.DisplayAlerts = False
    
    'General Variables
    Dim exportWB As Workbook
    Dim reportWB As Workbook
    Dim odooWB As Workbook
    Dim sapWB As Workbook
    
    Dim exportWS As Worksheet
    Dim reportWS As Worksheet
    Dim odooWS As Worksheet
    Dim sapWS As Worksheet
    
    'Exceptions
    Dim ex1 As String
    Dim ex2 As String
    
    ex1 = "Quantité manquante sur Odoo pour permettre la réception."
    ex2 = "Article non présent dans la commande Odoo ou Numéro de commande SAP non lisible."
    
    'Prepare files, create and open them
    Set exportWB = Workbooks.Add
    Set reportWB = Workbooks.Add
    Set odooWB = Workbooks.Open(odooPath)
    Set sapWB = Workbooks.Open(sapPath)

    
    Set exportWS = exportWB.Worksheets(1)
    Set reportWS = reportWB.Worksheets(1)
    Set odooWS = odooWB.Sheets(1)
    
    Set sapWS = sapWB.Sheets(1)
    
    'Add headers for the EXPORT
    exportWS.Range("A1").value = "ID Externe"
    exportWS.Range("B1").value = "Mouvements de stock non colisé/ID"
    exportWS.Range("C1").value = "Mouvements de stock non colisé/Quantité traitée"
    exportWS.Range("D1").value = "Info Reference du transfert Odoo"
    exportWS.Range("E1").value = "Info Commande SAP"
    exportWS.Range("F1").value = "Info Article"
    exportWS.Range("G1").value = "Info Rapprochement"
    
    'Add headers for the REPORT
    reportWS.Range("A1").value = "Référence"
    reportWS.Range("B1").value = "Article"
    reportWS.Range("C1").value = "Anomalie"
    reportWS.Range("D1").value = "Commande SAP"
    reportWS.Range("E1").value = "Quantité à valider en livraison sur Odoo"
    
    
        
    reportWB.Sheets.Add
    Set reportWS = reportWB.Sheets(1)
    
    reportWS.Range("A1").value = "Référence"
    
    Set reportWS = reportWB.Sheets(2)
        
    Dim Nb_Of_Rows_SAP As Integer
    Dim Nb_Of_Rows_ODOO As Integer
    Dim CountSAP As Integer
    Dim CountODOO As Integer
    Dim MaxDrop As Integer '###############IDEA 3
    Dim Row As Integer
    Dim refID As String
    Dim extID As String
    
    CountSAP = 2
    CountODOO = 2
    Row = 2
    
    '#################################
    '## Chargement des données ODOO ##
    '#################################
    
    Nb_Of_Rows_SAP = sapWS.Cells(Rows.Count, 3).End(xlUp).Row
    odooWB.Activate
    Nb_Of_Rows_ODOO = odooWS.Cells(Rows.Count, 5).End(xlUp).Row
    
    
    While (CountODOO <= Nb_Of_Rows_ODOO)

        If (odooWS.Range("A" & CountODOO).value <> "") Then
            extID = odooWS.Range("A" & CountODOO).value
        End If
        
        If (odooWS.Range("C" & CountODOO).value <> "") Then
            refID = odooWS.Range("C" & CountODOO).value
        End If
        
        
        exportWS.Range("A" & CountODOO).value = extID
        exportWS.Range("B" & CountODOO).value = odooWS.Range("L" & CountODOO).value
        exportWS.Range("C" & CountODOO).value = odooWS.Range("H" & CountODOO).value
        exportWS.Range("D" & CountODOO).value = refID
        exportWS.Range("E" & CountODOO).value = odooWS.Range("J" & CountODOO).value
        exportWS.Range("F" & CountODOO).value = odooWS.Range("D" & CountODOO).value
        exportWS.Range("G" & CountODOO).value = "Non"
        
        CountODOO = CountODOO + 1
    Wend
    
    Row = 2
    CountODOO = 2
    foundPosition = 2
    Dim found As Boolean
    Dim missingQt As Boolean
    Dim countREPORT As Integer
    countREPORT = 2
    found = False
    
    
    '#########################################
    '## Vérification des données ODOO       ##
    '## avec Création du rapport d'erreur   ##
    '#########################################
    
    While (CountSAP < Nb_Of_Rows_SAP)
        While (CountODOO <= Nb_Of_Rows_ODOO And ((Not found) Or (found And missingQt = False)))
            If (CheckForErrors(sapWS.Range("D" & CountSAP).value) <> -1) Then
            
                If (InStr(1, "" & exportWS.Range("E" & CountODOO).value, sapWS.Range("C" & CountSAP).value) > 0) _
                And (exportWS.Range("F" & CountODOO).value = CLng(sapWS.Range("D" & CountSAP).value)) _
                And sapWS.Range("G" & CountSAP).value <> "" Then
                
                    'Maximum déposable = Quantité initiale demandée - Quantité déjà ajoutée en reception
                    MaxDrop = odooWS.Range("G" & CountODOO).value - exportWS.Range("C" & CountODOO).value
                    
                    'Si la quantité déposable max est plus grande que la quantité livrée sur SAP, on ajoute toutes les pièces de SAP
                    If (MaxDrop >= sapWS.Range("G" & CountSAP).value) _
                    And (MaxDrop <> 0) Then
                        exportWS.Range("C" & CountODOO).value = exportWS.Range("C" & CountODOO).value + sapWS.Range("G" & CountSAP).value
                        sapWS.Range("G" & CountSAP).value = 0
                        exportWS.Range("G" & CountODOO).value = "Reception"
                    
                    'Si la quantité déposable max est plus petite que la quantité livrée sur SAP, on ajoute qu'une partie des pièces de SAP
                    ElseIf (MaxDrop < sapWS.Range("G" & CountSAP).value) _
                    And (MaxDrop <> 0) Then
                        exportWS.Range("C" & CountODOO).value = exportWS.Range("C" & CountODOO).value + MaxDrop
                        sapWS.Range("G" & CountSAP).value = sapWS.Range("G" & CountSAP).value - MaxDrop
                        exportWS.Range("G" & CountODOO).value = "Reception"
                    End If
                    
                    found = True
                    foundPosition = CountODOO
                End If
            End If
            CountODOO = CountODOO + 1
        Wend
        
        If (found = False) Then
            '#########################################
            '## Mis à jour du rapport pour ex2      ##
            '#########################################
            reportWS.Range("A" & countREPORT).value = "Non trouvée"
            reportWS.Range("B" & countREPORT).value = sapWS.Range("D" & CountSAP).value
            reportWS.Range("C" & countREPORT).value = ex2
            reportWS.Range("D" & countREPORT).value = sapWS.Range("C" & CountSAP).value
            reportWS.Range("E" & countREPORT).value = sapWS.Range("G" & CountSAP).value
            countREPORT = countREPORT + 1
        ElseIf (sapWS.Range("G" & CountSAP).value > 0) Then
            '#########################################
            '## Mis à jour du rapport pour ex1      ##
            '#########################################
            reportWS.Range("A" & countREPORT).value = exportWS.Range("D" & foundPosition).value
            reportWS.Range("B" & countREPORT).value = exportWS.Range("F" & foundPosition).value
            reportWS.Range("C" & countREPORT).value = ex1 'quantité manquante
            reportWS.Range("D" & countREPORT).value = sapWS.Range("C" & CountSAP).value
            reportWS.Range("E" & countREPORT).value = sapWS.Range("G" & CountSAP).value
            countREPORT = countREPORT + 1
        End If
        
        
        found = False
        missingQt = False
        CountODOO = 2
        CountSAP = CountSAP + 1
    Wend
    Debug.Print CountSAP
    Dim unique As String
    
    CountODOO = 2
    countREPORT = 2
    unique = ""
    
    '###############################################
    '## Transferts (réceptions) à valider en Odoo  #
    '###############################################
    
    Set reportWS = reportWB.Sheets(1)

    While (CountODOO <= Nb_Of_Rows_ODOO)
        If (exportWS.Range("D" & CountODOO).value <> "") And (exportWS.Range("D" & CountODOO).value <> unique) Then
            If (exportWS.Range("G" & CountODOO).value = "Reception") Then
                reportWS.Range("A" & countREPORT).value = exportWS.Range("D" & CountODOO).value
                unique = exportWS.Range("D" & CountODOO).value
                countREPORT = countREPORT + 1
            End If
        End If
        CountODOO = CountODOO + 1
    Wend
    
    '###############################################
    '## Nétoyage des lignes avec 0 mouvements      #
    '###############################################
    
    On Error Resume Next
    exportWS.Range("$C$1:$C$" & Nb_Of_Rows_ODOO).AutoFilter Field:=1, Criteria1:="0"
    exportWS.Range("$C$2:$C$" & Nb_Of_Rows_ODOO).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    exportWS.AutoFilterMode = False
    
    exportWS.Range("$G$1:$G$" & Nb_Of_Rows_ODOO).AutoFilter Field:=1, Criteria1:="Non"
    exportWS.Range("$G$2:$G$" & Nb_Of_Rows_ODOO).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    exportWS.AutoFilterMode = False

    exportWB.SaveAs exportPath
    reportWB.SaveAs reportPath
    exportWB.Close False
    reportWB.Close False
    odooWB.Close False
    sapWB.Close False
    
    Application.DisplayAlerts = True
End Sub

Function CheckForErrors(value As String) As Long

    On Error GoTo handling
    
    CheckForErrors = CLng(value)
    GoTo clean
 
handling:
    CheckForErrors = -1
clean:
End Function