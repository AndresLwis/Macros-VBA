Attribute VB_Name = "Módulo2"
Private Sub Concatenar()

n = InputBox("Digite el número de registros")

For i = 2 To n

    If Cells(i, 1) <> "" And Cells(i, 2) <> "" Then
               
        For k = 1 To 10
            If Cells(i + k, 1) = "" And Cells(i + k, 2) <> "" Then
             contar = contar + 1
            Else
            k = 100
            End If
        Next
        
        For j = 1 To contar
            If Cells(i + j, 1) = "" Then
                
            Cells(i, 2) = Cells(i, 2) & Chr(10) & Cells(i + j, 2)
            Cells(i + j, 2) = ""
            Else
            j = contar + 1
            
            End If
        Next
        


    End If


Next

End Sub
