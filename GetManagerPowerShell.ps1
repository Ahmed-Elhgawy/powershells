Sub GetManagersFromAD()

    Dim objAD AS Object
    Dim objUser AS Object
    Dim managerDN AS String
    Dim managerName AS String
    Dim i AS Long

    i = 2

    Do While Cells(i, 1).value <> ""

        On Error Resume Next

        Set objUser = GetObject("LDAP://CN=" & Cells(i, 1).Value & ",CN=Users,DC=HDBANK,DC=local")

        If Err.Number <> 0 Then
            Cells(i, 2).value = "user not found"
            Err.Clear
        Else
           managerDN = objUser.Get("manager")

           If managerDN <> "" Then
               Set objAD = GetObject("LDAP://" & managerDN)
               managerName = objAD.SamAccountName
               Cells(i, 2).value = managerName
           Else
               Cells(i, 2).value = "No Manager"
           End If
        End If

        i = i + 1

    Loop

    MsgBox "Done!"
End Sub