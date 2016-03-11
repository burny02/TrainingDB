Imports TemplateDB
Module Variables
    Public OverClass As OverClass
    Private Const TablePath As String = "M:\VOLUNTEER SCREENING SERVICES\Systems\Training\Backend.accdb"
    Private Const PWord As String = "RetroRetro*1"
    Private Const Connect2 As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & TablePath & ";Jet OLEDB:Database Password=" & PWord
    Private Const UserTable As String = "[Users]"
    Private Const UserField As String = "Username"
    Private Const LockTable As String = "[Locker]"
    Private Const AuditTable As String = "[Audit]"
    Private Contact As String = "Skye Firminger"
    Public Const SolutionName As String = "Training Tool"


    Public Sub StartUpCentral()

        OverClass = New OverClass
        OverClass.SetPrivate(UserTable,
                           UserField,
                           LockTable,
                           Contact,
                           Connect2,
                           AuditTable)

        OverClass.LockCheck()

        OverClass.LoginCheck()

        OverClass.AddAllDataItem(Form1)

        For Each ctl In OverClass.DataItemCollection
            If (TypeOf ctl Is ComboBox) Then
                Dim Com As ComboBox = ctl
                AddHandler Com.SelectionChangeCommitted, AddressOf GenericCombo
            End If
        Next
        For Each ctl In OverClass.DataItemCollection
            If (TypeOf ctl Is Button) Then
                Dim But As Button = ctl
                AddHandler But.Click, AddressOf ButtonSpecifics
            End If
        Next

    End Sub

End Module