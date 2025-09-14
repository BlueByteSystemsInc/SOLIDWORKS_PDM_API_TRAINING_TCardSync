Imports EPDM.Interop.epdm

Public Class SetupPage


    Public Sub LoadData(ByRef poCmd As EdmCmd)

        Dim props As IEdmTaskProperties
        props = poCmd.mpoExtra

        'Populate the edit box from a variable
        TextBox1.Text = props.GetValEx(AddIn.STORAGEKEY)

    End Sub

    Public Sub StoreData(ByRef poCmd As EdmCmd)

        Dim props As IEdmTaskProperties
        props = poCmd.mpoExtra

        props.SetValEx(AddIn.STORAGEKEY, TextBox1.Text)

    End Sub
End Class
