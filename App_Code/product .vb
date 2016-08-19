Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient

Public Class product_

    Public Function getProducts() As DataSet
        Dim conn As New SqlConnection(Classconn.strconnsql)
        Dim adapter As New SqlDataAdapter("SELECT [ProductID], [ProductName], [SupplierID], [CategoryID], [QuantityPerUnit], [UnitPrice] FROM [Products]", conn)
        Dim ds As New DataSet
        adapter.Fill(ds, "Products")
        Return ds
    End Function
    Public Sub updateProducts(ByVal ProductID As Integer, _
                          ByVal ProductName As String, _
                          ByVal SupplierID As Integer, _
                          ByVal CategoryID As Integer, _
                          ByVal QuantityPerUnit As String, _
                          ByVal UnitPrice As Double)
        Dim conn As New SqlConnection(Classconn.strconnsql)
        Dim adapter As New SqlDataAdapter("SELECT * FROM Products WHERE ProductID=" & ProductID, conn)
        Dim ds As New DataSet
        adapter.Fill(ds, "Products")
        With ds.Tables(0).Rows(0)
            .Item("ProductName") = ProductName
            .Item("SupplierID") = SupplierID
            .Item("CategoryID") = CategoryID
            .Item("QuantityPerUnit") = QuantityPerUnit
            .Item("UnitPrice") = UnitPrice
        End With
        Dim cb As New SqlCommandBuilder(adapter)
        adapter.Update(ds, "Products")


    End Sub



End Class
