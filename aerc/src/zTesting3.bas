Option Compare Database
Option Explicit

Public Sub ExportAllTableDataWithSchema()

    Dim tdf As DAO.TableDef
    Dim strExportPath As String
    strExportPath = "C:\ae\aegit\aerc\src\exp\"
    
    For Each tdf In CurrentDb.TableDefs
        If Not (Left$(tdf.Name, 4) = "MSys" _
            Or Left$(tdf.Name, 4) = "~TMP" _
            Or Left$(tdf.Name, 3) = "zzz") Then
            
            ExportTableDataWithSchema strExportPath, tdf.Name
        End If
    Next tdf
    Debug.Print "DONE !!!"
End Sub

Private Sub ExportTableDataWithSchema(expPath As String, strTableName As String)
    ' Ref: https://msdn.microsoft.com/en-us/library/office/ff193212.aspx
    ' expression .ExportXML(ObjectType, DataSource, DataTarget, SchemaTarget, PresentationTarget, ImageTarget, Encoding, OtherFlags, WhereCondition, AdditionalData)
    ' expression - A variable that represents an Application object.
    '
    ' Usage Example: ExportTableDataWithSchema "tblDummy3"

    Application.ExportXML ObjectType:=acExportTable, DataSource:=strTableName, DataTarget:=expPath & strTableName & ".xml", _
        SchemaTarget:=strTableName & ".xsd", Encoding:=acUTF8, OtherFlags:=acExportAllTableAndFieldProperties
    Debug.Print strTableName

End Sub

Public Sub ExportCustomerOrderData()
    ' Ref: https://msdn.microsoft.com/en-us/library/office/ff193212.aspx

    Dim objOrderInfo As AdditionalData
    Dim objOrderDetailsInfo As AdditionalData

    Set objOrderInfo = Application.CreateAdditionalData

    ' Add the Orders and Order Details tables to the data to be exported.
    Set objOrderDetailsInfo = objOrderInfo.Add("Orders")
    objOrderDetailsInfo.Add "Order Details"

    ' Export the contents of the Customers table. The Orders and Order
    ' Details tables will be included in the XML file.
    Application.ExportXML ObjectType:=acExportTable, DataSource:="Customers", _
        DataTarget:="Customer Orders.xml", AdditionalData:=objOrderInfo
    Debug.Print "DONE !!!"

End Sub