Option Compare Database
Option Explicit

Public Sub ExportTableDataWithSchema(strTableName As String)
' Ref: https://msdn.microsoft.com/en-us/library/office/ff193212.aspx
' expression .ExportXML(ObjectType, DataSource, DataTarget, SchemaTarget, PresentationTarget, ImageTarget, Encoding, OtherFlags, WhereCondition, AdditionalData)
' expression A variable that represents an Application object.
'
' Usage Example: ExportTableDataWithSchema "tblDummy3"

    Application.ExportXML ObjectType:=acExportTable, DataSource:=strTableName, DataTarget:=strTableName & ".xml", _
        SchemaTarget:=strTableName & ".sch", Encoding:=acUTF8, OtherFlags:=acExportAllTableAndFieldProperties
    Debug.Print "DONE !!!"

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