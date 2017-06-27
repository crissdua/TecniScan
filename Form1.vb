Imports SAPbobsCOM
Imports SAPbouiCOM
Imports System.Windows.Forms
Imports System.IO
Imports System.Xml

Public Class Form1

    Dim myStream As Stream
    Private Property _oCompany As SAPbobsCOM.Company

    Public Property oCompany() As SAPbobsCOM.Company
        Get
            Return _oCompany
        End Get
        Set(ByVal value As SAPbobsCOM.Company)
            _oCompany = value
        End Set
    End Property
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' MakeConnectionSAP()
    End Sub
    Public Sub IngresaFactura()
        Dim oReturn As Integer = -1
        Dim oError As Integer = 0
        Dim errMsg As String = ""
        If MakeConnectionSAP() Then
            Dim Invoice As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            Invoice.CardCode = "400001"
            Invoice.DocTotal = 10
            Invoice.DocType = BoDocumentTypes.dDocument_Service
            Invoice.Lines.ItemDescription = "prueba"
            Invoice.Lines.TaxCode = "EXE"
            Invoice.Lines.AccountCode = "_SYS00000000004"
            Invoice.Lines.LineTotal = 10
            oReturn = Invoice.Add()
            If oReturn <> 0 Then
                oCompany.GetLastError(oError, errMsg)
                MsgBox(errMsg)
            End If
        End If
    End Sub
    Public Function MakeConnectionSAP() As Boolean
        Dim Connected As Boolean = False
        '' Dim cnnParam As New Settings
        Try
            Connected = -1

            oCompany = New SAPbobsCOM.Company
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014
            oCompany.DbUserName = "sa"
            oCompany.DbPassword = "12345"
            oCompany.Server = "DESKTOP-13FMJTF"
            oCompany.CompanyDB = "FYA"
            oCompany.UserName = "manager"
            oCompany.Password = "alegria"
            oCompany.LicenseServer = "DESKTOP-13FMJTF:30000"
            oCompany.UseTrusted = False
            Connected = oCompany.Connect()

            If Connected <> 0 Then
                Connected = False
                MsgBox(oCompany.GetLastErrorDescription)
            Else
                'MsgBox("Conexión con SAP exitosa")
                Connected = True
            End If
            Return Connected
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        IngresaFactura()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim openFileDialog1 As OpenFileDialog = New OpenFileDialog()
        openFileDialog1.InitialDirectory = "c:\\"
        openFileDialog1.Filter = "xml files (*.xml)|"
        openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = True

        Dim oReturn As Integer = -1
        Dim oError As Integer = 0
        Dim errMsg As String = ""


        If (openFileDialog1.ShowDialog() = DialogResult.OK) Then
            Try
                myStream = openFileDialog1.OpenFile()

                Using (myStream)
                        Dim file As StreamWriter = New StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\Log.txt", True)
                        Dim archivo = openFileDialog1.SafeFileName.ToString()
                        Dim Xml As XmlDocument = New XmlDocument()
                        Xml.Load(myStream)
                        file.WriteLine("Archivo Cargado: " + archivo)
                    Dim ArticleList As XmlNodeList = Xml.SelectNodes("invoice/document")
                    For Each xnDoc As XmlNode In ArticleList
                        If MakeConnectionSAP() Then
                            Dim Invoice As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

                            Invoice.Series = xnDoc.SelectSingleNode("series").InnerText
                            Invoice.DocNum = xnDoc.SelectSingleNode("docnum").InnerText
                            Invoice.DocDate = xnDoc.SelectSingleNode("docdate").InnerText
                            Invoice.CardCode = xnDoc.SelectSingleNode("cardcode").InnerText
                            Invoice.DocTotal = xnDoc.SelectSingleNode("doctotal").InnerText
                            Invoice.DocType = BoDocumentTypes.dDocument_Items

                            Dim CatNodesList As XmlNodeList = xnDoc.SelectNodes("document_lines/line")
                            For Each xnDet As XmlNode In CatNodesList
                                Invoice.Lines.ItemCode = xnDet.SelectSingleNode("itemcode").InnerText
                                Invoice.Lines.Quantity = xnDet.SelectSingleNode("quantity").InnerText
                                Invoice.Lines.TaxCode = xnDet.SelectSingleNode("taxcode").InnerText
                                Invoice.Lines.LineTotal = Convert.ToDouble(xnDet.SelectSingleNode("linetotal").InnerText)
                                Invoice.Lines.Add()
                            Next
                            Invoice.Comments = "pruebas"
                            oReturn = Invoice.Add()
                            If oReturn <> 0 Then
                                oCompany.GetLastError(oError, errMsg)
                                MsgBox(errMsg)
                            End If
                        End If
                    Next

                End Using

            Catch ex As Exception

            End Try

        End If
    End Sub
End Class
