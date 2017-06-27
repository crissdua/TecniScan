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
        'Factura()
        'Cliente()
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


    Private Sub Factura()
        Dim oReturn As Integer = -1
        Dim oError As Integer = 0
        Dim errMsg As String = ""
        Try
            Dim entra As String = "C:\TS\Factura\invoice.xml"
            Dim sale As String = "C:\TS\Factura\temp\invoice" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
            Dim Xml As XmlDocument = New XmlDocument()
            Xml.Load(entra)
            Dim log As StreamWriter = New StreamWriter("C:\TS\Factura\log\Log.txt", True)
            log.WriteLine("Archivo Cargado: " + entra)
            Dim ArticleList As XmlNodeList = Xml.SelectNodes("invoice/document")
            For Each xnDoc As XmlNode In ArticleList
                If MakeConnectionSAP() Then
                    Dim Invoice As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

                    Invoice.Series = xnDoc.SelectSingleNode("series").InnerText
                    Invoice.DocNum = xnDoc.SelectSingleNode("docnum").InnerText
                    Invoice.DocDate = xnDoc.SelectSingleNode("docdate").InnerText
                    Invoice.CardCode = xnDoc.SelectSingleNode("cardcode").InnerText
                    Invoice.DocTotal = Convert.ToDouble(xnDoc.SelectSingleNode("doctotal").InnerText)
                    Dim tipo As Char = xnDoc.SelectSingleNode("doctype").InnerText

                    If tipo = "S" Then
                        Invoice.DocType = BoDocumentTypes.dDocument_Service


                        Dim CatNodesList As XmlNodeList = xnDoc.SelectNodes("document_lines/line")
                        For Each xnDet As XmlNode In CatNodesList
                            Invoice.Lines.ItemDescription = xnDet.SelectSingleNode("itemdescription").InnerText
                            Invoice.Lines.TaxCode = xnDet.SelectSingleNode("taxcode").InnerText
                            Invoice.Lines.AccountCode = xnDet.SelectSingleNode("accountcode").InnerText
                            Invoice.Lines.LineTotal = Convert.ToDouble(xnDet.SelectSingleNode("linetotal").InnerText)
                            Invoice.Lines.Add()
                        Next
                    Else

                        Invoice.DocType = BoDocumentTypes.dDocument_Items

                        Dim CatNodesList As XmlNodeList = xnDoc.SelectNodes("document_lines/line")
                        For Each xnDet As XmlNode In CatNodesList
                            Invoice.Lines.ItemCode = xnDet.SelectSingleNode("itemcode").InnerText
                            Invoice.Lines.Quantity = xnDet.SelectSingleNode("quantity").InnerText
                            Invoice.Lines.TaxCode = xnDet.SelectSingleNode("taxcode").InnerText
                            Invoice.Lines.LineTotal = Convert.ToDouble(xnDet.SelectSingleNode("linetotal").InnerText)
                            Invoice.Lines.Add()
                        Next
                    End If
                    Invoice.Comments = "pruebas"
                    oReturn = Invoice.Add()
                    If oReturn <> 0 Then
                        oCompany.GetLastError(oError, errMsg)
                        MsgBox(errMsg)
                    End If
                End If
            Next
            File.Move(entra, sale)
            log.WriteLine("Archivo en Temporal: " + sale)
            log.Close()

        Catch ex As Exception
            Dim log As StreamWriter = New StreamWriter("C:\TS\Factura\log\Log.txt", True)
            Dim sale As String = "C:\TS\Factura\temp\invoice" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
            log.WriteLine("Ocurrio un Error en : " + sale + " Error:" + ex.ToString())
            log.Close()
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Cliente()
        Dim oReturn As Integer = -1
        Dim oError As Integer = 0
        Dim errMsg As String = ""
        Try
            Dim entra As String = "C:\TS\Cliente\cliente.xml"
            Dim sale As String = "C:\TS\Cliente\temp\cliente" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
            Dim Xml As XmlDocument = New XmlDocument()
            Xml.Load(entra)
            Dim log As StreamWriter = New StreamWriter("C:\TS\Cliente\log\Log.txt", True)
            log.WriteLine("Archivo Cargado: " + entra)
            Dim ArticleList As XmlNodeList = Xml.SelectNodes("cliente/document")
            For Each xnDoc As XmlNode In ArticleList
                If MakeConnectionSAP() Then
                    Dim oBusinessPartners As SAPbobsCOM.BusinessPartners = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                    oBusinessPartners.CardCode = xnDoc.SelectSingleNode("cardcode").InnerText
                    oBusinessPartners.CardName = xnDoc.SelectSingleNode("cardname").InnerText
                    oBusinessPartners.GroupCode = xnDoc.SelectSingleNode("groupnumber").InnerText
                    oBusinessPartners.Currency = xnDoc.SelectSingleNode("doccurrency").InnerText
                    oBusinessPartners.FederalTaxID = "000000000000" 'xnDoc.SelectSingleNode("federaltaxid").InnerText

                    'If tipo = "S" Then


                    '    Dim CatNodesList As XmlNodeList = xnDoc.SelectNodes("document_lines/line")
                    '    For Each xnDet As XmlNode In CatNodesList
                    '        oBusinessPartners.Lines.ItemDescription = xnDet.SelectSingleNode("itemdescription").InnerText
                    '        oBusinessPartners.Lines.TaxCode = xnDet.SelectSingleNode("taxcode").InnerText
                    '        oBusinessPartners.Lines.AccountCode = xnDet.SelectSingleNode("accountcode").InnerText
                    '        oBusinessPartners.Lines.LineTotal = Convert.ToDouble(xnDet.SelectSingleNode("linetotal").InnerText)
                    '        oBusinessPartners.Lines.Add()
                    '    Next
                    'Else

                    '    oBusinessPartners.DocType = BoDocumentTypes.dDocument_Items

                    '    Dim CatNodesList As XmlNodeList = xnDoc.SelectNodes("document_lines/line")
                    '    For Each xnDet As XmlNode In CatNodesList
                    '        oBusinessPartners.Lines.ItemCode = xnDet.SelectSingleNode("itemcode").InnerText
                    '        oBusinessPartners.Lines.Quantity = xnDet.SelectSingleNode("quantity").InnerText
                    '        oBusinessPartners.Lines.TaxCode = xnDet.SelectSingleNode("taxcode").InnerText
                    '        oBusinessPartners.Lines.LineTotal = Convert.ToDouble(xnDet.SelectSingleNode("linetotal").InnerText)
                    '        oBusinessPartners.Lines.Add()
                    '    Next
                    'End If
                    'oBusinessPartners.Comments = "pruebas"
                    oReturn = oBusinessPartners.Add()
                    If oReturn <> 0 Then
                        oCompany.GetLastError(oError, errMsg)
                        MsgBox(errMsg)
                    End If
                End If
            Next
            File.Move(entra, sale)
            log.WriteLine("Archivo en Temporal: " + sale)
            log.Close()

        Catch ex As Exception
            Dim log As StreamWriter = New StreamWriter("C:\TS\Cliente\log\Log.txt", True)
            Dim sale As String = "C:\TS\Cliente\temp\cliente" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml"
            log.WriteLine("Ocurrio un Error en : " + sale + " Error:" + ex.ToString())
            log.Close()
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Cliente()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Factura()
    End Sub
End Class
