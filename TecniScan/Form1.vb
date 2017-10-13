Imports SAPbobsCOM
Imports SAPbouiCOM
Imports System.Windows.Forms
Imports System.IO
Imports System.Xml
Imports System.Threading
Imports System.Globalization

Public Class Form1
    Dim MyThread As Thread
    Dim FacturaStart As New ThreadStart(AddressOf BackgroundFactura)
    Dim CallFactura As New MethodInvoker(AddressOf Me.FacturaToma)
    Dim ClienteStart As New ThreadStart(AddressOf BackgroundCliente)
    Dim CallCliente As New MethodInvoker(AddressOf Me.ClienteToma)
    Dim myStream As Stream
    Dim count As Integer
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
        'MakeConnectionSAP()
        'oCompany = New SAPbobsCOM.Company
        'Dim cliente As String = oCompany.UserName
        'MessageBox.Show("Usuario sap" + cliente)
        'Try
        '    MyThread = New Thread(FacturaStart)
        '    MyThread.IsBackground = True
        '    MyThread.Name = "MyThreadFactura"
        '    MyThread.Start()
        '    MyThread = New Thread(ClienteStart)
        '    MyThread.IsBackground = True
        '    MyThread.Name = "MyThreadCliente"
        '    MyThread.Start()
        'Catch ex As Exception
        'End Try
    End Sub
    Public Sub BackgroundCliente()
        While True
            Me.BeginInvoke(CallCliente)
            Thread.Sleep(1500)
        End While
    End Sub

    Public Sub BackgroundFactura()
        While True
            Me.BeginInvoke(CallFactura)
            Thread.Sleep(1500)
        End While
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
    Private Sub pagosVendedor()
        Dim RetVal As Long 'Valor de retorno al agregar el documento
        Dim oError As Integer = 0
        Dim errMsg As String = ""
        Dim oPmt As SAPbobsCOM.Payments
        oPmt = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
        oPmt.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice
        oPmt.Invoices.DocEntry = 929
        oPmt.Invoices.SumApplied = 100
        oPmt.CardCode = "400001"
        oPmt.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
        oPmt.TransferAccount = "1130401"
        oPmt.TransferSum = 100
        oPmt.TransferDate = Now
        oPmt.DocDate = Now
        RetVal = oPmt.Add()
        If RetVal <> 0 Then
            oCompany.GetLastError(oError, errMsg)
            MsgBox(errMsg)
        End If
    End Sub
    Private Sub pagocheque()
        Dim RetVal As Long 'Valor de retorno al agregar el documento
        Dim oError As Integer = 0
        Dim errMsg As String = ""
        'Creamos el documento como objeto "Documento"
        Dim oPays As SAPbobsCOM.Payments
        Dim CodigoBanco As String
        'Establecemos el documento de tipo pago a cuenta
        MakeConnectionSAP()
        oPays = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)

        'Dim cheque As SAPbobsCOM.Payments_Checks = oCompany.GetBusinessObject(SAPbobsCOM.BoAccountTypes.at_Other)
        'cheque************************************

        oPays.CardCode = "400001"
        'Cheque
        oPays.Invoices.DocEntry = "878"
        oPays.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
        oPays.Invoices.SumApplied = 5.0
        oPays.Invoices.Add()

        oPays.Checks.CheckAccount = "_SYS00000000387"
        oPays.Checks.AccounttNum = "66-13987-3" 'RTrim(CuentaBancaria) 'estoy quemando estos datos
        oPays.Checks.BankCode = "BGT"
        oPays.Checks.CheckSum = 5.0
        'agregar el cheque
        oPays.Checks.Add()


        'Agregar el pago
        'Check the result
        RetVal = oPays.Add()
        If RetVal <> 0 Then
            oCompany.GetLastError(oError, errMsg)
            MsgBox(errMsg)
        Else
            MsgBox("pago realizado")
        End If
    End Sub
    Private Sub PagoChequeEfectivo()
        Dim RetVal As Long 'Valor de retorno al agregar el documento
        Dim oError As Integer = 0
        Dim errMsg As String = ""
        'Creamos el documento como objeto "Documento"
        Dim oPays As SAPbobsCOM.Payments
        Dim CodigoBanco As String
        'Establecemos el documento de tipo pago a cuenta
        oPays = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
        'cheque************************************

        'oPays.CardName = NombreCheque
        ' ''oPays.DocDueDate = DateTime.Today 'Fecha del documento.
        'oPays.Checks.CheckAccount = "_SYS00000000354"
        ''oPays.Checks.AccounttNum = "0-92000212-8" 'RTrim(CuentaBancaria)
        ''
        'oPays.Checks.CheckSum = Valor
        ' ''oPays.Checks.Currency = "QTZ"
        ''oPays.Checks.Details = Referencia
        ''oPays.Checks.DueDate = DateTime.Today
        ''oPays.Checks.Trnsfrable = 0
        ' ''agregar el cheque
        'oPays.Checks.Add()


        'Solicitud
        'oPays.CardName = NombreCheque
        oPays.AccountPayments.AccountCode = "_SYS00000000004"
        oPays.AccountPayments.Decription = "desc"
        oPays.AccountPayments.SumPaid = 100
        'oPays.DocObjectCode = BoPaymentsObjectType.bopot_OutgoingPayments
        oPays.DocType = BoRcptTypes.rAccount
        oPays.DocCurrency = "QTZ"
        oPays.DocDate = DateTime.Today
        'oPays.IsPayToBank = BoYesNoEnum.tNO
        oPays.JournalRemarks = "Pago: "
        oPays.CashAccount = "_SYS00000000004"
        'oPays.Checks.BankCode = CodigoBanco
        oPays.CashSum = 100
        'oPays.Checks.Add()
        'oPays.LocalCurrency = BoYesNoEnum.tNO

        'Agregar el pago
        RetVal = oPays.Add
        'Check the result
        RetVal = oPays.Add
        If RetVal <> 0 Then
            oCompany.GetLastError(oError, errMsg)
            MsgBox(errMsg)
        End If
    End Sub
    Private Sub Factura()
        Dim oReturn As Integer = -1
        Dim oError As Integer = 0
        Dim errMsg As String = ""
        Dim sql As String
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

                            'sql = ("select AcctCode from oact where ActId ='" + xnDet.SelectSingleNode("accountcode").InnerText + "';")

                            'Dim oRecordSet As SAPbobsCOM.Recordset
                            'oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'oRecordSet.DoQuery(sql)
                            'If oRecordSet.RecordCount > 0 Then
                            '    Dim var As String = oRecordSet.Fields.Item(1).Value
                            'End If

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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        PagoChequeEfectivo()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        pagocheque()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        pagosVendedor()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        FacturaToma()
        'Timer1.Start()
    End Sub
    Public Sub FacturaToma()
        Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
        Dim objSubFolder As Object = "C:\TS\Factura\"
        Dim objFolder As Object = objFSO.GetFolder(objSubFolder)
        Dim colFiles As Object = objFolder.Files

        For Each objFile In colFiles

            If existeFactura() = 0 Then
                Dim entra As String = "C:\TS\Factura\" + objFile.Name.ToString
                Dim sale As String = "C:\TS\Factura\Integration\invoice.xml"
                File.Move(entra, sale)
            ElseIf existeFactura() = 1 Then
                'Timer1.Start()
                Exit Sub
            End If
        Next
    End Sub

    Public Sub ClienteToma()
        Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
        Dim objSubFolder As Object = "C:\TS\Cliente\"
        Dim objFolder As Object = objFSO.GetFolder(objSubFolder)
        Dim colFiles As Object = objFolder.Files

        For Each objFile In colFiles

            If existeCliente() = 0 Then
                Dim entra As String = "C:\TS\Cliente\" + objFile.Name.ToString
                Dim sale As String = "C:\TS\Cliente\Integration\cliente.xml"
                File.Move(entra, sale)
            ElseIf existeFactura() = 1 Then
                'Timer1.Start()
                Exit Sub
            End If
        Next
    End Sub

    Private Function existeFactura()
        If My.Computer.FileSystem.FileExists("C:\TS\Factura\Integration\invoice.xml") Then
            'ListBox1.Items.Add("Si Existe")
            Return 1
        Else
            'ListBox1.Items.Add("No Existe")
            Return 0
        End If
    End Function
    Private Function existeCliente()
        If My.Computer.FileSystem.FileExists("C:\TS\Cliente\Integration\cliente.xml") Then
            'ListBox1.Items.Add("Si Existe")
            Return 1
        Else
            'ListBox1.Items.Add("No Existe")
            Return 0
        End If
    End Function

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'Timer1.Start()
        'Button6_Click(sender, e)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

    End Sub

    Private Sub CambiaPor(doc As String, porc As String, line1 As Integer, line2 As Integer)
        Try
            Dim linea1 As Integer
            Dim linea2 As Integer
            Dim Desc As Double
            linea1 = Convert.ToInt32(line1)
            linea2 = Convert.ToInt32(line2)

            Desc = Convert.ToDecimal(porc)

            Dim delivery As SAPbobsCOM.Documents
            Dim oError As Integer = -1
            Dim message As String = ""
            delivery = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)

            For a As Integer = linea1 To linea2


                If delivery.GetByKey(doc) Then
                    delivery.Lines.SetCurrentLine(a)
                    delivery.Lines.DiscountPercent = Desc
                    oError = delivery.Update()
                    If oError <> 0 Then
                        MessageBox.Show("error: " + oCompany.GetLastErrorDescription)
                    End If

                End If
            Next
        Catch ex As Exception
        End Try
    End Sub


    Private Sub FacturaDeOrden()
        Dim oReturn As Integer = -1
        Dim oError As Integer = 0
        Dim errMsg As String = ""
        Dim sql As String
        Try
            If MakeConnectionSAP() Then
                Dim Invoice As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

                Invoice.Series = "75"
                Invoice.DocNum = "123456"
                Invoice.DocDate = "2017-09-08"
                Invoice.CardCode = "400001"
                'Invoice.DocTotal = "150.00"
                Invoice.DocType = BoDocumentTypes.dDocument_Items

                Invoice.Lines.BaseEntry = "22"

                Invoice.Lines.BaseLine = 0
                Invoice.Lines.BaseType = "17"
                Invoice.Lines.Add()


                Invoice.Comments = "pruebas"
                oReturn = Invoice.Add()
                If oReturn <> 0 Then
                    oCompany.GetLastError(oError, errMsg)
                    MsgBox(errMsg)
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        FacturaDeOrden()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim oReturn As Integer = -1
        Dim oError As Integer = 0
        Dim errMsg As String = ""
        Dim sql As String
        Try
            If MakeConnectionSAP() Then
                Dim Journal As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

                Journal.Series = "75"
                Journal.DocNum = "123456"
                Journal.DocDate = "2017-09-08"
                Journal.CardCode = "400001"
                'Invoice.DocTotal = "150.00"
                Journal.DocType = BoDocumentTypes.dDocument_Items

                Journal.Lines.BaseEntry = "22"

                Journal.Lines.BaseLine = 0
                Journal.Lines.BaseType = "17"
                Journal.Lines.Add()


                Journal.Comments = "pruebas"
                oReturn = Journal.Add()
                If oReturn <> 0 Then
                    oCompany.GetLastError(oError, errMsg)
                    MsgBox(errMsg)
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim InPay As SAPbobsCOM.Payments
        Dim fecha As Date
        MakeConnectionSAP()
        InPay = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)


        InPay.CardCode = "400001"

        InPay.Invoices.DocEntry = "872" ' Invoice Number that we just created.
        InPay.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice

        InPay.CreditCards.CreditCard = 1  ' Mastercard = 1 , VISA = 2
        DateTime.TryParseExact(DateTime.Today, "MM/yy", Nothing, DateTimeStyles.None, fecha)
        fecha = "01/12/17"
        InPay.CreditCards.CardValidUntil = "31/12/17"
        'InPay.CreditCards.CardValidUntil = "12/17"
        InPay.CreditCards.CreditCardNumber = "1111" ' Just need 4 last digits
        InPay.CreditCards.CreditSum = 5 ' Total Amount of the Invoice
        InPay.CreditCards.VoucherNum = "1234" ' Need to give the Credit Card confirmation number.

        If InPay.Add() <> 0 Then
            MsgBox(oCompany.GetLastErrorDescription())
        Else
            MsgBox("Incoming payment Created!")
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim RetVal As Long 'Valor de retorno al agregar el documento
        Dim oError As Integer = 0
        Dim errMsg As String = ""
        'Creamos el documento como objeto "Documento"
        Dim oPays As SAPbobsCOM.Payments
        Dim CodigoBanco As String
        'Establecemos el documento de tipo pago a cuenta
        MakeConnectionSAP()
        oPays = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)

        'Dim cheque As SAPbobsCOM.Payments_Checks = oCompany.GetBusinessObject(SAPbobsCOM.BoAccountTypes.at_Other)
        'cheque************************************

        oPays.CardCode = "400001"
        'Cheque
        oPays.Invoices.DocEntry = "887"
        oPays.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
        oPays.Invoices.SumApplied = 5.0
        oPays.Invoices.Add()

        oPays.TransferAccount = "_SYS00000000004"
        oPays.TransferSum = 5 'RTrim(CuentaBancaria) 'estoy quemando estos datos
        oPays.TransferReference = "BGT"
        'agregar el cheque
        oPays.Checks.Add()


        'Agregar el pago
        'Check the result
        RetVal = oPays.Add()
        If RetVal <> 0 Then
            oCompany.GetLastError(oError, errMsg)
            MsgBox(errMsg)
        Else
            MsgBox("pago realizado")
        End If
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dim RetVal As Long 'Valor de retorno al agregar el documento
        Dim oError As Integer = 0
        Dim errMsg As String = ""
        MakeConnectionSAP()
        'Creamos el documento como objeto "Documento"
        'Dim oCancelDoc As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.o)
        Dim oInvoice As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        Dim oInvoice2 As SAPbobsCOM.Documents
        'Establecemos el documento de tipo pago a cuenta

        oInvoice.GetByKey(894)
        oInvoice2 = oInvoice.CreateCancellationDocument()

        RetVal = oInvoice2.Add()

        If RetVal <> 0 Then
            oCompany.GetLastError(oError, errMsg)
            MsgBox(errMsg)
        Else
            MsgBox("pago realizado")
        End If

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Dim oReturn As Integer = -1
        Dim oError As Integer = 0
        Dim errMsg As String = ""
        Dim sql As String
        Try
            If MakeConnectionSAP() Then
                Dim orders As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)

                orders.Series = "6"
                orders.DocNum = "120"
                orders.DocDate = "2017-10-12"
                orders.DocDueDate = "2017-10-12"
                orders.TaxDate = "2017-10-12"
                orders.CardCode = "400001"
                orders.DocType = BoDocumentTypes.dDocument_Items
                orders.Lines.ItemCode = "1198"
                orders.Lines.Quantity = "1"
                orders.Lines.TaxCode = "EXE"
                orders.Lines.Add()
                orders.Comments = "pruebas"
                oReturn = orders.Add()
                If oReturn <> 0 Then
                    oCompany.GetLastError(oError, errMsg)
                    MsgBox(errMsg)
                Else
                    MsgBox("Orden Creada correctamente")
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Function AplicacionFuncionando() As Boolean

        Dim aProceso() As Process
        aProceso = Process.GetProcesses()
        Dim oProceso As Process
        Dim Nombre_Proceso As String
        For Each oProceso In aProceso
            Nombre_Proceso = oProceso.ProcessName.ToString()
            If Nombre_Proceso = "Integration_SAP" Then
                Me.count += 1 'Debes declarar esta variable Global 
            End If
        Next
        If count = 2 Then
            Return 1
        End If
    End Function
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        If (AplicacionFuncionando() = True) Then
            MessageBox.Show("Ejecutada")
        Else
            MessageBox.Show("No Ejecutada")
        End If
    End Sub
End Class
