Imports System.Text
Imports System.IO
Imports System.Net.Mail
Imports System.Net
Imports System.Configuration
Imports log4net
Imports log4net.Config
Imports NPOI.HSSF.UserModel
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

Module Module1
	Private ReadOnly log As ILog = LogManager.GetLogger(GetType(Module1))

	Sub Main()
		XmlConfigurator.Configure()
		GetDataAndSendFile()
	End Sub

	Private Sub GetDataAndSendFile()
		log.Info("Execution started")
		Dim exportBal As ExportBAL = New ExportBAL()
		Dim dtData As DataTable = New DataTable()
		dtData = exportBal.GetALLAutoTransactionExportSettings()

		If (dtData IsNot Nothing) Then
			If (dtData.Rows.Count > 0) Then
				For Each dr As DataRow In dtData.Rows
					Dim companyId As Integer = dr("CompanyId")
					Dim ExportOption As Integer = dr("ExportOption")
					Dim FtpServerPath As String = dr("FtpServerPath")
					Dim FtpUsername As String = dr("FtpUsername")
					Dim FtpPassword As String = dr("FtpPassword")
					Dim EmailId As String = dr("EmailId")
					Dim OffSetValue As Integer = dr("OffSetValue")
					Dim ExecutionTime As TimeSpan = dr("ExecutionTime")
					Dim CustomizedExportTemplateId As String = dr("CustomizedExportTemplateId").ToString()
					Dim Separator As String = dr("Separator").ToString()
					Dim IncludePreviouslyExportTransactions As Boolean = dr("IncludePreviouslyExportTransactions").ToString()
					Dim ExportOnlyNewTransactions As Boolean = dr("ExportOnlyNewTransactions").ToString()
					Dim ExportZeroQtyTransactions As Boolean = dr("ExportZeroQtyTransactions").ToString()
					Dim DecimalPlace As String = dr("DecimalPlace").ToString()
					Dim DecimalType As String = dr("DecimalType").ToString()
					Dim DateType As String = dr("DateType").ToString()

					Dim currentDate As DateTime = DateTime.UtcNow.AddMinutes(OffSetValue)
					Dim executionDate As DateTime = Convert.ToDateTime(DateTime.UtcNow.AddMinutes(OffSetValue).ToString("yyyy-MM-dd") & " " & ExecutionTime.ToString())

					Dim startTime As TimeSpan = New TimeSpan(23, 30, 0)
					Dim endTime As TimeSpan = New TimeSpan(23, 59, 0)

					If (ExecutionTime >= startTime And ExecutionTime <= endTime) Then
						executionDate = executionDate.AddDays(-1)
					End If

					If (currentDate >= executionDate) Then

						Dim filePath As String = ""
						Dim fileName As String = ""
						Dim result As Integer = 0
						Dim checkTransReportSend As Integer = exportBal.CheckTransReportSend(companyId, executionDate)
						If (checkTransReportSend = 1) Then
							CreateFile(companyId, ExportOption, OffSetValue, filePath, fileName, executionDate,
									   CustomizedExportTemplateId, Separator, IncludePreviouslyExportTransactions,
									   ExportOnlyNewTransactions, ExportZeroQtyTransactions, DecimalPlace, DecimalType, DateType)
							If (filePath <> "") Then

								If (EmailId <> "") Then
									exportBal = New ExportBAL()
									result = exportBal.SaveUpdateAutoTransactionExportHistory(0, companyId, executionDate, fileName, True, False)
									SendEmail(EmailId, filePath, companyId, executionDate, fileName)
								End If
								If (FtpServerPath <> "") Then
									exportBal = New ExportBAL()
									exportBal.SaveUpdateAutoTransactionExportHistory(result, companyId, executionDate, fileName, IIf(result <> 0, True, False), True)

									SendFileToFTP(filePath, fileName, FtpServerPath, FtpUsername, FtpPassword, companyId, executionDate, result)
								End If
							End If
						End If
					End If
				Next
			End If
		End If
		log.Info("Execution Ended")
	End Sub

	Private Sub CreateFile(CompanyId As Integer, ExportOption As Integer, OffSetValue As Integer, ByRef filePath As String, ByRef fileName As String,
						   executionDate As DateTime, CustomizedExportTemplateId As Integer, Separator As String, IncludePreviouslyExportTransactions As Boolean,
						   ExportOnlyNewTransactions As Boolean, ExportZeroQtyTransactions As Boolean, DecimalPlace As String, DecimalType As String, DateType As String)

		Try

			Dim exportBal As ExportBAL = New ExportBAL()
			Dim dtTransactions As DataTable = New DataTable()
			Dim sb As New StringBuilder()


			dtTransactions = exportBal.GetTransactionByCompanyId(CompanyId, DecimalPlace, DecimalType)

			If (dtTransactions Is Nothing) Then
				log.Error("Data not found to CompanyId : " & CompanyId & ", ExportOption : " & ExportOption & ", executionDate: " & executionDate)
				Return
			End If
			If (dtTransactions.Rows.Count = 0) Then
				log.Error("Data not found to CompanyId : " & CompanyId & ", ExportOption : " & ExportOption & ", executionDate: " & executionDate)
				Return
			End If

			If (IncludePreviouslyExportTransactions = False And ExportOnlyNewTransactions = True) Then
				Dim dvTemp As DataView = dtTransactions.DefaultView
				dvTemp.RowFilter = "IsExportedTransaction=0"
				dtTransactions = dvTemp.ToTable()
			ElseIf (IncludePreviouslyExportTransactions = False And ExportOnlyNewTransactions = False) Then
				Dim dvTemp As DataView = dtTransactions.DefaultView
				dvTemp.RowFilter = "IsExportedTransaction=0"
				dtTransactions = dvTemp.ToTable()
			End If

			If (ExportZeroQtyTransactions = False) Then
				Dim dvTemp As DataView = dtTransactions.DefaultView
				dvTemp.RowFilter = "Quantity <> '0.0' and Quantity <> '0' and Quantity <> '0.00' and Quantity <> '00' and Quantity <> '000'"
				dtTransactions = dvTemp.ToTable()
			End If

			Dim dsCustomizedExportTemplates As DataSet = New DataSet()
			Dim dtCustomizedExportField As DataTable = New DataTable()

			dsCustomizedExportTemplates = exportBal.GetCustomizedExportTemplateById(CustomizedExportTemplateId)
			dtCustomizedExportField = dsCustomizedExportTemplates.Tables(1)

			Dim dv As DataView = dtTransactions.DefaultView
			Dim fields As String = String.Join(",", (From row In dtCustomizedExportField.AsEnumerable Select row("FieldName")).ToArray)
			Dim columns() As String = fields.Split(",")

			Dim dtTemp As DataTable = dv.ToTable(False, columns)

			If (ExportOption = "1") Then


				'For k As Integer = 0 To dtTemp.Columns.Count - 1
				'    sb.Append(dtTemp.Columns(k).ColumnName)
				'Next
				'append new line
				For i As Integer = 0 To dtTemp.Rows.Count - 1
					For k As Integer = 0 To dtTemp.Columns.Count - 1
						'add separator
						Dim dtTempCustomizedExportField As DataTable = New DataTable()
						Dim dvTemp As DataView = dtCustomizedExportField.DefaultView
						dvTemp.RowFilter = "FieldName='" & dtTemp.Columns(k).ColumnName & "'"
						dtTempCustomizedExportField = dvTemp.ToTable()

						Dim FieldLength As Integer = dtTempCustomizedExportField.Rows(0)("FieldLength").ToString()
						Dim FillCharacter As String = dtTempCustomizedExportField.Rows(0)("PaddingCharacter").ToString()
						Dim justify As String = dtTempCustomizedExportField.Rows(0)("Justify").ToString()

						If (dtTemp.Columns(k).ColumnName = "Date") Then

							Dim FieldColumnValue As String = Convert.ToDateTime(dtTemp.Rows(i)(k)).ToString(DateType) '.ToString("MMddyyyy")
							Dim FieldColumnValueLength As Integer = FieldColumnValue.Length
							'Dim paddingLength As Integer = FieldLength - FieldColumnValueLength
							Dim finalValue As String = ""

							If (FillCharacter = "") Then
								FillCharacter = " "
							End If

							If (FieldLength > FieldColumnValueLength) Then
								If (justify = "Left") Then
									finalValue = FieldColumnValue.PadLeft(FieldLength, FillCharacter)
								ElseIf (justify = "Right") Then
									finalValue = FieldColumnValue.PadRight(FieldLength, FillCharacter)
								Else
									finalValue = FieldColumnValue.PadLeft(FieldLength, FillCharacter)
								End If
							ElseIf (FieldLength < FieldColumnValueLength) Then
								finalValue = FieldColumnValue.Substring(0, FieldLength)
							Else
								finalValue = FieldColumnValue
							End If

							sb.Append(finalValue)

						ElseIf (dtTemp.Columns(k).ColumnName = "Time") Then
							Dim FieldColumnValue As String = Convert.ToDateTime(dtTemp.Rows(i)(k)).ToString("HHmm")
							Dim FieldColumnValueLength As Integer = FieldColumnValue.Length
							'Dim paddingLength As Integer = FieldLength - FieldColumnValue
							Dim finalValue As String = ""

							If (FillCharacter = "") Then
								FillCharacter = " "
							End If

							If (FieldLength > FieldColumnValueLength) Then
								If (justify = "Left") Then
									finalValue = FieldColumnValue.PadLeft(FieldLength, FillCharacter)
								ElseIf (justify = "Right") Then
									finalValue = FieldColumnValue.PadRight(FieldLength, FillCharacter)
								Else
									finalValue = FieldColumnValue.PadLeft(FieldLength, FillCharacter)
								End If
							ElseIf (FieldLength < FieldColumnValueLength) Then
								finalValue = FieldColumnValue.Substring(0, FieldLength)
							Else
								finalValue = FieldColumnValue
							End If

							sb.Append(finalValue)

							'sb.Append(Convert.ToDateTime(dtTemp.Rows(i)(k)).ToString("hhmm").PadRight(20, " "))
						Else
							'sb.Append(dtTemp.Rows(i)(k).ToString().Replace(",", ";"))

							Dim FieldColumnValue As String = dtTemp.Rows(i)(k).ToString().Replace(",", ";")
							Dim FieldColumnValueLength As Integer = FieldColumnValue.Length
							'Dim paddingLength As Integer = FieldLength - FieldColumnValue
							Dim finalValue As String = ""

							If (FillCharacter = "") Then
								FillCharacter = " "
							End If

							If (FieldLength > FieldColumnValueLength) Then
								If (justify = "Left") Then
									finalValue = FieldColumnValue.PadLeft(FieldLength, FillCharacter)
								ElseIf (justify = "Right") Then
									finalValue = FieldColumnValue.PadRight(FieldLength, FillCharacter)
								Else
									finalValue = FieldColumnValue.PadLeft(FieldLength, FillCharacter)
								End If
							ElseIf (FieldLength < FieldColumnValueLength) Then
								finalValue = FieldColumnValue.Substring(0, FieldLength)
							Else
								finalValue = FieldColumnValue
							End If

							sb.Append(finalValue)

						End If
					Next
					'append new line
					sb.Append(vbCr & vbLf)
				Next

				CreateFile(filePath, fileName, OffSetValue, sb, ExportOption)

			ElseIf (ExportOption = "2") Then

				sb = New StringBuilder()
				'For k As Integer = 0 To dtTemp.Columns.Count - 1
				'    sb.Append(dtTemp.Columns(k).ColumnName)
				'Next
				'append new line
				For i As Integer = 0 To dtTemp.Rows.Count - 1
					For k As Integer = 0 To dtTemp.Columns.Count - 1
						'add separator
						Dim dtTempCustomizedExportField As DataTable = New DataTable()
						Dim dvTemp As DataView = dtCustomizedExportField.DefaultView
						dvTemp.RowFilter = "FieldName='" & dtTemp.Columns(k).ColumnName & "'"
						dtTempCustomizedExportField = dvTemp.ToTable()

						Dim FieldLength As Integer = dtTempCustomizedExportField.Rows(0)("FieldLength").ToString()
						Dim FillCharacter As String = dtTempCustomizedExportField.Rows(0)("PaddingCharacter").ToString()
						Dim justify As String = dtTempCustomizedExportField.Rows(0)("Justify").ToString()

						If (dtTemp.Columns(k).ColumnName = "Date") Then

							Dim FieldColumnValue As String = Convert.ToDateTime(dtTemp.Rows(i)(k)).ToString(DateType) '.ToString("MMddyyyy")
							Dim FieldColumnValueLength As Integer = FieldColumnValue.Length
							'Dim paddingLength As Integer = FieldLength - FieldColumnValueLength
							Dim finalValue As String = ""

							If (FieldLength > FieldColumnValueLength) Then
								If (justify = "Left") Then
									finalValue = FieldColumnValue.PadLeft(FieldLength, FillCharacter)
								ElseIf (justify = "Right") Then
									finalValue = FieldColumnValue.PadRight(FieldLength, FillCharacter)
								Else
									finalValue = FieldColumnValue
								End If
							ElseIf (FieldLength < FieldColumnValueLength) Then
								finalValue = FieldColumnValue.Substring(0, FieldLength)
							Else
								finalValue = FieldColumnValue
							End If

							If (Separator = "none") Then
								sb.Append(finalValue)
							ElseIf (Separator = "comma") Then
								sb.Append(finalValue + ","c)
							Else
								sb.Append(finalValue + Separator)
							End If


						ElseIf (dtTemp.Columns(k).ColumnName = "Time") Then
							Dim FieldColumnValue As String = Convert.ToDateTime(dtTemp.Rows(i)(k)).ToString("HHmm")
							Dim FieldColumnValueLength As Integer = FieldColumnValue.Length
							'Dim paddingLength As Integer = FieldLength - FieldColumnValue
							Dim finalValue As String = ""

							If (FieldLength > FieldColumnValueLength) Then
								If (justify = "Left") Then
									finalValue = FieldColumnValue.PadLeft(FieldLength, FillCharacter)
								ElseIf (justify = "Right") Then
									finalValue = FieldColumnValue.PadRight(FieldLength, FillCharacter)
								Else
									finalValue = FieldColumnValue.PadLeft(FieldLength, FillCharacter)
								End If
							ElseIf (FieldLength < FieldColumnValueLength) Then
								finalValue = FieldColumnValue.Substring(0, FieldLength)
							Else
								finalValue = FieldColumnValue
							End If

							If (Separator = "none") Then
								sb.Append(finalValue)
							ElseIf (Separator = "comma") Then
								sb.Append(finalValue + ","c)
							Else
								sb.Append(finalValue + Separator)
							End If

							'sb.Append(Convert.ToDateTime(dtTemp.Rows(i)(k)).ToString("hhmm").PadRight(20, " "))
						Else
							'sb.Append(dtTemp.Rows(i)(k).ToString().Replace(",", ";"))

							Dim FieldColumnValue As String = dtTemp.Rows(i)(k).ToString().Replace(",", ";")
							Dim FieldColumnValueLength As Integer = FieldColumnValue.Length
							'Dim paddingLength As Integer = FieldLength - FieldColumnValue
							Dim finalValue As String = ""

							If (FillCharacter = "") Then
								FillCharacter = " "
							End If

							If (FieldLength > FieldColumnValueLength) Then
								If (justify = "Left") Then
									finalValue = FieldColumnValue.PadLeft(FieldLength, FillCharacter)
								ElseIf (justify = "Right") Then
									finalValue = FieldColumnValue.PadRight(FieldLength, FillCharacter)
								Else
									finalValue = FieldColumnValue
								End If
							ElseIf (FieldLength < FieldColumnValueLength) Then
								finalValue = FieldColumnValue.Substring(0, FieldLength)
							Else
								finalValue = FieldColumnValue
							End If

							If (Separator = "none") Then
								sb.Append(finalValue)
							ElseIf (Separator = "comma") Then
								sb.Append(finalValue + ","c)
							Else
								sb.Append(finalValue + Separator)
							End If

						End If
					Next
					sb = sb.Remove(sb.Length - 1, 1)
					'append new line
					sb.Append(vbCr & vbLf)
				Next

				CreateFile(filePath, fileName, OffSetValue, sb, ExportOption)

			Else

				sb = New StringBuilder()
				'For k As Integer = 0 To dtTemp.Columns.Count - 1
				'    sb.Append(dtTemp.Columns(k).ColumnName)
				'Next
				'append new line
				Dim dtFinal As DataTable = New DataTable()
				'dtFinal = dtTemp.Clone()

				For l As Integer = 0 To dtTemp.Columns.Count - 1
					dtFinal.Columns.Add(dtTemp.Columns(l).ColumnName, System.Type.[GetType]("System.String"))
				Next

				For i As Integer = 0 To dtTemp.Rows.Count - 1
					Dim drNew As DataRow = dtFinal.NewRow()

					For k As Integer = 0 To dtTemp.Columns.Count - 1
						'add separator
						Dim dtTempCustomizedExportField As DataTable = New DataTable()
						Dim dvTemp As DataView = dtCustomizedExportField.DefaultView
						dvTemp.RowFilter = "FieldName='" & dtTemp.Columns(k).ColumnName & "'"
						dtTempCustomizedExportField = dvTemp.ToTable()

						Dim FieldLength As Integer = dtTempCustomizedExportField.Rows(0)("FieldLength").ToString()
						Dim FillCharacter As String = dtTempCustomizedExportField.Rows(0)("PaddingCharacter").ToString()
						Dim justify As String = dtTempCustomizedExportField.Rows(0)("Justify").ToString()

						If (dtTemp.Columns(k).ColumnName = "Date") Then

							Dim FieldColumnValue As String = Convert.ToDateTime(dtTemp.Rows(i)(k)).ToString(DateType) '.ToString("MMddyyyy")
							Dim FieldColumnValueLength As Integer = FieldColumnValue.Length
							'Dim paddingLength As Integer = FieldLength - FieldColumnValueLength
							Dim finalValue As String = ""

							If (FieldLength > FieldColumnValueLength) Then
								If (justify = "Left") Then
									finalValue = FieldColumnValue.PadLeft(FieldLength, FillCharacter)
								ElseIf (justify = "Right") Then
									finalValue = FieldColumnValue.PadRight(FieldLength, FillCharacter)
								Else
									finalValue = FieldColumnValue
								End If
							ElseIf (FieldLength < FieldColumnValueLength) Then
								finalValue = FieldColumnValue.Substring(0, FieldLength)
							Else
								finalValue = FieldColumnValue
							End If

							'dtFinal.Rows(i)(k) = finalValue.ToString()
							drNew(k) = finalValue
							'sb.Append(finalValue + ","c)


						ElseIf (dtTemp.Columns(k).ColumnName = "Time") Then
							Dim FieldColumnValue As String = Convert.ToDateTime(dtTemp.Rows(i)(k)).ToString("HHmm")
							Dim FieldColumnValueLength As Integer = FieldColumnValue.Length
							'Dim paddingLength As Integer = FieldLength - FieldColumnValue
							Dim finalValue As String = ""

							If (FieldLength > FieldColumnValueLength) Then
								If (justify = "Left") Then
									finalValue = FieldColumnValue.PadLeft(FieldLength, FillCharacter)
								ElseIf (justify = "Right") Then
									finalValue = FieldColumnValue.PadRight(FieldLength, FillCharacter)
								Else
									finalValue = FieldColumnValue.PadLeft(FieldLength, FillCharacter)
								End If
							ElseIf (FieldLength < FieldColumnValueLength) Then
								finalValue = FieldColumnValue.Substring(0, FieldLength)
							Else
								finalValue = FieldColumnValue
							End If

							'sb.Append(finalValue + ","c)
							'dtFinal.Rows(i)(k) = finalValue
							drNew(k) = finalValue
							'sb.Append(Convert.ToDateTime(dtTemp.Rows(i)(k)).ToString("hhmm").PadRight(20, " "))
						Else
							'sb.Append(dtTemp.Rows(i)(k).ToString().Replace(",", ";"))

							Dim FieldColumnValue As String = dtTemp.Rows(i)(k).ToString().Replace(",", ";")
							Dim FieldColumnValueLength As Integer = FieldColumnValue.Length
							'Dim paddingLength As Integer = FieldLength - FieldColumnValue
							Dim finalValue As String = ""

							If (FillCharacter = "") Then
								FillCharacter = " "
							End If

							If (FieldLength > FieldColumnValueLength) Then
								If (justify = "Left") Then
									finalValue = FieldColumnValue.PadLeft(FieldLength, FillCharacter)
								ElseIf (justify = "Right") Then
									finalValue = FieldColumnValue.PadRight(FieldLength, FillCharacter)
								Else
									finalValue = FieldColumnValue
								End If
							ElseIf (FieldLength < FieldColumnValueLength) Then
								finalValue = FieldColumnValue.Substring(0, FieldLength)
							Else
								finalValue = FieldColumnValue
							End If

							'sb.Append(finalValue + ","c)
							'dtFinal.Rows(i)(k) = finalValue
							drNew(k) = finalValue

						End If
					Next

					dtFinal.Rows.Add(drNew)

					'sb = sb.Remove(sb.Length - 1, 1)
					''append new line
					'sb.Append(vbCr & vbLf)
				Next

				If (ExportOption = "3") Then
					WriteExcelWithNPOI("xls", dtFinal, False, False, OffSetValue, filePath, fileName)
				Else
					WriteExcelWithNPOI("xlsx", dtFinal, False, False, OffSetValue, filePath, fileName)
				End If

			End If


			Dim transactionIds As String = String.Join(",", (From row In dtTransactions.AsEnumerable Select row("TransactionId")).ToArray)

			exportBal = New ExportBAL()
			exportBal.UpdateExportedTransaction(transactionIds)

			'Dim exportBal As ExportBAL = New ExportBAL()

			'Dim dtTransactions As DataTable = New DataTable()

			'dtTransactions = exportBal.GetTransactionByCompanyId(CompanyId)

			'If (dtTransactions Is Nothing) Then
			'    log.Error("Data not found to CompanyId : " & CompanyId & ", ExportOption : " & ExportOption & ", executionDate: " & executionDate)
			'    Return
			'End If
			'If (dtTransactions.Rows.Count = 0) Then
			'    log.Error("Data not found to CompanyId : " & CompanyId & ", ExportOption : " & ExportOption & ", executionDate: " & executionDate)
			'    Return
			'End If

			'Dim sb As New StringBuilder()
			'For k As Integer = 0 To dtTransactions.Columns.Count - 1
			'    sb.Append(dtTransactions.Columns(k).ColumnName + ","c)
			'Next
			''append new line
			'sb.Append(vbCr & vbLf)
			'For i As Integer = 0 To dtTransactions.Rows.Count - 1
			'    For k As Integer = 0 To dtTransactions.Columns.Count - 1
			'        'add separator
			'        If (dtTransactions.Columns(k).ColumnName = "Date") Then
			'            sb.Append(Convert.ToDateTime(dtTransactions.Rows(i)(k)).ToString("dd-MMM-yyyy").Replace(",", ";") + ","c)
			'        ElseIf (dtTransactions.Columns(k).ColumnName = "Time") Then
			'            sb.Append(Convert.ToDateTime(dtTransactions.Rows(i)(k)).ToString("hh:mm tt").Replace(",", ";") + ","c)
			'        Else
			'            sb.Append(dtTransactions.Rows(i)(k).ToString().Replace(",", ";") + ","c)

			'        End If


			'    Next
			'    'append new line
			'    sb.Append(vbCr & vbLf)
			'Next

			'Dim FileDate As String = DateTime.UtcNow.AddMinutes(OffSetValue).ToString("yyyyMMdd")
			'fileName = "FluidSecure" & "_" & FileDate

			'Select Case ExportOption
			'    Case "1"
			'        fileName += ".txt"
			'    Case "2"
			'        fileName += ".csv"
			'    Case "3"
			'        fileName += ".xls"
			'    Case Else
			'        fileName += ".txt"
			'End Select


			'filePath = System.Environment.CurrentDirectory & "\" & fileName

			'Dim createdFile As StreamWriter = New StreamWriter(filePath)
			'createdFile.WriteLine(sb.ToString())
			'createdFile.Close()

		Catch ex As Exception

			log.Error("Exception occurred in CreateFile to CompanyId : " & CompanyId & ", ExportOption : " & ExportOption & ", OffSetValue: " & OffSetValue & ". ex is :" & ex.Message)
		End Try
	End Sub

	Private Sub SendEmail(EmailId As String, filePath As String, CompanyId As Integer, ExecutionDate As DateTime, FileName As String)
		Try

			Dim body As String = String.Empty
			Using sr As New StreamReader(ConfigurationManager.AppSettings("PathForExportTransactionEmailTemplate"))
				body = sr.ReadToEnd()
			End Using
			'------------------
			body = body.Replace("emailId", EmailId)

			Dim mailClient As New SmtpClient(ConfigurationManager.AppSettings("smtpServer"))

			mailClient.Credentials = New NetworkCredential(ConfigurationManager.AppSettings("emailAccount"), ConfigurationManager.AppSettings("emailPassword"))
			mailClient.Port = Convert.ToInt32(ConfigurationManager.AppSettings("smtpPort"))

			Dim messageSend As New MailMessage()
			messageSend.From = New MailAddress(ConfigurationManager.AppSettings("FromEmail"))
			'log.Info("Email send to: " + EmailId)
			'messageSend.To.Add(emailTo)

			messageSend.Subject = ConfigurationManager.AppSettings("TransSubject")
			messageSend.Body = body
			messageSend.Attachments.Add(New Attachment(filePath))
			messageSend.IsBodyHtml = True
			mailClient.EnableSsl = Convert.ToBoolean(ConfigurationManager.AppSettings("EnableSsl"))

			Dim emailArray As String() = EmailId.Split(",")

			For index = 0 To emailArray.Length - 1
				messageSend.To.Add(emailArray(index))
				mailClient.Send(messageSend)
				'log.Info("Email send to: " + emailArray(index))
				messageSend.To.Remove(New MailAddress(emailArray(index)))
			Next

			log.Debug("Transaction file sent to " & EmailId)

		Catch ex As Exception
			log.Error("Exception occurred in SendEmail to EmailId : " & EmailId & ". ex is :" & ex.Message)
		End Try

	End Sub

	Private Sub SendFileToFTP(filePath As String, FileName As String, FtpServerPath As String, FtpUsername As String, FtpPassword As String, companyId As Integer, ExecutionDate As DateTime, AutoTransactionExportHistoryId As Integer)
		Try
			If (Not FtpServerPath.Contains("ftp://")) Then
				FtpServerPath = "ftp://" & FtpServerPath
			End If
			'Create Request To Upload File'
			Dim wrUpload As FtpWebRequest = DirectCast(WebRequest.Create(FtpServerPath & "/" & FileName), FtpWebRequest)

			'Specify Username & Password'
			wrUpload.Credentials = New NetworkCredential(FtpUsername, FtpPassword)

			'Start Upload Process'
			wrUpload.Method = WebRequestMethods.Ftp.UploadFile

			'Locate File And Store It In Byte Array'
			Dim btfile() As Byte = File.ReadAllBytes(filePath)

			'Get File'
			Dim strFile As Stream = wrUpload.GetRequestStream()

			'Upload Each Byte'
			strFile.Write(btfile, 0, btfile.Length)

			'Close'
			strFile.Close()

			'Free Memory'
			strFile.Dispose()

			Dim response As FtpWebResponse = wrUpload.GetResponse()

			response.Close()


		Catch ex As Exception
			log.Error("Exception occurred in SendFileToFTP to FtpServerPath : " & FtpServerPath & ". ex is :" & ex.Message)
		End Try

	End Sub

	Private Sub CreateFile(ByRef filePath As String, ByRef fileName As String, OffSetValue As Integer, sb As StringBuilder, ExportOption As String)
		Dim FileDate As String = DateTime.UtcNow.AddMinutes(OffSetValue).ToString("yyyyMMdd")
		fileName = "TransactionDetails" & "_" & FileDate

		Select Case ExportOption
			Case "1"
				fileName += ".txt"
			Case "2"
				fileName += ".csv"
			Case Else
				fileName += ".txt"
		End Select


		filePath = System.Environment.CurrentDirectory & "\" & fileName

		Dim createdFile As StreamWriter = New StreamWriter(filePath)
		createdFile.WriteLine(sb.ToString())
		createdFile.Close()
	End Sub

	Public Sub WriteExcelWithNPOI(ByVal extension As String, ByVal dt As DataTable,
								  isDefaultTemplate As Boolean, isOnlyTemplate As Boolean, OffSetValue As Integer, ByRef filePath As String, ByRef fileName As String)
		Dim workbook As IWorkbook

		If extension = "xlsx" Then
			workbook = New XSSFWorkbook()
		ElseIf extension = "xls" Then
			workbook = New HSSFWorkbook()
		Else
			Throw New Exception("This format is not supported")
		End If

		Dim sheet1 As ISheet = workbook.CreateSheet("Sheet 1")
		If (isOnlyTemplate = True) Then
			Dim row1 As IRow = sheet1.CreateRow(0)

			For j As Integer = 0 To dt.Columns.Count - 1
				Dim cell As ICell = row1.CreateCell(j)
				Dim columnName As String = dt.Columns(j).ToString()
				cell.SetCellValue(columnName)
			Next
		Else
			If (isDefaultTemplate = True) Then

				Dim row1 As IRow = sheet1.CreateRow(0)

				For j As Integer = 0 To dt.Columns.Count - 1
					Dim cell As ICell = row1.CreateCell(j)
					Dim columnName As String = dt.Columns(j).ToString()
					cell.SetCellValue(columnName)
				Next

				For i As Integer = 0 To dt.Rows.Count - 1
					Dim row As IRow = sheet1.CreateRow(i + 1)

					For j As Integer = 0 To dt.Columns.Count - 1
						Dim cell As ICell = row.CreateCell(j)
						Dim columnName As String = dt.Columns(j).ToString()
						cell.SetCellValue(dt.Rows(i)(columnName).ToString())
					Next
				Next

			Else

				For i As Integer = 0 To dt.Rows.Count - 1
					Dim row As IRow = sheet1.CreateRow(i)

					For j As Integer = 0 To dt.Columns.Count - 1
						Dim cell As ICell = row.CreateCell(j)
						Dim columnName As String = dt.Columns(j).ToString()
						cell.SetCellValue(dt.Rows(i)(columnName).ToString())
					Next
				Next

			End If

		End If


		Using exportData = New MemoryStream()
			'Response.Clear()
			workbook.Write(exportData)
			Dim FileDate As String = DateTime.UtcNow.AddMinutes(OffSetValue).ToString("yyyyMMdd")
			fileName = "TransactionDetails" & "_" & FileDate & "." & extension
			filePath = System.Environment.CurrentDirectory & "\" & fileName

			Dim bw As BinaryWriter = New BinaryWriter(File.Open(filePath, FileMode.OpenOrCreate))
			If extension = "xlsx" Then
				bw.Write(exportData.ToArray())
			ElseIf extension = "xls" Then
				bw.Write(exportData.GetBuffer())
			End If

			bw.Close()

			'If extension = "xlsx" Then
			'	Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
			'	Response.AddHeader("Content-Disposition", String.Format("attachment;filename={0}", filename))
			'	Response.BinaryWrite(exportData.ToArray())
			'ElseIf extension = "xls" Then
			'	Response.ContentType = "application/vnd.ms-excel"
			'	Response.AddHeader("Content-Disposition", String.Format("attachment;filename={0}", filename))
			'	Response.BinaryWrite(exportData.GetBuffer())
			'End If

			'Response.[End]()
		End Using
	End Sub

End Module
