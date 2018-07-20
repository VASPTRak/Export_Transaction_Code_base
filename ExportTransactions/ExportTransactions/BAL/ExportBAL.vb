Imports log4net
Imports log4net.Config
Imports System.Data.SqlClient

Public Class ExportBAL
    Private Shared ReadOnly log As ILog = LogManager.GetLogger(GetType(ExportBAL))

    Shared Sub New()
        XmlConfigurator.Configure()
    End Sub

    Public Function GetALLAutoTransactionExportSettings() As DataTable
        Dim dal = New GeneralizedDAL()
        Try

            Dim ds As DataSet = New DataSet()

            Dim Param As SqlParameter() = New SqlParameter() {}

            ds = dal.ExecuteStoredProcedureGetDataSet("usp_tt_AutoTransactionExportSettings_GetALLAutoTransactionExportSettings", Param)

            Return ds.Tables(0)

        Catch ex As Exception
            log.Error("Error occurred in GetALLAutoTransactionExportSettings Exception is :" + ex.Message)
            Return Nothing
        Finally

        End Try
    End Function

    Public Function CheckTransReportSend(CompanyId As Integer, ExecutionDate As DateTime) As Integer
        Dim dal = New GeneralizedDAL()
        Try

            Dim ds As DataSet = New DataSet()

            Dim Param As SqlParameter() = New SqlParameter(1) {}

            Param(0) = New SqlParameter("@CompanyId", SqlDbType.Int)
            Param(0).Direction = ParameterDirection.Input
            Param(0).Value = CompanyId

            Param(1) = New SqlParameter("@ExecutionDate", SqlDbType.Date)
            Param(1).Direction = ParameterDirection.Input
            Param(1).Value = ExecutionDate

            Dim result As Integer = dal.ExecuteStoredProcedureGetInteger("usp_tt_AutoTransactionExportHistory_CheckTransReportSend", Param)

            Return result

        Catch ex As Exception
            log.Error("Error occurred in GetALLAutoTransactionExportSettings Exception is :" + ex.Message)
            Return 0
        Finally

        End Try
    End Function

	Public Function GetTransactionByCompanyId(CompanyId As Integer, DecimalPlace As Integer, DecimalType As Integer) As DataTable
		Dim dal = New GeneralizedDAL()
		Try

			Dim ds As DataSet = New DataSet()

			Dim Param As SqlParameter() = New SqlParameter(2) {}

			Param(0) = New SqlParameter("@CompanyId", SqlDbType.Int)
			Param(0).Direction = ParameterDirection.Input
			Param(0).Value = CompanyId

			Param(1) = New SqlParameter("@DecimalPlace", SqlDbType.Int)
			Param(1).Direction = ParameterDirection.Input
			Param(1).Value = DecimalPlace

			Param(2) = New SqlParameter("@DecimalType", SqlDbType.Int)
			Param(2).Direction = ParameterDirection.Input
			Param(2).Value = DecimalType

			ds = dal.ExecuteStoredProcedureGetDataSet("usp_tt_AutoTransactionExportSettings_GetTransactionByCompanyId", Param)

			Return ds.Tables(0)

		Catch ex As Exception
			log.Error("Error occurred in GetTransactionByCompanyId Exception is :" + ex.Message)
			Return Nothing
		Finally

		End Try
	End Function

	Public Function SaveUpdateAutoTransactionExportHistory(AutoTransactionExportHistoryId As Integer, CompanyId As Integer, ExecutionDate As DateTime, FileName As String, SendFileToEmail As Boolean, SendFileToFTP As Boolean) As Integer
        Try
            Dim result As Integer
            Dim dal = New GeneralizedDAL()
            Dim parcollection(5) As SqlParameter
            Dim ParAutoTransactionExportHistoryId = New SqlParameter("@AutoTransactionExportHistoryId", SqlDbType.Int)
            Dim ParCompanyId = New SqlParameter("@CompanyId", SqlDbType.Int)
            Dim ParExecutionDate = New SqlParameter("@ExecutionDate", SqlDbType.DateTime)
            Dim ParFileName = New SqlParameter("@FileName", SqlDbType.NVarChar, 2000)
            Dim ParSendFileToEmail = New SqlParameter("@SendFileToEmail", SqlDbType.Bit)
            Dim ParSendFileToFTP = New SqlParameter("@SendFileToFTP", SqlDbType.Bit)

            ParAutoTransactionExportHistoryId.Direction = ParameterDirection.Input
            ParCompanyId.Direction = ParameterDirection.Input
            ParExecutionDate.Direction = ParameterDirection.Input
            ParFileName.Direction = ParameterDirection.Input
            ParSendFileToEmail.Direction = ParameterDirection.Input
            ParSendFileToFTP.Direction = ParameterDirection.Input

            ParAutoTransactionExportHistoryId.Value = AutoTransactionExportHistoryId
            ParCompanyId.Value = CompanyId
            ParExecutionDate.Value = ExecutionDate
            ParFileName.Value = FileName
            ParSendFileToEmail.Value = SendFileToEmail
            ParSendFileToFTP.Value = SendFileToFTP

            parcollection(0) = ParAutoTransactionExportHistoryId
            parcollection(1) = ParCompanyId
            parcollection(2) = ParExecutionDate
            parcollection(3) = ParFileName
            parcollection(4) = ParSendFileToEmail
            parcollection(5) = ParSendFileToFTP

            result = dal.ExecuteStoredProcedureGetInteger("usp_tt_AutoTransactionExportHistory_InsertUpdate", parcollection)

            Return result
        Catch ex As Exception
            log.Error("Error occurred in SaveUpdateAutoTransactionExportHistory Exception is :" + ex.Message)
            Return 0

        Finally

        End Try

        Return 0

    End Function

    Public Function GetCustomizedExportTemplateById(CustomizedExportTemplateId As Integer) As DataSet
        Dim dal = New GeneralizedDAL()
        Try

            Dim ds As DataSet = New DataSet()


            Dim Param As SqlParameter() = New SqlParameter(0) {}

            Param(0) = New SqlParameter("@CustomizedExportTemplateId", SqlDbType.Int)
            Param(0).Direction = ParameterDirection.Input
            Param(0).Value = CustomizedExportTemplateId

            ds = dal.ExecuteStoredProcedureGetDataSet("usp_tt_CustomizedExport_GetCustomizedExportTemplateById", Param)

            Return ds

        Catch ex As Exception

            log.Error("Error occurred in GetCustomizedExportTemplateById Exception is :" + ex.Message)
            Return Nothing
        Finally

        End Try
    End Function

    Public Function UpdateExportedTransaction(TransactionIds As String) As Integer
        Dim dal = New GeneralizedDAL()
        Try

            Dim ds As DataSet = New DataSet()

            Dim Param As SqlParameter() = New SqlParameter(0) {}

            Param(0) = New SqlParameter("@TransactionIds", SqlDbType.NVarChar, 2000)
            Param(0).Direction = ParameterDirection.Input
            Param(0).Value = TransactionIds


            Dim result As Integer = dal.ExecuteStoredProcedureGetInteger("usp_tt_Transactions_UpdateExportedTransaction", Param)

            Return result

        Catch ex As Exception
            log.Error("Error occurred in UpdateExportedTransaction Exception is :" + ex.Message)
            Return 0
        Finally

        End Try
    End Function
End Class
