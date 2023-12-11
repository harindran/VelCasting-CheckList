Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace CheckList
    <FormAttribute("MICKLT", "Business Objects/FrmCheckList.b1f")>
    Friend Class FrmCheckList
        Inherits UserFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Private WithEvents pCFL As SAPbouiCOM.ISBOChooseFromListEventArg
        Dim FormCount As Integer = 0
        Private WithEvents objDBHeader As SAPbouiCOM.DBDataSource
        Private WithEvents odbdsDetails As SAPbouiCOM.DBDataSource
        Dim objRs As SAPbobsCOM.Recordset
        Dim StrSql As String
        Dim strCheckList As String = ""
        Private WithEvents objmatrix As SAPbouiCOM.Matrix
        Dim Row As Integer
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.StaticText0 = CType(Me.GetItem("lentry").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("tentry").Specific, SAPbouiCOM.EditText)
            Me.StaticText1 = CType(Me.GetItem("lprodnum").Specific, SAPbouiCOM.StaticText)
            Me.EditText1 = CType(Me.GetItem("tprodnum").Specific, SAPbouiCOM.EditText)
            Me.StaticText2 = CType(Me.GetItem("lnum").Specific, SAPbouiCOM.StaticText)
            Me.EditText2 = CType(Me.GetItem("tnum").Specific, SAPbouiCOM.EditText)
            Me.StaticText3 = CType(Me.GetItem("lproddate").Specific, SAPbouiCOM.StaticText)
            Me.EditText3 = CType(Me.GetItem("tproddate").Specific, SAPbouiCOM.EditText)
            Me.StaticText4 = CType(Me.GetItem("lremark").Specific, SAPbouiCOM.StaticText)
            Me.Folder0 = CType(Me.GetItem("fldrcont").Specific, SAPbouiCOM.Folder)
            Me.Folder1 = CType(Me.GetItem("fldr1").Specific, SAPbouiCOM.Folder)
            Me.Matrix0 = CType(Me.GetItem("mtxcont").Specific, SAPbouiCOM.Matrix)
            Me.EditText5 = CType(Me.GetItem("tremark").Specific, SAPbouiCOM.EditText)
            Me.LinkedButton0 = CType(Me.GetItem("lnkorder").Specific, SAPbouiCOM.LinkedButton)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler LoadAfter, AddressOf Me.Form_LoadAfter
            AddHandler DataLoadAfter, AddressOf Me.Form_DataLoadAfter
            AddHandler DataAddBefore, AddressOf Me.Form_DataAddBefore

        End Sub

        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()
            Try
                objform.Items.Item("tentry").Specific.String = objaddon.objglobalmethods.GetNextDocEntry_Value("@MICKLIST")
                objDBHeader = objform.DataSources.DBDataSources.Item("@MICKLIST")
                bModal = True
                If Link_Value <> "-1" Then
                    objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    EditText0.Item.Enabled = True
                    EditText0.Value = Link_Value
                    objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    'EditText3.Item.Enabled = False
                    'EditText0.Item.Enabled = False
                    Link_Value = "-1"
                End If
                objform.Settings.Enabled = True
                Matrix0.AutoResizeColumns()
                Folder0.Item.Click()

            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

#Region "Fields"
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents StaticText4 As SAPbouiCOM.StaticText
        Private WithEvents Folder0 As SAPbouiCOM.Folder
        Private WithEvents Folder1 As SAPbouiCOM.Folder
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents EditText5 As SAPbouiCOM.EditText
        Private WithEvents LinkedButton0 As SAPbouiCOM.LinkedButton

#End Region

        Private Sub Form_LoadAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                objform = objaddon.objapplication.Forms.GetForm("MICKLT", pVal.FormTypeCount)
            Catch ex As Exception
            End Try

        End Sub

        Public Sub LoadCheckList(ByVal RForm As SAPbouiCOM.Form, ByVal RowID As Integer, ByVal ItemCode As String, ByVal ProdOrderEntry As String)
            Try
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objmatrix = RForm.Items.Item("13").Specific
                Row = RowID
                objform.Left = RForm.Left + 20
                objform.Top = RForm.Top + 20
                If objaddon.HANA Then
                    StrSql = "Select T1.""U_Grade"",T1.""U_CheckList"" ""CheckList Name"",T1.""U_Min"",T1.""U_Max"" "
                    StrSql += vbCrLf + "from ""@MIGRADEM"" T0 join ""@MIGRADEM1"" T1 on T0.""Code""=T1.""Code"" where T0.""U_Grade""='" & ItemCode & "'"
                    StrSql += vbCrLf + "and T0.""Code""=(Select Max(cast(""Code"" as Integer)) from ""@MIGRADEM"")"
                Else
                    StrSql = "Select T1.U_Grade,T1.U_CheckList CheckList Name,T1.U_Min,T1.U_Max "
                    StrSql += vbCrLf + "from [@MIGRADEM] T0 join [@MIGRADEM1] T1 on T0.Code=T1.Code where T0.U_Grade='" & ItemCode & "'"
                    StrSql += vbCrLf + "and T0.Code=(Select Max(cast(Code as Integer)) from [@MIGRADEM])"
                End If
                objRs.DoQuery(StrSql)
                Dim i As Integer = 0
                If objRs.RecordCount > 0 Then
                    odbdsDetails = objform.DataSources.DBDataSources.Item("@MICKLIST1")
                    Matrix0.Clear()
                    odbdsDetails.Clear()
                    objform.Freeze(True)
                    objaddon.objapplication.StatusBar.SetText("Loading Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    While Not objRs.EoF
                        If objRs.Fields.Item("U_Grade").Value <> "" Then
                            Matrix0.AddRow()
                            'odbdsDetails.Clear()
                            Matrix0.GetLineData(Matrix0.VisualRowCount)
                            odbdsDetails.SetValue("LineId", 0, i + 1)
                            odbdsDetails.SetValue("U_Grade", 0, objRs.Fields.Item("U_Grade").Value.ToString)
                            odbdsDetails.SetValue("U_CheckList", 0, objRs.Fields.Item("CheckList Name").Value.ToString)
                            odbdsDetails.SetValue("U_Min", 0, objRs.Fields.Item("U_Min").Value.ToString)
                            odbdsDetails.SetValue("U_Max", 0, objRs.Fields.Item("U_Max").Value.ToString)
                            Matrix0.SetLineData(Matrix0.VisualRowCount)
                            i += 1
                        End If
                        objRs.MoveNext()
                    End While
                    objform.Freeze(False)

                    objRs = Nothing
                    objaddon.objapplication.StatusBar.SetText("Loaded Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Matrix0.AutoResizeColumns()
                Else
                    objaddon.objapplication.StatusBar.SetText("No records found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End If
                objform.Items.Item("tprodnum").Specific.String = ProdOrderEntry
                If objaddon.HANA Then
                    objform.Items.Item("tproddate").Specific.String = objaddon.objglobalmethods.getSingleValue("select To_Varchar(""PostDate"",'yyyyMMdd') from OWOR where ""DocEntry""='" & ProdOrderEntry & "'")
                Else
                    objform.Items.Item("tproddate").Specific.String = objaddon.objglobalmethods.getSingleValue("select Format(PostDate,'yyyyMMdd') from OWOR where DocEntry='" & ProdOrderEntry & "'")
                End If
                objform.Items.Item("tremark").Specific.String = "Created By " & objaddon.objcompany.UserName & " on " & Now.ToString("dd/MMM/yyyy HH:mm:ss")

            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

        Private Sub Form_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                objform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            Try
                If objform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Exit Sub
                Dim Input, Min, Max As Double
                For i As Integer = 1 To Matrix0.VisualRowCount
                    Input = IIf(Matrix0.Columns.Item("input").Cells.Item(i).Specific.String = "", 0, CDbl(Matrix0.Columns.Item("input").Cells.Item(i).Specific.String))
                    Min = IIf(Matrix0.Columns.Item("min").Cells.Item(i).Specific.String = "", 0, CDbl(Matrix0.Columns.Item("min").Cells.Item(i).Specific.String))
                    Max = IIf(Matrix0.Columns.Item("max").Cells.Item(i).Specific.String = "", 0, CDbl(Matrix0.Columns.Item("max").Cells.Item(i).Specific.String))
                    If Matrix0.Columns.Item("grade").Cells.Item(i).Specific.String <> "" Then
                        If Input < Min Or Input > Max Then
                            objaddon.objapplication.StatusBar.SetText("Input Data is not matching with the range... On Line " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False : Exit Sub
                        End If
                    End If
                Next
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_DataAddBefore(ByRef pVal As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
            Try
                objmatrix.Columns.Item("U_GCheckList").Cells.Item(Row).Specific.String = strCheckList
                If objmatrix.Columns.Item("U_GCheckList").Cells.Item(Row).Specific.String <> "" Then
                    objmatrix.Columns.Item("U_GCheckList").Editable = False
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button0_PressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.PressedAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ActionSuccess = True Then
                    If objmatrix.Columns.Item("U_GCheckList").Cells.Item(Row).Specific.String <> "" Then
                        objform.Close()
                    End If
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                strCheckList = objDBHeader.GetValue("DocEntry", 0)
            Catch ex As Exception

            End Try

        End Sub
    End Class
End Namespace
