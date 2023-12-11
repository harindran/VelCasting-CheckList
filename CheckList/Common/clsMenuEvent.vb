Imports SAPbouiCOM
Namespace CheckList

    Public Class clsMenuEvent
        Dim objform As SAPbouiCOM.Form
        Dim objglobalmethods As New clsGlobalMethods

        Public Sub MenuEvent_For_StandardMenu(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                Select Case objaddon.objapplication.Forms.ActiveForm.TypeEx

                    Case "141", "170", "-170", "426", "-426", "392", "-392"
                        Default_Sample_MenuEvent(pVal, BubbleEvent)
                    Case "MIGRADE"
                        GradeMaster_MenuEvent(pVal, BubbleEvent)
                    Case "MICKLT"
                        CheckList_MenuEvent(pVal, BubbleEvent)
                End Select
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Default_Sample_MenuEvent(ByVal pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Dim oUDFForm As SAPbouiCOM.Form
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "6005"
                            If objaddon.objapplication.Forms.ActiveForm.Items.Item("chkactive").Specific.Checked = True And objaddon.objapplication.Forms.ActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                BubbleEvent = False
                            End If
                        Case "6913"
                            If objform.TypeEx = "392" Then
                                oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                                oUDFForm.Items.Item("U_TransId").Enabled = False
                                oUDFForm.Items.Item("U_IEntry").Enabled = False
                                oUDFForm.Items.Item("U_OEntry").Enabled = False
                                oUDFForm.Items.Item("U_IntRecNo").Enabled = False
                            End If
                    End Select
                Else

                    oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                    Select Case pval.MenuUID
                        Case "1281" 'Find
                            If objform.TypeEx = "141" Then
                                oUDFForm.Items.Item("U_MBAPNo").Enabled = True
                            ElseIf objform.TypeEx = "170" Or objform.TypeEx = "426" Then
                                oUDFForm.Items.Item("U_JENo").Enabled = True
                                oUDFForm.Items.Item("U_Select").Enabled = True
                                objform.Items.Item("tjeno").Visible = False
                                objform.Items.Item("ljeno").Visible = False
                            ElseIf objform.TypeEx = "392" Then
                                oUDFForm.Items.Item("U_TransId").Enabled = True
                                oUDFForm.Items.Item("U_IEntry").Enabled = True
                                oUDFForm.Items.Item("U_OEntry").Enabled = True
                                oUDFForm.Items.Item("U_IntRecNo").Enabled = True
                            End If
                        Case "1287" 'Duplicate
                            If oUDFForm.Items.Item("U_MBAPNo").Enabled = False Then
                                oUDFForm.Items.Item("U_MBAPNo").Enabled = True
                            End If
                            oUDFForm.Items.Item("U_MBAPNo").Specific.String = ""
                        Case "1282"
                            If objform.TypeEx = "141" Then
                                oUDFForm.Items.Item("U_MBAPNo").Enabled = False
                            ElseIf objform.TypeEx = "170" Or objform.TypeEx = "426" Then
                                oUDFForm.Items.Item("U_JENo").Enabled = False
                                objform.Items.Item("tjeno").Visible = False
                                objform.Items.Item("ljeno").Visible = False
                                objform.Items.Item("chkactive").Visible = False
                            ElseIf objform.TypeEx = "392" Then
                                oUDFForm.Items.Item("U_TransId").Enabled = False
                                oUDFForm.Items.Item("U_IEntry").Enabled = False
                                oUDFForm.Items.Item("U_OEntry").Enabled = False
                                oUDFForm.Items.Item("U_IntRecNo").Enabled = False
                            End If

                        Case Else
                            If objform.TypeEx = "141" Then
                                oUDFForm.Items.Item("U_MBAPNo").Enabled = False
                            ElseIf objform.TypeEx = "170" Or objform.TypeEx = "426" Then
                                oUDFForm.Items.Item("U_JENo").Enabled = False
                                oUDFForm.Items.Item("U_Select").Enabled = False
                            ElseIf objform.TypeEx = "392" Then
                                oUDFForm.Items.Item("U_TransId").Enabled = False
                                oUDFForm.Items.Item("U_IEntry").Enabled = False
                                oUDFForm.Items.Item("U_OEntry").Enabled = False
                                oUDFForm.Items.Item("U_IntRecNo").Enabled = False
                            End If
                    End Select
                End If
            Catch ex As Exception
                'objaddon.objapplication.SetStatusBarMessage("Error in Standart Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

#Region "Grade Master & CheckList"

        Private Sub GradeMaster_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0 As SAPbouiCOM.Matrix
            Dim odbdsDetails As SAPbouiCOM.DBDataSource
            Dim FolderItem As SAPbouiCOM.Folder
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                odbdsDetails = objform.DataSources.DBDataSources.Item("@MIGRADEM1")
                Matrix0 = objform.Items.Item("mtxcont").Specific
                FolderItem = objform.Items.Item("fldrcont").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            'objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            'BubbleEvent = False
                            If objaddon.objapplication.MessageBox("Removal of an entry cannot be reversed. Do you want to continue?", 2, "Yes", "No") <> 1 Then BubbleEvent = False : Exit Sub
                        Case "1293"
                            If FolderItem.Selected Then
                                If Matrix0.VisualRowCount = 1 Then BubbleEvent = False

                            End If

                        Case "1292"

                    End Select
                Else
                    Dim DBSource As SAPbouiCOM.DBDataSource
                    DBSource = objform.DataSources.DBDataSources.Item("@MIGRADEM")
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode                           
                            objform.Items.Item("tcode").Enabled = True
                            objform.Items.Item("mtxcont").Enabled = False
                            objform.Items.Item("tgrade").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Case "1282" ' Add Mode
                            objform.Items.Item("tcode").Specific.String = objaddon.objglobalmethods.GetNextCode_Value("@MIGRADEM")
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "grade", "#")
                            Dim objcombo As SAPbouiCOM.ComboBox
                            objcombo = Matrix0.Columns.Item("checklist").Cells.Item(Matrix0.VisualRowCount).Specific
                            objaddon.objglobalmethods.LoadCombo(objcombo, "Select ""Code"",""Name"" from ""@CHECK""")

                        Case "1288", "1289", "1290", "1291"
                            'objform.Items.Item("btngendoc").Enabled = True
                            objaddon.objapplication.Menus.Item("1300").Activate()
                        Case "1293"
                            If FolderItem.Selected Then
                                DeleteRow(Matrix0, "@MIGRADEM1")
                            End If
                            If objform.Mode = BoFormMode.fm_OK_MODE Then objform.Mode = BoFormMode.fm_UPDATE_MODE
                            objform.Update()
                            objform.Refresh()
                        Case "1292"
                            Try
                                If FolderItem.Selected Then
                                    If Matrix0.VisualRowCount > 0 Then
                                        If odbdsDetails.GetValue("U_Grade", Matrix0.VisualRowCount - 1) = "" Then Exit Sub
                                        objform.Freeze(True)
                                        odbdsDetails.InsertRecord(odbdsDetails.Size)
                                        odbdsDetails.SetValue("LineId", Matrix0.VisualRowCount, Matrix0.VisualRowCount + 1)
                                        Matrix0.LoadFromDataSource()
                                        objform.Freeze(False)
                                    End If
                                End If
                            Catch ex As Exception
                            End Try
                        Case "1287"  'Duplicate
                            objform.Items.Item("tgrade").Specific.String = ""
                            'objform.Items.Item("txtcode").Specific.String = ""
                            'objform.Items.Item("txtname").Specific.String = ""
                            'objform.Items.Item("txtentry").Specific.String = ""

                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                'objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

        Private Sub CheckList_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Dim Matrix0 As SAPbouiCOM.Matrix
            Dim odbdsDetails As SAPbouiCOM.DBDataSource
            Dim FolderItem As SAPbouiCOM.Folder
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                odbdsDetails = objform.DataSources.DBDataSources.Item("@MICKLIST1")
                Matrix0 = objform.Items.Item("mtxcont").Specific
                FolderItem = objform.Items.Item("fldrcont").Specific
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                            'If objaddon.objapplication.MessageBox("Removal of an entry cannot be reversed. Do you want to continue?", 2, "Yes", "No") <> 1 Then BubbleEvent = False : Exit Sub
                        Case "1293"
                            If FolderItem.Selected Then
                                If Matrix0.VisualRowCount = 1 Then BubbleEvent = False
                            End If
                    End Select
                Else
                    Dim DBSource As SAPbouiCOM.DBDataSource
                    DBSource = objform.DataSources.DBDataSources.Item("@MICKLIST")
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode                           
                            objform.Items.Item("tentry").Enabled = True
                            objform.Items.Item("tnum").Enabled = True
                            objform.Items.Item("tprodnum").Enabled = True
                            objform.Items.Item("tproddate").Enabled = True
                            objform.Items.Item("mtxcont").Enabled = False
                            objform.Items.Item("tentry").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Case "1282" ' Add Mode


                        Case "1288", "1289", "1290", "1291"
                            'objform.Items.Item("btngendoc").Enabled = True
                            objaddon.objapplication.Menus.Item("1300").Activate()
                        Case "1293"
                            If FolderItem.Selected Then
                                DeleteRow(Matrix0, "@MICKLIST1")
                            End If
                            If objform.Mode = BoFormMode.fm_OK_MODE Then objform.Mode = BoFormMode.fm_UPDATE_MODE
                            objform.Update()
                            objform.Refresh()
                        Case "1292"

                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                'objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub
#End Region

        Sub DeleteRow(ByVal objMatrix As SAPbouiCOM.Matrix, ByVal TableName As String)
            Try
                Dim DBSource As SAPbouiCOM.DBDataSource
                'objMatrix = objform.Items.Item("20").Specific
                objMatrix.FlushToDataSource()
                DBSource = objform.DataSources.DBDataSources.Item(TableName) '"@MIREJDET1"
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.GetLineData(i)
                    DBSource.Offset = i - 1
                    DBSource.SetValue("LineId", DBSource.Offset, i)
                    objMatrix.SetLineData(i)
                    objMatrix.FlushToDataSource()
                Next
                DBSource.RemoveRecord(DBSource.Size - 1)
                objMatrix.LoadFromDataSource()

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("DeleteRow  Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Finally
            End Try
        End Sub
    End Class
End Namespace