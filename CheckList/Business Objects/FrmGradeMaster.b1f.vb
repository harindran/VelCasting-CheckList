Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace CheckList
    <FormAttribute("MIGRADE", "Business Objects/FrmGradeMaster.b1f")>
    Friend Class FrmGradeMaster
        Inherits UserFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Private WithEvents pCFL As SAPbouiCOM.ISBOChooseFromListEventArg
        Dim FormCount As Integer = 0
        Private WithEvents objDBHeader As SAPbouiCOM.DBDataSource
        Private WithEvents odbdsDetails As SAPbouiCOM.DBDataSource
        Dim objcombo As SAPbouiCOM.ComboBox
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.StaticText0 = CType(Me.GetItem("lcode").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("tcode").Specific, SAPbouiCOM.EditText)
            Me.StaticText1 = CType(Me.GetItem("lgrade").Specific, SAPbouiCOM.StaticText)
            Me.EditText1 = CType(Me.GetItem("tgrade").Specific, SAPbouiCOM.EditText)
            Me.Folder0 = CType(Me.GetItem("fldrcont").Specific, SAPbouiCOM.Folder)
            Me.Matrix0 = CType(Me.GetItem("mtxcont").Specific, SAPbouiCOM.Matrix)
            Me.Folder1 = CType(Me.GetItem("fldr1").Specific, SAPbouiCOM.Folder)
            Me.LinkedButton0 = CType(Me.GetItem("lnkgrade").Specific, SAPbouiCOM.LinkedButton)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler LoadAfter, AddressOf Me.Form_LoadAfter
            AddHandler DataLoadAfter, AddressOf Me.Form_DataLoadAfter

        End Sub

        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()
            Try
                'objform = objaddon.objapplication.Forms.GetForm("SUBBOM", Me.FormCount)
                'objform = objaddon.objapplication.Forms.ActiveForm
                objform.Freeze(True)
                'Dim objRs As SAPbobsCOM.Recordset
                objform.Items.Item("mtxcont").Enabled = False
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tcode", False, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tgrade", True, True, False)
                objDBHeader = objform.DataSources.DBDataSources.Item("@MIGRADEM")
                odbdsDetails = objform.DataSources.DBDataSources.Item("@MIGRADEM1")
                objform.ActiveItem = "tgrade"
                Matrix0.Columns.Item("checklist").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
                objform.Settings.Enabled = True
                objform.Freeze(False)
                Matrix0.AutoResizeColumns()
                Folder0.Item.Click()
            Catch ex As Exception
                'objform.Freeze(False)
            Finally
                objform.Freeze(False)
            End Try
        End Sub

#Region "Fields"
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents Folder0 As SAPbouiCOM.Folder
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents Folder1 As SAPbouiCOM.Folder
        Private WithEvents LinkedButton0 As SAPbouiCOM.LinkedButton

#End Region
        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            Try
                RemoveLastrow(Matrix0, "grade")
            Catch ex As Exception

            End Try

        End Sub

        Private Sub EditText1_ChooseFromListAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText1.ChooseFromListAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        'EditText2.Value = pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value
                        objDBHeader.SetValue("U_Grade", 0, pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value)
                    Catch ex As Exception
                    End Try

                End If
            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try

        End Sub

        Private Sub EditText1_ChooseFromListBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles EditText1.ChooseFromListBefore
            If pVal.ActionSuccess = True Then Exit Sub
            Try
                Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item("CFL_0")
                Dim oConds As SAPbouiCOM.Conditions
                Dim oCond As SAPbouiCOM.Condition
                Dim oEmptyConds As New SAPbouiCOM.Conditions
                oCFL.SetConditions(oEmptyConds)
                oConds = oCFL.GetConditions()

                oCond = oConds.Add()
                oCond.Alias = "validFor"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "Y"
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                oCond = oConds.Add()
                oCond.Alias = "U_CheckList"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "Y"

                oCFL.SetConditions(oConds)
            Catch ex As Exception
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try

        End Sub

        Private Sub Form_LoadAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                objform = objaddon.objapplication.Forms.GetForm("MIGRADE", pVal.FormTypeCount)
            Catch ex As Exception

            End Try


        End Sub

        Private Sub Matrix0_ChooseFromListBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix0.ChooseFromListBefore
            Dim oEmptyConds As New SAPbouiCOM.Conditions
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = objaddon.objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            Dim oConds As SAPbouiCOM.Conditions
            Dim oCond As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Try
                oCFL = objform.ChooseFromLists.Item("CFL_1")
                oCFL.SetConditions(oEmptyConds)
                oConds = oCFL.GetConditions()
                oCond = oConds.Add()
                oCond.Alias = "validFor"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "Y"
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                oCond = oConds.Add()
                oCond.Alias = "U_CheckList"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "Y"
                oCFL.SetConditions(oConds)
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Matrix0_ChooseFromListAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.ChooseFromListAfter
            Try
                If pVal.ColUID = "grade" And pVal.ActionSuccess = True Then
                    Try
                        'Dim cmbtype As SAPbouiCOM.ComboBox = Matrix0.Columns.Item("Type").Cells.Item(pVal.Row).Specific
                        'Dim UnitPrice As String = ""
                        If pVal.ActionSuccess = False Then Exit Sub
                        pCFL = pVal

                        If Not pCFL.SelectedObjects Is Nothing Then
                            Try
                                Matrix0.Columns.Item("grade").Cells.Item(pVal.Row).Specific.String = pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value
                            Catch ex As Exception
                            End Try
                            objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "grade", "#")
                        End If

                    Catch ex As Exception
                    End Try
                End If
                Matrix0.AutoResizeColumns()
            Catch ex As Exception
            End Try

        End Sub

        Private Sub RemoveLastrow(ByVal omatrix As SAPbouiCOM.Matrix, ByVal Columname_check As String)
            Try
                If omatrix.VisualRowCount = 0 Then Exit Sub
                If Columname_check.ToString = "" Then Exit Sub
                If omatrix.Columns.Item(Columname_check).Cells.Item(omatrix.VisualRowCount).Specific.string = "" Then
                    omatrix.DeleteRow(omatrix.VisualRowCount)
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Button0_PressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.PressedAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                If pVal.ActionSuccess = True And objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                    objform.Items.Item("tcode").Specific.String = objaddon.objglobalmethods.GetNextCode_Value("@MIGRADEM")
                    objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "grade", "#")
                    objcombo = Matrix0.Columns.Item("checklist").Cells.Item(Matrix0.VisualRowCount).Specific
                    objaddon.objglobalmethods.LoadCombo(objcombo, "Select ""Code"",""Name"" from ""@CHECK""")
                End If
            Catch ex As Exception
            End Try
        End Sub

        Private Sub Form_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                objcombo = Matrix0.Columns.Item("checklist").Cells.Item(Matrix0.VisualRowCount).Specific
                objaddon.objglobalmethods.LoadCombo(objcombo, "Select ""Code"",""Name"" from ""@CHECK""")
            Catch ex As Exception

            End Try


        End Sub

        Private Sub Matrix0_LostFocusAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.LostFocusAfter
            Try
                objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "grade", "#")
            Catch ex As Exception

            End Try

        End Sub
    End Class
End Namespace
