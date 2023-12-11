Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace CheckList
    <FormAttribute("65214", "Business Objects/SysReceiptFormProduction.b1f")>
    Friend Class SystemForm1
        Inherits SystemFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Dim StrSql As String
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.Matrix0 = CType(Me.GetItem("13").Specific, SAPbouiCOM.Matrix)
            Me.Button1 = CType(Me.GetItem("btncklist").Specific, SAPbouiCOM.Button)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler LoadAfter, AddressOf Me.Form_LoadAfter

        End Sub

        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub Button0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg)


        End Sub

        Private Sub OnCustomInitialize()
            Try

            Catch ex As Exception

            End Try
        End Sub

        Private Sub Form_LoadAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                objform = objaddon.objapplication.Forms.GetForm("65214", pVal.FormTypeCount)
            Catch ex As Exception

            End Try


        End Sub

        Private WithEvents Matrix0 As SAPbouiCOM.Matrix



        Private Sub Matrix0_KeyDownAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.KeyDownAfter
            Try
                If Matrix0.Columns.Item("1").Cells.Item(pVal.Row).Specific.String = "" Then Exit Sub
                If Matrix0.Columns.Item("U_GCheckList").Cells.Item(pVal.Row).Specific.String <> "" Then Exit Sub
                StrSql = objaddon.objglobalmethods.getSingleValue("Select 1 as ""CheckListFlag"" from OITM where ""ItemCode""='" & Matrix0.Columns.Item("1").Cells.Item(pVal.Row).Specific.String & "' and ifnull(""U_CheckList"",'')='Y'")
                If StrSql <> "1" Then Exit Sub
                Dim ProdEntry As String
                ProdEntry = objform.DataSources.DBDataSources.Item("OWOR").GetValue("DocEntry", 0)
                If pVal.ItemUID = "13" And pVal.CharPressed = 9 And pVal.ColUID = "U_GCheckList" Then
                    Dim activeform As New FrmCheckList
                    activeform.Show()
                    activeform.LoadCheckList(objform, pVal.Row, Matrix0.Columns.Item("1").Cells.Item(pVal.Row).Specific.String, ProdEntry)
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Matrix0_ChooseFromListBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix0.ChooseFromListBefore
            Try
                'BubbleEvent = False
            Catch ex As Exception

            End Try

        End Sub

        Private WithEvents Button1 As SAPbouiCOM.Button

        Private Sub Button1_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button1.ClickAfter
            Try
                If Matrix0.Columns.Item("U_GCheckList").Cells.Item(1).Specific.String = "" Then objaddon.objapplication.StatusBar.SetText("CheckList Not Found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) : Exit Sub
                Link_Value = Matrix0.Columns.Item("U_GCheckList").Cells.Item(1).Specific.String
                Dim activeform As New FrmCheckList
                activeform.Show()


            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            Try
                StrSql = objaddon.objglobalmethods.getSingleValue("Select 1 as ""CheckListFlag"" from OITM where ""ItemCode""='" & Matrix0.Columns.Item("1").Cells.Item(1).Specific.String & "' and ifnull(""U_CheckList"",'')='Y'")
                If StrSql <> "1" Then Exit Sub
                If Matrix0.Columns.Item("U_GCheckList").Cells.Item(1).Specific.String = "" Then
                    objaddon.objapplication.StatusBar.SetText("CheckList is Mandatory...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    BubbleEvent = False : Exit Sub
                End If

            Catch ex As Exception

            End Try

        End Sub

        Private Sub Matrix0_KeyDownBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix0.KeyDownBefore
            Try
                Select Case pVal.ColUID
                    Case "U_GCheckList"
                        If Matrix0.Columns.Item("U_GCheckList").Cells.Item(1).Specific.String = "" Then Exit Sub
                        BubbleEvent = False
                End Select

            Catch ex As Exception

            End Try

        End Sub
    End Class
End Namespace
