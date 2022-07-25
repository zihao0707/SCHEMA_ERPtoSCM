
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Configuration
Imports Oracle.ManagedDataAccess.Client
Imports System.Data
Imports System.IO
Imports System.Collections
Imports System.Diagnostics
Imports Oracle.ManagedDataAccess.Types

Public Class Form1
    Dim A As Int32 = 0
    Dim number As Int32 = 0
    Dim day As String
    Dim time As String
    Dim staf As String
    Dim timred As String
    Dim stafre As String
    Dim timre As String
    Dim DataChange As String = "N"
    '------------------採購單單頭SUT01_0000-----------------------------------------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles SUT01_0000.Click
        Label2.Text = "0"
        Label3.Text = "更新中..."
        Call update_SUT01_0000() 'ok
        Call update_SUT01_0100() 'ok
    End Sub

    Private Sub Button1_Click_2(sender As Object, e As EventArgs) Handles Button1.Click
        Call update_BDP080_G003()
        Call update_BDP080_G004()
    End Sub

    '------------------產品SUB01_0000-----------------------------------------------------------------
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles SUB01_0000.Click 'ok
        Label2.Text = 0
        Label3.Text = "更新中..."
        Call update_SUB01_0000()


    End Sub
    '------------------供應商SUB02_0000-----------------------------------------------------------------
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles SUB02_0000.Click 'ok
        Label2.Text = 0
        Label3.Text = "更新中..."
        Call update_SUB02_0000()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Not Get_Setting() Then
            Me.Close()
        End If

        Dim fileReader2 As System.IO.StreamReader
        fileReader2 = My.Computer.FileSystem.OpenTextFileReader("time.txt")
        timred = fileReader2.ReadLine() '上次更新時間

        Timer2.Interval = TextBox1.Text


        fileReader2.Close()

    End Sub


    Function update_SUT01_0000()

        Dim dtTmp1 As DataTable
        Dim sSql As String = ""
        Try
            dtTmp1 = Get_ErpDataTable("select  a.OS_NO as pur_code, a.CUS_NO as sup_code,
                                    CONVERT(varchar(100),a.OS_DD, 111)  as pur_date,
                                                '00' as pur_status,'已核准' as erp_status,'0000' as pur_version,
			                                    a.USR as usr_code,'' as dep_code,'' as ls_add,
		                                     case when b.CNT_MAN1 is NULL then '' else b.CNT_MAN1 end as ls_man,
		                                      case when a.CLS_REM is NULL then '' else a.CLS_REM end as tra_cond,
		                                       case when a.PAY_REM is NULL then '' else a.PAY_REM end as rec_cond,
		                                      case when a.INV_ID is NULL then '' else a.INV_ID end as inv_no,
                                                TAX_ID as tax_type,0 as tax_rate,
			                                    case when CUR_ID is NULL then '' else CUR_ID end as cur_type,
                                                0 as pur_amount,0 as pur_tax,0 as pur_total
                                          
                                                from [MF_POS] a 
                                                    left join [CUST] b on a.CUS_NO = b.CUS_NO
                                                    left join [CUST] c on a.CUS_NO = c.CUS_NO
                                                where a.CLS_ID ='F'  ")

            For ilop = 0 To dtTmp1.Rows.Count - 1

                sSql &= SetSQLArray("pur_code", dtTmp1.Rows(ilop).Item("pur_code"))
                sSql &= SetSQLArray("sup_code", dtTmp1.Rows(ilop).Item("sup_code"))
                sSql &= SetSQLArray("pur_date", dtTmp1.Rows(ilop).Item("pur_date"))
                sSql &= SetSQLArray("pur_status", dtTmp1.Rows(ilop).Item("pur_status"))
                sSql &= SetSQLArray("erp_status", dtTmp1.Rows(ilop).Item("erp_status"))
                sSql &= SetSQLArray("pur_version", dtTmp1.Rows(ilop).Item("pur_version"))
                sSql &= SetSQLArray("usr_code", dtTmp1.Rows(ilop).Item("usr_code"))
                sSql &= SetSQLArray("dep_code", dtTmp1.Rows(ilop).Item("dep_code"))
                sSql &= SetSQLArray("ls_add", dtTmp1.Rows(ilop).Item("ls_add"))
                sSql &= SetSQLArray("ls_man", dtTmp1.Rows(ilop).Item("ls_man"))
                sSql &= SetSQLArray("tra_cond", dtTmp1.Rows(ilop).Item("tra_cond"))
                sSql &= SetSQLArray("rec_cond", dtTmp1.Rows(ilop).Item("rec_cond"))
                sSql &= SetSQLArray("inv_no", dtTmp1.Rows(ilop).Item("inv_no"))
                sSql &= SetSQLArray("tax_type", dtTmp1.Rows(ilop).Item("tax_type"))
                sSql &= SetSQLArray("tax_rate", dtTmp1.Rows(ilop).Item("tax_rate"))
                sSql &= SetSQLArray("cur_type", dtTmp1.Rows(ilop).Item("cur_type"))
                sSql &= SetSQLArray("pur_amount", dtTmp1.Rows(ilop).Item("pur_amount"))
                sSql &= SetSQLArray("pur_total", dtTmp1.Rows(ilop).Item("pur_total"))
                sSql &= SetSQLArray("pur_tax", dtTmp1.Rows(ilop).Item("pur_tax"))

                If Chk_RelData("Select * from SUT01_0000 where pur_code='" & dtTmp1.Rows(ilop).Item("pur_code") & "'") Then
                    Get_InsertSql("SUT01_0000", sSql)
                    number += 1
                End If
                Label2.Text = number

            Next
            Label3.Text = "採購單單頭 " & vbCrLf & ""
        Catch ex As Exception
            DataChange = "Y"
            Call MsgBox(ex.ToString)
            Application.Exit()
        End Try
        Exit Function
    End Function

    Function update_SUT01_0100()

        Dim dtTmp1 As DataTable
        Dim sSql As String = ""

        Try
            dtTmp1 = Get_ErpDataTable("select a.OS_NO as pur_code,PRD_NO as pro_code,ITM as p_no,
                                            case when b.BAT_NO is NULL or b.BAT_NO ='' then '1' else b.BAT_NO end as lot_no,
                                            UNIT as pro_unit,CONVERT(varchar(100),a.EST_DD, 111) as sys_date ,
								             case when	UP is NULL then 0 else UP end  as pro_price,
                                            case when a.QTY is NULL then 0 else a.QTY end as pro_qty,
                                            case when a.REM is NULL then '' else a.REM end as cmemo,
									        case when  (a.QTY =0 or  a.UP=0 or a.QTY is NULL or  a.UP is NULL) then 0 
									        else a.QTY * a.UP end	as sub_total
                                    from [TF_POS] a left join [MF_POS] b on a.OS_NO = b.OS_NO
                                    where a.OS_NO in(select a.OS_NO as pur_code from [MF_POS] a 
                                                           left join [CUST] b on a.CUS_NO = b.CUS_NO
                                                           left join [CUST] c on a.CUS_NO = c.CUS_NO
                                                    where a.CLS_ID ='F') ")
            For ilop = 0 To dtTmp1.Rows.Count - 1
                sSql &= SetSQLArray("pur_code", dtTmp1.Rows(ilop).Item("pur_code"))
                sSql &= SetSQLArray("pro_code", dtTmp1.Rows(ilop).Item("pro_code"))
                sSql &= SetSQLArray("p_no", dtTmp1.Rows(ilop).Item("p_no"))
                sSql &= SetSQLArray("lot_no", dtTmp1.Rows(ilop).Item("lot_no"))
                sSql &= SetSQLArray("pro_unit", dtTmp1.Rows(ilop).Item("pro_unit"))
                sSql &= SetSQLArray("sys_date", dtTmp1.Rows(ilop).Item("sys_date"))
                sSql &= SetSQLArray("pro_qty", dtTmp1.Rows(ilop).Item("pro_qty"))
                sSql &= SetSQLArray("p_qty", dtTmp1.Rows(ilop).Item("pro_qty"))
                sSql &= SetSQLArray("pro_price", dtTmp1.Rows(ilop).Item("pro_price"))
                sSql &= SetSQLArray("sub_total", dtTmp1.Rows(ilop).Item("sub_total"))
                sSql &= SetSQLArray("cmemo", dtTmp1.Rows(ilop).Item("cmemo"))

                If Chk_RelData("select * FROM SUT01_0000 WHERE pur_code = '" & dtTmp1.Rows(ilop).Item("pur_code") & "'") = False Then
                    If Chk_RelData("select * FROM SUT01_0100 WHERE  pur_code + '-' + pro_code = '" &
                                   dtTmp1.Rows(ilop).Item("pur_code") & "-" &
                                   dtTmp1.Rows(ilop).Item("pro_code") &
                                   "' AND p_no = '" & dtTmp1.Rows(ilop).Item("p_no") & "'") Then
                        number += 1
                        Get_InsertSql("SUT01_0100", sSql)
                    End If
                End If

            Next
            Label2.Text = number
            Label3.Text += "採購單單身 " & vbCrLf & ""
        Catch ex As Exception
            DataChange = "Y"
            Call MsgBox(ex.ToString)
            Application.Exit()
        End Try
        Exit Function
    End Function

    Function update_SUT01_0120()

        Dim dtTmp1 As DataTable
        Dim sKey As String = ""
        Dim sSql As String = ""

        Try
            dtTmp1 = Get_ErpDataTable("select  a.OS_NO as pur_code
                                                from [MF_POS] a 
                                                    left join [CUST] b on a.CUS_NO = b.CUS_NO
                                                    left join [CUST] c on a.CUS_NO = c.CUS_NO
                                                where a.CLS_ID ='T' ")
            For ilop = 0 To dtTmp1.Rows.Count - 1

                sKey &= SetSQLArray("pur_code", dtTmp1.Rows(ilop).Item("pur_code"))
                sSql &= SetSQLArray("pur_version", dtTmp1.Rows(ilop).Item("pur_version"))

                Get_UpdateSql("SUT01_0000", sSql, sKey)

                number += 1
                Label2.Text = number


            Next
            Label3.Text += "採購變更單更新完成"
            number = 0
        Catch ex As Exception
            Call MsgBox(ex.ToString)
            Application.Exit()
        End Try
        Exit Function
    End Function

    Function update_SUB02_0000()

        Dim dtTmp1 As DataTable
        Dim sSql As String = ""
        Dim Remake As DataTable
        '///清空
        Try
            Remake = Get_DataTable("TRUNCATE TABLE SUB02_0000")
        Catch ex As Exception
            Call MsgBox(ex.ToString)
            Application.Exit()
        End Try

        Try
            dtTmp1 = Get_ErpDataTable("  select CUS_NO as sup_code,NAME as sup_name,
                             case when CNT_MAN1 is NULL then '' else CNT_MAN1 end as sup_man,
                                    case when TEL1 is NULL then '' else TEL1 end as sup_tel,
		                            case when E_MAIL is NULL then '' else E_MAIL end  as sup_mail,
		                            case when ADR2 is NULL then '' else ADR2 end as com_add,
                                    '' as sup_type,
									case when FAX is NULL then '' else FAX end as ls_fax,'' as rec_cond, 
                                    '' as cur_type
                            FROM [CUST] where OBJ_ID !='1'
")
            For ilop = 0 To dtTmp1.Rows.Count - 1

                sSql &= SetSQLArray("sup_code", dtTmp1.Rows(ilop).Item("sup_code"))
                sSql &= SetSQLArray("sup_name", dtTmp1.Rows(ilop).Item("sup_name"))
                sSql &= SetSQLArray("sup_tel ", dtTmp1.Rows(ilop).Item("sup_tel "))
                sSql &= SetSQLArray("com_add", dtTmp1.Rows(ilop).Item("com_add"))
                sSql &= SetSQLArray("sup_type", dtTmp1.Rows(ilop).Item("sup_type"))
                sSql &= SetSQLArray("rec_cond", dtTmp1.Rows(ilop).Item("rec_cond"))
                sSql &= SetSQLArray("sup_man", dtTmp1.Rows(ilop).Item("sup_man"))
                sSql &= SetSQLArray("sup_mail", dtTmp1.Rows(ilop).Item("sup_mail"))
                sSql &= SetSQLArray("ls_fax", dtTmp1.Rows(ilop).Item("ls_fax"))
                sSql &= SetSQLArray("cur_type", dtTmp1.Rows(ilop).Item("cur_type"))

                Get_InsertSql("SUB02_0000", sSql)

                number += 1
                Label2.Text = number

            Next
            Label3.Text = "供應商更新完成"
        Catch ex As Exception
            DataChange = "Y"
            Call MsgBox(ex.ToString)
            Application.Exit()
        End Try

        Exit Function
    End Function

    Function update_SUB01_0000()  '料件檔

        Dim number As Int32 = 0
        Dim dtTmp1 As DataTable
        Dim sSql As String = ""
        Dim Remake As DataTable

        '///清空
        Try
            Remake = Get_DataTable("TRUNCATE TABLE SUB01_0000")
        Catch ex As Exception
            Call MsgBox(ex.ToString)
            Application.Exit()
        End Try
        Try
            dtTmp1 = Get_ErpDataTable("select PRD_NO as pro_code,
                        case when SPC is NULL then '' else SPC end as pro_spec,
                        case when NAME is NULL then '' else NAME end as pro_name,0 as std_price,
                        '' as pro_unit
                        from [PRDT]")


            For ilop = 0 To dtTmp1.Rows.Count - 1

                sSql &= SetSQLArray("pro_code", dtTmp1.Rows(ilop).Item("pro_code"))
                sSql &= SetSQLArray("pro_spec", dtTmp1.Rows(ilop).Item("pro_spec"))
                sSql &= SetSQLArray("pro_name", dtTmp1.Rows(ilop).Item("pro_name"))
                sSql &= SetSQLArray("std_price", dtTmp1.Rows(ilop).Item("std_price"))
                sSql &= SetSQLArray("pro_unit", dtTmp1.Rows(ilop).Item("pro_unit"))
                Get_InsertSql("SUB01_0000", sSql)

                number += 1
                Label2.Text = number

            Next

            Label3.Text = "料件更新完成"
            number = 0
        Catch ex As Exception
            DataChange = "Y"
            Call MsgBox(ex.ToString)
            Application.Exit()
        End Try


        Exit Function
    End Function
    Function update_BDP080_G003() '採購人員帳號 廠商

        Dim number As Int32 = 0
        Dim dtTmp1 As DataTable
        Dim sSql As String = ""
        Dim Remake As DataTable

        Try
            Remake = Get_DataTable("DELETE FROM BDP080 WHERE grp_code = 'G003' ")
        Catch ex As Exception
            Call MsgBox(ex.ToString)
            Application.Exit()
        End Try

        Try
            dtTmp1 = Get_ErpDataTable("with usr as(
                                                select distinct a.USR as usr_code from [MF_POS] a 
                                                left join [CUST] b on a.CUS_NO = b.CUS_NO
                                                left join [CUST] c on a.CUS_NO = c.CUS_NO)

                            SELECT usr_code,usr_code AS usr_name,'151－151－860－254－337－515－272' AS usr_pass,
                                        'G003' AS grp_code ,'02-0000-0000' AS usr_tel1,
											  '02-0000-0000'  AS usr_tel2,
											  'a12345@gmail.com'   AS usr_mail
											  ,'B' AS limit_type,'Y' AS is_use
                                        FROM usr")
            For ilop = 1 To dtTmp1.Rows.Count - 1

                sSql &= SetSQLArray("usr_code", dtTmp1.Rows(ilop).Item("usr_code"))
                sSql &= SetSQLArray("usr_name", dtTmp1.Rows(ilop).Item("usr_name"))
                sSql &= SetSQLArray("usr_pass", dtTmp1.Rows(ilop).Item("usr_pass"))
                sSql &= SetSQLArray("grp_code", dtTmp1.Rows(ilop).Item("grp_code"))
                sSql &= SetSQLArray("usr_tel1", dtTmp1.Rows(ilop).Item("usr_tel1"))
                sSql &= SetSQLArray("usr_mail", dtTmp1.Rows(ilop).Item("usr_mail"))
                sSql &= SetSQLArray("limit_type", dtTmp1.Rows(ilop).Item("limit_type"))
                sSql &= SetSQLArray("is_use", dtTmp1.Rows(ilop).Item("is_use"))
                Get_InsertSql("BDP080", sSql)

                number += 1
                Label2.Text = number


            Next
            Label3.Text = "採購人員資料更新完成" & vbCrLf & ""
        Catch ex As Exception
            DataChange = "Y"
            Call MsgBox(ex.ToString)
            Application.Exit()
        End Try


        Exit Function
    End Function

    Function update_BDP080_G004() '供應商人員帳號  廠商客戶

        Dim number As Int32 = 0
        Dim dtTmp1 As DataTable
        Dim sSQL As String = ""
        Dim Remake As DataTable

        Try
            Remake = Get_DataTable("DELETE FROM BDP080 WHERE grp_code = 'G004' ")
        Catch ex As Exception
            Call MsgBox(ex.ToString)
            Application.Exit()
        End Try

        Try
            dtTmp1 = Get_ErpDataTable("SELECT  CUS_NO AS usr_code,NAME AS usr_name,'151－151－860－254－337－515－272' AS usr_pass,'G004' AS grp_code,
                                              case when TEL1 is NULL then '' else TEL1 end AS usr_tel1,
											  case when TEL2 is NULL then '' else TEL2 end  AS usr_tel2,
											  case when E_MAIL is NULL then '' else E_MAIL end   AS usr_mail,
											  'B' AS limit_type,'Y' AS is_use 
                                            FROM [CUST] where OBJ_ID !='1' ")
            For ilop = 1 To dtTmp1.Rows.Count - 1

                sSql &= SetSQLArray("usr_code", dtTmp1.Rows(ilop).Item("usr_code"))
                sSql &= SetSQLArray("usr_name", dtTmp1.Rows(ilop).Item("usr_name"))
                sSql &= SetSQLArray("usr_pass", dtTmp1.Rows(ilop).Item("usr_pass"))
                sSql &= SetSQLArray("grp_code", dtTmp1.Rows(ilop).Item("grp_code"))
                sSQL &= SetSQLArray("usr_tel1", dtTmp1.Rows(ilop).Item("usr_tel1"))
                sSQL &= SetSQLArray("usr_tel1", dtTmp1.Rows(ilop).Item("usr_tel2"))
                sSQL &= SetSQLArray("usr_mail", dtTmp1.Rows(ilop).Item("usr_mail"))
                sSql &= SetSQLArray("limit_type", dtTmp1.Rows(ilop).Item("limit_type"))
                sSQL &= SetSQLArray("is_use", dtTmp1.Rows(ilop).Item("is_use"))
                sSQL &= SetSQLArray("sup_code", dtTmp1.Rows(ilop).Item("usr_code"))
                Get_InsertSql("BDP080", sSql)
                number += 1
            Next
            Label2.Text = number
            Label3.Text += "供應商人員資料更新完成"
        Catch ex As Exception
            DataChange = "Y"
            Call MsgBox(ex.ToString)
            Application.Exit()
        End Try


        Exit Function
    End Function

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If (DataChange = "N") Then
            DataChange = "Y"
        Call update_SUT01_0000()
            Call update_SUT01_0100()
            Label3.Text = "採購單更新完成" & vbCrLf & "請購單更新完成"
        End If
        DataChange = "N"

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Dim dtTmp5 As DataTable
        Dim dtTmp6 As DataTable
        Dim update As String
        Dim par_name As String

        update = "Select par_name from BDP000 where par_value='Y'"
        dtTmp5 = Get_DataTable(update)
        If (dtTmp5.Rows.Count > 0 And DataChange = "N") Then
            For i = 0 To dtTmp5.Rows.Count - 1
                par_name = dtTmp5.Rows(i).Item("par_name")
                Select Case par_name
                    Case "sync_pro"
                        DataChange = "Y"
                        Call update_SUB01_0000()
                        dtTmp6 = Get_DataTable("UPDATE BDP000 SET par_value='N' WHERE par_name='" & par_name & "'")
                        DataChange = "N"
                    Case "sync_sup"
                        DataChange = "Y"
                        Call update_SUB02_0000()
                        dtTmp6 = Get_DataTable("UPDATE BDP000 SET par_value='N' WHERE par_name='" & par_name & "'")
                        DataChange = "N"
                    Case "sync_pur"
                        DataChange = "Y"
                        Call update_SUT01_0000()
                        Call update_SUT01_0100()
                        dtTmp6 = Get_DataTable("UPDATE BDP000 SET par_value='N' WHERE par_name='" & par_name & "'")
                        DataChange = "N"
                End Select
            Next
        End If


    End Sub


    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick

        day = Format(Now, "yyyy/MM/dd") '2016-1-10
        time = Format(Now, "hh:mm:ss") ' 23:38:20
        Label4.Text = day + " " + time


    End Sub

    Function Main()
        Dim sw As StreamWriter = New StreamWriter("time.txt")
        sw.WriteLine(Label4.Text)
        sw.Close()


        Exit Function
    End Function
End Class



