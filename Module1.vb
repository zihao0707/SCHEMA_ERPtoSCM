Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Module Module1
    '資料庫
    Public pConn As String
    Public pErp As String


    '紀錄SQL語法
    Public pRecSQL As String = "N"
    '發票機
    Public pInvoiceBaudRate As String
    Public pInvoiceParity As String
    Public pInvoiceDataBits As String
    Public pInvoiceStopBits As String
    Public pInvoicePortName As String

    Public Function Get_Setting() As Boolean
        '取得連線字串
        Try
            Dim objReader As New StreamReader(Application.StartupPath & "\setting.ini")

            'DB連線字串
            pConn = ChangeString(objReader.ReadLine(), "2")
            pErp = ChangeString(objReader.ReadLine(), "2")

            '紀錄SQL語法
            pRecSQL = objReader.ReadLine()

            '發票機COM port設定值
            'pInvoiceBaudRate = objReader.ReadLine()
            'pInvoiceParity = objReader.ReadLine()
            'pInvoiceDataBits = objReader.ReadLine()
            'pInvoiceStopBits = objReader.ReadLine()
            'pInvoicePortName = objReader.ReadLine()

            objReader.Close()
            Return True

        Catch ex As Exception
            Write_Log("找不到設定檔[setting.ini]")
            Return False
        End Try
    End Function

    Public Function ConnectDB() As SqlConnection
        ConnectDB = Nothing
        Try
            ConnectDB = New SqlConnection(pConn)
            ConnectDB.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function


    Public Function ERPDB() As SqlConnection
        ERPDB = Nothing
        Try
            ERPDB = New SqlConnection(pErp)
            ERPDB.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Public Function Get_ErpDataTable(sSql As String) As DataTable
        Get_ErpDataTable = New DataTable
        Try
            If sSql <> "" Then
                '建立連線
                Dim Con_db As SqlConnection = ERPDB()

                Using (Con_db)
                    Dim Fun_Adpt As New SqlDataAdapter(sSql, Con_db)
                    Fun_Adpt.Fill(Get_ErpDataTable)
                    Con_db.Close()
                End Using
            End If
        Catch ex As Exception
            'MsgBox(ex.Message)

            Using outfile As New StreamWriter(Application.StartupPath & "\error.txt", True)
                outfile.Write(sSql & vbNewLine)
            End Using
        End Try
    End Function

    Public Function Get_DataTable(sSql As String) As DataTable
        Get_DataTable = New DataTable
        Try
            If sSql <> "" Then
                '建立連線
                Dim Con_db As SqlConnection = ConnectDB()

                Using (Con_db)
                    Dim Fun_Adpt As New SqlDataAdapter(sSql, Con_db)
                    Fun_Adpt.Fill(Get_DataTable)
                    Con_db.Close()
                End Using
            End If
        Catch ex As Exception
            'MsgBox(ex.Message)

            Using outfile As New StreamWriter(Application.StartupPath & "\error.txt", True)
                outfile.Write(sSql & vbNewLine)
            End Using
        End Try
    End Function

    '------------------

    ''' <summary>
    ''' 設定SQL欄位
    ''' </summary>
    ''' <param name="pFieldCode">欄位編號</param>
    ''' <param name="pFieldValue">欄位值</param>
    ''' <param name="pFieldType">欄位類別 (空白:無特殊處理 A:SQL語法 B:加密)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SetSQLArray(pFieldCode As String, pFieldValue As String, Optional pFieldType As String = "") As String
        Return "^#|^#" & pFieldCode & "&$,&$" & pFieldValue & "&$,&$" & pFieldType
    End Function

    ''' <summary>
    ''' 取得新增SQL語法
    ''' </summary>
    ''' <param name="pTableCode">Table名稱</param>
    ''' <param name="pData">欄位與資料的陣列字串</param>
    ''' <returns>SQL字串</returns>
    ''' <remarks></remarks>
    Public Function Get_InsertSql(pTableCode As String, pData As String) As String
        If pTableCode = "" Or pData = "" Then Return ""

        Dim sField As String = ""
        For Each row As DataRow In SQLArrayToDataTable(pData).Rows
            sField &= IIf(sField = "", "", ",") & row(0).ToString.Trim
        Next

        Dim sValue As String = ""
        For Each row As DataRow In SQLArrayToDataTable(pData).Rows
            sValue &= IIf(sValue = "", "N'", ",N'") & RepSql(row(1).ToString.Trim) & "'"
        Next

        Return "INSERT INTO " & RepSql(pTableCode) & "(" & sField & ")VALUES(" & sValue & ");"
    End Function

    '--------------------------------------共用資料檢查類別/CMS.PageBase.Check--------------
    ''' <summary>
    ''' 傳入一Sql語法，檢查是否有重覆資料 [true:沒有重覆 false:有重覆]
    ''' </summary>
    ''' <param name="pSql">Sql語法</param>
    ''' <returns>true:沒有重覆 false:有重覆</returns>
    ''' <remarks></remarks>

    Public Function Chk_RelData(pSql As String) As Boolean
        Dim dtFun As DataTable = Get_DataTable(pSql)
        If dtFun.Rows.Count > 0 Then Return False Else Return True
    End Function


    ''' <summary>
    ''' 取得修改SQL語法
    ''' </summary>
    ''' <param name="pTableCode">Table名稱</param>
    ''' <param name="pData">修改欄位與資料的陣列字串</param>
    ''' <param name="pKey">KEY欄位與資料的陣列字串</param>
    ''' <returns>SQL字串</returns>
    ''' <remarks></remarks>
    Public Function Get_UpdateSql(pTableCode As String, pData As String, pKey As String) As String
        If pTableCode = "" Or pData = "" Or pKey = "" Then Return ""

        Dim sSet As String = ""
        For Each row As DataRow In SQLArrayToDataTable(pData).Rows

            sSet &= IIf(sSet = "", "", ",") & row(0).ToString.Trim & "=N'" & RepSql(row(1).ToString.Trim) & "'"

        Next
        Dim sKey As String = ""
        For Each row As DataRow In SQLArrayToDataTable(pKey).Rows
            sKey &= IIf(sKey = "", "", " AND ") & row(0).ToString.Trim & "=N'" & RepSql(row(1).ToString.Trim) & "'"
        Next

        Return "UPDATE " & RepSql(pTableCode) & " SET " & sSet & " WHERE " & sKey & ";"
    End Function

    Private Function SQLArrayToDataTable(pSQLArray As String) As DataTable
        Dim dtSQL As New DataTable
        dtSQL.Columns.Add("field_code", System.Type.GetType("System.String"))
        dtSQL.Columns.Add("field_name", System.Type.GetType("System.String"))
        dtSQL.Columns.Add("field_type", System.Type.GetType("System.String"))

        Dim sXArray As String() = Split(pSQLArray, "^#|^#")

        For i As Integer = 1 To sXArray.Length - 1
            Dim sYArray As String() = Split(sXArray(i), "&$,&$")
            If sYArray.Length = 3 Then
                Dim drow As System.Data.DataRow
                drow = dtSQL.NewRow
                drow("field_code") = sYArray(0).ToString.Trim
                drow("field_name") = sYArray(1).ToString.Trim
                drow("field_type") = sYArray(2).ToString.Trim
                dtSQL.Rows.Add(drow)
            End If
        Next
        Return dtSQL
    End Function

    Public Function Get_DataStr(sSql As String) As String
        Get_DataStr = ""
        Try
            If sSql <> "" Then
                '建立連線
                Dim Con_db As SqlConnection = ConnectDB()

                Using (Con_db)
                    Dim Fun_Adpt As New SqlCommand(sSql, Con_db)
                    Get_DataStr = Fun_Adpt.ToString()
                    Con_db.Close()
                End Using
            End If
        Catch ex As Exception

            Using outfile As New StreamWriter(Application.StartupPath & "\error.txt", True)
                outfile.Write(sSql & vbNewLine)
            End Using
        End Try

        Return Get_DataStr
    End Function

    Public Function SaveData(sSql As String) As Boolean
        If String.IsNullOrEmpty(sSql) = True Then Return False

        Dim Con_db As SqlConnection = ConnectDB()
        Try
            Using (Con_db)
                Dim Sql_cmd As New SqlCommand(sSql, Con_db)
                Sql_cmd.ExecuteNonQuery()
                SaveData = True
                Con_db.Close()
            End Using
        Catch ex As Exception
            'MsgBox(ex.Message)

            Using outfile As New StreamWriter(Application.StartupPath & "\error.txt", True)
                outfile.Write(sSql & vbNewLine)
            End Using
            Return False
        End Try
    End Function

    Public Function RepSql(pStr As String) As String
        If String.IsNullOrEmpty(pStr) = True Then Return ""
        '過濾攻擊字元
        Dim sStr As String = Regex.Replace(pStr, "\b(exec(ute)?|select|update|insert|delete|drop|create)\b|[;]|(-{2})|(/\*.*\*/)", String.Empty, RegexOptions.IgnoreCase)
        Return Replace(sStr.Trim, "'", "''")
    End Function

    Public Sub Write_Log(pStatus As String, Optional pMsg As String = "", Optional lbTmp As Label = Nothing)
        If pMsg <> "" Then pMsg = " - " & pMsg
        Dim sMsg As String = Format(Now, "yyyy/MM/dd HH:mm:ss") & " - " & pStatus & pMsg
        Dim sb As New StringBuilder()
        sb.Append(sMsg & vbNewLine)
        Using outfile As New StreamWriter(Application.StartupPath & "\log_" & Format(Now, "yyMM") & ".txt", True)
            outfile.Write(sb.ToString())
        End Using

        If lbTmp IsNot Nothing Then
            lbTmp.Text = sMsg
        End If
    End Sub

    Public Function Get_Space(pCnt As Integer) As String
        Get_Space = ""
        For i As Integer = 1 To pCnt
            Get_Space &= " "
        Next
    End Function
    '變更日期格式
    Public Function Get_date(sys_date As String) As String

        Dim YY As String = Mid(sys_date, 1, 4)
        Dim MM1 As String = Mid(sys_date, 1, 6)
        Dim MM As String = Mid(MM1, 5, 6)
        Dim DD As String = Mid(sys_date, 7, 8)
        sys_date = YY + "/" + MM + "/" + DD
        Return sys_date
    End Function

    '金錢格式化
    Public Function FormatPrice(pAmt As Integer) As String
        Return pAmt.ToString("###,##0").PadLeft(7)
    End Function
    Public Function ChangeString(ByVal C_Str As String, ByVal pCtype As String) As String
        Dim Ilop As Integer
        Select Case pCtype
            Case "1" '編碼用
                For Ilop = 1 To Len(C_Str)
                    If (Ilop Mod 2) = 0 Then
                        ChangeString = ChangeString & ChangeChar1(Mid(C_Str, Ilop, 1))
                    Else
                        ChangeString = ChangeString & ChangeChar3(Mid(C_Str, Ilop, 1))
                    End If
                Next
            Case "2" '解碼用
                For Ilop = 1 To Len(C_Str)
                    If (Ilop Mod 2) = 0 Then
                        ChangeString = ChangeString & ChangeChar2(Mid(C_Str, Ilop, 1))
                    Else
                        ChangeString = ChangeString & ChangeChar4(Mid(C_Str, Ilop, 1))
                    End If
                Next
        End Select
    End Function
    Private Function ChangeChar4(ByVal C_bit As String) As String
        Select Case C_bit
            Case "n" : ChangeChar4 = "!"
            Case "F" : ChangeChar4 = "#"
            Case "p" : ChangeChar4 = "$"
            Case "A" : ChangeChar4 = "%"
            Case "w" : ChangeChar4 = "&"
            Case "1" : ChangeChar4 = "("
            Case "]" : ChangeChar4 = ")"
            Case "f" : ChangeChar4 = "*"
            Case "E" : ChangeChar4 = "+"
            Case "5" : ChangeChar4 = ","
            Case "L" : ChangeChar4 = "-"
            Case "r" : ChangeChar4 = "."
            Case "T" : ChangeChar4 = "/"
            Case "i" : ChangeChar4 = "0"
            Case "H" : ChangeChar4 = "1"
            Case "<" : ChangeChar4 = "2"
            Case "@" : ChangeChar4 = "3"
            Case "{" : ChangeChar4 = "4"
            Case "c" : ChangeChar4 = "5"
            Case "B" : ChangeChar4 = "6"
            Case "k" : ChangeChar4 = "7"
            Case "^" : ChangeChar4 = "8"
            Case "y" : ChangeChar4 = "9"
            Case "W" : ChangeChar4 = ":"
            Case "+" : ChangeChar4 = ";"
            Case "6" : ChangeChar4 = "<"
            Case "u" : ChangeChar4 = "="
            Case "J" : ChangeChar4 = ">"
            Case "C" : ChangeChar4 = "?"
            Case "P" : ChangeChar4 = "@"
            Case "t" : ChangeChar4 = "A"
            Case "}" : ChangeChar4 = "B"
            Case "[" : ChangeChar4 = "C"
            Case "K" : ChangeChar4 = "D"
            Case "?" : ChangeChar4 = "E"
            Case "e" : ChangeChar4 = "F"
            Case "j" : ChangeChar4 = "G"
            Case "U" : ChangeChar4 = "H"
            Case "2" : ChangeChar4 = "I"
            Case ">" : ChangeChar4 = "J"
            Case "a" : ChangeChar4 = "K"
            Case "=" : ChangeChar4 = "L"
            Case "o" : ChangeChar4 = "M"
            Case "9" : ChangeChar4 = "N"
            Case "." : ChangeChar4 = "O"
            Case "(" : ChangeChar4 = "P"
            Case "_" : ChangeChar4 = "Q"
            Case "v" : ChangeChar4 = "R"
            Case "3" : ChangeChar4 = "S"
            Case "!" : ChangeChar4 = "T"
            Case "h" : ChangeChar4 = "U"
            Case "~" : ChangeChar4 = "V"
            Case "b" : ChangeChar4 = "W"
            Case "`" : ChangeChar4 = "X"
            Case "7" : ChangeChar4 = "Y"
            Case "*" : ChangeChar4 = "Z"
            Case "G" : ChangeChar4 = "["
            Case "l" : ChangeChar4 = "\"
            Case "R" : ChangeChar4 = "]"
            Case "m" : ChangeChar4 = "^"
            Case "s" : ChangeChar4 = "_"
            Case "0" : ChangeChar4 = "`"
            Case "-" : ChangeChar4 = "a"
            Case "\" : ChangeChar4 = "b"
            Case "Z" : ChangeChar4 = "c"
            Case ":" : ChangeChar4 = "d"
            Case "#" : ChangeChar4 = "e"
            Case "I" : ChangeChar4 = "f"
            Case "D" : ChangeChar4 = "g"
            Case "," : ChangeChar4 = "h"
            Case "N" : ChangeChar4 = "i"
            Case "O" : ChangeChar4 = "j"
            Case ";" : ChangeChar4 = "k"
            Case "%" : ChangeChar4 = "l"
            Case "Q" : ChangeChar4 = "m"
            Case "z" : ChangeChar4 = "n"
            Case "&" : ChangeChar4 = "o"
            Case "x" : ChangeChar4 = "p"
            Case "S" : ChangeChar4 = "q"
            Case ")" : ChangeChar4 = "r"
            Case "Y" : ChangeChar4 = "s"
            Case "4" : ChangeChar4 = "t"
            Case "M" : ChangeChar4 = "u"
            Case "/" : ChangeChar4 = "v"
            Case "8" : ChangeChar4 = "w"
            Case "g" : ChangeChar4 = "x"
            Case "X" : ChangeChar4 = "y"
            Case "V" : ChangeChar4 = "z"
            Case "d" : ChangeChar4 = "{"
            Case "q" : ChangeChar4 = "}"
            Case " " : ChangeChar4 = "~"
            Case "$" : ChangeChar4 = " "
            Case Else
                ChangeChar4 = C_bit
        End Select
    End Function
    Private Function ChangeChar3(ByVal C_bit As String) As String
        Select Case C_bit
            Case "!" : ChangeChar3 = "n"
            Case "#" : ChangeChar3 = "F"
            Case "$" : ChangeChar3 = "p"
            Case "%" : ChangeChar3 = "A"
            Case "&" : ChangeChar3 = "w"
            Case "(" : ChangeChar3 = "1"
            Case ")" : ChangeChar3 = "]"
            Case "*" : ChangeChar3 = "f"
            Case "+" : ChangeChar3 = "E"
            Case "," : ChangeChar3 = "5"
            Case "-" : ChangeChar3 = "L"
            Case "." : ChangeChar3 = "r"
            Case "/" : ChangeChar3 = "T"
            Case "0" : ChangeChar3 = "i"
            Case "1" : ChangeChar3 = "H"
            Case "2" : ChangeChar3 = "<"
            Case "3" : ChangeChar3 = "@"
            Case "4" : ChangeChar3 = "{"
            Case "5" : ChangeChar3 = "c"
            Case "6" : ChangeChar3 = "B"
            Case "7" : ChangeChar3 = "k"
            Case "8" : ChangeChar3 = "^"
            Case "9" : ChangeChar3 = "y"
            Case ":" : ChangeChar3 = "W"
            Case ";" : ChangeChar3 = "+"
            Case "<" : ChangeChar3 = "6"
            Case "=" : ChangeChar3 = "u"
            Case ">" : ChangeChar3 = "J"
            Case "?" : ChangeChar3 = "C"
            Case "@" : ChangeChar3 = "P"
            Case "A" : ChangeChar3 = "t"
            Case "B" : ChangeChar3 = "}"
            Case "C" : ChangeChar3 = "["
            Case "D" : ChangeChar3 = "K"
            Case "E" : ChangeChar3 = "?"
            Case "F" : ChangeChar3 = "e"
            Case "G" : ChangeChar3 = "j"
            Case "H" : ChangeChar3 = "U"
            Case "I" : ChangeChar3 = "2"
            Case "J" : ChangeChar3 = ">"
            Case "K" : ChangeChar3 = "a"
            Case "L" : ChangeChar3 = "="
            Case "M" : ChangeChar3 = "o"
            Case "N" : ChangeChar3 = "9"
            Case "O" : ChangeChar3 = "."
            Case "P" : ChangeChar3 = "("
            Case "Q" : ChangeChar3 = "_"
            Case "R" : ChangeChar3 = "v"
            Case "S" : ChangeChar3 = "3"
            Case "T" : ChangeChar3 = "!"
            Case "U" : ChangeChar3 = "h"
            Case "V" : ChangeChar3 = "~"
            Case "W" : ChangeChar3 = "b"
            Case "X" : ChangeChar3 = "`"
            Case "Y" : ChangeChar3 = "7"
            Case "Z" : ChangeChar3 = "*"
            Case "[" : ChangeChar3 = "G"
            Case "\" : ChangeChar3 = "l"
            Case "]" : ChangeChar3 = "R"
            Case "^" : ChangeChar3 = "m"
            Case "_" : ChangeChar3 = "s"
            Case "`" : ChangeChar3 = "0"
            Case "a" : ChangeChar3 = "-"
            Case "b" : ChangeChar3 = "\"
            Case "c" : ChangeChar3 = "Z"
            Case "d" : ChangeChar3 = ":"
            Case "e" : ChangeChar3 = "#"
            Case "f" : ChangeChar3 = "I"
            Case "g" : ChangeChar3 = "D"
            Case "h" : ChangeChar3 = ","
            Case "i" : ChangeChar3 = "N"
            Case "j" : ChangeChar3 = "O"
            Case "k" : ChangeChar3 = ";"
            Case "l" : ChangeChar3 = "%"
            Case "m" : ChangeChar3 = "Q"
            Case "n" : ChangeChar3 = "z"
            Case "o" : ChangeChar3 = "&"
            Case "p" : ChangeChar3 = "x"
            Case "q" : ChangeChar3 = "S"
            Case "r" : ChangeChar3 = ")"
            Case "s" : ChangeChar3 = "Y"
            Case "t" : ChangeChar3 = "4"
            Case "u" : ChangeChar3 = "M"
            Case "v" : ChangeChar3 = "/"
            Case "w" : ChangeChar3 = "8"
            Case "x" : ChangeChar3 = "g"
            Case "y" : ChangeChar3 = "X"
            Case "z" : ChangeChar3 = "V"
            Case "{" : ChangeChar3 = "d"
            Case "}" : ChangeChar3 = "q"
            Case "~" : ChangeChar3 = " "
            Case " " : ChangeChar3 = "$"
            Case Else
                ChangeChar3 = C_bit
        End Select
    End Function
    Private Function ChangeChar2(ByVal C_bit As String) As String
        Select Case C_bit
            Case "a" : ChangeChar2 = "!"
            Case "1" : ChangeChar2 = "#"
            Case "m" : ChangeChar2 = "$"
            Case "H" : ChangeChar2 = "%"
            Case "8" : ChangeChar2 = "&"
            Case "g" : ChangeChar2 = "("
            Case "o" : ChangeChar2 = ")"
            Case "0" : ChangeChar2 = "*"
            Case "O" : ChangeChar2 = "+"
            Case "k" : ChangeChar2 = ","
            Case "K" : ChangeChar2 = "-"
            Case "d" : ChangeChar2 = "."
            Case "x" : ChangeChar2 = "/"
            Case "R" : ChangeChar2 = "0"
            Case "{" : ChangeChar2 = "1"
            Case "," : ChangeChar2 = "2"
            Case "<" : ChangeChar2 = "3"
            Case "Y" : ChangeChar2 = "4"
            Case "q" : ChangeChar2 = "5"
            Case "A" : ChangeChar2 = "6"
            Case "(" : ChangeChar2 = "7"
            Case "/" : ChangeChar2 = "8"
            Case "&" : ChangeChar2 = "9"
            Case "I" : ChangeChar2 = ":"
            Case "J" : ChangeChar2 = ";"
            Case "e" : ChangeChar2 = "<"
            Case "u" : ChangeChar2 = "="
            Case "6" : ChangeChar2 = ">"
            Case "w" : ChangeChar2 = "?"
            Case "D" : ChangeChar2 = "@"
            Case "#" : ChangeChar2 = "A"
            Case "c" : ChangeChar2 = "B"
            Case "j" : ChangeChar2 = "C"
            Case "!" : ChangeChar2 = "D"
            Case ">" : ChangeChar2 = "E"
            Case "y" : ChangeChar2 = "F"
            Case "T" : ChangeChar2 = "G"
            Case "~" : ChangeChar2 = "H"
            Case "i" : ChangeChar2 = "I"
            Case "F" : ChangeChar2 = "J"
            Case ")" : ChangeChar2 = "K"
            Case "b" : ChangeChar2 = "L"
            Case "h" : ChangeChar2 = "M"
            Case "%" : ChangeChar2 = "N"
            Case "-" : ChangeChar2 = "O"
            Case "t" : ChangeChar2 = "P"
            Case "s" : ChangeChar2 = "Q"
            Case "X" : ChangeChar2 = "R"
            Case "f" : ChangeChar2 = "S"
            Case "*" : ChangeChar2 = "T"
            Case "l" : ChangeChar2 = "U"
            Case "_" : ChangeChar2 = "V"
            Case "^" : ChangeChar2 = "W"
            Case "n" : ChangeChar2 = "X"
            Case "+" : ChangeChar2 = "Y"
            Case "P" : ChangeChar2 = "Z"
            Case "U" : ChangeChar2 = "["
            Case "W" : ChangeChar2 = "\"
            Case "r" : ChangeChar2 = "]"
            Case "z" : ChangeChar2 = "^"
            Case "N" : ChangeChar2 = "_"
            Case "p" : ChangeChar2 = "`"
            Case "`" : ChangeChar2 = "a"
            Case "$" : ChangeChar2 = "b"
            Case "M" : ChangeChar2 = "c"
            Case "2" : ChangeChar2 = "d"
            Case "S" : ChangeChar2 = "e"
            Case "v" : ChangeChar2 = "f"
            Case "4" : ChangeChar2 = "g"
            Case "." : ChangeChar2 = "h"
            Case ";" : ChangeChar2 = "i"
            Case " " : ChangeChar2 = "j"
            Case "9" : ChangeChar2 = "k"
            Case "C" : ChangeChar2 = "l"
            Case "B" : ChangeChar2 = "m"
            Case "}" : ChangeChar2 = "n"
            Case "\" : ChangeChar2 = "o"
            Case "G" : ChangeChar2 = "p"
            Case "V" : ChangeChar2 = "q"
            Case "E" : ChangeChar2 = "r"
            Case "L" : ChangeChar2 = "s"
            Case "3" : ChangeChar2 = "t"
            Case "?" : ChangeChar2 = "u"
            Case ":" : ChangeChar2 = "v"
            Case "@" : ChangeChar2 = "w"
            Case "5" : ChangeChar2 = "x"
            Case "]" : ChangeChar2 = "y"
            Case "7" : ChangeChar2 = "z"
            Case "[" : ChangeChar2 = "{"
            Case "=" : ChangeChar2 = "}"
            Case "Q" : ChangeChar2 = "~"
            Case "Z" : ChangeChar2 = " "
            Case Else
                ChangeChar2 = C_bit
        End Select
    End Function
    Private Function ChangeChar1(ByVal C_bit As String) As String
        Select Case C_bit
            Case "!" : ChangeChar1 = "a"
            Case "#" : ChangeChar1 = "1"
            Case "$" : ChangeChar1 = "m"
            Case "%" : ChangeChar1 = "H"
            Case "&" : ChangeChar1 = "8"
            Case "(" : ChangeChar1 = "g"
            Case ")" : ChangeChar1 = "o"
            Case "*" : ChangeChar1 = "0"
            Case "+" : ChangeChar1 = "O"
            Case "," : ChangeChar1 = "k"
            Case "-" : ChangeChar1 = "K"
            Case "." : ChangeChar1 = "d"
            Case "/" : ChangeChar1 = "x"
            Case "0" : ChangeChar1 = "R"
            Case "1" : ChangeChar1 = "{"
            Case "2" : ChangeChar1 = ","
            Case "3" : ChangeChar1 = "<"
            Case "4" : ChangeChar1 = "Y"
            Case "5" : ChangeChar1 = "q"
            Case "6" : ChangeChar1 = "A"
            Case "7" : ChangeChar1 = "("
            Case "8" : ChangeChar1 = "/"
            Case "9" : ChangeChar1 = "&"
            Case ":" : ChangeChar1 = "I"
            Case ";" : ChangeChar1 = "J"
            Case "<" : ChangeChar1 = "e"
            Case "=" : ChangeChar1 = "u"
            Case ">" : ChangeChar1 = "6"
            Case "?" : ChangeChar1 = "w"
            Case "@" : ChangeChar1 = "D"
            Case "A" : ChangeChar1 = "#"
            Case "B" : ChangeChar1 = "c"
            Case "C" : ChangeChar1 = "j"
            Case "D" : ChangeChar1 = "!"
            Case "E" : ChangeChar1 = ">"
            Case "F" : ChangeChar1 = "y"
            Case "G" : ChangeChar1 = "T"
            Case "H" : ChangeChar1 = "~"
            Case "I" : ChangeChar1 = "i"
            Case "J" : ChangeChar1 = "F"
            Case "K" : ChangeChar1 = ")"
            Case "L" : ChangeChar1 = "b"
            Case "M" : ChangeChar1 = "h"
            Case "N" : ChangeChar1 = "%"
            Case "O" : ChangeChar1 = "-"
            Case "P" : ChangeChar1 = "t"
            Case "Q" : ChangeChar1 = "s"
            Case "R" : ChangeChar1 = "X"
            Case "S" : ChangeChar1 = "f"
            Case "T" : ChangeChar1 = "*"
            Case "U" : ChangeChar1 = "l"
            Case "V" : ChangeChar1 = "_"
            Case "W" : ChangeChar1 = "^"
            Case "X" : ChangeChar1 = "n"
            Case "Y" : ChangeChar1 = "+"
            Case "Z" : ChangeChar1 = "P"
            Case "[" : ChangeChar1 = "U"
            Case "\" : ChangeChar1 = "W"
            Case "]" : ChangeChar1 = "r"
            Case "^" : ChangeChar1 = "z"
            Case "_" : ChangeChar1 = "N"
            Case "`" : ChangeChar1 = "p"
            Case "a" : ChangeChar1 = "`"
            Case "b" : ChangeChar1 = "$"
            Case "c" : ChangeChar1 = "M"
            Case "d" : ChangeChar1 = "2"
            Case "e" : ChangeChar1 = "S"
            Case "f" : ChangeChar1 = "v"
            Case "g" : ChangeChar1 = "4"
            Case "h" : ChangeChar1 = "."
            Case "i" : ChangeChar1 = ";"
            Case "j" : ChangeChar1 = " "
            Case "k" : ChangeChar1 = "9"
            Case "l" : ChangeChar1 = "C"
            Case "m" : ChangeChar1 = "B"
            Case "n" : ChangeChar1 = "}"
            Case "o" : ChangeChar1 = "\"
            Case "p" : ChangeChar1 = "G"
            Case "q" : ChangeChar1 = "V"
            Case "r" : ChangeChar1 = "E"
            Case "s" : ChangeChar1 = "L"
            Case "t" : ChangeChar1 = "3"
            Case "u" : ChangeChar1 = "?"
            Case "v" : ChangeChar1 = ":"
            Case "w" : ChangeChar1 = "@"
            Case "x" : ChangeChar1 = "5"
            Case "y" : ChangeChar1 = "]"
            Case "z" : ChangeChar1 = "7"
            Case "{" : ChangeChar1 = "["
            Case "}" : ChangeChar1 = "="
            Case "~" : ChangeChar1 = "Q"
            Case " " : ChangeChar1 = "Z"
            Case Else
                ChangeChar1 = C_bit
        End Select
    End Function
End Module
