<%
' �̶��h���J��ӥ��Ϊ��ť�,�����ƪ����n(�u����2���H�W�Ψ�)
IF Request.Form("replylevel") = "2" Then
   SPACES = "�@�@"
END IF

' ���o�U��쪺���
Name = Request.Form("Name")
Sex = Request.Form("Sex")
'Job = Request.Form("Job")
Email = Request.Form("Email")
HomePage = Request.Form("HomePage")
messagebody = SPACES & Request.Form("messagebody")
'icq = Request.Form("icq")
'security = Request.Form("security")

' �p�G HTML �y�k���� �h�ϥΦ��ԭz
messagebody = Replace(messagebody,"<" , "&lt;")

' ������_���,<br>�@�˷|�Q�[�@,��J�O�_�j�� 80 �r��,���F���[�@,�ҥH��`�r,���N��<br>
messagebody = Replace(messagebody,vbCrLf , "`" & SPACES)

' textarea �۰��_�檺�P�_������s�X��,�ҥH�u�n�ε{���X Parse_text �ӫ����_��B�z,
parse_text(messagebody)


'�p�G�S�� e-mail
If Email="" Then
    Email=" "
End If

'�p�G�S�� icq
If icq="" Then
    icq=" "
End If

Homepage = trim(HomePage)
'�p�G�S�� HTTp
If len(trim(HomePage))  = 0 or mid(HomePage,1,4) <> "http" Then
    HomePage ="http://" & Homepage
End If

' ��~�K�O WEBHOSTME ���}�һy�k�p�U:
Set Conn = Server.CreateObject ("ADODB.Connection")
Conn.Provider = "Microsoft.Jet.OLEDB.4.0"
Conn.ConnectionString = "Data Source=" & Server.MapPath ("/cgi-bin/BOARD/GBOOK.mdb")
Conn.Open 

' �}�Ҹ�Ʈw
' Set conn = Server.CreateObject("ADODB.Connection")
' DBPath = Server.MapPath("GBook.mdb")
' conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath

Set rs = Server.CreateObject("ADODB.Recordset")
SortSql = "Select *From guestbook where Messageid=" & Request.form("Messageid")
rs.Open SortSql, conn, 1,3

' �x�s�ק�L�᪺���
rs("Name") = Name
'rs("Sex") = Sex
'rs("Job") = Job
rs("Email") = Email
rs("HomePage") = HomePage
rs("messagebody") = messagebody
'rs("icq") = icq
'rs("security") = security
rs.update
rs.close
SET CONN = NOTHING

Response.Redirect "book.asp" ' ��^�y�ͨ��O�z

Function Parse_text(m_output)
  '�W�@����m�B�ثe�֭p���סB��ڦ�m�B�`����
  Dim old_pos, tmp_Len, i, tot_Len,temp_Memo,temp2_Memo
  old_pos = 0   '�W�@����m
  tmp_Len = 0   '�O���޿�W������
  TOT_LEN = Len(Trim(m_output))
  temp_Memo = ""
  temp2_Memo = "" 

 ' i �N��r�ꤤ��ڪ���m
  For i = 1 To tot_Len

      If Len(Hex(Asc(Mid(m_output, i, 1)))) > 2 and Mid(m_output,i,1) <> "`" Then
         tmp_Len = tmp_Len  + 2         '�N���ץ[2
      ElseIf Mid(m_output,i,1) = "`" then
         temp_Memo =  Mid(m_output,old_pos+1,i-old_pos)
         '�٭�p�ƭ�,�]�� <BR> �]�|�Q��J,�ҥH�N<BR>�N����"`",�H�קK�Q�[�@
         old_pos = i
         tmp_Len = 0
      ELSE
         tmp_Len = tmp_Len + 1         '�_�h�[ 1
      End If

      '�p�G�j�󵥩� 80 �r���A�ǳƿ�X���G
      If tmp_Len >=  80  Then
         temp_Memo = Mid(m_output,old_pos+1,i-old_pos) & "<br>" & SPACES
        ' �٭�p�ƭ�
         old_pos = i
         tmp_Len = 0
      End If
         temp2_Memo = temp2_Memo & temp_Memo
         temp_Memo = ""
  Next

  '�N��l���r��(�p�G�٦�����)�@�ֿ�X
   messagebody = Replace(TEMP2_MEMO ,"`", "<br>") & Mid(m_output,old_pos+1,tot_Len-old_pos+2)

End Function

%>