<%
   ' ���o�U��쪺���
     Name = Request.Form("Name")
     Sex = Request.Form("Sex")
     Job = Request.Form("Job")
     Email = Request.Form("Email")
     HomePage = Request.Form("HomePage")
     messagebody = Request.Form("messagebody")
     icq = Request.Form("icq")
     security = Request.Form("security")

    ' �p�G HTML �y�k���� �h�ϥΦ��ԭz
      messagebody = Replace(messagebody,"<" , "&lt;")

      messagebody = Replace(messagebody,vbCrLf , "<br>")

    ' ������_���,<br>�@�˷|�Q�[�@,��J�O�_�j�� 80 �r��,���F���[�@,�ҥH��`�r,���N��<br>
      messagebody = Replace(messagebody, "<br>" ,"`")

    ' textarea �۰��_�檺�P�_������s�X��,�ҥH�u�n�ε{���X Parse_text �ӫ����_��B�z,
      parse_text(messagebody)

      messagebody = Replace(messagebody ,"`", "<br>" )
      messagebody = Replace(messagebody,vbCrLf,"<br>")

    ' �p�G�S�� e-mail
      If Email="" Then
         Email=" "
      End If

    ' �p�G�S�� icq
      If icq="" Then
         icq=" "
      End If

      Homepage = trim(HomePage)
    ' �p�G�S�� HTTp
      If len(trim(HomePage))  = 0 or mid(HomePage,1,4) <> "http" Then
         HomePage ="http://" & Homepage
      End If

    ' user data write to bowser...Cookies
      Domain = Request.ServerVariables("SERVER_NAME")

      Response.Cookies("Name") = Name
      Response.Cookies("Name").Domain = Domain
      Response.Cookies("Name").Expires = date()+30

      Response.Cookies("sex") = sex
      Response.Cookies("sex").Domain = Domain
      Response.Cookies("sex").Expires = date()+30

      Response.Cookies("Email") = Email
      Response.Cookies("Email").Domain = Domain
      Response.Cookies("Email").Expires = date()+30

      Response.Cookies("HomePage") = HomePage
      Response.Cookies("HomePage").Domain = Domain
      Response.Cookies("HomePage").Expires = date()+30

      Response.Cookies("icq") = icq
      Response.Cookies("icq").Domain = Domain
      Response.Cookies("icq").Expires = date()+30

' ��~�K�O WEBHOSTME ���}�һy�k�p�U:
'Set Conn = Server.CreateObject ("ADODB.Connection")
'Conn.Provider = "Microsoft.Jet.OLEDB.4.0"
'Conn.ConnectionString = "Data Source=" & Server.MapPath ("/cgi-bin/BOARD/GBOOK.mdb")
'Conn.Open 


    ' �P��Ʈw�s�u
      Set conn = Server.CreateObject("ADODB.Connection")
      DBPath = Server.MapPath("GBook.mdb")
      conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
     Set RS = Server.CreateObject("ADODB.Recordset")
     RS.Open "select * from guestbook ORDER BY MESSAGEID DESC", conn,1,3

    ' �]���Q������ SORT ���ӱƧǾ�ӱi�K�̻P�^���̪�����,���S���\�h���D,
    ' ���^����i�K�̪��ɭԧڷQ�����N�N�^���̪�����,�ƧǦb�L���U��,�o�˴N��
    ' ���� �Q�תO �@�˶��z�L�}�ҤG����Ʈw�ϥΰj��Ӱ���ƪ����o�P�Ƨ�,���سt��
    ' �|�C�ܦh,�ҥH���ܴN�b��Ƽg�J�ɴN�Ƨ�...�B�s�i�K�̪���Ʀb�̤W��
    ' �� COUNT �_�l�ȧڬO�]���q 999999 �}�Һ�_,�ҥH��A����������h�H�b�d����,
    ' �̦n��N 999999 ��j�@�I,�h�j���S���Y...
    ' �O�o SAVEREPLY.ASP ���� 999999 �]�n�@�_��@!!
    ' �`�N: �� count = 0 ��,�Ƨǰ��D�N�X�{���~,�]���ڷQ���X��L�覡�ӳ]�w count �ȤF...
    ' �p�G���H�Q�X�ӤF,�·Чi��...thanks.

   IF RS.RECORDCOUNT <= 0 THEN    ' �]�����N�s�i�K�̪���ƱƦb�̤W��,�ҥH�ڨS�ϥΦ۰ʽs�����\��
      count = 999999
   ELSE
      rs.movelast
      count = rs("messageid") - 1   ' ���messageid �̫᪺�ȨӴ� 1
   END IF

   ' �s�W�@����ƿ�,�üg�J��Ʈw
   rs.AddNew
   rs("messageid") = count
   rs("replyid") = count
   rs("Sort") = count
   rs("Name") = Name
   rs("Sex") = Sex
   rs("Job") = Job
   rs("Email") = Email
   rs("HomePage") = HomePage
   rs("messagebody") = messagebody
   rs("icq") = icq
   rs("security") = security
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
         '�N���ץ[2
         tmp_Len = tmp_Len  + 2
      ElseIf Mid(m_output,i,1) = "`" then
         temp_Memo =  Mid(m_output,old_pos+1,i-old_pos)
         '�٭�p�ƭ�,�]�� <BR> �]�|�Q��J,�ҥH�N<BR>�N����"`",�H�קK�Q�[�@
         old_pos = i
         tmp_Len = 0
      ELSE
         '�_�h�[1
         tmp_Len = tmp_Len + 1
      End If

      '�p�G�j�󵥩� 80 �r���A�ǳƿ�X���G
      If tmp_Len >=  80  Then
         temp_Memo =  Mid(m_output,old_pos+1,i-old_pos) & "<br>"
        ' �٭�p�ƭ�
         old_pos = i
         tmp_Len = 0
      End If
         temp2_Memo = temp2_Memo & temp_Memo
         temp_Memo = ""
  Next

  '�N��l���r��(�p�G�٦�����)�@�ֿ�X
  If tmp_Len > 0 Then
     messagebody = temp2_Memo & Mid(m_output,old_pos+1,tot_Len-old_pos)
     messagebody = Replace(messagebody ,"`", "<br>" )
  End If
End Function
%>