Const OPEN_SEN_DSC_ID = 203
Const SCD_VAT_EXEMPT_ID = 0
Const PWD_DSC_ID = 248
Const SCD_TRANS_TYPE_SEQ = 26
Const DIPLOMAT_TRANS_TYPE_SEQ = 27
Const DIPLOMAT_SALES_TYPE_SEQ = 88
'Const TENDER_BCODE_NUMBER_LIST = "0" 'LIST OF TENDER BCODE NUMBER, comma separated

Sub Execute (Context)
  Dim clsItem
  Dim bError
  Dim sUserId
  Dim sUserName
  Dim sInfo
  Dim vTmp
  Dim i
  'On Error Resume Next

  Set m_Context = Context
  Set clsItem = Nothing
  Call Context.GetArgumentValue("Item", clsItem)
  If clsItem Is Nothing Then Exit Sub
  
  If clsItem.v_DtlType = DTLTYPE_MI Then
  'Context.ShowMsg(Context.v_GlobalVar("CITIZENINFO"))
  'Context.ShowMsg("Step aaaaaa1")
  'Context.ShowMsg(Isbit(clsItem.x_PmMi.sTdef, 2))
 ' Context.ShowMsg(Context.v_GlobalVar("CONDICOUNT"))
  'Context.ShowMsg(GetItemName3(clsItem.v_Number))
  'Context.ShowMsg(Context.v_GlobalVar("CITIZENINFO"))
    If Not Isbit(clsItem.x_PmMi.sTdef, 2) And _
       Context.v_GlobalVar("CONDICOUNT") <> "" Then
      Context.ShowMsg("Condiment required")
      Context.v_ReturnValue = False
   ElseIf GetItemName3(clsItem.v_Number) = "1" And _ 
      Context.v_GlobalVar("CITIZENINFO") <> "" Then
	'Context.ShowMsg("Step 1")
      sInfo = Mid(Context.v_GlobalVar("CITIZENINFO"), 1, _
              instr(1, Context.v_GlobalVar("CITIZENINFO"), vbTab) - 1)
			  
			  ' Context.ShowMsg("Step 2")
      vTmp = Split(sInfo, "|")
	  
      sUserId = vTmp(0)
	   'Context.ShowMsg("Step 3")
	  Dim iremarks1
	  Dim iremarks2
	  
	  If Len(sUserId) <= 20 Then
	   'Context.ShowMsg("Step 4")
	    iremarks1 = sUserId
		iremarks2 = ""
	  ElseIf Len(sUserId) = 32 Then
	   'Context.ShowMsg("Step 5")
	   Context.ShowMsg(sUserId)
		iremarks1 = Mid(sUserId,1,16)
		iremarks2 = Mid(sUserId,17,16)
		'Context.ShowMsg(iremarks1 + "/" + iremarks2)
	  End If
	  
      clsItem.x_Ref.Add iremarks1
	   'Context.ShowMsg("Step 6")
      clsItem.x_Ref.Add iremarks2
	   'Context.ShowMsg("Step 7")
      clsItem.x_Ref.Add ""
	  ' Context.ShowMsg("Step 8")
     ' If (Context.v_GlobalVar("PAXDISC") - 1) = 0 Then 
     '   Context.v_GlobalVar("CITIZENINFO") = ""
        'Context.v_GlobalVar("PAXDISC") = ""
      'Else
        Context.v_GlobalVar("CITIZENINFO") = Mid(Context.v_GlobalVar("CITIZENINFO"), _
                                             Instr(1, Context.v_GlobalVar("CITIZENINFO"), vbTab) + 1)
											  'Context.ShowMsg("Step 9")
        'Context.v_GlobalVar("PAXDISC") = Context.v_GlobalVar("PAXDISC") - 1
      'End If	   
    End If
    'End If 
    'End If
    'End If
  ElseIf clsItem.v_DtlType = DTLTYPE_TND Then
    'Context.ShowMsg(Context.v_GlobalVar("CITIZENINFO"))
	'Context.ShowMsg(m_Context.IsBitOn(clsItem.v_Tdef,61))
    If Context.v_GlobalVar("SCD") = "" And m_Context.x_Bill.x_Chk.v_TxnTypeSeq = SCD_TRANS_TYPE_SEQ Then
      Call Context.ShowMsg("Please apply Senior Citizen Discount first before payment.")
      Context.v_ReturnValue = False    
    ElseIf m_Context.x_Bill.x_Chk.v_TxnTypeSeq = DIPLOMAT_TRANS_TYPE_SEQ Then
      With m_Context.x_Bill
        For i = 0 To .x_List.Count - 1
          Set clsItem = m_Context.x_Bill.x_List.Item(i)

          If .x_Chk.v_TxnTypeSeq <> DIPLOMAT_TRANS_TYPE_SEQ Or clsItem.v_SalesType <> DIPLOMAT_SALES_TYPE_SEQ Then
            Context.ShowMsg("Operation not allowed." & vbCrLf & "Please perform correct Diplomat transaction.")
            Context.v_ReturnValue = False    
            Exit Sub
          End If
        Next
      End With
	  'isNumberExist(TENDER_BCODE_NUMBER_LIST, clsItem.v_Number) And _
	  'm_Context.IsBitOn(clsItem.v_Tdef,61) And _ 
	  'Isbit(clsItem.x_PmMi.sTdef,61) And _
	  'Context.v_GlobalVar("PAXDISC") <> ""
	ElseIf m_Context.IsBitOn(clsItem.v_Tdef,61) And _ 
       Context.v_GlobalVar("CITIZENINFO") <> "" Then

      sInfo = Mid(Context.v_GlobalVar("CITIZENINFO"), 1, _
              instr(1, Context.v_GlobalVar("CITIZENINFO"), vbTab) - 1)
      vTmp = Split(sInfo, "|")
      sUserId = vTmp(0)
	  Dim remarks1
	  Dim remarks2
	  
	  If Len(sUserId) <= 20 Then
	    remarks1 = sUserId
		remarks2 = ""
	  ElseIf Len(sUserId) = 32 Then
		remarks1 = Mid(sUserId,1,16)
		remarks2 = Mid(sUserId,17,16)
	  End If
	  
      clsItem.x_Ref.Add remarks1
      clsItem.x_Ref.Add remarks2
      clsItem.x_Ref.Add ""
     ' If (Context.v_GlobalVar("PAXDISC") - 1) = 0 Then 
     '   Context.v_GlobalVar("CITIZENINFO") = ""
        'Context.v_GlobalVar("PAXDISC") = ""
      'Else
        Context.v_GlobalVar("CITIZENINFO") = Mid(Context.v_GlobalVar("CITIZENINFO"), _
                                             Instr(1, Context.v_GlobalVar("CITIZENINFO"), vbTab) + 1)
        'Context.v_GlobalVar("PAXDISC") = Context.v_GlobalVar("PAXDISC") - 1
      'End If	   
    End If
  End If
End Sub

Function GetItemName3(number)
	Call OpenPosConnection(m_PosConn)
	sSql = "SELECT name3 FROM menudef where number = " & number & ""
	Set rstRec = OpenRecordset(sSql,m_PosConn)
	Do While Not rstRec.EOF
		GetItemName3 = rstRec.Fields("name3")
		rstRec.MoveNext
	Loop
	Call CloseAdoObj(m_PosConn)
End Function

Function CatchError(source)
  If Err.Number <> 0 Then
    Call m_Context.ShowMsg("Error Occur in " & source, Err.Description, True, 0, 3)
    CatchError = True
    m_PosConn.Rollback
  End If
End Function

Public Sub OpenPosConnection (cnn)
  Set cnn = CreateObject("ADODB.COnnection")
  cnn.ConnectionTimeout = CONNTIMEOUT
  cnn.Open POS_SERVER & ";User Id=datascan;Password=DTSbsd7188228" 
End Sub

Public Sub OpenClientConnection (cnn)
  Set cnn = CreateObject("ADODB.COnnection")
  cnn.ConnectionTimeout = CONNTIMEOUT
  cnn.Open POS_CLIENT & ";User Id=datascan;Password=DTSbsd7188228"
End Sub

Public Sub CloseAdoObj(ObjAdo)
  On Error Resume Next
  If Not ObjAdo Is Nothing Then
    ObjAdo.Close
    Set objAdo = Nothing
  End If
End Sub

Public Function OpenRecordset(Source, ActiveConnection)
  Dim rst

  Set rst = CreateObject("ADODB.Recordset")
  rst.CursorLocation = 3
  rst.Open Source, ActiveConnection, 1, 3
  Set OpenRecordset = rst
End Function

Function isNumberExist(strList, NUM)
	On Error Resume Next
	Dim a
	Dim x
	a = Split(strList,",")
	isNumberExist = False
	
	For Each x In a
		m_Context.ShowMsg(x & " " & NUM)
		If Int(x) = Int(NUM) Then
			isNumberExist = True
			Exit For
		End If
	Next
	
End Function
