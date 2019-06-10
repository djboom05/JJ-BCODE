Const TRANSIGHT_BCODE_DLL = "Transight.CES.Adaptor.CESAdaptorBCODE"

'Const TENDER_BCODE_NUMBER_LIST = "1001,1002,1003" 'LIST OF TENDER BCODE NUMBER, comma separated
Dim m_Context

Sub Execute (Context)
  Dim bError
  Dim sMsg
  Dim bEnd
  'On Error Resume Next
  
  
  'On Error Resume Next
  Dim objBCODE
  Set objBCODE = CreateObject(TRANSIGHT_BCODE_DLL)

  'Validate BCODE before Closing
  Context.v_ReturnValue = False
  With m_Context.x_Bill
  	For fLoop = 0 To .x_List.Count - 1
		Set clsItem = m_Context.x_Bill.x_List.Item(fLoop)
	    If clsItem.v_DtlType = "T" Then
			If m_Context.IsBitOn(clsItem.v_Tdef,61) Then
				Dim bcodestr1
				Dim bcodestr2
				Dim itemCount
				itemCount = Cstr(clsItem.x_Ref.Count)
				bcodestr1 = ""
				bcodestr2 = ""
				If clsItem.x_Ref.Count >=1 Then
					bcodestr1 = clsItem.x_Ref.Item(1)
				End If
				If clsItem.x_Ref.Count >=2 Then
					bcodestr2 = clsItem.x_Ref.Item(2)
				End If
				If bcodestring <> "" Then
					Dim bcodeResponse
					bcodeResponse = objBCODE.BCODERequest(.x_Pc.lNum, .x_Chk.v_ChkNum, bcodestring,"1")
					 ' Call m_Context.ShowMsg("BCODE Response: " & bcodeResponse & " " & Cstr(ubound(a)))
					  
					  If Instr(Ucase(bcodeResponse),"EXPIRED") > 0 Then
						Call m_Context.ShowMsg(bcodeResponse)
						Context.v_ReturnValue = False
						Exit Sub
					  ElseIf Instr(Ucase(bcodeResponse),"INVALID") > 0 Then
						Call m_Context.ShowMsg("Invalid bCode! Please check.")
						Context.v_ReturnValue = False
						Exit Sub
					  ElseIf Instr(Ucase(bcodeResponse),"CLAIMED") > 0 Then
						Call m_Context.ShowMsg("Already claimed bCode!")
						Context.v_ReturnValue = False
						Exit Sub
					  
					  Else
						'Call m_Context.ShowMsg("BCODE Response: " & bcodeResponse & " " & Cstr(ubound(a)))
						Dim a
						a=Split(bcodeResponse,"|")
						If UBound(a) > 0 Then
							' UCase(a(0)) = "VALID" OR UCase(a(0)) = "SUCCESS"
							If UCase(a(0)) = "SUCCESS" Then
								Context.v_ReturnValue = True
							Else
								Context.v_ReturnValue = False
								Call m_Context.ShowMsg(a(0))		
								Exit Sub
							End If
						Else
							Call m_Context.ShowMsg("Invalid response bCode!")
							Context.v_ReturnValue = False
							Exit Sub
						End If

					  End If
					  
				End If
			End If
		End If
	Next
  End With
        
End Sub


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

Function GetTenderID(number)
	Call OpenPosConnection(m_PosConn)
	sSql = "SELECT id FROM tender where number = '" & number & "'"
	Set rstRec = OpenRecordset(sSql,m_PosConn)
	Do While Not rstRec.EOF
		GetTenderID = rstRec.Fields("ID")
		rstRec.MoveNext
	Loop
	Call CloseAdoObj(m_PosConn)
End Function

Function isNumberExist(strList, NUM)
	On Error Resume Next
	Dim a
	Dim x
	a = Split(strList,",")
	isNumberExist = False
	
	For Each x In a
		If Int(x) = Int(NUM) Then
			isNumberExist = True
			Exit For
		End If
	Next
	
End Function
