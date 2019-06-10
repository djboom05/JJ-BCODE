Const TRANSIGHT_BCODE_DLL = "Transight.CES.Adaptor.CESAdaptorBCODE"

Dim BCODE_ITEM_NUMBER
Dim ITEMNUMBER
Dim sBCODE
Dim m_Context
Dim clsItem

'*************************************
'******INITIALIZE    VARIABLES*******
'******NOTE: ASSIGN VALUES HERE*******
'*************************************
Sub InitVariables ()
	'INITIALIZE VARIABLE
	'ASSIGN PARAMETER WITH NUMBER COLUMN TO GET ITS CORRESPONDING ID 
	BCODE_ITEM_NUMBER = GetItemID(8000000) 'PUT BCODE Tender NUMBER in parameter
	ITEMNUMBER = 0
	sBCODE = ""
End Sub
'*************************************
'******INITIALIZE VARIABLES END*******
'*************************************

Sub Execute (Context)
  Dim bError
  Dim sMsg
  Dim bEnd
  'On Error Resume Next

  Call InitVariables()

  Set m_Context = Context
  Set clsItem = Nothing
  Context.v_ReturnValue = False
  Call Context.GetArgumentValue("Item", clsItem)
  If Context.x_Bill.v_BillStatus = 0 Then
    Context.ShowMsg("Idle operation not allowed.")
  Else
  	  
    With Context.x_UtilsForm
      Call .ClearScreen
      Call .DrawBox(0,1,32, 10 , "Please Scan bCode ",0, False)'697
      Call .AddTextBox(3, 3, "txtBCODE", "bCode Reference:  ", 32, "", 2)
      While Not bEnd
        If .ReadSave Then
          sBCODE = .x_TextBox("txtBCODE").Text 
          If Trim(sBCODE) = "" Then
            Call Context.ShowMsg("bCode cannot be blank.")
            '.x_TextBox("txtBCODE").SetFocus 
		  ElseIf Not( Len(sBCODE) = 32 Or Len(sBCODE) = 5 Or Len(sBCODE) = 10)  Then
		    Call Context.ShowMsg("Invalid bCode!")
          '  .x_TextBox("txtBCODE").SetFocus
		  ElseIf ValidateBCODE(sBCODE) = False Then
            .x_TextBox("txtBCODE").Text = ""
            '.x_TextBox("txtBCODE").SetFocus		  
          Else
			  'VALIDATE ITEMNUMBER
			  'If ITEMNUMBER > m_Context.x_Bill.x_Chk.v_SubTtl Then
				'Call Context.ShowMsg("bCode Tender Amount is greater than the SubTotal")
				'.x_TextBox("txtBCODE").Text = ""
			 ' Else
			'Call Context.GetArgumentValue("Item", clsItem)
			'If  ITEMNUMBER <> clsItem.v_Number Then 
			'Call m_Context.ShowMsg("Item Numbers do not match:" + vbCrlf + ITEMNUMBER + " And " + cStr(Trim(clsItem.v_Number)))
			'Context.v_ReturnValue = False
			'bEnd = True
			'else
				  Context.v_GlobalVar("CITIZENINFO") = sBCODE & vbTab
				  'Call Context.AddStepsAlphaNum(m_Context.FormatVal(ITEMNUMBER, "0.00"))
				  'Call m_Context.ShowMsg("Item Number: " + CStr(Trim(ITEMNUMBER)))
				  Call Context.AddStep(1, GetItemID(ITEMNUMBER))
				  Call Context.AddStep(91, 11)
				  Context.v_ReturnValue = True
				  bEnd = True
			  'end if
          End If  
        Else
          bEnd = True
        End If
      Wend
      .CleanUp
      
    End With
  End If
        
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

Function GetItemID(number)
	Call OpenPosConnection(m_PosConn)
	sSql = "SELECT id FROM menudef where number = '" & number & "'"
	Set rstRec = OpenRecordset(sSql,m_PosConn)
	Do While Not rstRec.EOF
		GetItemID = rstRec.Fields("ID")
		rstRec.MoveNext
	Loop
	Call CloseAdoObj(m_PosConn)
End Function

Function ValidateBCODE(bcodestring)


	  'Check If BCODE AlreadyScanned
	  Dim isDuplicate
	  isDuplicate = False
	  
	  With m_Context.x_Bill
		For fLoop = 0 To .x_List.Count - 1
			Set clsItem = m_Context.x_Bill.x_List.Item(fLoop)
			If clsItem.v_DtlType =  "M" Then
				Dim bcodestr1
				Dim bcodestr2
				Dim itemCount
				itemCount = Cstr(clsItem.x_Ref.Count)
				bcodestr1 = ""
				bcodestr2 = ""
				'Call m_Context.ShowMsg("BCODE Ref Count:" & Isbit(clsItem.x_PmMi.sTdef, 40))
				If Int(clsItem.x_Ref.Count) >=1 Then
					bcodestr1 = clsItem.x_Ref.Item(1)
				End If
				If Int(clsItem.x_Ref.Count) >=2 Then
					bcodestr2 = clsItem.x_Ref.Item(2)
					'Call m_Context.ShowMsg0("bcodestr2: " & bcodestr2)
				End If
				
				Dim Oldbcodestring
				Oldbcodestring = bcodestr1 & "" & bcodestr2
				'Call m_Context.ShowMsg("BCODE Ref Count:" & itemCount)
				'Call m_Context.ShowMsg("Previous BCODE:" & Oldbcodestring & " Length: " & Len(Oldbcodestring))
				'Call m_Context.ShowMsg("New BCODE:" & bcodestring & " Length: " & Len(bcodestring))

				If Mid(Trim(Oldbcodestring),1,12) = Mid(Trim(bcodestring),1,12) Then
				'If Int(BCODE_ITEM_NUMBER) = Int(clsItem.v_Number) Then
					isDuplicate = True
					ValidateBCODE = False
					Call m_Context.ShowMsg("Duplicate Entry! Already scanned bCode.")
					Exit For
				End If
			End If
		Next
	  End With


		If isDuplicate = False Then
		'Call m_Context.ShowMsg("No Duplicate found")
	  
  'On Error Resume Next
  Dim objBCODE
  Set objBCODE = CreateObject(TRANSIGHT_BCODE_DLL)

  Dim terminalNum
	terminalNum = m_Context.x_Bill.x_Pc.lNum
  Dim checkNum
	checkNum = m_Context.x_Bill.x_Chk.v_ChkNum
  ITEMNUMBER = 0

  
  Dim bcodeResponse
  bcodeResponse = objBCODE.BCODERequest(terminalNum, checkNum, bcodestring,"0")
  'Call m_Context.ShowMsg("Response","Response: " & bcodeResponse)
  
  If Instr(Ucase(bcodeResponse),"INVALID") > 0 Then
    ValidateBCODE = False
	Call m_Context.ShowMsg("Invalid bCode! Please check.")
  ElseIf Instr(Ucase(bcodeResponse),"CLAIMED") > 0 Then
    ValidateBCODE = False
	Call m_Context.ShowMsg("Already claimed bCode!")
  Else
	'Call m_Context.ShowMsg("BCODE Response: " & bcodeResponse & " " & Cstr(ubound(a)))
    Dim a
	a=Split(bcodeResponse,"|")
	If UBound(a) > 0  Then
		If UCase(a(0)) = "VALID" AND UCase(a(6)) <> "PESO" Then
			ValidateBCODE = True
			If a(5) = "DEMOITEM" Then
				ITEMNUMBER = 0
			Else
				ITEMNUMBER = a(5)
			End If
			sBCODE = a(2)
			Call m_Context.ShowMsg(ITEMNUMBER)
		Else
			ValidateBCODE = False
			Call m_Context.ShowMsg("bCode Invalid for Item")			
		End If
	Else
		Call m_Context.ShowMsg("Invalid response bCode!")
		ValidateBCODE = False
	End If

  End If

End If  
End Function
