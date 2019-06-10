'Transight POS 3.30c
'07/09/10  dtlim  Added SCD VAT Exempt
'07/10/10  dtlim  Added Credit Card Slip printing
'07/19/10  dtlim  Added DtlType to determine if clsItem is discount
'07/23/10  dtlim  Enabled Double Height and Width in printing Amount Due
'08/02/10  dtlim  Added TakeOut and Delivery Check Limit
'10/19/17  aries   No limit in all transaction for PWD, fix computation of Vat Exempt Sales
'10/20/17 aries   BIR TRAILER Configuration


Const DBH_H1W1 = 1
Const DBH_H1W2 = 2
Const DBH_H2W1 = 3
Const DBH_H2W2 = 4
Const INVAL_2  = 0
Const INVAL_1  = -1
Const VATEXEMPT = 1     '0- Not Vat Exempt; 1- Vat Exempt
Const PRINTERID = 18     'Terminal Printer ID
Const PRINTCOPY = 1   'Discount Print 2 Copy
Const PRINTERWIDTH = 32  'Thermal = 40, Dot Matrix = 32
Const SENIORDISC = 1003 'Senior Citizen Discount Number
Const VATSENIOR = 1007 'VAT on Senior Discount Number
Const PWDDISC = 1006 'Senior Citizen Discount Number
Const VATPWD = 1020 'VAT on PWD Discount Number
Const VATSCPROMO = 1021 'VAT exempt on Senior Discount Number Promo
Const MAX_CHECK_LIMIT_TO = 300
Const DINE_IN_ID = 25 'TRANS TYPE ID

'----BIRINFO TRAILER CONFIG---------
'----Put 0 if not required
Const TRAILER_BIRINFO_NUMBER_HEADER = 51 'TRAILER HEADER
Const TRAILER_BIRINFO_NUMBER_FOOTER = 52 'TRAILER FOOTER
'----BIRINFO TRAILER CONFIG End

Dim aSplit
Dim bSplit
Dim cSplit
Dim dSplit
Dim rSplit
Dim bPadChinese
Dim m_Context
Dim sChgDue
Dim sServiceChg
Dim bEatIn
Dim bTakeAway
Dim bOthers
Dim sRemarks
Dim mConn
Dim iPrintCnt 


Sub Execute(Context)
  Dim sRecptLogo
  Dim sRecptHeader
  Dim sRecptBody
  Dim sRecptDiscount
  Dim sRecptTrailer
  Dim sRecptOrigCopy
  Dim sRecptDupCopy
  Dim sRecptFooter
  Dim sPrtString
  Dim sTemp
  Dim sTemp2
  Dim clsTrailer1
  Dim bPrt32Col 
  Dim clsDesc
  Dim bIncludeTnd
  Dim sPayment
  Dim curLastTxnChgDue
  Dim i
  Dim k
  Dim bError
  Dim bChgTrail 
  Dim bPrintMultiCpy
  Dim bDiscount
  Dim sDBLWidth
  Dim sTrType
  Dim sUserId
  Dim sUserName 
  Dim sUserTemp
  Dim iDelimeter
  Dim vTmp

  On Error Resume Next
  
  Set m_Context = Context

  iPrintCnt = 1
  bEatIn = 0
  bTakeAway = 0
  bOthers = 0
  bPadChinese = (Context.x_Pos.v_RegionalString = "CH")
  bChgTrail = False
  
  With Context.x_Bill
    If m_Context.x_Bill.x_List.Count = 0 And Not .x_Chk.StatBit(14) Then
      Exit Sub
    End If

    Set clsDesc = Context.x_Pos.x_Rvc.GetDescriptor(Context.x_Pos.v_Language)
    sRecptLogo = GetHeaderLogo(PRINTERWIDTH, Context.x_Pos.x_Rvc, clsDesc)
    sRecptHeader = GetPrtStrHdr(PRINTERWIDTH, Context.x_Pos.x_Rvc, clsDesc)

    curLastTxnChgDue = "0.00"
    sRecptBody = .x_List.GetPrtStrBody(True, 0, False, 0, _
                 .x_Sys, .x_Rvc, Context.x_Pos.x_SysMsg, _
                 .x_Chk.v_FastTxn, True, PRINTERWIDTH, Context.x_Pos.v_Language, cStr(sPayment), _
                 cCur(curLastTxnChgDue)) & vbCrLf
    If Not bIncludeTnd Then sPayment = ""
    sTemp = GetPrtStrSumm(True, PRINTERWIDTH, clsDesc, 0, .x_Txn, .x_Txns.x_CsldtedHistory, True, False, True, True, sPayment, curLastTxnChgDue,VATEXEMPT)
    sRecptBody = sRecptBody & sTemp
    If .x_Rvc.DoNotPurgeTndOnPrt Then
      sRecptBody = sRecptBody & .x_List.ItemisedTndMed(0, .x_Sys, .x_Rvc, _
                   Context.x_Pos.x_SysMsg, .x_Chk.v_FastTxn, PRINTERWIDTH)
    End If
    sPrtString = sPrtString & _
                 GetPrtAddChkIndicator(PRINTERWIDTH, clsDesc, .x_Chk.x_Txn)
'    sRecptBody = sRecptBody & String(PRINTERWIDTH, "-") & vbCrLf & vbCrLf

    For lLoop = 0 To .x_List.Count - 1
      Set clsItem = .x_List.Item(lLoop)

      If clsItem.v_DtlType = "D" Then
        bDiscount = True
      End If

      If clsItem.v_DtlType = "D" And clsItem.v_Number = SENIORDISC Then 'Senior Citizen
        sUserID = clsItem.x_Ref.Item(1)
        sUserName = clsItem.x_Ref.Item(2)
        sTmp = m_Context.PadC("SENIOR DISC", PRINTERWIDTH, " ", 2)
        sDBLWidth = m_Context.FormatPrintStr(DBH_H1W2, bPadChinese, sTmp) & vbCrLf 
        If sRecptDiscount = "" Then
          sRecptDiscount = String(PRINTERWIDTH, "-") & vbCrLf & vbCrLf & sDBLWidth
        End If
        sRecptDiscount = sRecptDiscount & Context.PadR(CStr("Printed Name: " & sUserName), PRINTERWIDTH) & vbNewLine
        sRecptDiscount = sRecptDiscount & Context.PadR(CStr("ID Number   : " & sUserID), PRINTERWIDTH) & vbNewLine & vbNewLine
        sRecptDiscount = sRecptDiscount & Context.PadR(CStr("Signature   : ___________________"), PRINTERWIDTH) & vbNewLine & vbNewLine
        iPrintCnt = 2
        bChgTrail =True
        'Exit For
	  ElseIf clsItem.v_DtlType = "D" And clsItem.v_Number = VATSCPROMO Then 'SC PROMO
        sUserID = clsItem.x_Ref.Item(1)
        sUserName = clsItem.x_Ref.Item(2)
        sTmp = m_Context.PadC("SC PROMO", PRINTERWIDTH, " ", 2)
        sDBLWidth = m_Context.FormatPrintStr(DBH_H1W2, bPadChinese, sTmp) & vbCrLf 
        If sRecptDiscount = "" Then
          sRecptDiscount = String(PRINTERWIDTH, "-") & vbCrLf & vbCrLf & sDBLWidth
        Else
          sRecptDiscount = sRecptDiscount & String(PRINTERWIDTH, "-") & vbCrLf & vbCrLf & sDBLWidth
        End If
        sRecptDiscount = sRecptDiscount & Context.PadR(CStr("Printed Name : " & sUserName), 32) & vbNewLine
        sRecptDiscount = sRecptDiscount & Context.PadR(CStr("ID Number   : " & sUserID), 32) & vbNewLine & vbNewLine
        sRecptDiscount = sRecptDiscount & Context.PadR(CStr("Signature   : __________________________"), 32) & vbNewLine & vbNewLine
        iPrintCnt = 2
        bChgTrail =True
      

     ElseIf clsItem.v_DtlType = "D" And clsItem.v_Number = PWDDISC Then 'PWD Disc
        sUserTmp = Mid(Context.v_GlobalVar("CITIZENINFO"), 1, InStr(1, Context.v_GlobalVar("CITIZENINFO"), vbTab) - 1)
        vTmp = Split(sUserTmp, "|")
        sUserID = clsItem.x_Ref.Item(1)
        sUserName = clsItem.x_Ref.Item(2)
        'sUserTIN = clsItem.x_Ref.Item(3)
        sTmp = m_Context.PadC("PWD DISC", PRINTERWIDTH, " ", 2)
        sDBLWidth = m_Context.FormatPrintStr(DBH_H1W2, bPadChinese, sTmp) & vbCrLf 
        If sRecptDiscount = "" Then
          sRecptDiscount = String(PRINTERWIDTH, "-") & vbCrLf & vbCrLf & sDBLWidth
        End If
        sRecptDiscount = sRecptDiscount & Context.PadR(CStr("Printed Name: " & sUserName), PRINTERWIDTH) & vbNewLine
        sRecptDiscount = sRecptDiscount & Context.PadR(CStr("ID Number   : " & sUserID), PRINTERWIDTH) & vbNewLine
        'sRecptDiscount = sRecptDiscount & Context.PadR(CStr("TIN         : " & sUserTIN), PRINTERWIDTH) & vbNewLine & vbNewLine
        sRecptDiscount = sRecptDiscount & Context.PadR(CStr("Signature   : ___________________"), PRINTERWIDTH) & vbNewLine & vbNewLine
        iPrintCnt = 2
        bChgTrail =True
        'Exit For
      ElseIf clsItem.v_DtlType = "D" Then 'Other Discount
        iPrintCnt = 2
        bChgTrail =True
        'Exit For
      End If
    Next

    If .x_Chk.v_TxnStatBit(3) Then 'Reprint
      iPrintCnt = 1
    ElseIf m_Context.x_Bill.x_Chk.StatBit(28) Then 'Refund 
      iPrintCnt = 2
    End If

    sRecptDupCopy = m_Context.FormatPrintStr(DBH_H1W2, bPadChinese, Context.PadC("DUPLICATE COPY", PRINTERWIDTH)) & vbNewLine & vbNewLine

    Set clsTrailer1 = Context.x_PosMem.GetTrailer(Context.x_Pos.v_Language, _
                    .x_LastTender, .x_Rvc, False)   










	Context.v_Globalvar("surveycode") = ""

	If .x_Chk.v_TxnStatBit(3) Then 'Reprint
		'Call Context.ShowMsg("REPREF", "REPRINT", True, 0, 1)
		Call repref(cnn, Context, rst)
		'Call Context.ShowMsg("REPREF", "REPRINTED", True, 0, 1)
	ElseIf .x_Chk.StatBit(28) Then 'Refund 
		'Call Context.ShowMsg("REPREF", "REFUND", True, 0, 1)
		Call repref(cnn, Context, rst)
		'Call Context.ShowMsg("REPREF", "REFUNDED", True, 0, 1)
	Else
		'Call Context.ShowMsg("TXN", "ORDINARY TXN", True, 0, 1)
		Call checkersub(cnn, Context, txnnum , rst, rst0, rst1, rst2, rst3)
		'Call Context.ShowMsg("TXN", "TXN END", True, 0, 1)
	End If

    

    sRecptFooter = GetPrtStrTrailer(PRINTERWIDTH, .x_Chk.Closed, _
                   .x_Chk.x_Txn, clsTrailer1, (m_Context.x_Pos.v_RegionalString = "CH")) & Context.v_GLobalvar("surveycode") & vbCrLf & GetPOS_ProviderInfo
  End With
  
  For mLoop = 1 To iPrintCnt
    If PRINTERWIDTH = 32 Then
      Call Context.PrintString(sRecptHeader & sRecptBody & sRecptDiscount & sRecptFooter & Chr(31), "", PRINTERID, , sRecptLogo) 
      'Call Context.ShowMsg(sRecptHeader & sRecptBody & sRecptDiscount & sRecptFooter) 
    Else
      Call Context.PrintString("", sRecptHeader & sRecptBody & sRecptDiscount & sRecptFooter & Chr(31),PRINTERID, , sRecptLogo) 
      'Call Context.ShowMsg(sRecptHeader & sRecptBody & sRecptDiscount & sRecptFooter) 
    End If
  Next

  Set m_Context = Nothing
End Sub

Function GetHeaderLogo(byWidth, clsRvc, clsDesc)
  Dim clsHeader
  Dim byLang

  With m_Context.x_Bill
    byLang = m_Context.x_Pos.v_Language
  End With
  Set clsHeader = clsRvc.GetHeader(byLang)
  GetHeaderLogo = clsHeader.sLogo
End Function

Function GetPrtStrHdr(byWidth, clsRvc, clsDesc)
  Dim colItem
  Dim iCnt
  Dim iInnerCnt
  Dim sTmp
  Dim sTmp1
  Dim lTmp
  Dim byLang
  Dim sErr
  Dim sPageNumHold
  Dim sPlaceHolder
  Dim sSubString
  Dim clsHeader
  Dim byBlank
  Dim sFinishWaste
  Dim byTemp
  Dim bCtrlChar
  Dim iLenCheck
  Dim lTxnCount
  Dim sTemp
  Dim bPadChinese
  Dim sRemark
  Dim bDummy
  Dim sDBLWidth
  
  On Error Resume Next
  

  With m_Context.x_Bill
    byLang = m_Context.x_Pos.v_Language
    bPadChinese = (m_Context.x_Pos.v_RegionalString = "CH")
    If .x_Chk.v_TxnStatBit(6) Then
'      Call m_Context.showmsg("Here 1") 
      sSubString = sSubString & m_Context.PadC(m_Context.x_Pos.x_SysMsg.TimedChk(), byWidth) & vbCrLf
      If .x_Chk.x_Cust Is Nothing Then
        sSubString = sSubString & _
                     m_Context.PadC(Replace(m_Context.x_Pos.x_SysMsg.AdvTime(), "#", _
                     .x_Sys.DateTimeByRegion(.x_Chk.v_TblOpen, _
                     1)), byWidth) & vbCrLf
      Else
        sSubString = sSubString & m_Context.PadC(.x_Chk.QuoteStr(.x_Sys, _
                     m_Context.x_Pos.x_SysMsg.DlvrTime(), .x_Chk.v_TblOpen, _
                     byLang), byWidth) & vbCrLf
      End If
    Else
'      Call m_Context.showmsg("Here 2") 
      Set clsHeader = clsRvc.GetHeader(byLang)
      If .x_Chk.v_Training Then
'       Call m_Context.showmsg("Here 3")
        byBlank = 0
        For iCnt = 1 To 10
          If Trim(clsHeader.colTrainingHeader.Item(iCnt)) = "" Then
            byBlank = byBlank + 1
          Else
            While byBlank > 0
              sSubString = sSubString & vbCrLf
              byBlank = byBlank - 1
            Wend
            If clsHeader.colTrainingDW(iCnt) Then
              byTemp = byWidth / 2
            Else
              byTemp = byWidth
            End If
            
            'byTemp = IIf(clsHeader.colTrainingDW(iCnt), byWidth / 2, byWidth)
            sTemp = m_Context.PadC(Trim(clsHeader.colTrainingHeader.Item(iCnt)), _
                    byTemp) & vbCrLf
            If clsHeader.colTrainingDW(iCnt) Then
              sTemp = GetPrtChar(sTemp, DBH_H1W2, byWidth, bPadChinese)
            Else
              sTemp = GetPrtChar(sTemp, DBH_H1W1, byWidth, bPadChinese)
            End If                                 
'            sTemp = GetPrtChar(sTemp, IIf(clsHeader.colTrainingDW(iCnt), _
'                               DBH_H1W2, DBH_H1W1), byWidth, bPadChinese)
'              sTemp = HeaderEnhance(sTemp, clsHeader.colTrainingInvColor(iCnt), bCtrlChar)
            sSubString = sSubString & sTemp
          End If
        Next
      ElseIf .x_Chk.v_FastTxn or .x_Chk.Closed Then
'       Call m_Context.showmsg("Here 4")
        byBlank = 0
        For iCnt = 1 To 10
          If Trim(clsHeader.colRecHeader.Item(iCnt)) = "" Then
            byBlank = byBlank + 1
          Else
            While byBlank > 0
              sSubString = sSubString & vbCrLf
              byBlank = byBlank - 1
            Wend
            
            If clsHeader.colRecDW(iCnt) Then
              byTemp = byWidth / 2
            Else
              byTemp = byWidth
            End If
            'byTemp = IIf(clsHeader.colRecDW(iCnt), byWidth / 2, byWidth)
            sTemp = m_Context.PadC(Trim(clsHeader.colRecHeader.Item(iCnt)), _
                                   byTemp) & vbCrLf
            If clsHeader.colRecDW(iCnt) Then
              sTemp = GetPrtChar(sTemp, DBH_H1W2, byWidth, bPadChinese)
            Else
              sTemp = GetPrtChar(sTemp, DBH_H1W1, byWidth, bPadChinese)
            End If                         
'            sTemp = GetPrtChar(sTemp, IIf(clsHeader.colRecDW(iCnt), DBH_H1W2, DBH_H1W1), _
'                              byWidth)
'            sTemp = HeaderEnhance(sTemp, clsHeader.colRecInvColor(iCnt), bCtrlChar)
            sSubString = sSubString & sTemp
          End If
        Next
      Else
'       Call m_Context.showmsg("Here 5")
        byBlank = 0
        For iCnt = 1 To 10
          If Trim(clsHeader.colGstHeader.Item(iCnt)) = "" Then
            byBlank = byBlank + 1
          Else
            While byBlank > 0
              sSubString = sSubString & vbCrLf
              byBlank = byBlank - 1
            Wend
            If clsHeader.colGstDW(iCnt) Then
              byTemp = byWidth / 2
            Else
              byTemp = byWidth
            End If 
            sTemp = m_Context.PadC(Trim(clsHeader.colGstHeader.Item(iCnt)), _
                                  byTemp) & vbCrLf
            If clsHeader.colGstDW(iCnt) Then
              sTemp = GetPrtChar(sTemp, DBH_H1W2, byWidth, bPadChinese)
            Else
              sTemp = GetPrtChar(sTemp, DBH_H1W1, byWidth, bPadChinese)
            End If  
'            sTemp = HeaderEnhance(sTemp, clsHeader.colGstInvColor(iCnt), bCtrlChar) '& vbCrLf
            sSubString = sSubString & sTemp
          End If
        Next
      End If
	  
	  
'---------BIR INFO TRAILER-----------------------------------------------------------------------------
		sSubString = sSubString & GetBIRINFO_TRAILER(TRAILER_BIRINFO_NUMBER_HEADER)	
'---------BIR INFO TRAILER-----------------------------------------------------------------------------

 '     Call m_Context.ShowMsg("In invert color")
 '     sSubString = HeaderEnhance(sSubString, clsHeader.colGstInvColor(iCnt), bCtrlChar) & vbCrLf     

      'sSubString = sSubString & IIf(bCtrlChar, "" & vbCrLf & "", "")
      Set clsHeader = Nothing
    End If
'    sSubString = HeaderEnhance(sSubString, clsHeader.colGstInvColor(iCnt), bCtrlChar) & vbCrLf     
'    Call m_Context.showmsg("After Header Enhance") 
    'transaction history

'*********************Changes on Receipt Header************************'
'    If m_Context.IsBitOn(clsRvc.sOptions, 81) And (.x_Txns.v_HistoryRcdCount > 0) Then
'      sSubString = sSubString & _
'                   GetPrtChar(m_Context.PadR(m_Context.x_Pos.x_SysMsg.OrderChangedBy(), _
'                              byWidth) & vbCrLf, DBH_H1W1, byWidth, bPadChinese)
'      lTxnCount = .x_Txns.v_HistoryRcdCount
'      For iCnt = 1 To lTxnCount
'        If .x_Chk.x_Txns.Item(iCnt).v_Type = "A" Then
'        Else
'          Set colItem = m_Context.WrapText(.x_Sys.DateTimeByRegion( _
'                        .x_Txns.Item(iCnt).v_EndTime) & _
'                        " - " & .x_Txns.Item(iCnt).v_EmpNameTxn, _
'                        byWidth)
'          For iInnerCnt = 1 To colItem.Count
'            sSubString = sSubString & GetPrtChar(colItem.Item(iInnerCnt) & _
'                         vbCrLf, DBH_H1W1, byWidth, bPadChinese)
'          Next
'        End If
'      Next
'      sSubString = sSubString & vbCrLf
'      Set colItem = Nothing
'    End If


    'complaint
'    If m_Context.IsBitOn(clsRvc.sOptions, 82) And (.x_Chk.x_Complaint.Count > 0) Then
'      For iCnt = 1 To .x_Chk.x_Complaint.Count
'        With .x_Chk.x_Complaint.Item(iCnt)
'          sTmp = Replace("[%]", "%", .v_ComplaintID)
'          sSubString = sSubString & GetPrtChar(m_Context.x_Pos.x_SysMsg.Complaint(), _
'                       DBH_H1W1, byWidth, bPadChinese) & vbCrLf
'          sSubString = sSubString & GetPrtChar(.x_Sys.DateTimeByRegion(_
'                       .v_ComplaintDate), DBH_H1W1, byWidth, bPadChinese) & vbCrLf
'          sSubString = sSubString & GetPrtChar(sTmp & " " & _
'                       .v_ComplaintName, DBH_H1W1, byWidth, bPadChinese) & vbCrLf
'          sSubString = sSubString & GetPrtChar(m_Context.x_Pos.x_SysMsg.Remark(), _
'                       DBH_H1W1, byWidth, bPadChinese)
'          Set colItem = m_Context.WrapText(.v_Complaint, byWidth)
'          For iInnerCnt = 1 To colItem.Count
'            sSubString = sSubString & GetPrtChar(colItem.Item(iInnerCnt), _
'                         DBH_H1W1, byWidth, bPadChinese) & vbCrLf
'          Next
'          sTmp = Replace("[%]", "%", .v_SolutionID)
'          sSubString = sSubString & GetPrtChar(m_Context.x_Pos.x_SysMsg.Solution(), _
'                       DBH_H1W1, byWidth, bPadChinese) & vbCrLf
'          sSubString = sSubString & GetPrtChar(.x_Sys.DateTimeByRegion( _
'                       .v_SolutionDate), DBH_H1W1, byWidth, bPadChinese) & vbCrLf
'          sSubString = sSubString & GetPrtChar(sTmp & " " & .v_SolutionName, _
'                       DBH_H1W1, byWidth, bPadChinese) & vbCrLf
'          sSubString = sSubString & GetPrtChar(m_Context.x_Pos.x_SysMsg.Remark(), _
'                       DBH_H1W1, byWidth, bPadChinese)
'          Set colItem = m_Context.WrapText(.v_Solution, byWidth)
'        End With
'        For iInnerCnt = 1 To colItem.Count
'          sSubString = sSubString & GetPrtChar(colItem.Item(iInnerCnt), _
'                       DBH_H1W1, byWidth, bPadChinese) & vbCrLf
'        Next
'        sSubString = sSubString & vbCrLf
'      Next
'    End If
    sRemark = ""
    If .x_Chk.v_DelAddress <> "" Then
      Set colItem = m_Context.WrapText(.x_Chk.v_DelAddress, byWidth)
      For iCnt = 1 To colItem.Count
        sRemark = sRemark & colItem.Item(iCnt) & vbCrLf
      Next
      Set colItem = Nothing
    End If
    If sRemark <> "" Then
      sSubString = sSubString & _
                   GetPrtChar(sRemark, DBH_H1W1, byWidth, bPadChinese) & vbCrLf
      sRemark = ""
    End If
    sSubString = sSubString & GetPrtChar(GetThaiTaxHeaderStr(byWidth, bPadChinese), _
                                         DBH_H1W1, byWidth, bPadChinese)

    sTmp = .x_Owner.lNum & " " & .x_Owner.Name(byLang)
    sTmp1 = .x_Pc.lNum & " " & .x_Pc.Name(byLang)
    If Len(sTmp) + Len(sTmp1) + 2 > byWidth Then
      sSubString = sSubString & GetPrtChar(sTmp1 & vbCrLf & sTmp & vbCrLf, _
                                           DBH_H1W1, byWidth, bPadChinese)
    Else
      sSubString = sSubString & GetPrtChar(sTmp1 & m_Context.PadL(sTmp, byWidth - _
                   m_Context.StrLenGraphically(sTmp1)) & _
                   vbCrLf, DBH_H1W1, byWidth, bPadChinese)
    End If

    If m_Context.IsBitOn(clsRvc.sOptions, 86) Then
      sTmp = .x_TxnEmp.lNum & " " & .x_TxnEmp.Name(byLang)
      sTmp1 = m_Context.x_Pos.x_SysMsg.PrintedCnt() & .x_Chk.v_ChkPrtCnt
      sSubString = sSubString & GetPrtChar(sTmp1 & String(byWidth - _
                   m_Context.StrLenGraphically(sTmp & sTmp1), " ") & _
                   sTmp & vbCrLf, DBH_H1W1, byWidth, bPadChinese)
    End If
    sSubString = sSubString & String(byWidth, "-") & vbCrLf
    
    If m_Context.IsBitOn(clsRvc.sOptions, 40) Then
      sPlaceHolder = String(byWidth, "%")
      sTmp = sPlaceHolder
    Else
      sTmp = Trim(clsDesc.sChk & CStr(.x_Chk.v_ChkNum))
      iLenCheck = m_Context.StrLenGraphically(sTmp)
      If m_Context.IsBitOn(clsRvc.sOptions, 154) Then
        sTmp = GetPrtChar(sTmp, DBH_H1W2, byWidth, bPadChinese)
        iLenCheck = iLenCheck * 2
        sTmp = Chr(27) & "|2C" & sTmp & ""
      ElseIf m_Context.IsBitOn(clsRvc.sOptions, 153) Then
        sTmp = GetPrtChar(sTmp, DBH_H2W2, byWidth, bPadChinese)
        iLenCheck = iLenCheck * 2
      Else      
        sTmp = GetPrtChar(sTmp, DBH_H1W1, byWidth, bPadChinese)
      End If
      sPlaceHolder = String(byWidth - iLenCheck, "%")
      sTmp = sTmp & sPlaceHolder
    End If
    
    If clsDesc.sCover = "" Then
      sSubString = sSubString & Replace(sTmp, sPlaceHolder, " ") & vbCrLf
    Else
      sTmp1 = Trim(clsDesc.sCover & CStr(.x_Chk.v_Cover))
      If Len(sPlaceHolder) >= Len(sTmp1) Then
        sSubString = sSubString & Replace(sTmp, sPlaceHolder, _
                     GetPrtChar(m_Context.PadL(sTmp1, Len(sPlaceHolder)), _
                     DBH_H1W1, byWidth, bPadChinese)) & vbCrLf
      Else
        sSubString = sSubString & Replace(sTmp, sPlaceHolder, " ") & vbCrLf
        sSubString = sSubString & GetPrtChar(m_Context.PadR(sTmp1, byWidth) & _
                     vbCrLf, DBH_H1W1, byWidth, bPadChinese)
      End If
    End If
    sTmp = ""
    sTmp1 = ""
    GetPrtStrHdr = sSubString
    sSubString = ""

    If m_Context.IsBitOn(clsRvc.sOptions, 76) Then
      If byWidth = 24 Then
        sTmp = m_Context.PadC(clsRvc.Name(byLang), byWidth)
        GetPrtStrHdr = sTmp & vbCrLf & GetPrtStrHdr
      Else
        sTmp = m_Context.PadC(clsRvc.Name(byLang), byWidth, " ", 2)
        GetPrtStrHdr = m_Context.FormatPrintStr(DBH_H2W2, bPadChinese, sTmp) & _
                       vbCrLf & GetPrtStrHdr
      End If
    ElseIf m_Context.IsBitOn(clsRvc.sOptions, 75) Then
      If byWidth = 24 Then
        sTmp = m_Context.PadC(clsRvc.Name(byLang), byWidth)
        GetPrtStrHdr = sTmp & vbCrLf & GetPrtStrHdr
      Else
        sTmp = m_Context.PadC(clsRvc.Name(byLang), byWidth, " ", 2)
        GetPrtStrHdr = m_Context.FormatPrintStr(DBH_H1W2, bPadChinese, sTmp) & _
                       vbCrLf & GetPrtStrHdr
      End If 
    End If

    If .x_Chk.v_ID <> "" Then
      sTmp = " [" & .x_Chk.v_ID & "]"
      If byWidth = 24 Then
        GetPrtStrHdr = GetPrtStrHdr & sTmp & vbCrLf
      Else
        GetPrtStrHdr = GetPrtStrHdr & _
                       m_Context.FormatPrintStr(DBH_H1W1, bPadChinese, sTmp) & vbCrLf
      End If
    End If

    If .x_Table Is Nothing Then
      sTmp = m_Context.PadC(.x_Sys.DateTimeByRegion(.x_Chk.v_OpenDtTime), byWidth)
      If byWidth = 24 Then
        GetPrtStrHdr = GetPrtStrHdr & sTmp & vbCrLf
      Else
        GetPrtStrHdr = GetPrtStrHdr & m_Context.FormatPrintStr(DBH_H1W1, _
                       bPadChinese, sTmp) & vbCrLf
      End If
    ElseIf m_Context.IsBitOn(clsRvc.sOptions, 51) Or m_Context.IsBitOn(clsRvc.sOptions, 111) Then
      sTmp = m_Context.PadC(.x_Sys.DateTimeByRegion(.x_Chk.v_OpenDtTime), byWidth)
      If byWidth = 24 Then
        GetPrtStrHdr = GetPrtStrHdr & sTmp & vbCrLf
        sTmp = m_Context.PadR(clsDesc.sTblName & GetGrpStr(.x_Table.sName, _
                              .x_Chk.v_TableGrp), byWidth)
        GetPrtStrHdr = GetPrtStrHdr & sTmp & vbCrLf
      Else
        GetPrtStrHdr = GetPrtStrHdr & m_Context.FormatPrintStr(DBH_H1W1, _
                       bPadChinese, sTmp) & vbCrLf
        sTmp = m_Context.PadR(clsDesc.sTblName & GetGrpStr(.x_Table.sName, _
                             .x_Chk.v_TableGrp), byWidth, " ", 2)
        If m_Context.IsBitOn(clsRvc.sOptions, 111) Then
           GetPrtStrHdr = GetPrtStrHdr & m_Context.FormatPrintStr(DBH_H1W2, _
                          bPadChinese, sTmp) & vbCrLf
        Else
           GetPrtStrHdr = GetPrtStrHdr & m_Context.FormatPrintStr(DBH_H2W2, _
                          bPadChinese, sTmp) & vbCrLf
        End If
      End If
    Else
      If byWidth = 24 Then
        sTmp = m_Context.PadC(.x_Sys.DateTimeByRegion(.x_Chk.v_OpenDtTime), byWidth)
        GetPrtStrHdr = GetPrtStrHdr & sTmp & vbCrLf
        sTmp = m_Context.PadR(clsDesc.sTblName & GetGrpStr(.x_Table.sName, _
                             .x_Chk.v_TableGrp), byWidth)
        GetPrtStrHdr = GetPrtStrHdr & sTmp & vbCrLf
      Else
        sTmp = m_Context.PadR(.x_Sys.DateTimeByRegion(.x_Chk.v_OpenDtTime), _
               2 * byWidth / 3) & GetPrtBlk(byWidth - (2 * byWidth / 3), _
               clsDesc.sTblName, GetGrpStr(.x_Table.sName, .x_Chk.v_TableGrp), _
               bPadChinese)
        GetPrtStrHdr = GetPrtStrHdr & m_Context.FormatPrintStr(DBH_H1W1, _
                       bPadChinese, sTmp) & vbCrLf
      End If
    End If

    sSubString = sSubString & String(byWidth, "-") & vbCrLf
    If .x_Chk.v_Training Then
      sSubString = sSubString & _
                   m_Context.PadC(m_Context.x_Pos.x_SysMsg.TrainChk(), byWidth) & vbCrLf
    End If

    If byWidth = 24 Then
      GetPrtStrHdr = GetPrtStrHdr & sSubString
    Else
      GetPrtStrHdr = GetPrtStrHdr & m_Context.FormatPrintStr(DBH_H1W1, _
                     bPadChinese, sSubString)
    End If

    sSubString = ""
    If .x_Chk.StatBit(14) Then  'V1.31
      sTmp = m_Context.PadC(m_Context.x_Pos.x_SysMsg.VoidedChk(), byWidth)
      If byWidth = 24 Then
        GetPrtStrHdr = sTmp & vbCrLf & GetPrtStrHdr
      Else
        GetPrtStrHdr = m_Context.FormatPrintStr(DBH_H1W1, bPadChinese, sTmp) & _
                       vbCrLf & GetPrtStrHdr
      End If
    End If

    If .x_Chk.StatBit(28) Then
      sTmp = m_Context.PadC(m_Context.x_Pos.x_SysMsg.OverringChk(), byWidth)
      If byWidth = 24 Then
        GetPrtStrHdr = sTmp & vbCrLf & GetPrtStrHdr
      Else
        GetPrtStrHdr = m_Context.FormatPrintStr(DBH_H1W1, bPadChinese, sTmp) & _
                       vbCrLf & GetPrtStrHdr
      End If
    End If

    If .x_Chk.v_TxnStatBit(3) Then
      sTmp = m_Context.PadC(m_Context.x_Pos.x_SysMsg.ReprtChk(), byWidth)
      
      If byWidth = 24 Then
        GetPrtStrHdr = sTmp & vbCrLf & GetPrtStrHdr
      Else
        GetPrtStrHdr = m_Context.FormatPrintStr(DBH_H1W1, bPadChinese, sTmp) & _
                       vbCrLf & GetPrtStrHdr
      End If
    End If

    If .v_FinishWasteEntry Then
      sFinishWaste = m_Context.PadC(m_Context.x_Pos.x_SysMsg.FinishWaste(), _
                                    byWidth) & vbCrLf
    End If

    If .x_Sys.bDemo Then
      sTmp = m_Context.PadC(m_Context.x_Pos.x_SysMsg.DemoMode, byWidth)
      If byWidth = 24 Then
        GetPrtStrHdr = sTmp & vbCrLf & sFinishWaste & GetPrtStrHdr
      Else
        GetPrtStrHdr = m_Context.FormatPrintStr(DBH_H1W1, bPadChinese, sTmp) & _
                       vbCrLf & m_Context.FormatPrintStr(DBH_H1W1, bPadChinese, _
                       sFinishWaste) & GetPrtStrHdr
      End If
    ElseIf sFinishWaste <> "" Then
      If byWidth = 24 Then
        GetPrtStrHdr = sFinishWaste & GetPrtStrHdr
      Else
        GetPrtStrHdr = m_Context.FormatPrintStr(DBH_H1W1, bPadChinese, _
                       sFinishWaste) & GetPrtStrHdr
      End If
    End If
  End With
End Function

Function HeaderEnhance(sHeader,bInverse,bCtrlChar)
    Const N = ""
    Const ERED = ""
    Const RED = "|rC"
    Dim lLoop 
    Dim sTemp 
    
    HeaderEnhance = sHeader
    
    If bInverse Then
      HeaderEnhance = ""
      For lLoop = 1 To Len(sHeader)
        sTemp = Mid(sHeader, lLoop, 1)
        HeaderEnhance = HeaderEnhance & sTemp
        If (sTemp = ERED Or sTemp = N) And lLoop <> Len(sHeader) Then
          HeaderEnhance = HeaderEnhance & RED
        End If
      Next
 
    HeaderEnhance = InvertColor(HeaderEnhance)  
         
    End If
'562    If bCtrlChar Then HeaderEnhance = N & HeaderEnhance
    If Not bCtrlChar Then bCtrlChar = bInverse
End Function

Function DblWidth(sStr) 
    Const N = ""
    Const DHW = "|2C"
    DblWidth = Chr(27) & DHW & sStr & N
End Function

Function InvertColor(sStr) 
    Const N = ""
    Const REVERSE = "|rC"
    InvertColor = Chr(27) & REVERSE & sStr & N
End Function

Function GetPrtChar(sString, eDblByte, byWidth, bPrtChinese)
  If byWidth = 24 Then
    GetPrtChar = sString
  Else
    GetPrtChar = m_Context.FormatPrintStr(eDblByte, bPrtChinese, sString)
  End If
End Function

Function GetThaiTaxHeaderStr(byWidth, bPadChinese)
  Dim colWrap
  Dim sThaiTaxHdr
  Dim iWrapLine
  On Error Resume Next
  
  sThaiTaxHdr = m_Context.x_Bill.x_Chk.v_ThaiTaxHeader
  If sThaiTaxHdr <> "" Then
    Set colWrap = m_Context.WrapText(sThaiTaxHdr, byWidth)
    For iWrapLine = 1 To colWrap.Count
      If Trim(colWrap.Item(iWrapLine)) <> "" Then
         GetThaiTaxHeaderStr = GetThaiTaxHeaderStr & _
                               m_Context.PadC(colWrap.Item(iWrapLine), _
                               byWidth) & vbCrLf
      End If
    Next
  End If
End Function

Function GetThaiTaxDetailStr(byWidth, bPadChinese)
  Dim colWrap
  Dim sThaiTaxDet
  Dim iWrapLine
  On Error Resume Next
  
  sThaiTaxDet = m_Context.x_Bill.x_Chk.v_ThaiTaxDetail
  If sThaiTaxDet <> "" Then
    Set colWrap = m_Context.WrapText(sThaiTaxDet, byWidth)
    For iWrapLine = 1 To colWrap.Count
      If Trim(colWrap.Item(iWrapLine)) <> "" Then
        GetThaiTaxDetailStr = GetThaiTaxDetailStr & _
                              colWrap.Item(iWrapLine) & vbCrLf
      End If
    Next
  End If
      
End Function


Function GetGrpStr(sTbl, byGrp)
  GetGrpStr = sTbl & " / " & byGrp
End Function

Function GetPrtBlk(byWidth, sStr1, sStr2, bPadChinese)
  Dim lCnt
  
  On Error Resume Next
  
  If m_Context.StrLenGraphically(sStr2) > byWidth Then
    GetPrtBlk = m_Context.PadR(sStr2, byWidth)
  ElseIf (m_Context.StrLenGraphically(sStr1) + _
         m_Context.StrLenGraphically(sStr2)) > byWidth Then
    GetPrtBlk = m_Context.PadR(sStr1, _
                byWidth - m_Context.StrLenGraphically(sStr2)) & sStr2
  Else
    GetPrtBlk = m_Context.PadL(sStr1 & sStr2, byWidth)
  End If

End Function
                      
Function GetPrtStrSumm(bIncludeTnd, byWidth, clsDesc, _
                       lLastTxnSeq, clsTxn, clsTotal, _
                       bPrtSumm, bTaxByRound, bPrtVat, _
                       bTndOnly, sPaymentItems, curLastTxnChgDue,VATEXEMPT)
  Dim sTmp 
  Dim sTmp1
  Dim sTndTdef
  Dim curTmp
  Dim by
  Dim curPayment
  Dim lTmp
  Dim byLang
  Dim sErr
  Dim bNoSubTtl
  Dim curOutStd
  Dim curAmtDue
  Dim lActiveTxnRndSeq
  Dim bPrevTxnRnd
  Dim lTxnIndex
  Dim bPrtedVat
  Dim sVat
  Dim bPadChinese
  Dim sAddOn
  Dim iLen
  Dim sComboSaving
  Dim colWrap
  Dim iWrapLine
  Dim curTaxable
  Dim curChgDue
  Dim byCombine 
  Dim cur12VAT
  Dim curSubttl
  Dim sTender
  Dim bReturn
  Dim curReturnVAT
  Dim curDeposit
  Dim lItemCnt
  Dim curDisc
  Dim bSCD
  Dim curVATSenior
  Dim curVATSeniorExempt
  Dim curSeniorDisc
'PWD
  Dim bPWD
  Dim curPWDDisc
'PWD END
  Dim iSCDCover
  Dim iPWDCover
  Dim bVATZero

  Dim curVATPWD
  Dim curVATPWDExempt
  Dim curVatZeroRatedSales
  'SC Promo
  Dim curSCPromo
  Dim curSCPromoExempt
  Dim iSCPromoCover
  Dim bSCPromo


  On Error Resume Next

  byCombine = 2
  curReturnVAT = 0
  curDeposit = 0
  lItemCnt = 0
  curDisc = 0
  curVATSenior = 0
  curVATSeniorExempt = 0
  bSCD = False
  curVATPWD = 0
  curVATPWDExempt = 0
  bPWD = False
  bVATZero = False
  bSCPromo = False
  curSCPromoExempt = 0
  curSCPromo = 0
  curVatZeroRatedSales = 0
  
  With m_Context.x_Bill
    If m_Context.IsBitOn(.x_Rvc.sOptions, 122) Then 'Print canadian tax itemizer
      If m_Context.IsBitOn(.x_Chk.Status, 2) = m_Context.IsBitOn(.x_Chk.Status, 4) Then
        byCombine = 4
      End If
    ElseIf m_Context.IsBitOn(.x_Rvc.sOptions, 121) Then 'Print canadian tax detail
      If m_Context.IsBitOn(.x_Chk.Status, 2) = m_Context.IsBitOn(.x_Chk.Status, 3) Then
        byCombine = 3
      End If
    End If

    If lLastTxnSeq <> INVAL_2 Then bPrevTxnRnd = True
    lTxnIndex = .x_Txns.GetIndex(lLastTxnSeq)
    byLang = m_Context.x_Pos.v_Language()
    bPadChinese = Not (m_Context.x_Pos.v_RegionalString = "CH")

    If Not .x_LastTender Is Nothing Then
      sTndTdef = .x_LastTender.sTDef
    End If

    If x_Chk.Closed And Not bPrevTxnRnd Then
      If bTndOnly Then
        If .x_Txns.v_HistoryRcdCount > 0 And .x_Chk.v_SlipPrted And .x_Txns.v_PrintedWithSumm Then
          If clsTotal.v_PrevTxnSeq <> -1 Then
            curTmp = .x_List.GetTxnSubTtl(clsTotal.v_TxnSeq)
          Else
            curTmp = .x_List.GetTxnSubTtl(.x_List.GetActiveTxnSeq(clsTotal.v_TxnSeq))
          End If
          If .x_Chk.v_SubTtl = curTmp Then
            bNoSubTtl = True
          End If
        End If
      End If
    End If

    If (m_Context.IsBitOn(sTndTdef, 11) Or m_Context.IsBitOn(.x_Rvc.sOptions, 126)) And Not bPrevTxnRnd Then
      GetPrtStrSumm = GetPrtStrSumm & _
                      .x_List.GetPrtSalesItemizer(byWidth, _
                      .x_Sys, clsDesc)
    End If

    If clsTxn.v_PrevTxnSeq <> INVAL_1 Then
      lActiveTxnRndSeq = clsTxn.v_PrevTxnSeq
    ElseIf clsTxn.v_TxnSeq = INVAL_2 Then
      lActiveTxnRndSeq = INVAL_2
    Else
      lActiveTxnRndSeq = .x_List.GetActiveTxnSeq(lLastTxnSeq)
    End If
 
    For fLoop = 0 To .x_List.Count - 1
      Set clsItem = m_Context.x_Bill.x_List.Item(fLoop)

      If clsItem.v_SalesType = 85 Then
        bVATZero = True
      End If

      'INVAL_2  
      If lActiveTxnRndSeq = 0 Or (clsItem.v_TxnSeq <> 0 And clsItem.v_TxnSeq <= lActiveTxnRndSeq) Then
        If clsItem.v_DtlType = "D" Then
          bDisc = True
        ElseIf clsItem.v_DtlType <> "M" Then
        ElseIf clsItem.x_Mi.v_Ret = True Then
          lItemCnt = lItemCnt - clsItem.v_Qty 
        ElseIf m_Context.IsBitOn(clsItem.v_Tdef,2) Then
          lItemCnt = lItemCnt + clsItem.v_Qty
        ElseIf m_Context.IsBitOn(clsItem.v_Tdef,23) And (clsItem.v_Price = 0) Then
        ElseIf clsItem.StatBit(5) = True Then 
          lItemCnt = lItemCnt + clsItem.v_Qty
        ElseIf clsItem.StatBit(4) = True Then
          lItemCnt = lItemCnt + clsItem.v_Qty
        Else
          lItemCnt = lItemCnt + clsItem.v_Qty
        End If
      End If  
    Next

    If m_Context.IsBitOn(.x_Rvc.sOptions, 116) Then
'   With mclsDisp
'     lItemCnt = 0
'     For lCnt = .L_Bound To .U_Bound
'       If lTxnSeq = 0 Or (.Item(lCnt).v_TxnSeq <> INVAL_2 And .Item(lCnt).v_TxnSeq <= lTxnSeq) Then
'         If .Item(lCnt).v_DtlType <> DTLTYPE_MI Then
'         ElseIf IsBit(.Item(lCnt).v_Tdef, MENUDEF_COND_ITEM) Then
'         ElseIf IsBit(.Item(lCnt).v_Tdef, MENUDEF_SPC_FISH_PRICE_ITEM) And (.Item(lCnt).v_Price = 0) Then
'         ElseIf .Item(lCnt).StatBit(SBD_04_COMBO_MUM) Then
             '** 15/10/2004  cclaw   2.56ac  Total item sold for return item
'         ElseIf .Item(lCnt).x_Mi.v_Ret Then
'           lItemCnt = lItemCnt - .Item(lCnt).v_Qty
'         Else
'           lItemCnt = lItemCnt + .Item(lCnt).v_Qty
'         End If
'       End If
'     Next
'   End With

      GetPrtStrSumm = GetPrtStrSumm & Space(3) & _
                      FormatStr(m_Context.x_Pos.x_SysMsg.TtlItemSold(), _
                      byWidth - 19, lItemCnt, _
                      bPadChinese) & vbCrLf
    End If

    curSubttl = .x_Chk.v_SubTtl

  '--VAT ZERO RATED--------------------------------------------------------------
    'If bVATZero = True then
      For kLoop = 0 To .x_List.Count - 1
        Set clsItem = m_Context.x_Bill.x_List.Item(kLoop) 
        If (clsItem.v_SalesType = 101 Or clsItem.v_SalesType = 95) and clsItem.v_DtlType = "M" Then
          curVatZeroRatedSales = curVatZeroRatedSales + clsItem.v_Price
        End If
      Next 
   ' End If
'------------------------------------------------------------------------------

   '--Total--------------------------------------------------------------
    If Not .x_Chk.StatBit(14) Then
	  Dim totalamt
	  totalamt = 0
      For kLoop = 0 To .x_List.Count - 1
        Set clsItem = m_Context.x_Bill.x_List.Item(kLoop) 
        If clsItem.v_DtlType = "M" Then
          totalamt = totalamt + clsItem.v_Price
        End If
      Next
	
      GetPrtStrSumm = GetPrtStrSumm & GetPrtSummDtl(byWidth, 6, _
                      "Total:", totalamt, sTmp)
    End If
   '--End Total--------------------------------------------------------------
  
    '--Compute Discount--------------------------------------------------------------
    If bDisc Then
      For kLoop = 0 To .x_List.Count - 1
        Set clsItem = m_Context.x_Bill.x_List.Item(kLoop) 
        If clsItem.v_DtlType = "D" And clsItem.v_Number <> VATSENIOR And clsItem.v_Number <> VATPWD Then
          curDisc = curDisc + clsItem.v_Price
        End If
      Next

      '--Subtotal--------------------------------------------------------------
      If bNoSubTtl Then
      Else
        If bPrevTxnRnd Then
          curTmp = .x_List.GetTxnSubTtl(lActiveTxnRndSeq)
        Else
          curTmp = .x_Chk.v_SubTtl
        End If

		'10/19/17
		If .x_Chk.StatBit(14) Or .x_Chk.StatBit(28) Then
		    curSubttl = curTmp + (curDisc*-1)
		Else
			curSubttl = curTmp + Abs(curDisc)
		End If
        
        
      End If
      '--End Subtotal--------------------------------------------------------------

      '--Less Discount--------------------------------------------------------------
      iSCDCover = 0
      iPWDCover = 0
	  iSCPromoCover = 0
      For kLoop = 0 To .x_List.Count - 1
        Set clsItem = m_Context.x_Bill.x_List.Item(kLoop) 
        If clsItem.v_Number = SENIORDISC And clsItem.v_DtlType = "D" Then
          bSCD = True
          iSCDCover = iSCDCover + 1
          curSeniorDisc = curSeniorDisc + clsItem.v_Price
        ElseIf clsItem.v_Number = PWDDISC And clsItem.v_DtlType = "D" Then
          bPWD = True
          iPWDCover = iPWDCover + 1
          curPWDDisc = curPWDDisc + clsItem.v_Price
        ElseIf clsItem.v_Number = VATSENIOR And clsItem.v_DtlType = "D" Then
          curVATSenior = curVATSenior + clsItem.v_Price
        ElseIf clsItem.v_Number = VATPWD And clsItem.v_DtlType = "D" Then
          curVATPWD = curVATPWD + clsItem.v_Price
		ElseIf clsItem.v_Number = VATSCPROMO And clsItem.v_DtlType = "D" Then
          bSCPromo = True
		  iSCPromoCover = iSCPromoCover + 1
		  curSCPromo = curSCPromo + clsItem.v_Price
        End If
        If (clsItem.v_DtlType = "D") Then
'          GetPrtStrSumm = GetPrtStrSumm & GetPrtSummDtl(byWidth, 6, clsItem.v_Name, _
'                                                        clsItem.v_Price, "")
        End If
      Next
	  
	  If bSCPromo Then
          GetPrtStrSumm = GetPrtStrSumm & GetPrtSummDtl(byWidth, 6, "Less SC Vat:", _
                                                        curSCPromo, "")
	  Else
		  GetPrtStrSumm = GetPrtStrSumm & GetPrtSummDtl(byWidth, 6, "Discount:", _
                                                      curDisc, "")
        
	  End If

      '--VAT on Senior--------------------------------------------------------------
      If bSCD Then
        curVATSeniorExempt = (curSubttl - curVATSenior - curVATPWD) / 1 * iSCDCover * 0.8 / 1.12 ' .x_Chk.v_Cover
        If Abs(curSubttl) > MAX_CHECK_LIMIT_TO And _
          (.x_Chk.v_TxnTypeSeq <> DINE_IN_ID) Then
          curVATSeniorExempt = (MAX_CHECK_LIMIT_TO * iSCDCover) * 0.8 / 1.12
        End If
      End If
      '--VAT on PWD--------------------------------------------------------------
      If bPWD Then
        'No limit in all transaction for PWD 10/19/2017 by aries
		curVATPWDExempt = (curSubttl - curVATPWD) / 1 * iPWDCover * 0.8 / 1.12  '.x_Chk.v_Cover
      End If
	  If bSCD Then
	  GetPrtStrSumm = GetPrtStrSumm & GetPrtSummDtl(byWidth, 6, "Less SC Vat:", _
                                                    curVATSenior, "")
	  ElseIf bPWD Then
	  GetPrtStrSumm = GetPrtStrSumm & GetPrtSummDtl(byWidth, 6, "Less PWD Vat:", _
                                                    curVATPWD, "")
	  End If
	  '10/19/17
	  '--SC Promo Disc--------------------------------------------------------------------------------
      If bSCPromo Then
        curSCPromoExempt = curSubttl * iSCPromoCover / 1.12
      End If
    End If
    '--End Compute Discount--------------------------------------------------------------

    '--ADD ON Tax--------------------------------------------------------------
    bPrtedVat = .x_Txns.PrintedWithVATByTxn(lTxnIndex)
    sAddOn = ""
    If((m_Context.IsBitOn(sTndTdef, 14) Or m_Context.IsBitOn(.x_Rvc.sOptions, 105)) _
       And Not bPrtedVat And Not bNoSubTtl) Or ((sTndTdef = "" And bPrtVat) And Not bNoSubTtl) Then 
      For by = 1 To 8
        If (byCombine = 3 And by = 3) Or (byCombine = 4 And by = 4) Then by = by + 1
        If Trim(m_Context.x_PosMem.ItemTax(by).sName) <> "" And _
          (m_Context.x_PosMem.ItemTax(by).eType = 1 Or _
          m_Context.x_PosMem.ItemTax(by).eType = 4) Then
          If (m_Context.IsBitOn(sTndTdef, 14) And Not bPrtedVat) Or Not bTaxByRound Then
            curTmp = clsTxn.x_Tax.Value(by)
            curTaxable = clsTxn.x_Tax.TaxableAmount(by)
'comment for VAT
'            If by = 2 And byCombine <> 2 Then
'                curTmp = curTmp + clsTxn.x_Tax.Value(byCombine)
'                curTaxable = curTaxable + clsTxn.x_Tax.TaxableAmount(byCombine)
'            End If
          Else
            curTmp = clsTxn.x_Tax.Value(by) - clsTotal.x_Tax.Value(by)
            curTaxable = clsTxn.x_Tax.TaxableAmount(by) - clsTotal.x_Tax.TaxableAmount(by)
'comment for VAT
'            If by = 2 And byCombine <> 2 Then
'                curTmp = curTmp + clsTxn.x_Tax.Value(byCombine) - _
'                    clsTotal.x_Tax.Value(byCombine)
'                curTaxable = curTaxable + clsTxn.x_Tax.TaxableAmount(byCombine) - _
'                      clsTotal.x_Tax.TaxableAmount(byCombine)
'            End If
          End If
          If curTmp <> 0 Then
            sAddOn = sAddOn & GetPrtTax(by, byWidth, .x_Sys, _
                                  False, curTaxable, curTmp, bPadChinese)
            curAddOn = curAddOn + curTmp
          ElseIf .x_Chk.StatBit(by) Then
            sAddOn = sAddOn & GetPrtTax(by, byWidth, .x_Sys, _
                                  True, curTaxable, 0, bPadChinese)
            curAddOn = curAddOn + curTmp
          End If
        End If
      Next
    ElseIf (Len(clsDesc.sTaxName) > 0) And Not bNoSubTtl Then
      If bTaxByRound Then
        curTmp = clsTxn.v_TaxTtlAddOn - clsTotal.v_TaxTtlAddOn
      Else
        curTmp = clsTxn.v_TaxTtlAddOn
      End If
      If curTmp <> 0 Then
        GetPrtStrSumm = GetPrtStrSumm & _
                        GetPrtSummDtl(byWidth, 6, clsDesc.sTaxName, curTmp, "")
      End If
    End If
    GetPrtStrSumm = GetPrtStrSumm & sAddOn
    '--End ADD ON Tax--------------------------------------------------------------

    '--Service Charge--------------------------------------------------------------
    If (clsDesc.sServName <> "") And Not bNoSubTtl Then 
      If bTaxByRound Then
        curTmp = clsTxn.v_AutoSvc - clsTotal.v_AutoSvc
      Else
        curTmp = clsTxn.v_AutoSvc
      End If

      If curTmp <> 0 Then
        GetPrtStrSumm = GetPrtStrSumm & _
                        GetPrtSummDtl(byWidth, 6, clsDesc.sServName, curTmp, "")
      ElseIf m_Context.IsBitOn(sTndTdef, 14) And .x_Chk.StatBit(9) Then
        GetPrtStrSumm = GetPrtStrSumm & _
                        GetPrtSummDtl(byWidth, 4, clsDesc.sServName, curTmp, " *")
      End If
    End If
    '--End Service Charge--------------------------------------------------------------

    '--VAT Tax--------------------------------------------------------------
    bPrtedVat = .x_Txns.PrintedWithVATByTxn(lTxnIndex)
    If ((m_Context.IsBitOn(sTndTdef, 14) Or m_Context.IsBitOn(.x_Rvc.sOptions, 105)) And _
      Not bPrtedVat) Or ((sTndTdef = "" And bPrtVat) And Not bNoSubTtl) Then
      For by = 1 To 8
        If Trim(m_Context.x_PosMem.ItemTax(by).sName) <> "" And _
           ((m_Context.x_PosMem.ItemTax(by).eType = 2) Or _
           (m_Context.x_PosMem.ItemTax(by).eType = 3)) Then
          If (m_Context.IsBitOn(sTndTdef, 14) And Not bPrtedVat) Or Not bTaxByRound Then
            curTmp = clsTxn.x_Tax.Value(by)
            curTaxable = clsTxn.x_Tax.TaxableAmount(by)
'            If by = 2 And byCombine <> 2 Then
'              curTmp = curTmp + clsTxn.x_Tax.Value(byCombine)
'              curTaxable = curTaxable + clsTxn.x_Tax.TaxableAmount(byCombine)
'            End If              
          Else
            curTmp = clsTxn.x_Tax.Value(by) - clsTotal.x_Tax.Value(by)
            curTaxable = clsTxn.x_Tax.TaxableAmount(by) - clsTotal.x_Tax.TaxableAmount(by)
'            If by = 2 And byCombine <> 2 Then
'              curTmp = curTmp + clsTxn.x_Tax.Value(byCombine) - _
'                       clsTotal.x_Tax.Value(byCombine)
'              curTaxable = curTaxable + clsTxn.x_Tax.TaxableAmount(byCombine) - _
'                           clsTotal.x_Tax.TaxableAmount(byCombine)
'            End If
          End If

          If Abs(curTmp) > 0 Then
            cur12VAT = cur12VAT + curTmp
          End If

          If curTmp <> 0 Then           
            If Abs(curVATSenior + curVATPWD) Or Abs(curSCPromo) < Abs(curTmp) Then
              If Abs(curTmp + curVATSenior + curVATPWD + curSCPromo) > 0.1 Then
                sVat = sVat & GetPrtTax(by, byWidth, .x_Sys, _
                                        False, curTaxable, curTmp + curVATSenior + curVATPWD + curSCPromo, bPadChinese)
              Else
                sVat = sVat & GetPrtTax(by, byWidth, .x_Sys, _
                                        False, curTaxable, 0, bPadChinese)
              End If
            Else
              sVat = sVat & GetPrtTax(by, byWidth, .x_Sys, _
                                      False, curTaxable, 0, bPadChinese)
            End If
            curReturnVAT = curTmp                       
          ElseIf m_Context.x_Bill.x_Chk.StatBit(by) Then
            sVat = sVat & GetPrtTax(by, byWidth, .x_Sys, _
                                    True, curTaxable, 0, bPadChinese)                     
          End If
        End If
      Next
    End If
    cur12VAT = GetAmtStr(.x_Sys.byDecPt, cur12VAT)
    '--End VAT Tax--------------------------------------------------------------

    '--Tip--------------------------------------------------------------
    If (clsDesc.sTipName <> "") And Not bNoSubTtl Then
      If bPrevTxnRnd Then
        curTmp = .x_List.GetTxnOtherSvc(lActiveTxnRndSeq)
      Else
        curTmp = .x_Chk.v_OtherSvc
      End If
      If curTmp <> 0 Then
        GetPrtStrSumm = GetPrtStrSumm & _
                        GetPrtSummDtl(byWidth, 6, clsDesc.sTipName, curTmp, "")
      End If
    End If
    '--End Tip--------------------------------------------------------------

    '--Amount Due--------------------------------------------------------------
    sTmp = GetAmtStr(.x_Sys.byDecPt, .x_Chk.Total)
    lTmp = byWidth - 6
    If byWidth = 24 Then
      sTmp1 = FormatStr(clsDesc.sTtlDueName, lTmp, GetAmtStr(m_Context.x_Bill.x_Sys.byDecPt, sTmp), bPadChinese)
    Else
      If m_Context.IsBitOn(.x_Rvc.sOptions, 52) Or m_Context.IsBitOn(.x_Rvc.sOptions, 112) Then
        sTmp1 = clsDesc.sTtlDueName  '& GetAmtStr(m_Context.x_Bill.x_Sys.byDecPt, sTmp)
        iLen = m_Context.StrLenGraphically(sTmp1)
        iLen1 = m_Context.StrLenGraphically(sTmp)
        If (iLen * 2 + iLen1) > lTmp Then
          iLen = Len(sTmp) * 2
          If iLen > lTmp Then
            sTmp1 = FormatStr(clsDesc.sTtlDueName, lTmp, sTmp, bPadChinese)
            sTmp1 = m_Context.FormatPrintStr(DBH_H1W1, bPadChinese, sTmp1)
          Else
            lTmp = lTmp - iLen
            If lTmp = 0 Then
              sTmp1 = sTmp
            Else
              sTmp1 = m_Context.PadR(clsDesc.sTtlDueName, lTmp - 2, " ", 2) & _
                    " " & GetAmtStr(m_Context.x_Bill.x_Sys.byDecPt, sTmp)
            End If
            If m_Context.IsBitOn(.x_Rvc.sOptions, 112) Then
              sTmp1 = m_Context.FormatPrintStr(DBH_H1W2, bPadChinese, sTmp1)
            Else
              sTmp1 = m_Context.FormatPrintStr(DBH_H2W2, bPadChinese, sTmp1)
            End If
          End If
        Else
          lTmp = lTmp - (Len(sTmp1) * 2) - Len(sTmp)
          sTmp = m_Context.PadL(sTmp, Len(sTmp) + lTmp)
          If m_Context.IsBitOn(.x_Rvc.sOptions, 112) Then
            sTmp1 = m_Context.FormatPrintStr(DBH_H1W2, bPadChinese, sTmp1) & sTmp
          Else
            sTmp1 = m_Context.FormatPrintStr(DBH_H2W2, bPadChinese, sTmp1) & sTmp
          End If
        End If
      Else
        sTmp1 = FormatStr(clsDesc.sTtlDueName, lTmp, GetAmtStr(m_Context.x_Bill.x_Sys.byDecPt, sTmp), bPadChinese)
        sTmp1 = m_Context.FormatPrintStr(DBH_H1W1, bPadChinese, sTmp1)
      End If
    End If
    GetPrtStrSumm = GetPrtStrSumm & Space(3) & sTmp1 & vbCrLf & vbCrLf
'    GetPrtStrSumm = GetPrtStrSumm & GetPrtSummDtl(byWidth, 6, _
'                    clsDesc.sTtlDueName, .x_Chk.Total, sTmp)
'    GetPrtStrSumm = GetPrtStrSumm & GetPrtSummDtl(byWidth, 6, _
'                    m_Context.x_Pos.x_SysMsg.ChkTotal(), .x_Chk.Total, sTmp)

    If mcurComboSaving > 0 Then
      sComboSaving = FormatStr(m_Context.x_SysMsg.YouSaved(), byWidth - 7, _
                     Format(mcurComboSaving, m_Context.x_PosMem.ItemSystem.DecPtMask), _
                     bPadChinese)
      sComboSaving = m_Context.FormatPrintStr(DBH_H1W1, bPadChinese, sComboSaving)
  '      If IsBit(.x_Rvc.sOptions, 61) Then sComboSaving = InvertColor(sComboSaving)
      If m_Context.IsBitOn(.x_Rvc.sOptions, 61) Then
        GetPrtStrSumm = GetPrtStrSumm & Space(3) & sComboSaving & vbCrLf & ""
      Else
        GetPrtStrSumm = GetPrtStrSumm & Space(3) & sComboSaving & vbCrLf & ""
      End If
    End If

    If cur12VAT <> "" Then 
      For lLoop = 0 To m_Context.x_Bill.x_List.Count - 1
        Set clsItem = m_Context.x_Bill.x_List.Item(lLoop)

        If clsItem.x_Mi.v_Ret = False Then
          bReturn = False
          Exit For
        Else
          bReturn = True
          Exit For
        End If
      Next
	  
	  '10/19/17
	  If .x_Chk.StatBit(14) Or .x_Chk.StatBit(28) Then
		    curVATSeniorExempt = curVATSeniorExempt*-1
	  End If

      If Not .x_Chk.StatBit(14) Then
        If bReturn = False Then
          If bVATZero Then
			If Abs(.x_Chk.Total - cur12VAT - curAddOn - clsTxn.v_AutoSvc - curVATSeniorExempt - curSCPromoExempt - curVATSenior - curVATPWDExempt - curVATPWD - curSCPromo - curVatZeroRatedSales) > 0.1 Then
				GetPrtStrSumm = GetPrtStrSumm & _
                            GetPrtSummDtl(byWidth, 6, "VATable Sales", .x_Chk.Total - cur12VAT - curAddOn - clsTxn.v_AutoSvc - curVATSeniorExempt - curSCPromoExempt - curVATSenior - curVATPWDExempt - curVATPWD - curSCPromo - curVatZeroRatedSales, "")
			Else
                    GetPrtStrSumm = GetPrtStrSumm & _
                              GetPrtSummDtl(byWidth, 6, "VATable Sales", 0, "")
			End If
          Else
            If Not .x_Chk.Closed Then
				If Abs(.x_Chk.Total - cur12VAT - curAddOn - clsTxn.v_AutoSvc - curVATSeniorExempt - curSCPromoExempt - curVATSenior - curVATPWDExempt - curVATPWD - curSCPromo - curVatZeroRatedSales - curAdvDepo) > 0.1 Then
	              GetPrtStrSumm = GetPrtStrSumm & _
                              GetPrtSummDtl(byWidth, 6, "VATable Sales", .x_Chk.Total - cur12VAT - curAddOn - clsTxn.v_AutoSvc - curVATSeniorExempt - curSCPromoExempt - curVATSenior - curVATPWDExempt - curVATPWD - curSCPromo - curVatZeroRatedSales - curAdvDepo, "")

				Else
                    GetPrtStrSumm = GetPrtStrSumm & _
                              GetPrtSummDtl(byWidth, 6, "VATable Sales", 0, "")

				End If
            ElseIf(curAdvDepo > 0 And curLessDepo > 0) Then
				If Abs(.x_Chk.Total - cur12VAT - curAddOn - clsTxn.v_AutoSvc - curVATSeniorExempt - curSCPromoExempt - curVATSenior - curVATPWDExempt - curVATPWD - curSCPromo - curVatZeroRatedSales - curLessDepo) > 0.1 Then
					GetPrtStrSumm = GetPrtStrSumm & _
                              GetPrtSummDtl(byWidth, 6, "VATable Sales", .x_Chk.Total - cur12VAT - curAddOn - clsTxn.v_AutoSvc - curVATSeniorExempt - curSCPromoExempt - curVATSenior - curVATPWDExempt - curVATPWD - curSCPromo - curVatZeroRatedSales - curLessDepo, "")

				Else
                    GetPrtStrSumm = GetPrtStrSumm & _
                              GetPrtSummDtl(byWidth, 6, "VATable Sales", 0, "")

				End If
            Else
				
				If Abs(.x_Chk.Total - cur12VAT - curAddOn - clsTxn.v_AutoSvc - curVATSeniorExempt - curSCPromoExempt - curVATSenior - curSCPromo - curVATPWDExempt - curVATPWD - curVatZeroRatedSales) > 0.1 Then
					  GetPrtStrSumm = GetPrtStrSumm & _
									  GetPrtSummDtl(byWidth, 6, "VATable Sales", .x_Chk.Total - cur12VAT - curAddOn - clsTxn.v_AutoSvc - curVATSeniorExempt - curSCPromoExempt - curVATSenior - curSCPromo - curVATPWDExempt - curVATPWD - curVatZeroRatedSales, "")
				Else
					  GetPrtStrSumm = GetPrtStrSumm & _
									  GetPrtSummDtl(byWidth, 6, "VATable Sales", 0, "")
				End If			
			End If
          End If 'If bVATZero Then
        Else
          If Abs(.x_Chk.Total - curReturnVAT - curAddOn - clsTxn.v_AutoSvc - curVATSeniorExempt - curSCPromoExempt - curVATSenior - curSCPromo - curVATPWDExempt - curVATPWD) > 0.1 Then
              GetPrtStrSumm = GetPrtStrSumm & _
							  GetPrtSummDtl(byWidth, 6, "VATable Sales", .x_Chk.Total - curReturnVAT - curAddOn - clsTxn.v_AutoSvc - curVATSeniorExempt - curSCPromoExempt - curVATSenior - curSCPromo - curVATPWDExempt - curVATPWD, "")
          Else
              GetPrtStrSumm = GetPrtStrSumm & _
                              GetPrtSummDtl(byWidth, 6, "VATable Sales", 0, "")
          End If
	    End If 'If bReturn = False Then

	  End If 'If Not .x_Chk.StatBit(14) Then


	  
	  If sVat <> "" Then
		  If Right(GetPrtStrSumm, 2) = vbCrLf Then
			GetPrtStrSumm = GetPrtStrSumm & sVat
		  Else
			GetPrtStrSumm = GetPrtStrSumm & vbCrLf & sVat
		  End If
	  Else
		  GetPrtStrSumm = GetPrtStrSumm & _
						  GetPrtSummDtl(byWidth, 6, "VAT Amount", 0, "")
	  End If

        GetPrtStrSumm = GetPrtStrSumm & _
                        GetPrtSummDtl(byWidth, 6, "VAT Exempt Sales", curVATSeniorExempt + curVATPWDExempt + curSCPromoExempt - curSeniorDisc - curPWDDisc - curMCBDISC, "")
        GetPrtStrSumm = GetPrtStrSumm & _
                        GetPrtSummDtl(byWidth, 6, "VAT Zero-rated Sales", curVatZeroRatedSales , "")
	Else
		If sVat <> "" Then
		  If Right(GetPrtStrSumm, 2) = vbCrLf Then
			GetPrtStrSumm = GetPrtStrSumm & sVat
		  Else
			GetPrtStrSumm = GetPrtStrSumm & vbCrLf & sVat
		  End If
		Else
		  GetPrtStrSumm = GetPrtStrSumm & _
						  GetPrtSummDtl(byWidth, 6, "VAT Amount", 0, "")
		End If


	End If 'If cur12VAT <> "" Then 


'-----Payment-------------------------------------------------------------------------------

    For kLoop = 0 To .x_List.Count - 1
      Set clsItem = m_Context.x_Bill.x_List.Item(kLoop) 
      If (clsItem.v_DtlType = "T") Then
        If clsItem.x_Ref.Item(1) <> "" Then
          If clsItem.x_Tnd.v_Expiry = "" Then
            GetPrtStrSumm = GetPrtStrSumm & "   " & clsItem.x_Ref.Item(1) & vbCrLf
          Else
            GetPrtStrSumm = GetPrtStrSumm & "   " & clsItem.x_Ref.Item(1) & " " & _
              Right(clsItem.x_Tnd.v_Expiry, 2) & "/" & Left(clsItem.x_Tnd.v_Expiry, 2) & vbCrLf
          End If
        End If

         If clsItem.v_Price < 0 then
            Exit For

 	 End If
        GetPrtStrSumm = GetPrtStrSumm & GetPrtSummDtl(byWidth, 6, clsItem.v_Name, clsItem.v_Price, "")
      End If
    Next

'------Change Due---------------------------------------------------------------------------
    If .x_Chk.Closed Then
      GetPrtStrSumm = GetPrtStrSumm & sPaymentItems '& vbCrLf 
      curChgDue = .x_Chk.ChgDue
      If curChgDue <> 0 And bIncludeTnd Then
          GetPrtStrSumm = GetPrtStrSumm & _
                          GetPrtSummDtl(byWidth, 6, clsDesc.sChange, curChgDue, "")
      End If

	  'RESERVE FOR LOYALTY SCRIPT HERE
	  'LOYALTY START
	  
	  'LOYALTY END
	  	  
	  'OFFICIAL RECEIPT
      GetPrtStrSumm = GetPrtStrSumm & _
                      GetThaiTaxDetailStr(byWidth, bPadChinese)
'      If m_Context.IsBitOn(sTndTdef, 34) Then
'        GetPrtStrSumm = GetPrtStrSumm & Right(.x_Chk.v_Remark, 27) & vbCrLf
'      End If
      If m_Context.IsBitOn(sTndTdef, 30) Then
        iPrintCnt = 2
      End If
    End If
  End With
End Function


Function FormatStr(sPreStr, byLen, sAmt, bPadChinese)
  Dim sTmp
  On Error Resume Next
 
  FormatStr = m_Context.PadR(sPreStr, byLen - (Len(sAmt) + 1)) & " " & sAmt
End Function

Function GetPrtSummDtl(byWidth, byMargin, sDesc, _
                       curAmt, sExempt)
  Dim bPadChinese
  Dim sTemp
  
  On Error Resume Next

  bPadChinese = (m_Context.x_Pos.v_RegionalString = "CH")

  If byWidth = 24 Then
    GetPrtSummDtl = GetPrtSummDtl & Space(3) & _
                    FormatStr(sDesc, byWidth - byMargin, _
                              GetAmtStr(m_Context.x_Bill.x_Sys.byDecPt, curAmt) & sExempt, _
                              bPadChinese) & vbCrLf
  Else
    GetPrtSummDtl = GetPrtSummDtl & Space(3) & _
                    m_Context.FormatPrintStr(DBH_H1W1, bPadChinese,FormatStr(sDesc, byWidth - byMargin,GetAmtStr(m_Context.x_Bill.x_Sys.byDecPt, curAmt) & _
                    sExempt, bPadChinese)) & vbCrLf

  End If

End Function

Function GetAmtStr(byDec, curAmt)
  If curAmt >= 0 Then 
    GetAmtStr = m_Context.FormatVal(Abs(curAmt), DecPt(byDec)) & " "
  Else 
    GetAmtStr = m_Context.FormatVal(Abs(curAmt), DecPt(byDec)) & "-"
  End If
End Function

Function GetPrtTax(byTax, byWidth, clsSys, bExempted, _
                   curTaxable, curTax, bPadChinese)
  Dim sTaxable
  Dim sExempt
  Dim byMargin

  On Error Resume Next
 
  If bExempted Then
    byMargin = 4
    sExempt = " *"
  Else
    byMargin = 6
    sExempt = ""       
  End If
  
  If m_Context.IsBitOn(m_Context.x_Bill.x_Rvc.sOptions, 125) Then  'Print tax itemizer
'    byMargin = byMargin - 2
    sTaxable = GetAmtStr(clsSys.byDecPt, curTaxable) & " "
    If byWidth = 24 Then
      GetPrtTax = Space(3) & _
                  FormatStr(Trim(m_Context.x_PosMem.ItemTax(byTax).sName), _
                  byWidth - byMargin, _
                  GetAmtStr(clsSys.byDecPt, curTax) & sExempt, bPadChinese) & vbCrLf
    Else
      GetPrtTax = Space(3) & _
                  m_Context.FormatPrintStr(DBH_H1W1, bPadChinese, _
                  FormatStr(Trim(m_Context.x_PosMem.ItemTax(byTax).sName), _
                  byWidth - byMargin, _
                  GetAmtStr(clsSys.byDecPt, curTax) & sExempt, bPadChinese)) & vbCrLf          
    End If
  ElseIf byWidth = 24 Then
    GetPrtTax = Space(3) & _
                FormatStr(Trim(m_Context.x_PosMem.ItemTax(byTax).sName), _
                byWidth - byMargin, GetAmtStr(clsSys.byDecPt, curTax) & sExempt, _
                bPadChinese) & vbCrLf
  Else
    GetPrtTax = Space(3) & _
                m_Context.FormatPrintStr(DBH_H1W1, bPadChinese, _
                FormatStr(Trim(m_Context.x_PosMem.ItemTax(byTax).sName), _
                byWidth - byMargin, GetAmtStr(clsSys.byDecPt, curTax) & sExempt, _
                bPadChinese)) & vbCrLf             
  End If

End Function

Function GetPrtAddChkIndicator(byWidth, clsDesc, clsTxn)
  With m_Context.x_Bill
    If (.x_Chk.v_XferChkNum > 0) And (Left(clsTxn.v_Misc, 7) = "ADDEDTO") Then
      GetPrtAddChkIndicator = m_Context.PadC(Replace(m_Context.x_SysMsg.ChkAddedTo(), _
                              "%", clsDesc.sChk & CStr(.x_Chk.v_XferChkNum)), _
                               byWidth) & vbCrLf

      If Mid(clsTxn.v_Misc, 9) <> "" Then
        GetPrtAddChkIndicator = GetPrtAddChkIndicator & _
                                m_Context.PadC(Replace(m_Context.x_SysMsg.AddedToTbl(), _
                                "%", Mid(clsTxn.v_Misc, 9)), byWidth) & vbCrLf
      End If
    End If
  End With
End Function


Function GetPrtStrTrailer(byWidth, bChkClosed, clsTxn, _
                          clsTrailer, bPadChinese)
  Dim sTmp
  Dim by
  Dim byBlank
  Dim lTmp
  Dim sErr
  Dim sTrailerStr

  On Error Resume Next
  With m_Context.x_Bill
    sTrailerStr = ""
    'Call m_Context.showmsg("This 1")
    If bChkClosed Then
    ' Call m_Context.showmsg("This 2")
      sTmp = Replace(Replace(m_Context.x_Pos.x_SysMsg.ClosedAt(), "%", _
             .x_TxnEmp.lNum), "#", .x_Sys.DateTimeByRegion(Now))
      If m_Context.StrLenGraphically(sTmp) <= byWidth Then
        sTmp = m_Context.PadC(sTmp, byWidth, "-")
        If byWidth = 24 Then
          sTrailerStr = sTrailerStr & sTmp & vbCrLf
        Else
          sTrailerStr = sTrailerStr & _
                        m_Context.FormatPrintStr(DBH_H1W1, _
                        bPadChinese, sTmp) & vbCrLf
        End If
      Else
     '  Call m_Context.showmsg("This 3")
        sTmp = Trim(Replace(Replace(m_Context.x_Pos.x_SysMsg.ClosedAt(), "%", _
               .x_TxnEmp.lNum), "#", ""))
        sTmp = m_Context.PadC(sTmp, byWidth, "-")
        If byWidth = 24 Then
      '   Call m_Context.showmsg("This 4")
          sTrailerStr = sTrailerStr & sTmp & vbCrLf
        Else
      '   Call m_Context.showmsg("This 5")
          sTrailerStr = sTrailerStr & _
                        m_Context.FormatPrintStr(DBH_H1W1, _
                        bPadChinese, sTmp) & vbCrLf
        End If
    'Call m_Context.showmsg("This 6")
        sTmp = Trim(.x_Sys.DateTimeByRegion(.x_Chk.v_ChkCls))
        sTmp = m_Context.PadC(sTmp, byWidth, "-")
        If byWidth = 24 Then
        ' Call m_Context.showmsg("This 7")
          sTrailerStr = sTrailerStr & sTmp & vbCrLf
        Else
        ' Call m_Context.showmsg("This 8")
          sTrailerStr = sTrailerStr & _
                        m_Context.FormatPrintStr(DBH_H1W1, _
                        bPadChinese, sTmp) & vbCrLf
        End If
      End If
    Else
      'Call m_Context.showmsg("This 9")
      If m_Context.IsBitOn(.x_Rvc.sOptions, 16) Then
        'Call m_Context.showmsg("This 10")
        sTmp = m_Context.x_Pos.x_SysMsg.StoredAt()
        sTmp = Replace(Replace(sTmp, "%", .x_TxnEmp.lNum), "#", _
               .x_Sys.DateTimeByRegion(clsTxn.v_EndTime))
        If m_Context.StrLenGraphically(sTmp) <= byWidth Then
        ' Call m_Context.showmsg("This 11") 
          sTmp = m_Context.PadC(sTmp, byWidth, "-")
          If byWidth = 24 Then
         '  Call m_Context.showmsg("This 12")
            sTrailerStr = sTrailerStr & sTmp & vbCrLf
          Else
         '  Call m_Context.showmsg("This 13")
            sTrailerStr = sTrailerStr & _
                          m_Context.FormatPrintStr(DBH_H1W1, _
                          bPadChinese, sTmp) & vbCrLf
          End If
        Else
        ' Call m_Context.showmsg("This 14")
          sTmp = Trim(Replace(Replace(sTmp, "%", .x_TxnEmp.lNum), "#", ""))
          sTmp = m_Context.PadC(sTmp, byWidth, "-")
          If byWidth = 24 Then
         '  Call m_Context.showmsg("This 15")
            sTrailerStr = sTrailerStr & sTmp & vbCrLf
          Else
          ' Call m_Context.showmsg("This 16")
            sTrailerStr = sTrailerStr & _
                          m_Context.FormatPrintStr(DBH_H1W1, _
                          bPadChinese, sTmp) & vbCrLf
          End If
          'Call m_Context.showmsg("This 17")
          sTmp = Trim(.x_Sys.DateTimeByRegion(clsTxn.v_EndTime))
          sTmp = m_Context.PadC(sTmp, byWidth, "-")
          If byWidth = 24 Then
            'Call m_Context.showmsg("This 18")
            sTrailerStr = sTrailerStr & sTmp & vbCrLf
          Else
            'Call m_Context.showmsg("This 19")
            sTrailerStr = sTrailerStr & _
                          m_Context.FormatPrintStr(DBH_H1W1, _
                          bPadChinese, sTmp) & vbCrLf
          End If
        End If
      Else
        'Call m_Context.showmsg("This 20")
        sTrailerStr = sTrailerStr & String(byWidth, "-") & vbCrLf
      End If
    End If
  End With
  'Call m_Context.showmsg("This 21")  
  GetPrtStrTrailer = sTrailerStr & _
                     GetPrtTrailer(byWidth, clsTrailer, bPadChinese)  
                      

End Function

Function GetPrtTrailer(byWidth, clsTrailer, bPadChinese)
  Dim sTmp
  Dim by
  Dim byBlank
  Dim sErr
  Dim byTemp
  Dim bCtrlChar
       
  On Error Resume Next
  
  If Not clsTrailer Is Nothing Then
    For by = 1 To 12
      If Trim(clsTrailer.GstTrailer(by)) = "" Then
        byBlank = byBlank + 1
      Else
        While byBlank > 0
          GetPrtTrailer = GetPrtTrailer & vbCrLf
          byBlank = byBlank - 1
        Wend
        If clsTrailer.GstDW(by) Then
          byTemp = byWidth / 2
        Else
          byTemp = byWidth
        End If
        sTmp = m_Context.PadC(Trim(clsTrailer.GstTrailer(by)), byTemp)
        If clsTrailer.GstDW(by) Then   
          sTmp = GetPrtChar(sTmp, DBH_H1W2, byWidth, bPadChinese)
        Else
          sTmp = GetPrtChar(sTmp, DBH_H1W1, byWidth, bPadChinese)
        End If 
'        sTmp = HeaderEnhance(sTmp, clsTrailer.GstInvColor(by), bCtrlChar)
        GetPrtTrailer = GetPrtTrailer & sTmp & vbCrLf
      End If
    Next
'    GetPrtTrailer = GetPrtTrailer & IIf(bCtrlChar, "" & vbCrLf & "", "")
  End If
End Function


Function DecPt(byDec)
     Dim lCnt
     On Error Resume next

   If byDec = 0 Then
     DecPt = "0" 
   Else
     DecPt = "#,##0" & "." & String(byDec, "0")
   End If     
End Function

Public Sub checkersub(cnn, Context, txnnum , rst, rst0, rst1, rst2, rst3)
	Set cnn = CreateObject("ADODB.COnnection")
        cnn.ConnectionTimeout = CONNTIMEOUT
        cnn.Open POS_SERVER & ";User Id=datascan;Password=DTSbsd7188228"	
	
	'Check if there is an existing setting
	Set rst = CreateObject("ADODB.recordset")
	rst.Open "Select * from Settings",cnn

	Dim checker
	checker = False
	
	Dim CheckNumberOrig
	Dim CheckNumber
	CheckNumberOrig = CStr(Context.x_Bill.x_Chk.v_ChkNum)
	CheckNumber = CStr(Right("00000" & CheckNumberOrig, 5))

	If Not rst.EOF Then
		'Has Settings
		'Call Context.ShowMsg("Settings", "Has settings", True, 0, 1)
		
		'Check if date now is within start and end date
		If Now >= rst.Fields("StartDate").Value And Now <= rst.Fields("EndDate").Value Then
			'Call Context.ShowMsg("Date Range", "Within date range", True, 0, 1)
	                'Date now falls within date range

			Set rst1 = CreateObject("ADODB.recordset")
			rst1.Open "Select * from Settings",cnn
		
	                'Check which settings are enabled
        	        If rst1.Fields("ByTransactions").Value = -1 And rst1.Fields("ByMinPurchase").Value = -1 = False Then
                	    'Only By transactions is enabled, check if transaction setting is equal to counter

			    cnn.Execute "UPDATE Settings SET Counter = Counter + 1"
			    Set rst2 = CreateObject("ADODB.recordset")
			    rst2.Open "Select * from Settings",cnn
			    'Call Context.ShowMsg("Settings", "Only by transactions is enabled", True, 0, 1)

			    'Reset counter
                	    If rst2.Fields("Counter").Value = rst2.Fields("Transactions").Value Then
				'Call Context.ShowMsg("Generate code", "code generated 1", True, 0, 1)
				checker = True
                        	cnn.Execute "UPDATE Settings SET Counter = 0"
	                    ElseIf rst2.Fields("Counter").Value > rst2.Fields("Transactions").Value Then
				'Call Context.ShowMsg("Generate code", "code generated 2", True, 0, 1)
        	                checker = True
                        	cnn.Execute "UPDATE Settings SET Counter = " & rst2.Fields("Counter").Value - rst2.Fields("Transactions").Value
			    Else
				'counter < transactions
				'Call Context.ShowMsg("Generate code", "code not generated 1", True, 0, 1)
	                    End If

			    rst2.Close
			    Set rst2 = Nothing

        	        ElseIf rst1.Fields("ByTransactions").Value = False And rst1.Fields("ByMinPurchase").Value = True Then
                	    'Only By minimum purchase is enabled
			    'Call Context.ShowMsg("Settings", "Only by minimun purchase is enabled", True, 0, 1)

	                    'Check if check amount >= minamount
			    Set rst0 = CreateObject("ADODB.recordset")
	       		    rst0.Open "SELECT * FROM checks WHERE chk_num = " & CheckNumberOrig, cnn
			    'Call Context.ShowMsg("Settings", "Only by minimun purchase is enabled 2", True, 0, 1)

	                    If CDbl(rst0.Fields("sub_ttl").Value) >= CDbl(rst1.Fields("MinAmount").Value) Then
				'Call Context.ShowMsg("Generate code", "code generated 3", True, 0, 1)
        	                checker = True
			    Else
				'Call Context.ShowMsg("Generate code", "code not generated 2", True, 0, 1)
                	    End If

			    rst0.Close
			    Set rst0 = Nothing

	                ElseIf rst1.Fields("ByTransactions").Value = True And rst1.Fields("ByMinPurchase").Value = True Then
        	            'Both are enabled
			    'Call Context.ShowMsg("Settings", "Both settings are enabled", True, 0, 1)

                	    'Check if check transactions = counter, amount >= minamount
			    Set rst0 = CreateObject("ADODB.recordset")
	       		    rst0.Open "SELECT * FROM checks WHERE chk_num = " & CheckNumberOrig, cnn

	                    If CDbl(rst0.Fields("sub_ttl").Value) >= CDbl(rst1.Fields("MinAmount").Value) Then
				cnn.Execute "UPDATE Settings SET Counter = Counter + 1"
				'Call Context.ShowMsg("Generate code", "minimum amount met", True, 0, 1)
				
				set rst3 = CreateObject("ADODB.recordset")
				rst3.Open "Select * from Settings",cnn

				If rst3.Fields("Counter").Value = rst3.Fields("Transactions").Value Then
					checker = True
					'Call Context.ShowMsg("Generate code", "Code generated 5", True, 0, 1)
	                        	cnn.Execute "UPDATE Settings SET Counter = 0"
        	            	ElseIf rst3.Fields("Counter").Value > rst3.Fields("Transactions").Value Then
					checker = True
					'Call Context.ShowMsg("Generate code", "Code generated 6", True, 0, 1)
                	        	cnn.Execute "UPDATE Settings SET Counter = " & rst3.Fields("Counter").Value - rst3.Fields("Transactions").Value
				Else
					'Call Context.ShowMsg("Generate code", "Code notgenerated 7", True, 0, 1)
                    		End If

				rst3.Close
				Set rst3 = Nothing

			    Else
				checker = False
				'Call Context.ShowMsg("Generate code", "minimum amount not met", True, 0, 1)
				'Call Context.ShowMsg("Generate code", "code not generated 3", True, 0, 1)
	                    End If

			    rst0.Close
			    Set rst0 = Nothing

        	        Else
                	    'No settings are enabled
			    'Call Context.ShowMsg("Settings", "No settings are enabled", True, 0, 1)
	                    checker = False
        		End If

			rst1.Close
			Set rst1 = Nothing
		Else
			'Outside date range
			'Call Context.ShowMsg("Date Range", "outside date range", True, 0, 1)
		End If
	Else
		'No Settings
		'Call Context.ShowMsg("Settings", "No settings", True, 0, 1)
	End If
	
	'If checker = True method
	If checker = True Then
	    Set rst1 = CreateObject("ADODB.recordset")
	    rst1.Open "SELECT propertyno FROM System", cnn

	    Dim StoreCode
	    Dim DateTimeNow
	    Dim TimeNow
	    Dim DateToday

	    StoreCode = CStr(Right("0" & rst1.Fields("propertyno").Value, 4))
	    DateTimeNow = Now
	    TimeNow = CStr(Right("0" & Hour(DateTimeNow), 2) & Right("0" & Minute(DateTimeNow), 2))
	    DateToday = CStr(Right("0" & Month(DateTimeNow), 2) & Right("0" & Day(DateTimeNow), 2) & Year(DateTimeNow))

	    rst1.Close
	    Set rst1 = Nothing

	    Dim first
	    Dim second
            Dim third
            Dim fourth
            Dim fifth
            Dim sixth
            Dim seventh
            Dim eighth
            Dim ninth
            Dim tenth
            Dim eleventh
            Dim twelfth
            Dim thirteenth
            Dim fourteenth
            Dim fifteenth
	    Dim surveycode

	    first = Mid(CheckNumber, 5, 1)
	    second = Mid(TimeNow, 2, 1)
	    third = Mid(StoreCode, 4, 1)
	    fourth = Mid(CheckNumber, 4, 1)
	    fifth = Mid(StoreCode, 2, 1)
	    sixth = Mid(DateToday, 2, 1)
	    seventh = Mid(CheckNumber, 3, 1)
	    eighth = Mid(StoreCode, 1, 1)
	    ninth = Mid(CheckNumber, 2, 1)
            If CInt(Mid(DateToday, 1, 2)) <= 9 Then
            	tenth = "2"
            Else
            	tenth = "3"
            End If
	    eleventh = Mid(DateToday, 4, 1)
	    twelfth = Mid(StoreCode, 3, 1)
	    thirteenth = Mid(DateToday, 3, 1)
	    fourteenth = Mid(TimeNow, 1, 1)
	    fifteenth = Mid(CheckNumber, 1, 1)

            surveycode = first & second & third & fourth & fifth & sixth & seventh & eighth & ninth & tenth & eleventh & twelfth & thirteenth & fourteenth & fifteenth
	    'Call Context.ShowMsg("Code", "TXN NUMBER:" & CheckNumberOrig & " CHECKNUMBERCUT:" & CheckNumber & " STORE CODE:" & StoreCode & " TIMENOW:" & TimeNow & " DATETODAY:" & DateToday & " SURVEYCODE:" & surveycode, True, 0, 1)
	    cnn.Execute "INSERT INTO SurveyCode VALUES ('" & surveycode & "', " & CLng(CheckNumberOrig) & ", '" & CDate(DateTimeNow) & "')"
	    Context.v_Globalvar("surveycode") = vbNewLine & vbNewLine & Context.PadC("Tell us about your visit", 40) & vbNewLine & Context.PadC("and get a [free smoothie]", 40) & vbNewLine & Context.PadC("on your next purchase.", 40) & vbNewLine & vbNewLine & Context.PadC("Go to www.jambasatisfaction.com", 40) & vbNewLine & Context.PadC("within 7 days of your visit", 40) & vbNewLine & Context.PadC("to complete the survey.", 40) & vbNewLine & vbNewLine & Context.PadC("Survey code: " & surveycode, 40) & vbNewLine & vbNewLine
	Else
		Context.v_Globalvar("surveycode") = ""
        End If

	rst.Close
	set rst = Nothing

        cnn.Close 
        Set cnn = Nothing

End Sub

Public Sub repref(cnn, Context, rst)
	'Call Context.ShowMsg("REPREF", "1", True, 0, 1)
	Set cnn = CreateObject("ADODB.COnnection")
        cnn.ConnectionTimeout = CONNTIMEOUT
        cnn.Open POS_SERVER & ";User Id=Datascan;Password=DTSbsd7188228"	

	'Call Context.ShowMsg("REPREF", "2", True, 0, 1)

	Set rst = CreateObject("ADODB.recordset")
	rst.Open "Select * from SurveyCode where CheckNumber = " & CStr(Context.x_Bill.x_Chk.v_ChkNum),cnn

	'Call Context.ShowMsg("REPREF", "3", True, 0, 1)

	If Not rst.EOF Then
		'Call Context.ShowMsg("REFREF", "HAS SURVEY CODE", True, 0, 1)
		'Call Context.ShowMsg("REFREF", CStr(rst.Fields("SurveyCode").Value), True, 0, 1)
		Context.v_Globalvar("surveycode") = vbNewLine & vbNewLine & Context.PadC("Tell us about your visit", 40) & vbNewLine & Context.PadC("and get a [free smoothie]", 40) & vbNewLine & Context.PadC("on your next purchase.", 40) & vbNewLine & vbNewLine & Context.PadC("Go to www.jambasatisfaction.com", 40) & vbNewLine & Context.PadC("within 7 days of your visit", 40) & vbNewLine & Context.PadC("to complete the survey.", 40) & vbNewLine & vbNewLine & Context.PadC("Survey code: " & CStr(rst.Fields("SurveyCode").Value), 40) & vbNewLine & vbNewLine
	Else
		'Call Context.ShowMsg("REFREF", "NO SURVEY CODE", True, 0, 1)
		Context.v_Globalvar("surveycode") = ""
	End If

	'Call Context.ShowMsg("REPREF", "4", True, 0, 1)

	rst.Close
	set rst = Nothing

        cnn.Close 
        Set cnn = Nothing
End Sub

Function GetPOS_ProviderInfo()
	On Error Resume Next
	GetPOS_ProviderInfo = GetPOS_ProviderInfo & vbCrLf
	GetPOS_ProviderInfo = GetPOS_ProviderInfo & GetBIRINFO_TRAILER(TRAILER_BIRINFO_NUMBER_FOOTER)
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

Function CatchError(source)
  If Err.Number <> 0 Then
    Call m_Context.ShowMsg("Error Occur in " & source, Err.Description, True, 0, 3)
    CatchError = True
    m_PosConn.Rollback
  End If
End Function

Public Function SqlStr(vStr)
'Convert Null String to "''"
  If IsNull(vStr) Then
    SqlStr = "''"
  Else
    SqlStr = "'" & Replace(vStr, "'", "''") & "'"
  End If
End Function

Function GetBIRINFO_TRAILER(TRAILER_NUMBER)

On Error Resume Next

 Dim sSql
 Dim rstRec
 
  'On Error Resume Next
  If TRAILER_NUMBER = "" Then
	TRAILER_NUMBER = 0
  End If
  sSql = "select * from trailer where number=" & TRAILER_NUMBER

  Call OpenPosConnection(m_PosConn)
  Set rstRec = OpenRecordset(sSql,m_PosConn)

  If Not rstRec.EOF Then
	rstRec.MoveFirst
	If Trim(rstRec.Fields("guesttrailer1")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("guesttrailer1")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("guesttrailer2")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("guesttrailer2")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("guesttrailer3")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("guesttrailer3")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("guesttrailer4")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("guesttrailer4")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("guesttrailer5")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("guesttrailer5")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("guesttrailer6")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("guesttrailer6")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("guesttrailer7")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("guesttrailer7")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("guesttrailer8")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("guesttrailer8")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("guesttrailer9")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("guesttrailer9")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("guesttrailer10")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("guesttrailer10")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("guesttrailer11")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("guesttrailer11")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("guesttrailer12")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("guesttrailer12")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("cctrailer1")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("cctrailer1")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("cctrailer2")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("cctrailer2")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("cctrailer3")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("cctrailer3")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("cctrailer4")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("cctrailer4")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("cctrailer5")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("cctrailer5")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("cctrailer6")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("cctrailer6")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("cctrailer7")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("cctrailer7")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("cctrailer8")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("cctrailer8")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("cctrailer9")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("cctrailer9")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("cctrailer10")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("cctrailer10")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("cctrailer11")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("cctrailer11")), PRINTERWIDTH) & vbCrlf
	End IF
	If Trim(rstRec.Fields("cctrailer12")) <> "" Then
		GetBIRINFO_TRAILER = GetBIRINFO_TRAILER & m_Context.PadC(Trim(rstRec.Fields("cctrailer12")), PRINTERWIDTH) & vbCrlf
	End IF
  End If

  Call CloseAdoObj(m_PosConn)

  
End Function

