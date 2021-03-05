'**************************************************************************************************
'*                                                                                                *
'*          Desenvolvido por: Ana Caroline Passos                                                 *
'*          Utilizacao: Extracao da base de RCs ME5A_BR						                      *
'*                                                                                                *
'**************************************************************************************************



'============================ Extracao SAP ========================================================
    
    'definicao de variaveis
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set xlApp = CreateObject("Excel.Application")
	Set WshShell = WScript.CreateObject("WScript.Shell")

	Call ExecuteScript

    Function ExtractReport()
        OpenSAPLogon
        ConnectToSAP("01- EBP - SAP Corp")
		If Session.findById("wnd[0]/sbar").messagetype = "W" Then
			Session.findById("wnd[0]").sendVKey 0
		End If
        ExecuteScript
    End Function


    Sub OpenSAPLogon

        Set WshShell = WScript.CreateObject("WScript.Shell")
        sapLogonPadPath = """C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplgpad.exe"""
        WshShell.Run(sapLogonPadPath)

    End Sub

    Sub ConnectToSAP(site)

        wscript.sleep(3000)
        Set SapGui = GetObject("SAPGUI")
        Set Appl = SapGui.GetScriptingEngine
        Set Connection = Appl.Openconnection(site, True)

    End Sub

	Sub ExecuteScript
	
		'primeira extracao ME5A_BR
		If Not IsObject(application) Then
		   Set SapGuiAuto  = GetObject("SAPGUI")
		   Set application = SapGuiAuto.GetScriptingEngine
		End If
		If Not IsObject(connection) Then
		   Set connection = application.Children(0)
		End If
		If Not IsObject(session) Then
		   Set session    = connection.Children(0)
		End If
		If IsObject(WScript) Then
		   WScript.ConnectObject session,     "on"
		   WScript.ConnectObject application, "on"
		End If
		session.findById("wnd[0]").maximize
		session.findById("wnd[0]/tbar[0]/okcd").text = "me5a"
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]/tbar[1]/btn[17]").press
		session.findById("wnd[1]/usr/txtENAME-LOW").text = "diorodri"
		session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
		session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 8
		session.findById("wnd[1]/tbar[0]/btn[8]").press
		session.findById("wnd[0]/tbar[1]/btn[8]").press
		session.findById("wnd[0]/tbar[1]/btn[33]").press
		session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 664
		session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").setFocus
		session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell 761,"TEXT"
		session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 754
		session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "761"
		session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
		session.findById("wnd[0]/tbar[1]/btn[45]").press
		session.findById("wnd[1]/tbar[0]/btn[0]").press
		session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
		session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
		session.findById("wnd[1]").sendVKey 4
		session.findById("wnd[2]").close
		session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 43
		session.findById("wnd[1]").sendVKey 4
		session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus
		session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 0
		session.findById("wnd[2]").sendVKey 4
		session.findById("wnd[3]").close
		session.findById("wnd[2]").close
		session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Temp\RCSPOT\"
		session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ME5A_BR_COM.txt"
		session.findById("wnd[1]/usr/ctxtDY_FILENAME").setFocus
		session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 11
		session.findById("wnd[1]/tbar[0]/btn[11]").press
		session.findById("wnd[0]/tbar[0]/btn[3]").press
		session.findById("wnd[0]/tbar[0]/btn[3]").press
		
		'Segunda extracao ME5A_BR_CRI
		If Not IsObject(application) Then
		   Set SapGuiAuto  = GetObject("SAPGUI")
		   Set application = SapGuiAuto.GetScriptingEngine
		End If
		If Not IsObject(connection) Then
		   Set connection = application.Children(0)
		End If
		If Not IsObject(session) Then
		   Set session    = connection.Children(0)
		End If
		If IsObject(WScript) Then
		   WScript.ConnectObject session,     "on"
		   WScript.ConnectObject application, "on"
		End If
		session.findById("wnd[0]").maximize
		session.findById("wnd[0]/tbar[0]/okcd").text = "ME5A"
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]/tbar[1]/btn[17]").press
		session.findById("wnd[1]/usr/txtENAME-LOW").text = "diorodri"
		session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
		session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 8
		session.findById("wnd[1]/tbar[0]/btn[8]").press
		session.findById("wnd[0]/tbar[1]/btn[16]").press
		session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").selectNode "        159"
		session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").topNode = "        154"
		session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").doubleClickNode "        159"
		session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN002-LOW").setFocus
		session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN002-LOW").caretPosition = 0
		session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN002_%_APP_%-VALU_PUSH").press
		session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "CRI"
		session.findById("wnd[1]/tbar[0]/btn[8]").press
		session.findById("wnd[0]/tbar[1]/btn[8]").press
		session.findById("wnd[0]/tbar[1]/btn[33]").press
		session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 664
		session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").setFocus
		session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell 761,"TEXT"
		session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 754
		session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "761"
		session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
		session.findById("wnd[0]/tbar[1]/btn[45]").press
		session.findById("wnd[1]/tbar[0]/btn[0]").press
		session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Temp\RCSPOT\"
		session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ME5A_BR_COM_CRI.txt"
		session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 15
		session.findById("wnd[1]/tbar[0]/btn[11]").press
		session.findById("wnd[0]/tbar[0]/btn[3]").press
		session.findById("wnd[0]/tbar[0]/btn[3]").press
		
		'terceira extracao ME5A_BR_AOG
		If Not IsObject(application) Then
		   Set SapGuiAuto  = GetObject("SAPGUI")
		   Set application = SapGuiAuto.GetScriptingEngine
		End If
		If Not IsObject(connection) Then
		   Set connection = application.Children(0)
		End If
		If Not IsObject(session) Then
		   Set session    = connection.Children(0)
		End If
		If IsObject(WScript) Then
		   WScript.ConnectObject session,     "on"
		   WScript.ConnectObject application, "on"
		End If
		session.findById("wnd[0]").maximize
		session.findById("wnd[0]/tbar[0]/okcd").text = "ME5A"
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]/tbar[1]/btn[17]").press
		session.findById("wnd[1]/usr/txtENAME-LOW").text = "diorodri"
		session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
		session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 8
		session.findById("wnd[1]/tbar[0]/btn[8]").press
		session.findById("wnd[0]/tbar[1]/btn[16]").press
		session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").selectNode "        159"
		session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").topNode = "        154"
		session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").doubleClickNode "        159"
		session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN002-LOW").setFocus
		session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/ctxt%%DYN002-LOW").caretPosition = 0
		session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN002_%_APP_%-VALU_PUSH").press
		session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "AOG"
		session.findById("wnd[1]/tbar[0]/btn[8]").press
		session.findById("wnd[0]/tbar[1]/btn[8]").press
		session.findById("wnd[0]/tbar[1]/btn[33]").press
		session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 664
		session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").setFocus
		session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell 761,"TEXT"
		session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 754
		session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "761"
		session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
		session.findById("wnd[0]/tbar[1]/btn[45]").press
		session.findById("wnd[1]/tbar[0]/btn[0]").press
		session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Temp\RCSPOT\"
		session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ME5A_BR_COM_AOG.txt"
		session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 15
		session.findById("wnd[1]/tbar[0]/btn[11]").press
		session.findById("wnd[0]/tbar[0]/btn[3]").press
		session.findById("wnd[0]/tbar[0]/btn[3]").press
		
		session.findById("wnd[0]").close
		On Error Resume Next
		session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

	End Sub
	
	