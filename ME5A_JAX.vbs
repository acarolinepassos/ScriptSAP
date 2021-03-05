'**************************************************************************************************
'*                                                                                                *
'*          Desenvolvido por: Ana Caroline Passos                                                 *
'*          Utilizacao: Extracao da base de RCs ME5A_JAX						                      *
'*                                                                                                *
'**************************************************************************************************



'============================ Extracao SAP =========================================
    
    'definicao de variaveis
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set xlApp = CreateObject("Excel.Application")
	Set WshShell = WScript.CreateObject("WScript.Shell")

	Call ExecuteScript

    Function ExtractReport()
        OpenSAPLogon
        ConnectToSAP("EDP - Defence & Security (JAX)")
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
		session.findById("wnd[0]/tbar[0]/okcd").text = "ME5A"
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]/tbar[1]/btn[17]").press
		session.findById("wnd[1]/usr/txtENAME-LOW").text = "diorodri"
		session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
		session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 8
		session.findById("wnd[1]/tbar[0]/btn[8]").press
		session.findById("wnd[0]/tbar[1]/btn[8]").press
		session.findById("wnd[0]/tbar[1]/btn[33]").press
		session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell 52,"TEXT"
		session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 43
		session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "52"
		session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
		session.findById("wnd[0]/tbar[1]/btn[45]").press
		session.findById("wnd[1]/tbar[0]/btn[0]").press
		session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Temp\RCSPOT\"
		session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ME5A_JAX.txt"
		session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
		session.findById("wnd[1]/tbar[0]/btn[11]").press
		session.findById("wnd[0]/tbar[0]/btn[3]").press
		session.findById("wnd[0]/tbar[0]/btn[3]").press
		
		session.findById("wnd[0]").close
		On Error Resume Next
		session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

	End Sub
	
	