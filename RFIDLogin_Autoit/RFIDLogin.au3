#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile=RFIDLogin.exe
#AutoIt3Wrapper_Outfile_x64=RFIDLogin_x64.exe
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_UseX64=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;~ #RequireAdmin

#include <AutoItConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#Include <WinAPI.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIListBox.au3>
#include <TabConstants.au3>
#include <GuiListView.au3>
#include <MsgBoxConstants.au3>
#include <File.au3>
#include <Array.au3>
#include <GuiEdit.au3>
#include <TrayConstants.au3>
#include "MD5.au3"
#include "CommMG.au3"

Opt("TrayMenuMode", 3)
Opt("TrayOnEventMode", 1)

HotKeySet("+^c", "RunSettingForm")

Local $oMyError = ObjEvent("AutoIt.Error", "_Error")

$sDBFile = @ScriptDir & "\DatabaseSettings.mdb"
$sDBUsername = "Admin"
$sDBPassword = "1234578"
$Baudrate = 9600
$AdPassword = MD5("159875321598753698741236")

Local $NowLogin = False, $OnSettingRun = False
Local $Form4, $Form4Button1, $Form4Label2, $hTimer, $last, $Form3ListView1, $Form1, $Form2
$adoCon = ObjCreate("ADODB.Connection")

If Not FileExists($sDBFile) Then
	MsgBox(16, "ERROR", "ไม่พบไฟล์ฐานข้อมูล")
	Exit -1
EndIf

DBBegin()
$query = "SELECT AdPass FROM config WHERE ID = 1;"
$adoRs = ObjCreate("ADODB.Recordset")
$adoRs.CursorType = 1
$adoRs.LockType = 3
$adoRs.Open($query, $adoCon)
$AdPassword = $adoRs.Fields("AdPass").value
DBEnd()

TraySetOnEvent($TRAY_EVENT_PRIMARYDOUBLE, "TrayEvent")

Local $sportSetError, $find = False, $comName

While Not $find
	LoadDevByCOM()

	If Not $find Then
		$iMsgBoxAnswer = MsgBox($MB_CANCELTRYCONTINUE + $MB_ICONQUESTION, "ไม่พบอุปกรณ์", "ไม่พบอุปกรณ์ RFID เชื่อมต่ออยู่ คุณต้องการดำเนินการต่อหรือไม่ ?")
		If $iMsgBoxAnswer = $IDCONTINUE Then
			SetingForm()
		ElseIf $iMsgBoxAnswer = $IDTRYAGAIN Then
			ContinueLoop
		EndIf
		Exit
	EndIf
WEnd
While 1
	LoginRFID()
	ShowFormLogout()
WEnd



Func ShowFormLogout()
	$NowLogin = True
	#Region ### START Koda GUI section ### Form=C:\Users\Max\Dropbox\Autoit\RFIDLogin\Form4.kxf
	$Form4 = GUICreate("เวลาทำงาน", 210, 66, @DesktopWidth - 220, @DesktopHeight - 120, BitOR($WS_SYSMENU,$WS_POPUP), $WS_EX_TOOLWINDOW)
	$Form4Button1 = GUICtrlCreateButton("ออกจากระบบ", 120, 8, 83, 49)
	$Form4Label1 = GUICtrlCreateLabel("เวลาเข้าสู่ระบบ", 8, 8, 73, 17)
	$Form4Label2 = GUICtrlCreateLabel("00:00:00", 8, 24, 94, 33)
	GUICtrlSetFont(-1, 18, 400, 0, "MS Sans Serif")
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	$hTimer = TimerInit()
	$last = 0

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Exit
			Case $Form4Button1
				ExitLoop

		EndSwitch

		$fDiff = TimerDiff($hTimer)
		If $fDiff - $last >= 1000 Then
			$last = $fDiff
			$sec = Int(Mod($fDiff / 1000, 60))
			$min = Int(Mod($fDiff / 1000 / 60, 60))
			$hour = Int($fDiff / 1000 / 60 / 60)
			If $sec < 10 Then $sec = "0" & String($sec)
			If $min < 10 Then $min = "0" & String($min)
			If $hour < 10 Then $hour = "0" & String($hour)
			GUICtrlSetData($Form4Label2, $hour & ":" & $min & ":" & $sec)
		EndIf
	WEnd
	GUIDelete($Form4)
EndFunc

Func LoginRFID()
	$NowLogin = False
	#Region ### START Koda GUI section ### Form=C:\Users\Max\Dropbox\Autoit\RFIDLogin\Form.kxf
	$Form1 = GUICreate("Login By RFID", @DesktopWidth, @DesktopHeight, -1, -1, BitOR($WS_MAXIMIZE,$WS_POPUP,$DS_SETFOREGROUND))
;~ 	$Form1 = GUICreate("Login By RFID", 600, 600, -1, -1)
	$Form1Label1 = GUICtrlCreateLabel("แตะบัตรนักศึกษา", 136, 112, 538, 108, BitOR($SS_CENTER,$SS_CENTERIMAGE))
	GUICtrlSetFont(-1, 80, 400, 0, "Angsana New")
	GUICtrlSetColor(-1, 0xFF0000)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetPos(-1, 0, 0, @DesktopWidth, @DesktopHeight  - 100)
	$Form1Label2 = GUICtrlCreateLabel("เพื่อปลดล็อกเครื่องคอมพิวเตอร์", 200, 224, 406, 69, BitOR($SS_CENTER,$SS_CENTERIMAGE))
	GUICtrlSetFont(-1, 36, 400, 0, "Angsana New")
	GUICtrlSetColor(-1, 0x008000)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetPos(-1, 0, 100, @DesktopWidth, @DesktopHeight  - 100)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

;~ 	$pID = Run("taskmgr.exe", "", @SW_DISABLE)
	Opt('WINTITLEMATCHMODE', 4)
	ControlHide('classname=Shell_TrayWnd', '', '')

	_CommClearOutputBuffer()
    _CommClearInputBuffer()

	While 1
		If _CommGetInputCount() > 0 Then
			$rfid = getRFID()
			If IsString($rfid) And $rfid <> "" Then
				DBBegin()
				$query = "SELECT COUNT(*) As Count FROM users WHERE uID = '" & $rfid & "';"
				$adoRs = ObjCreate("ADODB.Recordset")
				$adoRs.CursorType = 1
				$adoRs.LockType = 3
				$adoRs.Open($query, $adoCon)
				$Count = $adoRs.Fields("Count").value
				DBEnd()
				If $Count >= 1 Then
					ExitLoop
				Else
					GUICtrlSetData($Form1Label2, "บัตรไม่ถูกต้อง ลองแตะใหม่อีกครั้ง")
				EndIf
			EndIf
		Else
			$comName = _CommPortConnection()
			If @error Then
				$find = False
				While Not $find
					LoadDevByCOM()
					If $find = False Then
						MsgBox(16, "ERROR", "ไม่พบการเชื่อมต่อเครื่อง RFID")
						SetingForm()
					EndIf
				WEnd
			EndIf
		EndIf

;~ 		Lock screen
		WinSetOnTop($Form1, "", $WINDOWS_ONTOP)
		WinActivate($Form1)
;~ 		BlockInput($BI_DISABLE)
;~ 		_WinAPI_ShowCursor(False)
		MouseMove(0, 0, 1)
	WEnd

;~ 	BlockInput($BI_ENABLE)
;~ 	_WinAPI_ShowCursor(True)
	ControlShow('classname=Shell_TrayWnd', '', '')

;~ 	ProcessClose($pID)

	GUIDelete($Form1)
EndFunc

Func SetingForm()
	#Region ### START Koda GUI section ### Form=c:\users\max\dropbox\autoit\rfidlogin\form2.kxf
	$Form2 = GUICreate("การตั้งค่า", 475, 277, -1, -1)
	$Form2Button1 = GUICtrlCreateButton("เข้าตั้งค่า", 168, 144, 155, 57)
	GUICtrlSetFont(-1, 26, 400, 0, "Angsana New")
	$Form2Input1 = GUICtrlCreateInput("", 112, 64, 273, 45, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER,$ES_PASSWORD))
	GUICtrlSetFont(-1, 23, 400, 0, "MS Sans Serif")

	$Enter_KEY = GUICtrlCreateDummy()
	Dim $Arr[1][2] = [["{ENTER}", $Enter_KEY]]
	GUISetAccelerators($Arr)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	GUICtrlSetState($Form2Input1, $GUI_FOCUS)
	WinSetOnTop($Form1, "", $WINDOWS_NOONTOP)
	WinSetOnTop($Form2, "", $WINDOWS_ONTOP)
	WinActivate($Form2)
	WinSetOnTop($Form2, "", $WINDOWS_NOONTOP)

	$Login = False

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $Form2Button1, $Enter_KEY
				If GUICtrlRead($Form2Input1) = "" Then
					MsgBox(16, "ERROR", "กรุณากรอกรหัสผ่าน")
					GUICtrlSetState($Form2Input1, $GUI_FOCUS)
					_GUICtrlEdit_SetSel(GUICtrlGetHandle($Form2Input1), 0, -1)
					ContinueLoop
				EndIf
				If MD5(GUICtrlRead($Form2Input1)) = $AdPassword Then
					$Login = True
					ExitLoop
				Else
					MsgBox(16, "ERROR", "รหัสผ่านไม่ถูกต้อง")
				EndIf

			Case $GUI_EVENT_CLOSE
				ExitLoop

		EndSwitch
		UpTime()
	WEnd
	GUIDelete($Form2)
	If $Login = True Then FormSetting()
EndFunc

Func FormSetting()
	#Region ### START Koda GUI section ### Form=c:\users\max\dropbox\autoit\rfidlogin\form3.kxf
	$Form3 = GUICreate("การตั้งค่า", 479, 281, -1, -1)
	$Form3Tab1 = GUICtrlCreateTab(8, 8, 465, 265)
	$Form3TabSheet1 = GUICtrlCreateTabItem("จัดการผู้ใช้งาน")
	$Form3ListView1 = GUICtrlCreateListView("", 16, 40, 449, 175)
	$Form3Button1 = GUICtrlCreateButton("เพิ่มใหม่", 16, 232, 75, 25)
	$Form3Button2 = GUICtrlCreateButton("ลบ", 112, 232, 75, 25)
	$Form3Button3 = GUICtrlCreateButton("ลบทั้งหมด", 208, 232, 75, 25)
	$Form3TabSheet2 = GUICtrlCreateTabItem("ตั้งค่า")
	$Form3Button4 = GUICtrlCreateButton("เปลี่ยนรหัสผ่าน", 32, 56, 419, 49)
	GUICtrlSetFont(-1, 18, 400, 0, "Angsana New")
	$Form3Button5 = GUICtrlCreateButton("ออกจากโปรแกรม", 29, 127, 419, 49)
	GUICtrlSetFont(-1, 18, 400, 0, "Angsana New")
	$Form3Button6 = GUICtrlCreateButton("เกี่ยวกับโปรแกรม", 30, 195, 419, 49)
	GUICtrlSetFont(-1, 18, 400, 0, "Angsana New")
	GUICtrlCreateTabItem("")
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	_GUICtrlListView_SetExtendedListViewStyle($Form3ListView1, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_CHECKBOXES, $LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT, $LVS_EX_SUBITEMIMAGES))
	_GUICtrlListView_AddColumn($Form3ListView1, "RFID", 160)
	_GUICtrlListView_AddColumn($Form3ListView1, "NAME", 250)

	DBBegin()
	$query = "SELECT * FROM users;"
	$adoRs = ObjCreate("ADODB.Recordset")
	$adoRs.CursorType = 1
	$adoRs.LockType = 3
	$adoRs.Open($query, $adoCon)

	$RowsCount = $adoRs.RecordCount
	If $RowsCount > 0 Then
		$adoRs.MoveFirst
		$i = 0
		While Not $adoRs.EOF
			_GUICtrlListView_AddItem($Form3ListView1, $adoRs.Fields("uID").value)
			_GUICtrlListView_AddSubItem($Form3ListView1, $i, $adoRs.Fields("uName").value, 1)
			$i = $i + 1
			$adoRs.MoveNext
		WEnd
	EndIf
	DBEnd()

;~ 	WinSetOnTop($Form3, "", $WINDOWS_ONTOP)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $Form3Button1
				WinSetOnTop($Form3, "", $WINDOWS_NOONTOP)
				GUISetState(@SW_DISABLE, $Form3)
				AddUser()
				GUISetState(@SW_ENABLE, $Form3)
				WinActivate($Form3)
			Case $Form3Button2
				GUICtrlSetState($Form3Button2, $GUI_DISABLE)
				$DeleteCount = 0
				$InSQL = ""
				$ItemCount = _GUICtrlListView_GetItemCount($Form3ListView1)
				For $i = $ItemCount - 1 To 0 Step -1
					If _GUICtrlListView_GetItemChecked($Form3ListView1, $i) = True Then
						$RFID = StringStripCR(_GUICtrlListView_GetItemText($Form3ListView1, $i))
						_GUICtrlListView_DeleteItem(GUICtrlGetHandle($Form3ListView1), $i)
						$InSQL = $InSQL & "'" & $RFID & "',"
						$DeleteCount = $DeleteCount + 1
					EndIf
				Next
				$InSQL = StringTrimRight($InSQL, 1)
				If $DeleteCount > 0 Then
					DBExe("DELETE * FROM users WHERE uID IN (" & $InSQL &");")
				Else
					MsgBox(16, "ERROR", "โปรดเลือกรายการที่ต้องการลบ")
				EndIf
				GUICtrlSetState($Form3Button2, $GUI_ENABLE)

			Case $Form3Button3
				$ItemCount = _GUICtrlListView_GetItemCount($Form3ListView1)
				If $ItemCount <= 0 Then
					MsgBox(16, "ผิดพลาด", "ไม่มีข้อมูลผู้ใช้ในตาราง")
					ContinueLoop
				EndIf

				$iMsgBoxAnswer = MsgBox(36, "ลบผู้ใช้", "คุณต้องการลบผู้ใช้งานทั้งหมดหรือไม่ ?")
				If $iMsgBoxAnswer = 7 Then ContinueLoop

				GUICtrlSetState($Form3Button3, $GUI_DISABLE)
				DBExe("DELETE * FROM users;")
				_GUICtrlListView_DeleteAllItems(GUICtrlGetHandle($Form3ListView1))
				GUICtrlSetState($Form3Button3, $GUI_ENABLE)

			Case $Form3Button4
				$PassOne = InputBox("กรอกรหัสผ่าน", "กรอกรหัสผ่านใหม่", "", "•")
				If @error Then ContinueLoop
				$PassTwo = InputBox("กรอกรหัสผ่าน", "กรอกรหัสผ่านใหม่อีกครั้ง", "", "•")
				If @error Then ContinueLoop
				If $PassOne = $PassTwo Then
					GUICtrlSetState($Form3Button4, $GUI_DISABLE)
					DBExe("UPDATE config SET AdPass = '" & MD5($PassOne) &"' WHERE ID = 1;")
					$AdPassword = MD5($PassOne)
					GUICtrlSetState($Form3Button4, $GUI_ENABLE)
					MsgBox($MB_ICONINFORMATION + $MB_OK, "สำเร็จ", "เปลี่ยนแปลงรหัสผ่านเรียบร้อยแล้ว")

				Else
					MsgBox(16, "ERROR", "รหัสผ่านไม่ตรงกัน")
				EndIf
			Case $Form3Button5
				$comName = _CommPortConnection()
				If Not @error Then
					_CommClosePort()
				EndIf
				ControlShow('classname=Shell_TrayWnd', '', '')
				Exit
			Case $Form3Button6
				MsgBox($MB_ICONINFORMATION + $MB_OK, "เกี่ยวกับ", ".:: โปรแกรม RFID Login PC ::." & @CRLF & _
					"     โปรแกรมนี้ถูกพัฒนาขึ้นด้วยภาษา Autoit มุ่งเน้นเพื่อศึกษาการเชื่อมต่อฮาร์ดแวร์เพื่อสื่อสาร และควบคุมคอมพิวเตอร์ โปรแกรมนี้เป็นเพียงตัวอย่างไม่ได้ถูกแบบมาให้ใช้งานได้จริง ผู้ใช้ควรคำนึถึงความปลอดภัยเป็นสำคัญ" & @CRLF & _
					@CRLF & _
					"ผู้พัฒนา: นายสนธยา  นงนุช" & @CRLF & _
					"Facebook: fb.me/maxthai" & @CRLF & _
					"ร้าน IOXhop จำหน่ายบอร์ด Arduino RFID : www.ioxhop.com" & @CRLF)

			Case $GUI_EVENT_CLOSE
				ExitLoop

		EndSwitch
		UpTime()
	WEnd
	GUIDelete($Form3)
EndFunc

Func AddUser()
	#Region ### START Koda GUI section ### Form=C:\Users\Max\Dropbox\Autoit\RFIDLogin\Form6.kxf
	$Form6 = GUICreate("เพิ่มผู้ใช้ใหม่", 414, 258, -1, -1)
	$Form6Label1 = GUICtrlCreateLabel("หมายเลข RFID", 16, 24, 83, 17)
	GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
	$Form6Label2 = GUICtrlCreateLabel("กรุณาแตะบัตร", 16, 48, 382, 47, $SS_CENTER)
	GUICtrlSetFont(-1, 23, 800, 0, "Angsana New")
	GUICtrlSetBkColor(-1, 0xFFFFFF)
	$Form6Label3 = GUICtrlCreateLabel("ชื่อ", 16, 120, 19, 17)
	GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
	$Form6Input1 = GUICtrlCreateInput("", 16, 144, 377, 28)
	GUICtrlSetFont(-1, 12, 400, 0, "MS Sans Serif")
	GUICtrlSetState(-1, $GUI_DISABLE)
	$Form6Button1 = GUICtrlCreateButton("บันทึก", 232, 208, 75, 33)
	GUICtrlSetState(-1, $GUI_DISABLE)
	$Form6Button2 = GUICtrlCreateButton("ยกเลิก", 320, 208, 75, 33)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $Form6Button1
				$rfid = GUICtrlRead($Form6Label2)
				$name = GUICtrlRead($Form6Input1)
				If $name = "" Then
					MsgBox(16, "ERROR", "กรุณากรอกชื่อ")
					ContinueLoop
				EndIf
				GUICtrlSetState($Form6Button1, $GUI_DISABLE)
				DBExe("INSERT INTO users (uID, uName) VALUES ('" & $rfid & "', '" & $name & "');")
				$ItemCount = _GUICtrlListView_GetItemCount($Form3ListView1)
				_GUICtrlListView_AddItem($Form3ListView1, $rfid)
				_GUICtrlListView_AddSubItem($Form3ListView1, $ItemCount, $name, 1)
				GUICtrlSetState($Form6Button1, $GUI_ENABLE)
				MsgBox($MB_ICONINFORMATION + $MB_OK, "สำเร็จ", "เพิ่มผู้ใช้งานแล้ว")
				ExitLoop
			Case $GUI_EVENT_CLOSE, $Form6Button2
				ExitLoop

		EndSwitch
		If _CommGetInputCount() > 0 Then
			$rfid = getRFID()

			If $rfid <> 0 And $rfid <> -1 Then
				GUICtrlSetData($Form6Label2, $rfid)
				GUICtrlSetState($Form6Input1, $GUI_ENABLE)
				GUICtrlSetState($Form6Button1, $GUI_ENABLE)
			EndIf
		Else
			$comName = _CommPortConnection()
			If @error Then
				LoadDevByCOM()
				If $find = False Then
					MsgBox(16, "ERROR", "ไม่พบการเชื่อมต่อเครื่อง RFID")
					ExitLoop
				EndIf
			EndIf
		EndIf
		UpTime()
	WEnd
	GUIDelete($Form6)
EndFunc

Func UpTime()
	If $NowLogin = True Then
		$fDiff = TimerDiff($hTimer)
		If $fDiff - $last >= 1000 Then
			$last = $fDiff
			$sec = Int(Mod($fDiff / 1000, 60))
			$min = Int(Mod($fDiff / 1000 / 60, 60))
			$hour = Int($fDiff / 1000 / 60 / 60)
			If $sec < 10 Then $sec = "0" & String($sec)
			If $min < 10 Then $min = "0" & String($min)
			If $hour < 10 Then $hour = "0" & String($hour)
			GUICtrlSetData($Form4Label2, $hour & ":" & $min & ":" & $sec)
		EndIf
	EndIf
EndFunc

Func DBBegin()
	$adoCon = ObjCreate("ADODB.Connection")
	$adoCon.Open("Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & $sDBFile, $sDBUsername, $sDBPassword)
EndFunc

Func DBEnd()
	$adoCon.Close
EndFunc

Func DBExe($query)
	DBBegin()
	$adoCon.Execute($query, $adoCon)
	DBEnd()
EndFunc

Func TestPort($portname)
	$setport = StringReplace($portname, 'COM', '')
	$resOpen = _CommSetPort($setport, $sportSetError, $Baudrate, 8, 0, 1, 0, 2, 2)
	If $resOpen = 0 Then
		_CommClosePort()
		Return False
	Else
		_CommSendString("RF?" & @CRLF, 1)
		$line = _CommGetLine(@CRLF, 0, 500)
		If @error Or Not $line = "Y" Then
;~ 			_CommClosePort()
			Return False
		EndIf
	EndIf
	Return True
EndFunc

Func LoadDevByCOM()
	$portlist = _CommListPorts(0)
	$comName = "COM0"
	For $pl = 1 To $portlist[0]
		If TestPort($portlist[$pl]) = True Then
			$find = True
			$comName = $portlist[$pl]
			ExitLoop
		EndIf
	Next
EndFunc

Func getRFID()
	Return _CommGetLine(@CR, 0, 50)
EndFunc

Func RunSettingForm()
	If $OnSettingRun = False Then
		$OnSettingRun = True
		GUISetState($Form4Button1, $GUI_DISABLE)
		SetingForm()
		GUICtrlSetState($Form4Button1, $GUI_ENABLE)
		$OnSettingRun = False
	Else
		WinSetOnTop($Form1, "", $WINDOWS_NOONTOP)
		WinSetOnTop($Form2, "", $WINDOWS_ONTOP)
		WinActivate($Form2)
		WinSetOnTop($Form2, "", $WINDOWS_NOONTOP)
	EndIf
EndFunc

Func TrayEvent()
	Switch @TRAY_ID
		Case $TRAY_EVENT_PRIMARYDOUBLE
			RunSettingForm()

	EndSwitch
EndFunc   ;==>TrayEvent

Func _Error()
	$HexNumber = Hex($oMyError.Number, 8)
	MsgBox(16, "ERROR", "Error Number: " & $HexNumber & @CRLF & _
		"WinDescription: " & $oMyError.WinDescription & @CRLF & _
		"Script Line: " & $oMyError.ScriptLine & @CRLF)
	Exit -1
EndFunc