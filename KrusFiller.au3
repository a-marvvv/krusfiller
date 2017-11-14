#include <MsgBoxConstants.au3>
#include <GUIConstants.au3>
#include "UIAutomate.au3"
#include <Array.au3>
#NoTrayIcon
#Region
#AutoIt3Wrapper_Compression=4
#EndRegion
Opt("WinTitleMatchMode", 2)
Global Const $Rus = '00000419'; Раскладка русского языка
Global Const $Eng = '00000409'; Раскладка английского языка
Global Const $Win_Name = "Карточка устройства "
Global Const $info_str="Заполнялка КРУС v0.1" & @CRLF & "Telegram: @DonAlexey"
Global Const $config = @ScriptDir & "\config.ini"
Monitor()

Func Monitor()
While 1
	If WinExists($Win_Name, '') Then
		Global $hWnd = WinGetHandle($Win_Name, '')
		Global $oParent = _UIA_GetElementFromHandle($hWnd)
		_WinAPI_SetKeyboardLayout($Eng, $hWnd)
		CreateMenu($hWnd)
	EndIf
	Sleep(100)
 WEnd
EndFunc

Func CheckTab($d)
    $s = StringSplit($d, "_")
	SelectTab($s[1])
   EndFunc

Func Fill()
	  $Data = GetData()
	  $CData = GetIni()
	  If $Data[1] <> $CData[1][1] Then
		 MsgBox($MB_ICONWARNING, "", "В буфере обмена данные не для " & $CData[1][1])
		 Return
	  EndIf

	  for $i = 2 To $CData[0][0]
		 $Name = $CData[$i][0]
		 $ControlID = $CData[$i][1]
		 CheckTab($Name)
		 if StringInStr($ControlID, "TDBEdit") Then
			SetControlTextCustom($ControlID,$Data[$i])
		 Else
			SetControlText($ControlID,$Data[$i])
		 EndIf
	  Next
	  MsgBox($MB_ICONWARNING, "", "Заполнено")
   EndFunc

   Func GetIni()
	  return IniReadSection($config, GetDev())
	  EndFunc

Func GetDev()
   return StringMid(WinGetTitle($hWnd),21)
EndFunc

Func GetData()
   $sArray = StringSplit(ClipGet(), "|")
   Return $sArray
EndFunc

Func CheckDev()
   IniReadSection($config, GetDev())
   If @error Then
	  MsgBox($MB_ICONWARNING, "", "Карточка устройства не поддерживается")
	  Return 1
   Else
	  Return 0
	  EndIf
EndFunc

Func CreateMenu($hGUI)

  $aPos = ControlGetPos($hGUI, "", "Ок")
  $gui = GUICreate("", $aPos[2]-10, $aPos[3], $aPos[0] - 100, $aPos[1], $WS_POPUP)
  DllCall("user32.dll", "hwnd", "SetParent", "hwnd", $gui, "hwnd", $hGUI)
  $menu = GUICtrlCreateMenu("Заполнить")
  $m_item1  = GUICtrlCreateMenuItem ("Из буфера обмена", $menu)
  GUICtrlCreateMenuitem ("",$menu) ; separator
  $m_item_i = GUICtrlCreateMenuitem ("Info",$menu)
  GUISetState()
  GUICtrlSetState($menu, $GUI_FOCUS)
  While 1
	 If not WinExists ($hGUI) Then ExitLoop
    Switch GUIGetMsg()
      Case $m_item1
        If Not CheckDev() Then Fill()
	 Case $m_item_i
	  MsgBox($MB_ICONWARNING, "", $info_str)
    EndSwitch
  WEnd
   EndFunc

Func SelectTab($Tab_Name)
   $oElement = _UIA_GetControlTypeElement($oParent, "UIA_TabItemControlTypeId", $Tab_Name)
   _UIA_ElementDoDefaultAction($oElement)
EndFunc

Func SetControlText($Control,$Text)
   WinActivate($hWnd)
   ControlSetText($hWnd, "",$Control,$Text)
EndFunc

Func SetControlTextCustom($Control,$Text)
   WinActivate($hWnd)
   ControlFocus($hWnd,"",$Control)
   Send($Text)
EndFunc

Func _WinAPI_SetKeyboardLayout($sLayout, $hWnd)
	If Not WinExists($hWnd) Then
		Return SetError(1, 0, 0)
	EndIf
	Local $Ret = DllCall('user32.dll', 'long', 'LoadKeyboardLayout', 'str', StringFormat('%08s', StringStripWS($sLayout, 8)), 'int', 0)
	If (@error) Or ($Ret[0] = 0) Then
		Return SetError(1, 0, 0)
	EndIf
	DllCall('user32.dll', 'ptr', 'SendMessage', 'hwnd', $hWnd, 'int', 0x0050, 'int', 1, 'int', $Ret[0])
	Return SetError(0, 0, 1)
EndFunc 
