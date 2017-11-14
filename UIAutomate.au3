#include-once
#include <UIAConstants.au3>

; ============================================================================================
; ��������      : UIAutomate.au3
; ������ AutoIt : 3.3.10.0 +
; ������ UDF    : 1.8
; ��������      : ����� ������� ��� ������ � ���������� GUI ����������� API UIAutomation
; �����         : InnI
; ============================================================================================

; #����������_����������# ====================================================================
Global $UIA_ConsoleWriteError = 1 ; ����� �������� ������ � �������
Global $UIA_DefaultWaitTime   = 5 ; ����� �������� �� ��������� � ��������
Global $UIA_ElementVersion    = 1 ; ������ ��������: 0-����; 1-WIN_7 � ����; 2-WIN_8; 3-WIN_81; 4,5-WIN_10; 6,7-WIN_10(1703)
; ============================================================================================

; #�������# ==================================================================================
; _UIA_CreateLogicalCondition        ������ ���������� ������� �� ������ �������� �������
; _UIA_CreatePropertyCondition       ������ ������� �� ������ �������� � ��� ��������
; _UIA_ElementDoDefaultAction        ���������� ��������� �������� �� ���������
; _UIA_ElementFindInArray            ������� �������, ��������������� ��������� �������� � ��� ��������
; _UIA_ElementGetBoundingRectangle   ���������� ������������� �������, �������������� �������
; _UIA_ElementGetFirstLastChild      ������� ������ � ��������� �������� �������� (�������) ���������� ��������
; _UIA_ElementGetParent              ���������� ������������ ������� (������) ���������� ��������
; _UIA_ElementGetPreviousNext        ������� ���������� � ��������� �������� (�������) ���� �� ������
; _UIA_ElementGetPropertyValue       ���������� �������� ��������� �������� ��������
; _UIA_ElementMouseClick             ��������� ���� ���� �� ��������
; _UIA_ElementScrollIntoView         ������������ ������� � ������� ���������
; _UIA_ElementSetFocus               ������������� �������� ����� �����
; _UIA_ElementTextSetValue           ������������� �������� (�����) � ��������� �������
; _UIA_FindAllElements               ������� ��� ��������, ��������������� ��������� �������� � ��� ��������
; _UIA_FindAllElementsEx             ������� ��� ��������, ��������������� ������� ������
; _UIA_FindElementsInArray           ������� ��� ��������, ��������������� ��������� �������� � ��� ��������
; _UIA_GetControlTypeElement         ������� ������� (������) ���������� ���� � �������� ��������� � ���������
; _UIA_GetElementFromCondition       ������� ������� (������) �� ������ ��������� �������
; _UIA_GetElementFromFocus           ������ ������� (������) �� ������ ������ �����
; _UIA_GetElementFromHandle          ������ ������� (������) �� ������ �����������
; _UIA_GetElementFromPoint           ������ ������� (������) �� ������ �������� ���������
; _UIA_ObjectCreate                  ������ ������ UIAutomation
; _UIA_WaitControlTypeElement        ������� ������� (������) ���������� ���� � �������� ��������� � ���������
; _UIA_WaitElementFromCondition      ������� ������� (������) �� ������ ��������� �������
; ============================================================================================

; #���_�����������_�������������# ============================================================
; __UIA_ConsoleWriteError
; __UIA_CreateElement
; __UIA_GetPropIdFromStr
; __UIA_GetTypeIdFromStr
; ============================================================================================

; ============================================================================================
; ��� ������� : _UIA_CreateLogicalCondition
; ��������    : ������ ���������� ������� �� ������ �������� �������
; ���������   : _UIA_CreateLogicalCondition($oCondition1[, $sOperator = "NOT"[, $oCondition2 = 0]])
; ���������   : $oCondition1 - ������ �������
;             : $sOperator   - ���������� �������� (������)
;             :              "NOT" - ���������� �� (�� ���������)
;             :              "AND" - ���������� �
;             :              "OR"  - ���������� ���
;             : $oCondition2 - ������ �������
; ����������  : �����   - �������
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - ������ ������� �� ����������
;             :         @error = 2 - ������ �������� ������� UIAutomation
;             :         @error = 3 - ������ ������� �� ����������
;             :         @error = 4 - ������ �������� �������
;             :         @error = 5 - ����������� �������� ��������
; �����       : InnI
; ����������  : ��� ������������� ��������� "NOT" ������ ������� ������������
; ============================================================================================
Func _UIA_CreateLogicalCondition($oCondition1, $sOperator = "NOT", $oCondition2 = 0)
  If Not IsObj($oCondition1) Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_CreateLogicalCondition : ������ ������� �� ����������"), 0)
  Local $pCondition, $oCondition, $oUIAutomation = _UIA_ObjectCreate()
  If @error Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_CreateLogicalCondition : ������ �������� ������� UIAutomation"), 0)
  Switch $sOperator
    Case "NOT"
      $oUIAutomation.CreateNotCondition($oCondition1, $pCondition)
      $oCondition = ObjCreateInterface($pCondition, $sIID_IUIAutomationNotCondition, $dtagIUIAutomationNotCondition)
      If Not IsObj($oCondition) Then Return SetError(4, __UIA_ConsoleWriteError("_UIA_CreateLogicalCondition : ������ �������� �������"), 0)
    Case "AND"
      If Not IsObj($oCondition2) Then Return SetError(3, __UIA_ConsoleWriteError("_UIA_CreateLogicalCondition : ������ ������� �� ����������"), 0)
      $oUIAutomation.CreateAndCondition($oCondition1, $oCondition2, $pCondition)
      $oCondition = ObjCreateInterface($pCondition, $sIID_IUIAutomationAndCondition, $dtagIUIAutomationAndCondition)
      If Not IsObj($oCondition) Then Return SetError(4, __UIA_ConsoleWriteError("_UIA_CreateLogicalCondition : ������ �������� �������"), 0)
    Case "OR"
      If Not IsObj($oCondition2) Then Return SetError(3, __UIA_ConsoleWriteError("_UIA_CreateLogicalCondition : ������ ������� �� ����������"), 0)
      $oUIAutomation.CreateOrCondition($oCondition1, $oCondition2, $pCondition)
      $oCondition = ObjCreateInterface($pCondition, $sIID_IUIAutomationOrCondition, $dtagIUIAutomationOrCondition)
      If Not IsObj($oCondition) Then Return SetError(4, __UIA_ConsoleWriteError("_UIA_CreateLogicalCondition : ������ �������� �������"), 0)
    Case Else
      Return SetError(5, __UIA_ConsoleWriteError("_UIA_CreateLogicalCondition : ����������� �������� �������� (" & $sOperator & ")"), 0)
  EndSwitch
  Return $oCondition
EndFunc ; _UIA_CreateLogicalCondition

; ============================================================================================
; ��� ������� : _UIA_CreatePropertyCondition
; ��������    : ������ ������� �� ������ �������� � ��� ��������
; ���������   : _UIA_CreatePropertyCondition($vProperty, $vPropertyValue)
; ���������   : $vProperty      - �������� ��������
;             : $vPropertyValue - �������� �������� ��������
; ����������  : �����   - �������
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - ������ �������� ������� UIAutomation
;             :         @error = 2 - ������ �������������� ��������
;             :         @error = 3 - ������ �������� ������� (��. ���������� � ���������)
; �����       : InnI
; ����������  : �������� �������� ����� ����������� �� ����� ����� ������ ������� Inspect
;             : �������� �������� ����� ����������� �� ������ ����� ������ ������� Inspect
;             : �������� ������� ������������. ��������, ��� �������� "IsEnabled" ����� ��������� �������� True, � �� 1 � �� "True"
; ============================================================================================
Func _UIA_CreatePropertyCondition($vProperty, $vPropertyValue)
  Local $pCondition, $oCondition, $oUIAutomation = _UIA_ObjectCreate()
  If @error Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_CreatePropertyCondition : ������ �������� ������� UIAutomation"), 0)
  $vProperty = __UIA_GetPropIdFromStr($vProperty, "_UIA_CreatePropertyCondition")
  If $vProperty = -1 Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_CreatePropertyCondition : ������ �������������� ��������"), 0)
  $oUIAutomation.CreatePropertyCondition($vProperty, $vPropertyValue, $pCondition)
  $oCondition = ObjCreateInterface($pCondition, $sIID_IUIAutomationPropertyCondition, $dtagIUIAutomationPropertyCondition)
  If Not IsObj($oCondition) Then Return SetError(3, __UIA_ConsoleWriteError("_UIA_CreatePropertyCondition : ������ �������� ������� (��. ���������� � ���������)"), 0)
  Return $oCondition
EndFunc ; _UIA_CreatePropertyCondition

; ============================================================================================
; ��� ������� : _UIA_ElementDoDefaultAction
; ��������    : ���������� ��������� �������� �� ���������
; ���������   : _UIA_ElementDoDefaultAction($oElement)
; ���������   : $oElement - ������� (������)
; ����������  : �����   - 1
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - �������� �� �������� ��������
;             :         @error = 2 - ������ �������� ������� �� ������ ������� LegacyIAccessible
;             :         @error = 3 - ������ ���������� ������
; �����       : InnI
; ����������  :
; ============================================================================================
Func _UIA_ElementDoDefaultAction($oElement)
  If Not IsObj($oElement) Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_ElementDoDefaultAction : �������� �� �������� ��������"), 0)
  Local $pIAccess, $oIAccess
  $oElement.GetCurrentPattern($UIA_LegacyIAccessiblePatternId, $pIAccess)
  $oIAccess = ObjCreateInterface($pIAccess, $sIID_IUIAutomationLegacyIAccessiblePattern, $dtagIUIAutomationLegacyIAccessiblePattern)
  If Not IsObj($oIAccess) Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_ElementDoDefaultAction : ������ �������� ������� �� ������ ������� LegacyIAccessible"), 0)
  Local $iErrorCode = $oIAccess.DoDefaultAction()
  If $iErrorCode Then Return SetError(3, __UIA_ConsoleWriteError("_UIA_ElementDoDefaultAction : ������ ���������� ������ (0x" & Hex($iErrorCode) & ")"), 0)
  Return 1
EndFunc ; _UIA_ElementDoDefaultAction

; ============================================================================================
; ��� ������� : _UIA_ElementFindInArray
; ��������    : ������� �������, ��������������� ��������� �������� � ��� ��������
; ���������   : _UIA_ElementFindInArray(Const ByRef $aElements, $vProperty, $vPropertyValue[, $fInStr = False[, $iStartIndex = 1[, $fToEnd = True]]])
; ���������   : $aElements      - ������ ��������� (��������)
;             : $vProperty      - �������� �������� ��������
;             : $vPropertyValue - �������� �������� �������� ��������
;             : $fInStr         - ������ ���������� �������� �������� (�� ���������) ��� ���������
;             : $iStartIndex    - ������ ������ ������ (�� ��������� 1)
;             : $fToEnd         - ����������� ������: � ����� ������� (�� ���������) ��� � ������ �������
; ����������  : �����   - ������ �������� (�������) � �������� �������
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - ������ �������� �� �������� ��������
;             :         @error = 2 - ������ ������
;             :         @error = 3 - ����������� ������
;             :         @error = 4 - ������� ������� [0] �� ������������� ���������� ��������� ���������
;             :         @error = 5 - ������ �� �������� ���������
;             :         @error = 6 - ������� ������� [������] �� �������� ��������
;             :         @error = 7 - ������ �������������� ��������
;             :         @error = 8 - ������������ ������ ������ ������
; �����       : InnI
; ����������  : ������� ������� � ������� �������� ������ ��������� ���������� ��������
;             : �������� �������� ����� ����������� �� ����� ����� ������ ������� Inspect
;             : �������� �������� ����� ����������� �� ������ ����� ������ ������� Inspect
; ============================================================================================
Func _UIA_ElementFindInArray(Const ByRef $aElements, $vProperty, $vPropertyValue, $fInStr = False, $iStartIndex = 1, $fToEnd = True)
  If Not IsArray($aElements) Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_ElementFindInArray : ������ �������� �� �������� ��������"), 0)
  If Not UBound($aElements) Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_ElementFindInArray : ������ ������"), 0)
  If UBound($aElements, 0) > 1 Then Return SetError(3, __UIA_ConsoleWriteError("_UIA_ElementFindInArray : ����������� ������"), 0)
  If $aElements[0] <> UBound($aElements) - 1 Then Return SetError(4, __UIA_ConsoleWriteError("_UIA_ElementFindInArray : ������� ������� [0] �� ������������� ���������� ��������� ���������"), 0)
  If Not $aElements[0] Then Return SetError(5, __UIA_ConsoleWriteError("_UIA_ElementFindInArray : ������ �� �������� ���������"), 0)
  For $i = 1 To $aElements[0]
    If Not IsObj($aElements[$i]) Then Return SetError(6, __UIA_ConsoleWriteError("_UIA_ElementFindInArray : ������� ������� [" & $i & "] �� �������� ��������"), 0)
  Next
  $vProperty = __UIA_GetPropIdFromStr($vProperty, "_UIA_ElementFindInArray")
  If $vProperty = -1 Then Return SetError(7, __UIA_ConsoleWriteError("_UIA_ElementFindInArray : ������ �������������� ��������"), 0)
  $iStartIndex = Int($iStartIndex)
  If $iStartIndex < 1 Or $iStartIndex > $aElements[0] Then Return SetError(8, __UIA_ConsoleWriteError("_UIA_ElementFindInArray : ������������ ������ ������ ������"), 0)
  Local $vValue, $iEndIndex = $aElements[0], $iStep = 1
  If Not $fToEnd Then
    $iEndIndex = 1
    $iStep = -1
  EndIf
  For $i = $iStartIndex To $iEndIndex Step $iStep
    $aElements[$i].GetCurrentPropertyValue($vProperty, $vValue)
    If $fInStr Then
      If StringInStr($vValue, $vPropertyValue) Then Return $i
    Else
      If $vPropertyValue = $vValue Then Return $i
    EndIf
    $vValue = 0
  Next
  Return 0
EndFunc ; _UIA_ElementFindInArray

; ============================================================================================
; ��� ������� : _UIA_ElementGetBoundingRectangle
; ��������    : ���������� ������������� �������, �������������� �������
; ���������   : _UIA_ElementGetBoundingRectangle($oElement)
; ���������   : $oElement - ������� (������)
; ����������  : �����   - ������
;             :         $aArray[0] - X ���������� �������� ������ ���� ��������
;             :         $aArray[1] - Y ���������� �������� ������ ���� ��������
;             :         $aArray[2] - X ���������� ������� ������� ���� ��������
;             :         $aArray[3] - Y ���������� ������� ������� ���� ��������
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - �������� �� �������� ��������
;             :         @error = 2 - ������ �������� ������� ���������
; �����       : InnI
; ����������  : ���������� �������������� - ��������
; ============================================================================================
Func _UIA_ElementGetBoundingRectangle($oElement)
  If Not IsObj($oElement) Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_ElementGetBoundingRectangle : �������� �� �������� ��������"), 0)
  Local $tRect = DllStructCreate("long Left;long Top;long Right;long Bottom")
  $oElement.CurrentBoundingRectangle($tRect)
  Local $aRect[4] = [$tRect.Left, $tRect.Top, $tRect.Right, $tRect.Bottom]
  $tRect = 0
  If Not IsArray($aRect) Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_ElementGetBoundingRectangle : ������ �������� ������� ���������"), 0)
  Return $aRect
EndFunc ; _UIA_ElementGetBoundingRectangle

; ============================================================================================
; ��� ������� : _UIA_ElementGetFirstLastChild
; ��������    : ������� ������ � ��������� �������� �������� (�������) ���������� ��������
; ���������   : _UIA_ElementGetFirstLastChild($oElement)
; ���������   : $oElement - ������� (������)
; ����������  : �����   - ������
;             :         $aArray[0] - ������ �������� ������� (������)
;             :         $aArray[1] - ��������� �������� ������� (������)
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - �������� �� �������� ��������
;             :         @error = 2 - ������ �������� ������� UIAutomation
;             :         @error = 3 - ������ �������� ������� ������
; �����       : InnI
; ����������  : ���� ������/��������� ������� �����������, �� ��������������� ��������� ������� ����� �� ������
; ============================================================================================
Func _UIA_ElementGetFirstLastChild($oElement)
  If Not IsObj($oElement) Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_ElementGetFirstLastChild : �������� �� �������� ��������"), 0)
  Local $aFirstLast[2], $pFirstLast, $pRawWalker, $oRawWalker, $oUIAutomation = _UIA_ObjectCreate()
  If @error Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_ElementGetFirstLastChild : ������ �������� ������� UIAutomation"), 0)
  $oUIAutomation.RawViewWalker($pRawWalker)
  $oRawWalker = ObjCreateInterface($pRawWalker, $sIID_IUIAutomationTreeWalker, $dtagIUIAutomationTreeWalker)
  If Not IsObj($oRawWalker) Then Return SetError(3, __UIA_ConsoleWriteError("_UIA_ElementGetFirstLastChild : ������ �������� ������� ������"), 0)
  $oRawWalker.GetFirstChildElement($oElement, $pFirstLast)
  $aFirstLast[0] = __UIA_CreateElement($pFirstLast)
  $pFirstLast = 0
  $oRawWalker.GetLastChildElement($oElement, $pFirstLast)
  $aFirstLast[1] = __UIA_CreateElement($pFirstLast)
  $pFirstLast = 0
  Return $aFirstLast
EndFunc ; _UIA_ElementGetFirstLastChild

; ============================================================================================
; ��� ������� : _UIA_ElementGetParent
; ��������    : ���������� ������������ ������� (������) ���������� ��������
; ���������   : _UIA_ElementGetParent($oElement)
; ���������   : $oElement - ������� (������)
; ����������  : �����   - ������������ ������� (������)
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - �������� �� �������� ��������
;             :         @error = 2 - ������ �������� ������� UIAutomation
;             :         @error = 3 - ������ �������� ������� ������
; �����       : InnI
; ����������  : ���� ������������ ������� �����������, �� ������� ����� �� ������
; ============================================================================================
Func _UIA_ElementGetParent($oElement)
  If Not IsObj($oElement) Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_ElementGetParent : �������� �� �������� ��������"), 0)
  Local $pParent, $pRawWalker, $oRawWalker, $oUIAutomation = _UIA_ObjectCreate()
  If @error Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_ElementGetParent : ������ �������� ������� UIAutomation"), 0)
  $oUIAutomation.RawViewWalker($pRawWalker)
  $oRawWalker = ObjCreateInterface($pRawWalker, $sIID_IUIAutomationTreeWalker, $dtagIUIAutomationTreeWalker)
  If Not IsObj($oRawWalker) Then Return SetError(3, __UIA_ConsoleWriteError("_UIA_ElementGetParent : ������ �������� ������� ������"), 0)
  $oRawWalker.GetParentElement($oElement, $pParent)
  Return __UIA_CreateElement($pParent)
EndFunc ; _UIA_ElementGetParent

; ============================================================================================
; ��� ������� : _UIA_ElementGetPreviousNext
; ��������    : ������� ���������� � ��������� �������� (�������) ���� �� ������
; ���������   : _UIA_ElementGetPreviousNext($oElement)
; ���������   : $oElement - ������� (������)
; ����������  : �����   - ������
;             :         $aArray[0] - ���������� ������� (������)
;             :         $aArray[1] - ��������� ������� (������)
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - �������� �� �������� ��������
;             :         @error = 2 - ������ �������� ������� UIAutomation
;             :         @error = 3 - ������ �������� ������� ������
; �����       : InnI
; ����������  : ���� ����������/��������� ������� �����������, �� ��������������� ��������� ������� ����� �� ������
;             : �� ������� ������������ �������� � �������� ����������� � ������ �������� � �������� ����������
; ============================================================================================
Func _UIA_ElementGetPreviousNext($oElement)
  If Not IsObj($oElement) Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_ElementGetPreviousNext : �������� �� �������� ��������"), 0)
  Local $aPrevNext[2], $pPrevNext, $pRawWalker, $oRawWalker, $oUIAutomation = _UIA_ObjectCreate()
  If @error Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_ElementGetPreviousNext : ������ �������� ������� UIAutomation"), 0)
  $oUIAutomation.RawViewWalker($pRawWalker)
  $oRawWalker = ObjCreateInterface($pRawWalker, $sIID_IUIAutomationTreeWalker, $dtagIUIAutomationTreeWalker)
  If Not IsObj($oRawWalker) Then Return SetError(3, __UIA_ConsoleWriteError("_UIA_ElementGetPreviousNext : ������ �������� ������� ������"), 0)
  $oRawWalker.GetPreviousSiblingElement($oElement, $pPrevNext)
  $aPrevNext[0] = __UIA_CreateElement($pPrevNext)
  $pPrevNext = 0
  $oRawWalker.GetNextSiblingElement($oElement, $pPrevNext)
  $aPrevNext[1] = __UIA_CreateElement($pPrevNext)
  $pPrevNext = 0
  Return $aPrevNext
EndFunc ; _UIA_ElementGetPreviousNext

; ============================================================================================
; ��� ������� : _UIA_ElementGetPropertyValue
; ��������    : ���������� �������� ��������� �������� ��������
; ���������   : _UIA_ElementGetPropertyValue($oElement, $vProperty)
; ���������   : $oElement  - ������� (������)
;             : $vProperty - ��������, �������� �������� ����� ��������
; ����������  : �����   - �������� ��������� ��������
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - ������ �������� �� �������� ��������
;             :         @error = 2 - ������ �������������� ��������
;             :         @error = 3 - ������ ���������� ������
; �����       : InnI
; ����������  : �������� �������� ����� ����������� �� ����� ����� ������ ������� Inspect
; ============================================================================================
Func _UIA_ElementGetPropertyValue($oElement, $vProperty)
  If Not IsObj($oElement) Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_ElementGetPropertyValue : ������ �������� �� �������� ��������"), 0)
  $vProperty = __UIA_GetPropIdFromStr($vProperty, "_UIA_ElementGetPropertyValue")
  If $vProperty = -1 Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_ElementGetPropertyValue : ������ �������������� ��������"), 0)
  Local $vValue, $iErrorCode = $oElement.GetCurrentPropertyValue($vProperty, $vValue)
  If $iErrorCode Then Return SetError(3, __UIA_ConsoleWriteError("_UIA_ElementGetPropertyValue : ������ ���������� ������ (0x" & Hex($iErrorCode) & ")"), 0)
  Return $vValue
EndFunc ; _UIA_ElementGetPropertyValue

; ============================================================================================
; ��� ������� : _UIA_ElementMouseClick
; ��������    : ��������� ���� ���� �� ��������
; ���������   : _UIA_ElementMouseClick($oElement[, $sButton = ""[, $iX = Default[, $iY = Default[, $iClicks = 1[, $fSetFocus = True]]]]])
; ���������   : $oElement - ������� (������)
;             : $sButton  - ������ ���� (�� ��������� �����)
;             : $iX       - X ���������� ����� (�� ��������� �������� ������ ��������)
;             : $iY       - Y ���������� ����� (�� ��������� �������� ������ ��������)
;             : $iClicks  - ���������� ������ (�� ��������� 1)
;             : $SetFocus - ��������� �������� ����� ������ (�� ��������� True)
; ����������  : �����   - 1
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - ������ �������� �� �������� ��������
;             :         @error = 2 - ������ �������� ������� ��������� ������������� �������
;             :         @error = 3 - ���������� ����� ������� �� ������� ��������
;             :         @error = 4 - ������ ���������� ������� MouseClick
; �����       : InnI
; ����������  : ������������ �������� ���������� ������������ ������ �������� ���� ��������
; ============================================================================================
Func _UIA_ElementMouseClick($oElement, $sButton = "", $iX = Default, $iY = Default, $iClicks = 1, $fSetFocus = True)
  If Not IsObj($oElement) Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_ElementMouseClick : ������ �������� �� �������� ��������"), 0)
  If $fSetFocus Then $oElement.SetFocus()
  Local $aRect = _UIA_ElementGetBoundingRectangle($oElement)
  If Not IsArray($aRect) Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_ElementMouseClick : ������ �������� ������� ��������� ������������� �������"), 0)
  If $iX = Default Then $iX = ($aRect[2] - $aRect[0]) / 2
  If $iY = Default Then $iY = ($aRect[3] - $aRect[1]) / 2
  If $iX < 0 Or $iY < 0 Or $iX >= $aRect[2] - $aRect[0] Or $iY >= $aRect[3] - $aRect[1] Then  Return SetError(3, __UIA_ConsoleWriteError("_UIA_ElementMouseClick : ���������� ����� ������� �� ������� ��������"), 0)
  Local $iMode = Opt("MouseCoordMode", 1)
  Local $iResult = MouseClick($sButton, $aRect[0] + $iX, $aRect[1] + $iY, $iClicks, 0)
  Opt("MouseCoordMode", $iMode)
  If Not $iResult Then Return SetError(4, __UIA_ConsoleWriteError("_UIA_ElementMouseClick : ������ ���������� ������� MouseClick"), 0)
  Return 1
EndFunc ; _UIA_ElementMouseClick

; ============================================================================================
; ��� ������� : _UIA_ElementScrollIntoView
; ��������    : ������������ ������� � ������� ���������
; ���������   : _UIA_ElementScrollIntoView($oElement)
; ���������   : $oElement - ������� (������)
; ����������  : �����   - 1
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - �������� �� �������� ��������
;             :         @error = 2 - ������ �������� ������� �� ������ ������� ScrollItem
;             :         @error = 3 - ������ ���������� ������
; �����       : InnI
; ����������  :
; ============================================================================================
Func _UIA_ElementScrollIntoView($oElement)
  If Not IsObj($oElement) Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_ElementScrollIntoView : �������� �� �������� ��������"), 0)
  Local $pScrollItem, $oScrollItem
  $oElement.GetCurrentPattern($UIA_ScrollItemPatternId, $pScrollItem)
  $oScrollItem = ObjCreateInterface($pScrollItem, $sIID_IUIAutomationScrollItemPattern, $dtagIUIAutomationScrollItemPattern)
  If Not IsObj($oScrollItem) Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_ElementScrollIntoView : ������ �������� ������� �� ������ ������� ScrollItem"), 0)
  Local $iErrorCode = $oScrollItem.ScrollIntoView()
  If $iErrorCode Then Return SetError(3, __UIA_ConsoleWriteError("_UIA_ElementScrollIntoView : ������ ���������� ������ (0x" & Hex($iErrorCode) & ")"), 0)
  Return 1
EndFunc ; _UIA_ElementScrollIntoView

; ============================================================================================
; ��� ������� : _UIA_ElementSetFocus
; ��������    : ������������� �������� ����� �����
; ���������   : _UIA_ElementSetFocus($oElement)
; ���������   : $oElement - ������� (������)
; ����������  : �����   - 1
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - �������� �� �������� ��������
;             :         @error = 2 - ������ ���������� ������
; �����       : InnI
; ����������  : ����� ������������ ����, ���������� �������
; ============================================================================================
Func _UIA_ElementSetFocus($oElement)
  If Not IsObj($oElement) Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_ElementSetFocus : �������� �� �������� ��������"), 0)
  Local $iErrorCode = $oElement.SetFocus()
  If $iErrorCode Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_ElementSetFocus : ������ ���������� ������ (0x" & Hex($iErrorCode) & ")"), 0)
  Return 1
EndFunc ; _UIA_ElementSetFocus

; ============================================================================================
; ��� ������� : _UIA_ElementTextSetValue
; ��������    : ������������� �������� (�����) � ��������� �������
; ���������   : _UIA_ElementTextSetValue($oElement, $sValue)
; ���������   : $oElement - ������� (������)
;             : $sValue   - ����� �������� (�����)
; ����������  : �����   - 1
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - ������ �������� �� �������� ��������
;             :         @error = 2 - ������ �������� ������� �� ������ ������� LegacyIAccessible
;             :         @error = 3 - ������ ���������� ������
; �����       : InnI
; ����������  : ������ ��� ��������� ���������, ������� ������ ValuePattern
; ============================================================================================
Func _UIA_ElementTextSetValue($oElement, $sValue)
  If Not IsObj($oElement) Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_ElementTextSetValue : ������ �������� �� �������� ��������"), 0)
  Local $pIAccess, $oIAccess
  $oElement.GetCurrentPattern($UIA_LegacyIAccessiblePatternId, $pIAccess)
  $oIAccess = ObjCreateInterface($pIAccess, $sIID_IUIAutomationLegacyIAccessiblePattern, $dtagIUIAutomationLegacyIAccessiblePattern)
  If Not IsObj($oIAccess) Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_ElementTextSetValue : ������ �������� ������� �� ������ ������� LegacyIAccessible"), 0)
  Local $iErrorCode = $oIAccess.SetValue($sValue)
  If $iErrorCode Then Return SetError(3, __UIA_ConsoleWriteError("_UIA_ElementTextSetValue : ������ ���������� ������ (0x" & Hex($iErrorCode) & ")"), 0)
  Return 1
EndFunc ; _UIA_ElementTextSetValue

; ============================================================================================
; ��� ������� : _UIA_FindAllElements
; ��������    : ������� ��� ��������, ��������������� ��������� �������� � ��� ��������
; ���������   : _UIA_FindAllElements($oElementFrom[, $vProperty = 0[, $vPropertyValue = ""]])
; ���������   : $oElementFrom   - ������� (������), �� �������� ���������� �����
;             : $vProperty      - �������� �������� �������� (�� ��������� 0 - �����)
;             : $vPropertyValue - �������� �������� �������� �������� (�� ��������� ������ ������)
; ����������  : �����   - ������ ��������� (��������)
;             :         $aArray[0] - ���������� ��������� ��������� (N)
;             :         $aArray[1] - ������ ��������� ������� (������)
;             :         $aArray[2] - ������ ��������� ������� (������)
;             :         $aArray[N] - N-� ��������� ������� (������)
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - ������ �������� �� �������� ��������
;             :         @error = 2 - ������ �������� ������� UIAutomation
;             :         @error = 3 - ������ �������������� ��������
;             :         @error = 4 - ������ �������� ������� ������ (��. ���������� � ���������)
;             :         @error = 5 - ������ �������� ������� ��������� (��������)
; �����       : InnI
; ����������  : ����� ������������ �� ���������� �������� �� ��� ������ � ���� �����������
;             : ��� �������, �� �������� ������������ �����, � ��������� �� ����������
;             : ��� ������ ����� ��������� ($vProperty = 0) �������� �������� �� �����������
;             : �������� �������� ����� ����������� �� ����� ����� ������ ������� Inspect
;             : �������� �������� ����� ����������� �� ������ ����� ������ ������� Inspect
;             : �������� ������� ������������. ��������, ��� �������� "IsEnabled" ����� ��������� �������� True, � �� 1 � �� "True"
; ============================================================================================
Func _UIA_FindAllElements($oElementFrom, $vProperty = 0, $vPropertyValue = "")
  If Not IsObj($oElementFrom) Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_FindAllElements : ������ �������� �� �������� ��������"), 0)
  Local $oCondition, $oUIAutomation = _UIA_ObjectCreate()
  If @error Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_FindAllElements : ������ �������� ������� UIAutomation"), 0)
  If Not $vProperty Then
    $oCondition = Default
  Else
    $vProperty = __UIA_GetPropIdFromStr($vProperty, "_UIA_FindAllElements")
    If $vProperty = -1 Then Return SetError(3, __UIA_ConsoleWriteError("_UIA_FindAllElements : ������ �������������� ��������"), 0)
    $oCondition = _UIA_CreatePropertyCondition($vProperty, $vPropertyValue)
    If Not IsObj($oCondition) Then Return SetError(4, __UIA_ConsoleWriteError("_UIA_FindAllElements : ������ �������� ������� ������ (��. ���������� � ���������)"), 0)
  EndIf
  Local $aElements = _UIA_FindAllElementsEx($oElementFrom, $oCondition)
  If Not IsArray($aElements) Then Return SetError(5, __UIA_ConsoleWriteError("_UIA_FindAllElements : ������ �������� ������� ��������� (��������)"), 0)
  Return $aElements
EndFunc ; _UIA_FindAllElements

; ============================================================================================
; ��� ������� : _UIA_FindAllElementsEx
; ��������    : ������� ��� ��������, ��������������� ������� ������
; ���������   : _UIA_FindAllElementsEx($oElementFrom[, $oCondition = Default[, $iTreeScope = Default]])
; ���������   : $oElementFrom - ������� (������), �� �������� ���������� �����
;             : $oCondition   - ������� ������ (�� ��������� �����)
;             : $iTreeScope   - ������� ������:
;             :               $TreeScope_Element - �������� ������ ��� �������
;             :               $TreeScope_Children - �������� ������ �������� ��������
;             :               $TreeScope_Descendants - �������� ���� �������� �������� (�� ���������)
;             :               $TreeScope_Subtree - �������� ��� ������� � ���� ��������
; ����������  : �����   - ������ ��������� (��������)
;             :         $aArray[0] - ���������� ��������� ��������� (N)
;             :         $aArray[1] - ������ ��������� ������� (������)
;             :         $aArray[2] - ������ ��������� ������� (������)
;             :         $aArray[N] - N-� ��������� ������� (������)
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - ������ �������� �� �������� ��������
;             :         @error = 2 - ������� ������ �� ����������
;             :         @error = 3 - ������ �������� ������� �������
;             :         @error = 4 - ������ �������� ������� ��������� (��������)
; �����       : InnI
; ����������  : ������� ������ $TreeScope_Parent � $TreeScope_Ancestors �� ��������������
; ============================================================================================
Func _UIA_FindAllElementsEx($oElementFrom, $oCondition = Default, $iTreeScope = Default)
  If Not IsObj($oElementFrom) Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_FindAllElementsEx : ������ �������� �� �������� ��������"), 0)
  If $oCondition = Default Then
    Local $pCondition, $oUIAutomation = _UIA_ObjectCreate()
    $oUIAutomation.CreateTrueCondition($pCondition)
    $oCondition = ObjCreateInterface($pCondition, $sIID_IUIAutomationBoolCondition, $dtagIUIAutomationBoolCondition)
  EndIf
  If Not IsObj($oCondition) Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_FindAllElementsEx : ������� ������ �� ����������"), 0)
  If $iTreeScope = Default Then $iTreeScope = $TreeScope_Descendants
  Local $pUIElementArray, $oUIElementArray, $pUIElement, $iElements
  $oElementFrom.FindAll($iTreeScope, $oCondition, $pUIElementArray)
  $oUIElementArray = ObjCreateInterface($pUIElementArray, $sIID_IUIAutomationElementArray, $dtagIUIAutomationElementArray)
  If Not IsObj($oUIElementArray) Then Return SetError(3, __UIA_ConsoleWriteError("_UIA_FindAllElementsEx : ������ �������� ������� �������"), 0)
  $oUIElementArray.Length($iElements)
  If Not $iElements Then
    Local $aElements[1] = [0]
    Return $aElements
  EndIf
  Local $aElements[$iElements + 1]
  $aElements[0] = $iElements
  For $i = 1 To $iElements
    $oUIElementArray.GetElement($i - 1, $pUIElement)
    $aElements[$i] = __UIA_CreateElement($pUIElement)
    $pUIElement = 0
  Next
  If Not IsArray($aElements) Then Return SetError(4, __UIA_ConsoleWriteError("_UIA_FindAllElementsEx : ������ �������� ������� ��������� (��������)"), 0)
  Return $aElements
EndFunc ; _UIA_FindAllElementsEx

; ============================================================================================
; ��� ������� : _UIA_FindElementsInArray
; ��������    : ������� ��� ��������, ��������������� ��������� �������� � ��� ��������
; ���������   : _UIA_FindElementsInArray(Const ByRef $aElements, $vProperty, $vPropertyValue[, $fInStr = False[, $fIndexArray = False]])
; ���������   : $aElements      - ������ ��������� (��������)
;             : $vProperty      - �������� �������� ��������
;             : $vPropertyValue - �������� �������� �������� ��������
;             : $fInStr         - ������ ���������� �������� �������� (�� ���������) ��� ���������
;             : $fIndexArray    - ���������� ������� �������� ������� (�� ���������) ��� ������� ��������� ��������� �������
; ����������  : �����   - ������ ��������� (��������)
;             :         $aArray[0] - ���������� ��������� ��������� (N)
;             :         $aArray[1] - ������ ��������� ������� (������)
;             :         $aArray[2] - ������ ��������� ������� (������)
;             :         $aArray[N] - N-� ��������� ������� (������)
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - ������ �������� �� �������� ��������
;             :         @error = 2 - ������ ������
;             :         @error = 3 - ����������� ������
;             :         @error = 4 - ������� ������� [0] �� ������������� ���������� ��������� ���������
;             :         @error = 5 - ������ �� �������� ���������
;             :         @error = 6 - ������� ������� [������] �� �������� ��������
;             :         @error = 7 - ������ �������������� ��������
; �����       : InnI
; ����������  : ������� ������� � ������� �������� ������ ��������� ���������� ��������
;             : �������� �������� ����� ����������� �� ����� ����� ������ ������� Inspect
;             : �������� �������� ����� ����������� �� ������ ����� ������ ������� Inspect
; ============================================================================================
Func _UIA_FindElementsInArray(Const ByRef $aElements, $vProperty, $vPropertyValue, $fInStr = False, $fIndexArray = False)
  If Not IsArray($aElements) Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_FindElementsInArray : ������ �������� �� �������� ��������"), 0)
  If Not UBound($aElements) Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_FindElementsInArray : ������ ������"), 0)
  If UBound($aElements, 0) > 1 Then Return SetError(3, __UIA_ConsoleWriteError("_UIA_FindElementsInArray : ����������� ������"), 0)
  If $aElements[0] <> UBound($aElements) - 1 Then Return SetError(4, __UIA_ConsoleWriteError("_UIA_FindElementsInArray : ������� ������� [0] �� ������������� ���������� ��������� ���������"), 0)
  If Not $aElements[0] Then Return SetError(5, __UIA_ConsoleWriteError("_UIA_FindElementsInArray : ������ �� �������� ���������"), 0)
  For $i = 1 To $aElements[0]
    If Not IsObj($aElements[$i]) Then Return SetError(6, __UIA_ConsoleWriteError("_UIA_FindElementsInArray : ������� ������� [" & $i & "] �� �������� ��������"), 0)
  Next
  $vProperty = __UIA_GetPropIdFromStr($vProperty, "_UIA_FindElementsInArray")
  If $vProperty = -1 Then Return SetError(7, __UIA_ConsoleWriteError("_UIA_FindElementsInArray : ������ �������������� ��������"), 0)
  Local $vValue, $aArray[$aElements[0] + 1], $j = 1
  For $i = 1 To $aElements[0]
    $aElements[$i].GetCurrentPropertyValue($vProperty, $vValue)
    If $fInStr Then
      If StringInStr($vValue, $vPropertyValue) Then
        If $fIndexArray Then
          $aArray[$j] = $i
        Else
          $aArray[$j] = $aElements[$i]
        EndIf
        $j += 1
      EndIf
    Else
      If $vPropertyValue = $vValue Then
        If $fIndexArray Then
          $aArray[$j] = $i
        Else
          $aArray[$j] = $aElements[$i]
        EndIf
        $j += 1
      EndIf
    EndIf
    $vValue = 0
  Next
  ReDim $aArray[$j]
  $aArray[0] = $j - 1
  Return $aArray
EndFunc ; _UIA_FindElementsInArray

; ============================================================================================
; ��� ������� : _UIA_GetControlTypeElement
; ��������    : ������� ������� (������) ���������� ���� � �������� ��������� � ���������
; ���������   : _UIA_GetControlTypeElement($oElementFrom, $vControlType, $vPropertyValue[, $vProperty = Default[, $fInStr = False]])
; ���������   : $oElementFrom   - ������� (������), �� �������� ���������� �����
;             : $vControlType   - ������������� ���� �������� ��������
;             : $vPropertyValue - �������� �������� �������� ��������
;             : $vProperty      - �������� �������� �������� (�� ��������� "Name" - $UIA_NamePropertyId)
;             : $fInStr         - ������ ���������� �������� �������� (�� ���������) ��� ���������
; ����������  : �����   - ������� (������)
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - ������ �������� �� �������� ��������
;             :         @error = 2 - ������ �������������� ����
;             :         @error = 3 - ������ �������������� ��������
;             :         @error = 4 - ������ �������� ������� ��������� (��������)
;             :         @error = 5 - �������� ���������� ���� �� �������
;             :         @error = 6 - �������� ���������� �������� ��������� ��������� �� ������������� ���������
; �����       : InnI
; ����������  : ����� ������������ �� ���������� �������� �� ��� ������ � ���� �����������
;             : ������������� ���� ����� ����������� �� �������� �������� "ControlType" ������� Inspect
;             : �������� �������� ����� ����������� �� ����� ����� ������ ������� Inspect
;             : �������� �������� ����� ����������� �� ������ ����� ������ ������� Inspect
; ============================================================================================
Func _UIA_GetControlTypeElement($oElementFrom, $vControlType, $vPropertyValue, $vProperty = Default, $fInStr = False)
  If Not IsObj($oElementFrom) Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_GetControlTypeElement : ������ �������� �� �������� ��������"), 0)
  $vControlType = __UIA_GetTypeIdFromStr($vControlType, "_UIA_GetControlTypeElement")
  If $vControlType = -1 Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_GetControlTypeElement : ������ �������������� ����"), 0)
  If $vProperty = Default Then
    $vProperty = $UIA_NamePropertyId
  Else
    $vProperty = __UIA_GetPropIdFromStr($vProperty, "_UIA_GetControlTypeElement")
    If $vProperty = -1 Then Return SetError(3, __UIA_ConsoleWriteError("_UIA_GetControlTypeElement : ������ �������������� ��������"), 0)
  EndIf
  Local $aElements = _UIA_FindAllElements($oElementFrom, $UIA_ControlTypePropertyId, $vControlType)
  If Not IsArray($aElements) Then Return SetError(4, __UIA_ConsoleWriteError("_UIA_GetControlTypeElement : ������ �������� ������� ��������� (��������)"), 0)
  If Not $aElements[0] Then Return SetError(5, __UIA_ConsoleWriteError("_UIA_GetControlTypeElement : �������� ���������� ���� �� �������"), 0)
  Local $vValue
  For $i = 1 To $aElements[0]
    $aElements[$i].GetCurrentPropertyValue($vProperty, $vValue)
    If $fInStr Then
      If StringInStr($vValue, $vPropertyValue) Then ExitLoop
    Else
      If $vPropertyValue = $vValue Then ExitLoop
    EndIf
    $vValue = 0
  Next
  If $i = $aElements[0] + 1 Then Return SetError(6, __UIA_ConsoleWriteError("_UIA_GetControlTypeElement : �������� ���������� �������� ��������� ��������� �� ������������� ���������"), 0)
  Return $aElements[$i]
EndFunc ; _UIA_GetControlTypeElement

; ============================================================================================
; ��� ������� : _UIA_GetElementFromCondition
; ��������    : ������� ������� (������) �� ������ ��������� �������
; ���������   : _UIA_GetElementFromCondition($oElementFrom, $oCondition[, $iTreeScope = Default])
; ���������   : $oElementFrom - ������� (������), �� �������� ���������� �����
;             : $oCondition   - ������� ������
;             : $iTreeScope   - ������� ������:
;             :               $TreeScope_Element - �������� ������ ��� �������
;             :               $TreeScope_Children - �������� ������ �������� ��������
;             :               $TreeScope_Descendants - �������� ���� �������� �������� (�� ���������)
;             :               $TreeScope_Subtree - �������� ��� ������� � ���� ��������
; ����������  : �����   - ������� (������)
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - ������ �������� �� �������� ��������
;             :         @error = 2 - ������� ������ �� ����������
;             :         @error = 3 - ������ ���������� ������
; �����       : InnI
; ����������  : ���� ������� �� ������, �� ������� ����� �� ������
;             : ������� ������ $TreeScope_Parent � $TreeScope_Ancestors �� ��������������
; ============================================================================================
Func _UIA_GetElementFromCondition($oElementFrom, $oCondition, $iTreeScope = Default)
  If Not IsObj($oElementFrom) Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_GetElementFromCondition : ������ �������� �� �������� ��������"), 0)
  If Not IsObj($oCondition) Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_GetElementFromCondition : ������� ������ �� ����������"), 0)
  If $iTreeScope = Default Then $iTreeScope = $TreeScope_Descendants
  Local $pUIElement, $iErrorCode = $oElementFrom.FindFirst($iTreeScope, $oCondition, $pUIElement)
  If $iErrorCode Then Return SetError(3, __UIA_ConsoleWriteError("_UIA_GetElementFromCondition : ������ ���������� ������ (0x" & Hex($iErrorCode) & ")"), 0)
  Return __UIA_CreateElement($pUIElement)
EndFunc ; _UIA_GetElementFromCondition

; ============================================================================================
; ��� ������� : _UIA_GetElementFromFocus
; ��������    : ������ ������� (������) �� ������ ������ �����
; ���������   : _UIA_GetElementFromFocus()
; ���������   :
; ����������  : �����   - ������� (������)
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - ������ �������� ������� UIAutomation
;             :         @error = 2 - ������ �������� �������� (�������)
; �����       : InnI
; ����������  :
; ============================================================================================
Func _UIA_GetElementFromFocus()
  Local $pElement, $oElement, $oUIAutomation = _UIA_ObjectCreate()
  If @error Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_GetElementFromFocus : ������ �������� ������� UIAutomation"), 0)
  $oUIAutomation.GetFocusedElement($pElement)
  $oElement = __UIA_CreateElement($pElement)
  $pElement = 0
  If Not IsObj($oElement) Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_GetElementFromFocus : ������ �������� �������� (�������)"), 0)
  Return $oElement
EndFunc ; _UIA_GetElementFromFocus

; ============================================================================================
; ��� ������� : _UIA_GetElementFromHandle
; ��������    : ������ ������� (������) �� ������ �����������
; ���������   : _UIA_GetElementFromHandle($hHandle)
; ���������   : $hHandle - ���������� ���� ��� ��������
; ����������  : �����   - ������� (������)
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - �������� �� �������� ������������
;             :         @error = 2 - ������ �������� ������� UIAutomation
;             :         @error = 3 - ������ �������� �������� (�������)
; �����       : InnI
; ����������  :
; ============================================================================================
Func _UIA_GetElementFromHandle($hHandle)
  If Not IsHWnd($hHandle) Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_GetElementFromHandle : �������� �� �������� ������������"), 0)
  Local $pElement, $oElement, $oUIAutomation = _UIA_ObjectCreate()
  If @error Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_GetElementFromHandle : ������ �������� ������� UIAutomation"), 0)
  $oUIAutomation.ElementFromHandle($hHandle, $pElement)
  $oElement = __UIA_CreateElement($pElement)
  $pElement = 0
  If Not IsObj($oElement) Then Return SetError(3, __UIA_ConsoleWriteError("_UIA_GetElementFromHandle : ������ �������� �������� (�������)"), 0)
  Return $oElement
EndFunc ; _UIA_GetElementFromHandle

; ============================================================================================
; ��� ������� : _UIA_GetElementFromPoint
; ��������    : ������ ������� (������) �� ������ �������� ���������
; ���������   : _UIA_GetElementFromPoint($iX = Default, $iY = Default)
; ���������   : $iX - X ���������� ������ (�� ��������� - ������� ����)
;             : $iY - Y ���������� ������ (�� ��������� - ������� ����)
; ����������  : �����   - ������� (������)
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - ������ �������� ������� UIAutomation
;             :         @error = 2 - ������ �������� �������� (�������)
; �����       : InnI
; ����������  :
; ============================================================================================
Func _UIA_GetElementFromPoint($iX = Default, $iY = Default)
  Local $pElement, $oElement, $tPoint = DllStructCreate("long X;long Y"), $oUIAutomation = _UIA_ObjectCreate()
  If @error Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_GetElementFromPoint : ������ �������� ������� UIAutomation"), 0)
  Local $iMode = Opt("MouseCoordMode", 1)
  $tPoint.X = ($iX = Default) ? MouseGetPos(0) : $iX
  $tPoint.Y = ($iY = Default) ? MouseGetPos(1) : $iY
  Opt("MouseCoordMode", $iMode)
  $oUIAutomation.ElementFromPoint($tPoint, $pElement)
  $oElement = __UIA_CreateElement($pElement)
  $pElement = 0
  If Not IsObj($oElement) Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_GetElementFromPoint : ������ �������� �������� (�������)"), 0)
  Return $oElement
EndFunc ; _UIA_GetElementFromPoint

; ============================================================================================
; ��� ������� : _UIA_ObjectCreate
; ��������    : ������ ������ UIAutomation
; ���������   : _UIA_ObjectCreate()
; ���������   :
; ����������  : �����   - ������ UIAutomation
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - ������ �������� ������� UIAutomation
;             :         @error = 2 - ����������� ������ ��������
; �����       : InnI
; ����������  :
; ============================================================================================
Func _UIA_ObjectCreate()
  Local $oUIAutomation
  Switch $UIA_ElementVersion
    Case 0
      For $i = 4 To 2 Step -1
        $oUIAutomation = ObjCreateInterface($sCLSID_CUIAutomation8, Eval("sIID_IUIAutomation" & $i), Eval("dtagIUIAutomation" & $i))
        If IsObj($oUIAutomation) Then Return $oUIAutomation
      Next
      ContinueCase
    Case 1
      $oUIAutomation = ObjCreateInterface($sCLSID_CUIAutomation, $sIID_IUIAutomation, $dtagIUIAutomation)
    Case 2
      $oUIAutomation = ObjCreateInterface($sCLSID_CUIAutomation8, $sIID_IUIAutomation2, $dtagIUIAutomation2)
    Case 3 To 5
      $oUIAutomation = ObjCreateInterface($sCLSID_CUIAutomation8, $sIID_IUIAutomation3, $dtagIUIAutomation3)
    Case 6 To 7
      $oUIAutomation = ObjCreateInterface($sCLSID_CUIAutomation8, $sIID_IUIAutomation4, $dtagIUIAutomation4)
    Case Else
      Return SetError(2, __UIA_ConsoleWriteError("_UIA_ObjectCreate : ����������� ������ �������� (" & $UIA_ElementVersion & ")"), 0)
  EndSwitch
  If Not IsObj($oUIAutomation) Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_ObjectCreate : ������ �������� ������� UIAutomation"), 0)
  Return $oUIAutomation
EndFunc ; _UIA_ObjectCreate

; ============================================================================================
; ��� ������� : _UIA_WaitControlTypeElement
; ��������    : ������� ������� (������) ���������� ���� � �������� ��������� � ���������
; ���������   : _UIA_WaitControlTypeElement($oElementFrom, $vControlType, $vPropertyValue[, $vProperty = Default[, $fInStr = False[, $iWaitTime = Default]]])
; ���������   : $oElementFrom   - ������� (������), �� �������� ���������� �����
;             : $vControlType   - ������������� ���� ���������� ��������
;             : $vPropertyValue - �������� �������� ���������� ��������
;             : $vProperty      - �������� ���������� �������� (�� ��������� "Name" - $UIA_NamePropertyId)
;             : $fInStr         - ������ ���������� �������� �������� (�� ���������) ��� ���������
;             : $iWaitTime      - ����� �������� � �������� (�� ��������� $UIA_DefaultWaitTime, 0 - ����������)
; ����������  : �����   - ������� (������)
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - ������ �������� �� �������� ��������
;             :         @error = 2 - ������ �������������� ����
;             :         @error = 3 - ������ �������������� ��������
;             :         @error = 4 - ��������� ����� ��������
; �����       : InnI
; ����������  : ����� ������������ �� ���������� �������� �� ��� ������ � ���� �����������
;             : ������������� ���� ����� ����������� �� �������� �������� "ControlType" ������� Inspect
;             : �������� �������� ����� ����������� �� ����� ����� ������ ������� Inspect
;             : �������� �������� ����� ����������� �� ������ ����� ������ ������� Inspect
; ============================================================================================
Func _UIA_WaitControlTypeElement($oElementFrom, $vControlType, $vPropertyValue, $vProperty = Default, $fInStr = False, $iWaitTime = Default)
  If Not IsObj($oElementFrom) Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_WaitControlTypeElement : ������ �������� �� �������� ��������"), 0)
  $vControlType = __UIA_GetTypeIdFromStr($vControlType, "_UIA_WaitControlTypeElement")
  If $vControlType = -1 Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_WaitControlTypeElement : ������ �������������� ����"), 0)
  If $vProperty = Default Then
    $vProperty = $UIA_NamePropertyId
  Else
    $vProperty = __UIA_GetPropIdFromStr($vProperty, "_UIA_WaitControlTypeElement")
    If $vProperty = -1 Then Return SetError(3, __UIA_ConsoleWriteError("_UIA_WaitControlTypeElement : ������ �������������� ��������"), 0)
  EndIf
  If $iWaitTime = Default Then $iWaitTime = $UIA_DefaultWaitTime
  Local $oElement, $iOption = $UIA_ConsoleWriteError, $iTimer = TimerInit()
  $UIA_ConsoleWriteError = 0
  Do
    $oElement = _UIA_GetControlTypeElement($oElementFrom, $vControlType, $vPropertyValue, $vProperty, $fInStr)
    If IsObj($oElement) Then
      $UIA_ConsoleWriteError = $iOption
      Return $oElement
    EndIf
    If $iWaitTime > 0 And TimerDiff($iTimer) > $iWaitTime * 1000 Then
      $UIA_ConsoleWriteError = $iOption
      Return SetError(4, __UIA_ConsoleWriteError("_UIA_WaitControlTypeElement : ��������� ����� ��������"), 0)
    EndIf
    Sleep(100)
  Until 0
EndFunc ; _UIA_WaitControlTypeElement

; ============================================================================================
; ��� ������� : _UIA_WaitElementFromCondition
; ��������    : ������� ������� (������) �� ������ ��������� �������
; ���������   : _UIA_WaitElementFromCondition($oElementFrom, $oCondition[, $iTreeScope = Default[, $iWaitTime = Default]])
; ���������   : $oElementFrom - ������� (������), �� �������� ���������� �����
;             : $oCondition   - ������� ������
;             : $iTreeScope   - ������� ������:
;             :               $TreeScope_Element - �������� ������ ��� �������
;             :               $TreeScope_Children - �������� ������ �������� ��������
;             :               $TreeScope_Descendants - �������� ���� �������� �������� (�� ���������)
;             :               $TreeScope_Subtree - �������� ��� ������� � ���� ��������
;             : $iWaitTime    - ����� �������� � �������� (�� ��������� $UIA_DefaultWaitTime, 0 - ����������)
; ����������  : �����   - ������� (������)
;             : ������� - 0 � ������������� @error
;             :         @error = 1 - ������ �������� �� �������� ��������
;             :         @error = 2 - ������� ������ �� ����������
;             :         @error = 3 - ��������� ����� ��������
; �����       : InnI
; ����������  : ������� ������ $TreeScope_Parent � $TreeScope_Ancestors �� ��������������
; ============================================================================================
Func _UIA_WaitElementFromCondition($oElementFrom, $oCondition, $iTreeScope = Default, $iWaitTime = Default)
  If Not IsObj($oElementFrom) Then Return SetError(1, __UIA_ConsoleWriteError("_UIA_WaitElementFromCondition : ������ �������� �� �������� ��������"), 0)
  If Not IsObj($oCondition) Then Return SetError(2, __UIA_ConsoleWriteError("_UIA_WaitElementFromCondition : ������� ������ �� ����������"), 0)
  If $iTreeScope = Default Then $iTreeScope = $TreeScope_Descendants
  If $iWaitTime = Default Then $iWaitTime = $UIA_DefaultWaitTime
  Local $oElement, $iOption = $UIA_ConsoleWriteError, $iTimer = TimerInit()
  $UIA_ConsoleWriteError = 0
  Do
    $oElement = _UIA_GetElementFromCondition($oElementFrom, $oCondition, $iTreeScope)
    If IsObj($oElement) Then
      $UIA_ConsoleWriteError = $iOption
      Return $oElement
    EndIf
    If $iWaitTime > 0 And TimerDiff($iTimer) > $iWaitTime * 1000 Then
      $UIA_ConsoleWriteError = $iOption
      Return SetError(3, __UIA_ConsoleWriteError("_UIA_WaitElementFromCondition : ��������� ����� ��������"), 0)
    EndIf
    Sleep(100)
  Until 0
EndFunc ; _UIA_WaitElementFromCondition

; ============================================================================================
; ��� ����������� �������������
; ============================================================================================
; ��� ������� : __UIA_ConsoleWriteError
; ��������    : ������� �������� ������ � �������
; ���������   : __UIA_ConsoleWriteError($sString)
; ���������   : $sString - ������ � ��������� ������
; ����������  : ������
; �����       : InnI
; ����������  :
; ============================================================================================
Func __UIA_ConsoleWriteError($sString)
  If $UIA_ConsoleWriteError Then ConsoleWrite("!> " & $sString & @CRLF)
EndFunc ; __UIA_ConsoleWriteError

; ============================================================================================
; ��� ������� : __UIA_CreateElement
; ��������    : ������ ������� (������)
; ���������   : __UIA_CreateElement($pElement)
; ���������   : $pElement - ��������� �� �������
; ����������  : �����   - ������� (������)
;             : ������� - 0
; �����       : InnI
; ����������  :
; ============================================================================================
Func __UIA_CreateElement($pElement)
  If Not $pElement Then Return 0
  Local $oElement, $Ver = ($UIA_ElementVersion = 1) ? "" : $UIA_ElementVersion
  If $UIA_ElementVersion Then
    Return ObjCreateInterface($pElement, Eval("sIID_IUIAutomationElement" & $Ver), Eval("dtagIUIAutomationElement" & $Ver))
  Else
    For $i = 7 To 2 Step -1
      $oElement = ObjCreateInterface($pElement, Eval("sIID_IUIAutomationElement" & $i), Eval("dtagIUIAutomationElement" & $i))
      If IsObj($oElement) Then Return $oElement
    Next
    Return ObjCreateInterface($pElement, $sIID_IUIAutomationElement, $dtagIUIAutomationElement)
  EndIf
EndFunc ; __UIA_CreateElement

; ============================================================================================
; ��� ������� : __UIA_GetPropIdFromStr
; ��������    : ����������� ������������ � ������� Inspect �������� � ��� �������� �������������
; ���������   : __UIA_GetPropIdFromStr($sString[, $sFuncName = ""])
; ���������   : $sString   - ������, ������������ � ����� ����� ������ ������� Inspect
;             : $sFuncName - ��� ���������� �������
; ����������  : �����   - �������� ������������� ��������
;             : ������� - -1 � ������������� @error
;             :         @error = 1 - ������ �������������� ������ � ������������� ��������
;             :         @error = 2 - ����������� ������������� ��������
; �����       : InnI
; ����������  :
; ============================================================================================
Func __UIA_GetPropIdFromStr($sString, $sFuncName = "")
  Local $iPropertyID, $sStr = StringReplace(StringRegExpReplace($sString, "[\.:\h]", ""), "property", "")
  If Int($sStr) Then
    $iPropertyID = Int($sStr)
  Else
    $iPropertyID = Eval("UIA_" & $sStr & "PropertyID")
    If @error Then Return SetError(1, __UIA_ConsoleWriteError($sFuncName & ' : ������ �������������� ������ "' & $sString & '" � ������������� ��������'), -1)
  EndIf
  If $iPropertyID < 30000 Or $iPropertyID > 30167 Then
    Return SetError(2, __UIA_ConsoleWriteError($sFuncName & " : ����������� ������������� �������� (" & $iPropertyID & ")"), -1)
  Else
    Return $iPropertyID
  EndIf
EndFunc ; __UIA_GetPropIdFromStr

; ============================================================================================
; ��� ������� : __UIA_GetTypeIdFromStr
; ��������    : ����������� ������������ � ������� Inspect �������� �������� ControlType � �������� �������������
; ���������   : __UIA_GetTypeIdFromStr($sString[, $sFuncName = ""])
; ���������   : $sString - ������, ������������ � �������� �������� ControlType ������� Inspect
;             : $sFuncName - ��� ���������� �������
; ����������  : �����   - �������� ������������� ����
;             : ������� - -1 � ������������� @error
;             :         @error = 1 - ������ �������������� ������ � ������������� ����
;             :         @error = 2 - ����������� ������������� ����
; �����       : InnI
; ����������  :
; ============================================================================================
Func __UIA_GetTypeIdFromStr($sString, $sFuncName = "")
  Local $iControlTypeID, $sStr = StringRegExpReplace($sString, "[\)\(\h]", "")
  If Int($sStr) Then
    $iControlTypeID = Int($sStr)
  Else
    $sStr = StringRegExp($sStr, "(?i)uia_.+typeid", 1)
    If IsArray($sStr) And IsDeclared($sStr[0]) Then
      $iControlTypeID = Eval($sStr[0])
    Else
      Return SetError(1, __UIA_ConsoleWriteError($sFuncName & ' : ������ �������������� ������ "' & $sString & '" � ������������� ����'), -1)
    EndIf
  EndIf
  If $iControlTypeID < 50000 Or $iControlTypeID > 50040 Then
    Return SetError(2, __UIA_ConsoleWriteError($sFuncName & " : ����������� ������������� ���� (" & $iControlTypeID & ")"), -1)
  Else
    Return $iControlTypeID
  EndIf
EndFunc ; __UIA_GetTypeIdFromStr

; #������_���_SCITE# =========================================================================
; ����������� ������ ������� (���� au3.user.calltips.api)
#cs
_UIA_CreateLogicalCondition($oCondition1[, $sOperator = "NOT"[, $oCondition2 = 0]]) ������ ���������� ������� �� ������ �������� ������� (���������: #include <UIAutomation.au3>)
_UIA_CreatePropertyCondition($vProperty, $vPropertyValue) ������ ������� �� ������ �������� � ��� �������� (���������: #include <UIAutomation.au3>)
_UIA_ElementDoDefaultAction($oElement) ���������� ��������� �������� �� ��������� (���������: #include <UIAutomation.au3>)
_UIA_ElementFindInArray(Const ByRef $aElements, $vProperty, $vPropertyValue[, $fInStr = False[, $iStartIndex = 1[, $fToEnd = True]]]) ������� �������, ��������������� ��������� �������� � ��� �������� (���������: #include <UIAutomation.au3>)
_UIA_ElementGetBoundingRectangle($oElement) ���������� ������������� �������, �������������� ������� (���������: #include <UIAutomation.au3>)
_UIA_ElementGetFirstLastChild($oElement) ������� ������ � ��������� �������� �������� (�������) ���������� �������� (���������: #include <UIAutomation.au3>)
_UIA_ElementGetParent($oElement) ���������� ������������ ������� (������) ���������� �������� (���������: #include <UIAutomation.au3>)
_UIA_ElementGetPreviousNext($oElement) ������� ���������� � ��������� �������� (�������) ���� �� ������ (���������: #include <UIAutomation.au3>)
_UIA_ElementGetPropertyValue($oElement, $vProperty) ���������� �������� ��������� �������� �������� (���������: #include <UIAutomation.au3>)
_UIA_ElementMouseClick($oElement[, $sButton = ""[, $iX = Default[, $iY = Default[, $iClicks = 1[, $fSetFocus = True]]]]]) ��������� ���� ���� �� �������� (���������: #include <UIAutomation.au3>)
_UIA_ElementScrollIntoView($oElement) ������������ ������� � ������� ��������� (���������: #include <UIAutomation.au3>)
_UIA_ElementSetFocus($oElement) ������������� �������� ����� ����� (���������: #include <UIAutomation.au3>)
_UIA_ElementTextSetValue($oElement, $sValue) ������������� �������� (�����) � ��������� ������� (���������: #include <UIAutomation.au3>)
_UIA_FindAllElements($oElementFrom[, $vProperty = 0[, $vPropertyValue = ""]]) ������� ��� ��������, ��������������� ��������� �������� � ��� �������� (���������: #include <UIAutomation.au3>)
_UIA_FindAllElementsEx($oElementFrom[, $oCondition = Default[, $iTreeScope = Default]]) ������� ��� ��������, ��������������� ������� ������ (���������: #include <UIAutomation.au3>)
_UIA_FindElementsInArray(Const ByRef $aElements, $vProperty, $vPropertyValue[, $fInStr = False[, $fIndexArray = False]]) ������� ��� ��������, ��������������� ��������� �������� � ��� �������� (���������: #include <UIAutomation.au3>)
_UIA_GetControlTypeElement($oElementFrom, $vControlType, $vPropertyValue[, $vProperty = Default[, $fInStr = False]]) ������� ������� (������) ���������� ���� � �������� ��������� � ��������� (���������: #include <UIAutomation.au3>)
_UIA_GetElementFromCondition($oElementFrom, $oCondition[, $iTreeScope = Default]) ������� ������� (������) �� ������ ��������� ������� (���������: #include <UIAutomation.au3>)
_UIA_GetElementFromFocus() ������ ������� (������) �� ������ ������ ����� (���������: #include <UIAutomation.au3>)
_UIA_GetElementFromHandle($hHandle) ������ ������� (������) �� ������ ����������� (���������: #include <UIAutomation.au3>)
_UIA_GetElementFromPoint($iX = Default, $iY = Default) ������ ������� (������) �� ������ �������� ��������� (���������: #include <UIAutomation.au3>)
_UIA_ObjectCreate() ������ ������ UIAutomation (���������: #include <UIAutomation.au3>)
_UIA_WaitControlTypeElement($oElementFrom, $vControlType, $vPropertyValue[, $vProperty = Default[, $fInStr = False[, $iWaitTime = Default]]]) ������� ������� (������) ���������� ���� � �������� ��������� � ��������� (���������: #include <UIAutomation.au3>)
_UIA_WaitElementFromCondition($oElementFrom, $oCondition[, $iTreeScope = Default[, $iWaitTime = Default]]) ������� ������� (������) �� ������ ��������� ������� (���������: #include <UIAutomation.au3>)
#ce
; ��������� ������� (���� au3.UserUdfs.properties)
#cs
au3.keywords.user.udfs=_uia_createlogicalcondition _uia_createpropertycondition _uia_elementdodefaultaction \
	_uia_elementfindinarray _uia_elementgetboundingrectangle _uia_elementgetfirstlastchild _uia_elementgetparent \
	_uia_elementgetpreviousnext _uia_elementgetpropertyvalue _uia_elementmouseclick _uia_elementscrollintoview \
	_uia_elementsetfocus _uia_elementtextsetvalue _uia_findallelements _uia_findallelementsex _uia_findelementsinarray \
	_uia_getcontroltypeelement _uia_getelementfromcondition _uia_getelementfromfocus _uia_getelementfromhandle \
	_uia_getelementfrompoint _uia_objectcreate _uia_waitcontroltypeelement _uia_waitelementfromcondition
#ce
; ============================================================================================