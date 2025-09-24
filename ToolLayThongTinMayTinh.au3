#RequireAdmin

#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=o.ico
#AutoIt3Wrapper_Res_Description=ToolLayThongTinMayTinh_v2.0 ; tên hiển thị trong Task Manager
#AutoIt3Wrapper_Outfile=ToolLayThongTinMayTinh_v2.0.exe ; tên file đầu ra (.exe) cho ứng dụng
#AutoIt3Wrapper_Res_Fileversion=2.0.0.0
#AutoIt3Wrapper_Res_Companyname=Copyright@lqviet_10.08.2025
#AutoIt3Wrapper_Res_Language=1066 ; Vietnamese
#AutoIt3Wrapper_Run_Obfuscator=y  ; Sử dụng bộ làm rối mã nguồn
#AutoIt3Wrapper_UseUpx=y   ;sử dụng công cụ UPX để nén file .exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

; ================================== Khai báo thư viện và chế độ xử lý sự kiện ==================================
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <MsgBoxConstants.au3>
#include <StringConstants.au3>
#include <Constants.au3>
#include <File.au3>
#include <Array.au3>
#include <WinAPIReg.au3>

Opt("GUIOnEventMode", 1) ; Chế độ xử lý sự kiện dựa trên hàm

; ================================== Khai báo biến toàn cục ==================================
Global $txt_pcName, $txt_manu, $txt_model, $txt_os
Global $txt_office, $txt_security
Global $txt_mainboard, $txt_year, $txt_cpu, $txt_gpu
Global $txt_ram1, $txt_ram2, $txt_disk1, $txt_disk2
Global $btn_check_win_license, $btn_check_office_license
Global $btn_export, $btn_exit
Global $hGUI
Global $txt_total_ram, $txt_total_disk

; ================================== Tạo giao diện người dùng (GUI) ==================================
Func CreateGUI()
	$hGUI = GUICreate("Xem Thông tin hệ thống - v2.0", 700, 530) ; Tăng chiều cao GUI lên 530
	GUISetFont(10)

	Local $lblAuthor = GUICtrlCreateLabel("📌 Phát triển bởi: Lê Quốc Việt", 20, 10, 300, 20)
	GUICtrlSetColor($lblAuthor, 0xFF0000) ; Đỏ
	GUICtrlSetFont($lblAuthor, 10, 400, 2, "Arial") ; Cỡ 10, thường 400 (đậm 800), nghiêng 2

	; ================================== Vùng "Thông tin chung" ==================================
	GUICtrlCreateGroup("💡 Thông tin chung", 20, 40, 310, 230) ; Tăng chiều rộng nhóm
	GUICtrlCreateLabel("💻 Tên máy:", 30, 60, 110, 20)
	$txt_pcName = GUICtrlCreateInput("", 140, 60, 180, 20, $ES_READONLY + $ES_AUTOHSCROLL + $ES_LEFT)

	GUICtrlCreateLabel("🕋 Hãng sản xuất:", 30, 85, 110, 20)
	$txt_manu = GUICtrlCreateInput("", 140, 85, 180, 20, $ES_READONLY + $ES_AUTOHSCROLL + $ES_LEFT)

	GUICtrlCreateLabel("📰 Model:", 30, 110, 110, 20)
	$txt_model = GUICtrlCreateInput("", 140, 110, 180, 20, $ES_READONLY + $ES_AUTOHSCROLL + $ES_LEFT)

	GUICtrlCreateLabel("🪟 Hệ điều hành:", 30, 135, 110, 20)
	$txt_os = GUICtrlCreateInput("", 140, 135, 180, 20, $ES_READONLY + $ES_AUTOHSCROLL + $ES_LEFT)

	GUICtrlCreateLabel("✅ Bản quyền Win:", 30, 160, 110, 20)
	$btn_check_win_license = GUICtrlCreateButton("Kiểm tra", 140, 160, 80, 20)
	$btn_oem_key = GUICtrlCreateButton("OEM Key", 230, 160, 90, 20)
	GUICtrlSetTip($btn_oem_key, "Chỉ hiện Key khi có bản quyền OEM kèm theo máy tính")

	GUICtrlCreateLabel("📦 Office:", 30, 185, 110, 20)
	$txt_office = GUICtrlCreateInput("", 140, 185, 180, 20, $ES_LEFT + $ES_AUTOHSCROLL + $ES_READONLY)
	GUICtrlSetTip($txt_office, "16.0: Office 2016, 2019, 2021 & 365" & @CRLF & "15.0: Office 2013" & @CRLF & "14.0: Office 2010" & @CRLF & "12.0: Office 2007")

	GUICtrlCreateLabel("✅ BQuyền Office:", 30, 210, 130, 20)
	$btn_check_office_license = GUICtrlCreateButton("Kiểm tra", 140, 210, 180, 20)

	GUICtrlCreateLabel("🔒 Security:", 30, 235, 110, 20)
	$txt_security = GUICtrlCreateInput("", 140, 235, 180, 20, $ES_READONLY + $ES_AUTOHSCROLL + $ES_LEFT)
	GUICtrlCreateGroup("", -99, -99, 1, 1)

	; ================================== Vùng "Thông tin phần cứng" ==================================
	GUICtrlCreateGroup("📣 Thông tin phần cứng", 350, 40, 330, 155)
	GUICtrlCreateLabel("🗂 Mainboard:", 360, 60, 100, 20)
	$txt_mainboard = GUICtrlCreateInput("", 460, 60, 210, 20, $ES_READONLY + $ES_AUTOHSCROLL + $ES_LEFT)

	GUICtrlCreateLabel("♾️ Năm sản xuất:", 360, 85, 130, 20)
	$txt_year = GUICtrlCreateInput("", 460, 85, 210, 20, $ES_READONLY + $ES_AUTOHSCROLL + $ES_LEFT)
	GUICtrlSetTip($txt_year, "Năm sản xuất được tính theo Mainboard")
	GUICtrlCreateLabel("🎟 CPU:", 360, 110, 60, 20)
	$txt_cpu = GUICtrlCreateInput("", 420, 110, 250, 20, $ES_READONLY + $ES_AUTOHSCROLL + $ES_LEFT)

	GUICtrlCreateLabel("🎞 Card đồ hoạ:", 360, 135, 100, 20)
	$txt_gpu = GUICtrlCreateInput("", 460, 135, 210, 20, $ES_READONLY + $ES_AUTOHSCROLL + $ES_LEFT)
	GUICtrlCreateGroup("", -99, -99, 1, 1)

	; ================================== Vùng "RAM" ==================================
	GUICtrlCreateGroup("💳 RAM", 20, 280, 660, 105) ; Tăng chiều cao nhóm RAM
	GUICtrlCreateLabel("🎞 Khe 1:", 30, 300, 70, 20)
	$txt_ram1 = GUICtrlCreateInput("", 100, 300, 560, 20, $ES_READONLY + $ES_AUTOHSCROLL + $ES_LEFT)
	GUICtrlCreateLabel("🎞 Khe 2:", 30, 325, 70, 20)
	$txt_ram2 = GUICtrlCreateInput("", 100, 325, 560, 20, $ES_READONLY + $ES_AUTOHSCROLL + $ES_LEFT)
	GUICtrlCreateLabel("Tổng cộng:", 30, 350, 70, 20) ; Thêm dòng "Tổng cộng"
	$txt_total_ram = GUICtrlCreateInput("", 100, 350, 560, 20, $ES_READONLY + $ES_AUTOHSCROLL + $ES_LEFT) ; Thêm ô input tổng RAM
	GUICtrlCreateGroup("", -99, -99, 1, 1)

	; ================================== Vùng "Ổ cứng" ==================================
	GUICtrlCreateGroup("🛢 Ổ cứng", 20, 390, 660, 105) ; Điều chỉnh vị trí và tăng chiều cao
	GUICtrlCreateLabel("🗄 Ổ đĩa 1:", 30, 410, 70, 20)
	$txt_disk1 = GUICtrlCreateInput("", 100, 410, 560, 20, $ES_READONLY + $ES_AUTOHSCROLL + $ES_LEFT)
	GUICtrlSetTip($txt_disk1, "Trên Win 7, chỉ có thể xác định là HDD, xác định là SSD chỉ hỗ trợ từ Win 8 trở lên")
	GUICtrlCreateLabel("🗄 Ổ đĩa 2:", 30, 435, 70, 20)
	$txt_disk2 = GUICtrlCreateInput("", 100, 435, 560, 20, $ES_READONLY + $ES_AUTOHSCROLL + $ES_LEFT)
	GUICtrlCreateLabel("Tổng cộng:", 30, 460, 70, 20) ; Thêm dòng "Tổng cộng"
	$txt_total_disk = GUICtrlCreateInput("", 100, 460, 560, 20, $ES_READONLY + $ES_AUTOHSCROLL + $ES_LEFT) ; Thêm ô input tổng ổ cứng
	GUICtrlCreateGroup("", -99, -99, 1, 1)

	; ================================== Các nút chức năng (căn giữa) ==================================
	Local $iBtnWidth = 150
	Local $iBtnSpacing = 20
	Local $iTotalWidth = (2 * $iBtnWidth) + $iBtnSpacing
	Local $iStartPos = (700 - $iTotalWidth) / 2

	$btn_export = GUICtrlCreateButton("📝 Xuất kết quả", $iStartPos, 495, $iBtnWidth, 30) ; Điều chỉnh vị trí nút
	$btn_exit = GUICtrlCreateButton("❌ Thoát", $iStartPos + $iBtnWidth + $iBtnSpacing, 495, $iBtnWidth, 30) ; Điều chỉnh vị trí nút

	GUICtrlSetOnEvent($btn_check_win_license, "CheckWinLicense")
	GUICtrlSetOnEvent($btn_oem_key, "CheckOemKey")
	GUICtrlSetOnEvent($btn_check_office_license, "CheckOfficeLicense")
	GUICtrlSetOnEvent($btn_export, "HandleExport")
	GUICtrlSetOnEvent($btn_exit, "HandleExit")
	GUISetOnEvent($GUI_EVENT_CLOSE, "HandleExit")

	GUISetState(@SW_SHOW)
EndFunc

; ================================== Hàm chính để lấy và hiển thị thông tin ==================================
Func GetAndDisplayInfo()
    Local $oWMI = ObjGet("winmgmts:\\.\root\cimv2")

    ; Thông tin chung
    GUICtrlSetData($txt_pcName, @ComputerName)
    Local $colCS = $oWMI.ExecQuery("SELECT * FROM Win32_ComputerSystem")
    If Not @error And $colCS.Count > 0 Then
        Local $oCS = $colCS.ItemIndex(0)
        GUICtrlSetData($txt_manu, $oCS.Manufacturer)
        GUICtrlSetData($txt_model, $oCS.Model)
    EndIf

    ; Thông tin hệ điều hành
    Local $colOS = $oWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")
    If Not @error And $colOS.Count > 0 Then
        Local $oOS = $colOS.ItemIndex(0)
        GUICtrlSetData($txt_os, $oOS.Caption & " (" & @OSArch & ")")
		GUICtrlSetData($txt_os, GUICtrlRead($txt_os))
    EndIf

    ; Thông tin Mainboard
    Local $colMB = $oWMI.ExecQuery("SELECT * FROM Win32_BaseBoard")
    If Not @error And $colMB.Count > 0 Then
        Local $oMB = $colMB.ItemIndex(0)
        GUICtrlSetData($txt_mainboard, $oMB.Product & " (" & $oMB.Manufacturer & ")")
    EndIf

    ; Thông tin năm sản xuất
    Local $colBIOS = $oWMI.ExecQuery("SELECT * FROM Win32_BIOS")
    If Not @error And $colBIOS.Count > 0 Then
        Local $oBIOS = $colBIOS.ItemIndex(0)
        Local $sBiosDate = $oBIOS.ReleaseDate
        If StringLen($sBiosDate) >= 4 Then
            GUICtrlSetData($txt_year, StringMid($sBiosDate, 1, 4))
        Else
            GUICtrlSetData($txt_year, "Chưa lấy được thông tin")
        EndIf
    EndIf

    ; Lấy thông tin Office
    GUICtrlSetData($txt_office, GetOfficeInfo())

    ; Phần mềm diệt virus
    GUICtrlSetData($txt_security, GetSecuritySoftwareInfo())

    ; Thông tin CPU
    Local $colCPU = $oWMI.ExecQuery("SELECT * FROM Win32_Processor")
    If Not @error And $colCPU.Count > 0 Then
        Local $oCPU = $colCPU.ItemIndex(0)
        GUICtrlSetData($txt_cpu, $oCPU.Name)
    EndIf

    ; Thông tin GPU
    Local $colGPU = $oWMI.ExecQuery("SELECT * FROM Win32_VideoController")
    If Not @error And $colGPU.Count > 0 Then
        Local $oGPU = $colGPU.ItemIndex(0)
        GUICtrlSetData($txt_gpu, $oGPU.Name)
    EndIf

    ; Thông tin RAM (có thêm tên hãng)
    Local $colMem = $oWMI.ExecQuery("SELECT * FROM Win32_PhysicalMemory")
    If Not @error And IsObj($colMem) Then
        Local $iTotalRAM_GB = 0
        For $i = 0 To $colMem.Count - 1
            Local $oMem = $colMem.ItemIndex($i)

            ; Lấy dung lượng và chuyển đổi sang GB
            Local $sizeGB = Round($oMem.Capacity / (1024 * 1024 * 1024), 0)
            $iTotalRAM_GB += $sizeGB ; Cộng dồn vào tổng

            Local $sManufacturer = $oMem.Manufacturer
            If StringLen($sManufacturer) = 0 Then
                $sManufacturer = "Unknown"
            EndIf

            Local $sRamInfo = $sizeGB & "GB - " & $sManufacturer & " - Bus " & $oMem.Speed & " MHz"
            If $i = 0 Then
                GUICtrlSetData($txt_ram1, $sRamInfo)
            ElseIf $i = 1 Then
                GUICtrlSetData($txt_ram2, $sRamInfo)
            EndIf
        Next
        ; Hiển thị tổng dung lượng RAM sau khi vòng lặp kết thúc
		GUICtrlSetData($txt_total_ram, $iTotalRAM_GB & " GB")
	EndIf

    ; Thông tin ổ cứng
	Local $colDisk = $oWMI.ExecQuery("SELECT * FROM Win32_DiskDrive")
	If Not @error And IsObj($colDisk) Then
		Local $iTotalDisk_GB = 0 ; Biến tổng dung lượng theo giá trị thông thường

		For $i = 0 To $colDisk.Count - 1
			Local $oDisk = $colDisk.ItemIndex($i)

			; Lấy dung lượng thực tế và làm tròn sang GB
			Local $sizeActualGB = Round($oDisk.Size / (1024 * 1024 * 1024), 0)

			Local $sDiskSizeDisplay = ""
			Local $iDiskSizeConventional = 0

			; Xac dinh dung luong chuan thuc te va dung luong de hien thi
			If $sizeActualGB > 1600 Then
				$iDiskSizeConventional = 2048
				$sDiskSizeDisplay = $iDiskSizeConventional & " GB (2TB)"
			ElseIf $sizeActualGB > 900 Then
				$iDiskSizeConventional = 1024
				$sDiskSizeDisplay = $iDiskSizeConventional & " GB (1TB)"
			ElseIf $sizeActualGB > 400 Then
				$iDiskSizeConventional = 512
				$sDiskSizeDisplay = $iDiskSizeConventional & " GB"
			ElseIf $sizeActualGB > 220 Then
				$iDiskSizeConventional = 256
				$sDiskSizeDisplay = $iDiskSizeConventional & " GB"
			ElseIf $sizeActualGB > 110 Then
				$iDiskSizeConventional = 128
				$sDiskSizeDisplay = $iDiskSizeConventional & " GB"
			Else
				$iDiskSizeConventional = $sizeActualGB
				$sDiskSizeDisplay = $iDiskSizeConventional & " GB"
			EndIf

			; Cộng dồn dung lượng thông thường vào tổng
			$iTotalDisk_GB += $iDiskSizeConventional

			; Xác định loại ổ đĩa
			Local $sDiskType = "Unknown"
			If $oDisk.InterfaceType = "USB" Then
				$sDiskType = "USB"
			Else
				If @OSVersion = "WIN_7" Then
					If $oDisk.MediaType = "Fixed hard disk media" Then
						$sDiskType = "HDD"
					EndIf
				Else
					Local $oStorageWMI = ObjGet("winmgmts:\\.\root\microsoft\windows\storage")
					If IsObj($oStorageWMI) Then
						Local $colParts = $oStorageWMI.ExecQuery("SELECT * FROM MSFT_PhysicalDisk WHERE FriendlyName LIKE '%" & StringStripWS($oDisk.Model, 3) & "%'")
						If Not @error And IsObj($colParts) And $colParts.Count > 0 Then
							Local $oPart = $colParts.ItemIndex(0)
							If $oPart.MediaType = 4 Then
								$sDiskType = "SSD"
							ElseIf $oPart.MediaType = 3 Then
								$sDiskType = "HDD"
							EndIf
						EndIf
					EndIf
				EndIf
			EndIf

			; Tạo chuỗi thông tin cuối cùng
			Local $sDiskInfo = $sDiskType & " - " & $sDiskSizeDisplay & " - " & $oDisk.Model & "-" & $sizeActualGB & "GB"

			If $i = 0 Then
				GUICtrlSetData($txt_disk1, $sDiskInfo)
			ElseIf $i = 1 Then
				GUICtrlSetData($txt_disk2, $sDiskInfo)
			EndIf
		Next

		; Hiển thị tổng dung lượng ổ cứng sau khi vòng lặp kết thúc
		Local $sTotalDisplay = $iTotalDisk_GB & " GB"
		If $iTotalDisk_GB >= 1024 Then
			Local $dTotalTB = Round($iTotalDisk_GB / 1024, 0)
			$sTotalDisplay &= " (" & $dTotalTB & "TB)"
		EndIf
		GUICtrlSetData($txt_total_disk, $sTotalDisplay)
	EndIf

	;======================================================================
	; Tạo một mảng chứa ID của tất cả các ô input cần xử lý
    Local $aInputControls = [ _
        $txt_pcName, $txt_manu, $txt_model, $txt_os, $txt_office, $txt_security, _
        $txt_mainboard, $txt_year, $txt_cpu, $txt_gpu, $txt_ram1, $txt_ram2, _
        $txt_disk1, $txt_disk2 _
    ]

    For $i = 0 To UBound($aInputControls) - 1
        ; Kiểm tra xem điều khiển có tồn tại không
        If IsHWnd(GUICtrlGetHandle($aInputControls[$i])) Then
            ; Tập trung vào ô input đó
            ControlFocus($hGUI, "", $aInputControls[$i])
            ; Gửi phím "Home" để di chuyển con trỏ về đầu
            Send("{Home}")
        EndIf
    Next
    ; Sau khi hoàn tất, bạn có thể tập trung lại vào một điều khiển khác hoặc để mặc định
    ControlFocus($hGUI, "", $btn_export)

EndFunc

; ================================== Hàm: CheckWinLicense - Kiểm tra bản quyền Windows ==================================
Func CheckWinLicense()
    Run(@ComSpec & " /k slmgr /xpr", "", @SW_SHOW)
EndFunc

; ================================== Hàm: CheckOemKey - Lấy và hiển thị OEM Key ==================================
Func CheckOemKey()
    Run(@ComSpec & " /k wmic path softwareLicensingService get OA3xOriginalProductKey" & @CRLF & "PAUSE", "", @SW_SHOW)
EndFunc

; ================================== Hàm: GetOfficeInfo - Lấy thông tin Office ==================================
Func GetOfficeInfo()
    Local $aOfficeDirs[2] = [ _
        @ProgramFilesDir & "\Microsoft Office", _
        @ProgramFilesDir & " (x86)\Microsoft Office" _
    ]

    Local $iHighestVersion = 0

    ; Duyệt cả hai thư mục Office
    For $i = 0 To 1
        If FileExists($aOfficeDirs[$i]) Then
            Local $aFiles = _FileListToArray($aOfficeDirs[$i], "Office*", $FLTA_FOLDERS)
            If IsArray($aFiles) Then
                For $j = 1 To $aFiles[0] ; Bỏ phần tử 0 vì là số lượng
                    ; Kiểm tra tên dạng "Office16", "Office15", ...
                    If StringRegExp($aFiles[$j], "^Office(\d{2})$") Then
                        Local $iVersion = Number(StringMid($aFiles[$j], 7, 2))
                        If $iVersion > $iHighestVersion Then
                            $iHighestVersion = $iVersion
                        EndIf
                    EndIf
                Next
            EndIf
        EndIf
    Next

    ; Trả về kết quả
    If $iHighestVersion > 0 Then
        Local $sOfficeVersionName = ""
        Switch $iHighestVersion
            Case 16
                $sOfficeVersionName = "16.0 - Office 2016/365/2019/2021/2024"
            Case 15
                $sOfficeVersionName = "15.0 - Office 2013"
            Case 14
                $sOfficeVersionName = "14.0 - Office 2010"
            Case 12
                $sOfficeVersionName = "12.0 - Office 2007"
            Case 11
                $sOfficeVersionName = "11.0 - Office 2003"
            Case Else
                $sOfficeVersionName = "Phiên bản không xác định"
        EndSwitch

        ; Lấy bitness thực tế
        Local $sOfficeBit = _GetOfficeBitness()

        Return "Microsoft Office (" & $sOfficeVersionName & " - " & $sOfficeBit & ")"
    EndIf

    Return "Chưa lấy được thông tin"
EndFunc

; ======================
; _GetOfficeBitness()
; - Ưu tiên dò trong Common Files\Microsoft Shared (OfficeXX)
; - Nếu không thấy -> kiểm tra thư mục cài đặt chính
; - Trả: "32-bit", "64-bit", "Cả hai (32-bit & 64-bit)" hoặc "Không xác định"
; ======================
Func _GetOfficeBitness()
    ; Lấy đường dẫn Program Files chính xác
    Local $sPF64 = EnvGet("ProgramW6432")         ; trên Win64: C:\Program Files
    Local $sPF32 = EnvGet("ProgramFiles(x86)")    ; trên Win64: C:\Program Files (x86)
    If $sPF64 = "" Then $sPF64 = @ProgramFilesDir ; fallback
    ; Nếu hệ 32-bit thì EnvGet("ProgramFiles(x86)") có thể rỗng -> xử lý
    If $sPF32 = "" And @OSArch = "X86" Then $sPF32 = ""

    ; Common Files paths
    Local $sCommon64 = $sPF64 & "\Common Files\Microsoft Shared"
    Local $sCommon32 = ($sPF32 <> "") ? $sPF32 & "\Common Files\Microsoft Shared" : ""

    ; Tìm version cao nhất trong mỗi path
    Local $iHighest64 = _FindHighestOfficeInPath($sCommon64)
    Local $iHighest32 = _FindHighestOfficeInPath($sCommon32)

    ; Nếu tìm thấy ở Common Files -> quyết định dựa vào version tìm được
    If $iHighest64 > 0 Or $iHighest32 > 0 Then
        If $iHighest64 > $iHighest32 Then Return "64-bit"
        If $iHighest32 > $iHighest64 Then Return "32-bit"
        ; cùng version cao nhất ở cả hai nhánh
        If $iHighest64 > 0 Then Return "Cả hai (32-bit & 64-bit)"
    EndIf

    ; Fallback: kiểm tra thư mục cài đặt chính
    Local $sMain64 = $sPF64 & "\Microsoft Office"
    Local $sMain32 = ($sPF32 <> "") ? $sPF32 & "\Microsoft Office" : ""
    Local $bMain64 = FileExists($sMain64)
    Local $bMain32 = ($sMain32 <> "") ? FileExists($sMain32) : False

    If $bMain64 And Not $bMain32 Then Return "64-bit"
    If $bMain32 And Not $bMain64 Then Return "32-bit"
    If $bMain64 And $bMain32 Then Return "Cả hai (32-bit & 64-bit)"

    Return "Không xác định"
EndFunc

; ======================
; _FindHighestOfficeInPath($sPath)
; - Trả về số nguyên version cao nhất tìm thấy trong $sPath (ví dụ 16 cho Office16)
; - Nếu không tìm thấy trả 0
; ======================
Func _FindHighestOfficeInPath($sPath)
    If $sPath = "" Then Return 0
    If Not FileExists($sPath) Then Return 0

    ; Lấy danh sách thư mục con
    Local $aDirs = _FileListToArray($sPath, "*", $FLTA_FOLDERS)
    If Not IsArray($aDirs) Then Return 0

    Local $iHighest = 0
    For $i = 1 To $aDirs[0]
        Local $sName = $aDirs[$i]
        ; Tìm vị trí "Office" (không phân biệt hoa thường)
        Local $posOffice = StringInStr(StringUpper($sName), "OFFICE")
        If $posOffice > 0 Then
            ; Lấy phần nằm sau chữ "Office"
            Local $sAfter = StringTrimLeft($sName, $posOffice + 5) ; pos + len("Office") -1
            ; Lấy chuỗi số liên tiếp đầu tiên trong phần còn lại
            Local $sDigits = StringRegExpReplace($sAfter, "\D", "") ; loại hết ký tự không số
            If StringLen($sDigits) > 0 Then
                ; Nếu có nhiều chữ số (không thường), lấy 2 chữ số đầu để biểu diễn 11/12/14/15/16...
                Local $sTake = (StringLen($sDigits) >= 2) ? StringLeft($sDigits, 2) : $sDigits
                Local $iVer = Number($sTake)
                If $iVer > $iHighest Then $iHighest = $iVer
            EndIf
        EndIf
    Next

    Return $iHighest
EndFunc

; ======================
; Hàm xác định bitness của Office từ Registry
; ======================
;Func _GetOfficeBitness()
;    Local Const $HKEY_LOCAL_MACHINE = "HKEY_LOCAL_MACHINE"
;    Local $sKey64 = "SOFTWARE\Microsoft\Office"
;    Local $sKey32 = "SOFTWARE\WOW6432Node\Microsoft\Office"

;    If RegEnumKey($HKEY_LOCAL_MACHINE & "\" & $sKey32, 1) <> "" Then
;        Return "32-bit"
;    EndIf
;    If RegEnumKey($HKEY_LOCAL_MACHINE & "\" & $sKey64, 1) <> "" Then
;        Return "64-bit"
;    EndIf
;    Return "Không xác định"
;EndFunc

; ================================== Hàm: CheckOfficeLicense - Kiểm tra bản quyền Office ==================================
Func CheckOfficeLicense()
    Local $sOfficePath = ""
    Local $sOfficePath64 = @ProgramFilesDir & "\Microsoft Office"
    Local $sOfficePath32 = @ProgramFilesDir & "(x86)\Microsoft Office"

    Local $sSearchPath = ""
    Local $iVersion = 16
    For $iVersion = 16 To 12 Step -1
        If FileExists($sOfficePath64 & "\Office" & $iVersion & "\ospp.vbs") Then
            $sSearchPath = $sOfficePath64 & "\Office" & $iVersion & "\"
            ExitLoop
        ElseIf FileExists($sOfficePath32 & "\Office" & $iVersion & "\ospp.vbs") Then
            $sSearchPath = $sOfficePath32 & "\Office" & $iVersion & "\"
            ExitLoop
        EndIf
    Next

    If $sSearchPath <> "" Then
        Local $sCmd = 'cscript.exe //Nologo "' & $sSearchPath & 'ospp.vbs" /dstatus'
        ; Sử dụng Run thay vì RunWait
        Run(@ComSpec & " /k " & $sCmd, "", @SW_SHOW)
    Else
        MsgBox($MB_ICONERROR, "Lỗi", "Không tìm thấy file 'ospp.vbs'. Vui lòng kiểm tra lại đường dẫn cài đặt Office.")
    EndIf
EndFunc

; ================================== Hàm: GetSecuritySoftwareInfo - Lấy thông tin phần mềm diệt virus ==================================
Func GetSecuritySoftwareInfo()
    Local $sSecurityInfo = ""

    ; --- Phương pháp 1: Sử dụng WMI ---
    Local $oWMI = ""
    If @OSVersion = "WIN_7" Then
        $oWMI = ObjGet("winmgmts:\\.\root\SecurityCenter")
    Else
        $oWMI = ObjGet("winmgmts:\\.\root\SecurityCenter2")
    EndIf

    If IsObj($oWMI) Then
        Local $colItems = $oWMI.ExecQuery("SELECT * FROM AntiVirusProduct")
        If Not @error And IsObj($colItems) Then
            For $oItem In $colItems
                If StringLen($oItem.displayName) > 0 Then
                    If Not StringInStr($sSecurityInfo, $oItem.displayName, $STR_NOCASESENSEBASIC) Then
                        $sSecurityInfo &= $oItem.displayName & " "
                    EndIf
                EndIf
            Next
        EndIf
    EndIf

    ; Nếu có kết quả từ WMI → trả về ngay
    If StringLen($sSecurityInfo) > 0 Then
        Return StringStripWS($sSecurityInfo, $STR_STRIPTRAILING)
    EndIf

   ; --- Phương pháp 2: Kiểm tra thư mục cài đặt Antivirus ---
	Local $aAntivirusPaths[] = [ _
		@ProgramFilesDir & "\Windows Defender", _
		@ProgramFilesDir & "\ESET", _
		@ProgramFilesDir & "\Kaspersky Lab", _
		@ProgramFilesDir & "\McAfee", _
		@ProgramFilesDir & "\Norton Security", _
		@ProgramFilesDir & "\AVG", _
		@ProgramFilesDir & "\Avast Software", _
		@ProgramFilesDir & "\Bitdefender", _
		@ProgramFilesDir & "\Trend Micro", _
		@ProgramFilesDir & "\Sophos" _
	]

	; Thêm thư mục Program Files (x86) nếu là Windows 64-bit
	If @OSArch = "X64" Then
		Local $iUBound = UBound($aAntivirusPaths)
		ReDim $aAntivirusPaths[$iUBound + 10]
		$aAntivirusPaths[$iUBound + 0] = @ProgramFilesDir & " (x86)\ESET"
		$aAntivirusPaths[$iUBound + 1] = @ProgramFilesDir & " (x86)\Kaspersky Lab"
		$aAntivirusPaths[$iUBound + 2] = @ProgramFilesDir & " (x86)\McAfee"
		$aAntivirusPaths[$iUBound + 3] = @ProgramFilesDir & " (x86)\Norton Security"
		$aAntivirusPaths[$iUBound + 4] = @ProgramFilesDir & " (x86)\AVG"
		$aAntivirusPaths[$iUBound + 5] = @ProgramFilesDir & " (x86)\Avast Software"
		$aAntivirusPaths[$iUBound + 6] = @ProgramFilesDir & " (x86)\Bitdefender"
		$aAntivirusPaths[$iUBound + 7] = @ProgramFilesDir & " (x86)\Trend Micro"
		$aAntivirusPaths[$iUBound + 8] = @ProgramFilesDir & " (x86)\Sophos"
		$aAntivirusPaths[$iUBound + 9] = @ProgramFilesDir & " (x86)\Windows Defender"
	EndIf

	; Duyệt danh sách thư mục để kiểm tra
	For $i = 0 To UBound($aAntivirusPaths) - 1
		If FileExists($aAntivirusPaths[$i]) Then
			Local $sProductName = StringTrimLeft($aAntivirusPaths[$i], StringInStr($aAntivirusPaths[$i], "\", 0, -1))
			If Not StringInStr($sSecurityInfo, $sProductName, $STR_NOCASESENSEBASIC) Then
				$sSecurityInfo &= $sProductName & " "
			EndIf
		EndIf
	Next

    If StringLen($sSecurityInfo) > 0 Then
        Return StringStripWS($sSecurityInfo, $STR_STRIPTRAILING)
    Else
        Return "Chưa lấy được thông tin"
    EndIf
EndFunc

; ================================== Vòng lặp chính và các hàm xử lý sự kiện ==================================
CreateGUI()
Sleep(1000)
GetAndDisplayInfo()

While 1
    Sleep(100)
WEnd

; === Hàm: HandleExport - Xử lý khi nhấn nút "Xuất kết quả" ===
Func HandleExport()
    ; Local $sFilePath = @ScriptDir & "\SystemInfo.txt"
	Local $sFilePath = @ScriptDir & "\" & GUICtrlRead($txt_pcName) & "-TTinPC.txt"
	; Chế độ ghi đè file:
    ; 1 = Mở file ở chế độ ghi, nếu file đã tồn tại thì nội dung cũ sẽ bị xóa và ghi lại từ đầu.
    Local $hFile = FileOpen($sFilePath, 1)
    If $hFile = -1 Then
        MsgBox($MB_ICONERROR, "Lỗi", "Không thể tạo file: " & $sFilePath)
        Return
    EndIf

    FileWriteLine($hFile, "📌 Thông tin hệ thống - Phát triển bởi: Lê Quốc Việt")
    FileWriteLine($hFile,"ToolLayThongTinMayTinh_v2.0")
    FileWriteLine($hFile, "=========================================================")
    FileWriteLine($hFile, "Tên máy: " & GUICtrlRead($txt_pcName))
    FileWriteLine($hFile, "Hãng sản xuất: " & GUICtrlRead($txt_manu))
    FileWriteLine($hFile, "Model: " & GUICtrlRead($txt_model))
    FileWriteLine($hFile, "Mainboard: " & GUICtrlRead($txt_mainboard))
    FileWriteLine($hFile, "Năm sản xuất: " & GUICtrlRead($txt_year))
    FileWriteLine($hFile, "Hệ điều hành: " & GUICtrlRead($txt_os))
    FileWriteLine($hFile, "Bản quyền Win: " & "Nhấn nút 'Kiểm tra' trên giao diện để xem")
    FileWriteLine($hFile, "Office: " & GUICtrlRead($txt_office))
    FileWriteLine($hFile, "BQ Office: " & "Nhấn nút 'Kiểm tra' trên giao diện để xem")
    FileWriteLine($hFile, "Security: " & GUICtrlRead($txt_security))
    FileWriteLine($hFile, "CPU: " & GUICtrlRead($txt_cpu))
    FileWriteLine($hFile, "Card đồ hoạ: " & GUICtrlRead($txt_gpu))
    FileWriteLine($hFile, "RAM (Khe 1): " & GUICtrlRead($txt_ram1))
    FileWriteLine($hFile, "RAM (Khe 2): " & GUICtrlRead($txt_ram2))
    FileWriteLine($hFile, "Tổng RAM: " & GUICtrlRead($txt_total_ram))
    FileWriteLine($hFile, "Ổ đĩa 1: " & GUICtrlRead($txt_disk1))
    FileWriteLine($hFile, "Ổ đĩa 2: " & GUICtrlRead($txt_disk2))
	FileWriteLine($hFile, "Tổng ổ đĩa: " & GUICtrlRead($txt_total_disk))

    FileClose($hFile)
    MsgBox($MB_ICONINFORMATION, "Thành công", "Đã xuất thông tin ra file: " & $sFilePath)
EndFunc

; === Hàm: HandleExit - Xử lý khi nhấn nút "Thoát" hoặc đóng cửa sổ ===
Func HandleExit()
    Exit
EndFunc