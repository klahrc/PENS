Attribute VB_Name = "MGlobals"
Option Explicit

'---------------------------------------------------------------------------------------
' Module Constant Declarations Follow
'---------------------------------------------------------------------------------------
Public Enum enumAnchorStyles
    enumAnchorStyleNone = 0
    enumAnchorStyletop = 1
    enumAnchorStyleBottom = 2
    enumAnchorStyleLeft = 4
    enumAnchorStyleRight = 8
End Enum


' Debug file handle
Public Const gsAPP_NAME As String = "PENS"
'''''Public Const GSAPPNAME As String = "Update Addin Demo"

Public Const gsCCC_Filename As String = "CCC.xltm"
Public Const gsOCXFilename As String = "iGrid500_10Tec.OCX"

Public Const gsFILE_DEBUG_LOG As String = "Debug.log"        ' The name of the file where debug messages will be logged to.

Public Const gsPENS_VERSION As String = "C1"

Public Const gsFILENAME_PP As String = "C3"
Public Const gsREM_FOLDER_PP As String = "C4"

Public Const gsFILENAME_RS As String = "C7"
Public Const gsREM_FOLDER_RS As String = "C8"

Public Const gsLOCAL_FOLDERS As String = "C10"

Public Const gsINFORM_LOCAL_COPY As String = "C11"

Public Const gsUSE_LOCAL_DATA As String = "C12"

Public Const gsDB_NAME As String = "C14"

Public Const gsUPDATE_FOLDER As String = "C15"

Public Const gsDEBUG_MODE As String = "C16"

Public Const gsRELEASE_TYPE As String = "E1"


'''Public gFrmExtPanel As frmExtraction
Public gFrmNavPanel As frmNavigation
Public gFrmSettings As frmSettings
Public gFrmCostDet As frmCostDetails
Public gFrmDetStatus As frmDetStatus
Public gFrmReleaseNotes As frmReleaseNotes
Public gFrmAbout As frmAbout
Public gFrmTips As frmTips

Public gbPPFileOK As Boolean        ' Portfolio Plan File found
Public gbRSFileOK As Boolean        ' Resource Spreadsheet File found

Public giDebugFile As Integer

Public gbInitialized As Boolean

Public gwsConfig As Worksheet

Public gbDebugMode As Boolean        ' True enables debug mode, False disables it.

Public gbConnected2Network As Boolean

Public guDictProj As Object

Public gColPrjInfo As Collection

Public gbCompletedNavPanelLoad As Boolean

Public gbComboZapCompleted As Boolean

Public gsPP_Filename As String
Public gsPP_Network_Folder As String
Public gsRS_Filename As String
Public gsRS_Network_Folder As String
Public gsLocal_Folder As String
Public gbUse_Local_Folder As Boolean

Public gsLastReport As String

Public gbJoin_BETA_Program As Boolean

Public gsTipsArray() As String

Public glPosTip As Long

Public gbFY17 As Boolean





