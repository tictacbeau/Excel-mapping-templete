Attribute VB_Name = "modConstants"
Option Explicit

' ============================================================
'  APPLICATION CONSTANTS
'  Remittance Allocation Engine v1.0
' ============================================================

' Application Info
Public Const APP_NAME As String = "Remittance Allocation Engine"
Public Const APP_VERSION As String = "1.0.0"
Public Const APP_TITLE As String = "Remittance Allocation Engine v1.0"

' ---- Sheet Names -------------------------------------------
Public Const SHEET_DASHBOARD As String = "Dashboard"
Public Const SHEET_PAYOR_REGISTRY As String = "PayorRegistry"
Public Const SHEET_TEMPLATE_DB As String = "TemplateDB"
Public Const SHEET_AUDIT_LOG As String = "AuditLog"
Public Const SHEET_SETTINGS As String = "Settings"
Public Const SHEET_PROCESSING As String = "Processing"
Public Const SHEET_OUTPUT As String = "Output"

' ---- PayorRegistry Sheet Columns ---------------------------
Public Const REG_COL_ID As Integer = 1            ' Unique payor ID (GUID-like)
Public Const REG_COL_NAME As Integer = 2           ' Display name
Public Const REG_COL_ALIASES As Integer = 3        ' JSON array of aliases
Public Const REG_COL_FILE_TYPE As Integer = 4      ' Expected input file type
Public Const REG_COL_PARSE_RULES As Integer = 5    ' JSON parse rule object
Public Const REG_COL_LOOKUP_CONFIG As Integer = 6  ' JSON lookup config object
Public Const REG_COL_PRINT_SETTINGS As Integer = 7 ' JSON print settings
Public Const REG_COL_NOTES As Integer = 8          ' Free-text notes
Public Const REG_COL_LAST_MODIFIED As Integer = 9  ' ISO timestamp
Public Const REG_COL_CREATED As Integer = 10       ' ISO timestamp
Public Const REG_HEADER_ROW As Integer = 1
Public Const REG_DATA_START_ROW As Integer = 2

' ---- TemplateDB Sheet Columns ------------------------------
' Each row = one cell definition in a template
Public Const TDB_COL_PAYOR_ID As Integer = 1       ' Foreign key to PayorRegistry
Public Const TDB_COL_TEMPLATE_TYPE As Integer = 2  ' "SINGLE" or "MULTI"
Public Const TDB_COL_SECTION As Integer = 3        ' "HEADER", "BODY", "FOOTER"
Public Const TDB_COL_ROW_OFFSET As Integer = 4     ' Row within section (1-based)
Public Const TDB_COL_COL_NUM As Integer = 5        ' Column number (1-based)
Public Const TDB_COL_CELL_TYPE As Integer = 6      ' DATA/LOOKUP/CALC/BLANK/LABEL
Public Const TDB_COL_FIELD_MAP As Integer = 7      ' Field name / formula reference
Public Const TDB_COL_DEFAULT_VALUE As Integer = 8  ' Static label text or default
Public Const TDB_COL_FORMAT_JSON As Integer = 9    ' JSON: font, fill, border, etc.
Public Const TDB_COL_IS_DYNAMIC As Integer = 10    ' TRUE if this row repeats per invoice
Public Const TDB_COL_MERGE_REF As Integer = 11     ' Merge range e.g. "A1:D1"
Public Const TDB_COL_FORMULA As Integer = 12       ' Formula string (CALC type)
Public Const TDB_HEADER_ROW As Integer = 1
Public Const TDB_DATA_START_ROW As Integer = 2

' ---- Settings Sheet Keys -----------------------------------
Public Const SETTING_LAST_PAYOR As String = "LastPayor"
Public Const SETTING_DEFAULT_OUTPUT_PATH As String = "DefaultOutputPath"
Public Const SETTING_DATA_SOURCE_PATH As String = "DataSourcePath"
Public Const SETTING_AUDIT_ENABLED As String = "AuditEnabled"
Public Const SETTING_BACKUP_DATA_PATH As String = "BackupDataPath"
Public Const SETTING_THEME As String = "Theme"

' ---- Cell Type Constants -----------------------------------
Public Const CELLTYPE_DATA As String = "DATA"       ' From remittance (blue)
Public Const CELLTYPE_LOOKUP As String = "LOOKUP"   ' From internal DB (green)
Public Const CELLTYPE_CALC As String = "CALC"       ' Calculated / formula (yellow)
Public Const CELLTYPE_BLANK As String = "BLANK"     ' Intentionally empty (white/hatched)
Public Const CELLTYPE_LABEL As String = "LABEL"     ' Static label (gray)

' ---- Cell Type Display Colors (Excel Long values) ----------
' Blue for DATA fields
Public Const COLOR_DATA As Long = 13395456          ' #CCE5FF (soft blue)
' Green for LOOKUP fields
Public Const COLOR_LOOKUP As Long = 13434828        ' #CCFFCC (soft green)
' Yellow for CALC fields
Public Const COLOR_CALC As Long = 16777164          ' #FFFF9C (soft yellow)
' White / hatched for BLANK
Public Const COLOR_BLANK As Long = 16777215         ' #FFFFFF
' Light gray for LABEL
Public Const COLOR_LABEL As Long = 15921906         ' #F2F2F2

' ---- Template Section Names --------------------------------
Public Const SECTION_HEADER As String = "HEADER"
Public Const SECTION_BODY As String = "BODY"
Public Const SECTION_FOOTER As String = "FOOTER"

' ---- Template Types ----------------------------------------
Public Const TMPL_SINGLE As String = "SINGLE"
Public Const TMPL_MULTI As String = "MULTI"

' ---- Input File Types --------------------------------------
Public Const FILETYPE_XLSX As String = "xlsx"
Public Const FILETYPE_XLS_GENUINE As String = "xls_genuine"
Public Const FILETYPE_XLS_HTML As String = "xls_html"
Public Const FILETYPE_CSV As String = "csv"

' OLE2 compound document magic bytes (genuine .xls)
Public Const OLE2_MAGIC As String = Chr(208) & Chr(207) & Chr(17) & Chr(224)

' ---- Delimiters --------------------------------------------
Public Const DELIM_COMMA As String = ","
Public Const DELIM_TAB As String = Chr(9)
Public Const DELIM_PIPE As String = "|"
Public Const DELIM_SEMICOLON As String = ";"
Public Const DELIM_SPACE As String = " "
Public Const DELIM_NEWLINE As String = Chr(10)

' ---- Power Query / Data Model ------------------------------
Public Const PQ_CONNECTION_NAME As String = "RemittanceInternalData"
Public Const DM_TABLE_NAME As String = "InternalData"
Public Const PQ_BACKUP_SUFFIX As String = "_BACKUP"

' ---- Audit Log ---------------------------------------------
Public Const AUDIT_MAX_ROWS As Long = 10000
Public Const AUDIT_COL_TIMESTAMP As Integer = 1
Public Const AUDIT_COL_ACTION As Integer = 2
Public Const AUDIT_COL_PAYOR As Integer = 3
Public Const AUDIT_COL_FILE As Integer = 4
Public Const AUDIT_COL_DETAIL As Integer = 5
Public Const AUDIT_COL_STATUS As Integer = 6
Public Const AUDIT_HEADER_ROW As Integer = 1
Public Const AUDIT_DATA_START_ROW As Integer = 2

' ---- Error Codes (user-defined) ----------------------------
Public Const ERR_FILE_NOT_FOUND As Long = 1001
Public Const ERR_CONVERSION_FAILED As Long = 1002
Public Const ERR_PAYOR_NOT_FOUND As Long = 1003
Public Const ERR_TEMPLATE_NOT_FOUND As Long = 1004
Public Const ERR_LOOKUP_FAILED As Long = 1005
Public Const ERR_REGEX_INVALID As Long = 1006
Public Const ERR_SCHEMA_DRIFT As Long = 1007
Public Const ERR_INVALID_JSON As Long = 1008
Public Const ERR_DATA_MODEL_FAILED As Long = 1009
Public Const ERR_NO_INVOICES As Long = 1010

' ---- Named Ranges ------------------------------------------
Public Const NR_SETTINGS_TABLE As String = "tblSettings"
Public Const NR_PAYOR_LIST As String = "tblPayors"

' ---- UI Colors (Dashboard theme) ---------------------------
Public Const UI_COLOR_PRIMARY As Long = 3494464     ' #354EA0 (navy blue - title bar)
Public Const UI_COLOR_SECONDARY As Long = 5197615   ' #4F7BEF (accent blue)
Public Const UI_COLOR_SUCCESS As Long = 5025616     ' #4CAF50 (green - success)
Public Const UI_COLOR_WARNING As Long = 2915072     ' #FF9800 (orange - warning)
Public Const UI_COLOR_ERROR As Long = 255           ' #FF0000 (red - error)
Public Const UI_COLOR_BG As Long = 16777215         ' #FFFFFF (white - background)
Public Const UI_COLOR_PANEL As Long = 15921906      ' #F2F2F2 (light gray - panels)
