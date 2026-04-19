Module molVar

    ' biến hệ thống
    Public Const SYS_STEEL_NODE_BLOCK_NAME As String = "KCS_STEEL"
    Public Const SYS_LAYER_AXIS_NAME As String = "KCS_AXIS"
    Public Const SYS_LAYER_THIN_NAME As String = "KCS_CHI"
    Public Const SYS_LAYER_BORDER_NAME As String = "KCS_BORDER"
    Public Const SYS_LAYER_TEXT_NAME As String = "KCS_TEXT"
    Public Const SYS_LAYER_STEEL_NAME As String = "KCS_STEEL"
    Public Const SYS_LAYER_STIRRUP_NAME As String = "THEP DAI"
    Public Const SYS_LAYER_HATCH_NAME As String = "HATCH"
    Public Const SYS_LAYER_DIM_NAME As String = "KCS_DIM"
    Public Const SYS_TEXT_WIDTH_FACTOR As Double = 0.75
    Public Const SYS_TEXT_STYLE_NAME As String = "KCS_TEXT"
    Public Const SYS_DIM_STYLE_NAME As String = "KCS_DIM25"
    Public Const SYS_TAG_CIRCLE_DIA As Double = 5
    Public Const SYS_TEXT_HEIGHT As Double = 2.5

    Public Const SYS_TextStyle As String = "KCS_SMTEXT"

    Public t As Double
    Public Abv As Double = 20
    Public Diembatdauvethep As Double
    Public sothanhmoc As Double
    Public Chieucaomoc As Double
    Dim a_bardot As Integer = 22

    Public Const SYS_DIM_STYLE As String = "KCS_DIM25"

    Public SYS_D_TextH_SM As Decimal = 62.5
    Public SYS_D_TextH_MID As Decimal = 100
    Public SYS_D_TextH_BIG As Decimal = 125
    Public SYS_D_TextWF As Decimal = 0.7

    Public Const SYS_D_DimFoot1 As Decimal = 200
    Public Const SYS_D_DimFoot2 As Decimal = 175
    Public Const SYS_D_DimFoot3 As Decimal = 125
    Public Const SYS_D_DimFoot4 As Decimal = 50
    Public Const SYS_D_DimFoot5 As Decimal = 150

    Public SYS_L_DIM As String
    Public SYS_L_SOTHEP_CIRCLE As String
    Public SYS_L_SOTHEP_TEXT As String
    Public SYS_L_TEXT As String

    Public SYS_KCS_CONFIG_BarNote_KyHieuDuongKinh As String = "%%c"
    Public SYS_KCS_CONFIG_BarNote_KyHieuKhoangCach As String = "@"

    Public TK_Dot_L2 As Decimal

    Public ThangBo_Config As STR_ThangBo_CONFIG
End Module
