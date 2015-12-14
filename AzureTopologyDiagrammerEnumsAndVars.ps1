# ===================================================================================
# Name: Azure Topology Diagrammer
# Desc: Enums and Vars for Visio diagramming
# ===================================================================================

# RGB Codes
$rgbAzure = "0,120,215"
$rgbGeneral = "150,150,150"
$rgbOffice365 = "220,60,0"
$rgbOnPrem = "0,24,143"

# visSectionIndices:  https://msdn.microsoft.com/EN-US/library/office/ff765983.aspx    
[int]$visSectionObject = 1
[int]$visSectionCharacter = 3
[int]$visSectionParagraph = 4

# visRowIndices:      https://msdn.microsoft.com/EN-US/library/office/ff765539.aspx
[int]$visRowRulerGrid = 18
[int]$visRowPageLayout = 24
[int]$visRowCharacter = 0
[int]$visRowLine = 2
[int]$visRowParagraph = 0

# visCellIndices:     https://msdn.microsoft.com/EN-US/library/office/ff767991.aspx
[int]$visCharacterColor = 1
[int]$visCharacterDblUnderline = 8
[int]$visHorzAlign = 6
[int]$visLineColor = 1
[int]$visLinePattern = 2
[int]$visPLOPlaceStyle = 8
[int]$visPLORouteStyle = 9
[int]$visXRulerOrigin = 4
[int]$visXGridOrigin = 10
[int]$visYRulerOrigin = 5
[int]$visYGridOrigin = 11

# visDiagramServices: https://msdn.microsoft.com/en-us/library/office/ff768414(v=office.15).aspx
[int]$visServiceVersion150 = 8

# visUnitCodes: https://msdn.microsoft.com/EN-US/library/office/ff769148.aspx
[int]$visInches = 65

# visListDirection: https://msdn.microsoft.com/EN-US/library/office/ff766886.aspx
[int]$visListDirTopToBottom = 2