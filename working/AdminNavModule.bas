Attribute VB_Name = "AdminNavModule"
'
'
'
Public Const IconAdvanced = "<img src=""/cclib/images/NavAdv.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconAdvancedClosed = "<img src=""/cclib/images/NavAdvClosed.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconAdvancedOpened = "<img src=""/cclib/images/NavAdvOpened.gif"" class=""ccImgA"" title=""{title}"">"
'Public Const IconAdvanced = "<img src=""/cclib/images/NavAdv.gif"" width=18 x height=18 border=0 style=""vertical-align:middle;"" title=""{title}"">"
'Public Const IconAdvancedClosed = "<img src=""/cclib/images/NavAdvClosed.gif"" width=18 x height=18 border=0 style=""vertical-align:middle;"" title=""{title}"">"
'Public Const IconAdvancedOpened = "<img src=""/cclib/images/NavAdvOpened.gif"" width=18 x height=18 border=0 style=""vertical-align:middle;"" title=""{title}"">"
'
Public Const IconContent = "<img src=""/cclib/images/NavContent.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconContentOpened = "<img src=""/cclib/images/NavContentOpened.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconContentClosed = "<img src=""/cclib/images/NavContentClosed.gif"" class=""ccImgA"" title=""{title}"">"
'
Public Const IconFolder = "<img src=""/cclib/images/NavFolderClosed.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconFolderClosed = "<img src=""/cclib/images/NavFolderClosed.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconFolderOpened = "<img src=""/cclib/images/NavFolderOpened.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconFolderNoSubNodes = "<img src=""/cclib/images/NavFolder.gif"" class=""ccImgA"" title=""{title}"">"
'
Public Const IconEmail = "<img src=""/cclib/images/NavEmail.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconEmailClosed = "<img src=""/cclib/images/NavEmailClosed.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconEmailOpened = "<img src=""/cclib/images/NavEmailOpened.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconSystemEmail = "<img src=""/cclib/images/NavEmail.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconConditionalEmail = "<img src=""/cclib/images/NavEmail.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconGroupEmail = "<img src=""/cclib/images/NavEmail.gif"" class=""ccImgA"" title=""{title}"">"
'
Public Const IconUsers = "<img src=""/cclib/images/NavUsers.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconUsersOpened = "<img src=""/cclib/images/NavUsersOpened.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconUsersClosed = "<img src=""/cclib/images/NavUsersClosed.gif"" class=""ccImgA"" title=""{title}"">"
'
Public Const IconReports = "<img src=""/cclib/images/NavReports.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconReportsOpened = "<img src=""/cclib/images/NavReportsOpened.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconReportsClosed = "<img src=""/cclib/images/NavReportsClosed.gif"" class=""ccImgA"" title=""{title}"">"
'
Public Const IconSettings = "<img src=""/cclib/images/NavSettings.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconSettingsOpened = "<img src=""/cclib/images/NavSettingsOpened.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconSettingsClosed = "<img src=""/cclib/images/NavSettingsClosed.gif"" class=""ccImgA"" title=""{title}"">"
'
Public Const IconTools = "<img src=""/cclib/images/NavTools.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconToolsOpened = "<img src=""/cclib/images/NavToolsOpened.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconToolsClosed = "<img src=""/cclib/images/NavToolsClosed.gif"" class=""ccImgA"" title=""{title}"">"
'
Public Const IconRecord = "<img src=""/cclib/images/NavRecord.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconRecordOpened = "<img src=""/cclib/images/NavRecord.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconRecordClosed = "<img src=""/cclib/images/NavRecord.gif"" class=""ccImgA"" title=""{title}"">"
'
Public Const IconHelp = "<img src=""/cclib/images/NavHelp.gif"" class=""ccImgA"" title=""{title}"">"
'
Public Const IconAddon = "<img src=""/cclib/images/NavAddons.gif"" class=""ccImgA"" title=""{title}"">"
'
Public Const IconAddons = "<img src=""/cclib/images/NavAddons.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconAddonsOpened = "<img src=""/cclib/images/NavAddonsOpened.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconAddonsClosed = "<img src=""/cclib/images/NavAddonsClosed.gif"" class=""ccImgA"" title=""{title}"">"
'
Public Const IconPublicHome = "<img src=""/cclib/images/NavPublicHome.gif"" class=""ccImgA"" title=""Public Home"">"
Public Const IconAdminHome = "<img src=""/cclib/images/NavHome.gif"" class=""ccImgA"" title=""Admin Home"">"
'
Public Const IconBox = "<img src=""/cclib/mktree/box.gif"" width=19 height=19 border=0 style=""vertical-align:middle;"" title=""{title}"">"
'
Public Const IconManageContent = "<img src=""/cclib/images/NavContent.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconManageContentOpened = "<img src=""/cclib/images/NavContentOpened.gif"" class=""ccImgA"" title=""{title}"">"
Public Const IconManageContentClosed = "<img src=""/cclib/images/NavContentClosed.gif"" class=""ccImgA"" title=""{title}"">"
'
Public Const NavigatorContentName = "Navigator Entries"
Public Const LegacyMenuContentName = "Menu Entries"
Public Const NodeIDAllContentList = "NodeIDAllContentList"
Public Const NodeIDAddonsNoCollection = "NodeIDAddonsNoCollection"
Public Const NodeIDLegacyMenu = "legacymenunode"
Public Const NodeIDManageAddons = "manageaddons"
Public Const NodeIDManageAddonsAdvanced = "manageaddonsadvanced"
Public Const NodeIDManageAddonsCollectionPrefix = "collection"
Public Const NodeIDSettings = "settings"
Public Const NodeIDTools = "tools"
Public Const NodeIDReports = "reports"
'
Public Type SortNodeType
    name As String
    addonid As Long
    ContentControlID As Long
    CollectionID As Long
    NavigatorID As Long
    NewWindow As Boolean
    ContentID As Long
    Link As String
    NavIconType As String
    NavIconTitle As String
    HelpAddonID As Long
    helpCollectionID As Long
    NodeIDString As String
End Type
'
Public Enum NodeTypeEnum
    NodeTypeEntry = 0
    NodeTypeCollection = 1
    NodeTypeAddon = 2
    NodeTypeContent = 3
End Enum

