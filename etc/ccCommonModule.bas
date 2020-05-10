Attribute VB_Name = "ccCommonModule"

Option Explicit
'
#Const DebugBuild = False
'
'=======================================================================
'   sitepropertyNames
'=======================================================================
'
Public Const siteproperty_serverPageDefault_name = "serverPageDefault"
Public Const siteproperty_serverPageDefault_defaultValue = "index.php"
'
'=======================================================================
'   content replacements
'=======================================================================
'
Public Const contentReplaceEscapeStart = "{%"
Public Const contentReplaceEscapeEnd = "%}"
'
Public Type fieldEditorType
    fieldId As Long
    addonid As Long
End Type
'
Public Const protectedContentSetControlFieldList = "ID,CREATEDBY,DATEADDED,MODIFIEDBY,MODIFIEDDATE,EDITSOURCEID,EDITARCHIVE,EDITBLANK,CONTENTCONTROLID"
'Public Const protectedContentSetControlFieldList = "ID,CREATEDBY,DATEADDED,MODIFIEDBY,MODIFIEDDATE,EDITSOURCEID,EDITARCHIVE,EDITBLANK"
'
Public Const HTMLEditorDefaultCopyStartMark = "<!-- cc -->"
Public Const HTMLEditorDefaultCopyEndMark = "<!-- /cc -->"
Public Const HTMLEditorDefaultCopyNoCr = HTMLEditorDefaultCopyStartMark & "<p><br></p>" & HTMLEditorDefaultCopyEndMark
Public Const HTMLEditorDefaultCopyNoCr2 = "<p><br></p>"
'
Public Const IconWidthHeight = " width=21 height=22 "
'Public Const IconWidthHeight = " width=18 height=22 "
'
Public Const CoreCollectionGuid = "{8DAABAE6-8E45-4CEE-A42C-B02D180E799B}" ' contains core Contensive objects, loaded from Library
Public Const ApplicationCollectionGuid = "{C58A76E2-248B-4DE8-BF9C-849A960F79C6}" ' exported from application during upgrade
'
Public Const adminCommonAddonGuid = "{76E7F79E-489F-4B0F-8EE5-0BAC3E4CD782}"
Public Const DashboardAddonGuid = "{4BA7B4A2-ED6C-46C5-9C7B-8CE251FC8FF5}"
Public Const PersonalizationGuid = "{C82CB8A6-D7B9-4288-97FF-934080F5FC9C}"
Public Const TextBoxGuid = "{7010002E-5371-41F7-9C77-0BBFF1F8B728}"
Public Const ContentBoxGuid = "{E341695F-C444-4E10-9295-9BEEC41874D8}"
Public Const DynamicMenuGuid = "{DB1821B3-F6E4-4766-A46E-48CA6C9E4C6E}"
Public Const ChildListGuid = "{D291F133-AB50-4640-9A9A-18DB68FF363B}"
Public Const DynamicFormGuid = "{8284FA0C-6C9D-43E1-9E57-8E9DD35D2DCC}"
Public Const AddonManagerGuid = "{1DC06F61-1837-419B-AF36-D5CC41E1C9FD}"
Public Const FormWizardGuid = "{2B1384C4-FD0E-4893-B3EA-11C48429382F}"
Public Const ImportWizardGuid = "{37F66F90-C0E0-4EAF-84B1-53E90A5B3B3F}"
Public Const JQueryGuid = "{9C882078-0DAC-48E3-AD4B-CF2AA230DF80}"
Public Const JQueryUIGuid = "{840B9AEF-9470-4599-BD47-7EC0C9298614}"
Public Const ImportProcessAddonGuid = "{5254FAC6-A7A6-4199-8599-0777CC014A13}"
Public Const StructuredDataProcessorGuid = "{65D58FE9-8B76-4490-A2BE-C863B372A6A4}"
Public Const jQueryFancyBoxGuid = "{24C2DBCF-3D84-44B6-A5F7-C2DE7EFCCE3D}"
'
Public Const DefaultLandingPageGuid = "{925F4A57-32F7-44D9-9027-A91EF966FB0D}"
Public Const DefaultLandingSectionGuid = "{D882ED77-DB8F-4183-B12C-F83BD616E2E1}"
Public Const DefaultTemplateGuid = "{47BE95E4-5D21-42CC-9193-A343241E2513}"
Public Const DefaultDynamicMenuGuid = "{E8D575B9-54AE-4BF9-93B7-C7E7FE6F2DB3}"
'
Public Const fpoContentBox = "{1571E62A-972A-4BFF-A161-5F6075720791}"
'
Public Const sfImageExtList = "jpg,jpeg,gif,png"
'
Public Const PageChildListInstanceID = "{ChildPageList}"
'
Public Const cr = vbCrLf & vbTab
Public Const cr2 = cr & vbTab
Public Const cr3 = cr2 & vbTab
Public Const cr4 = cr3 & vbTab
Public Const cr5 = cr4 & vbTab
Public Const cr6 = cr5 & vbTab
'
Public Const AddonOptionConstructor_BlockNoAjax = "Wrapper=[Default:0|None:-1|ListID(Wrappers)]" & vbCrLf & "css Container id" & vbCrLf & "css Container class"
Public Const AddonOptionConstructor_Block = "Wrapper=[Default:0|None:-1|ListID(Wrappers)]" & vbCrLf & "As Ajax=[If Add-on is Ajax:0|Yes:1]" & vbCrLf & "css Container id" & vbCrLf & "css Container class"
Public Const AddonOptionConstructor_Inline = "As Ajax=[If Add-on is Ajax:0|Yes:1]" & vbCrLf & "css Container id" & vbCrLf & "css Container class"
'
' Constants used as arguments to SiteBuilderClass.CreateNewSite
'
Public Const SiteTypeBaseAsp = 1
Public Const sitetypebaseaspx = 2
Public Const SiteTypeDemoAsp = 3
Public Const SiteTypeBasePhp = 4
'
'Public Const AddonNewParse = True
'
Public Const AddonOptionConstructor_ForBlockText = "AllowGroups=[listid(groups)]checkbox"
Public Const AddonOptionConstructor_ForBlockTextEnd = ""
Public Const BlockTextStartMarker = "<!-- BLOCKTEXTSTART -->"
Public Const BlockTextEndMarker = "<!-- BLOCKTEXTEND -->"
'
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'
Public Const InstallFolderName = "Install"
Public Const DownloadFileRootNode = "collectiondownload"
Public Const CollectionFileRootNode = "collection"
Public Const CollectionFileRootNodeOld = "addoncollection"
Public Const CollectionListRootNode = "collectionlist"
'
Public Const LegacyLandingPageName = "Landing Page Content"
Public Const DefaultNewLandingPageName = "Home"
Public Const DefaultLandingSectionName = "Home"
'
' ----- Errors Specific to the Contensive Objects
'
Public Const KmaccErrorUpgrading = KmaObjectError + 1
Public Const KmaccErrorServiceStopped = KmaObjectError + 2
'
Public Const UserErrorHeadline = "<p class=""ccError"">There was a problem with this page.</p>"
'
' ----- Errors connecting to server
'
Public Const ccError_InvalidAppName = 100
Public Const ccError_ErrorAddingApp = 101
Public Const ccError_ErrorDeletingApp = 102
Public Const ccError_InvalidFieldName = 103     ' Invalid parameter name
Public Const ccError_InvalidCommand = 104
Public Const ccError_InvalidAuthentication = 105
Public Const ccError_NotConnected = 106             ' Attempt to execute a command without a connection
'
'
'
Public Const ccStatusCode_Base = KmaErrorBase
Public Const ccStatusCode_ControllerCreateFailed = ccStatusCode_Base + 1
Public Const ccStatusCode_ControllerInProcess = ccStatusCode_Base + 2
Public Const ccStatusCode_ControllerStartedWithoutService = ccStatusCode_Base + 3
'
' ----- Previous errors, can be replaced
'
'Public Const KmaError_UnderlyingObject_Msg = "An error occurred in an underlying routine."
Public Const KmaccErrorServiceStopped_Msg = "The Contensive CSv Service is not running."
Public Const KmaError_BadObject_Msg = "Server Object is not valid."
Public Const KmaError_UpgradeInProgress_Msg = "Server is busy with internal upgrade."
'
'Public Const KmaError_InvalidArgument_Msg = "Invalid Argument"
'Public Const KmaError_UnderlyingObject_Msg = "An error occurred in an underlying routine."
'Public Const KmaccErrorServiceStopped_Msg = "The Contensive CSv Service is not running."
'Public Const KmaError_BadObject_Msg = "Server Object is not valid."
'Public Const KmaError_UpgradeInProgress_Msg = "Server is busy with internal upgrade."
'Public Const KmaError_InvalidArgument_Msg = "Invalid Argument"
'
'-----------------------------------------------------------------------
'   GetApplicationList indexes
'-----------------------------------------------------------------------
'
Public Const AppList_Name = 0
Public Const AppList_Status = 1
Public Const AppList_ConnectionsActive = 2
Public Const AppList_ConnectionString = 3
Public Const AppList_DataBuildVersion = 4
Public Const AppList_LicenseKey = 5
Public Const AppList_RootPath = 6
Public Const AppList_PhysicalFilePath = 7
Public Const AppList_DomainName = 8
Public Const AppList_DefaultPage = 9
Public Const AppList_AllowSiteMonitor = 10
Public Const AppList_HitCounter = 11
Public Const AppList_ErrorCount = 12
Public Const AppList_DateStarted = 13
Public Const AppList_AutoStart = 14
Public Const AppList_Progress = 15
Public Const AppList_PhysicalWWWPath = 16
Public Const AppListCount = 17
'
'-----------------------------------------------------------------------
'   System MemberID - when the system does an update, it uses this member
'-----------------------------------------------------------------------
'
Public Const SystemMemberID = 0
'
'-----------------------------------------------------------------------
' ----- old (OptionKeys for available Options)
'-----------------------------------------------------------------------
'
Public Const OptionKeyProductionLicense = 0
Public Const OptionKeyDeveloperLicense = 1
'
'-----------------------------------------------------------------------
' ----- LicenseTypes, replaced OptionKeys
'-----------------------------------------------------------------------
'
Public Const LicenseTypeInvalid = -1
Public Const LicenseTypeProduction = 0
Public Const LicenseTypeTrial = 1
'
'-----------------------------------------------------------------------
' ----- Active Content Definitions
'-----------------------------------------------------------------------
'
Public Const ACTypeDate = "DATE"
Public Const ACTypeVisit = "VISIT"
Public Const ACTypeVisitor = "VISITOR"
Public Const ACTypeMember = "MEMBER"
Public Const ACTypeOrganization = "ORGANIZATION"
Public Const ACTypeChildList = "CHILDLIST"
Public Const ACTypeContact = "CONTACT"
Public Const ACTypeFeedback = "FEEDBACK"
Public Const ACTypeLanguage = "LANGUAGE"
Public Const ACTypeAggregateFunction = "AGGREGATEFUNCTION"
Public Const ACTypeAddon = "ADDON"
Public Const ACTypeImage = "IMAGE"
Public Const ACTypeDownload = "DOWNLOAD"
Public Const ACTypeEnd = "END"
Public Const ACTypeTemplateContent = "CONTENT"
Public Const ACTypeTemplateText = "TEXT"
Public Const ACTypeDynamicMenu = "DYNAMICMENU"
Public Const ACTypeWatchList = "WATCHLIST"
Public Const ACTypeRSSLink = "RSSLINK"
Public Const ACTypePersonalization = "PERSONALIZATION"
Public Const ACTypeDynamicForm = "DYNAMICFORM"
'
Public Const ACTagEnd = "<ac type=""" & ACTypeEnd & """>"
'
' ----- PropertyType Definitions
'
Public Const PropertyTypeMember = 0
Public Const PropertyTypeVisit = 1
Public Const PropertyTypeVisitor = 2
'
'-----------------------------------------------------------------------
' ----- Port Assignments
'-----------------------------------------------------------------------
'
Public Const WinsockPortWebOut = 4000
Public Const WinsockPortServerFromWeb = 4001
Public Const WinsockPortServerToClient = 4002
'
Public Const Port_ContentServerControlDefault = 4531
Public Const Port_SiteMonitorDefault = 4532
'
Public Const RMBMethodHandShake = 1
Public Const RMBMethodMessage = 3
Public Const RMBMethodTestPoint = 4
Public Const RMBMethodInit = 5
Public Const RMBMethodClosePage = 6
Public Const RMBMethodOpenCSContent = 7
'
' ----- Position equates for the Remote Method Block
'
Const RMBPositionLength = 0             ' Length of the RMB
Const RMBPositionSourceHandle = 4       ' Handle generated by the source of the command
Const RMBPositionMethod = 8             ' Method in the method block
Const RMBPositionArgumentCount = 12     ' The number of arguments in the Block
Const RMBPositionFirstArgument = 16     ' The offset to the first argu
'
'-----------------------------------------------------------------------
'   Remote Connections
'   List of current remove connections for Remote Monitoring/administration
'-----------------------------------------------------------------------
'
Public Type RemoteAdministratorType
    RemoteIP As String
    RemotePort As Long
End Type
'
' Default username/password
'
Public Const DefaultServerUsername = "root"
Public Const DefaultServerPassword = "contensive"
'
'-----------------------------------------------------------------------
'   Form Contension Strategy
'
'       all Contensive Forms contain a hidden "ccFormSN"
'       The value in the hidden is the FormID string. All input
'       elements of the form are named FormID & "ElementName"
'
'       This prevents form elements from different forms from interfearing
'       with each other, and with developer generated forms.
'
'       GetFormSN gets a new valid random formid to be used.
'       All forms requires:
'           a FormId (text), containing the formid string
'           a [formid]Type (text), as defined in FormTypexxx in CommonModule
'
'       Forms have two primary sections: GetForm and ProcessForm
'
'       Any form that has a GetForm method, should have the process form
'       in the main.init, selected with this [formid]type hidden (not the
'       GetForm method). This is so the process can alter the stream
'       output for areas before the GetForm call.
'
'       System forms, like tools panel, that may appear on any page, have
'       their process call in the main.init.
'
'       Popup forms, like ImageSelector have their processform call in the
'       main.init because no .asp page exists that might contain a call
'       the process section.
'
'-----------------------------------------------------------------------
'
Public Const FormTypeToolsPanel = "do30a8vl29"
Public Const FormTypeActiveEditor = "l1gk70al9n"
Public Const FormTypeImageSelector = "ila9c5s01m"
Public Const FormTypePageAuthoring = "2s09lmpalb"
Public Const FormTypeMyProfile = "89aLi180j5"
Public Const FormTypeLogin = "login"
'Public Const FormTypeLogin = "l09H58a195"
Public Const FormTypeSendPassword = "lk0q56am09"
Public Const FormTypeJoin = "6df38abv00"
Public Const FormTypeHelpBubbleEditor = "9df019d77sA"
Public Const FormTypeAddonSettingsEditor = "4ed923aFGw9d"
Public Const FormTypeAddonStyleEditor = "ar5028jklkfd0s"
Public Const FormTypeSiteStyleEditor = "fjkq4w8794kdvse"
'Public Const FormTypeAggregateFunctionProperties = "9wI751270"
'
'-----------------------------------------------------------------------
'   Hardcoded profile form const
'-----------------------------------------------------------------------
'
Public Const rnMyProfileTopics = "profileTopics"
'
'-----------------------------------------------------------------------
' Legacy - replaced with HardCodedPages
'   Intercept Page Strategy
'
'       RequestnameInterceptpage = InterceptPage number from the input stream
'       InterceptPage = Global variant with RequestnameInterceptpage value read during early Init
'
'       Intercept pages are complete pages that appear instead of what
'       the physical page calls.
'-----------------------------------------------------------------------
'
Public Const RequestNameInterceptpage = "ccIPage"
'
Public Const LegacyInterceptPageSNResourceLibrary = "s033l8dm15"
Public Const LegacyInterceptPageSNSiteExplorer = "kdif3318sd"
Public Const LegacyInterceptPageSNImageUpload = "ka983lm039"
Public Const LegacyInterceptPageSNMyProfile = "k09ddk9105"
Public Const LegacyInterceptPageSNLogin = "6ge42an09a"
Public Const LegacyInterceptPageSNPrinterVersion = "l6d09a10sP"
Public Const LegacyInterceptPageSNUploadEditor = "k0hxp2aiOZ"
'
'-----------------------------------------------------------------------
' Ajax functions intercepted during init, answered and response closed
'   These are hard-coded internal Contensive functions
'   These should eventually be replaced with (HardcodedAddons) remote methods
'   They should all be prefixed "cc"
'   They are called with cj.ajax.qs(), setting RequestNameAjaxFunction=name in the qs
'   These name=value pairs go in the QueryString argument of the javascript cj.ajax.qs() function
'-----------------------------------------------------------------------
'
'Public Const RequestNameOpenSettingPage = "settingpageid"
Public Const RequestNameAjaxFunction = "ajaxfn"
Public Const RequestNameAjaxFastFunction = "ajaxfastfn"
'
Public Const AjaxOpenAdminNav = "aps89102kd"
Public Const AjaxOpenAdminNavGetContent = "d8475jkdmfj2"
Public Const AjaxCloseAdminNav = "3857fdjdskf91"
Public Const AjaxAdminNavOpenNode = "8395j2hf6jdjf"
Public Const AjaxAdminNavOpenNodeGetContent = "eieofdwl34efvclaeoi234598"
Public Const AjaxAdminNavCloseNode = "w325gfd73fhdf4rgcvjk2"
'
Public Const AjaxCloseIndexFilter = "k48smckdhorle0"
Public Const AjaxOpenIndexFilter = "Ls8jCDt87kpU45YH"
Public Const AjaxOpenIndexFilterGetContent = "llL98bbJQ38JC0KJm"
Public Const AjaxStyleEditorAddStyle = "ajaxstyleeditoradd"
Public Const AjaxPing = "ajaxalive"
Public Const AjaxGetFormEditTabContent = "ajaxgetformedittabcontent"
Public Const AjaxData = "data"
Public Const AjaxGetVisitProperty = "getvisitproperty"
Public Const AjaxSetVisitProperty = "setvisitproperty"
Public Const AjaxGetDefaultAddonOptionString = "ccGetDefaultAddonOptionString"
Public Const ajaxGetFieldEditorPreferenceForm = "ajaxgetfieldeditorpreference"
'
'-----------------------------------------------------------------------
'
' no - for now just use ajaxfn in the cj.ajax.qs call
'   this is more work, and I do not see why it buys anything new or better
'
'   Hard-coded addons
'       these are internal Contensive functions
'       can be called with just /addonname?querystring
'       call them with cj.ajax.addon() or cj.ajax.addonCallback()
'       are first in the list of checks when a URL rewrite is detected in Init()
'       should all be prefixed with 'cc'
'-----------------------------------------------------------------------
'
'Public Const HardcodedAddonGetDefaultAddonOptionString = "ccGetDefaultAddonOptionString"
'
'-----------------------------------------------------------------------
'   Remote Methods
'       ?RemoteMethodAddon=string
'       calls an addon (if marked to run as a remote method)
'       blocks all other Contensive output (tools panel, javascript, etc)
'-----------------------------------------------------------------------
'
Public Const RequestNameRemoteMethodAddon = "remotemethodaddon"
'
'-----------------------------------------------------------------------
'   Hard Coded Pages
'       ?Method=string
'       Querystring based so they can be added to URLs, preserving the current page for a return
'       replaces output stream with html output
'-----------------------------------------------------------------------
'
Public Const RequestNameHardCodedPage = "method"
'
Public Const HardCodedPageLogin = "login"
Public Const HardCodedPageLoginDefault = "logindefault"
Public Const HardCodedPageMyProfile = "myprofile"
Public Const HardCodedPagePrinterVersion = "printerversion"
Public Const HardCodedPageResourceLibrary = "resourcelibrary"
Public Const HardCodedPageLogoutLogin = "logoutlogin"
Public Const HardCodedPageLogout = "logout"
Public Const HardCodedPageSiteExplorer = "siteexplorer"
'Public Const HardCodedPageForceMobile = "forcemobile"
'Public Const HardCodedPageForceNonMobile = "forcenonmobile"
Public Const HardCodedPageNewOrder = "neworderpage"
Public Const HardCodedPageStatus = "status"
Public Const HardCodedPageGetJSPage = "getjspage"
Public Const HardCodedPageGetJSLogin = "getjslogin"
Public Const HardCodedPageRedirect = "redirect"
Public Const HardCodedPageExportAscii = "exportascii"
Public Const HardCodedPagePayPalConfirm = "paypalconfirm"
Public Const HardCodedPageSendPassword = "sendpassword"
'
'-----------------------------------------------------------------------
'   Option values
'       does not effect output directly
'-----------------------------------------------------------------------
'
Public Const RequestNamePageOptions = "ccoptions"
'
Public Const PageOptionForceMobile = "forcemobile"
Public Const PageOptionForceNonMobile = "forcenonmobile"
Public Const PageOptionLogout = "logout"
Public Const PageOptionPrinterVersion = "printerversion"
'
' convert to options later
'
Public Const RequestNameDashboardReset = "ResetDashboard"
'
'-----------------------------------------------------------------------
'   DataSource constants
'-----------------------------------------------------------------------
'
Public Const DefaultDataSourceID = -1
'
'-----------------------------------------------------------------------
' ----- Type compatibility between databases
'       Boolean
'           Access      YesNo       true=1, false=0
'           SQL Server  bit         true=1, false=0
'           MySQL       bit         true=1, false=0
'           Oracle      integer(1)  true=1, false=0
'           Note: false does not equal NOT true
'       Integer (Number)
'           Access      Long        8 bytes, about E308
'           SQL Server  int
'           MySQL       integer
'           Oracle      integer(8)
'       Float
'           Access      Double      8 bytes, about E308
'           SQL Server  Float
'           MySQL
'           Oracle
'       Text
'           Access
'           SQL Server
'           MySQL
'           Oracle
'-----------------------------------------------------------------------
'
'Public Const SQLFalse = "0"
'Public Const SQLTrue = "1"
'
'-----------------------------------------------------------------------
' ----- Style sheet definitions
'-----------------------------------------------------------------------
'
Public Const defaultStyleFilename = "ccDefault.r5.css"
Public Const StyleSheetStart = "<STYLE TYPE=""text/css"">"
Public Const StyleSheetEnd = "</STYLE>"
'
Public Const SpanClassAdminNormal = "<span class=""ccAdminNormal"">"
Public Const SpanClassAdminSmall = "<span class=""ccAdminSmall"">"
'
' remove these from ccWebx
'
Public Const SpanClassNormal = "<span class=""ccNormal"">"
Public Const SpanClassSmall = "<span class=""ccSmall"">"
Public Const SpanClassLarge = "<span class=""ccLarge"">"
Public Const SpanClassHeadline = "<span class=""ccHeadline"">"
Public Const SpanClassList = "<span class=""ccList"">"
Public Const SpanClassListCopy = "<span class=""ccListCopy"">"
Public Const SpanClassError = "<span class=""ccError"">"
Public Const SpanClassSeeAlso = "<span class=""ccSeeAlso"">"
Public Const SpanClassEnd = "</span>"
'
'-----------------------------------------------------------------------
' ----- XHTML definitions
'-----------------------------------------------------------------------
'
Public Const DTDTransitional = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"
'
Public Const BR = "<br>"
'
'-----------------------------------------------------------------------
' AuthoringControl Types
'-----------------------------------------------------------------------
'
Public Const AuthoringControlsEditing = 1
Public Const AuthoringControlsSubmitted = 2
Public Const AuthoringControlsApproved = 3
Public Const AuthoringControlsModified = 4
'
'-----------------------------------------------------------------------
' ----- Panel and header colors
'-----------------------------------------------------------------------
'
'Public Const "ccPanel" = "#E0E0E0"    ' The background color of a panel (black copy visible on it)
'Public Const "ccPanelHilite" = "#F8F8F8"  '
'Public Const "ccPanelShadow" = "#808080"  '
'
'Public Const HeaderColorBase = "#0320B0"   ' The background color of a panel header (reverse copy visible)
'Public Const "ccPanelHeaderHilite" = "#8080FF" '
'Public Const "ccPanelHeaderShadow" = "#000000" '
'
'-----------------------------------------------------------------------
' ----- Field type Definitions
'       Field Types are numeric values that describe how to treat values
'       stored as ContentFieldDefinitionType (FieldType property of FieldType Type.. ;)
'-----------------------------------------------------------------------
'
Public Const FieldTypeInteger = 1       ' An long number
Public Const FieldTypeText = 2          ' A text field (up to 255 characters)
Public Const FieldTypeLongText = 3      ' A memo field (up to 8000 characters)
Public Const FieldTypeBoolean = 4       ' A yes/no field
Public Const FieldTypeDate = 5          ' A date field
Public Const FieldTypeFile = 6          ' A filename of a file in the files directory.
Public Const FieldTypeLookup = 7        ' A lookup is a FieldTypeInteger that indexes into another table
Public Const FieldTypeRedirect = 8      ' creates a link to another section
Public Const FieldTypeCurrency = 9      ' A Float that prints in dollars
Public Const FieldTypeTextFile = 10     ' Text saved in a file in the files area.
Public Const FieldTypeImage = 11        ' A filename of a file in the files directory.
Public Const FieldTypeFloat = 12        ' A float number
Public Const FieldTypeAutoIncrement = 13 'long that automatically increments with the new record
Public Const FieldTypeManyToMany = 14    ' no database field - sets up a relationship through a Rule table to another table
Public Const FieldTypeMemberSelect = 15 ' This ID is a ccMembers record in a group defined by the MemberSelectGroupID field
Public Const FieldTypeCSSFile = 16      ' A filename of a CSS compatible file
Public Const FieldTypeXMLFile = 17      ' the filename of an XML compatible file
Public Const FieldTypeJavascriptFile = 18 ' the filename of a javascript compatible file
Public Const FieldTypeLink = 19           ' Links used in href tags -- can go to pages or resources
Public Const FieldTypeResourceLink = 20   ' Links used in resources, link <img or <object. Should not be pages
Public Const FieldTypeHTML = 21           ' LongText field that expects HTML content
Public Const FieldTypeHTMLFile = 22       ' TextFile field that expects HTML content
Public Const FieldTypeMax = 22
'
' ----- Field Descriptors for these type
'       These are what are publicly displayed for each type
'       See GetFieldDescriptorByType and vise-versa to translater
'
Public Const FieldDescriptorInteger = "Integer"
Public Const FieldDescriptorText = "Text"
Public Const FieldDescriptorLongText = "LongText"
Public Const FieldDescriptorBoolean = "Boolean"
Public Const FieldDescriptorDate = "Date"
Public Const FieldDescriptorFile = "File"
Public Const FieldDescriptorLookup = "Lookup"
Public Const FieldDescriptorRedirect = "Redirect"
Public Const FieldDescriptorCurrency = "Currency"
Public Const FieldDescriptorImage = "Image"
Public Const FieldDescriptorFloat = "Float"
Public Const FieldDescriptorManyToMany = "ManyToMany"
Public Const FieldDescriptorTextFile = "TextFile"
Public Const FieldDescriptorCSSFile = "CSSFile"
Public Const FieldDescriptorXMLFile = "XMLFile"
Public Const FieldDescriptorJavascriptFile = "JavascriptFile"
Public Const FieldDescriptorLink = "Link"
Public Const FieldDescriptorResourceLink = "ResourceLink"
Public Const FieldDescriptorMemberSelect = "MemberSelect"
Public Const FieldDescriptorHTML = "HTML"
Public Const FieldDescriptorHTMLFile = "HTMLFile"
'
Public Const FieldDescriptorLcaseInteger = "integer"
Public Const FieldDescriptorLcaseText = "text"
Public Const FieldDescriptorLcaseLongText = "longtext"
Public Const FieldDescriptorLcaseBoolean = "boolean"
Public Const FieldDescriptorLcaseDate = "date"
Public Const FieldDescriptorLcaseFile = "file"
Public Const FieldDescriptorLcaseLookup = "lookup"
Public Const FieldDescriptorLcaseRedirect = "redirect"
Public Const FieldDescriptorLcaseCurrency = "currency"
Public Const FieldDescriptorLcaseImage = "image"
Public Const FieldDescriptorLcaseFloat = "float"
Public Const FieldDescriptorLcaseManyToMany = "manytomany"
Public Const FieldDescriptorLcaseTextFile = "textfile"
Public Const FieldDescriptorLcaseCSSFile = "cssfile"
Public Const FieldDescriptorLcaseXMLFile = "xmlfile"
Public Const FieldDescriptorLcaseJavascriptFile = "javascriptfile"
Public Const FieldDescriptorLcaseLink = "link"
Public Const FieldDescriptorLcaseResourceLink = "resourcelink"
Public Const FieldDescriptorLcaseMemberSelect = "memberselect"
Public Const FieldDescriptorLcaseHTML = "html"
Public Const FieldDescriptorLcaseHTMLFile = "htmlfile"
'
'------------------------------------------------------------------------
' ----- Payment Options
'------------------------------------------------------------------------
'
Public Const PayTypeCreditCardOnline = 0   ' Pay by credit card online
Public Const PayTypeCreditCardByPhone = 1  ' Phone in a credit card
Public Const PayTypeCreditCardByFax = 9    ' Phone in a credit card
Public Const PayTypeCHECK = 2              ' pay by check to be mailed
Public Const PayTypeCREDIT = 3             ' pay on account
Public Const PayTypeNONE = 4               ' order total is $0.00. Nothing due
Public Const PayTypeCHECKCOMPANY = 5       ' pay by company check
Public Const PayTypeNetTerms = 6
Public Const PayTypeCODCompanyCheck = 7
Public Const PayTypeCODCertifiedFunds = 8
Public Const PayTypePAYPAL = 10
Public Const PAYDEFAULT = 0
'
'------------------------------------------------------------------------
' ----- Credit card options
'------------------------------------------------------------------------
'
Public Const CCTYPEVISA = 0                ' Visa
Public Const CCTYPEMC = 1                  ' MasterCard
Public Const CCTYPEAMEX = 2                ' American Express
Public Const CCTYPEDISCOVER = 3            ' Discover
Public Const CCTYPENOVUS = 4               ' Novus Card
Public Const CCTYPEDEFAULT = 0
'
'------------------------------------------------------------------------
' ----- Shipping Options
'------------------------------------------------------------------------
'
Public Const SHIPGROUND = 0                ' ground
Public Const SHIPOVERNIGHT = 1             ' overnight
Public Const SHIPSTANDARD = 2              ' standard, whatever that is
Public Const SHIPOVERSEAS = 3              ' to overseas
Public Const SHIPCANADA = 4                ' to Canada
Public Const SHIPDEFAULT = 0
'
'------------------------------------------------------------------------
' Debugging info
'------------------------------------------------------------------------
'
Public Const TestPointTab = 2
Public Const TestPointTabChr = "-"
Public CPTickCountBase As Double
'
'------------------------------------------------------------------------
'   project width button defintions
'------------------------------------------------------------------------
'
Public Const ButtonApply = "  Apply "
Public Const ButtonLogin = "  Login  "
Public Const ButtonLogout = "  Logout  "
Public Const ButtonSendPassword = "  Send Password  "
Public Const ButtonJoin = "   Join   "
Public Const ButtonSave = "  Save  "
Public Const ButtonOK = "     OK     "
Public Const ButtonReset = "  Reset  "
Public Const ButtonSaveAddNew = " Save + Add "
'Public Const ButtonSaveAddNew = " Save > Add "
Public Const ButtonCancel = " Cancel "
Public Const ButtonRestartContensiveApplication = " Restart Contensive Application "
Public Const ButtonCancelAll = "  Cancel  "
Public Const ButtonFind = "   Find   "
Public Const ButtonDelete = "  Delete  "
Public Const ButtonDeletePerson = " Delete Person "
Public Const ButtonDeleteRecord = " Delete Record "
Public Const ButtonDeleteEmail = " Delete Email "
Public Const ButtonDeletePage = " Delete Page "
Public Const ButtonFileChange = "   Upload   "
Public Const ButtonFileDelete = "    Delete    "
Public Const ButtonClose = "  Close   "
Public Const ButtonAdd = "   Add    "
Public Const ButtonAddChildPage = " Add Child "
Public Const ButtonAddSiblingPage = " Add Sibling "
Public Const ButtonContinue = " Continue >> "
Public Const ButtonBack = "  << Back  "
Public Const ButtonNext = "   Next   "
Public Const ButtonPrevious = " Previous "
Public Const ButtonFirst = "  First   "
Public Const ButtonSend = "  Send   "
Public Const ButtonSendTest = "Send Test"
Public Const ButtonCreateDuplicate = " Create Duplicate "
Public Const ButtonActivate = "  Activate   "
Public Const ButtonDeactivate = "  Deactivate   "
Public Const ButtonOpenActiveEditor = "Active Edit"
Public Const ButtonPublish = " Publish Changes "
Public Const ButtonAbortEdit = " Abort Edits "
Public Const ButtonPublishSubmit = " Submit for Publishing "
Public Const ButtonPublishApprove = " Approve for Publishing "
Public Const ButtonPublishDeny = " Deny for Publishing "
Public Const ButtonWorkflowPublishApproved = " Publish Approved Records "
Public Const ButtonWorkflowPublishSelected = " Publish Selected Records "
Public Const ButtonSetHTMLEdit = " Edit WYSIWYG "
Public Const ButtonSetTextEdit = " Edit HTML "
Public Const ButtonRefresh = " Refresh "
Public Const ButtonOrder = " Order "
Public Const ButtonSearch = " Search "
Public Const ButtonSpellCheck = " Spell Check "
Public Const ButtonLibraryUpload = " Upload "
Public Const ButtonCreateReport = " Create Report "
Public Const ButtonClearTrapLog = " Clear Trap Log "
Public Const ButtonNewSearch = " New Search "
Public Const ButtonReloadCDef = " Reload Content Definitions "
Public Const ButtonImportTemplates = " Import Templates "
Public Const ButtonRSSRefresh = " Update RSS Feeds Now "
Public Const ButtonRequestDownload = " Request Download "
Public Const ButtonFinish = " Finish "
Public Const ButtonRegister = " Register "
Public Const ButtonBegin = "Begin"
Public Const ButtonAbort = "Abort"
Public Const ButtonCreateGUID = " Create GUID "
Public Const ButtonEnable = " Enable "
Public Const ButtonDisable = " Disable "
Public Const ButtonMarkReviewed = " Mark Reviewed "
'
'------------------------------------------------------------------------
'   member actions
'------------------------------------------------------------------------
'
Public Const MemberActionNOP = 0
Public Const MemberActionLogin = 1
Public Const MemberActionLogout = 2
Public Const MemberActionForceLogin = 3
Public Const MemberActionSendPassword = 4
Public Const MemberActionForceLogout = 5
Public Const MemberActionToolsApply = 6
Public Const MemberActionJoin = 7
Public Const MemberActionSaveProfile = 8
Public Const MemberActionEditProfile = 9
'
'-----------------------------------------------------------------------
' ----- note pad info
'-----------------------------------------------------------------------
'
Public Const NoteFormList = 1
Public Const NoteFormRead = 2
'
Public Const NoteButtonPrevious = " Previous "
Public Const NoteButtonNext = "   Next   "
Public Const NoteButtonDelete = "  Delete  "
Public Const NoteButtonClose = "  Close   "
'                       ' Submit button is created in CommonDim, so it is simple
Public Const NoteButtonSubmit = "Submit"
'
'-----------------------------------------------------------------------
' ----- Admin site storage
'-----------------------------------------------------------------------
'
Public Const AdminMenuModeHidden = 0       '   menu is hidden
Public Const AdminMenuModeLeft = 1     '   menu on the left
Public Const AdminMenuModeTop = 2      '   menu as dropdowns from the top
'
' ----- AdminActions - describes the form processing to do
'
Public Const AdminActionNop = 0            ' do nothing
Public Const AdminActionDelete = 4         ' delete record
Public Const AdminActionFind = 5           '
Public Const AdminActionDeleteFilex = 6        '
Public Const AdminActionUpload = 7         '
Public Const AdminActionSaveNormal = 3         ' save fields to database
Public Const AdminActionSaveEmail = 8          ' save email record (and update EmailGroups) to database
Public Const AdminActionSaveMember = 11        '
Public Const AdminActionSaveSystem = 12
Public Const AdminActionSavePaths = 13     ' Save a record that is in the BathBlocking Format
Public Const AdminActionSendEmail = 9          '
Public Const AdminActionSendEmailTest = 10     '
Public Const AdminActionNext = 14               '
Public Const AdminActionPrevious = 15           '
Public Const AdminActionFirst = 16              '
Public Const AdminActionSaveContent = 17        '
Public Const AdminActionSaveField = 18          ' Save a single field, fieldname = fn input
Public Const AdminActionPublish = 19            ' Publish record live
Public Const AdminActionAbortEdit = 20          ' Publish record live
Public Const AdminActionPublishSubmit = 21      ' Submit for Workflow Publishing
Public Const AdminActionPublishApprove = 22     ' Approve for Workflow Publishing
Public Const AdminActionWorkflowPublishApproved = 23    ' Publish what was approved
Public Const AdminActionSetHTMLEdit = 24        ' Set Member Property for this field to HTML Edit
Public Const AdminActionSetTextEdit = 25        ' Set Member Property for this field to Text Edit
Public Const AdminActionSave = 26               ' Save Record
Public Const AdminActionActivateEmail = 27      ' Activate a Conditional Email
Public Const AdminActionDeactivateEmail = 28    ' Deactivate a conditional email
Public Const AdminActionDuplicate = 29          ' Duplicate the (sent email) record
Public Const AdminActionDeleteRows = 30         ' Delete from rows of records, row0 is boolean, rowid0 is ID, rowcnt is count
Public Const AdminActionSaveAddNew = 31         ' Save Record and add a new record
Public Const AdminActionReloadCDef = 32         ' Load Content Definitions
Public Const AdminActionWorkflowPublishSelected = 33 ' Publish what was selected
Public Const AdminActionMarkReviewed = 34       ' Mark the record reviewed without making any changes
Public Const AdminActionEditRefresh = 35        ' reload the page just like a save, but do not save
'
' ----- Adminforms (0-99)
'
Public Const AdminFormRoot = 0             ' intro page
Public Const AdminFormIndex = 1            ' record list page
Public Const AdminFormHelp = 2             ' popup help window
Public Const AdminFormUpload = 3           ' encoded file upload form
Public Const AdminFormEdit = 4             ' Edit form for system format records
Public Const AdminFormEditSystem = 5       ' Edit form for system format records
Public Const AdminFormEditNormal = 6       ' record edit page
Public Const AdminFormEditEmail = 7        ' Edit form for Email format records
Public Const AdminFormEditMember = 8       ' Edit form for Member format records
Public Const AdminFormEditPaths = 9        ' Edit form for Paths format records
Public Const AdminFormClose = 10           ' Special Case - do a window close instead of displaying a form
Public Const AdminFormReports = 12         ' Call Reports form (admin only)
'Public Const AdminFormSpider = 13          ' Call Spider
Public Const AdminFormEditContent = 14     ' Edit form for Content records
Public Const AdminFormDHTMLEdit = 15       ' ActiveX DHTMLEdit form
Public Const AdminFormEditPageContent = 16 '
Public Const AdminFormPublishing = 17       ' Workflow Authoring Publish Control form
Public Const AdminFormQuickStats = 18       ' Quick Stats (from Admin root)
Public Const AdminFormResourceLibrary = 19  ' Resource Library without Selects
Public Const AdminFormEDGControl = 20       ' Control Form for the EDG publishing controls
Public Const AdminFormSpiderControl = 21    ' Control Form for the Content Spider
Public Const AdminFormContentChildTool = 22 ' Admin Create Content Child tool
Public Const AdminformPageContentMap = 23   ' Map all content to a single map
Public Const AdminformHousekeepingControl = 24 ' Housekeeping control
Public Const AdminFormCommerceControl = 25
Public Const AdminFormContactManager = 26
Public Const AdminFormStyleEditor = 27
Public Const AdminFormEmailControl = 28
Public Const AdminFormCommerceInterface = 29
Public Const AdminFormDownloads = 30
Public Const AdminformRSSControl = 31
Public Const AdminFormMeetingSmart = 32
Public Const AdminFormMemberSmart = 33
Public Const AdminFormEmailWizard = 34
Public Const AdminFormImportWizard = 35
Public Const AdminFormCustomReports = 36
Public Const AdminFormFormWizard = 37
Public Const AdminFormLegacyAddonManager = 38
Public Const AdminFormIndex_SubFormAdvancedSearch = 39
Public Const AdminFormIndex_SubFormSetColumns = 40
Public Const AdminFormPageControl = 41
Public Const AdminFormSecurityControl = 42
Public Const AdminFormEditorConfig = 43
Public Const AdminFormBuilderCollection = 44
Public Const AdminFormClearCache = 45
Public Const AdminFormMobileBrowserControl = 46
Public Const AdminFormMetaKeywordTool = 47
Public Const AdminFormIndex_SubFormExport = 48
'
' ----- AdminFormTools (11,100-199)
'
Public Const AdminFormTools = 11           ' Call Tools form (developer only)
Public Const AdminFormToolRoot = 11         ' These should match for compatibility
Public Const AdminFormToolCreateContentDefinition = 101
Public Const AdminFormToolContentTest = 102
Public Const AdminFormToolConfigureMenu = 103
Public Const AdminFormToolConfigureListing = 104
Public Const AdminFormToolConfigureEdit = 105
Public Const AdminFormToolManualQuery = 106
Public Const AdminFormToolWriteUpdateMacro = 107
Public Const AdminFormToolDuplicateContent = 108
Public Const AdminFormToolDuplicateDataSource = 109
Public Const AdminFormToolDefineContentFieldsFromTable = 110
Public Const AdminFormToolContentDiagnostic = 111
Public Const AdminFormToolCreateChildContent = 112
Public Const AdminFormToolClearContentWatchLink = 113
Public Const AdminFormToolSyncTables = 114
Public Const AdminFormToolBenchmark = 115
Public Const AdminFormToolSchema = 116
Public Const AdminFormToolContentFileView = 117
Public Const AdminFormToolDbIndex = 118
Public Const AdminFormToolContentDbSchema = 119
Public Const AdminFormToolLogFileView = 120
Public Const AdminFormToolLoadCDef = 121
Public Const AdminFormToolLoadTemplates = 122
Public Const AdminformToolFindAndReplace = 123
Public Const AdminformToolCreateGUID = 124
Public Const AdminformToolIISReset = 125
Public Const AdminFormToolRestart = 126
Public Const AdminFormToolWebsiteFileView = 127
'
' ----- Define the index column structure
'       IndexColumnVariant( 0, n ) is the first column on the left
'       IndexColumnVariant( 0, IndexColumnField ) = the index into the fields array
'
Public Const IndexColumnField = 0          ' The field displayed in the column
Public Const IndexColumnWIDTH = 1          ' The width of the column
Public Const IndexColumnSORTPRIORITY = 2       ' lowest columns sorts first
Public Const IndexColumnSORTDIRECTION = 3      ' direction of the sort on this column
Public Const IndexColumnSATTRIBUTEMAX = 3      ' the number of attributes here
Public Const IndexColumnsMax = 50
'
' ----- ReportID Constants, moved to ccCommonModule
'
Public Const ReportFormRoot = 1
Public Const ReportFormDailyVisits = 2
Public Const ReportFormWeeklyVisits = 12
Public Const ReportFormSitePath = 4
Public Const ReportFormSearchKeywords = 5
Public Const ReportFormReferers = 6
Public Const ReportFormBrowserList = 8
Public Const ReportFormAddressList = 9
Public Const ReportFormContentProperties = 14
Public Const ReportFormSurveyList = 15
Public Const ReportFormOrdersList = 13
Public Const ReportFormOrderDetails = 21
Public Const ReportFormVisitorList = 11
Public Const ReportFormMemberDetails = 16
Public Const ReportFormPageList = 10
Public Const ReportFormVisitList = 3
Public Const ReportFormVisitDetails = 17
Public Const ReportFormVisitorDetails = 20
Public Const ReportFormSpiderDocList = 22
Public Const ReportFormSpiderErrorList = 23
Public Const ReportFormEDGDocErrors = 24
Public Const ReportFormDownloadLog = 25
Public Const ReportFormSpiderDocDetails = 26
Public Const ReportFormSurveyDetails = 27
Public Const ReportFormEmailDropList = 28
Public Const ReportFormPageTraffic = 29
Public Const ReportFormPagePerformance = 30
Public Const ReportFormEmailDropDetails = 31
Public Const ReportFormEmailOpenDetails = 32
Public Const ReportFormEmailClickDetails = 33
Public Const ReportFormGroupList = 34
Public Const ReportFormGroupMemberList = 35
Public Const ReportFormTrapList = 36
Public Const ReportFormCount = 36
'
'=============================================================================
' Page Scope Meetings Related Storage
'=============================================================================
'
Public Const MeetingFormIndex = 0
Public Const MeetingFormAttendees = 1
Public Const MeetingFormLinks = 2
Public Const MeetingFormFacility = 3
Public Const MeetingFormHotel = 4
Public Const MeetingFormDetails = 5
'
'------------------------------------------------------------------------------
' Form actions
'------------------------------------------------------------------------------
'
' ----- DataSource Types
'
Public Const DataSourceTypeODBCSQL99 = 0
Public Const DataSourceTypeODBCAccess = 1
Public Const DataSourceTypeODBCSQLServer = 2
Public Const DataSourceTypeODBCMySQL = 3
Public Const DataSourceTypeXMLFile = 4      ' Use MSXML Interface to open a file
'
'------------------------------------------------------------------------------
'   Application Status
'------------------------------------------------------------------------------
'
Public Const ApplicationStatusNotFound = 0
Public Const ApplicationStatusLoadedNotRunning = 1
Public Const ApplicationStatusRunning = 2
Public Const ApplicationStatusStarting = 3
Public Const ApplicationStatusUpgrading = 4
' Public Const ApplicationStatusConnectionBusy = 5    ' can not open connection because already open
Public Const ApplicationStatusKernelFailure = 6     ' can not create Kernel
Public Const ApplicationStatusNoHostService = 7     ' host service process ID not set
Public Const ApplicationStatusLicenseFailure = 8    ' failed to start because of License failure
Public Const ApplicationStatusDbFailure = 9         ' failed to start because ccSetup table not found
Public Const ApplicationStatusUnknownFailure = 10   ' failed to start because of unknown error, see trace log
Public Const ApplicationStatusDbBad = 11            ' ccContent,ccFields no records found
Public Const ApplicationStatusConnectionObjectFailure = 12 ' Connection Object FAiled
Public Const ApplicationStatusConnectionStringFailure = 13 ' Connection String FAiled to open the ODBC connection
Public Const ApplicationStatusDataSourceFailure = 14 ' DataSource failed to open
Public Const ApplicationStatusDuplicateDomains = 15 ' Can not locate application because there are 1+ apps that match the domain
Public Const ApplicationStatusPaused = 16           ' Running, but all activity is blocked (for backup)
'
' Document (HTML, graphic, etc) retrieved from site
'
Public Const ResponseHeaderCountMax = 20
Public Const ResponseCookieCountMax = 20
'
' ----- text delimiter that divides the text and html parts of an email message stored in the queue folder
'
Public Const EmailTextHTMLDelimiter = vbCrLf & " ----- End Text Begin HTML -----" & vbCrLf
'
'------------------------------------------------------------------------
'   Common RequestName Variables
'------------------------------------------------------------------------
'
Public Const RequestNameDynamicFormID = "dformid"
'
Public Const RequestNameRunAddon = "addonid"
Public Const RequestNameEditReferer = "EditReferer"
Public Const RequestNameRefreshBlock = "ccFormRefreshBlockSN"
Public Const RequestNameCatalogOrder = "CatalogOrderID"
Public Const RequestNameCatalogCategoryID = "CatalogCatID"
Public Const RequestNameCatalogForm = "CatalogFormID"
Public Const RequestNameCatalogItemID = "CatalogItemID"
Public Const RequestNameCatalogItemAge = "CatalogItemAge"
Public Const RequestNameCatalogRecordTop = "CatalogTop"
Public Const RequestNameCatalogFeatured = "CatalogFeatured"
Public Const RequestNameCatalogSpan = "CatalogSpan"
Public Const RequestNameCatalogKeywords = "CatalogKeywords"
Public Const RequestNameCatalogSource = "CatalogSource"
'
Public Const RequestNameLibraryFileID = "fileEID"
Public Const RequestNameDownloadID = "downloadid"
Public Const RequestNameLibraryUpload = "LibraryUpload"
Public Const RequestNameLibraryName = "LibraryName"
Public Const RequestNameLibraryDescription = "LibraryDescription"

Public Const RequestNameTestHook = "CC"       ' input request that sets debugging hooks

Public Const RequestNameRootPage = "RootPageName"
Public Const RequestNameRootPageID = "RootPageID"
Public Const RequestNameContent = "ContentName"
Public Const RequestNameOrderByClause = "OrderByClause"
Public Const RequestNameAllowChildPageList = "AllowChildPageList"
'
Public Const RequestNameCRKey = "crkey"
Public Const RequestNameAdminForm = "af"
Public Const RequestNameAdminSubForm = "subform"
Public Const RequestNameButton = "button"
Public Const RequestNameAdminSourceForm = "asf"
Public Const RequestNameAdminFormSpelling = "SpellingRequest"
Public Const RequestNameInlineStyles = "InlineStyles"
Public Const RequestNameAllowCSSReset = "AllowCSSReset"
'
Public Const RequestNameReportForm = "rid"
'
Public Const RequestNameToolContentID = "ContentID"
'
Public Const RequestNameCut = "a904o2pa0cut"
Public Const RequestNamePaste = "dp29a7dsa6paste"
Public Const RequestNamePasteParentContentID = "dp29a7dsa6cid"
Public Const RequestNamePasteParentRecordID = "dp29a7dsa6rid"
Public Const RequestNamePasteFieldList = "dp29a7dsa6key"
Public Const RequestNameCutClear = "dp29a7dsa6clear"
'
Public Const RequestNameRequestBinary = "RequestBinary"
' removed -- this was an old method of blocking form input for file uploads
'Public Const RequestNameFormBlock = "RB"
Public Const RequestNameJSForm = "RequestJSForm"
Public Const RequestNameJSProcess = "ProcessJSForm"
'
Public Const RequestNameFolderID = "FolderID"
'
Public Const RequestNameEmailMemberID = "emi8s9Kj"
Public Const RequestNameEmailOpenFlag = "eof9as88"
Public Const RequestNameEmailOpenCssFlag = "8aa41pM3"
Public Const RequestNameEmailClickFlag = "ecf34Msi"
Public Const RequestNameEmailSpamFlag = "9dq8Nh61"
Public Const RequestNameEmailBlockRequestDropID = "BlockEmailRequest"
Public Const RequestNameVisitTracking = "s9lD1088"
Public Const RequestNameBlockContentTracking = "BlockContentTracking"
Public Const RequestNameCookieDetectVisitID = "f92vo2a8d"

Public Const RequestNamePageNumber = "PageNumber"
Public Const RequestNamePageSize = "PageSize"
'
Public Const RequestValueNull = "[NULL]"
'
Public Const SpellCheckUserDictionaryFilename = "SpellCheck\UserDictionary.txt"
'
Public Const RequestNameStateString = "vstate"
'
'------------------------------------------------------------------------------
' name value pairs
'------------------------------------------------------------------------------
'
Public Type NameValuePairType
    Name As String
    Value As String
    End Type
''
'' ----- ContentSetMirror Type
''       Used on the WebClient, not the CSv
''       Stores info about the ContentSet, and caches the current row
''
'Public Type ContentSetMirrorType
'    Open As Boolean                     ' If true, it is in use
'    Updateable As Boolean               ' Can not update an OpenCSSQL because Fields are not accessable
'    ContentName As String               ' If updateable, this is the contentname
'    CSPointer As Long                ' CSPointer for this ContentSet
'    '
'    ' ----- a cache of the current row, passed in during open and nextrecord, back during save and nextrecord
'    '
'    EOF As Boolean                      ' if true, Row is empty and at end of records
'    RowCache() As ContentSetRowCacheType ' array of fields buffered for this set
'    RowCacheSize As Long             ' the total number of fields in the row
'    RowCacheCount As Long            ' the number of field() values to write
'    End Type
'
' ----- Dataset for graphing
'
Public Type ColumnDataType
    Name As String
    row() As Long
    End Type
'
Public Type ChartDataType
    Title As String
    XLabel As String
    YLabel As String
    RowCount As Long
    RowLabel() As String
    ColumnCount As Long
    Column() As ColumnDataType
    End Type
''
' PrivateStorage to hold the DebugTimer
'
Type TimerStackType
    Label As String
    StartTicks As Long
    End Type
Private Const TimerStackMax = 20
Private TimerStack(TimerStackMax) As TimerStackType
Private TimerStackCount As Long
'
Public Const TextSearchStartTagDefault = "<!--TextSearchStart-->"
Public Const TextSearchEndTagDefault = "<!--TextSearchEnd-->"
'
'-------------------------------------------------------------------------------------
'   IPDaemon communication objects
'-------------------------------------------------------------------------------------
'
Type IPDaemonConnectionType
    ConnectionID As Integer
    BytesToSend As Long
    HTTPVersion As String
    HTTPMethod As String
    Path As String
    Query As String
    Headers As String
    PostData As String
    SendData As Boolean
    State As Integer
    ContentLength As Integer
    End Type

Global IPDaemonConnection() As IPDaemonConnectionType

Global Const IPDaemon_DISCONNECTED = 0
Global Const IPDaemon_CONNECTED = 1
Global Const IPDaemon_HEADERS = 2
Global Const IPDaemon_POSTDATA = 3
Global Const IPDaemon_SERVE = 4
Global Const IPDaemon_SERVEDIR = 5
Global Const IPDaemon_SERVEFILE = 6
'
'-------------------------------------------------------------------------------------
'   Email
'-------------------------------------------------------------------------------------
'
Public Const EmailLogTypeDrop = 1                   ' Email was dropped
Public Const EmailLogTypeOpen = 2                   ' System detected the email was opened
Public Const EmailLogTypeClick = 3                  ' System detected a click from a link on the email
Public Const EmailLogTypeBounce = 4                 ' Email was processed by bounce processing
Public Const EmailLogTypeBlockRequest = 5           ' recipient asked us to stop sending email
Public Const EmailLogTypeImmediateSend = 6        ' Email was dropped
'
Public Const DefaultSpamFooter = "<p>To block future emails from this site, <link>click here</link></p>"
'
Public Const FeedbackFormNotSupportedComment = "<!--" & vbCrLf & "Feedback form is not supported in this context" & vbCrLf & "-->"
'
'-------------------------------------------------------------------------------------
'   Page Content constants
'-------------------------------------------------------------------------------------
'
Public Const ContentBlockCopyName = "Content Block Copy"
'
Public Const BubbleCopy_AdminAddPage = "Use the Add page to create new content records. The save button puts you in edit mode. The OK button creates the record and exits."
Public Const BubbleCopy_AdminIndexPage = "Use the Admin Listing page to locate content records through the Admin Site."
Public Const BubbleCopy_SpellCheckPage = "Use the Spell Check page to verify and correct spelling throught the content."
Public Const BubbleCopy_AdminEditPage = "Use the Edit page to add and modify content."
'
'
Public Const TemplateDefaultName = "Default"
'Public Const TemplateDefaultBody = "<!--" & vbCrLf & "Default Template - edit this Page Template, or select a different template for your page or section" & vbCrLf & "-->{{DYNAMICMENU?MENU=}}<br>{{CONTENT}}"
Public Const TemplateDefaultBody = "" _
    & vbCrLf & vbTab & "<!--" _
    & vbCrLf & vbTab & "Default Template - edit this Page Template, or select a different template for your page or section" _
    & vbCrLf & vbTab & "-->" _
    & vbCrLf & vbTab & "<ac type=""AGGREGATEFUNCTION"" name=""Dynamic Menu"" querystring=""Menu Name=Default"" acinstanceid=""{6CBADABB-5B0D-43E1-B3CA-46A3D60DA3E1}"" >" _
    & vbCrLf & vbTab & "<ac type=""AGGREGATEFUNCTION"" name=""Content Box"" acinstanceid=""{49E0D0C0-D323-49B6-B211-B9599673A265}"" >"
Public Const TemplateDefaultBodyTag = "<body class=""ccBodyWeb"">"
'
'=======================================================================
'   Internal Tab interface storage
'=======================================================================
'
Private Type TabType
    Caption As String
    Link As String
    StylePrefix As String
    IsHit As Boolean
    LiveBody As String
End Type
Private Tabs() As TabType
Private TabsCnt As Long
Private TabsSize As Long
'
' Admin Navigator Nodes
'
Public Const NavigatorNodeCollectionList = -1
Public Const NavigatorNodeAddonList = -1
'
' Pointers into index of PCC (Page Content Cache) array
'
Public Const PCC_ID = 0
Public Const PCC_Active = 1
Public Const PCC_ParentID = 2
Public Const PCC_Name = 3
Public Const PCC_Headline = 4
Public Const PCC_MenuHeadline = 5
Public Const PCC_DateArchive = 6
Public Const PCC_DateExpires = 7
Public Const PCC_PubDate = 8
Public Const PCC_ChildListSortMethodID = 9
Public Const PCC_ContentControlID = 10
Public Const PCC_TemplateID = 11
Public Const PCC_BlockContent = 12
Public Const PCC_BlockPage = 13
Public Const PCC_Link = 14
Public Const PCC_RegistrationGroupID = 15
Public Const PCC_BlockSourceID = 16
Public Const PCC_CustomBlockMessageFilename = 17
Public Const PCC_JSOnLoad = 18
Public Const PCC_JSHead = 19
Public Const PCC_JSEndBody = 20
Public Const PCC_Viewings = 21
Public Const PCC_ContactMemberID = 22
Public Const PCC_AllowHitNotification = 23
Public Const PCC_TriggerSendSystemEmailID = 24
Public Const PCC_TriggerConditionID = 25
Public Const PCC_TriggerConditionGroupID = 26
Public Const PCC_TriggerAddGroupID = 27
Public Const PCC_TriggerRemoveGroupID = 28
Public Const PCC_AllowMetaContentNoFollow = 29
Public Const PCC_ParentListName = 30
Public Const PCC_CopyFilename = 31
Public Const PCC_BriefFilename = 32
Public Const PCC_AllowChildListDisplay = 33
Public Const PCC_SortOrder = 34
Public Const PCC_DateAdded = 35
Public Const PCC_ModifiedDate = 36
Public Const PCC_ChildPagesFound = 37
Public Const PCC_AllowInMenus = 38
Public Const PCC_AllowInChildLists = 39
Public Const PCC_JSFilename = 40
Public Const PCC_ChildListInstanceOptions = 41
Public Const PCC_IsSecure = 42
Public Const PCC_AllowBrief = 43
Public Const PCC_ColCnt = 44
'
' Indexes into the SiteSectionCache
' Created from "ID, Name,TemplateID,ContentID,MenuImageFilename,Caption,MenuImageOverFilename,HideMenu,BlockSection,RootPageID,JSOnLoad,JSHead,JSEndBody"
'
Public Const SSC_ID = 0
Public Const SSC_Name = 1
Public Const SSC_TemplateID = 2
Public Const SSC_ContentID = 3
Public Const SSC_MenuImageFilename = 4
Public Const SSC_Caption = 5
Public Const SSC_MenuImageOverFilename = 6
Public Const SSC_HideMenu = 7
Public Const SSC_BlockSection = 8
Public Const SSC_RootPageID = 9
Public Const SSC_JSOnLoad = 10
Public Const SSC_JSHead = 11
Public Const SSC_JSEndBody = 12
Public Const SSC_JSFilename = 13
Public Const SSC_cnt = 14
'
' Indexes into the TemplateCache
' Created from "t.ID,t.Name,t.Link,t.BodyHTML,t.JSOnLoad,t.JSHead,t.JSEndBody,t.StylesFilename,r.StyleID"
'
Public Const TC_ID = 0
Public Const TC_Name = 1
Public Const TC_Link = 2
Public Const TC_BodyHTML = 3
Public Const TC_JSOnLoad = 4
Public Const TC_JSInHeadLegacy = 5
'Public Const TC_JSHead = 5
Public Const TC_JSEndBody = 6
Public Const TC_StylesFilename = 7
Public Const TC_SharedStylesIDList = 8
Public Const TC_MobileBodyHTML = 9
Public Const TC_MobileStylesFilename = 10
Public Const TC_OtherHeadTags = 11
Public Const TC_BodyTag = 12
Public Const TC_JSInHeadFilename = 13
'Public Const TC_JSFilename = 13
Public Const TC_IsSecure = 14
Public Const TC_DomainIdList = 15
' for now, Mobile templates do not have shared styles
'Public Const TC_MobileSharedStylesIDList = 11
Public Const TC_cnt = 16
'
' DTD
'
Public Const DTDDefault = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3.org/TR/html4/loose.dtd"">"
Public Const DTDDefaultAdmin = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3.org/TR/html4/loose.dtd"">"
'
' innova Editor feature list
'
Public Const InnovaEditorFeaturefilename = "Config\EditorCongif.txt"
Public Const InnovaEditorFeatureList = "FullScreen,Preview,Print,Search,Cut,Copy,Paste,PasteWord,PasteText,SpellCheck,Undo,Redo,Image,Flash,Media,CustomObject,CustomTag,Bookmark,Hyperlink,HTMLSource,XHTMLSource,Numbering,Bullets,Indent,Outdent,JustifyLeft,JustifyCenter,JustifyRight,JustifyFull,Table,Guidelines,Absolute,Characters,Line,Form,RemoveFormat,ClearAll,StyleAndFormatting,TextFormatting,ListFormatting,BoxFormatting,ParagraphFormatting,CssText,Styles,Paragraph,FontName,FontSize,Bold,Italic,Underline,Strikethrough,Superscript,Subscript,ForeColor,BackColor"
Public Const InnovaEditorPublicFeatureList = "FullScreen,Preview,Print,Search,Cut,Copy,Paste,PasteWord,PasteText,SpellCheck,Undo,Redo,Bookmark,Hyperlink,HTMLSource,XHTMLSource,Numbering,Bullets,Indent,Outdent,JustifyLeft,JustifyCenter,JustifyRight,JustifyFull,Table,Guidelines,Absolute,Characters,Line,Form,RemoveFormat,ClearAll,StyleAndFormatting,TextFormatting,ListFormatting,BoxFormatting,ParagraphFormatting,CssText,Styles,Paragraph,FontName,FontSize,Bold,Italic,Underline,Strikethrough,Superscript,Subscript,ForeColor,BackColor"
''
'' Content Type
''
'Enum contentTypeEnum
'    contentTypeWeb = 1
'    ContentTypeEmail = 2
'    contentTypeWebTemplate = 3
'    contentTypeEmailTemplate = 4
'End Enum
'Public EditorContext As contentTypeEnum
'Enum EditorContextEnum
'    contentTypeWeb = 1
'    contentTypeEmail = 2
'End Enum
'Public EditorContext As EditorContextEnum
''
'Public Const EditorAddonMenuEmailTemplateFilename = "templates/EditorAddonMenuTemplateEmail.js"
'Public Const EditorAddonMenuEmailContentFilename = "templates/EditorAddonMenuContentEmail.js"
'Public Const EditorAddonMenuWebTemplateFilename = "templates/EditorAddonMenuTemplateWeb.js"
'Public Const EditorAddonMenuWebContentFilename = "templates/EditorAddonMenuContentWeb.js"
'
Public Const DynamicStylesFilename = "templates/styles.css"
Public Const AdminSiteStylesFilename = "templates/AdminSiteStyles.css"
Public Const EditorStyleRulesFilenamePattern = "templates/EditorStyleRules$TemplateID$.js"
' deprecated 11/24/3009 - StyleRules destinction between web/email not needed b/c body background blocked
'Public Const EditorStyleWebRulesFilename = "templates/EditorStyleWebRules.js"
'Public Const EditorStyleEmailRulesFilename = "templates/EditorStyleEmailRules.js"
'
' ----- ccGroupRules storage for list of Content that a group can author
'
Public Type ContentGroupRuleType
    ContentID As Long
    GroupID As Long
    AllowAdd As Boolean
    AllowDelete As Boolean
End Type
'
' ----- This should match the Lookup List in the NavIconType field in the Navigator Entry content definition
'
Public Const navTypeIDList = "Add-on,Report,Setting,Tool"
Public Const NavTypeIDAddon = 1
Public Const NavTypeIDReport = 2
Public Const NavTypeIDSetting = 3
Public Const NavTypeIDTool = 4
'
Public Const NavIconTypeList = "Custom,Advanced,Content,Folder,Email,User,Report,Setting,Tool,Record,Addon,help"
Public Const NavIconTypeCustom = 1
Public Const NavIconTypeAdvanced = 2
Public Const NavIconTypeContent = 3
Public Const NavIconTypeFolder = 4
Public Const NavIconTypeEmail = 5
Public Const NavIconTypeUser = 6
Public Const NavIconTypeReport = 7
Public Const NavIconTypeSetting = 8
Public Const NavIconTypeTool = 9
Public Const NavIconTypeRecord = 10
Public Const NavIconTypeAddon = 11
Public Const NavIconTypeHelp = 12
'
Public Const QueryTypeSQL = 1
Public Const QueryTypeOpenContent = 2
Public Const QueryTypeUpdateContent = 3
Public Const QueryTypeInsertContent = 4
'
' Google Data Object construction in GetRemoteQuery
'
Public Type ColsType
    Type As String
    Id As String
    Label As String
    Pattern As String
End Type
'
Public Type CellType
    v As String
    f As String
    p As String
End Type
'
Public Type RowsType
    Cell() As CellType
End Type
'
Public Type GoogleDataType
    IsEmpty As Boolean
    col() As ColsType
    row() As RowsType
End Type
'
Public Enum GoogleVisualizationStatusEnum
    OK = 1
    warning = 2
    Error = 3
End Enum
'
Public Type GoogleVisualizationType
    version As String
    reqid As String
    status As GoogleVisualizationStatusEnum
    warnings() As String
    errors() As String
    sig As String
    table As GoogleDataType
End Type

'Public Const ReturnFormatTypeGoogleTable = 1
'Public Const ReturnFormatTypeNameValue = 2

Public Enum RemoteFormatEnum
    RemoteFormatJsonTable = 1
    RemoteFormatJsonNameArray = 2
    RemoteFormatJsonNameValue = 3
End Enum
'
'
'
Public Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey&)
Public Declare Function RegOpenKeyExA& Lib "advapi32.dll" (ByVal hKey&, ByVal lpszSubKey$, dwOptions&, ByVal samDesired&, lpHKey&)
Public Declare Function RegQueryValueExA& Lib "advapi32.dll" (ByVal hKey&, ByVal lpszValueName$, ByVal lpdwRes&, lpdwType&, ByVal lpDataBuff$, nSize&)
Public Declare Function RegQueryValueEx& Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey&, ByVal lpszValueName$, ByVal lpdwRes&, lpdwType&, lpDataBuff&, nSize&)

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1&                          ' Unicode nul terminated string
Public Const REG_DWORD = 4&                       ' 32-bit number

Public Const KEY_QUERY_VALUE = &H1&
Public Const KEY_SET_VALUE = &H2&
Public Const KEY_CREATE_SUB_KEY = &H4&
Public Const KEY_ENUMERATE_SUB_KEYS = &H8&
Public Const KEY_NOTIFY = &H10&
Public Const KEY_CREATE_LINK = &H20&
Public Const READ_CONTROL = &H20000
Public Const WRITE_DAC = &H40000
Public Const WRITE_OWNER = &H80000
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const STANDARD_RIGHTS_READ = READ_CONTROL
Public Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Public Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Public Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Public Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Public Const KEY_EXECUTE = KEY_READ

'======================================================================================
'
'======================================================================================
'
Public Sub StartDebugTimer(Enabled As Boolean, Label As String)
    ' ##### removed to catch err<>0 problem on error resume next
    If Enabled Then
        If TimerStackCount < TimerStackMax Then
            TimerStack(TimerStackCount).Label = Label
            TimerStack(TimerStackCount).StartTicks = GetTickCount
        Else
            Call AppendLogFile(App.EXEName & ".?.StartDebugTimer, " & "Timer Stack overflow, attempting push # [" & TimerStackCount & "], but max = [" & TimerStackMax & "]")
            End If
        TimerStackCount = TimerStackCount + 1
        End If
    End Sub
'
Public Sub StopDebugTimer(Enabled As Boolean, Label As String)
    ' ##### removed to catch err<>0 problem on error resume next
    If Enabled Then
        If TimerStackCount <= 0 Then
            Call AppendLogFile(App.EXEName & ".?.StopDebugTimer, " & "Timer Error, attempting to Pop, but the stack is empty")
        Else
            If TimerStackCount <= TimerStackMax Then
                If TimerStack(TimerStackCount - 1).Label = Label Then
                    Call AppendLogFile(App.EXEName & ".?.StopDebugTimer, " & "Timer [" & String(2 * TimerStackCount, ".") & Label & "] took " & (GetTickCount - TimerStack(TimerStackCount - 1).StartTicks) & " msec")
                Else
                    Call AppendLogFile(App.EXEName & ".?.StopDebugTimer, " & "Timer Error, [" & Label & "] was popped, but [" & TimerStack(TimerStackCount).Label & "] was on the top of the stack")
                    End If
                End If
            TimerStackCount = TimerStackCount - 1
            End If
        End If
    End Sub
'
'
'
Public Function PayString(Index) As String
    ' ##### removed to catch err<>0 problem on error resume next
    Select Case Index
        Case PayTypeCreditCardOnline
            PayString = "Credit Card"
        Case PayTypeCreditCardByPhone
            PayString = "Credit Card by phone"
        Case PayTypeCreditCardByFax
            PayString = "Credit Card by fax"
        Case PayTypeCHECK
            PayString = "Personal Check"
        Case PayTypeCHECKCOMPANY
            PayString = "Company Check"
        Case PayTypeCREDIT
            PayString = "You will be billed"
        Case PayTypeNetTerms
            PayString = "Net Terms (Approved customers only)"
        Case PayTypeCODCompanyCheck
            PayString = "COD- Pre-Approved Only"
        Case PayTypeCODCertifiedFunds
            PayString = "COD- Certified Funds"
        Case PayTypePAYPAL
            PayString = "PayPal"
        Case Else
            ' Case PayTypeNONE
            PayString = "No payment required"
        End Select
    End Function
'
'
'
Public Function CCString(Index) As String
    ' ##### removed to catch err<>0 problem on error resume next
    Select Case Index
        Case CCTYPEVISA
            CCString = "Visa"
        Case CCTYPEMC
            CCString = "MasterCard"
        Case CCTYPEAMEX
            CCString = "American Express"
        Case CCTYPEDISCOVER
            CCString = "Discover"
        Case Else
            ' Case CCTYPENOVUS
            CCString = "Novus Card"
        End Select
    End Function
'
'========================================================================
' Get a Long from a CommandPacket
'   position+0, 4 byte value
'========================================================================
'
Public Function GetLongFromByteArray(ByteArray() As Byte, Position As Long) As Long
    ' ##### removed to catch err<>0 problem on error resume next
    '
    GetLongFromByteArray = ByteArray(Position + 3)
    GetLongFromByteArray = ByteArray(Position + 2) + (256 * GetLongFromByteArray)
    GetLongFromByteArray = ByteArray(Position + 1) + (256 * GetLongFromByteArray)
    GetLongFromByteArray = ByteArray(Position + 0) + (256 * GetLongFromByteArray)
    Position = Position + 4
    '
    End Function
'
'========================================================================
' Get a Long from a byte array
'   position+0, 4 byte size of the number
'   position+3, start of the number
'========================================================================
'
Public Function GetNumberFromByteArray(ByteArray() As Byte, Position As Long) As Long
    ' ##### removed to catch err<>0 problem on error resume next
    '
    Dim ArgumentCount As Long
    Dim ArgumentLength As Long
    '
    ArgumentLength = GetLongFromByteArray(ByteArray(), Position)
    '
    If ArgumentLength > 0 Then
        GetNumberFromByteArray = 0
        For ArgumentCount = ArgumentLength - 1 To 0 Step -1
            GetNumberFromByteArray = ByteArray(Position + ArgumentCount) + (256 * GetNumberFromByteArray)
            Next
        End If
    Position = Position + ArgumentLength
    '
    End Function
'
'========================================================================
' Get a String a byte array
'   position+0, 4 byte length of the string
'   position+3, start of the string
'========================================================================
'
Public Function GetStringFromByteArray(ByteArray() As Byte, Position As Long) As String
    ' ##### removed to catch err<>0 problem on error resume next
    '
    Dim Pointer As Long
    Dim ArgumentLength As Long
    '
    ArgumentLength = GetLongFromByteArray(ByteArray(), Position)
    '
    GetStringFromByteArray = ""
    If ArgumentLength > 0 Then
        For Pointer = 0 To ArgumentLength - 1
            GetStringFromByteArray = GetStringFromByteArray & chr(ByteArray(Position + Pointer))
            Next
        End If
    Position = Position + ArgumentLength
    '
    End Function
'
'========================================================================
' Get a Long from a byte array
'========================================================================
'
Public Sub SetLongByteArray(ByRef ByteArray() As Byte, Position As Long, LongValue As Long)
    ' ##### removed to catch err<>0 problem on error resume next
    '
    ByteArray(Position + 0) = LongValue And (&HFF)
    ByteArray(Position + 1) = Int(LongValue / 256) And (&HFF)
    ByteArray(Position + 2) = Int(LongValue / (256 ^ 2)) And (&HFF)
    ByteArray(Position + 3) = Int(LongValue / (256 ^ 3)) And (&HFF)
    Position = Position + 4
    '
    End Sub
'
'========================================================================
' Set a string in a byte array
'========================================================================
'
Public Sub SetStringByteArray(ByRef ByteArray() As Byte, Position As Long, StringValue As String)
    ' ##### removed to catch err<>0 problem on error resume next
    '
    Dim Pointer As Long
    Dim LenStringValue As Long
    '
    LenStringValue = Len(StringValue)
    If LenStringValue > 0 Then
        For Pointer = 0 To LenStringValue - 1
            ByteArray(Position + Pointer) = Asc(Mid(StringValue, Pointer + 1, 1)) And (&HFF)
            Next
        Position = Position + LenStringValue
        End If
    '
    End Sub

'
'========================================================================
'   Set a Long long on the end of a RMB (Remote Method Block)
'       You determine the position, or it will add it to the end
'========================================================================
'
Public Sub SetRMBLong(ByRef ByteArray() As Byte, LongValue As Long, Optional Position)
    ' ##### removed to catch err<>0 problem on error resume next
    '
    Dim Temp As Long
    Dim MyPosition As Long
    Dim ByteArraySize As Long
    '
    ' ----- determine the position
    '
    If Not IsMissing(Position) Then
        MyPosition = Position
    Else
        '
        ' ----- Add it to the end, determine length
        '
        MyPosition = ByteArray(RMBPositionLength + 3)
        MyPosition = ByteArray(RMBPositionLength + 2) + (256 * MyPosition)
        MyPosition = ByteArray(RMBPositionLength + 1) + (256 * MyPosition)
        MyPosition = ByteArray(RMBPositionLength + 0) + (256 * MyPosition)
        '
        ' ----- adjust size of array if necessary
        '
        ByteArraySize = UBound(ByteArray)
        If ByteArraySize < (MyPosition + 8) Then
            ReDim Preserve ByteArray(ByteArraySize + 8)
            End If
        End If
    '
    ' ----- set the length
    '
    'ByteArray(MyPosition + 0) = 4
    'ByteArray(MyPosition + 1) = 0
    'ByteArray(MyPosition + 2) = 0
    'ByteArray(MyPosition + 3) = 0
    'MyPosition = MyPosition + 4
    '
    ' ----- set the value
    '
    ByteArray(MyPosition + 0) = LongValue And (&HFF)
    ByteArray(MyPosition + 1) = Int(LongValue / 256) And (&HFF)
    ByteArray(MyPosition + 2) = Int(LongValue / (256 ^ 2)) And (&HFF)
    ByteArray(MyPosition + 3) = Int(LongValue / (256 ^ 3)) And (&HFF)
    MyPosition = MyPosition + 4
    '
    If IsMissing(Position) Then
        '
        ' ----- Adjust the RMB length if length not given
        '
        ByteArray(RMBPositionLength + 0) = MyPosition And (&HFF)
        ByteArray(RMBPositionLength + 1) = Int(MyPosition / 256) And (&HFF)
        ByteArray(RMBPositionLength + 2) = Int(MyPosition / (256 ^ 2)) And (&HFF)
        ByteArray(RMBPositionLength + 3) = Int(MyPosition / (256 ^ 3)) And (&HFF)
        End If
    '
    End Sub
'
'========================================================================
'   Set a Long long on the end of a RMB (Remote Method Block)
'       You determine the position, or it will add it to the end
'========================================================================
'
Public Sub SetRMBString(ByRef ByteArray() As Byte, StringValue As String, Optional Position)
    ' ##### removed to catch err<>0 problem on error resume next
    '
    Dim Temp As Long
    Dim MyPosition As Long
    Dim ByteArraySize As Long
    '
    ' ----- determine the position
    '
    If Not IsMissing(Position) Then
        MyPosition = Position
    Else
        '
        ' ----- Add it to the end, determine length
        '
        MyPosition = ByteArray(RMBPositionLength + 3)
        MyPosition = ByteArray(RMBPositionLength + 2) + (256 * MyPosition)
        MyPosition = ByteArray(RMBPositionLength + 1) + (256 * MyPosition)
        MyPosition = ByteArray(RMBPositionLength + 0) + (256 * MyPosition)
        '
        ' ----- adjust size of array if necessary
        '
        ByteArraySize = UBound(ByteArray)
        If ByteArraySize < (MyPosition + 8) Then
            ReDim Preserve ByteArray(ByteArraySize + 4 + Len(StringValue))
            End If
        End If
    '
    ' ----- set the value
    '
    
    '
    Dim Pointer As Long
    Dim LenStringValue As Long
    '
    LenStringValue = Len(StringValue)
    If LenStringValue > 0 Then
        For Pointer = 0 To LenStringValue - 1
            ByteArray(MyPosition + Pointer) = Asc(Mid(StringValue, Pointer + 1, 1)) And (&HFF)
            Next
        MyPosition = MyPosition + LenStringValue
        End If
    '
    If IsMissing(Position) Then
        '
        ' ----- Adjust the RMB length if length not given
        '
        ByteArray(RMBPositionLength + 0) = MyPosition And (&HFF)
        ByteArray(RMBPositionLength + 1) = Int(MyPosition / 256) And (&HFF)
        ByteArray(RMBPositionLength + 2) = Int(MyPosition / (256 ^ 2)) And (&HFF)
        ByteArray(RMBPositionLength + 3) = Int(MyPosition / (256 ^ 3)) And (&HFF)
        End If
    '
    End Sub
'
'========================================================================
'   IsTrue
'       returns true or false depending on the state of the variant input
'========================================================================
'
Function IsTrue(ValueVariant) As Boolean
    IsTrue = kmaEncodeBoolean(ValueVariant)
    End Function
'
'========================================================================
' EncodeXML
'
'========================================================================
'
Function EncodeXML(ValueVariant As Variant, fieldType As Long) As String
    ' ##### removed to catch err<>0 problem on error resume next
    '
    Dim TimeValuething As Single
    Dim TimeHours As Long
    Dim TimeMinutes As Long
    Dim TimeSeconds As Long
    '
    Select Case fieldType
        Case FieldTypeInteger, FieldTypeLookup, FieldTypeRedirect, FieldTypeManyToMany, FieldTypeMemberSelect
            If IsNull(ValueVariant) Then
                EncodeXML = "null"
            ElseIf ValueVariant = "" Then
                EncodeXML = "null"
            ElseIf IsNumeric(ValueVariant) Then
                EncodeXML = Int(ValueVariant)
            Else
                EncodeXML = "null"
                End If
        Case FieldTypeBoolean
            If IsNull(ValueVariant) Then
                EncodeXML = "0"
            ElseIf ValueVariant <> False Then
                EncodeXML = "1"
            Else
                EncodeXML = "0"
                End If
        Case FieldTypeCurrency
            If IsNull(ValueVariant) Then
                EncodeXML = "null"
            ElseIf ValueVariant = "" Then
                EncodeXML = "null"
            ElseIf IsNumeric(ValueVariant) Then
                EncodeXML = ValueVariant
            Else
                EncodeXML = "null"
                End If
        Case FieldTypeFloat
            If IsNull(ValueVariant) Then
                EncodeXML = "null"
            ElseIf ValueVariant = "" Then
                EncodeXML = "null"
            ElseIf IsNumeric(ValueVariant) Then
                EncodeXML = ValueVariant
            Else
                EncodeXML = "null"
                End If
        Case FieldTypeDate
            If IsNull(ValueVariant) Then
                EncodeXML = "null"
            ElseIf ValueVariant = "" Then
                EncodeXML = "null"
            ElseIf IsDate(ValueVariant) Then
                'TimeVar = CDate(ValueVariant)
                'TimeValuething = 86400! * (TimeVar - Int(TimeVar))
                'TimeHours = Int(TimeValuething / 3600!)
                'TimeMinutes = Int(TimeValuething / 60!) - (TimeHours * 60)
                'TimeSeconds = TimeValuething - (TimeHours * 3600!) - (TimeMinutes * 60!)
                'EncodeXML = Year(ValueVariant) & "-" & Right("0" & Month(ValueVariant), 2) & "-" & Right("0" & Day(ValueVariant), 2) & " " & Right("0" & TimeHours, 2) & ":" & Right("0" & TimeMinutes, 2) & ":" & Right("0" & TimeSeconds, 2)
                EncodeXML = kmaEncodeText(ValueVariant)
                End If
        Case Else
            '
            ' ----- FieldTypeText
            ' ----- FieldTypeLongText
            ' ----- FieldTypeFile
            ' ----- FieldTypeImage
            ' ----- FieldTypeTextFile
            ' ----- FieldTypeCSSFile
            ' ----- FieldTypeXMLFile
            ' ----- FieldTypeJavascriptFile
            ' ----- FieldTypeLink
            ' ----- FieldTypeResourceLink
            ' ----- FieldTypeHTML
            ' ----- FieldTypeHTMLFile
            '
            If IsNull(ValueVariant) Then
                EncodeXML = "null"
            ElseIf ValueVariant = "" Then
                EncodeXML = ""
            Else
                'EncodeXML = ASPServer.HTMLEncode(ValueVariant)
                'EncodeXML = Replace(ValueVariant, "&", "&lt;")
                'EncodeXML = Replace(ValueVariant, "<", "&lt;")
                'EncodeXML = Replace(EncodeXML, ">", "&gt;")
                End If
        End Select
    '
    End Function
'
'========================================================================
' EncodeFilename
'
'========================================================================
'
Public Function encodeFilename(Source As String) As String
    Dim allowed As String
    Dim chr As String
    Dim Ptr As Long
    Dim cnt As Long
    Dim returnString As String
    '
    returnString = ""
    cnt = Len(Source)
    If cnt > 254 Then
        cnt = 254
    End If
    allowed = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ^&'@{}[],$-#()%.+~_"
    For Ptr = 1 To cnt
        chr = Mid(Source, Ptr, 1)
        If InStr(1, allowed, chr, vbBinaryCompare) Then
            returnString = returnString & chr
        End If
    Next
    encodeFilename = returnString
End Function
'
'Function encodeFilename(Filename As String) As String
'    ' ##### removed to catch err<>0 problem on error resume next
'    '
'    Dim Source() As Variant
'    Dim Replacement() As Variant
'    '
'    Source = Array("""", "*", "/", ":", "<", ">", "?", "\", "|", "=")
'    Replacement = Array("_", "_", "_", "_", "_", "_", "_", "_", "_", "_")
'    '
'    encodeFilename = ReplaceMany(Filename, Source, Replacement)
'    If Len(encodeFilename) > 254 Then
'        encodeFilename = Left(encodeFilename, 254)
'    End If
'    encodeFilename = Replace(encodeFilename, vbCr, "_")
'    encodeFilename = Replace(encodeFilename, vbLf, "_")
'    '
'    End Function
'
'
'

'
'========================================================================
' DecodeHTML
'
'========================================================================
'
Function DecodeHTML(Source As String) As String
    ' ##### removed to catch err<>0 problem on error resume next
    '
    DecodeHTML = kmaDecodeHTML(Source)
    'Dim SourceChr() As Variant
    'Dim ReplacementChr() As Variant
    ''
    'SourceChr = Array("&gt;", "&lt;", "&nbsp;", "&amp;")
    'ReplacementChr = Array(">", "<", " ", "&")
    ''
    'DecodeHTML = ReplaceMany(Source, SourceChr, ReplacementChr)
    '
    End Function
'
'========================================================================
' EncodeFilename
'
'========================================================================
'
Function ReplaceMany(Source As String, ArrayOfSource() As Variant, ArrayOfReplacement() As Variant) As String
    ' ##### removed to catch err<>0 problem on error resume next
    '
    Dim Count As Long
    Dim Pointer As Long
    '
    Count = UBound(ArrayOfSource) + 1
    ReplaceMany = Source
    For Pointer = 0 To Count - 1
        ReplaceMany = Replace(ReplaceMany, ArrayOfSource(Pointer), ArrayOfReplacement(Pointer))
        Next
    '
    End Function
'
'
'
Public Function GetURIHost(URI) As String
    ' ##### removed to catch err<>0 problem on error resume next
    '
    '   Divide the URI into URIHost, URIPath, and URIPage
    '
    Dim URIWorking As String
    Dim Slash As Long
    Dim LastSlash As Long
    Dim URIHost As String
    Dim URIPath As String
    Dim URIPage As String
    URIWorking = URI
    If Mid(UCase(URIWorking), 1, 4) = "HTTP" Then
        URIWorking = Mid(URIWorking, InStr(1, URIWorking, "//") + 2)
        End If
    URIHost = URIWorking
    Slash = InStr(1, URIHost, "/")
    If Slash = 0 Then
        URIPath = "/"
        URIPage = ""
    Else
        URIPath = Mid(URIHost, Slash)
        URIHost = Mid(URIHost, 1, Slash - 1)
        Slash = InStr(1, URIPath, "/")
        Do While Slash <> 0
            LastSlash = Slash
            Slash = InStr(LastSlash + 1, URIPath, "/")
            DoEvents
            Loop
        URIPage = Mid(URIPath, LastSlash + 1)
        URIPath = Mid(URIPath, 1, LastSlash)
        End If
    GetURIHost = URIHost
    '
    End Function
'
'
'
Public Function GetURIPage(URI) As String
    ' ##### removed to catch err<>0 problem on error resume next
    '
    '   Divide the URI into URIHost, URIPath, and URIPage
    '
    Dim Slash As Long
    Dim LastSlash As Long
    Dim URIHost As String
    Dim URIPath As String
    Dim URIPage As String
    Dim URIWorking As String
    URIWorking = URI
    If Mid(UCase(URIWorking), 1, 4) = "HTTP" Then
        URIWorking = Mid(URIWorking, InStr(1, URIWorking, "//") + 2)
        End If
    URIHost = URIWorking
    Slash = InStr(1, URIHost, "/")
    If Slash = 0 Then
        URIPath = "/"
        URIPage = ""
    Else
        URIPath = Mid(URIHost, Slash)
        URIHost = Mid(URIHost, 1, Slash - 1)
        Slash = InStr(1, URIPath, "/")
        Do While Slash <> 0
            LastSlash = Slash
            Slash = InStr(LastSlash + 1, URIPath, "/")
            DoEvents
            Loop
        URIPage = Mid(URIPath, LastSlash + 1)
        URIPath = Mid(URIPath, 1, LastSlash)
        End If
    GetURIPage = URIPage
    '
    End Function
'
'
'
Function GetDateFromGMT(GMTDate As String) As Date
    ' ##### removed to catch err<>0 problem on error resume next
    '
    Dim WorkString As String
    GetDateFromGMT = 0
    If GMTDate <> "" Then
        WorkString = Mid(GMTDate, 6, 11)
        If IsDate(WorkString) Then
            GetDateFromGMT = CDate(WorkString)
            WorkString = Mid(GMTDate, 18, 8)
            If IsDate(WorkString) Then
                GetDateFromGMT = GetDateFromGMT + CDate(WorkString) + 4 / 24
                End If
            End If
        End If
    '
    End Function
'
' Wdy, DD-Mon-YYYY HH:MM:SS GMT
'
Function GetGMTFromDate(DateValue As Date) As String
    '
    Dim WorkString As String
    Dim WorkLong As Long
    '
    If IsDate(DateValue) Then
        Select Case Weekday(DateValue)
            Case vbSunday
                GetGMTFromDate = "Sun, "
            Case vbMonday
                GetGMTFromDate = "Mon, "
            Case vbTuesday
                GetGMTFromDate = "Tue, "
            Case vbWednesday
                GetGMTFromDate = "Wed, "
            Case vbThursday
                GetGMTFromDate = "Thu, "
            Case vbFriday
                GetGMTFromDate = "Fri, "
            Case vbSaturday
                GetGMTFromDate = "Sat, "
        End Select
        '
        WorkLong = Day(DateValue)
        If WorkLong < 10 Then
            GetGMTFromDate = GetGMTFromDate & "0" & CStr(WorkLong) & " "
        Else
            GetGMTFromDate = GetGMTFromDate & CStr(WorkLong) & " "
        End If
        '
        Select Case Month(DateValue)
            Case 1
                GetGMTFromDate = GetGMTFromDate & "Jan "
            Case 2
                GetGMTFromDate = GetGMTFromDate & "Feb "
            Case 3
                GetGMTFromDate = GetGMTFromDate & "Mar "
            Case 4
                GetGMTFromDate = GetGMTFromDate & "Apr "
            Case 5
                GetGMTFromDate = GetGMTFromDate & "May "
            Case 6
                GetGMTFromDate = GetGMTFromDate & "Jun "
            Case 7
                GetGMTFromDate = GetGMTFromDate & "Jul "
            Case 8
                GetGMTFromDate = GetGMTFromDate & "Aug "
            Case 9
                GetGMTFromDate = GetGMTFromDate & "Sep "
            Case 10
                GetGMTFromDate = GetGMTFromDate & "Oct "
            Case 11
                GetGMTFromDate = GetGMTFromDate & "Nov "
            Case 12
                GetGMTFromDate = GetGMTFromDate & "Dec "
        End Select
        '
        GetGMTFromDate = GetGMTFromDate & CStr(Year(DateValue)) & " "
        '
        WorkLong = Hour(DateValue)
        If WorkLong < 10 Then
            GetGMTFromDate = GetGMTFromDate & "0" & CStr(WorkLong) & ":"
        Else
            GetGMTFromDate = GetGMTFromDate & CStr(WorkLong) & ":"
        End If
        '
        WorkLong = Minute(DateValue)
        If WorkLong < 10 Then
            GetGMTFromDate = GetGMTFromDate & "0" & CStr(WorkLong) & ":"
        Else
            GetGMTFromDate = GetGMTFromDate & CStr(WorkLong) & ":"
        End If
        '
        WorkLong = Second(DateValue)
        If WorkLong < 10 Then
            GetGMTFromDate = GetGMTFromDate & "0" & CStr(WorkLong)
        Else
            GetGMTFromDate = GetGMTFromDate & CStr(WorkLong)
        End If
        '
        GetGMTFromDate = GetGMTFromDate & " GMT"
    End If
    '
    End Function
'
'========================================================================
'   EncodeSQL
'       encode a variable to go in an sql expression
'       NOT supported
'========================================================================
'
Public Function EncodeSQL(ExpressionVariant As Variant, Optional fieldType As Variant) As String
    ' ##### removed to catch err<>0 problem on error resume next
    '
    Dim iFieldType As Long
    Dim MethodName As String
    '
    MethodName = "EncodeSQL"
    '
    iFieldType = KmaEncodeMissingInteger(fieldType, FieldTypeText)
    Select Case iFieldType
        Case FieldTypeBoolean
            EncodeSQL = KmaEncodeSQLBoolean(ExpressionVariant)
        Case FieldTypeCurrency, FieldTypeAutoIncrement, FieldTypeFloat, FieldTypeInteger, FieldTypeLookup, FieldTypeMemberSelect
            EncodeSQL = KmaEncodeSQLNumber(ExpressionVariant)
        Case FieldTypeDate
            EncodeSQL = KmaEncodeSQLDate(ExpressionVariant)
        Case FieldTypeLongText, FieldTypeHTML
            EncodeSQL = KmaEncodeSQLLongText(ExpressionVariant)
        Case FieldTypeFile, FieldTypeImage, FieldTypeLink, FieldTypeResourceLink, FieldTypeRedirect, FieldTypeManyToMany, FieldTypeText, FieldTypeTextFile, FieldTypeJavascriptFile, FieldTypeXMLFile, FieldTypeCSSFile, FieldTypeHTMLFile
            EncodeSQL = KmaEncodeSQLText(ExpressionVariant)
        Case Else
            EncodeSQL = KmaEncodeSQLText(ExpressionVariant)
            On Error GoTo 0
            Call Err.Raise(KmaErrorBase, App.EXEName, "Unknown Field Type [" & fieldType & "] used FieldTypeText.")
        End Select
    '
    End Function
''
''
''
'Public Sub AppendLogFile(Text)
'    On Error GoTo 0
'    Dim kmafs As New kmaFileSystem3.FileSystemClass
'    Dim Filename As String
'    Filename = GetProgramPath() & "\logs\Trace" & kmaEncodeText(CLng(Int(Now()))) & ".log"
'    Call kmafs.AppendLogFile2(Filename, """" & FormatDateTime(Now(), vbGeneralDate) & """,""" & Text & """" & vbCrLf)
'    End Sub
''
''========================================================================
''   HandleError
''       Logs the error and either resumes next, or raises it to the next level
''========================================================================
''
'Public Function HandleError(ClassName As String, MethodName As String, ErrNumber As Long, ErrSource As String, ErrDescription As String, ErrorTrap As Boolean, ResumeNext As Boolean, Optional URL As String) As String
'    ' ##### removed to catch err<>0 problem on error resume next
'    '
'    Dim ErrorMessage As String
'    '
'    If ErrorTrap Then
'        ErrorMessage = ErrorMessage & " Unexpected ErrorTrap"
'    Else
'        ErrorMessage = ErrorMessage & " Error"
'        End If
'    '
'    If URL <> "" Then
'        ErrorMessage = ErrorMessage & " on page [" & URL & "]"
'        End If
'    '
'    If ErrorTrap Then
'        If ResumeNext Then
'            Call AppendLogFile(App.EXEName & "." & ClassName & "." & MethodName & ErrorMessage & ", will resume after logging [" & ErrSource & " #" & ErrNumber & ", " & ErrDescription & "]")
'        Else
'            Call AppendLogFile(App.EXEName & "." & ClassName & "." & MethodName & ErrorMessage & ", will abort after logging [" & ErrSource & " #" & ErrNumber & ", " & ErrDescription & "]")
'            On Error GoTo 0
'            Call Err.Raise(ErrNumber, ErrSource, ErrDescription)
'            End If
'    Else
'        If ResumeNext Then
'            Call AppendLogFile(App.EXEName & "." & ClassName & "." & MethodName & ErrorMessage & ", will resume after logging  [" & ErrSource & " #" & ErrNumber & ", " & ErrDescription & "]")
'        Else
'            Call AppendLogFile(App.EXEName & "." & ClassName & "." & MethodName & ErrorMessage & ", will abort after logging [" & ErrSource & " #" & ErrNumber & ", " & ErrDescription & "]")
'            On Error GoTo 0
'            Call Err.Raise(ErrNumber, ErrSource, ErrDescription, , -1)
'            End If
'        End If
'    '
'    End Function
'
'
'
Public Sub cpTick(Text As String)
    ' ##### removed to catch err<>0 problem on error resume next
    '
    Dim iText As String
    Dim Duration As Long
    If CPTickCountBase <> 0 Then
        Duration = (GetTickCount - CPTickCountBase)
        End If
    iText = "cpTick " & Format(Duration / 1000, "0000.000") & " " & App.EXEName & " " & Text
    Call AppendLogFile(App.EXEName & ".?.cpTick, " & iText)
    CPTickCountBase = GetTickCount
    '
    End Sub
'
'=====================================================================================================
'   Set a value in a name/value pair
'=====================================================================================================
'
Public Sub SetNameValueArrays(InputName As String, InputValue As String, SQLName() As String, SQLValue() As String, Index As Long)
    ' ##### removed to catch err<>0 problem on error resume next
    '
    SQLName(Index) = InputName
    SQLValue(Index) = InputValue
    Index = Index + 1
    '
    End Sub
'
'
'
Public Function GetApplicationStatusMessage(ApplicationStatus As Long) As String
    Select Case ApplicationStatus
        Case ApplicationStatusNoHostService
            GetApplicationStatusMessage = "Contensive server not running"
        Case ApplicationStatusNotFound
            GetApplicationStatusMessage = "Contensive application not found"
        Case ApplicationStatusLoadedNotRunning
            GetApplicationStatusMessage = "Contensive application not running"
        Case ApplicationStatusRunning
            GetApplicationStatusMessage = "Contensive application running"
        Case ApplicationStatusStarting
            GetApplicationStatusMessage = "Contensive application starting"
        Case ApplicationStatusUpgrading
            GetApplicationStatusMessage = "Contensive database upgrading"
        Case ApplicationStatusDbBad
            GetApplicationStatusMessage = "Error verifying core database records"
        Case ApplicationStatusDbFailure
            GetApplicationStatusMessage = "Error opening application database"
        Case ApplicationStatusKernelFailure
            GetApplicationStatusMessage = "Error contacting Contensive kernel services"
        Case ApplicationStatusLicenseFailure
            GetApplicationStatusMessage = "Error verifying Contensive site license, see Http://www.Contensive.com/License"
        Case ApplicationStatusConnectionObjectFailure
            GetApplicationStatusMessage = "Error creating ODBC Connection object"
        Case ApplicationStatusConnectionStringFailure
            GetApplicationStatusMessage = "ODBC Data Source connection failed"
        Case ApplicationStatusDataSourceFailure
            GetApplicationStatusMessage = "Error opening default data source"
        Case ApplicationStatusDuplicateDomains
            GetApplicationStatusMessage = "Can not determine application because there are multiple applications with domain names that match this site's domain (See Application Manager)"
        Case ApplicationStatusUnknownFailure
            GetApplicationStatusMessage = "Unknown error, see trace log for details (/Contensive/Logs/trace____.log)"
        Case ApplicationStatusPaused
            GetApplicationStatusMessage = "Contensive application paused"
        Case Else
            GetApplicationStatusMessage = "Unknown status code [" & ApplicationStatus & "], see trace log for details"
        End Select
    End Function
'
'
'
Public Function GetFormInputSelectNameValue(SelectName As String, NameValueArray() As NameValuePairType) As String
    Dim Pointer As Long
    Dim Source() As NameValuePairType
    '
    Source = NameValueArray
    GetFormInputSelectNameValue = "<SELECT name=""" & SelectName & """ Size=""1"">"
    For Pointer = 0 To UBound(NameValueArray)
        GetFormInputSelectNameValue = GetFormInputSelectNameValue & "<OPTION value=""" & Source(Pointer).Value & """>" & Source(Pointer).Name & "</OPTION>"
        Next
    GetFormInputSelectNameValue = GetFormInputSelectNameValue & "</SELECT>"
    End Function
'
'
'
Public Function kmaGetSpacer(Width As Long, Height As Long) As String
    kmaGetSpacer = "<img alt=""space"" src=""/ccLib/images/spacer.gif"" width=""" & Width & """ height=""" & Height & """ border=""0"">"
    End Function
'
'
'
Public Function kmaProcessReplacement(NameValueLines As Variant, Source As Variant) As String
    '
    Dim iNameValueLines As String
    Dim Lines() As String
    Dim LineCnt As Long
    Dim LinePtr As Long
    '
    Dim Names() As String
    Dim Values() As String
    Dim PairPtr As Long
    Dim PairCnt As Long
    Dim Splits() As String
    '
    ' ----- read pairs in from NameValueLines
    '
    iNameValueLines = kmaEncodeText(NameValueLines)
    If InStr(1, iNameValueLines, "=") <> 0 Then
        PairCnt = 0
        Lines = SplitCRLF(iNameValueLines)
        For LinePtr = 0 To UBound(Lines)
            If InStr(1, Lines(LinePtr), "=") <> 0 Then
                Splits = Split(Lines(LinePtr), "=")
                ReDim Preserve Names(PairCnt)
                ReDim Preserve Names(PairCnt)
                ReDim Preserve Values(PairCnt)
                Names(PairCnt) = Trim(Splits(0))
                Names(PairCnt) = Replace(Names(PairCnt), vbTab, "")
                Splits(0) = ""
                Values(PairCnt) = Trim(Splits(1))
                PairCnt = PairCnt + 1
            End If
        Next
    End If
    '
    ' ----- Process replacements on Source
    '
    kmaProcessReplacement = kmaEncodeText(Source)
    If PairCnt > 0 Then
        For PairPtr = 0 To PairCnt - 1
            kmaProcessReplacement = Replace(kmaProcessReplacement, Names(PairPtr), Values(PairPtr), 1, 999, 1)
        Next
    End If
    '
    End Function
'
'==========================================================================================================================
'   To convert from site license to server licenses, we still need the URLEncoder in the site license
'   This routine generates a site license that is just the URL encoder.
'==========================================================================================================================
'
Public Function GetURLEncoder() As String
    Randomize
    GetURLEncoder = CStr(Int(1 + (Rnd() * 8))) & CStr(Int(1 + (Rnd() * 8))) & CStr(Int(1000000000 + (Rnd() * 899999999)))
End Function
'
'==========================================================================================================================
'   To convert from site license to server licenses, we still need the URLEncoder in the site license
'   This routine generates a site license that is just the URL encoder.
'==========================================================================================================================
'
Public Function GetSiteLicenseKey() As String
    GetSiteLicenseKey = "00000-00000-00000-" & GetURLEncoder
End Function
'
'
'
Public Sub ccAddTabEntry(Caption As String, Link As String, IsHit As Boolean, Optional StylePrefix As String, Optional LiveBody As String)
    On Error GoTo ErrorTrap
    '
    If TabsCnt <= TabsSize Then
        TabsSize = TabsSize + 10
        ReDim Preserve Tabs(TabsSize)
    End If
    With Tabs(TabsCnt)
        .Caption = Caption
        .Link = Link
        .IsHit = IsHit
        .StylePrefix = KmaEncodeMissingText(StylePrefix, "cc")
        .LiveBody = KmaEncodeMissingText(LiveBody, "")
    End With
    TabsCnt = TabsCnt + 1
    '
    Exit Sub
    '
ErrorTrap:
    Call Err.Raise(Err.Number, Err.Source, "Error in ccAddTabEntry-" & Err.Description)
End Sub
'
'
'
Public Function OldccGetTabs() As String
    On Error GoTo ErrorTrap
    '
    Dim TabPtr As Long
    Dim HitPtr As Long
    Dim IsLiveTab As Boolean
    Dim TabBody As String
    Dim TabLink As String
    Dim TabID As String
    Dim FirstLiveBodyShown As Boolean
    '
    If TabsCnt > 0 Then
        HitPtr = 0
        '
        ' Create TabBar
        '
        OldccGetTabs = OldccGetTabs & "<table border=0 cellspacing=0 cellpadding=0 align=center ><tr>"
        For TabPtr = 0 To TabsCnt - 1
            TabID = CStr(GetRandomInteger)
            If Tabs(TabPtr).LiveBody = "" Then
                '
                ' This tab is linked to a page
                '
                TabLink = kmaEncodeHTML(Tabs(TabPtr).Link)
            Else
                '
                ' This tab has a live body
                '
                TabLink = kmaEncodeHTML(Tabs(TabPtr).Link)
                If Not FirstLiveBodyShown Then
                    FirstLiveBodyShown = True
                    TabBody = TabBody & "<div style=""visibility: visible; position: absolute; left: 0px;"" class=""" & Tabs(TabPtr).StylePrefix & "Body"" id=""" & TabID & """></div>"
                Else
                    TabBody = TabBody & "<div style=""visibility: hidden; position: absolute; left: 0px;"" class=""" & Tabs(TabPtr).StylePrefix & "Body"" id=""" & TabID & """></div>"
                End If
            End If
            OldccGetTabs = OldccGetTabs & "<td valign=bottom>"
            If Tabs(TabPtr).IsHit And (HitPtr = 0) Then
                HitPtr = TabPtr
                '
                ' This tab is hit
                '
                OldccGetTabs = OldccGetTabs _
                    & "<table cellspacing=0 cellPadding=0 border=0>"
                OldccGetTabs = OldccGetTabs _
                    & "<tr>" _
                    & "<td colspan=2 height=1 width=2></td>" _
                    & "<td colspan=1 height=1 bgcolor=black></td>" _
                    & "<td colspan=3 height=1 width=3></td>" _
                    & "</tr>"
                OldccGetTabs = OldccGetTabs _
                    & "<tr>" _
                    & "<td colspan=1 height=1 width=1></td>" _
                    & "<td colspan=1 height=1 width=1 bgcolor=black></td>" _
                    & "<td colspan=1 height=1></td>" _
                    & "<td colspan=1 height=1 width=1 bgcolor=black></td>" _
                    & "<td colspan=2 height=1 width=2></td>" _
                    & "</tr>"
                OldccGetTabs = OldccGetTabs _
                    & "<tr>" _
                    & "<td colspan=1 height=2 bgcolor=black></td>" _
                    & "<td colspan=1 height=2></td>" _
                    & "<td colspan=1 height=2></td>" _
                    & "<td colspan=1 height=2></td>" _
                    & "<td colspan=1 height=2 width=1 bgcolor=black></td>" _
                    & "<td colspan=1 height=2 width=1></td>" _
                    & "</tr>"
                OldccGetTabs = OldccGetTabs _
                    & "<tr>" _
                    & "<td bgcolor=black></td>" _
                    & "<td></td>" _
                    & "<td>" _
                    & "<table cellspacing=0 cellPadding=2 border=0><tr>" _
                    & "<td Class=""ccTabHit"">&nbsp;<a href=""" & TabLink & """ Class=""ccTabHit"">" & Tabs(TabPtr).Caption & "</a>&nbsp;</td>" _
                    & "</tr></table >" _
                    & "</td>" _
                    & "<td></td>" _
                    & "<td bgcolor=black></td>" _
                    & "<td></td>" _
                    & "</tr>"
                OldccGetTabs = OldccGetTabs _
                    & "<tr>" _
                    & "<td bgcolor=black></td>" _
                    & "<td></td>" _
                    & "<td></td>" _
                    & "<td></td>" _
                    & "<td bgcolor=black></td>" _
                    & "<td bgcolor=black></td>" _
                    & "</tr>" _
                    & "</table >"
            Else
                '
                ' This tab is not hit
                '
                OldccGetTabs = OldccGetTabs _
                    & "<table cellspacing=0 cellPadding=0 border=0>"
                OldccGetTabs = OldccGetTabs _
                    & "<tr>" _
                    & "<td colspan=6 height=1></td>" _
                    & "</tr>"
                OldccGetTabs = OldccGetTabs _
                    & "<tr>" _
                    & "<td colspan=2 height=1></td>" _
                    & "<td colspan=1 height=1 bgcolor=black></td>" _
                    & "<td colspan=3 height=1></td>" _
                    & "</tr>"
                OldccGetTabs = OldccGetTabs _
                    & "<tr>" _
                    & "<td width=1></td>" _
                    & "<td width=1 bgcolor=black></td>" _
                    & "<td></td>" _
                    & "<td width=1 bgcolor=black></td>" _
                    & "<td width=2 colspan=2></td>" _
                    & "</tr>"
                OldccGetTabs = OldccGetTabs _
                    & "<tr>" _
                    & "<td width=1 bgcolor=black></td>" _
                    & "<td width=1></td>" _
                    & "<td nowrap>" _
                    & "<table cellspacing=0 cellPadding=2 border=0><tr>" _
                    & "<td Class=""ccTab"">&nbsp;<a href=""" & TabLink & """ Class=""ccTab"">" & Tabs(TabPtr).Caption & "</a>&nbsp;</td>" _
                    & "</tr></table >" _
                    & "</td>" _
                    & "<td width=1></td>" _
                    & "<td width=1 bgcolor=black></td>" _
                    & "<td width=1></td>" _
                    & "</tr>"
                OldccGetTabs = OldccGetTabs _
                    & "<tr>" _
                    & "<td colspan=6 height=1 bgcolor=black></td>" _
                    & "</tr>" _
                    & "</table >"
            End If
            OldccGetTabs = OldccGetTabs & "</td>"
        Next
        OldccGetTabs = OldccGetTabs & "<td class=""ccTabEnd"">&nbsp;</td></tr>"
        If TabBody <> "" Then
            OldccGetTabs = OldccGetTabs & "<tr><td colspan=6>" & TabBody & "</td></tr>"
        End If
        OldccGetTabs = OldccGetTabs & "</tr></table >"
        TabsCnt = 0
    End If
    '
    Exit Function
    '
ErrorTrap:
    Call Err.Raise(Err.Number, Err.Source, "Error in OldccGetTabs-" & Err.Description)
End Function


'
'
'
Public Function ccGetTabs() As String
    On Error GoTo ErrorTrap
    '
    Dim TabPtr As Long
    Dim HitPtr As Long
    Dim IsLiveTab As Boolean
    Dim TabBody As String
    Dim TabLink As String
    Dim TabID As String
    Dim FirstLiveBodyShown As Boolean
    '
    If TabsCnt > 0 Then
        HitPtr = 0
        '
        ' Create TabBar
        '
        ccGetTabs = ccGetTabs & "<table border=0 cellspacing=0 cellpadding=0 align=center ><tr>"
        For TabPtr = 0 To TabsCnt - 1
            TabID = CStr(GetRandomInteger)
            If Tabs(TabPtr).LiveBody = "" Then
                '
                ' This tab is linked to a page
                '
                TabLink = kmaEncodeHTML(Tabs(TabPtr).Link)
            Else
                '
                ' This tab has a live body
                '
                TabLink = kmaEncodeHTML(Tabs(TabPtr).Link)
                If Not FirstLiveBodyShown Then
                    FirstLiveBodyShown = True
                    TabBody = TabBody & "<div style=""visibility: visible; position: absolute; left: 0px;"" class=""" & Tabs(TabPtr).StylePrefix & "Body"" id=""" & TabID & """>" & Tabs(TabPtr).LiveBody & "</div>"
                Else
                    TabBody = TabBody & "<div style=""visibility: hidden; position: absolute; left: 0px;"" class=""" & Tabs(TabPtr).StylePrefix & "Body"" id=""" & TabID & """>" & Tabs(TabPtr).LiveBody & "</div>"
                End If
            End If
            ccGetTabs = ccGetTabs & "<td valign=bottom>"
            If Tabs(TabPtr).IsHit And (HitPtr = 0) Then
                HitPtr = TabPtr
                '
                ' This tab is hit
                '
                ccGetTabs = ccGetTabs _
                    & "<table cellspacing=0 cellPadding=0 border=0>"
                ccGetTabs = ccGetTabs _
                    & "<tr>" _
                    & "<td colspan=2 height=1 width=2></td>" _
                    & "<td colspan=1 height=1 bgcolor=black></td>" _
                    & "<td colspan=3 height=1 width=3></td>" _
                    & "</tr>"
                ccGetTabs = ccGetTabs _
                    & "<tr>" _
                    & "<td colspan=1 height=1 width=1></td>" _
                    & "<td colspan=1 height=1 width=1 bgcolor=black></td>" _
                    & "<td Class=""ccTabHit"" colspan=1 height=1></td>" _
                    & "<td colspan=1 height=1 width=1 bgcolor=black></td>" _
                    & "<td colspan=2 height=1 width=2></td>" _
                    & "</tr>"
                ccGetTabs = ccGetTabs _
                    & "<tr>" _
                    & "<td colspan=1 height=2 bgcolor=black></td>" _
                    & "<td Class=""ccTabHit"" colspan=1 height=2></td>" _
                    & "<td Class=""ccTabHit"" colspan=1 height=2></td>" _
                    & "<td Class=""ccTabHit"" colspan=1 height=2></td>" _
                    & "<td colspan=1 height=2 bgcolor=black></td>" _
                    & "<td colspan=1 height=2></td>" _
                    & "</tr>"
                ccGetTabs = ccGetTabs _
                    & "<tr>" _
                    & "<td bgcolor=black></td>" _
                    & "<td Class=""ccTabHit""></td>" _
                    & "<td Class=""ccTabHit"">" _
                    & "<table cellspacing=0 cellPadding=2 border=0><tr>" _
                    & "<td Class=""ccTabHit"">&nbsp;<a href=""" & TabLink & """ Class=""ccTabHit"">" & Tabs(TabPtr).Caption & "</a>&nbsp;</td>" _
                    & "</tr></table >" _
                    & "</td>" _
                    & "<td Class=""ccTabHit""></td>" _
                    & "<td bgcolor=black></td>" _
                    & "<td></td>" _
                    & "</tr>"
                ccGetTabs = ccGetTabs _
                    & "<tr>" _
                    & "<td bgcolor=black></td>" _
                    & "<td Class=""ccTabHit""></td>" _
                    & "<td Class=""ccTabHit""></td>" _
                    & "<td Class=""ccTabHit""></td>" _
                    & "<td bgcolor=black></td>" _
                    & "<td bgcolor=black></td>" _
                    & "</tr>" _
                    & "</table >"
            Else
                '
                ' This tab is not hit
                '
                ccGetTabs = ccGetTabs _
                    & "<table cellspacing=0 cellPadding=0 border=0>"
                ccGetTabs = ccGetTabs _
                    & "<tr>" _
                    & "<td colspan=6 height=1></td>" _
                    & "</tr>"
                ccGetTabs = ccGetTabs _
                    & "<tr>" _
                    & "<td colspan=2 height=1></td>" _
                    & "<td colspan=1 height=1 bgcolor=black></td>" _
                    & "<td colspan=3 height=1></td>" _
                    & "</tr>"
                ccGetTabs = ccGetTabs _
                    & "<tr>" _
                    & "<td width=1></td>" _
                    & "<td width=1 bgcolor=black></td>" _
                    & "<td Class=""ccTab""></td>" _
                    & "<td width=1 bgcolor=black></td>" _
                    & "<td width=2 colspan=2></td>" _
                    & "</tr>"
                ccGetTabs = ccGetTabs _
                    & "<tr>" _
                    & "<td width=1 bgcolor=black></td>" _
                    & "<td width=1 Class=""ccTab""></td>" _
                    & "<td nowrap Class=""ccTab"">" _
                    & "<table cellspacing=0 cellPadding=2 border=0><tr>" _
                    & "<td Class=""ccTab"">&nbsp;<a href=""" & TabLink & """ Class=""ccTab"">" & Tabs(TabPtr).Caption & "</a>&nbsp;</td>" _
                    & "</tr></table >" _
                    & "</td>" _
                    & "<td width=1 Class=""ccTab""></td>" _
                    & "<td width=1 bgcolor=black></td>" _
                    & "<td width=1></td>" _
                    & "</tr>"
                ccGetTabs = ccGetTabs _
                    & "<tr>" _
                    & "<td colspan=6 height=1 bgcolor=black></td>" _
                    & "</tr>" _
                    & "</table >"
            End If
            ccGetTabs = ccGetTabs & "</td>"
        Next
        ccGetTabs = ccGetTabs & "<td class=""ccTabEnd"">&nbsp;</td></tr>"
        If TabBody <> "" Then
            ccGetTabs = ccGetTabs & "<tr><td colspan=6>" & TabBody & "</td></tr>"
        End If
        ccGetTabs = ccGetTabs & "</tr></table >"
        TabsCnt = 0
    End If
    '
    Exit Function
    '
ErrorTrap:
    Call Err.Raise(Err.Number, Err.Source, "Error in ccGetTabs-" & Err.Description)
End Function
'
'
'
Public Function ConvertLinksToAbsolute(Source As String, RootLink As String) As String
    On Error GoTo ErrorTrap
    '
    Dim s As String
    '
    s = Source
    '
    s = Replace(s, " href=""", " href=""/", , , vbTextCompare)
    s = Replace(s, " href=""/http", " href=""http", , , vbTextCompare)
    s = Replace(s, " href=""/mailto", " href=""mailto", , , vbTextCompare)
    s = Replace(s, " href=""//", " href=""" & RootLink, , , vbTextCompare)
    s = Replace(s, " href=""/?", " href=""" & RootLink & "?", , , vbTextCompare)
    s = Replace(s, " href=""/", " href=""" & RootLink, , , vbTextCompare)
    '
    s = Replace(s, " href=", " href=/", , , vbTextCompare)
    s = Replace(s, " href=/""", " href=""", , , vbTextCompare)
    s = Replace(s, " href=/http", " href=http", , , vbTextCompare)
    s = Replace(s, " href=//", " href=" & RootLink, , , vbTextCompare)
    s = Replace(s, " href=/?", " href=" & RootLink & "?", , , vbTextCompare)
    s = Replace(s, " href=/", " href=" & RootLink, , , vbTextCompare)
    '
    s = Replace(s, " src=""", " src=""/", , , vbTextCompare)
    s = Replace(s, " src=""/http", " src=""http", , , vbTextCompare)
    s = Replace(s, " src=""//", " src=""" & RootLink, , , vbTextCompare)
    s = Replace(s, " src=""/?", " src=""" & RootLink & "?", , , vbTextCompare)
    s = Replace(s, " src=""/", " src=""" & RootLink, , , vbTextCompare)
    '
    s = Replace(s, " src=", " src=/", , , vbTextCompare)
    s = Replace(s, " src=/""", " src=""", , , vbTextCompare)
    s = Replace(s, " src=/http", " src=http", , , vbTextCompare)
    s = Replace(s, " src=//", " src=" & RootLink, , , vbTextCompare)
    s = Replace(s, " src=/?", " src=" & RootLink & "?", , , vbTextCompare)
    s = Replace(s, " src=/", " src=" & RootLink, , , vbTextCompare)
    '
    ConvertLinksToAbsolute = s
    '
    Exit Function
    '
ErrorTrap:
    Call Err.Raise(Err.Number, Err.Source, "Error in ConvertLinksToAbsolute-" & Err.Description)
End Function
'
'
'
Public Function GetProgramPath() As String
    GetProgramPath = App.Path
    If InStr(1, GetProgramPath, "c:\h\contensive\", vbTextCompare) <> 0 Then
        GetProgramPath = "c:\h\Contensive"
    ElseIf InStr(1, GetProgramPath, "c:\release\", vbTextCompare) <> 0 Then
        GetProgramPath = "c:\h\Contensive"
    End If
End Function
'
'
'
Public Function GetAddonRootPath() As String
    GetAddonRootPath = GetProgramPath
    If InStr(1, GetAddonRootPath, "c:\h\contensive", vbTextCompare) <> 0 Then
        '
        ' debugging - change program path to dummy path so addon builds all copy to
        '
        GetAddonRootPath = "c:\program files\kma\contensive"
    End If
    GetAddonRootPath = GetAddonRootPath & "\addons"
End Function
'
'
'
Public Function GetHTMLComment(Comment) As String
    GetHTMLComment = "<!-- " & Comment & " -->"
End Function
'
'
'
Public Function SplitCRLF(Expression As String) As String()
    Dim Args() As String
    Dim Ptr As Long
    '
    If InStr(1, Expression, vbCrLf) <> 0 Then
        SplitCRLF = Split(Expression, vbCrLf, , vbTextCompare)
    ElseIf InStr(1, Expression, vbCr) <> 0 Then
        SplitCRLF = Split(Expression, vbCr, , vbTextCompare)
    ElseIf InStr(1, Expression, vbLf) <> 0 Then
        SplitCRLF = Split(Expression, vbLf, , vbTextCompare)
    Else
        ReDim SplitCRLF(0)
        SplitCRLF = Split(Expression, vbCrLf)
    End If
End Function
'
'
'
Public Sub kmaShell(Cmd As String, Optional ByVal eWindowStyle As VBA.VbAppWinStyle = vbHide, Optional WaitForReturn As Boolean)
    On Error GoTo ErrorTrap
    '
    Dim ShellObj As Object
    '
    Set ShellObj = CreateObject("WScript.Shell")
    If Not (ShellObj Is Nothing) Then
        Call ShellObj.Run(Cmd, 0, WaitForReturn)
    End If
    Set ShellObj = Nothing
    Exit Sub
    '
ErrorTrap:
    Call AppendLogFile("ErrorTrap, kmaShell running command [" & Cmd & "], WaitForReturn=" & WaitForReturn & ", err=" & GetErrString(Err))
End Sub
'
'------------------------------------------------------------------------------------------------------------
'   Encodes an argument in an Addon OptionString (QueryString) for all non-allowed characters
'       call this before parsing them together
'       call decodeAddonConstructorArgument after parsing them apart
'
'       Arg0,Arg1,Arg2,Arg3,Name=Value&Name=VAlue[Option1|Option2]
'
'       This routine is needed for all Arg, Name, Value, Option values
'
'------------------------------------------------------------------------------------------------------------
'
Public Function EncodeAddonConstructorArgument(Arg As String) As String
    Dim a As String
    If Arg <> "" Then
        a = Arg
        If True Then
        'If AddonNewParse Then
            a = Replace(a, "\", "\\")
            a = Replace(a, vbCrLf, "\n")
            a = Replace(a, vbTab, "\t")
            a = Replace(a, "&", "\&")
            a = Replace(a, "=", "\=")
            a = Replace(a, ",", "\,")
            a = Replace(a, """", "\""")
            a = Replace(a, "'", "\'")
            a = Replace(a, "|", "\|")
            a = Replace(a, "[", "\[")
            a = Replace(a, "]", "\]")
            a = Replace(a, ":", "\:")
        End If
        EncodeAddonConstructorArgument = a
    End If
End Function
'
'------------------------------------------------------------------------------------------------------------
'   Decodes an argument parsed from an AddonConstructorString for all non-allowed characters
'       AddonConstructorString is a & delimited string of name=value[selector]descriptor
'
'       to get a value from an AddonConstructorString, first use getargument() to get the correct value[selector]descriptor
'       then remove everything to the right of any '['
'
'       call encodeAddonConstructorargument before parsing them together
'       call decodeAddonConstructorArgument after parsing them apart
'
'       Arg0,Arg1,Arg2,Arg3,Name=Value&Name=VAlue[Option1|Option2]
'
'       This routine is needed for all Arg, Name, Value, Option values
'
'------------------------------------------------------------------------------------------------------------
'
Public Function DecodeAddonConstructorArgument(EncodedArg As String) As String
    Dim a As String
    Dim Pos As Long
    '
    a = EncodedArg
    If True Then
    'If AddonNewParse Then
        a = Replace(a, "\:", ":")
        a = Replace(a, "\]", "]")
        a = Replace(a, "\[", "[")
        a = Replace(a, "\|", "|")
        a = Replace(a, "\'", "'")
        a = Replace(a, "\""", """")
        a = Replace(a, "\,", ",")
        a = Replace(a, "\=", "=")
        a = Replace(a, "\&", "&")
        a = Replace(a, "\t", vbTab)
        a = Replace(a, "\n", vbCrLf)
        a = Replace(a, "\\", "\")
    End If
    DecodeAddonConstructorArgument = a
End Function
'
'------------------------------------------------------------------------------------------------------------
'   use only internally
'
'   encode an argument to be used in a name=value& (N-V-A) string
'
'   an argument can be any one of these is this format:
'       Arg0,Arg1,Arg2,Arg3,Name=Value&Name=Value[Option1|Option2]descriptor
'
'   to create an nva string
'       string = encodeNvaArgument( name ) & "=" & encodeNvaArgument( value ) & "&"
'
'   to decode an nva string
'       split on ampersand then on equal, and decodeNvaArgument() each part
'
'------------------------------------------------------------------------------------------------------------
'
Public Function encodeNvaArgument(Arg As String) As String
    Dim a As String
    a = Arg
    If a <> "" Then
        a = Replace(a, vbCrLf, "#0013#")
        a = Replace(a, vbLf, "#0013#")
        a = Replace(a, vbCr, "#0013#")
        a = Replace(a, "&", "#0038#")
        a = Replace(a, "=", "#0061#")
        a = Replace(a, ",", "#0044#")
        a = Replace(a, """", "#0034#")
        a = Replace(a, "'", "#0039#")
        a = Replace(a, "|", "#0124#")
        a = Replace(a, "[", "#0091#")
        a = Replace(a, "]", "#0093#")
        a = Replace(a, ":", "#0058#")
    End If
    encodeNvaArgument = a
End Function
'
'------------------------------------------------------------------------------------------------------------
'   use only internally
'       decode an argument removed from a name=value& string
'       see encodeNvaArgument for details on how to use this
'------------------------------------------------------------------------------------------------------------
'
Public Function decodeNvaArgument(EncodedArg As String) As String
    Dim a As String
    '
    a = EncodedArg
    a = Replace(a, "#0058#", ":")
    a = Replace(a, "#0093#", "]")
    a = Replace(a, "#0091#", "[")
    a = Replace(a, "#0124#", "|")
    a = Replace(a, "#0039#", "'")
    a = Replace(a, "#0034#", """")
    a = Replace(a, "#0044#", ",")
    a = Replace(a, "#0061#", "=")
    a = Replace(a, "#0038#", "&")
    a = Replace(a, "#0013#", vbCrLf)
    decodeNvaArgument = a
End Function
'
' returns true of the link is a valid link on the source host
'
Public Function IsLinkToThisHost(Host As String, Link As String) As Boolean
    '
    Dim LinkHost As String
    Dim Pos As Long
    '
    If Trim(Link) = "" Then
        '
        ' Blank is not a link
        '
        IsLinkToThisHost = False
    ElseIf InStr(1, Link, "://") <> 0 Then
        '
        ' includes protocol, may be link to another site
        '
        LinkHost = LCase(Link)
        Pos = 1
        Pos = InStr(Pos, LinkHost, "://")
        If Pos > 0 Then
            Pos = InStr(Pos + 3, LinkHost, "/")
            If Pos > 0 Then
                LinkHost = Mid(LinkHost, 1, Pos - 1)
            End If
            IsLinkToThisHost = (LCase(Host) = LinkHost)
            If Not IsLinkToThisHost Then
                '
                ' try combinations including/excluding www.
                '
                If InStr(1, LinkHost, "www.", vbTextCompare) <> 0 Then
                    '
                    ' remove it
                    '
                    LinkHost = Replace(LinkHost, "www.", "", 1, -1, vbTextCompare)
                    IsLinkToThisHost = (LCase(Host) = LinkHost)
                Else
                    '
                    ' add it
                    '
                    LinkHost = Replace(LinkHost, "://", "://www.", 1, -1, vbTextCompare)
                    IsLinkToThisHost = (LCase(Host) = LinkHost)
                End If
            End If
        End If
    ElseIf InStr(1, Link, "#") = 1 Then
        '
        ' Is a bookmark, not a link
        '
        IsLinkToThisHost = False
    Else
        '
        ' all others are links on the source
        '
        IsLinkToThisHost = True
    End If
    If Not IsLinkToThisHost Then
        Link = Link
    End If
End Function
'
'========================================================================================================
'   ConvertLinkToRootRelative
'
'   /images/logo-main.jpg with any Basepath to /images/logo-main.jpg
'   http://gcm.brandeveolve.com/images/logo-main.jpg with any BasePath  to /images/logo-main.jpg
'   images/logo-main.jpg with Basepath '/' to /images/logo-main.jpg
'   logo-main.jpg with Basepath '/images/' to /images/logo-main.jpg
'
'========================================================================================================
'
Public Function ConvertLinkToRootRelative(Link As String, BasePath As String) As String
    '
    Dim Pos As Long
    '
    ConvertLinkToRootRelative = Link
    If InStr(1, Link, "/") = 1 Then
        '
        '   case /images/logo-main.jpg with any Basepath to /images/logo-main.jpg
        '
    ElseIf InStr(1, Link, "://") <> 0 Then
        '
        '   case http://gcm.brandeveolve.com/images/logo-main.jpg with any BasePath  to /images/logo-main.jpg
        '
        Pos = InStr(1, Link, "://")
        If Pos > 0 Then
            Pos = InStr(Pos + 3, Link, "/")
            If Pos > 0 Then
                ConvertLinkToRootRelative = Mid(Link, Pos)
            Else
                '
                ' This is just the domain name, RootRelative is the root
                '
                ConvertLinkToRootRelative = "/"
            End If
        End If
    Else
        '
        '   case images/logo-main.jpg with Basepath '/' to /images/logo-main.jpg
        '   case logo-main.jpg with Basepath '/images/' to /images/logo-main.jpg
        '
        ConvertLinkToRootRelative = BasePath & Link
    End If
    '
End Function
'
'
'
Public Function GetAddonIconImg(AdminURL As String, IconWidth As Long, IconHeight As Long, IconSprites As Long, IconIsInline As Boolean, IconImgID As String, IconFilename As String, serverFilePath As String, IconAlt As String, IconTitle As String, ACInstanceID As String, IconSpriteColumn As Long) As String
    '
    Dim ImgStyle As String
    Dim IconHeightNumeric As Long
    '
    If IconAlt = "" Then
        IconAlt = "Add-on"
    End If
    If IconTitle = "" Then
        IconTitle = "Rendered as Add-on"
    End If
    If IconFilename = "" Then
        '
        ' No icon given, use the default
        '
        If IconIsInline Then
            IconFilename = "/ccLib/images/IconAddonInlineDefault.png"
            IconWidth = 62
            IconHeight = 17
            IconSprites = 0
        Else
            IconFilename = "/ccLib/images/IconAddonBlockDefault.png"
            IconWidth = 57
            IconHeight = 59
            IconSprites = 4
        End If
    ElseIf InStr(1, IconFilename, "://") <> 0 Then
        '
        ' icon is an Absolute URL - leave it
        '
    ElseIf Left(IconFilename, 1) = "/" Then
        '
        ' icon is Root Relative, leave it
        '
    Else
        '
        ' icon is a virtual file, add the serverfilepath
        '
        IconFilename = serverFilePath & IconFilename
    End If
    'IconFilename = kmaEncodeJavascript(IconFilename)
    If (IconWidth = 0) Or (IconHeight = 0) Then
        IconSprites = 0
    End If
    
    If IconSprites = 0 Then
        '
        ' just the icon
        '
        GetAddonIconImg = "<img" _
            & " border=0" _
            & " id=""" & IconImgID & """" _
            & " onDblClick=""window.parent.OpenAddonPropertyWindow(this,'" & AdminURL & "');""" _
            & " alt=""" & IconAlt & """" _
            & " title=""" & IconTitle & """" _
            & " src=""" & IconFilename & """"
        'GetAddonIconImg = "<img" _
        '    & " id=""AC,AGGREGATEFUNCTION,0," & FieldName & "," & ArgumentList & """" _
        '    & " onDblClick=""window.parent.OpenAddonPropertyWindow(this);""" _
        '    & " alt=""" & IconAlt & """" _
        '    & " title=""" & IconTitle & """" _
        '    & " src=""" & IconFilename & """"
        If IconWidth <> 0 Then
            GetAddonIconImg = GetAddonIconImg & " width=""" & IconWidth & "px"""
        End If
        If IconHeight <> 0 Then
            GetAddonIconImg = GetAddonIconImg & " height=""" & IconHeight & "px"""
        End If
        If IconIsInline Then
            GetAddonIconImg = GetAddonIconImg & " style=""vertical-align:middle;display:inline;"" "
        Else
            GetAddonIconImg = GetAddonIconImg & " style=""display:block"" "
        End If
        If ACInstanceID <> "" Then
            GetAddonIconImg = GetAddonIconImg & " ACInstanceID=""" & ACInstanceID & """"
        End If
        GetAddonIconImg = GetAddonIconImg & ">"
    Else
        '
        ' Sprite Icon
        '
        GetAddonIconImg = GetIconSprite(IconImgID, IconSpriteColumn, IconFilename, IconWidth, IconHeight, IconAlt, IconTitle, "window.parent.OpenAddonPropertyWindow(this,'" & AdminURL & "');", IconIsInline, ACInstanceID)
'        GetAddonIconImg = "<img" _
'            & " border=0" _
'            & " id=""" & IconImgID & """" _
'            & " onMouseOver=""this.style.backgroundPosition='" & (-1 * IconSpriteColumn * IconWidth) & "px -" & (2 * IconHeight) & "px'""" _
'            & " onMouseOut=""this.style.backgroundPosition='" & (-1 * IconSpriteColumn * IconWidth) & "px 0px'""" _
'            & " onDblClick=""window.parent.OpenAddonPropertyWindow(this,'" & AdminURL & "');""" _
'            & " alt=""" & IconAlt & """" _
'            & " title=""" & IconTitle & """" _
'            & " src=""/ccLib/images/spacer.gif"""
'        ImgStyle = "background:url(" & IconFilename & ") " & (-1 * IconSpriteColumn * IconWidth) & "px 0px no-repeat;"
'        ImgStyle = ImgStyle & "width:" & IconWidth & "px;"
'        ImgStyle = ImgStyle & "height:" & IconHeight & "px;"
'        If IconIsInline Then
'            'GetAddonIconImg = GetAddonIconImg & " align=""middle"""
'            ImgStyle = ImgStyle & "vertical-align:middle;display:inline;"
'        Else
'            ImgStyle = ImgStyle & "display:block;"
'        End If
'
'
'        'Return_IconStyleMenuEntries = Return_IconStyleMenuEntries & vbCrLf & ",["".icon" & AddonID & """,false,"".icon" & AddonID & """,""background:url(" & IconFilename & ") 0px 0px no-repeat;"
'        'GetAddonIconImg = "<img" _
'        '    & " id=""AC,AGGREGATEFUNCTION,0," & FieldName & "," & ArgumentList & """" _
'        '    & " onMouseOver=""this.style.backgroundPosition=\'0px -" & (2 * IconHeight) & "px\'""" _
'        '    & " onMouseOut=""this.style.backgroundPosition=\'0px 0px\'""" _
'        '    & " onDblClick=""window.parent.OpenAddonPropertyWindow(this);""" _
'        '    & " alt=""" & IconAlt & """" _
'        '    & " title=""" & IconTitle & """" _
'        '    & " src=""/ccLib/images/spacer.gif"""
'        If ACInstanceID <> "" Then
'            GetAddonIconImg = GetAddonIconImg & " ACInstanceID=""" & ACInstanceID & """"
'        End If
'        GetAddonIconImg = GetAddonIconImg & " style=""" & ImgStyle & """>"
'        'Return_IconStyleMenuEntries = Return_IconStyleMenuEntries & """]"
    End If
End Function
'
'
'
Public Function ConvertRSTypeToGoogleType(RecordFieldType As Long) As String
    Select Case RecordFieldType
        Case 2, 3, 4, 5, 6, 14, 16, 17, 18, 19, 20, 21, 131
            ConvertRSTypeToGoogleType = "number"
        Case Else
            ConvertRSTypeToGoogleType = "string"
    End Select
End Function

'
'========================================================================
'   HandleError
'       Logs the error and either resumes next, or raises it to the next level
'========================================================================
'
Public Sub AppendLogFile2(ContensiveAppName As String, Context As String, ProgramName As String, ClassName As String, MethodName As String, ErrNumber As Long, ErrSource As String, ErrDescription As String, ErrorTrap As Boolean, ResumeNextAfterLogging As Boolean, URL As String, LogFolder As String, LogNamePrefix As String)
    On Error GoTo ErrorTrap
    '
    Dim MonthNumber As Long
    Dim DayNumber As Long
    Dim FilenameNoExt As String
    Dim kmafs As Object
    Dim ErrorMessage As String
    Dim LogLine As String
    Dim ResumeMessage As String
    Dim FolderFileList As String
    Dim FolderFiles() As String
    Dim Ptr As Long
    Dim PathFilenameNoExt As String
    Dim FileDetails() As String
    Dim fileSize As Long
Dim RetryCnt As Long
Dim SaveOK As Boolean
Dim FileSuffix As String
Dim iLogFolder As String
    '
    iLogFolder = LogFolder
    '
    If ErrorTrap Then
        ErrorMessage = "Error Trap"
    Else
        ErrorMessage = "Log Entry"
    End If
    '
    If ResumeNextAfterLogging Then
        ResumeMessage = "Resume after logging"
    Else
        ResumeMessage = "Abort after logging"
    End If
    '
    LogLine = "" _
        & LogFileCopyPrep(FormatDateTime(Now(), vbGeneralDate)) _
        & "," & LogFileCopyPrep(ContensiveAppName) _
        & "," & LogFileCopyPrep(ProgramName) _
        & "," & LogFileCopyPrep(ClassName) _
        & "," & LogFileCopyPrep(MethodName) _
        & "," & LogFileCopyPrep(Context) _
        & "," & LogFileCopyPrep(ErrorMessage) _
        & "," & LogFileCopyPrep(ResumeMessage) _
        & "," & LogFileCopyPrep(ErrSource) _
        & "," & LogFileCopyPrep(ErrNumber) _
        & "," & LogFileCopyPrep(ErrDescription) _
        & "," & LogFileCopyPrep(URL) _
        & vbCrLf
    '
    DayNumber = Day(Now)
    MonthNumber = Month(Now)
    FilenameNoExt = Year(Now)
    If MonthNumber < 10 Then
        FilenameNoExt = FilenameNoExt & "0"
    End If
    FilenameNoExt = FilenameNoExt & MonthNumber
    If DayNumber < 10 Then
        FilenameNoExt = FilenameNoExt & "0"
    End If
    FilenameNoExt = LogNamePrefix & FilenameNoExt & DayNumber
    If iLogFolder <> "" Then
        iLogFolder = iLogFolder & "\"
    End If
    iLogFolder = GetProgramPath & "\logs\" & iLogFolder
    PathFilenameNoExt = iLogFolder & FilenameNoExt
    '
    Set kmafs = CreateObject("kmaFileSystem3.FileSystemClass")
    FolderFileList = kmafs.GetFolderFiles2(iLogFolder)
    FolderFiles = Split(FolderFileList, vbCrLf)
    For Ptr = 0 To UBound(FolderFiles)
        If InStr(1, FolderFiles(Ptr), FilenameNoExt & ".log" & ",", vbTextCompare) <> 0 Then
            FileDetails = Split(FolderFiles(Ptr), vbTab)
            fileSize = kmaEncodeInteger(FileDetails(5))
            Exit For
        End If
    Next
    If fileSize < 10000000 Then
        RetryCnt = 0
        SaveOK = False
        FileSuffix = ""
        On Error Resume Next
        Do While (Not SaveOK) And (RetryCnt < 10)
            SaveOK = True
            Call kmafs.AppendFile(LCase(PathFilenameNoExt & FileSuffix & ".log"), LogLine)
            If Err.Number <> 0 Then
                If Err.Number = 70 Then
                    '
                    ' permission denied - happens when more then one process are writing at once, go to the next suffix
                    '
                    FileSuffix = "-" & CStr(RetryCnt + 1)
                    SaveOK = False
                Else
                    '
                    ' ignore all other errors - this routine logs errors, so there is nothing to do if it fails
                    '
                End If
                RetryCnt = RetryCnt + 1
                Err.Clear
            End If
        Loop
    End If
    Set kmafs = Nothing
    Exit Sub
    '
ErrorTrap:
    Err.Clear
End Sub
'
'========================================================================
'   HandleError
'       Logs the error and either resumes next, or raises it to the next level
'========================================================================
'
Public Sub HandleError2(ContensiveAppName As String, Context As String, ProgramName As String, ClassName As String, MethodName As String, ErrNumber As Long, ErrSource As String, ErrDescription As String, ErrorTrap As Boolean, ResumeNext As Boolean, URL As String)
    '
    Call AppendLogFile2(ContensiveAppName, Context, ProgramName, ClassName, MethodName, ErrNumber, ErrSource, ErrDescription, ErrorTrap, ResumeNext, URL, "", "Trace")
    '
    If Not ResumeNext Then
        On Error GoTo 0
        If ErrNumber = 0 Then
            Call Err.Raise(KmaErrorInternal, ErrSource, Context)
        Else
            Call Err.Raise(ErrNumber, ErrSource, ErrDescription)
        End If
    End If
    '
    End Sub
'
'
'
Private Function LogFileCopyPrep(Source) As String
    Dim Copy As String
    Copy = Source
    Copy = Replace(Copy, vbCrLf, " ")
    Copy = Replace(Copy, vbLf, " ")
    Copy = Replace(Copy, vbCr, " ")
    Copy = Replace(Copy, """", """""")
    Copy = """" & Copy & """"
    LogFileCopyPrep = Copy
End Function
' moved to csv
''
''=================================================================================================================
''   GetAddonOptionStringValue
''
''   gets the value from a list matching the name
''
''   InstanceOptionstring is an "AddonEncoded" name=AddonEncodedValue[selector]descriptor&name=value string
''=================================================================================================================
''
'Public Function GetAddonOptionStringValue(OptionName As String, AddonOptionString As String) As String
'    On Error GoTo ErrorTrap
'    '
'    Dim Pos As Long
'    Dim s As String
'    '
'    s = GetArgument(OptionName, AddonOptionString, "", "&")
'    Pos = InStr(1, s, "[")
'    If Pos > 0 Then
'        s = Left(s, Pos - 1)
'    End If
'    s = decodeNvaArgument(s)
'    '
'    GetAddonOptionStringValue = Trim(s)
'    '
'    Exit Function
'ErrorTrap:
'    Call HandleError2("", "", App.EXEName, "ccCommonModule", "GetAddonOptionStringValue", Err.Number, Err.Source, Err.Description, True, False, "")
'End Function
'
'
'
Public Function GetIconSprite(TagID As String, SpriteColumn As Long, IconSrc As String, IconWidth As Long, IconHeight As Long, IconAlt As String, IconTitle As String, onDblClick As String, IconIsInline As Boolean, ACInstanceID As String) As String
    '
    Dim ImgStyle As String
    '
        GetIconSprite = "<img" _
            & " border=0" _
            & " id=""" & TagID & """" _
            & " onMouseOver=""this.style.backgroundPosition='" & (-1 * SpriteColumn * IconWidth) & "px -" & (2 * IconHeight) & "px';""" _
            & " onMouseOut=""this.style.backgroundPosition='" & (-1 * SpriteColumn * IconWidth) & "px 0px'""" _
            & " onDblClick=""" & onDblClick & """" _
            & " alt=""" & IconAlt & """" _
            & " title=""" & IconTitle & """" _
            & " src=""/ccLib/images/spacer.gif"""
        ImgStyle = "background:url(" & IconSrc & ") " & (-1 * SpriteColumn * IconWidth) & "px 0px no-repeat;"
        ImgStyle = ImgStyle & "width:" & IconWidth & "px;"
        ImgStyle = ImgStyle & "height:" & IconHeight & "px;"
        If IconIsInline Then
            ImgStyle = ImgStyle & "vertical-align:middle;display:inline;"
        Else
            ImgStyle = ImgStyle & "display:block;"
        End If
        If ACInstanceID <> "" Then
            GetIconSprite = GetIconSprite & " ACInstanceID=""" & ACInstanceID & """"
        End If
        GetIconSprite = GetIconSprite & " style=""" & ImgStyle & """>"
End Function
'
'
'
Public Function RegGetValue$(MainKey&, SubKey$, Value$)
   ' MainKey must be one of the Publicly declared HKEY constants.
   Dim sKeyType&       'to return the key type.  This function expects REG_SZ or REG_DWORD
   Dim ret&            'returned by registry functions, should be 0&
   Dim lpHKey&         'return handle to opened key
   Dim lpcbData&       'length of data in returned string
   Dim ReturnedString$ 'returned string value
   Dim ReturnedLong&   'returned long value
   If MainKey >= &H80000000 And MainKey <= &H80000006 Then
      ' Open key
      ret = RegOpenKeyExA(MainKey, SubKey, 0&, KEY_READ, lpHKey)
      If ret <> ERROR_SUCCESS Then
         RegGetValue = ""
         Exit Function     'No key open, so leave
      End If
      
      ' Set up buffer for data to be returned in.
      ' Adjust next value for larger buffers.
      lpcbData = 255
      ReturnedString = Space$(lpcbData)

      ' Read key
      ret& = RegQueryValueExA(lpHKey, Value, ByVal 0&, sKeyType, ReturnedString, lpcbData)
      If ret <> ERROR_SUCCESS Then
         RegGetValue = ""   'Value probably doesn't exist
      Else
        If sKeyType = REG_DWORD Then
            ret = RegQueryValueEx(lpHKey, Value, ByVal 0&, sKeyType, ReturnedLong, 4)
            If ret = ERROR_SUCCESS Then RegGetValue = CStr(ReturnedLong)
        Else
            RegGetValue = Left$(ReturnedString, lpcbData - 1)
        End If
    End If
      ' Always close opened keys.
      ret = RegCloseKey(lpHKey)
   End If
End Function

