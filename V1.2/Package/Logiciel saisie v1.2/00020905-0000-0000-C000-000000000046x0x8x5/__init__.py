# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.1 (v3.3.1:d9893d13c628, Apr  6 2013, 20:25:12) [MSC v.1600 32 bit (Intel)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Mon Apr 15 10:41:18 2013
'Microsoft Word 14.0 Object Library'
makepy_version = '0.5.01'
python_version = 0x30301f0

import win32com.client.CLSIDToClass, pythoncom, pywintypes
import win32com.client.util
from pywintypes import IID
from win32com.client import Dispatch

# The following 3 lines may need tweaking for the particular server
# Candidates are pythoncom.Missing, .Empty and .ArgNotFound
defaultNamedOptArg=pythoncom.Empty
defaultNamedNotOptArg=pythoncom.Empty
defaultUnnamedArg=pythoncom.Empty

CLSID = IID('{00020905-0000-0000-C000-000000000046}')
MajorVersion = 8
MinorVersion = 5
LibraryFlags = 8
LCID = 0x0

class constants:
	wdAlertsAll                   =-1         # from enum WdAlertLevel
	wdAlertsMessageBox            =-2         # from enum WdAlertLevel
	wdAlertsNone                  =0          # from enum WdAlertLevel
	wdCenter                      =1          # from enum WdAlignmentTabAlignment
	wdLeft                        =0          # from enum WdAlignmentTabAlignment
	wdRight                       =2          # from enum WdAlignmentTabAlignment
	wdIndent                      =1          # from enum WdAlignmentTabRelative
	wdMargin                      =0          # from enum WdAlignmentTabRelative
	wdAnimationBlinkingBackground =2          # from enum WdAnimation
	wdAnimationLasVegasLights     =1          # from enum WdAnimation
	wdAnimationMarchingBlackAnts  =4          # from enum WdAnimation
	wdAnimationMarchingRedAnts    =5          # from enum WdAnimation
	wdAnimationNone               =0          # from enum WdAnimation
	wdAnimationShimmer            =6          # from enum WdAnimation
	wdAnimationSparkleText        =3          # from enum WdAnimation
	wdSessionStartSet             =1          # from enum WdApplyQuickStyleSets
	wdTemplateSet                 =2          # from enum WdApplyQuickStyleSets
	wdBoth                        =3          # from enum WdAraSpeller
	wdFinalYaa                    =2          # from enum WdAraSpeller
	wdInitialAlef                 =1          # from enum WdAraSpeller
	wdNone                        =0          # from enum WdAraSpeller
	wdNumeralArabic               =0          # from enum WdArabicNumeral
	wdNumeralContext              =2          # from enum WdArabicNumeral
	wdNumeralHindi                =1          # from enum WdArabicNumeral
	wdNumeralSystem               =3          # from enum WdArabicNumeral
	wdIcons                       =1          # from enum WdArrangeStyle
	wdTiled                       =0          # from enum WdArrangeStyle
	wdAutoFitContent              =1          # from enum WdAutoFitBehavior
	wdAutoFitFixed                =0          # from enum WdAutoFitBehavior
	wdAutoFitWindow               =2          # from enum WdAutoFitBehavior
	wdAutoClose                   =3          # from enum WdAutoMacros
	wdAutoExec                    =0          # from enum WdAutoMacros
	wdAutoExit                    =4          # from enum WdAutoMacros
	wdAutoNew                     =1          # from enum WdAutoMacros
	wdAutoOpen                    =2          # from enum WdAutoMacros
	wdAutoSync                    =5          # from enum WdAutoMacros
	wdAutoVersionOff              =0          # from enum WdAutoVersions
	wdAutoVersionOnClose          =1          # from enum WdAutoVersions
	wdBaselineAlignAuto           =4          # from enum WdBaselineAlignment
	wdBaselineAlignBaseline       =2          # from enum WdBaselineAlignment
	wdBaselineAlignCenter         =1          # from enum WdBaselineAlignment
	wdBaselineAlignFarEast50      =3          # from enum WdBaselineAlignment
	wdBaselineAlignTop            =0          # from enum WdBaselineAlignment
	wdSortByLocation              =1          # from enum WdBookmarkSortBy
	wdSortByName                  =0          # from enum WdBookmarkSortBy
	wdBorderDistanceFromPageEdge  =1          # from enum WdBorderDistanceFrom
	wdBorderDistanceFromText      =0          # from enum WdBorderDistanceFrom
	wdBorderBottom                =-3         # from enum WdBorderType
	wdBorderDiagonalDown          =-7         # from enum WdBorderType
	wdBorderDiagonalUp            =-8         # from enum WdBorderType
	wdBorderHorizontal            =-5         # from enum WdBorderType
	wdBorderLeft                  =-2         # from enum WdBorderType
	wdBorderRight                 =-4         # from enum WdBorderType
	wdBorderTop                   =-1         # from enum WdBorderType
	wdBorderVertical              =-6         # from enum WdBorderType
	emptyenum                     =0          # from enum WdBorderTypeHID
	wdColumnBreak                 =8          # from enum WdBreakType
	wdLineBreak                   =6          # from enum WdBreakType
	wdLineBreakClearLeft          =9          # from enum WdBreakType
	wdLineBreakClearRight         =10         # from enum WdBreakType
	wdPageBreak                   =7          # from enum WdBreakType
	wdSectionBreakContinuous      =3          # from enum WdBreakType
	wdSectionBreakEvenPage        =4          # from enum WdBreakType
	wdSectionBreakNextPage        =2          # from enum WdBreakType
	wdSectionBreakOddPage         =5          # from enum WdBreakType
	wdTextWrappingBreak           =11         # from enum WdBreakType
	wdBrowseComment               =3          # from enum WdBrowseTarget
	wdBrowseEdit                  =10         # from enum WdBrowseTarget
	wdBrowseEndnote               =5          # from enum WdBrowseTarget
	wdBrowseField                 =6          # from enum WdBrowseTarget
	wdBrowseFind                  =11         # from enum WdBrowseTarget
	wdBrowseFootnote              =4          # from enum WdBrowseTarget
	wdBrowseGoTo                  =12         # from enum WdBrowseTarget
	wdBrowseGraphic               =8          # from enum WdBrowseTarget
	wdBrowseHeading               =9          # from enum WdBrowseTarget
	wdBrowsePage                  =1          # from enum WdBrowseTarget
	wdBrowseSection               =2          # from enum WdBrowseTarget
	wdBrowseTable                 =7          # from enum WdBrowseTarget
	wdBrowserLevelMicrosoftInternetExplorer5=1          # from enum WdBrowserLevel
	wdBrowserLevelMicrosoftInternetExplorer6=2          # from enum WdBrowserLevel
	wdBrowserLevelV4              =0          # from enum WdBrowserLevel
	wdTypeAutoText                =9          # from enum WdBuildingBlockTypes
	wdTypeBibliography            =34         # from enum WdBuildingBlockTypes
	wdTypeCoverPage               =2          # from enum WdBuildingBlockTypes
	wdTypeCustom1                 =29         # from enum WdBuildingBlockTypes
	wdTypeCustom2                 =30         # from enum WdBuildingBlockTypes
	wdTypeCustom3                 =31         # from enum WdBuildingBlockTypes
	wdTypeCustom4                 =32         # from enum WdBuildingBlockTypes
	wdTypeCustom5                 =33         # from enum WdBuildingBlockTypes
	wdTypeCustomAutoText          =23         # from enum WdBuildingBlockTypes
	wdTypeCustomBibliography      =35         # from enum WdBuildingBlockTypes
	wdTypeCustomCoverPage         =16         # from enum WdBuildingBlockTypes
	wdTypeCustomEquations         =17         # from enum WdBuildingBlockTypes
	wdTypeCustomFooters           =18         # from enum WdBuildingBlockTypes
	wdTypeCustomHeaders           =19         # from enum WdBuildingBlockTypes
	wdTypeCustomPageNumber        =20         # from enum WdBuildingBlockTypes
	wdTypeCustomPageNumberBottom  =26         # from enum WdBuildingBlockTypes
	wdTypeCustomPageNumberPage    =27         # from enum WdBuildingBlockTypes
	wdTypeCustomPageNumberTop     =25         # from enum WdBuildingBlockTypes
	wdTypeCustomQuickParts        =15         # from enum WdBuildingBlockTypes
	wdTypeCustomTableOfContents   =28         # from enum WdBuildingBlockTypes
	wdTypeCustomTables            =21         # from enum WdBuildingBlockTypes
	wdTypeCustomTextBox           =24         # from enum WdBuildingBlockTypes
	wdTypeCustomWatermarks        =22         # from enum WdBuildingBlockTypes
	wdTypeEquations               =3          # from enum WdBuildingBlockTypes
	wdTypeFooters                 =4          # from enum WdBuildingBlockTypes
	wdTypeHeaders                 =5          # from enum WdBuildingBlockTypes
	wdTypePageNumber              =6          # from enum WdBuildingBlockTypes
	wdTypePageNumberBottom        =12         # from enum WdBuildingBlockTypes
	wdTypePageNumberPage          =13         # from enum WdBuildingBlockTypes
	wdTypePageNumberTop           =11         # from enum WdBuildingBlockTypes
	wdTypeQuickParts              =1          # from enum WdBuildingBlockTypes
	wdTypeTableOfContents         =14         # from enum WdBuildingBlockTypes
	wdTypeTables                  =7          # from enum WdBuildingBlockTypes
	wdTypeTextBox                 =10         # from enum WdBuildingBlockTypes
	wdTypeWatermarks              =8          # from enum WdBuildingBlockTypes
	wdPropertyAppName             =9          # from enum WdBuiltInProperty
	wdPropertyAuthor              =3          # from enum WdBuiltInProperty
	wdPropertyBytes               =22         # from enum WdBuiltInProperty
	wdPropertyCategory            =18         # from enum WdBuiltInProperty
	wdPropertyCharacters          =16         # from enum WdBuiltInProperty
	wdPropertyCharsWSpaces        =30         # from enum WdBuiltInProperty
	wdPropertyComments            =5          # from enum WdBuiltInProperty
	wdPropertyCompany             =21         # from enum WdBuiltInProperty
	wdPropertyFormat              =19         # from enum WdBuiltInProperty
	wdPropertyHiddenSlides        =27         # from enum WdBuiltInProperty
	wdPropertyHyperlinkBase       =29         # from enum WdBuiltInProperty
	wdPropertyKeywords            =4          # from enum WdBuiltInProperty
	wdPropertyLastAuthor          =7          # from enum WdBuiltInProperty
	wdPropertyLines               =23         # from enum WdBuiltInProperty
	wdPropertyMMClips             =28         # from enum WdBuiltInProperty
	wdPropertyManager             =20         # from enum WdBuiltInProperty
	wdPropertyNotes               =26         # from enum WdBuiltInProperty
	wdPropertyPages               =14         # from enum WdBuiltInProperty
	wdPropertyParas               =24         # from enum WdBuiltInProperty
	wdPropertyRevision            =8          # from enum WdBuiltInProperty
	wdPropertySecurity            =17         # from enum WdBuiltInProperty
	wdPropertySlides              =25         # from enum WdBuiltInProperty
	wdPropertySubject             =2          # from enum WdBuiltInProperty
	wdPropertyTemplate            =6          # from enum WdBuiltInProperty
	wdPropertyTimeCreated         =11         # from enum WdBuiltInProperty
	wdPropertyTimeLastPrinted     =10         # from enum WdBuiltInProperty
	wdPropertyTimeLastSaved       =12         # from enum WdBuiltInProperty
	wdPropertyTitle               =1          # from enum WdBuiltInProperty
	wdPropertyVBATotalEdit        =13         # from enum WdBuiltInProperty
	wdPropertyWords               =15         # from enum WdBuiltInProperty
	wdStyleBibliography           =-266       # from enum WdBuiltinStyle
	wdStyleBlockQuotation         =-85        # from enum WdBuiltinStyle
	wdStyleBodyText               =-67        # from enum WdBuiltinStyle
	wdStyleBodyText2              =-81        # from enum WdBuiltinStyle
	wdStyleBodyText3              =-82        # from enum WdBuiltinStyle
	wdStyleBodyTextFirstIndent    =-78        # from enum WdBuiltinStyle
	wdStyleBodyTextFirstIndent2   =-79        # from enum WdBuiltinStyle
	wdStyleBodyTextIndent         =-68        # from enum WdBuiltinStyle
	wdStyleBodyTextIndent2        =-83        # from enum WdBuiltinStyle
	wdStyleBodyTextIndent3        =-84        # from enum WdBuiltinStyle
	wdStyleBookTitle              =-265       # from enum WdBuiltinStyle
	wdStyleCaption                =-35        # from enum WdBuiltinStyle
	wdStyleClosing                =-64        # from enum WdBuiltinStyle
	wdStyleCommentReference       =-40        # from enum WdBuiltinStyle
	wdStyleCommentText            =-31        # from enum WdBuiltinStyle
	wdStyleDate                   =-77        # from enum WdBuiltinStyle
	wdStyleDefaultParagraphFont   =-66        # from enum WdBuiltinStyle
	wdStyleEmphasis               =-89        # from enum WdBuiltinStyle
	wdStyleEndnoteReference       =-43        # from enum WdBuiltinStyle
	wdStyleEndnoteText            =-44        # from enum WdBuiltinStyle
	wdStyleEnvelopeAddress        =-37        # from enum WdBuiltinStyle
	wdStyleEnvelopeReturn         =-38        # from enum WdBuiltinStyle
	wdStyleFooter                 =-33        # from enum WdBuiltinStyle
	wdStyleFootnoteReference      =-39        # from enum WdBuiltinStyle
	wdStyleFootnoteText           =-30        # from enum WdBuiltinStyle
	wdStyleHeader                 =-32        # from enum WdBuiltinStyle
	wdStyleHeading1               =-2         # from enum WdBuiltinStyle
	wdStyleHeading2               =-3         # from enum WdBuiltinStyle
	wdStyleHeading3               =-4         # from enum WdBuiltinStyle
	wdStyleHeading4               =-5         # from enum WdBuiltinStyle
	wdStyleHeading5               =-6         # from enum WdBuiltinStyle
	wdStyleHeading6               =-7         # from enum WdBuiltinStyle
	wdStyleHeading7               =-8         # from enum WdBuiltinStyle
	wdStyleHeading8               =-9         # from enum WdBuiltinStyle
	wdStyleHeading9               =-10        # from enum WdBuiltinStyle
	wdStyleHtmlAcronym            =-96        # from enum WdBuiltinStyle
	wdStyleHtmlAddress            =-97        # from enum WdBuiltinStyle
	wdStyleHtmlCite               =-98        # from enum WdBuiltinStyle
	wdStyleHtmlCode               =-99        # from enum WdBuiltinStyle
	wdStyleHtmlDfn                =-100       # from enum WdBuiltinStyle
	wdStyleHtmlKbd                =-101       # from enum WdBuiltinStyle
	wdStyleHtmlNormal             =-95        # from enum WdBuiltinStyle
	wdStyleHtmlPre                =-102       # from enum WdBuiltinStyle
	wdStyleHtmlSamp               =-103       # from enum WdBuiltinStyle
	wdStyleHtmlTt                 =-104       # from enum WdBuiltinStyle
	wdStyleHtmlVar                =-105       # from enum WdBuiltinStyle
	wdStyleHyperlink              =-86        # from enum WdBuiltinStyle
	wdStyleHyperlinkFollowed      =-87        # from enum WdBuiltinStyle
	wdStyleIndex1                 =-11        # from enum WdBuiltinStyle
	wdStyleIndex2                 =-12        # from enum WdBuiltinStyle
	wdStyleIndex3                 =-13        # from enum WdBuiltinStyle
	wdStyleIndex4                 =-14        # from enum WdBuiltinStyle
	wdStyleIndex5                 =-15        # from enum WdBuiltinStyle
	wdStyleIndex6                 =-16        # from enum WdBuiltinStyle
	wdStyleIndex7                 =-17        # from enum WdBuiltinStyle
	wdStyleIndex8                 =-18        # from enum WdBuiltinStyle
	wdStyleIndex9                 =-19        # from enum WdBuiltinStyle
	wdStyleIndexHeading           =-34        # from enum WdBuiltinStyle
	wdStyleIntenseEmphasis        =-262       # from enum WdBuiltinStyle
	wdStyleIntenseQuote           =-182       # from enum WdBuiltinStyle
	wdStyleIntenseReference       =-264       # from enum WdBuiltinStyle
	wdStyleLineNumber             =-41        # from enum WdBuiltinStyle
	wdStyleList                   =-48        # from enum WdBuiltinStyle
	wdStyleList2                  =-51        # from enum WdBuiltinStyle
	wdStyleList3                  =-52        # from enum WdBuiltinStyle
	wdStyleList4                  =-53        # from enum WdBuiltinStyle
	wdStyleList5                  =-54        # from enum WdBuiltinStyle
	wdStyleListBullet             =-49        # from enum WdBuiltinStyle
	wdStyleListBullet2            =-55        # from enum WdBuiltinStyle
	wdStyleListBullet3            =-56        # from enum WdBuiltinStyle
	wdStyleListBullet4            =-57        # from enum WdBuiltinStyle
	wdStyleListBullet5            =-58        # from enum WdBuiltinStyle
	wdStyleListContinue           =-69        # from enum WdBuiltinStyle
	wdStyleListContinue2          =-70        # from enum WdBuiltinStyle
	wdStyleListContinue3          =-71        # from enum WdBuiltinStyle
	wdStyleListContinue4          =-72        # from enum WdBuiltinStyle
	wdStyleListContinue5          =-73        # from enum WdBuiltinStyle
	wdStyleListNumber             =-50        # from enum WdBuiltinStyle
	wdStyleListNumber2            =-59        # from enum WdBuiltinStyle
	wdStyleListNumber3            =-60        # from enum WdBuiltinStyle
	wdStyleListNumber4            =-61        # from enum WdBuiltinStyle
	wdStyleListNumber5            =-62        # from enum WdBuiltinStyle
	wdStyleListParagraph          =-180       # from enum WdBuiltinStyle
	wdStyleMacroText              =-46        # from enum WdBuiltinStyle
	wdStyleMessageHeader          =-74        # from enum WdBuiltinStyle
	wdStyleNavPane                =-90        # from enum WdBuiltinStyle
	wdStyleNormal                 =-1         # from enum WdBuiltinStyle
	wdStyleNormalIndent           =-29        # from enum WdBuiltinStyle
	wdStyleNormalObject           =-158       # from enum WdBuiltinStyle
	wdStyleNormalTable            =-106       # from enum WdBuiltinStyle
	wdStyleNoteHeading            =-80        # from enum WdBuiltinStyle
	wdStylePageNumber             =-42        # from enum WdBuiltinStyle
	wdStylePlainText              =-91        # from enum WdBuiltinStyle
	wdStyleQuote                  =-181       # from enum WdBuiltinStyle
	wdStyleSalutation             =-76        # from enum WdBuiltinStyle
	wdStyleSignature              =-65        # from enum WdBuiltinStyle
	wdStyleStrong                 =-88        # from enum WdBuiltinStyle
	wdStyleSubtitle               =-75        # from enum WdBuiltinStyle
	wdStyleSubtleEmphasis         =-261       # from enum WdBuiltinStyle
	wdStyleSubtleReference        =-263       # from enum WdBuiltinStyle
	wdStyleTOAHeading             =-47        # from enum WdBuiltinStyle
	wdStyleTOC1                   =-20        # from enum WdBuiltinStyle
	wdStyleTOC2                   =-21        # from enum WdBuiltinStyle
	wdStyleTOC3                   =-22        # from enum WdBuiltinStyle
	wdStyleTOC4                   =-23        # from enum WdBuiltinStyle
	wdStyleTOC5                   =-24        # from enum WdBuiltinStyle
	wdStyleTOC6                   =-25        # from enum WdBuiltinStyle
	wdStyleTOC7                   =-26        # from enum WdBuiltinStyle
	wdStyleTOC8                   =-27        # from enum WdBuiltinStyle
	wdStyleTOC9                   =-28        # from enum WdBuiltinStyle
	wdStyleTableColorfulGrid      =-172       # from enum WdBuiltinStyle
	wdStyleTableColorfulList      =-171       # from enum WdBuiltinStyle
	wdStyleTableColorfulShading   =-170       # from enum WdBuiltinStyle
	wdStyleTableDarkList          =-169       # from enum WdBuiltinStyle
	wdStyleTableLightGrid         =-161       # from enum WdBuiltinStyle
	wdStyleTableLightGridAccent1  =-175       # from enum WdBuiltinStyle
	wdStyleTableLightList         =-160       # from enum WdBuiltinStyle
	wdStyleTableLightListAccent1  =-174       # from enum WdBuiltinStyle
	wdStyleTableLightShading      =-159       # from enum WdBuiltinStyle
	wdStyleTableLightShadingAccent1=-173       # from enum WdBuiltinStyle
	wdStyleTableMediumGrid1       =-166       # from enum WdBuiltinStyle
	wdStyleTableMediumGrid2       =-167       # from enum WdBuiltinStyle
	wdStyleTableMediumGrid3       =-168       # from enum WdBuiltinStyle
	wdStyleTableMediumList1       =-164       # from enum WdBuiltinStyle
	wdStyleTableMediumList1Accent1=-178       # from enum WdBuiltinStyle
	wdStyleTableMediumList2       =-165       # from enum WdBuiltinStyle
	wdStyleTableMediumShading1    =-162       # from enum WdBuiltinStyle
	wdStyleTableMediumShading1Accent1=-176       # from enum WdBuiltinStyle
	wdStyleTableMediumShading2    =-163       # from enum WdBuiltinStyle
	wdStyleTableMediumShading2Accent1=-177       # from enum WdBuiltinStyle
	wdStyleTableOfAuthorities     =-45        # from enum WdBuiltinStyle
	wdStyleTableOfFigures         =-36        # from enum WdBuiltinStyle
	wdStyleTitle                  =-63        # from enum WdBuiltinStyle
	wdStyleTocHeading             =-267       # from enum WdBuiltinStyle
	wdCalendarArabic              =1          # from enum WdCalendarType
	wdCalendarHebrew              =2          # from enum WdCalendarType
	wdCalendarJapan               =4          # from enum WdCalendarType
	wdCalendarKorean              =6          # from enum WdCalendarType
	wdCalendarSakaEra             =7          # from enum WdCalendarType
	wdCalendarTaiwan              =3          # from enum WdCalendarType
	wdCalendarThai                =5          # from enum WdCalendarType
	wdCalendarTranslitEnglish     =8          # from enum WdCalendarType
	wdCalendarTranslitFrench      =9          # from enum WdCalendarType
	wdCalendarUmalqura            =13         # from enum WdCalendarType
	wdCalendarWestern             =0          # from enum WdCalendarType
	wdCalendarTypeBidi            =99         # from enum WdCalendarTypeBi
	wdCalendarTypeGregorian       =100        # from enum WdCalendarTypeBi
	wdCaptionEquation             =-3         # from enum WdCaptionLabelID
	wdCaptionFigure               =-1         # from enum WdCaptionLabelID
	wdCaptionTable                =-2         # from enum WdCaptionLabelID
	wdCaptionNumberStyleArabic    =0          # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleArabicFullWidth=14         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleArabicLetter1=46         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleArabicLetter2=48         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleChosung   =25         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleGanada    =24         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleHanjaRead =41         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleHanjaReadDigit=42         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleHebrewLetter1=45         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleHebrewLetter2=47         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleHindiArabic=51         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleHindiCardinalText=52         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleHindiLetter1=49         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleHindiLetter2=50         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleKanji     =10         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleKanjiDigit=11         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleKanjiTraditional=16         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleLowercaseLetter=4          # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleLowercaseRoman=2          # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleNumberInCircle=18         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleSimpChinNum2=38         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleSimpChinNum3=39         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleThaiArabic=54         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleThaiCardinalText=55         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleThaiLetter=53         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleTradChinNum2=34         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleTradChinNum3=35         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleUppercaseLetter=3          # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleUppercaseRoman=1          # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleVietCardinalText=56         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleZodiac1   =30         # from enum WdCaptionNumberStyle
	wdCaptionNumberStyleZodiac2   =31         # from enum WdCaptionNumberStyle
	emptyenum                     =0          # from enum WdCaptionNumberStyleHID
	wdCaptionPositionAbove        =0          # from enum WdCaptionPosition
	wdCaptionPositionBelow        =1          # from enum WdCaptionPosition
	wdCellColorByAuthor           =-1         # from enum WdCellColor
	wdCellColorLightBlue          =2          # from enum WdCellColor
	wdCellColorLightGray          =7          # from enum WdCellColor
	wdCellColorLightGreen         =6          # from enum WdCellColor
	wdCellColorLightOrange        =5          # from enum WdCellColor
	wdCellColorLightPurple        =4          # from enum WdCellColor
	wdCellColorLightYellow        =3          # from enum WdCellColor
	wdCellColorNoHighlight        =0          # from enum WdCellColor
	wdCellColorPink               =1          # from enum WdCellColor
	wdCellAlignVerticalBottom     =3          # from enum WdCellVerticalAlignment
	wdCellAlignVerticalCenter     =1          # from enum WdCellVerticalAlignment
	wdCellAlignVerticalTop        =0          # from enum WdCellVerticalAlignment
	wdFullWidth                   =7          # from enum WdCharacterCase
	wdHalfWidth                   =6          # from enum WdCharacterCase
	wdHiragana                    =9          # from enum WdCharacterCase
	wdKatakana                    =8          # from enum WdCharacterCase
	wdLowerCase                   =0          # from enum WdCharacterCase
	wdNextCase                    =-1         # from enum WdCharacterCase
	wdTitleSentence               =4          # from enum WdCharacterCase
	wdTitleWord                   =2          # from enum WdCharacterCase
	wdToggleCase                  =5          # from enum WdCharacterCase
	wdUpperCase                   =1          # from enum WdCharacterCase
	emptyenum                     =0          # from enum WdCharacterCaseHID
	wdWidthFullWidth              =7          # from enum WdCharacterWidth
	wdWidthHalfWidth              =6          # from enum WdCharacterWidth
	wdCheckInMajorVersion         =1          # from enum WdCheckInVersionType
	wdCheckInMinorVersion         =0          # from enum WdCheckInVersionType
	wdCheckInOverwriteVersion     =2          # from enum WdCheckInVersionType
	wdAlwaysConvert               =1          # from enum WdChevronConvertRule
	wdAskToConvert                =3          # from enum WdChevronConvertRule
	wdAskToNotConvert             =2          # from enum WdChevronConvertRule
	wdNeverConvert                =0          # from enum WdChevronConvertRule
	wdCollapseEnd                 =0          # from enum WdCollapseDirection
	wdCollapseStart               =1          # from enum WdCollapseDirection
	wdColorAqua                   =13421619   # from enum WdColor
	wdColorAutomatic              =-16777216  # from enum WdColor
	wdColorBlack                  =0          # from enum WdColor
	wdColorBlue                   =16711680   # from enum WdColor
	wdColorBlueGray               =10053222   # from enum WdColor
	wdColorBrightGreen            =65280      # from enum WdColor
	wdColorBrown                  =13209      # from enum WdColor
	wdColorDarkBlue               =8388608    # from enum WdColor
	wdColorDarkGreen              =13056      # from enum WdColor
	wdColorDarkRed                =128        # from enum WdColor
	wdColorDarkTeal               =6697728    # from enum WdColor
	wdColorDarkYellow             =32896      # from enum WdColor
	wdColorGold                   =52479      # from enum WdColor
	wdColorGray05                 =15987699   # from enum WdColor
	wdColorGray10                 =15132390   # from enum WdColor
	wdColorGray125                =14737632   # from enum WdColor
	wdColorGray15                 =14277081   # from enum WdColor
	wdColorGray20                 =13421772   # from enum WdColor
	wdColorGray25                 =12632256   # from enum WdColor
	wdColorGray30                 =11776947   # from enum WdColor
	wdColorGray35                 =10921638   # from enum WdColor
	wdColorGray375                =10526880   # from enum WdColor
	wdColorGray40                 =10066329   # from enum WdColor
	wdColorGray45                 =9211020    # from enum WdColor
	wdColorGray50                 =8421504    # from enum WdColor
	wdColorGray55                 =7566195    # from enum WdColor
	wdColorGray60                 =6710886    # from enum WdColor
	wdColorGray625                =6316128    # from enum WdColor
	wdColorGray65                 =5855577    # from enum WdColor
	wdColorGray70                 =5000268    # from enum WdColor
	wdColorGray75                 =4210752    # from enum WdColor
	wdColorGray80                 =3355443    # from enum WdColor
	wdColorGray85                 =2500134    # from enum WdColor
	wdColorGray875                =2105376    # from enum WdColor
	wdColorGray90                 =1644825    # from enum WdColor
	wdColorGray95                 =789516     # from enum WdColor
	wdColorGreen                  =32768      # from enum WdColor
	wdColorIndigo                 =10040115   # from enum WdColor
	wdColorLavender               =16751052   # from enum WdColor
	wdColorLightBlue              =16737843   # from enum WdColor
	wdColorLightGreen             =13434828   # from enum WdColor
	wdColorLightOrange            =39423      # from enum WdColor
	wdColorLightTurquoise         =16777164   # from enum WdColor
	wdColorLightYellow            =10092543   # from enum WdColor
	wdColorLime                   =52377      # from enum WdColor
	wdColorOliveGreen             =13107      # from enum WdColor
	wdColorOrange                 =26367      # from enum WdColor
	wdColorPaleBlue               =16764057   # from enum WdColor
	wdColorPink                   =16711935   # from enum WdColor
	wdColorPlum                   =6697881    # from enum WdColor
	wdColorRed                    =255        # from enum WdColor
	wdColorRose                   =13408767   # from enum WdColor
	wdColorSeaGreen               =6723891    # from enum WdColor
	wdColorSkyBlue                =16763904   # from enum WdColor
	wdColorTan                    =10079487   # from enum WdColor
	wdColorTeal                   =8421376    # from enum WdColor
	wdColorTurquoise              =16776960   # from enum WdColor
	wdColorViolet                 =8388736    # from enum WdColor
	wdColorWhite                  =16777215   # from enum WdColor
	wdColorYellow                 =65535      # from enum WdColor
	wdAuto                        =0          # from enum WdColorIndex
	wdBlack                       =1          # from enum WdColorIndex
	wdBlue                        =2          # from enum WdColorIndex
	wdBrightGreen                 =4          # from enum WdColorIndex
	wdByAuthor                    =-1         # from enum WdColorIndex
	wdDarkBlue                    =9          # from enum WdColorIndex
	wdDarkRed                     =13         # from enum WdColorIndex
	wdDarkYellow                  =14         # from enum WdColorIndex
	wdGray25                      =16         # from enum WdColorIndex
	wdGray50                      =15         # from enum WdColorIndex
	wdGreen                       =11         # from enum WdColorIndex
	wdNoHighlight                 =0          # from enum WdColorIndex
	wdPink                        =5          # from enum WdColorIndex
	wdRed                         =6          # from enum WdColorIndex
	wdTeal                        =10         # from enum WdColorIndex
	wdTurquoise                   =3          # from enum WdColorIndex
	wdViolet                      =12         # from enum WdColorIndex
	wdWhite                       =8          # from enum WdColorIndex
	wdYellow                      =7          # from enum WdColorIndex
	wdCompareDestinationNew       =2          # from enum WdCompareDestination
	wdCompareDestinationOriginal  =0          # from enum WdCompareDestination
	wdCompareDestinationRevised   =1          # from enum WdCompareDestination
	wdCompareTargetCurrent        =1          # from enum WdCompareTarget
	wdCompareTargetNew            =2          # from enum WdCompareTarget
	wdCompareTargetSelected       =0          # from enum WdCompareTarget
	wdAlignTablesRowByRow         =39         # from enum WdCompatibility
	wdAllowSpaceOfSameStyleInTable=54         # from enum WdCompatibility
	wdApplyBreakingRules          =46         # from enum WdCompatibility
	wdAutofitLikeWW11             =57         # from enum WdCompatibility
	wdAutospaceLikeWW7            =38         # from enum WdCompatibility
	wdCachedColBalance            =65         # from enum WdCompatibility
	wdConvMailMergeEsc            =6          # from enum WdCompatibility
	wdDisableOTKerning            =66         # from enum WdCompatibility
	wdDontAdjustLineHeightInTable =36         # from enum WdCompatibility
	wdDontAutofitConstrainedTables=56         # from enum WdCompatibility
	wdDontBalanceSingleByteDoubleByteWidth=16         # from enum WdCompatibility
	wdDontBreakConstrainedForcedTables=62         # from enum WdCompatibility
	wdDontBreakWrappedTables      =43         # from enum WdCompatibility
	wdDontOverrideTableStyleFontSzAndJustification=68         # from enum WdCompatibility
	wdDontSnapTextToGridInTableWithObjects=44         # from enum WdCompatibility
	wdDontULTrailSpace            =15         # from enum WdCompatibility
	wdDontUseAsianBreakRulesInGrid=48         # from enum WdCompatibility
	wdDontUseHTMLParagraphAutoSpacing=35         # from enum WdCompatibility
	wdDontUseIndentAsNumberingTabStop=52         # from enum WdCompatibility
	wdDontVertAlignCellWithShape  =61         # from enum WdCompatibility
	wdDontVertAlignInTextbox      =63         # from enum WdCompatibility
	wdDontWrapTextWithPunctuation =47         # from enum WdCompatibility
	wdExactOnTop                  =28         # from enum WdCompatibility
	wdExpandShiftReturn           =14         # from enum WdCompatibility
	wdFELineBreak11               =53         # from enum WdCompatibility
	wdFlipMirrorIndents           =67         # from enum WdCompatibility
	wdFootnoteLayoutLikeWW8       =34         # from enum WdCompatibility
	wdForgetLastTabAlignment      =37         # from enum WdCompatibility
	wdGrowAutofit                 =50         # from enum WdCompatibility
	wdHangulWidthLikeWW11         =59         # from enum WdCompatibility
	wdLayoutRawTableWidth         =40         # from enum WdCompatibility
	wdLayoutTableRowsApart        =41         # from enum WdCompatibility
	wdLeaveBackslashAlone         =13         # from enum WdCompatibility
	wdLineWrapLikeWord6           =32         # from enum WdCompatibility
	wdMWSmallCaps                 =22         # from enum WdCompatibility
	wdNoColumnBalance             =5          # from enum WdCompatibility
	wdNoExtraLineSpacing          =23         # from enum WdCompatibility
	wdNoLeading                   =20         # from enum WdCompatibility
	wdNoSpaceForUL                =21         # from enum WdCompatibility
	wdNoSpaceRaiseLower           =2          # from enum WdCompatibility
	wdNoTabHangIndent             =1          # from enum WdCompatibility
	wdOrigWordTableRules          =9          # from enum WdCompatibility
	wdPrintBodyTextBeforeHeader   =19         # from enum WdCompatibility
	wdPrintColBlack               =3          # from enum WdCompatibility
	wdSelectFieldWithFirstOrLastCharacter=45         # from enum WdCompatibility
	wdShapeLayoutLikeWW8          =33         # from enum WdCompatibility
	wdShowBreaksInFrames          =11         # from enum WdCompatibility
	wdSpacingInWholePoints        =18         # from enum WdCompatibility
	wdSplitPgBreakAndParaMark     =60         # from enum WdCompatibility
	wdSubFontBySize               =25         # from enum WdCompatibility
	wdSuppressBottomSpacing       =29         # from enum WdCompatibility
	wdSuppressSpBfAfterPgBrk      =7          # from enum WdCompatibility
	wdSuppressTopSpacing          =8          # from enum WdCompatibility
	wdSuppressTopSpacingMac5      =17         # from enum WdCompatibility
	wdSwapBordersFacingPages      =12         # from enum WdCompatibility
	wdTransparentMetafiles        =10         # from enum WdCompatibility
	wdTruncateFontHeight          =24         # from enum WdCompatibility
	wdUnderlineTabInNumList       =58         # from enum WdCompatibility
	wdUseNormalStyleForList       =51         # from enum WdCompatibility
	wdUsePrinterMetrics           =26         # from enum WdCompatibility
	wdUseWord2002TableStyleRules  =49         # from enum WdCompatibility
	wdUseWord97LineBreakingRules  =42         # from enum WdCompatibility
	wdWPJustification             =31         # from enum WdCompatibility
	wdWPSpaceWidth                =30         # from enum WdCompatibility
	wdWW11IndentRules             =55         # from enum WdCompatibility
	wdWW6BorderRules              =27         # from enum WdCompatibility
	wdWord11KerningPairs          =64         # from enum WdCompatibility
	wdWrapTrailSpaces             =4          # from enum WdCompatibility
	wdCurrent                     =65535      # from enum WdCompatibilityMode
	wdWord2003                    =11         # from enum WdCompatibilityMode
	wdWord2007                    =12         # from enum WdCompatibilityMode
	wdWord2010                    =14         # from enum WdCompatibilityMode
	wdEvenColumnBanding           =7          # from enum WdConditionCode
	wdEvenRowBanding              =3          # from enum WdConditionCode
	wdFirstColumn                 =4          # from enum WdConditionCode
	wdFirstRow                    =0          # from enum WdConditionCode
	wdLastColumn                  =5          # from enum WdConditionCode
	wdLastRow                     =1          # from enum WdConditionCode
	wdNECell                      =8          # from enum WdConditionCode
	wdNWCell                      =9          # from enum WdConditionCode
	wdOddColumnBanding            =6          # from enum WdConditionCode
	wdOddRowBanding               =2          # from enum WdConditionCode
	wdSECell                      =10         # from enum WdConditionCode
	wdSWCell                      =11         # from enum WdConditionCode
	wdAutoPosition                =0          # from enum WdConstants
	wdBackward                    =-1073741823 # from enum WdConstants
	wdCreatorCode                 =1297307460 # from enum WdConstants
	wdFirst                       =1          # from enum WdConstants
	wdForward                     =1073741823 # from enum WdConstants
	wdToggle                      =9999998    # from enum WdConstants
	wdUndefined                   =9999999    # from enum WdConstants
	wdContentControlDateStorageDate=1          # from enum WdContentControlDateStorageFormat
	wdContentControlDateStorageDateTime=2          # from enum WdContentControlDateStorageFormat
	wdContentControlDateStorageText=0          # from enum WdContentControlDateStorageFormat
	wdContentControlBuildingBlockGallery=5          # from enum WdContentControlType
	wdContentControlCheckBox      =8          # from enum WdContentControlType
	wdContentControlComboBox      =3          # from enum WdContentControlType
	wdContentControlDate          =6          # from enum WdContentControlType
	wdContentControlDropdownList  =4          # from enum WdContentControlType
	wdContentControlGroup         =7          # from enum WdContentControlType
	wdContentControlPicture       =2          # from enum WdContentControlType
	wdContentControlRichText      =0          # from enum WdContentControlType
	wdContentControlText          =1          # from enum WdContentControlType
	wdContinueDisabled            =0          # from enum WdContinue
	wdContinueList                =2          # from enum WdContinue
	wdResetList                   =1          # from enum WdContinue
	wdArgentina                   =54         # from enum WdCountry
	wdBrazil                      =55         # from enum WdCountry
	wdCanada                      =2          # from enum WdCountry
	wdChile                       =56         # from enum WdCountry
	wdChina                       =86         # from enum WdCountry
	wdDenmark                     =45         # from enum WdCountry
	wdFinland                     =358        # from enum WdCountry
	wdFrance                      =33         # from enum WdCountry
	wdGermany                     =49         # from enum WdCountry
	wdIceland                     =354        # from enum WdCountry
	wdItaly                       =39         # from enum WdCountry
	wdJapan                       =81         # from enum WdCountry
	wdKorea                       =82         # from enum WdCountry
	wdLatinAmerica                =3          # from enum WdCountry
	wdMexico                      =52         # from enum WdCountry
	wdNetherlands                 =31         # from enum WdCountry
	wdNorway                      =47         # from enum WdCountry
	wdPeru                        =51         # from enum WdCountry
	wdSpain                       =34         # from enum WdCountry
	wdSweden                      =46         # from enum WdCountry
	wdTaiwan                      =886        # from enum WdCountry
	wdUK                          =44         # from enum WdCountry
	wdUS                          =1          # from enum WdCountry
	wdVenezuela                   =58         # from enum WdCountry
	wdCursorMovementLogical       =0          # from enum WdCursorMovement
	wdCursorMovementVisual        =1          # from enum WdCursorMovement
	wdCursorIBeam                 =1          # from enum WdCursorType
	wdCursorNormal                =2          # from enum WdCursorType
	wdCursorNorthwestArrow        =3          # from enum WdCursorType
	wdCursorWait                  =0          # from enum WdCursorType
	wdCustomLabelA4               =2          # from enum WdCustomLabelPageSize
	wdCustomLabelA4LS             =3          # from enum WdCustomLabelPageSize
	wdCustomLabelA5               =4          # from enum WdCustomLabelPageSize
	wdCustomLabelA5LS             =5          # from enum WdCustomLabelPageSize
	wdCustomLabelB4JIS            =13         # from enum WdCustomLabelPageSize
	wdCustomLabelB5               =6          # from enum WdCustomLabelPageSize
	wdCustomLabelFanfold          =8          # from enum WdCustomLabelPageSize
	wdCustomLabelHigaki           =11         # from enum WdCustomLabelPageSize
	wdCustomLabelHigakiLS         =12         # from enum WdCustomLabelPageSize
	wdCustomLabelLetter           =0          # from enum WdCustomLabelPageSize
	wdCustomLabelLetterLS         =1          # from enum WdCustomLabelPageSize
	wdCustomLabelMini             =7          # from enum WdCustomLabelPageSize
	wdCustomLabelVertHalfSheet    =9          # from enum WdCustomLabelPageSize
	wdCustomLabelVertHalfSheetLS  =10         # from enum WdCustomLabelPageSize
	wdDateLanguageBidi            =10         # from enum WdDateLanguage
	wdDateLanguageLatin           =1033       # from enum WdDateLanguage
	wdAutoRecoverPath             =5          # from enum WdDefaultFilePath
	wdBorderArtPath               =19         # from enum WdDefaultFilePath
	wdCurrentFolderPath           =14         # from enum WdDefaultFilePath
	wdDocumentsPath               =0          # from enum WdDefaultFilePath
	wdGraphicsFiltersPath         =10         # from enum WdDefaultFilePath
	wdPicturesPath                =1          # from enum WdDefaultFilePath
	wdProgramPath                 =9          # from enum WdDefaultFilePath
	wdProofingToolsPath           =12         # from enum WdDefaultFilePath
	wdStartupPath                 =8          # from enum WdDefaultFilePath
	wdStyleGalleryPath            =15         # from enum WdDefaultFilePath
	wdTempFilePath                =13         # from enum WdDefaultFilePath
	wdTextConvertersPath          =11         # from enum WdDefaultFilePath
	wdToolsPath                   =6          # from enum WdDefaultFilePath
	wdTutorialPath                =7          # from enum WdDefaultFilePath
	wdUserOptionsPath             =4          # from enum WdDefaultFilePath
	wdUserTemplatesPath           =2          # from enum WdDefaultFilePath
	wdWorkgroupTemplatesPath      =3          # from enum WdDefaultFilePath
	wdWord10ListBehavior          =2          # from enum WdDefaultListBehavior
	wdWord8ListBehavior           =0          # from enum WdDefaultListBehavior
	wdWord9ListBehavior           =1          # from enum WdDefaultListBehavior
	wdWord8TableBehavior          =0          # from enum WdDefaultTableBehavior
	wdWord9TableBehavior          =1          # from enum WdDefaultTableBehavior
	wdDeleteCellsEntireColumn     =3          # from enum WdDeleteCells
	wdDeleteCellsEntireRow        =2          # from enum WdDeleteCells
	wdDeleteCellsShiftLeft        =0          # from enum WdDeleteCells
	wdDeleteCellsShiftUp          =1          # from enum WdDeleteCells
	wdDeletedTextMarkBold         =5          # from enum WdDeletedTextMark
	wdDeletedTextMarkCaret        =2          # from enum WdDeletedTextMark
	wdDeletedTextMarkColorOnly    =9          # from enum WdDeletedTextMark
	wdDeletedTextMarkDoubleStrikeThrough=10         # from enum WdDeletedTextMark
	wdDeletedTextMarkDoubleUnderline=8          # from enum WdDeletedTextMark
	wdDeletedTextMarkHidden       =0          # from enum WdDeletedTextMark
	wdDeletedTextMarkItalic       =6          # from enum WdDeletedTextMark
	wdDeletedTextMarkNone         =4          # from enum WdDeletedTextMark
	wdDeletedTextMarkPound        =3          # from enum WdDeletedTextMark
	wdDeletedTextMarkStrikeThrough=1          # from enum WdDeletedTextMark
	wdDeletedTextMarkUnderline    =7          # from enum WdDeletedTextMark
	wdDiacriticColorBidi          =0          # from enum WdDiacriticColor
	wdDiacriticColorLatin         =1          # from enum WdDiacriticColor
	wdGrammar                     =1          # from enum WdDictionaryType
	wdHangulHanjaConversion       =8          # from enum WdDictionaryType
	wdHangulHanjaConversionCustom =9          # from enum WdDictionaryType
	wdHyphenation                 =3          # from enum WdDictionaryType
	wdSpelling                    =0          # from enum WdDictionaryType
	wdSpellingComplete            =4          # from enum WdDictionaryType
	wdSpellingCustom              =5          # from enum WdDictionaryType
	wdSpellingLegal               =6          # from enum WdDictionaryType
	wdSpellingMedical             =7          # from enum WdDictionaryType
	wdThesaurus                   =2          # from enum WdDictionaryType
	emptyenum                     =0          # from enum WdDictionaryTypeHID
	wd70                          =0          # from enum WdDisableFeaturesIntroducedAfter
	wd70FE                        =1          # from enum WdDisableFeaturesIntroducedAfter
	wd80                          =2          # from enum WdDisableFeaturesIntroducedAfter
	wdInsertContent               =0          # from enum WdDocPartInsertOptions
	wdInsertPage                  =2          # from enum WdDocPartInsertOptions
	wdInsertParagraph             =1          # from enum WdDocPartInsertOptions
	wdLeftToRight                 =0          # from enum WdDocumentDirection
	wdRightToLeft                 =1          # from enum WdDocumentDirection
	wdDocumentEmail               =2          # from enum WdDocumentKind
	wdDocumentLetter              =1          # from enum WdDocumentKind
	wdDocumentNotSpecified        =0          # from enum WdDocumentKind
	wdDocument                    =1          # from enum WdDocumentMedium
	wdEmailMessage                =0          # from enum WdDocumentMedium
	wdWebPage                     =2          # from enum WdDocumentMedium
	wdTypeDocument                =0          # from enum WdDocumentType
	wdTypeFrameset                =2          # from enum WdDocumentType
	wdTypeTemplate                =1          # from enum WdDocumentType
	wdDocumentViewLtr             =1          # from enum WdDocumentViewDirection
	wdDocumentViewRtl             =0          # from enum WdDocumentViewDirection
	wdDropMargin                  =2          # from enum WdDropPosition
	wdDropNone                    =0          # from enum WdDropPosition
	wdDropNormal                  =1          # from enum WdDropPosition
	wdAutomaticUpdate             =3          # from enum WdEditionOption
	wdCancelPublisher             =0          # from enum WdEditionOption
	wdChangeAttributes            =5          # from enum WdEditionOption
	wdManualUpdate                =4          # from enum WdEditionOption
	wdOpenSource                  =7          # from enum WdEditionOption
	wdSelectPublisher             =2          # from enum WdEditionOption
	wdSendPublisher               =1          # from enum WdEditionOption
	wdUpdateSubscriber            =6          # from enum WdEditionOption
	wdPublisher                   =0          # from enum WdEditionType
	wdSubscriber                  =1          # from enum WdEditionType
	wdEditorCurrent               =-6         # from enum WdEditorType
	wdEditorEditors               =-5         # from enum WdEditorType
	wdEditorEveryone              =-1         # from enum WdEditorType
	wdEditorOwners                =-4         # from enum WdEditorType
	wdEmailHTMLFidelityHigh       =3          # from enum WdEmailHTMLFidelity
	wdEmailHTMLFidelityLow        =1          # from enum WdEmailHTMLFidelity
	wdEmailHTMLFidelityMedium     =2          # from enum WdEmailHTMLFidelity
	wdEmphasisMarkNone            =0          # from enum WdEmphasisMark
	wdEmphasisMarkOverComma       =2          # from enum WdEmphasisMark
	wdEmphasisMarkOverSolidCircle =1          # from enum WdEmphasisMark
	wdEmphasisMarkOverWhiteCircle =3          # from enum WdEmphasisMark
	wdEmphasisMarkUnderSolidCircle=4          # from enum WdEmphasisMark
	wdCancelDisabled              =0          # from enum WdEnableCancelKey
	wdCancelInterrupt             =1          # from enum WdEnableCancelKey
	wdEncloseStyleLarge           =2          # from enum WdEncloseStyle
	wdEncloseStyleNone            =0          # from enum WdEncloseStyle
	wdEncloseStyleSmall           =1          # from enum WdEncloseStyle
	wdEnclosureCircle             =0          # from enum WdEnclosureType
	wdEnclosureDiamond            =3          # from enum WdEnclosureType
	wdEnclosureSquare             =1          # from enum WdEnclosureType
	wdEnclosureTriangle           =2          # from enum WdEnclosureType
	wdEndOfDocument               =1          # from enum WdEndnoteLocation
	wdEndOfSection                =0          # from enum WdEndnoteLocation
	wdCenterClockwise             =7          # from enum WdEnvelopeOrientation
	wdCenterLandscape             =4          # from enum WdEnvelopeOrientation
	wdCenterPortrait              =1          # from enum WdEnvelopeOrientation
	wdLeftClockwise               =6          # from enum WdEnvelopeOrientation
	wdLeftLandscape               =3          # from enum WdEnvelopeOrientation
	wdLeftPortrait                =0          # from enum WdEnvelopeOrientation
	wdRightClockwise              =8          # from enum WdEnvelopeOrientation
	wdRightLandscape              =5          # from enum WdEnvelopeOrientation
	wdRightPortrait               =2          # from enum WdEnvelopeOrientation
	wdExportCreateHeadingBookmarks=1          # from enum WdExportCreateBookmarks
	wdExportCreateNoBookmarks     =0          # from enum WdExportCreateBookmarks
	wdExportCreateWordBookmarks   =2          # from enum WdExportCreateBookmarks
	wdExportFormatPDF             =17         # from enum WdExportFormat
	wdExportFormatXPS             =18         # from enum WdExportFormat
	wdExportDocumentContent       =0          # from enum WdExportItem
	wdExportDocumentWithMarkup    =7          # from enum WdExportItem
	wdExportOptimizeForOnScreen   =1          # from enum WdExportOptimizeFor
	wdExportOptimizeForPrint      =0          # from enum WdExportOptimizeFor
	wdExportAllDocument           =0          # from enum WdExportRange
	wdExportCurrentPage           =2          # from enum WdExportRange
	wdExportFromTo                =3          # from enum WdExportRange
	wdExportSelection             =1          # from enum WdExportRange
	wdLineBreakJapanese           =1041       # from enum WdFarEastLineBreakLanguageID
	wdLineBreakKorean             =1042       # from enum WdFarEastLineBreakLanguageID
	wdLineBreakSimplifiedChinese  =2052       # from enum WdFarEastLineBreakLanguageID
	wdLineBreakTraditionalChinese =1028       # from enum WdFarEastLineBreakLanguageID
	wdFarEastLineBreakLevelCustom =2          # from enum WdFarEastLineBreakLevel
	wdFarEastLineBreakLevelNormal =0          # from enum WdFarEastLineBreakLevel
	wdFarEastLineBreakLevelStrict =1          # from enum WdFarEastLineBreakLevel
	wdFieldKindCold               =3          # from enum WdFieldKind
	wdFieldKindHot                =1          # from enum WdFieldKind
	wdFieldKindNone               =0          # from enum WdFieldKind
	wdFieldKindWarm               =2          # from enum WdFieldKind
	wdFieldShadingAlways          =1          # from enum WdFieldShading
	wdFieldShadingNever           =0          # from enum WdFieldShading
	wdFieldShadingWhenSelected    =2          # from enum WdFieldShading
	wdFieldAddin                  =81         # from enum WdFieldType
	wdFieldAddressBlock           =93         # from enum WdFieldType
	wdFieldAdvance                =84         # from enum WdFieldType
	wdFieldAsk                    =38         # from enum WdFieldType
	wdFieldAuthor                 =17         # from enum WdFieldType
	wdFieldAutoNum                =54         # from enum WdFieldType
	wdFieldAutoNumLegal           =53         # from enum WdFieldType
	wdFieldAutoNumOutline         =52         # from enum WdFieldType
	wdFieldAutoText               =79         # from enum WdFieldType
	wdFieldAutoTextList           =89         # from enum WdFieldType
	wdFieldBarCode                =63         # from enum WdFieldType
	wdFieldBibliography           =97         # from enum WdFieldType
	wdFieldBidiOutline            =92         # from enum WdFieldType
	wdFieldCitation               =96         # from enum WdFieldType
	wdFieldComments               =19         # from enum WdFieldType
	wdFieldCompare                =80         # from enum WdFieldType
	wdFieldCreateDate             =21         # from enum WdFieldType
	wdFieldDDE                    =45         # from enum WdFieldType
	wdFieldDDEAuto                =46         # from enum WdFieldType
	wdFieldData                   =40         # from enum WdFieldType
	wdFieldDatabase               =78         # from enum WdFieldType
	wdFieldDate                   =31         # from enum WdFieldType
	wdFieldDocProperty            =85         # from enum WdFieldType
	wdFieldDocVariable            =64         # from enum WdFieldType
	wdFieldEditTime               =25         # from enum WdFieldType
	wdFieldEmbed                  =58         # from enum WdFieldType
	wdFieldEmpty                  =-1         # from enum WdFieldType
	wdFieldExpression             =34         # from enum WdFieldType
	wdFieldFileName               =29         # from enum WdFieldType
	wdFieldFileSize               =69         # from enum WdFieldType
	wdFieldFillIn                 =39         # from enum WdFieldType
	wdFieldFootnoteRef            =5          # from enum WdFieldType
	wdFieldFormCheckBox           =71         # from enum WdFieldType
	wdFieldFormDropDown           =83         # from enum WdFieldType
	wdFieldFormTextInput          =70         # from enum WdFieldType
	wdFieldFormula                =49         # from enum WdFieldType
	wdFieldGlossary               =47         # from enum WdFieldType
	wdFieldGoToButton             =50         # from enum WdFieldType
	wdFieldGreetingLine           =94         # from enum WdFieldType
	wdFieldHTMLActiveX            =91         # from enum WdFieldType
	wdFieldHyperlink              =88         # from enum WdFieldType
	wdFieldIf                     =7          # from enum WdFieldType
	wdFieldImport                 =55         # from enum WdFieldType
	wdFieldInclude                =36         # from enum WdFieldType
	wdFieldIncludePicture         =67         # from enum WdFieldType
	wdFieldIncludeText            =68         # from enum WdFieldType
	wdFieldIndex                  =8          # from enum WdFieldType
	wdFieldIndexEntry             =4          # from enum WdFieldType
	wdFieldInfo                   =14         # from enum WdFieldType
	wdFieldKeyWord                =18         # from enum WdFieldType
	wdFieldLastSavedBy            =20         # from enum WdFieldType
	wdFieldLink                   =56         # from enum WdFieldType
	wdFieldListNum                =90         # from enum WdFieldType
	wdFieldMacroButton            =51         # from enum WdFieldType
	wdFieldMergeField             =59         # from enum WdFieldType
	wdFieldMergeRec               =44         # from enum WdFieldType
	wdFieldMergeSeq               =75         # from enum WdFieldType
	wdFieldNext                   =41         # from enum WdFieldType
	wdFieldNextIf                 =42         # from enum WdFieldType
	wdFieldNoteRef                =72         # from enum WdFieldType
	wdFieldNumChars               =28         # from enum WdFieldType
	wdFieldNumPages               =26         # from enum WdFieldType
	wdFieldNumWords               =27         # from enum WdFieldType
	wdFieldOCX                    =87         # from enum WdFieldType
	wdFieldPage                   =33         # from enum WdFieldType
	wdFieldPageRef                =37         # from enum WdFieldType
	wdFieldPrint                  =48         # from enum WdFieldType
	wdFieldPrintDate              =23         # from enum WdFieldType
	wdFieldPrivate                =77         # from enum WdFieldType
	wdFieldQuote                  =35         # from enum WdFieldType
	wdFieldRef                    =3          # from enum WdFieldType
	wdFieldRefDoc                 =11         # from enum WdFieldType
	wdFieldRevisionNum            =24         # from enum WdFieldType
	wdFieldSaveDate               =22         # from enum WdFieldType
	wdFieldSection                =65         # from enum WdFieldType
	wdFieldSectionPages           =66         # from enum WdFieldType
	wdFieldSequence               =12         # from enum WdFieldType
	wdFieldSet                    =6          # from enum WdFieldType
	wdFieldShape                  =95         # from enum WdFieldType
	wdFieldSkipIf                 =43         # from enum WdFieldType
	wdFieldStyleRef               =10         # from enum WdFieldType
	wdFieldSubject                =16         # from enum WdFieldType
	wdFieldSubscriber             =82         # from enum WdFieldType
	wdFieldSymbol                 =57         # from enum WdFieldType
	wdFieldTOA                    =73         # from enum WdFieldType
	wdFieldTOAEntry               =74         # from enum WdFieldType
	wdFieldTOC                    =13         # from enum WdFieldType
	wdFieldTOCEntry               =9          # from enum WdFieldType
	wdFieldTemplate               =30         # from enum WdFieldType
	wdFieldTime                   =32         # from enum WdFieldType
	wdFieldTitle                  =15         # from enum WdFieldType
	wdFieldUserAddress            =62         # from enum WdFieldType
	wdFieldUserInitials           =61         # from enum WdFieldType
	wdFieldUserName               =60         # from enum WdFieldType
	wdMatchAnyCharacter           =65599      # from enum WdFindMatch
	wdMatchAnyDigit               =65567      # from enum WdFindMatch
	wdMatchAnyLetter              =65583      # from enum WdFindMatch
	wdMatchCaretCharacter         =11         # from enum WdFindMatch
	wdMatchColumnBreak            =14         # from enum WdFindMatch
	wdMatchCommentMark            =5          # from enum WdFindMatch
	wdMatchEmDash                 =8212       # from enum WdFindMatch
	wdMatchEnDash                 =8211       # from enum WdFindMatch
	wdMatchEndnoteMark            =65555      # from enum WdFindMatch
	wdMatchField                  =19         # from enum WdFindMatch
	wdMatchFootnoteMark           =65554      # from enum WdFindMatch
	wdMatchGraphic                =1          # from enum WdFindMatch
	wdMatchManualLineBreak        =65551      # from enum WdFindMatch
	wdMatchManualPageBreak        =65564      # from enum WdFindMatch
	wdMatchNonbreakingHyphen      =30         # from enum WdFindMatch
	wdMatchNonbreakingSpace       =160        # from enum WdFindMatch
	wdMatchOptionalHyphen         =31         # from enum WdFindMatch
	wdMatchParagraphMark          =65551      # from enum WdFindMatch
	wdMatchSectionBreak           =65580      # from enum WdFindMatch
	wdMatchTabCharacter           =9          # from enum WdFindMatch
	wdMatchWhiteSpace             =65655      # from enum WdFindMatch
	wdFindAsk                     =2          # from enum WdFindWrap
	wdFindContinue                =1          # from enum WdFindWrap
	wdFindStop                    =0          # from enum WdFindWrap
	wdFlowLtr                     =0          # from enum WdFlowDirection
	wdFlowRtl                     =1          # from enum WdFlowDirection
	wdFontBiasDefault             =0          # from enum WdFontBias
	wdFontBiasDontCare            =255        # from enum WdFontBias
	wdFontBiasFareast             =1          # from enum WdFontBias
	wdBeneathText                 =1          # from enum WdFootnoteLocation
	wdBottomOfPage                =0          # from enum WdFootnoteLocation
	wdFrameBottom                 =-999997    # from enum WdFramePosition
	wdFrameCenter                 =-999995    # from enum WdFramePosition
	wdFrameInside                 =-999994    # from enum WdFramePosition
	wdFrameLeft                   =-999998    # from enum WdFramePosition
	wdFrameOutside                =-999993    # from enum WdFramePosition
	wdFrameRight                  =-999996    # from enum WdFramePosition
	wdFrameTop                    =-999999    # from enum WdFramePosition
	wdFrameAtLeast                =1          # from enum WdFrameSizeRule
	wdFrameAuto                   =0          # from enum WdFrameSizeRule
	wdFrameExact                  =2          # from enum WdFrameSizeRule
	wdFramesetNewFrameAbove       =0          # from enum WdFramesetNewFrameLocation
	wdFramesetNewFrameBelow       =1          # from enum WdFramesetNewFrameLocation
	wdFramesetNewFrameLeft        =3          # from enum WdFramesetNewFrameLocation
	wdFramesetNewFrameRight       =2          # from enum WdFramesetNewFrameLocation
	wdFramesetSizeTypeFixed       =1          # from enum WdFramesetSizeType
	wdFramesetSizeTypePercent     =0          # from enum WdFramesetSizeType
	wdFramesetSizeTypeRelative    =2          # from enum WdFramesetSizeType
	wdFramesetTypeFrame           =1          # from enum WdFramesetType
	wdFramesetTypeFrameset        =0          # from enum WdFramesetType
	wdFrenchBoth                  =0          # from enum WdFrenchSpeller
	wdFrenchPostReform            =2          # from enum WdFrenchSpeller
	wdFrenchPreReform             =1          # from enum WdFrenchSpeller
	wdGoToAbsolute                =1          # from enum WdGoToDirection
	wdGoToFirst                   =1          # from enum WdGoToDirection
	wdGoToLast                    =-1         # from enum WdGoToDirection
	wdGoToNext                    =2          # from enum WdGoToDirection
	wdGoToPrevious                =3          # from enum WdGoToDirection
	wdGoToRelative                =2          # from enum WdGoToDirection
	wdGoToBookmark                =-1         # from enum WdGoToItem
	wdGoToComment                 =6          # from enum WdGoToItem
	wdGoToEndnote                 =5          # from enum WdGoToItem
	wdGoToEquation                =10         # from enum WdGoToItem
	wdGoToField                   =7          # from enum WdGoToItem
	wdGoToFootnote                =4          # from enum WdGoToItem
	wdGoToGrammaticalError        =14         # from enum WdGoToItem
	wdGoToGraphic                 =8          # from enum WdGoToItem
	wdGoToHeading                 =11         # from enum WdGoToItem
	wdGoToLine                    =3          # from enum WdGoToItem
	wdGoToObject                  =9          # from enum WdGoToItem
	wdGoToPage                    =1          # from enum WdGoToItem
	wdGoToPercent                 =12         # from enum WdGoToItem
	wdGoToProofreadingError       =15         # from enum WdGoToItem
	wdGoToSection                 =0          # from enum WdGoToItem
	wdGoToSpellingError           =13         # from enum WdGoToItem
	wdGoToTable                   =2          # from enum WdGoToItem
	wdGranularityCharLevel        =0          # from enum WdGranularity
	wdGranularityWordLevel        =1          # from enum WdGranularity
	wdGutterPosLeft               =0          # from enum WdGutterStyle
	wdGutterPosRight              =2          # from enum WdGutterStyle
	wdGutterPosTop                =1          # from enum WdGutterStyle
	wdGutterStyleBidi             =2          # from enum WdGutterStyleOld
	wdGutterStyleLatin            =-10        # from enum WdGutterStyleOld
	wdHeaderFooterEvenPages       =3          # from enum WdHeaderFooterIndex
	wdHeaderFooterFirstPage       =2          # from enum WdHeaderFooterIndex
	wdHeaderFooterPrimary         =1          # from enum WdHeaderFooterIndex
	wdHeadingSeparatorBlankLine   =1          # from enum WdHeadingSeparator
	wdHeadingSeparatorLetter      =2          # from enum WdHeadingSeparator
	wdHeadingSeparatorLetterFull  =4          # from enum WdHeadingSeparator
	wdHeadingSeparatorLetterLow   =3          # from enum WdHeadingSeparator
	wdHeadingSeparatorNone        =0          # from enum WdHeadingSeparator
	wdFullScript                  =0          # from enum WdHebSpellStart
	wdMixedAuthorizedScript       =3          # from enum WdHebSpellStart
	wdMixedScript                 =2          # from enum WdHebSpellStart
	wdPartialScript               =1          # from enum WdHebSpellStart
	wdHelp                        =0          # from enum WdHelpType
	wdHelpAbout                   =1          # from enum WdHelpType
	wdHelpActiveWindow            =2          # from enum WdHelpType
	wdHelpContents                =3          # from enum WdHelpType
	wdHelpExamplesAndDemos        =4          # from enum WdHelpType
	wdHelpHWP                     =13         # from enum WdHelpType
	wdHelpIchitaro                =11         # from enum WdHelpType
	wdHelpIndex                   =5          # from enum WdHelpType
	wdHelpKeyboard                =6          # from enum WdHelpType
	wdHelpPE2                     =12         # from enum WdHelpType
	wdHelpPSSHelp                 =7          # from enum WdHelpType
	wdHelpQuickPreview            =8          # from enum WdHelpType
	wdHelpSearch                  =9          # from enum WdHelpType
	wdHelpUsingHelp               =10         # from enum WdHelpType
	emptyenum                     =0          # from enum WdHelpTypeHID
	wdAutoDetectHighAnsiFarEast   =2          # from enum WdHighAnsiText
	wdHighAnsiIsFarEast           =0          # from enum WdHighAnsiText
	wdHighAnsiIsHighAnsi          =1          # from enum WdHighAnsiText
	wdHorizontalInVerticalFitInLine=1          # from enum WdHorizontalInVerticalType
	wdHorizontalInVerticalNone    =0          # from enum WdHorizontalInVerticalType
	wdHorizontalInVerticalResizeLine=2          # from enum WdHorizontalInVerticalType
	wdHorizontalLineAlignCenter   =1          # from enum WdHorizontalLineAlignment
	wdHorizontalLineAlignLeft     =0          # from enum WdHorizontalLineAlignment
	wdHorizontalLineAlignRight    =2          # from enum WdHorizontalLineAlignment
	wdHorizontalLineFixedWidth    =-2         # from enum WdHorizontalLineWidthType
	wdHorizontalLinePercentWidth  =-1         # from enum WdHorizontalLineWidthType
	wdIMEModeAlpha                =8          # from enum WdIMEMode
	wdIMEModeAlphaFull            =7          # from enum WdIMEMode
	wdIMEModeHangul               =10         # from enum WdIMEMode
	wdIMEModeHangulFull           =9          # from enum WdIMEMode
	wdIMEModeHiragana             =4          # from enum WdIMEMode
	wdIMEModeKatakana             =5          # from enum WdIMEMode
	wdIMEModeKatakanaHalf         =6          # from enum WdIMEMode
	wdIMEModeNoControl            =0          # from enum WdIMEMode
	wdIMEModeOff                  =2          # from enum WdIMEMode
	wdIMEModeOn                   =1          # from enum WdIMEMode
	wdIndexFilterAiueo            =1          # from enum WdIndexFilter
	wdIndexFilterAkasatana        =2          # from enum WdIndexFilter
	wdIndexFilterChosung          =3          # from enum WdIndexFilter
	wdIndexFilterFull             =6          # from enum WdIndexFilter
	wdIndexFilterLow              =4          # from enum WdIndexFilter
	wdIndexFilterMedium           =5          # from enum WdIndexFilter
	wdIndexFilterNone             =0          # from enum WdIndexFilter
	wdIndexBulleted               =4          # from enum WdIndexFormat
	wdIndexClassic                =1          # from enum WdIndexFormat
	wdIndexFancy                  =2          # from enum WdIndexFormat
	wdIndexFormal                 =5          # from enum WdIndexFormat
	wdIndexModern                 =3          # from enum WdIndexFormat
	wdIndexSimple                 =6          # from enum WdIndexFormat
	wdIndexTemplate               =0          # from enum WdIndexFormat
	wdIndexSortByStroke           =0          # from enum WdIndexSortBy
	wdIndexSortBySyllable         =1          # from enum WdIndexSortBy
	wdIndexIndent                 =0          # from enum WdIndexType
	wdIndexRunin                  =1          # from enum WdIndexType
	wdActiveEndAdjustedPageNumber =1          # from enum WdInformation
	wdActiveEndPageNumber         =3          # from enum WdInformation
	wdActiveEndSectionNumber      =2          # from enum WdInformation
	wdAtEndOfRowMarker            =31         # from enum WdInformation
	wdCapsLock                    =21         # from enum WdInformation
	wdEndOfRangeColumnNumber      =17         # from enum WdInformation
	wdEndOfRangeRowNumber         =14         # from enum WdInformation
	wdFirstCharacterColumnNumber  =9          # from enum WdInformation
	wdFirstCharacterLineNumber    =10         # from enum WdInformation
	wdFrameIsSelected             =11         # from enum WdInformation
	wdHeaderFooterType            =33         # from enum WdInformation
	wdHorizontalPositionRelativeToPage=5          # from enum WdInformation
	wdHorizontalPositionRelativeToTextBoundary=7          # from enum WdInformation
	wdInClipboard                 =38         # from enum WdInformation
	wdInCommentPane               =26         # from enum WdInformation
	wdInEndnote                   =36         # from enum WdInformation
	wdInFootnote                  =35         # from enum WdInformation
	wdInFootnoteEndnotePane       =25         # from enum WdInformation
	wdInHeaderFooter              =28         # from enum WdInformation
	wdInMasterDocument            =34         # from enum WdInformation
	wdInWordMail                  =37         # from enum WdInformation
	wdMaximumNumberOfColumns      =18         # from enum WdInformation
	wdMaximumNumberOfRows         =15         # from enum WdInformation
	wdNumLock                     =22         # from enum WdInformation
	wdNumberOfPagesInDocument     =4          # from enum WdInformation
	wdOverType                    =23         # from enum WdInformation
	wdReferenceOfType             =32         # from enum WdInformation
	wdRevisionMarking             =24         # from enum WdInformation
	wdSelectionMode               =20         # from enum WdInformation
	wdStartOfRangeColumnNumber    =16         # from enum WdInformation
	wdStartOfRangeRowNumber       =13         # from enum WdInformation
	wdVerticalPositionRelativeToPage=6          # from enum WdInformation
	wdVerticalPositionRelativeToTextBoundary=8          # from enum WdInformation
	wdWithInTable                 =12         # from enum WdInformation
	wdZoomPercentage              =19         # from enum WdInformation
	wdInlineShapeChart            =12         # from enum WdInlineShapeType
	wdInlineShapeDiagram          =13         # from enum WdInlineShapeType
	wdInlineShapeEmbeddedOLEObject=1          # from enum WdInlineShapeType
	wdInlineShapeHorizontalLine   =6          # from enum WdInlineShapeType
	wdInlineShapeLinkedOLEObject  =2          # from enum WdInlineShapeType
	wdInlineShapeLinkedPicture    =4          # from enum WdInlineShapeType
	wdInlineShapeLinkedPictureHorizontalLine=8          # from enum WdInlineShapeType
	wdInlineShapeLockedCanvas     =14         # from enum WdInlineShapeType
	wdInlineShapeOLEControlObject =5          # from enum WdInlineShapeType
	wdInlineShapeOWSAnchor        =11         # from enum WdInlineShapeType
	wdInlineShapePicture          =3          # from enum WdInlineShapeType
	wdInlineShapePictureBullet    =9          # from enum WdInlineShapeType
	wdInlineShapePictureHorizontalLine=7          # from enum WdInlineShapeType
	wdInlineShapeScriptAnchor     =10         # from enum WdInlineShapeType
	wdInlineShapeSmartArt         =15         # from enum WdInlineShapeType
	wdInsertCellsEntireColumn     =3          # from enum WdInsertCells
	wdInsertCellsEntireRow        =2          # from enum WdInsertCells
	wdInsertCellsShiftDown        =1          # from enum WdInsertCells
	wdInsertCellsShiftRight       =0          # from enum WdInsertCells
	wdInsertedTextMarkBold        =1          # from enum WdInsertedTextMark
	wdInsertedTextMarkColorOnly   =5          # from enum WdInsertedTextMark
	wdInsertedTextMarkDoubleStrikeThrough=7          # from enum WdInsertedTextMark
	wdInsertedTextMarkDoubleUnderline=4          # from enum WdInsertedTextMark
	wdInsertedTextMarkItalic      =2          # from enum WdInsertedTextMark
	wdInsertedTextMarkNone        =0          # from enum WdInsertedTextMark
	wdInsertedTextMarkStrikeThrough=6          # from enum WdInsertedTextMark
	wdInsertedTextMarkUnderline   =3          # from enum WdInsertedTextMark
	wd24HourClock                 =21         # from enum WdInternationalIndex
	wdCurrencyCode                =20         # from enum WdInternationalIndex
	wdDateSeparator               =25         # from enum WdInternationalIndex
	wdDecimalSeparator            =18         # from enum WdInternationalIndex
	wdInternationalAM             =22         # from enum WdInternationalIndex
	wdInternationalPM             =23         # from enum WdInternationalIndex
	wdListSeparator               =17         # from enum WdInternationalIndex
	wdProductLanguageID           =26         # from enum WdInternationalIndex
	wdThousandsSeparator          =19         # from enum WdInternationalIndex
	wdTimeSeparator               =24         # from enum WdInternationalIndex
	wdJustificationModeCompress   =1          # from enum WdJustificationMode
	wdJustificationModeCompressKana=2          # from enum WdJustificationMode
	wdJustificationModeExpand     =0          # from enum WdJustificationMode
	wdKanaHiragana                =9          # from enum WdKana
	wdKanaKatakana                =8          # from enum WdKana
	wdKey0                        =48         # from enum WdKey
	wdKey1                        =49         # from enum WdKey
	wdKey2                        =50         # from enum WdKey
	wdKey3                        =51         # from enum WdKey
	wdKey4                        =52         # from enum WdKey
	wdKey5                        =53         # from enum WdKey
	wdKey6                        =54         # from enum WdKey
	wdKey7                        =55         # from enum WdKey
	wdKey8                        =56         # from enum WdKey
	wdKey9                        =57         # from enum WdKey
	wdKeyA                        =65         # from enum WdKey
	wdKeyAlt                      =1024       # from enum WdKey
	wdKeyB                        =66         # from enum WdKey
	wdKeyBackSingleQuote          =192        # from enum WdKey
	wdKeyBackSlash                =220        # from enum WdKey
	wdKeyBackspace                =8          # from enum WdKey
	wdKeyC                        =67         # from enum WdKey
	wdKeyCloseSquareBrace         =221        # from enum WdKey
	wdKeyComma                    =188        # from enum WdKey
	wdKeyCommand                  =512        # from enum WdKey
	wdKeyControl                  =512        # from enum WdKey
	wdKeyD                        =68         # from enum WdKey
	wdKeyDelete                   =46         # from enum WdKey
	wdKeyE                        =69         # from enum WdKey
	wdKeyEnd                      =35         # from enum WdKey
	wdKeyEquals                   =187        # from enum WdKey
	wdKeyEsc                      =27         # from enum WdKey
	wdKeyF                        =70         # from enum WdKey
	wdKeyF1                       =112        # from enum WdKey
	wdKeyF10                      =121        # from enum WdKey
	wdKeyF11                      =122        # from enum WdKey
	wdKeyF12                      =123        # from enum WdKey
	wdKeyF13                      =124        # from enum WdKey
	wdKeyF14                      =125        # from enum WdKey
	wdKeyF15                      =126        # from enum WdKey
	wdKeyF16                      =127        # from enum WdKey
	wdKeyF2                       =113        # from enum WdKey
	wdKeyF3                       =114        # from enum WdKey
	wdKeyF4                       =115        # from enum WdKey
	wdKeyF5                       =116        # from enum WdKey
	wdKeyF6                       =117        # from enum WdKey
	wdKeyF7                       =118        # from enum WdKey
	wdKeyF8                       =119        # from enum WdKey
	wdKeyF9                       =120        # from enum WdKey
	wdKeyG                        =71         # from enum WdKey
	wdKeyH                        =72         # from enum WdKey
	wdKeyHome                     =36         # from enum WdKey
	wdKeyHyphen                   =189        # from enum WdKey
	wdKeyI                        =73         # from enum WdKey
	wdKeyInsert                   =45         # from enum WdKey
	wdKeyJ                        =74         # from enum WdKey
	wdKeyK                        =75         # from enum WdKey
	wdKeyL                        =76         # from enum WdKey
	wdKeyM                        =77         # from enum WdKey
	wdKeyN                        =78         # from enum WdKey
	wdKeyNumeric0                 =96         # from enum WdKey
	wdKeyNumeric1                 =97         # from enum WdKey
	wdKeyNumeric2                 =98         # from enum WdKey
	wdKeyNumeric3                 =99         # from enum WdKey
	wdKeyNumeric4                 =100        # from enum WdKey
	wdKeyNumeric5                 =101        # from enum WdKey
	wdKeyNumeric5Special          =12         # from enum WdKey
	wdKeyNumeric6                 =102        # from enum WdKey
	wdKeyNumeric7                 =103        # from enum WdKey
	wdKeyNumeric8                 =104        # from enum WdKey
	wdKeyNumeric9                 =105        # from enum WdKey
	wdKeyNumericAdd               =107        # from enum WdKey
	wdKeyNumericDecimal           =110        # from enum WdKey
	wdKeyNumericDivide            =111        # from enum WdKey
	wdKeyNumericMultiply          =106        # from enum WdKey
	wdKeyNumericSubtract          =109        # from enum WdKey
	wdKeyO                        =79         # from enum WdKey
	wdKeyOpenSquareBrace          =219        # from enum WdKey
	wdKeyOption                   =1024       # from enum WdKey
	wdKeyP                        =80         # from enum WdKey
	wdKeyPageDown                 =34         # from enum WdKey
	wdKeyPageUp                   =33         # from enum WdKey
	wdKeyPause                    =19         # from enum WdKey
	wdKeyPeriod                   =190        # from enum WdKey
	wdKeyQ                        =81         # from enum WdKey
	wdKeyR                        =82         # from enum WdKey
	wdKeyReturn                   =13         # from enum WdKey
	wdKeyS                        =83         # from enum WdKey
	wdKeyScrollLock               =145        # from enum WdKey
	wdKeySemiColon                =186        # from enum WdKey
	wdKeyShift                    =256        # from enum WdKey
	wdKeySingleQuote              =222        # from enum WdKey
	wdKeySlash                    =191        # from enum WdKey
	wdKeySpacebar                 =32         # from enum WdKey
	wdKeyT                        =84         # from enum WdKey
	wdKeyTab                      =9          # from enum WdKey
	wdKeyU                        =85         # from enum WdKey
	wdKeyV                        =86         # from enum WdKey
	wdKeyW                        =87         # from enum WdKey
	wdKeyX                        =88         # from enum WdKey
	wdKeyY                        =89         # from enum WdKey
	wdKeyZ                        =90         # from enum WdKey
	wdNoKey                       =255        # from enum WdKey
	wdKeyCategoryAutoText         =4          # from enum WdKeyCategory
	wdKeyCategoryCommand          =1          # from enum WdKeyCategory
	wdKeyCategoryDisable          =0          # from enum WdKeyCategory
	wdKeyCategoryFont             =3          # from enum WdKeyCategory
	wdKeyCategoryMacro            =2          # from enum WdKeyCategory
	wdKeyCategoryNil              =-1         # from enum WdKeyCategory
	wdKeyCategoryPrefix           =7          # from enum WdKeyCategory
	wdKeyCategoryStyle            =5          # from enum WdKeyCategory
	wdKeyCategorySymbol           =6          # from enum WdKeyCategory
	wdAfrikaans                   =1078       # from enum WdLanguageID
	wdAlbanian                    =1052       # from enum WdLanguageID
	wdAmharic                     =1118       # from enum WdLanguageID
	wdArabic                      =1025       # from enum WdLanguageID
	wdArabicAlgeria               =5121       # from enum WdLanguageID
	wdArabicBahrain               =15361      # from enum WdLanguageID
	wdArabicEgypt                 =3073       # from enum WdLanguageID
	wdArabicIraq                  =2049       # from enum WdLanguageID
	wdArabicJordan                =11265      # from enum WdLanguageID
	wdArabicKuwait                =13313      # from enum WdLanguageID
	wdArabicLebanon               =12289      # from enum WdLanguageID
	wdArabicLibya                 =4097       # from enum WdLanguageID
	wdArabicMorocco               =6145       # from enum WdLanguageID
	wdArabicOman                  =8193       # from enum WdLanguageID
	wdArabicQatar                 =16385      # from enum WdLanguageID
	wdArabicSyria                 =10241      # from enum WdLanguageID
	wdArabicTunisia               =7169       # from enum WdLanguageID
	wdArabicUAE                   =14337      # from enum WdLanguageID
	wdArabicYemen                 =9217       # from enum WdLanguageID
	wdArmenian                    =1067       # from enum WdLanguageID
	wdAssamese                    =1101       # from enum WdLanguageID
	wdAzeriCyrillic               =2092       # from enum WdLanguageID
	wdAzeriLatin                  =1068       # from enum WdLanguageID
	wdBasque                      =1069       # from enum WdLanguageID
	wdBelgianDutch                =2067       # from enum WdLanguageID
	wdBelgianFrench               =2060       # from enum WdLanguageID
	wdBengali                     =1093       # from enum WdLanguageID
	wdBulgarian                   =1026       # from enum WdLanguageID
	wdBurmese                     =1109       # from enum WdLanguageID
	wdByelorussian                =1059       # from enum WdLanguageID
	wdCatalan                     =1027       # from enum WdLanguageID
	wdCherokee                    =1116       # from enum WdLanguageID
	wdChineseHongKongSAR          =3076       # from enum WdLanguageID
	wdChineseMacaoSAR             =5124       # from enum WdLanguageID
	wdChineseSingapore            =4100       # from enum WdLanguageID
	wdCroatian                    =1050       # from enum WdLanguageID
	wdCzech                       =1029       # from enum WdLanguageID
	wdDanish                      =1030       # from enum WdLanguageID
	wdDivehi                      =1125       # from enum WdLanguageID
	wdDutch                       =1043       # from enum WdLanguageID
	wdEdo                         =1126       # from enum WdLanguageID
	wdEnglishAUS                  =3081       # from enum WdLanguageID
	wdEnglishBelize               =10249      # from enum WdLanguageID
	wdEnglishCanadian             =4105       # from enum WdLanguageID
	wdEnglishCaribbean            =9225       # from enum WdLanguageID
	wdEnglishIndonesia            =14345      # from enum WdLanguageID
	wdEnglishIreland              =6153       # from enum WdLanguageID
	wdEnglishJamaica              =8201       # from enum WdLanguageID
	wdEnglishNewZealand           =5129       # from enum WdLanguageID
	wdEnglishPhilippines          =13321      # from enum WdLanguageID
	wdEnglishSouthAfrica          =7177       # from enum WdLanguageID
	wdEnglishTrinidadTobago       =11273      # from enum WdLanguageID
	wdEnglishUK                   =2057       # from enum WdLanguageID
	wdEnglishUS                   =1033       # from enum WdLanguageID
	wdEnglishZimbabwe             =12297      # from enum WdLanguageID
	wdEstonian                    =1061       # from enum WdLanguageID
	wdFaeroese                    =1080       # from enum WdLanguageID
	wdFilipino                    =1124       # from enum WdLanguageID
	wdFinnish                     =1035       # from enum WdLanguageID
	wdFrench                      =1036       # from enum WdLanguageID
	wdFrenchCameroon              =11276      # from enum WdLanguageID
	wdFrenchCanadian              =3084       # from enum WdLanguageID
	wdFrenchCongoDRC              =9228       # from enum WdLanguageID
	wdFrenchCotedIvoire           =12300      # from enum WdLanguageID
	wdFrenchHaiti                 =15372      # from enum WdLanguageID
	wdFrenchLuxembourg            =5132       # from enum WdLanguageID
	wdFrenchMali                  =13324      # from enum WdLanguageID
	wdFrenchMonaco                =6156       # from enum WdLanguageID
	wdFrenchMorocco               =14348      # from enum WdLanguageID
	wdFrenchReunion               =8204       # from enum WdLanguageID
	wdFrenchSenegal               =10252      # from enum WdLanguageID
	wdFrenchWestIndies            =7180       # from enum WdLanguageID
	wdFrisianNetherlands          =1122       # from enum WdLanguageID
	wdFulfulde                    =1127       # from enum WdLanguageID
	wdGaelicIreland               =2108       # from enum WdLanguageID
	wdGaelicScotland              =1084       # from enum WdLanguageID
	wdGalician                    =1110       # from enum WdLanguageID
	wdGeorgian                    =1079       # from enum WdLanguageID
	wdGerman                      =1031       # from enum WdLanguageID
	wdGermanAustria               =3079       # from enum WdLanguageID
	wdGermanLiechtenstein         =5127       # from enum WdLanguageID
	wdGermanLuxembourg            =4103       # from enum WdLanguageID
	wdGreek                       =1032       # from enum WdLanguageID
	wdGuarani                     =1140       # from enum WdLanguageID
	wdGujarati                    =1095       # from enum WdLanguageID
	wdHausa                       =1128       # from enum WdLanguageID
	wdHawaiian                    =1141       # from enum WdLanguageID
	wdHebrew                      =1037       # from enum WdLanguageID
	wdHindi                       =1081       # from enum WdLanguageID
	wdHungarian                   =1038       # from enum WdLanguageID
	wdIbibio                      =1129       # from enum WdLanguageID
	wdIcelandic                   =1039       # from enum WdLanguageID
	wdIgbo                        =1136       # from enum WdLanguageID
	wdIndonesian                  =1057       # from enum WdLanguageID
	wdInuktitut                   =1117       # from enum WdLanguageID
	wdItalian                     =1040       # from enum WdLanguageID
	wdJapanese                    =1041       # from enum WdLanguageID
	wdKannada                     =1099       # from enum WdLanguageID
	wdKanuri                      =1137       # from enum WdLanguageID
	wdKashmiri                    =1120       # from enum WdLanguageID
	wdKazakh                      =1087       # from enum WdLanguageID
	wdKhmer                       =1107       # from enum WdLanguageID
	wdKirghiz                     =1088       # from enum WdLanguageID
	wdKonkani                     =1111       # from enum WdLanguageID
	wdKorean                      =1042       # from enum WdLanguageID
	wdKyrgyz                      =1088       # from enum WdLanguageID
	wdLanguageNone                =0          # from enum WdLanguageID
	wdLao                         =1108       # from enum WdLanguageID
	wdLatin                       =1142       # from enum WdLanguageID
	wdLatvian                     =1062       # from enum WdLanguageID
	wdLithuanian                  =1063       # from enum WdLanguageID
	wdMacedonianFYROM             =1071       # from enum WdLanguageID
	wdMalayBruneiDarussalam       =2110       # from enum WdLanguageID
	wdMalayalam                   =1100       # from enum WdLanguageID
	wdMalaysian                   =1086       # from enum WdLanguageID
	wdMaltese                     =1082       # from enum WdLanguageID
	wdManipuri                    =1112       # from enum WdLanguageID
	wdMarathi                     =1102       # from enum WdLanguageID
	wdMexicanSpanish              =2058       # from enum WdLanguageID
	wdMongolian                   =1104       # from enum WdLanguageID
	wdNepali                      =1121       # from enum WdLanguageID
	wdNoProofing                  =1024       # from enum WdLanguageID
	wdNorwegianBokmol             =1044       # from enum WdLanguageID
	wdNorwegianNynorsk            =2068       # from enum WdLanguageID
	wdOriya                       =1096       # from enum WdLanguageID
	wdOromo                       =1138       # from enum WdLanguageID
	wdPashto                      =1123       # from enum WdLanguageID
	wdPersian                     =1065       # from enum WdLanguageID
	wdPolish                      =1045       # from enum WdLanguageID
	wdPortuguese                  =2070       # from enum WdLanguageID
	wdPortugueseBrazil            =1046       # from enum WdLanguageID
	wdPunjabi                     =1094       # from enum WdLanguageID
	wdRhaetoRomanic               =1047       # from enum WdLanguageID
	wdRomanian                    =1048       # from enum WdLanguageID
	wdRomanianMoldova             =2072       # from enum WdLanguageID
	wdRussian                     =1049       # from enum WdLanguageID
	wdRussianMoldova              =2073       # from enum WdLanguageID
	wdSamiLappish                 =1083       # from enum WdLanguageID
	wdSanskrit                    =1103       # from enum WdLanguageID
	wdSerbianCyrillic             =3098       # from enum WdLanguageID
	wdSerbianLatin                =2074       # from enum WdLanguageID
	wdSesotho                     =1072       # from enum WdLanguageID
	wdSimplifiedChinese           =2052       # from enum WdLanguageID
	wdSindhi                      =1113       # from enum WdLanguageID
	wdSindhiPakistan              =2137       # from enum WdLanguageID
	wdSinhalese                   =1115       # from enum WdLanguageID
	wdSlovak                      =1051       # from enum WdLanguageID
	wdSlovenian                   =1060       # from enum WdLanguageID
	wdSomali                      =1143       # from enum WdLanguageID
	wdSorbian                     =1070       # from enum WdLanguageID
	wdSpanish                     =1034       # from enum WdLanguageID
	wdSpanishArgentina            =11274      # from enum WdLanguageID
	wdSpanishBolivia              =16394      # from enum WdLanguageID
	wdSpanishChile                =13322      # from enum WdLanguageID
	wdSpanishColombia             =9226       # from enum WdLanguageID
	wdSpanishCostaRica            =5130       # from enum WdLanguageID
	wdSpanishDominicanRepublic    =7178       # from enum WdLanguageID
	wdSpanishEcuador              =12298      # from enum WdLanguageID
	wdSpanishElSalvador           =17418      # from enum WdLanguageID
	wdSpanishGuatemala            =4106       # from enum WdLanguageID
	wdSpanishHonduras             =18442      # from enum WdLanguageID
	wdSpanishModernSort           =3082       # from enum WdLanguageID
	wdSpanishNicaragua            =19466      # from enum WdLanguageID
	wdSpanishPanama               =6154       # from enum WdLanguageID
	wdSpanishParaguay             =15370      # from enum WdLanguageID
	wdSpanishPeru                 =10250      # from enum WdLanguageID
	wdSpanishPuertoRico           =20490      # from enum WdLanguageID
	wdSpanishUruguay              =14346      # from enum WdLanguageID
	wdSpanishVenezuela            =8202       # from enum WdLanguageID
	wdSutu                        =1072       # from enum WdLanguageID
	wdSwahili                     =1089       # from enum WdLanguageID
	wdSwedish                     =1053       # from enum WdLanguageID
	wdSwedishFinland              =2077       # from enum WdLanguageID
	wdSwissFrench                 =4108       # from enum WdLanguageID
	wdSwissGerman                 =2055       # from enum WdLanguageID
	wdSwissItalian                =2064       # from enum WdLanguageID
	wdSyriac                      =1114       # from enum WdLanguageID
	wdTajik                       =1064       # from enum WdLanguageID
	wdTamazight                   =1119       # from enum WdLanguageID
	wdTamazightLatin              =2143       # from enum WdLanguageID
	wdTamil                       =1097       # from enum WdLanguageID
	wdTatar                       =1092       # from enum WdLanguageID
	wdTelugu                      =1098       # from enum WdLanguageID
	wdThai                        =1054       # from enum WdLanguageID
	wdTibetan                     =1105       # from enum WdLanguageID
	wdTigrignaEritrea             =2163       # from enum WdLanguageID
	wdTigrignaEthiopic            =1139       # from enum WdLanguageID
	wdTraditionalChinese          =1028       # from enum WdLanguageID
	wdTsonga                      =1073       # from enum WdLanguageID
	wdTswana                      =1074       # from enum WdLanguageID
	wdTurkish                     =1055       # from enum WdLanguageID
	wdTurkmen                     =1090       # from enum WdLanguageID
	wdUkrainian                   =1058       # from enum WdLanguageID
	wdUrdu                        =1056       # from enum WdLanguageID
	wdUzbekCyrillic               =2115       # from enum WdLanguageID
	wdUzbekLatin                  =1091       # from enum WdLanguageID
	wdVenda                       =1075       # from enum WdLanguageID
	wdVietnamese                  =1066       # from enum WdLanguageID
	wdWelsh                       =1106       # from enum WdLanguageID
	wdXhosa                       =1076       # from enum WdLanguageID
	wdYi                          =1144       # from enum WdLanguageID
	wdYiddish                     =1085       # from enum WdLanguageID
	wdYoruba                      =1130       # from enum WdLanguageID
	wdZulu                        =1077       # from enum WdLanguageID
	wdChineseHongKong             =3076       # from enum WdLanguageID2000
	wdChineseMacao                =5124       # from enum WdLanguageID2000
	wdEnglishTrinidad             =11273      # from enum WdLanguageID2000
	wdLayoutModeDefault           =0          # from enum WdLayoutMode
	wdLayoutModeGenko             =3          # from enum WdLayoutMode
	wdLayoutModeGrid              =1          # from enum WdLayoutMode
	wdLayoutModeLineGrid          =2          # from enum WdLayoutMode
	wdFullBlock                   =0          # from enum WdLetterStyle
	wdModifiedBlock               =1          # from enum WdLetterStyle
	wdSemiBlock                   =2          # from enum WdLetterStyle
	wdLetterBottom                =1          # from enum WdLetterheadLocation
	wdLetterLeft                  =2          # from enum WdLetterheadLocation
	wdLetterRight                 =3          # from enum WdLetterheadLocation
	wdLetterTop                   =0          # from enum WdLetterheadLocation
	wdLigaturesAll                =15         # from enum WdLigatures
	wdLigaturesContextual         =2          # from enum WdLigatures
	wdLigaturesContextualDiscretional=10         # from enum WdLigatures
	wdLigaturesContextualHistorical=6          # from enum WdLigatures
	wdLigaturesContextualHistoricalDiscretional=14         # from enum WdLigatures
	wdLigaturesDiscretional       =8          # from enum WdLigatures
	wdLigaturesHistorical         =4          # from enum WdLigatures
	wdLigaturesHistoricalDiscretional=12         # from enum WdLigatures
	wdLigaturesNone               =0          # from enum WdLigatures
	wdLigaturesStandard           =1          # from enum WdLigatures
	wdLigaturesStandardContextual =3          # from enum WdLigatures
	wdLigaturesStandardContextualDiscretional=11         # from enum WdLigatures
	wdLigaturesStandardContextualHistorical=7          # from enum WdLigatures
	wdLigaturesStandardDiscretional=9          # from enum WdLigatures
	wdLigaturesStandardHistorical =5          # from enum WdLigatures
	wdLigaturesStandardHistoricalDiscretional=13         # from enum WdLigatures
	wdCRLF                        =0          # from enum WdLineEndingType
	wdCROnly                      =1          # from enum WdLineEndingType
	wdLFCR                        =3          # from enum WdLineEndingType
	wdLFOnly                      =2          # from enum WdLineEndingType
	wdLSPS                        =4          # from enum WdLineEndingType
	wdLineSpace1pt5               =1          # from enum WdLineSpacing
	wdLineSpaceAtLeast            =3          # from enum WdLineSpacing
	wdLineSpaceDouble             =2          # from enum WdLineSpacing
	wdLineSpaceExactly            =4          # from enum WdLineSpacing
	wdLineSpaceMultiple           =5          # from enum WdLineSpacing
	wdLineSpaceSingle             =0          # from enum WdLineSpacing
	wdLineStyleDashDot            =5          # from enum WdLineStyle
	wdLineStyleDashDotDot         =6          # from enum WdLineStyle
	wdLineStyleDashDotStroked     =20         # from enum WdLineStyle
	wdLineStyleDashLargeGap       =4          # from enum WdLineStyle
	wdLineStyleDashSmallGap       =3          # from enum WdLineStyle
	wdLineStyleDot                =2          # from enum WdLineStyle
	wdLineStyleDouble             =7          # from enum WdLineStyle
	wdLineStyleDoubleWavy         =19         # from enum WdLineStyle
	wdLineStyleEmboss3D           =21         # from enum WdLineStyle
	wdLineStyleEngrave3D          =22         # from enum WdLineStyle
	wdLineStyleInset              =24         # from enum WdLineStyle
	wdLineStyleNone               =0          # from enum WdLineStyle
	wdLineStyleOutset             =23         # from enum WdLineStyle
	wdLineStyleSingle             =1          # from enum WdLineStyle
	wdLineStyleSingleWavy         =18         # from enum WdLineStyle
	wdLineStyleThickThinLargeGap  =16         # from enum WdLineStyle
	wdLineStyleThickThinMedGap    =13         # from enum WdLineStyle
	wdLineStyleThickThinSmallGap  =10         # from enum WdLineStyle
	wdLineStyleThinThickLargeGap  =15         # from enum WdLineStyle
	wdLineStyleThinThickMedGap    =12         # from enum WdLineStyle
	wdLineStyleThinThickSmallGap  =9          # from enum WdLineStyle
	wdLineStyleThinThickThinLargeGap=17         # from enum WdLineStyle
	wdLineStyleThinThickThinMedGap=14         # from enum WdLineStyle
	wdLineStyleThinThickThinSmallGap=11         # from enum WdLineStyle
	wdLineStyleTriple             =8          # from enum WdLineStyle
	wdTableRow                    =1          # from enum WdLineType
	wdTextLine                    =0          # from enum WdLineType
	wdLineWidth025pt              =2          # from enum WdLineWidth
	wdLineWidth050pt              =4          # from enum WdLineWidth
	wdLineWidth075pt              =6          # from enum WdLineWidth
	wdLineWidth100pt              =8          # from enum WdLineWidth
	wdLineWidth150pt              =12         # from enum WdLineWidth
	wdLineWidth225pt              =18         # from enum WdLineWidth
	wdLineWidth300pt              =24         # from enum WdLineWidth
	wdLineWidth450pt              =36         # from enum WdLineWidth
	wdLineWidth600pt              =48         # from enum WdLineWidth
	wdLinkTypeChart               =8          # from enum WdLinkType
	wdLinkTypeDDE                 =6          # from enum WdLinkType
	wdLinkTypeDDEAuto             =7          # from enum WdLinkType
	wdLinkTypeImport              =5          # from enum WdLinkType
	wdLinkTypeInclude             =4          # from enum WdLinkType
	wdLinkTypeOLE                 =0          # from enum WdLinkType
	wdLinkTypePicture             =1          # from enum WdLinkType
	wdLinkTypeReference           =3          # from enum WdLinkType
	wdLinkTypeText                =2          # from enum WdLinkType
	wdListApplyToSelection        =2          # from enum WdListApplyTo
	wdListApplyToThisPointForward =1          # from enum WdListApplyTo
	wdListApplyToWholeList        =0          # from enum WdListApplyTo
	wdBulletGallery               =1          # from enum WdListGalleryType
	wdNumberGallery               =2          # from enum WdListGalleryType
	wdOutlineNumberGallery        =3          # from enum WdListGalleryType
	wdListLevelAlignCenter        =1          # from enum WdListLevelAlignment
	wdListLevelAlignLeft          =0          # from enum WdListLevelAlignment
	wdListLevelAlignRight         =2          # from enum WdListLevelAlignment
	wdListNumberStyleAiueo        =20         # from enum WdListNumberStyle
	wdListNumberStyleAiueoHalfWidth=12         # from enum WdListNumberStyle
	wdListNumberStyleArabic       =0          # from enum WdListNumberStyle
	wdListNumberStyleArabic1      =46         # from enum WdListNumberStyle
	wdListNumberStyleArabic2      =48         # from enum WdListNumberStyle
	wdListNumberStyleArabicFullWidth=14         # from enum WdListNumberStyle
	wdListNumberStyleArabicLZ     =22         # from enum WdListNumberStyle
	wdListNumberStyleArabicLZ2    =62         # from enum WdListNumberStyle
	wdListNumberStyleArabicLZ3    =63         # from enum WdListNumberStyle
	wdListNumberStyleArabicLZ4    =64         # from enum WdListNumberStyle
	wdListNumberStyleBullet       =23         # from enum WdListNumberStyle
	wdListNumberStyleCardinalText =6          # from enum WdListNumberStyle
	wdListNumberStyleChosung      =25         # from enum WdListNumberStyle
	wdListNumberStyleGBNum1       =26         # from enum WdListNumberStyle
	wdListNumberStyleGBNum2       =27         # from enum WdListNumberStyle
	wdListNumberStyleGBNum3       =28         # from enum WdListNumberStyle
	wdListNumberStyleGBNum4       =29         # from enum WdListNumberStyle
	wdListNumberStyleGanada       =24         # from enum WdListNumberStyle
	wdListNumberStyleHangul       =43         # from enum WdListNumberStyle
	wdListNumberStyleHanja        =44         # from enum WdListNumberStyle
	wdListNumberStyleHanjaRead    =41         # from enum WdListNumberStyle
	wdListNumberStyleHanjaReadDigit=42         # from enum WdListNumberStyle
	wdListNumberStyleHebrew1      =45         # from enum WdListNumberStyle
	wdListNumberStyleHebrew2      =47         # from enum WdListNumberStyle
	wdListNumberStyleHindiArabic  =51         # from enum WdListNumberStyle
	wdListNumberStyleHindiCardinalText=52         # from enum WdListNumberStyle
	wdListNumberStyleHindiLetter1 =49         # from enum WdListNumberStyle
	wdListNumberStyleHindiLetter2 =50         # from enum WdListNumberStyle
	wdListNumberStyleIroha        =21         # from enum WdListNumberStyle
	wdListNumberStyleIrohaHalfWidth=13         # from enum WdListNumberStyle
	wdListNumberStyleKanji        =10         # from enum WdListNumberStyle
	wdListNumberStyleKanjiDigit   =11         # from enum WdListNumberStyle
	wdListNumberStyleKanjiTraditional=16         # from enum WdListNumberStyle
	wdListNumberStyleKanjiTraditional2=17         # from enum WdListNumberStyle
	wdListNumberStyleLegal        =253        # from enum WdListNumberStyle
	wdListNumberStyleLegalLZ      =254        # from enum WdListNumberStyle
	wdListNumberStyleLowercaseBulgarian=67         # from enum WdListNumberStyle
	wdListNumberStyleLowercaseGreek=60         # from enum WdListNumberStyle
	wdListNumberStyleLowercaseLetter=4          # from enum WdListNumberStyle
	wdListNumberStyleLowercaseRoman=2          # from enum WdListNumberStyle
	wdListNumberStyleLowercaseRussian=58         # from enum WdListNumberStyle
	wdListNumberStyleLowercaseTurkish=65         # from enum WdListNumberStyle
	wdListNumberStyleNone         =255        # from enum WdListNumberStyle
	wdListNumberStyleNumberInCircle=18         # from enum WdListNumberStyle
	wdListNumberStyleOrdinal      =5          # from enum WdListNumberStyle
	wdListNumberStyleOrdinalText  =7          # from enum WdListNumberStyle
	wdListNumberStylePictureBullet=249        # from enum WdListNumberStyle
	wdListNumberStyleSimpChinNum1 =37         # from enum WdListNumberStyle
	wdListNumberStyleSimpChinNum2 =38         # from enum WdListNumberStyle
	wdListNumberStyleSimpChinNum3 =39         # from enum WdListNumberStyle
	wdListNumberStyleSimpChinNum4 =40         # from enum WdListNumberStyle
	wdListNumberStyleThaiArabic   =54         # from enum WdListNumberStyle
	wdListNumberStyleThaiCardinalText=55         # from enum WdListNumberStyle
	wdListNumberStyleThaiLetter   =53         # from enum WdListNumberStyle
	wdListNumberStyleTradChinNum1 =33         # from enum WdListNumberStyle
	wdListNumberStyleTradChinNum2 =34         # from enum WdListNumberStyle
	wdListNumberStyleTradChinNum3 =35         # from enum WdListNumberStyle
	wdListNumberStyleTradChinNum4 =36         # from enum WdListNumberStyle
	wdListNumberStyleUppercaseBulgarian=68         # from enum WdListNumberStyle
	wdListNumberStyleUppercaseGreek=61         # from enum WdListNumberStyle
	wdListNumberStyleUppercaseLetter=3          # from enum WdListNumberStyle
	wdListNumberStyleUppercaseRoman=1          # from enum WdListNumberStyle
	wdListNumberStyleUppercaseRussian=59         # from enum WdListNumberStyle
	wdListNumberStyleUppercaseTurkish=66         # from enum WdListNumberStyle
	wdListNumberStyleVietCardinalText=56         # from enum WdListNumberStyle
	wdListNumberStyleZodiac1      =30         # from enum WdListNumberStyle
	wdListNumberStyleZodiac2      =31         # from enum WdListNumberStyle
	wdListNumberStyleZodiac3      =32         # from enum WdListNumberStyle
	emptyenum                     =0          # from enum WdListNumberStyleHID
	wdListBullet                  =2          # from enum WdListType
	wdListListNumOnly             =1          # from enum WdListType
	wdListMixedNumbering          =5          # from enum WdListType
	wdListNoNumbering             =0          # from enum WdListType
	wdListOutlineNumbering        =4          # from enum WdListType
	wdListPictureBullet           =6          # from enum WdListType
	wdListSimpleNumbering         =3          # from enum WdListType
	wdLockChanged                 =3          # from enum WdLockType
	wdLockEphemeral               =2          # from enum WdLockType
	wdLockNone                    =0          # from enum WdLockType
	wdLockReservation             =1          # from enum WdLockType
	wdFirstDataSourceRecord       =-6         # from enum WdMailMergeActiveRecord
	wdFirstRecord                 =-4         # from enum WdMailMergeActiveRecord
	wdLastDataSourceRecord        =-7         # from enum WdMailMergeActiveRecord
	wdLastRecord                  =-5         # from enum WdMailMergeActiveRecord
	wdNextDataSourceRecord        =-8         # from enum WdMailMergeActiveRecord
	wdNextRecord                  =-2         # from enum WdMailMergeActiveRecord
	wdNoActiveRecord              =-1         # from enum WdMailMergeActiveRecord
	wdPreviousDataSourceRecord    =-9         # from enum WdMailMergeActiveRecord
	wdPreviousRecord              =-3         # from enum WdMailMergeActiveRecord
	wdMergeIfEqual                =0          # from enum WdMailMergeComparison
	wdMergeIfGreaterThan          =3          # from enum WdMailMergeComparison
	wdMergeIfGreaterThanOrEqual   =5          # from enum WdMailMergeComparison
	wdMergeIfIsBlank              =6          # from enum WdMailMergeComparison
	wdMergeIfIsNotBlank           =7          # from enum WdMailMergeComparison
	wdMergeIfLessThan             =2          # from enum WdMailMergeComparison
	wdMergeIfLessThanOrEqual      =4          # from enum WdMailMergeComparison
	wdMergeIfNotEqual             =1          # from enum WdMailMergeComparison
	wdMergeInfoFromAccessDDE      =1          # from enum WdMailMergeDataSource
	wdMergeInfoFromExcelDDE       =2          # from enum WdMailMergeDataSource
	wdMergeInfoFromMSQueryDDE     =3          # from enum WdMailMergeDataSource
	wdMergeInfoFromODBC           =4          # from enum WdMailMergeDataSource
	wdMergeInfoFromODSO           =5          # from enum WdMailMergeDataSource
	wdMergeInfoFromWord           =0          # from enum WdMailMergeDataSource
	wdNoMergeInfo                 =-1         # from enum WdMailMergeDataSource
	wdDefaultFirstRecord          =1          # from enum WdMailMergeDefaultRecord
	wdDefaultLastRecord           =-16        # from enum WdMailMergeDefaultRecord
	wdSendToEmail                 =2          # from enum WdMailMergeDestination
	wdSendToFax                   =3          # from enum WdMailMergeDestination
	wdSendToNewDocument           =0          # from enum WdMailMergeDestination
	wdSendToPrinter               =1          # from enum WdMailMergeDestination
	wdMailFormatHTML              =1          # from enum WdMailMergeMailFormat
	wdMailFormatPlainText         =0          # from enum WdMailMergeMailFormat
	wdCatalog                     =3          # from enum WdMailMergeMainDocType
	wdDirectory                   =3          # from enum WdMailMergeMainDocType
	wdEMail                       =4          # from enum WdMailMergeMainDocType
	wdEnvelopes                   =2          # from enum WdMailMergeMainDocType
	wdFax                         =5          # from enum WdMailMergeMainDocType
	wdFormLetters                 =0          # from enum WdMailMergeMainDocType
	wdMailingLabels               =1          # from enum WdMailMergeMainDocType
	wdNotAMergeDocument           =-1         # from enum WdMailMergeMainDocType
	wdDataSource                  =5          # from enum WdMailMergeState
	wdMainAndDataSource           =2          # from enum WdMailMergeState
	wdMainAndHeader               =3          # from enum WdMailMergeState
	wdMainAndSourceAndHeader      =4          # from enum WdMailMergeState
	wdMainDocumentOnly            =1          # from enum WdMailMergeState
	wdNormalDocument              =0          # from enum WdMailMergeState
	wdMAPI                        =1          # from enum WdMailSystem
	wdMAPIandPowerTalk            =3          # from enum WdMailSystem
	wdNoMailSystem                =0          # from enum WdMailSystem
	wdPowerTalk                   =2          # from enum WdMailSystem
	wdPriorityHigh                =3          # from enum WdMailerPriority
	wdPriorityLow                 =2          # from enum WdMailerPriority
	wdPriorityNormal              =1          # from enum WdMailerPriority
	wdAddress1                    =10         # from enum WdMappedDataFields
	wdAddress2                    =11         # from enum WdMappedDataFields
	wdAddress3                    =29         # from enum WdMappedDataFields
	wdBusinessFax                 =17         # from enum WdMappedDataFields
	wdBusinessPhone               =16         # from enum WdMappedDataFields
	wdCity                        =12         # from enum WdMappedDataFields
	wdCompany                     =9          # from enum WdMappedDataFields
	wdCountryRegion               =15         # from enum WdMappedDataFields
	wdCourtesyTitle               =2          # from enum WdMappedDataFields
	wdDepartment                  =30         # from enum WdMappedDataFields
	wdEmailAddress                =20         # from enum WdMappedDataFields
	wdFirstName                   =3          # from enum WdMappedDataFields
	wdHomeFax                     =19         # from enum WdMappedDataFields
	wdHomePhone                   =18         # from enum WdMappedDataFields
	wdJobTitle                    =8          # from enum WdMappedDataFields
	wdLastName                    =5          # from enum WdMappedDataFields
	wdMiddleName                  =4          # from enum WdMappedDataFields
	wdNickname                    =7          # from enum WdMappedDataFields
	wdPostalCode                  =14         # from enum WdMappedDataFields
	wdRubyFirstName               =27         # from enum WdMappedDataFields
	wdRubyLastName                =28         # from enum WdMappedDataFields
	wdSpouseCourtesyTitle         =22         # from enum WdMappedDataFields
	wdSpouseFirstName             =23         # from enum WdMappedDataFields
	wdSpouseLastName              =25         # from enum WdMappedDataFields
	wdSpouseMiddleName            =24         # from enum WdMappedDataFields
	wdSpouseNickname              =26         # from enum WdMappedDataFields
	wdState                       =13         # from enum WdMappedDataFields
	wdSuffix                      =6          # from enum WdMappedDataFields
	wdUniqueIdentifier            =1          # from enum WdMappedDataFields
	wdWebPageURL                  =21         # from enum WdMappedDataFields
	wdCentimeters                 =1          # from enum WdMeasurementUnits
	wdInches                      =0          # from enum WdMeasurementUnits
	wdMillimeters                 =2          # from enum WdMeasurementUnits
	wdPicas                       =4          # from enum WdMeasurementUnits
	wdPoints                      =3          # from enum WdMeasurementUnits
	emptyenum                     =0          # from enum WdMeasurementUnitsHID
	wdMergeFormatFromOriginal     =0          # from enum WdMergeFormatFrom
	wdMergeFormatFromPrompt       =2          # from enum WdMergeFormatFrom
	wdMergeFormatFromRevised      =1          # from enum WdMergeFormatFrom
	wdMergeSubTypeAccess          =1          # from enum WdMergeSubType
	wdMergeSubTypeOAL             =2          # from enum WdMergeSubType
	wdMergeSubTypeOLEDBText       =5          # from enum WdMergeSubType
	wdMergeSubTypeOLEDBWord       =3          # from enum WdMergeSubType
	wdMergeSubTypeOther           =0          # from enum WdMergeSubType
	wdMergeSubTypeOutlook         =6          # from enum WdMergeSubType
	wdMergeSubTypeWord            =7          # from enum WdMergeSubType
	wdMergeSubTypeWord2000        =8          # from enum WdMergeSubType
	wdMergeSubTypeWorks           =4          # from enum WdMergeSubType
	wdMergeTargetCurrent          =1          # from enum WdMergeTarget
	wdMergeTargetNew              =2          # from enum WdMergeTarget
	wdMergeTargetSelected         =0          # from enum WdMergeTarget
	wdMonthNamesArabic            =0          # from enum WdMonthNames
	wdMonthNamesEnglish           =1          # from enum WdMonthNames
	wdMonthNamesFrench            =2          # from enum WdMonthNames
	wdMoveFromTextMarkBold        =6          # from enum WdMoveFromTextMark
	wdMoveFromTextMarkCaret       =3          # from enum WdMoveFromTextMark
	wdMoveFromTextMarkColorOnly   =10         # from enum WdMoveFromTextMark
	wdMoveFromTextMarkDoubleStrikeThrough=1          # from enum WdMoveFromTextMark
	wdMoveFromTextMarkDoubleUnderline=9          # from enum WdMoveFromTextMark
	wdMoveFromTextMarkHidden      =0          # from enum WdMoveFromTextMark
	wdMoveFromTextMarkItalic      =7          # from enum WdMoveFromTextMark
	wdMoveFromTextMarkNone        =5          # from enum WdMoveFromTextMark
	wdMoveFromTextMarkPound       =4          # from enum WdMoveFromTextMark
	wdMoveFromTextMarkStrikeThrough=2          # from enum WdMoveFromTextMark
	wdMoveFromTextMarkUnderline   =8          # from enum WdMoveFromTextMark
	wdMoveToTextMarkBold          =1          # from enum WdMoveToTextMark
	wdMoveToTextMarkColorOnly     =5          # from enum WdMoveToTextMark
	wdMoveToTextMarkDoubleStrikeThrough=7          # from enum WdMoveToTextMark
	wdMoveToTextMarkDoubleUnderline=4          # from enum WdMoveToTextMark
	wdMoveToTextMarkItalic        =2          # from enum WdMoveToTextMark
	wdMoveToTextMarkNone          =0          # from enum WdMoveToTextMark
	wdMoveToTextMarkStrikeThrough =6          # from enum WdMoveToTextMark
	wdMoveToTextMarkUnderline     =3          # from enum WdMoveToTextMark
	wdExtend                      =1          # from enum WdMovementType
	wdMove                        =0          # from enum WdMovementType
	wdHangulToHanja               =0          # from enum WdMultipleWordConversionsMode
	wdHanjaToHangul               =1          # from enum WdMultipleWordConversionsMode
	wdNewBlankDocument            =0          # from enum WdNewDocumentType
	wdNewEmailMessage             =2          # from enum WdNewDocumentType
	wdNewFrameset                 =3          # from enum WdNewDocumentType
	wdNewWebPage                  =1          # from enum WdNewDocumentType
	wdNewXMLDocument              =4          # from enum WdNewDocumentType
	wdNoteNumberStyleArabic       =0          # from enum WdNoteNumberStyle
	wdNoteNumberStyleArabicFullWidth=14         # from enum WdNoteNumberStyle
	wdNoteNumberStyleArabicLetter1=46         # from enum WdNoteNumberStyle
	wdNoteNumberStyleArabicLetter2=48         # from enum WdNoteNumberStyle
	wdNoteNumberStyleHanjaRead    =41         # from enum WdNoteNumberStyle
	wdNoteNumberStyleHanjaReadDigit=42         # from enum WdNoteNumberStyle
	wdNoteNumberStyleHebrewLetter1=45         # from enum WdNoteNumberStyle
	wdNoteNumberStyleHebrewLetter2=47         # from enum WdNoteNumberStyle
	wdNoteNumberStyleHindiArabic  =51         # from enum WdNoteNumberStyle
	wdNoteNumberStyleHindiCardinalText=52         # from enum WdNoteNumberStyle
	wdNoteNumberStyleHindiLetter1 =49         # from enum WdNoteNumberStyle
	wdNoteNumberStyleHindiLetter2 =50         # from enum WdNoteNumberStyle
	wdNoteNumberStyleKanji        =10         # from enum WdNoteNumberStyle
	wdNoteNumberStyleKanjiDigit   =11         # from enum WdNoteNumberStyle
	wdNoteNumberStyleKanjiTraditional=16         # from enum WdNoteNumberStyle
	wdNoteNumberStyleLowercaseLetter=4          # from enum WdNoteNumberStyle
	wdNoteNumberStyleLowercaseRoman=2          # from enum WdNoteNumberStyle
	wdNoteNumberStyleNumberInCircle=18         # from enum WdNoteNumberStyle
	wdNoteNumberStyleSimpChinNum1 =37         # from enum WdNoteNumberStyle
	wdNoteNumberStyleSimpChinNum2 =38         # from enum WdNoteNumberStyle
	wdNoteNumberStyleSymbol       =9          # from enum WdNoteNumberStyle
	wdNoteNumberStyleThaiArabic   =54         # from enum WdNoteNumberStyle
	wdNoteNumberStyleThaiCardinalText=55         # from enum WdNoteNumberStyle
	wdNoteNumberStyleThaiLetter   =53         # from enum WdNoteNumberStyle
	wdNoteNumberStyleTradChinNum1 =33         # from enum WdNoteNumberStyle
	wdNoteNumberStyleTradChinNum2 =34         # from enum WdNoteNumberStyle
	wdNoteNumberStyleUppercaseLetter=3          # from enum WdNoteNumberStyle
	wdNoteNumberStyleUppercaseRoman=1          # from enum WdNoteNumberStyle
	wdNoteNumberStyleVietCardinalText=56         # from enum WdNoteNumberStyle
	emptyenum                     =0          # from enum WdNoteNumberStyleHID
	wdNumberFormDefault           =0          # from enum WdNumberForm
	wdNumberFormLining            =1          # from enum WdNumberForm
	wdNumberFormOldStyle          =2          # from enum WdNumberForm
	wdNumberSpacingDefault        =0          # from enum WdNumberSpacing
	wdNumberSpacingProportional   =1          # from enum WdNumberSpacing
	wdNumberSpacingTabular        =2          # from enum WdNumberSpacing
	wdCaptionNumberStyleBidiLetter1=49         # from enum WdNumberStyleWordBasicBiDi
	wdCaptionNumberStyleBidiLetter2=50         # from enum WdNumberStyleWordBasicBiDi
	wdListNumberStyleBidi1        =49         # from enum WdNumberStyleWordBasicBiDi
	wdListNumberStyleBidi2        =50         # from enum WdNumberStyleWordBasicBiDi
	wdNoteNumberStyleBidiLetter1  =49         # from enum WdNumberStyleWordBasicBiDi
	wdNoteNumberStyleBidiLetter2  =50         # from enum WdNumberStyleWordBasicBiDi
	wdPageNumberStyleBidiLetter1  =49         # from enum WdNumberStyleWordBasicBiDi
	wdPageNumberStyleBidiLetter2  =50         # from enum WdNumberStyleWordBasicBiDi
	wdNumberAllNumbers            =3          # from enum WdNumberType
	wdNumberListNum               =2          # from enum WdNumberType
	wdNumberParagraph             =1          # from enum WdNumberType
	wdRestartContinuous           =0          # from enum WdNumberingRule
	wdRestartPage                 =2          # from enum WdNumberingRule
	wdRestartSection              =1          # from enum WdNumberingRule
	wdFloatOverText               =1          # from enum WdOLEPlacement
	wdInLine                      =0          # from enum WdOLEPlacement
	wdOLEControl                  =2          # from enum WdOLEType
	wdOLEEmbed                    =1          # from enum WdOLEType
	wdOLELink                     =0          # from enum WdOLEType
	wdOLEVerbDiscardUndoState     =-6         # from enum WdOLEVerb
	wdOLEVerbHide                 =-3         # from enum WdOLEVerb
	wdOLEVerbInPlaceActivate      =-5         # from enum WdOLEVerb
	wdOLEVerbOpen                 =-2         # from enum WdOLEVerb
	wdOLEVerbPrimary              =0          # from enum WdOLEVerb
	wdOLEVerbShow                 =-1         # from enum WdOLEVerb
	wdOLEVerbUIActivate           =-4         # from enum WdOLEVerb
	wdOMathBreakBinAfter          =1          # from enum WdOMathBreakBin
	wdOMathBreakBinBefore         =0          # from enum WdOMathBreakBin
	wdOMathBreakBinRepeat         =2          # from enum WdOMathBreakBin
	wdOMathBreakSubMinusMinus     =0          # from enum WdOMathBreakSub
	wdOMathBreakSubMinusPlus      =2          # from enum WdOMathBreakSub
	wdOMathBreakSubPlusMinus      =1          # from enum WdOMathBreakSub
	wdOMathFracBar                =0          # from enum WdOMathFracType
	wdOMathFracLin                =3          # from enum WdOMathFracType
	wdOMathFracNoBar              =1          # from enum WdOMathFracType
	wdOMathFracSkw                =2          # from enum WdOMathFracType
	wdOMathFunctionAcc            =1          # from enum WdOMathFunctionType
	wdOMathFunctionBar            =2          # from enum WdOMathFunctionType
	wdOMathFunctionBorderBox      =4          # from enum WdOMathFunctionType
	wdOMathFunctionBox            =3          # from enum WdOMathFunctionType
	wdOMathFunctionDelim          =5          # from enum WdOMathFunctionType
	wdOMathFunctionEqArray        =6          # from enum WdOMathFunctionType
	wdOMathFunctionFrac           =7          # from enum WdOMathFunctionType
	wdOMathFunctionFunc           =8          # from enum WdOMathFunctionType
	wdOMathFunctionGroupChar      =9          # from enum WdOMathFunctionType
	wdOMathFunctionLimLow         =10         # from enum WdOMathFunctionType
	wdOMathFunctionLimUpp         =11         # from enum WdOMathFunctionType
	wdOMathFunctionLiteralText    =22         # from enum WdOMathFunctionType
	wdOMathFunctionMat            =12         # from enum WdOMathFunctionType
	wdOMathFunctionNary           =13         # from enum WdOMathFunctionType
	wdOMathFunctionNormalText     =21         # from enum WdOMathFunctionType
	wdOMathFunctionPhantom        =14         # from enum WdOMathFunctionType
	wdOMathFunctionRad            =16         # from enum WdOMathFunctionType
	wdOMathFunctionScrPre         =15         # from enum WdOMathFunctionType
	wdOMathFunctionScrSub         =17         # from enum WdOMathFunctionType
	wdOMathFunctionScrSubSup      =18         # from enum WdOMathFunctionType
	wdOMathFunctionScrSup         =19         # from enum WdOMathFunctionType
	wdOMathFunctionText           =20         # from enum WdOMathFunctionType
	wdOMathHorizAlignCenter       =0          # from enum WdOMathHorizAlignType
	wdOMathHorizAlignLeft         =1          # from enum WdOMathHorizAlignType
	wdOMathHorizAlignRight        =2          # from enum WdOMathHorizAlignType
	wdOMathJcCenter               =2          # from enum WdOMathJc
	wdOMathJcCenterGroup          =1          # from enum WdOMathJc
	wdOMathJcInline               =7          # from enum WdOMathJc
	wdOMathJcLeft                 =3          # from enum WdOMathJc
	wdOMathJcRight                =4          # from enum WdOMathJc
	wdOMathShapeCentered          =0          # from enum WdOMathShapeType
	wdOMathShapeMatch             =1          # from enum WdOMathShapeType
	wdOMathSpacing1pt5            =1          # from enum WdOMathSpacingRule
	wdOMathSpacingDouble          =2          # from enum WdOMathSpacingRule
	wdOMathSpacingExactly         =3          # from enum WdOMathSpacingRule
	wdOMathSpacingMultiple        =4          # from enum WdOMathSpacingRule
	wdOMathSpacingSingle          =0          # from enum WdOMathSpacingRule
	wdOMathDisplay                =0          # from enum WdOMathType
	wdOMathInline                 =1          # from enum WdOMathType
	wdOMathVertAlignBottom        =2          # from enum WdOMathVertAlignType
	wdOMathVertAlignCenter        =0          # from enum WdOMathVertAlignType
	wdOMathVertAlignTop           =1          # from enum WdOMathVertAlignType
	wdOpenFormatAllWord           =6          # from enum WdOpenFormat
	wdOpenFormatAllWordTemplates  =13         # from enum WdOpenFormat
	wdOpenFormatAuto              =0          # from enum WdOpenFormat
	wdOpenFormatDocument          =1          # from enum WdOpenFormat
	wdOpenFormatDocument97        =1          # from enum WdOpenFormat
	wdOpenFormatEncodedText       =5          # from enum WdOpenFormat
	wdOpenFormatOpenDocumentText  =18         # from enum WdOpenFormat
	wdOpenFormatRTF               =3          # from enum WdOpenFormat
	wdOpenFormatTemplate          =2          # from enum WdOpenFormat
	wdOpenFormatTemplate97        =2          # from enum WdOpenFormat
	wdOpenFormatText              =4          # from enum WdOpenFormat
	wdOpenFormatUnicodeText       =5          # from enum WdOpenFormat
	wdOpenFormatWebPages          =7          # from enum WdOpenFormat
	wdOpenFormatXML               =8          # from enum WdOpenFormat
	wdOpenFormatXMLDocument       =9          # from enum WdOpenFormat
	wdOpenFormatXMLDocumentMacroEnabled=10         # from enum WdOpenFormat
	wdOpenFormatXMLDocumentMacroEnabledSerialized=15         # from enum WdOpenFormat
	wdOpenFormatXMLDocumentSerialized=14         # from enum WdOpenFormat
	wdOpenFormatXMLTemplate       =11         # from enum WdOpenFormat
	wdOpenFormatXMLTemplateMacroEnabled=12         # from enum WdOpenFormat
	wdOpenFormatXMLTemplateMacroEnabledSerialized=17         # from enum WdOpenFormat
	wdOpenFormatXMLTemplateSerialized=16         # from enum WdOpenFormat
	wdOrganizerObjectAutoText     =1          # from enum WdOrganizerObject
	wdOrganizerObjectCommandBars  =2          # from enum WdOrganizerObject
	wdOrganizerObjectProjectItems =3          # from enum WdOrganizerObject
	wdOrganizerObjectStyles       =0          # from enum WdOrganizerObject
	wdOrientLandscape             =1          # from enum WdOrientation
	wdOrientPortrait              =0          # from enum WdOrientation
	wdOriginalDocumentFormat      =1          # from enum WdOriginalFormat
	wdPromptUser                  =2          # from enum WdOriginalFormat
	wdWordDocument                =0          # from enum WdOriginalFormat
	wdOutlineLevel1               =1          # from enum WdOutlineLevel
	wdOutlineLevel2               =2          # from enum WdOutlineLevel
	wdOutlineLevel3               =3          # from enum WdOutlineLevel
	wdOutlineLevel4               =4          # from enum WdOutlineLevel
	wdOutlineLevel5               =5          # from enum WdOutlineLevel
	wdOutlineLevel6               =6          # from enum WdOutlineLevel
	wdOutlineLevel7               =7          # from enum WdOutlineLevel
	wdOutlineLevel8               =8          # from enum WdOutlineLevel
	wdOutlineLevel9               =9          # from enum WdOutlineLevel
	wdOutlineLevelBodyText        =10         # from enum WdOutlineLevel
	wdArtApples                   =1          # from enum WdPageBorderArt
	wdArtArchedScallops           =97         # from enum WdPageBorderArt
	wdArtBabyPacifier             =70         # from enum WdPageBorderArt
	wdArtBabyRattle               =71         # from enum WdPageBorderArt
	wdArtBalloons3Colors          =11         # from enum WdPageBorderArt
	wdArtBalloonsHotAir           =12         # from enum WdPageBorderArt
	wdArtBasicBlackDashes         =155        # from enum WdPageBorderArt
	wdArtBasicBlackDots           =156        # from enum WdPageBorderArt
	wdArtBasicBlackSquares        =154        # from enum WdPageBorderArt
	wdArtBasicThinLines           =151        # from enum WdPageBorderArt
	wdArtBasicWhiteDashes         =152        # from enum WdPageBorderArt
	wdArtBasicWhiteDots           =147        # from enum WdPageBorderArt
	wdArtBasicWhiteSquares        =153        # from enum WdPageBorderArt
	wdArtBasicWideInline          =150        # from enum WdPageBorderArt
	wdArtBasicWideMidline         =148        # from enum WdPageBorderArt
	wdArtBasicWideOutline         =149        # from enum WdPageBorderArt
	wdArtBats                     =37         # from enum WdPageBorderArt
	wdArtBirds                    =102        # from enum WdPageBorderArt
	wdArtBirdsFlight              =35         # from enum WdPageBorderArt
	wdArtCabins                   =72         # from enum WdPageBorderArt
	wdArtCakeSlice                =3          # from enum WdPageBorderArt
	wdArtCandyCorn                =4          # from enum WdPageBorderArt
	wdArtCelticKnotwork           =99         # from enum WdPageBorderArt
	wdArtCertificateBanner        =158        # from enum WdPageBorderArt
	wdArtChainLink                =128        # from enum WdPageBorderArt
	wdArtChampagneBottle          =6          # from enum WdPageBorderArt
	wdArtCheckedBarBlack          =145        # from enum WdPageBorderArt
	wdArtCheckedBarColor          =61         # from enum WdPageBorderArt
	wdArtCheckered                =144        # from enum WdPageBorderArt
	wdArtChristmasTree            =8          # from enum WdPageBorderArt
	wdArtCirclesLines             =91         # from enum WdPageBorderArt
	wdArtCirclesRectangles        =140        # from enum WdPageBorderArt
	wdArtClassicalWave            =56         # from enum WdPageBorderArt
	wdArtClocks                   =27         # from enum WdPageBorderArt
	wdArtCompass                  =54         # from enum WdPageBorderArt
	wdArtConfetti                 =31         # from enum WdPageBorderArt
	wdArtConfettiGrays            =115        # from enum WdPageBorderArt
	wdArtConfettiOutline          =116        # from enum WdPageBorderArt
	wdArtConfettiStreamers        =14         # from enum WdPageBorderArt
	wdArtConfettiWhite            =117        # from enum WdPageBorderArt
	wdArtCornerTriangles          =141        # from enum WdPageBorderArt
	wdArtCouponCutoutDashes       =163        # from enum WdPageBorderArt
	wdArtCouponCutoutDots         =164        # from enum WdPageBorderArt
	wdArtCrazyMaze                =100        # from enum WdPageBorderArt
	wdArtCreaturesButterfly       =32         # from enum WdPageBorderArt
	wdArtCreaturesFish            =34         # from enum WdPageBorderArt
	wdArtCreaturesInsects         =142        # from enum WdPageBorderArt
	wdArtCreaturesLadyBug         =33         # from enum WdPageBorderArt
	wdArtCrossStitch              =138        # from enum WdPageBorderArt
	wdArtCup                      =67         # from enum WdPageBorderArt
	wdArtDecoArch                 =89         # from enum WdPageBorderArt
	wdArtDecoArchColor            =50         # from enum WdPageBorderArt
	wdArtDecoBlocks               =90         # from enum WdPageBorderArt
	wdArtDiamondsGray             =88         # from enum WdPageBorderArt
	wdArtDoubleD                  =55         # from enum WdPageBorderArt
	wdArtDoubleDiamonds           =127        # from enum WdPageBorderArt
	wdArtEarth1                   =22         # from enum WdPageBorderArt
	wdArtEarth2                   =21         # from enum WdPageBorderArt
	wdArtEclipsingSquares1        =101        # from enum WdPageBorderArt
	wdArtEclipsingSquares2        =86         # from enum WdPageBorderArt
	wdArtEggsBlack                =66         # from enum WdPageBorderArt
	wdArtFans                     =51         # from enum WdPageBorderArt
	wdArtFilm                     =52         # from enum WdPageBorderArt
	wdArtFirecrackers             =28         # from enum WdPageBorderArt
	wdArtFlowersBlockPrint        =49         # from enum WdPageBorderArt
	wdArtFlowersDaisies           =48         # from enum WdPageBorderArt
	wdArtFlowersModern1           =45         # from enum WdPageBorderArt
	wdArtFlowersModern2           =44         # from enum WdPageBorderArt
	wdArtFlowersPansy             =43         # from enum WdPageBorderArt
	wdArtFlowersRedRose           =39         # from enum WdPageBorderArt
	wdArtFlowersRoses             =38         # from enum WdPageBorderArt
	wdArtFlowersTeacup            =103        # from enum WdPageBorderArt
	wdArtFlowersTiny              =42         # from enum WdPageBorderArt
	wdArtGems                     =139        # from enum WdPageBorderArt
	wdArtGingerbreadMan           =69         # from enum WdPageBorderArt
	wdArtGradient                 =122        # from enum WdPageBorderArt
	wdArtHandmade1                =159        # from enum WdPageBorderArt
	wdArtHandmade2                =160        # from enum WdPageBorderArt
	wdArtHeartBalloon             =16         # from enum WdPageBorderArt
	wdArtHeartGray                =68         # from enum WdPageBorderArt
	wdArtHearts                   =15         # from enum WdPageBorderArt
	wdArtHeebieJeebies            =120        # from enum WdPageBorderArt
	wdArtHolly                    =41         # from enum WdPageBorderArt
	wdArtHouseFunky               =73         # from enum WdPageBorderArt
	wdArtHypnotic                 =87         # from enum WdPageBorderArt
	wdArtIceCreamCones            =5          # from enum WdPageBorderArt
	wdArtLightBulb                =121        # from enum WdPageBorderArt
	wdArtLightning1               =53         # from enum WdPageBorderArt
	wdArtLightning2               =119        # from enum WdPageBorderArt
	wdArtMapPins                  =30         # from enum WdPageBorderArt
	wdArtMapleLeaf                =81         # from enum WdPageBorderArt
	wdArtMapleMuffins             =2          # from enum WdPageBorderArt
	wdArtMarquee                  =146        # from enum WdPageBorderArt
	wdArtMarqueeToothed           =131        # from enum WdPageBorderArt
	wdArtMoons                    =125        # from enum WdPageBorderArt
	wdArtMosaic                   =118        # from enum WdPageBorderArt
	wdArtMusicNotes               =79         # from enum WdPageBorderArt
	wdArtNorthwest                =104        # from enum WdPageBorderArt
	wdArtOvals                    =126        # from enum WdPageBorderArt
	wdArtPackages                 =26         # from enum WdPageBorderArt
	wdArtPalmsBlack               =80         # from enum WdPageBorderArt
	wdArtPalmsColor               =10         # from enum WdPageBorderArt
	wdArtPaperClips               =82         # from enum WdPageBorderArt
	wdArtPapyrus                  =92         # from enum WdPageBorderArt
	wdArtPartyFavor               =13         # from enum WdPageBorderArt
	wdArtPartyGlass               =7          # from enum WdPageBorderArt
	wdArtPencils                  =25         # from enum WdPageBorderArt
	wdArtPeople                   =84         # from enum WdPageBorderArt
	wdArtPeopleHats               =23         # from enum WdPageBorderArt
	wdArtPeopleWaving             =85         # from enum WdPageBorderArt
	wdArtPoinsettias              =40         # from enum WdPageBorderArt
	wdArtPostageStamp             =135        # from enum WdPageBorderArt
	wdArtPumpkin1                 =65         # from enum WdPageBorderArt
	wdArtPushPinNote1             =63         # from enum WdPageBorderArt
	wdArtPushPinNote2             =64         # from enum WdPageBorderArt
	wdArtPyramids                 =113        # from enum WdPageBorderArt
	wdArtPyramidsAbove            =114        # from enum WdPageBorderArt
	wdArtQuadrants                =60         # from enum WdPageBorderArt
	wdArtRings                    =29         # from enum WdPageBorderArt
	wdArtSafari                   =98         # from enum WdPageBorderArt
	wdArtSawtooth                 =133        # from enum WdPageBorderArt
	wdArtSawtoothGray             =134        # from enum WdPageBorderArt
	wdArtScaredCat                =36         # from enum WdPageBorderArt
	wdArtSeattle                  =78         # from enum WdPageBorderArt
	wdArtShadowedSquares          =57         # from enum WdPageBorderArt
	wdArtSharksTeeth              =132        # from enum WdPageBorderArt
	wdArtShorebirdTracks          =83         # from enum WdPageBorderArt
	wdArtSkyrocket                =77         # from enum WdPageBorderArt
	wdArtSnowflakeFancy           =76         # from enum WdPageBorderArt
	wdArtSnowflakes               =75         # from enum WdPageBorderArt
	wdArtSombrero                 =24         # from enum WdPageBorderArt
	wdArtSouthwest                =105        # from enum WdPageBorderArt
	wdArtStars                    =19         # from enum WdPageBorderArt
	wdArtStars3D                  =17         # from enum WdPageBorderArt
	wdArtStarsBlack               =74         # from enum WdPageBorderArt
	wdArtStarsShadowed            =18         # from enum WdPageBorderArt
	wdArtStarsTop                 =157        # from enum WdPageBorderArt
	wdArtSun                      =20         # from enum WdPageBorderArt
	wdArtSwirligig                =62         # from enum WdPageBorderArt
	wdArtTornPaper                =161        # from enum WdPageBorderArt
	wdArtTornPaperBlack           =162        # from enum WdPageBorderArt
	wdArtTrees                    =9          # from enum WdPageBorderArt
	wdArtTriangleParty            =123        # from enum WdPageBorderArt
	wdArtTriangles                =129        # from enum WdPageBorderArt
	wdArtTribal1                  =130        # from enum WdPageBorderArt
	wdArtTribal2                  =109        # from enum WdPageBorderArt
	wdArtTribal3                  =108        # from enum WdPageBorderArt
	wdArtTribal4                  =107        # from enum WdPageBorderArt
	wdArtTribal5                  =110        # from enum WdPageBorderArt
	wdArtTribal6                  =106        # from enum WdPageBorderArt
	wdArtTwistedLines1            =58         # from enum WdPageBorderArt
	wdArtTwistedLines2            =124        # from enum WdPageBorderArt
	wdArtVine                     =47         # from enum WdPageBorderArt
	wdArtWaveline                 =59         # from enum WdPageBorderArt
	wdArtWeavingAngles            =96         # from enum WdPageBorderArt
	wdArtWeavingBraid             =94         # from enum WdPageBorderArt
	wdArtWeavingRibbon            =95         # from enum WdPageBorderArt
	wdArtWeavingStrips            =136        # from enum WdPageBorderArt
	wdArtWhiteFlowers             =46         # from enum WdPageBorderArt
	wdArtWoodwork                 =93         # from enum WdPageBorderArt
	wdArtXIllusions               =111        # from enum WdPageBorderArt
	wdArtZanyTriangles            =112        # from enum WdPageBorderArt
	wdArtZigZag                   =137        # from enum WdPageBorderArt
	wdArtZigZagStitch             =143        # from enum WdPageBorderArt
	wdPageFitBestFit              =2          # from enum WdPageFit
	wdPageFitFullPage             =1          # from enum WdPageFit
	wdPageFitNone                 =0          # from enum WdPageFit
	wdPageFitTextFit              =3          # from enum WdPageFit
	wdAlignPageNumberCenter       =1          # from enum WdPageNumberAlignment
	wdAlignPageNumberInside       =3          # from enum WdPageNumberAlignment
	wdAlignPageNumberLeft         =0          # from enum WdPageNumberAlignment
	wdAlignPageNumberOutside      =4          # from enum WdPageNumberAlignment
	wdAlignPageNumberRight        =2          # from enum WdPageNumberAlignment
	wdPageNumberStyleArabic       =0          # from enum WdPageNumberStyle
	wdPageNumberStyleArabicFullWidth=14         # from enum WdPageNumberStyle
	wdPageNumberStyleArabicLetter1=46         # from enum WdPageNumberStyle
	wdPageNumberStyleArabicLetter2=48         # from enum WdPageNumberStyle
	wdPageNumberStyleHanjaRead    =41         # from enum WdPageNumberStyle
	wdPageNumberStyleHanjaReadDigit=42         # from enum WdPageNumberStyle
	wdPageNumberStyleHebrewLetter1=45         # from enum WdPageNumberStyle
	wdPageNumberStyleHebrewLetter2=47         # from enum WdPageNumberStyle
	wdPageNumberStyleHindiArabic  =51         # from enum WdPageNumberStyle
	wdPageNumberStyleHindiCardinalText=52         # from enum WdPageNumberStyle
	wdPageNumberStyleHindiLetter1 =49         # from enum WdPageNumberStyle
	wdPageNumberStyleHindiLetter2 =50         # from enum WdPageNumberStyle
	wdPageNumberStyleKanji        =10         # from enum WdPageNumberStyle
	wdPageNumberStyleKanjiDigit   =11         # from enum WdPageNumberStyle
	wdPageNumberStyleKanjiTraditional=16         # from enum WdPageNumberStyle
	wdPageNumberStyleLowercaseLetter=4          # from enum WdPageNumberStyle
	wdPageNumberStyleLowercaseRoman=2          # from enum WdPageNumberStyle
	wdPageNumberStyleNumberInCircle=18         # from enum WdPageNumberStyle
	wdPageNumberStyleNumberInDash =57         # from enum WdPageNumberStyle
	wdPageNumberStyleSimpChinNum1 =37         # from enum WdPageNumberStyle
	wdPageNumberStyleSimpChinNum2 =38         # from enum WdPageNumberStyle
	wdPageNumberStyleThaiArabic   =54         # from enum WdPageNumberStyle
	wdPageNumberStyleThaiCardinalText=55         # from enum WdPageNumberStyle
	wdPageNumberStyleThaiLetter   =53         # from enum WdPageNumberStyle
	wdPageNumberStyleTradChinNum1 =33         # from enum WdPageNumberStyle
	wdPageNumberStyleTradChinNum2 =34         # from enum WdPageNumberStyle
	wdPageNumberStyleUppercaseLetter=3          # from enum WdPageNumberStyle
	wdPageNumberStyleUppercaseRoman=1          # from enum WdPageNumberStyle
	wdPageNumberStyleVietCardinalText=56         # from enum WdPageNumberStyle
	emptyenum                     =0          # from enum WdPageNumberStyleHID
	wdPaper10x14                  =0          # from enum WdPaperSize
	wdPaper11x17                  =1          # from enum WdPaperSize
	wdPaperA3                     =6          # from enum WdPaperSize
	wdPaperA4                     =7          # from enum WdPaperSize
	wdPaperA4Small                =8          # from enum WdPaperSize
	wdPaperA5                     =9          # from enum WdPaperSize
	wdPaperB4                     =10         # from enum WdPaperSize
	wdPaperB5                     =11         # from enum WdPaperSize
	wdPaperCSheet                 =12         # from enum WdPaperSize
	wdPaperCustom                 =41         # from enum WdPaperSize
	wdPaperDSheet                 =13         # from enum WdPaperSize
	wdPaperESheet                 =14         # from enum WdPaperSize
	wdPaperEnvelope10             =25         # from enum WdPaperSize
	wdPaperEnvelope11             =26         # from enum WdPaperSize
	wdPaperEnvelope12             =27         # from enum WdPaperSize
	wdPaperEnvelope14             =28         # from enum WdPaperSize
	wdPaperEnvelope9              =24         # from enum WdPaperSize
	wdPaperEnvelopeB4             =29         # from enum WdPaperSize
	wdPaperEnvelopeB5             =30         # from enum WdPaperSize
	wdPaperEnvelopeB6             =31         # from enum WdPaperSize
	wdPaperEnvelopeC3             =32         # from enum WdPaperSize
	wdPaperEnvelopeC4             =33         # from enum WdPaperSize
	wdPaperEnvelopeC5             =34         # from enum WdPaperSize
	wdPaperEnvelopeC6             =35         # from enum WdPaperSize
	wdPaperEnvelopeC65            =36         # from enum WdPaperSize
	wdPaperEnvelopeDL             =37         # from enum WdPaperSize
	wdPaperEnvelopeItaly          =38         # from enum WdPaperSize
	wdPaperEnvelopeMonarch        =39         # from enum WdPaperSize
	wdPaperEnvelopePersonal       =40         # from enum WdPaperSize
	wdPaperExecutive              =5          # from enum WdPaperSize
	wdPaperFanfoldLegalGerman     =15         # from enum WdPaperSize
	wdPaperFanfoldStdGerman       =16         # from enum WdPaperSize
	wdPaperFanfoldUS              =17         # from enum WdPaperSize
	wdPaperFolio                  =18         # from enum WdPaperSize
	wdPaperLedger                 =19         # from enum WdPaperSize
	wdPaperLegal                  =4          # from enum WdPaperSize
	wdPaperLetter                 =2          # from enum WdPaperSize
	wdPaperLetterSmall            =3          # from enum WdPaperSize
	wdPaperNote                   =20         # from enum WdPaperSize
	wdPaperQuarto                 =21         # from enum WdPaperSize
	wdPaperStatement              =22         # from enum WdPaperSize
	wdPaperTabloid                =23         # from enum WdPaperSize
	wdPrinterAutomaticSheetFeed   =7          # from enum WdPaperTray
	wdPrinterDefaultBin           =0          # from enum WdPaperTray
	wdPrinterEnvelopeFeed         =5          # from enum WdPaperTray
	wdPrinterFormSource           =15         # from enum WdPaperTray
	wdPrinterLargeCapacityBin     =11         # from enum WdPaperTray
	wdPrinterLargeFormatBin       =10         # from enum WdPaperTray
	wdPrinterLowerBin             =2          # from enum WdPaperTray
	wdPrinterManualEnvelopeFeed   =6          # from enum WdPaperTray
	wdPrinterManualFeed           =4          # from enum WdPaperTray
	wdPrinterMiddleBin            =3          # from enum WdPaperTray
	wdPrinterOnlyBin              =1          # from enum WdPaperTray
	wdPrinterPaperCassette        =14         # from enum WdPaperTray
	wdPrinterSmallFormatBin       =9          # from enum WdPaperTray
	wdPrinterTractorFeed          =8          # from enum WdPaperTray
	wdPrinterUpperBin             =1          # from enum WdPaperTray
	wdAlignParagraphCenter        =1          # from enum WdParagraphAlignment
	wdAlignParagraphDistribute    =4          # from enum WdParagraphAlignment
	wdAlignParagraphJustify       =3          # from enum WdParagraphAlignment
	wdAlignParagraphJustifyHi     =7          # from enum WdParagraphAlignment
	wdAlignParagraphJustifyLow    =8          # from enum WdParagraphAlignment
	wdAlignParagraphJustifyMed    =5          # from enum WdParagraphAlignment
	wdAlignParagraphLeft          =0          # from enum WdParagraphAlignment
	wdAlignParagraphRight         =2          # from enum WdParagraphAlignment
	wdAlignParagraphThaiJustify   =9          # from enum WdParagraphAlignment
	emptyenum                     =0          # from enum WdParagraphAlignmentHID
	wdAdjective                   =0          # from enum WdPartOfSpeech
	wdAdverb                      =2          # from enum WdPartOfSpeech
	wdConjunction                 =5          # from enum WdPartOfSpeech
	wdIdiom                       =8          # from enum WdPartOfSpeech
	wdInterjection                =7          # from enum WdPartOfSpeech
	wdNoun                        =1          # from enum WdPartOfSpeech
	wdOther                       =9          # from enum WdPartOfSpeech
	wdPreposition                 =6          # from enum WdPartOfSpeech
	wdPronoun                     =4          # from enum WdPartOfSpeech
	wdVerb                        =3          # from enum WdPartOfSpeech
	wdPasteBitmap                 =4          # from enum WdPasteDataType
	wdPasteDeviceIndependentBitmap=5          # from enum WdPasteDataType
	wdPasteEnhancedMetafile       =9          # from enum WdPasteDataType
	wdPasteHTML                   =10         # from enum WdPasteDataType
	wdPasteHyperlink              =7          # from enum WdPasteDataType
	wdPasteMetafilePicture        =3          # from enum WdPasteDataType
	wdPasteOLEObject              =0          # from enum WdPasteDataType
	wdPasteRTF                    =1          # from enum WdPasteDataType
	wdPasteShape                  =8          # from enum WdPasteDataType
	wdPasteText                   =2          # from enum WdPasteDataType
	wdKeepSourceFormatting        =0          # from enum WdPasteOptions
	wdKeepTextOnly                =2          # from enum WdPasteOptions
	wdMatchDestinationFormatting  =1          # from enum WdPasteOptions
	wdUseDestinationStyles        =3          # from enum WdPasteOptions
	wdPhoneticGuideAlignmentCenter=0          # from enum WdPhoneticGuideAlignmentType
	wdPhoneticGuideAlignmentLeft  =3          # from enum WdPhoneticGuideAlignmentType
	wdPhoneticGuideAlignmentOneTwoOne=2          # from enum WdPhoneticGuideAlignmentType
	wdPhoneticGuideAlignmentRight =4          # from enum WdPhoneticGuideAlignmentType
	wdPhoneticGuideAlignmentRightVertical=5          # from enum WdPhoneticGuideAlignmentType
	wdPhoneticGuideAlignmentZeroOneZero=1          # from enum WdPhoneticGuideAlignmentType
	wdLinkDataInDoc               =1          # from enum WdPictureLinkType
	wdLinkDataOnDisk              =2          # from enum WdPictureLinkType
	wdLinkNone                    =0          # from enum WdPictureLinkType
	wdPortugueseBoth              =3          # from enum WdPortugueseReform
	wdPortuguesePostReform        =2          # from enum WdPortugueseReform
	wdPortuguesePreReform         =1          # from enum WdPortugueseReform
	wdPreferredWidthAuto          =1          # from enum WdPreferredWidthType
	wdPreferredWidthPercent       =2          # from enum WdPreferredWidthType
	wdPreferredWidthPoints        =3          # from enum WdPreferredWidthType
	wdPrintAutoTextEntries        =4          # from enum WdPrintOutItem
	wdPrintComments               =2          # from enum WdPrintOutItem
	wdPrintDocumentContent        =0          # from enum WdPrintOutItem
	wdPrintDocumentWithMarkup     =7          # from enum WdPrintOutItem
	wdPrintEnvelope               =6          # from enum WdPrintOutItem
	wdPrintKeyAssignments         =5          # from enum WdPrintOutItem
	wdPrintMarkup                 =2          # from enum WdPrintOutItem
	wdPrintProperties             =1          # from enum WdPrintOutItem
	wdPrintStyles                 =3          # from enum WdPrintOutItem
	wdPrintAllPages               =0          # from enum WdPrintOutPages
	wdPrintEvenPagesOnly          =2          # from enum WdPrintOutPages
	wdPrintOddPagesOnly           =1          # from enum WdPrintOutPages
	wdPrintAllDocument            =0          # from enum WdPrintOutRange
	wdPrintCurrentPage            =2          # from enum WdPrintOutRange
	wdPrintFromTo                 =3          # from enum WdPrintOutRange
	wdPrintRangeOfPages           =4          # from enum WdPrintOutRange
	wdPrintSelection              =1          # from enum WdPrintOutRange
	wdGrammaticalError            =1          # from enum WdProofreadingErrorType
	wdSpellingError               =0          # from enum WdProofreadingErrorType
	wdProtectedViewCloseEdit      =1          # from enum WdProtectedViewCloseReason
	wdProtectedViewCloseForced    =2          # from enum WdProtectedViewCloseReason
	wdProtectedViewCloseNormal    =0          # from enum WdProtectedViewCloseReason
	wdAllowOnlyComments           =1          # from enum WdProtectionType
	wdAllowOnlyFormFields         =2          # from enum WdProtectionType
	wdAllowOnlyReading            =3          # from enum WdProtectionType
	wdAllowOnlyRevisions          =0          # from enum WdProtectionType
	wdNoProtection                =-1         # from enum WdProtectionType
	wdAutomaticMargin             =0          # from enum WdReadingLayoutMargin
	wdFullMargin                  =2          # from enum WdReadingLayoutMargin
	wdSuppressMargin              =1          # from enum WdReadingLayoutMargin
	wdReadingOrderLtr             =1          # from enum WdReadingOrder
	wdReadingOrderRtl             =0          # from enum WdReadingOrder
	wdChart                       =14         # from enum WdRecoveryType
	wdChartLinked                 =15         # from enum WdRecoveryType
	wdChartPicture                =13         # from enum WdRecoveryType
	wdFormatOriginalFormatting    =16         # from enum WdRecoveryType
	wdFormatPlainText             =22         # from enum WdRecoveryType
	wdFormatSurroundingFormattingWithEmphasis=20         # from enum WdRecoveryType
	wdListCombineWithExistingList =24         # from enum WdRecoveryType
	wdListContinueNumbering       =7          # from enum WdRecoveryType
	wdListDontMerge               =25         # from enum WdRecoveryType
	wdListRestartNumbering        =8          # from enum WdRecoveryType
	wdPasteDefault                =0          # from enum WdRecoveryType
	wdSingleCellTable             =6          # from enum WdRecoveryType
	wdSingleCellText              =5          # from enum WdRecoveryType
	wdTableAppendTable            =10         # from enum WdRecoveryType
	wdTableInsertAsRows           =11         # from enum WdRecoveryType
	wdTableOriginalFormatting     =12         # from enum WdRecoveryType
	wdTableOverwriteCells         =23         # from enum WdRecoveryType
	wdUseDestinationStylesRecovery=19         # from enum WdRecoveryType
	wdDocumentControlRectangle    =13         # from enum WdRectangleType
	wdLineBetweenColumnRectangle  =5          # from enum WdRectangleType
	wdMailNavArea                 =12         # from enum WdRectangleType
	wdMarkupRectangle             =2          # from enum WdRectangleType
	wdMarkupRectangleArea         =8          # from enum WdRectangleType
	wdMarkupRectangleButton       =3          # from enum WdRectangleType
	wdMarkupRectangleMoveMatch    =10         # from enum WdRectangleType
	wdPageBorderRectangle         =4          # from enum WdRectangleType
	wdReadingModeNavigation       =9          # from enum WdRectangleType
	wdReadingModePanningArea      =11         # from enum WdRectangleType
	wdSelection                   =6          # from enum WdRectangleType
	wdShapeRectangle              =1          # from enum WdRectangleType
	wdSystem                      =7          # from enum WdRectangleType
	wdTextRectangle               =0          # from enum WdRectangleType
	wdContentText                 =-1         # from enum WdReferenceKind
	wdEndnoteNumber               =6          # from enum WdReferenceKind
	wdEndnoteNumberFormatted      =17         # from enum WdReferenceKind
	wdEntireCaption               =2          # from enum WdReferenceKind
	wdFootnoteNumber              =5          # from enum WdReferenceKind
	wdFootnoteNumberFormatted     =16         # from enum WdReferenceKind
	wdNumberFullContext           =-4         # from enum WdReferenceKind
	wdNumberNoContext             =-3         # from enum WdReferenceKind
	wdNumberRelativeContext       =-2         # from enum WdReferenceKind
	wdOnlyCaptionText             =4          # from enum WdReferenceKind
	wdOnlyLabelAndNumber          =3          # from enum WdReferenceKind
	wdPageNumber                  =7          # from enum WdReferenceKind
	wdPosition                    =15         # from enum WdReferenceKind
	wdRefTypeBookmark             =2          # from enum WdReferenceType
	wdRefTypeEndnote              =4          # from enum WdReferenceType
	wdRefTypeFootnote             =3          # from enum WdReferenceType
	wdRefTypeHeading              =1          # from enum WdReferenceType
	wdRefTypeNumberedItem         =0          # from enum WdReferenceType
	wdRelativeHorizontalPositionCharacter=3          # from enum WdRelativeHorizontalPosition
	wdRelativeHorizontalPositionColumn=2          # from enum WdRelativeHorizontalPosition
	wdRelativeHorizontalPositionInnerMarginArea=6          # from enum WdRelativeHorizontalPosition
	wdRelativeHorizontalPositionLeftMarginArea=4          # from enum WdRelativeHorizontalPosition
	wdRelativeHorizontalPositionMargin=0          # from enum WdRelativeHorizontalPosition
	wdRelativeHorizontalPositionOuterMarginArea=7          # from enum WdRelativeHorizontalPosition
	wdRelativeHorizontalPositionPage=1          # from enum WdRelativeHorizontalPosition
	wdRelativeHorizontalPositionRightMarginArea=5          # from enum WdRelativeHorizontalPosition
	wdRelativeHorizontalSizeInnerMarginArea=4          # from enum WdRelativeHorizontalSize
	wdRelativeHorizontalSizeLeftMarginArea=2          # from enum WdRelativeHorizontalSize
	wdRelativeHorizontalSizeMargin=0          # from enum WdRelativeHorizontalSize
	wdRelativeHorizontalSizeOuterMarginArea=5          # from enum WdRelativeHorizontalSize
	wdRelativeHorizontalSizePage  =1          # from enum WdRelativeHorizontalSize
	wdRelativeHorizontalSizeRightMarginArea=3          # from enum WdRelativeHorizontalSize
	wdRelativeVerticalPositionBottomMarginArea=5          # from enum WdRelativeVerticalPosition
	wdRelativeVerticalPositionInnerMarginArea=6          # from enum WdRelativeVerticalPosition
	wdRelativeVerticalPositionLine=3          # from enum WdRelativeVerticalPosition
	wdRelativeVerticalPositionMargin=0          # from enum WdRelativeVerticalPosition
	wdRelativeVerticalPositionOuterMarginArea=7          # from enum WdRelativeVerticalPosition
	wdRelativeVerticalPositionPage=1          # from enum WdRelativeVerticalPosition
	wdRelativeVerticalPositionParagraph=2          # from enum WdRelativeVerticalPosition
	wdRelativeVerticalPositionTopMarginArea=4          # from enum WdRelativeVerticalPosition
	wdRelativeVerticalSizeBottomMarginArea=3          # from enum WdRelativeVerticalSize
	wdRelativeVerticalSizeInnerMarginArea=4          # from enum WdRelativeVerticalSize
	wdRelativeVerticalSizeMargin  =0          # from enum WdRelativeVerticalSize
	wdRelativeVerticalSizeOuterMarginArea=5          # from enum WdRelativeVerticalSize
	wdRelativeVerticalSizePage    =1          # from enum WdRelativeVerticalSize
	wdRelativeVerticalSizeTopMarginArea=2          # from enum WdRelativeVerticalSize
	wdRelocateDown                =1          # from enum WdRelocate
	wdRelocateUp                  =0          # from enum WdRelocate
	wdRDIAll                      =99         # from enum WdRemoveDocInfoType
	wdRDIComments                 =1          # from enum WdRemoveDocInfoType
	wdRDIContentType              =16         # from enum WdRemoveDocInfoType
	wdRDIDocumentManagementPolicy =15         # from enum WdRemoveDocInfoType
	wdRDIDocumentProperties       =8          # from enum WdRemoveDocInfoType
	wdRDIDocumentServerProperties =14         # from enum WdRemoveDocInfoType
	wdRDIDocumentWorkspace        =10         # from enum WdRemoveDocInfoType
	wdRDIEmailHeader              =5          # from enum WdRemoveDocInfoType
	wdRDIInkAnnotations           =11         # from enum WdRemoveDocInfoType
	wdRDIRemovePersonalInformation=4          # from enum WdRemoveDocInfoType
	wdRDIRevisions                =2          # from enum WdRemoveDocInfoType
	wdRDIRoutingSlip              =6          # from enum WdRemoveDocInfoType
	wdRDISendForReview            =7          # from enum WdRemoveDocInfoType
	wdRDITemplate                 =9          # from enum WdRemoveDocInfoType
	wdRDIVersions                 =3          # from enum WdRemoveDocInfoType
	wdReplaceAll                  =2          # from enum WdReplace
	wdReplaceNone                 =0          # from enum WdReplace
	wdReplaceOne                  =1          # from enum WdReplace
	wdRevisedLinesMarkLeftBorder  =1          # from enum WdRevisedLinesMark
	wdRevisedLinesMarkNone        =0          # from enum WdRevisedLinesMark
	wdRevisedLinesMarkOutsideBorder=3          # from enum WdRevisedLinesMark
	wdRevisedLinesMarkRightBorder =2          # from enum WdRevisedLinesMark
	wdRevisedPropertiesMarkBold   =1          # from enum WdRevisedPropertiesMark
	wdRevisedPropertiesMarkColorOnly=5          # from enum WdRevisedPropertiesMark
	wdRevisedPropertiesMarkDoubleStrikeThrough=7          # from enum WdRevisedPropertiesMark
	wdRevisedPropertiesMarkDoubleUnderline=4          # from enum WdRevisedPropertiesMark
	wdRevisedPropertiesMarkItalic =2          # from enum WdRevisedPropertiesMark
	wdRevisedPropertiesMarkNone   =0          # from enum WdRevisedPropertiesMark
	wdRevisedPropertiesMarkStrikeThrough=6          # from enum WdRevisedPropertiesMark
	wdRevisedPropertiesMarkUnderline=3          # from enum WdRevisedPropertiesMark
	wdNoRevision                  =0          # from enum WdRevisionType
	wdRevisionCellDeletion        =17         # from enum WdRevisionType
	wdRevisionCellInsertion       =16         # from enum WdRevisionType
	wdRevisionCellMerge           =18         # from enum WdRevisionType
	wdRevisionCellSplit           =19         # from enum WdRevisionType
	wdRevisionConflict            =7          # from enum WdRevisionType
	wdRevisionConflictDelete      =21         # from enum WdRevisionType
	wdRevisionConflictInsert      =20         # from enum WdRevisionType
	wdRevisionDelete              =2          # from enum WdRevisionType
	wdRevisionDisplayField        =5          # from enum WdRevisionType
	wdRevisionInsert              =1          # from enum WdRevisionType
	wdRevisionMovedFrom           =14         # from enum WdRevisionType
	wdRevisionMovedTo             =15         # from enum WdRevisionType
	wdRevisionParagraphNumber     =4          # from enum WdRevisionType
	wdRevisionParagraphProperty   =10         # from enum WdRevisionType
	wdRevisionProperty            =3          # from enum WdRevisionType
	wdRevisionReconcile           =6          # from enum WdRevisionType
	wdRevisionReplace             =9          # from enum WdRevisionType
	wdRevisionSectionProperty     =12         # from enum WdRevisionType
	wdRevisionStyle               =8          # from enum WdRevisionType
	wdRevisionStyleDefinition     =13         # from enum WdRevisionType
	wdRevisionTableProperty       =11         # from enum WdRevisionType
	wdLeftMargin                  =0          # from enum WdRevisionsBalloonMargin
	wdRightMargin                 =1          # from enum WdRevisionsBalloonMargin
	wdBalloonPrintOrientationAuto =0          # from enum WdRevisionsBalloonPrintOrientation
	wdBalloonPrintOrientationForceLandscape=2          # from enum WdRevisionsBalloonPrintOrientation
	wdBalloonPrintOrientationPreserve=1          # from enum WdRevisionsBalloonPrintOrientation
	wdBalloonWidthPercent         =0          # from enum WdRevisionsBalloonWidthType
	wdBalloonWidthPoints          =1          # from enum WdRevisionsBalloonWidthType
	wdBalloonRevisions            =0          # from enum WdRevisionsMode
	wdInLineRevisions             =1          # from enum WdRevisionsMode
	wdMixedRevisions              =2          # from enum WdRevisionsMode
	wdRevisionsViewFinal          =0          # from enum WdRevisionsView
	wdRevisionsViewOriginal       =1          # from enum WdRevisionsView
	wdWrapAlways                  =1          # from enum WdRevisionsWrap
	wdWrapAsk                     =2          # from enum WdRevisionsWrap
	wdWrapNever                   =0          # from enum WdRevisionsWrap
	wdAllAtOnce                   =1          # from enum WdRoutingSlipDelivery
	wdOneAfterAnother             =0          # from enum WdRoutingSlipDelivery
	wdNotYetRouted                =0          # from enum WdRoutingSlipStatus
	wdRouteComplete               =2          # from enum WdRoutingSlipStatus
	wdRouteInProgress             =1          # from enum WdRoutingSlipStatus
	wdAlignRowCenter              =1          # from enum WdRowAlignment
	wdAlignRowLeft                =0          # from enum WdRowAlignment
	wdAlignRowRight               =2          # from enum WdRowAlignment
	wdRowHeightAtLeast            =1          # from enum WdRowHeightRule
	wdRowHeightAuto               =0          # from enum WdRowHeightRule
	wdRowHeightExactly            =2          # from enum WdRowHeightRule
	wdAdjustFirstColumn           =2          # from enum WdRulerStyle
	wdAdjustNone                  =0          # from enum WdRulerStyle
	wdAdjustProportional          =1          # from enum WdRulerStyle
	wdAdjustSameWidth             =3          # from enum WdRulerStyle
	wdGenderFemale                =0          # from enum WdSalutationGender
	wdGenderMale                  =1          # from enum WdSalutationGender
	wdGenderNeutral               =2          # from enum WdSalutationGender
	wdGenderUnknown               =3          # from enum WdSalutationGender
	wdSalutationBusiness          =2          # from enum WdSalutationType
	wdSalutationFormal            =1          # from enum WdSalutationType
	wdSalutationInformal          =0          # from enum WdSalutationType
	wdSalutationOther             =3          # from enum WdSalutationType
	wdFormatDOSText               =4          # from enum WdSaveFormat
	wdFormatDOSTextLineBreaks     =5          # from enum WdSaveFormat
	wdFormatDocument              =0          # from enum WdSaveFormat
	wdFormatDocument97            =0          # from enum WdSaveFormat
	wdFormatDocumentDefault       =16         # from enum WdSaveFormat
	wdFormatEncodedText           =7          # from enum WdSaveFormat
	wdFormatFilteredHTML          =10         # from enum WdSaveFormat
	wdFormatFlatXML               =19         # from enum WdSaveFormat
	wdFormatFlatXMLMacroEnabled   =20         # from enum WdSaveFormat
	wdFormatFlatXMLTemplate       =21         # from enum WdSaveFormat
	wdFormatFlatXMLTemplateMacroEnabled=22         # from enum WdSaveFormat
	wdFormatHTML                  =8          # from enum WdSaveFormat
	wdFormatOpenDocumentText      =23         # from enum WdSaveFormat
	wdFormatPDF                   =17         # from enum WdSaveFormat
	wdFormatRTF                   =6          # from enum WdSaveFormat
	wdFormatTemplate              =1          # from enum WdSaveFormat
	wdFormatTemplate97            =1          # from enum WdSaveFormat
	wdFormatText                  =2          # from enum WdSaveFormat
	wdFormatTextLineBreaks        =3          # from enum WdSaveFormat
	wdFormatUnicodeText           =7          # from enum WdSaveFormat
	wdFormatWebArchive            =9          # from enum WdSaveFormat
	wdFormatXML                   =11         # from enum WdSaveFormat
	wdFormatXMLDocument           =12         # from enum WdSaveFormat
	wdFormatXMLDocumentMacroEnabled=13         # from enum WdSaveFormat
	wdFormatXMLTemplate           =14         # from enum WdSaveFormat
	wdFormatXMLTemplateMacroEnabled=15         # from enum WdSaveFormat
	wdFormatXPS                   =18         # from enum WdSaveFormat
	wdDoNotSaveChanges            =0          # from enum WdSaveOptions
	wdPromptToSaveChanges         =-2         # from enum WdSaveOptions
	wdSaveChanges                 =-1         # from enum WdSaveOptions
	wdScrollbarTypeAuto           =0          # from enum WdScrollbarType
	wdScrollbarTypeNo             =2          # from enum WdScrollbarType
	wdScrollbarTypeYes            =1          # from enum WdScrollbarType
	wdSectionDirectionLtr         =1          # from enum WdSectionDirection
	wdSectionDirectionRtl         =0          # from enum WdSectionDirection
	wdSectionContinuous           =0          # from enum WdSectionStart
	wdSectionEvenPage             =3          # from enum WdSectionStart
	wdSectionNewColumn            =1          # from enum WdSectionStart
	wdSectionNewPage              =2          # from enum WdSectionStart
	wdSectionOddPage              =4          # from enum WdSectionStart
	wdSeekCurrentPageFooter       =10         # from enum WdSeekView
	wdSeekCurrentPageHeader       =9          # from enum WdSeekView
	wdSeekEndnotes                =8          # from enum WdSeekView
	wdSeekEvenPagesFooter         =6          # from enum WdSeekView
	wdSeekEvenPagesHeader         =3          # from enum WdSeekView
	wdSeekFirstPageFooter         =5          # from enum WdSeekView
	wdSeekFirstPageHeader         =2          # from enum WdSeekView
	wdSeekFootnotes               =7          # from enum WdSeekView
	wdSeekMainDocument            =0          # from enum WdSeekView
	wdSeekPrimaryFooter           =4          # from enum WdSeekView
	wdSeekPrimaryHeader           =1          # from enum WdSeekView
	wdSelActive                   =8          # from enum WdSelectionFlags
	wdSelAtEOL                    =2          # from enum WdSelectionFlags
	wdSelOvertype                 =4          # from enum WdSelectionFlags
	wdSelReplace                  =16         # from enum WdSelectionFlags
	wdSelStartActive              =1          # from enum WdSelectionFlags
	wdNoSelection                 =0          # from enum WdSelectionType
	wdSelectionBlock              =6          # from enum WdSelectionType
	wdSelectionColumn             =4          # from enum WdSelectionType
	wdSelectionFrame              =3          # from enum WdSelectionType
	wdSelectionIP                 =1          # from enum WdSelectionType
	wdSelectionInlineShape        =7          # from enum WdSelectionType
	wdSelectionNormal             =2          # from enum WdSelectionType
	wdSelectionRow                =5          # from enum WdSelectionType
	wdSelectionShape              =8          # from enum WdSelectionType
	wdSeparatorColon              =2          # from enum WdSeparatorType
	wdSeparatorEmDash             =3          # from enum WdSeparatorType
	wdSeparatorEnDash             =4          # from enum WdSeparatorType
	wdSeparatorHyphen             =0          # from enum WdSeparatorType
	wdSeparatorPeriod             =1          # from enum WdSeparatorType
	wdShapeBottom                 =-999997    # from enum WdShapePosition
	wdShapeCenter                 =-999995    # from enum WdShapePosition
	wdShapeInside                 =-999994    # from enum WdShapePosition
	wdShapeLeft                   =-999998    # from enum WdShapePosition
	wdShapeOutside                =-999993    # from enum WdShapePosition
	wdShapeRight                  =-999996    # from enum WdShapePosition
	wdShapeTop                    =-999999    # from enum WdShapePosition
	wdShapePositionRelativeNone   =-999999    # from enum WdShapePositionRelative
	wdShapeSizeRelativeNone       =-999999    # from enum WdShapeSizeRelative
	wdShowFilterFormattingAvailable=4          # from enum WdShowFilter
	wdShowFilterFormattingInUse   =3          # from enum WdShowFilter
	wdShowFilterFormattingRecommended=5          # from enum WdShowFilter
	wdShowFilterStylesAll         =2          # from enum WdShowFilter
	wdShowFilterStylesAvailable   =0          # from enum WdShowFilter
	wdShowFilterStylesInUse       =1          # from enum WdShowFilter
	wdShowSourceDocumentsBoth     =3          # from enum WdShowSourceDocuments
	wdShowSourceDocumentsNone     =0          # from enum WdShowSourceDocuments
	wdShowSourceDocumentsOriginal =1          # from enum WdShowSourceDocuments
	wdShowSourceDocumentsRevised  =2          # from enum WdShowSourceDocuments
	wdControlActiveX              =13         # from enum WdSmartTagControlType
	wdControlButton               =6          # from enum WdSmartTagControlType
	wdControlCheckbox             =9          # from enum WdSmartTagControlType
	wdControlCombo                =12         # from enum WdSmartTagControlType
	wdControlDocumentFragment     =14         # from enum WdSmartTagControlType
	wdControlDocumentFragmentURL  =15         # from enum WdSmartTagControlType
	wdControlHelp                 =3          # from enum WdSmartTagControlType
	wdControlHelpURL              =4          # from enum WdSmartTagControlType
	wdControlImage                =8          # from enum WdSmartTagControlType
	wdControlLabel                =7          # from enum WdSmartTagControlType
	wdControlLink                 =2          # from enum WdSmartTagControlType
	wdControlListbox              =11         # from enum WdSmartTagControlType
	wdControlRadioGroup           =16         # from enum WdSmartTagControlType
	wdControlSeparator            =5          # from enum WdSmartTagControlType
	wdControlSmartTag             =1          # from enum WdSmartTagControlType
	wdControlTextbox              =10         # from enum WdSmartTagControlType
	wdSortFieldAlphanumeric       =0          # from enum WdSortFieldType
	wdSortFieldDate               =2          # from enum WdSortFieldType
	wdSortFieldJapanJIS           =4          # from enum WdSortFieldType
	wdSortFieldKoreaKS            =6          # from enum WdSortFieldType
	wdSortFieldNumeric            =1          # from enum WdSortFieldType
	wdSortFieldStroke             =5          # from enum WdSortFieldType
	wdSortFieldSyllable           =3          # from enum WdSortFieldType
	emptyenum                     =0          # from enum WdSortFieldTypeHID
	wdSortOrderAscending          =0          # from enum WdSortOrder
	wdSortOrderDescending         =1          # from enum WdSortOrder
	wdSortSeparateByCommas        =1          # from enum WdSortSeparator
	wdSortSeparateByDefaultTableSeparator=2          # from enum WdSortSeparator
	wdSortSeparateByTabs          =0          # from enum WdSortSeparator
	wdSpanishTuteoAndVoseo        =1          # from enum WdSpanishSpeller
	wdSpanishTuteoOnly            =0          # from enum WdSpanishSpeller
	wdSpanishVoseoOnly            =2          # from enum WdSpanishSpeller
	wdPaneComments                =15         # from enum WdSpecialPane
	wdPaneCurrentPageFooter       =17         # from enum WdSpecialPane
	wdPaneCurrentPageHeader       =16         # from enum WdSpecialPane
	wdPaneEndnoteContinuationNotice=12         # from enum WdSpecialPane
	wdPaneEndnoteContinuationSeparator=13         # from enum WdSpecialPane
	wdPaneEndnoteSeparator        =14         # from enum WdSpecialPane
	wdPaneEndnotes                =8          # from enum WdSpecialPane
	wdPaneEvenPagesFooter         =6          # from enum WdSpecialPane
	wdPaneEvenPagesHeader         =3          # from enum WdSpecialPane
	wdPaneFirstPageFooter         =5          # from enum WdSpecialPane
	wdPaneFirstPageHeader         =2          # from enum WdSpecialPane
	wdPaneFootnoteContinuationNotice=9          # from enum WdSpecialPane
	wdPaneFootnoteContinuationSeparator=10         # from enum WdSpecialPane
	wdPaneFootnoteSeparator       =11         # from enum WdSpecialPane
	wdPaneFootnotes               =7          # from enum WdSpecialPane
	wdPaneNone                    =0          # from enum WdSpecialPane
	wdPanePrimaryFooter           =4          # from enum WdSpecialPane
	wdPanePrimaryHeader           =1          # from enum WdSpecialPane
	wdPaneRevisions               =18         # from enum WdSpecialPane
	wdPaneRevisionsHoriz          =19         # from enum WdSpecialPane
	wdPaneRevisionsVert           =20         # from enum WdSpecialPane
	wdSpellingCapitalization      =2          # from enum WdSpellingErrorType
	wdSpellingCorrect             =0          # from enum WdSpellingErrorType
	wdSpellingNotInDictionary     =1          # from enum WdSpellingErrorType
	wdAnagram                     =2          # from enum WdSpellingWordType
	wdSpellword                   =0          # from enum WdSpellingWordType
	wdWildcard                    =1          # from enum WdSpellingWordType
	wdStatisticCharacters         =3          # from enum WdStatistic
	wdStatisticCharactersWithSpaces=5          # from enum WdStatistic
	wdStatisticFarEastCharacters  =6          # from enum WdStatistic
	wdStatisticLines              =1          # from enum WdStatistic
	wdStatisticPages              =2          # from enum WdStatistic
	wdStatisticParagraphs         =4          # from enum WdStatistic
	wdStatisticWords              =0          # from enum WdStatistic
	emptyenum                     =0          # from enum WdStatisticHID
	wdCommentsStory               =4          # from enum WdStoryType
	wdEndnoteContinuationNoticeStory=17         # from enum WdStoryType
	wdEndnoteContinuationSeparatorStory=16         # from enum WdStoryType
	wdEndnoteSeparatorStory       =15         # from enum WdStoryType
	wdEndnotesStory               =3          # from enum WdStoryType
	wdEvenPagesFooterStory        =8          # from enum WdStoryType
	wdEvenPagesHeaderStory        =6          # from enum WdStoryType
	wdFirstPageFooterStory        =11         # from enum WdStoryType
	wdFirstPageHeaderStory        =10         # from enum WdStoryType
	wdFootnoteContinuationNoticeStory=14         # from enum WdStoryType
	wdFootnoteContinuationSeparatorStory=13         # from enum WdStoryType
	wdFootnoteSeparatorStory      =12         # from enum WdStoryType
	wdFootnotesStory              =2          # from enum WdStoryType
	wdMainTextStory               =1          # from enum WdStoryType
	wdPrimaryFooterStory          =9          # from enum WdStoryType
	wdPrimaryHeaderStory          =7          # from enum WdStoryType
	wdTextFrameStory              =5          # from enum WdStoryType
	wdStyleSheetLinkTypeImported  =1          # from enum WdStyleSheetLinkType
	wdStyleSheetLinkTypeLinked    =0          # from enum WdStyleSheetLinkType
	wdStyleSheetPrecedenceHigher  =-1         # from enum WdStyleSheetPrecedence
	wdStyleSheetPrecedenceHighest =1          # from enum WdStyleSheetPrecedence
	wdStyleSheetPrecedenceLower   =-2         # from enum WdStyleSheetPrecedence
	wdStyleSheetPrecedenceLowest  =0          # from enum WdStyleSheetPrecedence
	wdStyleSortByBasedOn          =3          # from enum WdStyleSort
	wdStyleSortByFont             =2          # from enum WdStyleSort
	wdStyleSortByName             =0          # from enum WdStyleSort
	wdStyleSortByType             =4          # from enum WdStyleSort
	wdStyleSortRecommended        =1          # from enum WdStyleSort
	wdStyleTypeCharacter          =2          # from enum WdStyleType
	wdStyleTypeLinked             =6          # from enum WdStyleType
	wdStyleTypeList               =4          # from enum WdStyleType
	wdStyleTypeParagraph          =1          # from enum WdStyleType
	wdStyleTypeParagraphOnly      =5          # from enum WdStyleType
	wdStyleTypeTable              =3          # from enum WdStyleType
	wdStylisticSet01              =1          # from enum WdStylisticSet
	wdStylisticSet02              =2          # from enum WdStylisticSet
	wdStylisticSet03              =4          # from enum WdStylisticSet
	wdStylisticSet04              =8          # from enum WdStylisticSet
	wdStylisticSet05              =16         # from enum WdStylisticSet
	wdStylisticSet06              =32         # from enum WdStylisticSet
	wdStylisticSet07              =64         # from enum WdStylisticSet
	wdStylisticSet08              =128        # from enum WdStylisticSet
	wdStylisticSet09              =256        # from enum WdStylisticSet
	wdStylisticSet10              =512        # from enum WdStylisticSet
	wdStylisticSet11              =1024       # from enum WdStylisticSet
	wdStylisticSet12              =2048       # from enum WdStylisticSet
	wdStylisticSet13              =4096       # from enum WdStylisticSet
	wdStylisticSet14              =8192       # from enum WdStylisticSet
	wdStylisticSet15              =16384      # from enum WdStylisticSet
	wdStylisticSet16              =32768      # from enum WdStylisticSet
	wdStylisticSet17              =65536      # from enum WdStylisticSet
	wdStylisticSet18              =131072     # from enum WdStylisticSet
	wdStylisticSet19              =262144     # from enum WdStylisticSet
	wdStylisticSet20              =524288     # from enum WdStylisticSet
	wdStylisticSetDefault         =0          # from enum WdStylisticSet
	wdSubscriberBestFormat        =0          # from enum WdSubscriberFormats
	wdSubscriberPict              =4          # from enum WdSubscriberFormats
	wdSubscriberRTF               =1          # from enum WdSubscriberFormats
	wdSubscriberText              =2          # from enum WdSubscriberFormats
	wd100Words                    =-4         # from enum WdSummaryLength
	wd10Percent                   =-6         # from enum WdSummaryLength
	wd10Sentences                 =-2         # from enum WdSummaryLength
	wd20Sentences                 =-3         # from enum WdSummaryLength
	wd25Percent                   =-7         # from enum WdSummaryLength
	wd500Words                    =-5         # from enum WdSummaryLength
	wd50Percent                   =-8         # from enum WdSummaryLength
	wd75Percent                   =-9         # from enum WdSummaryLength
	wdSummaryModeCreateNew        =3          # from enum WdSummaryMode
	wdSummaryModeHideAllButSummary=1          # from enum WdSummaryMode
	wdSummaryModeHighlight        =0          # from enum WdSummaryMode
	wdSummaryModeInsert           =2          # from enum WdSummaryMode
	wdTCSCConverterDirectionAuto  =2          # from enum WdTCSCConverterDirection
	wdTCSCConverterDirectionSCTC  =0          # from enum WdTCSCConverterDirection
	wdTCSCConverterDirectionTCSC  =1          # from enum WdTCSCConverterDirection
	wdAlignTabBar                 =4          # from enum WdTabAlignment
	wdAlignTabCenter              =1          # from enum WdTabAlignment
	wdAlignTabDecimal             =3          # from enum WdTabAlignment
	wdAlignTabLeft                =0          # from enum WdTabAlignment
	wdAlignTabList                =6          # from enum WdTabAlignment
	wdAlignTabRight               =2          # from enum WdTabAlignment
	wdTabLeaderDashes             =2          # from enum WdTabLeader
	wdTabLeaderDots               =1          # from enum WdTabLeader
	wdTabLeaderHeavy              =4          # from enum WdTabLeader
	wdTabLeaderLines              =3          # from enum WdTabLeader
	wdTabLeaderMiddleDot          =5          # from enum WdTabLeader
	wdTabLeaderSpaces             =0          # from enum WdTabLeader
	emptyenum                     =0          # from enum WdTabLeaderHID
	wdTableDirectionLtr           =1          # from enum WdTableDirection
	wdTableDirectionRtl           =0          # from enum WdTableDirection
	wdSeparateByCommas            =2          # from enum WdTableFieldSeparator
	wdSeparateByDefaultListSeparator=3          # from enum WdTableFieldSeparator
	wdSeparateByParagraphs        =0          # from enum WdTableFieldSeparator
	wdSeparateByTabs              =1          # from enum WdTableFieldSeparator
	wdTableFormat3DEffects1       =32         # from enum WdTableFormat
	wdTableFormat3DEffects2       =33         # from enum WdTableFormat
	wdTableFormat3DEffects3       =34         # from enum WdTableFormat
	wdTableFormatClassic1         =4          # from enum WdTableFormat
	wdTableFormatClassic2         =5          # from enum WdTableFormat
	wdTableFormatClassic3         =6          # from enum WdTableFormat
	wdTableFormatClassic4         =7          # from enum WdTableFormat
	wdTableFormatColorful1        =8          # from enum WdTableFormat
	wdTableFormatColorful2        =9          # from enum WdTableFormat
	wdTableFormatColorful3        =10         # from enum WdTableFormat
	wdTableFormatColumns1         =11         # from enum WdTableFormat
	wdTableFormatColumns2         =12         # from enum WdTableFormat
	wdTableFormatColumns3         =13         # from enum WdTableFormat
	wdTableFormatColumns4         =14         # from enum WdTableFormat
	wdTableFormatColumns5         =15         # from enum WdTableFormat
	wdTableFormatContemporary     =35         # from enum WdTableFormat
	wdTableFormatElegant          =36         # from enum WdTableFormat
	wdTableFormatGrid1            =16         # from enum WdTableFormat
	wdTableFormatGrid2            =17         # from enum WdTableFormat
	wdTableFormatGrid3            =18         # from enum WdTableFormat
	wdTableFormatGrid4            =19         # from enum WdTableFormat
	wdTableFormatGrid5            =20         # from enum WdTableFormat
	wdTableFormatGrid6            =21         # from enum WdTableFormat
	wdTableFormatGrid7            =22         # from enum WdTableFormat
	wdTableFormatGrid8            =23         # from enum WdTableFormat
	wdTableFormatList1            =24         # from enum WdTableFormat
	wdTableFormatList2            =25         # from enum WdTableFormat
	wdTableFormatList3            =26         # from enum WdTableFormat
	wdTableFormatList4            =27         # from enum WdTableFormat
	wdTableFormatList5            =28         # from enum WdTableFormat
	wdTableFormatList6            =29         # from enum WdTableFormat
	wdTableFormatList7            =30         # from enum WdTableFormat
	wdTableFormatList8            =31         # from enum WdTableFormat
	wdTableFormatNone             =0          # from enum WdTableFormat
	wdTableFormatProfessional     =37         # from enum WdTableFormat
	wdTableFormatSimple1          =1          # from enum WdTableFormat
	wdTableFormatSimple2          =2          # from enum WdTableFormat
	wdTableFormatSimple3          =3          # from enum WdTableFormat
	wdTableFormatSubtle1          =38         # from enum WdTableFormat
	wdTableFormatSubtle2          =39         # from enum WdTableFormat
	wdTableFormatWeb1             =40         # from enum WdTableFormat
	wdTableFormatWeb2             =41         # from enum WdTableFormat
	wdTableFormatWeb3             =42         # from enum WdTableFormat
	wdTableFormatApplyAutoFit     =16         # from enum WdTableFormatApply
	wdTableFormatApplyBorders     =1          # from enum WdTableFormatApply
	wdTableFormatApplyColor       =8          # from enum WdTableFormatApply
	wdTableFormatApplyFirstColumn =128        # from enum WdTableFormatApply
	wdTableFormatApplyFont        =4          # from enum WdTableFormatApply
	wdTableFormatApplyHeadingRows =32         # from enum WdTableFormatApply
	wdTableFormatApplyLastColumn  =256        # from enum WdTableFormatApply
	wdTableFormatApplyLastRow     =64         # from enum WdTableFormatApply
	wdTableFormatApplyShading     =2          # from enum WdTableFormatApply
	wdTableBottom                 =-999997    # from enum WdTablePosition
	wdTableCenter                 =-999995    # from enum WdTablePosition
	wdTableInside                 =-999994    # from enum WdTablePosition
	wdTableLeft                   =-999998    # from enum WdTablePosition
	wdTableOutside                =-999993    # from enum WdTablePosition
	wdTableRight                  =-999996    # from enum WdTablePosition
	wdTableTop                    =-999999    # from enum WdTablePosition
	wdTaskPaneApplyStyles         =17         # from enum WdTaskPanes
	wdTaskPaneDocumentActions     =7          # from enum WdTaskPanes
	wdTaskPaneDocumentManagement  =16         # from enum WdTaskPanes
	wdTaskPaneDocumentProtection  =6          # from enum WdTaskPanes
	wdTaskPaneDocumentUpdates     =13         # from enum WdTaskPanes
	wdTaskPaneFaxService          =11         # from enum WdTaskPanes
	wdTaskPaneFormatting          =0          # from enum WdTaskPanes
	wdTaskPaneHelp                =9          # from enum WdTaskPanes
	wdTaskPaneMailMerge           =2          # from enum WdTaskPanes
	wdTaskPaneNav                 =18         # from enum WdTaskPanes
	wdTaskPaneResearch            =10         # from enum WdTaskPanes
	wdTaskPaneRevealFormatting    =1          # from enum WdTaskPanes
	wdTaskPaneSearch              =4          # from enum WdTaskPanes
	wdTaskPaneSelection           =19         # from enum WdTaskPanes
	wdTaskPaneSharedWorkspace     =8          # from enum WdTaskPanes
	wdTaskPaneSignature           =14         # from enum WdTaskPanes
	wdTaskPaneStyleInspector      =15         # from enum WdTaskPanes
	wdTaskPaneTranslate           =3          # from enum WdTaskPanes
	wdTaskPaneXMLDocument         =12         # from enum WdTaskPanes
	wdTaskPaneXMLStructure        =5          # from enum WdTaskPanes
	wdAttachedTemplate            =2          # from enum WdTemplateType
	wdGlobalTemplate              =1          # from enum WdTemplateType
	wdNormalTemplate              =0          # from enum WdTemplateType
	wdCalculationText             =5          # from enum WdTextFormFieldType
	wdCurrentDateText             =3          # from enum WdTextFormFieldType
	wdCurrentTimeText             =4          # from enum WdTextFormFieldType
	wdDateText                    =2          # from enum WdTextFormFieldType
	wdNumberText                  =1          # from enum WdTextFormFieldType
	wdRegularText                 =0          # from enum WdTextFormFieldType
	wdTextOrientationDownward     =3          # from enum WdTextOrientation
	wdTextOrientationHorizontal   =0          # from enum WdTextOrientation
	wdTextOrientationHorizontalRotatedFarEast=4          # from enum WdTextOrientation
	wdTextOrientationUpward       =2          # from enum WdTextOrientation
	wdTextOrientationVertical     =5          # from enum WdTextOrientation
	wdTextOrientationVerticalFarEast=1          # from enum WdTextOrientation
	emptyenum                     =0          # from enum WdTextOrientationHID
	wdTightAll                    =1          # from enum WdTextboxTightWrap
	wdTightFirstAndLastLines      =2          # from enum WdTextboxTightWrap
	wdTightFirstLineOnly          =3          # from enum WdTextboxTightWrap
	wdTightLastLineOnly           =4          # from enum WdTextboxTightWrap
	wdTightNone                   =0          # from enum WdTextboxTightWrap
	wdTexture10Percent            =100        # from enum WdTextureIndex
	wdTexture12Pt5Percent         =125        # from enum WdTextureIndex
	wdTexture15Percent            =150        # from enum WdTextureIndex
	wdTexture17Pt5Percent         =175        # from enum WdTextureIndex
	wdTexture20Percent            =200        # from enum WdTextureIndex
	wdTexture22Pt5Percent         =225        # from enum WdTextureIndex
	wdTexture25Percent            =250        # from enum WdTextureIndex
	wdTexture27Pt5Percent         =275        # from enum WdTextureIndex
	wdTexture2Pt5Percent          =25         # from enum WdTextureIndex
	wdTexture30Percent            =300        # from enum WdTextureIndex
	wdTexture32Pt5Percent         =325        # from enum WdTextureIndex
	wdTexture35Percent            =350        # from enum WdTextureIndex
	wdTexture37Pt5Percent         =375        # from enum WdTextureIndex
	wdTexture40Percent            =400        # from enum WdTextureIndex
	wdTexture42Pt5Percent         =425        # from enum WdTextureIndex
	wdTexture45Percent            =450        # from enum WdTextureIndex
	wdTexture47Pt5Percent         =475        # from enum WdTextureIndex
	wdTexture50Percent            =500        # from enum WdTextureIndex
	wdTexture52Pt5Percent         =525        # from enum WdTextureIndex
	wdTexture55Percent            =550        # from enum WdTextureIndex
	wdTexture57Pt5Percent         =575        # from enum WdTextureIndex
	wdTexture5Percent             =50         # from enum WdTextureIndex
	wdTexture60Percent            =600        # from enum WdTextureIndex
	wdTexture62Pt5Percent         =625        # from enum WdTextureIndex
	wdTexture65Percent            =650        # from enum WdTextureIndex
	wdTexture67Pt5Percent         =675        # from enum WdTextureIndex
	wdTexture70Percent            =700        # from enum WdTextureIndex
	wdTexture72Pt5Percent         =725        # from enum WdTextureIndex
	wdTexture75Percent            =750        # from enum WdTextureIndex
	wdTexture77Pt5Percent         =775        # from enum WdTextureIndex
	wdTexture7Pt5Percent          =75         # from enum WdTextureIndex
	wdTexture80Percent            =800        # from enum WdTextureIndex
	wdTexture82Pt5Percent         =825        # from enum WdTextureIndex
	wdTexture85Percent            =850        # from enum WdTextureIndex
	wdTexture87Pt5Percent         =875        # from enum WdTextureIndex
	wdTexture90Percent            =900        # from enum WdTextureIndex
	wdTexture92Pt5Percent         =925        # from enum WdTextureIndex
	wdTexture95Percent            =950        # from enum WdTextureIndex
	wdTexture97Pt5Percent         =975        # from enum WdTextureIndex
	wdTextureCross                =-11        # from enum WdTextureIndex
	wdTextureDarkCross            =-5         # from enum WdTextureIndex
	wdTextureDarkDiagonalCross    =-6         # from enum WdTextureIndex
	wdTextureDarkDiagonalDown     =-3         # from enum WdTextureIndex
	wdTextureDarkDiagonalUp       =-4         # from enum WdTextureIndex
	wdTextureDarkHorizontal       =-1         # from enum WdTextureIndex
	wdTextureDarkVertical         =-2         # from enum WdTextureIndex
	wdTextureDiagonalCross        =-12        # from enum WdTextureIndex
	wdTextureDiagonalDown         =-9         # from enum WdTextureIndex
	wdTextureDiagonalUp           =-10        # from enum WdTextureIndex
	wdTextureHorizontal           =-7         # from enum WdTextureIndex
	wdTextureNone                 =0          # from enum WdTextureIndex
	wdTextureSolid                =1000       # from enum WdTextureIndex
	wdTextureVertical             =-8         # from enum WdTextureIndex
	wdNotThemeColor               =-1         # from enum WdThemeColorIndex
	wdThemeColorAccent1           =4          # from enum WdThemeColorIndex
	wdThemeColorAccent2           =5          # from enum WdThemeColorIndex
	wdThemeColorAccent3           =6          # from enum WdThemeColorIndex
	wdThemeColorAccent4           =7          # from enum WdThemeColorIndex
	wdThemeColorAccent5           =8          # from enum WdThemeColorIndex
	wdThemeColorAccent6           =9          # from enum WdThemeColorIndex
	wdThemeColorBackground1       =12         # from enum WdThemeColorIndex
	wdThemeColorBackground2       =14         # from enum WdThemeColorIndex
	wdThemeColorHyperlink         =10         # from enum WdThemeColorIndex
	wdThemeColorHyperlinkFollowed =11         # from enum WdThemeColorIndex
	wdThemeColorMainDark1         =0          # from enum WdThemeColorIndex
	wdThemeColorMainDark2         =2          # from enum WdThemeColorIndex
	wdThemeColorMainLight1        =1          # from enum WdThemeColorIndex
	wdThemeColorMainLight2        =3          # from enum WdThemeColorIndex
	wdThemeColorText1             =13         # from enum WdThemeColorIndex
	wdThemeColorText2             =15         # from enum WdThemeColorIndex
	wdTOAClassic                  =1          # from enum WdToaFormat
	wdTOADistinctive              =2          # from enum WdToaFormat
	wdTOAFormal                   =3          # from enum WdToaFormat
	wdTOASimple                   =4          # from enum WdToaFormat
	wdTOATemplate                 =0          # from enum WdToaFormat
	wdTOCClassic                  =1          # from enum WdTocFormat
	wdTOCDistinctive              =2          # from enum WdTocFormat
	wdTOCFancy                    =3          # from enum WdTocFormat
	wdTOCFormal                   =5          # from enum WdTocFormat
	wdTOCModern                   =4          # from enum WdTocFormat
	wdTOCSimple                   =6          # from enum WdTocFormat
	wdTOCTemplate                 =0          # from enum WdTocFormat
	wdTOFCentered                 =3          # from enum WdTofFormat
	wdTOFClassic                  =1          # from enum WdTofFormat
	wdTOFDistinctive              =2          # from enum WdTofFormat
	wdTOFFormal                   =4          # from enum WdTofFormat
	wdTOFSimple                   =5          # from enum WdTofFormat
	wdTOFTemplate                 =0          # from enum WdTofFormat
	wdTrailingNone                =2          # from enum WdTrailingCharacter
	wdTrailingSpace               =1          # from enum WdTrailingCharacter
	wdTrailingTab                 =0          # from enum WdTrailingCharacter
	wdTwoLinesInOneAngleBrackets  =4          # from enum WdTwoLinesInOneType
	wdTwoLinesInOneCurlyBrackets  =5          # from enum WdTwoLinesInOneType
	wdTwoLinesInOneNoBrackets     =1          # from enum WdTwoLinesInOneType
	wdTwoLinesInOneNone           =0          # from enum WdTwoLinesInOneType
	wdTwoLinesInOneParentheses    =2          # from enum WdTwoLinesInOneType
	wdTwoLinesInOneSquareBrackets =3          # from enum WdTwoLinesInOneType
	wdUnderlineDash               =7          # from enum WdUnderline
	wdUnderlineDashHeavy          =23         # from enum WdUnderline
	wdUnderlineDashLong           =39         # from enum WdUnderline
	wdUnderlineDashLongHeavy      =55         # from enum WdUnderline
	wdUnderlineDotDash            =9          # from enum WdUnderline
	wdUnderlineDotDashHeavy       =25         # from enum WdUnderline
	wdUnderlineDotDotDash         =10         # from enum WdUnderline
	wdUnderlineDotDotDashHeavy    =26         # from enum WdUnderline
	wdUnderlineDotted             =4          # from enum WdUnderline
	wdUnderlineDottedHeavy        =20         # from enum WdUnderline
	wdUnderlineDouble             =3          # from enum WdUnderline
	wdUnderlineNone               =0          # from enum WdUnderline
	wdUnderlineSingle             =1          # from enum WdUnderline
	wdUnderlineThick              =6          # from enum WdUnderline
	wdUnderlineWavy               =11         # from enum WdUnderline
	wdUnderlineWavyDouble         =43         # from enum WdUnderline
	wdUnderlineWavyHeavy          =27         # from enum WdUnderline
	wdUnderlineWords              =2          # from enum WdUnderline
	wdCell                        =12         # from enum WdUnits
	wdCharacter                   =1          # from enum WdUnits
	wdCharacterFormatting         =13         # from enum WdUnits
	wdColumn                      =9          # from enum WdUnits
	wdItem                        =16         # from enum WdUnits
	wdLine                        =5          # from enum WdUnits
	wdParagraph                   =4          # from enum WdUnits
	wdParagraphFormatting         =14         # from enum WdUnits
	wdRow                         =10         # from enum WdUnits
	wdScreen                      =7          # from enum WdUnits
	wdSection                     =8          # from enum WdUnits
	wdSentence                    =3          # from enum WdUnits
	wdStory                       =6          # from enum WdUnits
	wdTable                       =15         # from enum WdUnits
	wdWindow                      =11         # from enum WdUnits
	wdWord                        =2          # from enum WdUnits
	wdListBehaviorAddBulletsNumbering=1          # from enum WdUpdateStyleListBehavior
	wdListBehaviorKeepPreviousPattern=0          # from enum WdUpdateStyleListBehavior
	wdFormattingFromCurrent       =0          # from enum WdUseFormattingFrom
	wdFormattingFromPrompt        =2          # from enum WdUseFormattingFrom
	wdFormattingFromSelected      =1          # from enum WdUseFormattingFrom
	wdAlignVerticalBottom         =3          # from enum WdVerticalAlignment
	wdAlignVerticalCenter         =1          # from enum WdVerticalAlignment
	wdAlignVerticalJustify        =2          # from enum WdVerticalAlignment
	wdAlignVerticalTop            =0          # from enum WdVerticalAlignment
	wdConflictView                =8          # from enum WdViewType
	wdMasterView                  =5          # from enum WdViewType
	wdNormalView                  =1          # from enum WdViewType
	wdOutlineView                 =2          # from enum WdViewType
	wdPrintPreview                =4          # from enum WdViewType
	wdPrintView                   =3          # from enum WdViewType
	wdReadingView                 =7          # from enum WdViewType
	wdWebView                     =6          # from enum WdViewType
	wdOnlineView                  =6          # from enum WdViewTypeOld
	wdPageView                    =3          # from enum WdViewTypeOld
	wdVisualSelectionBlock        =0          # from enum WdVisualSelection
	wdVisualSelectionContinuous   =1          # from enum WdVisualSelection
	wdWindowStateMaximize         =1          # from enum WdWindowState
	wdWindowStateMinimize         =2          # from enum WdWindowState
	wdWindowStateNormal           =0          # from enum WdWindowState
	wdWindowDocument              =0          # from enum WdWindowType
	wdWindowTemplate              =1          # from enum WdWindowType
	wdDialogBuildingBlockOrganizer=2067       # from enum WdWordDialog
	wdDialogCSSLinks              =1261       # from enum WdWordDialog
	wdDialogCompatibilityChecker  =2439       # from enum WdWordDialog
	wdDialogConnect               =420        # from enum WdWordDialog
	wdDialogConsistencyChecker    =1121       # from enum WdWordDialog
	wdDialogContentControlProperties=2394       # from enum WdWordDialog
	wdDialogControlRun            =235        # from enum WdWordDialog
	wdDialogConvertObject         =392        # from enum WdWordDialog
	wdDialogCopyFile              =300        # from enum WdWordDialog
	wdDialogCreateAutoText        =872        # from enum WdWordDialog
	wdDialogCreateSource          =1922       # from enum WdWordDialog
	wdDialogDocumentInspector     =1482       # from enum WdWordDialog
	wdDialogDocumentStatistics    =78         # from enum WdWordDialog
	wdDialogDrawAlign             =634        # from enum WdWordDialog
	wdDialogDrawSnapToGrid        =633        # from enum WdWordDialog
	wdDialogEditAutoText          =985        # from enum WdWordDialog
	wdDialogEditCreatePublisher   =732        # from enum WdWordDialog
	wdDialogEditFind              =112        # from enum WdWordDialog
	wdDialogEditFrame             =458        # from enum WdWordDialog
	wdDialogEditGoTo              =896        # from enum WdWordDialog
	wdDialogEditGoToOld           =811        # from enum WdWordDialog
	wdDialogEditLinks             =124        # from enum WdWordDialog
	wdDialogEditObject            =125        # from enum WdWordDialog
	wdDialogEditPasteSpecial      =111        # from enum WdWordDialog
	wdDialogEditPublishOptions    =735        # from enum WdWordDialog
	wdDialogEditReplace           =117        # from enum WdWordDialog
	wdDialogEditStyle             =120        # from enum WdWordDialog
	wdDialogEditSubscribeOptions  =736        # from enum WdWordDialog
	wdDialogEditSubscribeTo       =733        # from enum WdWordDialog
	wdDialogEditTOACategory       =625        # from enum WdWordDialog
	wdDialogEmailOptions          =863        # from enum WdWordDialog
	wdDialogExportAsFixedFormat   =2349       # from enum WdWordDialog
	wdDialogFileDocumentLayout    =178        # from enum WdWordDialog
	wdDialogFileFind              =99         # from enum WdWordDialog
	wdDialogFileMacCustomPageSetupGX=737        # from enum WdWordDialog
	wdDialogFileMacPageSetup      =685        # from enum WdWordDialog
	wdDialogFileMacPageSetupGX    =444        # from enum WdWordDialog
	wdDialogFileNew               =79         # from enum WdWordDialog
	wdDialogFileNew2007           =1116       # from enum WdWordDialog
	wdDialogFileOpen              =80         # from enum WdWordDialog
	wdDialogFilePageSetup         =178        # from enum WdWordDialog
	wdDialogFilePrint             =88         # from enum WdWordDialog
	wdDialogFilePrintOneCopy      =445        # from enum WdWordDialog
	wdDialogFilePrintSetup        =97         # from enum WdWordDialog
	wdDialogFileRoutingSlip       =624        # from enum WdWordDialog
	wdDialogFileSaveAs            =84         # from enum WdWordDialog
	wdDialogFileSaveVersion       =1007       # from enum WdWordDialog
	wdDialogFileSummaryInfo       =86         # from enum WdWordDialog
	wdDialogFileVersions          =945        # from enum WdWordDialog
	wdDialogFitText               =983        # from enum WdWordDialog
	wdDialogFontSubstitution      =581        # from enum WdWordDialog
	wdDialogFormFieldHelp         =361        # from enum WdWordDialog
	wdDialogFormFieldOptions      =353        # from enum WdWordDialog
	wdDialogFormatAddrFonts       =103        # from enum WdWordDialog
	wdDialogFormatBordersAndShading=189        # from enum WdWordDialog
	wdDialogFormatBulletsAndNumbering=824        # from enum WdWordDialog
	wdDialogFormatCallout         =610        # from enum WdWordDialog
	wdDialogFormatChangeCase      =322        # from enum WdWordDialog
	wdDialogFormatColumns         =177        # from enum WdWordDialog
	wdDialogFormatDefineStyleBorders=185        # from enum WdWordDialog
	wdDialogFormatDefineStyleFont =181        # from enum WdWordDialog
	wdDialogFormatDefineStyleFrame=184        # from enum WdWordDialog
	wdDialogFormatDefineStyleLang =186        # from enum WdWordDialog
	wdDialogFormatDefineStylePara =182        # from enum WdWordDialog
	wdDialogFormatDefineStyleTabs =183        # from enum WdWordDialog
	wdDialogFormatDrawingObject   =960        # from enum WdWordDialog
	wdDialogFormatDropCap         =488        # from enum WdWordDialog
	wdDialogFormatEncloseCharacters=1162       # from enum WdWordDialog
	wdDialogFormatFont            =174        # from enum WdWordDialog
	wdDialogFormatFrame           =190        # from enum WdWordDialog
	wdDialogFormatPageNumber      =298        # from enum WdWordDialog
	wdDialogFormatParagraph       =175        # from enum WdWordDialog
	wdDialogFormatPicture         =187        # from enum WdWordDialog
	wdDialogFormatRetAddrFonts    =221        # from enum WdWordDialog
	wdDialogFormatSectionLayout   =176        # from enum WdWordDialog
	wdDialogFormatStyle           =180        # from enum WdWordDialog
	wdDialogFormatStyleGallery    =505        # from enum WdWordDialog
	wdDialogFormatStylesCustom    =1248       # from enum WdWordDialog
	wdDialogFormatTabs            =179        # from enum WdWordDialog
	wdDialogFormatTheme           =855        # from enum WdWordDialog
	wdDialogFormattingRestrictions=1427       # from enum WdWordDialog
	wdDialogFrameSetProperties    =1074       # from enum WdWordDialog
	wdDialogHelpAbout             =9          # from enum WdWordDialog
	wdDialogHelpWordPerfectHelp   =10         # from enum WdWordDialog
	wdDialogHelpWordPerfectHelpOptions=511        # from enum WdWordDialog
	wdDialogHorizontalInVertical  =1160       # from enum WdWordDialog
	wdDialogIMESetDefault         =1094       # from enum WdWordDialog
	wdDialogInsertAddCaption      =402        # from enum WdWordDialog
	wdDialogInsertAutoCaption     =359        # from enum WdWordDialog
	wdDialogInsertBookmark        =168        # from enum WdWordDialog
	wdDialogInsertBreak           =159        # from enum WdWordDialog
	wdDialogInsertCaption         =357        # from enum WdWordDialog
	wdDialogInsertCaptionNumbering=358        # from enum WdWordDialog
	wdDialogInsertCrossReference  =367        # from enum WdWordDialog
	wdDialogInsertDatabase        =341        # from enum WdWordDialog
	wdDialogInsertDateTime        =165        # from enum WdWordDialog
	wdDialogInsertField           =166        # from enum WdWordDialog
	wdDialogInsertFile            =164        # from enum WdWordDialog
	wdDialogInsertFootnote        =370        # from enum WdWordDialog
	wdDialogInsertFormField       =483        # from enum WdWordDialog
	wdDialogInsertHyperlink       =925        # from enum WdWordDialog
	wdDialogInsertIndex           =170        # from enum WdWordDialog
	wdDialogInsertIndexAndTables  =473        # from enum WdWordDialog
	wdDialogInsertMergeField      =167        # from enum WdWordDialog
	wdDialogInsertNumber          =812        # from enum WdWordDialog
	wdDialogInsertObject          =172        # from enum WdWordDialog
	wdDialogInsertPageNumbers     =294        # from enum WdWordDialog
	wdDialogInsertPicture         =163        # from enum WdWordDialog
	wdDialogInsertPlaceholder     =2348       # from enum WdWordDialog
	wdDialogInsertSource          =2120       # from enum WdWordDialog
	wdDialogInsertSubdocument     =583        # from enum WdWordDialog
	wdDialogInsertSymbol          =162        # from enum WdWordDialog
	wdDialogInsertTableOfAuthorities=471        # from enum WdWordDialog
	wdDialogInsertTableOfContents =171        # from enum WdWordDialog
	wdDialogInsertTableOfFigures  =472        # from enum WdWordDialog
	wdDialogInsertWebComponent    =1324       # from enum WdWordDialog
	wdDialogLabelOptions          =1367       # from enum WdWordDialog
	wdDialogLetterWizard          =821        # from enum WdWordDialog
	wdDialogListCommands          =723        # from enum WdWordDialog
	wdDialogMailMerge             =676        # from enum WdWordDialog
	wdDialogMailMergeCheck        =677        # from enum WdWordDialog
	wdDialogMailMergeCreateDataSource=642        # from enum WdWordDialog
	wdDialogMailMergeCreateHeaderSource=643        # from enum WdWordDialog
	wdDialogMailMergeFieldMapping =1304       # from enum WdWordDialog
	wdDialogMailMergeFindRecipient=1326       # from enum WdWordDialog
	wdDialogMailMergeFindRecord   =569        # from enum WdWordDialog
	wdDialogMailMergeHelper       =680        # from enum WdWordDialog
	wdDialogMailMergeInsertAddressBlock=1305       # from enum WdWordDialog
	wdDialogMailMergeInsertAsk    =4047       # from enum WdWordDialog
	wdDialogMailMergeInsertFields =1307       # from enum WdWordDialog
	wdDialogMailMergeInsertFillIn =4048       # from enum WdWordDialog
	wdDialogMailMergeInsertGreetingLine=1306       # from enum WdWordDialog
	wdDialogMailMergeInsertIf     =4049       # from enum WdWordDialog
	wdDialogMailMergeInsertNextIf =4053       # from enum WdWordDialog
	wdDialogMailMergeInsertSet    =4054       # from enum WdWordDialog
	wdDialogMailMergeInsertSkipIf =4055       # from enum WdWordDialog
	wdDialogMailMergeOpenDataSource=81         # from enum WdWordDialog
	wdDialogMailMergeOpenHeaderSource=82         # from enum WdWordDialog
	wdDialogMailMergeQueryOptions =681        # from enum WdWordDialog
	wdDialogMailMergeRecipients   =1308       # from enum WdWordDialog
	wdDialogMailMergeSetDocumentType=1339       # from enum WdWordDialog
	wdDialogMailMergeUseAddressBook=779        # from enum WdWordDialog
	wdDialogMarkCitation          =463        # from enum WdWordDialog
	wdDialogMarkIndexEntry        =169        # from enum WdWordDialog
	wdDialogMarkTableOfContentsEntry=442        # from enum WdWordDialog
	wdDialogMyPermission          =1437       # from enum WdWordDialog
	wdDialogNewToolbar            =586        # from enum WdWordDialog
	wdDialogNoteOptions           =373        # from enum WdWordDialog
	wdDialogOMathRecognizedFunctions=2165       # from enum WdWordDialog
	wdDialogOrganizer             =222        # from enum WdWordDialog
	wdDialogPermission            =1469       # from enum WdWordDialog
	wdDialogPhoneticGuide         =986        # from enum WdWordDialog
	wdDialogReviewAfmtRevisions   =570        # from enum WdWordDialog
	wdDialogSchemaLibrary         =1417       # from enum WdWordDialog
	wdDialogSearch                =1363       # from enum WdWordDialog
	wdDialogShowRepairs           =1381       # from enum WdWordDialog
	wdDialogSourceManager         =1920       # from enum WdWordDialog
	wdDialogStyleManagement       =1948       # from enum WdWordDialog
	wdDialogTCSCTranslator        =1156       # from enum WdWordDialog
	wdDialogTableAutoFormat       =563        # from enum WdWordDialog
	wdDialogTableCellOptions      =1081       # from enum WdWordDialog
	wdDialogTableColumnWidth      =143        # from enum WdWordDialog
	wdDialogTableDeleteCells      =133        # from enum WdWordDialog
	wdDialogTableFormatCell       =612        # from enum WdWordDialog
	wdDialogTableFormula          =348        # from enum WdWordDialog
	wdDialogTableInsertCells      =130        # from enum WdWordDialog
	wdDialogTableInsertRow        =131        # from enum WdWordDialog
	wdDialogTableInsertTable      =129        # from enum WdWordDialog
	wdDialogTableOfCaptionsOptions=551        # from enum WdWordDialog
	wdDialogTableOfContentsOptions=470        # from enum WdWordDialog
	wdDialogTableProperties       =861        # from enum WdWordDialog
	wdDialogTableRowHeight        =142        # from enum WdWordDialog
	wdDialogTableSort             =199        # from enum WdWordDialog
	wdDialogTableSplitCells       =137        # from enum WdWordDialog
	wdDialogTableTableOptions     =1080       # from enum WdWordDialog
	wdDialogTableToText           =128        # from enum WdWordDialog
	wdDialogTableWrapping         =854        # from enum WdWordDialog
	wdDialogTextToTable           =127        # from enum WdWordDialog
	wdDialogToolsAcceptRejectChanges=506        # from enum WdWordDialog
	wdDialogToolsAdvancedSettings =206        # from enum WdWordDialog
	wdDialogToolsAutoCorrect      =378        # from enum WdWordDialog
	wdDialogToolsAutoCorrectExceptions=762        # from enum WdWordDialog
	wdDialogToolsAutoManager      =915        # from enum WdWordDialog
	wdDialogToolsAutoSummarize    =874        # from enum WdWordDialog
	wdDialogToolsBulletsNumbers   =196        # from enum WdWordDialog
	wdDialogToolsCompareDocuments =198        # from enum WdWordDialog
	wdDialogToolsCreateDirectory  =833        # from enum WdWordDialog
	wdDialogToolsCreateEnvelope   =173        # from enum WdWordDialog
	wdDialogToolsCreateLabels     =489        # from enum WdWordDialog
	wdDialogToolsCustomize        =152        # from enum WdWordDialog
	wdDialogToolsCustomizeKeyboard=432        # from enum WdWordDialog
	wdDialogToolsCustomizeMenuBar =615        # from enum WdWordDialog
	wdDialogToolsCustomizeMenus   =433        # from enum WdWordDialog
	wdDialogToolsDictionary       =989        # from enum WdWordDialog
	wdDialogToolsEnvelopesAndLabels=607        # from enum WdWordDialog
	wdDialogToolsGrammarSettings  =885        # from enum WdWordDialog
	wdDialogToolsHangulHanjaConversion=784        # from enum WdWordDialog
	wdDialogToolsHighlightChanges =197        # from enum WdWordDialog
	wdDialogToolsHyphenation      =195        # from enum WdWordDialog
	wdDialogToolsLanguage         =188        # from enum WdWordDialog
	wdDialogToolsMacro            =215        # from enum WdWordDialog
	wdDialogToolsMacroRecord      =214        # from enum WdWordDialog
	wdDialogToolsManageFields     =631        # from enum WdWordDialog
	wdDialogToolsMergeDocuments   =435        # from enum WdWordDialog
	wdDialogToolsOptions          =974        # from enum WdWordDialog
	wdDialogToolsOptionsAutoFormat=959        # from enum WdWordDialog
	wdDialogToolsOptionsAutoFormatAsYouType=778        # from enum WdWordDialog
	wdDialogToolsOptionsBidi      =1029       # from enum WdWordDialog
	wdDialogToolsOptionsCompatibility=525        # from enum WdWordDialog
	wdDialogToolsOptionsEdit      =224        # from enum WdWordDialog
	wdDialogToolsOptionsEditCopyPaste=1356       # from enum WdWordDialog
	wdDialogToolsOptionsFileLocations=225        # from enum WdWordDialog
	wdDialogToolsOptionsFuzzy     =790        # from enum WdWordDialog
	wdDialogToolsOptionsGeneral   =203        # from enum WdWordDialog
	wdDialogToolsOptionsPrint     =208        # from enum WdWordDialog
	wdDialogToolsOptionsSave      =209        # from enum WdWordDialog
	wdDialogToolsOptionsSecurity  =1361       # from enum WdWordDialog
	wdDialogToolsOptionsSmartTag  =1395       # from enum WdWordDialog
	wdDialogToolsOptionsSpellingAndGrammar=211        # from enum WdWordDialog
	wdDialogToolsOptionsTrackChanges=386        # from enum WdWordDialog
	wdDialogToolsOptionsTypography=739        # from enum WdWordDialog
	wdDialogToolsOptionsUserInfo  =213        # from enum WdWordDialog
	wdDialogToolsOptionsView      =204        # from enum WdWordDialog
	wdDialogToolsProtectDocument  =503        # from enum WdWordDialog
	wdDialogToolsProtectSection   =578        # from enum WdWordDialog
	wdDialogToolsRevisions        =197        # from enum WdWordDialog
	wdDialogToolsSpellingAndGrammar=828        # from enum WdWordDialog
	wdDialogToolsTemplates        =87         # from enum WdWordDialog
	wdDialogToolsThesaurus        =194        # from enum WdWordDialog
	wdDialogToolsUnprotectDocument=521        # from enum WdWordDialog
	wdDialogToolsWordCount        =228        # from enum WdWordDialog
	wdDialogTwoLinesInOne         =1161       # from enum WdWordDialog
	wdDialogUpdateTOC             =331        # from enum WdWordDialog
	wdDialogViewZoom              =577        # from enum WdWordDialog
	wdDialogWebOptions            =898        # from enum WdWordDialog
	wdDialogWindowActivate        =220        # from enum WdWordDialog
	wdDialogXMLElementAttributes  =1460       # from enum WdWordDialog
	wdDialogXMLOptions            =1425       # from enum WdWordDialog
	emptyenum                     =0          # from enum WdWordDialogHID
	wdDialogEmailOptionsTabQuoting=1900002    # from enum WdWordDialogTab
	wdDialogEmailOptionsTabSignature=1900000    # from enum WdWordDialogTab
	wdDialogEmailOptionsTabStationary=1900001    # from enum WdWordDialogTab
	wdDialogFilePageSetupTabCharsLines=150004     # from enum WdWordDialogTab
	wdDialogFilePageSetupTabLayout=150003     # from enum WdWordDialogTab
	wdDialogFilePageSetupTabMargins=150000     # from enum WdWordDialogTab
	wdDialogFilePageSetupTabPaper =150001     # from enum WdWordDialogTab
	wdDialogFormatBordersAndShadingTabBorders=700000     # from enum WdWordDialogTab
	wdDialogFormatBordersAndShadingTabPageBorder=700001     # from enum WdWordDialogTab
	wdDialogFormatBordersAndShadingTabShading=700002     # from enum WdWordDialogTab
	wdDialogFormatBulletsAndNumberingTabBulleted=1500000    # from enum WdWordDialogTab
	wdDialogFormatBulletsAndNumberingTabNumbered=1500001    # from enum WdWordDialogTab
	wdDialogFormatBulletsAndNumberingTabOutlineNumbered=1500002    # from enum WdWordDialogTab
	wdDialogFormatDrawingObjectTabColorsAndLines=1200000    # from enum WdWordDialogTab
	wdDialogFormatDrawingObjectTabHR=1200007    # from enum WdWordDialogTab
	wdDialogFormatDrawingObjectTabPicture=1200004    # from enum WdWordDialogTab
	wdDialogFormatDrawingObjectTabPosition=1200002    # from enum WdWordDialogTab
	wdDialogFormatDrawingObjectTabSize=1200001    # from enum WdWordDialogTab
	wdDialogFormatDrawingObjectTabTextbox=1200005    # from enum WdWordDialogTab
	wdDialogFormatDrawingObjectTabWeb=1200006    # from enum WdWordDialogTab
	wdDialogFormatDrawingObjectTabWrapping=1200003    # from enum WdWordDialogTab
	wdDialogFormatFontTabAnimation=600002     # from enum WdWordDialogTab
	wdDialogFormatFontTabCharacterSpacing=600001     # from enum WdWordDialogTab
	wdDialogFormatFontTabFont     =600000     # from enum WdWordDialogTab
	wdDialogFormatParagraphTabIndentsAndSpacing=1000000    # from enum WdWordDialogTab
	wdDialogFormatParagraphTabTeisai=1000002    # from enum WdWordDialogTab
	wdDialogFormatParagraphTabTextFlow=1000001    # from enum WdWordDialogTab
	wdDialogInsertIndexAndTablesTabIndex=400000     # from enum WdWordDialogTab
	wdDialogInsertIndexAndTablesTabTableOfAuthorities=400003     # from enum WdWordDialogTab
	wdDialogInsertIndexAndTablesTabTableOfContents=400001     # from enum WdWordDialogTab
	wdDialogInsertIndexAndTablesTabTableOfFigures=400002     # from enum WdWordDialogTab
	wdDialogInsertSymbolTabSpecialCharacters=200001     # from enum WdWordDialogTab
	wdDialogInsertSymbolTabSymbols=200000     # from enum WdWordDialogTab
	wdDialogLetterWizardTabLetterFormat=1600000    # from enum WdWordDialogTab
	wdDialogLetterWizardTabOtherElements=1600002    # from enum WdWordDialogTab
	wdDialogLetterWizardTabRecipientInfo=1600001    # from enum WdWordDialogTab
	wdDialogLetterWizardTabSenderInfo=1600003    # from enum WdWordDialogTab
	wdDialogNoteOptionsTabAllEndnotes=300001     # from enum WdWordDialogTab
	wdDialogNoteOptionsTabAllFootnotes=300000     # from enum WdWordDialogTab
	wdDialogOrganizerTabAutoText  =500001     # from enum WdWordDialogTab
	wdDialogOrganizerTabCommandBars=500002     # from enum WdWordDialogTab
	wdDialogOrganizerTabMacros    =500003     # from enum WdWordDialogTab
	wdDialogOrganizerTabStyles    =500000     # from enum WdWordDialogTab
	wdDialogStyleManagementTabEdit=2200000    # from enum WdWordDialogTab
	wdDialogStyleManagementTabRecommend=2200001    # from enum WdWordDialogTab
	wdDialogStyleManagementTabRestrict=2200002    # from enum WdWordDialogTab
	wdDialogTablePropertiesTabCell=1800003    # from enum WdWordDialogTab
	wdDialogTablePropertiesTabColumn=1800002    # from enum WdWordDialogTab
	wdDialogTablePropertiesTabRow =1800001    # from enum WdWordDialogTab
	wdDialogTablePropertiesTabTable=1800000    # from enum WdWordDialogTab
	wdDialogTemplates             =2100000    # from enum WdWordDialogTab
	wdDialogTemplatesLinkedCSS    =2100003    # from enum WdWordDialogTab
	wdDialogTemplatesXMLExpansionPacks=2100002    # from enum WdWordDialogTab
	wdDialogTemplatesXMLSchema    =2100001    # from enum WdWordDialogTab
	wdDialogToolsAutoCorrectExceptionsTabFirstLetter=1400000    # from enum WdWordDialogTab
	wdDialogToolsAutoCorrectExceptionsTabHangulAndAlphabet=1400002    # from enum WdWordDialogTab
	wdDialogToolsAutoCorrectExceptionsTabIac=1400003    # from enum WdWordDialogTab
	wdDialogToolsAutoCorrectExceptionsTabInitialCaps=1400001    # from enum WdWordDialogTab
	wdDialogToolsAutoManagerTabAutoCorrect=1700000    # from enum WdWordDialogTab
	wdDialogToolsAutoManagerTabAutoFormat=1700003    # from enum WdWordDialogTab
	wdDialogToolsAutoManagerTabAutoFormatAsYouType=1700001    # from enum WdWordDialogTab
	wdDialogToolsAutoManagerTabAutoText=1700002    # from enum WdWordDialogTab
	wdDialogToolsAutoManagerTabSmartTags=1700004    # from enum WdWordDialogTab
	wdDialogToolsEnvelopesAndLabelsTabEnvelopes=800000     # from enum WdWordDialogTab
	wdDialogToolsEnvelopesAndLabelsTabLabels=800001     # from enum WdWordDialogTab
	wdDialogToolsOptionsTabAcetate=1266       # from enum WdWordDialogTab
	wdDialogToolsOptionsTabBidi   =1029       # from enum WdWordDialogTab
	wdDialogToolsOptionsTabCompatibility=525        # from enum WdWordDialogTab
	wdDialogToolsOptionsTabEdit   =224        # from enum WdWordDialogTab
	wdDialogToolsOptionsTabFileLocations=225        # from enum WdWordDialogTab
	wdDialogToolsOptionsTabFuzzy  =790        # from enum WdWordDialogTab
	wdDialogToolsOptionsTabGeneral=203        # from enum WdWordDialogTab
	wdDialogToolsOptionsTabHangulHanjaConversion=786        # from enum WdWordDialogTab
	wdDialogToolsOptionsTabPrint  =208        # from enum WdWordDialogTab
	wdDialogToolsOptionsTabProofread=211        # from enum WdWordDialogTab
	wdDialogToolsOptionsTabSave   =209        # from enum WdWordDialogTab
	wdDialogToolsOptionsTabSecurity=1361       # from enum WdWordDialogTab
	wdDialogToolsOptionsTabTrackChanges=386        # from enum WdWordDialogTab
	wdDialogToolsOptionsTabTypography=739        # from enum WdWordDialogTab
	wdDialogToolsOptionsTabUserInfo=213        # from enum WdWordDialogTab
	wdDialogToolsOptionsTabView   =204        # from enum WdWordDialogTab
	wdDialogWebOptionsBrowsers    =2000000    # from enum WdWordDialogTab
	wdDialogWebOptionsEncoding    =2000003    # from enum WdWordDialogTab
	wdDialogWebOptionsFiles       =2000001    # from enum WdWordDialogTab
	wdDialogWebOptionsFonts       =2000004    # from enum WdWordDialogTab
	wdDialogWebOptionsGeneral     =2000000    # from enum WdWordDialogTab
	wdDialogWebOptionsPictures    =2000002    # from enum WdWordDialogTab
	wdDialogFilePageSetupTabPaperSize=150001     # from enum WdWordDialogTabHID
	wdDialogFilePageSetupTabPaperSource=150002     # from enum WdWordDialogTabHID
	wdWrapBoth                    =0          # from enum WdWrapSideType
	wdWrapLargest                 =3          # from enum WdWrapSideType
	wdWrapLeft                    =1          # from enum WdWrapSideType
	wdWrapRight                   =2          # from enum WdWrapSideType
	wdWrapBehind                  =5          # from enum WdWrapType
	wdWrapFront                   =3          # from enum WdWrapType
	wdWrapInline                  =7          # from enum WdWrapType
	wdWrapNone                    =3          # from enum WdWrapType
	wdWrapSquare                  =0          # from enum WdWrapType
	wdWrapThrough                 =2          # from enum WdWrapType
	wdWrapTight                   =1          # from enum WdWrapType
	wdWrapTopBottom               =4          # from enum WdWrapType
	wdWrapMergeBehind             =3          # from enum WdWrapTypeMerged
	wdWrapMergeFront              =4          # from enum WdWrapTypeMerged
	wdWrapMergeInline             =0          # from enum WdWrapTypeMerged
	wdWrapMergeSquare             =1          # from enum WdWrapTypeMerged
	wdWrapMergeThrough            =5          # from enum WdWrapTypeMerged
	wdWrapMergeTight              =2          # from enum WdWrapTypeMerged
	wdWrapMergeTopBottom          =6          # from enum WdWrapTypeMerged
	wdXMLNodeLevelCell            =3          # from enum WdXMLNodeLevel
	wdXMLNodeLevelInline          =0          # from enum WdXMLNodeLevel
	wdXMLNodeLevelParagraph       =1          # from enum WdXMLNodeLevel
	wdXMLNodeLevelRow             =2          # from enum WdXMLNodeLevel
	wdXMLNodeAttribute            =2          # from enum WdXMLNodeType
	wdXMLNodeElement              =1          # from enum WdXMLNodeType
	wdXMLSelectionChangeReasonDelete=2          # from enum WdXMLSelectionChangeReason
	wdXMLSelectionChangeReasonInsert=1          # from enum WdXMLSelectionChangeReason
	wdXMLSelectionChangeReasonMove=0          # from enum WdXMLSelectionChangeReason
	wdXMLValidationStatusCustom   =-1072898048 # from enum WdXMLValidationStatus
	wdXMLValidationStatusOK       =0          # from enum WdXMLValidationStatus
	xlAxisCrossesAutomatic        =-4105      # from enum XlAxisCrosses
	xlAxisCrossesCustom           =-4114      # from enum XlAxisCrosses
	xlAxisCrossesMaximum          =2          # from enum XlAxisCrosses
	xlAxisCrossesMinimum          =4          # from enum XlAxisCrosses
	xlPrimary                     =1          # from enum XlAxisGroup
	xlSecondary                   =2          # from enum XlAxisGroup
	xlCategory                    =1          # from enum XlAxisType
	xlSeriesAxis                  =3          # from enum XlAxisType
	xlValue                       =2          # from enum XlAxisType
	xlBackgroundAutomatic         =-4105      # from enum XlBackground
	xlBackgroundOpaque            =3          # from enum XlBackground
	xlBackgroundTransparent       =2          # from enum XlBackground
	xlBox                         =0          # from enum XlBarShape
	xlConeToMax                   =5          # from enum XlBarShape
	xlConeToPoint                 =4          # from enum XlBarShape
	xlCylinder                    =3          # from enum XlBarShape
	xlPyramidToMax                =2          # from enum XlBarShape
	xlPyramidToPoint              =1          # from enum XlBarShape
	xlHairline                    =1          # from enum XlBorderWeight
	xlMedium                      =-4138      # from enum XlBorderWeight
	xlThick                       =4          # from enum XlBorderWeight
	xlThin                        =2          # from enum XlBorderWeight
	xlAutomaticScale              =-4105      # from enum XlCategoryType
	xlCategoryScale               =2          # from enum XlCategoryType
	xlTimeScale                   =3          # from enum XlCategoryType
	xlChartElementPositionAutomatic=-4105      # from enum XlChartElementPosition
	xlChartElementPositionCustom  =-4114      # from enum XlChartElementPosition
	xlAnyGallery                  =23         # from enum XlChartGallery
	xlBuiltIn                     =21         # from enum XlChartGallery
	xlUserDefined                 =22         # from enum XlChartGallery
	xlAxis                        =21         # from enum XlChartItem
	xlAxisTitle                   =17         # from enum XlChartItem
	xlChartArea                   =2          # from enum XlChartItem
	xlChartTitle                  =4          # from enum XlChartItem
	xlCorners                     =6          # from enum XlChartItem
	xlDataLabel                   =0          # from enum XlChartItem
	xlDataTable                   =7          # from enum XlChartItem
	xlDisplayUnitLabel            =30         # from enum XlChartItem
	xlDownBars                    =20         # from enum XlChartItem
	xlDropLines                   =26         # from enum XlChartItem
	xlErrorBars                   =9          # from enum XlChartItem
	xlFloor                       =23         # from enum XlChartItem
	xlHiLoLines                   =25         # from enum XlChartItem
	xlLeaderLines                 =29         # from enum XlChartItem
	xlLegend                      =24         # from enum XlChartItem
	xlLegendEntry                 =12         # from enum XlChartItem
	xlLegendKey                   =13         # from enum XlChartItem
	xlMajorGridlines              =15         # from enum XlChartItem
	xlMinorGridlines              =16         # from enum XlChartItem
	xlNothing                     =28         # from enum XlChartItem
	xlPivotChartDropZone          =32         # from enum XlChartItem
	xlPivotChartFieldButton       =31         # from enum XlChartItem
	xlPlotArea                    =19         # from enum XlChartItem
	xlRadarAxisLabels             =27         # from enum XlChartItem
	xlSeries                      =3          # from enum XlChartItem
	xlSeriesLines                 =22         # from enum XlChartItem
	xlShape                       =14         # from enum XlChartItem
	xlTrendline                   =8          # from enum XlChartItem
	xlUpBars                      =18         # from enum XlChartItem
	xlWalls                       =5          # from enum XlChartItem
	xlXErrorBars                  =10         # from enum XlChartItem
	xlYErrorBars                  =11         # from enum XlChartItem
	xlAllFaces                    =7          # from enum XlChartPicturePlacement
	xlEnd                         =2          # from enum XlChartPicturePlacement
	xlEndSides                    =3          # from enum XlChartPicturePlacement
	xlFront                       =4          # from enum XlChartPicturePlacement
	xlFrontEnd                    =6          # from enum XlChartPicturePlacement
	xlFrontSides                  =5          # from enum XlChartPicturePlacement
	xlSides                       =1          # from enum XlChartPicturePlacement
	xlStack                       =2          # from enum XlChartPictureType
	xlStackScale                  =3          # from enum XlChartPictureType
	xlStretch                     =1          # from enum XlChartPictureType
	xlSplitByCustomSplit          =4          # from enum XlChartSplitType
	xlSplitByPercentValue         =3          # from enum XlChartSplitType
	xlSplitByPosition             =1          # from enum XlChartSplitType
	xlSplitByValue                =2          # from enum XlChartSplitType
	xlColorIndexAutomatic         =-4105      # from enum XlColorIndex
	xlColorIndexNone              =-4142      # from enum XlColorIndex
	xl3DBar                       =-4099      # from enum XlConstants
	xl3DSurface                   =-4103      # from enum XlConstants
	xlAbove                       =0          # from enum XlConstants
	xlAutomatic                   =-4105      # from enum XlConstants
	xlBar                         =2          # from enum XlConstants
	xlBelow                       =1          # from enum XlConstants
	xlBoth                        =1          # from enum XlConstants
	xlBottom                      =-4107      # from enum XlConstants
	xlCenter                      =-4108      # from enum XlConstants
	xlChecker                     =9          # from enum XlConstants
	xlCircle                      =8          # from enum XlConstants
	xlColumn                      =3          # from enum XlConstants
	xlCombination                 =-4111      # from enum XlConstants
	xlCorner                      =2          # from enum XlConstants
	xlCrissCross                  =16         # from enum XlConstants
	xlCross                       =4          # from enum XlConstants
	xlCustom                      =-4114      # from enum XlConstants
	xlDefaultAutoFormat           =-1         # from enum XlConstants
	xlDiamond                     =2          # from enum XlConstants
	xlDistributed                 =-4117      # from enum XlConstants
	xlFill                        =5          # from enum XlConstants
	xlFixedValue                  =1          # from enum XlConstants
	xlGeneral                     =1          # from enum XlConstants
	xlGray16                      =17         # from enum XlConstants
	xlGray25                      =-4124      # from enum XlConstants
	xlGray50                      =-4125      # from enum XlConstants
	xlGray75                      =-4126      # from enum XlConstants
	xlGray8                       =18         # from enum XlConstants
	xlGrid                        =15         # from enum XlConstants
	xlHigh                        =-4127      # from enum XlConstants
	xlInside                      =2          # from enum XlConstants
	xlJustify                     =-4130      # from enum XlConstants
	xlLeft                        =-4131      # from enum XlConstants
	xlLightDown                   =13         # from enum XlConstants
	xlLightHorizontal             =11         # from enum XlConstants
	xlLightUp                     =14         # from enum XlConstants
	xlLightVertical               =12         # from enum XlConstants
	xlLow                         =-4134      # from enum XlConstants
	xlMaximum                     =2          # from enum XlConstants
	xlMinimum                     =4          # from enum XlConstants
	xlMinusValues                 =3          # from enum XlConstants
	xlNextToAxis                  =4          # from enum XlConstants
	xlNone                        =-4142      # from enum XlConstants
	xlOpaque                      =3          # from enum XlConstants
	xlOutside                     =3          # from enum XlConstants
	xlPercent                     =2          # from enum XlConstants
	xlPlus                        =9          # from enum XlConstants
	xlPlusValues                  =2          # from enum XlConstants
	xlRight                       =-4152      # from enum XlConstants
	xlScale                       =3          # from enum XlConstants
	xlSemiGray75                  =10         # from enum XlConstants
	xlShowLabel                   =4          # from enum XlConstants
	xlShowLabelAndPercent         =5          # from enum XlConstants
	xlShowPercent                 =3          # from enum XlConstants
	xlShowValue                   =2          # from enum XlConstants
	xlSingle                      =2          # from enum XlConstants
	xlSolid                       =1          # from enum XlConstants
	xlSquare                      =1          # from enum XlConstants
	xlStError                     =4          # from enum XlConstants
	xlStar                        =5          # from enum XlConstants
	xlTop                         =-4160      # from enum XlConstants
	xlTransparent                 =2          # from enum XlConstants
	xlTriangle                    =3          # from enum XlConstants
	xlBitmap                      =2          # from enum XlCopyPictureFormat
	xlPicture                     =-4147      # from enum XlCopyPictureFormat
	xlLabelPositionAbove          =0          # from enum XlDataLabelPosition
	xlLabelPositionBelow          =1          # from enum XlDataLabelPosition
	xlLabelPositionBestFit        =5          # from enum XlDataLabelPosition
	xlLabelPositionCenter         =-4108      # from enum XlDataLabelPosition
	xlLabelPositionCustom         =7          # from enum XlDataLabelPosition
	xlLabelPositionInsideBase     =4          # from enum XlDataLabelPosition
	xlLabelPositionInsideEnd      =3          # from enum XlDataLabelPosition
	xlLabelPositionLeft           =-4131      # from enum XlDataLabelPosition
	xlLabelPositionMixed          =6          # from enum XlDataLabelPosition
	xlLabelPositionOutsideEnd     =2          # from enum XlDataLabelPosition
	xlLabelPositionRight          =-4152      # from enum XlDataLabelPosition
	xlDataLabelSeparatorDefault   =1          # from enum XlDataLabelSeparator
	xlDataLabelsShowBubbleSizes   =6          # from enum XlDataLabelsType
	xlDataLabelsShowLabel         =4          # from enum XlDataLabelsType
	xlDataLabelsShowLabelAndPercent=5          # from enum XlDataLabelsType
	xlDataLabelsShowNone          =-4142      # from enum XlDataLabelsType
	xlDataLabelsShowPercent       =3          # from enum XlDataLabelsType
	xlDataLabelsShowValue         =2          # from enum XlDataLabelsType
	xlInterpolated                =3          # from enum XlDisplayBlanksAs
	xlNotPlotted                  =1          # from enum XlDisplayBlanksAs
	xlZero                        =2          # from enum XlDisplayBlanksAs
	xlHundredMillions             =-8         # from enum XlDisplayUnit
	xlHundredThousands            =-5         # from enum XlDisplayUnit
	xlHundreds                    =-2         # from enum XlDisplayUnit
	xlMillionMillions             =-10        # from enum XlDisplayUnit
	xlMillions                    =-6         # from enum XlDisplayUnit
	xlTenMillions                 =-7         # from enum XlDisplayUnit
	xlTenThousands                =-4         # from enum XlDisplayUnit
	xlThousandMillions            =-9         # from enum XlDisplayUnit
	xlThousands                   =-3         # from enum XlDisplayUnit
	xlCap                         =1          # from enum XlEndStyleCap
	xlNoCap                       =2          # from enum XlEndStyleCap
	xlChartX                      =-4168      # from enum XlErrorBarDirection
	xlChartY                      =1          # from enum XlErrorBarDirection
	xlErrorBarIncludeBoth         =1          # from enum XlErrorBarInclude
	xlErrorBarIncludeMinusValues  =3          # from enum XlErrorBarInclude
	xlErrorBarIncludeNone         =-4142      # from enum XlErrorBarInclude
	xlErrorBarIncludePlusValues   =2          # from enum XlErrorBarInclude
	xlErrorBarTypeCustom          =-4114      # from enum XlErrorBarType
	xlErrorBarTypeFixedValue      =1          # from enum XlErrorBarType
	xlErrorBarTypePercent         =2          # from enum XlErrorBarType
	xlErrorBarTypeStDev           =-4155      # from enum XlErrorBarType
	xlErrorBarTypeStError         =4          # from enum XlErrorBarType
	xlHAlignCenter                =-4108      # from enum XlHAlign
	xlHAlignCenterAcrossSelection =7          # from enum XlHAlign
	xlHAlignDistributed           =-4117      # from enum XlHAlign
	xlHAlignFill                  =5          # from enum XlHAlign
	xlHAlignGeneral               =1          # from enum XlHAlign
	xlHAlignJustify               =-4130      # from enum XlHAlign
	xlHAlignLeft                  =-4131      # from enum XlHAlign
	xlHAlignRight                 =-4152      # from enum XlHAlign
	xlLegendPositionBottom        =-4107      # from enum XlLegendPosition
	xlLegendPositionCorner        =2          # from enum XlLegendPosition
	xlLegendPositionCustom        =-4161      # from enum XlLegendPosition
	xlLegendPositionLeft          =-4131      # from enum XlLegendPosition
	xlLegendPositionRight         =-4152      # from enum XlLegendPosition
	xlLegendPositionTop           =-4160      # from enum XlLegendPosition
	xlContinuous                  =1          # from enum XlLineStyle
	xlDash                        =-4115      # from enum XlLineStyle
	xlDashDot                     =4          # from enum XlLineStyle
	xlDashDotDot                  =5          # from enum XlLineStyle
	xlDot                         =-4118      # from enum XlLineStyle
	xlDouble                      =-4119      # from enum XlLineStyle
	xlLineStyleNone               =-4142      # from enum XlLineStyle
	xlSlantDashDot                =13         # from enum XlLineStyle
	xlMarkerStyleAutomatic        =-4105      # from enum XlMarkerStyle
	xlMarkerStyleCircle           =8          # from enum XlMarkerStyle
	xlMarkerStyleDash             =-4115      # from enum XlMarkerStyle
	xlMarkerStyleDiamond          =2          # from enum XlMarkerStyle
	xlMarkerStyleDot              =-4118      # from enum XlMarkerStyle
	xlMarkerStyleNone             =-4142      # from enum XlMarkerStyle
	xlMarkerStylePicture          =-4147      # from enum XlMarkerStyle
	xlMarkerStylePlus             =9          # from enum XlMarkerStyle
	xlMarkerStyleSquare           =1          # from enum XlMarkerStyle
	xlMarkerStyleStar             =5          # from enum XlMarkerStyle
	xlMarkerStyleTriangle         =3          # from enum XlMarkerStyle
	xlMarkerStyleX                =-4168      # from enum XlMarkerStyle
	xlDownward                    =-4170      # from enum XlOrientation
	xlHorizontal                  =-4128      # from enum XlOrientation
	xlUpward                      =-4171      # from enum XlOrientation
	xlVertical                    =-4166      # from enum XlOrientation
	xlPatternAutomatic            =-4105      # from enum XlPattern
	xlPatternChecker              =9          # from enum XlPattern
	xlPatternCrissCross           =16         # from enum XlPattern
	xlPatternDown                 =-4121      # from enum XlPattern
	xlPatternGray16               =17         # from enum XlPattern
	xlPatternGray25               =-4124      # from enum XlPattern
	xlPatternGray50               =-4125      # from enum XlPattern
	xlPatternGray75               =-4126      # from enum XlPattern
	xlPatternGray8                =18         # from enum XlPattern
	xlPatternGrid                 =15         # from enum XlPattern
	xlPatternHorizontal           =-4128      # from enum XlPattern
	xlPatternLightDown            =13         # from enum XlPattern
	xlPatternLightHorizontal      =11         # from enum XlPattern
	xlPatternLightUp              =14         # from enum XlPattern
	xlPatternLightVertical        =12         # from enum XlPattern
	xlPatternLinearGradient       =4000       # from enum XlPattern
	xlPatternNone                 =-4142      # from enum XlPattern
	xlPatternRectangularGradient  =4001       # from enum XlPattern
	xlPatternSemiGray75           =10         # from enum XlPattern
	xlPatternSolid                =1          # from enum XlPattern
	xlPatternUp                   =-4162      # from enum XlPattern
	xlPatternVertical             =-4166      # from enum XlPattern
	xlPrinter                     =2          # from enum XlPictureAppearance
	xlScreen                      =1          # from enum XlPictureAppearance
	xlCenterPoint                 =5          # from enum XlPieSliceIndex
	xlInnerCenterPoint            =8          # from enum XlPieSliceIndex
	xlInnerClockwisePoint         =7          # from enum XlPieSliceIndex
	xlInnerCounterClockwisePoint  =9          # from enum XlPieSliceIndex
	xlMidClockwiseRadiusPoint     =4          # from enum XlPieSliceIndex
	xlMidCounterClockwiseRadiusPoint=6          # from enum XlPieSliceIndex
	xlOuterCenterPoint            =2          # from enum XlPieSliceIndex
	xlOuterClockwisePoint         =3          # from enum XlPieSliceIndex
	xlOuterCounterClockwisePoint  =1          # from enum XlPieSliceIndex
	xlHorizontalCoordinate        =1          # from enum XlPieSliceLocation
	xlVerticalCoordinate          =2          # from enum XlPieSliceLocation
	xlColumnField                 =2          # from enum XlPivotFieldOrientation
	xlDataField                   =4          # from enum XlPivotFieldOrientation
	xlHidden                      =0          # from enum XlPivotFieldOrientation
	xlPageField                   =3          # from enum XlPivotFieldOrientation
	xlRowField                    =1          # from enum XlPivotFieldOrientation
	xlContext                     =-5002      # from enum XlReadingOrder
	xlLTR                         =-5003      # from enum XlReadingOrder
	xlRTL                         =-5004      # from enum XlReadingOrder
	xlAliceBlue                   =16775408   # from enum XlRgbColor
	xlAntiqueWhite                =14150650   # from enum XlRgbColor
	xlAqua                        =16776960   # from enum XlRgbColor
	xlAquamarine                  =13959039   # from enum XlRgbColor
	xlAzure                       =16777200   # from enum XlRgbColor
	xlBeige                       =14480885   # from enum XlRgbColor
	xlBisque                      =12903679   # from enum XlRgbColor
	xlBlack                       =0          # from enum XlRgbColor
	xlBlanchedAlmond              =13495295   # from enum XlRgbColor
	xlBlue                        =16711680   # from enum XlRgbColor
	xlBlueViolet                  =14822282   # from enum XlRgbColor
	xlBrown                       =2763429    # from enum XlRgbColor
	xlBurlyWood                   =8894686    # from enum XlRgbColor
	xlCadetBlue                   =10526303   # from enum XlRgbColor
	xlChartreuse                  =65407      # from enum XlRgbColor
	xlCoral                       =5275647    # from enum XlRgbColor
	xlCornflowerBlue              =15570276   # from enum XlRgbColor
	xlCornsilk                    =14481663   # from enum XlRgbColor
	xlCrimson                     =3937500    # from enum XlRgbColor
	xlDarkBlue                    =9109504    # from enum XlRgbColor
	xlDarkCyan                    =9145088    # from enum XlRgbColor
	xlDarkGoldenrod               =755384     # from enum XlRgbColor
	xlDarkGray                    =11119017   # from enum XlRgbColor
	xlDarkGreen                   =25600      # from enum XlRgbColor
	xlDarkGrey                    =11119017   # from enum XlRgbColor
	xlDarkKhaki                   =7059389    # from enum XlRgbColor
	xlDarkMagenta                 =9109643    # from enum XlRgbColor
	xlDarkOliveGreen              =3107669    # from enum XlRgbColor
	xlDarkOrange                  =36095      # from enum XlRgbColor
	xlDarkOrchid                  =13382297   # from enum XlRgbColor
	xlDarkRed                     =139        # from enum XlRgbColor
	xlDarkSalmon                  =8034025    # from enum XlRgbColor
	xlDarkSeaGreen                =9419919    # from enum XlRgbColor
	xlDarkSlateBlue               =9125192    # from enum XlRgbColor
	xlDarkSlateGray               =5197615    # from enum XlRgbColor
	xlDarkSlateGrey               =5197615    # from enum XlRgbColor
	xlDarkTurquoise               =13749760   # from enum XlRgbColor
	xlDarkViolet                  =13828244   # from enum XlRgbColor
	xlDeepPink                    =9639167    # from enum XlRgbColor
	xlDeepSkyBlue                 =16760576   # from enum XlRgbColor
	xlDimGray                     =6908265    # from enum XlRgbColor
	xlDimGrey                     =6908265    # from enum XlRgbColor
	xlDodgerBlue                  =16748574   # from enum XlRgbColor
	xlFireBrick                   =2237106    # from enum XlRgbColor
	xlFloralWhite                 =15792895   # from enum XlRgbColor
	xlForestGreen                 =2263842    # from enum XlRgbColor
	xlFuchsia                     =16711935   # from enum XlRgbColor
	xlGainsboro                   =14474460   # from enum XlRgbColor
	xlGhostWhite                  =16775416   # from enum XlRgbColor
	xlGold                        =55295      # from enum XlRgbColor
	xlGoldenrod                   =2139610    # from enum XlRgbColor
	xlGray                        =8421504    # from enum XlRgbColor
	xlGreen                       =32768      # from enum XlRgbColor
	xlGreenYellow                 =3145645    # from enum XlRgbColor
	xlGrey                        =8421504    # from enum XlRgbColor
	xlHoneydew                    =15794160   # from enum XlRgbColor
	xlHotPink                     =11823615   # from enum XlRgbColor
	xlIndianRed                   =6053069    # from enum XlRgbColor
	xlIndigo                      =8519755    # from enum XlRgbColor
	xlIvory                       =15794175   # from enum XlRgbColor
	xlKhaki                       =9234160    # from enum XlRgbColor
	xlLavender                    =16443110   # from enum XlRgbColor
	xlLavenderBlush               =16118015   # from enum XlRgbColor
	xlLawnGreen                   =64636      # from enum XlRgbColor
	xlLemonChiffon                =13499135   # from enum XlRgbColor
	xlLightBlue                   =15128749   # from enum XlRgbColor
	xlLightCoral                  =8421616    # from enum XlRgbColor
	xlLightCyan                   =9145088    # from enum XlRgbColor
	xlLightGoldenrodYellow        =13826810   # from enum XlRgbColor
	xlLightGray                   =13882323   # from enum XlRgbColor
	xlLightGreen                  =9498256    # from enum XlRgbColor
	xlLightGrey                   =13882323   # from enum XlRgbColor
	xlLightPink                   =12695295   # from enum XlRgbColor
	xlLightSalmon                 =8036607    # from enum XlRgbColor
	xlLightSeaGreen               =11186720   # from enum XlRgbColor
	xlLightSkyBlue                =16436871   # from enum XlRgbColor
	xlLightSlateGray              =10061943   # from enum XlRgbColor
	xlLightSlateGrey              =10061943   # from enum XlRgbColor
	xlLightSteelBlue              =14599344   # from enum XlRgbColor
	xlLightYellow                 =14745599   # from enum XlRgbColor
	xlLime                        =65280      # from enum XlRgbColor
	xlLimeGreen                   =3329330    # from enum XlRgbColor
	xlLinen                       =15134970   # from enum XlRgbColor
	xlMaroon                      =128        # from enum XlRgbColor
	xlMediumAquamarine            =11206502   # from enum XlRgbColor
	xlMediumBlue                  =13434880   # from enum XlRgbColor
	xlMediumOrchid                =13850042   # from enum XlRgbColor
	xlMediumPurple                =14381203   # from enum XlRgbColor
	xlMediumSeaGreen              =7451452    # from enum XlRgbColor
	xlMediumSlateBlue             =15624315   # from enum XlRgbColor
	xlMediumSpringGreen           =10156544   # from enum XlRgbColor
	xlMediumTurquoise             =13422920   # from enum XlRgbColor
	xlMediumVioletRed             =8721863    # from enum XlRgbColor
	xlMidnightBlue                =7346457    # from enum XlRgbColor
	xlMintCream                   =16449525   # from enum XlRgbColor
	xlMistyRose                   =14804223   # from enum XlRgbColor
	xlMoccasin                    =11920639   # from enum XlRgbColor
	xlNavajoWhite                 =11394815   # from enum XlRgbColor
	xlNavy                        =8388608    # from enum XlRgbColor
	xlNavyBlue                    =8388608    # from enum XlRgbColor
	xlOldLace                     =15136253   # from enum XlRgbColor
	xlOlive                       =32896      # from enum XlRgbColor
	xlOliveDrab                   =2330219    # from enum XlRgbColor
	xlOrange                      =42495      # from enum XlRgbColor
	xlOrangeRed                   =17919      # from enum XlRgbColor
	xlOrchid                      =14053594   # from enum XlRgbColor
	xlPaleGoldenrod               =7071982    # from enum XlRgbColor
	xlPaleGreen                   =10025880   # from enum XlRgbColor
	xlPaleTurquoise               =15658671   # from enum XlRgbColor
	xlPaleVioletRed               =9662683    # from enum XlRgbColor
	xlPapayaWhip                  =14020607   # from enum XlRgbColor
	xlPeachPuff                   =12180223   # from enum XlRgbColor
	xlPeru                        =4163021    # from enum XlRgbColor
	xlPink                        =13353215   # from enum XlRgbColor
	xlPlum                        =14524637   # from enum XlRgbColor
	xlPowderBlue                  =15130800   # from enum XlRgbColor
	xlPurple                      =8388736    # from enum XlRgbColor
	xlRed                         =255        # from enum XlRgbColor
	xlRosyBrown                   =9408444    # from enum XlRgbColor
	xlRoyalBlue                   =14772545   # from enum XlRgbColor
	xlSalmon                      =7504122    # from enum XlRgbColor
	xlSandyBrown                  =6333684    # from enum XlRgbColor
	xlSeaGreen                    =5737262    # from enum XlRgbColor
	xlSeashell                    =15660543   # from enum XlRgbColor
	xlSienna                      =2970272    # from enum XlRgbColor
	xlSilver                      =12632256   # from enum XlRgbColor
	xlSkyBlue                     =15453831   # from enum XlRgbColor
	xlSlateBlue                   =13458026   # from enum XlRgbColor
	xlSlateGray                   =9470064    # from enum XlRgbColor
	xlSlateGrey                   =9470064    # from enum XlRgbColor
	xlSnow                        =16448255   # from enum XlRgbColor
	xlSpringGreen                 =8388352    # from enum XlRgbColor
	xlSteelBlue                   =11829830   # from enum XlRgbColor
	xlTan                         =9221330    # from enum XlRgbColor
	xlTeal                        =8421376    # from enum XlRgbColor
	xlThistle                     =14204888   # from enum XlRgbColor
	xlTomato                      =4678655    # from enum XlRgbColor
	xlTurquoise                   =13688896   # from enum XlRgbColor
	xlViolet                      =15631086   # from enum XlRgbColor
	xlWheat                       =11788021   # from enum XlRgbColor
	xlWhite                       =16777215   # from enum XlRgbColor
	xlWhiteSmoke                  =16119285   # from enum XlRgbColor
	xlYellow                      =65535      # from enum XlRgbColor
	xlYellowGreen                 =3329434    # from enum XlRgbColor
	xlColumns                     =2          # from enum XlRowCol
	xlRows                        =1          # from enum XlRowCol
	xlScaleLinear                 =-4132      # from enum XlScaleType
	xlScaleLogarithmic            =-4133      # from enum XlScaleType
	xlSizeIsArea                  =1          # from enum XlSizeRepresents
	xlSizeIsWidth                 =2          # from enum XlSizeRepresents
	xlTickLabelOrientationAutomatic=-4105      # from enum XlTickLabelOrientation
	xlTickLabelOrientationDownward=-4170      # from enum XlTickLabelOrientation
	xlTickLabelOrientationHorizontal=-4128      # from enum XlTickLabelOrientation
	xlTickLabelOrientationUpward  =-4171      # from enum XlTickLabelOrientation
	xlTickLabelOrientationVertical=-4166      # from enum XlTickLabelOrientation
	xlTickLabelPositionHigh       =-4127      # from enum XlTickLabelPosition
	xlTickLabelPositionLow        =-4134      # from enum XlTickLabelPosition
	xlTickLabelPositionNextToAxis =4          # from enum XlTickLabelPosition
	xlTickLabelPositionNone       =-4142      # from enum XlTickLabelPosition
	xlTickMarkCross               =4          # from enum XlTickMark
	xlTickMarkInside              =2          # from enum XlTickMark
	xlTickMarkNone                =-4142      # from enum XlTickMark
	xlTickMarkOutside             =3          # from enum XlTickMark
	xlDays                        =0          # from enum XlTimeUnit
	xlMonths                      =1          # from enum XlTimeUnit
	xlYears                       =2          # from enum XlTimeUnit
	xlExponential                 =5          # from enum XlTrendlineType
	xlLinear                      =-4132      # from enum XlTrendlineType
	xlLogarithmic                 =-4133      # from enum XlTrendlineType
	xlMovingAvg                   =6          # from enum XlTrendlineType
	xlPolynomial                  =3          # from enum XlTrendlineType
	xlPower                       =4          # from enum XlTrendlineType
	xlUnderlineStyleDouble        =-4119      # from enum XlUnderlineStyle
	xlUnderlineStyleDoubleAccounting=5          # from enum XlUnderlineStyle
	xlUnderlineStyleNone          =-4142      # from enum XlUnderlineStyle
	xlUnderlineStyleSingle        =2          # from enum XlUnderlineStyle
	xlUnderlineStyleSingleAccounting=4          # from enum XlUnderlineStyle
	xlVAlignBottom                =-4107      # from enum XlVAlign
	xlVAlignCenter                =-4108      # from enum XlVAlign
	xlVAlignDistributed           =-4117      # from enum XlVAlign
	xlVAlignJustify               =-4130      # from enum XlVAlign
	xlVAlignTop                   =-4160      # from enum XlVAlign

RecordMap = {
}

CLSIDToClassMap = {}
CLSIDToPackageMap = {
	'{935D59F5-E365-4F92-B7F5-1C499A63ECA8}' : 'TickLabels',
	'{AFAF0C0E-8603-40F6-8FD1-42726CAC21E3}' : 'OMathScrPre',
	'{CDB0FF41-E862-47BB-AE77-3FA7B1AE3189}' : 'ChartFont',
	'{AE6CE2F5-B9D3-407D-85A8-0F10C63289A4}' : 'Line',
	'{74DE9576-8E99-4E28-912B-CB30747C60CE}' : 'OMathLimLow',
	'{7D0F7985-68D9-4D93-91CB-8109280E76CC}' : 'Rectangles',
	'{54F46DC4-F6A6-48CC-BD66-46C1DDEADD22}' : 'ContentControlListEntries',
	'{0C1FABE7-F737-406F-9CA3-B07661F9D1A2}' : 'XMLMapping',
	'{CAE36175-3818-4C60-BCBF-0645D51EB33B}' : 'OMathMatCol',
	'{79635BF1-BD1D-4B3F-A520-C1106F1AAAD8}' : 'Break',
	'{B5828B50-0E3D-448A-962D-A40702A5868D}' : 'BuildingBlockTypes',
	'{256B6ABA-6A38-4D39-971C-91FDA9922814}' : 'CoAuthors',
	'{A2E94180-7564-4D97-806B-BBC0D0A1350C}' : 'Walls',
	'{B3A1E8C6-E1CE-4A46-8D12-E017157B03D7}' : 'Legend',
	'{C75AD98A-74E9-49FE-8BF1-544839CC08A5}' : 'ChartArea',
	'{6F9D1F68-06F7-49EF-8902-185E54EB5E87}' : 'OMathAutoCorrect',
	'{56AFD330-440C-4F4C-A39C-ED306D084D5F}' : 'PlotArea',
	'{FF06FEF2-DA89-41C0-A0A8-5CD434E210AD}' : 'ChartCharacters',
	'{86488FB4-9633-4C93-8057-FC1FA7A847AE}' : 'ChartGroup',
	'{DD8F80B8-9B80-4E89-9BEC-F12DF35E43B3}' : 'ChartColorFormat',
	'{DFB6AA6C-1068-420F-969D-01280FCC1630}' : 'SmartTagAction',
	'{000209FF-0000-0000-C000-000000000046}' : 'Application',
	'{D8779F01-4869-4403-B334-D60C5F9C9175}' : 'OMathAutoCorrectEntry',
	'{7EBC66BD-F788-42C3-91F4-E8C841A69005}' : 'Axis',
	'{C774F5EA-A539-4284-A1BE-30AEC052D899}' : 'XSLTransforms',
	'{00020A00-0001-0000-C000-000000000046}' : 'IApplicationEvents3',
	'{8FEB78F7-35C6-4871-918C-193C3CDD886D}' : 'SeriesCollection',
	'{91807402-6C6F-47CD-B8FA-C42FEE8EE924}' : 'Pages',
	'{00020A01-0001-0000-C000-000000000046}' : 'IApplicationEvents4',
	'{AED7E08C-14F0-4F33-921D-4C5353137BF6}' : 'Editors',
	'{5D311669-EA51-11D3-87CC-00105AA31A34}' : 'MappedDataField',
	'{F1B14F40-5C32-4C8C-B5B2-DE537BB6B89D}' : 'GlowFormat',
	'{C6D50987-25D7-408A-BFF2-90BF86A24E93}' : 'BuildingBlocks',
	'{E59544D5-C299-46A0-84C1-C51AB38F9759}' : 'CoAuthor',
	'{BEA85A24-D7DA-4F3D-B58C-ED90FB01D615}' : 'FootnoteOptions',
	'{5D7F6C15-36CE-44CC-9692-5A1F8B8C906D}' : 'SeriesLines',
	'{DF076FDE-8781-4051-A5BC-99F6B7DC04D4}' : 'LegendKey',
	'{A87E00E9-3AC3-4B53-ABE3-7379653D0E82}' : 'XMLChildNodeSuggestion',
	'{36162C62-B59A-4278-AF3D-F2AC1EB999D9}' : 'LeaderLines',
	'{DD947D72-F33C-4198-9BDF-F86181D05E41}' : 'Editor',
	'{C04865A3-9F8A-486C-BB58-B4C3E6563136}' : 'DisplayUnitLabel',
	'{02B17CB4-7D55-4B34-B38B-10381433441F}' : 'OMathGroupChar',
	'{B66D3C1A-4541-4961-B35B-A353C03F6A99}' : 'ChartFormat',
	'{F2B60A10-DED5-46FB-A914-3C6F4EBB6451}' : 'SmartTagRecognizers',
	'{E2E0F3A7-204C-40C5-BAA5-290F374FDF5A}' : 'OMathBreaks',
	'{DE63B5AC-CA4F-46FE-9184-A5719AB9ED5B}' : 'XMLChildNodeSuggestions',
	'{5E9A888C-E5DC-4DCB-8308-3C91FB61E6F4}' : 'SmartTagType',
	'{D040DAF9-6CE4-4BE8-839D-F4538A4301CF}' : 'SoftEdgeFormat',
	'{12DCDC9A-5418-48A3-BBE6-EB57BAE275E8}' : 'Reviewers',
	'{F01943FF-1985-445E-8602-8FB8F39CCA75}' : 'ReflectionFormat',
	'{C1A870A0-850E-4D38-98A7-741CB8C3BCA4}' : 'Points',
	'{9F1DF642-3CCE-4D83-A770-D2634A05D278}' : 'DropLines',
	'{396F9073-F9FD-11D3-8EA0-0050049A1A01}' : 'CanvasShapes',
	'{67A7EEC5-285D-4024-B071-BD6B33B88547}' : 'OMathRad',
	'{DB77D541-85C3-42E8-8649-AFBD7CF87866}' : 'OMathPhantom',
	'{E598E358-2852-42D4-8775-160BD91B7244}' : 'UndoRecord',
	'{15EBE471-0182-4CCE-98D0-B6614D1C32A1}' : 'SmartTagRecognizer',
	'{1B426348-607D-433C-9216-C5D2BF0EF31F}' : 'OMathMatRows',
	'{FE0971F0-5E60-4985-BCDA-95CB0B8E0308}' : 'XMLSchemaReference',
	'{6FFA84BB-A350-4442-BB53-A43653459A84}' : 'Chart',
	'{C1AD33E4-F088-40A9-9D2F-D94017D115C4}' : 'ChartTitle',
	'{47CEF4AE-DC32-4220-8AA5-19CCC0E6633A}' : 'Reviewer',
	'{FC9086C6-0287-4997-B2E1-816C334A22F8}' : 'OMathLimUpp',
	'{40810760-068A-4486-BEC9-8EA58C7029F5}' : 'Series',
	'{FD0A74E8-C719-49F6-BA1B-F6D9839D1AB9}' : 'ProtectedViewWindows',
	'{AE6D45E5-981E-4547-8752-674BB55420A5}' : 'Corners',
	'{799A6814-EA41-11D3-87CC-00105AA31A34}' : 'MappedDataFields',
	'{CEBD4184-4E6D-4FC6-A42D-2142B1B76AF5}' : 'OMathNary',
	'{65E515D5-F50B-4951-8F38-FA6AC8707387}' : 'OMathBreak',
	'{18CD5EC8-8B7B-42C8-992A-2A407468642C}' : 'OMathAutoCorrectEntries',
	'{1498F56D-ED33-41F9-B37B-EF30E50B08AC}' : 'ConditionalStyle',
	'{B923FDE0-F08C-11D3-91B0-00105A0A19FD}' : 'CustomProperty',
	'{B923FDE1-F08C-11D3-91B0-00105A0A19FD}' : 'CustomProperties',
	'{8A342FA0-5831-4B5E-82E1-003D0A0C635D}' : 'Point',
	'{98DFBD12-96CB-4F07-90EA-749FF1D6B89D}' : 'OMathScrSub',
	'{E2E8A400-0615-427D-ADCC-CAD39FFEBD42}' : 'Lines',
	'{B184502B-587A-4C6A-8DC4-ECE4354883C6}' : 'Interior',
	'{CDE12CD8-767B-4757-8A31-13029A086305}' : 'SmartTagActions',
	'{09760240-0B89-49F7-A79D-479F24723F56}' : 'XMLNode',
	'{00020906-0000-0000-C000-000000000046}' : 'Document',
	'{C2B83A65-B061-4469-83B6-8877437CB8A0}' : 'Conflicts',
	'{817F99FA-CCC4-4971-8E9D-1238F735AAFF}' : 'BuildingBlockType',
	'{ECFBDB5E-ACD2-4530-AD79-4560B7FF055C}' : 'Category',
	'{B6511068-70BF-4751-A741-55C1D41AD96F}' : 'LegendEntries',
	'{DCE9F2C4-4C02-43BA-840E-B4276550EF79}' : 'DataTable',
	'{D0A95726-678A-4B9D-8103-1E2B86735AE7}' : 'OMathScrSup',
	'{DFF99AC2-CD2A-43AD-91B1-A2BE40BC7146}' : 'CoAuthLocks',
	'{E3124493-7D6A-410F-9A48-CC822C033CEC}' : 'XSLTransform',
	'{2503B6EE-0889-44DF-B920-6D6F9659DEA3}' : 'OMathBorderBox',
	'{FA02A26B-6550-45C5-B6F0-80E757CD3482}' : 'Sources',
	'{B140A023-4850-4DA6-BC5F-CC459C4507BC}' : 'XMLNamespace',
	'{99755F80-FE96-4F7D-B636-B8E800E54F44}' : 'CoAuthLock',
	'{54B7061A-D56C-40E5-B85B-58146446C782}' : 'Trendlines',
	'{DC489AD4-23C4-4F4B-990F-45A51C7C0C4F}' : 'OMathScrSubSup',
	'{00020910-0000-0000-C000-000000000046}' : 'Dialogs',
	'{00020911-0000-0000-C000-000000000046}' : 'TableOfAuthorities',
	'{00020912-0000-0000-C000-000000000046}' : 'TablesOfAuthorities',
	'{00020913-0000-0000-C000-000000000046}' : 'TableOfContents',
	'{00020914-0000-0000-C000-000000000046}' : 'TablesOfContents',
	'{00020915-0000-0000-C000-000000000046}' : 'CustomLabel',
	'{00020916-0000-0000-C000-000000000046}' : 'CustomLabels',
	'{00020917-0000-0000-C000-000000000046}' : 'MailingLabel',
	'{00020918-0000-0000-C000-000000000046}' : 'Envelope',
	'{00020919-0000-0000-C000-000000000046}' : 'MailMergeDataField',
	'{0002091A-0000-0000-C000-000000000046}' : 'MailMergeDataFields',
	'{0002091B-0000-0000-C000-000000000046}' : 'MailMergeFieldName',
	'{0002091C-0000-0000-C000-000000000046}' : 'MailMergeFieldNames',
	'{1F998A61-71C6-44C2-A0F2-1D66169B47CB}' : 'OMathEqArray',
	'{0002091E-0000-0000-C000-000000000046}' : 'MailMergeField',
	'{0002091F-0000-0000-C000-000000000046}' : 'MailMergeFields',
	'{00020920-0000-0000-C000-000000000046}' : 'MailMerge',
	'{00020921-0000-0000-C000-000000000046}' : 'TableOfFigures',
	'{44FEE887-6600-41AB-95A5-DE33C605116C}' : 'OMathRecognizedFunctions',
	'{00020923-0000-0000-C000-000000000046}' : 'ListEntry',
	'{00020924-0000-0000-C000-000000000046}' : 'ListEntries',
	'{00020925-0000-0000-C000-000000000046}' : 'DropDown',
	'{00020926-0000-0000-C000-000000000046}' : 'CheckBox',
	'{00020927-0000-0000-C000-000000000046}' : 'TextInput',
	'{00020928-0000-0000-C000-000000000046}' : 'FormField',
	'{00020929-0000-0000-C000-000000000046}' : 'FormFields',
	'{0002092A-0000-0000-C000-000000000046}' : 'Frame',
	'{0002092B-0000-0000-C000-000000000046}' : 'Frames',
	'{0002092C-0000-0000-C000-000000000046}' : 'Style',
	'{0002092D-0000-0000-C000-000000000046}' : 'Styles',
	'{0002092E-0000-0000-C000-000000000046}' : 'Browser',
	'{0002092F-0000-0000-C000-000000000046}' : 'Field',
	'{00020930-0000-0000-C000-000000000046}' : 'Fields',
	'{00020931-0000-0000-C000-000000000046}' : 'LinkFormat',
	'{00020933-0000-0000-C000-000000000046}' : 'OLEFormat',
	'{00020935-0000-0000-C000-000000000046}' : 'System',
	'{00020936-0000-0000-C000-000000000046}' : 'AutoTextEntry',
	'{00020937-0000-0000-C000-000000000046}' : 'AutoTextEntries',
	'{00020939-0000-0000-C000-000000000046}' : 'TextRetrievalMode',
	'{0002093A-0000-0000-C000-000000000046}' : 'Shading',
	'{0002093B-0000-0000-C000-000000000046}' : 'Border',
	'{0002093C-0000-0000-C000-000000000046}' : 'Borders',
	'{0002093D-0000-0000-C000-000000000046}' : 'Comment',
	'{0002093E-0000-0000-C000-000000000046}' : 'Endnote',
	'{0002093F-0000-0000-C000-000000000046}' : 'Footnote',
	'{00020940-0000-0000-C000-000000000046}' : 'Comments',
	'{00020941-0000-0000-C000-000000000046}' : 'Endnotes',
	'{00020942-0000-0000-C000-000000000046}' : 'Footnotes',
	'{00020943-0000-0000-C000-000000000046}' : 'TwoInitialCapsException',
	'{00020944-0000-0000-C000-000000000046}' : 'TwoInitialCapsExceptions',
	'{00020945-0000-0000-C000-000000000046}' : 'FirstLetterException',
	'{00020946-0000-0000-C000-000000000046}' : 'FirstLetterExceptions',
	'{00020947-0000-0000-C000-000000000046}' : 'AutoCorrectEntry',
	'{00020948-0000-0000-C000-000000000046}' : 'AutoCorrectEntries',
	'{00020949-0000-0000-C000-000000000046}' : 'AutoCorrect',
	'{0002094A-0000-0000-C000-000000000046}' : 'Cells',
	'{0002094B-0000-0000-C000-000000000046}' : 'Columns',
	'{0002094C-0000-0000-C000-000000000046}' : 'Rows',
	'{0002094D-0000-0000-C000-000000000046}' : 'Tables',
	'{0002094E-0000-0000-C000-000000000046}' : 'Cell',
	'{0002094F-0000-0000-C000-000000000046}' : 'Column',
	'{00020950-0000-0000-C000-000000000046}' : 'Row',
	'{00020951-0000-0000-C000-000000000046}' : 'Table',
	'{00020952-0000-0000-C000-000000000046}' : '_Font',
	'{00020953-0000-0000-C000-000000000046}' : '_ParagraphFormat',
	'{00020954-0000-0000-C000-000000000046}' : 'TabStop',
	'{00020955-0000-0000-C000-000000000046}' : 'TabStops',
	'{00020956-0000-0000-C000-000000000046}' : 'DropCap',
	'{00020957-0000-0000-C000-000000000046}' : 'Paragraph',
	'{00020958-0000-0000-C000-000000000046}' : 'Paragraphs',
	'{00020959-0000-0000-C000-000000000046}' : 'Section',
	'{0002095A-0000-0000-C000-000000000046}' : 'Sections',
	'{0002095B-0000-0000-C000-000000000046}' : 'Sentences',
	'{0002095C-0000-0000-C000-000000000046}' : 'Words',
	'{0002095D-0000-0000-C000-000000000046}' : 'Characters',
	'{0002095E-0000-0000-C000-000000000046}' : 'Range',
	'{0002095F-0000-0000-C000-000000000046}' : 'Panes',
	'{00020960-0000-0000-C000-000000000046}' : 'Pane',
	'{00020961-0000-0000-C000-000000000046}' : 'Windows',
	'{00020962-0000-0000-C000-000000000046}' : 'Window',
	'{00020963-0000-0000-C000-000000000046}' : 'RecentFiles',
	'{00020964-0000-0000-C000-000000000046}' : 'RecentFile',
	'{00020965-0000-0000-C000-000000000046}' : 'Variables',
	'{00020966-0000-0000-C000-000000000046}' : 'Variable',
	'{00020967-0000-0000-C000-000000000046}' : 'Bookmarks',
	'{00020968-0000-0000-C000-000000000046}' : 'Bookmark',
	'{00020969-0000-0000-C000-000000000046}' : 'RoutingSlip',
	'{0002096A-0000-0000-C000-000000000046}' : 'Template',
	'{0002096B-0000-0000-C000-000000000046}' : '_Document',
	'{0002096C-0000-0000-C000-000000000046}' : 'Documents',
	'{0002096D-0000-0000-C000-000000000046}' : 'Language',
	'{0002096E-0000-0000-C000-000000000046}' : 'Languages',
	'{0002096F-0000-0000-C000-000000000046}' : 'FontNames',
	'{00020970-0000-0000-C000-000000000046}' : '_Application',
	'{00020971-0000-0000-C000-000000000046}' : 'PageSetup',
	'{00020972-0000-0000-C000-000000000046}' : 'LineNumbering',
	'{00020973-0000-0000-C000-000000000046}' : 'TextColumns',
	'{00020974-0000-0000-C000-000000000046}' : 'TextColumn',
	'{00020975-0000-0000-C000-000000000046}' : 'Selection',
	'{00020976-0000-0000-C000-000000000046}' : 'TablesOfAuthoritiesCategories',
	'{00020977-0000-0000-C000-000000000046}' : 'TableOfAuthoritiesCategory',
	'{00020978-0000-0000-C000-000000000046}' : 'CaptionLabels',
	'{00020979-0000-0000-C000-000000000046}' : 'CaptionLabel',
	'{0002097A-0000-0000-C000-000000000046}' : 'AutoCaptions',
	'{0002097B-0000-0000-C000-000000000046}' : 'AutoCaption',
	'{0002097C-0000-0000-C000-000000000046}' : 'Indexes',
	'{0002097D-0000-0000-C000-000000000046}' : 'Index',
	'{0002097E-0000-0000-C000-000000000046}' : 'AddIn',
	'{0002097F-0000-0000-C000-000000000046}' : 'AddIns',
	'{00020980-0000-0000-C000-000000000046}' : 'Revisions',
	'{00020981-0000-0000-C000-000000000046}' : 'Revision',
	'{00020982-0000-0000-C000-000000000046}' : 'Task',
	'{00020983-0000-0000-C000-000000000046}' : 'Tasks',
	'{00020984-0000-0000-C000-000000000046}' : 'HeadersFooters',
	'{00020985-0000-0000-C000-000000000046}' : 'HeaderFooter',
	'{00020986-0000-0000-C000-000000000046}' : 'PageNumbers',
	'{00020987-0000-0000-C000-000000000046}' : 'PageNumber',
	'{00020988-0000-0000-C000-000000000046}' : 'Subdocuments',
	'{00020989-0000-0000-C000-000000000046}' : 'Subdocument',
	'{0002098A-0000-0000-C000-000000000046}' : 'HeadingStyles',
	'{0002098B-0000-0000-C000-000000000046}' : 'HeadingStyle',
	'{0002098C-0000-0000-C000-000000000046}' : 'StoryRanges',
	'{0002098D-0000-0000-C000-000000000046}' : 'ListLevel',
	'{0002098E-0000-0000-C000-000000000046}' : 'ListLevels',
	'{0002098F-0000-0000-C000-000000000046}' : 'ListTemplate',
	'{00020990-0000-0000-C000-000000000046}' : 'ListTemplates',
	'{00020991-0000-0000-C000-000000000046}' : 'ListParagraphs',
	'{00020992-0000-0000-C000-000000000046}' : 'List',
	'{00020993-0000-0000-C000-000000000046}' : 'Lists',
	'{00020994-0000-0000-C000-000000000046}' : 'ListGallery',
	'{00020995-0000-0000-C000-000000000046}' : 'ListGalleries',
	'{00020996-0000-0000-C000-000000000046}' : 'KeyBindings',
	'{00020997-0000-0000-C000-000000000046}' : 'KeysBoundTo',
	'{00020998-0000-0000-C000-000000000046}' : 'KeyBinding',
	'{00020999-0000-0000-C000-000000000046}' : 'FileConverter',
	'{0002099A-0000-0000-C000-000000000046}' : 'FileConverters',
	'{0002099B-0000-0000-C000-000000000046}' : 'SynonymInfo',
	'{0002099C-0000-0000-C000-000000000046}' : 'Hyperlinks',
	'{0002099D-0000-0000-C000-000000000046}' : 'Hyperlink',
	'{0002099F-0000-0000-C000-000000000046}' : 'Shapes',
	'{000209A0-0000-0000-C000-000000000046}' : 'Shape',
	'{000209A1-0000-0000-C000-000000000046}' : '_LetterContent',
	'{000209A2-0000-0000-C000-000000000046}' : 'Templates',
	'{000209A4-0000-0000-C000-000000000046}' : '_OLEControl',
	'{000209A5-0000-0000-C000-000000000046}' : 'View',
	'{000209A6-0000-0000-C000-000000000046}' : 'Zoom',
	'{000209A7-0000-0000-C000-000000000046}' : 'Zooms',
	'{000209A8-0000-0000-C000-000000000046}' : 'InlineShape',
	'{000209A9-0000-0000-C000-000000000046}' : 'InlineShapes',
	'{000209AA-0000-0000-C000-000000000046}' : 'SpellingSuggestions',
	'{000209AB-0000-0000-C000-000000000046}' : 'SpellingSuggestion',
	'{000209AC-0000-0000-C000-000000000046}' : 'Dictionaries',
	'{000209AD-0000-0000-C000-000000000046}' : 'Dictionary',
	'{000209AE-0000-0000-C000-000000000046}' : 'ReadabilityStatistics',
	'{000209AF-0000-0000-C000-000000000046}' : 'ReadabilityStatistic',
	'{000209B0-0000-0000-C000-000000000046}' : 'Find',
	'{000209B1-0000-0000-C000-000000000046}' : 'Replacement',
	'{000209B2-0000-0000-C000-000000000046}' : 'TextFrame',
	'{000209B3-0000-0000-C000-000000000046}' : 'Versions',
	'{000209B4-0000-0000-C000-000000000046}' : 'Version',
	'{000209B5-0000-0000-C000-000000000046}' : 'ShapeRange',
	'{000209B6-0000-0000-C000-000000000046}' : 'GroupShapes',
	'{000209B7-0000-0000-C000-000000000046}' : 'Options',
	'{000209B8-0000-0000-C000-000000000046}' : 'Dialog',
	'{000209B9-0000-0000-C000-000000000046}' : '_Global',
	'{000209BA-0000-0000-C000-000000000046}' : 'MailMessage',
	'{000209BB-0000-0000-C000-000000000046}' : 'ProofreadingErrors',
	'{000209BD-0000-0000-C000-000000000046}' : 'Mailer',
	'{000209C0-0000-0000-C000-000000000046}' : 'ListFormat',
	'{194F8476-B79D-4572-A609-294207DE77C1}' : 'ErrorBars',
	'{000209C3-0000-0000-C000-000000000046}' : 'WrapFormat',
	'{000209C4-0000-0000-C000-000000000046}' : 'Adjustments',
	'{000209C5-0000-0000-C000-000000000046}' : 'CalloutFormat',
	'{000209C6-0000-0000-C000-000000000046}' : 'ColorFormat',
	'{000209C7-0000-0000-C000-000000000046}' : 'ConnectorFormat',
	'{000209C8-0000-0000-C000-000000000046}' : 'FillFormat',
	'{000209C9-0000-0000-C000-000000000046}' : 'FreeformBuilder',
	'{000209CA-0000-0000-C000-000000000046}' : 'LineFormat',
	'{000209CB-0000-0000-C000-000000000046}' : 'PictureFormat',
	'{000209CC-0000-0000-C000-000000000046}' : 'ShadowFormat',
	'{000209CD-0000-0000-C000-000000000046}' : 'ShapeNode',
	'{000209CE-0000-0000-C000-000000000046}' : 'ShapeNodes',
	'{000209CF-0000-0000-C000-000000000046}' : 'TextEffectFormat',
	'{000209D0-0000-0000-C000-000000000046}' : 'ThreeDFormat',
	'{000209D1-0000-0000-C000-000000000046}' : 'HangulAndAlphabetExceptions',
	'{000209D2-0000-0000-C000-000000000046}' : 'HangulAndAlphabetException',
	'{EE95AFE3-3026-4172-B078-0E79DAB5CC3D}' : 'ContentControl',
	'{000209D7-0000-0000-C000-000000000046}' : 'EmailAuthor',
	'{000209DB-0000-0000-C000-000000000046}' : 'EmailOptions',
	'{000209DC-0000-0000-C000-000000000046}' : 'EmailSignature',
	'{000209DD-0000-0000-C000-000000000046}' : 'Email',
	'{000209DE-0000-0000-C000-000000000046}' : 'HorizontalLineFormat',
	'{000209DF-0000-0000-C000-000000000046}' : 'OtherCorrectionsExceptions',
	'{000209E0-0000-0000-C000-000000000046}' : 'HangulHanjaConversionDictionaries',
	'{000209E1-0000-0000-C000-000000000046}' : 'OtherCorrectionsException',
	'{000209E2-0000-0000-C000-000000000046}' : 'Frameset',
	'{000209E3-0000-0000-C000-000000000046}' : 'DefaultWebOptions',
	'{000209E4-0000-0000-C000-000000000046}' : 'WebOptions',
	'{000209E5-0000-0000-C000-000000000046}' : 'EmailSignatureEntries',
	'{000209E6-0000-0000-C000-000000000046}' : 'EmailSignatureEntry',
	'{000209E7-0000-0000-C000-000000000046}' : 'HTMLDivision',
	'{000209E8-0000-0000-C000-000000000046}' : 'HTMLDivisions',
	'{000209E9-0000-0000-C000-000000000046}' : 'DiagramNode',
	'{000209EA-0000-0000-C000-000000000046}' : 'DiagramNodeChildren',
	'{000209EB-0000-0000-C000-000000000046}' : 'DiagramNodes',
	'{000209EC-0000-0000-C000-000000000046}' : 'Diagram',
	'{000209ED-0000-0000-C000-000000000046}' : 'SmartTag',
	'{000209EE-0000-0000-C000-000000000046}' : 'SmartTags',
	'{000209EF-0000-0000-C000-000000000046}' : 'StyleSheet',
	'{000209F0-0000-0000-C000-000000000046}' : 'Global',
	'{000209F1-0000-0000-C000-000000000046}' : 'LetterContent',
	'{000209F2-0000-0000-C000-000000000046}' : 'OLEControl',
	'{000209F3-0000-0000-C000-000000000046}' : 'OCXEvents',
	'{000209F4-0000-0000-C000-000000000046}' : 'ParagraphFormat',
	'{000209F5-0000-0000-C000-000000000046}' : 'Font',
	'{000209F6-0000-0000-C000-000000000046}' : 'DocumentEvents',
	'{000209F7-0000-0000-C000-000000000046}' : 'ApplicationEvents',
	'{000209F7-0001-0000-C000-000000000046}' : 'IApplicationEvents',
	'{000209FE-0000-0000-C000-000000000046}' : 'ApplicationEvents2',
	'{000209FE-0001-0000-C000-000000000046}' : 'IApplicationEvents2',
	'{00020A00-0000-0000-C000-000000000046}' : 'ApplicationEvents3',
	'{00020A01-0000-0000-C000-000000000046}' : 'ApplicationEvents4',
	'{00020A02-0000-0000-C000-000000000046}' : 'DocumentEvents2',
	'{C94688A6-A2A7-4133-A26D-726CD569D5F3}' : 'OMathDelim',
	'{7E64D2BE-2818-48CB-8F8A-CC7B61D9E860}' : 'Floor',
	'{86905AC9-33F3-4A88-96C8-B289B0390BCA}' : 'UpBars',
	'{E4442A83-F623-459C-8E95-8BFB44DCF23A}' : 'OMath',
	'{842C37FE-C76F-4B2B-9B60-C408CB5E838E}' : 'OMathBox',
	'{AB0D33A3-C9EA-485B-9443-4C1BB3656CEA}' : 'ChartBorder',
	'{07B7CC7E-E66C-11D3-9454-00105AA31A08}' : 'StyleSheets',
	'{4A6AE865-199D-4EA3-9F6B-125BD9C40EDF}' : 'Source',
	'{354AB591-A217-48B4-99E4-14F58F15667D}' : 'Axes',
	'{6E47678B-A879-4E56-8698-3B7CF169FAD4}' : 'Categories',
	'{B9F1A4E2-0D0A-43B7-8495-139E7ACBD840}' : 'TaskPane',
	'{E6AAEC05-E543-4085-BA92-9BF7D2474F51}' : 'Research',
	'{BFD3FC23-F763-4FF8-826E-1AFBF598A4E7}' : 'BuildingBlock',
	'{EFC71F9C-7F42-4CD4-A7A7-970D7A48CD27}' : 'OMathMatCols',
	'{F743EDD0-9B97-4B09-89CC-77BE19B51481}' : 'ProtectedViewWindow',
	'{0D951ADF-10A6-4C9B-BCD9-0FB8CBAD9A87}' : 'OMathFunc',
	'{1FD94DF1-3569-4465-94FF-E8B22D28EEB0}' : 'DataLabel',
	'{39709229-56A0-4E29-9112-B31DD067EBFD}' : 'BuildingBlockEntries',
	'{0002091D-0000-0000-C000-000000000046}' : 'MailMergeDataSource',
	'{9E6B5EC5-E8E4-40AF-9540-6203F71E2823}' : 'CoAuthUpdate',
	'{16BE9309-D708-4322-BB1A-B056F58D17EA}' : 'Breaks',
	'{352840A9-AF7D-4CA4-87FC-21C68FDAB3E4}' : 'Page',
	'{FC9090AF-0DDB-4EC1-86E8-8751F2199F2C}' : 'Gridlines',
	'{B7564E97-0519-4C68-B400-3803E7C63242}' : 'TableStyle',
	'{00020922-0000-0000-C000-000000000046}' : 'TablesOfFigures',
	'{30225CFC-5A71-4FE6-B527-90A52C54AE77}' : 'CoAuthUpdates',
	'{5C04BD93-2F3F-4668-918D-9738EC901039}' : 'OMathRecognizedFunction',
	'{D8252C5E-EB9F-4D74-AA72-C178B128FAC4}' : 'DataLabels',
	'{3E061A7E-67AD-4EAA-BC1E-55057D5E596F}' : 'OMathMat',
	'{0C6FA8CA-E65F-4FC7-AB8F-20729EECBB14}' : 'ContentControlListEntry',
	'{DB8E3072-E106-4453-8E7C-53056F1D30DC}' : 'SmartTagTypes',
	'{804CD967-F83B-432D-9446-C61A45CFEFF0}' : 'ContentControls',
	'{873E774B-926A-4CB1-878D-635A45187595}' : 'OMaths',
	'{BF043168-F4DE-4E7C-B206-741A8B3EF71A}' : 'EndnoteOptions',
	'{F8DDB497-CA6C-4711-9BA4-2718FA3BB6FE}' : 'ChartGroups',
	'{F08B45F1-8F23-4156-9D63-1820C0ED229A}' : 'OMathBar',
	'{7A1BCE11-5783-4C7D-BD02-F3D84AB40E7F}' : 'HiLoLines',
	'{3834F60F-EE8C-455D-A441-D766675D6D3B}' : 'Bibliography',
	'{D36C1F42-7044-4B9E-9CA3-85919454DB04}' : 'XMLNodes',
	'{656BBED7-E82D-4B0A-8F97-EC742BA11FFA}' : 'XMLNamespaces',
	'{F258DE05-C41B-4C33-A778-F0D3F98CEEB3}' : 'OMathAcc',
	'{5DAA8BB6-054E-48F6-BEAC-EAAD02BE0CC7}' : 'OMathMatRow',
	'{356B06EC-4908-42A4-81FC-4B5A51F3483B}' : 'XMLSchemaReferences',
	'{50209974-BA32-4A03-8FA6-BAC56CC056FD}' : 'OMathFrac',
	'{ADD4EDF3-2F33-4734-9CE6-D476097C5ADA}' : 'Rectangle',
	'{4A304B59-31FF-42DD-B436-7FC9C5DB7559}' : 'ChartData',
	'{E6AAEC05-E543-4085-BA92-9BF7D2474F5C}' : 'TaskPanes',
	'{8245795B-9AED-4943-A16D-E586ED8180D1}' : 'OMathArgs',
	'{65DF9F31-B1E3-4651-87E8-51D55F302161}' : 'CoAuthoring',
	'{C4A02049-024C-4273-8934-E48CC21479A9}' : 'LegendEntry',
	'{84A6A663-AEF4-4FCD-83FD-9BB707F157CA}' : 'DownBars',
	'{91C46192-3124-4346-A815-10B8873F5A06}' : 'Trendline',
	'{8B0E45DB-3A7B-42EE-9D17-A92AF69B79C1}' : 'AxisTitle',
	'{6215E4B1-545A-406E-9824-0A5B5AC8AD21}' : 'Conflict',
	'{497142A4-16FD-42C6-BC58-15D89345FC21}' : 'OMathFunctions',
	'{F152D349-7D20-4C01-A42B-2D6DE4F3891C}' : 'ChartFillFormat',
	'{F1F37152-1DB1-4901-AD9A-C740F99464B4}' : 'OMathFunction',
}
VTablesToClassMap = {}
VTablesToPackageMap = {
	'{935D59F5-E365-4F92-B7F5-1C499A63ECA8}' : 'TickLabels',
	'{AFAF0C0E-8603-40F6-8FD1-42726CAC21E3}' : 'OMathScrPre',
	'{CDB0FF41-E862-47BB-AE77-3FA7B1AE3189}' : 'ChartFont',
	'{AE6CE2F5-B9D3-407D-85A8-0F10C63289A4}' : 'Line',
	'{74DE9576-8E99-4E28-912B-CB30747C60CE}' : 'OMathLimLow',
	'{7D0F7985-68D9-4D93-91CB-8109280E76CC}' : 'Rectangles',
	'{54F46DC4-F6A6-48CC-BD66-46C1DDEADD22}' : 'ContentControlListEntries',
	'{0C1FABE7-F737-406F-9CA3-B07661F9D1A2}' : 'XMLMapping',
	'{CAE36175-3818-4C60-BCBF-0645D51EB33B}' : 'OMathMatCol',
	'{79635BF1-BD1D-4B3F-A520-C1106F1AAAD8}' : 'Break',
	'{B5828B50-0E3D-448A-962D-A40702A5868D}' : 'BuildingBlockTypes',
	'{256B6ABA-6A38-4D39-971C-91FDA9922814}' : 'CoAuthors',
	'{A2E94180-7564-4D97-806B-BBC0D0A1350C}' : 'Walls',
	'{B3A1E8C6-E1CE-4A46-8D12-E017157B03D7}' : 'Legend',
	'{C75AD98A-74E9-49FE-8BF1-544839CC08A5}' : 'ChartArea',
	'{6F9D1F68-06F7-49EF-8902-185E54EB5E87}' : 'OMathAutoCorrect',
	'{56AFD330-440C-4F4C-A39C-ED306D084D5F}' : 'PlotArea',
	'{FF06FEF2-DA89-41C0-A0A8-5CD434E210AD}' : 'ChartCharacters',
	'{86488FB4-9633-4C93-8057-FC1FA7A847AE}' : 'ChartGroup',
	'{DD8F80B8-9B80-4E89-9BEC-F12DF35E43B3}' : 'ChartColorFormat',
	'{DFB6AA6C-1068-420F-969D-01280FCC1630}' : 'SmartTagAction',
	'{D8779F01-4869-4403-B334-D60C5F9C9175}' : 'OMathAutoCorrectEntry',
	'{7EBC66BD-F788-42C3-91F4-E8C841A69005}' : 'Axis',
	'{C774F5EA-A539-4284-A1BE-30AEC052D899}' : 'XSLTransforms',
	'{16BE9309-D708-4322-BB1A-B056F58D17EA}' : 'Breaks',
	'{8FEB78F7-35C6-4871-918C-193C3CDD886D}' : 'SeriesCollection',
	'{91807402-6C6F-47CD-B8FA-C42FEE8EE924}' : 'Pages',
	'{AED7E08C-14F0-4F33-921D-4C5353137BF6}' : 'Editors',
	'{5D311669-EA51-11D3-87CC-00105AA31A34}' : 'MappedDataField',
	'{F1B14F40-5C32-4C8C-B5B2-DE537BB6B89D}' : 'GlowFormat',
	'{C6D50987-25D7-408A-BFF2-90BF86A24E93}' : 'BuildingBlocks',
	'{E59544D5-C299-46A0-84C1-C51AB38F9759}' : 'CoAuthor',
	'{BEA85A24-D7DA-4F3D-B58C-ED90FB01D615}' : 'FootnoteOptions',
	'{5D7F6C15-36CE-44CC-9692-5A1F8B8C906D}' : 'SeriesLines',
	'{DF076FDE-8781-4051-A5BC-99F6B7DC04D4}' : 'LegendKey',
	'{A87E00E9-3AC3-4B53-ABE3-7379653D0E82}' : 'XMLChildNodeSuggestion',
	'{36162C62-B59A-4278-AF3D-F2AC1EB999D9}' : 'LeaderLines',
	'{DD947D72-F33C-4198-9BDF-F86181D05E41}' : 'Editor',
	'{C04865A3-9F8A-486C-BB58-B4C3E6563136}' : 'DisplayUnitLabel',
	'{02B17CB4-7D55-4B34-B38B-10381433441F}' : 'OMathGroupChar',
	'{B66D3C1A-4541-4961-B35B-A353C03F6A99}' : 'ChartFormat',
	'{F2B60A10-DED5-46FB-A914-3C6F4EBB6451}' : 'SmartTagRecognizers',
	'{E2E0F3A7-204C-40C5-BAA5-290F374FDF5A}' : 'OMathBreaks',
	'{DE63B5AC-CA4F-46FE-9184-A5719AB9ED5B}' : 'XMLChildNodeSuggestions',
	'{5E9A888C-E5DC-4DCB-8308-3C91FB61E6F4}' : 'SmartTagType',
	'{D040DAF9-6CE4-4BE8-839D-F4538A4301CF}' : 'SoftEdgeFormat',
	'{12DCDC9A-5418-48A3-BBE6-EB57BAE275E8}' : 'Reviewers',
	'{F01943FF-1985-445E-8602-8FB8F39CCA75}' : 'ReflectionFormat',
	'{C1A870A0-850E-4D38-98A7-741CB8C3BCA4}' : 'Points',
	'{9F1DF642-3CCE-4D83-A770-D2634A05D278}' : 'DropLines',
	'{396F9073-F9FD-11D3-8EA0-0050049A1A01}' : 'CanvasShapes',
	'{67A7EEC5-285D-4024-B071-BD6B33B88547}' : 'OMathRad',
	'{DB77D541-85C3-42E8-8649-AFBD7CF87866}' : 'OMathPhantom',
	'{E598E358-2852-42D4-8775-160BD91B7244}' : 'UndoRecord',
	'{15EBE471-0182-4CCE-98D0-B6614D1C32A1}' : 'SmartTagRecognizer',
	'{1B426348-607D-433C-9216-C5D2BF0EF31F}' : 'OMathMatRows',
	'{FE0971F0-5E60-4985-BCDA-95CB0B8E0308}' : 'XMLSchemaReference',
	'{6FFA84BB-A350-4442-BB53-A43653459A84}' : 'Chart',
	'{C1AD33E4-F088-40A9-9D2F-D94017D115C4}' : 'ChartTitle',
	'{47CEF4AE-DC32-4220-8AA5-19CCC0E6633A}' : 'Reviewer',
	'{FC9086C6-0287-4997-B2E1-816C334A22F8}' : 'OMathLimUpp',
	'{40810760-068A-4486-BEC9-8EA58C7029F5}' : 'Series',
	'{FD0A74E8-C719-49F6-BA1B-F6D9839D1AB9}' : 'ProtectedViewWindows',
	'{AE6D45E5-981E-4547-8752-674BB55420A5}' : 'Corners',
	'{799A6814-EA41-11D3-87CC-00105AA31A34}' : 'MappedDataFields',
	'{CEBD4184-4E6D-4FC6-A42D-2142B1B76AF5}' : 'OMathNary',
	'{65E515D5-F50B-4951-8F38-FA6AC8707387}' : 'OMathBreak',
	'{18CD5EC8-8B7B-42C8-992A-2A407468642C}' : 'OMathAutoCorrectEntries',
	'{1498F56D-ED33-41F9-B37B-EF30E50B08AC}' : 'ConditionalStyle',
	'{B923FDE0-F08C-11D3-91B0-00105A0A19FD}' : 'CustomProperty',
	'{B923FDE1-F08C-11D3-91B0-00105A0A19FD}' : 'CustomProperties',
	'{8A342FA0-5831-4B5E-82E1-003D0A0C635D}' : 'Point',
	'{98DFBD12-96CB-4F07-90EA-749FF1D6B89D}' : 'OMathScrSub',
	'{E2E8A400-0615-427D-ADCC-CAD39FFEBD42}' : 'Lines',
	'{B184502B-587A-4C6A-8DC4-ECE4354883C6}' : 'Interior',
	'{CDE12CD8-767B-4757-8A31-13029A086305}' : 'SmartTagActions',
	'{09760240-0B89-49F7-A79D-479F24723F56}' : 'XMLNode',
	'{C2B83A65-B061-4469-83B6-8877437CB8A0}' : 'Conflicts',
	'{817F99FA-CCC4-4971-8E9D-1238F735AAFF}' : 'BuildingBlockType',
	'{ECFBDB5E-ACD2-4530-AD79-4560B7FF055C}' : 'Category',
	'{54B7061A-D56C-40E5-B85B-58146446C782}' : 'Trendlines',
	'{DCE9F2C4-4C02-43BA-840E-B4276550EF79}' : 'DataTable',
	'{D0A95726-678A-4B9D-8103-1E2B86735AE7}' : 'OMathScrSup',
	'{DFF99AC2-CD2A-43AD-91B1-A2BE40BC7146}' : 'CoAuthLocks',
	'{E3124493-7D6A-410F-9A48-CC822C033CEC}' : 'XSLTransform',
	'{2503B6EE-0889-44DF-B920-6D6F9659DEA3}' : 'OMathBorderBox',
	'{FA02A26B-6550-45C5-B6F0-80E757CD3482}' : 'Sources',
	'{B140A023-4850-4DA6-BC5F-CC459C4507BC}' : 'XMLNamespace',
	'{99755F80-FE96-4F7D-B636-B8E800E54F44}' : 'CoAuthLock',
	'{B6511068-70BF-4751-A741-55C1D41AD96F}' : 'LegendEntries',
	'{DC489AD4-23C4-4F4B-990F-45A51C7C0C4F}' : 'OMathScrSubSup',
	'{354AB591-A217-48B4-99E4-14F58F15667D}' : 'Axes',
	'{00020911-0000-0000-C000-000000000046}' : 'TableOfAuthorities',
	'{00020912-0000-0000-C000-000000000046}' : 'TablesOfAuthorities',
	'{00020913-0000-0000-C000-000000000046}' : 'TableOfContents',
	'{00020914-0000-0000-C000-000000000046}' : 'TablesOfContents',
	'{00020915-0000-0000-C000-000000000046}' : 'CustomLabel',
	'{00020916-0000-0000-C000-000000000046}' : 'CustomLabels',
	'{00020917-0000-0000-C000-000000000046}' : 'MailingLabel',
	'{00020918-0000-0000-C000-000000000046}' : 'Envelope',
	'{00020919-0000-0000-C000-000000000046}' : 'MailMergeDataField',
	'{0002091A-0000-0000-C000-000000000046}' : 'MailMergeDataFields',
	'{0002091B-0000-0000-C000-000000000046}' : 'MailMergeFieldName',
	'{0002091C-0000-0000-C000-000000000046}' : 'MailMergeFieldNames',
	'{1F998A61-71C6-44C2-A0F2-1D66169B47CB}' : 'OMathEqArray',
	'{0002091E-0000-0000-C000-000000000046}' : 'MailMergeField',
	'{0002091F-0000-0000-C000-000000000046}' : 'MailMergeFields',
	'{00020920-0000-0000-C000-000000000046}' : 'MailMerge',
	'{00020921-0000-0000-C000-000000000046}' : 'TableOfFigures',
	'{44FEE887-6600-41AB-95A5-DE33C605116C}' : 'OMathRecognizedFunctions',
	'{00020923-0000-0000-C000-000000000046}' : 'ListEntry',
	'{00020924-0000-0000-C000-000000000046}' : 'ListEntries',
	'{00020925-0000-0000-C000-000000000046}' : 'DropDown',
	'{00020926-0000-0000-C000-000000000046}' : 'CheckBox',
	'{00020927-0000-0000-C000-000000000046}' : 'TextInput',
	'{00020928-0000-0000-C000-000000000046}' : 'FormField',
	'{00020929-0000-0000-C000-000000000046}' : 'FormFields',
	'{0002092A-0000-0000-C000-000000000046}' : 'Frame',
	'{0002092B-0000-0000-C000-000000000046}' : 'Frames',
	'{0002092C-0000-0000-C000-000000000046}' : 'Style',
	'{0002092D-0000-0000-C000-000000000046}' : 'Styles',
	'{0002092E-0000-0000-C000-000000000046}' : 'Browser',
	'{0002092F-0000-0000-C000-000000000046}' : 'Field',
	'{00020930-0000-0000-C000-000000000046}' : 'Fields',
	'{00020931-0000-0000-C000-000000000046}' : 'LinkFormat',
	'{00020933-0000-0000-C000-000000000046}' : 'OLEFormat',
	'{00020935-0000-0000-C000-000000000046}' : 'System',
	'{00020936-0000-0000-C000-000000000046}' : 'AutoTextEntry',
	'{00020937-0000-0000-C000-000000000046}' : 'AutoTextEntries',
	'{00020939-0000-0000-C000-000000000046}' : 'TextRetrievalMode',
	'{0002093A-0000-0000-C000-000000000046}' : 'Shading',
	'{0002093B-0000-0000-C000-000000000046}' : 'Border',
	'{0002093C-0000-0000-C000-000000000046}' : 'Borders',
	'{0002093D-0000-0000-C000-000000000046}' : 'Comment',
	'{0002093E-0000-0000-C000-000000000046}' : 'Endnote',
	'{0002093F-0000-0000-C000-000000000046}' : 'Footnote',
	'{00020940-0000-0000-C000-000000000046}' : 'Comments',
	'{00020941-0000-0000-C000-000000000046}' : 'Endnotes',
	'{00020942-0000-0000-C000-000000000046}' : 'Footnotes',
	'{00020943-0000-0000-C000-000000000046}' : 'TwoInitialCapsException',
	'{00020944-0000-0000-C000-000000000046}' : 'TwoInitialCapsExceptions',
	'{00020945-0000-0000-C000-000000000046}' : 'FirstLetterException',
	'{00020946-0000-0000-C000-000000000046}' : 'FirstLetterExceptions',
	'{00020947-0000-0000-C000-000000000046}' : 'AutoCorrectEntry',
	'{00020948-0000-0000-C000-000000000046}' : 'AutoCorrectEntries',
	'{00020949-0000-0000-C000-000000000046}' : 'AutoCorrect',
	'{0002094A-0000-0000-C000-000000000046}' : 'Cells',
	'{0002094B-0000-0000-C000-000000000046}' : 'Columns',
	'{0002094C-0000-0000-C000-000000000046}' : 'Rows',
	'{0002094D-0000-0000-C000-000000000046}' : 'Tables',
	'{0002094E-0000-0000-C000-000000000046}' : 'Cell',
	'{0002094F-0000-0000-C000-000000000046}' : 'Column',
	'{00020950-0000-0000-C000-000000000046}' : 'Row',
	'{00020951-0000-0000-C000-000000000046}' : 'Table',
	'{00020952-0000-0000-C000-000000000046}' : '_Font',
	'{00020953-0000-0000-C000-000000000046}' : '_ParagraphFormat',
	'{00020954-0000-0000-C000-000000000046}' : 'TabStop',
	'{00020955-0000-0000-C000-000000000046}' : 'TabStops',
	'{00020956-0000-0000-C000-000000000046}' : 'DropCap',
	'{00020957-0000-0000-C000-000000000046}' : 'Paragraph',
	'{00020958-0000-0000-C000-000000000046}' : 'Paragraphs',
	'{00020959-0000-0000-C000-000000000046}' : 'Section',
	'{0002095A-0000-0000-C000-000000000046}' : 'Sections',
	'{0002095B-0000-0000-C000-000000000046}' : 'Sentences',
	'{0002095C-0000-0000-C000-000000000046}' : 'Words',
	'{0002095D-0000-0000-C000-000000000046}' : 'Characters',
	'{0002095E-0000-0000-C000-000000000046}' : 'Range',
	'{0002095F-0000-0000-C000-000000000046}' : 'Panes',
	'{00020960-0000-0000-C000-000000000046}' : 'Pane',
	'{00020961-0000-0000-C000-000000000046}' : 'Windows',
	'{00020962-0000-0000-C000-000000000046}' : 'Window',
	'{00020963-0000-0000-C000-000000000046}' : 'RecentFiles',
	'{00020964-0000-0000-C000-000000000046}' : 'RecentFile',
	'{00020965-0000-0000-C000-000000000046}' : 'Variables',
	'{00020966-0000-0000-C000-000000000046}' : 'Variable',
	'{00020967-0000-0000-C000-000000000046}' : 'Bookmarks',
	'{00020968-0000-0000-C000-000000000046}' : 'Bookmark',
	'{00020969-0000-0000-C000-000000000046}' : 'RoutingSlip',
	'{0002096A-0000-0000-C000-000000000046}' : 'Template',
	'{0002096B-0000-0000-C000-000000000046}' : '_Document',
	'{0002096C-0000-0000-C000-000000000046}' : 'Documents',
	'{0002096D-0000-0000-C000-000000000046}' : 'Language',
	'{0002096E-0000-0000-C000-000000000046}' : 'Languages',
	'{0002096F-0000-0000-C000-000000000046}' : 'FontNames',
	'{00020970-0000-0000-C000-000000000046}' : '_Application',
	'{00020971-0000-0000-C000-000000000046}' : 'PageSetup',
	'{00020972-0000-0000-C000-000000000046}' : 'LineNumbering',
	'{00020973-0000-0000-C000-000000000046}' : 'TextColumns',
	'{00020974-0000-0000-C000-000000000046}' : 'TextColumn',
	'{00020975-0000-0000-C000-000000000046}' : 'Selection',
	'{00020976-0000-0000-C000-000000000046}' : 'TablesOfAuthoritiesCategories',
	'{00020977-0000-0000-C000-000000000046}' : 'TableOfAuthoritiesCategory',
	'{00020978-0000-0000-C000-000000000046}' : 'CaptionLabels',
	'{00020979-0000-0000-C000-000000000046}' : 'CaptionLabel',
	'{0002097A-0000-0000-C000-000000000046}' : 'AutoCaptions',
	'{0002097B-0000-0000-C000-000000000046}' : 'AutoCaption',
	'{0002097C-0000-0000-C000-000000000046}' : 'Indexes',
	'{0002097D-0000-0000-C000-000000000046}' : 'Index',
	'{0002097E-0000-0000-C000-000000000046}' : 'AddIn',
	'{0002097F-0000-0000-C000-000000000046}' : 'AddIns',
	'{00020980-0000-0000-C000-000000000046}' : 'Revisions',
	'{00020981-0000-0000-C000-000000000046}' : 'Revision',
	'{00020982-0000-0000-C000-000000000046}' : 'Task',
	'{00020983-0000-0000-C000-000000000046}' : 'Tasks',
	'{00020984-0000-0000-C000-000000000046}' : 'HeadersFooters',
	'{00020985-0000-0000-C000-000000000046}' : 'HeaderFooter',
	'{00020986-0000-0000-C000-000000000046}' : 'PageNumbers',
	'{00020987-0000-0000-C000-000000000046}' : 'PageNumber',
	'{00020988-0000-0000-C000-000000000046}' : 'Subdocuments',
	'{00020989-0000-0000-C000-000000000046}' : 'Subdocument',
	'{0002098A-0000-0000-C000-000000000046}' : 'HeadingStyles',
	'{0002098B-0000-0000-C000-000000000046}' : 'HeadingStyle',
	'{0002098C-0000-0000-C000-000000000046}' : 'StoryRanges',
	'{0002098D-0000-0000-C000-000000000046}' : 'ListLevel',
	'{0002098E-0000-0000-C000-000000000046}' : 'ListLevels',
	'{0002098F-0000-0000-C000-000000000046}' : 'ListTemplate',
	'{00020990-0000-0000-C000-000000000046}' : 'ListTemplates',
	'{00020991-0000-0000-C000-000000000046}' : 'ListParagraphs',
	'{00020992-0000-0000-C000-000000000046}' : 'List',
	'{00020993-0000-0000-C000-000000000046}' : 'Lists',
	'{00020994-0000-0000-C000-000000000046}' : 'ListGallery',
	'{00020995-0000-0000-C000-000000000046}' : 'ListGalleries',
	'{00020996-0000-0000-C000-000000000046}' : 'KeyBindings',
	'{00020997-0000-0000-C000-000000000046}' : 'KeysBoundTo',
	'{00020998-0000-0000-C000-000000000046}' : 'KeyBinding',
	'{00020999-0000-0000-C000-000000000046}' : 'FileConverter',
	'{0002099A-0000-0000-C000-000000000046}' : 'FileConverters',
	'{0002099B-0000-0000-C000-000000000046}' : 'SynonymInfo',
	'{0002099C-0000-0000-C000-000000000046}' : 'Hyperlinks',
	'{0002099D-0000-0000-C000-000000000046}' : 'Hyperlink',
	'{0002099F-0000-0000-C000-000000000046}' : 'Shapes',
	'{000209A0-0000-0000-C000-000000000046}' : 'Shape',
	'{000209A1-0000-0000-C000-000000000046}' : '_LetterContent',
	'{000209A2-0000-0000-C000-000000000046}' : 'Templates',
	'{000209A4-0000-0000-C000-000000000046}' : '_OLEControl',
	'{000209A5-0000-0000-C000-000000000046}' : 'View',
	'{000209A6-0000-0000-C000-000000000046}' : 'Zoom',
	'{000209A7-0000-0000-C000-000000000046}' : 'Zooms',
	'{000209A8-0000-0000-C000-000000000046}' : 'InlineShape',
	'{000209A9-0000-0000-C000-000000000046}' : 'InlineShapes',
	'{000209AA-0000-0000-C000-000000000046}' : 'SpellingSuggestions',
	'{000209AB-0000-0000-C000-000000000046}' : 'SpellingSuggestion',
	'{000209AC-0000-0000-C000-000000000046}' : 'Dictionaries',
	'{000209AD-0000-0000-C000-000000000046}' : 'Dictionary',
	'{000209AE-0000-0000-C000-000000000046}' : 'ReadabilityStatistics',
	'{000209AF-0000-0000-C000-000000000046}' : 'ReadabilityStatistic',
	'{000209B0-0000-0000-C000-000000000046}' : 'Find',
	'{000209B1-0000-0000-C000-000000000046}' : 'Replacement',
	'{000209B2-0000-0000-C000-000000000046}' : 'TextFrame',
	'{000209B3-0000-0000-C000-000000000046}' : 'Versions',
	'{000209B4-0000-0000-C000-000000000046}' : 'Version',
	'{000209B5-0000-0000-C000-000000000046}' : 'ShapeRange',
	'{000209B6-0000-0000-C000-000000000046}' : 'GroupShapes',
	'{000209B7-0000-0000-C000-000000000046}' : 'Options',
	'{000209B8-0000-0000-C000-000000000046}' : 'Dialog',
	'{000209B9-0000-0000-C000-000000000046}' : '_Global',
	'{000209BA-0000-0000-C000-000000000046}' : 'MailMessage',
	'{000209BB-0000-0000-C000-000000000046}' : 'ProofreadingErrors',
	'{000209BD-0000-0000-C000-000000000046}' : 'Mailer',
	'{000209C0-0000-0000-C000-000000000046}' : 'ListFormat',
	'{F258DE05-C41B-4C33-A778-F0D3F98CEEB3}' : 'OMathAcc',
	'{000209C3-0000-0000-C000-000000000046}' : 'WrapFormat',
	'{000209C4-0000-0000-C000-000000000046}' : 'Adjustments',
	'{000209C5-0000-0000-C000-000000000046}' : 'CalloutFormat',
	'{000209C6-0000-0000-C000-000000000046}' : 'ColorFormat',
	'{000209C7-0000-0000-C000-000000000046}' : 'ConnectorFormat',
	'{000209C8-0000-0000-C000-000000000046}' : 'FillFormat',
	'{000209C9-0000-0000-C000-000000000046}' : 'FreeformBuilder',
	'{000209CA-0000-0000-C000-000000000046}' : 'LineFormat',
	'{000209CB-0000-0000-C000-000000000046}' : 'PictureFormat',
	'{000209CC-0000-0000-C000-000000000046}' : 'ShadowFormat',
	'{000209CD-0000-0000-C000-000000000046}' : 'ShapeNode',
	'{000209CE-0000-0000-C000-000000000046}' : 'ShapeNodes',
	'{000209CF-0000-0000-C000-000000000046}' : 'TextEffectFormat',
	'{000209D0-0000-0000-C000-000000000046}' : 'ThreeDFormat',
	'{000209D1-0000-0000-C000-000000000046}' : 'HangulAndAlphabetExceptions',
	'{000209D2-0000-0000-C000-000000000046}' : 'HangulAndAlphabetException',
	'{EE95AFE3-3026-4172-B078-0E79DAB5CC3D}' : 'ContentControl',
	'{000209D7-0000-0000-C000-000000000046}' : 'EmailAuthor',
	'{000209DB-0000-0000-C000-000000000046}' : 'EmailOptions',
	'{000209DC-0000-0000-C000-000000000046}' : 'EmailSignature',
	'{000209DD-0000-0000-C000-000000000046}' : 'Email',
	'{000209DE-0000-0000-C000-000000000046}' : 'HorizontalLineFormat',
	'{000209DF-0000-0000-C000-000000000046}' : 'OtherCorrectionsExceptions',
	'{000209E0-0000-0000-C000-000000000046}' : 'HangulHanjaConversionDictionaries',
	'{000209E1-0000-0000-C000-000000000046}' : 'OtherCorrectionsException',
	'{000209E2-0000-0000-C000-000000000046}' : 'Frameset',
	'{000209E3-0000-0000-C000-000000000046}' : 'DefaultWebOptions',
	'{000209E4-0000-0000-C000-000000000046}' : 'WebOptions',
	'{000209E5-0000-0000-C000-000000000046}' : 'EmailSignatureEntries',
	'{000209E6-0000-0000-C000-000000000046}' : 'EmailSignatureEntry',
	'{000209E7-0000-0000-C000-000000000046}' : 'HTMLDivision',
	'{000209E8-0000-0000-C000-000000000046}' : 'HTMLDivisions',
	'{000209E9-0000-0000-C000-000000000046}' : 'DiagramNode',
	'{000209EA-0000-0000-C000-000000000046}' : 'DiagramNodeChildren',
	'{000209EB-0000-0000-C000-000000000046}' : 'DiagramNodes',
	'{000209EC-0000-0000-C000-000000000046}' : 'Diagram',
	'{000209ED-0000-0000-C000-000000000046}' : 'SmartTag',
	'{000209EE-0000-0000-C000-000000000046}' : 'SmartTags',
	'{000209EF-0000-0000-C000-000000000046}' : 'StyleSheet',
	'{000209F7-0001-0000-C000-000000000046}' : 'IApplicationEvents',
	'{000209FE-0001-0000-C000-000000000046}' : 'IApplicationEvents2',
	'{00020A00-0001-0000-C000-000000000046}' : 'IApplicationEvents3',
	'{C94688A6-A2A7-4133-A26D-726CD569D5F3}' : 'OMathDelim',
	'{7E64D2BE-2818-48CB-8F8A-CC7B61D9E860}' : 'Floor',
	'{86905AC9-33F3-4A88-96C8-B289B0390BCA}' : 'UpBars',
	'{E4442A83-F623-459C-8E95-8BFB44DCF23A}' : 'OMath',
	'{842C37FE-C76F-4B2B-9B60-C408CB5E838E}' : 'OMathBox',
	'{AB0D33A3-C9EA-485B-9443-4C1BB3656CEA}' : 'ChartBorder',
	'{07B7CC7E-E66C-11D3-9454-00105AA31A08}' : 'StyleSheets',
	'{4A6AE865-199D-4EA3-9F6B-125BD9C40EDF}' : 'Source',
	'{00020910-0000-0000-C000-000000000046}' : 'Dialogs',
	'{6E47678B-A879-4E56-8698-3B7CF169FAD4}' : 'Categories',
	'{B9F1A4E2-0D0A-43B7-8495-139E7ACBD840}' : 'TaskPane',
	'{E6AAEC05-E543-4085-BA92-9BF7D2474F51}' : 'Research',
	'{BFD3FC23-F763-4FF8-826E-1AFBF598A4E7}' : 'BuildingBlock',
	'{EFC71F9C-7F42-4CD4-A7A7-970D7A48CD27}' : 'OMathMatCols',
	'{F743EDD0-9B97-4B09-89CC-77BE19B51481}' : 'ProtectedViewWindow',
	'{0D951ADF-10A6-4C9B-BCD9-0FB8CBAD9A87}' : 'OMathFunc',
	'{1FD94DF1-3569-4465-94FF-E8B22D28EEB0}' : 'DataLabel',
	'{39709229-56A0-4E29-9112-B31DD067EBFD}' : 'BuildingBlockEntries',
	'{0002091D-0000-0000-C000-000000000046}' : 'MailMergeDataSource',
	'{9E6B5EC5-E8E4-40AF-9540-6203F71E2823}' : 'CoAuthUpdate',
	'{352840A9-AF7D-4CA4-87FC-21C68FDAB3E4}' : 'Page',
	'{FC9090AF-0DDB-4EC1-86E8-8751F2199F2C}' : 'Gridlines',
	'{B7564E97-0519-4C68-B400-3803E7C63242}' : 'TableStyle',
	'{00020922-0000-0000-C000-000000000046}' : 'TablesOfFigures',
	'{30225CFC-5A71-4FE6-B527-90A52C54AE77}' : 'CoAuthUpdates',
	'{5C04BD93-2F3F-4668-918D-9738EC901039}' : 'OMathRecognizedFunction',
	'{D8252C5E-EB9F-4D74-AA72-C178B128FAC4}' : 'DataLabels',
	'{3E061A7E-67AD-4EAA-BC1E-55057D5E596F}' : 'OMathMat',
	'{0C6FA8CA-E65F-4FC7-AB8F-20729EECBB14}' : 'ContentControlListEntry',
	'{DB8E3072-E106-4453-8E7C-53056F1D30DC}' : 'SmartTagTypes',
	'{804CD967-F83B-432D-9446-C61A45CFEFF0}' : 'ContentControls',
	'{873E774B-926A-4CB1-878D-635A45187595}' : 'OMaths',
	'{BF043168-F4DE-4E7C-B206-741A8B3EF71A}' : 'EndnoteOptions',
	'{F8DDB497-CA6C-4711-9BA4-2718FA3BB6FE}' : 'ChartGroups',
	'{F08B45F1-8F23-4156-9D63-1820C0ED229A}' : 'OMathBar',
	'{7A1BCE11-5783-4C7D-BD02-F3D84AB40E7F}' : 'HiLoLines',
	'{3834F60F-EE8C-455D-A441-D766675D6D3B}' : 'Bibliography',
	'{D36C1F42-7044-4B9E-9CA3-85919454DB04}' : 'XMLNodes',
	'{656BBED7-E82D-4B0A-8F97-EC742BA11FFA}' : 'XMLNamespaces',
	'{194F8476-B79D-4572-A609-294207DE77C1}' : 'ErrorBars',
	'{5DAA8BB6-054E-48F6-BEAC-EAAD02BE0CC7}' : 'OMathMatRow',
	'{356B06EC-4908-42A4-81FC-4B5A51F3483B}' : 'XMLSchemaReferences',
	'{50209974-BA32-4A03-8FA6-BAC56CC056FD}' : 'OMathFrac',
	'{ADD4EDF3-2F33-4734-9CE6-D476097C5ADA}' : 'Rectangle',
	'{4A304B59-31FF-42DD-B436-7FC9C5DB7559}' : 'ChartData',
	'{E6AAEC05-E543-4085-BA92-9BF7D2474F5C}' : 'TaskPanes',
	'{8245795B-9AED-4943-A16D-E586ED8180D1}' : 'OMathArgs',
	'{65DF9F31-B1E3-4651-87E8-51D55F302161}' : 'CoAuthoring',
	'{C4A02049-024C-4273-8934-E48CC21479A9}' : 'LegendEntry',
	'{84A6A663-AEF4-4FCD-83FD-9BB707F157CA}' : 'DownBars',
	'{91C46192-3124-4346-A815-10B8873F5A06}' : 'Trendline',
	'{8B0E45DB-3A7B-42EE-9D17-A92AF69B79C1}' : 'AxisTitle',
	'{6215E4B1-545A-406E-9824-0A5B5AC8AD21}' : 'Conflict',
	'{497142A4-16FD-42C6-BC58-15D89345FC21}' : 'OMathFunctions',
	'{F152D349-7D20-4C01-A42B-2D6DE4F3891C}' : 'ChartFillFormat',
	'{F1F37152-1DB1-4901-AD9A-C740F99464B4}' : 'OMathFunction',
}


NamesToIIDMap = {
	'DataTable' : '{DCE9F2C4-4C02-43BA-840E-B4276550EF79}',
	'HeadingStyles' : '{0002098A-0000-0000-C000-000000000046}',
	'ChartCharacters' : '{FF06FEF2-DA89-41C0-A0A8-5CD434E210AD}',
	'OMathMatCol' : '{CAE36175-3818-4C60-BCBF-0645D51EB33B}',
	'Comments' : '{00020940-0000-0000-C000-000000000046}',
	'Trendlines' : '{54B7061A-D56C-40E5-B85B-58146446C782}',
	'Selection' : '{00020975-0000-0000-C000-000000000046}',
	'Section' : '{00020959-0000-0000-C000-000000000046}',
	'OMath' : '{E4442A83-F623-459C-8E95-8BFB44DCF23A}',
	'KeysBoundTo' : '{00020997-0000-0000-C000-000000000046}',
	'OMathScrSubSup' : '{DC489AD4-23C4-4F4B-990F-45A51C7C0C4F}',
	'ChartGroups' : '{F8DDB497-CA6C-4711-9BA4-2718FA3BB6FE}',
	'Column' : '{0002094F-0000-0000-C000-000000000046}',
	'_OLEControl' : '{000209A4-0000-0000-C000-000000000046}',
	'SeriesLines' : '{5D7F6C15-36CE-44CC-9692-5A1F8B8C906D}',
	'Cell' : '{0002094E-0000-0000-C000-000000000046}',
	'RecentFiles' : '{00020963-0000-0000-C000-000000000046}',
	'_Application' : '{00020970-0000-0000-C000-000000000046}',
	'Hyperlink' : '{0002099D-0000-0000-C000-000000000046}',
	'Corners' : '{AE6D45E5-981E-4547-8752-674BB55420A5}',
	'OMathAutoCorrectEntries' : '{18CD5EC8-8B7B-42C8-992A-2A407468642C}',
	'ShapeRange' : '{000209B5-0000-0000-C000-000000000046}',
	'MailMergeDataField' : '{00020919-0000-0000-C000-000000000046}',
	'OMathFunction' : '{F1F37152-1DB1-4901-AD9A-C740F99464B4}',
	'SmartTagAction' : '{DFB6AA6C-1068-420F-969D-01280FCC1630}',
	'OMathEqArray' : '{1F998A61-71C6-44C2-A0F2-1D66169B47CB}',
	'OMathBorderBox' : '{2503B6EE-0889-44DF-B920-6D6F9659DEA3}',
	'Reviewers' : '{12DCDC9A-5418-48A3-BBE6-EB57BAE275E8}',
	'Shape' : '{000209A0-0000-0000-C000-000000000046}',
	'Pages' : '{91807402-6C6F-47CD-B8FA-C42FEE8EE924}',
	'RecentFile' : '{00020964-0000-0000-C000-000000000046}',
	'OMathRecognizedFunction' : '{5C04BD93-2F3F-4668-918D-9738EC901039}',
	'Tasks' : '{00020983-0000-0000-C000-000000000046}',
	'KeyBindings' : '{00020996-0000-0000-C000-000000000046}',
	'TextColumn' : '{00020974-0000-0000-C000-000000000046}',
	'ApplicationEvents2' : '{000209FE-0000-0000-C000-000000000046}',
	'TablesOfContents' : '{00020914-0000-0000-C000-000000000046}',
	'LegendEntries' : '{B6511068-70BF-4751-A741-55C1D41AD96F}',
	'ApplicationEvents4' : '{00020A01-0000-0000-C000-000000000046}',
	'TableOfContents' : '{00020913-0000-0000-C000-000000000046}',
	'OMathFrac' : '{50209974-BA32-4A03-8FA6-BAC56CC056FD}',
	'ReadabilityStatistics' : '{000209AE-0000-0000-C000-000000000046}',
	'ColorFormat' : '{000209C6-0000-0000-C000-000000000046}',
	'PlotArea' : '{56AFD330-440C-4F4C-A39C-ED306D084D5F}',
	'AutoTextEntry' : '{00020936-0000-0000-C000-000000000046}',
	'Break' : '{79635BF1-BD1D-4B3F-A520-C1106F1AAAD8}',
	'Source' : '{4A6AE865-199D-4EA3-9F6B-125BD9C40EDF}',
	'CustomLabels' : '{00020916-0000-0000-C000-000000000046}',
	'_Font' : '{00020952-0000-0000-C000-000000000046}',
	'OtherCorrectionsExceptions' : '{000209DF-0000-0000-C000-000000000046}',
	'ListEntries' : '{00020924-0000-0000-C000-000000000046}',
	'ListFormat' : '{000209C0-0000-0000-C000-000000000046}',
	'EmailSignature' : '{000209DC-0000-0000-C000-000000000046}',
	'ListLevel' : '{0002098D-0000-0000-C000-000000000046}',
	'HangulAndAlphabetExceptions' : '{000209D1-0000-0000-C000-000000000046}',
	'CheckBox' : '{00020926-0000-0000-C000-000000000046}',
	'Variable' : '{00020966-0000-0000-C000-000000000046}',
	'Zooms' : '{000209A7-0000-0000-C000-000000000046}',
	'TableOfAuthoritiesCategory' : '{00020977-0000-0000-C000-000000000046}',
	'SmartTagRecognizers' : '{F2B60A10-DED5-46FB-A914-3C6F4EBB6451}',
	'HeaderFooter' : '{00020985-0000-0000-C000-000000000046}',
	'Revisions' : '{00020980-0000-0000-C000-000000000046}',
	'Words' : '{0002095C-0000-0000-C000-000000000046}',
	'DefaultWebOptions' : '{000209E3-0000-0000-C000-000000000046}',
	'ChartFont' : '{CDB0FF41-E862-47BB-AE77-3FA7B1AE3189}',
	'LinkFormat' : '{00020931-0000-0000-C000-000000000046}',
	'MailMergeField' : '{0002091E-0000-0000-C000-000000000046}',
	'Fields' : '{00020930-0000-0000-C000-000000000046}',
	'OMathBreak' : '{65E515D5-F50B-4951-8F38-FA6AC8707387}',
	'SmartTagTypes' : '{DB8E3072-E106-4453-8E7C-53056F1D30DC}',
	'MappedDataField' : '{5D311669-EA51-11D3-87CC-00105AA31A34}',
	'OLEFormat' : '{00020933-0000-0000-C000-000000000046}',
	'OMathScrSub' : '{98DFBD12-96CB-4F07-90EA-749FF1D6B89D}',
	'Rows' : '{0002094C-0000-0000-C000-000000000046}',
	'GroupShapes' : '{000209B6-0000-0000-C000-000000000046}',
	'Bookmark' : '{00020968-0000-0000-C000-000000000046}',
	'Options' : '{000209B7-0000-0000-C000-000000000046}',
	'CoAuthLocks' : '{DFF99AC2-CD2A-43AD-91B1-A2BE40BC7146}',
	'PageSetup' : '{00020971-0000-0000-C000-000000000046}',
	'OMathScrSup' : '{D0A95726-678A-4B9D-8103-1E2B86735AE7}',
	'MailMergeFields' : '{0002091F-0000-0000-C000-000000000046}',
	'OMathRecognizedFunctions' : '{44FEE887-6600-41AB-95A5-DE33C605116C}',
	'Lines' : '{E2E8A400-0615-427D-ADCC-CAD39FFEBD42}',
	'Bibliography' : '{3834F60F-EE8C-455D-A441-D766675D6D3B}',
	'Dialogs' : '{00020910-0000-0000-C000-000000000046}',
	'OMathBox' : '{842C37FE-C76F-4B2B-9B60-C408CB5E838E}',
	'Reviewer' : '{47CEF4AE-DC32-4220-8AA5-19CCC0E6633A}',
	'Line' : '{AE6CE2F5-B9D3-407D-85A8-0F10C63289A4}',
	'Revision' : '{00020981-0000-0000-C000-000000000046}',
	'Page' : '{352840A9-AF7D-4CA4-87FC-21C68FDAB3E4}',
	'LineFormat' : '{000209CA-0000-0000-C000-000000000046}',
	'Variables' : '{00020965-0000-0000-C000-000000000046}',
	'FileConverter' : '{00020999-0000-0000-C000-000000000046}',
	'OMathScrPre' : '{AFAF0C0E-8603-40F6-8FD1-42726CAC21E3}',
	'Point' : '{8A342FA0-5831-4B5E-82E1-003D0A0C635D}',
	'Lists' : '{00020993-0000-0000-C000-000000000046}',
	'InlineShapes' : '{000209A9-0000-0000-C000-000000000046}',
	'Shading' : '{0002093A-0000-0000-C000-000000000046}',
	'Envelope' : '{00020918-0000-0000-C000-000000000046}',
	'Template' : '{0002096A-0000-0000-C000-000000000046}',
	'LeaderLines' : '{36162C62-B59A-4278-AF3D-F2AC1EB999D9}',
	'SmartTagType' : '{5E9A888C-E5DC-4DCB-8308-3C91FB61E6F4}',
	'MappedDataFields' : '{799A6814-EA41-11D3-87CC-00105AA31A34}',
	'HeadingStyle' : '{0002098B-0000-0000-C000-000000000046}',
	'CaptionLabel' : '{00020979-0000-0000-C000-000000000046}',
	'Task' : '{00020982-0000-0000-C000-000000000046}',
	'Pane' : '{00020960-0000-0000-C000-000000000046}',
	'Tables' : '{0002094D-0000-0000-C000-000000000046}',
	'BuildingBlockTypes' : '{B5828B50-0E3D-448A-962D-A40702A5868D}',
	'CustomProperty' : '{B923FDE0-F08C-11D3-91B0-00105A0A19FD}',
	'Chart' : '{6FFA84BB-A350-4442-BB53-A43653459A84}',
	'View' : '{000209A5-0000-0000-C000-000000000046}',
	'CustomProperties' : '{B923FDE1-F08C-11D3-91B0-00105A0A19FD}',
	'Rectangle' : '{ADD4EDF3-2F33-4734-9CE6-D476097C5ADA}',
	'AutoCaptions' : '{0002097A-0000-0000-C000-000000000046}',
	'TextRetrievalMode' : '{00020939-0000-0000-C000-000000000046}',
	'MailMergeFieldName' : '{0002091B-0000-0000-C000-000000000046}',
	'FillFormat' : '{000209C8-0000-0000-C000-000000000046}',
	'FootnoteOptions' : '{BEA85A24-D7DA-4F3D-B58C-ED90FB01D615}',
	'KeyBinding' : '{00020998-0000-0000-C000-000000000046}',
	'Dictionaries' : '{000209AC-0000-0000-C000-000000000046}',
	'Sections' : '{0002095A-0000-0000-C000-000000000046}',
	'PageNumbers' : '{00020986-0000-0000-C000-000000000046}',
	'_Global' : '{000209B9-0000-0000-C000-000000000046}',
	'Language' : '{0002096D-0000-0000-C000-000000000046}',
	'OMathBreaks' : '{E2E0F3A7-204C-40C5-BAA5-290F374FDF5A}',
	'FormFields' : '{00020929-0000-0000-C000-000000000046}',
	'CoAuthors' : '{256B6ABA-6A38-4D39-971C-91FDA9922814}',
	'SmartTags' : '{000209EE-0000-0000-C000-000000000046}',
	'ListGallery' : '{00020994-0000-0000-C000-000000000046}',
	'StyleSheet' : '{000209EF-0000-0000-C000-000000000046}',
	'MailingLabel' : '{00020917-0000-0000-C000-000000000046}',
	'TabStop' : '{00020954-0000-0000-C000-000000000046}',
	'ErrorBars' : '{194F8476-B79D-4572-A609-294207DE77C1}',
	'DiagramNode' : '{000209E9-0000-0000-C000-000000000046}',
	'ReflectionFormat' : '{F01943FF-1985-445E-8602-8FB8F39CCA75}',
	'Subdocuments' : '{00020988-0000-0000-C000-000000000046}',
	'Footnotes' : '{00020942-0000-0000-C000-000000000046}',
	'TextColumns' : '{00020973-0000-0000-C000-000000000046}',
	'TablesOfFigures' : '{00020922-0000-0000-C000-000000000046}',
	'OMathMatRow' : '{5DAA8BB6-054E-48F6-BEAC-EAAD02BE0CC7}',
	'FirstLetterExceptions' : '{00020946-0000-0000-C000-000000000046}',
	'Shapes' : '{0002099F-0000-0000-C000-000000000046}',
	'Window' : '{00020962-0000-0000-C000-000000000046}',
	'OMathLimUpp' : '{FC9086C6-0287-4997-B2E1-816C334A22F8}',
	'CoAuthUpdates' : '{30225CFC-5A71-4FE6-B527-90A52C54AE77}',
	'AddIn' : '{0002097E-0000-0000-C000-000000000046}',
	'Rectangles' : '{7D0F7985-68D9-4D93-91CB-8109280E76CC}',
	'ChartFillFormat' : '{F152D349-7D20-4C01-A42B-2D6DE4F3891C}',
	'XMLSchemaReference' : '{FE0971F0-5E60-4985-BCDA-95CB0B8E0308}',
	'_LetterContent' : '{000209A1-0000-0000-C000-000000000046}',
	'TabStops' : '{00020955-0000-0000-C000-000000000046}',
	'Categories' : '{6E47678B-A879-4E56-8698-3B7CF169FAD4}',
	'OMaths' : '{873E774B-926A-4CB1-878D-635A45187595}',
	'ChartFormat' : '{B66D3C1A-4541-4961-B35B-A353C03F6A99}',
	'Table' : '{00020951-0000-0000-C000-000000000046}',
	'ConnectorFormat' : '{000209C7-0000-0000-C000-000000000046}',
	'Legend' : '{B3A1E8C6-E1CE-4A46-8D12-E017157B03D7}',
	'CustomLabel' : '{00020915-0000-0000-C000-000000000046}',
	'XMLNodes' : '{D36C1F42-7044-4B9E-9CA3-85919454DB04}',
	'Range' : '{0002095E-0000-0000-C000-000000000046}',
	'ListParagraphs' : '{00020991-0000-0000-C000-000000000046}',
	'System' : '{00020935-0000-0000-C000-000000000046}',
	'ChartGroup' : '{86488FB4-9633-4C93-8057-FC1FA7A847AE}',
	'OMathMatCols' : '{EFC71F9C-7F42-4CD4-A7A7-970D7A48CD27}',
	'LineNumbering' : '{00020972-0000-0000-C000-000000000046}',
	'IApplicationEvents' : '{000209F7-0001-0000-C000-000000000046}',
	'FontNames' : '{0002096F-0000-0000-C000-000000000046}',
	'OMathAutoCorrectEntry' : '{D8779F01-4869-4403-B334-D60C5F9C9175}',
	'OMathAcc' : '{F258DE05-C41B-4C33-A778-F0D3F98CEEB3}',
	'ReadabilityStatistic' : '{000209AF-0000-0000-C000-000000000046}',
	'Zoom' : '{000209A6-0000-0000-C000-000000000046}',
	'MailMergeFieldNames' : '{0002091C-0000-0000-C000-000000000046}',
	'CoAuthoring' : '{65DF9F31-B1E3-4651-87E8-51D55F302161}',
	'SmartTagActions' : '{CDE12CD8-767B-4757-8A31-13029A086305}',
	'OMathFunctions' : '{497142A4-16FD-42C6-BC58-15D89345FC21}',
	'HiLoLines' : '{7A1BCE11-5783-4C7D-BD02-F3D84AB40E7F}',
	'ThreeDFormat' : '{000209D0-0000-0000-C000-000000000046}',
	'ProtectedViewWindows' : '{FD0A74E8-C719-49F6-BA1B-F6D9839D1AB9}',
	'Indexes' : '{0002097C-0000-0000-C000-000000000046}',
	'Windows' : '{00020961-0000-0000-C000-000000000046}',
	'DropDown' : '{00020925-0000-0000-C000-000000000046}',
	'Axes' : '{354AB591-A217-48B4-99E4-14F58F15667D}',
	'XMLNode' : '{09760240-0B89-49F7-A79D-479F24723F56}',
	'Adjustments' : '{000209C4-0000-0000-C000-000000000046}',
	'StoryRanges' : '{0002098C-0000-0000-C000-000000000046}',
	'Trendline' : '{91C46192-3124-4346-A815-10B8873F5A06}',
	'Mailer' : '{000209BD-0000-0000-C000-000000000046}',
	'List' : '{00020992-0000-0000-C000-000000000046}',
	'Paragraphs' : '{00020958-0000-0000-C000-000000000046}',
	'_Document' : '{0002096B-0000-0000-C000-000000000046}',
	'XMLChildNodeSuggestion' : '{A87E00E9-3AC3-4B53-ABE3-7379653D0E82}',
	'DropLines' : '{9F1DF642-3CCE-4D83-A770-D2634A05D278}',
	'BuildingBlockType' : '{817F99FA-CCC4-4971-8E9D-1238F735AAFF}',
	'Cells' : '{0002094A-0000-0000-C000-000000000046}',
	'EndnoteOptions' : '{BF043168-F4DE-4E7C-B206-741A8B3EF71A}',
	'ListEntry' : '{00020923-0000-0000-C000-000000000046}',
	'Endnotes' : '{00020941-0000-0000-C000-000000000046}',
	'AutoCaption' : '{0002097B-0000-0000-C000-000000000046}',
	'TextFrame' : '{000209B2-0000-0000-C000-000000000046}',
	'Panes' : '{0002095F-0000-0000-C000-000000000046}',
	'Points' : '{C1A870A0-850E-4D38-98A7-741CB8C3BCA4}',
	'AutoCorrectEntry' : '{00020947-0000-0000-C000-000000000046}',
	'Columns' : '{0002094B-0000-0000-C000-000000000046}',
	'Series' : '{40810760-068A-4486-BEC9-8EA58C7029F5}',
	'TablesOfAuthorities' : '{00020912-0000-0000-C000-000000000046}',
	'ListLevels' : '{0002098E-0000-0000-C000-000000000046}',
	'Gridlines' : '{FC9090AF-0DDB-4EC1-86E8-8751F2199F2C}',
	'BuildingBlockEntries' : '{39709229-56A0-4E29-9112-B31DD067EBFD}',
	'SoftEdgeFormat' : '{D040DAF9-6CE4-4BE8-839D-F4538A4301CF}',
	'Comment' : '{0002093D-0000-0000-C000-000000000046}',
	'OMathAutoCorrect' : '{6F9D1F68-06F7-49EF-8902-185E54EB5E87}',
	'OMathGroupChar' : '{02B17CB4-7D55-4B34-B38B-10381433441F}',
	'HorizontalLineFormat' : '{000209DE-0000-0000-C000-000000000046}',
	'WebOptions' : '{000209E4-0000-0000-C000-000000000046}',
	'MailMergeDataSource' : '{0002091D-0000-0000-C000-000000000046}',
	'_ParagraphFormat' : '{00020953-0000-0000-C000-000000000046}',
	'Versions' : '{000209B3-0000-0000-C000-000000000046}',
	'Bookmarks' : '{00020967-0000-0000-C000-000000000046}',
	'SeriesCollection' : '{8FEB78F7-35C6-4871-918C-193C3CDD886D}',
	'DataLabels' : '{D8252C5E-EB9F-4D74-AA72-C178B128FAC4}',
	'TwoInitialCapsExceptions' : '{00020944-0000-0000-C000-000000000046}',
	'SmartTagRecognizer' : '{15EBE471-0182-4CCE-98D0-B6614D1C32A1}',
	'AutoTextEntries' : '{00020937-0000-0000-C000-000000000046}',
	'Style' : '{0002092C-0000-0000-C000-000000000046}',
	'GlowFormat' : '{F1B14F40-5C32-4C8C-B5B2-DE537BB6B89D}',
	'Sources' : '{FA02A26B-6550-45C5-B6F0-80E757CD3482}',
	'Breaks' : '{16BE9309-D708-4322-BB1A-B056F58D17EA}',
	'ChartArea' : '{C75AD98A-74E9-49FE-8BF1-544839CC08A5}',
	'XSLTransforms' : '{C774F5EA-A539-4284-A1BE-30AEC052D899}',
	'Browser' : '{0002092E-0000-0000-C000-000000000046}',
	'StyleSheets' : '{07B7CC7E-E66C-11D3-9454-00105AA31A08}',
	'ChartData' : '{4A304B59-31FF-42DD-B436-7FC9C5DB7559}',
	'ListTemplates' : '{00020990-0000-0000-C000-000000000046}',
	'OMathArgs' : '{8245795B-9AED-4943-A16D-E586ED8180D1}',
	'BuildingBlock' : '{BFD3FC23-F763-4FF8-826E-1AFBF598A4E7}',
	'ApplicationEvents' : '{000209F7-0000-0000-C000-000000000046}',
	'PageNumber' : '{00020987-0000-0000-C000-000000000046}',
	'UpBars' : '{86905AC9-33F3-4A88-96C8-B289B0390BCA}',
	'TaskPane' : '{B9F1A4E2-0D0A-43B7-8495-139E7ACBD840}',
	'ApplicationEvents3' : '{00020A00-0000-0000-C000-000000000046}',
	'OMathRad' : '{67A7EEC5-285D-4024-B071-BD6B33B88547}',
	'Replacement' : '{000209B1-0000-0000-C000-000000000046}',
	'LegendEntry' : '{C4A02049-024C-4273-8934-E48CC21479A9}',
	'XMLChildNodeSuggestions' : '{DE63B5AC-CA4F-46FE-9184-A5719AB9ED5B}',
	'DiagramNodes' : '{000209EB-0000-0000-C000-000000000046}',
	'OMathLimLow' : '{74DE9576-8E99-4E28-912B-CB30747C60CE}',
	'Subdocument' : '{00020989-0000-0000-C000-000000000046}',
	'TableStyle' : '{B7564E97-0519-4C68-B400-3803E7C63242}',
	'CanvasShapes' : '{396F9073-F9FD-11D3-8EA0-0050049A1A01}',
	'SmartTag' : '{000209ED-0000-0000-C000-000000000046}',
	'CoAuthor' : '{E59544D5-C299-46A0-84C1-C51AB38F9759}',
	'Dictionary' : '{000209AD-0000-0000-C000-000000000046}',
	'Axis' : '{7EBC66BD-F788-42C3-91F4-E8C841A69005}',
	'TextInput' : '{00020927-0000-0000-C000-000000000046}',
	'OMathMatRows' : '{1B426348-607D-433C-9216-C5D2BF0EF31F}',
	'TwoInitialCapsException' : '{00020943-0000-0000-C000-000000000046}',
	'DropCap' : '{00020956-0000-0000-C000-000000000046}',
	'DocumentEvents2' : '{00020A02-0000-0000-C000-000000000046}',
	'MailMerge' : '{00020920-0000-0000-C000-000000000046}',
	'Hyperlinks' : '{0002099C-0000-0000-C000-000000000046}',
	'ShadowFormat' : '{000209CC-0000-0000-C000-000000000046}',
	'DataLabel' : '{1FD94DF1-3569-4465-94FF-E8B22D28EEB0}',
	'Research' : '{E6AAEC05-E543-4085-BA92-9BF7D2474F51}',
	'BuildingBlocks' : '{C6D50987-25D7-408A-BFF2-90BF86A24E93}',
	'Conflicts' : '{C2B83A65-B061-4469-83B6-8877437CB8A0}',
	'ChartBorder' : '{AB0D33A3-C9EA-485B-9443-4C1BB3656CEA}',
	'Frame' : '{0002092A-0000-0000-C000-000000000046}',
	'Row' : '{00020950-0000-0000-C000-000000000046}',
	'Footnote' : '{0002093F-0000-0000-C000-000000000046}',
	'DiagramNodeChildren' : '{000209EA-0000-0000-C000-000000000046}',
	'FreeformBuilder' : '{000209C9-0000-0000-C000-000000000046}',
	'InlineShape' : '{000209A8-0000-0000-C000-000000000046}',
	'Dialog' : '{000209B8-0000-0000-C000-000000000046}',
	'Frameset' : '{000209E2-0000-0000-C000-000000000046}',
	'ChartTitle' : '{C1AD33E4-F088-40A9-9D2F-D94017D115C4}',
	'OtherCorrectionsException' : '{000209E1-0000-0000-C000-000000000046}',
	'OMathFunc' : '{0D951ADF-10A6-4C9B-BCD9-0FB8CBAD9A87}',
	'SpellingSuggestion' : '{000209AB-0000-0000-C000-000000000046}',
	'Email' : '{000209DD-0000-0000-C000-000000000046}',
	'ContentControl' : '{EE95AFE3-3026-4172-B078-0E79DAB5CC3D}',
	'Category' : '{ECFBDB5E-ACD2-4530-AD79-4560B7FF055C}',
	'Sentences' : '{0002095B-0000-0000-C000-000000000046}',
	'Walls' : '{A2E94180-7564-4D97-806B-BBC0D0A1350C}',
	'AxisTitle' : '{8B0E45DB-3A7B-42EE-9D17-A92AF69B79C1}',
	'FileConverters' : '{0002099A-0000-0000-C000-000000000046}',
	'Editor' : '{DD947D72-F33C-4198-9BDF-F86181D05E41}',
	'XMLNamespaces' : '{656BBED7-E82D-4B0A-8F97-EC742BA11FFA}',
	'Languages' : '{0002096E-0000-0000-C000-000000000046}',
	'TableOfFigures' : '{00020921-0000-0000-C000-000000000046}',
	'CoAuthUpdate' : '{9E6B5EC5-E8E4-40AF-9540-6203F71E2823}',
	'MailMessage' : '{000209BA-0000-0000-C000-000000000046}',
	'TablesOfAuthoritiesCategories' : '{00020976-0000-0000-C000-000000000046}',
	'Templates' : '{000209A2-0000-0000-C000-000000000046}',
	'Version' : '{000209B4-0000-0000-C000-000000000046}',
	'TaskPanes' : '{E6AAEC05-E543-4085-BA92-9BF7D2474F5C}',
	'Find' : '{000209B0-0000-0000-C000-000000000046}',
	'ShapeNodes' : '{000209CE-0000-0000-C000-000000000046}',
	'Styles' : '{0002092D-0000-0000-C000-000000000046}',
	'CalloutFormat' : '{000209C5-0000-0000-C000-000000000046}',
	'EmailSignatureEntry' : '{000209E6-0000-0000-C000-000000000046}',
	'SpellingSuggestions' : '{000209AA-0000-0000-C000-000000000046}',
	'ContentControlListEntry' : '{0C6FA8CA-E65F-4FC7-AB8F-20729EECBB14}',
	'Conflict' : '{6215E4B1-545A-406E-9824-0A5B5AC8AD21}',
	'FormField' : '{00020928-0000-0000-C000-000000000046}',
	'XMLNamespace' : '{B140A023-4850-4DA6-BC5F-CC459C4507BC}',
	'OMathMat' : '{3E061A7E-67AD-4EAA-BC1E-55057D5E596F}',
	'HangulAndAlphabetException' : '{000209D2-0000-0000-C000-000000000046}',
	'HangulHanjaConversionDictionaries' : '{000209E0-0000-0000-C000-000000000046}',
	'UndoRecord' : '{E598E358-2852-42D4-8775-160BD91B7244}',
	'HeadersFooters' : '{00020984-0000-0000-C000-000000000046}',
	'Endnote' : '{0002093E-0000-0000-C000-000000000046}',
	'AutoCorrectEntries' : '{00020948-0000-0000-C000-000000000046}',
	'ShapeNode' : '{000209CD-0000-0000-C000-000000000046}',
	'TextEffectFormat' : '{000209CF-0000-0000-C000-000000000046}',
	'IApplicationEvents2' : '{000209FE-0001-0000-C000-000000000046}',
	'IApplicationEvents3' : '{00020A00-0001-0000-C000-000000000046}',
	'IApplicationEvents4' : '{00020A01-0001-0000-C000-000000000046}',
	'Characters' : '{0002095D-0000-0000-C000-000000000046}',
	'ContentControls' : '{804CD967-F83B-432D-9446-C61A45CFEFF0}',
	'OMathBar' : '{F08B45F1-8F23-4156-9D63-1820C0ED229A}',
	'EmailSignatureEntries' : '{000209E5-0000-0000-C000-000000000046}',
	'PictureFormat' : '{000209CB-0000-0000-C000-000000000046}',
	'Field' : '{0002092F-0000-0000-C000-000000000046}',
	'ContentControlListEntries' : '{54F46DC4-F6A6-48CC-BD66-46C1DDEADD22}',
	'XSLTransform' : '{E3124493-7D6A-410F-9A48-CC822C033CEC}',
	'RoutingSlip' : '{00020969-0000-0000-C000-000000000046}',
	'EmailAuthor' : '{000209D7-0000-0000-C000-000000000046}',
	'Editors' : '{AED7E08C-14F0-4F33-921D-4C5353137BF6}',
	'Interior' : '{B184502B-587A-4C6A-8DC4-ECE4354883C6}',
	'OCXEvents' : '{000209F3-0000-0000-C000-000000000046}',
	'AddIns' : '{0002097F-0000-0000-C000-000000000046}',
	'ProtectedViewWindow' : '{F743EDD0-9B97-4B09-89CC-77BE19B51481}',
	'OMathDelim' : '{C94688A6-A2A7-4133-A26D-726CD569D5F3}',
	'ListTemplate' : '{0002098F-0000-0000-C000-000000000046}',
	'Paragraph' : '{00020957-0000-0000-C000-000000000046}',
	'DisplayUnitLabel' : '{C04865A3-9F8A-486C-BB58-B4C3E6563136}',
	'TickLabels' : '{935D59F5-E365-4F92-B7F5-1C499A63ECA8}',
	'MailMergeDataFields' : '{0002091A-0000-0000-C000-000000000046}',
	'ChartColorFormat' : '{DD8F80B8-9B80-4E89-9BEC-F12DF35E43B3}',
	'ConditionalStyle' : '{1498F56D-ED33-41F9-B37B-EF30E50B08AC}',
	'ProofreadingErrors' : '{000209BB-0000-0000-C000-000000000046}',
	'Frames' : '{0002092B-0000-0000-C000-000000000046}',
	'ListGalleries' : '{00020995-0000-0000-C000-000000000046}',
	'Index' : '{0002097D-0000-0000-C000-000000000046}',
	'Diagram' : '{000209EC-0000-0000-C000-000000000046}',
	'CaptionLabels' : '{00020978-0000-0000-C000-000000000046}',
	'FirstLetterException' : '{00020945-0000-0000-C000-000000000046}',
	'EmailOptions' : '{000209DB-0000-0000-C000-000000000046}',
	'CoAuthLock' : '{99755F80-FE96-4F7D-B636-B8E800E54F44}',
	'XMLSchemaReferences' : '{356B06EC-4908-42A4-81FC-4B5A51F3483B}',
	'TableOfAuthorities' : '{00020911-0000-0000-C000-000000000046}',
	'OMathNary' : '{CEBD4184-4E6D-4FC6-A42D-2142B1B76AF5}',
	'HTMLDivision' : '{000209E7-0000-0000-C000-000000000046}',
	'Border' : '{0002093B-0000-0000-C000-000000000046}',
	'HTMLDivisions' : '{000209E8-0000-0000-C000-000000000046}',
	'Borders' : '{0002093C-0000-0000-C000-000000000046}',
	'Documents' : '{0002096C-0000-0000-C000-000000000046}',
	'OMathPhantom' : '{DB77D541-85C3-42E8-8649-AFBD7CF87866}',
	'LegendKey' : '{DF076FDE-8781-4051-A5BC-99F6B7DC04D4}',
	'DownBars' : '{84A6A663-AEF4-4FCD-83FD-9BB707F157CA}',
	'AutoCorrect' : '{00020949-0000-0000-C000-000000000046}',
	'WrapFormat' : '{000209C3-0000-0000-C000-000000000046}',
	'Floor' : '{7E64D2BE-2818-48CB-8F8A-CC7B61D9E860}',
	'DocumentEvents' : '{000209F6-0000-0000-C000-000000000046}',
	'SynonymInfo' : '{0002099B-0000-0000-C000-000000000046}',
	'XMLMapping' : '{0C1FABE7-F737-406F-9CA3-B07661F9D1A2}',
}

win32com.client.constants.__dicts__.append(constants.__dict__)

