# -*- coding: utf-8 -*-
# Created by makepy.py version 0.5.01
# By python version 3.13.2 (tags/v3.13.2:4f8bb39, Feb  4 2025, 15:23:48) [MSC v.1942 64 bit (AMD64)]
# From type library 'composer.exe'
# On Mon Feb 24 15:25:59 2025
''
makepy_version = '0.5.01'
python_version = 0x30d02f0

import win32com.client.CLSIDToClass, pythoncom, pywintypes
import win32com.client.util
from pywintypes import IID
from win32com.client import Dispatch

# The following 3 lines may need tweaking for the particular server
# Candidates are pythoncom.Missing, .Empty and .ArgNotFound
defaultNamedOptArg=pythoncom.Empty
defaultNamedNotOptArg=pythoncom.Empty
defaultUnnamedArg=pythoncom.Empty

CLSID = IID('{A71E1897-7597-441A-B1D8-725087CBF5FD}')
MajorVersion = 1
MinorVersion = 0
LibraryFlags = 8
LCID = 0x0

class constants:
	crAtvAllActors                =0          # from enum CrActorVisibility
	crAtvGhost                    =3          # from enum CrActorVisibility
	crAtvInvisible                =2          # from enum CrActorVisibility
	crAtvKeepActorVisibility      =0          # from enum CrActorVisibility
	crAtvUndefinedValues          =0          # from enum CrActorVisibility
	crAtvVisible                  =1          # from enum CrActorVisibility
	crApplicationProfileHighQuality=3          # from enum CrApplicationProfile
	crApplicationProfileHighSpeed =2          # from enum CrApplicationProfile
	crApplicationProfileSafe      =4          # from enum CrApplicationProfile
	crApplicationProfileStandard  =0          # from enum CrApplicationProfile
	crbAssemblyTree               =1          # from enum CrBars
	crbBOMTree                    =5          # from enum CrBars
	crbBarsFirst                  =1          # from enum CrBars
	crbBarsLast                   =11         # from enum CrBars
	crbCollabTree                 =2          # from enum CrBars
	crbInformation                =10         # from enum CrBars
	crbLayers                     =6          # from enum CrBars
	crbMarkers                    =7          # from enum CrBars
	crbProperty                   =3          # from enum CrBars
	crbTaskPanes                  =9          # from enum CrBars
	crbTimeline                   =8          # from enum CrBars
	crbView                       =4          # from enum CrBars
	crCalloutGenAttachCenterBoundingBox=0          # from enum CrCalloutGenAttach
	crCalloutGenAttachNearestToCenter=1          # from enum CrCalloutGenAttach
	crCalloutGenModeEachBomIDMulti=2          # from enum CrCalloutGenMode
	crCalloutGenModeEachBomIDSingle=1          # from enum CrCalloutGenMode
	crCalloutGenModeEachInstance  =0          # from enum CrCalloutGenMode
	crClfAnnotation               =1          # from enum CrCollaborationFamily
	crClfCallout                  =2          # from enum CrCollaborationFamily
	crClfCamera                   =3          # from enum CrCollaborationFamily
	crClfCoordinateSystem         =4          # from enum CrCollaborationFamily
	crClfCuttingPlane             =5          # from enum CrCollaborationFamily
	crClfEnvironment              =6          # from enum CrCollaborationFamily
	crClfIntersectionLines        =7          # from enum CrCollaborationFamily
	crClfLighting                 =8          # from enum CrCollaborationFamily
	crClfMagneticLine             =9          # from enum CrCollaborationFamily
	crClfMarkups                  =10         # from enum CrCollaborationFamily
	crClfMeasurement              =11         # from enum CrCollaborationFamily
	crClfNoCollaboration          =0          # from enum CrCollaborationFamily
	crClfPMI                      =14         # from enum CrCollaborationFamily
	crClfPanel                    =12         # from enum CrCollaborationFamily
	crClfPaths                    =13         # from enum CrCollaborationFamily
	crDgrCuttingPlane             =2          # from enum CrDiggerRender
	crDgrOignonSkin               =0          # from enum CrDiggerRender
	crDgrOnionSkin                =0          # from enum CrDiggerRender
	crDgrXRay                     =1          # from enum CrDiggerRender
	crDgrZoom                     =3          # from enum CrDiggerRender
	crAskForAgreement             =0          # from enum CrDisplaySignatureMode
	crNoDisplay                   =2          # from enum CrDisplaySignatureMode
	crShowOnly                    =1          # from enum CrDisplaySignatureMode
	crDt3D                        =2          # from enum CrDocumentType
	crDtNoOpenedDoc               =1          # from enum CrDocumentType
	crFlrAdvancedActiveX          =16         # from enum CrFileRights
	crFlrAdvancedMetaProperties   =8192       # from enum CrFileRights
	crFlrAnnotations              =1          # from enum CrFileRights
	crFlrBOMs                     =256        # from enum CrFileRights
	crFlrClashDetection           =8          # from enum CrFileRights
	crFlrCrossHighlight           =4096       # from enum CrFileRights
	crFlrCuttingPlanes            =2          # from enum CrFileRights
	crFlrForbidDeleteActors       =16384      # from enum CrFileRights
	crFlrForbidDeleteParts        =2048       # from enum CrFileRights
	crFlrFreezePublication        =1024       # from enum CrFileRights
	crFlrLoadingOnDemand          =128        # from enum CrFileRights
	crFlrNone                     =0          # from enum CrFileRights
	crFlrSave                     =4          # from enum CrFileRights
	crFlrSimulation               =64         # from enum CrFileRights
	crFlrTrees                    =32         # from enum CrFileRights
	crLycAssy                     =1          # from enum CrLayerColor
	crLycAssyHighlight            =3          # from enum CrLayerColor
	crLycDefaultSelection         =0          # from enum CrLayerColor
	crLycHighlight                =2          # from enum CrLayerColor
	crLycPMICrsHLHighlight        =5          # from enum CrLayerColor
	crLycPMICrsHLSelection        =4          # from enum CrLayerColor
	crMpaAnimable                 =2          # from enum CrMetaPropertyAttribute
	crMpaAnimatable               =2          # from enum CrMetaPropertyAttribute
	crMpaMergeable                =4          # from enum CrMetaPropertyAttribute
	crMpaVisible                  =8          # from enum CrMetaPropertyAttribute
	crMptBoolean                  =0          # from enum CrMetaPropertyType
	crMptColor                    =5          # from enum CrMetaPropertyType
	crMptDate                     =4          # from enum CrMetaPropertyType
	crMptFloating                 =2          # from enum CrMetaPropertyType
	crMptInteger                  =1          # from enum CrMetaPropertyType
	crMptMetaPropertyTypeFirst    =0          # from enum CrMetaPropertyType
	crMptMetaPropertyTypeLast     =6          # from enum CrMetaPropertyType
	crMptString                   =3          # from enum CrMetaPropertyType
	crMpuAcceleration             =1          # from enum CrMetaPropertyUnit
	crMpuAngle                    =2          # from enum CrMetaPropertyUnit
	crMpuArea                     =3          # from enum CrMetaPropertyUnit
	crMpuLength                   =4          # from enum CrMetaPropertyUnit
	crMpuNone                     =0          # from enum CrMetaPropertyUnit
	crMpuSpeed                    =5          # from enum CrMetaPropertyUnit
	crMpuTime                     =6          # from enum CrMetaPropertyUnit
	crMpuVolume                   =7          # from enum CrMetaPropertyUnit
	crNavigationFlyThrough        =4          # from enum CrMouseNavigation
	crNavigationPan               =2          # from enum CrMouseNavigation
	crNavigationRotate            =1          # from enum CrMouseNavigation
	crNavigationSelect            =0          # from enum CrMouseNavigation
	crNavigationZoom              =3          # from enum CrMouseNavigation
	crObjAllChildren              =7          # from enum CrObjectType
	crObjAssys                    =1          # from enum CrObjectType
	crObjCADCaptures              =4          # from enum CrObjectType
	crObjCADViews                 =2          # from enum CrObjectType
	crRfsHighSpeedRefresh         =0          # from enum CrRefreshType
	crRfsRefreshAntiAliasing      =2          # from enum CrRefreshType
	crRfsScene                    =1          # from enum CrRefreshType
	crRrmCustomRendering          =11         # from enum CrRender
	crRrmFlat                     =5          # from enum CrRender
	crRrmPoint                    =7          # from enum CrRender
	crRrmRenderFirst              =0          # from enum CrRender
	crRrmRenderLast               =12         # from enum CrRender
	crRrmShadedIllustration       =8          # from enum CrRender
	crRrmShadedIllustrationSimple =10         # from enum CrRender
	crRrmShadedIllustrationWithColors=9          # from enum CrRender
	crRrmSilhouette               =4          # from enum CrRender
	crRrmSmooth                   =0          # from enum CrRender
	crRrmSmoothWithOutlines       =1          # from enum CrRender
	crRrmTechnical                =2          # from enum CrRender
	crRrmTechnicalFlat            =3          # from enum CrRender
	crRrmWireframe                =6          # from enum CrRender
	crRCmd1Pane                   =287        # from enum CrRibbonCommand
	crRCmd2HorizontalPanes        =288        # from enum CrRibbonCommand
	crRCmd2VerticalPanes          =289        # from enum CrRibbonCommand
	crRCmd34XMYMZP                =32         # from enum CrRibbonCommand
	crRCmd34XMYPZP                =33         # from enum CrRibbonCommand
	crRCmd34XPYMZP                =31         # from enum CrRibbonCommand
	crRCmd34XPYPZP                =30         # from enum CrRibbonCommand
	crRCmd3Panes1Down2VerticalUp  =293        # from enum CrRibbonCommand
	crRCmd3Panes1Left2HorizontalRight=290        # from enum CrRibbonCommand
	crRCmd3Panes1Right2HorizontalLeft=291        # from enum CrRibbonCommand
	crRCmd3Panes1Up2VerticalDown  =292        # from enum CrRibbonCommand
	crRCmd4Panes                  =294        # from enum CrRibbonCommand
	crRCmdAO                      =98         # from enum CrRibbonCommand
	crRCmdAddCamera               =62         # from enum CrRibbonCommand
	crRCmdAddCurrentLinksToRigidSet=248        # from enum CrRibbonCommand
	crRCmdAddNode                 =115        # from enum CrRibbonCommand
	crRCmdAddRigidParentLevel     =244        # from enum CrRibbonCommand
	crRCmdAlignCameraOnFace       =26         # from enum CrRibbonCommand
	crRCmdAlignMagneticLine       =111        # from enum CrRibbonCommand
	crRCmdAlignPivotWithGivenPivot=229        # from enum CrRibbonCommand
	crRCmdAlignPivotWithLineAxis  =232        # from enum CrRibbonCommand
	crRCmdAlignPivotWithLocalAxis =230        # from enum CrRibbonCommand
	crRCmdAlignPivotWithWorldAxis =231        # from enum CrRibbonCommand
	crRCmdAngle                   =173        # from enum CrRibbonCommand
	crRCmdAnimation               =63         # from enum CrRibbonCommand
	crRCmdAnimationLibrary        =331        # from enum CrRibbonCommand
	crRCmdApplyToSelection        =188        # from enum CrRibbonCommand
	crRCmdArcLength               =155        # from enum CrRibbonCommand
	crRCmdArea                    =146        # from enum CrRibbonCommand
	crRCmdArrow                   =121        # from enum CrRibbonCommand
	crRCmdAssemblyTree            =297        # from enum CrRibbonCommand
	crRCmdAttachCOI               =55         # from enum CrRibbonCommand
	crRCmdAttachCamera            =54         # from enum CrRibbonCommand
	crRCmdAttachCenter            =56         # from enum CrRibbonCommand
	crRCmdBOM                     =278        # from enum CrRibbonCommand
	crRCmdBOMTable                =23         # from enum CrRibbonCommand
	crRCmdBOMTree                 =299        # from enum CrRibbonCommand
	crRCmdBoundingBoxes           =81         # from enum CrRibbonCommand
	crRCmdBrokenEdges             =83         # from enum CrRibbonCommand
	crRCmdCallout                 =150        # from enum CrRibbonCommand
	crRCmdCallouts                =22         # from enum CrRibbonCommand
	crRCmdCameraFollowObjectKeepPosition=60         # from enum CrRibbonCommand
	crRCmdCameraFollowObjectRigid =61         # from enum CrRibbonCommand
	crRCmdCameraPathCustom        =319        # from enum CrRibbonCommand
	crRCmdCameraPathLinear        =318        # from enum CrRibbonCommand
	crRCmdCameraPathMenu          =316        # from enum CrRibbonCommand
	crRCmdCameraPathSmooth        =317        # from enum CrRibbonCommand
	crRCmdCameraPlayMode          =326        # from enum CrRibbonCommand
	crRCmdChainMode               =171        # from enum CrRibbonCommand
	crRCmdCircle                  =124        # from enum CrRibbonCommand
	crRCmdCircularArrow           =122        # from enum CrRibbonCommand
	crRCmdClearRigidLinksSet      =250        # from enum CrRibbonCommand
	crRCmdClearTracks             =307        # from enum CrRibbonCommand
	crRCmdClearanceChecking       =283        # from enum CrRibbonCommand
	crRCmdCollaboration           =21         # from enum CrRibbonCommand
	crRCmdCollaborationTree       =298        # from enum CrRibbonCommand
	crRCmdCompositePolyline       =119        # from enum CrRibbonCommand
	crRCmdConfigs                 =332        # from enum CrRibbonCommand
	crRCmdConstructionEdges       =84         # from enum CrRibbonCommand
	crRCmdCoordinateSystem        =182        # from enum CrRibbonCommand
	crRCmdCoordinateSystemOnGeometry=181        # from enum CrRibbonCommand
	crRCmdCoordinates             =149        # from enum CrRibbonCommand
	crRCmdCopy                    =257        # from enum CrRibbonCommand
	crRCmdCopyFromActor           =311        # from enum CrRibbonCommand
	crRCmdCopyLocationInParent    =207        # from enum CrRibbonCommand
	crRCmdCopyLocationInWorldCoordinateSystem=206        # from enum CrRibbonCommand
	crRCmdCopyMetas               =66         # from enum CrRibbonCommand
	crRCmdCopySelection           =65         # from enum CrRibbonCommand
	crRCmdCopyTransformationBetween2Actors=204        # from enum CrRibbonCommand
	crRCmdCopyTransformationInParent=203        # from enum CrRibbonCommand
	crRCmdCopyTransformationInWorldCoordinateSystem=202        # from enum CrRibbonCommand
	crRCmdCornerArrow             =123        # from enum CrRibbonCommand
	crRCmdCreate                  =184        # from enum CrRibbonCommand
	crRCmdCreateAndAdd            =185        # from enum CrRibbonCommand
	crRCmdCreateAssociativePathFromAnimation=137        # from enum CrRibbonCommand
	crRCmdCreateAssociativePathFromNeutral=139        # from enum CrRibbonCommand
	crRCmdCreateCameraKey         =315        # from enum CrRibbonCommand
	crRCmdCreateCuttingLineActor  =192        # from enum CrRibbonCommand
	crRCmdCreateDirectionalLight  =94         # from enum CrRibbonCommand
	crRCmdCreateFrom3Points       =186        # from enum CrRibbonCommand
	crRCmdCreateGrid              =106        # from enum CrRibbonCommand
	crRCmdCreateGridOnGeometry    =105        # from enum CrRibbonCommand
	crRCmdCreateMagneticLine      =110        # from enum CrRibbonCommand
	crRCmdCreateNonAssociativePathFromAnimation=138        # from enum CrRibbonCommand
	crRCmdCreateNonAssociativePathFromNeutral=140        # from enum CrRibbonCommand
	crRCmdCreatePathFromLineLinearSpeed=313        # from enum CrRibbonCommand
	crRCmdCreatePathFromLineLinearTime=314        # from enum CrRibbonCommand
	crRCmdCreatePositionalLight   =95         # from enum CrRibbonCommand
	crRCmdCreateSpotLight         =93         # from enum CrRibbonCommand
	crRCmdCurveLength             =154        # from enum CrRibbonCommand
	crRCmdCustom                  =69         # from enum CrRibbonCommand
	crRCmdCustom1                 =34         # from enum CrRibbonCommand
	crRCmdCustom2                 =35         # from enum CrRibbonCommand
	crRCmdCustom3                 =36         # from enum CrRibbonCommand
	crRCmdCustom4                 =37         # from enum CrRibbonCommand
	crRCmdCut                     =64         # from enum CrRibbonCommand
	crRCmdCylindrical             =215        # from enum CrRibbonCommand
	crRCmdDOF                     =240        # from enum CrRibbonCommand
	crRCmdDeleteAllKeys           =322        # from enum CrRibbonCommand
	crRCmdDeleteCameraKeys        =324        # from enum CrRibbonCommand
	crRCmdDeleteChildParent       =245        # from enum CrRibbonCommand
	crRCmdDeleteDiggerKeys        =325        # from enum CrRibbonCommand
	crRCmdDeleteKeysOfSelectedActors=323        # from enum CrRibbonCommand
	crRCmdDeleteUnusedKeys        =320        # from enum CrRibbonCommand
	crRCmdDepthOfFieldAutomatic   =87         # from enum CrRibbonCommand
	crRCmdDepthOfFieldOnOff       =86         # from enum CrRibbonCommand
	crRCmdDepthOfFieldSetPos      =89         # from enum CrRibbonCommand
	crRCmdDepthOfFieldVisible     =88         # from enum CrRibbonCommand
	crRCmdDetachCOI               =58         # from enum CrRibbonCommand
	crRCmdDetachCamera            =57         # from enum CrRibbonCommand
	crRCmdDetachCenter            =59         # from enum CrRibbonCommand
	crRCmdDetectClones            =263        # from enum CrRibbonCommand
	crRCmdDetectCurves            =221        # from enum CrRibbonCommand
	crRCmdDiameter                =158        # from enum CrRibbonCommand
	crRCmdDiameterFrom3Points     =159        # from enum CrRibbonCommand
	crRCmdDigger                  =24         # from enum CrRibbonCommand
	crRCmdDiggerPlayMode          =327        # from enum CrRibbonCommand
	crRCmdDimensions              =178        # from enum CrRibbonCommand
	crRCmdDimensionsInCoordinateSystem=179        # from enum CrRibbonCommand
	crRCmdDistanceAngleBetween2Lines=164        # from enum CrRibbonCommand
	crRCmdDistanceAngleBetween2Planes=165        # from enum CrRibbonCommand
	crRCmdDistanceBetween2Actors  =168        # from enum CrRibbonCommand
	crRCmdDistanceBetween2Axes    =166        # from enum CrRibbonCommand
	crRCmdDistanceBetween2Centers =167        # from enum CrRibbonCommand
	crRCmdDistanceBetween2Points  =163        # from enum CrRibbonCommand
	crRCmdDistanceBetweenLineAndPoint=170        # from enum CrRibbonCommand
	crRCmdDistanceBetweenPlaneAndPoint=169        # from enum CrRibbonCommand
	crRCmdExplode                 =255        # from enum CrRibbonCommand
	crRCmdExplodeByColor          =256        # from enum CrRibbonCommand
	crRCmdExport                  =310        # from enum CrRibbonCommand
	crRCmdFanMode                 =172        # from enum CrRibbonCommand
	crRCmdFileNewProject          =0          # from enum CrRibbonCommand
	crRCmdFileOpen                =1          # from enum CrRibbonCommand
	crRCmdFilters                 =275        # from enum CrRibbonCommand
	crRCmdFlat                    =78         # from enum CrRibbonCommand
	crRCmdFlatTechnical           =76         # from enum CrRibbonCommand
	crRCmdFlip                    =190        # from enum CrRibbonCommand
	crRCmdFlipFaces               =261        # from enum CrRibbonCommand
	crRCmdFlipNormals             =262        # from enum CrRibbonCommand
	crRCmdFreeDrag                =217        # from enum CrRibbonCommand
	crRCmdFrontBack               =27         # from enum CrRibbonCommand
	crRCmdFullScreen              =296        # from enum CrRibbonCommand
	crRCmdGDT                     =145        # from enum CrRibbonCommand
	crRCmdGenerateFreePath        =141        # from enum CrRibbonCommand
	crRCmdGettingStarted          =272        # from enum CrRibbonCommand
	crRCmdGhostSelection          =16         # from enum CrRibbonCommand
	crRCmdGlow                    =99         # from enum CrRibbonCommand
	crRCmdGround                  =100        # from enum CrRibbonCommand
	crRCmdHideSelection           =14         # from enum CrRibbonCommand
	crRCmdHighQuality             =103        # from enum CrRibbonCommand
	crRCmdHighResolutionImage     =280        # from enum CrRibbonCommand
	crRCmdImage                   =125        # from enum CrRibbonCommand
	crRCmdImage2D                 =128        # from enum CrRibbonCommand
	crRCmdImageLibrary            =329        # from enum CrRibbonCommand
	crRCmdInteractiveCollisionDetection=284        # from enum CrRibbonCommand
	crRCmdInvertSelection         =45         # from enum CrRibbonCommand
	crRCmdInvertVisibility        =17         # from enum CrRibbonCommand
	crRCmdInvertVisibilityWithGhost=18         # from enum CrRibbonCommand
	crRCmdJoinPolylines           =118        # from enum CrRibbonCommand
	crRCmdKeepAnimation           =254        # from enum CrRibbonCommand
	crRCmdKineFreeDrag            =239        # from enum CrRibbonCommand
	crRCmdKineMode                =238        # from enum CrRibbonCommand
	crRCmdKinematicRigidityMode   =247        # from enum CrRibbonCommand
	crRCmdLabel                   =143        # from enum CrRibbonCommand
	crRCmdLabelWithMultipleActors =144        # from enum CrRibbonCommand
	crRCmdLayers                  =301        # from enum CrRibbonCommand
	crRCmdLight                   =91         # from enum CrRibbonCommand
	crRCmdLineAxisToLineAxis      =196        # from enum CrRibbonCommand
	crRCmdLineAxisToPlane         =198        # from enum CrRibbonCommand
	crRCmdLinear                  =209        # from enum CrRibbonCommand
	crRCmdLinearLength            =153        # from enum CrRibbonCommand
	crRCmdLink                    =151        # from enum CrRibbonCommand
	crRCmdLinkChildToParent       =242        # from enum CrRibbonCommand
	crRCmdLinkSelectionToParent   =243        # from enum CrRibbonCommand
	crRCmdLoadRoot                =308        # from enum CrRibbonCommand
	crRCmdLoadonVisibility        =20         # from enum CrRibbonCommand
	crRCmdLocalTransform          =226        # from enum CrRibbonCommand
	crRCmdLockPosition            =200        # from enum CrRibbonCommand
	crRCmdMarkers                 =302        # from enum CrRibbonCommand
	crRCmdMenuAlign               =194        # from enum CrRibbonCommand
	crRCmdMenuAlignCamera         =25         # from enum CrRibbonCommand
	crRCmdMenuAlignPivot          =228        # from enum CrRibbonCommand
	crRCmdMenuArrow               =120        # from enum CrRibbonCommand
	crRCmdMenuAttachCamera        =53         # from enum CrRibbonCommand
	crRCmdMenuCoordinateSystem    =180        # from enum CrRibbonCommand
	crRCmdMenuCopyLocation        =205        # from enum CrRibbonCommand
	crRCmdMenuCopyTransformation  =201        # from enum CrRibbonCommand
	crRCmdMenuCreateFromLine      =312        # from enum CrRibbonCommand
	crRCmdMenuCuttingLine         =191        # from enum CrRibbonCommand
	crRCmdMenuCuttingPlanesCreate =183        # from enum CrRibbonCommand
	crRCmdMenuCylindrical         =214        # from enum CrRibbonCommand
	crRCmdMenuDelete              =321        # from enum CrRibbonCommand
	crRCmdMenuDiameter            =157        # from enum CrRibbonCommand
	crRCmdMenuDimensions          =177        # from enum CrRibbonCommand
	crRCmdMenuDistanceAngle       =162        # from enum CrRibbonCommand
	crRCmdMenuFile3DLive          =3          # from enum CrRibbonCommand
	crRCmdMenuFileUpdate          =2          # from enum CrRibbonCommand
	crRCmdMenuImage2D             =127        # from enum CrRibbonCommand
	crRCmdMenuLabel               =142        # from enum CrRibbonCommand
	crRCmdMenuLayout              =286        # from enum CrRibbonCommand
	crRCmdMenuLength              =152        # from enum CrRibbonCommand
	crRCmdMenuLighting            =92         # from enum CrRibbonCommand
	crRCmdMenuLightingMode        =90         # from enum CrRibbonCommand
	crRCmdMenuLinear              =208        # from enum CrRibbonCommand
	crRCmdMenuLink                =241        # from enum CrRibbonCommand
	crRCmdMenuMagnet              =109        # from enum CrRibbonCommand
	crRCmdMenuOutlineStyle        =82         # from enum CrRibbonCommand
	crRCmdMenuPath                =136        # from enum CrRibbonCommand
	crRCmdMenuPivotMode           =223        # from enum CrRibbonCommand
	crRCmdMenuPolyline            =113        # from enum CrRibbonCommand
	crRCmdMenuRenderingMode       =68         # from enum CrRibbonCommand
	crRCmdMenuRigidity            =246        # from enum CrRibbonCommand
	crRCmdMenuSelect              =38         # from enum CrRibbonCommand
	crRCmdMenuSetPivot            =233        # from enum CrRibbonCommand
	crRCmdMenuSpherical           =211        # from enum CrRibbonCommand
	crRCmdMenuTools               =104        # from enum CrRibbonCommand
	crRCmdMenuVisibility          =9          # from enum CrRibbonCommand
	crRCmdMerge                   =252        # from enum CrRibbonCommand
	crRCmdMergeByColor            =253        # from enum CrRibbonCommand
	crRCmdMinimalDistance         =176        # from enum CrRibbonCommand
	crRCmdMirror                  =102        # from enum CrRibbonCommand
	crRCmdModelBrowser            =273        # from enum CrRibbonCommand
	crRCmdMove                    =187        # from enum CrRibbonCommand
	crRCmdMultipleGizmos          =225        # from enum CrRibbonCommand
	crRCmdNoTransformations       =220        # from enum CrRibbonCommand
	crRCmdPPL                     =96         # from enum CrRibbonCommand
	crRCmdPartByPartCylindrical   =216        # from enum CrRibbonCommand
	crRCmdPartByPartLinear        =210        # from enum CrRibbonCommand
	crRCmdPartByPartSpherical     =213        # from enum CrRibbonCommand
	crRCmdPaste                   =67         # from enum CrRibbonCommand
	crRCmdPathPlanning            =285        # from enum CrRibbonCommand
	crRCmdPipeLength              =148        # from enum CrRibbonCommand
	crRCmdPivotTransform          =224        # from enum CrRibbonCommand
	crRCmdPlaneProjectedMeasurement=174        # from enum CrRibbonCommand
	crRCmdPlaneToPlane            =199        # from enum CrRibbonCommand
	crRCmdPlayWidgetComplete      =133        # from enum CrRibbonCommand
	crRCmdPlayWidgetEnd           =132        # from enum CrRibbonCommand
	crRCmdPlayWidgetNext          =129        # from enum CrRibbonCommand
	crRCmdPlayWidgetPrevious      =130        # from enum CrRibbonCommand
	crRCmdPlayWidgetStart         =131        # from enum CrRibbonCommand
	crRCmdPointToPlane            =197        # from enum CrRibbonCommand
	crRCmdPointToPoint            =195        # from enum CrRibbonCommand
	crRCmdPoints                  =80         # from enum CrRibbonCommand
	crRCmdPolyline                =114        # from enum CrRibbonCommand
	crRCmdPrimitivesCreateCube    =268        # from enum CrRibbonCommand
	crRCmdPrimitivesCreateCylinder=270        # from enum CrRibbonCommand
	crRCmdPrimitivesCreateDisc    =267        # from enum CrRibbonCommand
	crRCmdPrimitivesCreateLine    =265        # from enum CrRibbonCommand
	crRCmdPrimitivesCreatePoint   =264        # from enum CrRibbonCommand
	crRCmdPrimitivesCreateSphere  =269        # from enum CrRibbonCommand
	crRCmdPrimitivesCreateSquare  =266        # from enum CrRibbonCommand
	crRCmdProjectPlaneMeasurementOnAxis=175        # from enum CrRibbonCommand
	crRCmdProperties              =303        # from enum CrRibbonCommand
	crRCmdRadius                  =160        # from enum CrRibbonCommand
	crRCmdRadiusFrom3Points       =161        # from enum CrRibbonCommand
	crRCmdRedlining               =112        # from enum CrRibbonCommand
	crRCmdRefresh                 =306        # from enum CrRibbonCommand
	crRCmdRemoveCurrentLinksfromRigidSet=249        # from enum CrRibbonCommand
	crRCmdRemoveNode              =116        # from enum CrRibbonCommand
	crRCmdReplace                 =258        # from enum CrRibbonCommand
	crRCmdRestoreActiveViewVisibility=330        # from enum CrRibbonCommand
	crRCmdRestoreNeutralPosition  =222        # from enum CrRibbonCommand
	crRCmdRightLeft               =28         # from enum CrRibbonCommand
	crRCmdRotate                  =219        # from enum CrRibbonCommand
	crRCmdSaveCuttingLine         =193        # from enum CrRibbonCommand
	crRCmdSaveRoot                =309        # from enum CrRibbonCommand
	crRCmdScale                   =259        # from enum CrRibbonCommand
	crRCmdSecure3DBrush           =271        # from enum CrRibbonCommand
	crRCmdSelect                  =39         # from enum CrRibbonCommand
	crRCmdSelectAcrossSphere      =43         # from enum CrRibbonCommand
	crRCmdSelectAll               =44         # from enum CrRibbonCommand
	crRCmdSelectByColor           =40         # from enum CrRibbonCommand
	crRCmdSelectCollaboration     =50         # from enum CrRibbonCommand
	crRCmdSelectContactActors     =48         # from enum CrRibbonCommand
	crRCmdSelectGeometry          =49         # from enum CrRibbonCommand
	crRCmdSelectInsideSphere      =42         # from enum CrRibbonCommand
	crRCmdSelectInstances         =41         # from enum CrRibbonCommand
	crRCmdSelectUnsupportedStudio =47         # from enum CrRibbonCommand
	crRCmdSetPivotOnBoundingBoxCenter=235        # from enum CrRibbonCommand
	crRCmdSetPivotOnLineAxis      =237        # from enum CrRibbonCommand
	crRCmdSetPivotOnLocalCoordinateSystem=234        # from enum CrRibbonCommand
	crRCmdSetPivotOnPoint         =236        # from enum CrRibbonCommand
	crRCmdShadedIllustration      =72         # from enum CrRibbonCommand
	crRCmdShadedIllustrationWithColor=73         # from enum CrRibbonCommand
	crRCmdShadedIllustrationWithoutSilhouette=74         # from enum CrRibbonCommand
	crRCmdShadow                  =101        # from enum CrRibbonCommand
	crRCmdShadowMap               =97         # from enum CrRibbonCommand
	crRCmdShowAll                 =10         # from enum CrRibbonCommand
	crRCmdShowAllGeometry         =11         # from enum CrRibbonCommand
	crRCmdShowParentAxis          =227        # from enum CrRibbonCommand
	crRCmdShowSelection           =15         # from enum CrRibbonCommand
	crRCmdShowSelectionAndGhostOthers=13         # from enum CrRibbonCommand
	crRCmdShowSelectionAndHideOthers=12         # from enum CrRibbonCommand
	crRCmdSilhouette              =77         # from enum CrRibbonCommand
	crRCmdSimplification          =282        # from enum CrRibbonCommand
	crRCmdSmartOutlines           =85         # from enum CrRibbonCommand
	crRCmdSmooth                  =70         # from enum CrRibbonCommand
	crRCmdSmoothWithOutline       =71         # from enum CrRibbonCommand
	crRCmdSpherical               =212        # from enum CrRibbonCommand
	crRCmdSplitPolyline           =117        # from enum CrRibbonCommand
	crRCmdStyles                  =274        # from enum CrRibbonCommand
	crRCmdStylesApply             =6          # from enum CrRibbonCommand
	crRCmdStylesAutoSubscribe     =8          # from enum CrRibbonCommand
	crRCmdStylesCombo             =7          # from enum CrRibbonCommand
	crRCmdStylesSubscribe         =4          # from enum CrRibbonCommand
	crRCmdStylesUnsubscribe       =5          # from enum CrRibbonCommand
	crRCmdSymmetrize              =260        # from enum CrRibbonCommand
	crRCmdTechnical               =75         # from enum CrRibbonCommand
	crRCmdTechnicalIllustration   =279        # from enum CrRibbonCommand
	crRCmdText2D                  =135        # from enum CrRibbonCommand
	crRCmdTextures                =276        # from enum CrRibbonCommand
	crRCmdThickness               =156        # from enum CrRibbonCommand
	crRCmdTimeSettings            =328        # from enum CrRibbonCommand
	crRCmdTimeline                =304        # from enum CrRibbonCommand
	crRCmdTopBottom               =29         # from enum CrRibbonCommand
	crRCmdTransformGrid           =108        # from enum CrRibbonCommand
	crRCmdTransformGridOnGeometry =107        # from enum CrRibbonCommand
	crRCmdTranslate               =218        # from enum CrRibbonCommand
	crRCmdUnghost                 =19         # from enum CrRibbonCommand
	crRCmdUnionMode               =189        # from enum CrRibbonCommand
	crRCmdUnselectHiddenAndGhosted=46         # from enum CrRibbonCommand
	crRCmdUpdate                  =251        # from enum CrRibbonCommand
	crRCmdVectorImage             =126        # from enum CrRibbonCommand
	crRCmdVectorImage2D           =134        # from enum CrRibbonCommand
	crRCmdVectorView              =295        # from enum CrRibbonCommand
	crRCmdVideo                   =281        # from enum CrRibbonCommand
	crRCmdViews                   =277        # from enum CrRibbonCommand
	crRCmdViewsPane               =300        # from enum CrRibbonCommand
	crRCmdVolume                  =147        # from enum CrRibbonCommand
	crRCmdWireframe               =79         # from enum CrRibbonCommand
	crRCmdWorkshops               =305        # from enum CrRibbonCommand
	crRCmdZoomFitAll              =51         # from enum CrRibbonCommand
	crRCmdZoomSelection           =52         # from enum CrRibbonCommand
	crRGrpAlign                   =18         # from enum CrRibbonGroup
	crRGrpAlignPivots             =21         # from enum CrRibbonGroup
	crRGrpAnimClean               =34         # from enum CrRibbonGroup
	crRGrpAnimOther               =36         # from enum CrRibbonGroup
	crRGrpAnimPaths               =33         # from enum CrRibbonGroup
	crRGrpAnimPlay                =35         # from enum CrRibbonGroup
	crRGrpAnimScenarios           =32         # from enum CrRibbonGroup
	crRGrpAnnotations             =15         # from enum CrRibbonGroup
	crRGrpBars                    =31         # from enum CrRibbonGroup
	crRGrpCopyPaste               =5          # from enum CrRibbonGroup
	crRGrpCuttingPlanes           =17         # from enum CrRibbonGroup
	crRGrpDepthOfField            =7          # from enum CrRibbonGroup
	crRGrpDigger                  =2          # from enum CrRibbonGroup
	crRGrpExplode                 =19         # from enum CrRibbonGroup
	crRGrpGeometry                =23         # from enum CrRibbonGroup
	crRGrpGround                  =9          # from enum CrRibbonGroup
	crRGrpKinematic               =22         # from enum CrRibbonGroup
	crRGrpLighting                =8          # from enum CrRibbonGroup
	crRGrpMarkups                 =12         # from enum CrRibbonGroup
	crRGrpMeasurements            =16         # from enum CrRibbonGroup
	crRGrpMove                    =20         # from enum CrRibbonGroup
	crRGrpNavigate                =3          # from enum CrRibbonGroup
	crRGrpOnDemand                =10         # from enum CrRibbonGroup
	crRGrpPDMExplore              =41         # from enum CrRibbonGroup
	crRGrpPDMFile                 =37         # from enum CrRibbonGroup
	crRGrpPDMInfo                 =42         # from enum CrRibbonGroup
	crRGrpPDMProperties           =39         # from enum CrRibbonGroup
	crRGrpPDMStates               =40         # from enum CrRibbonGroup
	crRGrpPDMVersion              =38         # from enum CrRibbonGroup
	crRGrpPanels                  =13         # from enum CrRibbonGroup
	crRGrpPaths                   =14         # from enum CrRibbonGroup
	crRGrpPrimitives              =24         # from enum CrRibbonGroup
	crRGrpRendering               =6          # from enum CrRibbonGroup
	crRGrpSecure                  =25         # from enum CrRibbonGroup
	crRGrpShowHide                =4          # from enum CrRibbonGroup
	crRGrpStyles                  =0          # from enum CrRibbonGroup
	crRGrpTools                   =11         # from enum CrRibbonGroup
	crRGrpUser                    =43         # from enum CrRibbonGroup
	crRGrpViewport                =30         # from enum CrRibbonGroup
	crRGrpVisibility              =1          # from enum CrRibbonGroup
	crRGrpWsGeometry              =29         # from enum CrRibbonGroup
	crRGrpWsProperties            =27         # from enum CrRibbonGroup
	crRGrpWsPublishing            =28         # from enum CrRibbonGroup
	crRGrpWsStart                 =26         # from enum CrRibbonGroup
	crRTabAnimate                 =7          # from enum CrRibbonTab
	crRTabAuthor                  =2          # from enum CrRibbonTab
	crRTabDisplay                 =6          # from enum CrRibbonTab
	crRTabGeometry                =4          # from enum CrRibbonTab
	crRTabHome                    =0          # from enum CrRibbonTab
	crRTabPDM                     =9          # from enum CrRibbonTab
	crRTabRender                  =1          # from enum CrRibbonTab
	crRTabStyle                   =8          # from enum CrRibbonTab
	crRTabTransform               =3          # from enum CrRibbonTab
	crRTabWorkshops               =5          # from enum CrRibbonTab
	crTbAnnotation                =3          # from enum CrToolbars
	crTbGeometry                  =5          # from enum CrToolbars
	crTbKinematic                 =8          # from enum CrToolbars
	crTbPivotAlignment            =7          # from enum CrToolbars
	crTbRenderNavigateTransform   =2          # from enum CrToolbars
	crTbStandard                  =1          # from enum CrToolbars
	crTbStyles                    =6          # from enum CrToolbars
	crTbTexture                   =4          # from enum CrToolbars
	crTearoffToolbarsFirst        =9          # from enum CrToolbars
	crTearoffToolbarsLast         =24         # from enum CrToolbars
	crToAlignment                 =15         # from enum CrToolbars
	crToAnnotation                =18         # from enum CrToolbars
	crToCameraView                =13         # from enum CrToolbars
	crToCuttingPlane              =20         # from enum CrToolbars
	crToExplode                   =16         # from enum CrToolbars
	crToGrid                      =23         # from enum CrToolbars
	crToLayout                    =9          # from enum CrToolbars
	crToLighing                   =22         # from enum CrToolbars
	crToMarkup                    =17         # from enum CrToolbars
	crToMeasurement               =19         # from enum CrToolbars
	crToNavigate                  =12         # from enum CrToolbars
	crToPaths                     =21         # from enum CrToolbars
	crToRender                    =10         # from enum CrToolbars
	crToTransform                 =14         # from enum CrToolbars
	crToVisibility                =11         # from enum CrToolbars
	crToolbarsFirst               =1          # from enum CrToolbars
	crToolbarsLast                =9          # from enum CrToolbars
	crChecked                     =2          # from enum CrUpdateCommandState
	crEnabled                     =1          # from enum CrUpdateCommandState
	crRadio                       =4          # from enum CrUpdateCommandState
	crViHide                      =1          # from enum CrWindowVisibility
	crViMaximize                  =4          # from enum CrWindowVisibility
	crViMinimize                  =3          # from enum CrWindowVisibility
	crViRestore                   =7          # from enum CrWindowVisibility
	crViShow                      =2          # from enum CrWindowVisibility
	crViShowMaximized             =6          # from enum CrWindowVisibility
	crViShowMinimized             =5          # from enum CrWindowVisibility
	crWsBOM                       =10         # from enum CrWorkShop
	crWsClearanceChecking         =12         # from enum CrWorkShop
	crWsFilters                   =5          # from enum CrWorkShop
	crWsGettingStarted            =1          # from enum CrWorkShop
	crWsHighResolutionImage       =6          # from enum CrWorkShop
	crWsImageLibraryBrowser       =14         # from enum CrWorkShop
	crWsInteractiveCollisionDetection=13         # from enum CrWorkShop
	crWsModelBrowser              =2          # from enum CrWorkShop
	crWsNone                      =0          # from enum CrWorkShop
	crWsSimplification            =11         # from enum CrWorkShop
	crWsStyles                    =4          # from enum CrWorkShop
	crWsTechnicalIllustration     =7          # from enum CrWorkShop
	crWsTextureBrowser            =3          # from enum CrWorkShop
	crWsVideo                     =8          # from enum CrWorkShop
	crWsView                      =9          # from enum CrWorkShop

from win32com.client import DispatchBaseClass
class IApplication(DispatchBaseClass):
	CLSID = IID('{48CDCE04-6ECF-4D0A-BC25-F48FA1B1AC78}')
	coclass_clsid = IID('{3F41659D-BAB2-466C-A4D2-6E3CECFFB964}')

	def AboutBox(self):
		'method AboutBox'
		return self._oleobj_.InvokeTypes(20, LCID, 1, (24, 0), (),)

	def AddCustomRibbonItem(self, label=defaultNamedNotOptArg, userData=defaultNamedNotOptArg, smallImageIndex=defaultNamedNotOptArg, largeImageIndex=defaultNamedNotOptArg):
		'method AddCustomRibbonItem'
		return self._oleobj_.InvokeTypes(31, LCID, 1, (3, 0), ((8, 1), (8, 1), (2, 1), (2, 1)),label
			, userData, smallImageIndex, largeImageIndex)

	def AddCustomRibbonTab(self, position=defaultNamedNotOptArg, label=defaultNamedNotOptArg, smallImageList=defaultNamedNotOptArg, largeImageList=defaultNamedNotOptArg):
		'method AddCustomRibbonTab'
		return self._oleobj_.InvokeTypes(29, LCID, 1, (3, 0), ((2, 1), (8, 1), (8, 1), (8, 1)),position
			, label, smallImageList, largeImageList)

	def AddMenu(self, documentType=defaultNamedNotOptArg, position=defaultNamedNotOptArg, menuLabel=defaultNamedNotOptArg):
		'method AddMenu'
		return self._oleobj_.InvokeTypes(11, LCID, 1, (3, 0), ((2, 1), (2, 1), (8, 1)),documentType
			, position, menuLabel)

	def AddMenuItem(self, documentType=defaultNamedNotOptArg, menuItem=defaultNamedNotOptArg, userData=defaultNamedNotOptArg):
		'method AddMenuItem'
		return self._oleobj_.InvokeTypes(12, LCID, 1, (3, 0), ((2, 1), (8, 1), (8, 1)),documentType
			, menuItem, userData)

	# Result is of type IDocument
	def GetActiveDocument(self):
		'method GetActiveDocument'
		ret = self._oleobj_.InvokeTypes(5, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, 'GetActiveDocument', '{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')
		return ret

	def GetActiveWorkShop(self):
		'property GetActiveWorkShop'
		return self._oleobj_.InvokeTypes(18, LCID, 1, (3, 0), (),)

	def GetBarVisibility(self, barID=defaultNamedNotOptArg):
		'property GetBarVisibility'
		return self._oleobj_.InvokeTypes(16, LCID, 1, (11, 0), ((3, 1),),barID
			)

	# Result is of type IDocument
	def GetDocument(self, docID=defaultNamedNotOptArg):
		'method GetDocument'
		ret = self._oleobj_.InvokeTypes(4, LCID, 1, (9, 0), ((12, 1),),docID
			)
		if ret is not None:
			ret = Dispatch(ret, 'GetDocument', '{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')
		return ret

	def GetDocumentCount(self):
		'method GetDocumentCount'
		return self._oleobj_.InvokeTypes(3, LCID, 1, (3, 0), (),)

	def GetToolbarVisibility(self, toolbarID=defaultNamedNotOptArg):
		'property GetToolbarVisibility'
		return self._oleobj_.InvokeTypes(14, LCID, 1, (11, 0), ((3, 1),),toolbarID
			)

	def LaunchCommandMessage(self, commandCode=defaultNamedNotOptArg):
		'method LaunchCommandMessage'
		return self._oleobj_.InvokeTypes(21, LCID, 1, (11, 0), ((8, 1),),commandCode
			)

	def LoadDefaultDocumentProperties(self, strFilePath=defaultNamedNotOptArg):
		'method LoadDefaultDocumentProperties'
		return self._oleobj_.InvokeTypes(331, LCID, 1, (11, 0), ((8, 1),),strFilePath
			)

	# Result is of type IDocument
	def Open(self, filePath=defaultNamedNotOptArg):
		'method Open'
		ret = self._oleobj_.InvokeTypes(8, LCID, 1, (9, 0), ((8, 1),),filePath
			)
		if ret is not None:
			ret = Dispatch(ret, 'Open', '{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')
		return ret

	def PrintSetup(self):
		'method PrintSetup'
		return self._oleobj_.InvokeTypes(22, LCID, 1, (24, 0), (),)

	def Quit(self, Save=defaultNamedNotOptArg):
		'method Quit'
		return self._oleobj_.InvokeTypes(7, LCID, 1, (11, 0), ((11, 1),),Save
			)

	def RemoveCustomRibbonTab(self, CustomRibbonTabLabel=defaultNamedNotOptArg):
		'method RemoveCustomRibbonTab'
		return self._oleobj_.InvokeTypes(36, LCID, 1, (3, 0), ((8, 1),),CustomRibbonTabLabel
			)

	def RemoveRibbonCommand(self, ribbonCommand=defaultNamedNotOptArg):
		'method RemoveRibbonCommand'
		return self._oleobj_.InvokeTypes(27, LCID, 1, (3, 0), ((3, 1),),ribbonCommand
			)

	def RemoveRibbonGroup(self, ribbonGroup=defaultNamedNotOptArg):
		'method RemoveRibbonGroup'
		return self._oleobj_.InvokeTypes(26, LCID, 1, (3, 0), ((3, 1),),ribbonGroup
			)

	def RemoveRibbonTab(self, ribbonTab=defaultNamedNotOptArg):
		'method RemoveRibbonTab'
		return self._oleobj_.InvokeTypes(25, LCID, 1, (3, 0), ((3, 1),),ribbonTab
			)

	def RunRibbonCommand(self, ribbonCommand=defaultNamedNotOptArg):
		'method RunRibbonCommand'
		return self._oleobj_.InvokeTypes(28, LCID, 1, (3, 0), ((3, 1),),ribbonCommand
			)

	def SaveDefaultDocumentProperties(self, strFilePath=defaultNamedNotOptArg):
		'method SaveDefaultDocumentProperties'
		return self._oleobj_.InvokeTypes(332, LCID, 1, (11, 0), ((8, 1),),strFilePath
			)

	def SetActiveDocument(self, doc=defaultNamedNotOptArg):
		'method SetActiveDocument'
		return self._oleobj_.InvokeTypes(6, LCID, 1, (24, 0), ((9, 1),),doc
			)

	def SetActiveWorkShop(self, activeWorkShop=defaultNamedNotOptArg):
		'property SetActiveWorkShop'
		return self._oleobj_.InvokeTypes(19, LCID, 1, (11, 0), ((3, 1),),activeWorkShop
			)

	def SetAdvancedLicense(self, vendorCode=defaultNamedNotOptArg):
		'method SetAdvancedLicense'
		return self._oleobj_.InvokeTypes(33, LCID, 1, (11, 0), ((8, 1),),vendorCode
			)

	def SetBarVisibility(self, barID=defaultNamedNotOptArg, Visible=defaultNamedNotOptArg):
		'property SetBarVisibility'
		return self._oleobj_.InvokeTypes(17, LCID, 1, (11, 0), ((3, 1), (11, 1)),barID
			, Visible)

	def SetFocus(self):
		'method SetFocus'
		return self._oleobj_.InvokeTypes(9, LCID, 1, (24, 0), (),)

	def SetInformationTemplate(self, strHTMLTemplatePath=defaultNamedNotOptArg):
		'method SetInformationTemplate'
		return self._oleobj_.InvokeTypes(34, LCID, 1, (24, 0), ((8, 1),),strHTMLTemplatePath
			)

	def SetInformationTemplateContent(self, strInformationTemplateContent=defaultNamedNotOptArg):
		'method SetInformationTemplateContent'
		return self._oleobj_.InvokeTypes(35, LCID, 1, (24, 0), ((8, 1),),strInformationTemplateContent
			)

	def SetToolbarVisibility(self, toolbarID=defaultNamedNotOptArg, Visible=defaultNamedNotOptArg):
		'property SetToolbarVisibility'
		return self._oleobj_.InvokeTypes(15, LCID, 1, (11, 0), ((3, 1), (11, 1)),toolbarID
			, Visible)

	def Show(self, showCommand=defaultNamedNotOptArg):
		'method Show'
		return self._oleobj_.InvokeTypes(23, LCID, 1, (24, 0), ((3, 1),),showCommand
			)

	def UpdateCustomRibbon(self):
		'method UpdateCustomRibbon'
		return self._oleobj_.InvokeTypes(32, LCID, 1, (24, 0), (),)

	def UpdateMenus(self):
		'method UpdateMenus'
		return self._oleobj_.InvokeTypes(13, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
		"MainWindowHandle": (2, 2, (3, 0), (), "MainWindowHandle", None),
		"TooltipVisibility": (305, 2, (11, 0), (), "TooltipVisibility", None),
		"Version": (10, 2, (8, 0), (), "Version", None),
		"Visible": (1, 2, (11, 0), (), "Visible", None),
	}
	_prop_map_put_ = {
		"TooltipVisibility": ((305, LCID, 4, 0),()),
		"Visible": ((1, LCID, 4, 0),()),
	}
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

class IApplicationEvents:
	CLSID = CLSID_Sink = IID('{21268FAA-FC76-44CA-84AB-217807808491}')
	coclass_clsid = IID('{3F41659D-BAB2-466C-A4D2-6E3CECFFB964}')
	_public_methods_ = [] # For COM Server support
	_dispid_to_func_ = {
		1610612736 : "OnQueryInterface",
		1610612737 : "OnAddRef",
		1610612738 : "OnRelease",
		1610678272 : "OnGetTypeInfoCount",
		1610678273 : "OnGetTypeInfo",
		1610678274 : "OnGetIDsOfNames",
		1610678275 : "OnInvoke",
		        1 : "OnCommand",
		        2 : "OnUpdateCommandUI",
		        3 : "OnOpenDocument",
		        4 : "OnActivateDocument",
		        5 : "OnCloseApplication",
		}

	def __init__(self, oobj = None):
		if oobj is None:
			self._olecp = None
		else:
			import win32com.server.util
			from win32com.server.policy import EventHandlerPolicy
			cpc=oobj._oleobj_.QueryInterface(pythoncom.IID_IConnectionPointContainer)
			cp=cpc.FindConnectionPoint(self.CLSID_Sink)
			cookie=cp.Advise(win32com.server.util.wrap(self, usePolicy=EventHandlerPolicy))
			self._olecp,self._olecp_cookie = cp,cookie
	def __del__(self):
		try:
			self.close()
		except pythoncom.com_error:
			pass
	def close(self):
		if self._olecp is not None:
			cp,cookie,self._olecp,self._olecp_cookie = self._olecp,self._olecp_cookie,None,None
			cp.Unadvise(cookie)
	def _query_interface_(self, iid):
		import win32com.server.util
		if iid==self.CLSID_Sink: return win32com.server.util.wrap(self)

	# Event Handlers
	# If you create handlers, they should have the following prototypes:
#	def OnQueryInterface(self, riid=defaultNamedNotOptArg, ppvObj=pythoncom.Missing):
#	def OnAddRef(self):
#	def OnRelease(self):
#	def OnGetTypeInfoCount(self, pctinfo=pythoncom.Missing):
#	def OnGetTypeInfo(self, itinfo=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, pptinfo=pythoncom.Missing):
#	def OnGetIDsOfNames(self, riid=defaultNamedNotOptArg, rgszNames=defaultNamedNotOptArg, cNames=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg
#			, rgdispid=pythoncom.Missing):
#	def OnInvoke(self, dispidMember=defaultNamedNotOptArg, riid=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, wFlags=defaultNamedNotOptArg
#			, pdispparams=defaultNamedNotOptArg, pvarResult=pythoncom.Missing, pexcepinfo=pythoncom.Missing, puArgErr=pythoncom.Missing):
#	def OnCommand(self, sender=defaultNamedNotOptArg, menuLabel=defaultNamedNotOptArg, userData=defaultNamedNotOptArg):
#	def OnUpdateCommandUI(self, sender=defaultNamedNotOptArg, menuLabel=defaultNamedNotOptArg, state=pythoncom.Missing, userData=defaultNamedNotOptArg):
#	def OnOpenDocument(self, sender=defaultNamedNotOptArg, Document=defaultNamedNotOptArg):
#	def OnActivateDocument(self, sender=defaultNamedNotOptArg, Document=defaultNamedNotOptArg, activated=defaultNamedNotOptArg):
#	def OnCloseApplication(self, sender=defaultNamedNotOptArg):


class IDocument(DispatchBaseClass):
	CLSID = IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')
	coclass_clsid = IID('{E5B985AA-AC9B-44E3-A3FF-44D9E7B90714}')

	def AbortQueries(self):
		'method AbortQueries'
		return self._oleobj_.InvokeTypes(234, LCID, 1, (11, 0), (),)

	def AddLink(self, Selection=defaultNamedNotOptArg, link=defaultNamedNotOptArg, autoLink=defaultNamedNotOptArg):
		'method AddLink'
		return self._oleobj_.InvokeTypes(134, LCID, 1, (11, 0), ((8, 1), (8, 1), (11, 1)),Selection
			, link, autoLink)

	def AddMetaProperties(self, Selection=defaultNamedNotOptArg, propertySet=defaultNamedNotOptArg):
		'method AddMetaProperties'
		return self._oleobj_.InvokeTypes(208, LCID, 1, (11, 0), ((8, 1), (8, 1)),Selection
			, propertySet)

	def AddMetaPropertyDefinition(self, name=defaultNamedNotOptArg, metaType=defaultNamedNotOptArg, metaLabel=defaultNamedNotOptArg, stateFlags=defaultNamedNotOptArg
			, unit=defaultNamedNotOptArg, bounded=defaultNamedNotOptArg, minValue=defaultNamedNotOptArg, maxValue=defaultNamedNotOptArg):
		'method AddMetaPropertyDefinition'
		return self._oleobj_.InvokeTypes(206, LCID, 1, (11, 0), ((8, 1), (3, 1), (8, 1), (3, 1), (3, 1), (11, 1), (4, 1), (4, 1)),name
			, metaType, metaLabel, stateFlags, unit, bounded
			, minValue, maxValue)

	def AddProductIn(self, FileName=defaultNamedNotOptArg, netGuid=defaultNamedNotOptArg):
		'method AddProductIn'
		return self._oleobj_.InvokeTypes(196, LCID, 1, (11, 0), ((8, 1), (8, 1)),FileName
			, netGuid)

	def AddProductIn2(self, FileName=defaultNamedNotOptArg, netGuidToAttach=defaultNamedNotOptArg, instanceGuid=defaultNamedNotOptArg):
		'method AddProductIn2'
		return self._oleobj_.InvokeTypes(232, LCID, 1, (11, 0), ((8, 1), (8, 1), (8, 1)),FileName
			, netGuidToAttach, instanceGuid)

	def AddSignature(self, LogoPath=defaultNamedNotOptArg, Agreement=defaultNamedNotOptArg, eSignatureMode=defaultNamedNotOptArg):
		'method AddSignature'
		return self._oleobj_.InvokeTypes(325, LCID, 1, (11, 0), ((8, 1), (8, 1), (3, 1)),LogoPath
			, Agreement, eSignatureMode)

	def Antialiasing(self):
		'method Antialiasing'
		return self._oleobj_.InvokeTypes(106, LCID, 1, (24, 0), (),)

	def ChangeParent(self, Selection=defaultNamedNotOptArg, selectionNewParent=defaultNamedNotOptArg):
		'méthode ChangeParent'
		return self._oleobj_.InvokeTypes(254, LCID, 1, (11, 0), ((8, 1), (8, 1)),Selection
			, selectionNewParent)

	def Close(self, Save=defaultNamedNotOptArg):
		'method Close'
		return self._oleobj_.InvokeTypes(1003, LCID, 1, (11, 0), ((11, 1),),Save
			)

	def ConfigureColumnsForAssemblyPane(self, strFilePath=defaultNamedNotOptArg):
		'method ConfigureColumnsForAssemblyPane'
		return self._oleobj_.InvokeTypes(335, LCID, 1, (11, 0), ((8, 1),),strFilePath
			)

	def CreateAssyGroup(self, Selection=defaultNamedNotOptArg):
		'méthode CreateAssyGroup'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(252, LCID, 1, (8, 0), ((8, 1),),Selection
			)

	def CreateBBox(self, ptAx=defaultNamedNotOptArg, ptAy=defaultNamedNotOptArg, ptAz=defaultNamedNotOptArg, ptBx=defaultNamedNotOptArg
			, ptBy=defaultNamedNotOptArg, ptBz=defaultNamedNotOptArg):
		'méthode Create BBox'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(250, LCID, 1, (8, 0), ((4, 1), (4, 1), (4, 1), (4, 1), (4, 1), (4, 1)),ptAx
			, ptAy, ptAz, ptBx, ptBy, ptBz
			)

	def CreateCallout(self, Selection=defaultNamedNotOptArg, eGenMode=defaultNamedNotOptArg, eGenAttach=defaultNamedNotOptArg):
		'method CreateCallout'
		return self._oleobj_.InvokeTypes(314, LCID, 1, (11, 0), ((8, 1), (3, 1), (3, 1)),Selection
			, eGenMode, eGenAttach)

	def CreateCameraView(self, name=defaultNamedNotOptArg, position=defaultNamedNotOptArg):
		'method CreateCameraView'
		return self._oleobj_.InvokeTypes(290, LCID, 1, (11, 0), ((8, 1), (3, 1)),name
			, position)

	def CreateDummyObject(self, Selection=defaultNamedNotOptArg):
		'méthode CreateDummyObject'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(253, LCID, 1, (8, 0), ((8, 1),),Selection
			)

	def CreateSelectionSet(self, SelectionSetName=defaultNamedNotOptArg, Selection=defaultNamedNotOptArg):
		'method CreateSelectionSet'
		return self._oleobj_.InvokeTypes(320, LCID, 1, (11, 0), ((8, 1), (8, 1)),SelectionSetName
			, Selection)

	def CreateView(self, name=defaultNamedNotOptArg, position=defaultNamedNotOptArg):
		'method CreateView'
		return self._oleobj_.InvokeTypes(289, LCID, 1, (11, 0), ((8, 1), (3, 1)),name
			, position)

	def DeleteGenericAnimation(self, TagID=defaultNamedNotOptArg):
		'method DeleteGenericAnimation'
		return self._oleobj_.InvokeTypes(326, LCID, 1, (11, 0), ((3, 1),),TagID
			)

	def DeleteView(self, view=defaultNamedNotOptArg):
		'method DeleteView'
		return self._oleobj_.InvokeTypes(291, LCID, 1, (11, 0), ((8, 1),),view
			)

	def ExportAllPubs(self, FileName=defaultNamedNotOptArg, visibility=defaultNamedNotOptArg):
		'method ExportAllPubs'
		return self._oleobj_.InvokeTypes(227, LCID, 1, (11, 0), ((8, 1), (3, 1)),FileName
			, visibility)

	def ExportAllViews(self, FileName=defaultNamedNotOptArg):
		'method ExportAllViews'
		return self._oleobj_.InvokeTypes(224, LCID, 1, (11, 0), ((8, 1),),FileName
			)

	def ExportResources(self, folderPath=defaultNamedNotOptArg):
		'méthode ExportResources'
		return self._oleobj_.InvokeTypes(256, LCID, 1, (11, 0), ((8, 1),),folderPath
			)

	def ExportViews(self, viewsSelection=defaultNamedNotOptArg, FileName=defaultNamedNotOptArg):
		'method ExportViews'
		return self._oleobj_.InvokeTypes(225, LCID, 1, (11, 0), ((8, 1), (8, 1)),viewsSelection
			, FileName)

	def ExportViewsPubs(self, viewsIDSelection=defaultNamedNotOptArg, FileName=defaultNamedNotOptArg):
		'method ExportViewsPubs'
		return self._oleobj_.InvokeTypes(226, LCID, 1, (11, 0), ((8, 1), (8, 1)),viewsIDSelection
			, FileName)

	def FileBrowseDialog(self):
		'method FileBrowseDialog'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(310, LCID, 1, (8, 0), (),)

	def FilterViews(self, Selection=defaultNamedNotOptArg):
		'method FilterViews'
		return self._oleobj_.InvokeTypes(283, LCID, 1, (11, 0), ((8, 1),),Selection
			)

	def GetActorsClassName(self, Selection=defaultNamedNotOptArg):
		'method GetActorsClassName'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(274, LCID, 1, (8, 0), ((8, 1),),Selection
			)

	def GetActorsProperties(self, Selection=defaultNamedNotOptArg):
		'method GetActorsProperties'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(285, LCID, 1, (8, 0), ((8, 1),),Selection
			)

	def GetAllActors(self, visibility=defaultNamedNotOptArg):
		'method GetAllActors'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(158, LCID, 1, (8, 0), ((3, 1),),visibility
			)

	def GetAllActorsByType(self, className=defaultNamedNotOptArg):
		'method GetAllActorsByType'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(223, LCID, 1, (8, 0), ((8, 1),),className
			)

	def GetAllActorsOnFilter(self, propertySet=defaultNamedNotOptArg, matchValue=defaultNamedNotOptArg):
		'method GetAllActorsOnFilter'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(162, LCID, 1, (8, 0), ((8, 1), (11, 1)),propertySet
			, matchValue)

	def GetAllCollaborations(self):
		'method GetAllCollaborations'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(161, LCID, 1, (8, 0), (),)

	def GetAllConfigurations(self):
		'method GetAllConfigurations'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(163, LCID, 1, (8, 0), (),)

	def GetAllGenericAnimation(self):
		'method GetAllGenericAnimation'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(308, LCID, 1, (8, 0), (),)

	def GetAllHotspots(self):
		'method GetAllHotspots'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(316, LCID, 1, (8, 0), (),)

	def GetAllLayers(self):
		'method GetAllLayers'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(168, LCID, 1, (8, 0), (),)

	def GetAllLinks(self):
		'method GetAllLinks'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(159, LCID, 1, (8, 0), (),)

	def GetAllLockedViews(self):
		'methode GetAllLockedViews'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(323, LCID, 1, (8, 0), (),)

	def GetAllMarkers(self):
		'method GetAllMarkers'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(164, LCID, 1, (8, 0), (),)

	def GetAllMarkersDetails(self):
		'method GetAllMarkersDetails'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(306, LCID, 1, (8, 0), (),)

	def GetAllMetaPropertyDefinitions(self):
		'method GetAllMetaPropertyDefinitions'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(266, LCID, 1, (8, 0), (),)

	def GetAllSelectionSets(self):
		'method GetAllSelectionSets'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(312, LCID, 1, (8, 0), (),)

	def GetAllTexts(self):
		'method GetAllTexts'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(160, LCID, 1, (8, 0), (),)

	def GetAllViews(self):
		'méthode GetAllViews'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(246, LCID, 1, (8, 0), (),)

	def GetAssyChild(self, Selection=defaultNamedNotOptArg):
		'method GetAssyChild'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(184, LCID, 1, (8, 0), ((8, 1),),Selection
			)

	def GetAssyParents(self, Selection=defaultNamedNotOptArg):
		'method GetAssyParents'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(185, LCID, 1, (8, 0), ((8, 1),),Selection
			)

	def GetAssyRoot(self):
		'method GetAssyRoot'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(183, LCID, 1, (8, 0), (),)

	def GetCADCaptures(self):
		'méthode GetCADCaptures'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(244, LCID, 1, (8, 0), (),)

	def GetCADViews(self):
		'méthode GetCADViews'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(245, LCID, 1, (8, 0), (),)

	def GetCamera(self):
		'method GetCamera'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(180, LCID, 1, (8, 0), (),)

	def GetChildren(self, Selection=defaultNamedNotOptArg, mask=defaultNamedNotOptArg):
		'méthode GetChildren'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(247, LCID, 1, (8, 0), ((8, 1), (3, 1)),Selection
			, mask)

	def GetChildrenView(self, viewName=defaultNamedNotOptArg):
		'méthode GetChildrenView'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(249, LCID, 1, (8, 0), ((8, 1),),viewName
			)

	def GetCrossHighlightGeometry(self, strGUID=defaultNamedNotOptArg, bNormals=defaultNamedNotOptArg):
		'method GetCrossHighlightGeometry'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(287, LCID, 1, (8, 0), ((8, 1), (11, 1)),strGUID
			, bNormals)

	def GetDigger(self):
		'method GetDigger'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(192, LCID, 1, (8, 0), (),)

	def GetDiggerRenderer(self):
		'method GetDiggerRenderer'
		return self._oleobj_.InvokeTypes(219, LCID, 1, (3, 0), (),)

	def GetGenericAnimationProperty(self, nIndex=defaultNamedNotOptArg):
		'method GetGenericAnimationProperty'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(307, LCID, 1, (8, 0), ((3, 1),),nIndex
			)

	def GetGround(self):
		'method GetGround'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(265, LCID, 1, (8, 0), (),)

	def GetInvertedSelection(self, Selection=defaultNamedNotOptArg):
		'method GetInvertedSelection'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(220, LCID, 1, (8, 0), ((8, 1),),Selection
			)

	def GetLinks(self, Selection=defaultNamedNotOptArg):
		'method GetLinks'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(156, LCID, 1, (8, 0), ((8, 1),),Selection
			)

	def GetLocationInParent(self, strSelection=defaultNamedNotOptArg):
		'method GetLocationInParent'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(258, LCID, 1, (8, 0), ((8, 1),),strSelection
			)

	def GetLocationInPart(self, strSelectionSource=defaultNamedNotOptArg, strSelectionDest=defaultNamedNotOptArg):
		'method GetLocationInPart'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(259, LCID, 1, (8, 0), ((8, 1), (8, 1)),strSelectionSource
			, strSelectionDest)

	def GetLocationInWorld(self, strSelection=defaultNamedNotOptArg):
		'method GetLocationInWorld'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(257, LCID, 1, (8, 0), ((8, 1),),strSelection
			)

	def GetMeasurementPersistency(self):
		'method GetMeasurementPersistency'
		return self._oleobj_.InvokeTypes(281, LCID, 1, (11, 0), (),)

	def GetOnFilter(self, propertiesSet=defaultNamedNotOptArg, matchValue=defaultNamedNotOptArg):
		'méthode GetOnFilter'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(251, LCID, 1, (8, 0), ((8, 1), (11, 1)),propertiesSet
			, matchValue)

	def GetOptionsPro(self):
		'method GetOptionsPro'
		return self._oleobj_.InvokeTypes(143, LCID, 1, (3, 0), (),)

	def GetProperties(self, Selection=defaultNamedNotOptArg, intersect=defaultNamedNotOptArg, matchValue=defaultNamedNotOptArg):
		'method GetProperties'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(167, LCID, 1, (8, 0), ((8, 1), (11, 1), (11, 1)),Selection
			, intersect, matchValue)

	def GetPropertyValue(self, Selection=defaultNamedNotOptArg, propertyName=defaultNamedNotOptArg):
		'method GetPropertyValue'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(217, LCID, 1, (8, 0), ((8, 1), (8, 1)),Selection
			, propertyName)

	def GetViewport(self):
		'method GetViewport'
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(190, LCID, 1, (8, 0), (),)

	def GetVisibility(self, Selection=defaultNamedNotOptArg):
		'method GetVisibility'
		return self._oleobj_.InvokeTypes(166, LCID, 1, (3, 0), ((8, 1),),Selection
			)

	def GoToConfiguration(self, configurationName=defaultNamedNotOptArg):
		'method GoToConfiguration'
		return self._oleobj_.InvokeTypes(129, LCID, 1, (11, 0), ((8, 1),),configurationName
			)

	def GoToLayer(self, layerName=defaultNamedNotOptArg):
		'method GoToLayer'
		return self._oleobj_.InvokeTypes(132, LCID, 1, (11, 0), ((8, 1),),layerName
			)

	def GoToMarker(self, markerName=defaultNamedNotOptArg):
		'method GoToMarker'
		return self._oleobj_.InvokeTypes(130, LCID, 1, (11, 0), ((8, 1),),markerName
			)

	def GoToView(self, viewName=defaultNamedNotOptArg):
		'méthode GoToView'
		return self._oleobj_.InvokeTypes(248, LCID, 1, (11, 0), ((8, 1),),viewName
			)

	def HasChilds(self, Selection=defaultNamedNotOptArg):
		'method HasChilds'
		return self._oleobj_.InvokeTypes(187, LCID, 1, (11, 0), ((8, 1),),Selection
			)

	def ImportPubs(self, FileName=defaultNamedNotOptArg):
		'method ImportPubs'
		return self._oleobj_.InvokeTypes(229, LCID, 1, (11, 0), ((8, 1),),FileName
			)

	def ImportViewsData(self, FileName=defaultNamedNotOptArg, viewIDInsertAfter=defaultNamedNotOptArg):
		'method ImportViewsData'
		return self._oleobj_.InvokeTypes(228, LCID, 1, (11, 0), ((8, 1), (8, 1)),FileName
			, viewIDInsertAfter)

	def ImportViewsData2(self, FileName=defaultNamedNotOptArg, viewIDInsertAfter=defaultNamedNotOptArg, forceResolution=defaultNamedNotOptArg, allInstances=defaultNamedNotOptArg):
		'method ImportViewsData2'
		return self._oleobj_.InvokeTypes(275, LCID, 1, (11, 0), ((8, 1), (8, 1), (11, 1), (11, 1)),FileName
			, viewIDInsertAfter, forceResolution, allInstances)

	def InvertVisibility(self, withGhost=defaultNamedNotOptArg):
		'method InvertVisibility'
		return self._oleobj_.InvokeTypes(140, LCID, 1, (24, 0), ((11, 1),),withGhost
			)

	def IsCollaboration(self, Selection=defaultNamedNotOptArg):
		'méthode IsCollaboration'
		return self._oleobj_.InvokeTypes(243, LCID, 1, (11, 0), ((8, 1),),Selection
			)

	def IsPureAssy(self, Selection=defaultNamedNotOptArg):
		'method IsPureAssy'
		return self._oleobj_.InvokeTypes(186, LCID, 1, (11, 0), ((8, 1),),Selection
			)

	def LaunchQueryBBox(self):
		'method LaunchQueryBBox'
		return self._oleobj_.InvokeTypes(222, LCID, 1, (11, 0), (),)

	def LinearExplode(self, Selection=defaultNamedNotOptArg, vectorX=defaultNamedNotOptArg, vectorY=defaultNamedNotOptArg, vectorZ=defaultNamedNotOptArg
			, fDistance=defaultNamedNotOptArg):
		'method CreateLinearExplode'
		return self._oleobj_.InvokeTypes(315, LCID, 1, (11, 0), ((8, 1), (3, 1), (3, 1), (3, 1), (4, 1)),Selection
			, vectorX, vectorY, vectorZ, fDistance)

	def LoadDocumentProperties(self, strFilePath=defaultNamedNotOptArg):
		'method LoadDocumentProperties'
		return self._oleobj_.InvokeTypes(329, LCID, 1, (11, 0), ((8, 1),),strFilePath
			)

	def LoadGeometries(self, netGuid=defaultNamedNotOptArg):
		'method LoadGeometries'
		return self._oleobj_.InvokeTypes(238, LCID, 1, (11, 0), ((8, 1),),netGuid
			)

	def LockView(self, view=defaultNamedNotOptArg):
		'method LockView'
		return self._oleobj_.InvokeTypes(321, LCID, 1, (11, 0), ((8, 1),),view
			)

	def Merge(self, FileName=defaultNamedNotOptArg):
		'method Merge'
		return self._oleobj_.InvokeTypes(142, LCID, 1, (24, 0), ((8, 1),),FileName
			)

	def Pause(self):
		'method Pause'
		return self._oleobj_.InvokeTypes(169, LCID, 1, (24, 0), (),)

	def Play(self):
		'method Play'
		return self._oleobj_.InvokeTypes(101, LCID, 1, (24, 0), (),)

	def PlayMarkerSequence(self, markerName=defaultNamedNotOptArg):
		'method PlayMarkerSequence'
		return self._oleobj_.InvokeTypes(131, LCID, 1, (11, 0), ((8, 1),),markerName
			)

	def Print(self):
		'method Print'
		return self._oleobj_.InvokeTypes(203, LCID, 1, (24, 0), (),)

	def PrintPreview(self):
		'method PrintPreview'
		return self._oleobj_.InvokeTypes(202, LCID, 1, (24, 0), (),)

	def Redo(self):
		'method Redo'
		return self._oleobj_.InvokeTypes(112, LCID, 1, (24, 0), (),)

	def RefreshScene(self, refreshType=defaultNamedNotOptArg):
		'method RefreshScene'
		return self._oleobj_.InvokeTypes(136, LCID, 1, (24, 0), ((3, 1),),refreshType
			)

	def Remove(self, Selection=defaultNamedNotOptArg):
		'method Remove'
		return self._oleobj_.InvokeTypes(152, LCID, 1, (11, 0), ((8, 1),),Selection
			)

	def RemoveMetaProperties(self, propertyMap=defaultNamedNotOptArg):
		'method RemoveMetaProperties'
		return self._oleobj_.InvokeTypes(209, LCID, 1, (11, 0), ((8, 1),),propertyMap
			)

	def RemoveMetaPropertyDefinition(self, name=defaultNamedNotOptArg):
		'method RemoveMetaPropertyDefinition'
		return self._oleobj_.InvokeTypes(207, LCID, 1, (11, 0), ((8, 1),),name
			)

	def RemoveProductNodes(self, Selection=defaultNamedNotOptArg):
		'method RemoveProductNodes'
		return self._oleobj_.InvokeTypes(233, LCID, 1, (11, 0), ((8, 1),),Selection
			)

	def RemoveViews(self, viewsIDSelection=defaultNamedNotOptArg):
		'method RemoveViews'
		return self._oleobj_.InvokeTypes(230, LCID, 1, (11, 0), ((8, 1),),viewsIDSelection
			)

	def RenameView(self, strViewName=defaultNamedNotOptArg, strNewViewName=defaultNamedNotOptArg):
		'method RenameView'
		return self._oleobj_.InvokeTypes(300, LCID, 1, (11, 0), ((8, 1), (8, 1)),strViewName
			, strNewViewName)

	def ReplaceGeometry(self, FileName=defaultNamedNotOptArg, strNetGUIDSource=defaultNamedNotOptArg, selectionDest=defaultNamedNotOptArg):
		'méthode ReplaceGeometry'
		return self._oleobj_.InvokeTypes(255, LCID, 1, (11, 0), ((8, 1), (8, 1), (8, 1)),FileName
			, strNetGUIDSource, selectionDest)

	def ResetZoomFitAll(self, viewList=defaultNamedNotOptArg):
		'method ResetZoomFitAll'
		return self._oleobj_.InvokeTypes(311, LCID, 1, (24, 0), ((8, 1),),viewList
			)

	def RestoreNeutralProperties(self, Selection=defaultNamedNotOptArg, propertySet=defaultNamedNotOptArg):
		'method RestoreNeutralProperties'
		return self._oleobj_.InvokeTypes(216, LCID, 1, (11, 0), ((8, 1), (8, 1)),Selection
			, propertySet)

	def RotateInLCS(self, Selection=defaultNamedNotOptArg, originX=defaultNamedNotOptArg, originY=defaultNamedNotOptArg, originZ=defaultNamedNotOptArg
			, vectorX=defaultNamedNotOptArg, vectorY=defaultNamedNotOptArg, vectorZ=defaultNamedNotOptArg, angle=defaultNamedNotOptArg):
		'method RotateInLCS'
		return self._oleobj_.InvokeTypes(177, LCID, 1, (11, 0), ((8, 1), (4, 1), (4, 1), (4, 1), (4, 1), (4, 1), (4, 1), (4, 1)),Selection
			, originX, originY, originZ, vectorX, vectorY
			, vectorZ, angle)

	def RotateInParent(self, Selection=defaultNamedNotOptArg, originX=defaultNamedNotOptArg, originY=defaultNamedNotOptArg, originZ=defaultNamedNotOptArg
			, vectorX=defaultNamedNotOptArg, vectorY=defaultNamedNotOptArg, vectorZ=defaultNamedNotOptArg, angle=defaultNamedNotOptArg):
		'method RotateInParent'
		return self._oleobj_.InvokeTypes(173, LCID, 1, (11, 0), ((8, 1), (4, 1), (4, 1), (4, 1), (4, 1), (4, 1), (4, 1), (4, 1)),Selection
			, originX, originY, originZ, vectorX, vectorY
			, vectorZ, angle)

	def RotateInWCS(self, Selection=defaultNamedNotOptArg, originX=defaultNamedNotOptArg, originY=defaultNamedNotOptArg, originZ=defaultNamedNotOptArg
			, vectorX=defaultNamedNotOptArg, vectorY=defaultNamedNotOptArg, vectorZ=defaultNamedNotOptArg, angle=defaultNamedNotOptArg):
		'method RotateInWCS'
		return self._oleobj_.InvokeTypes(172, LCID, 1, (11, 0), ((8, 1), (4, 1), (4, 1), (4, 1), (4, 1), (4, 1), (4, 1), (4, 1)),Selection
			, originX, originY, originZ, vectorX, vectorY
			, vectorZ, angle)

	def Save(self):
		'method Save'
		return self._oleobj_.InvokeTypes(236, LCID, 1, (11, 0), (),)

	def SaveAs(self, FileName=defaultNamedNotOptArg):
		'method SaveAs'
		return self._oleobj_.InvokeTypes(237, LCID, 1, (11, 0), ((8, 1),),FileName
			)

	def SaveBOMTree(self, strFilePath=defaultNamedNotOptArg):
		'method SaveBOMTree'
		return self._oleobj_.InvokeTypes(334, LCID, 1, (11, 0), ((8, 1),),strFilePath
			)

	def SaveDocumentAs(self, FileName=defaultNamedNotOptArg):
		'method SaveDocumentAs'
		return self._oleobj_.InvokeTypes(324, LCID, 1, (11, 0), ((8, 1),),FileName
			)

	def SaveDocumentProperties(self, strFilePath=defaultNamedNotOptArg):
		'method SaveDocumentProperties'
		return self._oleobj_.InvokeTypes(330, LCID, 1, (11, 0), ((8, 1),),strFilePath
			)

	def SaveImage(self, pathName=defaultNamedNotOptArg):
		'method SaveImage'
		return self._oleobj_.InvokeTypes(150, LCID, 1, (11, 0), ((8, 1),),pathName
			)

	def SaveImageFullViewport(self, pathName=defaultNamedNotOptArg):
		'method SaveImageFullViewport'
		return self._oleobj_.InvokeTypes(268, LCID, 1, (11, 0), ((8, 1),),pathName
			)

	def SaveImageUsingProfile(self, pathName=defaultNamedNotOptArg, profileFullPath=defaultNamedNotOptArg):
		'method SaveImageUsingProfile'
		return self._oleobj_.InvokeTypes(319, LCID, 1, (11, 0), ((8, 1), (8, 1)),pathName
			, profileFullPath)

	def SaveSMG(self, filePath=defaultNamedNotOptArg, password=defaultNamedNotOptArg):
		'method SaveSMG'
		return self._oleobj_.InvokeTypes(148, LCID, 1, (11, 0), ((8, 1), (8, 1)),filePath
			, password)

	def ScrollViewToTop(self, strName=defaultNamedNotOptArg):
		'method ScrollViewToTop'
		return self._oleobj_.InvokeTypes(271, LCID, 1, (24, 0), ((8, 1),),strName
			)

	def SelectOnFilter(self, propertySet=defaultNamedNotOptArg, matchValue=defaultNamedNotOptArg):
		'method SelectOnFilter'
		return self._oleobj_.InvokeTypes(154, LCID, 1, (11, 0), ((8, 1), (11, 1)),propertySet
			, matchValue)

	def SetCamera(self, defCamera=defaultNamedNotOptArg):
		'method SetCamera'
		return self._oleobj_.InvokeTypes(179, LCID, 1, (11, 0), ((8, 1),),defCamera
			)

	def SetCameraKey(self):
		'method SetCameraKey'
		return self._oleobj_.InvokeTypes(328, LCID, 1, (24, 0), (),)

	def SetColor(self, layerColor=defaultNamedNotOptArg, red=defaultNamedNotOptArg, green=defaultNamedNotOptArg, blue=defaultNamedNotOptArg):
		'method SetColor'
		return self._oleobj_.InvokeTypes(221, LCID, 1, (11, 0), ((3, 1), (3, 1), (3, 1), (3, 1)),layerColor
			, red, green, blue)

	def SetDigger(self, defDigger=defaultNamedNotOptArg):
		'method SetDigger'
		return self._oleobj_.InvokeTypes(191, LCID, 1, (11, 0), ((8, 1),),defDigger
			)

	def SetDiggerRenderer(self, digRenderer=defaultNamedNotOptArg):
		'method SetDiggerRenderer'
		return self._oleobj_.InvokeTypes(218, LCID, 1, (11, 0), ((3, 1),),digRenderer
			)

	def SetGround(self, pGround=defaultNamedNotOptArg):
		'method SetGround'
		return self._oleobj_.InvokeTypes(264, LCID, 1, (24, 0), ((16392, 1),),pGround
			)

	def SetMeasurementPersistency(self, result=defaultNamedNotOptArg):
		'method SetMeasurementPersistency'
		return self._oleobj_.InvokeTypes(282, LCID, 1, (24, 0), ((11, 1),),result
			)

	def SetMouseNavigationMode(self, navigationType=defaultNamedNotOptArg):
		'method SetMouseNavigationMode'
		return self._oleobj_.InvokeTypes(267, LCID, 1, (11, 0), ((3, 1),),navigationType
			)

	def SetNeutralProperties(self, Selection=defaultNamedNotOptArg, propertySet=defaultNamedNotOptArg):
		'method SetNeutralProperties'
		return self._oleobj_.InvokeTypes(215, LCID, 1, (11, 0), ((8, 1), (8, 1)),Selection
			, propertySet)

	def SetPropertyMap(self, propertyMap=defaultNamedNotOptArg):
		'method SetPropertyMap'
		return self._oleobj_.InvokeTypes(135, LCID, 1, (11, 0), ((8, 1),),propertyMap
			)

	def SetPropertySet(self, Selection=defaultNamedNotOptArg, propertySet=defaultNamedNotOptArg):
		'method SetPropertySet'
		return self._oleobj_.InvokeTypes(137, LCID, 1, (11, 0), ((8, 1), (8, 1)),Selection
			, propertySet)

	def SetViewPosition(self, position=defaultNamedNotOptArg, Selection=defaultNamedNotOptArg):
		'method SetViewPosition'
		return self._oleobj_.InvokeTypes(288, LCID, 1, (24, 0), ((3, 1), (8, 1)),position
			, Selection)

	def SetViewport(self, defViewport=defaultNamedNotOptArg):
		'method SetViewport'
		return self._oleobj_.InvokeTypes(189, LCID, 1, (11, 0), ((8, 1),),defViewport
			)

	def SetVisibility(self, Selection=defaultNamedNotOptArg, visibility=defaultNamedNotOptArg, invVisibility=defaultNamedNotOptArg):
		'method SetVisibility'
		return self._oleobj_.InvokeTypes(139, LCID, 1, (24, 0), ((8, 1), (3, 1), (3, 1)),Selection
			, visibility, invVisibility)

	def ShowNoShowCollaboration(self):
		'method ShowNoShowCollaboration'
		return self._oleobj_.InvokeTypes(141, LCID, 1, (24, 0), (),)

	def Stop(self):
		'method Stop'
		return self._oleobj_.InvokeTypes(102, LCID, 1, (24, 0), (),)

	def TranslateInLCS(self, strSelection=defaultNamedNotOptArg, translationX=defaultNamedNotOptArg, translationY=defaultNamedNotOptArg, translationZ=defaultNamedNotOptArg):
		'method TranslateInLCS'
		return self._oleobj_.InvokeTypes(176, LCID, 1, (11, 0), ((8, 1), (4, 1), (4, 1), (4, 1)),strSelection
			, translationX, translationY, translationZ)

	def TranslateInParent(self, strSelection=defaultNamedNotOptArg, translationX=defaultNamedNotOptArg, translationY=defaultNamedNotOptArg, translationZ=defaultNamedNotOptArg):
		'method TranslateInParent'
		return self._oleobj_.InvokeTypes(175, LCID, 1, (11, 0), ((8, 1), (4, 1), (4, 1), (4, 1)),strSelection
			, translationX, translationY, translationZ)

	def TranslateInWCS(self, Selection=defaultNamedNotOptArg, translationX=defaultNamedNotOptArg, translationY=defaultNamedNotOptArg, translationZ=defaultNamedNotOptArg):
		'method TranslateInWCS'
		return self._oleobj_.InvokeTypes(174, LCID, 1, (11, 0), ((8, 1), (4, 1), (4, 1), (4, 1)),Selection
			, translationX, translationY, translationZ)

	def Undo(self):
		'method Undo'
		return self._oleobj_.InvokeTypes(111, LCID, 1, (24, 0), (),)

	def UnloadGeometries(self, netGuid=defaultNamedNotOptArg):
		'method UnloadGeometries'
		return self._oleobj_.InvokeTypes(239, LCID, 1, (11, 0), ((8, 1),),netGuid
			)

	def UnlockView(self, view=defaultNamedNotOptArg):
		'method UnlockView'
		return self._oleobj_.InvokeTypes(322, LCID, 1, (11, 0), ((8, 1),),view
			)

	def Update(self, FileName=defaultNamedNotOptArg):
		'method Update'
		return self._oleobj_.InvokeTypes(241, LCID, 1, (11, 0), ((8, 1),),FileName
			)

	def UpdateGenericAnimation(self, TagID=defaultNamedNotOptArg, SelectedActors=defaultNamedNotOptArg):
		'method UpdateGenericAnimation'
		return self._oleobj_.InvokeTypes(327, LCID, 1, (11, 0), ((3, 1), (8, 1)),TagID
			, SelectedActors)

	def UpdateView(self, viewID=defaultNamedNotOptArg):
		'method UpdateView'
		return self._oleobj_.InvokeTypes(231, LCID, 1, (11, 0), ((8, 1),),viewID
			)

	def ZoomFitAll(self):
		'method ZoomFitAll'
		return self._oleobj_.InvokeTypes(109, LCID, 1, (24, 0), (),)

	def ZoomSelection(self):
		'method ZoomSelection'
		return self._oleobj_.InvokeTypes(110, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
		"AllNameOccurences": (309, 2, (11, 0), (), "AllNameOccurences", None),
		"AssySelectionMode": (182, 2, (11, 0), (), "AssySelectionMode", None),
		"AutoKeyMode": (242, 2, (11, 0), (), "AutoKeyMode", None),
		"CameraInertia": (115, 2, (11, 0), (), "CameraInertia", None),
		"CameraPlayMode": (107, 2, (11, 0), (), "CameraPlayMode", None),
		"CameraSmoothTime": (170, 2, (2, 0), (), "CameraSmoothTime", None),
		"DropInViewport": (299, 2, (11, 0), (), "DropInViewport", None),
		"EventsMask": (235, 2, (3, 0), (), "EventsMask", None),
		"FileName": (113, 2, (8, 0), (), "FileName", None),
		"GroundGrid": (116, 2, (11, 0), (), "GroundGrid", None),
		"HighlightInvisibleActorsinViewport": (333, 2, (11, 0), (), "HighlightInvisibleActorsinViewport", None),
		"HighlightedObject": (138, 2, (8, 0), (), "HighlightedObject", None),
		"KinematicActivated": (193, 2, (11, 0), (), "KinematicActivated", None),
		"NoHighlight": (181, 2, (11, 0), (), "NoHighlight", None),
		"PMICrossHLOpacity": (298, 2, (4, 0), (), "PMICrossHLOpacity", None),
		"PMICrossHLStayOnTop": (297, 2, (11, 0), (), "PMICrossHLStayOnTop", None),
		"Pos": (153, 2, (3, 0), (), "Pos", None),
		"RenderGroundShadow": (104, 2, (11, 0), (), "RenderGroundShadow", None),
		"RenderMirror": (240, 2, (11, 0), (), "RenderMirror", None),
		"RenderMode": (105, 2, (3, 0), (), "RenderMode", None),
		"Selection": (133, 2, (8, 0), (), "Selection", None),
		"StartupView": (205, 2, (8, 0), (), "StartupView", None),
		"UseGUID": (171, 2, (11, 0), (), "UseGUID", None),
		"ViewNameTemplate": (301, 2, (8, 0), (), "ViewNameTemplate", None),
	}
	_prop_map_put_ = {
		"AllNameOccurences": ((309, LCID, 4, 0),()),
		"AssySelectionMode": ((182, LCID, 4, 0),()),
		"AutoKeyMode": ((242, LCID, 4, 0),()),
		"CameraInertia": ((115, LCID, 4, 0),()),
		"CameraPlayMode": ((107, LCID, 4, 0),()),
		"CameraSmoothTime": ((170, LCID, 4, 0),()),
		"DropInViewport": ((299, LCID, 4, 0),()),
		"EventsMask": ((235, LCID, 4, 0),()),
		"GroundGrid": ((116, LCID, 4, 0),()),
		"HighlightInvisibleActorsinViewport": ((333, LCID, 4, 0),()),
		"HighlightedObject": ((138, LCID, 4, 0),()),
		"KinematicActivated": ((193, LCID, 4, 0),()),
		"NoHighlight": ((181, LCID, 4, 0),()),
		"PMICrossHLOpacity": ((298, LCID, 4, 0),()),
		"PMICrossHLStayOnTop": ((297, LCID, 4, 0),()),
		"Pos": ((153, LCID, 4, 0),()),
		"RenderGroundShadow": ((104, LCID, 4, 0),()),
		"RenderMirror": ((240, LCID, 4, 0),()),
		"RenderMode": ((105, LCID, 4, 0),()),
		"Selection": ((133, LCID, 4, 0),()),
		"StartupView": ((205, LCID, 4, 0),()),
		"UseGUID": ((171, LCID, 4, 0),()),
		"ViewNameTemplate": ((301, LCID, 4, 0),()),
	}
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

class IDocumentEvents:
	CLSID = CLSID_Sink = IID('{DE79BC3C-1CE6-4DB7-816A-6307D8ADF418}')
	coclass_clsid = IID('{E5B985AA-AC9B-44E3-A3FF-44D9E7B90714}')
	_public_methods_ = [] # For COM Server support
	_dispid_to_func_ = {
		1610612736 : "OnQueryInterface",
		1610612737 : "OnAddRef",
		1610612738 : "OnRelease",
		1610678272 : "OnGetTypeInfoCount",
		1610678273 : "OnGetTypeInfo",
		1610678274 : "OnGetIDsOfNames",
		1610678275 : "OnInvoke",
		      130 : "OnAddMarker",
		      131 : "OnDeleteMarker",
		      101 : "OnChangeSelection",
		      102 : "Onlink",
		      103 : "OnGetHighlight",
		      104 : "OnLooseHighlight",
		      105 : "OnChangeProperties",
		      106 : "OnChangePos",
		      107 : "OnChangeConfiguration",
		      108 : "OnChangeLayer",
		      109 : "OnAddActors",
		      110 : "OnRemoveActors",
		      111 : "OnPlayChangePos",
		      112 : "OnBeginLoadModel",
		      113 : "OnEndLoadModel",
		      114 : "OnChangeVisibility",
		      115 : "OnClick",
		      116 : "OnDblClick",
		      117 : "OnPlayBegin",
		      118 : "OnPlayEnd",
		      119 : "OnChangeMarker",
		      120 : "OnPlayChangeMarker",
		      121 : "OnDisplayLessThan32Bits",
		      122 : "OnChangeViews",
		      123 : "OnCloseDocument",
		      125 : "OnAddView",
		      126 : "OnRemoveViews",
		      127 : "OnMoveViews",
		      128 : "OnSave",
		      129 : "OnPasteViews",
		      132 : "OnAddGenericAnimation",
		      135 : "OnClickGenericAnimation",
		      133 : "OnMoveGenericAnimation",
		      134 : "OnDeleteGenericAnimation",
		      137 : "OnSelectFromSelectionSets",
		      136 : "OnChangeAllViews",
		      138 : "OnUndo",
		      139 : "OnRedo",
		      142 : "OnCreateSelectionSet",
		      141 : "OnDeleteSelectionSet",
		      140 : "OnUpdateSelectionSet",
		      145 : "OnCreateHotspot",
		      144 : "OnDeleteHotspot",
		      143 : "OnUpdateHotspot",
		}

	def __init__(self, oobj = None):
		if oobj is None:
			self._olecp = None
		else:
			import win32com.server.util
			from win32com.server.policy import EventHandlerPolicy
			cpc=oobj._oleobj_.QueryInterface(pythoncom.IID_IConnectionPointContainer)
			cp=cpc.FindConnectionPoint(self.CLSID_Sink)
			cookie=cp.Advise(win32com.server.util.wrap(self, usePolicy=EventHandlerPolicy))
			self._olecp,self._olecp_cookie = cp,cookie
	def __del__(self):
		try:
			self.close()
		except pythoncom.com_error:
			pass
	def close(self):
		if self._olecp is not None:
			cp,cookie,self._olecp,self._olecp_cookie = self._olecp,self._olecp_cookie,None,None
			cp.Unadvise(cookie)
	def _query_interface_(self, iid):
		import win32com.server.util
		if iid==self.CLSID_Sink: return win32com.server.util.wrap(self)

	# Event Handlers
	# If you create handlers, they should have the following prototypes:
#	def OnQueryInterface(self, riid=defaultNamedNotOptArg, ppvObj=pythoncom.Missing):
#	def OnAddRef(self):
#	def OnRelease(self):
#	def OnGetTypeInfoCount(self, pctinfo=pythoncom.Missing):
#	def OnGetTypeInfo(self, itinfo=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, pptinfo=pythoncom.Missing):
#	def OnGetIDsOfNames(self, riid=defaultNamedNotOptArg, rgszNames=defaultNamedNotOptArg, cNames=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg
#			, rgdispid=pythoncom.Missing):
#	def OnInvoke(self, dispidMember=defaultNamedNotOptArg, riid=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, wFlags=defaultNamedNotOptArg
#			, pdispparams=defaultNamedNotOptArg, pvarResult=pythoncom.Missing, pexcepinfo=pythoncom.Missing, puArgErr=pythoncom.Missing):
#	def OnAddMarker(self, sender=defaultNamedNotOptArg, label=defaultNamedNotOptArg, frame=defaultNamedNotOptArg):
#	def OnDeleteMarker(self, sender=defaultNamedNotOptArg, frame=defaultNamedNotOptArg):
#	def OnChangeSelection(self, sender=defaultNamedNotOptArg, Selection=defaultNamedNotOptArg):
#	def Onlink(self, sender=defaultNamedNotOptArg, link=defaultNamedNotOptArg, autoLink=defaultNamedNotOptArg):
#	def OnGetHighlight(self, sender=defaultNamedNotOptArg, name=defaultNamedNotOptArg):
#	def OnLooseHighlight(self, sender=defaultNamedNotOptArg, name=defaultNamedNotOptArg):
#	def OnChangeProperties(self, sender=defaultNamedNotOptArg, propertyMap=defaultNamedNotOptArg):
#	def OnChangePos(self, sender=defaultNamedNotOptArg, Pos=defaultNamedNotOptArg):
#	def OnChangeConfiguration(self, sender=defaultNamedNotOptArg, configurationName=defaultNamedNotOptArg):
#	def OnChangeLayer(self, sender=defaultNamedNotOptArg, layerName=defaultNamedNotOptArg):
#	def OnAddActors(self, sender=defaultNamedNotOptArg, Selection=defaultNamedNotOptArg):
#	def OnRemoveActors(self, sender=defaultNamedNotOptArg, Selection=defaultNamedNotOptArg):
#	def OnPlayChangePos(self, sender=defaultNamedNotOptArg, Pos=defaultNamedNotOptArg):
#	def OnBeginLoadModel(self, sender=defaultNamedNotOptArg):
#	def OnEndLoadModel(self, sender=defaultNamedNotOptArg, Selection=defaultNamedNotOptArg):
#	def OnChangeVisibility(self, sender=defaultNamedNotOptArg, Selection=defaultNamedNotOptArg):
#	def OnClick(self, sender=defaultNamedNotOptArg, name=defaultNamedNotOptArg):
#	def OnDblClick(self, sender=defaultNamedNotOptArg, name=defaultNamedNotOptArg):
#	def OnPlayBegin(self, sender=defaultNamedNotOptArg, Pos=defaultNamedNotOptArg):
#	def OnPlayEnd(self, sender=defaultNamedNotOptArg, Pos=defaultNamedNotOptArg):
#	def OnChangeMarker(self, sender=defaultNamedNotOptArg, label=defaultNamedNotOptArg):
#	def OnPlayChangeMarker(self, sender=defaultNamedNotOptArg, label=defaultNamedNotOptArg):
#	def OnDisplayLessThan32Bits(self, sender=defaultNamedNotOptArg):
#	def OnChangeViews(self, sender=defaultNamedNotOptArg, viewID=defaultNamedNotOptArg):
#	def OnCloseDocument(self, sender=defaultNamedNotOptArg):
#	def OnAddView(self, sender=defaultNamedNotOptArg, Selection=defaultNamedNotOptArg):
#	def OnRemoveViews(self, sender=defaultNamedNotOptArg, Selection=defaultNamedNotOptArg):
#	def OnMoveViews(self, sender=defaultNamedNotOptArg, Selection=defaultNamedNotOptArg, DropIndex=defaultNamedNotOptArg):
#	def OnSave(self, sender=defaultNamedNotOptArg, status=defaultNamedNotOptArg):
#	def OnPasteViews(self, sender=defaultNamedNotOptArg, selectionCopy=defaultNamedNotOptArg, selectionPaste=defaultNamedNotOptArg):
#	def OnAddGenericAnimation(self, sender=defaultNamedNotOptArg, stepPropMap=defaultNamedNotOptArg):
#	def OnClickGenericAnimation(self, sender=defaultNamedNotOptArg, stepPropMap=defaultNamedNotOptArg):
#	def OnMoveGenericAnimation(self, sender=defaultNamedNotOptArg, stepPropMap=defaultNamedNotOptArg):
#	def OnDeleteGenericAnimation(self, sender=defaultNamedNotOptArg, stepPropMap=defaultNamedNotOptArg):
#	def OnSelectFromSelectionSets(self, sender=defaultNamedNotOptArg, SelectSelection=defaultNamedNotOptArg):
#	def OnChangeAllViews(self, sender=defaultNamedNotOptArg, strViews=defaultNamedNotOptArg):
#	def OnUndo(self, sender=defaultNamedNotOptArg, commandName=defaultNamedNotOptArg, Selection=defaultNamedNotOptArg):
#	def OnRedo(self, sender=defaultNamedNotOptArg, commandName=defaultNamedNotOptArg, Selection=defaultNamedNotOptArg):
#	def OnCreateSelectionSet(self, sender=defaultNamedNotOptArg, Selection=defaultNamedNotOptArg):
#	def OnDeleteSelectionSet(self, sender=defaultNamedNotOptArg, Selection=defaultNamedNotOptArg):
#	def OnUpdateSelectionSet(self, sender=defaultNamedNotOptArg, Selection=defaultNamedNotOptArg):
#	def OnCreateHotspot(self, sender=defaultNamedNotOptArg, Selection=defaultNamedNotOptArg):
#	def OnDeleteHotspot(self, sender=defaultNamedNotOptArg, Selection=defaultNamedNotOptArg):
#	def OnUpdateHotspot(self, sender=defaultNamedNotOptArg, Selection=defaultNamedNotOptArg):


from win32com.client import CoClassBaseClass
# This CoClass is known by the name 'DSComposer.Application'
class Application(CoClassBaseClass): # A CoClass
	CLSID = IID('{3F41659D-BAB2-466C-A4D2-6E3CECFFB964}')
	coclass_sources = [
		IApplicationEvents,
	]
	default_source = IApplicationEvents
	coclass_interfaces = [
		IApplication,
	]
	default_interface = IApplication

# This CoClass is known by the name 'SOLIDWORKSComposer.Document'
class Document(CoClassBaseClass): # A CoClass
	CLSID = IID('{E5B985AA-AC9B-44E3-A3FF-44D9E7B90714}')
	coclass_sources = [
		IDocumentEvents,
	]
	default_source = IDocumentEvents
	coclass_interfaces = [
		IDocument,
	]
	default_interface = IDocument

IApplication_vtables_dispatch_ = 1
IApplication_vtables_ = [
	(( 'Visible' , 'Visible' , ), 1, (1, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'Visible' , 'Visible' , ), 1, (1, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'MainWindowHandle' , 'hWnd' , ), 2, (2, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'GetDocumentCount' , 'count' , ), 3, (3, (), [ (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'GetDocument' , 'docID' , 'doc' , ), 4, (4, (), [ (12, 1, None, None) , 
			 (16393, 10, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , ], 1 , 1 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'GetActiveDocument' , 'doc' , ), 5, (5, (), [ (16393, 10, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , ], 1 , 1 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'SaveDefaultDocumentProperties' , 'strFilePath' , 'result' , ), 332, (332, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'LoadDefaultDocumentProperties' , 'strFilePath' , 'result' , ), 331, (331, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'SetActiveDocument' , 'doc' , ), 6, (6, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , ], 1 , 1 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'Quit' , 'Save' , 'status' , ), 7, (7, (), [ (11, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'Open' , 'filePath' , 'doc' , ), 8, (8, (), [ (8, 1, None, None) , 
			 (16393, 10, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , ], 1 , 1 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( 'SetFocus' , ), 9, (9, (), [ ], 1 , 1 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'Version' , 'Version' , ), 10, (10, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'AddMenu' , 'documentType' , 'position' , 'menuLabel' , 'result' , 
			 ), 11, (11, (), [ (2, 1, None, None) , (2, 1, None, None) , (8, 1, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'AddMenuItem' , 'documentType' , 'menuItem' , 'userData' , 'result' , 
			 ), 12, (12, (), [ (2, 1, None, None) , (8, 1, None, None) , (8, 1, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 168 , (3, 0, None, None) , 0 , )),
	(( 'UpdateMenus' , ), 13, (13, (), [ ], 1 , 1 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( 'GetToolbarVisibility' , 'toolbarID' , 'Visible' , ), 14, (14, (), [ (3, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'SetToolbarVisibility' , 'toolbarID' , 'Visible' , 'result' , ), 15, (15, (), [ 
			 (3, 1, None, None) , (11, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'GetBarVisibility' , 'barID' , 'Visible' , ), 16, (16, (), [ (3, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'SetBarVisibility' , 'barID' , 'Visible' , 'result' , ), 17, (17, (), [ 
			 (3, 1, None, None) , (11, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'GetActiveWorkShop' , 'activeWorkShop' , ), 18, (18, (), [ (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( 'SetActiveWorkShop' , 'activeWorkShop' , 'result' , ), 19, (19, (), [ (3, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 224 , (3, 0, None, None) , 0 , )),
	(( 'AboutBox' , ), 20, (20, (), [ ], 1 , 1 , 4 , 0 , 232 , (3, 0, None, None) , 0 , )),
	(( 'LaunchCommandMessage' , 'commandCode' , 'result' , ), 21, (21, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 240 , (3, 0, None, None) , 64 , )),
	(( 'PrintSetup' , ), 22, (22, (), [ ], 1 , 1 , 4 , 0 , 248 , (3, 0, None, None) , 0 , )),
	(( 'Show' , 'showCommand' , ), 23, (23, (), [ (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 256 , (3, 0, None, None) , 0 , )),
	(( 'RemoveRibbonTab' , 'ribbonTab' , 'result' , ), 25, (25, (), [ (3, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 264 , (3, 0, None, None) , 0 , )),
	(( 'RemoveRibbonGroup' , 'ribbonGroup' , 'result' , ), 26, (26, (), [ (3, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 272 , (3, 0, None, None) , 0 , )),
	(( 'RemoveRibbonCommand' , 'ribbonCommand' , 'result' , ), 27, (27, (), [ (3, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 280 , (3, 0, None, None) , 0 , )),
	(( 'RunRibbonCommand' , 'ribbonCommand' , 'result' , ), 28, (28, (), [ (3, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 288 , (3, 0, None, None) , 0 , )),
	(( 'AddCustomRibbonTab' , 'position' , 'label' , 'smallImageList' , 'largeImageList' , 
			 'result' , ), 29, (29, (), [ (2, 1, None, None) , (8, 1, None, None) , (8, 1, None, None) , 
			 (8, 1, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 296 , (3, 0, None, None) , 0 , )),
	(( 'AddCustomRibbonItem' , 'label' , 'userData' , 'smallImageIndex' , 'largeImageIndex' , 
			 'result' , ), 31, (31, (), [ (8, 1, None, None) , (8, 1, None, None) , (2, 1, None, None) , 
			 (2, 1, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 304 , (3, 0, None, None) , 0 , )),
	(( 'UpdateCustomRibbon' , ), 32, (32, (), [ ], 1 , 1 , 4 , 0 , 312 , (3, 0, None, None) , 0 , )),
	(( 'SetAdvancedLicense' , 'vendorCode' , 'result' , ), 33, (33, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 320 , (3, 0, None, None) , 0 , )),
	(( 'SetInformationTemplate' , 'strHTMLTemplatePath' , ), 34, (34, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 328 , (3, 0, None, None) , 0 , )),
	(( 'SetInformationTemplateContent' , 'strInformationTemplateContent' , ), 35, (35, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 336 , (3, 0, None, None) , 0 , )),
	(( 'TooltipVisibility' , 'TooltipVisibility' , ), 305, (305, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 344 , (3, 0, None, None) , 0 , )),
	(( 'TooltipVisibility' , 'TooltipVisibility' , ), 305, (305, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 352 , (3, 0, None, None) , 0 , )),
	(( 'RemoveCustomRibbonTab' , 'CustomRibbonTabLabel' , 'result' , ), 36, (36, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 360 , (3, 0, None, None) , 0 , )),
]

IApplicationEvents_vtables_dispatch_ = 1
IApplicationEvents_vtables_ = [
	(( 'Command' , 'sender' , 'menuLabel' , 'userData' , 'status' , 
			 ), 1, (1, (), [ (9, 1, None, "IID('{48CDCE04-6ECF-4D0A-BC25-F48FA1B1AC78}')") , (8, 1, None, None) , (8, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'UpdateCommandUI' , 'sender' , 'menuLabel' , 'state' , 'userData' , 
			 'status' , ), 2, (2, (), [ (9, 1, None, "IID('{48CDCE04-6ECF-4D0A-BC25-F48FA1B1AC78}')") , (8, 1, None, None) , (16387, 2, None, None) , 
			 (8, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'OpenDocument' , 'sender' , 'Document' , ), 3, (3, (), [ (9, 1, None, "IID('{48CDCE04-6ECF-4D0A-BC25-F48FA1B1AC78}')") , 
			 (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , ], 1 , 1 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'ActivateDocument' , 'sender' , 'Document' , 'activated' , ), 4, (4, (), [ 
			 (9, 1, None, "IID('{48CDCE04-6ECF-4D0A-BC25-F48FA1B1AC78}')") , (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , (11, 1, None, None) , ], 1 , 1 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'CloseApplication' , 'sender' , ), 5, (5, (), [ (9, 1, None, "IID('{48CDCE04-6ECF-4D0A-BC25-F48FA1B1AC78}')") , ], 1 , 1 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
]

IDocument_vtables_dispatch_ = 1
IDocument_vtables_ = [
	(( 'FileName' , 'FileName' , ), 113, (113, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'RenderGroundShadow' , 'RenderGroundShadow' , ), 104, (104, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'RenderGroundShadow' , 'RenderGroundShadow' , ), 104, (104, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'RenderMode' , 'RenderMode' , ), 105, (105, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'RenderMode' , 'RenderMode' , ), 105, (105, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'CameraPlayMode' , 'CameraPlayMode' , ), 107, (107, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'CameraPlayMode' , 'CameraPlayMode' , ), 107, (107, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'Selection' , 'Selection' , ), 133, (133, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'Selection' , 'Selection' , ), 133, (133, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'CameraInertia' , 'CameraInertia' , ), 115, (115, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'CameraInertia' , 'CameraInertia' , ), 115, (115, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( 'GroundGrid' , 'GroundGrid' , ), 116, (116, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'GroundGrid' , 'GroundGrid' , ), 116, (116, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'HighlightedObject' , 'HighlightedObject' , ), 138, (138, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'HighlightedObject' , 'HighlightedObject' , ), 138, (138, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 168 , (3, 0, None, None) , 0 , )),
	(( 'Pos' , 'Pos' , ), 153, (153, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( 'Pos' , 'Pos' , ), 153, (153, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'CameraSmoothTime' , 'CameraSmoothTime' , ), 170, (170, (), [ (16386, 10, None, None) , ], 1 , 2 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'CameraSmoothTime' , 'CameraSmoothTime' , ), 170, (170, (), [ (2, 1, None, None) , ], 1 , 4 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'UseGUID' , 'UseGUID' , ), 171, (171, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'UseGUID' , 'UseGUID' , ), 171, (171, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( 'NoHighlight' , 'NoHighlight' , ), 181, (181, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 224 , (3, 0, None, None) , 0 , )),
	(( 'NoHighlight' , 'NoHighlight' , ), 181, (181, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 232 , (3, 0, None, None) , 0 , )),
	(( 'AssySelectionMode' , 'AssySelectionMode' , ), 182, (182, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 240 , (3, 0, None, None) , 0 , )),
	(( 'AssySelectionMode' , 'AssySelectionMode' , ), 182, (182, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 248 , (3, 0, None, None) , 0 , )),
	(( 'KinematicActivated' , 'KinematicActivated' , ), 193, (193, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 256 , (3, 0, None, None) , 0 , )),
	(( 'KinematicActivated' , 'KinematicActivated' , ), 193, (193, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 264 , (3, 0, None, None) , 0 , )),
	(( 'StartupView' , 'StartupView' , ), 205, (205, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 272 , (3, 0, None, None) , 0 , )),
	(( 'StartupView' , 'StartupView' , ), 205, (205, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 280 , (3, 0, None, None) , 0 , )),
	(( 'EventsMask' , 'EventsMask' , ), 235, (235, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 288 , (3, 0, None, None) , 0 , )),
	(( 'EventsMask' , 'EventsMask' , ), 235, (235, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 296 , (3, 0, None, None) , 0 , )),
	(( 'RenderMirror' , 'RenderMirror' , ), 240, (240, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 304 , (3, 0, None, None) , 0 , )),
	(( 'RenderMirror' , 'RenderMirror' , ), 240, (240, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 312 , (3, 0, None, None) , 0 , )),
	(( 'Close' , 'Save' , 'status' , ), 1003, (1003, (), [ (11, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 320 , (3, 0, None, None) , 0 , )),
	(( 'Play' , ), 101, (101, (), [ ], 1 , 1 , 4 , 0 , 328 , (3, 0, None, None) , 0 , )),
	(( 'Stop' , ), 102, (102, (), [ ], 1 , 1 , 4 , 0 , 336 , (3, 0, None, None) , 0 , )),
	(( 'Antialiasing' , ), 106, (106, (), [ ], 1 , 1 , 4 , 0 , 344 , (3, 0, None, None) , 0 , )),
	(( 'ZoomFitAll' , ), 109, (109, (), [ ], 1 , 1 , 4 , 0 , 352 , (3, 0, None, None) , 0 , )),
	(( 'ZoomSelection' , ), 110, (110, (), [ ], 1 , 1 , 4 , 0 , 360 , (3, 0, None, None) , 0 , )),
	(( 'Undo' , ), 111, (111, (), [ ], 1 , 1 , 4 , 0 , 368 , (3, 0, None, None) , 0 , )),
	(( 'Redo' , ), 112, (112, (), [ ], 1 , 1 , 4 , 0 , 376 , (3, 0, None, None) , 0 , )),
	(( 'GoToConfiguration' , 'configurationName' , 'result' , ), 129, (129, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 384 , (3, 0, None, None) , 0 , )),
	(( 'GoToMarker' , 'markerName' , 'result' , ), 130, (130, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 392 , (3, 0, None, None) , 0 , )),
	(( 'PlayMarkerSequence' , 'markerName' , 'result' , ), 131, (131, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 400 , (3, 0, None, None) , 0 , )),
	(( 'GoToLayer' , 'layerName' , 'result' , ), 132, (132, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 408 , (3, 0, None, None) , 0 , )),
	(( 'AddLink' , 'Selection' , 'link' , 'autoLink' , 'result' , 
			 ), 134, (134, (), [ (8, 1, None, None) , (8, 1, None, None) , (11, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 416 , (3, 0, None, None) , 0 , )),
	(( 'SetPropertyMap' , 'propertyMap' , 'result' , ), 135, (135, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 424 , (3, 0, None, None) , 0 , )),
	(( 'RefreshScene' , 'refreshType' , ), 136, (136, (), [ (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 432 , (3, 0, None, None) , 0 , )),
	(( 'SetPropertySet' , 'Selection' , 'propertySet' , 'result' , ), 137, (137, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 440 , (3, 0, None, None) , 0 , )),
	(( 'SetVisibility' , 'Selection' , 'visibility' , 'invVisibility' , ), 139, (139, (), [ 
			 (8, 1, None, None) , (3, 1, None, None) , (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 448 , (3, 0, None, None) , 0 , )),
	(( 'InvertVisibility' , 'withGhost' , ), 140, (140, (), [ (11, 1, None, None) , ], 1 , 1 , 4 , 0 , 456 , (3, 0, None, None) , 0 , )),
	(( 'ShowNoShowCollaboration' , ), 141, (141, (), [ ], 1 , 1 , 4 , 0 , 464 , (3, 0, None, None) , 0 , )),
	(( 'Merge' , 'FileName' , ), 142, (142, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 472 , (3, 0, None, None) , 0 , )),
	(( 'GetOptionsPro' , 'optionsPro' , ), 143, (143, (), [ (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 480 , (3, 0, None, None) , 0 , )),
	(( 'SaveSMG' , 'filePath' , 'password' , 'result' , ), 148, (148, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 488 , (3, 0, None, None) , 0 , )),
	(( 'SaveImage' , 'pathName' , 'result' , ), 150, (150, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 496 , (3, 0, None, None) , 0 , )),
	(( 'SaveImageFullViewport' , 'pathName' , 'result' , ), 268, (268, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 504 , (3, 0, None, None) , 0 , )),
	(( 'Remove' , 'Selection' , 'result' , ), 152, (152, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 512 , (3, 0, None, None) , 0 , )),
	(( 'SelectOnFilter' , 'propertySet' , 'matchValue' , 'result' , ), 154, (154, (), [ 
			 (8, 1, None, None) , (11, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 520 , (3, 0, None, None) , 0 , )),
	(( 'GetLinks' , 'Selection' , 'selectionLinks' , ), 156, (156, (), [ (8, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 528 , (3, 0, None, None) , 0 , )),
	(( 'GetAllActors' , 'visibility' , 'actorsByVisibility' , ), 158, (158, (), [ (3, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 536 , (3, 0, None, None) , 0 , )),
	(( 'GetAllLinks' , 'links' , ), 159, (159, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 544 , (3, 0, None, None) , 0 , )),
	(( 'GetAllTexts' , 'texts' , ), 160, (160, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 552 , (3, 0, None, None) , 0 , )),
	(( 'GetAllCollaborations' , 'collaborations' , ), 161, (161, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 560 , (3, 0, None, None) , 0 , )),
	(( 'GetAllActorsOnFilter' , 'propertySet' , 'matchValue' , 'filteredActors' , ), 162, (162, (), [ 
			 (8, 1, None, None) , (11, 1, None, None) , (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 568 , (3, 0, None, None) , 0 , )),
	(( 'GetAllConfigurations' , 'configurations' , ), 163, (163, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 576 , (3, 0, None, None) , 0 , )),
	(( 'GetAllMarkers' , 'markers' , ), 164, (164, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 584 , (3, 0, None, None) , 0 , )),
	(( 'GetAllMetaPropertyDefinitions' , 'definition' , ), 266, (266, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 592 , (3, 0, None, None) , 0 , )),
	(( 'GetVisibility' , 'Selection' , 'visibility' , ), 166, (166, (), [ (8, 1, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 600 , (3, 0, None, None) , 0 , )),
	(( 'GetProperties' , 'Selection' , 'intersect' , 'matchValue' , 'selectionProperties' , 
			 ), 167, (167, (), [ (8, 1, None, None) , (11, 1, None, None) , (11, 1, None, None) , (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 608 , (3, 0, None, None) , 0 , )),
	(( 'GetAllLayers' , 'layers' , ), 168, (168, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 616 , (3, 0, None, None) , 0 , )),
	(( 'Pause' , ), 169, (169, (), [ ], 1 , 1 , 4 , 0 , 624 , (3, 0, None, None) , 0 , )),
	(( 'RotateInWCS' , 'Selection' , 'originX' , 'originY' , 'originZ' , 
			 'vectorX' , 'vectorY' , 'vectorZ' , 'angle' , 'result' , 
			 ), 172, (172, (), [ (8, 1, None, None) , (4, 1, None, None) , (4, 1, None, None) , (4, 1, None, None) , 
			 (4, 1, None, None) , (4, 1, None, None) , (4, 1, None, None) , (4, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 632 , (3, 0, None, None) , 0 , )),
	(( 'RotateInParent' , 'Selection' , 'originX' , 'originY' , 'originZ' , 
			 'vectorX' , 'vectorY' , 'vectorZ' , 'angle' , 'result' , 
			 ), 173, (173, (), [ (8, 1, None, None) , (4, 1, None, None) , (4, 1, None, None) , (4, 1, None, None) , 
			 (4, 1, None, None) , (4, 1, None, None) , (4, 1, None, None) , (4, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 640 , (3, 0, None, None) , 0 , )),
	(( 'TranslateInWCS' , 'Selection' , 'translationX' , 'translationY' , 'translationZ' , 
			 'result' , ), 174, (174, (), [ (8, 1, None, None) , (4, 1, None, None) , (4, 1, None, None) , 
			 (4, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 648 , (3, 0, None, None) , 0 , )),
	(( 'TranslateInParent' , 'strSelection' , 'translationX' , 'translationY' , 'translationZ' , 
			 'result' , ), 175, (175, (), [ (8, 1, None, None) , (4, 1, None, None) , (4, 1, None, None) , 
			 (4, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 656 , (3, 0, None, None) , 0 , )),
	(( 'TranslateInLCS' , 'strSelection' , 'translationX' , 'translationY' , 'translationZ' , 
			 'result' , ), 176, (176, (), [ (8, 1, None, None) , (4, 1, None, None) , (4, 1, None, None) , 
			 (4, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 664 , (3, 0, None, None) , 0 , )),
	(( 'RotateInLCS' , 'Selection' , 'originX' , 'originY' , 'originZ' , 
			 'vectorX' , 'vectorY' , 'vectorZ' , 'angle' , 'result' , 
			 ), 177, (177, (), [ (8, 1, None, None) , (4, 1, None, None) , (4, 1, None, None) , (4, 1, None, None) , 
			 (4, 1, None, None) , (4, 1, None, None) , (4, 1, None, None) , (4, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 672 , (3, 0, None, None) , 0 , )),
	(( 'SetCamera' , 'defCamera' , 'result' , ), 179, (179, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 680 , (3, 0, None, None) , 0 , )),
	(( 'GetCamera' , 'camera' , ), 180, (180, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 688 , (3, 0, None, None) , 0 , )),
	(( 'GetAssyRoot' , 'assyRoot' , ), 183, (183, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 696 , (3, 0, None, None) , 0 , )),
	(( 'GetAssyChild' , 'Selection' , 'assyChild' , ), 184, (184, (), [ (8, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 704 , (3, 0, None, None) , 0 , )),
	(( 'GetAssyParents' , 'Selection' , 'assyParents' , ), 185, (185, (), [ (8, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 712 , (3, 0, None, None) , 0 , )),
	(( 'IsPureAssy' , 'Selection' , 'IsPureAssy' , ), 186, (186, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 720 , (3, 0, None, None) , 0 , )),
	(( 'HasChilds' , 'Selection' , 'hasChild' , ), 187, (187, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 728 , (3, 0, None, None) , 0 , )),
	(( 'SetViewport' , 'defViewport' , 'result' , ), 189, (189, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 736 , (3, 0, None, None) , 0 , )),
	(( 'GetViewport' , 'viewport' , ), 190, (190, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 744 , (3, 0, None, None) , 0 , )),
	(( 'SetDigger' , 'defDigger' , 'result' , ), 191, (191, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 752 , (3, 0, None, None) , 0 , )),
	(( 'GetDigger' , 'digger' , ), 192, (192, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 760 , (3, 0, None, None) , 0 , )),
	(( 'AddProductIn' , 'FileName' , 'netGuid' , 'result' , ), 196, (196, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 768 , (3, 0, None, None) , 0 , )),
	(( 'PrintPreview' , ), 202, (202, (), [ ], 1 , 1 , 4 , 0 , 776 , (3, 0, None, None) , 0 , )),
	(( 'Print' , ), 203, (203, (), [ ], 1 , 1 , 4 , 0 , 784 , (3, 0, None, None) , 0 , )),
	(( 'AddMetaPropertyDefinition' , 'name' , 'metaType' , 'metaLabel' , 'stateFlags' , 
			 'unit' , 'bounded' , 'minValue' , 'maxValue' , 'result' , 
			 ), 206, (206, (), [ (8, 1, None, None) , (3, 1, None, None) , (8, 1, None, None) , (3, 1, None, None) , 
			 (3, 1, None, None) , (11, 1, None, None) , (4, 1, None, None) , (4, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 792 , (3, 0, None, None) , 0 , )),
	(( 'RemoveMetaPropertyDefinition' , 'name' , 'result' , ), 207, (207, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 800 , (3, 0, None, None) , 0 , )),
	(( 'AddMetaProperties' , 'Selection' , 'propertySet' , 'result' , ), 208, (208, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 808 , (3, 0, None, None) , 0 , )),
	(( 'RemoveMetaProperties' , 'propertyMap' , 'result' , ), 209, (209, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 816 , (3, 0, None, None) , 0 , )),
	(( 'SetNeutralProperties' , 'Selection' , 'propertySet' , 'result' , ), 215, (215, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 824 , (3, 0, None, None) , 0 , )),
	(( 'RestoreNeutralProperties' , 'Selection' , 'propertySet' , 'result' , ), 216, (216, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 832 , (3, 0, None, None) , 0 , )),
	(( 'GetPropertyValue' , 'Selection' , 'propertyName' , 'value' , ), 217, (217, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 840 , (3, 0, None, None) , 0 , )),
	(( 'SetDiggerRenderer' , 'digRenderer' , 'result' , ), 218, (218, (), [ (3, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 848 , (3, 0, None, None) , 0 , )),
	(( 'GetDiggerRenderer' , 'renderer' , ), 219, (219, (), [ (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 856 , (3, 0, None, None) , 0 , )),
	(( 'GetInvertedSelection' , 'Selection' , 'invertedSelection' , ), 220, (220, (), [ (8, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 864 , (3, 0, None, None) , 0 , )),
	(( 'SetColor' , 'layerColor' , 'red' , 'green' , 'blue' , 
			 'result' , ), 221, (221, (), [ (3, 1, None, None) , (3, 1, None, None) , (3, 1, None, None) , 
			 (3, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 872 , (3, 0, None, None) , 0 , )),
	(( 'LaunchQueryBBox' , 'result' , ), 222, (222, (), [ (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 880 , (3, 0, None, None) , 0 , )),
	(( 'GetAllActorsByType' , 'className' , 'filteredActors' , ), 223, (223, (), [ (8, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 888 , (3, 0, None, None) , 0 , )),
	(( 'ExportAllViews' , 'FileName' , 'result' , ), 224, (224, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 896 , (3, 0, None, None) , 0 , )),
	(( 'ExportViews' , 'viewsSelection' , 'FileName' , 'result' , ), 225, (225, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 904 , (3, 0, None, None) , 0 , )),
	(( 'ExportViewsPubs' , 'viewsIDSelection' , 'FileName' , 'result' , ), 226, (226, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 912 , (3, 0, None, None) , 0 , )),
	(( 'ExportAllPubs' , 'FileName' , 'visibility' , 'result' , ), 227, (227, (), [ 
			 (8, 1, None, None) , (3, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 920 , (3, 0, None, None) , 0 , )),
	(( 'ImportViewsData' , 'FileName' , 'viewIDInsertAfter' , 'result' , ), 228, (228, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 928 , (3, 0, None, None) , 0 , )),
	(( 'ImportPubs' , 'FileName' , 'result' , ), 229, (229, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 936 , (3, 0, None, None) , 0 , )),
	(( 'RemoveViews' , 'viewsIDSelection' , 'result' , ), 230, (230, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 944 , (3, 0, None, None) , 0 , )),
	(( 'UpdateView' , 'viewID' , 'result' , ), 231, (231, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 952 , (3, 0, None, None) , 0 , )),
	(( 'AddProductIn2' , 'FileName' , 'netGuidToAttach' , 'instanceGuid' , 'result' , 
			 ), 232, (232, (), [ (8, 1, None, None) , (8, 1, None, None) , (8, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 960 , (3, 0, None, None) , 0 , )),
	(( 'RemoveProductNodes' , 'Selection' , 'result' , ), 233, (233, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 968 , (3, 0, None, None) , 0 , )),
	(( 'AbortQueries' , 'result' , ), 234, (234, (), [ (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 976 , (3, 0, None, None) , 0 , )),
	(( 'Save' , 'result' , ), 236, (236, (), [ (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 984 , (3, 0, None, None) , 0 , )),
	(( 'SaveAs' , 'FileName' , 'result' , ), 237, (237, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 992 , (3, 0, None, None) , 0 , )),
	(( 'SaveDocumentAs' , 'FileName' , 'result' , ), 324, (324, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1000 , (3, 0, None, None) , 0 , )),
	(( 'LoadGeometries' , 'netGuid' , 'result' , ), 238, (238, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1008 , (3, 0, None, None) , 0 , )),
	(( 'UnloadGeometries' , 'netGuid' , 'result' , ), 239, (239, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1016 , (3, 0, None, None) , 0 , )),
	(( 'Update' , 'FileName' , 'result' , ), 241, (241, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1024 , (3, 0, None, None) , 0 , )),
	(( 'AutoKeyMode' , 'AutoKeyMode' , ), 242, (242, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1032 , (3, 0, None, None) , 0 , )),
	(( 'AutoKeyMode' , 'AutoKeyMode' , ), 242, (242, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1040 , (3, 0, None, None) , 0 , )),
	(( 'IsCollaboration' , 'Selection' , 'bResult' , ), 243, (243, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1048 , (3, 0, None, None) , 0 , )),
	(( 'GetCADCaptures' , 'Selection' , ), 244, (244, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1056 , (3, 0, None, None) , 0 , )),
	(( 'GetCADViews' , 'Selection' , ), 245, (245, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1064 , (3, 0, None, None) , 0 , )),
	(( 'GetAllViews' , 'Selection' , ), 246, (246, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1072 , (3, 0, None, None) , 0 , )),
	(( 'GetChildren' , 'Selection' , 'mask' , 'resultSelection' , ), 247, (247, (), [ 
			 (8, 1, None, None) , (3, 1, None, None) , (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1080 , (3, 0, None, None) , 0 , )),
	(( 'GoToView' , 'viewName' , 'result' , ), 248, (248, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1088 , (3, 0, None, None) , 0 , )),
	(( 'GetChildrenView' , 'viewName' , 'resultSelection' , ), 249, (249, (), [ (8, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1096 , (3, 0, None, None) , 0 , )),
	(( 'CreateBBox' , 'ptAx' , 'ptAy' , 'ptAz' , 'ptBx' , 
			 'ptBy' , 'ptBz' , 'resultSelection' , ), 250, (250, (), [ (4, 1, None, None) , 
			 (4, 1, None, None) , (4, 1, None, None) , (4, 1, None, None) , (4, 1, None, None) , (4, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1104 , (3, 0, None, None) , 0 , )),
	(( 'GetOnFilter' , 'propertiesSet' , 'matchValue' , 'resultSelection' , ), 251, (251, (), [ 
			 (8, 1, None, None) , (11, 1, None, None) , (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1112 , (3, 0, None, None) , 0 , )),
	(( 'CreateAssyGroup' , 'Selection' , 'resultSelection' , ), 252, (252, (), [ (8, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1120 , (3, 0, None, None) , 0 , )),
	(( 'CreateDummyObject' , 'Selection' , 'resultSelection' , ), 253, (253, (), [ (8, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1128 , (3, 0, None, None) , 0 , )),
	(( 'ChangeParent' , 'Selection' , 'selectionNewParent' , 'result' , ), 254, (254, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1136 , (3, 0, None, None) , 0 , )),
	(( 'ReplaceGeometry' , 'FileName' , 'strNetGUIDSource' , 'selectionDest' , 'result' , 
			 ), 255, (255, (), [ (8, 1, None, None) , (8, 1, None, None) , (8, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1144 , (3, 0, None, None) , 0 , )),
	(( 'ExportResources' , 'folderPath' , 'result' , ), 256, (256, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1152 , (3, 0, None, None) , 0 , )),
	(( 'ImportViewsData2' , 'FileName' , 'viewIDInsertAfter' , 'forceResolution' , 'allInstances' , 
			 'result' , ), 275, (275, (), [ (8, 1, None, None) , (8, 1, None, None) , (11, 1, None, None) , 
			 (11, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1160 , (3, 0, None, None) , 0 , )),
	(( 'GetMeasurementPersistency' , 'result' , ), 281, (281, (), [ (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1168 , (3, 0, None, None) , 0 , )),
	(( 'SetMeasurementPersistency' , 'result' , ), 282, (282, (), [ (11, 1, None, None) , ], 1 , 1 , 4 , 0 , 1176 , (3, 0, None, None) , 0 , )),
	(( 'FilterViews' , 'Selection' , 'result' , ), 283, (283, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1184 , (3, 0, None, None) , 0 , )),
	(( 'GetActorsProperties' , 'Selection' , 'result' , ), 285, (285, (), [ (8, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1192 , (3, 0, None, None) , 0 , )),
	(( 'GetCrossHighlightGeometry' , 'strGUID' , 'bNormals' , 'result' , ), 287, (287, (), [ 
			 (8, 1, None, None) , (11, 1, None, None) , (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1200 , (3, 0, None, None) , 0 , )),
	(( 'SetViewPosition' , 'position' , 'Selection' , ), 288, (288, (), [ (3, 1, None, None) , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 1208 , (3, 0, None, None) , 0 , )),
	(( 'SetMouseNavigationMode' , 'navigationType' , 'result' , ), 267, (267, (), [ (3, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1216 , (3, 0, None, None) , 0 , )),
	(( 'GetActorsClassName' , 'Selection' , 'result' , ), 274, (274, (), [ (8, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1224 , (3, 0, None, None) , 0 , )),
	(( 'CreateView' , 'name' , 'position' , 'result' , ), 289, (289, (), [ 
			 (8, 1, None, None) , (3, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1232 , (3, 0, None, None) , 0 , )),
	(( 'CreateCameraView' , 'name' , 'position' , 'result' , ), 290, (290, (), [ 
			 (8, 1, None, None) , (3, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1240 , (3, 0, None, None) , 0 , )),
	(( 'DeleteView' , 'view' , 'result' , ), 291, (291, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1248 , (3, 0, None, None) , 0 , )),
	(( 'ScrollViewToTop' , 'strName' , ), 271, (271, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 1256 , (3, 0, None, None) , 0 , )),
	(( 'PMICrossHLStayOnTop' , 'PMICrossHLStayOnTop' , ), 297, (297, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1264 , (3, 0, None, None) , 0 , )),
	(( 'PMICrossHLStayOnTop' , 'PMICrossHLStayOnTop' , ), 297, (297, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1272 , (3, 0, None, None) , 0 , )),
	(( 'PMICrossHLOpacity' , 'PMICrossHLOpacity' , ), 298, (298, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 1280 , (3, 0, None, None) , 0 , )),
	(( 'PMICrossHLOpacity' , 'PMICrossHLOpacity' , ), 298, (298, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 1288 , (3, 0, None, None) , 0 , )),
	(( 'DropInViewport' , 'DropInViewport' , ), 299, (299, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1296 , (3, 0, None, None) , 0 , )),
	(( 'DropInViewport' , 'DropInViewport' , ), 299, (299, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1304 , (3, 0, None, None) , 0 , )),
	(( 'ViewNameTemplate' , 'ViewNameTemplate' , ), 301, (301, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1312 , (3, 0, None, None) , 0 , )),
	(( 'ViewNameTemplate' , 'ViewNameTemplate' , ), 301, (301, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 1320 , (3, 0, None, None) , 0 , )),
	(( 'RenameView' , 'strViewName' , 'strNewViewName' , 'result' , ), 300, (300, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1328 , (3, 0, None, None) , 0 , )),
	(( 'GetAllMarkersDetails' , 'markers' , ), 306, (306, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1336 , (3, 0, None, None) , 0 , )),
	(( 'GetAllGenericAnimation' , 'strList' , ), 308, (308, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1344 , (3, 0, None, None) , 0 , )),
	(( 'GetGenericAnimationProperty' , 'nIndex' , 'strStep' , ), 307, (307, (), [ (3, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1352 , (3, 0, None, None) , 0 , )),
	(( 'AllNameOccurences' , 'AllNameOccurences' , ), 309, (309, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1360 , (3, 0, None, None) , 0 , )),
	(( 'AllNameOccurences' , 'AllNameOccurences' , ), 309, (309, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1368 , (3, 0, None, None) , 0 , )),
	(( 'FileBrowseDialog' , 'filePath' , ), 310, (310, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1376 , (3, 0, None, None) , 0 , )),
	(( 'ResetZoomFitAll' , 'viewList' , ), 311, (311, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 1384 , (3, 0, None, None) , 0 , )),
	(( 'GetAllSelectionSets' , 'selectionsets' , ), 312, (312, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1392 , (3, 0, None, None) , 0 , )),
	(( 'CreateCallout' , 'Selection' , 'eGenMode' , 'eGenAttach' , 'result' , 
			 ), 314, (314, (), [ (8, 1, None, None) , (3, 1, None, None) , (3, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1400 , (3, 0, None, None) , 0 , )),
	(( 'LinearExplode' , 'Selection' , 'vectorX' , 'vectorY' , 'vectorZ' , 
			 'fDistance' , 'result' , ), 315, (315, (), [ (8, 1, None, None) , (3, 1, None, None) , 
			 (3, 1, None, None) , (3, 1, None, None) , (4, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1408 , (3, 0, None, None) , 0 , )),
	(( 'GetAllHotspots' , 'Selection' , ), 316, (316, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1416 , (3, 0, None, None) , 0 , )),
	(( 'SaveImageUsingProfile' , 'pathName' , 'profileFullPath' , 'result' , ), 319, (319, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1424 , (3, 0, None, None) , 0 , )),
	(( 'CreateSelectionSet' , 'SelectionSetName' , 'Selection' , 'result' , ), 320, (320, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1432 , (3, 0, None, None) , 0 , )),
	(( 'LockView' , 'view' , 'result' , ), 321, (321, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1440 , (3, 0, None, None) , 0 , )),
	(( 'UnlockView' , 'view' , 'result' , ), 322, (322, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1448 , (3, 0, None, None) , 0 , )),
	(( 'GetAllLockedViews' , 'Selection' , ), 323, (323, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1456 , (3, 0, None, None) , 0 , )),
	(( 'AddSignature' , 'LogoPath' , 'Agreement' , 'eSignatureMode' , 'result' , 
			 ), 325, (325, (), [ (8, 1, None, None) , (8, 1, None, None) , (3, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1464 , (3, 0, None, None) , 0 , )),
	(( 'GetGround' , 'pGround' , ), 265, (265, (), [ (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1472 , (3, 0, None, None) , 0 , )),
	(( 'SetGround' , 'pGround' , ), 264, (264, (), [ (16392, 1, None, None) , ], 1 , 1 , 4 , 0 , 1480 , (3, 0, None, None) , 0 , )),
	(( 'DeleteGenericAnimation' , 'TagID' , 'result' , ), 326, (326, (), [ (3, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1488 , (3, 0, None, None) , 0 , )),
	(( 'UpdateGenericAnimation' , 'TagID' , 'SelectedActors' , 'result' , ), 327, (327, (), [ 
			 (3, 1, None, None) , (8, 1, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1496 , (3, 0, None, None) , 0 , )),
	(( 'SetCameraKey' , ), 328, (328, (), [ ], 1 , 1 , 4 , 0 , 1504 , (3, 0, None, None) , 0 , )),
	(( 'SaveDocumentProperties' , 'strFilePath' , 'result' , ), 330, (330, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1512 , (3, 0, None, None) , 0 , )),
	(( 'LoadDocumentProperties' , 'strFilePath' , 'result' , ), 329, (329, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1520 , (3, 0, None, None) , 0 , )),
	(( 'GetLocationInWorld' , 'strSelection' , 'result' , ), 257, (257, (), [ (8, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1528 , (3, 0, None, None) , 0 , )),
	(( 'GetLocationInParent' , 'strSelection' , 'result' , ), 258, (258, (), [ (8, 1, None, None) , 
			 (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1536 , (3, 0, None, None) , 0 , )),
	(( 'GetLocationInPart' , 'strSelectionSource' , 'strSelectionDest' , 'result' , ), 259, (259, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 1544 , (3, 0, None, None) , 0 , )),
	(( 'HighlightInvisibleActorsinViewport' , 'bHighlight' , ), 333, (333, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1552 , (3, 0, None, None) , 0 , )),
	(( 'HighlightInvisibleActorsinViewport' , 'bHighlight' , ), 333, (333, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1560 , (3, 0, None, None) , 0 , )),
	(( 'SaveBOMTree' , 'strFilePath' , 'result' , ), 334, (334, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1568 , (3, 0, None, None) , 0 , )),
	(( 'ConfigureColumnsForAssemblyPane' , 'strFilePath' , 'result' , ), 335, (335, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 1576 , (3, 0, None, None) , 0 , )),
]

IDocumentEvents_vtables_dispatch_ = 1
IDocumentEvents_vtables_ = [
	(( 'AddMarker' , 'sender' , 'label' , 'frame' , ), 130, (130, (), [ 
			 (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , (8, 1, None, None) , (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'DeleteMarker' , 'sender' , 'frame' , ), 131, (131, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'ChangeSelection' , 'sender' , 'Selection' , ), 101, (101, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'link' , 'sender' , 'link' , 'autoLink' , ), 102, (102, (), [ 
			 (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , (8, 1, None, None) , (11, 1, None, None) , ], 1 , 1 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'GetHighlight' , 'sender' , 'name' , ), 103, (103, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'LooseHighlight' , 'sender' , 'name' , ), 104, (104, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'ChangeProperties' , 'sender' , 'propertyMap' , ), 105, (105, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'ChangePos' , 'sender' , 'Pos' , ), 106, (106, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'ChangeConfiguration' , 'sender' , 'configurationName' , ), 107, (107, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'ChangeLayer' , 'sender' , 'layerName' , ), 108, (108, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'AddActors' , 'sender' , 'Selection' , ), 109, (109, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( 'RemoveActors' , 'sender' , 'Selection' , ), 110, (110, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'PlayChangePos' , 'sender' , 'Pos' , ), 111, (111, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'BeginLoadModel' , 'sender' , ), 112, (112, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , ], 1 , 1 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'EndLoadModel' , 'sender' , 'Selection' , ), 113, (113, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 168 , (3, 0, None, None) , 0 , )),
	(( 'ChangeVisibility' , 'sender' , 'Selection' , ), 114, (114, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( 'Click' , 'sender' , 'name' , ), 115, (115, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'DblClick' , 'sender' , 'name' , ), 116, (116, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'PlayBegin' , 'sender' , 'Pos' , ), 117, (117, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'PlayEnd' , 'sender' , 'Pos' , ), 118, (118, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'ChangeMarker' , 'sender' , 'label' , ), 119, (119, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( 'PlayChangeMarker' , 'sender' , 'label' , ), 120, (120, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 224 , (3, 0, None, None) , 0 , )),
	(( 'DisplayLessThan32Bits' , 'sender' , ), 121, (121, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , ], 1 , 1 , 4 , 0 , 232 , (3, 0, None, None) , 0 , )),
	(( 'ChangeViews' , 'sender' , 'viewID' , ), 122, (122, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 240 , (3, 0, None, None) , 0 , )),
	(( 'CloseDocument' , 'sender' , ), 123, (123, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , ], 1 , 1 , 4 , 0 , 248 , (3, 0, None, None) , 0 , )),
	(( 'AddView' , 'sender' , 'Selection' , ), 125, (125, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 256 , (3, 0, None, None) , 0 , )),
	(( 'RemoveViews' , 'sender' , 'Selection' , ), 126, (126, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 264 , (3, 0, None, None) , 0 , )),
	(( 'MoveViews' , 'sender' , 'Selection' , 'DropIndex' , ), 127, (127, (), [ 
			 (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , (8, 1, None, None) , (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 272 , (3, 0, None, None) , 0 , )),
	(( 'Save' , 'sender' , 'status' , ), 128, (128, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 280 , (3, 0, None, None) , 0 , )),
	(( 'PasteViews' , 'sender' , 'selectionCopy' , 'selectionPaste' , ), 129, (129, (), [ 
			 (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , (8, 1, None, None) , (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 288 , (3, 0, None, None) , 0 , )),
	(( 'AddGenericAnimation' , 'sender' , 'stepPropMap' , ), 132, (132, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 296 , (3, 0, None, None) , 0 , )),
	(( 'ClickGenericAnimation' , 'sender' , 'stepPropMap' , ), 135, (135, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 304 , (3, 0, None, None) , 0 , )),
	(( 'MoveGenericAnimation' , 'sender' , 'stepPropMap' , ), 133, (133, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 312 , (3, 0, None, None) , 0 , )),
	(( 'DeleteGenericAnimation' , 'sender' , 'stepPropMap' , ), 134, (134, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 320 , (3, 0, None, None) , 0 , )),
	(( 'SelectFromSelectionSets' , 'sender' , 'SelectSelection' , ), 137, (137, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 328 , (3, 0, None, None) , 0 , )),
	(( 'ChangeAllViews' , 'sender' , 'strViews' , ), 136, (136, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 336 , (3, 0, None, None) , 0 , )),
	(( 'OnUndo' , 'sender' , 'commandName' , 'Selection' , ), 138, (138, (), [ 
			 (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , (8, 1, None, None) , (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 344 , (3, 0, None, None) , 0 , )),
	(( 'OnRedo' , 'sender' , 'commandName' , 'Selection' , ), 139, (139, (), [ 
			 (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , (8, 1, None, None) , (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 352 , (3, 0, None, None) , 0 , )),
	(( 'CreateSelectionSet' , 'sender' , 'Selection' , ), 142, (142, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 360 , (3, 0, None, None) , 0 , )),
	(( 'DeleteSelectionSet' , 'sender' , 'Selection' , ), 141, (141, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 368 , (3, 0, None, None) , 0 , )),
	(( 'UpdateSelectionSet' , 'sender' , 'Selection' , ), 140, (140, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 376 , (3, 0, None, None) , 0 , )),
	(( 'CreateHotspot' , 'sender' , 'Selection' , ), 145, (145, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 384 , (3, 0, None, None) , 0 , )),
	(( 'DeleteHotspot' , 'sender' , 'Selection' , ), 144, (144, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 392 , (3, 0, None, None) , 0 , )),
	(( 'UpdateHotspot' , 'sender' , 'Selection' , ), 143, (143, (), [ (9, 1, None, "IID('{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}')") , 
			 (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 400 , (3, 0, None, None) , 0 , )),
]

RecordMap = {
}

CLSIDToClassMap = {
	'{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}' : IDocument,
	'{48CDCE04-6ECF-4D0A-BC25-F48FA1B1AC78}' : IApplication,
	'{DE79BC3C-1CE6-4DB7-816A-6307D8ADF418}' : IDocumentEvents,
	'{E5B985AA-AC9B-44E3-A3FF-44D9E7B90714}' : Document,
	'{21268FAA-FC76-44CA-84AB-217807808491}' : IApplicationEvents,
	'{3F41659D-BAB2-466C-A4D2-6E3CECFFB964}' : Application,
}
CLSIDToPackageMap = {}
win32com.client.CLSIDToClass.RegisterCLSIDsFromDict( CLSIDToClassMap )
VTablesToPackageMap = {}
VTablesToClassMap = {
	'{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}' : 'IDocument',
	'{48CDCE04-6ECF-4D0A-BC25-F48FA1B1AC78}' : 'IApplication',
	'{DE79BC3C-1CE6-4DB7-816A-6307D8ADF418}' : 'IDocumentEvents',
	'{21268FAA-FC76-44CA-84AB-217807808491}' : 'IApplicationEvents',
}


NamesToIIDMap = {
	'IDocument' : '{A69EAF5C-D6B5-4F8F-AC94-0F3761682E97}',
	'IApplication' : '{48CDCE04-6ECF-4D0A-BC25-F48FA1B1AC78}',
	'IDocumentEvents' : '{DE79BC3C-1CE6-4DB7-816A-6307D8ADF418}',
	'IApplicationEvents' : '{21268FAA-FC76-44CA-84AB-217807808491}',
}

win32com.client.constants.__dicts__.append(constants.__dict__)

