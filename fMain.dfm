�
 TFRMMAIN 00  TPF0TfrmMainfrmMainLeft�Top� Caption   Учёт пропусковClientHeightNClientWidthfColor	clBtnFaceConstraints.MinHeight� Constraints.MinWidthMFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
KeyPreview	Menu	MainMenu1OldCreateOrderPositionpoScreenCenter
OnActivateFormActivateOnCreate
FormCreate	OnDestroyFormDestroyOnShowFormShow
DesignSizefN PixelsPerInch`
TextHeight TLabelLabel1LeftTopWidth#HeightCaption   @C??0  TLabelLabel2LeftTop6Width(HeightCaption   !BC45=B  	TComboBoxcbGroupsLeftTopWidthYHeightStylecsDropDownList
ItemHeightSorted	TabOrder OnChangecbGroupsChangeOnKeyUpcbGroupsKeyUp  	TComboBoxcbStudyLeftTopIWidth� HeightStylecsDropDownList
ItemHeightSorted	TabOrderOnChangecbStudyChangeOnKeyUpcbStudyKeyUp  TPageControlPageControl1Left� TopWidth]Height`
ActivePage	TabSheet1TabOrder 	TTabSheet	TabSheet1Caption!   Добавить пропуски TLabelLabel3LeftTopWidthHeightCaption   0B0  TLabelLabel4LeftwTopWidthHeightCaption   #206  TLabelLabel5Left� TopWidthHeightCaption   5C2  TDateTimePickerdtpAddDelayLeftTopWidthYHeightDate h�s�@�@Time h�s�@�@TabOrder   TEdit	edtHoursULeftwTopWidth!HeightTabOrder  TEdit	edtHoursNLeft� TopWidth!HeightTabOrder  TButtonbtnAddDelayLeft� TopWidthbHeightCaption   BACBAB2>20;Default	TabOrderOnClickbtnAddDelayClick   	TTabSheet	TabSheet2Caption   От и до
ImageIndex TLabelLabel6LeftTopWidthHeightCaption   B  TLabelLabel7LeftwTopWidthHeightCaption   >  TDateTimePickerdtpFromLeftTopWidthUHeightDate X
���@�@Time X
���@�@TabOrder   TDateTimePickerdtpToLeftwTopWidthUHeightDate P���@�@Time P���@�@TabOrder  TButtonbtnSearchDelaysLeft� TopWidthKHeightCaption   >8A:TabOrderOnClickbtnSearchDelaysClick   	TTabSheet	TabSheet3Caption
   !B0B8AB8:0
ImageIndexOnShowTabSheet3Show TLabelLabel8LeftTopWidthdHeightCaption   За прошлый месяц:  TLabelLabel9LeftTopWidthHHeightCaption   Уважительно:  TLabelLabel10LeftTop)WidthSHeightCaption   Неуважительно:  TLabellblPrevULeftmTopWidthHeight  TLabellblPrevNLeftmTop)WidthHeight  TLabellblStatNLeft%Top)WidthHeight  TLabellblStatULeft%TopWidthHeight  TLabelLabel13Left� TopWidthHHeightCaption   Уважительно:  TLabelLabel14Left� Top)WidthSHeightCaption   Неуважительно:  TLabelLabel15Left� TopWidthaHeightCaption   За текущий месяц:  TPanelPanel1Left� TopWidthHeight<
BevelOuter	bvLoweredTabOrder    	TTabSheet	TabSheet4Caption   $8;LB@
ImageIndex TButtonbtnSelectGroupLeft>TopWidth� HeightCaption:   Все прогулы по выбранной группеTabOrder OnClickbtnSelectGroupClick  TButtonbtnSelectStudyLeft>Top%Width� HeightCaption@   Все прогулы по выбранному студентуTabOrderOnClickbtnSelectStudyClick    
TStatusBar
StatusBar1Left Top:WidthfHeightPanelsWidth2    TPanelPanel2Left ToptWidthfHeight� AnchorsakLeftakTopakRightakBottom TabOrder
DesignSizef�   TLabelLabel11Left� TopWidthHeightCaption   B  TLabelLabel12Left/TopWidthHeightCaption   >  TLabelLabel16LeftTopWidth<HeightCaption   >8A:Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  	TCheckBox	cbOnGroupLeft
Top-WidthPHeightCaption   По группеTabOrder   	TCheckBox	cbOnStudyLeft`Top-WidthZHeightCaption   По студентуTabOrder  TDateTimePickerdtpSearchFromLeft� Top-WidthUHeightDate X
���@�@Time X
���@�@TabOrder  TDateTimePickerdtpSearchToLeft/Top-WidthUHeightDate P���@�@Time P���@�@TabOrder  TButton
btnFSearchLeft�Top)Width3HeightCaption   >8A:TabOrderOnClickbtnFSearchClick  TButtonbtnPrintLeftTop)WidthVHeightAnchorsakTopakRight Caption   Вывод в ExcelTabOrderOnClickbtnPrintClick  TStringGrid
StringGridLeftTopJWidth`Height{AlignalCustomAnchorsakLeftakTopakRightakBottom OptionsgoFixedVertLinegoFixedHorzLine
goVertLine
goHorzLinegoRangeSelectgoColSizing TabOrder
OnDblClickStringGridDblClickOnKeyUpStringGridKeyUp   	TMainMenu	MainMenu1Left�  	TMenuItemmFileCaption   $09; 	TMenuItemmExitCaption   KE>4OnClick
mExitClick   	TMenuItemmEditCaption   @02:0 	TMenuItem	mAddGroupCaption   Добавить группуOnClickmAddGroupClick  	TMenuItem
mEditGroupCaption   Изменить группуOnClickmEditGroupClick  	TMenuItem	mDelGroupCaption   Удалить группуOnClickmDelGroupClick  	TMenuItemN5Caption-  	TMenuItem	mAddStudyCaption!   Добавить студентаOnClickmAddStudyClick  	TMenuItem
mEditStudyCaption!   Изменить студентаOnClickmEditStudyClick  	TMenuItem	mDelStudyCaption   Удалить студентаOnClickmDelStudyClick   	TMenuItemN6Caption   ><>ILOnClickN6Click    