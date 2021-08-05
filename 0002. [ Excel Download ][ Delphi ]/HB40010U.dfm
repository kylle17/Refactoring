inherited HB40010: THB40010
  Caption = ''
  ClientHeight = 903
  ClientWidth = 1892
  Font.Height = -11
  Font.Name = 'Tahoma'
  OnShow = FormShow
  ExplicitWidth = 1908
  ExplicitHeight = 942
  PixelsPerInch = 96
  TextHeight = 13
  inherited P_HEAD: TPanel
    Width = 1892
    ExplicitWidth = 1892
    inherited fmTITLE: TLabel
      Width = 175
      Caption = #54616#44592#55092#44032#44228#54925#49436
      ExplicitWidth = 175
    end
    inherited P_HEAD_BUTTON: TPanel
      Left = 1448
      ExplicitLeft = 1448
      inherited SB_INIT: TSpeedButton
        OnClick = SB_INITClick
      end
      inherited SB_QUERY: TSpeedButton
        OnClick = SB_QUERYClick
        ExplicitTop = 7
      end
      inherited SB_SAVE: TSpeedButton
        ExplicitTop = 2
      end
      inherited SB_PRINT: TSpeedButton
        OnClick = SB_PRINTClick
      end
      inherited SB_EXCEL: TSpeedButton
        OnClick = SB_EXCELClick
      end
    end
    inherited EdtKey: TEdit
      Color = 8454143
    end
  end
  inherited Panel1: TPanel
    Width = 1892
    Color = 16052463
    ExplicitWidth = 1892
  end
  inherited Panel2: TPanel
    Top = 896
    Width = 1892
    Color = 16052463
    ExplicitTop = 896
    ExplicitWidth = 1892
  end
  inherited Panel3: TPanel
    Height = 822
    Color = 16052463
    ExplicitHeight = 822
  end
  inherited Panel4: TPanel
    Left = 1885
    Height = 822
    Color = 16052463
    ExplicitLeft = 1885
    ExplicitHeight = 822
  end
  inherited ScrollBox1: TScrollBox
    Width = 1878
    Height = 822
    Color = 16052463
    ExplicitWidth = 1878
    ExplicitHeight = 822
    object Splitter1: TSplitter
      Left = 377
      Top = 0
      Width = 6
      Height = 822
      Color = 16052463
      ParentColor = False
      ExplicitLeft = 469
      ExplicitTop = -3
    end
    object Panel12: TPanel
      Left = 383
      Top = 0
      Width = 1495
      Height = 822
      Align = alClient
      BevelOuter = bvNone
      Color = clWhite
      Ctl3D = False
      ParentBackground = False
      ParentCtl3D = False
      TabOrder = 0
      object Shape26: TShape
        Left = 0
        Top = 745
        Width = 1
        Height = 4
        Align = alLeft
        Brush.Color = 15592941
        Pen.Color = 14540253
        ExplicitLeft = 8
        ExplicitTop = 45
        ExplicitHeight = 704
      end
      object Shape28: TShape
        Left = 0
        Top = 695
        Width = 1495
        Height = 1
        Align = alBottom
        Brush.Color = 15592941
        Pen.Color = 14540253
        ExplicitLeft = 1
        ExplicitTop = 45
        ExplicitWidth = 704
      end
      object Shape27: TShape
        Left = 1494
        Top = 745
        Width = 1
        Height = 4
        Align = alRight
        Brush.Color = 15592941
        Pen.Color = 14540253
        ExplicitLeft = 8
        ExplicitTop = 45
        ExplicitHeight = 704
      end
      object pnl1: TPanel
        Left = 0
        Top = 45
        Width = 1495
        Height = 350
        Align = alTop
        Caption = 'pnl1'
        TabOrder = 2
        Visible = False
        object asgDetail: TAdvStringGrid
          Left = 1
          Top = 1
          Width = 1493
          Height = 348
          Cursor = crDefault
          TabStop = False
          Align = alClient
          BevelInner = bvNone
          BevelOuter = bvNone
          BorderStyle = bsNone
          ColCount = 32
          Ctl3D = False
          DefaultColWidth = 30
          DefaultRowHeight = 35
          DrawingStyle = gdsClassic
          FixedColor = 528832347
          FixedRows = 2
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = #45208#45588#44256#46357
          Font.Style = []
          Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goTabs]
          ParentCtl3D = False
          ParentFont = False
          ParentShowHint = False
          ScrollBars = ssBoth
          ShowHint = True
          TabOrder = 0
          Visible = False
          OnKeyDown = asgDetailKeyDown
          OnKeyPress = asgDetailKeyPress
          GridLineColor = 14013909
          GridFixedLineColor = 14013909
          HoverRowCells = [hcNormal, hcSelected]
          OnGetCellColor = asgDetailGetCellColor
          OnGetAlignment = asgDetailGetAlignment
          OnGetEditorType = asgDetailGetEditorType
          ActiveCellFont.Charset = DEFAULT_CHARSET
          ActiveCellFont.Color = clWindowText
          ActiveCellFont.Height = -11
          ActiveCellFont.Name = 'Tahoma'
          ActiveCellFont.Style = [fsBold]
          Bands.Active = True
          Bands.PrimaryColor = clWhite
          Bands.SecondaryColor = 16316664
          BorderColor = 15592941
          ColumnSize.Stretch = True
          ControlLook.FixedGradientFrom = 15658734
          ControlLook.FixedGradientTo = 15658734
          ControlLook.FixedGradientHoverFrom = clGray
          ControlLook.FixedGradientHoverTo = clWhite
          ControlLook.FixedGradientDownFrom = clGray
          ControlLook.FixedGradientDownTo = clSilver
          ControlLook.DropDownHeader.Font.Charset = DEFAULT_CHARSET
          ControlLook.DropDownHeader.Font.Color = clWindowText
          ControlLook.DropDownHeader.Font.Height = -11
          ControlLook.DropDownHeader.Font.Name = 'Tahoma'
          ControlLook.DropDownHeader.Font.Style = []
          ControlLook.DropDownHeader.Visible = True
          ControlLook.DropDownHeader.Buttons = <>
          ControlLook.DropDownFooter.Font.Charset = DEFAULT_CHARSET
          ControlLook.DropDownFooter.Font.Color = clWindowText
          ControlLook.DropDownFooter.Font.Height = -11
          ControlLook.DropDownFooter.Font.Name = 'Tahoma'
          ControlLook.DropDownFooter.Font.Style = []
          ControlLook.DropDownFooter.Visible = True
          ControlLook.DropDownFooter.Buttons = <>
          DefaultAlignment = taCenter
          EnhTextSize = True
          Filter = <>
          FilterDropDown.Font.Charset = DEFAULT_CHARSET
          FilterDropDown.Font.Color = clWindowText
          FilterDropDown.Font.Height = -11
          FilterDropDown.Font.Name = 'Tahoma'
          FilterDropDown.Font.Style = []
          FilterDropDown.TextChecked = 'Checked'
          FilterDropDown.TextUnChecked = 'Unchecked'
          FilterDropDownClear = '(All)'
          FilterEdit.TypeNames.Strings = (
            'Starts with'
            'Ends with'
            'Contains'
            'Not contains'
            'Equal'
            'Not equal'
            'Larger than'
            'Smaller than'
            'Clear')
          FixedColWidth = 30
          FixedRowHeight = 30
          FixedFont.Charset = DEFAULT_CHARSET
          FixedFont.Color = clBlack
          FixedFont.Height = -13
          FixedFont.Name = #45208#45588#44256#46357
          FixedFont.Style = []
          Flat = True
          FloatFormat = '%.2f'
          HideFocusRect = True
          HoverButtons.Buttons = <>
          HoverButtons.Position = hbLeftFromColumnLeft
          HTMLSettings.ImageFolder = 'images'
          HTMLSettings.ImageBaseName = 'img'
          MouseActions.DisjunctCellSelect = True
          MouseActions.RangeSelectAndEdit = True
          MouseActions.RowSelect = True
          Navigation.AdvanceOnEnter = True
          Navigation.AllowClipboardShortCuts = True
          Navigation.AllowSmartClipboard = True
          PrintSettings.DateFormat = 'dd/mm/yyyy'
          PrintSettings.Font.Charset = DEFAULT_CHARSET
          PrintSettings.Font.Color = clWindowText
          PrintSettings.Font.Height = -11
          PrintSettings.Font.Name = 'Tahoma'
          PrintSettings.Font.Style = []
          PrintSettings.FixedFont.Charset = DEFAULT_CHARSET
          PrintSettings.FixedFont.Color = clWindowText
          PrintSettings.FixedFont.Height = -11
          PrintSettings.FixedFont.Name = 'Tahoma'
          PrintSettings.FixedFont.Style = []
          PrintSettings.HeaderFont.Charset = DEFAULT_CHARSET
          PrintSettings.HeaderFont.Color = clWindowText
          PrintSettings.HeaderFont.Height = -11
          PrintSettings.HeaderFont.Name = 'Tahoma'
          PrintSettings.HeaderFont.Style = []
          PrintSettings.FooterFont.Charset = DEFAULT_CHARSET
          PrintSettings.FooterFont.Color = clWindowText
          PrintSettings.FooterFont.Height = -11
          PrintSettings.FooterFont.Name = 'Tahoma'
          PrintSettings.FooterFont.Style = []
          PrintSettings.PageNumSep = '/'
          SearchFooter.FindNextCaption = 'Find &next'
          SearchFooter.FindPrevCaption = 'Find &previous'
          SearchFooter.Font.Charset = DEFAULT_CHARSET
          SearchFooter.Font.Color = clWindowText
          SearchFooter.Font.Height = -11
          SearchFooter.Font.Name = 'Tahoma'
          SearchFooter.Font.Style = []
          SearchFooter.HighLightCaption = 'Highlight'
          SearchFooter.HintClose = 'Close'
          SearchFooter.HintFindNext = 'Find next occurrence'
          SearchFooter.HintFindPrev = 'Find previous occurrence'
          SearchFooter.HintHighlight = 'Highlight occurrences'
          SearchFooter.MatchCaseCaption = 'Match case'
          SearchFooter.ResultFormat = '(%d of %d)'
          SelectionColor = 16704470
          SelectionRectangle = True
          SelectionResizer = True
          SelectionRTFKeep = True
          ShowDesignHelper = False
          ShowFocusedSelectionColor = False
          SortSettings.DefaultFormat = ssAutomatic
          VAlignment = vtaCenter
          Version = '8.4.2.2'
          WordWrap = False
          ColWidths = (
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            563)
          RowHeights = (
            30
            30
            30
            35
            35
            35
            35
            35
            35
            35)
        end
      end
      object Panel5: TPanel
        Left = 0
        Top = 395
        Width = 1495
        Height = 350
        Align = alTop
        Caption = 'Panel5'
        TabOrder = 3
        Visible = False
        object asgDetail2: TAdvStringGrid
          Left = 1
          Top = 1
          Width = 1493
          Height = 348
          Cursor = crDefault
          TabStop = False
          Align = alClient
          BevelInner = bvNone
          BevelOuter = bvNone
          BorderStyle = bsNone
          ColCount = 32
          Ctl3D = False
          DefaultColWidth = 30
          DefaultRowHeight = 35
          DrawingStyle = gdsClassic
          FixedColor = 528832347
          FixedRows = 2
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = #45208#45588#44256#46357
          Font.Style = []
          Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goTabs]
          ParentCtl3D = False
          ParentFont = False
          ParentShowHint = False
          ScrollBars = ssBoth
          ShowHint = True
          TabOrder = 0
          Visible = False
          OnKeyDown = asgDetailKeyDown
          OnKeyPress = asgDetailKeyPress
          GridLineColor = 14013909
          GridFixedLineColor = 14013909
          HoverRowCells = [hcNormal, hcSelected]
          OnGetCellColor = asgDetailGetCellColor
          OnGetAlignment = asgDetailGetAlignment
          OnGetEditorType = asgDetailGetEditorType
          ActiveCellFont.Charset = DEFAULT_CHARSET
          ActiveCellFont.Color = clWindowText
          ActiveCellFont.Height = -11
          ActiveCellFont.Name = 'Tahoma'
          ActiveCellFont.Style = [fsBold]
          ActiveCellColor = clNone
          Bands.Active = True
          Bands.PrimaryColor = clWhite
          Bands.SecondaryColor = 16316664
          BorderColor = 15592941
          ColumnSize.Stretch = True
          ControlLook.FixedGradientFrom = 15658734
          ControlLook.FixedGradientTo = 15658734
          ControlLook.FixedGradientHoverFrom = clGray
          ControlLook.FixedGradientHoverTo = clWhite
          ControlLook.FixedGradientDownFrom = clGray
          ControlLook.FixedGradientDownTo = clSilver
          ControlLook.DropDownHeader.Font.Charset = DEFAULT_CHARSET
          ControlLook.DropDownHeader.Font.Color = clWindowText
          ControlLook.DropDownHeader.Font.Height = -11
          ControlLook.DropDownHeader.Font.Name = 'Tahoma'
          ControlLook.DropDownHeader.Font.Style = []
          ControlLook.DropDownHeader.Visible = True
          ControlLook.DropDownHeader.Buttons = <>
          ControlLook.DropDownFooter.Font.Charset = DEFAULT_CHARSET
          ControlLook.DropDownFooter.Font.Color = clWindowText
          ControlLook.DropDownFooter.Font.Height = -11
          ControlLook.DropDownFooter.Font.Name = 'Tahoma'
          ControlLook.DropDownFooter.Font.Style = []
          ControlLook.DropDownFooter.Visible = True
          ControlLook.DropDownFooter.Buttons = <>
          DefaultAlignment = taCenter
          EnhTextSize = True
          Filter = <>
          FilterDropDown.Font.Charset = DEFAULT_CHARSET
          FilterDropDown.Font.Color = clWindowText
          FilterDropDown.Font.Height = -11
          FilterDropDown.Font.Name = 'Tahoma'
          FilterDropDown.Font.Style = []
          FilterDropDown.TextChecked = 'Checked'
          FilterDropDown.TextUnChecked = 'Unchecked'
          FilterDropDownClear = '(All)'
          FilterEdit.TypeNames.Strings = (
            'Starts with'
            'Ends with'
            'Contains'
            'Not contains'
            'Equal'
            'Not equal'
            'Larger than'
            'Smaller than'
            'Clear')
          FixedColWidth = 30
          FixedRowHeight = 30
          FixedFont.Charset = DEFAULT_CHARSET
          FixedFont.Color = clBlack
          FixedFont.Height = -13
          FixedFont.Name = #45208#45588#44256#46357
          FixedFont.Style = []
          Flat = True
          FloatFormat = '%.2f'
          HideFocusRect = True
          HoverButtons.Buttons = <>
          HoverButtons.Position = hbLeftFromColumnLeft
          HTMLSettings.ImageFolder = 'images'
          HTMLSettings.ImageBaseName = 'img'
          MouseActions.DisjunctCellSelect = True
          MouseActions.RangeSelectAndEdit = True
          MouseActions.RowSelect = True
          Navigation.AdvanceOnEnter = True
          Navigation.AllowClipboardShortCuts = True
          Navigation.AllowSmartClipboard = True
          PrintSettings.DateFormat = 'dd/mm/yyyy'
          PrintSettings.Font.Charset = DEFAULT_CHARSET
          PrintSettings.Font.Color = clWindowText
          PrintSettings.Font.Height = -11
          PrintSettings.Font.Name = 'Tahoma'
          PrintSettings.Font.Style = []
          PrintSettings.FixedFont.Charset = DEFAULT_CHARSET
          PrintSettings.FixedFont.Color = clWindowText
          PrintSettings.FixedFont.Height = -11
          PrintSettings.FixedFont.Name = 'Tahoma'
          PrintSettings.FixedFont.Style = []
          PrintSettings.HeaderFont.Charset = DEFAULT_CHARSET
          PrintSettings.HeaderFont.Color = clWindowText
          PrintSettings.HeaderFont.Height = -11
          PrintSettings.HeaderFont.Name = 'Tahoma'
          PrintSettings.HeaderFont.Style = []
          PrintSettings.FooterFont.Charset = DEFAULT_CHARSET
          PrintSettings.FooterFont.Color = clWindowText
          PrintSettings.FooterFont.Height = -11
          PrintSettings.FooterFont.Name = 'Tahoma'
          PrintSettings.FooterFont.Style = []
          PrintSettings.PageNumSep = '/'
          SearchFooter.FindNextCaption = 'Find &next'
          SearchFooter.FindPrevCaption = 'Find &previous'
          SearchFooter.Font.Charset = DEFAULT_CHARSET
          SearchFooter.Font.Color = clWindowText
          SearchFooter.Font.Height = -11
          SearchFooter.Font.Name = 'Tahoma'
          SearchFooter.Font.Style = []
          SearchFooter.HighLightCaption = 'Highlight'
          SearchFooter.HintClose = 'Close'
          SearchFooter.HintFindNext = 'Find next occurrence'
          SearchFooter.HintFindPrev = 'Find previous occurrence'
          SearchFooter.HintHighlight = 'Highlight occurrences'
          SearchFooter.MatchCaseCaption = 'Match case'
          SearchFooter.ResultFormat = '(%d of %d)'
          SelectionColor = 16704470
          SelectionRectangle = True
          SelectionResizer = True
          SelectionRTFKeep = True
          ShowDesignHelper = False
          ShowFocusedSelectionColor = False
          SortSettings.DefaultFormat = ssAutomatic
          VAlignment = vtaCenter
          Version = '8.4.2.2'
          WordWrap = False
          ColWidths = (
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            30
            563)
          RowHeights = (
            30
            30
            30
            35
            35
            35
            35
            35
            35
            35)
        end
      end
      object Panel13: TPanel
        Left = 0
        Top = 0
        Width = 1495
        Height = 45
        Align = alTop
        BevelOuter = bvNone
        Color = 7362381
        Ctl3D = False
        ParentBackground = False
        ParentCtl3D = False
        TabOrder = 0
        object lblTitleDetailGrid: TLabel
          Left = 31
          Top = 11
          Width = 64
          Height = 20
          Caption = #49345#49464#51221#48372
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWhite
          Font.Height = -17
          Font.Name = #45208#45588#44256#46357
          Font.Style = [fsBold]
          ParentFont = False
        end
        object Image2: TImage
          Left = 15
          Top = 14
          Width = 9
          Height = 16
          Center = True
          Picture.Data = {
            0A544A504547496D616765B6050000FFD8FFE1001845786966000049492A0008
            0000000000000000000000FFEC00114475636B7900010004000000640000FFE1
            038D687474703A2F2F6E732E61646F62652E636F6D2F7861702F312E302F003C
            3F787061636B657420626567696E3D22EFBBBF222069643D2257354D304D7043
            656869487A7265537A4E54637A6B633964223F3E203C783A786D706D65746120
            786D6C6E733A783D2261646F62653A6E733A6D6574612F2220783A786D70746B
            3D2241646F626520584D5020436F726520352E362D633133382037392E313539
            3832342C20323031362F30392F31342D30313A30393A30312020202020202020
            223E203C7264663A52444620786D6C6E733A7264663D22687474703A2F2F7777
            772E77332E6F72672F313939392F30322F32322D7264662D73796E7461782D6E
            7323223E203C7264663A4465736372697074696F6E207264663A61626F75743D
            222220786D6C6E733A786D704D4D3D22687474703A2F2F6E732E61646F62652E
            636F6D2F7861702F312E302F6D6D2F2220786D6C6E733A73745265663D226874
            74703A2F2F6E732E61646F62652E636F6D2F7861702F312E302F73547970652F
            5265736F75726365526566232220786D6C6E733A786D703D22687474703A2F2F
            6E732E61646F62652E636F6D2F7861702F312E302F2220786D704D4D3A4F7269
            67696E616C446F63756D656E7449443D22786D702E6469643A37666466663838
            372D336365632D313334622D623064322D376533633032646664343132222078
            6D704D4D3A446F63756D656E7449443D22786D702E6469643A44343343314530
            303732434231314539424435353845324546314146304342412220786D704D4D
            3A496E7374616E636549443D22786D702E6969643A4434334331444646373243
            4231314539424435353845324546314146304342412220786D703A4372656174
            6F72546F6F6C3D2241646F62652050686F746F73686F70204343203230313720
            2857696E646F777329223E203C786D704D4D3A4465726976656446726F6D2073
            745265663A696E7374616E636549443D22786D702E6969643A61303332303666
            362D366365622D663634612D396262342D613337393034333836363061222073
            745265663A646F63756D656E7449443D2261646F62653A646F6369643A70686F
            746F73686F703A34636439363735632D366166652D313165392D613162362D63
            6165656361363963303965222F3E203C2F7264663A4465736372697074696F6E
            3E203C2F7264663A5244463E203C2F783A786D706D6574613E203C3F78706163
            6B657420656E643D2272223F3EFFEE000E41646F62650064C000000001FFDB00
            8400010101010101010101010101010101010101010101010101010101010101
            0101010101010101010101010102020202020202020202020303030303030303
            0303010101010101010201010202020102020303030303030303030303030303
            0303030303030303030303030303030303030303030303030303030303030303
            030303FFC00011080009000903011100021101031101FFC40078000101010000
            0000000000000000000000080109010002030000000000000000000000000008
            0900020A100000060101090000000000000000000001020304050600071172D3
            B49556170838110001030301040A030000000000000000010203040006070511
            12D294217172B2B35516561708380939FFDA000C03010002110311003F006ED2
            E97A628E9943D7E021EBEFA80FABED14211468C9C47CC3070C88A1E4A48CA10C
            9BC72F086151659411389C444476E66F320641CC123304FB9EE79FAA46C9F1B5
            47524A5D750F45790E9018600214D36D101B69A6C0404000020F4B74B56D4B01
            AB022E8BA2C584F596F424100A10A6DF6D4804BAE92085A960EF2D6ADAADE249
            3B6809E36F48BBD49D79C71B19B7CB7FB13F6F2B936F82833F41FD4AF361CC2B
            8AA69B7C4175DCB0732BE5F2D7F456DEED42F0D152C4FC49D5BAA4779559A78D
            D2809AFFD9}
        end
      end
      object pnlPrepareDetailGrid: TPanel
        Left = 0
        Top = 696
        Width = 1495
        Height = 126
        Align = alBottom
        BevelOuter = bvNone
        Caption = #54616#44592#55092#44032#44228#54925#49436#47484' '#51312#54924#54644#51452#49901#49884#50724'.'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -20
        Font.Name = #45208#45588#44256#46357
        Font.Style = []
        ParentFont = False
        TabOrder = 1
      end
    end
    object Panel8: TPanel
      Left = 0
      Top = 0
      Width = 377
      Height = 822
      Align = alLeft
      BevelOuter = bvNone
      Color = clWhite
      Ctl3D = False
      ParentBackground = False
      ParentCtl3D = False
      TabOrder = 1
      object Shape13: TShape
        Left = 0
        Top = 105
        Width = 1
        Height = 716
        Align = alLeft
        Brush.Color = 15592941
        Pen.Color = 14540253
        ExplicitLeft = 8
        ExplicitTop = 45
        ExplicitHeight = 704
      end
      object Shape15: TShape
        Left = 0
        Top = 821
        Width = 377
        Height = 1
        Align = alBottom
        Brush.Color = 15592941
        Pen.Color = 14540253
        ExplicitLeft = 1
        ExplicitTop = 45
        ExplicitWidth = 704
      end
      object Panel9: TPanel
        Left = 0
        Top = 0
        Width = 377
        Height = 45
        Align = alTop
        BevelOuter = bvNone
        Color = 7362381
        Ctl3D = False
        ParentBackground = False
        ParentCtl3D = False
        TabOrder = 0
        object Image3: TImage
          Left = 13
          Top = 14
          Width = 9
          Height = 16
          Center = True
          Picture.Data = {
            0A544A504547496D616765A2050000FFD8FFE1001845786966000049492A0008
            0000000000000000000000FFEC00114475636B7900010004000000640000FFE1
            038D687474703A2F2F6E732E61646F62652E636F6D2F7861702F312E302F003C
            3F787061636B657420626567696E3D22EFBBBF222069643D2257354D304D7043
            656869487A7265537A4E54637A6B633964223F3E203C783A786D706D65746120
            786D6C6E733A783D2261646F62653A6E733A6D6574612F2220783A786D70746B
            3D2241646F626520584D5020436F726520352E362D633133382037392E313539
            3832342C20323031362F30392F31342D30313A30393A30312020202020202020
            223E203C7264663A52444620786D6C6E733A7264663D22687474703A2F2F7777
            772E77332E6F72672F313939392F30322F32322D7264662D73796E7461782D6E
            7323223E203C7264663A4465736372697074696F6E207264663A61626F75743D
            222220786D6C6E733A786D704D4D3D22687474703A2F2F6E732E61646F62652E
            636F6D2F7861702F312E302F6D6D2F2220786D6C6E733A73745265663D226874
            74703A2F2F6E732E61646F62652E636F6D2F7861702F312E302F73547970652F
            5265736F75726365526566232220786D6C6E733A786D703D22687474703A2F2F
            6E732E61646F62652E636F6D2F7861702F312E302F2220786D704D4D3A4F7269
            67696E616C446F63756D656E7449443D22786D702E6469643A37666466663838
            372D336365632D313334622D623064322D376533633032646664343132222078
            6D704D4D3A446F63756D656E7449443D22786D702E6469643A44343343314446
            433732434231314539424435353845324546314146304342412220786D704D4D
            3A496E7374616E636549443D22786D702E6969643A4434334331444642373243
            4231314539424435353845324546314146304342412220786D703A4372656174
            6F72546F6F6C3D2241646F62652050686F746F73686F70204343203230313720
            2857696E646F777329223E203C786D704D4D3A4465726976656446726F6D2073
            745265663A696E7374616E636549443D22786D702E6969643A61303332303666
            362D366365622D663634612D396262342D613337393034333836363061222073
            745265663A646F63756D656E7449443D2261646F62653A646F6369643A70686F
            746F73686F703A34636439363735632D366166652D313165392D613162362D63
            6165656361363963303965222F3E203C2F7264663A4465736372697074696F6E
            3E203C2F7264663A5244463E203C2F783A786D706D6574613E203C3F78706163
            6B657420656E643D2272223F3EFFEE000E41646F62650064C000000001FFDB00
            8400010101010101010101010101010101010101010101010101010101010101
            0101010101010101010101010102020202020202020202020303030303030303
            0303010101010101010201010202020102020303030303030303030303030303
            0303030303030303030303030303030303030303030303030303030303030303
            030303FFC00011080009000903011100021101031101FFC40073000003000000
            0000000000000000000000010508010002030000000000000000000000000008
            0900010A1000010402000603000000000000000000020103040500061272B3D3
            5595160737110002010204030900000000000000000001020300061104050771
            B255D122D2939416170838FFDA000C03010002110311003F00454745AA86AB0A
            BABA156C8D72456B24884CB0E479B1DC6108A54A2245179D785548CC954B8955
            5571E3DBF6FD9F1D9F0699A6419592D8932AA402A8C92A320264909183330EF3
            33627124938D65DEF8BE374E7DD3CEDC57167751CBEE241A8C809124A9365A54
            94810C20106348C8091C6802850001854EFF0017FA1BCF0FB177B98327B4BEBB
            7511E71EDA61BF277DEFE84DE913C34358FC1AF796C7A8EE4B4BF3B6A1C26E66
            ABDCEFDDDA171CA722D4B9822D342AFFD9}
        end
        object Label2: TLabel
          Left = 29
          Top = 11
          Width = 64
          Height = 20
          Caption = #48512#49436#49440#53469
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWhite
          Font.Height = -17
          Font.Name = #45208#45588#44256#46357
          Font.Style = [fsBold]
          ParentFont = False
        end
      end
      object ScrollBox2: TScrollBox
        Left = 0
        Top = 45
        Width = 377
        Height = 60
        Align = alTop
        BevelInner = bvNone
        BevelOuter = bvNone
        BorderStyle = bsNone
        Color = 16448250
        Ctl3D = False
        ParentColor = False
        ParentCtl3D = False
        TabOrder = 1
        object Shape1: TShape
          Left = 0
          Top = 1
          Width = 1
          Height = 58
          Align = alLeft
          Brush.Color = 15592941
          Pen.Color = 14540253
          ExplicitLeft = 8
          ExplicitTop = 0
          ExplicitHeight = 66
        end
        object Shape2: TShape
          Left = 376
          Top = 1
          Width = 1
          Height = 58
          Align = alRight
          Brush.Color = 15592941
          Pen.Color = 14540253
          ExplicitLeft = 8
          ExplicitTop = 0
          ExplicitHeight = 66
        end
        object Shape3: TShape
          Left = 0
          Top = 0
          Width = 377
          Height = 1
          Align = alTop
          Brush.Color = 15592941
          Pen.Color = 14540253
          ExplicitLeft = 1
          ExplicitWidth = 66
        end
        object Shape11: TShape
          Left = 0
          Top = 59
          Width = 377
          Height = 1
          Align = alBottom
          Brush.Color = 15592941
          Pen.Color = 14540253
          ExplicitTop = 1
          ExplicitWidth = 65
        end
        object Label3: TLabel
          Left = 26
          Top = 20
          Width = 48
          Height = 15
          Caption = #44228#54925#45380#50900
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = #45208#45588#44256#46357
          Font.Style = []
          ParentFont = False
        end
        object Label5: TLabel
          Left = 149
          Top = 20
          Width = 12
          Height = 15
          Caption = #45380
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = #45208#45588#44256#46357
          Font.Style = []
          ParentFont = False
        end
        object Label8: TLabel
          Left = 234
          Top = 20
          Width = 48
          Height = 15
          Caption = #48512'      '#49436
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = #45208#45588#44256#46357
          Font.Style = []
          ParentFont = False
          Visible = False
        end
        object cbxBFYear: TComboBox
          Left = 85
          Top = 16
          Width = 58
          Height = 22
          AutoDropDown = True
          AutoCloseUp = True
          Style = csOwnerDrawFixed
          DropDownCount = 10
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = #45208#45588#44256#46357
          Font.Style = []
          ParentFont = False
          TabOrder = 0
          TextHint = #45380
          OnChange = cbxConDateChange
          Items.Strings = (
            '2018'
            '2019'
            '2020')
        end
        object cbxSDept: TComboBox
          Left = 293
          Top = 14
          Width = 125
          Height = 22
          AutoDropDown = True
          AutoCloseUp = True
          Style = csOwnerDrawFixed
          DropDownCount = 20
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = #45208#45588#44256#46357
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          TextHint = #48512#49436#47484' '#49440#53469#54616#49464#50836
          Visible = False
          OnChange = ChangeEnrollInfo
          Items.Strings = (
            '2018'
            '2019'
            '2020')
        end
      end
      object asgMaster: TAdvStringGrid
        Left = 1
        Top = 105
        Width = 376
        Height = 716
        Cursor = crDefault
        TabStop = False
        Align = alClient
        BevelInner = bvNone
        BevelOuter = bvNone
        BorderStyle = bsNone
        ColCount = 2
        Ctl3D = False
        DefaultRowHeight = 30
        DrawingStyle = gdsClassic
        FixedColor = 528832347
        FixedCols = 0
        RowCount = 2
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #45208#45588#44256#46357
        Font.Style = []
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goRowSelect]
        ParentCtl3D = False
        ParentFont = False
        ParentShowHint = False
        ScrollBars = ssBoth
        ShowHint = False
        TabOrder = 2
        GridLineColor = 14013909
        GridFixedLineColor = 14013909
        HoverRow = True
        HoverRowCells = [hcNormal, hcSelected]
        OnClickCell = asgMasterClickCell
        ActiveCellFont.Charset = DEFAULT_CHARSET
        ActiveCellFont.Color = clWindowText
        ActiveCellFont.Height = -11
        ActiveCellFont.Name = 'Tahoma'
        ActiveCellFont.Style = [fsBold]
        Bands.Active = True
        Bands.PrimaryColor = clWhite
        Bands.SecondaryColor = 16316664
        BorderColor = 15592941
        ColumnHeaders.Strings = (
          #49692#48264
          #48512#49436)
        ColumnSize.Stretch = True
        ControlLook.FixedGradientFrom = 15658734
        ControlLook.FixedGradientTo = 15658734
        ControlLook.FixedGradientHoverFrom = clGray
        ControlLook.FixedGradientHoverTo = clWhite
        ControlLook.FixedGradientDownFrom = clGray
        ControlLook.FixedGradientDownTo = clSilver
        ControlLook.DropDownHeader.Font.Charset = DEFAULT_CHARSET
        ControlLook.DropDownHeader.Font.Color = clWindowText
        ControlLook.DropDownHeader.Font.Height = -11
        ControlLook.DropDownHeader.Font.Name = 'Tahoma'
        ControlLook.DropDownHeader.Font.Style = []
        ControlLook.DropDownHeader.Visible = True
        ControlLook.DropDownHeader.Buttons = <>
        ControlLook.DropDownFooter.Font.Charset = DEFAULT_CHARSET
        ControlLook.DropDownFooter.Font.Color = clWindowText
        ControlLook.DropDownFooter.Font.Height = -11
        ControlLook.DropDownFooter.Font.Name = 'Tahoma'
        ControlLook.DropDownFooter.Font.Style = []
        ControlLook.DropDownFooter.Visible = True
        ControlLook.DropDownFooter.Buttons = <>
        DefaultAlignment = taCenter
        EnhTextSize = True
        Filter = <>
        FilterDropDown.Font.Charset = DEFAULT_CHARSET
        FilterDropDown.Font.Color = clWindowText
        FilterDropDown.Font.Height = -11
        FilterDropDown.Font.Name = 'Tahoma'
        FilterDropDown.Font.Style = []
        FilterDropDown.TextChecked = 'Checked'
        FilterDropDown.TextUnChecked = 'Unchecked'
        FilterDropDownClear = '(All)'
        FilterEdit.TypeNames.Strings = (
          'Starts with'
          'Ends with'
          'Contains'
          'Not contains'
          'Equal'
          'Not equal'
          'Larger than'
          'Smaller than'
          'Clear')
        FixedColWidth = 63
        FixedRowHeight = 30
        FixedFont.Charset = DEFAULT_CHARSET
        FixedFont.Color = clBlack
        FixedFont.Height = -13
        FixedFont.Name = #45208#45588#44256#46357
        FixedFont.Style = []
        Flat = True
        FloatFormat = '%.2f'
        HideFocusRect = True
        HoverButtons.Buttons = <>
        HoverButtons.Position = hbLeftFromColumnLeft
        HTMLSettings.ImageFolder = 'images'
        HTMLSettings.ImageBaseName = 'img'
        MouseActions.RowSelect = True
        PrintSettings.DateFormat = 'dd/mm/yyyy'
        PrintSettings.Font.Charset = DEFAULT_CHARSET
        PrintSettings.Font.Color = clWindowText
        PrintSettings.Font.Height = -11
        PrintSettings.Font.Name = 'Tahoma'
        PrintSettings.Font.Style = []
        PrintSettings.FixedFont.Charset = DEFAULT_CHARSET
        PrintSettings.FixedFont.Color = clWindowText
        PrintSettings.FixedFont.Height = -11
        PrintSettings.FixedFont.Name = 'Tahoma'
        PrintSettings.FixedFont.Style = []
        PrintSettings.HeaderFont.Charset = DEFAULT_CHARSET
        PrintSettings.HeaderFont.Color = clWindowText
        PrintSettings.HeaderFont.Height = -11
        PrintSettings.HeaderFont.Name = 'Tahoma'
        PrintSettings.HeaderFont.Style = []
        PrintSettings.FooterFont.Charset = DEFAULT_CHARSET
        PrintSettings.FooterFont.Color = clWindowText
        PrintSettings.FooterFont.Height = -11
        PrintSettings.FooterFont.Name = 'Tahoma'
        PrintSettings.FooterFont.Style = []
        PrintSettings.PageNumSep = '/'
        SearchFooter.FindNextCaption = 'Find &next'
        SearchFooter.FindPrevCaption = 'Find &previous'
        SearchFooter.Font.Charset = DEFAULT_CHARSET
        SearchFooter.Font.Color = clWindowText
        SearchFooter.Font.Height = -11
        SearchFooter.Font.Name = 'Tahoma'
        SearchFooter.Font.Style = []
        SearchFooter.HighLightCaption = 'Highlight'
        SearchFooter.HintClose = 'Close'
        SearchFooter.HintFindNext = 'Find next occurrence'
        SearchFooter.HintFindPrev = 'Find previous occurrence'
        SearchFooter.HintHighlight = 'Highlight occurrences'
        SearchFooter.MatchCaseCaption = 'Match case'
        SearchFooter.ResultFormat = '(%d of %d)'
        SelectionColor = 16704470
        ShowDesignHelper = False
        ShowFocusedSelectionColor = False
        SortSettings.DefaultFormat = ssAutomatic
        VAlignment = vtaCenter
        Version = '8.4.2.2'
        WordWrap = False
        ColWidths = (
          63
          313)
      end
    end
  end
  inherited Query1: TFDQuery
    FetchOptions.AssignedValues = [evItems, evRowsetSize, evCache, evCursorKind]
    FetchOptions.RowsetSize = 200
    FetchOptions.Items = [fiBlobs, fiDetails]
    FetchOptions.Cache = [fiBlobs, fiDetails]
  end
  inherited Query2: TFDQuery
    FetchOptions.AssignedValues = [evItems, evRowsetSize, evCache, evCursorKind]
    FetchOptions.RowsetSize = 200
    FetchOptions.Items = [fiBlobs, fiDetails]
    FetchOptions.Cache = [fiBlobs, fiDetails]
  end
  object popMenu: TPopupMenu
    AutoHotkeys = maManual
    Left = 872
    Top = 20
    object N00: TMenuItem
      Caption = #51221#49345
      object s1: TMenuItem
        Caption = 'A  08:00~17:00'
      end
      object s2: TMenuItem
        Caption = 'B  08:30~17:30'
      end
      object s3: TMenuItem
        Caption = 'C  09:00~18:00'
      end
      object s4: TMenuItem
        Caption = 'D  09:30~18:30'
      end
      object s5: TMenuItem
        Caption = 'E  10:00~19:00'
      end
      object s6: TMenuItem
        Caption = 'F  10:30~19:30'
      end
      object s7: TMenuItem
        Caption = 'G  11:00~20:00'
      end
      object xxxxx: TMenuItem
        Caption = '-----------------'
      end
      object p1: TMenuItem
        Caption = 'H  13:00~19:30'
      end
      object p2: TMenuItem
        Caption = 'I  13:30~20:00'
      end
      object p3: TMenuItem
        Caption = 'J  14:00~18:00'
      end
      object p4: TMenuItem
        Caption = 'K  08:00~19:00'
      end
      object p5: TMenuItem
        Caption = 'L  08:00~19:30'
      end
      object p6: TMenuItem
        Caption = '------------------'
      end
      object p7: TMenuItem
        Caption = 'N  09:00~15:00'
      end
      object p8: TMenuItem
        Caption = 'O  09:00~12:00'
      end
      object p9: TMenuItem
        Caption = 'P  10:00~15:00'
      end
      object p10: TMenuItem
        Caption = 'Q  09:00~19:00'
      end
      object p15: TMenuItem
        Caption = 'R  10:00~19:30'
      end
      object p11: TMenuItem
        Caption = #9424'  08:00~19:00'
      end
      object p12: TMenuItem
        Caption = #9425'  08:30~19:00'
      end
      object p13: TMenuItem
        Caption = #9426'  09:00~19:00'
      end
      object p14: TMenuItem
        Caption = #9427'  09:30~19:00'
      end
    end
    object N01: TMenuItem
      Caption = #45380#52264
    end
    object N02: TMenuItem
      Caption = #48152#52264
    end
    object N1: TMenuItem
      Caption = #48152#52264'('#20195')'
    end
    object N05: TMenuItem
      Caption = #45824#55092
    end
    object N04: TMenuItem
      Caption = #53945#55092
    end
    object N07: TMenuItem
      Caption = #44221#51312
    end
    object N08: TMenuItem
      Caption = #52636#51109
    end
    object N06: TMenuItem
      Caption = #48372#44148
    end
    object N12: TMenuItem
      Caption = #51204#51076
    end
    object N13: TMenuItem
      Caption = #44368#50977
    end
    object N15: TMenuItem
      Caption = #51221#51648
    end
  end
end
