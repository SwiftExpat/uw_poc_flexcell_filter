object frmExcelFilterExport: TfrmExcelFilterExport
  Left = 0
  Top = 0
  Caption = 'Excel Filter Export'
  ClientHeight = 516
  ClientWidth = 712
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  OnDestroy = FormDestroy
  OnShow = FormShow
  TextHeight = 15
  object pcMain: TTMSFNCPageControl
    Left = 0
    Top = 0
    Width = 712
    Height = 496
    Align = alClient
    ParentDoubleBuffered = False
    DoubleBuffered = True
    ParentColor = True
    TabOrder = 0
    Fill.Kind = gfkNone
    Stroke.Kind = gskNone
    TabAppearance.Font.Charset = DEFAULT_CHARSET
    TabAppearance.Font.Color = clWindowText
    TabAppearance.Font.Height = -11
    TabAppearance.Font.Name = 'Segoe UI'
    TabAppearance.Font.Style = []
    TabAppearance.Fill.Color = 16578806
    TabAppearance.DownFill.Color = 15776936
    TabAppearance.DownStroke.Color = 15702829
    TabAppearance.HoverFill.Color = 16380654
    TabAppearance.HoverStroke.Color = 15702829
    TabAppearance.TextSpacing = 0.000000000000000000
    TabAppearance.TextColor = 8026746
    TabAppearance.ActiveTextColor = 4539717
    TabAppearance.HoverTextColor = 8026746
    TabAppearance.DownTextColor = 4539717
    TabAppearance.ShowFocus = False
    TabAppearance.CloseStroke.Width = 2.000000000000000000
    TabAppearance.CloseDownStroke.Width = 2.000000000000000000
    TabAppearance.CloseHoverStroke.Width = 2.000000000000000000
    TabAppearance.BadgeFont.Charset = DEFAULT_CHARSET
    TabAppearance.BadgeFont.Color = 139
    TabAppearance.BadgeFont.Height = -11
    TabAppearance.BadgeFont.Name = 'Segoe UI'
    TabAppearance.BadgeFont.Style = [fsBold]
    ButtonAppearance.CloseIcon.Data = {
      1054544D53464E435356474269746D6170EF0000003C7376672077696474683D
      223234707822206865696768743D2232347078222076696577426F783D223020
      302032342032342220786D6C6E733D22687474703A2F2F7777772E77332E6F72
      672F323030302F737667222066696C6C3D226E6F6E65223E3C70617468207374
      726F6B653D2263757272656E74436F6C6F7222207374726F6B652D6C696E6563
      61703D22726F756E6422207374726F6B652D6C696E656A6F696E3D22726F756E
      6422207374726F6B652D77696474683D22322220643D224D3132203132203720
      376D352035203520356D2D352D3520352D356D2D3520352D352035222F3E3C2F
      7376673E}
    ButtonAppearance.InsertIcon.Data = {
      1054544D53464E435356474269746D6170DF0200003C3F786D6C207665727369
      6F6E3D22312E302220656E636F64696E673D225554462D38223F3E0A3C737667
      20786D6C6E733D22687474703A2F2F7777772E77332E6F72672F323030302F73
      76672220786D6C6E733A786C696E6B3D22687474703A2F2F7777772E77332E6F
      72672F313939392F786C696E6B222077696474683D2232347078222068656967
      68743D2232347078222076696577426F783D2230203020323420323422207665
      7273696F6E3D22312E31223E0A3C672069643D227375726661636531223E0A3C
      70617468207374796C653D22207374726F6B653A6E6F6E653B66696C6C2D7275
      6C653A6E6F6E7A65726F3B66696C6C3A7267622830252C30252C3025293B6669
      6C6C2D6F7061636974793A313B2220643D224D203132203620432031322E3431
      3430363220362031322E373520362E3333353933382031322E373520362E3735
      204C2031322E37352031312E3235204C2031372E32352031312E323520432031
      372E3636343036322031312E32352031382031312E3538353933382031382031
      3220432031382031322E3431343036322031372E3636343036322031322E3735
      2031372E32352031322E3735204C2031322E37352031322E3735204C2031322E
      37352031372E323520432031322E37352031372E3636343036322031322E3431
      3430363220313820313220313820432031312E3538353933382031382031312E
      32352031372E3636343036322031312E32352031372E3235204C2031312E3235
      2031322E3735204C20362E37352031322E3735204320362E3333353933382031
      322E373520362031322E3431343036322036203132204320362031312E353835
      39333820362E3333353933382031312E323520362E37352031312E3235204C20
      31312E32352031312E3235204C2031312E323520362E373520432031312E3235
      20362E3333353933382031312E35383539333820362031322036205A204D2031
      32203620222F3E0A3C2F673E0A3C2F7376673E0A}
    ButtonAppearance.TabListIcon.Data = {
      1054544D53464E435356474269746D6170D90000003C7376672077696474683D
      223234707822206865696768743D2232347078222076696577426F783D223020
      302032342032342220786D6C6E733D22687474703A2F2F7777772E77332E6F72
      672F323030302F737667223E3C70617468207374726F6B653D2263757272656E
      74436F6C6F7222207374726F6B652D6C696E656361703D22726F756E64222073
      74726F6B652D6C696E656A6F696E3D22726F756E6422207374726F6B652D7769
      6474683D22322220643D226D31372031302D3520352D352D35222066696C6C3D
      226E6F6E65222F3E3C2F7376673E}
    ButtonAppearance.ScrollNextIcon.Data = {
      1054544D53464E435356474269746D6170D80000003C7376672077696474683D
      223234707822206865696768743D2232347078222076696577426F783D223020
      302032342032342220786D6C6E733D22687474703A2F2F7777772E77332E6F72
      672F323030302F737667223E3C706174682066696C6C3D226E6F6E6522207374
      726F6B653D2263757272656E74436F6C6F7222207374726F6B652D6C696E6563
      61703D22726F756E6422207374726F6B652D6C696E656A6F696E3D22726F756E
      6422207374726F6B652D77696474683D22322220643D226D3130203720352035
      2D352035222F3E3C2F7376673E}
    ButtonAppearance.ScrollPreviousIcon.Data = {
      1054544D53464E435356474269746D6170D80000003C7376672077696474683D
      223234707822206865696768743D2232347078222076696577426F783D223020
      302032342032342220786D6C6E733D22687474703A2F2F7777772E77332E6F72
      672F323030302F737667223E3C706174682066696C6C3D226E6F6E6522207374
      726F6B653D2263757272656E74436F6C6F7222207374726F6B652D6C696E6563
      61703D22726F756E6422207374726F6B652D6C696E656A6F696E3D22726F756E
      6422207374726F6B652D77696474683D22322220643D226D313420372D352035
      20352035222F3E3C2F7376673E}
    TabSize.Margins.Left = 8.000000000000000000
    TabSize.Margins.Top = 8.000000000000000000
    TabSize.Margins.Right = 8.000000000000000000
    TabSize.Margins.Bottom = 8.000000000000000000
    Pages = <
      item
        Text = 'Data'
        Bitmaps = <>
        DisabledBitmaps = <>
        UseDefaultAppearance = True
      end
      item
        Text = 'Logger'
        Bitmaps = <>
        DisabledBitmaps = <>
        UseDefaultAppearance = True
      end>
    object TMSFNCPageControl1Page1: TTMSFNCPageControlContainer
      AlignWithMargins = True
      Left = 0
      Top = 29
      Width = 712
      Height = 467
      Margins.Left = 0
      Margins.Top = 29
      Margins.Right = 0
      Margins.Bottom = 0
      Align = alClient
      ParentDoubleBuffered = False
      DoubleBuffered = True
      TabStop = False
      TabOrder = 1
      PageIndex = 1
      IsActive = False
      object Memo1: TMemo
        Left = 0
        Top = 0
        Width = 712
        Height = 467
        Align = alClient
        Lines.Strings = (
          'Memo1')
        TabOrder = 0
      end
    end
    object TMSFNCPageControl1Page0: TTMSFNCPageControlContainer
      AlignWithMargins = True
      Left = 0
      Top = 29
      Width = 712
      Height = 467
      Margins.Left = 0
      Margins.Top = 29
      Margins.Right = 0
      Margins.Bottom = 0
      Align = alClient
      ParentDoubleBuffered = False
      DoubleBuffered = True
      TabStop = False
      TabOrder = 0
      PageIndex = 0
      IsActive = True
      object TMSFNCToolBar1: TTMSFNCToolBar
        Left = 0
        Top = 0
        Width = 712
        Height = 30
        Align = alTop
        ParentDoubleBuffered = False
        DoubleBuffered = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Segoe UI'
        Font.Style = []
        TabOrder = 0
        CompactBitmaps = <>
        Text = ''
        object btnFileOpen: TTMSFNCToolBarButton
          Left = 12
          Top = 3
          Width = 30
          Height = 24
          ParentDoubleBuffered = False
          DoubleBuffered = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Segoe UI'
          Font.Style = []
          ParentColor = True
          TabOrder = 0
          OnClick = btnFileOpenClick
          Text = 'Open'
          Bitmaps = <>
          LargeLayoutBitmaps = <>
          DisabledBitmaps = <>
          HoverBitmaps = <>
          LargeLayoutDisabledBitmaps = <>
          LargeLayoutHoverBitmaps = <>
          ControlIndex = 0
        end
        object btnSave: TTMSFNCToolBarButton
          Left = 45
          Top = 3
          Width = 30
          Height = 24
          ParentDoubleBuffered = False
          DoubleBuffered = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Segoe UI'
          Font.Style = []
          ParentColor = True
          TabOrder = 1
          OnClick = btnSaveClick
          Text = 'Save'
          Bitmaps = <>
          LargeLayoutBitmaps = <>
          DisabledBitmaps = <>
          HoverBitmaps = <>
          LargeLayoutDisabledBitmaps = <>
          LargeLayoutHoverBitmaps = <>
          ControlIndex = 1
        end
        object TMSFNCToolBarSeparator1: TTMSFNCToolBarSeparator
          Left = 78
          Top = 3
          Width = 3
          Height = 24
          ParentDoubleBuffered = False
          DoubleBuffered = True
          ParentColor = True
          TabOrder = 3
          ControlIndex = 2
        end
        object cbSheets: TTMSFNCToolBarItemPicker
          Left = 84
          Top = 3
          Width = 100
          Height = 24
          ParentDoubleBuffered = False
          DoubleBuffered = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Segoe UI'
          Font.Style = []
          ParentColor = True
          TabOrder = 2
          DropDownKind = ddkDropDownButton
          DropDownHeight = 150.000000000000000000
          DropDownWidth = 200.000000000000000000
          Text = ''
          HorizontalTextAlign = gtaLeading
          Bitmaps = <>
          LargeLayoutBitmaps = <>
          DisabledBitmaps = <>
          HoverBitmaps = <>
          LargeLayoutDisabledBitmaps = <>
          LargeLayoutHoverBitmaps = <>
          OnItemSelected = cbSheetsItemSelected
          ControlIndex = 3
        end
      end
      object DataGrid: TTMSFNCGrid
        Left = 0
        Top = 30
        Width = 712
        Height = 437
        Align = alClient
        ParentDoubleBuffered = False
        DoubleBuffered = True
        TabOrder = 1
        DefaultRowHeight = 40.000000000000000000
        FixedColumns = 0
        ColumnCount = 4
        Options.Bands.Enabled = True
        Options.Editing.CalcFormat = '%g'
        Options.Filtering.DropDown = True
        Options.Filtering.MultiColumn = True
        Options.Grouping.CalcFormat = '%g'
        Options.Grouping.GroupCountFormat = '(%d)'
        Options.IO.XMLEncoding = 'ISO-8859-1'
        Options.Mouse.ClickMargin = 0
        Options.Mouse.ColumnSizeMargin = 6
        Options.Mouse.RowSizeMargin = 6
        Columns = <
          item
            BorderWidth = 1
            FixedFont.Charset = DEFAULT_CHARSET
            FixedFont.Color = clWindowText
            FixedFont.Height = -11
            FixedFont.Name = 'Segoe UI'
            FixedFont.Style = [fsBold]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Segoe UI'
            Font.Style = []
            ID = ''
            Width = 70.000000000000000000
          end
          item
            BorderWidth = 1
            FixedFont.Charset = DEFAULT_CHARSET
            FixedFont.Color = clWindowText
            FixedFont.Height = -11
            FixedFont.Name = 'Segoe UI'
            FixedFont.Style = [fsBold]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Segoe UI'
            Font.Style = []
            ID = ''
            Width = 250.000000000000000000
          end
          item
            BorderWidth = 1
            FixedFont.Charset = DEFAULT_CHARSET
            FixedFont.Color = clWindowText
            FixedFont.Height = -11
            FixedFont.Name = 'Segoe UI'
            FixedFont.Style = [fsBold]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Segoe UI'
            Font.Style = []
            ID = ''
            Width = 100.000000000000000000
          end
          item
            BorderWidth = 1
            FixedFont.Charset = DEFAULT_CHARSET
            FixedFont.Color = clWindowText
            FixedFont.Height = -11
            FixedFont.Name = 'Segoe UI'
            FixedFont.Style = [fsBold]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Segoe UI'
            Font.Style = []
            ID = ''
            Width = 100.000000000000000000
          end
          item
            BorderWidth = 1
            FixedFont.Charset = DEFAULT_CHARSET
            FixedFont.Color = clWindowText
            FixedFont.Height = -11
            FixedFont.Name = 'Segoe UI'
            FixedFont.Style = [fsBold]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Segoe UI'
            Font.Style = []
            ID = ''
            Width = 90.000000000000000000
          end
          item
            BorderWidth = 1
            FixedFont.Charset = DEFAULT_CHARSET
            FixedFont.Color = clWindowText
            FixedFont.Height = -11
            FixedFont.Name = 'Segoe UI'
            FixedFont.Style = [fsBold]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Segoe UI'
            Font.Style = []
            ID = ''
            Width = 68.000000000000000000
          end
          item
            BorderWidth = 1
            FixedFont.Charset = DEFAULT_CHARSET
            FixedFont.Color = clWindowText
            FixedFont.Height = -11
            FixedFont.Name = 'Segoe UI'
            FixedFont.Style = [fsBold]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Segoe UI'
            Font.Style = []
            ID = ''
            Width = 68.000000000000000000
          end
          item
            BorderWidth = 1
            FixedFont.Charset = DEFAULT_CHARSET
            FixedFont.Color = clWindowText
            FixedFont.Height = -11
            FixedFont.Name = 'Segoe UI'
            FixedFont.Style = [fsBold]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Segoe UI'
            Font.Style = []
            ID = ''
            Width = 68.000000000000000000
          end
          item
            BorderWidth = 1
            FixedFont.Charset = DEFAULT_CHARSET
            FixedFont.Color = clWindowText
            FixedFont.Height = -11
            FixedFont.Name = 'Segoe UI'
            FixedFont.Style = [fsBold]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Segoe UI'
            Font.Style = []
            ID = ''
            Width = 68.000000000000000000
          end
          item
            BorderWidth = 1
            FixedFont.Charset = DEFAULT_CHARSET
            FixedFont.Color = clWindowText
            FixedFont.Height = -11
            FixedFont.Name = 'Segoe UI'
            FixedFont.Style = [fsBold]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Segoe UI'
            Font.Style = []
            ID = ''
            Width = 68.000000000000000000
          end
          item
            BorderWidth = 1
            FixedFont.Charset = DEFAULT_CHARSET
            FixedFont.Color = clWindowText
            FixedFont.Height = -11
            FixedFont.Name = 'Segoe UI'
            FixedFont.Style = [fsBold]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Segoe UI'
            Font.Style = []
            ID = ''
            Width = 68.000000000000000000
          end
          item
            BorderWidth = 1
            FixedFont.Charset = DEFAULT_CHARSET
            FixedFont.Color = clWindowText
            FixedFont.Height = -11
            FixedFont.Name = 'Segoe UI'
            FixedFont.Style = [fsBold]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Segoe UI'
            Font.Style = []
            ID = ''
            Width = 68.000000000000000000
          end
          item
            BorderWidth = 1
            FixedFont.Charset = DEFAULT_CHARSET
            FixedFont.Color = clWindowText
            FixedFont.Height = -11
            FixedFont.Name = 'Segoe UI'
            FixedFont.Style = [fsBold]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Segoe UI'
            Font.Style = []
            ID = ''
            Width = 68.000000000000000000
          end
          item
            BorderWidth = 1
            FixedFont.Charset = DEFAULT_CHARSET
            FixedFont.Color = clWindowText
            FixedFont.Height = -11
            FixedFont.Name = 'Segoe UI'
            FixedFont.Style = [fsBold]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Segoe UI'
            Font.Style = []
            ID = ''
            Width = 68.000000000000000000
          end
          item
            BorderWidth = 1
            FixedFont.Charset = DEFAULT_CHARSET
            FixedFont.Color = clWindowText
            FixedFont.Height = -11
            FixedFont.Name = 'Segoe UI'
            FixedFont.Style = [fsBold]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Segoe UI'
            Font.Style = []
            ID = ''
            Width = 68.000000000000000000
          end
          item
            BorderWidth = 1
            FixedFont.Charset = DEFAULT_CHARSET
            FixedFont.Color = clWindowText
            FixedFont.Height = -11
            FixedFont.Name = 'Segoe UI'
            FixedFont.Style = [fsBold]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Segoe UI'
            Font.Style = []
            ID = ''
            Width = 68.000000000000000000
          end
          item
            BorderWidth = 1
            FixedFont.Charset = DEFAULT_CHARSET
            FixedFont.Color = clWindowText
            FixedFont.Height = -11
            FixedFont.Name = 'Segoe UI'
            FixedFont.Style = [fsBold]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Segoe UI'
            Font.Style = []
            ID = ''
            Width = 68.000000000000000000
          end
          item
            BorderWidth = 1
            FixedFont.Charset = DEFAULT_CHARSET
            FixedFont.Color = clWindowText
            FixedFont.Height = -11
            FixedFont.Name = 'Segoe UI'
            FixedFont.Style = [fsBold]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Segoe UI'
            Font.Style = []
            ID = ''
            Width = 68.000000000000000000
          end
          item
            BorderWidth = 1
            FixedFont.Charset = DEFAULT_CHARSET
            FixedFont.Color = clWindowText
            FixedFont.Height = -11
            FixedFont.Name = 'Segoe UI'
            FixedFont.Style = [fsBold]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Segoe UI'
            Font.Style = []
            ID = ''
            Width = 68.000000000000000000
          end
          item
            BorderWidth = 1
            FixedFont.Charset = DEFAULT_CHARSET
            FixedFont.Color = clWindowText
            FixedFont.Height = -11
            FixedFont.Name = 'Segoe UI'
            FixedFont.Style = [fsBold]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Segoe UI'
            Font.Style = []
            ID = ''
            Width = 68.000000000000000000
          end
          item
            BorderWidth = 1
            FixedFont.Charset = DEFAULT_CHARSET
            FixedFont.Color = clWindowText
            FixedFont.Height = -11
            FixedFont.Name = 'Segoe UI'
            FixedFont.Style = [fsBold]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Segoe UI'
            Font.Style = []
            ID = ''
            Width = 68.000000000000000000
          end>
        DefaultFont.Charset = DEFAULT_CHARSET
        DefaultFont.Color = clWindowText
        DefaultFont.Height = -11
        DefaultFont.Name = 'Segoe UI'
        DefaultFont.Style = []
        TopRow = 1
        Appearance.FixedLayout.Fill.Color = 16380654
        Appearance.FixedLayout.Font.Charset = DEFAULT_CHARSET
        Appearance.FixedLayout.Font.Color = 4539717
        Appearance.FixedLayout.Font.Height = -13
        Appearance.FixedLayout.Font.Name = 'Segoe UI'
        Appearance.FixedLayout.Font.Style = [fsBold]
        Appearance.NormalLayout.Fill.Color = 16578806
        Appearance.NormalLayout.Font.Charset = DEFAULT_CHARSET
        Appearance.NormalLayout.Font.Color = 8026746
        Appearance.NormalLayout.Font.Height = -11
        Appearance.NormalLayout.Font.Name = 'Segoe UI'
        Appearance.NormalLayout.Font.Style = []
        Appearance.GroupLayout.Fill.Color = 12817262
        Appearance.GroupLayout.Font.Charset = DEFAULT_CHARSET
        Appearance.GroupLayout.Font.Color = clBlack
        Appearance.GroupLayout.Font.Height = -11
        Appearance.GroupLayout.Font.Name = 'Segoe UI'
        Appearance.GroupLayout.Font.Style = []
        Appearance.SummaryLayout.Fill.Color = 14009785
        Appearance.SummaryLayout.Font.Charset = DEFAULT_CHARSET
        Appearance.SummaryLayout.Font.Color = clBlack
        Appearance.SummaryLayout.Font.Height = -11
        Appearance.SummaryLayout.Font.Name = 'Segoe UI'
        Appearance.SummaryLayout.Font.Style = []
        Appearance.SelectedLayout.Fill.Color = 16441019
        Appearance.SelectedLayout.Font.Charset = DEFAULT_CHARSET
        Appearance.SelectedLayout.Font.Color = 4539717
        Appearance.SelectedLayout.Font.Height = -11
        Appearance.SelectedLayout.Font.Name = 'Segoe UI'
        Appearance.SelectedLayout.Font.Style = []
        Appearance.FocusedLayout.Fill.Color = 16039284
        Appearance.FocusedLayout.Font.Charset = DEFAULT_CHARSET
        Appearance.FocusedLayout.Font.Color = 4539717
        Appearance.FocusedLayout.Font.Height = -11
        Appearance.FocusedLayout.Font.Name = 'Segoe UI'
        Appearance.FocusedLayout.Font.Style = []
        Appearance.FixedSelectedLayout.Fill.Color = 14599344
        Appearance.FixedSelectedLayout.Font.Charset = DEFAULT_CHARSET
        Appearance.FixedSelectedLayout.Font.Color = clBlack
        Appearance.FixedSelectedLayout.Font.Height = -11
        Appearance.FixedSelectedLayout.Font.Name = 'Segoe UI'
        Appearance.FixedSelectedLayout.Font.Style = []
        Appearance.BandLayout.Fill.Color = 16711679
        Appearance.BandLayout.Font.Charset = DEFAULT_CHARSET
        Appearance.BandLayout.Font.Color = 8026746
        Appearance.BandLayout.Font.Height = -11
        Appearance.BandLayout.Font.Name = 'Segoe UI'
        Appearance.BandLayout.Font.Style = []
        Appearance.ProgressLayout.Format = '%.0f%%'
        LeftCol = 0
        ScrollMode = scmItemScrolling
        GlobalFont.Scale = 1.000000000000000000
        GlobalFont.Style = []
        DesignTimeSampleData = True
      end
    end
  end
  object sbMain: TTMSFNCStatusBar
    Left = 0
    Top = 496
    Width = 712
    Height = 20
    ParentDoubleBuffered = False
    DoubleBuffered = True
    TabOrder = 1
    Panels = <
      item
        Width = 50
        Progress.Level0Fill.Color = clLime
        Progress.Level1Fill.Color = clYellow
        Progress.Level2Fill.Color = 42495
        Progress.Level3Fill.Color = clRed
        Progress.Position = 0
      end
      item
        AutoSize = True
        Text = '123'
        Style = spsProgress
        Alignment = taCenter
        Width = 30
        Progress.Level0Fill.Color = clLime
        Progress.Level1Fill.Color = clYellow
        Progress.Level2Fill.Color = 42495
        Progress.Level3Fill.Color = clRed
        Progress.Position = 0
        Visible = False
      end>
    PanelAppearance.Font.Charset = DEFAULT_CHARSET
    PanelAppearance.Font.Color = clWindowText
    PanelAppearance.Font.Height = -11
    PanelAppearance.Font.Name = 'Segoe UI'
    PanelAppearance.Font.Style = []
  end
end
