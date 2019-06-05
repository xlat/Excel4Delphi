unit zexmlss;

interface

uses
  Classes, Sysutils, Graphics, UITypes, Math, Windows, RegularExpressions, Generics.Collections, System.Contnrs,
  zsspxml;

var ZE_XLSX_APPLICATION: string;

const _PointToMM: real = 0.3528;  // 1 типографский пункт = 0.3528 мм

type
  /// <summary>
  /// Data type of cell
  /// </summary>
  TZCellType = (ZENumber, ZEDateTime, ZEBoolean, ZEString, ZEError, ZEGeneral);
type
  /// <summary>
  /// Style lines of the cell border
  /// </summary>
  TZBorderType = (ZENone, ZEContinuous, ZEHair, ZEDot, ZEDash, ZEDashDot, ZEDashDotDot, ZESlantDashDot, ZEDouble);

  /// <summary>
  /// Horizontal alignment
  /// </summary>
  TZHorizontalAlignment = (ZHAutomatic, ZHLeft, ZHCenter, ZHRight, ZHFill, ZHJustify, ZHCenterAcrossSelection, ZHDistributed, ZHJustifyDistributed);

  /// <summary>
  /// Vertical alignment
  /// </summary>
  TZVerticalAlignment = (ZVAutomatic, ZVTop, ZVBottom, ZVCenter, ZVJustify, ZVDistributed, ZVJustifyDistributed);

  /// <summary>
  /// Fill pattern of the cell
  /// </summary>
  TZCellPattern = (ZPNone, ZPSolid, ZPGray75, ZPGray50, ZPGray25, ZPGray125, ZPGray0625, ZPHorzStripe, ZPVertStripe,
                  ZPReverseDiagStripe, ZPDiagStripe, ZPDiagCross, ZPThickDiagCross, ZPThinHorzStripe, ZPThinVertStripe,
                  ZPThinReverseDiagStripe, ZPThinDiagStripe, ZPThinHorzCross, ZPThinDiagCross);

  /// <summary>
  /// Vertical/horizontal split/freeze mode.
  /// </summary>
  TZSplitMode = (ZSplitNone, ZSplitFrozen, ZSplitSplit);

  TZViewMode = (zvmNormal, zvmPageBreakPreview);

  /// <summary>
  /// Inherited from TPersistent. It's a cell of spreadsheet.
  /// </summary>
  TZCell = class(TPersistent)
  private
    FFormula: string;
    FData: string;
    FHref: string;
    FHRefScreenTip: string;
    FComment: string;
    FCommentAuthor: string;
    FAlwaysShowComment: boolean; //default = false;
    FShowComment: boolean;       //default = false
    FCellType: TZCellType;
    FCellStyle: integer;
    function GetDataAsDouble: double;
    procedure SetDataAsDouble(const Value: double);
    procedure SetDataAsInteger(const Value: integer);
    function GetDataAsInteger: integer;

    function GetDataAsDateTime(): TDateTime;
    procedure SetDataAsDateTime(const Value: TDateTime);
    procedure SetDataAsString(const Value: string);
  public
    constructor Create(); virtual;
    procedure Assign(Source: TPersistent); override;
    /// <summary>
    /// Clear cell.
    /// </summary>
    procedure Clear();
    /// <summary>
    /// Always show comment. <br />False by default.
    /// </summary>
    property AlwaysShowComment: boolean read FAlwaysShowComment write FAlwaysShowComment default false;
    /// <summary>
    /// Comment text.
    /// </summary>
    property Comment: string read FComment write FComment;
    /// <summary>
    /// Author of comment.
    /// </summary>
    property CommentAuthor: string read FCommentAuthor write FCommentAuthor;
    /// <summary>
    /// Cell style number. <br />-1 by default.
    /// </summary>
    property CellStyle: integer read FCellStyle write FCellStyle default -1;
    /// <summary>
    /// Cell type. <br />ZEString by default.
    /// </summary>
    property CellType: TZCellType read FCellType write FCellType default ZEString;
    /// <summary>
    /// Specifies the value of this cell to show on screen.
    /// </summary>
    property Data: string read FData write FData;
    /// <summary>
    /// Formula in R1C1 style.
    /// </summary>
    property Formula: string read FFormula write FFormula;
    /// <summary>
    /// Specifies the URL to link this cell.
    /// </summary>
    property HRef: string read FHref write FHref;
    /// <summary>
    /// Displays the caption of URL to show on screen.
    /// </summary>
    property HRefScreenTip: string read FHRefScreenTip write FHRefScreenTip;
    /// <summary>
    /// Show comment. <br />False by default.
    /// </summary>
    property ShowComment: boolean read FShowComment write FShowComment default false;
    /// <summary>
    /// Present cell data as double value
    /// </summary>
    property AsDouble: double read GetDataAsDouble write SetDataAsDouble;
    /// <summary>
    /// Present cell data as integer value
    /// </summary>
    property AsInteger: integer read GetDataAsInteger write SetDataAsInteger;
    /// <summary>
    /// Present cell data as TDateTime value
    /// </summary>
    property AsDateTime: TDateTime read GetDataAsDateTime write SetDataAsDateTime;
    /// <summary>
    /// Present cell data as string value
    /// </summary>
    property AsString: string read FData write SetDataAsString;
  end;

  /// <summary>
  /// Border's style.
  /// </summary>
  TZBorderStyle = class (TPersistent)
  private
    FLineStyle: TZBorderType;
    FWeight: byte;
    FColor: TColor;
    procedure SetLineStyle(const Value: TZBorderType);
    procedure SetWeight(const Value: byte);
    procedure SetColor(const Value: TColor);
  public
    constructor Create();virtual;
    procedure Assign(Source: TPersistent); override;
    /// <returns>
    /// True, when border equal to Source.
    /// </returns>
    function IsEqual(Source: TPersistent): boolean; virtual;
  published
    /// <summary>
    /// Line style. <br />ZENone by default.
    /// </summary>
    property LineStyle: TZBorderType read FLineStyle write SetLineStyle default ZENone;
    /// <summary>
    /// Specifies the thickness of this border (0-3). <br />0 by default.
    /// </summary>
    property Weight: byte read FWeight write SetWeight default 0;
    /// <summary>
    /// Specifies the color of this border. <br />ClBlack by default.
    /// </summary>
    property Color: TColor read FColor write SetColor default ClBlack;
  end;

  /// <summary>
  /// Borders position.
  /// </summary>
  TZBordersPos = ( bpLeft, bpTop, bpRight, bpBottom, bpDiagonalLeft, bpDiagonalRight );

  /// <summary>
  /// Borders of cell.
  /// </summary>
  TZBorder = class (TPersistent)
  private
    FBorder: array [0..5] of TZBorderStyle;
    procedure SetBorder(Num: TZBordersPos; Const Value: TZBorderStyle);
    function GetBorder(Num: TZBordersPos):TZBorderStyle;
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent);override;
    /// <summary>
    /// Set border by border position.
    /// </summary>
    property Border[Num: TZBordersPos]: TZBorderStyle read GetBorder write SetBorder; default;
    /// <returns>
    /// True when borders equal Source.
    /// </returns>
    function IsEqual(Source: TPersistent): boolean; virtual;
  published
    /// <summary>
    /// Left border.
    /// </summary>
    property Left         : TZBorderStyle index bpLeft          read GetBorder write SetBorder;
    /// <summary>
    /// Top border.
    /// </summary>
    property Top          : TZBorderStyle index bpTop           read GetBorder write SetBorder;
    /// <summary>
    /// Right border.
    /// </summary>
    property Right        : TZBorderStyle index bpRight         read GetBorder write SetBorder;
    /// <summary>
    /// Bottom border.
    /// </summary>
    property Bottom       : TZBorderStyle index bpBottom        read GetBorder write SetBorder;
    /// <summary>
    /// Diagonal from upper left to lower right.
    /// </summary>
    property DiagonalLeft : TZBorderStyle index bpDiagonalLeft  read GetBorder write SetBorder;
    /// <summary>
    /// Diagonal from lower left to upper right.
    /// </summary>
    property DiagonalRight: TZBorderStyle index bpDiagonalRight read GetBorder write SetBorder;
  end;

  /// Угол поворота текста в ячейке. Целое со знаком.
  ///    Базовый диапазон -90 .. +90, расширенный -180 .. +180 (градусов)
  ///    Значения из расширенного диапазона НЕ МОГУТ работать в Microsoft Excel!
  /// Определение по XML SS совпадает с базовым
  /// Определение по OpenDocument 1.2 позволяет любое дробное число
  ///    в разных еденицах измерения (градусы, грады, радианы).
  ///    практически используется целое от 0 до 359
  /// Определение по OfficeXML наиболее запутанное:
  ///    254 и 255 - специальные значения,
  ///        255 работает в Excel хотя в стандарте не описано
  ///        254 принимается, но после открытия файла не работает
  ///    от 0 до +90 соответствует обоим вышеописанным стандартам
  ///    от +91 до +180 соответствует значениям -1 .. -90 [ = 90 - value]
  ///
  /// 1. MSDN aa140066.aspx: XML Spreadsheet Reference => ss:Rotate
  /// 2. OASIS OpenDocument 1.2 (Sep.2011) Part 1 => 18.3 Other Datatypes => 18.3.1 angle
  /// 3. ECMA-376 4th Edition (Dec.2012) Part 1 => 18.8 Styles => 18.8.1 alignment
  ///
  ///  An angle of the rotation (direction) for a text within a cell. Signed Integer;
  ///    Nominative range is -90 .. +90, extended is -180 .. +180 (in degrees)
  ///       Extended range values can not be loaded by Microsoft Excel!
  /// XML SS specifications matches the nominative range.
  /// OpenDocument 1.2 specification, atchign mathematics and SVG, permits FP
  ///    numbers in different unit (deg, grad, rad) with degrees by default;
  ///    practical implementations use non-negative integer numbers in 0..359 range.
  /// OfficeXML specification is the most perplexed
  ///    254 and 255 are special values, and:
  ///        255 works within MS Excel despite being not defined in the specs;
  ///        254 can be opened by excel, but such a cell can not be rendered.
  ///    The 0 to +90 range matches both standards aforementioned
  ///    The +91 to +180 range matches -1 to -90 range. [ = 90 - value]
  ///
  TZCellTextRotate = -180 .. +359;

  /// <summary>
  /// Specifies how text is aligned within the cell.
  /// </summary>
  TZAlignment = class (TPersistent)
  private
    FHorizontal: TZHorizontalAlignment;
    FVertical: TZVerticalAlignment;
    FIndent: integer;
    FRotate: TZCellTextRotate;
    FWrapText: boolean;
    FShrinkToFit: boolean;              // true - уменьшает размер шрифта, чтобы текст поместился в ячейку,
                                        // false - текст не уменьшается
    FVerticalText: boolean;             // true  - текст по одной букве в строке вертикально
                                        // false - по дефолту
    procedure SetHorizontal(const Value: TZHorizontalAlignment);
    procedure SetIndent(const Value: integer);
    procedure SetRotate(const Value: TZCellTextRotate);
    procedure SetShrinkToFit(const Value: boolean);
    procedure SetVertical(const Value: TZVerticalAlignment);
    procedure SetVerticalText(const Value:boolean);
    procedure SetWrapText(const Value: boolean);
  public
    constructor Create(); virtual;
    procedure Assign(Source: TPersistent); override;
    /// <returns>
    /// True when all properties equal to Source's properties.
    /// </returns>
    function IsEqual(Source: TPersistent): boolean; virtual;
  published
    /// <summary>
    /// Specifies how text is aligned by horizontally within the cell. <br />ZHAutomatic by default.
    /// </summary>
    property Horizontal: TZHorizontalAlignment read FHorizontal write SetHorizontal default ZHAutomatic;
    /// <summary>
    /// Specifies how far the cell's text is indented. <br />0 by default.
    /// </summary>
    property Indent: integer read FIndent write SetIndent default 0;
    /// <summary>
    /// Specifies the rotation of the text within the cell (from -90 to 90). <br />0 by default.
    /// </summary>
    property Rotate: TZCellTextRotate read FRotate write SetRotate;
    /// <summary>
    /// If True then the text size will shrunk so to all of the text fits within the cell. <br />False by default.
    /// </summary>
    property ShrinkToFit: boolean read FShrinkToFit write SetShrinkToFit default false;
    /// <summary>
    /// Specifies how text is aligned by vertically within the cell. <br />ZVAutomatic by default.
    /// </summary>
    property Vertical: TZVerticalAlignment read FVertical write SetVertical default ZVAutomatic;
    /// <summary>
    /// If True each letter is drawn horizontally, one above the other. <br />False by default.
    /// </summary>
    property VerticalText: boolean read FVerticalText write SetVerticalText default false;
    /// <summary>
    /// Specifies whether the text in cell should wrap at the cell boundary. <br />False by default.
    /// </summary>
    property WrapText: boolean read FWrapText write SetWrapText default false;
  end;

  /// <summary>
  /// Cell style.
  /// </summary>
  TZStyle = class (TPersistent)
  private
    FBorder: TZBorder;
    FAlignment: TZAlignment;
    FFont: TFont;
    FBGColor: TColor;
    FPatternColor: TColor;
    FCellPattern: TZCellPattern;
    FNumberFormatId: Integer;
    FNumberFormat: string;
    FProtect: boolean;
    FHideFormula: boolean;
    procedure SetFont(const Value: TFont);
    procedure SetBorder(const Value: TZBorder);
    procedure SetAlignment(const Value: TZAlignment);
    procedure SetBGColor(const Value: TColor);
    procedure SetPatternColor(const Value: TColor);
    procedure SetCellPattern(const Value: TZCellPattern);
  protected
    procedure SetNumberFormat(const Value: string); virtual;
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    /// <summary>
    /// True when style equal Source.
    /// </summary>
    function IsEqual(Source: TPersistent): boolean; virtual;
  published
    /// <summary>
    /// Cell font.
    /// </summary>
    property Font: TFont read FFont write SetFont;
    /// <summary>
    /// Cell borders.
    /// </summary>
    property Border: TZBorder read FBorder write SetBorder;
    /// <summary>
    /// Specifies how text is aligned within the cell.
    /// </summary>
    property Alignment: TZAlignment read FAlignment write SetAlignment;
    /// <summary>
    /// Background color of the cell. <br />clWindow by default.
    /// </summary>
    property BGColor: TColor read FBGColor write SetBGColor default clWindow;
    /// <summary>
    /// Color of fill pattern. <br />clWindow by default.
    /// </summary>
    property PatternColor: TColor read FPatternColor write SetPatternColor default clWindow;
    /// <summary>
    /// Indicates whether or not this cell is protected. <br />True by default.
    /// </summary>
    property Protect: boolean read FProtect write FProtect default true;
    /// <summary>
    /// Indicates whether or not this cell's formula should be hidden when sheet protection is enabled. <br />False by default.
    /// </summary>
    property HideFormula: boolean read FHideFormula write FHideFormula default false;
    /// <summary>
    /// Fill pattern of the cell. <br />ZPNone by default.
    /// </summary>
    property CellPattern: TZCellPattern read FCellPattern write SetCellPattern default ZPNone;
    /// <summary>
    /// Defines the number format that should be in cells referencing this style.
    /// </summary>
    property NumberFormat: string read FNumberFormat write SetNumberFormat;
    property NumberFormatId: Integer read FNumberFormatId write FNumberFormatId default -1;
  end;

  /// <summary>
  /// Contains styles of the document.
  /// </summary>
  TZStyles = class (TPersistent)
  private
    FDefaultStyle: TZStyle;
    FStyles: array of TZStyle;
    FCount: integer;
    procedure SetDefaultStyle(const Value: TZStyle);
    function  GetStyle(num: integer): TZStyle;
    procedure SetStyle(num: integer; const Value: TZStyle);
    procedure SetCount(const Value: integer);
  protected
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    /// <summary>
    /// Add a Style.
    /// </summary>
    /// <param name="Style">
    /// Style.
    /// </param>
    /// <param name="CheckMatch">
    /// Checks the coincidence of this style with introduced in earlier styles.
    /// <br />Return number of added (or, if CheckMatch = True, previously introduced) style.
    /// </param>
    function Add(const Style: TZStyle; CheckMatch: boolean = false): integer;
    procedure Assign(Source: TPersistent); override;
    /// <summary>
    /// Delete all styles.
    /// </summary>
    procedure Clear(); virtual;
    /// <summary>
    /// Delete style num, styles with a larger number are shifting.
    /// </summary>
    /// <param name="num">
    /// Style number
    /// </param>
    /// <returns>
    /// Return: 0 - if successfully deleted, -1 - can not delete style
    /// </returns>
    function DeleteStyle(num: integer):integer; virtual;
    /// <summary>
    /// Find number to match the Style introduced in earlier styles.
    /// </summary>
    /// <param name="Style">
    /// Style.
    /// </param>
    /// <returns>
    /// Return: -2 - if style not found, -1 - if style = DefaultStyle, 0..Count-1 - Style
    /// </returns>
    function Find(const Style: TZStyle): integer;
    /// <summary>
    /// Style num (-1 - DefaultStyle).
    /// </summary>
    property Items[num: integer]: TZStyle read GetStyle write SetStyle; default;
    /// <summary>
    /// Count styles in the document.
    /// </summary>
    property Count: integer read FCount write SetCount;
  published
    /// <summary>
    /// Default style ( = Items[-1]).
    /// </summary>
    property DefaultStyle: TZStyle read FDefaultStyle write SetDefaultStyle;
  end;

  TZSheet = class;

  /// <summary>
  /// Merged cells
  /// </summary>
  TZMergeCells = class
  private
    FSheet: TZSheet;
    FCount: integer;
    FMergeArea: Array of TRect;
    function GetItem(Num: integer): TRect;
    procedure SetItem(Num: integer; const rect: TRect);
  public
    constructor Create(ASheet: TZSheet); virtual;
    destructor Destroy(); override;
    /// <summary>
    /// Adds a merged cell enclosed into rectangle.
    /// </summary>
    function AddRect(Rct:TRect): byte;
    /// <summary>
    /// Adds a merged cell enclosed into rectangle
    /// </summary>
    function AddRectXY(x1, y1, x2, y2: integer): byte;
    /// <summary>
    /// Delete merged cell num. <br />Return True if the cell is successfully deleted.
    /// </summary>
    function DeleteItem(num: integer): boolean;
    /// <summary>
    /// Returns the number of merged cell, in which the cell [ACol, ARow] is top left.
    /// If returns a negative value - there is no such area.
    /// </summary>
    function InLeftTopCorner(ACol, ARow: integer): integer;
    /// <summary>
    /// Returns the number of merged cell that includes cell [ACol, ARow].
    /// If returns a negative value - cell [ACol, ARow] is not contained in the Merged area.
    /// </summary>
    function InMergeRange(ACol, ARow: integer): integer;
    function IsCrossWithArea(AID,AC1,AR1,AC2,AR2: integer): Boolean;
    /// <summary>
    /// Removes all merged cells.
    /// </summary>
    procedure Clear();
    property Count: integer read FCount;
    property Items[Num: Integer]: TRect read GetItem write SetItem; default;
  end;

  TZCellColumn = array of TZCell;

  TZEXMLSS = class;

  /// <summary>
  /// Common options for columns and rows. Ancestor for TZColOptions and TZRowOptions.
  /// </summary>
  TZRowColOptions = class (TPersistent)
  private
    FSheet: TZSheet;
    FHidden: boolean;
    FStyleID: integer;
    FSize: real;
    FAuto: boolean;
    FBreaked: boolean;
  protected
    function  GetAuto(): boolean;
    procedure SetAuto(Value: boolean);
    function  GetSizePoint(): real;
    procedure SetSizePoint(Value: real);
    function  GetSizeMM(): real;
    procedure SetSizeMM(Value: real);
    function  GetSizePix(): integer; virtual; abstract;
    procedure SetSizePix(Value: integer); virtual; abstract;
  public
    constructor Create(ASheet: TZSheet); virtual;
    procedure Assign(Source: TPersistent); override;
    /// <summary>
    /// True specifies that column or row is hidden. <br />False (not hidden) by default.
    /// </summary>
    property Hidden: boolean read FHidden write FHidden default false;
    /// <summary>
    /// Specifies a style for column or row. <br />-1 by default.
    /// </summary>
    property StyleID: integer read FStyleID write FStyleID default -1;
    /// <summary>
    /// Page break after column or row. <br />False (no break) by default.
    /// </summary>
    property Breaked: boolean read FBreaked write FBreaked default false;
  end;

  /// <summary>
  /// Column options.
  /// </summary>
  TZColOptions = class (TZRowColOptions)
  protected
    function  GetSizePix(): integer; override;
    procedure SetSizePix(Value: integer); override;
  public
    constructor Create(ASheet: TZSheet); override;
    /// <summary>
    /// If True, it means that this column should be autosized.
    /// </summary>
    property AutoFitWidth: boolean read GetAuto write SetAuto;
    /// <summary>
    /// Column width in points.
    /// </summary>
    property Width: real read GetSizePoint write SetSizePoint;
    /// <summary>
    /// Column width in mm.
    /// </summary>
    property WidthMM: real read GetSizeMM write SetSizeMM;
    /// <summary>
    /// Column width in pixels.
    /// </summary>
    property WidthPix: integer read GetSizePix write SetSizePix;
  end;

  /// <summary>
  /// Row options.
  /// </summary>
  TZRowOptions = class (TZRowColOptions)
  protected
    function  GetSizePix(): integer; override;
    procedure SetSizePix(Value: integer); override;
  public
    constructor Create(ASheet: TZSheet); override;
    /// <summary>
    /// If True, it means that this row should be autosized.
    /// </summary>
    property AutoFitHeight: boolean read GetAuto write SetAuto;
    /// <summary>
    /// Row height in points (1 point = 1/72" = 0.3528 mm).
    /// </summary>
    property Height: real read GetSizePoint write SetSizePoint;
    /// <summary>
    /// Row height in mm.
    /// </summary>
    property HeightMM: real read GetSizeMM write SetSizeMM;
    /// <summary>
    /// Row height in pixels.
    /// </summary>
    property HeightPix: integer read GetSizePix write SetSizePix;
  end;

  /// repeating columns or rows when printing the sheet
  ///   implemented as binary mess in XLSX
  ///   implemented as named range and overriding cell names in XML SS
  ///   implemented as special container tags in OpenDocument
  TZSheetPrintTitles = class(TPersistent)
  private
    FOwner: TZSheet;
    FColumns: boolean;
    FActive: boolean;
    FTill: word;
    FFrom: word;
    procedure SetActive(const Value: boolean);
    procedure SetFrom(const Value: word);
    procedure SetTill(const Value: word);
    function  Valid(const AFrom, ATill: word): boolean;
    procedure RequireValid(const AFrom, ATill: word);
  public
    procedure Assign(Source: TPersistent); override;
    constructor Create(const owner: TZSheet; const ForColumns: boolean);
    function ToString: string; override;
  published

    property From: word read FFrom write SetFrom;
    property Till: word read FTill write SetTill;
    property Active: boolean read FActive write SetActive;
  end;

  /// <summary>
  /// Footer or header margins.
  /// </summary>
  TZHeaderFooterMargins = class (TPersistent)
  private
    FMarginTopBottom: word;       //Bottom or top margin
    FMarginLeft: word;
    FMarginRight: word;
    FHeight: word;
    FUseAutoFitHeight: boolean;   //If true then in LibreOffice Calc
                                  //in window "Page style settings" on tabs "Header" or "Footer"
                                  //check box "AutoFit height" will be checked.
  public
    constructor Create();
    procedure Assign(Source: TPersistent); override;
    function IsEqual(Source: TPersistent): boolean; virtual;
  published
    /// <summary>
    /// Bottom / top margin of the footer / header in mm. <br />13 by default.
    /// </summary>
    property MarginTopBottom: word read FMarginTopBottom write FMarginTopBottom default 13;
    /// <summary>
    /// Left margin in mm. <br />0 by default.
    /// </summary>
    property MarginLeft: word read FMarginLeft write FMarginLeft default 0;
    /// <summary>
    /// Right margin in mm. <br />0 by default.
    /// </summary>
    property MarginRight: word read FMarginRight write FMarginRight default 0;
    /// <summary>
    /// Height of footer / header in mm. <br />7 by default.
    /// </summary>
    property Height: word read FHeight write FHeight default 7;
    /// <summary>
    /// Automatically fit height. <br />true by default.
    /// </summary>
    property UseAutoFitHeight: boolean read FUseAutoFitHeight write FUseAutoFitHeight default true;
  end;

  /// <summary>
  /// Inherited from TPersistent. Sheet options.
  /// </summary>
  TZSheetOptions = class (TPersistent)
  private
    FActiveCol: word;
    FActiveRow: word;
    FMarginBottom: word;
    FMarginLeft: word;
    FMarginTop: word;
    FMarginRight: word;
    FHeaderMargins: TZHeaderFooterMargins;
    FFooterMargins: TZHeaderFooterMargins;
    FPortraitOrientation: boolean;
    FCenterHorizontal: boolean;
    FCenterVertical: boolean;
    FStartPageNumber: integer;
    FDifferentOddEven: Boolean;
    FDifferentFirst: Boolean;

    FHeader: string;     //Header for all pages. When IsEvenFooterEqual = false - only for odd pages.
    FFooter: string;     //Footer for all pages. When IsEvenFooterEqual = false - only for odd pages.
    FEvenHeader: string; //Header for even pages. Used only if IsEvenFooterEqual = false
    FEvenFooter: string; //Footer for even pages. Used only if IsEvenFooterEqual = false
    FFirstPageHeader: string;
    FFirstPageFooter: string;

    FHeaderBGColor: TColor;
    FFooterBGColor: TColor;

    FScaleToPercent: integer;         //Document must be scaled to percentage value (100 - no scale)
    FScaleToPages: integer;           //Document scaled to fit a number of pages (1 - no scale)

    FPaperSize: byte;
    FPaperWidth: integer;
    FPaperHeight: integer;

    FFitToHeight: integer; //Number of vertical pages to fit on
    FFitToWidth: integer;  //Number of horizontal pages to fit on

    FSplitVerticalMode: TZSplitMode;
    FSplitHorizontalMode: TZSplitMode;
    FSplitVerticalValue: integer;       //Вроде можно вводить отрицательные
    FSplitHorizontalValue: integer;     //Измеряться будут:
                                        //    в пикселях, если SplitMode = ZSplitSplit
                                        //    в кол-ве строк/столбцов, если SplitMode = ZSplitFrozen
                                        // Если SplitMode = ZSplitNone, то фиксация столбцов/ячеек не работает
    function GetHeaderMargin(): word;
    procedure SetHeaderMargin(Value: word);
    function GetFooterMargin(): word;
    procedure SetFooterMargin(Value: word);
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
  published
    /// <summary>
    /// Column number with active cell. <br />0 by default.
    /// </summary>
    property ActiveCol: word read FActiveCol write FActiveCol default 0;
    /// <summary>
    /// Row number with active cell. <br />0 by default.
    /// </summary>
    property ActiveRow: word read FActiveRow write FActiveRow default 0;
    /// <summary>
    /// Specifies the bottom margin on the page in millimeters. <br />25 mm by default.
    /// </summary>
    property MarginBottom: word read FMarginBottom write FMarginBottom default 25;
    /// <summary>
    /// Specifies the left margin on the page in millimeters. <br />20 mm by default.
    /// </summary>
    property MarginLeft: word read FMarginLeft write FMarginLeft default 20;
    /// <summary>
    /// Specifies the top margin on the page in millimeters. <br />25 mm by default.
    /// </summary>
    property MarginTop: word read FMarginTop write FMarginTop default 25;
    /// <summary>
    /// Specifies the right margin on the page in millimeters. <br />20 mm by default.
    /// </summary>
    property MarginRight: word read FMarginRight write FMarginRight default 20;
    /// <summary>
    /// Paper Size (Paper size table). <br />9 (A4) by default.
    /// </summary>
    property PaperSize: byte read FPaperSize write FPaperSize default 9;
    /// <summary>
    /// Paper width in mm. Used only when PaperSize = 0!
    /// </summary>
    property PaperWidth: integer read FPaperWidth write FPaperWidth default 0;
    /// <summary>
    /// Paper height in mm. Used only when PaperSize = 0!
    /// </summary>
    property PaperHeight: integer read FPaperHeight write FPaperHeight default 0;
    property FitToHeight: integer read FFitToHeight write FFitToHeight default -1;
    property FitToWidth: integer read FFitToWidth write FFitToWidth default -1;
    /// <summary>
    /// Specifies the orientation of the page (True - Portrait, False - Landscape). <br />True by default.
    /// </summary>
    property PortraitOrientation: boolean read FPortraitOrientation write FPortraitOrientation default true;
    /// <summary>
    /// If True, the document should be centered horizontally on the page. <br />False by default.
    /// </summary>
    property CenterHorizontal: boolean read FCenterHorizontal write FCenterHorizontal default false;
    /// <summary>
    /// If True, the document should be centered vertically on the page. <br />False by default.
    /// </summary>
    property CenterVertical: boolean read FCenterVertical write FCenterVertical default false;
    /// <summary>
    /// Specifies the starting page number for print. <br />1 by default.
    /// </summary>
    property StartPageNumber: integer read FStartPageNumber write FStartPageNumber default 1;
    /// <summary>
    /// The size of header in millimeters. <br />13 mm by default. <br />
    /// </summary>
    /// <remarks>
    /// deprecated 'Use HeaderMargins.Height!'
    /// </remarks>
    property HeaderMargin: word read GetHeaderMargin write SetHeaderMargin default 13; //deprecated!
    /// <summary>
    /// The size of footer in millimeters. <br />13 mm by default. <br />
    /// </summary>
    /// <remarks>
    /// deprecated 'Use FooterMargins.Height!'
    /// </remarks>
    property FooterMargin: word read GetFooterMargin write SetFooterMargin default 13; //deprecated!
    /// <summary>
    /// Sizes and margins for header in mm.
    /// </summary>
    property HeaderMargins: TZHeaderFooterMargins read FHeaderMargins;
    /// <summary>
    ///Sizes and margins for footer in mm.
    /// </summary>
    property FooterMargins: TZHeaderFooterMargins read FFooterMargins;
    /// <summary>
    ///
    /// </summary>
    property IsDifferentFirst: Boolean read FDifferentFirst write FDifferentFirst;
    /// <summary>
    ///
    /// </summary>
    property IsDifferentOddEven: Boolean read FDifferentOddEven write FDifferentOddEven;
    property Header: string read FHeader write FHeader;
    property Footer: string read FFooter write FFooter;
    property EvenHeader: string read FEvenHeader write FEvenHeader;
    property EvenFooter: string read FEvenFooter write FEvenFooter;
    property FirstPageHeader: string read FFirstPageHeader write FFirstPageHeader;
    property FirstPageFooter: string read FFirstPageFooter write FFirstPageFooter;
    /// <summary>
    /// Background color for header. <br />clWindow by default.
    /// </summary>
    property HeaderBGColor: TColor read FHeaderBGColor write FHeaderBGColor default clWindow;
    /// <summary>
    /// Background color for footer. <br />clWindow by default.
    /// </summary>
    property FooterBGColor: TColor read FFooterBGColor write FFooterBGColor default clWindow;
    /// <summary>
    /// Document must be scaled to percentage value (100 - no scale). <br />100 by default.
    /// </summary>
    property ScaleToPercent: integer read FScaleToPercent write FScaleToPercent default 100;
    /// <summary>
    /// Document scaled to fit a number of pages (1 - no scale). <br />1 by default.
    /// </summary>
    property ScaleToPages: integer read FScaleToPages write FScaleToPages default 1;
    /// <summary>
    /// Vertical columns split/freeze mode.
    /// Does the same thing as LibreOffice Calc commands "Window - Freeze"/ "Window - Split" <br />ZSplitNone by default.
    /// </summary>
    property SplitVerticalMode: TZSplitMode read FSplitVerticalMode write FSplitVerticalMode default ZSplitNone;
    /// <summary>
    /// Horizontal rows split/freeze mode. <br />ZSplitNone by default.
    /// </summary>
    property SplitHorizontalMode: TZSplitMode read FSplitHorizontalMode write FSplitHorizontalMode default ZSplitNone;
    /// <summary>
    /// If SplitVerticalMode = ZSplitFrozen then count of column for freeze.
    /// If SplitVerticalMode = ZSplitSplit then size in pixels. <br />0 by default.
    /// </summary>
    property SplitVerticalValue: integer read FSplitVerticalValue write FSplitVerticalValue;
    /// <summary>
    /// If SplitHorizontalMode = ZSplitFrozen then count of rows for freeze.
    /// If SplitHorizontalMode = ZSplitSplit then size if pixels. <br />0 by default.
    /// </summary>
    property SplitHorizontalValue: integer read FSplitHorizontalValue write FSplitHorizontalValue;
  end;

  /// <summary>
  /// Conditions
  /// </summary>
  TZCondition = (
    ZCFIsTrueFormula,
    ZCFCellContentIsBetween,
    ZCFCellContentIsNotBetween,
    ZCFCellContentOperator,
    ZCFNumberValue,
    ZCFString,
    ZCFBoolTrue,
    ZCFBoolFalse,
    ZCFFormula,
    ZCFContainsText,
    ZCFNotContainsText,
    ZCFBeginsWithText,
    ZCFEndsWithText,
    ZCFCellIsEmpty,
    ZCFDuplicate,
    ZCFUnique,
    ZCFAboveAverage,
    ZCFBellowAverage,
    ZCFAboveEqualAverage,
    ZCFBelowEqualAverage,
    ZCFTopElements,
    ZCFBottomElements,
    ZCFTopPercent,
    ZCFBottomPercent,
    ZCFIsError,
    ZCFIsNoError
  );

  //Оператор для условного форматирования
  TZConditionalOperator = (
    ZCFOpGT,       //  > (Greater Than)
    ZCFOpLT,       //  < (Less Than)
    ZCFOpGTE,      //  >= (Greater or Equal)
    ZCFOpLTE,      //  <= (Less or Equal)
    ZCFOpEqual,    //  = (Equal)
    ZCFOpNotEqual  //  <> (Not Equal)
  );

  //Условное форматирование: стиль на условие (conditional formatting)
  TZConditionalStyleItem = class (TPersistent)
  private
    FCondition: TZCondition;                    //условие
    FConditionOperator: TZConditionalOperator;  //Оператор
    FValue1: String;
    FValue2: String;
    FApplyStyleID: integer;                     //номер применяемого стиля

                                                //Базовая ячейка (только для формул):
    FBaseCellPageIndex: integer;                //  Номер страницы для адреса базовой ячейки
                                                //    -1 - текущая страница
    FBaseCellRowIndex: integer;                 //  Номер строки
    FBaseCellColumnIndex: integer;              //  Номер столбца
  protected
  public
    constructor Create(); virtual;
    procedure Clear();
    procedure Assign(Source: TPersistent); override;
    function IsEqual(Source: TPersistent): boolean; virtual;
    property ApplyStyleID: integer read FApplyStyleID write FApplyStyleID;
    property BaseCellColumnIndex: integer read FBaseCellColumnIndex write FBaseCellColumnIndex;
    property BaseCellPageIndex: integer read FBaseCellPageIndex write FBaseCellPageIndex;
    property BaseCellRowIndex: integer read FBaseCellRowIndex write FBaseCellRowIndex;
    property Condition: TZCondition read FCondition write FCondition;
    property ConditionOperator: TZConditionalOperator read FConditionOperator write FConditionOperator;
    property Value1: String read FValue1 write FValue1;
    property Value2: String read FValue2 write FValue2;
  end;

  //Область для применения условного форматирования
  TZConditionalAreaItem = class (TPersistent)
  private
    FRow: integer;
    FColumn: integer;
    FWidth: integer;
    FHeight: integer;
    procedure SetRow(Value: integer);
    procedure SetColumn(Value: integer);
    procedure SetWidth(Value: integer);
    procedure SetHeight(Value: integer);
  protected
  public
    constructor Create(); overload; virtual;
    constructor Create(ColumnNum, RowNum, AreaWidth, AreaHeight: integer); overload; virtual;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(Source: TPersistent): boolean; virtual;
    property Row: integer read FRow write SetRow;
    property Column: integer read FColumn write SetColumn;
    property Width: integer read FWidth write SetWidth;
    property Height: integer read FHeight write SetHeight;
  end;

  //Области для применения условного форматирования
  TZConditionalAreas = class (TPersistent)
  private
    FCount: integer;
    FItems: array of TZConditionalAreaItem;
    procedure SetCount(Value: integer);
    function GetItem(num: integer): TZConditionalAreaItem;
    procedure SetItem(num: integer; Value: TZConditionalAreaItem);
  protected
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    function Add(): TZConditionalAreaItem; overload;
    function Add(ColumnNum, RowNum, AreaWidth, AreaHeight: integer): TZConditionalAreaItem; overload;
    procedure Assign(Source: TPersistent); override;
    procedure Delete(num: integer);
    function IsCellInArea(ColumnNum, RowNum: integer): boolean;
    function IsEqual(Source: TPersistent): boolean; virtual;
    property Count: integer read FCount write SetCount;
    property Items[num: integer]: TZConditionalAreaItem read GetItem write SetItem; default;
  end;

  //Условное форматирование: список условий
  TZConditionalStyle = class (TPersistent)
  private
    FCount: integer;                    //кол-во условий
    FMaxCount: integer;
    FAreas: TZConditionalAreas;
    FConditions: array of TZConditionalStyleItem;
    function GetItem(num: integer): TZConditionalStyleItem;
    procedure SetItem(num: integer; Value: TZConditionalStyleItem);
    procedure SetCount(value: integer);
    procedure SetAreas(Value: TZConditionalAreas);
  protected
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    function Add(): TZConditionalStyleItem; overload;
    function Add(StyleItem: TZConditionalStyleItem): TZConditionalStyleItem; overload;
    procedure Delete(num: integer);
    procedure Insert(num: integer); overload;
    procedure Insert(num: integer; StyleItem: TZConditionalStyleItem); overload;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(Source: TPersistent): boolean; virtual;
    property Areas: TZConditionalAreas read FAreas write SetAreas;
    property Count: integer read FCount write SetCount;
    property Items[num: integer]: TZConditionalStyleItem read GetItem write SetItem; default;
  end;

  TZConditionalFormatting = class (TPersistent)
  private
    FStyles: array of TZConditionalStyle;
    FCount: integer;
    procedure SetCount(Value: integer);
    function GetItem(num: integer): TZConditionalStyle;
    procedure SetItem(num: integer; Value: TZConditionalStyle);
  protected
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    function Add(): TZConditionalStyle; overload;
    function Add(Style: TZConditionalStyle): TZConditionalStyle; overload;
    function Add(ColumnNum, RowNum, AreaWidth, AreaHeight: integer): TZConditionalStyle; overload;
    procedure Clear();
    function Delete(num: integer): boolean;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(Source: TPersistent): boolean; virtual;
    property Count: integer read FCount write SetCount;
    property Items[num: integer]: TZConditionalStyle read GetItem write SetItem; default;
  end;

  //Transformations that can applied to a chart/image.
  TZETransform = class (TPersistent)
  private
    FRotate: double;
    FScaleX: double;
    FScaleY: double;
    FSkewX: double;
    FSkewY: double;
    FTranslateX: double;
    FTranslateY: double;
  protected
  public
    constructor Create();
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): boolean;
    procedure Clear();
  published
    property Rotate: double read FRotate write FRotate;
    property ScaleX: double read FScaleX write FScaleX;
    property ScaleY: double read FScaleY write FScaleY;
    property SkewX: double read FSkewX write FSkewX;
    property SkewY: double read FSkewY write FSkewY;
    property TranslateX: double read FTranslateX write FTranslateX;
    property TranslateY: double read FTranslateY write FTranslateY;
  end;

  //Common frame ancestor for Charts and Images
  TZECommonFrameAncestor = class (TPersistent)
  private
    FX: integer;
    FY: integer;
    FWidth: integer;
    FHeight: integer;
    FTransform: TZETransform;
  protected
    procedure SetX(value: integer); virtual;
    procedure SetY(value: integer); virtual;
    procedure SetWidth(value: integer); virtual;
    procedure SetHeight(value: integer); virtual;
    procedure SetTransform(const value: TZETransform); virtual;
  public
    constructor Create(); overload; virtual;
    constructor Create(AX, AY, AWidth, AHeight: integer); overload; virtual;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): boolean; virtual;
  published
    property X: integer read FX write SetX default 0;
    property Y: integer read FY write SetY default 0;
    property Width: integer read FWidth write SetWidth default 10;
    property Height: integer read FHeight write SetHeight default 10;
    property Transform: TZETransform read FTransform write SetTransform;
  end;

  //Possible chart types
  TZEChartType = (
    ZEChartTypeArea,
    ZEChartTypeBar,
    ZEChartTypeBubble,
    ZEChartTypeCircle,
    ZEChartTypeGantt,
    ZEChartTypeLine,
    ZEChartTypeRadar,
    ZEChartTypeRing,
    ZEChartTypeScatter,
    ZEChartTypeStock,
    ZEChartTypeSurface
  );

  //Specifies the rendering of bars for 3D bar charts
  TZEChartSolidType = (
    ZEChartSolidTypeCone,
    ZEChartSolidTypeCuboid,
    ZEChartSolidTypeCylinder,
    ZEChartSolidTypePyramid
  );

  //Type of symbols for a data point in a chart
  TZEChartSymbolType = (
    ZEChartSymbolTypeNone,            //No symbol should be used
    ZEChartSymbolTypeAutomatic,       //Auto select from TZEChartSymbol
    ZEChartSymbolTypeNamedSymbol      //Use selected from TZEChartSymbol
    //ZEChartSymbolTypeImage          //not for now
   );

  //Symbol to be used for a data point in a chart, used only for chart type = ZEChartSymbolTypeNamedSymbol
  TZEChartSymbol = (
    ZEChartSymbolArrowDown,
    ZEChartSymbolArrowUp,
    ZEChartSymbolArrowRight,
    ZEChartSymbolArrowLeft,
    ZEChartSymbolAsterisk,
    ZEChartSymbolCircle,
    ZEChartSymbolBowTie,
    ZEChartSymbolDiamond,
    ZEChartSymbolHorizontalBar,
    ZEChartSymbolHourglass,
    ZEChartSymbolPlus,
    ZEChartSymbolStar,
    ZEChartSymbolSquare,
    ZEChartSymbolX,
    ZEChartSymbolVerticalBar
  );

  //Position of Legend in chart
  TZEChartLegendPosition = (
    ZELegendBottom,     //Legend below the plot area
    ZELegendEnd,        //Legend on the right side of the plot area
    ZELegendStart,      //Legend on the left side of the plot area
    ZELegendTop,        //Legend above the plot area
    ZELegendBottomEnd,  //Legend in the bottom right corner of the plot area
    ZELegendBottomStart,//Legend in the bottom left corner
    ZELegendTopEnd,     //Legend in the top right corner
    ZELegendTopStart    //Legend in the top left corner
   );

  //Alignment of a legend
  TZELegendAlign = (
    ZELegengAlignStart,  //Legend aligned at the beginning of a plot area (left or top)
    ZELegendAlignCenter, //Legend aligned at the center of a plot area
    ZELegendAlignEnd     //Legend aligned at the end of a plot area (right or bottom)
  );

  //Range item
  TZEChartRangeItem = class (TPersistent)
  private
    FSheetNum: integer;
    FCol: integer;
    FRow: integer;
    FWidth: integer;
    FHeight: integer;
  protected
  public
    constructor Create(); overload;
    constructor Create(ASheetNum: integer; ACol, ARow, AWidth, AHeight: integer); overload;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): boolean; virtual;
  published
    property SheetNum: integer read FSheetNum write FSheetNum;
    property Col: integer read FCol write FCol;
    property Row: integer read FRow write FRow;
    property Width: integer read FWidth write FWidth;
    property Height: integer read FHeight write FHeight;
  end;

  //Store for Range items
  TZEChartRange = class (TPersistent)
  private
    FItems: array of TZEChartRangeItem;
    FCount: integer;
  protected
    function GetItem(num: integer): TZEChartRangeItem;
    procedure SetItem(num: integer; const Value: TZEChartRangeItem);
  public
    constructor Create();
    destructor Destroy(); override;
    function Add(): TZEChartRangeItem; overload;
    function Add(const ItemForClone: TZEChartRangeItem): TZEChartRangeItem; overload;
    function Delete(num: integer): boolean;
    procedure Clear();
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): boolean; virtual;
    property Items[num: integer]: TZEChartRangeItem read GetItem write SetItem; default;
    property Count: integer read FCount;
  end;

  //Title or Subtitle for chart, axis etc
  TZEChartTitleItem = class (TPersistent)
  private
    FText: string;
    FFont: TFont;
    FRotationAngle: integer;
    FIsDisplay: boolean;
  protected
    procedure SetFont(const value: TFont);
  public
    constructor Create(); virtual;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): boolean; virtual;
  published
    property Text: string read FText write FText;
    property Font: TFont read FFont write SetFont;
    property RotationAngle: integer read FRotationAngle write FRotationAngle default 0;
    property IsDisplay: boolean read FIsDisplay write FIsDisplay default true;
  end;

  //Chart legend
  TZEChartLegend = class (TZEChartTitleItem)
  private
    FPosition: TZEChartLegendPosition;
    FAlign: TZELegendAlign;
  protected
  public
    constructor Create(); override;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): boolean; override;
  published
    property Position: TZEChartLegendPosition read FPosition write FPosition;
    property Align: TZELegendAlign read FAlign write FAlign;
  end;

  //Chart axis
  TZEChartAxis = class (TZEChartTitleItem)
  private
    FLogarithmic: boolean;
    FReverseDirection: boolean;
    FScaleMin: double;
    FScaleMax: double;
    FAutoScaleMin: boolean;
    FAutoScaleMax: boolean;
  protected
  public
    constructor Create(); override;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): boolean; override;
  published
    property Logarithmic: boolean read FLogarithmic write FLogarithmic default false;
    property ReverseDirection: boolean read FReverseDirection write FReverseDirection default false;
    property ScaleMin: double read FScaleMin write FScaleMin;
    property ScaleMax: double read FScaleMax write FScaleMax;
    property AutoScaleMin: boolean read FAutoScaleMin write FAutoScaleMin;
    property AutoScaleMax: boolean read FAutoScaleMax write FAutoScaleMax;
  end;

  //Chart/series settings
  TZEChartSettings = class (TPersistent)
  private
    FJapanCandle: boolean;
  protected
  public
    constructor Create();
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): boolean; virtual;
  published
    property JapanCandle: boolean read FJapanCandle write FJapanCandle; // True - japanese candle. Used only for ZEChartTypeStock chart type.
  end;

  //Chart series
  TZEChartSeries = class (TPersistent)
  private
    FChartType: TZEChartType;
    FSeriesName: string;
    FSeriesNameSheet: integer;
    FSeriesNameRow: integer;
    FSeriesNameCol: integer;
    FRanges: TZEChartRange;
  protected
  public
    constructor Create();
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): boolean; virtual;
  published
                                    //For each series it's own ChartType
    property ChartType: TZEChartType read FChartType write FChartType;

                                    //If (SeriesNameRow >= 0) and (SeriesNameCol >= 0) then
                                    // text for Series label will be from cell[SeriesNameCol, SeriesNameRow] and
                                    // from property SeriesName otherwise.
    property SeriesName: string read FSeriesName write FSeriesName;

                                    //If SeriesNameSheet < 0 then uses current sheet for Series label
    property SeriesNameSheet: integer read FSeriesNameSheet write FSeriesNameSheet;
    property SeriesNameRow: integer read FSeriesNameRow write FSeriesNameRow;
    property SeriesNameCol: integer read FSeriesNameCol write FSeriesNameCol;
    property Ranges: TZEChartRange read FRanges;
  end;

  //Chart item
  TZEChart = class (TZECommonFrameAncestor)
  private
    FTitle: TZEChartTitleItem;
    FSubtitle: TZEChartTitleItem;
    FFooter: TZEChartTitleItem;
    FLegend: TZEChartLegend;
    FAxisX: TZEChartAxis;
    FAxisY: TZEChartAxis;
    FAxisZ: TZEChartAxis;
    FSecondaryAxisX: TZEChartAxis;
    FSecondaryAxisY: TZEChartAxis;
    FSecondaryAxisZ: TZEChartAxis;
    FDefaultChartType: TZEChartType;
    FView3D: boolean;
    FViewDeep: boolean;
  protected
    procedure SetTitle(const value: TZEChartTitleItem);
    procedure SetSubtitle(const value: TZEChartTitleItem);
    procedure SetFooter(const value: TZEChartTitleItem);
    procedure SetLegend(const value: TZEChartLegend);
    procedure CommonInit();
    procedure SetAxisX(const value: TZEChartAxis);
    procedure SetAxisY(const value: TZEChartAxis);
    procedure SetAxisZ(const value: TZEChartAxis);
    procedure SetSecondaryAxisX(const value: TZEChartAxis);
    procedure SetSecondaryAxisY(const value: TZEChartAxis);
    procedure SetSecondaryAxisZ(const value: TZEChartAxis);
  public
    constructor Create(); overload; override;
    constructor Create(AX, AY, AWidth, AHeight: integer); overload; override;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): boolean; override;
  published
    property AxisX: TZEChartAxis read FAxisX write SetAxisX;
    property AxisY: TZEChartAxis read FAxisY write SetAxisY;
    property AxisZ: TZEChartAxis read FAxisZ write SetAxisZ;

    property SecondaryAxisX: TZEChartAxis read FSecondaryAxisX write SetSecondaryAxisX;
    property SecondaryAxisY: TZEChartAxis read FSecondaryAxisY write SetSecondaryAxisY;
    property SecondaryAxisZ: TZEChartAxis read FSecondaryAxisZ write SetSecondaryAxisZ;

    property DefaultChartType: TZEChartType read FDefaultChartType write FDefaultChartType;

    property Title: TZEChartTitleItem read FTitle write SetTitle;
    property Subtitle: TZEChartTitleItem read FSubtitle write SetSubtitle;
    property Footer: TZEChartTitleItem read FFooter write SetFooter;
    property Legend: TZEChartLegend read FLegend write SetLegend;

    property View3D: boolean read FView3D write FView3D;
    property ViewDeep: boolean read FViewDeep write FViewDeep;
  end;

  //Store for charts on a sheet
  TZEChartStore = class (TPersistent)
  private
    FCount: integer;
    FItems: array of TZEChart;
  protected
    function GetItem(num: integer): TZEChart;
    procedure SetItem(num: integer; const Value: TZEChart);
  public
    constructor Create();
    destructor Destroy(); override;
    function Add(): TZEChart; overload;
    function Add(const ItemForClone: TZEChart): TZEChart; overload;
    function Delete(num: integer): boolean;
    procedure Clear();
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): boolean; virtual;
    property Items[num: integer]: TZEChart read GetItem write SetItem; default;
  published
    property Count: integer read FCount;
  end;

  TZCellAnchor = (ZACell, ZAAbsolute);

  // Picture item
  TZEPicture = class(TZECommonFrameAncestor)
  private
    FId: Integer;
    FRelId: Integer;
    FTitle: string;
    FDescription: string;
    FRow: Integer;
    FCol: Integer;
    FCellAnchor: TZCellAnchor;
    FHidden: Boolean;
    FFileName: string;
    FDataStream: TStream;
    function GetRelIdStr(): string;
    function GetName(): string;
  protected
    procedure CommonInit();
  public
    constructor Create(); overload; override;
    constructor Create(AX, AY, AWidth, AHeight: integer); overload; override;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): boolean; override;

    procedure AssignFromFile(AFileName: string);
    procedure AssignFromStream(AStream: TStream);
  published
    // through document
    property Id: Integer read FId write FId;
    // through worksheet
    property RelId: Integer read FRelId write FRelId;
    property RelIdStr: string read GetRelIdStr;
    property Name: string read GetName;
    property Title: string read FTitle write FTitle;
    property Description: string read FDescription write FDescription;
    property Row: Integer read FRow write FRow;
    property Col: Integer read FCol write FCol;
    property CellAnchor: TZCellAnchor read FCellAnchor write FCellAnchor;

    property Hidden: Boolean read FHidden write FHidden;
    property DataStream: TStream read FDataStream;
  end;

  // Store for pictures on a sheet
  TZEPictureStore = class (TPersistent)
  private
    FItems: TObjectList;
    function GetCount: integer;
  protected
    function GetItem(num: integer): TZEPicture;
    procedure SetItem(num: integer; const Value: TZEPicture);
  public
    constructor Create();
    destructor Destroy(); override;
    function Add(): TZEPicture; overload;
    function Add(const ItemForClone: TZEPicture): TZEPicture; overload;
    function Delete(num: integer): boolean;
    procedure Clear();
    procedure Assign(Source: TPersistent); override;
    function IsEqual(const Source: TPersistent): boolean; virtual;
    property Items[num: integer]: TZEPicture read GetItem write SetItem; default;
  published
    property Count: integer read GetCount;
  end;

  { Store pictures for worksheet }
  TZEDrawing = class(TPersistent)
  private
    FId: Integer;
    FPictureStore: TZEPictureStore;
    function GetIsEmpty(): Boolean;
  public
    constructor Create();
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;

    function AddPictureStream(ARow, ACol: Integer; AStream: TStream): TZEPicture;
    function AddPictureFile(ARow, ACol: Integer; AFileName: string): TZEPicture;

    property Id: Integer read FId write FId;
    property PictureStore: TZEPictureStore read FPictureStore;
    property IsEmpty: Boolean read GetIsEmpty;
  end;

  TZRange = class;

  /// <summary>
  /// Inherited from TPersistent. Contains a sheet of the document.
  /// </summary>
  TZSheet = class (TPersistent)
  private
    FStore: TZEXMLSS;
    FCells: array of TZCellColumn;
    FRows: array of TZRowOptions;
    FColumns: array of TZColOptions;
    FAutoFilter: string;
    FTitle: string;                     //заголовок листа
    FRowCount: integer;
    FColCount: integer;
    FTabColor: TColor;                  //цвет закладки
    FFitToPage: Boolean;
    FDefaultRowHeight: real;
    FDefaultColWidth: real;
    FMergeCells: TZMergeCells;
    FProtect: boolean;
    FRightToLeft: boolean;
    FSheetOptions: TZSheetOptions;
    FSelected: boolean;
    FPrintRows, FPrintCols: TZSheetPrintTitles;

    FCharts: TZEChartStore;
    FDrawing: TZEDrawing;

    FViewMode: TZViewMode;
    FRowBreaks:TArray<integer>;
    FColBreaks:TArray<integer>;
    FConditionalFormatting: TZConditionalFormatting;
    procedure SetConditionalFormatting(Value: TZConditionalFormatting);
    procedure SetCharts(const Value: TZEChartStore);
    procedure SetColumn(num: integer; const Value:TZColOptions);
    function  GetColumn(num: integer): TZColOptions;
    procedure SetRow(num: integer; const Value:TZRowOptions);
    function  GetRow(num: integer): TZRowOptions;
    function  GetSheetOptions(): TZSheetOptions;
    procedure SetSheetOptions(Value: TZSheetOptions);
    procedure SetPrintCols(const Value: TZSheetPrintTitles);
    procedure SetPrintRows(const Value: TZSheetPrintTitles);
  protected
    procedure SetColWidth(num: integer; const Value: real); virtual;
    function  GetColWidth(num: integer): real; virtual;
    procedure SetRowHeight(num: integer; const Value: real); virtual;
    function  GetRowHeight(num: integer): real; virtual;
    procedure SetDefaultColWidth(const Value: real); virtual;
    procedure SetDefaultRowHeight(const Value: real); virtual;
    function  GetCell(ACol, ARow: integer): TZCell; virtual;
    procedure SetCell(ACol, ARow: integer; const Value: TZCell); virtual;
    function  GetCellRef(ACol: string; ARow: integer): TZCell; virtual;
    procedure SetCellRef(ACol: string; ARow: integer; const Value: TZCell); virtual;
    function  GetRowCount: integer; virtual;
    procedure SetRowCount(const Value: integer); virtual;
    function  GetColCount: integer; virtual;
    procedure SetColCount(const Value: integer); virtual;
    function  GetRange(AC1,AR1,AC2,AR2: integer): TZRange; virtual;
    //procedure SetRange(AC1,AR1,AC2,AR2: integer; const Value: TZRange); virtual;
    function  GetRangeRef(AFromCol: string; AFromRow: Integer; AToCol: string; AToRow: integer): TZRange; virtual;
    //procedure SetRangeRef(AFrom, ATo: string; const Value: TZRange); virtual;
    function GetSheetIndex(): integer;
  public
    constructor Create(AStore: TZEXMLSS); virtual;
    destructor Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    procedure Clear(); virtual;
    procedure InsertRows(ARow, ACount: Integer);
    procedure CopyRows(ARowDst, ARowSrc, ACount: Integer);
    procedure SetCorrectTitle(const Value: string);
    /// <summary>
    /// Get or set the width (in points) of column num in the sheet.
    /// </summary>
    property ColWidths[num: integer]: real read GetColWidth write SetColWidth;
    /// <summary>
    /// Options of column num.
    /// </summary>
    property Columns[num: integer]: TZColOptions read GetColumn write SetColumn;
    /// <summary>
    /// Specifies various properties of the Row num.
    /// </summary>
    property Rows[num: integer]: TZRowOptions read GetRow write SetRow;
    property Range[AC1,AR1,AC2,AR2: integer]: TZRange read GetRange{ write SetRange};
    property RangeRef[AFromCol: string; AFromRow: Integer; AToCol: string; AToRow: integer]: TZRange read GetRangeRef{ write SetRangeRef};
    /// <summary>
    /// Get or set the height (in points) of row num in the sheet.
    /// </summary>
    property RowHeights[num: integer]: real read GetRowHeight write SetRowHeight;
    /// <summary>
    /// Default column width.
    /// </summary>
    property DefaultColWidth: real read FDefaultColwidth write SetDefaultColWidth;// default 48;
    /// <summary>
    /// Default row height.
    /// </summary>
    property DefaultRowHeight: real read FDefaultRowHeight write SetDefaultRowHeight;// default 12.75;
    /// <summary>
    /// Cell at the intersection of column ACol and row ARow.
    /// </summary>
    property Cell[ACol, ARow: integer]: TZCell read GetCell write SetCell; default;
    property CellRef[ACol: string; ARow: integer]: TZCell read GetCellRef write SetCellRef;
    property AutoFilter: string read FAutoFilter write FAutoFilter;
    /// <summary>
    /// Indicates whether or not this sheet is protected. <br />False by default.
    /// </summary>
    property Protect: boolean read FProtect write FProtect default false;
    property TabColor: TColor read FTabColor write FTabColor default ClWindow;
    property FitToPage: Boolean read FFitToPage write FFitToPage default false;
    /// <summary>
    /// Sheet title.
    /// </summary>
    property Title: string read FTitle write FTitle;
    property SheetIndex: integer read GetSheetIndex;
    /// <summary>
    /// Specifies the number of rows in the sheet.
    /// </summary>
    property RowCount: integer read GetRowCount write SetRowCount;
    /// <summary>
    /// If True, the window displays from right to left else window displays from left to right. <br />False by default.
    /// </summary>
    property RightToLeft: boolean read FRightToLeft write FRightToLeft default false;
    /// <summary>
    /// Specifies the number of columns in the sheet.
    /// </summary>
    property ColCount: integer read GetColCount write SetColCount;
    /// <summary>
    /// Merged cells.
    /// </summary>
    property MergeCells: TZMergeCells read FMergeCells write FMergeCells;
    /// <summary>
    /// Specifies various properties of the sheet.
    /// </summary>
    property SheetOptions: TZSheetOptions read GetSheetOptions write SetSheetOptions;
    /// <summary>
    /// Indicates whether or not this sheet is selecteded.
    /// </summary>
    property Selected: boolean read FSelected write FSelected;
    property WorkBook: TZEXMLSS read FStore;
    property RowsToRepeat: TZSheetPrintTitles read FPrintRows write SetPrintRows;
    property ColsToRepeat: TZSheetPrintTitles read FPrintCols write SetPrintCols;
    /// <summary>
    /// Conditional formatting for sheet.
    /// </summary>
    property ConditionalFormatting: TZConditionalFormatting read FConditionalFormatting write SetConditionalFormatting;
    property Charts: TZEChartStore read FCharts write SetCharts;
    property Drawing: TZEDrawing read FDrawing;
    property ViewMode: TZViewMode read FViewMode write FViewMode;
    property RowBreaks: TArray<integer> read FRowBreaks write FRowBreaks;
    property ColBreaks: TArray<integer> read FColBreaks write FColBreaks;
  end;

  /// <summary>
  /// Document sheets.
  /// </summary>
  TZSheets = class (TPersistent)
  private
    FStore: TZEXMLSS;
    FSheets: array of TZSheet;
    FCount : integer;
    procedure SetSheetCount(const Value: integer);
    procedure SetSheet(num: integer; Const Value: TZSheet);
    function  GetSheet(num: integer): TZSheet;
  protected
  public
    constructor Create(AStore: TZEXMLSS); virtual;
    destructor  Destroy(); override;
    procedure Assign(Source: TPersistent); override;
    /// <summary>
    /// Number of sheets in the document.
    /// </summary>
    property Count: integer read FCount write SetSheetCount;
    /// <summary>
    /// Document's sheet num.
    /// </summary>
    property Sheet[num: integer]: TZSheet read GetSheet write SetSheet; default;
  end;

  TZRange = class(TPersistent)
  private
    FSheet: TZSheet;
    FC1,FR1,FC2,FR2: Integer;

    function HasStyle: Boolean;
    procedure ApplyStyleValue(proc: TProc<TZStyle>);
    function GetVerticalAlignment(): TZVerticalAlignment;
    procedure SetVerticalAlignment(const Value: TZVerticalAlignment);
    function GetHorizontalAlignment(): TZHorizontalAlignment;
    procedure SetHorizontalAlignment(const Value: TZHorizontalAlignment);
    function GetBgColor(): TColor;
    procedure SetBgColor(const Value: TColor);
    function GetFontColor(): TColor;
    procedure SetFontColor(const Value: TColor);
    function GetFontSize(): Byte;
    procedure SetFontSize(const Value: Byte);
    function GetFontStyle(): TFontStyles;
    procedure SetFontStyle(const Value: TFontStyles);
    function GetBorderStyle(Num: TZBordersPos): TZBorderType;
    procedure SetBorderStyle(Num: TZBordersPos; const Value: TZBorderType);
    function GetBorderWidht(Num: TZBordersPos): Byte;
    procedure SetBorderWidht(Num: TZBordersPos; const Value: Byte);
    function GetBorderColor(Num: TZBordersPos): TColor;
    procedure SetBorderColor(Num: TZBordersPos; const Value: TColor);
    function GetWrapText(): Boolean;
    procedure SetWrapText(const Value: Boolean);
    function GetVerticalText(): Boolean;
    procedure SetVerticalText(const Value: Boolean);
    function GetNumberFormat(): string;
    procedure SetNumberFormat(const Value: string);
  protected
  public
    constructor Create(ASheet: TZSheet; AFC1, AFR1, AFC2, AFR2: Integer); virtual;
    procedure Assign(Source: TPersistent); override;
    destructor Destroy(); override;
    property VerticalAlignment: TZVerticalAlignment read GetVerticalAlignment write SetVerticalAlignment;
    property HorizontalAlignment: TZHorizontalAlignment read GetHorizontalAlignment write SetHorizontalAlignment;
    property BgColor: TColor read GetBgColor write SetBgColor;
    property FontColor: TColor read GetFontColor write SetFontColor;
    property FontSize: Byte read GetFontSize write SetFontSize;
    property FontStyle: TFontStyles read GetFontStyle write SetFontStyle;
    property BorderStyle[num: TZBordersPos]: TZBorderType read GetBorderStyle write SetBorderStyle;
    property BorderWidht[num: TZBordersPos]: Byte read GetBorderWidht write SetBorderWidht;
    property BorderColor[num: TZBordersPos]: TColor read GetBorderColor write SetBorderColor;
    property WrapText: Boolean read GetWrapText write SetWrapText;
    property VerticalText: Boolean read GetVerticalText write SetVerticalText;
    property NumberFormat: string read GetNumberFormat write SetNumberFormat;
    procedure Merge();
    procedure Clear();
  end;

  /// <summary>
  /// Document properties
  /// </summary>
  TZEXMLDocumentProperties = class(TPersistent)
  private
    FAuthor      : string;
    FLastAuthor  : string;
    FCreated     : TdateTime;
    FCompany     : string;
    FVersion     : string; // - should be integer by Spec but hardcoded float in real MS Office apps
    FWindowHeight: word;
    FWindowWidth : word;
    FWindowTopX  : integer;
    FWindowTopY  : integer;
    FModeR1C1    : boolean;
  protected
    procedure SetAuthor(const Value: string);
    procedure SetLastAuthor(const Value: string);
    procedure SetCompany(const Value: string);
    procedure SetVersion(const Value: string);
  public
    constructor Create(); virtual;
    procedure Assign(Source: TPersistent); override;
  published
    /// <summary>
    /// Author of the document
    /// </summary>
    property Author: string read FAuthor write SetAuthor;
    /// <summary>
    /// Author of last changes in the document
    /// </summary>
    property LastAuthor: string read FLastAuthor write SetLastAuthor;
    /// <summary>
    /// Date and time of document creation
    /// </summary>
    property Created: TDateTime read FCreated write FCreated;
    /// <summary>
    /// Company name
    /// </summary>
    property Company: string read FCompany write SetCompany;
    /// <summary>
    /// Document version
    /// </summary>
    property Version: string read FVersion write SetVersion;
    /// <summary>
    /// Enabled R1C1 style in Excel. <br />False by default
    /// </summary>
    property ModeR1C1: boolean read FModeR1C1 write FModeR1C1 default false;
    property WindowHeight: word read FWindowHeight write FWindowHeight default 20000;
    property WindowWidth: word read FWindowWidth write FWindowWidth default 20000;
    property WindowTopX: integer read FWindowTopX write FWindowTopX default 150;
    property WindowTopY: integer read FWindowTopY write FWindowTopY default 150;
  end;

  /// <summary>
  /// Contains spreadsheet document
  /// </summary>
  TZEXMLSS = class (TComponent)
  private
    FSheets: TZSheets;
    FDocumentProperties: TZEXMLDocumentProperties;
    FStyles: TZStyles;
    FHorPixelSize: real;
    FVertPixelSize: real;
    FDefaultSheetOptions: TZSheetOptions;
    procedure SetHorPixelSize(Value: real);
    procedure SetVertPixelSize(Value: real);
    function  GetDefaultSheetOptions(): TZSheetOptions;
    procedure SetDefaultSheetOptions(Value: TZSheetOptions);
  protected
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy();override;
    procedure Assign(Source: TPersistent); override;
    procedure GetPixelSize(hdc: HWND);
    property Sheets: TZSheets read FSheets write FSheets;
    /// <summary>
    /// Total not-empty drawings count from all sheets
    /// </summary>
    function DrawingCount(): Integer;
    function GetDrawing(num: Integer): TZEDrawing;
    function GetDrawingSheetNum(Value: TZEDrawing): Integer;
  published
    property Styles: TZStyles read FStyles write FStyles;
    property DefaultSheetOptions: TZSheetOptions read GetDefaultSheetOptions write SetDefaultSheetOptions;
    property DocumentProperties: TZEXMLDocumentProperties read FDocumentProperties write FDocumentProperties;
    property HorPixelSize: real read FHorPixelSize write SetHorPixelSize;  //размер пикселя по горизонтали
    property VertPixelSize: real read FVertPixelSize write SetVertPixelSize;  //размер пикселя по вертикали
  end;

/// <summary>
/// Convert TColor to Hex RGB
/// </summary>
function ColorToHTMLHex(Color: TColor): string;
/// <summary>
/// Convert Hex RGB (string) to TColor
/// </summary>
function HTMLHexToColor(value: string): TColor;
/// <summary>
/// Convert ARGB (string) to TColor
/// </summary>
function ARGBToColor(value: string): TColor;
/// <summary>
/// Convert pixels to point
/// </summary>
function PixelToPoint(inPixel: integer; PixelSizeMM: real = 0.265): real;

//Перевести типографский пункт (point) в пиксели
function PointToPixel(inPoint: real; PixelSizeMM: real = 0.265): integer;

//Перевести типографский пункт (point) в мм
function PointToMM(inPoint: real): real;

//Перевести мм в типографский пункт (point)
function MMToPoint(inMM: real): real;

function ZEIsFontsEquals(const Font1, Font2: TFont): boolean;

//Переводит дату в строку для XML (YYYY-MM-DDTHH:MM:SS[.mmm])
function ZEDateTimeToStr(ATime: TDateTime; Addmms: boolean = false): string;

//YYYY-MM-DDTHH:MM:SS[.mmm] to DateTime
function TryZEStrToDateTime(const AStrDateTime: string; out retDateTime: TDateTime): boolean;

function IntToStrN(value: integer; NullCount: integer): string;

implementation

uses zeformula;

//Переводит число в строку минимальной длины NullCount
//TODO: надо глянуть что с функциями в FlyLogReader-е
//INPUT
//      value: integer     - число
//      NullCount: integer - кол-во знаков в строке
//RETURN
//      string
function IntToStrN(value: integer; NullCount: integer): string;
var t, k: integer;
begin
  t := value;
  k := 0;
  if (t = 0) then
    k := 1;
  while t > 0 do begin
    inc(k);
    t := t div 10;
  end;
  result := IntToStr(value);
  for t := 1 to (NullCount - k) do
    result := '0' + result;
end; //IntToStrN

//Переводит дату в строку для XML (YYYY-MM-DDTHH:MM:SS[.mmm])
//INPUT
//      ATime: TDateTime - нужная дата/время
//      Addmms: boolean  - need add ms to result
function ZEDateTimeToStr(ATime: TDateTime; Addmms: boolean = false): string;
var HH, MM, SS, MS: word;
begin
  DecodeDate(ATime, HH, MM, SS);
  result := IntToStrN(HH, 4) + '-' + IntToStrN(MM, 2) + '-' + IntToStrN(SS, 2) + 'T';
  DecodeTime(ATime, HH, MM, SS, MS);
  result := result + IntToStrN(HH, 2) + ':' + IntToStrN(MM, 2) + ':' + IntToStrN(SS, 2);
  if (Addmms) then
    Result := Result + '.' + IntToStrN(MS, 3);
end;

//Try convert string (YYYY-MM-DDTHH:MM:SS[.mmm]) to datetime
//INPUT
//  const AStrDateTime: string  - string with datetime
//  out retDateTime: TDateTime  - returns on success datetime
//RETURN
//      boolean - true - ok
function TryZEStrToDateTime(const AStrDateTime: string; out retDateTime: TDateTime): boolean;
var a: array [0..10] of word;
  i, l: integer;
  s, ss: string;
  count: integer;
  ch: char;
  datedelimeters: integer;
  istimesign: boolean;
  timedelimeters: integer;
  istimezone: boolean;
  lastdateindex: integer;
  tmp: integer;
  msindex: integer;
  tzindex: integer;
  timezonemul: integer;
  _ms: word;

  function TryAddToArray(const ST: string): boolean;
  begin
    if (count > 10) then begin
      Result := false;
      exit;
    end;
    Result := TryStrToInt(ST, tmp);
    if (Result) then begin
      a[Count] := word(tmp);
      inc(Count);
    end
  end;

  procedure _CheckDigits();
  var _l: integer;
  begin
    _l := length(s);
    if (_l > 0) then begin
      if (_l > 4) then begin//it is not good
        if (istimesign) then begin
          // HHMMSS?
          if (_l = 6) then begin
            ss := copy(s, 1, 2);
            if (TryAddToArray(ss)) then begin
              ss := copy(s, 3, 2);
              if (TryAddToArray(ss)) then begin
                ss := copy(s, 5, 2);
                if (not TryAddToArray(ss)) then
                  Result := false;
              end else
                Result := false;
            end else
              Result := false
          end else
            Result := false;
        end else begin
          // YYYYMMDD?
          if (_l = 8) then begin
            ss := copy(s, 1, 4);
            if (not TryAddToArray(ss)) then
              Result := false
            else begin
              ss := copy(s, 5, 2);
              if (not TryAddToArray(ss)) then
                Result := false
              else begin
                ss := copy(s, 7, 2);
                if (not TryAddToArray(ss)) then
                  Result := false;
              end;
            end;
          end else
            Result := false;
        end;
      end else
        if (not TryAddToArray(s)) then
          Result := false;
    end; //if
    if (Count > 10) then
      Result := false;
    s := '';
  end;

  procedure _processDigit();
  begin
    s := s + ch;
  end;

  procedure _processTimeSign();
  begin
    istimesign := true;
    if (count > 0) then
      lastdateindex := count;

    _CheckDigits();
  end;

  procedure _processTimeDelimiter();
  begin
    _CheckDigits();
    inc(timedelimeters)
  end;

  procedure _processDateDelimiter();
  begin
    _CheckDigits();
    if (istimesign) then begin
      tzindex := count;
      istimezone := true;
      timezonemul := -1;
    end else
      inc(datedelimeters);
  end;

  procedure _processMSDelimiter();
  begin
    _CheckDigits();
    msindex := count;
  end;

  procedure _processTimeZoneSign();
  begin
    _CheckDigits();
    istimezone := true;
  end;

  procedure _processTimeZonePlus();
  begin
    _CheckDigits();
    istimezone := true;
    timezonemul := -1;
  end;

  function _TryGetDateTime(): boolean;
  var _time, _date: TDateTime;
  begin
    //Result := true;
    if (msindex >= 0) then
      _ms := a[msindex];
    if (lastdateindex >= 0) then begin
      Result := TryEncodeDate(a[0], a[1], a[2], _date);
      if (Result) then begin
        Result := TryEncodeTime(a[lastdateindex + 1], a[lastdateindex + 2], a[lastdateindex + 3], _ms, _time);
        if (Result) then
          retDateTime := _date + _time;
      end;
    end else
      Result := TryEncodeTime(a[lastdateindex + 1], a[lastdateindex + 2], a[lastdateindex + 3], _ms, retDateTime);
  end;

  function _TryGetDate(): boolean;
  begin
    if (datedelimeters = 0) and (timedelimeters >= 2) then begin
      if (msindex >= 0) then
        _ms := a[msindex];
      result := TryEncodeTime(a[0], a[1], a[2], _ms, retDateTime);
    end else if (count >= 3) then
      Result := TryEncodeDate(a[0], a[1], a[2], retDateTime)
    else
      Result := false;
  end;

begin
  Result := true;
  datedelimeters := 0;
  istimesign := false;
  timedelimeters := 0;
  istimezone := false;
  lastdateindex := -1;
  msindex := -1;
  tzindex := -1;
  timezonemul := 0;
  _ms := 0;
  FillChar(a, sizeof(a), 0);

  l := length(AStrDateTime);
  s := '';
  count := 0;
  for i := 1 to l do begin
    ch := AStrDateTime[i];
    case (ch) of
      '0'..'9': _processDigit();
      't', 'T': _processTimeSign();
      '-':      _processDateDelimiter();
      ':':      _processTimeDelimiter();
      '.', ',': _processMSDelimiter();
      'z', 'Z': _processTimeZoneSign();
      '+':      _processTimeZonePlus();
    end;
    if (not Result) then
      break
  end;

  if (Result and (s <> '')) then
    _CheckDigits();

  if (Result) then begin
    if (istimesign) then
      Result := _TryGetDateTime()
    else
      Result := _TryGetDate();
  end;
end; //TryZEStrToDateTime

//Checks is Font1 equal Font2
//INPUT
//  const Font1: TFont
//  const Font2: TFont
//RETURN
//    boolean - true - fonts are equals
function ZEIsFontsEquals(const Font1, Font2: TFont): boolean;
begin
  Result := Assigned(Font1) and (Assigned(Font2));
  if (Result) then begin
    Result := false;
    if (Font1.Color <> Font2.Color) then
      exit;

    if (Font1.Height <> Font2.Height) then
      exit;

    if (Font1.Name <> Font2.Name) then
      exit;

    if (Font1.Pitch <> Font2.Pitch) then
      exit;

    if (Font1.Size <> Font2.Size) then
      exit;

    if (Font1.Style <> Font2.Style) then
      exit;

    Result := true;
  end;
end; //ZEIsFontsEquals

function ColorToHTMLHex(Color: TColor): string;
var _RGB: integer;
begin
  _RGB := ColorToRGB(Color);
  //result := IntToHex(GetRValue(_RGB), 2) + IntToHex(GetGValue(_RGB), 2) + IntToHex(GetBValue(_RGB), 2);
  result := IntToHex(byte(_RGB), 2) + IntToHex(byte(_RGB shr 8), 2) + IntToHex(byte(_RGB shr 16), 2);
end;

function HTMLHexToColor(value: string): TColor;
var a: array [0..2] of integer;
  i, n, t: integer;
begin
  result := 0;
  if (value > '') then begin
    value := UpperCase(value);
    {$HINTS OFF}
    FillChar(a, sizeof(a), 0);
    {$HINTS ON}
    n := 0;
    if value[1] = '#' then delete(value, 1, 1);
    //А что, если будут цвета типа "black"?  {tut}
    for i := 1 to length(value) do begin
      if n > 2 then break;
      case value[i] of
        '0'..'9': t := ord(value[i]) - 48;
        'A'..'F': t := 10 + ord(value[i]) - 65;
        else
          t := 0;
      end;
      a[n] := a[n] * 16 + t;
      if i mod 2 = 0 then inc(n);
    end;
    result := a[2] shl 16 or a[1] shl 8 or a[0]//RGB(a[0], a[1], a[2]);
    //a[2] shl 16 or a[1] shl 8 or a[0]; = RGB
  end;
end;

function ARGBToColor(value: string): TColor;
var a: array [0..2] of integer; i, n, t: integer;
begin
  result := 0;
  if (value > '') then begin
    value := UpperCase(value);
    {$HINTS OFF}
    FillChar(a, sizeof(a), 0);
    {$HINTS ON}
    n := 0;
    if value[1] = '#' then delete(value, 1, 1);
    for i := 3 to length(value) do begin
      if n > 2 then break;
      case value[i] of
        '0'..'9': t := ord(value[i]) - 48;
        'A'..'F': t := 10 + ord(value[i]) - 65;
        else
          t := 0;
      end;
      a[n] := a[n] * 16 + t;
      if i mod 2 = 0 then
        inc(n);
    end;
    result := a[2] shl 16 or a[1] shl 8 or a[0]//RGB(a[0], a[1], a[2]);
    //a[2] shl 16 or a[1] shl 8 or a[0]; = RGB
  end;
end;

//Перевести пиксели в типографский пункт (point) ###
//Input
//      inPixel: integer          - размер в пикселях
//      PixelSizeMM: real = 0.265 - размер пикселя
function PixelToPoint(inPixel: integer; PixelSizeMM: real = 0.265): real;
begin
  result := inPixel * PixelSizeMM / _PointToMM;
  //и оставим 2 знака после запятой ^_^
  result := round(result * 100) / 100;
end;

//Перевести типографский пункт (point) в пиксели ###
//Input
//      inPoint: integer          - размер в пикселях
//      PixelSizeMM: real = 0.265 - размер пикселя
function PointToPixel(inPoint: real; PixelSizeMM: real = 0.265): integer;
begin
  result := round(inPoint * _PointToMM / PixelSizeMM);
end;

//Перевести типографский пункт (point) в мм  ###
//Input
//      inPoint: integer - размер в пунктах
function PointToMM(inPoint: real): real;
begin
  result := round(inPoint * _PointToMM * 100) / 100;
end;

//Перевести мм в типографский пункт (point)###
//Input
//      inMM: integer - размер в пунктах
function MMToPoint(inMM: real): real;
begin
  result := round(inMM / _PointToMM * 100) / 100;
end;

////::::::::::::: TZBorderStyle :::::::::::::::::////

constructor TZBorderStyle.Create();
begin
  FWeight := 0;
  FColor := clBlack;
  FLineStyle := ZENone;
end;

procedure TZBorderStyle.Assign(Source: TPersistent);
begin
  if Source is TZBorderStyle then begin
    Weight := (Source as TZBorderStyle).Weight;
    Color := (Source as TZBorderStyle).Color;
    LineStyle := (Source as TZBorderStyle).LineStyle;
  end else
    inherited Assign(Source);
end;

function TZBorderStyle.IsEqual(Source: TPersistent): boolean;
var zSource: TZBorderStyle;
begin
  Result := false;
  if not (Source is TZBorderStyle) then
    exit;
  zSource := Source as TZBorderStyle;

  if Self.LineStyle <> zSource.LineStyle then
    exit;

  if Self.Color <> zSource.Color then
    exit;

  if Self.Weight <> zSource.Weight then
    exit;

  Result := true;
end;

procedure TZBorderStyle.SetLineStyle(const Value: TZBorderType);
begin
  FLineStyle := Value;
end;

procedure TZBorderStyle.SetWeight(const Value: byte);
begin
  FWeight := Value;
  if FWeight > 3 then
    FWeight := 3;
end;

procedure TZBorderStyle.SetColor(const Value: TColor);
begin
  FColor := Value;
end;

////::::::::::::: TZBorder :::::::::::::::::////

constructor TZBorder.Create();
var i:integer;
begin
  for i := 0 to 5 do
    FBorder[i]:= TZBorderStyle.Create();
end;

destructor TZBorder.Destroy();
var i: integer;
begin
  for i := 0 to 5 do
    FreeAndNil(FBorder[i]);
  inherited Destroy;
end;

procedure TZBorder.Assign(Source: TPersistent);
var zSource: TZBorder; i: TZBordersPos;
begin
  if (Source is TZBorder) then begin
    zSource := Source as TZBorder;
    for i := bpLeft to bpDiagonalRight do
      Border[i].Assign(zSource.Border[i]);
  end else
    inherited Assign(Source);
end;

function TZBorder.IsEqual(Source: TPersistent): boolean;
var zSource: TZBorder; i: TZBordersPos;
begin
  Result := false;
  if not (Source is TZBorder) then
    exit;
  zSource := Source as TZBorder;
  for i := bpLeft to bpDiagonalRight do
    if not FBorder[Ord(i)].IsEqual(zSource.Border[i]) then
      exit;
  Result := True;
end;

//установить стиль границы
procedure TZBorder.SetBorder(Num: TZBordersPos; Const Value: TZBorderStyle);
begin
  if (Num >= bpLeft) and (Num <= bpDiagonalRight) then
    Border[num].Assign(Value);
end;

//прочитать стиль границы
function TZBorder.GetBorder(Num: TZBordersPos): TZBorderStyle;
begin
  if (Num >= bpLeft) and (Num <= bpDiagonalRight) then
    result := FBorder[Ord(Num)]
  else
    result := nil;
end;

////::::::::::::: TZAlignment :::::::::::::::::////

constructor TZAlignment.Create();
begin
  FIndent       := 0;
  FRotate       := 0;
  FHorizontal   := ZHAutomatic;
  FShrinkToFit  := false;
  FVertical     := ZVAutomatic;
  FVerticalText := false;
  FWrapText     := false;
end;

procedure TZAlignment.Assign(Source: TPersistent);
var zSource: TZAlignment;
begin
  if (Source is TZAlignment) then begin
    zSource := Source as TZAlignment;
    FHorizontal := zSource.Horizontal;
    FIndent := zSource.Indent;
    FRotate := zSource.Rotate;
    FShrinkToFit := zSource.ShrinkToFit;
    FVertical := zSource.Vertical;
    FVerticalText := zSource.VerticalText;
    FWrapText := zSource.WrapText;
  end else
    inherited Assign(Source);
end;

function TZAlignment.IsEqual(Source: TPersistent): boolean;
var zSource: TZAlignment;
begin
  result := true;
  if (Source is TZAlignment) then begin
    zSource := Source as TZAlignment;
    if Horizontal <> zSource.Horizontal then
      exit(false);
    if Indent <> zSource.Indent then
      exit(false);
    if Rotate <> zSource.Rotate then
      exit(false);
    if ShrinkToFit <> zSource.ShrinkToFit then
      exit(false);
    if Vertical <> zSource.Vertical then
      exit(false);
    if VerticalText <> zSource.VerticalText then
      exit(false);
    if WrapText <> zSource.WrapText then
      exit(false);
  end else
    result := false;
end;

procedure TZAlignment.SetHorizontal(const Value: TZHorizontalAlignment);
begin
  FHorizontal := Value;
end;

procedure TZAlignment.SetIndent(const Value: integer);
begin
  FIndent := Value;
end;

procedure TZAlignment.SetRotate(const Value: TZCellTextRotate);
begin
  FRotate := Value;
end;

procedure TZAlignment.SetShrinkToFit(const Value: boolean);
begin
  FShrinkToFit := Value;
end;

procedure TZAlignment.SetVertical(const Value: TZVerticalAlignment);
begin
  FVertical := Value;
end;

procedure TZAlignment.SetVerticalText(const Value:boolean);
begin
  FVerticalText := Value;
end;

procedure TZAlignment.SetWrapText(const Value: boolean);
begin
  FWrapText := Value;
end;

////::::::::::::: TZStyle :::::::::::::::::////
// about default font in Excel - http://support.microsoft.com/kb/214123
constructor TZStyle.Create();
begin
  FFont           := TFont.Create();
  FFont.Size      := 10;
  FFont.Name      := 'Arial';
  FFont.Color     := ClBlack;
  FBorder         := TZBorder.Create();
  FAlignment      := TZAlignment.Create();
  FBGColor        := clWindow;
  FPatternColor   := clWindow;
  FCellPattern    := ZPNone;
  FNumberFormat   := '';
  FNumberFormatId := -1;
  FProtect        := true;
  FHideFormula    := false;
end;

destructor TZStyle.Destroy();
begin
  FreeAndNil(FFont);
  FreeAndNil(FBorder);
  FreeAndNil(FAlignment);
  inherited Destroy();
end;

procedure TZStyle.Assign(Source: TPersistent);
var zSource: TZStyle;
begin
  if (Source is TZStyle) then begin
    zSource := Source as TZStyle;
    FFont.Assign(zSource.Font);
    FBorder.Assign(zSource.Border);
    FAlignment.Assign(zSource.Alignment);
    FBGColor        := zSource.BGColor;
    FPatternColor   := zSource.PatternColor;
    FCellPattern    := zSource.CellPattern;
    FNumberFormat   := zSource.NumberFormat;
    FNumberFormatId := zSource.NumberFormatId;
    FProtect        := zSource.Protect;
    FHideFormula    := zSource.HideFormula;
  end else
    inherited Assign(Source);
end;

function TZStyle.IsEqual(Source: TPersistent): boolean;
var zSource: TZStyle;
begin
  Result := False;
  if not (Source is TZStyle) then exit;
  zSource := Source as TZStyle;

  if not Border.IsEqual(zSource.Border) then
    exit;
  if not self.Alignment.IsEqual(zSource.Alignment) then
    exit;
  if BGColor <> zSource.BGColor then
    exit;
  if PatternColor <> zSource.PatternColor then
    exit;
  if CellPattern <> zSource.CellPattern then
    exit;
  if NumberFormat <> zSource.NumberFormat then
    exit;
  if NumberFormatId <> zSource.NumberFormatId then
    exit;
  if Protect <> zSource.Protect then
    exit;
  if HideFormula <> zSource.HideFormula then
    exit;
  Result := ZEIsFontsEquals(FFont, zSource.Font);
end;

procedure TZStyle.SetFont(const Value: TFont);
begin
  FFont.Assign(Value);
end;

procedure TZStyle.SetBorder(const Value: TZBorder);
begin
  FBorder.Assign(Value);
end;

procedure TZStyle.SetAlignment(const Value: TZAlignment);
begin
  FAlignment.Assign(Value);
end;

procedure TZStyle.SetBGColor(const Value: TColor);
begin
  FBGColor := Value;
end;

procedure TZStyle.SetPatternColor(const Value: TColor);
begin
  FPatternColor := Value;
end;

procedure TZStyle.SetCellPattern(const Value: TZCellPattern);
begin
  FCellPattern := Value;
end;

procedure TZStyle.SetNumberFormat(const Value: string);
begin
  FNumberFormat := Value;
  FNumberFormatId := -1;
end;

////::::::::::::: TZStyles :::::::::::::::::////

constructor TZStyles.Create();
begin
  FDefaultStyle := TZStyle.Create();
  FCount := 0;
end;

destructor TZStyles.Destroy();
var i: integer;
begin
  FreeAndNil(FDefaultStyle);
  for i:= 0 to FCount - 1  do
    FreeAndNil(FStyles[i]);
  SetLength(FStyles, 0);
  FStyles := nil;
  inherited Destroy();
end;

function TZStyles.Find(const Style: TZStyle): integer;
var i: integer;
begin
  result := -2;
  if DefaultStyle.IsEqual(Style) then
    exit(-1);

  for i := 0 to Count - 1 do
    if Items[i].IsEqual(Style) then begin
      Exit(i);
    end;
end;

function TZStyles.Add(const Style: TZStyle; CheckMatch: boolean = false): integer;
begin
  result := -2;
  if CheckMatch then
     result := Find(Style);
  if result = -2 then begin
    Count := Count + 1;
    Items[Count - 1].Assign(Style);
    result := Count - 1;
  end;
end;

procedure TZStyles.Assign(Source: TPersistent);
var zzz: TZStyles; i: integer;
begin
  if (Source is TZStyles) then begin
    zzz := Source as TZStyles;
    FDefaultStyle.Assign(zzz.DefaultStyle);
    Count := zzz.Count;
    for i := 0 to Count - 1 do
      FStyles[i].Assign(zzz[i]);
  end else
    inherited Assign(Source);
end;

procedure TZStyles.SetDefaultStyle(const Value: TZStyle);
begin
  if not(Value <> nil) then
    FDefaultStyle.Assign(Value);
end;

procedure TZStyles.SetCount(const Value: integer);
var i: integer;
begin
  if FCount < Value then begin
    SetLength(FStyles, Value);
    for i := FCount to Value - 1 do begin
      FStyles[i] := TZStyle.Create;
      FStyles[i].Assign(FDefaultStyle);
    end;
  end else begin
    for i:= Value to FCount - 1  do
      FreeAndNil(FStyles[i]);
    SetLength(FStyles, Value);
  end;
  FCount := Value;
end;

function TZStyles.GetStyle(num: integer): TZStyle;
begin
  if (num >= 0) and (num < Count) then
    result := FStyles[num]
  else
    result := DefaultStyle;
end;

procedure TZStyles.SetStyle(num: integer; const Value: TZStyle);
begin
  if (num >= 0) and (num < Count) then
    FStyles[num].Assign(Value)
  else
  if num = -1 then
    DefaultStyle.Assign(Value);
end;

function TZStyles.DeleteStyle(num: integer):integer;
begin
  if (num >= 0) and (num < Count) then begin
    FreeAndNil(FStyles[num]);
    System.Move(FStyles[num+1], FStyles[num], (Count - num - 1) * SizeOf(FStyles[num]));
    FStyles[Count - 1] := nil;
    dec(FCount);
    setlength(FStyles, FCount);
    // при удалении глянуть на ячейки - изменить num на 0, остальные сдвинуть.
    result := 0;
  end else
    result := -1;
end;

procedure TZStyles.Clear();
var i: integer;
begin
  for i:= 0 to FCount - 1 do
    FreeAndNil(FStyles[i]);
  SetLength(FStyles,0);
  FCount := 0;
end;

////::::::::::::: TZCell :::::::::::::::::////

constructor TZCell.Create();
begin
  FFormula           := '';
  FData              := '';
  FHref              := '';
  FComment           := '';
  FCommentAuthor     := '';
  FHRefScreenTip     := '';
  FCellType          := ZEString;
  FCellStyle         := -1; //по дефолту
  FAlwaysShowComment := false;
  FShowComment       := false;
end;

function TZCell.GetDataAsDateTime(): TDateTime;
var b: boolean;
begin
  b := false;
  if (FData = '') then
    Result := 0
  else if (not TryZEStrToDateTime(FData, Result)) then begin
    //If cell type is number then try convert "float" to datetime
    if (CellType = ZENumber) then
      Result := AsDouble
    else
      b := true;
  end;
  if (b) then
    Raise EConvertError.Create('ZxCell: Cannot cast data to DateTime');;
end;

procedure TZCell.SetDataAsDateTime(const Value: TDateTime);
begin
  FCellType := ZEDateTime;
  //FData := ZEDateTimeToStr(Value, true);
  FData := FloatToStr(Value).Replace(',','.');
end;

function TZCell.GetDataAsDouble: double;
var err: integer; b: boolean; _dt: TDateTime;
begin
  Val(FData, Result, err); // need old-school to ignore regional settings
  if (err > 0) then begin
    b := true;
    //If datetime ...
    if (CellType = ZEDateTime) then
      b := not TryZEStrToDateTime(FData, _dt);

    if (b) then
      Raise EConvertError.Create('ZxCell: Cannot cast data to number')
    else
      Result := _dt;
  end;
end;

function TZCell.GetDataAsInteger: integer;
begin
  Result := StrToInt(Data);
end;

procedure TZCell.SetDataAsInteger(const Value: integer);
begin
  Data := Trim(IntToStr(Value));
  CellType := ZENumber;
// Val adds the prepending space, maybe some I2S implementation would adds too
// and Excel dislikes it. Better safe than sorry.
end;

procedure TZCell.SetDataAsDouble(const Value: double);
var s: String; ss: ShortString;
begin
  if Value = 0 then
    SetDataAsInteger(0) // work around Excel 2010 weird XLSX bug
  else begin
    Str(Value, ss); // need old-school to ignore regional settings
    s := string(ss); // make XE2 happy
    s := UpperCase(Trim(s));
    // UpperCase for exponent w.r.t OpenXML format
    // Trim for leading space w.r.t XML SS format
    FData := s;
    CellType := ZENumber;
    // Seem natural and logical thing to do w.r.t further export...
    // Seem out of "brain-dead no-automation overall aproach of a component...
    // Correct choice? dunno. I prefer making export better
  end
end;

procedure TZCell.SetDataAsString(const Value: string);
begin
  FData := Value;
  CellType := ZEString;
end;

procedure TZCell.Assign(Source: TPersistent);
var zSource: TZCell;
begin
  if (Source is TZCell) then begin
    zSource := Source as TZCell;
    FFormula           := zSource.Formula;
    FData              := zSource.Data;
    FHref              := zSource.Href;
    FComment           := zSource.Comment;
    FCommentAuthor     := zSource.CommentAuthor;
    FCellStyle         := zSource.CellStyle;
    FCellType          := zSource.CellType;
    FAlwaysShowComment := zSource.AlwaysShowComment;
    FShowComment       := zSource.ShowComment;
  end else
    inherited Assign(Source);
end;

procedure TZCell.Clear();
begin
  FFormula           := '';
  FData              := '';
  FHref              := '';
  FComment           := '';
  FCommentAuthor     := '';
  FHRefScreenTip     := '';
  FCellType          := ZEString;
  FCellStyle         := -1;
  FAlwaysShowComment := false;
  FShowComment       := false;
end;

////::::::::::::: TZMergeCells :::::::::::::::::////

constructor TZMergeCells.Create(ASheet: TZSheet);
begin
  FSheet := ASheet;
  FCount := 0;
end;

destructor TZMergeCells.Destroy();
begin
  Clear();
  inherited Destroy();
end;

procedure TZMergeCells.Clear();
begin
  SetLength(FMergeArea, 0);
  FCount := 0;
end;

function TZMergeCells.GetItem(num: integer):TRect;
begin
  if (num >= 0) and (num < Count) then
    Result := FMergeArea[num]
  else begin
    result.Left   := 0;
    result.Top    := 0;
    result.Right  := 0;
    result.Bottom := 0;
  end;
end;

procedure TZMergeCells.SetItem(Num: integer; const rect: TRect);
begin
  FMergeArea[Num] := rect;
end;

function TZMergeCells.AddRect(Rct:TRect): byte;
var i: integer;
function usl(rct1, rct2: TRect): boolean;
begin
  result := (((rct1.Left >= rct2.Left) and (rct1.Left <= rct2.Right)) or
      ((rct1.right >= rct2.Left) and (rct1.right <= rct2.Right))) and
     (((rct1.Top >= rct2.Top) and (rct1.Top <= rct2.Bottom)) or
      ((rct1.Bottom >= rct2.Top) and (rct1.Bottom <= rct2.Bottom)));
end;

begin
  if Rct.Left > Rct.Right then begin
    i := Rct.Left;
    Rct.Left := Rct.Right;
    Rct.Right := i;
  end;
  if Rct.Top > Rct.Bottom then begin
    i := Rct.Top;
    Rct.Top := Rct.Bottom;
    Rct.Bottom := i;
  end;
  if (Rct.Left < 0) or (Rct.Top < 0) then begin
    result := 1;
    exit;
  end;
  if (Rct.Right - Rct.Left = 0) and (Rct.Bottom - Rct.Top = 0) then begin
    result := 3;
    exit;
  end;
  for i := 0 to Count-1 do
    if usl(FMergeArea[i], Rct) or usl(rct, FMergeArea[i]) then  begin
      result := 2;
      exit;
    end;
  //если надо, увеличиваем кол-во строк/столбцов в хранилище
  if (FSheet <> nil) then begin
    if Rct.Right > FSheet.ColCount - 1 then FSheet.ColCount := Rct.Right {+ 1};
    if Rct.Bottom > FSheet.RowCount - 1 then FSheet.RowCount := Rct.Bottom {+ 1};  
  end;
  inc(FCount);
  setlength(FMergeArea, FCount);
  FMergeArea[FCount - 1] := Rct;
  result := 0;
end;

function TZMergeCells.InLeftTopCorner(ACol, ARow: integer): integer;
var i: integer;
begin
  result := -1;
  for i := 0 to FCount - 1 do
  if (ACol = FMergeArea[i].Left) and (ARow = FMergeArea[i].top) then begin
    result := i;
    break;
  end;
end;

function TZMergeCells.InMergeRange(ACol, ARow: integer): integer;
var i: integer;
begin
  result := -1;
  for i := 0 to FCount - 1 do
  if (ACol >= FMergeArea[i].Left) and (ACol <= FMergeArea[i].Right) and
     (ARow >= FMergeArea[i].Top) and (ARow <= FMergeArea[i].Bottom) then begin
    result := i;
    break;
  end;
end;

function TZMergeCells.IsCrossWithArea(AID, AC1, AR1, AC2, AR2: integer): Boolean;
begin
  result :=
    (((Items[AID].Left    >= AC1) and (Items[AID].Left   <= AC2)) or
      ((Items[AID].right  >= AC1) and (Items[AID].right  <= AC2))) and
     (((Items[AID].Top    >= AR1) and (Items[AID].Top    <= AR2)) or
      ((Items[AID].Bottom >= AR1) and (Items[AID].Bottom <= AR2)));
end;

function TZMergeCells.DeleteItem(num: integer): boolean;
var i: integer;
begin
  if (num>count-1)or(num<0) then begin
    result := false;
    exit;
  end;
  for i:= num to Count-2 do
    FMergeArea[i] := FMergeArea[i+1];
  dec(FCount);
  setlength(FMergeArea,FCount);
  result := true;
end;

function TZMergeCells.AddRectXY(x1,y1,x2,y2: integer): byte;
var Rct: TRect;
begin
  Rct.left   := x1;
  Rct.Top    := y1;
  Rct.Right  := x2;
  Rct.Bottom := y2;
  result := AddRect(Rct);
end;

////::::::::::::: TZRowColOptions :::::::::::::::::////

constructor TZRowColOptions.Create(ASheet: TZSheet);
begin
  inherited Create();
  FSheet   := ASheet;
  FHidden  := false;
  FAuto    := true;
  FStyleID := -1;
  FBreaked := false;
end;

procedure TZRowColOptions.Assign(Source: TPersistent);
begin
  if Source is TZRowColOptions then begin
    Hidden   := (Source as TZRowColOptions).Hidden;
    StyleID  := (Source as TZRowColOptions).StyleID;
    FSize    := (Source as TZRowColOptions).FSize;
    FAuto    := (Source as TZRowColOptions).FAuto;
    FBreaked := (Source as TZRowColOptions).Breaked;
  end else
    inherited Assign(Source);
end;

function  TZRowColOptions.GetAuto(): boolean;
begin
  result := FAuto;
end;

procedure TZRowColOptions.SetAuto(Value: boolean);
begin
  FAuto := Value;
end;

function  TZRowColOptions.GetSizePoint(): real;
begin
  result := FSize;
end;

procedure TZRowColOptions.SetSizePoint(Value: real);
begin
  if Value >= 0 then
    FSize := Value;
end;

function  TZRowColOptions.GetSizeMM(): real;
begin
  result := PointToMM(FSize);
end;

procedure TZRowColOptions.SetSizeMM(Value: real);
begin
  FSize := MMToPoint(Value);
end;

////::::::::::::: TZRowOptions :::::::::::::::::////

constructor TZRowOptions.Create(ASheet: TZSheet);
begin
  inherited Create(ASheet);
  FSize := 48;
end;

function TZRowOptions.GetSizePix(): integer;
var t: real;
begin
  t := 0.265;
  if FSheet <> nil then
    if FSheet.FStore <> nil then
      t := FSheet.FStore.HorPixelSize;
  result := PointToPixel(FSize, t);
end;

procedure TZRowOptions.SetSizePix(Value: integer);
var t: real;
begin
  if Value < 0 then  exit;
  t := 0.265;
  if FSheet <> nil then
    if FSheet.FStore <> nil then
      t := FSheet.FStore.HorPixelSize;
  FSize := PixelToPoint(Value, t);
end;

////::::::::::::: TZColOptions :::::::::::::::::////

constructor TZColOptions.Create(ASheet: TZSheet);
begin
  inherited Create(ASheet);
  FSize := 12.75;
end;

function TZColOptions.GetSizePix(): integer;
var t: real;
begin
  t := 0.265;
  if FSheet <> nil then
    if FSheet.FStore <> nil then
      t := FSheet.FStore.VertPixelSize;
  result := PointToPixel(FSize, t);
end;

procedure TZColOptions.SetSizePix(Value: integer);
var t: real;
begin
  if Value < 0 then  exit;
  t := 0.265;
  if FSheet <> nil then
    if FSheet.FStore <> nil then
      t := FSheet.FStore.VertPixelSize;
  FSize := PixelToPoint(Value, t);
end;

////::::::::::::: TZHeaderFooterMargins :::::::::::::::::////

constructor TZHeaderFooterMargins.Create();
begin
  FMarginTopBottom := 13;
  FMarginLeft := 0;
  FMarginRight := 0;
  FHeight := 7;
  FUseAutoFitHeight := true;
end;

procedure TZHeaderFooterMargins.Assign(Source: TPersistent);
var t: TZHeaderFooterMargins;
begin
  if (Source is TZHeaderFooterMargins) then begin
    t := Source as TZHeaderFooterMargins;
    FMarginTopBottom  := t.MarginTopBottom;
    FMarginLeft       := t.MarginLeft;
    FMarginRight      := t.MarginRight;
    FHeight           := t.Height;
    FUseAutoFitHeight := t.UseAutoFitHeight;
  end else
    inherited Assign(Source);
end;

function TZHeaderFooterMargins.IsEqual(Source: TPersistent): boolean;
var t: TZHeaderFooterMargins;
begin
  result := false;
  if (Assigned(Source)) then
    if (Source is TZHeaderFooterMargins) then begin
      t := Source as TZHeaderFooterMargins;
      result := (FMarginTopBottom = t.MarginTopBottom) and
                (FMarginLeft = t.MarginLeft) and
                (FMarginRight = t.MarginRight) and
                (FHeight = t.FHeight) and
                (FUseAutoFitHeight = t.UseAutoFitHeight);
    end;
end;

////::::::::::::: TZSheetOptions :::::::::::::::::////

constructor TZSheetOptions.Create();
begin
  inherited;
  FHeaderMargins := TZHeaderFooterMargins.Create();
  FFooterMargins := TZHeaderFooterMargins.Create();
  FActiveCol           := 0;
  FActiveRow           := 0;
  FMarginBottom        := 25;
  FMarginLeft          := 20;
  FMarginTop           := 25;
  FMarginRight         := 20;
  FPortraitOrientation := true;
  FCenterHorizontal    := false;
  FCenterVertical      := false;
  FStartPageNumber     := 1 ;
  HeaderMargin         := 13;
  FooterMargin         := 13;
  FPaperSize           := 9;
  FFitToHeight         :=-1;
  FFitToWidth          :=-1;
  FDifferentFirst      := False;
  FDifferentOddEven    := False;
  FHeader              := '';
  FFooter              := '';
  FEvenHeader          := '';
  FEvenFooter          := '';
  FFirstPageHeader     := '';
  FFirstPageFooter     := '';
  FHeaderBGColor       := clWindow;
  FFooterBGColor       := clWindow;
  FSplitVerticalMode   := ZSplitNone;
  FSplitHorizontalMode := ZSplitNone;
  FSplitVerticalValue  := 0;
  FSplitHorizontalValue:= 0;
  FPaperWidth          := 0;
  FPaperHeight         := 0;
  FScaleToPercent      := 100;
  FScaleToPages        := 1;
end;

destructor TZSheetOptions.Destroy();
begin
  FreeAndNil(FHeaderMargins);
  FreeAndNil(FFooterMargins);
  inherited;
end;

function TZSheetOptions.GetHeaderMargin(): word;
begin
  result := FHeaderMargins.Height;
end;

procedure TZSheetOptions.SetHeaderMargin(Value: word);
begin
  FHeaderMargins.Height := Value;
end;

function TZSheetOptions.GetFooterMargin(): word;
begin
  result := FFooterMargins.Height;
end;

procedure TZSheetOptions.SetFooterMargin(Value: word);
begin
  FFooterMargins.Height := Value;
end;

procedure TZSheetOptions.Assign(Source: TPersistent);
var t: TZSheetOptions;
begin
  if Source is TZSheetOptions then begin
    t := Source as TZSheetOptions;
    ActiveCol            := t.ActiveCol;
    ActiveRow            := t.ActiveRow;
    MarginBottom         := t.MarginBottom;
    MarginLeft           := t.MarginLeft;
    MarginTop            := t.MarginTop;
    MarginRight          := t.MarginRight;
    PortraitOrientation  := t.PortraitOrientation;
    CenterHorizontal     := t.CenterHorizontal;
    CenterVertical       := t.CenterVertical;
    StartPageNumber      := t.StartPageNumber;
    PaperSize            := t.PaperSize;
    FitToHeight          := t.FitToHeight;
    FitToWidth           := t.FitToWidth;
    SplitVerticalMode    := t.SplitVerticalMode;
    SplitHorizontalMode  := t.SplitHorizontalMode;
    SplitVerticalValue   := t.SplitVerticalValue;
    SplitHorizontalValue := t.SplitHorizontalValue;
    Footer               := t.Footer;
    Header               := t.Header;
    EvenHeader           := t.EvenHeader;
    EvenFooter           := t.EvenFooter;
    FirstPageHeader      := t.FirstPageHeader;
    FirstPageFooter      := t.FirstPageFooter;
    HeaderBGColor        := t.HeaderBGColor;
    FooterBGColor        := t.FooterBGColor;
    IsDifferentFirst     := t.IsDifferentFirst;
    IsDifferentOddEven   := t.IsDifferentOddEven;
    ScaleToPercent       := t.ScaleToPercent;
    ScaleToPages         := t.ScaleToPages;
    HeaderMargins.Assign(t.HeaderMargins);
    FooterMargins.Assign(t.FooterMargins);
  end else
    inherited Assign(Source);
end;

////::::::::::::: TZSheet :::::::::::::::::////

constructor TZSheet.Create(AStore: TZEXMLSS);
var i, j: integer;
begin
  FStore := AStore;
  FRowCount := 0;
  FColCount := 0;
  FDefaultRowHeight := 12.75;//16;
  FDefaultColWidth := 48;//60;
  FMergeCells := TZMergeCells.Create(self);
  setlength(FCells, FColCount);
  FTabColor := ClWindow;
  FProtect := false;
  FRightToLeft := false;
  FSelected := false;
  SetLength(FRows, FRowCount);
  SetLength(FColumns, FColCount);
  for i := 0 to FColCount - 1 do begin
    setlength(FCells[i], FRowCount);
    for j := 0 to FRowCount - 1 do begin
      FCells[i,j] := TZCell.Create();
    end;
    FColumns[i] := TZColOptions.create(self);
    FColumns[i].Width := DefaultColWidth;
  end;
  for i := 0 to FRowCount - 1 do begin
    FRows[i] := TZRowOptions.create(self);
    FRows[i].Height := DefaultRowHeight;
  end;
  FSheetOptions := TZSheetOptions.Create();
  if FStore <> nil then
    if FStore.DefaultSheetOptions <> nil then
      FSheetOptions.Assign(FStore.DefaultSheetOptions);

  FPrintRows := TZSheetPrintTitles.Create(Self, false);
  FPrintCols := TZSheetPrintTitles.Create(Self, true);
  FConditionalFormatting := TZConditionalFormatting.Create();
  FCharts := TZEChartStore.Create();
  FDrawing := TZEDrawing.Create();
end;

destructor TZSheet.Destroy();
begin
  try
    FreeAndNil(FMergeCells);
    FreeAndNil(FSheetOptions);
    FPrintRows.Free;
    FPrintCols.Free;
    Clear();
    FCells := nil;
    FRows := nil;
    FColumns := nil;
    FreeAndNil(FConditionalFormatting);
    FreeAndNil(FCharts);
    FreeAndNil(FDrawing);
  finally
    inherited Destroy();
  end;
end;

procedure TZSheet.Assign(Source: TPersistent);
var zSource: TZSheet; i, j: integer;

begin
  if (Source is TZSheet) then begin
    ZSource     := Source as TZSheet;
    ////////////////////////////////
    RowCount    := zSource.RowCount;
    ColCount    := zSource.ColCount;
    TabColor    := zSource.TabColor;
    FitToPage   := zSource.FitToPage;
    Title       := zSource.Title;
    Protect     := zSource.Protect;
    RightToLeft := zSource.RightToLeft;
    DefaultRowHeight := zSource.DefaultRowHeight;
    DefaultColWidth  := zSource.DefaultColWidth;
    FViewMode  := zSource.FViewMode;
    FRowBreaks := zSource.FRowBreaks;
    FColBreaks := zSource.FColBreaks;
    FViewMode  := zSource.FViewMode;

    for i := 0 to RowCount - 1 do
      Rows[i] := ZSource.Rows[i];

    for i := 0 to ColCount - 1 do begin
      Columns[i] := ZSource.Columns[i];
      for j := 0 to RowCount - 1 do
        Cell[i, j] := zSource.Cell[i, j];
    end;

    FSelected := zSource.Selected;
    SheetOptions.Assign(zSource.SheetOptions);

    MergeCells.Clear();
    for i := 0 to ZSource.MergeCells.Count - 1 do
      MergeCells.AddRect(ZSource.MergeCells.GetItem(i));

    ConditionalFormatting.Assign(zSource.ConditionalFormatting);
    FDrawing := TZEDrawing.Create();
    FDrawing.Assign(zSource.FDrawing);
    RowsToRepeat.Assign(zSource.RowsToRepeat);
    ColsToRepeat.Assign(zSource.ColsToRepeat);
  end else
    inherited Assign(Source);
end;

procedure TZSheet.SetConditionalFormatting(Value: TZConditionalFormatting);
begin
  if (Assigned(Value)) then
    FConditionalFormatting.Assign(Value);
end; //SetConditionalFormatting

procedure TZSheet.SetCharts(const Value: TZEChartStore);
begin
  if (Assigned(Value)) then
    FCharts.Assign(Value);
end;

function TZSheet.GetSheetIndex: integer;
var i: Integer;
begin
  Result := 0;
  for I := 0 to WorkBook.FSheets.Count-1 do
    if WorkBook.FSheets[i] = Self then
        Exit(i);
end;

function TZSheet.GetSheetOptions(): TZSheetOptions;
begin
  result := FSheetOptions;
end;

procedure TZSheet.InsertRows(ARow, ACount: Integer);
var r, c: Integer;
begin
  // resize
  SetRowCount(FRowCount + ACount);

  // append and reloc cells
  for r := Length(FRows) - 1 downto ARow do begin
    // reloc rows
    if (r - ACount) < ARow then begin
      FRows[r] := TZRowOptions.Create(Self);
      FRows[r].Height := DefaultRowHeight;
    end else begin
      FRows[r] := FRows[r - ACount];
    end;
    // reloc cells
    for c := 0 to ColCount-1 do begin
      if (r - ACount) < ARow then begin
        FCells[c][r].Clear();
      end else begin
        FCells[c][r].Assign(FCells[c][r-ACount]);
      end;
    end;
  end;

  // reloc merged areas
  for r := 0 to MergeCells.Count-1 do begin
    if MergeCells[r].Top >= ARow then begin
      MergeCells.Items[r] := TRect.Create(
        MergeCells.Items[r].Left,
        MergeCells.Items[r].Top + ACount,
        MergeCells.Items[r].Right,
        MergeCells.Items[r].Bottom + ACount);
    end;
  end;
end;

procedure TZSheet.CopyRows(ARowDst, ARowSrc, ACount: Integer);
var r, c, delta: integer;
begin
  // copy row's and cell's info
  for r := 0 to ACount-1 do begin
    FRows[ARowDst+r].Assign( FRows[ARowSrc+r] );
    for c := 0 to FColCount-1 do begin
      FCells[c][ARowDst+r].Assign( FCells[c][ARowSrc+r] );
    end;
  end;

  delta := ARowDst - ARowSrc;

  // reloc merged areas
  for r := 0 to MergeCells.Count-1 do begin
    if (MergeCells[r].Top >= ARowSrc) and (MergeCells[r].Bottom < ARowSrc + ACount) then begin
      MergeCells.AddRect(TRect.Create(
        MergeCells.Items[r].Left,
        MergeCells.Items[r].Top + delta,
        MergeCells.Items[r].Right,
        MergeCells.Items[r].Bottom + delta));
    end;
  end;
end;

procedure TZSheet.SetSheetOptions(Value: TZSheetOptions);
begin
  if Value <> nil then
   FSheetOptions.Assign(Value);
end;

procedure TZSheet.SetCorrectTitle(const Value: string);
var
  i: integer;
  suffix, newTitle: string;
  sheetNames: TDictionary<string, string>;
  regEx: TRegEx;
begin
  newTitle := trim(Value);
  if newTitle.Length > 31 then
    newTitle := newTitle.Substring(0, 31);
  regEx := TRegEx.Create('[\\/\*\[\]\?:]');
  newTitle := regEx.Replace(newTitle, ' ');
  if newTitle.Trim.IsEmpty then
    newTitle := 'Лист1';
  if assigned(workbook) then
  begin
    sheetNames := TDictionary<string, string>.Create;
    try
      for i := 0 to self.workbook.Sheets.Count - 1 do
        if self <> workbook.Sheets[i] then
          sheetNames.AddOrSetValue(self.workbook.Sheets[i].Title.ToLower, '');
      i := 1;
      suffix := '';
      repeat
        if i > 1 then
          suffix := ' (' + IntToStr(i) + ')';
        if newTitle.Length + suffix.Length > 31 then
          newTitle := newTitle.Substring(0, 31 - suffix.Length);
        Inc(i);
      until not sheetNames.ContainsKey(newTitle.ToLower + suffix);
      newTitle := newTitle + suffix;
    finally
      sheetNames.Free;
    end;
  end;
  self.FTitle := newTitle;
end;

procedure TZSheet.SetColumn(num: integer; const Value:TZColOptions);
begin
  if (num >= 0) and (num < FColCount) then
    FColumns[num].Assign(Value);
end;

function  TZSheet.GetColumn(num: integer): TZColOptions;
begin
  if (num >= 0) and (num < FColCount) then
    result := FColumns[num]
  else
    result := nil;
end;

{
procedure TZSheet.SetRange(AC1,AR1,AC2,AR2: integer; const Value: TZRange);
begin

end;

procedure TZSheet.SetRangeRef(AFrom, ATo: string; const Value: TZRange);
begin

end;
}

function TZSheet.GetRange(AC1,AR1,AC2,AR2: integer): TZRange;
begin
  Result := TZRange.Create(Self, AC1,AR1,AC2,AR2);
end;

function TZSheet.GetRangeRef(AFromCol: string; AFromRow: Integer; AToCol: string; AToRow: integer): TZRange;
var AC1,AR1,AC2,AR2: Integer;
begin
  AC1 := ZEGetColByA1(AFromCol);
  AR1 := AFromRow;
  AC2 := ZEGetColByA1(AToCol);
  AR2 := AToRow;
  Result := TZRange.Create(Self, AC1,AR1,AC2,AR2);
end;

procedure TZSheet.SetRow(num: integer; const Value: TZRowOptions);
begin
  if (num >= 0) and (num < FRowCount) then
    FRows[num].Assign(Value);
end;

function  TZSheet.GetRow(num: integer): TZRowOptions;
begin
  if (num >= 0) and (num < FRowCount) then
    result := FRows[num]
  else
    result := nil;
end;

procedure TZSheet.SetColWidth(num: integer; const Value: real);
begin
  if (num < ColCount) and (num >= 0) and (Value >= 0) then
  if FColumns[num] <> nil then
    FColumns[num].Width := Value;
end;

function TZSheet.GetColWidth(num: integer): real;
begin
  result := 0;
  if (num < ColCount) and (num >= 0) then
    if FColumns[num] <> nil then
      result := FColumns[num].Width
end;

procedure TZSheet.SetRowHeight(num: integer; const Value: real);
begin
  if (num < RowCount) and (num >= 0) and (Value >= 0) then
    if FRows[num] <> nil then
      FRows[num].Height := Value;
end;

function TZSheet.GetRowHeight(num: integer): real;
begin
  result := 0;
  if (num < RowCount) and (num >= 0) then
    if FRows[num] <> nil then
      result := FRows[num].Height
end;

procedure TZSheet.SetDefaultColWidth(const Value: real);
begin
  if Value >= 0 then
    FDefaultColWidth := round(Value*100)/100;
end;

procedure TZSheet.SetDefaultRowHeight(const Value: real);
begin
  if Value >= 0 then
    FDefaultRowHeight := round(Value*100)/100;
end;

procedure TZSheet.SetPrintCols(const Value: TZSheetPrintTitles);
begin
  FPrintCols.Assign( Value );
end;

procedure TZSheet.SetPrintRows(const Value: TZSheetPrintTitles);
begin
  FPrintRows.Assign( Value );
end;

procedure TZSheet.Clear();
var i, j: integer;
begin
  for i:= 0 to FColCount - 1 do begin
    for j:= 0 to FRowCount - 1 do
      FreeAndNil(FCells[i][j]);
    SetLength(FCells[i], 0);
    FColumns[i].Free;
    FCells[i] := nil;
  end;
  for i := 0 to FRowCount - 1 do
    FRows[i].Free;
  SetLength(FCells, 0);
  FRowCount := 0;
  FColCount := 0;
  SetLength(FRows, 0);
  SetLength(FColumns, 0);
end;

procedure TZSheet.SetCell(ACol, ARow: integer; const Value: TZCell);
begin
  //добавить проверку на корректность текста
  if (ACol >= 0) and (ACol < FColCount) and
     (ARow >= 0) and (ARow < FRowCount) then
  FCells[ACol, ARow].Assign(Value);
end;

procedure TZSheet.SetCellRef(ACol: string; ARow: integer; const Value: TZCell);
begin
  SetCell(ZEGetColByA1(ACol), ARow, Value);
end;

function TZSheet.GetCell(ACol, ARow: integer):TZCell;
begin
  result := nil;
  if (ACol >= 0) and (ACol < FColCount) and
     (ARow >= 0) and (ARow < FRowCount) then
    result := FCells[ACol, ARow];
end;

function TZSheet.GetCellRef(ACol: string; ARow: integer): TZCell;
begin
    Result := GetCell(ZEGetColByA1(ACol), ARow);
end;

procedure TZSheet.SetColCount(const Value: integer);
var i, j: integer;
begin
  if Value < 0 then
    exit;
  if FColCount > Value then begin // todo Repeatable columns may be affected
    for i := Value to FColCount - 1 do begin
      for j := 0 to FRowCount - 1 do
        FreeAndNil(FCells[i][j]);
      setlength(FCells[i], 0);
      FreeAndNil(FColumns[i]);
    end;
    setlength(FCells, Value);
    SetLength(FColumns, Value);
  end else begin
    setlength(FCells, Value);
    SetLength(FColumns, Value);
    for i := FColCount to Value - 1 do
    begin
      setlength(FCells[i], FRowCount);
      FColumns[i] := TZColOptions.Create(self);
      FColumns[i].Width := DefaultColWidth;
      for j := 0 to FRowCount - 1 do
        FCells[i][j] := TZCell.Create();
    end;
  end;
  FColCount := Value;
end;

function TZSheet.GetColCount: integer;
begin
  result := FColCount;
end;

procedure TZSheet.SetRowCount(const Value: integer);
var i, j: integer;
begin
  if Value < 0 then
    exit;
  if FRowCount > Value then begin // todo Repeatable rows may be affected
    for i := 0 to FColCount - 1 do begin
      for j := Value  to FRowCount - 1 do
        FreeAndNil(FCells[i][j]);
      setlength(FCells[i], Value);
    end;
    for i := Value to FRowCount - 1 do
      FreeAndNil(FRows[i]);
    SetLength(FRows, Value);
  end else begin
    for i := 0 to FColCount - 1 do begin
      setlength(FCells[i], Value);
      for j := FRowCount to Value - 1 do
        FCells[i][j] := TZCell.Create();
    end;
    SetLength(FRows, Value);
    for i := FRowCount to Value - 1 do begin
      FRows[i] := TZRowOptions.Create(self);
      FRows[i].Height := DefaultRowHeight;
    end;
  end;
  FRowCount := Value;
end;

function TZSheet.GetRowCount: integer;
begin
  result := FRowCount;
end;

////::::::::::::: TZSheets :::::::::::::::::////
constructor TZSheets.Create(AStore: TZEXMLSS);
begin
  FStore := AStore;
  FCount := 0;
  SetLength(FSheets, 0);
end;

destructor TZSheets.Destroy();
var i: integer;
begin
  for i := 0 to FCount - 1 do
    FreeAndNil(FSheets[i]);
  SetLength(FSheets, 0);
  FSheets := nil;
  FStore := nil;
  inherited Destroy();
end;

procedure TZSheets.Assign(Source: TPersistent);
var t: TZSheets; i: integer;
begin
  if (Source is TZSheets) then begin
    t := Source as TZSheets;
    Count := t.Count;
    for i := 0 to Count - 1 do
      Sheet[i].Assign(t.Sheet[i]);
  end else
    inherited Assign(Source);
end; //Assign

procedure TZSheets.SetSheetCount(const Value: integer);
var i: integer;
begin
  if Count < Value then begin
    SetLength(FSheets, Value);
    for i := Count to Value - 1 do
      FSheets[i] := TZSheet.Create(FStore);
  end else begin
    for i:= Value to Count - 1  do
      FreeAndNil(FSheets[i]);
    SetLength(FSheets, Value);
  end;
  FCount := Value;
end;

procedure TZSheets.SetSheet(num: integer; Const Value: TZSheet);
begin
  if (num < Count) and (num >= 0) then
    FSheets[num].Assign(Value);
end;

function TZSheets.GetSheet(num: Integer): TZSheet;
begin
  if (num < Count) and (num >= 0) then
    result := FSheets[num]
  else
    result := nil;
end;

////::::::::::::: TZEXMLDocumentProperties :::::::::::::::::////

constructor TZEXMLDocumentProperties.Create();
begin
  FAuthor       := 'none';
  FLastAuthor   := 'none';
  FCompany      := 'none';
  FVersion      := '11.9999';
  FCreated      := Now();
  FWindowHeight := 20000;
  FWindowWidth  := 20000;
  FWindowTopX   := 150;
  FWindowTopY   := 150;
  FModeR1C1     := false;
end;

procedure TZEXMLDocumentProperties.SetAuthor(const Value: string);
begin
  //поставить проверку на валидность ввода
  FAuthor := Value;
end;

procedure TZEXMLDocumentProperties.SetLastAuthor(const Value: string);
begin
  //поставить проверку на валидность ввода
  FLastAuthor := Value;
end;

procedure TZEXMLDocumentProperties.SetCompany(const Value: string);
begin
  //поставить проверку на валидность ввода
  FCompany := Value;
end;

procedure TZEXMLDocumentProperties.SetVersion(const Value: string);
begin
  //поставить проверку на валидность ввода
  FVersion := Value;
end;

procedure TZEXMLDocumentProperties.Assign(Source: TPersistent);
var props: TZEXMLDocumentProperties;
begin
  if Source is TZEXMLDocumentProperties then begin
    props := Source as TZEXMLDocumentProperties;
    Author       := props.Author;
    LastAuthor   := props.LastAuthor;
    Created      := props.Created;
    Company      := props.Company;
    Version      := props.Version;
    WindowHeight := props.WindowHeight;
    WindowWidth  := props.WindowWidth;
    WindowTopX   := props.WindowTopX;
    WindowTopY   := props.WindowTopY;
    ModeR1C1     := props.ModeR1C1;
  end else
    inherited Assign(Source);
end;

////::::::::::::: TZEXMLSS :::::::::::::::::////

constructor TZEXMLSS.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FDocumentProperties  := TZEXMLDocumentProperties.Create;
  FStyles              := TZStyles.Create();
  FSheets              := TZSheets.Create(self);
  FHorPixelSize        := 0.265;
  FVertPixelSize       := 0.265;
  FDefaultSheetOptions := TZSheetOptions.Create();
end;

destructor TZEXMLSS.Destroy();
begin
  FreeAndNil(FDefaultSheetOptions);
  FreeAndNil(FDocumentProperties);
  FreeAndNil(FStyles);
  FreeAndNil(FSheets);
  inherited Destroy();
end;

procedure TZEXMLSS.Assign(Source: TPersistent);
var t: TZEXMLSS;
begin
  if (Source is TZEXMLSS) then begin
    t := Source as TZEXMLSS;
    Styles.Assign(t.Styles);
    Sheets.Assign(t.Sheets);
  end else if (Source is TZStyles) then
    Styles.Assign(Source as TZStyles)
  else if (Source is TZSheets) then
    Sheets.Assign(Source as TZSheets)
  else
    inherited Assign(Source);
end; //Assign

procedure TZEXMLSS.SetHorPixelSize(Value: real);
begin
  if Value > 0 then
    FHorPixelSize := Value;
end;

procedure TZEXMLSS.SetVertPixelSize(Value: real);
begin
  if Value > 0 then
    FVertPixelSize := Value;
end;

procedure TZEXMLSS.GetPixelSize(hdc: HWND);
begin
  HorPixelSize  := GetDeviceCaps(hdc, HORZSIZE) / GetDeviceCaps(hdc, HORZRES); //горизонтальный размер пикселя в миллиметрах
  VertPixelSize := GetDeviceCaps(hdc, VERTSIZE) / GetDeviceCaps(hdc, VERTRES); //вертикальный размер пикселя в миллиметрах
end;

function TZEXMLSS.GetDefaultSheetOptions(): TZSheetOptions;
begin
  result := FDefaultSheetOptions;
end;

procedure TZEXMLSS.SetDefaultSheetOptions(Value: TZSheetOptions);
begin
  if Value <> nil then
   FDefaultSheetOptions.Assign(Value);
end;

function TZEXMLSS.DrawingCount(): Integer;
var i, ii, n: Integer;
    tmp: TZEDrawing;
begin
  Result := 0;
  n := 0;
  for i := 0 to Sheets.Count - 1 do begin
    tmp := Sheets[i].Drawing;
    if (not tmp.IsEmpty) then begin
      Inc(Result);
      tmp.Id := Result;
      // set id for pictures
      for ii := 0 to tmp.PictureStore.Count - 1 do begin
        Inc(n);
        tmp.PictureStore[ii].Id := n;
      end;
    end else
      tmp.Id := 0;
  end;
end;

function TZEXMLSS.GetDrawing(num: Integer): TZEDrawing;
var i, n: Integer;
begin
  Result := nil;
  n := 0;
  for i := 0 to Sheets.Count - 1 do begin
    if not Sheets[i].Drawing.IsEmpty then begin
      if n = num then begin
        Result := Sheets[i].Drawing;
        Exit;
      end;
      Inc(n);
    end;
  end;
end;

function TZEXMLSS.GetDrawingSheetNum(Value: TZEDrawing): Integer;
var i: Integer;
begin
  Result := 0;
  for i := 0 to Sheets.Count - 1 do begin
    if Value = Sheets[i].Drawing then begin
      Result := i;
      Exit;
    end;
  end;
end;

{ TZSheetPrintTitles }

procedure TZSheetPrintTitles.Assign(Source: TPersistent);
var f, t: word;  a: boolean;
begin
  if Source is TZSheetPrintTitles then begin
      F := TZSheetPrintTitles(Source).From;
      T := TZSheetPrintTitles(Source).Till;
      A := TZSheetPrintTitles(Source).Active;

      if A then
        RequireValid(F, T);

      FFrom   := F;
      FTill   := T;
      FActive := A;
    end
  else inherited;
end;

constructor TZSheetPrintTitles.Create(const owner: TZSheet; const ForColumns: boolean);
begin
  if nil = owner then
    raise Exception.Create(Self.ClassName+' requires an existing worksheet for owners.');
  Self.FOwner := owner;
  Self.FColumns := ForColumns;
end;

procedure TZSheetPrintTitles.SetActive(const Value: boolean);
begin
  if Value then
    RequireValid(FFrom, FTill);
  FActive := Value
end;

procedure TZSheetPrintTitles.SetFrom(const Value: word);
begin
  if Active then
    RequireValid(Value, FTill);
  FFrom := Value;
end;

procedure TZSheetPrintTitles.SetTill(const Value: word);
begin
  if Active then
    RequireValid(FFrom, Value);
  FTill := Value;
end;

function TZSheetPrintTitles.ToString: string;
var c: char;
begin
  If Active then begin
     if FColumns then c := 'C' else c := 'R';
     Result := c + IntToStr(From + 1) + ':' + c + IntToStr(Till + 1);
  end else
     Result := '';
end;

procedure TZSheetPrintTitles.RequireValid(const AFrom, ATill: word);
begin
  if not Valid(AFrom, ATill) then
    raise Exception.Create('Invalid printable titles for the worksheet.');
end;

function TZSheetPrintTitles.Valid(const AFrom, ATill: word): boolean;
var UpperLimit: word;
begin
  Result := False;
  if AFrom > ATill then
    exit;

  if FColumns then
    UpperLimit := FOwner.ColCount
  else
    UpperLimit := FOwner.RowCount;

  if ATill >= UpperLimit then
    exit;

  Result := True;
end;

////::::::::::::: TZConditionalStyleItem :::::::::::::::::////

constructor TZConditionalStyleItem.Create();
begin
  Clear();
end;

procedure TZConditionalStyleItem.Clear();
begin
  FCondition           := ZCFCellContentOperator;
  FConditionOperator   := ZCFOpEqual;
  FApplyStyleID        := -1;
  FValue1              := '';
  FValue2              := '';
  FBaseCellPageIndex   := -1;
  FBaseCellRowIndex    := 0;
  FBaseCellColumnIndex := 0;
end; //Clear

procedure TZConditionalStyleItem.Assign(Source: TPersistent);
var t: TZConditionalStyleItem;
begin
  if (Source is TZConditionalStyleItem) then begin
    t := (Source as TZConditionalStyleItem);
    Condition           := t.Condition;
    ConditionOperator   := t.ConditionOperator;
    Value1              := t.Value1;
    Value2              := t.Value2;
    ApplyStyleID        := t.ApplyStyleID;
    BaseCellColumnIndex := t.BaseCellColumnIndex;
    BaseCellPageIndex   := t.BaseCellPageIndex;
    BaseCellRowIndex    := t.BaseCellRowIndex;
  end else
    inherited Assign(Source);
end; //Assign

function TZConditionalStyleItem.IsEqual(Source: TPersistent): boolean;
var t: TZConditionalStyleItem;
begin
  result := false;
  if (Source is TZConditionalStyleItem) then begin
    t := (Source as TZConditionalStyleItem);
    if (Condition  <> t.Condition) then
      exit;
    if (ConditionOperator <> t.ConditionOperator) then
      exit;
    if (ApplyStyleID <> t.ApplyStyleID) then
      exit;
    if (BaseCellColumnIndex <> t.BaseCellColumnIndex) then
      exit;
    if (BaseCellPageIndex <> t.BaseCellPageIndex) then
      exit;
    if (BaseCellRowIndex <> t.BaseCellRowIndex) then
      exit;
    if (Value1 <> t.Value1) then
      exit;
    if (Value2 <> t.Value2) then
      exit;
    result := true;
  end;
end; //IsEqual

////::::::::::::: TZConditionalStyle :::::::::::::::::////

constructor TZConditionalStyle.Create();
var i: integer;
begin
  FCount := 0;
  FMaxCount := 3;
  SetLength(FConditions, FMaxCount);
  for i := 0 to FMaxCount - 1 do
    FConditions[i] := TZConditionalStyleItem.Create();
  FAreas := TZConditionalAreas.Create();
end; //Create

destructor TZConditionalStyle.Destroy();
var i: integer;
begin
  for i := 0 to FMaxCount - 1 do
    if (Assigned(FConditions[i])) then
      FreeAndNil(FConditions[i]);
  SetLength(FConditions, 0);
  FConditions := nil;
  FreeAndNil(FAreas);
  inherited;
end; //Destroy

procedure TZConditionalStyle.Assign(Source: TPersistent);
var t: TZConditionalStyle; i: integer;
begin
  if (Source is TZConditionalStyle) then begin
    t := Source as TZConditionalStyle;
    Count := t.Count;
    FAreas.Assign(t.Areas);
    for i := 0 to Count - 1 do
      FConditions[i].Assign(t.Items[i]);
  end else
    inherited Assign(Source);
end; //Assign

function TZConditionalStyle.IsEqual(Source: TPersistent): boolean;
var t: TZConditionalStyle; i: integer;
begin
  result := false;
  if (Source is TZConditionalStyle) then begin
    t := Source as TZConditionalStyle;
    if (Count <> t.Count) then
      exit;
    for i := 0 to Count - 1 do
      if (not FConditions[i].IsEqual(t.Items[i])) then
        exit;
    result := true;
  end;
end; //IsEqual

function TZConditionalStyle.GetItem(num: integer): TZConditionalStyleItem;
begin
  result := nil;
  if (num >= 0) and (num < Count) then
    result := FConditions[num];
end; //GetItem

procedure TZConditionalStyle.SetItem(num: integer; Value: TZConditionalStyleItem);
begin
  if (num >= 0) and (num < Count) then
    FConditions[num].Assign(Value);
end; //SetItem

procedure TZConditionalStyle.SetCount(value: integer);
var i: integer;
begin
  //TODO: нужно ли ограничение на максимальное кол-во?
  if (value >= 0) then begin
    if (value < FCount) then begin
      for i := value to FCount - 1 do
        FConditions[i].Clear();
    end else if (value > FMaxCount) then begin
      SetLength(FConditions, value);
      for i := FMaxCount to value - 1 do
        FConditions[i] := TZConditionalStyleItem.Create();
      FMaxCount := value;
    end;
    FCount := value;
  end;
end; //SetCount

procedure TZConditionalStyle.SetAreas(Value: TZConditionalAreas);
begin
  if (Assigned(Value)) then
    FAreas.Assign(Value);
end;

function TZConditionalStyle.Add(): TZConditionalStyleItem;
begin
  Count := Count + 1;
  result := FConditions[Count - 1];
end; //Add

function TZConditionalStyle.Add(StyleItem: TZConditionalStyleItem): TZConditionalStyleItem;
begin
  result := Add();
  if (Assigned(StyleItem)) then
    result.Assign(StyleItem);
end; //Add

procedure TZConditionalStyle.Delete(num: integer);
var i: integer; t: TZConditionalStyleItem;
begin
  if (num >= 0) and (num < Count) then begin
    t := FConditions[num];
    for i := num to Count - 2 do
      FConditions[i] := FConditions[i + 1];
    if (Count > 0) then
      FConditions[Count - 1] := t;
    Count := Count - 1;
  end;
end; //Delete

procedure TZConditionalStyle.Insert(num: integer);
begin
  Insert(num, nil);
end; //Insert

procedure TZConditionalStyle.Insert(num: integer; StyleItem: TZConditionalStyleItem);
var i: integer; t: TZConditionalStyleItem;
begin
  if (num >= 0) and (num < Count) then begin
    Add();
    t := FConditions[Count - 1];
    for i := Count - 1 downto num + 1 do
      FConditions[i] := FConditions[i - 1];
    FConditions[num] := t;
    if (Assigned(StyleItem)) then
      FConditions[num].Assign(StyleItem);
  end;
end; //Insert

////::::::::::::: TZConditionalAreaItem :::::::::::::::::////

constructor TZConditionalAreaItem.Create();
begin
  Create(0, 0, 1, 1);
end;

constructor TZConditionalAreaItem.Create(ColumnNum, RowNum, AreaWidth, AreaHeight: integer);
begin
  Row    := RowNum;
  Column := ColumnNum;
  Width  := AreaWidth;
  Height := AreaHeight;
end;

procedure TZConditionalAreaItem.SetRow(Value: integer);
begin
  if (Value >= 0) then
    FRow := Value;
end; //SetRow

procedure TZConditionalAreaItem.SetColumn(Value: integer);
begin
  if (Value >= 0) then
    FColumn := Value;
end; //SetColumn

procedure TZConditionalAreaItem.SetWidth(Value: integer);
begin
  if (Value >= 0) then
    FWidth := Value;
end; //SetWidth

procedure TZConditionalAreaItem.SetHeight(Value: integer);
begin
  if (Value >= 0) then
    FHeight := Value;
end; //SetHeight

procedure TZConditionalAreaItem.Assign(Source: TPersistent);
var t: TZConditionalAreaItem;
begin
  if (Source is TZConditionalAreaItem) then begin
    t := Source as TZConditionalAreaItem;
    Row    := t.Row;
    Column := t.Column;
    Height := t.Height;
    Width  := t.Width;
  end else
    inherited Assign(Source);
end; //Assign

function TZConditionalAreaItem.IsEqual(Source: TPersistent): boolean;
var t: TZConditionalAreaItem;
begin
  result := false;
  if (Source is TZConditionalAreaItem) then begin
    t := Source as TZConditionalAreaItem;
    if (FRow <> t.Row) then
      exit;
    if (FColumn <> t.Column) then
      exit;
    if (FWidth <> t.Width) then
      exit;
    if (FHeight <> t.Height) then
      exit;
    result := true;
  end;
end; //IsEqual

////::::::::::::: TZConditionalAreas :::::::::::::::::////

constructor TZConditionalAreas.Create();
begin
  FCount := 1;
  SetLength(FItems, FCount);
  FItems[0] := TZConditionalAreaItem.Create();
end; //Create

destructor TZConditionalAreas.Destroy();
var i: integer;
begin
  for i := 0 to FCount - 1 do
    if (Assigned(FItems[i])) then
      FreeAndNil(FItems[i]);
  Setlength(FItems, 0);
  FItems := nil;
  inherited;
end; //Destroy

procedure TZConditionalAreas.SetCount(Value: integer);
var i: integer;
begin
  if ((Value >= 0) and (Value <> Count)) then begin
    if (Value < Count) then begin
      for i := Value to Count - 1 do
        if (Assigned(FItems[i])) then
          FreeAndNil(FItems[i]);
      Setlength(FItems, Value);
    end else if (Value > Count) then begin
      Setlength(FItems, Value);
      for i := Count to Value - 1 do
        FItems[i] := TZConditionalAreaItem.Create();
    end;
    FCount := Value;
  end;
end; //SetCount

function TZConditionalAreas.GetItem(num: integer): TZConditionalAreaItem;
begin
  result := nil;
  if ((num >= 0) and (num < FCount)) then
    result := FItems[num];
end; //GetItem

procedure TZConditionalAreas.SetItem(num: integer; Value: TZConditionalAreaItem);
begin
  if ((num >= 0) and (num < Count)) then
    if (Assigned(Value)) then
      FItems[num].Assign(Value);
end; //SetItem

function TZConditionalAreas.Add(): TZConditionalAreaItem;
begin
  SetCount(FCount + 1);
  Result  := FItems[FCount - 1];
end; //Add

function TZConditionalAreas.Add(ColumnNum, RowNum, AreaWidth, AreaHeight: integer): TZConditionalAreaItem;
begin
  Result        := Add();
  Result.Row    := RowNum;
  Result.Column := ColumnNum;
  Result.Width  := AreaWidth;
  Result.Height := AreaHeight;
end; //Add

procedure TZConditionalAreas.Assign(Source: TPersistent);
var t: TZConditionalAreas; i: integer;
begin
  if (Source is TZConditionalAreas) then begin
    t := Source as TZConditionalAreas;
    Count := t.Count;
    for i := 0 to Count - 1 do
        FItems[i].Assign(t.Items[i]);
  end else
    inherited Assign(Source);
end; //Assign

procedure TZConditionalAreas.Delete(num: integer);
var t: TZConditionalAreaItem; i: integer;
begin
  if ((num >= 0) and (num < Count)) then begin
    t := FItems[num];
    for i := num to Count - 2 do
      FItems[i] := FItems[i + 1];
    FItems[Count - 1] := t;
    Count := Count -1;
  end;
end; //Delete

//Определяет, находится ли ячейка в области
//INPUT
//      ColumnNum: integer  - номер столбца ячейки
//      RowNum: integer     - номер строки ячейки
//RETURN
//      boolean - true - ячейка входит в область
function TZConditionalAreas.IsCellInArea(ColumnNum, RowNum: integer): boolean;
var i, x, y, xx, yy: integer;
begin
  result := false;
  for i := 0 to Count - 1 do begin
    x := FItems[i].Column;
    y := FItems[i].Row;
    xx := x + FItems[i].Width;
    yy := y + FItems[i].Height;
    if ((ColumnNum >= x) and (ColumnNum < xx) and (RowNum >= y) and (RowNum < yy)) then begin
      result := true;
      break;
    end;
  end;
end; //IsCellInArea

function TZConditionalAreas.IsEqual(Source: TPersistent): boolean;
var t: TZConditionalAreas; i: integer;
begin
  result := false;
  if (Source is TZConditionalAreas) then begin
    t := Source as TZConditionalAreas;
    if (Count <> t.Count) then
      exit;
    for i := 0 to Count - 1 do
      if (not FItems[i].IsEqual(t.Items[i])) then
        exit;
    result := true;
  end;
end; //IsEqual

////::::::::::::: TConditionalFormatting :::::::::::::::::////

constructor TZConditionalFormatting.Create();
begin
  FCount := 0;
  SetLength(FStyles, 0);
end; //Create

destructor TZConditionalFormatting.Destroy();
var i: integer;
begin
  for i := 0 to FCount - 1 do
    if (Assigned(FStyles[i])) then
      FreeAndNil(FStyles[i]);
  SetLength(FStyles, 0);
  FStyles := nil;
  inherited;
end; //Destroy

procedure TZConditionalFormatting.SetCount(Value: integer);
var i: integer;
begin
  if (Value >= 0) then begin
    if (Value < FCount) then begin
      for i := Value to FCount - 1 do
         FreeAndNil(FStyles[i]);
      SetLength(FStyles, Value);
    end else if (Value > FCount) then begin
      SetLength(FStyles, Value);
      for i := FCount to Value - 1 do
        FStyles[i] := TZConditionalStyle.Create();
    end;
    FCount := value;
  end;
end; //SetCount

function TZConditionalFormatting.Add(): TZConditionalStyle;
begin
  result := Add(nil);
end;

function TZConditionalFormatting.GetItem(num: integer): TZConditionalStyle;
begin
  result := nil;
  if ((num >= 0) and (num < Count)) then
    result := FStyles[num];
end;

procedure TZConditionalFormatting.SetItem(num: integer; Value: TZConditionalStyle);
begin
  if ((num >= 0) and (num < Count)) then
    if (Assigned(Value)) then
      FStyles[num].Assign(Value);
end;

function TZConditionalFormatting.Add(Style: TZConditionalStyle): TZConditionalStyle;
begin
  Count := Count + 1;
  result := FStyles[Count - 1];
  if (Assigned(Style)) then
    result.Assign(Style);
end;

//Добавить условное форматирование с областью
//INPUT
//      ColumnNum: integer  - номер колонки
//      RowNum: integer     - номер строки
//      AreaWidth: integer  - ширина области
//      AreaHeight: integer - высота области
//RETURN
//      TZConditionalStyle - добавленный стиль
function TZConditionalFormatting.Add(ColumnNum, RowNum, AreaWidth, AreaHeight: integer): TZConditionalStyle;
var t: TZConditionalAreaItem;
begin
  result   := Add(nil);
  t        := result.Areas[0];
  t.Row    := RowNum;
  t.Column := ColumnNum;
  t.Width  := AreaWidth;
  t.Height := AreaHeight;
end;

procedure TZConditionalFormatting.Clear();
begin
  SetCount(0);
end;

//Delete condition formatting item
//INPUT
//      num: integer - number of CF item
//RETURN
//      boolean - true - item deleted
function TZConditionalFormatting.Delete(num: integer): boolean;
var i: integer;
begin
  Result := false;
  if (num >= 0) and (num < Count) then begin
    FreeAndNil(FStyles[num]);
    for i := num to FCount - 2 do
      FStyles[num] := FStyles[num + 1];
    Dec(FCount);
  end;
end;

procedure TZConditionalFormatting.Assign(Source: TPersistent);
var t: TZConditionalFormatting; i: integer;
begin
  if (Source is TZConditionalFormatting) then begin
    t := Source as TZConditionalFormatting;
    FCount := t.Count;
    for i := 0 to FCount - 1 do
      FStyles[i].Assign(t.Items[i]);
  end else
    inherited Assign(Source);
end;

function TZConditionalFormatting.IsEqual(Source: TPersistent): boolean;
var t: TZConditionalFormatting; i: integer;
begin
  result := false;
  if (Source is TZConditionalFormatting) then begin
    t := Source as TZConditionalFormatting;
    if (Count <> t.Count) then
      exit;
    for i := 0 to Count - 1 do
      if (not FStyles[i].IsEqual(t.Items[i])) then
        exit;
    result := true;
  end;
end;

////::::::::::::: TZECommonFrameAncestor :::::::::::::::::////

constructor TZECommonFrameAncestor.Create();
begin
  FX      := 0;
  FY      := 0;
  FWidth  := 10;
  FHeight := 10;
  FTransform := TZETransform.Create();
end;

constructor TZECommonFrameAncestor.Create(AX, AY, AWidth, AHeight: integer);
begin
  FX      := AX;
  FY      := AY;
  FWidth  := AWidth;
  FHeight := AHeight;
  FTransform := TZETransform.Create();
end;

destructor TZECommonFrameAncestor.Destroy();
begin
  FreeAndNil(FTransform);
  inherited;
end;

function TZECommonFrameAncestor.IsEqual(const Source: TPersistent): boolean;
var tmp: TZECommonFrameAncestor;
begin
  Result := false;
  if (Assigned(Source)) then
    if (Source is TZECommonFrameAncestor) then begin
      tmp := Source as TZECommonFrameAncestor;
      Result := (FX = tmp.X) and (FY = tmp.Y) and (FWidth = tmp.Width) and (FHeight = tmp.Height);
      if (Result) then
        Result := FTransform.IsEqual(tmp.Transform);
    end;
end;

procedure TZECommonFrameAncestor.Assign(Source: TPersistent);
var b: boolean; tmp: TZECommonFrameAncestor;
begin
  b := Source <> nil;
  if (b) then
    b := Source is TZECommonFrameAncestor;

  if (b) then begin
    tmp := Source as TZECommonFrameAncestor;
    FX      := tmp.X;
    FY      := tmp.Y;
    FWidth  := tmp.Width;
    FHeight := tmp.Height;
    FTransform.Assign(tmp.Transform);
  end
  else
    inherited Assign(Source);
end;

procedure TZECommonFrameAncestor.SetHeight(value: integer);
begin
  FHeight := value;
end;

procedure TZECommonFrameAncestor.SetTransform(const value: TZETransform);
begin
  if (Assigned(value)) then
    FTransform.Assign(value);
end;

procedure TZECommonFrameAncestor.SetWidth(value: integer);
begin
  FWidth := value;
end;

procedure TZECommonFrameAncestor.SetX(value: integer);
begin
  FX := value;
end;

procedure TZECommonFrameAncestor.SetY(value: integer);
begin
  FY := value;
end;

////::::::::::::: TZETransform :::::::::::::::::////

constructor TZETransform.Create();
begin
  Clear();
end;

procedure TZETransform.Assign(Source: TPersistent);
var b: boolean; tmp: TZETransform;
begin
  b := Assigned(Source);
  if (b) then
    b := Source is TZETransform;

  if (b) then begin
    tmp := Source as TZETransform;
    FRotate     := tmp.Rotate;
    FScaleX     := tmp.ScaleX;
    FScaleY     := tmp.ScaleY;
    FSkewX      := tmp.SkewX;
    FSkewY      := tmp.SkewY;
    FTranslateX := tmp.TranslateX;
    FTranslateY := tmp.TranslateY;
  end else
    inherited Assign(Source);
end;

procedure TZETransform.Clear();
begin
  FRotate     := 0;
  FScaleX     := 1;
  FScaleY     := 1;
  FSkewX      := 0;
  FSkewY      := 0;
  FTranslateX := 0;
  FTranslateY := 0;
end;

function TZETransform.IsEqual(const Source: TPersistent): boolean;
var tmp: TZETransform;
begin
  Result := Assigned(Source);
  if (Result) then
    Result := Source is TZETransform;

  if (Result) then begin
    tmp := Source as TZETransform;
    Result := (FRotate = tmp.Rotate) and
              (FScaleX = tmp.ScaleX) and
              (FScaleY = tmp.ScaleY) and
              (FSkewX = tmp.SkewX) and
              (FSkewY = tmp.SkewY) and
              (FTranslateX = tmp.TranslateX) and
              (FTranslateY = tmp.TranslateY);
  end;
end;

////::::::::::::: TZEChartRangeItem :::::::::::::::::////

constructor TZEChartRangeItem.Create();
begin
  FSheetNum := -1;
  FRow      := 0;
  FCol      := 0;
  FWidth    := 1;
  FHeight   := 1;
end;

constructor TZEChartRangeItem.Create(ASheetNum: integer; ACol, ARow, AWidth, AHeight: integer);
begin
  FSheetNum := ASheetNum;
  FRow      := ARow;
  FCol      := ACol;
  FWidth    := AWidth;
  FHeight   := AHeight;
end;

procedure TZEChartRangeItem.Assign(Source: TPersistent);
var tmp: TZEChartRangeItem; b: boolean;
begin
  b := Assigned(Source);
  if (b) then begin
    b := Source is TZEChartRangeItem;
    if (b) then begin
      tmp := Source as TZEChartRangeItem;
      FSheetNum := tmp.SheetNum;
      FRow      := tmp.Row;
      FCol      := tmp.Col;
      FWidth    := tmp.Width;
      FHeight   := tmp.Height;
    end;
  end;

  if (not b) then
    inherited Assign(Source);
end;

function TZEChartRangeItem.IsEqual(const Source: TPersistent): boolean;
var tmp: TZEChartRangeItem;
begin
  Result := Assigned(Source);
  if (Result) then begin
    Result := Source is TZEChartRangeItem;
    if (Result) then begin
      tmp := Source as TZEChartRangeItem;
      Result := (FSheetNum = tmp.SheetNum) and
                (FCol = tmp.Col) and
                (FRow = tmp.Row) and
                (FHeight = tmp.Height) and
                (FWidth = tmp.Width);
    end;
  end;
end;

////::::::::::::: TZEChartRange :::::::::::::::::////

constructor TZEChartRange.Create();
begin
  FCount := 0;
end;

destructor TZEChartRange.Destroy();
begin
  Clear();
  inherited;
end;

function TZEChartRange.GetItem(num: integer): TZEChartRangeItem;
begin
  Result := nil;
  if ((num >= 0) and (num < FCount)) then
    Result := FItems[num];
end;

procedure TZEChartRange.SetItem(num: integer; const Value: TZEChartRangeItem);
begin
  if ((num >= 0) and (num < FCount)) then
    FItems[num].Assign(Value);
end;

function TZEChartRange.Add(): TZEChartRangeItem;
begin
  SetLength(FItems, FCount + 1);
  Result := TZEChartRangeItem.Create();
  FItems[FCount] := result;
  inc(FCount);
end;

function TZEChartRange.Add(const ItemForClone: TZEChartRangeItem): TZEChartRangeItem;
begin
  Result := Add();
  if (Assigned(ItemForClone)) then
    Result.Assign(ItemForClone);
end;

function TZEChartRange.Delete(num: integer): boolean;
var i: integer;
begin
  Result := (num >= 0) and (num < FCount);
  if (Result) then begin
    FreeAndNil(FItems[num]);
    for i := num to FCount - 2 do
      FItems[i] := FItems[i + 1];
    Dec(FCount);
    SetLength(FItems, FCount);
  end;
end;

procedure TZEChartRange.Clear();
var i: integer;
begin
  for i := 0 to FCount - 1 do
    FreeAndNil(FItems[i]);
  FCount := 0;
  SetLength(FItems, 0);
end;

procedure TZEChartRange.Assign(Source: TPersistent);
var tmp: TZEChartRange; b: boolean; i: integer;
begin
  b := Assigned(Source);
  if (b) then begin
    b := Source is TZEChartRange;
    if (b) then begin
      tmp := Source as TZEChartRange;
      if (FCount > tmp.Count) then begin
        for i := tmp.Count to FCount - 1 do
          FreeAndNil(FItems[i]);
        FCount := tmp.Count;
        SetLength(FItems, FCount);
      end else if (FCount < tmp.Count) then begin
        SetLength(FItems, tmp.Count);
        for i := FCount to tmp.Count - 1 do
          FItems[i] := TZEChartRangeItem.Create();
        FCount := tmp.Count;
        SetLength(FItems, FCount);
      end;

      for i := 0 to FCount - 1 do
        FItems[i].Assign(tmp.Items[i]);
    end;
  end;

  if (not b) then
    inherited Assign(Source);
end;

function TZEChartRange.IsEqual(const Source: TPersistent): boolean;
var tmp: TZEChartRange; i: integer;
begin
  Result := Assigned(Source);
  if (Result) then begin
    Result := Source is TZEChartRange;
    if (Result) then begin
      tmp := Source as TZEChartRange;
      Result := FCount = tmp.Count;
      if (Result) then
        for i := 0 to FCount - 1 do
          if (not FItems[i].IsEqual(tmp.Items[i])) then begin
            Result := false;
            break;
          end;
    end;
  end;
end;

////::::::::::::: TZEChartTitleItem :::::::::::::::::////

constructor TZEChartTitleItem.Create();
begin
  FFont          := TFont.Create();
  FText          := '';
  FRotationAngle := 0;
  FIsDisplay     := true;
end;

destructor TZEChartTitleItem.Destroy();
begin
  FreeAndNil(FFont);
  inherited;
end;

procedure TZEChartTitleItem.Assign(Source: TPersistent);
var tmp: TZEChartTitleItem; b: boolean;
begin
  b := Assigned(Source);
  if (b) then begin
    b := Source is TZEChartTitleItem;
    if (b) then begin
      tmp := Source as TZEChartTitleItem;

      FFont.Assign(tmp.Font);
      FText          := tmp.Text;
      FRotationAngle := tmp.RotationAngle;
      FIsDisplay     := tmp.IsDisplay;
    end;
  end;

  if (not b) then
    inherited Assign(Source);
end;

function TZEChartTitleItem.IsEqual(const Source: TPersistent): boolean;
var tmp: TZEChartTitleItem;
begin
  Result := Assigned(Source);
  if (Result) then begin
    Result := Source is TZEChartTitleItem;
    if (Result) then begin
      tmp := Source as TZEChartTitleItem;
      Result := false;
      if (FIsDisplay = tmp.IsDisplay) then
        if (FRotationAngle = tmp.RotationAngle) then
          if (FText = tmp.Text) then
            Result := ZEIsFontsEquals(FFont, tmp.Font);
    end;
  end;
end; //IsEqual

procedure TZEChartTitleItem.SetFont(const value: TFont);
begin
  if (Assigned(value)) then
    FFont.Assign(value);
end;

////::::::::::::: TZEChartLegend :::::::::::::::::////

constructor TZEChartLegend.Create();
begin
  inherited;
  FPosition := ZELegendStart;
  FAlign := ZELegendAlignCenter;
end;

procedure TZEChartLegend.Assign(Source: TPersistent);
var tmp: TZEChartLegend;
begin
  inherited Assign(Source);

  if Assigned(Source) and (Source is TZEChartLegend) then begin
    tmp := Source as TZEChartLegend;
    FAlign    := tmp.Align;
    FPosition := tmp.Position;
  end;
end; //Assign

function TZEChartLegend.IsEqual(const Source: TPersistent): boolean;
var tmp: TZEChartLegend;
begin
  Result := inherited IsEqual(Source);
  if (Result) then begin
    Result := false;
    if (Source is TZEChartLegend) then begin
      tmp := Source as TZEChartLegend;
      Result := (FPosition = tmp.Position) and
                (FAlign = tmp.Align);
    end;
  end;
end; //IsEqual


////::::::::::::: TZEChartAxis :::::::::::::::::////

constructor TZEChartAxis.Create();
begin
  inherited;
  FLogarithmic := false;
  FReverseDirection := false;
  FScaleMin := 0;
  FScaleMax := 20000;
  FAutoScaleMin := true;
  FAutoScaleMax := true;
end;

procedure TZEChartAxis.Assign(Source: TPersistent);
var tmp: TZEChartAxis;
begin
  inherited Assign(Source);
  if (Source is TZEChartAxis) then begin
    tmp := Source as TZEChartAxis;
    FLogarithmic      := tmp.Logarithmic;
    FReverseDirection := tmp.ReverseDirection;
    FScaleMin         := tmp.ScaleMin;
    FScaleMax         := tmp.ScaleMax;
    FAutoScaleMin     := tmp.AutoScaleMin;
    FAutoScaleMax     := tmp.AutoScaleMax;
  end
end;

function TZEChartAxis.IsEqual(const Source: TPersistent): boolean;
var tmp: TZEChartAxis;
begin
  Result := inherited IsEqual(Source);
  if (Result) then begin
    Result := Source is TZEChartAxis;
    if (Result) then begin
      tmp := Source as TZEChartAxis;
      Result := (FLogarithmic = tmp.FLogarithmic) and
                (FReverseDirection = tmp.FReverseDirection) and
                (FAutoScaleMin = tmp.AutoScaleMin) and
                (FAutoScaleMax = tmp.AutoScaleMax);
      if (Result) then begin
        //TODO: for comparsion double values need check observational errors?
        if (not FAutoScaleMin) then
          Result := FScaleMin <> tmp.ScaleMin;
        if (Result and (not FAutoScaleMin)) then
          Result := FScaleMax <> tmp.ScaleMax;
      end;
    end;
  end;
end;

////::::::::::::: TZEChartSettings :::::::::::::::::////

constructor TZEChartSettings.Create();
begin
  FJapanCandle := true;
end;

destructor TZEChartSettings.Destroy();
begin
  inherited;
end;

procedure TZEChartSettings.Assign(Source: TPersistent);
var tmp: TZEChartSettings; b: boolean;
begin
  b := Source <> nil;
  if (b) then begin
    b := Source is TZEChartSettings;
    if (b) then begin
      tmp := Source as TZEChartSettings;
      FJapanCandle := tmp.JapanCandle;
    end;
  end;
  if (not b) then
    inherited Assign(Source);
end;

function TZEChartSettings.IsEqual(const Source: TPersistent): boolean;
var tmp: TZEChartSettings;
begin
  Result := Assigned(Source);
  if (Result) then begin
    Result := Source is TZEChartSettings;
    if (Result) then begin
      tmp := Source as TZEChartSettings;
      Result := FJapanCandle = tmp.JapanCandle;
    end;
  end;
end;

////::::::::::::: TZEChartSeries :::::::::::::::::////

constructor TZEChartSeries.Create();
begin
  FChartType := ZEChartTypeBar;
  FSeriesName := '';
  FSeriesNameSheet := -1;
  FSeriesNameRow := -1;
  FSeriesNameCol := -1;
  FRanges := TZEChartRange.Create();
end;

destructor TZEChartSeries.Destroy();
begin
  FreeAndNil(FRanges);
  inherited;
end;

procedure TZEChartSeries.Assign(Source: TPersistent);
var tmp: TZEChartSeries; b: boolean;
begin
  b := Assigned(Source);

  if (b) then begin
    b := Source is TZEChartSeries;
    if (b) then begin
      tmp := TZEChartSeries.Create();
      FChartType       := tmp.ChartType;
      FSeriesName      := tmp.SeriesName;
      FSeriesNameSheet := tmp.SeriesNameSheet;
      FSeriesNameRow   := tmp.SeriesNameRow;
      FSeriesNameCol   := tmp.SeriesNameCol;
      FRanges.Assign(tmp.Ranges);
    end;
  end;

  if (not b) then
    inherited Assign(Source);
end;

function TZEChartSeries.IsEqual(const Source: TPersistent): boolean;
var tmp: TZEChartSeries;
begin
  Result := Source <> nil;
  if (Result) then
    Result := Source is TZEChartSeries;
  if (Result) then begin
    tmp := TZEChartSeries.Create();
    Result := (FChartType = tmp.ChartType) and
              (FSeriesNameSheet = tmp.SeriesNameSheet);
    if (Result) then begin
      if (((FSeriesNameRow < 0) or (FSeriesNameCol < 0)) and
           (tmp.SeriesNameRow < 0) or (tmp.SeriesNameCol < 0)) then
        Result := tmp.SeriesName = FSeriesName
      else
        Result := (tmp.SeriesNameRow = FSeriesNameRow) and
                  (tmp.SeriesNameCol = FSeriesNameCol);
    end;
    if (Result) then
      Result := FRanges.IsEqual(tmp.Ranges);
  end;
end;

////::::::::::::: TZEChart :::::::::::::::::////

procedure TZEChart.CommonInit();
begin
  FTitle            := TZEChartTitleItem.Create();
  FSubtitle         := TZEChartTitleItem.Create();
  FLegend           := TZEChartLegend.Create();
  FFooter           := TZEChartTitleItem.Create();
  FAxisX            := TZEChartAxis.Create();
  FAxisY            := TZEChartAxis.Create();
  FAxisZ            := TZEChartAxis.Create();
  FSecondaryAxisX   := TZEChartAxis.Create();
  FSecondaryAxisY   := TZEChartAxis.Create();
  FSecondaryAxisZ   := TZEChartAxis.Create();
  FDefaultChartType := ZEChartTypeBar;
  FView3D           := false;
  FViewDeep         := false;
end;

constructor TZEChart.Create(AX, AY, AWidth, AHeight: integer);
begin
  inherited;
  CommonInit();
  X := AX;
  Y := AY;
  Width  := AWidth;
  Height := AHeight;
end;

constructor TZEChart.Create();
begin
  inherited;
  CommonInit();
end;

destructor TZEChart.Destroy();
begin
  FreeAndNil(FSubtitle);
  FreeAndNil(FTitle);
  FreeAndNil(FLegend);
  FreeAndNil(FFooter);
  FreeAndNil(FAxisX);
  FreeAndNil(FAxisY);
  FreeAndNil(FAxisZ);
  FreeAndNil(FSecondaryAxisX);
  FreeAndNil(FSecondaryAxisY);
  FreeAndNil(FSecondaryAxisZ);

  inherited;
end;

function TZEChart.IsEqual(const Source: TPersistent): boolean;
var tmp: TZEChart;
begin
  Result := inherited IsEqual(Source);
  if Result and (Source is TZEChart) then begin
    tmp := Source as TZEChart;
    Result := inherited IsEqual(Source);

    if (Result) then
      Result := FSubtitle.IsEqual(tmp.Subtitle) and
                FTitle.IsEqual(tmp.Title) and
                FLegend.IsEqual(tmp.Legend) and
                FAxisX.IsEqual(tmp.AxisX) and
                FAxisY.IsEqual(tmp.AxisY) and
                FAxisZ.IsEqual(tmp.AxisZ) and
                FSecondaryAxisX.IsEqual(tmp.SecondaryAxisX) and
                FSecondaryAxisY.IsEqual(tmp.SecondaryAxisY) and
                FSecondaryAxisZ.IsEqual(tmp.SecondaryAxisZ) and
                (FView3D = tmp.View3D) and
                (FViewDeep = tmp.ViewDeep);
  end;
end;

procedure TZEChart.SetAxisX(const value: TZEChartAxis);
begin
  if (Assigned(value)) then
    FAxisX.Assign(Value);
end;

procedure TZEChart.SetAxisY(const value: TZEChartAxis);
begin
  if (Assigned(value)) then
    FAxisY.Assign(Value);
end;

procedure TZEChart.SetAxisZ(const value: TZEChartAxis);
begin
  if (Assigned(value)) then
    FAxisZ.Assign(Value);
end;

procedure TZEChart.SetSecondaryAxisX(const value: TZEChartAxis);
begin
  if (Assigned(value)) then
    FSecondaryAxisX.Assign(value);
end;

procedure TZEChart.SetSecondaryAxisY(const value: TZEChartAxis);
begin
  if (Assigned(value)) then
    FSecondaryAxisY.Assign(value);
end;

procedure TZEChart.SetSecondaryAxisZ(const value: TZEChartAxis);
begin
  if (Assigned(value)) then
    FSecondaryAxisZ.Assign(value);
end;

procedure TZEChart.SetFooter(const value: TZEChartTitleItem);
begin
  if (Assigned(value)) then
    FFooter.Assign(value);
end;

procedure TZEChart.SetLegend(const value: TZEChartLegend);
begin
  if (Assigned(value)) then
    FLegend.Assign(value);
end;

procedure TZEChart.SetSubtitle(const value: TZEChartTitleItem);
begin
  if (Assigned(value)) then
    FSubtitle.Assign(value);
end;

procedure TZEChart.SetTitle(const value: TZEChartTitleItem);
begin
  if (Assigned(value)) then
    FTitle.Assign(value);
end;

procedure TZEChart.Assign(Source: TPersistent);
var tmp: TZEChart;
begin
  inherited;
  if Assigned(Source) and (Source is TZEChart) then begin
    tmp := Source as TZEChart;
    FSubtitle.Assign(tmp.Subtitle);
    FTitle.Assign(tmp.Title);
    FLegend.Assign(tmp.Legend);
    FAxisX.Assign(tmp.AxisX);
    FAxisY.Assign(tmp.AxisY);
    FAxisZ.Assign(tmp.AxisZ);
    FSecondaryAxisX.Assign(tmp.SecondaryAxisX);
    FSecondaryAxisY.Assign(tmp.SecondaryAxisY);
    FSecondaryAxisZ.Assign(tmp.SecondaryAxisZ);
    FView3D := tmp.View3D;
    FViewDeep := tmp.ViewDeep;
  end;
end;

////::::::::::::: TZEChartStore :::::::::::::::::////

constructor TZEChartStore.Create();
begin
  FCount := 0;
  SetLength(FItems, 0);
end;

destructor TZEChartStore.Destroy();
begin
  Clear();
  inherited;
end;

function TZEChartStore.Add(): TZEChart;
begin
  Result := TZEChart.Create();
  SetLength(FItems, FCount + 1);
  FItems[FCount] := Result;
  inc(FCount);
end;

function TZEChartStore.Add(const ItemForClone: TZEChart): TZEChart;
begin
  Result := Add();
  Result.Assign(ItemForClone);
end;

procedure TZEChartStore.Assign(Source: TPersistent);
var tmp: TZEChartStore; b: boolean; i: integer;
begin
  b := Assigned(Source);
  if (b) then begin
    b := Source is TZEChartStore;
    if (b) then begin
      tmp := Source as TZEChartStore;

      if (FCount > tmp.Count) then begin
        for i := tmp.Count to FCount - 1 do
          FreeAndNil(FItems[i]);
        Setlength(FItems, tmp.Count);
      end else if (FCount < tmp.Count) then begin
        Setlength(FItems, tmp.Count);
        for i := FCount to tmp.Count - 1 do
          FItems[i] := TZEChart.Create();
      end;
      FCount := tmp.Count;

      for i := 0 to FCount - 1 do
        FItems[i].Assign(tmp[i]);
    end;
  end;

  if (not b) then
    inherited Assign(Source);
end;

function TZEChartStore.Delete(num: integer): boolean;
var i: integer;
begin
  Result := (num >= 0) and (num < FCount);
  if (Result) then begin
    FreeAndNil(FItems[num]);
    for i := num to FCount - 2 do
      FItems[i] := FItems[i + 1];
    dec(FCount);
    SetLength(FItems, FCount);
  end;
end;

procedure TZEChartStore.Clear();
var i: integer;
begin
  for i := 0 to FCount - 1 do
    FreeAndNil(FItems[i]);
  FCount := 0;
  SetLength(FItems, 0);
end;

function TZEChartStore.GetItem(num: integer): TZEChart;
begin
  Result := nil;
  if ((num >= 0) and (num < FCount)) then
    Result := FItems[num];
end;

procedure TZEChartStore.SetItem(num: integer; const Value: TZEChart);
begin
  if ((num >= 0) and (num < FCount)) then
    FItems[num].Assign(Value);
end;

function TZEChartStore.IsEqual(const Source: TPersistent): boolean;
var tmp: TZEChartStore; i: integer;
begin
  Result := Assigned(Source);
  if (Result) then begin
    Result := Source is TZEChartStore;
    if (Result) then begin
      tmp := Source as TZEChartStore;
      Result := FCount = tmp.Count;
      if (Result) then
        for i := 0 to FCount - 1 do
          if (not FItems[i].IsEqual(tmp[i])) then begin
            Result := false;
            break;
          end;
    end;
  end;
end;

////::::::::::::: TZEPicture :::::::::::::::::////

procedure TZEPicture.Assign(Source: TPersistent);
var tmp: TZEPicture;
begin
  inherited;
  if Assigned(Source) and (Source is TZEPicture) then begin
    tmp := Source as TZEPicture;
    FId := tmp.Id;
    FTitle := tmp.Title;
    FDescription := tmp.Description;
    FRow := tmp.Row;
    FCol := tmp.Col;
    if Assigned(tmp.DataStream) then
      AssignFromStream(tmp.DataStream);
  end;
end;

procedure TZEPicture.AssignFromFile(AFileName: string);
begin
  if Assigned(FDataStream) then
    FreeAndNil(FDataStream);
  FFileName := AFileName;
end;

procedure TZEPicture.AssignFromStream(AStream: TStream);
var fcc: LongWord;
begin
  if Assigned(FDataStream) then
    FreeAndNil(FDataStream);
  FFileName := '';
  FDataStream := TMemoryStream.Create();
  AStream.Position := 0;
  // detect file type
  AStream.Read(fcc, SizeOf(fcc));
  if fcc = $474E5089 then
    FFileName := '.png'
  else if fcc = $E0FFD8FF then
    FFileName := '.jpeg';
  AStream.Position := 0;
  FDataStream.CopyFrom(AStream, AStream.Size);
end;

procedure TZEPicture.CommonInit();
begin
  FId := 0;
  FRelId := 0;
  FTitle := '';
  FDescription := '';
  FRow := 0;
  FCol := 0;
  FCellAnchor := ZACell;
end;

constructor TZEPicture.Create(AX, AY, AWidth, AHeight: integer);
begin
  inherited;
  CommonInit();
  X := AX;
  Y := AY;
  Width := AWidth;
  Height := AHeight;
end;

constructor TZEPicture.Create();
begin
  inherited;
  CommonInit();
end;

destructor TZEPicture.Destroy;
begin
  if Assigned(FDataStream) then
    FreeAndNil(FDataStream);
  inherited;
end;

function TZEPicture.GetName(): string;
begin
  if Pos('.', FFileName) > 0 then
    Result := 'image' + IntToStr(Id) + ExtractFileExt(FFileName)
  else
    Result := 'image' + IntToStr(Id) + '.png';
end;

function TZEPicture.GetRelIdStr(): string;
begin
  Result := 'rId' + IntToStr(RelId);
end;

function TZEPicture.IsEqual(const Source: TPersistent): boolean;
var tmp: TZEPicture;
begin
  Result := inherited IsEqual(Source);
  if (Result) and (Source is TZEPicture) then begin
    tmp := Source as TZEPicture;
    Result := (FId = tmp.Id)
          and (FTitle = tmp.Title)
          and (FDescription = tmp.Description)
          and (FHidden = tmp.Hidden);
          // TODO: compare filename/streams
  end;
end;

{ TZEPictureStore }

function TZEPictureStore.Add(const ItemForClone: TZEPicture): TZEPicture;
begin
  Result := Add();
  Result.Assign(ItemForClone);
end;

function TZEPictureStore.Add(): TZEPicture;
begin
  Result := TZEPicture.Create();
  FItems.Add(Result);
end;

procedure TZEPictureStore.Assign(Source: TPersistent);
var tmp: TZEPictureStore;
    b: boolean;
    i: integer;
begin
  b := Assigned(Source);
  if (b) then begin
    b := Source is TZEPictureStore;
    if (b) then begin
      tmp := Source as TZEPictureStore;

      if (Count > tmp.Count) then begin
        for i := Count - 1 downto tmp.Count do
          Delete(i);
      end else if (Count < tmp.Count) then begin
        for i := Count to tmp.Count - 1 do
          Add();
      end;

      for i := 0 to Count - 1 do
        Items[i].Assign(tmp[i]);
    end;
  end;

  if (not b) then
    inherited Assign(Source);
end;

procedure TZEPictureStore.Clear();
var i: integer;
begin
  for i := Count - 1 downto 0 do
    Delete(i);
end;

constructor TZEPictureStore.Create();
begin
  FItems := TObjectList.Create(true);
end;

function TZEPictureStore.Delete(num: integer): boolean;
begin
  Result := (num >= 0) and (num < Count);
  if (Result) then begin
    TZEPicture(FItems[num]).Free();
    FItems.Delete(num);
  end;
end;

destructor TZEPictureStore.Destroy();
begin
  Clear();
  FItems.Free();
  inherited;
end;

function TZEPictureStore.GetCount: integer;
begin
  Result := FItems.Count;
end;

function TZEPictureStore.GetItem(num: integer): TZEPicture;
begin
  Result := nil;
  if ((num >= 0) and (num < Count)) then
    Result := TZEPicture(FItems[num]);
end;

function TZEPictureStore.IsEqual(const Source: TPersistent): boolean;
var tmp: TZEPictureStore;
    i: integer;
begin
  Result := False;
  if Assigned(Source) and (Source is TZEPictureStore) then begin
    tmp := Source as TZEPictureStore;
    Result := (Count = tmp.Count);
    if (Result) then begin
      for i := 0 to Count - 1 do begin
        if (not Items[i].IsEqual(tmp[i])) then begin
          Result := False;
          break;
        end;
      end;
    end;
  end;
end;

procedure TZEPictureStore.SetItem(num: integer; const Value: TZEPicture);
begin
  if ((num >= 0) and (num < FItems.Count)) then
    TZEPicture(FItems[num]).Assign(Value);
end;

{ TZEDrawing }

constructor TZEDrawing.Create();
begin
  inherited;
  FPictureStore := TZEPictureStore.Create();
end;

destructor TZEDrawing.Destroy();
begin
  FPictureStore.Clear();
  FPictureStore.Free();
  inherited;
end;

procedure TZEDrawing.Assign(Source: TPersistent);
var tmp: TZEDrawing; b: boolean;
begin
  b := Assigned(Source);
  if (b) then begin
    b := Source is TZEDrawing;
    if (b) then begin
      tmp := Source as TZEDrawing;

      self.FId := tmp.FId;
      self.FPictureStore := TZEPictureStore.Create();
      self.FPictureStore.Assign(tmp.FPictureStore);
    end;
  end;

  if (not b) then
    inherited Assign(Source);
end;

function TZEDrawing.AddPictureFile(ARow, ACol: Integer; AFileName: string): TZEPicture;
begin
  Result := PictureStore.Add();
  Result.AssignFromFile(AFileName);
  Result.RelId := PictureStore.Count;
  Result.Row := ARow;
  Result.Col := ACol;
  Result.CellAnchor := ZACell;
end;

function TZEDrawing.AddPictureStream(ARow, ACol: Integer; AStream: TStream): TZEPicture;
begin
  Result := PictureStore.Add();
  Result.AssignFromStream(AStream);
  Result.RelId := PictureStore.Count;
  Result.Row := ARow;
  Result.Col := ACol;
  Result.CellAnchor := ZACell;
end;

function TZEDrawing.GetIsEmpty(): Boolean;
begin
  Result := (PictureStore.Count = 0);
end;

{
GetDeviceCaps(hdc, HORZSIZE) / GetDeviceCaps(hdc, HORZRES); //горизонтальный размер пикселя в миллиметрах
GetDeviceCaps(hdc, VERTSIZE) / GetDeviceCaps(hdc, VERTRES); //вертикальный размер пикселя в миллиметрах
}

{
  Paper Size Table
 Index    Paper type               Paper size
 ----------------------------------------------------
 0        Undefined
 1        Letter                   8 1/2" x 11"
 2        Letter small             8 1/2" x 11"
 3        Tabloid                     11" x 17"
 4        Ledger                      17" x 11"
 5        Legal                    8 1/2" x 14"
 6        Statement                5 1/2" x 8 1/2"
 7        Executive                7 1/4" x 10 1/2"
 8        A3                        297mm x 420mm
 9        A4                        210mm x 297mm
 10       A4 small                  210mm x 297mm
 11       A5                        148mm x 210mm
 12       B4                        250mm x 354mm
 13       B5                        182mm x 257mm
 14       Folio                    8 1/2" x 13"
 15       Quarto                    215mm x 275mm
 16                                   10" x 14"
 17                                   11" x 17"
 18       Note                     8 1/2" x 11"
 19       #9 Envelope              3 7/8" x 8 7/8"
 20       #10 Envelope             4 1/8" x 9 1/2"
 21       #11 Envelope             4 1/2" x 10 3/8"
 22       #12 Envelope             4 3/4" x 11"
 23       #14 Envelope                 5" x 11 1/2"
 24       C Sheet                     17" x 22"
 25       D Sheet                     22" x 34"
 26       E Sheet                     34" x 44"
 27       DL Envelope               110mm x 220mm
 28       C5 Envelope               162mm x 229mm
 29       C3 Envelope               324mm x 458mm
 30       C4 Envelope               229mm x 324mm
 31       C6 Envelope               114mm x 162mm
 32       C65 Envelope              114mm x 229mm
 33       B4 Envelope               250mm x 353mm
 34       B5 Envelope               176mm x 250mm
 35       B6 Envelope               125mm x 176mm
 36       Italy Envelope            110mm x 230mm
 37       Monarch Envelope         3 7/8" x 7 1/2"
 38       6 3/4 Envelope           3 5/8" x 6 1/2"
 39       US Standard Fanfold     14 7/8" x 11"
 40       German Std. Fanfold      8 1/2" x 12"
 41       German Legal Fanfold     8 1/2" x 13"
}

{ TZRange }
procedure TZRange.Assign(Source: TPersistent);
var src: TZRange; r,c,id: Integer; style: TZStyle; rect: TRect;
begin
  inherited;
  if Source is TZRange then begin
    src := TZRange(Source);

    // resize columns or rows if need
    if FSheet.RowCount <= Max(src.FR2, FR2) then
      FSheet.RowCount := Max(src.FR2, FR2);

    if FSheet.ColCount <= Max(src.FC2, FC2) then
      FSheet.ColCount := Max(src.FC2, FC2);

    // copy cells and styles
    for c := 0 to src.FC2-src.FC1 do begin
      for r := 0 to src.FR2-src.FR1 do begin
        FSheet.Cell[FC1+c, FR1+r].Assign(src.FSheet.Cell[src.FC1+c, src.FR1+r]);
        id := FSheet.Cell[FC1+c, FR1+r].CellStyle;
        if id > -1 then begin
          style := TZStyle.Create();
          // style must copy from source sheet
          style.Assign( src.FSheet.WorkBook.Styles[id] );
          id := FSheet.WorkBook.Styles.Add(style, true);
          FSheet.Cell[FC1+c, FR1+r].CellStyle := id;
        end;
      end;
    end;

    // remove cross merges
    for id := FSheet.MergeCells.Count-1 downto 0 do begin
      if FSheet.MergeCells.IsCrossWithArea(id, FC1, FR1, FC1+(src.FC2-src.FC1), FR1+(src.FR2-src.FR1)) then
        FSheet.MergeCells.DeleteItem(id);
    end;

    // copy and reloc merges
    for id := 0 to src.FSheet.MergeCells.Count-1 do begin
      if src.FSheet.MergeCells.IsCrossWithArea(id, src.FC1, src.FR1, src.FC2, src.FR2) then begin
        rect := src.FSheet.MergeCells[id];
        FSheet.MergeCells.AddRectXY(
          FC1+(rect.Left-src.FC1),
          FR1+(rect.Top-src.FR1),
          FC1+(rect.Right-src.FC1),
          FR1+(rect.Bottom-src.FR1)
        );
      end;
    end;
  end;
end;

constructor TZRange.Create(ASheet: TZSheet; AFC1, AFR1, AFC2, AFR2: Integer);
begin
  FSheet := ASheet;
  FC1 := AFC1;
  FR1 := AFR1;
  FC2 := AFC2;
  FR2 := AFR2;
end;

destructor TZRange.Destroy;
begin
  FSheet := nil;
  FC1 := 0;
  FR1 := 0;
  FC2 := 0;
  FR2 := 0;
  inherited;
end;

function TZRange.HasStyle: Boolean;
begin
  Result := FSheet.Cell[FC1, FR1].FCellStyle > -1;
end;

procedure TZRange.Merge;
var I: Integer;
begin
  for I := FSheet.MergeCells.Count-1 downto 0 do begin
    if FSheet.MergeCells.IsCrossWithArea(I, FC1, FR1, FC2, FR2) then
      FSheet.MergeCells.DeleteItem(I);
  end;
  FSheet.MergeCells.AddRectXY(FC1, FR1, FC2, FR2);
end;

procedure TZRange.ApplyStyleValue(proc: TProc<TZStyle>);
var col, row, id: Integer; style: TZStyle;
begin
  for col := FC1 to FC2 do begin
    for row := FR1 to FR2 do begin
      style := TZStyle.Create();
      id := FSheet.Cell[col, row].CellStyle;
      if id > -1 then
         style.Assign(FSheet.FStore.Styles[id])
      else if (FSheet.FStore.Styles.Count > 0) then
         style.Assign(FSheet.FStore.Styles[0]);

      proc(style);

      id := FSheet.FStore.Styles.Add(style, true);
      FSheet.Cell[col, row].CellStyle := id;
    end;
  end;
end;

procedure TZRange.Clear;
var col, row: Integer;
begin
  for col := FC1 to FC2 do begin
    for row := FR1 to FR2 do begin
      FSheet.Cell[col, row].Data := '';
      FSheet.Cell[col, row].Formula := '';
      FSheet.Cell[col, row].Comment := '';
      FSheet.Cell[col, row].CommentAuthor := '';
    end;
  end;
end;

function TZRange.GetBgColor: TColor;
begin
  Result := 0;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FC1,FR1].FCellStyle].BGColor;
end;

function TZRange.GetBorderColor(Num: TZBordersPos): TColor;
begin
  Result := 0;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FC1,FR1].FCellStyle].Border[Num].Color;
end;

function TZRange.GetBorderStyle(Num: TZBordersPos): TZBorderType;
begin
  Result := TZBorderType.ZENone;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FC1,FR1].FCellStyle].Border[Num].LineStyle;
end;

function TZRange.GetBorderWidht(Num: TZBordersPos): Byte;
begin
  Result := 0;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FC1,FR1].FCellStyle].Border[Num].Weight;
end;

function TZRange.GetFontColor: TColor;
begin
  Result := 0;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FC1,FR1].FCellStyle].Font.Color;
end;

function TZRange.GetFontSize: Byte;
begin
  Result := 0;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FC1,FR1].FCellStyle].Font.Size;
end;

function TZRange.GetFontStyle: TFontStyles;
begin
  Result := [];
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FC1,FR1].FCellStyle].Font.Style;
end;

function TZRange.GetHorizontalAlignment: TZHorizontalAlignment;
begin
  Result := TZHorizontalAlignment.ZHAutomatic;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FC1,FR1].FCellStyle].Alignment.Horizontal;
end;

function TZRange.GetNumberFormat: string;
begin
  Result := '';
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FC1,FR1].FCellStyle].NumberFormat;
end;

function TZRange.GetVerticalAlignment: TZVerticalAlignment;
begin
  Result := TZVerticalAlignment.ZVAutomatic;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FC1,FR1].FCellStyle].Alignment.Vertical;
end;

function TZRange.GetVerticalText: Boolean;
begin
  Result := false;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FC1,FR1].FCellStyle].Alignment.VerticalText;
end;

function TZRange.GetWrapText: Boolean;
begin
  Result := false;
  if HasStyle then
    Result := FSheet.FStore.FStyles[FSheet.Cell[FC1,FR1].FCellStyle].Alignment.WrapText;
end;

procedure TZRange.SetBgColor(const Value: TColor);
begin
  ApplyStyleValue(procedure (style: TZStyle) begin
    style.BGColor := Value;
  end);
end;

procedure TZRange.SetBorderColor(Num: TZBordersPos; const Value: TColor);
begin
  ApplyStyleValue(procedure (style: TZStyle) begin
    style.Border[Num].Color := Value;
  end);
end;

procedure TZRange.SetBorderStyle(Num: TZBordersPos; const Value: TZBorderType);
begin
  ApplyStyleValue(procedure (style: TZStyle) begin
    style.Border[Num].LineStyle := Value;
  end);
end;

procedure TZRange.SetBorderWidht(Num: TZBordersPos; const Value: Byte);
begin
  ApplyStyleValue(procedure (style: TZStyle) begin
    style.Border[Num].Weight := Value;
  end);
end;

procedure TZRange.SetFontColor(const Value: TColor);
begin
  ApplyStyleValue(procedure (style: TZStyle) begin
    style.Font.Color := Value;
  end);
end;

procedure TZRange.SetFontSize(const Value: Byte);
begin
  ApplyStyleValue(procedure (style: TZStyle) begin
    style.Font.Size := Value;
  end);
end;

procedure TZRange.SetFontStyle(const Value: TFontStyles);
begin
  ApplyStyleValue(procedure (style: TZStyle) begin
    style.Font.Style := Value;
  end);
end;

procedure TZRange.SetHorizontalAlignment(const Value: TZHorizontalAlignment);
begin
  ApplyStyleValue(procedure (style: TZStyle) begin
    style.Alignment.Horizontal := Value;
  end);
end;

procedure TZRange.SetNumberFormat(const Value: string);
begin
  ApplyStyleValue(procedure (style: TZStyle) begin
    style.NumberFormat := Value;
  end);
end;

procedure TZRange.SetVerticalAlignment(const Value: TZVerticalAlignment);
begin
  ApplyStyleValue(procedure (style: TZStyle) begin
    style.Alignment.Vertical := Value;
  end);
end;

procedure TZRange.SetVerticalText(const Value: Boolean);
begin
  ApplyStyleValue(procedure (style: TZStyle) begin
    style.Alignment.VerticalText := Value;
  end);
end;

procedure TZRange.SetWrapText(const Value: Boolean);
begin
  ApplyStyleValue(procedure (style: TZStyle) begin
    style.Alignment.WrapText := Value;
  end);
end;

end.
