import * as docx from 'docx'
import xml2js from 'xml2js'

// =====================
// Temel Tipler
// =====================

export enum Version {
  v1_7 = '1.7',
  v1_8 = '1.8',
}

export enum OrientationEnum {
  portrait = 1,
  landscape = 2,
}

export const Orientations: Record<OrientationEnum, keyof typeof OrientationEnum> = {
  [OrientationEnum.portrait]: 'portrait',
  [OrientationEnum.landscape]: 'landscape',
}

export enum GlobalStyleEnum {
  default = 'default',
  hvl_default = 'hvl-default',
}

export enum AlignmentEnum {
  center = 0,
  both = 2,
  right = 1,
  left = 3,
}

export const Alignments: Record<AlignmentEnum, keyof typeof AlignmentEnum> = {
  [AlignmentEnum.center]: 'center',
  [AlignmentEnum.both]: 'both',
  [AlignmentEnum.right]: 'right',
  [AlignmentEnum.left]: 'left',
}

// =====================
// Özellikler
// =====================

export interface ITextPosition {
  startOffset: number
  length: number
}

export interface ITextStyle {
  family?: string
  size?: number
  foreground?: number
  underline?: boolean
  italic?: boolean
  bold?: boolean
}

export interface ITabAttributes extends ITextStyle, ITextPosition {

}

export interface IContentAttributes extends ITextPosition, ITextStyle {
  strikethrough: boolean
  height: number
  width: number
}

export interface IFieldAttributes extends IContentAttributes {
  fieldType: string
  fieldName: string
}

export interface IPageBreakAttributes {}

export interface IHeaderFooterAttributes extends ITextStyle {}

export interface IImageAttributes extends IContentAttributes {
  imageData: string
  height: number
  width: number
  size?: number
}

export interface IParagraphAttributes {
  resolver: string
  Alignment?: AlignmentEnum
  FirstLineIndent?: number
  LeftIndent?: number
  RightIndent?: number
  SpaceAbove?: number
  Hanging?: number
  Numbered?: boolean
  Bulleted?: boolean
  SpaceBelow?: number
  LineSpacing?: number
  TabSet?: string
  RepeatingLabel?: boolean
  GroupName?: string
  NumberType?: string
  ListLevel?: number
  ListId?: number
  SecListTypeLevel1?: string
}

export interface ITableAttributes {
  tableName: string
  columnCount: string
  border: string
  columnSpans: string
}

export interface IRowAttributes {
  height: number
  width: number
  rowType: string
  rowName: string
}

export interface ITemplateAttributes {
  format_id: string
}

export interface IPageFormatAttributes {
  mediaSizeName: string
  paperOrientation: OrientationEnum
  leftMargin: number
  rightMargin: number
  topMargin: number
  bottomMargin: number
}

export interface IGlobalStyle extends ITextStyle {
  name: GlobalStyleEnum
  description: string
}

export interface IContent {
  '#name': NodeEnum.Content
  '_': string
}

// =====================
// Şablon Tipleri
// =====================

export enum NodeEnum {
  Template = 'template', // bitti
  Content = 'content', // bitti
  Properties = 'properties', // bitti
  PageFormat = 'pageFormat', // bitti
  PageBreak = 'page-break', // bitti
  Elements = 'elements', // bitti
  Table = 'table',
  Row = 'row',
  Cell = 'cell',
  Paragraph = 'paragraph', // bitti
  Image = 'image', // TODO: sadece png calisiyor
  Field = 'field',
  Space = 'space',
  Tab = 'tab',
  Header = 'header', // bitti
  Footer = 'footer', // bitti
  Styles = 'styles', // bitti
  Style = 'style',
}

export interface Node<TName extends NodeEnum, TAttributes, TChildren = never> {
  '#name': TName
  '$': TAttributes
  '$$'?: TChildren[]
}

export type ParagraphElement
  = | Content
    | Field
    | Image
    | Space
    | Tab

export type Paragraph = Node<NodeEnum.Paragraph, IParagraphAttributes, ParagraphElement>
export type Content = Node<NodeEnum.Content, IContentAttributes>
export type Field = Node<NodeEnum.Field, IFieldAttributes>
export type Image = Node<NodeEnum.Image, IImageAttributes>
export type Space = Node<NodeEnum.Space, IContentAttributes>
export type Tab = Node<NodeEnum.Tab, ITabAttributes>

export type Cell = Node<NodeEnum.Cell, object, Paragraph>
export type Row = Node<NodeEnum.Row, IRowAttributes, Cell>
export type Table = Node<NodeEnum.Table, ITableAttributes, Row>

export type Header = Node<NodeEnum.Header, IHeaderFooterAttributes, Paragraph>
export type Footer = Node<NodeEnum.Footer, IHeaderFooterAttributes, Paragraph>

export type DocumentElement = Table | Paragraph | Header | Footer | PageBreak
export type Elements = Node<NodeEnum.Elements, object, DocumentElement>

export type PropertiesElement = PageFormat
export type Properties = Node<NodeEnum.Properties, object, PropertiesElement>

export type PageFormat = Node<NodeEnum.PageFormat, IPageFormatAttributes>
export type PageBreak = Node<NodeEnum.PageBreak, IPageBreakAttributes, DocumentElement>
export type Style = Node<NodeEnum.Style, IGlobalStyle>
export type Styles = Node<NodeEnum.Styles, object, Style>

export type TemplateElement = IContent | Properties | Elements | Styles
export type Template = Node<NodeEnum.Template, ITemplateAttributes, TemplateElement>

export interface IUdfDocument {
  template: Template
}

// =====================
// Sınıflar
// =====================

class Utils {
  public static ptToTwip(pt?: number | string): number {
    const parsed = typeof pt === 'string' ? Number.parseFloat(pt) : pt
    return !isNaN(parsed as number) ? (parsed as number) * 20 : 0
  }

  public static docxNumberToNormal(value: string): number {
    const output = Number.parseFloat(value.slice(0, -2))
    return !isNaN(output) ? output : 0
  }

  public static convertColor(value?: number): string {
    if (value == null || isNaN(value))
      return '#ffffff'

    let color = value
    if (color < 0)
      color = 0xFFFFFFFF + color + 1

    const r = (color >> 16) & 0xFF
    const g = (color >> 8) & 0xFF
    const b = color & 0xFF
    return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`
  }

  public static mergeObjects<T>(...styles: Partial<T>[]): T {
    return styles.reduce<Partial<T>>((acc, style) => {
      for (const key in style) {
        if (acc[key as keyof T] == null && style[key as keyof T] != null)
          acc[key as keyof T] = style[key as keyof T]
      }
      return acc
    }, {}) as T
  }

  public static textStyle(style: IContentAttributes): docx.IRunStylePropertiesOptions {
    const attrs = style ?? {}
    return {
      bold: attrs.bold,
      italics: attrs.italic,
      size: attrs.size ? `${attrs.size}pt` : undefined,
      underline: attrs.underline ? { type: 'dash', color: '#000000' } : undefined,
      color: attrs.foreground ? Utils.convertColor(attrs.foreground) : undefined,
      font: attrs.family,
      strike: attrs.strikethrough,
    }
  }

  public static namedStyle(style?: ITextStyle): docx.IRunStylePropertiesOptions {
    const attrs = style ?? {}
    return {
      bold: attrs.bold,
      italics: attrs.italic,
      size: attrs.size ? `${attrs.size}pt` : undefined,
      font: attrs.family,
    }
  }
}

class ToDocx {
  private referenceText: string = ''
  private styles: { [name: string]: IGlobalStyle } = {}

  private sectionProperties: docx.ISectionOptions['properties'] = undefined
  private sectionChildren: (docx.Paragraph | docx.Table)[] = []
  private sectionHeaders: docx.ISectionOptions['headers'] = undefined
  private sectionFooters: docx.ISectionOptions['footers'] = undefined

  public build(document: IUdfDocument): docx.Document {
    const { template } = document
    template.$$ ??= []

    this.parseTemplate(template)

    const elementsNode = template.$$.find(child => child['#name'] === NodeEnum.Elements)
    if (elementsNode == null)
      throw new Error('Document doesn\'t have elements')

    elementsNode.$$ ??= []

    for (const el of elementsNode.$$)
      this.sectionChildren.push(...this.parseDocumentElement(el))

    return new docx.Document({
      sections: [{
        properties: this.sectionProperties,
        headers: this.sectionHeaders,
        footers: this.sectionFooters,
        children: this.sectionChildren,
      }],
    })
  }

  private parseTemplate(template: Template): void {
    template.$$ ??= []

    for (const node of template.$$) {
      switch (node['#name']) {
        case NodeEnum.Content:
          this.referenceText = node._
          break

        case NodeEnum.Properties:
          this.parseProperties(node)
          break

        case NodeEnum.Styles:
          this.parseStyles(node)
          break
      }
    }
  }

  private parseProperties(node: Properties): void {
    if (node.$$ == null)
      return

    for (const prop of node.$$) {
      switch (prop['#name']) {
        case NodeEnum.PageFormat:
          this.parsePageFormat(prop)
          break
      }
    }
  }

  private parsePageFormat(prop: PageFormat): void {
    if (prop.$ == null)
      return

    const { paperOrientation, topMargin, rightMargin, bottomMargin, leftMargin } = prop.$

    this.sectionProperties = {
      page: {
        size: {
          orientation: Orientations[paperOrientation],
        },
        margin: {
          top: `${topMargin || Document.DEFAULT_MARGIN}pt`,
          right: `${rightMargin || Document.DEFAULT_MARGIN}pt`,
          bottom: `${bottomMargin || Document.DEFAULT_MARGIN}pt`,
          left: `${leftMargin || Document.DEFAULT_MARGIN}pt`,
        },
      },
    }
  }

  private parseStyles(node: Styles): void {
    if (node.$$ == null)
      return

    for (const style of node.$$) {
      if (style?.$.name)
        this.styles[style.$.name] = style.$
    }
  }

  private parseDocumentElement(element: DocumentElement): typeof this.sectionChildren {
    const sectionElements: typeof this.sectionChildren = []

    switch (element['#name']) {
      case NodeEnum.Paragraph:
        sectionElements.push(this.parseParagraph(element))
        break

      case NodeEnum.PageBreak:
        sectionElements.push(new docx.Paragraph({
          children: [
            new docx.PageBreak(),
            ...element.$$?.flatMap(elem1 => this.parseDocumentElement(elem1)) ?? [],
          ],
        }))
        break

      case NodeEnum.Table:
        sectionElements.push(this.parseTable(element))
        break

      case NodeEnum.Header: {
        const children = element.$$?.flatMap(elem1 => this.parseDocumentElement(elem1)) ?? []
        this.sectionHeaders = { default: new docx.Header({ children }) }
      } break

      case NodeEnum.Footer: {
        const children = element.$$?.flatMap(elem1 => this.parseDocumentElement(elem1)) ?? []
        this.sectionFooters = { default: new docx.Footer({ children }) }
      } break
    }

    return sectionElements
  }

  private parseContent(
    node: Content,
    paragraphStyle: docx.IRunStylePropertiesOptions,
  ): docx.TextRun {
    const startOffset = node.$?.startOffset
    const textLength = node.$?.length

    if (isNaN(startOffset) || isNaN(textLength) || textLength <= 0 || startOffset + textLength > this.referenceText.length)
      throw new Error('Paragraph content skipped due wrong offset/length')

    const txt = this.referenceText.substring(startOffset, startOffset + textLength)
    const runStyle = Utils.mergeObjects<docx.IRunStylePropertiesOptions>(
      Utils.textStyle(node.$),
      paragraphStyle,
    )

    console.log(startOffset, txt)
    return new docx.TextRun({ text: txt, ...runStyle })
  }

  private parseTable(element: Table): docx.Table {
    const colWidthsInTwips = element.$.columnSpans != null
      ? element.$.columnSpans.toString().split(',').map(num => Utils.ptToTwip(Number(num)))
      : [2000, 2000, 2000, 2000]

    // Determine total number of grid columns available in this table
    const totalGridColumns = colWidthsInTwips.length

    const rows = element.$$?.map((row) => {
      let currentGridIndex = 0
      const rawCells = row.$$ ?? []
      const isSingleCellRow = rawCells.length === 1

      const cells = rawCells.map((cell) => {
        const children = cell.$$?.flatMap(para =>
          this.parseDocumentElement(para),
        ).filter((el): el is docx.Paragraph => el instanceof docx.Paragraph) ?? []

        // FIX: If the row has only one cell, force it to take the full column span
        const cellSpan = isSingleCellRow
          ? totalGridColumns
          : (1)

        // Slice your layout configuration based on the calculated span
        const targetWidths = colWidthsInTwips.slice(currentGridIndex, currentGridIndex + cellSpan)
        const totalCellWidth = targetWidths.reduce((sum, w) => sum + w, 0)
        currentGridIndex += cellSpan

        return new docx.TableCell({
          children,
          width: {
            size: totalCellWidth,
            type: docx.WidthType.DXA,
          },
          columnSpan: cellSpan,
        })
      }) ?? []

      return new docx.TableRow({
        children: cells,
        height: {
          value: Utils.ptToTwip(row.$.height),
          rule: docx.HeightRule.AUTO,
        },
      })
    }) ?? []

    return new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      columnWidths: colWidthsInTwips,
      rows,
    })
  }

  private parseParagraphTabStops(element: Paragraph): docx.TabStopDefinition[] | undefined {
    const TAB_TYPES: Record<number, string> = {
      0: docx.TabStopType.LEFT,
      1: docx.TabStopType.CENTER,
      2: docx.TabStopType.RIGHT,
      3: docx.TabStopType.DECIMAL,
      4: docx.TabStopType.BAR,
    }

    const TAB_LEADERS: Record<number, string> = {
      0: docx.LeaderType.NONE,
      1: docx.LeaderType.DOT,
      2: docx.LeaderType.HYPHEN,
      3: docx.LeaderType.UNDERSCORE,
      4: docx.LeaderType.MIDDLE_DOT,
    }

    return element.$?.TabSet
      ?.toString()
      .split(',')
      .map(tab => tab.split(':').map(Number))
      .map(([position, type = 0, leader = 0]) => ({
        position: Utils.ptToTwip(position), // direkt pt → twip
        type: TAB_TYPES[type] ?? docx.TabStopType.LEFT,
        leader: TAB_LEADERS[leader] ?? 'none',
      } as docx.TabStopDefinition)) ?? []
  }

  private parseParagraph(node: Paragraph): docx.Paragraph {
    node.$$ ??= []

    const paragraphStyle = Utils.mergeObjects<docx.IRunStylePropertiesOptions>(
      Utils.namedStyle(this.styles[node.$?.resolver]),
      Utils.namedStyle(this.styles[GlobalStyleEnum.hvl_default]),
    )

    const baseParagraphSize: number = typeof paragraphStyle.size === 'string'
      ? Utils.docxNumberToNormal(paragraphStyle.size)
      : (paragraphStyle.size || 11)

    const targetAlignment = Alignments[node.$?.Alignment
      ? (node.$!.Alignment - 1) as AlignmentEnum
      : AlignmentEnum.left]

    const hanging = node.$?.Hanging ?? 0
    const leftIndent = node.$?.LeftIndent ?? 0
    const rightIndent = node.$?.RightIndent ?? 0
    const firstLineIndent = node.$?.FirstLineIndent ?? 0
    const spaceAbove = node.$?.SpaceAbove ?? 0
    const spaceBelow = node.$?.SpaceBelow ?? 0

    const docxParagraph = new docx.Paragraph({
      alignment: targetAlignment,
      spacing: {
        line: node.$?.LineSpacing != null
          ? Math.round(Utils.ptToTwip(baseParagraphSize * (node.$!.LineSpacing + 1.0)))
          : Math.round(Utils.ptToTwip(baseParagraphSize)),
        lineRule: 'auto',
        before: spaceAbove > 0 ? Utils.ptToTwip(spaceAbove) : undefined,
        after: spaceBelow > 0 ? Utils.ptToTwip(spaceBelow) : undefined,
      },
      indent: {
        hanging: hanging > 0 ? Utils.ptToTwip(hanging) : undefined,
        left: hanging > 0
          ? Utils.ptToTwip(leftIndent + hanging)
          : leftIndent > 0
            ? Utils.ptToTwip(leftIndent)
            : undefined,
        firstLine: !hanging && firstLineIndent !== 0
          ? Utils.ptToTwip(firstLineIndent)
          : undefined,
        right: rightIndent > 0 ? Utils.ptToTwip(rightIndent) : undefined,
      },
      tabStops: this.parseParagraphTabStops(node),
    })

    for (const child of node.$$) {
      switch (child['#name']) {
        case NodeEnum.Content:
          docxParagraph.addChildElement(this.parseContent(child, paragraphStyle))
          break

          // case NodeEnum.Space: // TODO: Wrong implemantation
          //   docxParagraph.addChildElement(new docx.TextRun({ text: String.prototype.padStart(child.$.length, ' '), ...paragraphStyle }))
          //   break

          // case NodeEnum.Tab: // TODO: Wrong implemantation
          //   docxParagraph.addChildElement(new docx.TextRun({ text: String.prototype.padStart(child.$.length, '\t'), ...paragraphStyle }))
          //   break

        case NodeEnum.Field:
          break

        case NodeEnum.Image:
          docxParagraph.addChildElement(new docx.ImageRun({
            data: Uint8Array.from(window.atob(child.$.imageData), c => c.charCodeAt(0)),
            type: 'png',
            transformation: {
              height: child.$.size ?? child.$.height,
              width: child.$.size ?? child.$.width,
            },
          }))
          break
      }
    }

    return docxParagraph
  }
}

export class Document {
  public readonly document: IUdfDocument
  public static readonly DEFAULT_MARGIN = 42.5

  constructor(xmlString: string) {
    this.document = Document.parseDocument(xmlString)
  }

  public static parseDocument(xmlString: string): IUdfDocument {
    const parser = new xml2js.Parser({
      preserveChildrenOrder: true,
      explicitArray: true,
      explicitChildren: true,
      attrValueProcessors: [
        xml2js.processors.parseNumbers,
        xml2js.processors.parseBooleans,
      ],
    })

    let parsed: IUdfDocument | undefined

    parser.parseString(xmlString, (err, result: IUdfDocument) => {
      if (err)
        throw new Error(`Failed to parse UDF XML: ${err.message}`)
      parsed = result
    })

    if (!parsed)
      throw new Error('XML parsing resulted in undefined document.')

    this.validateVersion(parsed)
    return parsed
  }

  public static validateVersion(document: IUdfDocument): void {
    try {
      const version = document.template?.$?.format_id?.toString()
      if (!version)
        throw new Error('Missing UDF format version.')

      if (!Object.values(Version).includes(version as Version))
        throw new Error(`Unsupported UDF Version: ${version}`)
    }
    catch (error) {
      throw new Error(`Failed during UDF version validation: ${(error as Error).message}`)
    }
  }

  public toDocx(): docx.Document {
    return new ToDocx().build(this.document)
  }
}
