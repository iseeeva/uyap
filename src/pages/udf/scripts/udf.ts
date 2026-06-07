import * as docx from 'docx'
import xml2js from 'xml2js'

export enum OrientationEnum {
  portrait = 1,
  landscape = 2,
}

export enum AlignmentEnum {
  center = 0,
  both = 2,
  right = 1,
  left = 3,
}

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

export interface IContentAttributes extends ITextPosition, ITextStyle {
  strikethrough: boolean
  height: number
  width: number
}

export interface IFieldAttributes extends IContentAttributes {
  fieldType: string
  fieldName: string
}

export interface IImageAttributes extends IContentAttributes {
  imageData: string
  size?: number
}

export interface ITabAttributes extends ITextStyle, ITextPosition {}

export interface IHeaderFooterAttributes extends ITextStyle {}

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
  ListLevel?: number
  ListId?: number
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

export interface IPageFormatAttributes {
  mediaSizeName: string
  paperOrientation: OrientationEnum
  leftMargin: number
  rightMargin: number
  topMargin: number
  bottomMargin: number
}

export interface IGlobalStyle extends ITextStyle {
  name: string
  description: string
}

export enum NodeEnum {
  Template = 'template',
  Content = 'content',
  Properties = 'properties',
  PageFormat = 'pageFormat',
  PageBreak = 'page-break',
  Elements = 'elements',
  Table = 'table',
  Row = 'row',
  Cell = 'cell',
  Paragraph = 'paragraph',
  Image = 'image',
  Field = 'field',
  Space = 'space',
  Tab = 'tab',
  Header = 'header',
  Footer = 'footer',
  Styles = 'styles',
  Style = 'style',
}

export interface Node<TName extends NodeEnum, TAttributes, TChildren = never> {
  '#name': TName
  '$'?: TAttributes
  '$$'?: TChildren[]
}

export type ParagraphElement = Content | Field | Image | Space | Tab

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
export type PageFormat = Node<NodeEnum.PageFormat, IPageFormatAttributes>
export type PageBreak = Node<NodeEnum.PageBreak, object, ParagraphElement>
export type Style = Node<NodeEnum.Style, IGlobalStyle>
export type Styles = Node<NodeEnum.Styles, object, Style>
export type Properties = Node<NodeEnum.Properties, object, PageFormat>
export type Elements = Node<NodeEnum.Elements, object, DocumentElement>

export type DocumentElement = Table | Paragraph | PageBreak
export type TopLevelElement = DocumentElement | Header | Footer

export interface IContentNode {
  '#name': NodeEnum.Content
  '_': string
}

export type TemplateElement = Properties | Elements | Styles | IContentNode
export type Template = Node<NodeEnum.Template, { format_id: string }, TemplateElement>

export interface IUdfDocument {
  template: Template
}

// #region Utils

function convertColor(value?: number): string {
  if (value == null || isNaN(value))
    return '#ffffff'
  const color = value < 0 ? 0xFFFFFFFF + value + 1 : value
  const r = (color >> 16) & 0xFF
  const g = (color >> 8) & 0xFF
  const b = color & 0xFF
  return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`
}

function ptToTwip(pt: number | string | undefined): number {
  const parsed = typeof pt === 'string' ? parseFloat(pt) : pt
  return parsed != null && !isNaN(parsed) ? parsed * 20 : 0
}

function docxSizeToNumber(value: string): number {
  const output = parseFloat(value.slice(0, -2))
  return !isNaN(output) ? output : 0
}

function mergeStyles<T extends object>(...styles: Partial<T>[]): T {
  const result = {} as T
  for (const style of styles) {
    for (const _key in style) {
      const key = _key as keyof T
      if (result[key] == null && style[key] != null)
        result[key] = style[key] as T[keyof T]
    }
  }
  return result
}

function isContentNode(child: TemplateElement): child is IContentNode {
  return child['#name'] === NodeEnum.Content
}

// #endregion

// #region Style

function runStyleFromContent(attrs: IContentAttributes): docx.IRunStylePropertiesOptions {
  return {
    bold: attrs.bold,
    italics: attrs.italic,
    size: attrs.size ? `${attrs.size}pt` : undefined,
    underline: attrs.underline ? { type: 'dash', color: '#000000' } : undefined,
    color: attrs.foreground ? convertColor(attrs.foreground) : undefined,
    font: attrs.family,
    strike: attrs.strikethrough,
  }
}

function runStyleFromTextStyle(style?: ITextStyle): docx.IRunStylePropertiesOptions {
  return {
    bold: style?.bold,
    italics: style?.italic,
    size: style?.size ? `${style.size}pt` : undefined,
    font: style?.family,
  }
}

class StyleResolver {
  private readonly registry: Record<string, IGlobalStyle>

  constructor(styles: Record<string, IGlobalStyle>) {
    this.registry = styles
  }

  public static fromStylesNode(node: Styles): StyleResolver {
    const registry: Record<string, IGlobalStyle> = {}
    for (const style of node.$$ ?? []) {
      const attrs = style?.$
      if (attrs?.name)
        registry[attrs.name] = attrs
    }
    return new StyleResolver(registry)
  }

  public resolveForParagraph(resolver?: string): docx.IRunStylePropertiesOptions {
    return mergeStyles<docx.IRunStylePropertiesOptions>(
      resolver ? runStyleFromTextStyle(this.registry[resolver]) : {},
      runStyleFromTextStyle(this.registry['hvl-default']),
    )
  }
}

// #endregion

// #region TabStop

const TabTypes: Record<number, string> = {
  0: docx.TabStopType.LEFT,
  1: docx.TabStopType.CENTER,
  2: docx.TabStopType.RIGHT,
  3: docx.TabStopType.DECIMAL,
  4: docx.TabStopType.BAR,
}

const TabLeaders: Record<number, string> = {
  0: docx.LeaderType.NONE,
  1: docx.LeaderType.DOT,
  2: docx.LeaderType.HYPHEN,
  3: docx.LeaderType.UNDERSCORE,
  4: docx.LeaderType.MIDDLE_DOT,
}

function parseTabStops(tabSet?: string): docx.TabStopDefinition[] {
  if (!tabSet)
    return []
  return tabSet
    .toString()
    .split(',')
    .map(tab => tab.split(':').map(Number))
    .map(([position, type = 0, leader = 0]): docx.TabStopDefinition => ({
      position: ptToTwip(position),
      type: (TabTypes[type] ?? docx.TabStopType.LEFT) as any,
      leader: (TabLeaders[leader] ?? docx.LeaderType.NONE) as any,
    }))
}

// #endregion

// #region Alignment

const AlignmentMap: Record<AlignmentEnum, string> = {
  [AlignmentEnum.center]: docx.AlignmentType.CENTER,
  [AlignmentEnum.both]: docx.AlignmentType.BOTH,
  [AlignmentEnum.right]: docx.AlignmentType.RIGHT,
  [AlignmentEnum.left]: docx.AlignmentType.LEFT,
}

function resolveAlignment(raw?: AlignmentEnum): string {
  if (raw == null)
    return docx.AlignmentType.LEFT

  const key = (raw - 1) as AlignmentEnum
  return AlignmentMap[key] ?? docx.AlignmentType.LEFT
}

// #endregion

// #region ElementParser

class ElementParser {
  constructor(
    private readonly referenceText: string,
    private readonly styleResolver: StyleResolver,
  ) {}

  public parseDocumentElement(element: DocumentElement): docx.Paragraph | docx.Table {
    switch (element['#name']) {
      case NodeEnum.Paragraph:
        return this.parseParagraph(element)

      case NodeEnum.Table:
        return this.parseTable(element)

      case NodeEnum.PageBreak:
        return this.parsePageBreak(element)
    }
  }

  public parseParagraphsFrom(elements: Paragraph[]): docx.Paragraph[] {
    return elements.map(el => this.parseParagraph(el))
  }

  private parsePageBreak(element: PageBreak): docx.Paragraph {
    const children: docx.ParagraphChild[] = [new docx.PageBreak()]
    for (const child of element.$$ ?? []) {
      const run = this.parseParagraphChild(child as ParagraphElement, {})
      if (run)
        children.push(run)
    }
    return new docx.Paragraph({ children })
  }

  private parseParagraph(node: Paragraph): docx.Paragraph {
    node.$$ ??= []

    const attrs = node.$
    const paragraphStyle = this.styleResolver.resolveForParagraph(attrs?.resolver)

    const baseSize: number = typeof paragraphStyle.size === 'string'
      ? docxSizeToNumber(paragraphStyle.size)
      : (paragraphStyle.size ?? 11)

    const hanging = attrs?.Hanging ?? 0
    const leftIndent = attrs?.LeftIndent ?? 0
    const rightIndent = attrs?.RightIndent ?? 0
    const firstLine = attrs?.FirstLineIndent ?? 0
    const spaceAbove = attrs?.SpaceAbove ?? 0
    const spaceBelow = attrs?.SpaceBelow ?? 0
    const lineSpacing = attrs?.LineSpacing

    const paragraph = new docx.Paragraph({
      alignment: resolveAlignment(attrs?.Alignment) as any,
      spacing: {
        line: lineSpacing != null
          ? Math.round(ptToTwip(baseSize * (lineSpacing + 1.0)))
          : Math.round(ptToTwip(baseSize)),
        lineRule: docx.LineRuleType.AUTO,
        before: spaceAbove > 0 ? ptToTwip(spaceAbove) : undefined,
        after: spaceBelow > 0 ? ptToTwip(spaceBelow) : undefined,
      },
      indent: {
        hanging: hanging > 0 ? ptToTwip(hanging) : undefined,
        left: hanging > 0
          ? ptToTwip(leftIndent + hanging)
          : leftIndent > 0 ? ptToTwip(leftIndent) : undefined,
        firstLine: !hanging && firstLine !== 0 ? ptToTwip(firstLine) : undefined,
        right: rightIndent > 0 ? ptToTwip(rightIndent) : undefined,
      },
      tabStops: parseTabStops(attrs?.TabSet),
    })

    for (const child of node.$$) {
      const run = this.parseParagraphChild(child, paragraphStyle)
      if (run)
        paragraph.addChildElement(run)
    }

    return paragraph
  }

  private parseParagraphChild(
    child: ParagraphElement,
    paragraphStyle: docx.IRunStylePropertiesOptions,
  ): docx.TextRun | docx.ImageRun | undefined {
    switch (child['#name']) {
      case NodeEnum.Content:
        return this.parseContent(child, paragraphStyle)

      case NodeEnum.Field: // TODO: implement that types
      case NodeEnum.Space: // TODO: implement that types
      case NodeEnum.Tab: // TODO: implement that types
        break

      case NodeEnum.Image: {
        const attrs = child.$
        if (!attrs)
          return undefined

        return new docx.ImageRun({
          data: Uint8Array.from(window.atob(attrs.imageData), c => c.charCodeAt(0)),
          type: 'png',
          transformation: {
            height: attrs.size ?? attrs.height,
            width: attrs.size ?? attrs.width,
          },
        })
      }

      default:
        return undefined
    }
  }

  private parseContent(
    node: Content,
    paragraphStyle: docx.IRunStylePropertiesOptions,
  ): docx.TextRun | undefined {
    const attrs = node.$
    if (!attrs)
      return undefined

    const { startOffset, length } = attrs

    if (isNaN(startOffset) || isNaN(length) || length <= 0 || startOffset + length > this.referenceText.length)
      throw new Error('Paragraph content skipped due to wrong offset/length')

    const text = this.referenceText.substring(startOffset, startOffset + length)
    const runStyle = mergeStyles<docx.IRunStylePropertiesOptions>(
      runStyleFromContent(attrs),
      paragraphStyle,
    )

    return new docx.TextRun({ text, ...runStyle })
  }

  private parseTable(element: Table): docx.Table {
    const colWidths = element.$?.columnSpans != null
      ? element.$.columnSpans.toString().split(',').map(n => ptToTwip(Number(n)))
      : [2000, 2000, 2000, 2000]

    return new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      columnWidths: colWidths,
      rows: (element.$$ ?? []).map(row => this.parseRow(row, colWidths)),
    })
  }

  private parseRow(row: Row, colWidths: number[]): docx.TableRow {
    const rawCells = row.$$ ?? []
    const isSingleCell = rawCells.length === 1
    let gridIndex = 0

    const cells = rawCells.map((cell) => {
      const children = this.parseParagraphsFrom(cell.$$ ?? [])
      const span = isSingleCell ? colWidths.length : 1
      const width = colWidths.slice(gridIndex, gridIndex + span).reduce((a, b) => a + b, 0)
      gridIndex += span

      return new docx.TableCell({
        children,
        width: { size: width, type: docx.WidthType.DXA },
        columnSpan: span,
      })
    })

    return new docx.TableRow({
      children: cells,
      height: { value: ptToTwip(row.$?.height), rule: docx.HeightRule.AUTO }, // TODO: if height doesn't exist.
    })
  }
}

// #endregion

// #region SectionBuilder

class SectionBuilder {
  public properties: docx.ISectionOptions['properties'] = undefined
  public headers: docx.ISectionOptions['headers'] = undefined
  public footers: docx.ISectionOptions['footers'] = undefined
  public readonly children: (docx.Paragraph | docx.Table)[] = []

  public setPageFormat(prop: PageFormat): void {
    if (prop.$ == null)
      return

    const { paperOrientation, topMargin, rightMargin, bottomMargin, leftMargin } = prop.$
    this.properties = {
      page: {
        size: {
          orientation: paperOrientation === OrientationEnum.landscape
            ? docx.PageOrientation.LANDSCAPE
            : docx.PageOrientation.PORTRAIT,
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

  public toSection(): docx.ISectionOptions {
    return {
      properties: this.properties,
      headers: this.headers,
      footers: this.footers,
      children: this.children,
    }
  }
}

// #endregion

// #region DocumentConverter

class DocumentConverter {
  public convert(udfDocument: IUdfDocument): docx.Document {
    const { template } = udfDocument
    template.$$ ??= []

    const referenceText = this.extractReferenceText(template)
    const styleResolver = this.extractStyleResolver(template)
    const section = new SectionBuilder()

    this.applyProperties(template, section)

    const parser = new ElementParser(referenceText, styleResolver)

    const elementsNode = template.$$.find(
      (child): child is Elements => child['#name'] === NodeEnum.Elements,
    )
    if (!elementsNode)
      throw new Error('Document doesn\'t have elements')

    for (const el of elementsNode.$$ ?? []) {
      this.dispatchTopLevelElement(el, section, parser)
    }

    return new docx.Document({ sections: [section.toSection()] })
  }

  private extractReferenceText(template: Template): string {
    const contentNode = template.$$?.find(isContentNode)
    return contentNode?._ ?? ''
  }

  private extractStyleResolver(template: Template): StyleResolver {
    const stylesNode = template.$$?.find(
      (child): child is Styles => child['#name'] === NodeEnum.Styles,
    )
    return stylesNode ? StyleResolver.fromStylesNode(stylesNode) : new StyleResolver({})
  }

  private applyProperties(template: Template, section: SectionBuilder): void {
    const propertiesNode = template.$$?.find(
      (child): child is Properties => child['#name'] === NodeEnum.Properties,
    )
    if (!propertiesNode)
      return

    for (const prop of propertiesNode.$$ ?? []) {
      if (prop['#name'] === NodeEnum.PageFormat)
        section.setPageFormat(prop)
    }
  }

  private dispatchTopLevelElement(
    element: TopLevelElement,
    section: SectionBuilder,
    parser: ElementParser,
  ): void {
    switch (element['#name']) {
      case NodeEnum.Header:
        section.headers = { default: new docx.Header({ children: parser.parseParagraphsFrom(element.$$ ?? []) }) }
        break

      case NodeEnum.Footer:
        section.footers = { default: new docx.Footer({ children: parser.parseParagraphsFrom(element.$$ ?? []) }) }
        break

      default:
        section.children.push(parser.parseDocumentElement(element))
    }
  }
}

// #endregion

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

    return parsed
  }

  public toDocx(): docx.Document {
    return new DocumentConverter().convert(this.document)
  }
}
