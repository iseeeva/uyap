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
  Portrait = 1,
  Landscape = 2,
}

export const Orientations: Record<OrientationEnum, 'portrait' | 'landscape' | undefined> = {
  [OrientationEnum.Portrait]: 'portrait',
  [OrientationEnum.Landscape]: 'landscape',
}

export enum GlobalStyleEnum {
  default = 'default',
  hvl_default = 'hvl-default',
}

export enum AlignmentEnum {
  center = 0,
  right = 1,
  both = 2,
  left = 3,
}

export const Alignments: Record<AlignmentEnum, 'center' | 'right' | 'both' | 'left' | undefined> = {
  [AlignmentEnum.center]: 'center',
  [AlignmentEnum.right]: 'right',
  [AlignmentEnum.both]: 'both',
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

export interface IContentAttributes extends ITextPosition, ITextStyle {
  strikethrough: boolean
}

export interface IFieldAttributes extends IContentAttributes {
  fieldType: string
  fieldName: string
}

export interface IPageBreakAttributes {

}

export interface IHeaderFooterAttributes extends ITextStyle {

}

export interface IImageAttributes extends IContentAttributes {
  imageData: string
  height: number
  width: number
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
}

export interface ITableAttributes {
  tableName: string
  columnCount: string
  border: string
  columnSpans: string
}

export interface IRowAttributes {
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
export type Tab = Node<NodeEnum.Tab, ITextStyle>

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

  public static docxNumberToNormal(value: string) {
    const output = Number.parseFloat(value.slice(0, -2))
    return !isNaN(output) ? output : 0
  }

  public static convertColor(value?: number): string {
    if (value == null || isNaN(value)) {
      return '#ffffff'
    }
    let color = value
    if (color < 0) {
      color = 0xFFFFFFFF + color + 1
    }
    const r = (color >> 16) & 0xFF
    const g = (color >> 8) & 0xFF
    const b = color & 0xFF
    return `#${r.toString(16).padStart(2, '0')}${g
      .toString(16)
      .padStart(2, '0')}${b.toString(16).padStart(2, '0')}`
  }

  public static mergeObjects<T>(...styles: Partial<T>[]): T {
    return styles.reduce<Partial<T>>((acc, style) => {
      for (const key in style) {
        if (acc[key as keyof T] == null && style[key as keyof T] != null) {
          acc[key as keyof T] = style[key as keyof T]
        }
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
  public static parseDocument(document: IUdfDocument): docx.Document {
    const { template } = document
    template.$$ ??= []

    const paragraphs: Array<docx.Paragraph> = []
    const sections: docx.ISectionPropertiesOptions[] = []
    const styles: { [name: string]: IGlobalStyle } = {}
    let referenceText: string = ''

    for (const test of template.$$) {
      switch (test['#name']) {
        // Reference content
        case NodeEnum.Content: {
          const content = test
          referenceText = content._
        } break

          // Page properties
        case NodeEnum.Properties: {
          const propertiesNode = test

          if (propertiesNode.$$ == null) {
            continue
          }

          for (const test of propertiesNode.$$) {
            switch (test['#name']) {
              case NodeEnum.PageFormat: {
                const pageFmt = test

                if (pageFmt.$ == null) {
                  continue
                }

                const { paperOrientation, topMargin, rightMargin, bottomMargin, leftMargin } = pageFmt.$
                sections.push({
                  page: {
                    size: { orientation: Orientations[paperOrientation] },
                    margin: {
                      top: `${(topMargin || Document.DEFAULT_MARGIN)}pt`,
                      right: `${(rightMargin || Document.DEFAULT_MARGIN)}pt`,
                      bottom: `${(bottomMargin || Document.DEFAULT_MARGIN)}pt`,
                      left: `${(leftMargin || Document.DEFAULT_MARGIN)}pt`,
                    },
                  },
                })

                console.log('Page format configured.')
              } break
            }
          }

          console.log('Page properties configured.')
        } break

          // Pre-defined styles
        case NodeEnum.Styles: {
          const stylesNode = test

          if (stylesNode.$$ == null) {
            continue
          }

          stylesNode.$$?.forEach((group) => {
            if (group?.$.name) {
              styles[group.$.name] = group.$
              console.log(`Loaded style: ${group.$.name}`)
            }
          })
        } break
      }
    }

    // Elements
    const elementsNode = template.$$.find(child => child['#name'] === NodeEnum.Elements)
    if (elementsNode == null)
      throw new Error('Document doesn\'t have elements')

    elementsNode.$$ ??= []

    for (const test of elementsNode.$$)
      paragraphs.push(...this.parseDocumentElement(test, referenceText, styles))

    // elementsNode.$$.forEach((el, eIdx) => {
    //   const elementStyle = UDFUtils.mergeStyles<docx.IRunStylePropertiesOptions>(
    //     UDFUtils.namedStyle(styles[el.$?.resolver ?? '']),
    //     UDFUtils.namedStyle(styles[StyleDictionaryEnum.default]),
    //   )

    //   // Paragraphs
    //   el.paragraph?.forEach((p, pIdx) => {
    //     const para = new docx.Paragraph({
    //       // tabStops: [
    //       //   { position: 3000, type: "", leader: 'none' },
    //       // ],
    //       alignment: p.$?.Alignment ? AlignmentKeys[Number(p.$.Alignment) - 1] : undefined,
    //       spacing: {
    //         line: p.$?.LineSpacing && elementStyle.size
    //           ? Math.round(UDFUtils.ptToTwip(UDFUtils.docxNumberToNormal(elementStyle.size) * (Number.parseFloat(p.$?.LineSpacing) + 1.0)))
    //           : Math.round(UDFUtils.ptToTwip(UDFUtils.docxNumberToNormal(elementStyle.size) * 1.0)), // default line height * multiplier
    //         lineRule: 'auto',
    //       },
    //       indent: {
    //         firstLine: p.$?.FirstLineIndent ? `${Number.parseFloat(p.$.FirstLineIndent) / 28.346457}cm` : undefined,
    //         left: p.$?.LeftIndent ? `${Number.parseFloat(p.$.LeftIndent) / 28.346457}cm` : undefined,
    //         right: p.$?.RightIndent ? `${Number.parseFloat(p.$.RightIndent) / 28.346457}cm` : undefined,
    //       },
    //     })

    //     p.content?.forEach((c) => {
    //       const start = Number(c.$?.startOffset)
    //       const length = Number(c.$?.length)
    //       if (isNaN(start) || isNaN(length) || length <= 0)
    //         return
    //       const txt = textRef.substring(start, start + length)

    //       const runStyle = UDFUtils.mergeStyles<docx.IRunStylePropertiesOptions>(
    //         UDFUtils.textStyle(c),
    //         UDFUtils.namedStyle(styles[p.$?.resolver ?? '']),
    //         elementStyle,
    //       )

    //       para.addChildElement(new docx.TextRun({ text: txt, ...runStyle }))
    //       this.log(`${start} ${p.$?.Alignment ? AlignmentKeys[Number(p.$.Alignment) - 1] : undefined} ${txt}`)
    //     })

    //     paragraphs.push(para)
    //     this.log(`Parsed paragraph ${eIdx}.${pIdx}`)
    //   })

    //   // Tables
    //   el.table?.forEach((tbl, tIdx) => {
    //     const rows: docx.TableRow[] = []
    //     tbl.row?.forEach((r, rIdx) => {
    //       const cells = r.cell?.map((cell) => {
    //         const cellParas = cell.paragraph?.map((cp) => {
    //           const childRuns = cp.content?.map((cnt) => {
    //             const start = Number(cnt.$?.startOffset)
    //             const length = Number(cnt.$?.length)
    //             const text = textRef.substring(start, start + length)
    //             const style = UDFUtils.mergeStyles<docx.IRunStylePropertiesOptions>(
    //               UDFUtils.textStyle(cnt),
    //               UDFUtils.namedStyle(styles[cp.$?.resolver ?? '']),
    //               UDFUtils.namedStyle(styles[StyleDictionaryEnum.default]),
    //             )

    //             return new docx.TextRun({ text, ...style })
    //           }) ?? []

    //           return new docx.Paragraph({ children: childRuns })
    //         }) ?? []

    //         return new docx.TableCell({ children: cellParas })
    //       }) ?? []

    //       rows.push(new docx.TableRow({ children: cells }))
    //     })

    //     paragraphs.push(new docx.Table({ rows }))
    //     this.log(`Parsed table ${eIdx}.${tIdx}`)
    //   })
    // })

    const doc = new docx.Document({
      sections: [{ properties: sections[0], children: paragraphs }],
    })

    return doc
  }

  public static parseDocumentElement(
    element: DocumentElement,
    referenceText: string,
    styles: { [name: string]: IGlobalStyle },
  ): docx.Paragraph[] {
    const paragraphs: docx.Paragraph[] = []

    switch (element['#name']) {
      case NodeEnum.Paragraph:
        paragraphs.push(this.parseParagraph(element, referenceText, styles))
        break

      case NodeEnum.PageBreak:
        paragraphs.push(new docx.Paragraph({ children: [new docx.PageBreak(), ...element.$$?.map(elem1 => this.parseDocumentElement(elem1, referenceText, styles)).flat() || []] }))
        break

      case NodeEnum.Table:
        break

      case NodeEnum.Footer:
        paragraphs.push(...element.$$?.map(elem1 => this.parseDocumentElement(elem1, referenceText, styles)).flat() || [])
        break
    }

    return paragraphs
  }

  public static parseContent(
    node: Content,
    referenceText: string,
    paragraphStyle: docx.IRunStylePropertiesOptions,
  ): docx.TextRun {
    const content = node
    const start = Number(content.$?.startOffset)
    const length = Number(content.$?.length)

    if (isNaN(start) || isNaN(length) || length <= 0 || start + length > referenceText.length) {
      throw new Error('Paragraph content skipped due wrong offset/length')
    }

    const txt = referenceText.substring(start, start + length)
    const runStyle = Utils.mergeObjects<docx.IRunStylePropertiesOptions>(
      Utils.textStyle(content.$),
      paragraphStyle,
    )

    return new docx.TextRun({ text: txt, ...runStyle, strike: content.$.strikethrough })
  }

  public static parseParagraph(
    node: Paragraph,
    referenceText: string,
    styles: { [name: string]: IGlobalStyle },
  ): docx.Paragraph {
    const paragraph = node
    paragraph.$$ ??= []

    const paragraphStyle = Utils.mergeObjects<docx.IRunStylePropertiesOptions>(
      Utils.namedStyle(styles[paragraph.$?.resolver]),
      Utils.namedStyle(styles[GlobalStyleEnum.hvl_default]),
    )

    const baseSize: number = typeof paragraphStyle.size === 'string'
      ? Utils.docxNumberToNormal(paragraphStyle.size)
      : (paragraphStyle.size || 11)

    const docxParagraph = new docx.Paragraph({
      // tabStops: [
      //   { position: 3000, type: "", leader: 'none' },
      // ],
      alignment: Alignments[paragraph.$?.Alignment ? (paragraph.$!.Alignment - 1) as AlignmentEnum : AlignmentEnum.left],
      spacing: {
        line: paragraph.$?.LineSpacing != null
          ? Math.round(Utils.ptToTwip(baseSize * (paragraph.$!.LineSpacing + 1.0)))
          : Math.round(Utils.ptToTwip(baseSize)), // default line height * multiplier
        lineRule: 'auto',
      },
      indent: {
        firstLine: paragraph.$?.FirstLineIndent ? `${paragraph.$.FirstLineIndent}pt` : undefined,
        left: paragraph.$?.LeftIndent ? `${paragraph.$.LeftIndent}pt` : undefined,
        right: paragraph.$?.RightIndent ? `${paragraph.$.RightIndent}pt` : undefined,
      },
    })

    for (const test of paragraph.$$) {
      switch (test['#name']) {
        case NodeEnum.Content:
          docxParagraph.addChildElement(this.parseContent(test, referenceText, paragraphStyle))
          break

        case NodeEnum.Tab:
        case NodeEnum.Space:
        case NodeEnum.Field:
          break

        case NodeEnum.Image:
          {
            const imageOptions: docx.IImageOptions = {
              data: Uint8Array.from(window.atob(test.$.imageData), c => c.charCodeAt(0)),
              type: 'png', // TODO: type check needs
              transformation: {
                height: test.$.size ? test.$.size : test.$.height,
                width: test.$.size ? test.$.size : test.$.width,
              },
            }

            const imageObject = new docx.ImageRun(imageOptions)
            docxParagraph.addChildElement(imageObject)
          } break
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
      if (err) {
        throw new Error(`Failed to parse UDF XML: ${err.message}`)
      }
      parsed = result
    })

    if (!parsed) {
      throw new Error('XML parsing resulted in undefined document.')
    }

    this.validateVersion(parsed)
    return parsed
  }

  public static validateVersion(document: IUdfDocument): void {
    try {
      const version = document.template?.$?.format_id?.toString()
      if (!version) {
        throw new Error('Missing UDF format version.')
      }

      if (!Object.values(Version).includes(version as Version)) {
        throw new Error(`Unsupported UDF Version: ${version}`)
      }
    }
    catch (error) {
      throw new Error(`Failed during UDF version validation: ${(error as Error).message}`)
    }
  }

  public toDocx() {
    return ToDocx.parseDocument(this.document)
  }
}
