import type { INumberingOptions, IPropertiesOptions } from "docx"
import {
  AlignmentType,
  BorderStyle,
  ImportedXmlComponent,
  LevelFormat,
  ShadingType,
  convertInchesToTwip,
} from "docx"
import { listIndent } from "./list"

interface FontSizes {
  normal: number
  heading1: number
  heading2: number
  heading3: number
  heading456: number
  codeBlock: number
  footerText: number
}

function fontSizes(basePt: number): FontSizes {
  const hp = (pt: number) => pt * 2
  return {
    normal: hp(basePt),
    heading1: hp(basePt + 14),
    heading2: hp(basePt + 7),
    heading3: hp(basePt + 4),
    heading456: hp(basePt + 2),
    codeBlock: hp(basePt),
    footerText: hp(10),
  }
}

export function buildStyleOptions(
  basePt: number,
  externalStylesXml?: string,
): Pick<IPropertiesOptions, "externalStyles" | "styles"> {
  if (externalStylesXml) {
    return { externalStyles: externalStylesXml }
  }

  const sizes = fontSizes(basePt)

  return {
    styles: {
      paragraphStyles: [
        {
          id: "Normal",
          name: "Normal",
          run: { font: "Aptos", size: sizes.normal, color: "333333" },
          paragraph: { spacing: { after: 160, line: 276 } },
        },
        {
          id: "Heading1",
          name: "Heading 1",
          basedOn: "Normal",
          next: "Normal",
          run: { size: sizes.heading1, bold: true },
          paragraph: { spacing: { after: 160 } },
        },
        {
          id: "Heading2",
          name: "Heading 2",
          basedOn: "Normal",
          next: "Normal",
          run: { size: sizes.heading2, bold: true },
          paragraph: { spacing: { before: 280, after: 160 } },
        },
        {
          id: "Heading3",
          name: "Heading 3",
          basedOn: "Normal",
          next: "Normal",
          run: { size: sizes.heading3, bold: true },
          paragraph: { spacing: { before: 240, after: 160 } },
        },
        {
          id: "Heading4",
          name: "Heading 4",
          basedOn: "Normal",
          next: "Normal",
          run: { size: sizes.heading456, bold: true },
          paragraph: { spacing: { before: 200, after: 160 } },
        },
        {
          id: "Heading5",
          name: "Heading 5",
          basedOn: "Normal",
          next: "Normal",
          run: { size: sizes.heading456, bold: true, italics: true },
          paragraph: { spacing: { before: 200, after: 160 } },
        },
        {
          id: "Heading6",
          name: "Heading 6",
          basedOn: "Normal",
          next: "Normal",
          run: { size: sizes.heading456, underline: {} },
          paragraph: { spacing: { before: 200, after: 160 } },
        },
        {
          id: "ListItem",
          name: "List Item",
          basedOn: "Normal",
          next: "ListItem",
          paragraph: { spacing: { before: 60, after: 160, line: 276 }, contextualSpacing: true },
        },
        {
          id: "Blockquote",
          name: "Blockquote",
          basedOn: "Normal",
          next: "Normal",
          run: { italics: true, color: "777777" },
          paragraph: {
            spacing: { before: 0, after: 0, line: 276, lineRule: "auto" as const },
            indent: { left: convertInchesToTwip(0.15) },
            border: {
              left: { style: BorderStyle.SINGLE, size: 24, color: "A6A6A6", space: 6 },
            },
          },
        },
        {
          id: "CodeBlock",
          name: "Code Block",
          basedOn: "Normal",
          next: "Normal",
          run: { font: "Consolas", size: sizes.codeBlock, color: "555555" },
          paragraph: { spacing: { before: 0, after: 0, line: 240 } },
        },
        {
          id: "FooterText",
          name: "Footer Text",
          basedOn: "Normal",
          paragraph: { spacing: { before: 0, after: 0 } },
          run: { size: sizes.footerText, color: "888888" },
        },
      ],
      characterStyles: [
        {
          id: "lineNumber",
          name: "Line Number",
          run: { color: "888888" },
        },
        {
          id: "InlineCode",
          name: "Inline Code",
          run: {
            font: "Consolas",
            shading: { type: ShadingType.CLEAR, color: "F2F2F2", fill: "F2F2F2" },
          },
        },
        {
          id: "Hyperlink",
          name: "Hyperlink",
          run: { color: "0563C1", underline: {} },
        },
      ],
      importedStyles: [
        // fromXmlString returns a wrapper root; extract the inner w:style element
        (
          ImportedXmlComponent.fromXmlString(
            `<w:style w:type="table" w:styleId="TableGridLight">
            <w:name w:val="Grid Table Light"/>
            <w:basedOn w:val="TableNormal"/>
            <w:uiPriority w:val="40"/>
            <w:tblPr>
              <w:tblBorders>
                <w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF" w:themeColor="background1" w:themeShade="BF"/>
                <w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF" w:themeColor="background1" w:themeShade="BF"/>
                <w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF" w:themeColor="background1" w:themeShade="BF"/>
                <w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF" w:themeColor="background1" w:themeShade="BF"/>
                <w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF" w:themeColor="background1" w:themeShade="BF"/>
                <w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF" w:themeColor="background1" w:themeShade="BF"/>
              </w:tblBorders>
              <w:tblCellMar>
                <w:top w:w="0" w:type="dxa"/>
                <w:left w:w="108" w:type="dxa"/>
                <w:bottom w:w="0" w:type="dxa"/>
                <w:right w:w="108" w:type="dxa"/>
              </w:tblCellMar>
            </w:tblPr>
          </w:style>`,
          ) as unknown as { root: ImportedXmlComponent[] }
        ).root[0],
      ],
    },
  }
}

export function buildNumbering(orderedRefs: string[] = []): INumberingOptions {
  const ll = [0, 1, 2, 3, 4, 5].map((level) => listIndent(level)) as [
    ReturnType<typeof listIndent>,
    ReturnType<typeof listIndent>,
    ReturnType<typeof listIndent>,
    ReturnType<typeof listIndent>,
    ReturnType<typeof listIndent>,
    ReturnType<typeof listIndent>,
  ]

  const orderedLevels = [
    {
      level: 0,
      format: LevelFormat.DECIMAL,
      text: "%1.",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: ll[0] } },
    },
    {
      level: 1,
      format: LevelFormat.LOWER_LETTER,
      text: "%2.",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: ll[1] } },
    },
    {
      level: 2,
      format: LevelFormat.LOWER_ROMAN,
      text: "%3.",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: ll[2] } },
    },
    {
      level: 3,
      format: LevelFormat.DECIMAL,
      text: "%4.",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: ll[3] } },
    },
    {
      level: 4,
      format: LevelFormat.LOWER_LETTER,
      text: "%5.",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: ll[4] } },
    },
    {
      level: 5,
      format: LevelFormat.LOWER_ROMAN,
      text: "%6.",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: ll[5] } },
    },
  ]

  return {
    config: [
      ...orderedRefs.map((reference) => ({ reference, levels: orderedLevels })),
      {
        reference: "bullet-numbering",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "\u2022",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: ll[0] }, run: { font: "Aptos" } },
          },
          {
            level: 1,
            format: LevelFormat.BULLET,
            text: "\u25e6",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: ll[1] }, run: { font: "Aptos" } },
          },
          {
            level: 2,
            format: LevelFormat.BULLET,
            text: "\u25aa",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: ll[2] }, run: { font: "Aptos" } },
          },
          {
            level: 3,
            format: LevelFormat.BULLET,
            text: "\u2022",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: ll[3] }, run: { font: "Aptos" } },
          },
          {
            level: 4,
            format: LevelFormat.BULLET,
            text: "\u25e6",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: ll[4] }, run: { font: "Aptos" } },
          },
          {
            level: 5,
            format: LevelFormat.BULLET,
            text: "\u25aa",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: ll[5] }, run: { font: "Aptos" } },
          },
        ],
      },
    ],
  }
}
