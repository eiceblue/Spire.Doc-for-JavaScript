<template>
  <div>
    <el-tree style="max-width: 600px" :data="data" :props="defaultProps" accordion @node-click="handleNodeClick" />
  </div>
</template>
<script setup lang="ts">
interface Tree {
  label: string;
  children?: Tree[];
  $treeNodeId?: string | number;
  [key: string]: any;
}
const emits = defineEmits(["item-id"]);
const handleNodeClick = (data: Tree) => {
  if ((data as any).id !== undefined) {
    emits("item-id", (data as any).id);
  }
};
const data: Tree[] = [
  {
    label: "Spire.Doc",
    children: [
      // 01_QuickGuide
      {
        label: "01_QuickGuide",
        children: [
          {
            label: "HelloWorld",
            id: "1-1",
          },
        ],
      },
      // 02_FindAndReplace
      {
        label: "02_FindAndReplace",
        children: [
          {
            label: "FindAndHighlight",
            id: "2-1",
          },
          {
            label: "ReplaceContentWithDoc",
            id: "2-2",
          },
          {
            label: "ReplaceTextByRegex",
            id: "2-3",
          },
          {
            label: "ReplaceTextWithField",
            id: "2-4",
          },
          {
            label: "ReplaceTextWithTable",
            id: "2-5",
          },
          {
            label: "ReplaceWithDocument",
            id: "2-6",
          },
          {
            label: "ReplaceWithHtml",
            id: "2-7",
          },
          {
            label: "ReplaceWithImage",
            id: "2-8",
          },
          {
            label: "ReplaceWithText",
            id: "2-9",
          },
        ],
      },
      // 03_ExtractContent
      {
        label: "03_ExtractContent",
        children: [
          {
            label: "BetweenParagraphs",
            id: "3-1",
          },
          {
            label: "BetweenParagraphStyles",
            id: "3-2",
          },
          {
            label: "ExtractParagraphBasedOnStyle",
            id: "3-3",
          },
          {
            label: "FromBookmark",
            id: "3-4",
          },
          {
            label: "FromCommentRange",
            id: "3-5",
          },
          {
            label: "FromParagraphToTable",
            id: "3-6",
          },
          {
            label: "StartFromFormField",
            id: "3-7",
          },
        ],
      },
      // 04_Sections
      {
        label: "04_Sections",
        children: [
          {
            label: "AddAndDeleteSections",
            id: "4-1",
          },
          {
            label: "CloneSection",
            id: "4-2",
          },
          {
            label: "CloneSectionContent",
            id: "4-3",
          },
          {
            label: "ModifyPageSetupOfSection",
            id: "4-4",
          },
          {
            label: "RemoveSectionContent",
            id: "4-5",
          },
        ],
      },
      // 05_Paragraph
      {
        label: "05_Paragraph",
        children: [
          {
            label: "AddTabStopsToParagraph",
            id: "5-1",
          },
          {
            label: "AllowLatinTextWrapInMiddleOfAWord",
            id: "5-2",
          },
          {
            label: "CopyParagraph",
            id: "5-3",
          },
          {
            label: "FormACatalogue",
            id: "5-4",
          },
          {
            label: "GetParagraphByStyleName",
            id: "5-5",
          },
          {
            label: "GetParagraphRevisionsDetails",
            id: "5-6",
          },
          {
            label: "HideParagraph",
            id: "5-7",
          },
          {
            label: "InsertRtfStringToDoc",
            id: "5-8",
          },
          {
            label: "ManagePagination",
            id: "5-9",
          },
          {
            label: "RemoveAllParagraphs",
            id: "5-10",
          },
          {
            label: "RemoveEmptyLines",
            id: "5-11",
          },
          {
            label: "RemoveSpecificParagraph",
            id: "5-12",
          },
          {
            label: "SetFramePosition",
            id: "5-13",
          },
          {
            label: "SetParagraphShading",
            id: "5-14",
          },
          {
            label: "SetSpaceBetweenAsianAndLatinText",
            id: "5-15",
          },
          {
            label: "SetSpacing",
            id: "5-16",
          },
        ],
      },
      // 06_Text
      {
        label: "06_Text",
        children: [
          {
            label: "ApplyEmphasisMark",
            id: "6-1",
          },
          {
            label: "ChangeCase",
            id: "6-2",
          },
          {
            label: "CreateBarcode",
            id: "6-3",
          },
          {
            label: "GetText",
            id: "6-4",
          },
          {
            label: "InsertNewText",
            id: "6-5",
          },
          {
            label: "InsertSymbol",
            id: "6-6",
          },
          {
            label: "LoadTextWithEncoding",
            id: "6-7",
          },
          {
            label: "SetSuperscriptAndSubscript",
            id: "6-8",
          },
          {
            label: "SetTextDirection",
            id: "6-9",
          },
          {
            label: "SplitText",
            id: "6-10",
          },
        ],
      },
      // 07_DocumentOperation
      {
        label: "07_DocumentOperation",
        children: [
          {
            label: "AlterLanguageDictionary",
            id: "7-1",
          },
          {
            label: "CheckFileFormat",
            id: "7-2",
          },
          {
            label: "CompareDocuments",
            id: "7-3",
          },
          {
            label: "CompareDocumentsWithOptions",
            id: "7-4",
          },
          {
            label: "CountWordsNumber",
            id: "7-5",
          },
          {
            label: "DocumentProperty",
            id: "7-6",
          },
          {
            label: "LoadAndSaveToStream",
            id: "7-7",
          },
          {
            label: "RecurseAllDocumentObject",
            id: "7-8",
          },
          {
            label: "SetWordViewModes",
            id: "7-9",
          },
          {
            label: "UpdateLastSavedDate",
            id: "7-10",
          },
          {
            label: "AppendDocument",
            id: "7-11",
            children: [
              {
                label: "AddSectionFromOtherDoc",
                id: "7-11-1",
              },
              {
                label: "CloneWordDocument",
                id: "7-11-2",
              },
              {
                label: "CopyContentToAnotherDoc",
                id: "7-11-3",
              },
              {
                label: "KeepSameFormat",
                id: "7-11-4",
              },
              {
                label: "LinkHeadersFooters",
                id: "7-11-5",
              },
              {
                label: "Merge",
                id: "7-11-6",
              },
              {
                label: "MergeDocsOnSamePage",
                id: "7-11-7",
              },
              {
                label: "PreserveTheme",
                id: "7-11-8",
              },
              {
                label: "SectionBreakContinuous",
                id: "7-11-9",
              },
              {
                label: "SimpleInsertFile",
                id: "7-11-10",
              },
            ],
          },
          {
            label: "SplitDocument",
            id: "7-12",
            children: [
              {
                label: "SplitDocByPageBreak",
                id: "7-12-1",
              },
              {
                label: "SplitDocBySectionBreak",
                id: "7-12-2",
              },
              {
                label: "SplitDocIntoHtmlPages",
                id: "7-12-3",
              },
            ],
          },
          {
            label: "TrackChanges",
            id: "7-13",
            children: [
              {
                label: "AcceptOrRejectTrackedChange",
                id: "7-13-1",
              },
              {
                label: "EnableTrackChanges",
                id: "7-13-2",
              },
              {
                label: "GetRevisions",
                id: "7-13-3",
              },
            ],
          },
          {
            label: "Variables",
            id: "7-14",
            children: [
              {
                label: "AddVariables",
                id: "7-14-1",
              },
              {
                label: "CountVariables",
                id: "7-14-2",
              },
              {
                label: "GetVariables",
                id: "7-14-3",
              },
              {
                label: "RemoveVariables",
                id: "7-14-4",
              },
              {
                label: "RetrieveVariables",
                id: "7-14-5",
              },
            ],
          },
        ]
      },
      // 08_PageBackground
      {
        label: "08_PageBackground",
        children: [
          {
            label: "SetGradientBackground",
            id: "8-1",
          },
          {
            label: "SetImageBackground",
            id: "8-2",
          },
        ],
      },
      // 09_PageSetup
      {
        label: "09_PageSetup",
        children: [
          {
            label: "AddGutter",
            id: "9-1",
          },
          {
            label: "AddLineNumbers",
            id: "9-2",
          },
          {
            label: "AddPageBorders",
            id: "9-3",
          },
          {
            label: "AddPageNumbersInSections",
            id: "9-4",
          },
          {
            label: "DifferentPageSetup",
            id: "9-5",
          },
          {
            label: "InsertBreak",
            id: "9-6",
          },
          {
            label: "InsertPageBreakFirstApproach",
            id: "9-7",
          },
          {
            label: "InsertPageBreakSecondApproach",
            id: "9-8",
          },
          {
            label: "InsertSectionBreak",
            id: "9-9",
          },
          {
            label: "PageSetup",
            id: "9-10",
          },
          {
            label: "RemovePageBreaks",
            id: "9-11",
          },
          {
            label: "ResetPageNumber",
            id: "9-12",
          },
        ],
      },
      // 10_Conversion
      {
        label: "10_Conversion",
        children: [
          {
            label: "ConvertDocToByte",
            id: "10-1",
          },
          {
            label: "HtmlToImage",
            id: "10-2",
          },
          {
            label: "HtmlToPdf",
            id: "10-3",
          },
          {
            label: "HtmlToXml",
            id: "10-4",
          },
          {
            label: "HtmlToXps",
            id: "10-5",
          },
          {
            label: "ImageToPdf",
            id: "10-6",
          },
          {
            label: "OdtToWord",
            id: "10-7",
          },
          {
            label: "RtfToHtml",
            id: "10-8",
          },
          {
            label: "RtfToPDF",
            id: "10-9",
          },
          {
            label: "ToImage",
            id: "10-10",
          },
          {
            label: "ToOdt",
            id: "10-11",
          },
          {
            label: "ToPCL",
            id: "10-12",
          },
          {
            label: "ToPostScript",
            id: "10-13",
          },
          {
            label: "ToRtf",
            id: "10-14",
          },
          {
            label: "ToSVG",
            id: "10-15",
          },
          {
            label: "ToXML",
            id: "10-16",
          },
          {
            label: "ToXPS",
            id: "10-17",
          },
          {
            label: "TxtToWord",
            id: "10-18",
          },
          {
            label: "WordToPDFA",
            id: "10-19",
          },
          {
            label: "WordToTxt",
            id: "10-20",
          },
          {
            label: "WordToWordXML",
            id: "10-21",
          },
          {
            label: "XMLToPDF",
            id: "10-22",
          },
          {
            label: "XMLToWord",
            id: "10-23",
          },
          {
            label: "FromHtml",
            children: [
              {
                label: "HtmlFileToWord",
                id: "10-24-1",
              },
              {
                label: "HtmlStringToWord",
                id: "10-24-2",
              },
            ]
          },
          {
            label: "ToEPUB",
            children: [
              {
                label: "AddCoverImage",
                id: "10-25-1",
              },
              {
                label: "ToEpub",
                id: "10-25-2",
              },
            ]
          },
          {
            label: "ToHtml",
            children: [
              {
                label: "ToHtml",
                id: "10-26-1",
              },
              {
                label: "ToHtmlExportOption",
                id: "10-26-2",
              },
            ]
          },
          {
            label: "ToPDF",
            children: [
              {
                label: "DisableHyperlinks",
                id: "10-27-1",
              },
              {
                label: "EmbedAllFontsInPDF",
                id: "10-27-2",
              },
              {
                label: "EmbedNoninstalledFonts",
                id: "10-27-3",
              },
              {
                label: "KeepHiddenText",
                id: "10-27-4",
              },
              {
                label: "SetImageQuality",
                id: "10-27-5",
              },
              {
                label: "SpecifyEmbeddedFont",
                id: "10-27-6",
              },
              {
                label: "ToPDF",
                id: "10-27-7",
              },
              {
                label: "ToPDFAndCreateBookmarks",
                id: "10-27-8",
              },
              {
                label: "ToPdfWithPassword",
                id: "10-27-9",
              },
            ]
          },
        ],
      },
      // 11_Font
      {
        label: "11_Font",
        children: [
          {
            label: "ChangeFontColor",
            id: "11-1",
          },
          {
            label: "EmbedPrivateFont",
            id: "11-2",
          },
          {
            label: "GetListOfUsingFonts",
            id: "11-3",
          },
          {
            label: "SetFont",
            id: "11-4",
          },
        ],
      },
      // 12_Styles
      {
        label: "12_Styles",
        children: [
          {
            label: "ASCIICharactersBulletStyle",
            id: "12-1",
          },
          {
            label: "CharacterFormatting",
            id: "12-2",
          },
          {
            label: "CopyDocumentStyles",
            id: "12-3",
          },
          {
            label: "GetCharacterSpacing",
            id: "12-4",
          },
          {
            label: "GetTextByStyleName",
            id: "12-5",
          },
          {
            label: "Lists",
            id: "12-6",
          },
          {
            label: "MultiStylesInAParagraph",
            id: "12-7",
          },
          {
            label: "ParagraphFormatting",
            id: "12-8",
          },
          {
            label: "RestartList",
            id: "12-9",
          },
          {
            label: "RetrieveStyle",
            id: "12-10",
          },
          {
            label: "Styles",
            id: "12-11",
          },
        ],
      },
      // 13_MailMerge
      {
        label: "13_MailMerge",
        children: [
          {
            label: "ChangeLocale",
            id: "13-1",
          },
          {
            label: "ExecuteConditionalField",
            id: "13-2",
          },
          {
            label: "HideEmptyRegions",
            id: "13-3",
          },
          {
            label: "IdentifyMergeFieldName",
            id: "13-4",
          },
          {
            label: "MailMerge",
            id: "13-5",
          },
          {
            label: "MailMergeSwitches",
            id: "13-6",
          },
        ],
      },
      // 14_Bookmarks
      {
        label: "14_Bookmarks",
        children: [
          {
            label: "CopyBookmarkContent",
            id: "14-1",
          },
          {
            label: "CreateBookmark",
            id: "14-2",
          },
          {
            label: "CreateBookmarkForTable",
            id: "14-3",
          },
          {
            label: "ExtractBookmarkText",
            id: "14-4",
          },
          {
            label: "GetBookmarks",
            id: "14-5",
          },
          {
            label: "InsertDocAtBookmark",
            id: "14-6",
          },
          {
            label: "InsertImageAtBookmark",
            id: "14-7",
          },
          {
            label: "RemoveBookmark",
            id: "14-8",
          },
          {
            label: "RemoveBookmarkContent",
            id: "14-9",
          },
          {
            label: "ReplaceBookmarkContent",
            id: "14-10",
          },
          {
            label: "ReplaceWithTable",
            id: "14-11",
          },
        ],
      },
      // 15_Comments
      {
        label: "15_Comments",
        children: [
          {
            label: "AddCommentForSpecificText",
            id: "15-1",
          },
          {
            label: "Comment",
            id: "15-2",
          },
          {
            label: "ExtractComment",
            id: "15-3",
          },
          {
            label: "InsertPictureIntoComment",
            id: "15-4",
          },
          {
            label: "RemoveAndReplaceComment",
            id: "15-5",
          },
          {
            label: "RemoveContentWithComment",
            id: "15-6",
          },
          {
            label: "ReplyToComment",
            id: "15-7",
          }
        ],
      },
      // 16_ImageAndShape
      {
        label: "16_ImageAndShape",
        children: [
          {
            label: "AddBarcodeImage",
            id: "16-1",
          },
          {
            label: "AddHorizontalLine",
            id: "16-2",
          },
          {
            label: "AddImageToEachPage",
            id: "16-3",
          },
          {
            label: "AddShapeGroup",
            id: "16-4",
          },
          {
            label: "AddShapes",
            id: "16-5",
          },
          {
            label: "AlignShape",
            id: "16-6",
          },
          {
            label: "GetAlternativeText",
            id: "16-7",
          },
          {
            label: "Image",
            id: "16-8",
          },
          {
            label: "InsertImage",
            id: "16-9",
          },
          {
            label: "InsertWordArt",
            id: "16-10",
          },
          {
            label: "RemoveShape",
            id: "16-11",
          },
          {
            label: "ReplaceImageWithText",
            id: "16-12",
          },
          {
            label: "ResetImageSize",
            id: "16-13",
          },
          {
            label: "ResetShapeSize",
            id: "16-14",
          },
          {
            label: "RotateShape",
            id: "16-15",
          },
          {
            label: "SetLineShapeStyle",
            id: "16-16",
          },
          {
            label: "SetTextWrap",
            id: "16-17",
          },
          {
            label: "SetTransparencyForTextbox",
            id: "16-18",
          },
          {
            label: "SetTransparentColorForImage",
            id: "16-19",
          },
          {
            label: "UpdateImage",
            id: "16-20",
          }
        ],
      },
      // 17_HeaderAndFooter
      {
        label: "17_HeaderAndFooter",
        children: [
          {
            label: "AddHeaderOnlyFirstPage",
            id: "17-1",
          },
          {
            label: "AdjustHeaderFooterHeight",
            id: "17-2",
          },
          {
            label: "CopyHeaderAndFooter",
            id: "17-3",
          },
          {
            label: "DifferentFirstPage",
            id: "17-4",
          },
          {
            label: "HeaderAndFooter",
            id: "17-5",
          },
          {
            label: "ImageHeaderAndFooter",
            id: "17-6",
          },
          {
            label: "LockHeader",
            id: "17-7",
          },
          {
            label: "OddAndEvenHeaderFooter",
            id: "17-8",
          },
          {
            label: "PageBorderSurround",
            id: "17-9",
          },
          {
            label: "RemoveFooter",
            id: "17-10",
          },
          {
            label: "RemoveHeader",
            id: "17-11",
          },
        ],
      },
      // 18_Tables
      {
        label: "18_Tables",
        children: [
          {
            label: "AddAlternativeText",
            id: "18-1",
          },
          {
            label: "AddOrDeleteRow",
            id: "18-2",
          },
          {
            label: "AddOrRemoveColumn",
            id: "18-3",
          },
          {
            label: "AddPictureToTableCell",
            id: "18-4",
          },
          {
            label: "AllowBreakAcrossPages",
            id: "18-5",
          },
          {
            label: "AutoFitToContents",
            id: "18-6",
          },
          {
            label: "AutoFitToFixed",
            id: "18-7",
          },
          {
            label: "AutoFitToWindow",
            id: "18-8",
          },
          {
            label: "CellMergeStatus",
            id: "18-9",
          },
          {
            label: "CloneRow",
            id: "18-10",
          },
          {
            label: "CloneTable",
            id: "18-11",
          },
          {
            label: "CombineAndSplitTables",
            id: "18-12",
          },
          {
            label: "CreateNestedTable",
            id: "18-13",
          },
          {
            label: "CreateTable",
            id: "18-14",
          },
          {
            label: "CreateTableDirectly",
            id: "18-15",
          },
          {
            label: "CreateTableFromHTML",
            id: "18-16",
          },
          {
            label: "CreateVerticalTable",
            id: "18-17",
          },
          {
            label: "DifferentBorders",
            id: "18-18",
          },
          {
            label: "FormatMergedCells",
            id: "18-19",
          },
          {
            label: "GetDiagonalBorder",
            id: "18-20",
          },
          {
            label: "GetRowCellIndex",
            id: "18-21",
          },
          {
            label: "GetTablePosition",
            id: "18-22",
          },
          {
            label: "MergeAndSplitTableCell",
            id: "18-23",
          },
          {
            label: "ModifyTableFormat",
            id: "18-24",
          },
          {
            label: "PreventPageBreakInTable",
            id: "18-25",
          },
          {
            label: "RemoveTable",
            id: "18-26",
          },
          {
            label: "RepeatRowOnEachPage",
            id: "18-27",
          },
          {
            label: "ReplaceTextInTable",
            id: "18-28",
          },
          {
            label: "SetColumnWidth",
            id: "18-29",
          },
          {
            label: "SetOutsidePosition",
            id: "18-30",
          },
          {
            label: "SetTableStyleAndBorder",
            id: "18-31",
          },
          {
            label: "SetVerticalAlignment",
            id: "18-32",
          },
        ],
      },
      // 19_Hyperlinks
      {
        label: "19_Hyperlinks",
        children: [
          {
            label: "CreateImageHyperlink",
            id: "19-1",
          },
          {
            label: "FindHyperlinks",
            id: "19-2",
          },
          {
            label: "Hyperlink",
            id: "19-3",
          },
          {
            label: "ModifyHyperlinkText",
            id: "19-4",
          },
          {
            label: "RemoveHyperlinks",
            id: "19-5",
          },
          {
            label: "SetHyperlinkFormat",
            id: "19-6",
          },
        ],
      },
      // 20_Security
      {
        label: "20_Security",
        children: [
          {
            label: "Decrypt",
            id: "20-1",
          },
          {
            label: "Encrypt",
            id: "20-2",
          },
          {
            label: "LockSpecifiedSections",
            id: "20-3",
          },
          {
            label: "RemoveEditableRange",
            id: "20-4",
          },
          {
            label: "RemoveReadOnlyRestriction",
            id: "20-5",
          },
          {
            label: "SetEditableRange",
            id: "20-6",
          },
          {
            label: "SpecifiedProtectionType",
            id: "20-7",
          },
          {
            label: "WordToPdfEncrypt",
            id: "20-8",
          },
        ],
      },
      // 21_Fields
      {
        label: "21_Fields",
        children: [
          {
            label: "AddTCField",
            id: "21-1",
          },
          {
            label: "ConvertFieldToBodyText",
            id: "21-2",
          },
          {
            label: "ConvertIfFieldToText",
            id: "21-3",
          },
          {
            label: "CreateCrossReference",
            id: "21-4",
          },
          {
            label: "CreateFormField",
            id: "21-5",
          },
          {
            label: "CreateIFField",
            id: "21-6",
          },
          {
            label: "CreateNestedField",
            id: "21-7",
          },
          {
            label: "FillFormField",
            id: "21-8",
          },
          {
            label: "FormFieldsProperties",
            id: "21-9",
          },
          {
            label: "GetFieldText",
            id: "21-10",
          },
          {
            label: "GetFormFieldByName",
            id: "21-11",
          },
          {
            label: "GetFormFieldsCollection",
            id: "21-12",
          },
          {
            label: "GetMergeFieldName",
            id: "21-13",
          },
          {
            label: "InsertAddressBlockField",
            id: "21-14",
          },
          {
            label: "InsertAdvanceField",
            id: "21-15",
          },
          {
            label: "InsertMergeField",
            id: "21-16",
          },
          {
            label: "InsertNoneField",
            id: "21-17",
          },
          {
            label: "InsertPageRefField",
            id: "21-18",
          },
          {
            label: "RemoveCustomPropertyFields",
            id: "21-19",
          },
          {
            label: "RemoveField",
            id: "21-20",
          },
          {
            label: "ReplaceTextWithMergeField",
            id: "21-21",
          },
          {
            label: "UpdateFields",
            id: "21-22",
          },
        ],
      },
      // 22_TableOfContent
      {
        label: "22_TableOfContent",
        children: [
          {
            label: "ChangeTOCStyle",
            id: "22-1",
          },
          {
            label: "ChangeTOCTabStyle",
            id: "22-2",
          },
          {
            label: "CreateTableOfContentByDefault",
            id: "22-3",
          },
          {
            label: "CustomizeTableOfContent",
            id: "22-4",
          },
          {
            label: "RemoveTableOfContent",
            id: "22-5",
          },
        ],
      },
      // 23_Textbox
      {
        label: "23_Textbox",
        children: [
          {
            label: "DeleteTableFromTextBox",
            id: "23-1",
          },
          {
            label: "ExtractTextFromTextBoxes",
            id: "23-2",
          },
          {
            label: "InsertImageIntoTextBox",
            id: "23-3",
          },
          {
            label: "InsertTableIntoTextBox",
            id: "23-4",
          },
          {
            label: "ReadTableFromTextBox",
            id: "23-5",
          },
          {
            label: "RemoveTextBox",
            id: "23-6",
          },
          {
            label: "TextBox",
            id: "23-7",
          },
          {
            label: "TextBoxFormat",
            id: "23-8",
          },
        ],
      },
      // 24_Watermark
      {
        label: "24_Watermark",
        children: [
          {
            label: "ImageWaterMark",
            id: "24-1",
          },
          {
            label: "RemoveWatermark",
            id: "24-2",
          },
          {
            label: "TextWaterMark",
            id: "24-3",
          },
        ],
      },

      // 25_OLE
      {
        label: "25_OLE",
        children: [
          {
            label: "ExtractOLE",
            id: "25-1",
          },
          {
            label: "InsertOLE",
            id: "25-2",
          },
          {
            label: "InsertOLEAsIconViaStream",
            id: "25-3",
          },
        ],
      },
      // 26_StructuredDocumentTag
      {
        label: "26_StructuredDocumentTag",
        children: [
          {
            label: "AddCheckBoxContentControl",
            id: "26-1",
          },
          {
            label: "AddContentControls",
            id: "26-2",
          },
          {
            label: "AddRichTextContentControl",
            id: "26-3",
          },
          {
            label: "ComboBoxItem",
            id: "26-4",
          },
          {
            label: "GetContentControlProperty",
            id: "26-5",
          },
          {
            label: "LockContentControlContent",
            id: "26-6",
          },
          {
            label: "RemoveContentControls",
            id: "26-7",
          },
          {
            label: "UpdateCheckBox",
            id: "26-8",
          },
        ],
      },

      // 27_Footnote
      {
        label: "27_Footnote",
        children: [
          {
            label: "InsertEndnote",
            id: "27-1",
          },
          {
            label: "InsertFootnote",
            id: "27-2",
          },
          {
            label: "RemoveFootnote",
            id: "27-3",
          },
          {
            label: "SetPositionAndNumberFormat",
            id: "27-4",
          },
        ],
      },

      // 28_Macros
      {
        label: "28_Macros",
        children: [
          {
            label: "DetectAndRemoveVBAMacros",
            id: "28-1",
          },
          {
            label: "Macros",
            id: "28-2",
          },
        ],
      },

      // 29_Caption
      {
        label: "29_Caption",
        children: [
          {
            label: "AddPictureCaption",
            id: "29-1",
          },
          {
            label: "AddTableCaption",
            id: "29-2",
          },
          {
            label: "PictureCaptionCrossReference",
            id: "29-3",
          },
          {
            label: "SetCaptionWithChapterNumber",
            id: "29-4",
          },
          {
            label: "TableCaptionCrossReference",
            id: "29-5",
          },
        ],
      },

      // 30_Page
      {
        label: "30_Page",
        children: [
          {
            label: "FixedLayout",
            id: "30-1",
          },
        ],
      },

    ],
  },
];
const defaultProps = {
  children: "children",
  label: "label",
};
</script>