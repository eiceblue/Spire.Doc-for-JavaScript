<template>
  <div class="worker_page">
    <div class="tree_item">
      <SpireTree @item-id="handleChildID" />
    </div>
    <div class="file_item">
      <div class="file-reader">
        <div v-if="filePageId == '0'" class="non-centered">
          <Description />
        </div>
        
        <!--01_QuickGuide -->
        <div v-if="filePageId == '1-1'">
          <HelloWorld />
        </div>

        <!--02_FindAndReplace -->
        <div v-if="filePageId == '2-1'">
          <FindAndHighlight />
        </div>
        <div v-if="filePageId === '2-2'">
          <ReplaceContentWithDoc />
        </div>
        <div v-if="filePageId === '2-3'">
          <ReplaceTextByRegex />
        </div>
        <div v-if="filePageId === '2-4'">
          <ReplaceTextWithField />
        </div>
        <div v-if="filePageId === '2-5'">
          <ReplaceTextWithTable />
        </div>
        <div v-if="filePageId === '2-6'">
          <ReplaceWithDocument />
        </div>
        <div v-if="filePageId === '2-7'">
          <ReplaceWithHtml />
        </div>
        <div v-if="filePageId === '2-8'">
          <ReplaceWithImage />
        </div>
        <div v-if="filePageId === '2-9'">
          <ReplaceWithText />
        </div>
        
        <!--03_ExtractContent -->
        <div v-if="filePageId === '3-1'">
          <BetweenParagraphs />
        </div>
        <div v-if="filePageId === '3-2'">
          <BetweenParagraphStyles />
        </div>
        <div v-if="filePageId === '3-3'">
          <ExtractParagraphBasedOnStyle />
        </div>
        <div v-if="filePageId === '3-4'">
          <FromBookmark />
        </div>
        <div v-if="filePageId === '3-5'">
          <FromCommentRange />
        </div>
        <div v-if="filePageId === '3-6'">
          <FromParagraphToTable />
        </div>
        <div v-if="filePageId === '3-7'">
          <StartFromFormField />
        </div>
        
        <!--04_Sections -->
        <div v-if="filePageId === '4-1'">
          <AddAndDeleteSections />
        </div>
        <div v-if="filePageId === '4-2'">
          <CloneSection />
        </div>
        <div v-if="filePageId === '4-3'">
          <CloneSectionContent />
        </div>
        <div v-if="filePageId === '4-4'">
          <ModifyPageSetupOfSection />
        </div>
        <div v-if="filePageId === '4-5'">
          <RemoveSectionContent />
        </div>

        <!--05_Paragraph -->
        <div v-if="filePageId === '5-1'">
          <AddTabStopsToParagraph />
        </div>
        <div v-if="filePageId === '5-2'">
          <AllowLatinTextWrapInMiddleOfAWord />
        </div>
        <div v-if="filePageId === '5-3'">
          <CopyParagraph />
        </div>
        <div v-if="filePageId === '5-4'">
          <FormACatalogue />
        </div>
        <div v-if="filePageId === '5-5'">
          <GetParagraphByStyleName />
        </div>
        <div v-if="filePageId === '5-6'">
          <GetParagraphRevisionsDetails />
        </div>
        <div v-if="filePageId === '5-7'">
          <HideParagraph />
        </div>
        <div v-if="filePageId === '5-8'">
          <InsertRtfStringToDoc />
        </div>
        <div v-if="filePageId === '5-9'">
          <ManagePagination />
        </div>
        <div v-if="filePageId === '5-10'">
          <RemoveAllParagraphs />
        </div>
        <div v-if="filePageId === '5-11'">
          <RemoveEmptyLines />
        </div>
        <div v-if="filePageId === '5-12'">
          <RemoveSpecificParagraph />
        </div>
        <div v-if="filePageId === '5-13'">
          <SetFramePosition />
        </div>
        <div v-if="filePageId === '5-14'">
          <SetParagraphShading />
        </div>
        <div v-if="filePageId === '5-15'">
          <SetSpaceBetweenAsianAndLatinText />
        </div>
        <div v-if="filePageId === '5-16'">
          <SetSpacing />
        </div>
        
        <!--06_Text -->
        <div v-if="filePageId === '6-1'">
          <ApplyEmphasisMark />
        </div>
        <div v-if="filePageId === '6-2'">
          <ChangeCase />
        </div>
        <div v-if="filePageId === '6-3'">
          <CreateBarcode />
        </div>
        <div v-if="filePageId === '6-4'">
          <GetText />
        </div>
        <div v-if="filePageId === '6-5'">
          <InsertNewText />
        </div>
        <div v-if="filePageId === '6-6'">
          <InsertSymbol />
        </div>
        <div v-if="filePageId === '6-7'">
          <LoadTextWithEncoding />
        </div>
        <div v-if="filePageId === '6-8'">
          <SetSuperscriptAndSubscript />
        </div>
        <div v-if="filePageId === '6-9'">
          <SetTextDirection />
        </div>
        <div v-if="filePageId === '6-10'">
          <SplitText />
        </div>

        <!--07_DocumentOperation -->
        <div v-if="filePageId == '7-1'">
          <AlterLanguageDictionary />
        </div>
        <div v-if="filePageId == '7-2'">
          <CheckFileFormat />
        </div>
        <div v-if="filePageId == '7-3'">
          <CompareDocuments />
        </div>
        <div v-if="filePageId == '7-4'">
          <CompareDocumentsWithOptions />
        </div>
        <div v-if="filePageId == '7-5'">
          <CountWordsNumber />
        </div>
        <div v-if="filePageId == '7-6'">
          <DocumentProperty />
        </div>
        <div v-if="filePageId == '7-7'">
          <LoadAndSaveToStream />
        </div>
        <div v-if="filePageId == '7-8'">
          <RecurseAllDocumentObject />
        </div>
        <div v-if="filePageId == '7-9'">
          <SetWordViewModes />
        </div>
        <div v-if="filePageId == '7-10'">
          <UpdateLastSavedDate />
        </div>
        <div v-if="filePageId == '7-11-1'">
          <AddSectionFromOtherDoc />
        </div>
        <div v-if="filePageId == '7-11-2'">
          <CloneWordDocument />
        </div>
        <div v-if="filePageId == '7-11-3'">
          <CopyContentToAnotherDoc />
        </div>
        <div v-if="filePageId == '7-11-4'">
          <KeepSameFormat />
        </div>
        <div v-if="filePageId == '7-11-5'">
          <LinkHeadersFooters />
        </div>
        <div v-if="filePageId == '7-11-6'">
          <Merge />
        </div>
        <div v-if="filePageId == '7-11-7'">
          <MergeDocsOnSamePage />
        </div>
        <div v-if="filePageId == '7-11-8'">
          <PreserveTheme />
        </div>
        <div v-if="filePageId == '7-11-9'">
          <SectionBreakContinuous />
        </div>
        <div v-if="filePageId == '7-11-10'">
          <SimpleInsertFile />
        </div>
        <div v-if="filePageId == '7-12-1'">
          <SplitDocByPageBreak />
        </div>
        <div v-if="filePageId == '7-12-2'">
          <SplitDocBySectionBreak />
        </div>
        <div v-if="filePageId == '7-12-3'">
          <SplitDocIntoHtmlPages />
        </div>
        <div v-if="filePageId == '7-13-1'">
          <AcceptOrRejectTrackedChange />
        </div>
        <div v-if="filePageId == '7-13-2'">
          <EnableTrackChanges />
        </div>
        <div v-if="filePageId == '7-13-3'">
          <GetRevisions />
        </div>
        <div v-if="filePageId == '7-14-1'">
          <AddVariables />
        </div>
        <div v-if="filePageId == '7-14-2'">
          <CountVariables />
        </div>
        <div v-if="filePageId == '7-14-3'">
          <GetVariables />
        </div>
        <div v-if="filePageId == '7-14-4'">
          <RemoveVariables />
        </div>
        <div v-if="filePageId == '7-14-5'">
          <RetrieveVariables />
        </div>
        <!--08_PageBackground -->
        <div v-if="filePageId == '8-1'">
          <SetGradientBackground />
        </div>
        <div v-if="filePageId == '8-2'">
          <SetImageBackground />
        </div>
        <!--09_PageSetup-->
        <div v-if="filePageId == '9-1'">
          <AddGutter />
        </div>
        <div v-if="filePageId == '9-2'">
          <AddLineNumbers />
        </div>
        <div v-if="filePageId == '9-3'">
          <AddPageBorders />
        </div>
        <div v-if="filePageId == '9-4'">
          <AddPageNumbersInSections />
        </div>
        <div v-if="filePageId == '9-5'">
          <DifferentPageSetup />
        </div>
        <div v-if="filePageId == '9-6'">
          <InsertBreak />
        </div>
        <div v-if="filePageId == '9-7'">
          <InsertPageBreakFirstApproach />
        </div>
        <div v-if="filePageId == '9-8'">
          <InsertPageBreakSecondApproach />
        </div>
        <div v-if="filePageId == '9-9'">
          <InsertSectionBreak />
        </div>
        <div v-if="filePageId == '9-10'">
          <PageSetup />
        </div>
        <div v-if="filePageId == '9-11'">
          <RemovePageBreaks />
        </div>
        <div v-if="filePageId == '9-12'">
          <ResetPageNumber />
        </div>
        <!--10_Conversion -->
        <div v-if="filePageId == '10-1'">
          <ConvertDocToByte/>
        </div>
        <div v-if="filePageId == '10-2'">
          <HtmlToImage/>
        </div>
        <div v-if="filePageId == '10-3'">
          <HtmlToPdf/>
        </div>
        <div v-if="filePageId == '10-4'">
          <HtmlToXml/>
        </div>
        <div v-if="filePageId == '10-5'">
          <HtmlToXps/>
        </div>
        <div v-if="filePageId == '10-6'">
          <ImageToPdf/>
        </div>
        <div v-if="filePageId == '10-7'">
          <OdtToWord/>
        </div>
        <div v-if="filePageId == '10-8'">
          <RtfToHtml/>
        </div>
        <div v-if="filePageId == '10-9'">
          <RtfToPDF/>
        </div>
        <div v-if="filePageId == '10-10'">
          <ToImage/>
        </div>
        <div v-if="filePageId == '10-11'">
          <ToOdt/>
        </div>
        <div v-if="filePageId == '10-12'">
          <ToPCL/>
        </div>
        <div v-if="filePageId == '10-13'">
          <ToPostScript/>
        </div>
        <div v-if="filePageId == '10-14'">
          <ToRtf/>
        </div>
        <div v-if="filePageId == '10-15'">
          <ToSVG/>
        </div>
        <div v-if="filePageId == '10-16'">
          <ToXML/>
        </div>
        <div v-if="filePageId == '10-17'">
          <ToXPS/>
        </div>
        <div v-if="filePageId == '10-18'">
          <TxtToWord/>
        </div>
        <div v-if="filePageId == '10-19'">
          <WordToPDFA/>
        </div>        
        <div v-if="filePageId == '10-20'">
          <WordToTxt/>
        </div>
        <div v-if="filePageId == '10-21'">
          <WordToWordXML/>
        </div>
        <div v-if="filePageId == '10-22'">
          <XMLToPDF/>
        </div>
        <div v-if="filePageId == '10-23'">
          <XMLToWord/>
        </div>
        <div v-if="filePageId == '10-24-1'">
          <HtmlFileToWord/>
        </div>
        <div v-if="filePageId == '10-24-2'">
          <HtmlStringToWord/>
        </div>
        <div v-if="filePageId == '10-25-1'">
          <AddCoverImage/>
        </div>
        <div v-if="filePageId == '10-25-2'">
          <ToEpub/>
        </div>
        <div v-if="filePageId == '10-26-1'">
          <ToHtml/>
        </div>
        <div v-if="filePageId == '10-26-2'">
          <ToHtmlExportOption/>
        </div>
        <div v-if="filePageId == '10-27-1'">
          <DisableHyperlinks/>
        </div>
        <div v-if="filePageId == '10-27-2'">
          <EmbedAllFontsInPDF/>
        </div>
        <div v-if="filePageId == '10-27-3'">
          <EmbedNoninstalledFonts/>
        </div>
        <div v-if="filePageId == '10-27-4'">
          <KeepHiddenText/>
        </div>
        <div v-if="filePageId == '10-27-5'">
          <SetImageQuality/>
        </div>
        <div v-if="filePageId == '10-27-6'">
          <SpecifyEmbeddedFont/>
        </div>
        <div v-if="filePageId == '10-27-7'">
          <ToPDF/>
        </div>
        <div v-if="filePageId == '10-27-8'">
          <ToPDFAndCreateBookmarks/>
        </div>
        <div v-if="filePageId == '10-27-9'">
          <ToPdfWithPassword/>
        </div>
        <!--11_Font -->
        <div v-if="filePageId == '11-1'">
          <ChangeFontColor/>
        </div>
        <div v-if="filePageId == '11-2'">
          <EmbedPrivateFont/>
        </div>
        <div v-if="filePageId == '11-3'">
          <GetListOfUsingFonts/>
        </div>
        <div v-if="filePageId == '11-4'">
          <SetFont/>
        </div>
         <!--12_Styles -->
        <div v-if="filePageId == '12-1'">
          <ASCIICharactersBulletStyle/>
        </div>
        <div v-if="filePageId == '12-2'">
          <CharacterFormatting/>
        </div>
        <div v-if="filePageId == '12-3'">
          <CopyDocumentStyles/>
        </div>
        <div v-if="filePageId == '12-4'">
          <GetCharacterSpacing/>
        </div>
        <div v-if="filePageId == '12-5'">
          <GetTextByStyleName/>
        </div>
        <div v-if="filePageId == '12-6'">
          <Lists/>
        </div>
        <div v-if="filePageId == '12-7'">
          <MultiStylesInAParagraph/>
        </div>
        <div v-if="filePageId == '12-8'">
          <ParagraphFormatting/>
        </div>
        <div v-if="filePageId == '12-9'">
          <RestartList/>
        </div>
        <div v-if="filePageId == '12-10'">
          <RetrieveStyle/>
        </div>
        <div v-if="filePageId == '12-11'">
          <Styles/>
        </div>
        <!--13_MailMerge -->
        <div v-if="filePageId == '13-1'">
          <ChangeLocale/>
        </div>
        <div v-if="filePageId == '13-2'">
          <ExecuteConditionalField/>
        </div>
        <div v-if="filePageId == '13-3'">
          <HideEmptyRegions/>
        </div>
        <div v-if="filePageId == '13-4'">
          <IdentifyMergeFieldName/>
        </div>
        <div v-if="filePageId == '13-5'">
          <MailMerge/>
        </div>
        <div v-if="filePageId == '13-6'">
          <MailMergeSwitches/>
        </div>
        <div v-if="filePageId == '14-1'">
          <CopyBookmarkContent/>
        </div>
        <div v-if="filePageId == '14-2'">
          <CreateBookmark/>
        </div>
        <div v-if="filePageId == '14-3'">
          <CreateBookmarkForTable/>
        </div>
        <div v-if="filePageId == '14-4'">
          <ExtractBookmarkText/>
        </div>
        <div v-if="filePageId == '14-5'">
          <GetBookmarks/>
        </div>
        <div v-if="filePageId == '14-6'">
          <InsertDocAtBookmark/>
        </div>
        <div v-if="filePageId == '14-7'">
          <InsertImageAtBookmark/>
        </div>
        <div v-if="filePageId == '14-8'">
          <RemoveBookmark/>
        </div>
        <div v-if="filePageId == '14-9'">
          <RemoveBookmarkContent/>
        </div>
        <div v-if="filePageId == '14-10'">
          <ReplaceBookmarkContent/>
        </div>
        <div v-if="filePageId == '14-11'">
          <ReplaceWithTable/>
        </div>
        <!-- 15_Comments -->
        <div v-if="filePageId == '15-1'">
          <AddCommentForSpecificText />
        </div>
        <div v-if="filePageId == '15-2'">
          <Comment />
        </div>
        <div v-if="filePageId == '15-3'">
          <ExtractComment />
        </div>
        <div v-if="filePageId == '15-4'">
          <InsertPictureIntoComment />
        </div>
        <div v-if="filePageId == '15-5'">
          <RemoveAndReplaceComment />
        </div>
        <div v-if="filePageId == '15-6'">
          <RemoveContentWithComment />
        </div>
        <div v-if="filePageId == '15-7'">
          <ReplyToComment />
        </div>
        <!-- 16_ImageAndShape -->
        <div v-if="filePageId == '16-1'">
          <AddBarcodeImage />
        </div>
        <div v-if="filePageId == '16-2'">
          <AddHorizontalLine />
        </div>
        <div v-if="filePageId == '16-3'">
          <AddImageToEachPage />
        </div>
        <div v-if="filePageId == '16-4'">
          <AddShapeGroup />
        </div>
        <div v-if="filePageId == '16-5'">
          <AddShapes />
        </div>
        <div v-if="filePageId == '16-6'">
          <AlignShape />
        </div>
        <div v-if="filePageId == '16-7'">
          <GetAlternativeText />
        </div>
        <div v-if="filePageId == '16-8'">
          <Image />
        </div>
        <div v-if="filePageId == '16-9'">
          <InsertImage />
        </div>
        <div v-if="filePageId == '16-10'">
          <InsertWordArt />
        </div>
        <div v-if="filePageId == '16-11'">
          <RemoveShape />
        </div>
        <div v-if="filePageId == '16-12'">
          <ReplaceImageWithText />
        </div>
        <div v-if="filePageId == '16-13'">
          <ResetImageSize />
        </div>
        <div v-if="filePageId == '16-14'">
          <ResetShapeSize />
        </div>
        <div v-if="filePageId == '16-15'">
          <RotateShape />
        </div>
        <div v-if="filePageId == '16-16'">
          <SetLineShapeStyle />
        </div>
        <div v-if="filePageId == '16-17'">
          <SetTextWrap />
        </div>
        <div v-if="filePageId == '16-18'">
          <SetTransparencyForTextbox />
        </div>
        <div v-if="filePageId == '16-19'">
          <SetTransparentColorForImage />
        </div>
        <div v-if="filePageId == '16-20'">
          <UpdateImage />
        </div>
        <!--17_HeaderAndFooter -->
        <div v-if="filePageId == '17-1'">
          <AddHeaderOnlyFirstPage />
        </div>
        <div v-if="filePageId == '17-2'">
          <AdjustHeaderFooterHeight />
        </div>  
        <div v-if="filePageId == '17-3'">
          <CopyHeaderAndFooter />
        </div>
        <div v-if="filePageId == '17-4'">
          <DifferentFirstPage />
        </div>  
        <div v-if="filePageId == '17-5'">
          <HeaderAndFooter />
        </div>    
        <div v-if="filePageId == '17-6'">
          <ImageHeaderAndFooter />
        </div>  
        <div v-if="filePageId == '17-7'">
          <LockHeader/>
        </div>    
        <div v-if="filePageId == '17-8'">
          <OddAndEvenHeaderFooter />
        </div>  
        <div v-if="filePageId == '17-9'">
          <PageBorderSurround />
        </div>    
        <div v-if="filePageId == '17-10'">
          <RemoveFooter />
        </div>  
        <div v-if="filePageId == '17-11'">
          <RemoveHeader />
        </div>     
        <!--18_Tables -->
        <div v-if="filePageId == '18-1'">
          <AddAlternativeText />
        </div>
        <div v-if="filePageId == '18-2'">
          <AddOrDeleteRow />
        </div>  
        <div v-if="filePageId == '18-3'">
          <AddOrRemoveColumn />
        </div>
        <div v-if="filePageId == '18-4'">
          <AddPictureToTableCell />
        </div>  
        <div v-if="filePageId == '18-5'">
          <AllowBreakAcrossPages />
        </div>    
        <div v-if="filePageId == '18-6'">
          <AutoFitToContents />
        </div>  
        <div v-if="filePageId == '18-7'">
          <AutoFitToFixed/>
        </div>    
        <div v-if="filePageId == '18-8'">
          <AutoFitToWindow />
        </div>  
        <div v-if="filePageId == '18-9'">
          <CellMergeStatus />
        </div>    
        <div v-if="filePageId == '18-10'">
          <CloneRow />
        </div>  
        <div v-if="filePageId == '18-11'">
          <CloneTable />
        </div>            
        <div v-if="filePageId == '18-12'">
          <CombineAndSplitTables />
        </div>
        <div v-if="filePageId == '18-13'">
          <CreateNestedTable />
        </div>  
        <div v-if="filePageId == '18-14'">
          <CreateTable />
        </div>
        <div v-if="filePageId == '18-15'">
          <CreateTableDirectly />
        </div>  
        <div v-if="filePageId == '18-16'">
          <CreateTableFromHTML />
        </div>    
        <div v-if="filePageId == '18-17'">
          <CreateVerticalTable />
        </div>  
        <div v-if="filePageId == '18-18'">
          <DifferentBorders/>
        </div>    
        <div v-if="filePageId == '18-19'">
          <FormatMergedCells />
        </div>  
        <div v-if="filePageId == '18-20'">
          <GetDiagonalBorder />
        </div>    
        <div v-if="filePageId == '18-21'">
          <GetRowCellIndex />
        </div>  
        <div v-if="filePageId == '18-22'">
          <GetTablePosition/>
        </div>                    
          <div v-if="filePageId == '18-23'">
          <MergeAndSplitTableCell />
        </div>
        <div v-if="filePageId == '18-24'">
          <ModifyTableFormat />
        </div>  
        <div v-if="filePageId == '18-25'">
          <PreventPageBreakInTable />
        </div>
        <div v-if="filePageId == '18-26'">
          <RemoveTable />
        </div>  
        <div v-if="filePageId == '18-27'">
          <RepeatRowOnEachPage />
        </div>    
        <div v-if="filePageId == '18-28'">
          <ReplaceTextInTable />
        </div>  
        <div v-if="filePageId == '18-29'">
          <SetColumnWidth/>
        </div>    
        <div v-if="filePageId == '18-30'">
          <SetOutsidePosition/>
        </div>  
        <div v-if="filePageId == '18-31'">
          <SetTableStyleAndBorder />
        </div>    
        <div v-if="filePageId == '18-32'">
          <SetVerticalAlignment />
        </div>    
        <!--19_Hyperlinks -->
        <div v-if="filePageId == '19-1'">
          <CreateImageHyperlink />
        </div>
        <div v-if="filePageId == '19-2'">
          <FindHyperlinks />
        </div>
        <div v-if="filePageId == '19-3'">
          <Hyperlink />
        </div>
        <div v-if="filePageId == '19-4'">
          <ModifyHyperlinkText />
        </div>
        <div v-if="filePageId == '19-5'">
          <RemoveHyperlinks />
        </div>
        <div v-if="filePageId == '19-6'">
          <SetHyperlinkFormat />
        </div>
        <!--20_Security -->
        <div v-if="filePageId == '20-1'">
          <Decrypt />
        </div>
        <div v-if="filePageId == '20-2'">
          <Encrypt />
        </div>
        <div v-if="filePageId == '20-3'">
          <LockSpecifiedSections />
        </div>
        <div v-if="filePageId == '20-4'">
          <RemoveEditableRange />
        </div>
        <div v-if="filePageId == '20-5'">
          <RemoveReadOnlyRestriction />
        </div>
        <div v-if="filePageId == '20-6'">
          <SetEditableRange />
        </div>
        <div v-if="filePageId == '20-7'">
          <SpecifiedProtectionType />
        </div>
        <div v-if="filePageId == '20-8'">
          <WordToPdfEncrypt />
        </div>
        <!--21_Fields -->
        <div v-if="filePageId == '21-1'">
          <AddTCField />
        </div>
        <div v-if="filePageId == '21-2'">
          <ConvertFieldToBodyText />
        </div>
        <div v-if="filePageId == '21-3'">
          <ConvertIfFieldToText />
        </div>
        <div v-if="filePageId == '21-4'">
          <CreateCrossReference />
        </div>
        <div v-if="filePageId == '21-5'">
          <CreateFormField />
        </div>
        <div v-if="filePageId == '21-6'">
          <CreateIFField />
        </div>
        <div v-if="filePageId == '21-7'">
          <CreateNestedField />
        </div>
        <div v-if="filePageId == '21-8'">
          <FillFormField />
        </div>
        <div v-if="filePageId == '21-9'">
          <FormFieldsProperties />
        </div>
        <div v-if="filePageId == '21-10'">
          <GetFieldText />
        </div>
        <div v-if="filePageId == '21-11'">
          <GetFormFieldByName />
        </div>
        <div v-if="filePageId == '21-12'">
          <GetFormFieldsCollection />
        </div>
        <div v-if="filePageId == '21-13'">
          <GetMergeFieldName />
        </div>
        <div v-if="filePageId == '21-14'">
          <InsertAddressBlockField />
        </div>
        <div v-if="filePageId == '21-15'">
          <InsertAdvanceField />
        </div>
        <div v-if="filePageId == '21-16'">
          <InsertMergeField />
        </div>
        <div v-if="filePageId == '21-17'">
          <InsertNoneField />
        </div>
        <div v-if="filePageId == '21-18'">
          <InsertPageRefField />
        </div>
        <div v-if="filePageId == '21-19'">
          <RemoveCustomPropertyFields />
        </div>
        <div v-if="filePageId == '21-20'">
          <RemoveField />
        </div>
        <div v-if="filePageId == '21-21'">
          <ReplaceTextWithMergeField />
        </div>
        <div v-if="filePageId == '21-22'">
          <UpdateFields />
        </div>
        <!--22_TableOfContent -->
        <div v-if="filePageId == '22-1'">
          <ChangeTOCStyle />
        </div>
        <div v-if="filePageId == '22-2'">
          <ChangeTOCTabStyle />
        </div>
        <div v-if="filePageId == '22-3'">
          <CreateTableOfContentByDefault />
        </div>
        <div v-if="filePageId == '22-4'">
          <CustomizeTableOfContent />
        </div>
        <div v-if="filePageId == '22-5'">
          <RemoveTableOfContent />
        </div>
        <!--23_TextBox -->
        <div v-if="filePageId == '23-1'">
          <DeleteTableFromTextBox />
        </div>
        <div v-if="filePageId == '23-2'">
          <ExtractTextFromTextBoxes />
        </div>
        <div v-if="filePageId == '23-3'">
          <InsertImageIntoTextBox />
        </div>
        <div v-if="filePageId == '23-4'">
          <InsertTableIntoTextBox />
        </div>
        <div v-if="filePageId == '23-5'">
          <ReadTableFromTextBox />
        </div>
        <div v-if="filePageId == '23-6'">
          <RemoveTextBox />
        </div>
        <div v-if="filePageId == '23-7'">
          <TextBox />
        </div>
        <div v-if="filePageId == '23-8'">
          <TextBoxFormat />
        </div>
        <!--24_Watermark -->
        <div v-if="filePageId == '24-1'">
          <ImageWaterMark />
        </div>
        <div v-if="filePageId == '24-2'">
          <RemoveWatermark />
        </div>
        <div v-if="filePageId == '24-3'">
          <TextWaterMark />
        </div>
        <!--25_OLE -->
        <div v-if="filePageId == '25-1'">
          <ExtractOLE />
        </div>
        <div v-if="filePageId == '25-2'">
          <InsertOLE />
        </div>
        <div v-if="filePageId == '25-3'">
          <InsertOLEAsIconViaStream />
        </div>
        <!--26_StructuredDocumentTag -->
        <div v-if="filePageId == '26-1'">
          <AddCheckBoxContentControl />
        </div>
        <div v-if="filePageId == '26-2'">
          <AddContentControls />
        </div>
        <div v-if="filePageId == '26-3'">
          <AddRichTextContentControl />
        </div>
        <div v-if="filePageId == '26-4'">
          <ComboBoxItem />
        </div>
        <div v-if="filePageId == '26-5'">
          <GetContentControlProperty />
        </div>
        <div v-if="filePageId == '26-6'">
          <LockContentControlContent />
        </div>
        <div v-if="filePageId == '26-7'">
          <RemoveContentControls />
        </div>
        <div v-if="filePageId == '26-8'">
          <UpdateCheckBox />
        </div>
        <!--27_Footnote -->
        <div v-if="filePageId == '27-1'">
          <InsertEndnote />
        </div>
        <div v-if="filePageId == '27-2'">
          <InsertFootnote />
        </div>
        <div v-if="filePageId == '27-3'">
          <RemoveFootnote />
        </div>
        <div v-if="filePageId == '27-4'">
          <SetPositionAndNumberFormat />
        </div>
        <!--28_Macros -->
        <div v-if="filePageId == '28-1'">
          <DetectAndRemoveVBAMacros />
        </div>
        <div v-if="filePageId == '28-2'">
          <Macros />
        </div>
        <!--29_Caption -->
        <div v-if="filePageId == '29-1'">
          <AddPictureCaption />
        </div>
        <div v-if="filePageId == '29-2'">
          <AddTableCaption />
        </div>
        <div v-if="filePageId == '29-3'">
          <PictureCaptionCrossReference />
        </div>
        <div v-if="filePageId == '29-4'">
          <SetCaptionWithChapterNumber />
        </div>
        <div v-if="filePageId == '29-5'">
          <TableCaptionCrossReference />
        </div>
        <!--30_Page -->
        <div v-if="filePageId == '30-1'">
          <FixedLayout />
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref } from "vue";
import SpireTree from "@/components/SpireTree.vue";
import Description from "@/components/Description.vue";

// 01_quickguide
import HelloWorld from "@/components/01_quickguide/HelloWorld.vue";

// 02_FindAndReplace
import FindAndHighlight from "@/components/02_findandreplace/FindAndHighlight.vue";
import ReplaceContentWithDoc from "@/components/02_findandreplace/ReplaceContentWithDoc.vue";
import ReplaceTextByRegex from "@/components/02_findandreplace/ReplaceTextByRegex.vue";
import ReplaceTextWithField from "@/components/02_findandreplace/ReplaceTextWithField.vue";
import ReplaceTextWithTable from "@/components/02_findandreplace/ReplaceTextWithTable.vue";
import ReplaceWithDocument from "@/components/02_findandreplace/ReplaceWithDocument.vue";
import ReplaceWithHtml from "@/components/02_findandreplace/ReplaceWithHtml.vue";
import ReplaceWithImage from "@/components/02_findandreplace/ReplaceWithImage.vue";
import ReplaceWithText from "@/components/02_findandreplace/ReplaceWithText.vue";

// 03_ExtractContent
import BetweenParagraphs from "@/components/03_extractContent/BetweenParagraphs.vue";
import BetweenParagraphStyles from "@/components/03_extractContent/BetweenParagraphStyles.vue";
import ExtractParagraphBasedOnStyle from "@/components/03_extractContent/ExtractParagraphBasedOnStyle.vue";
import FromBookmark from "@/components/03_extractContent/FromBookmark.vue";
import FromCommentRange from "@/components/03_extractContent/FromCommentRange.vue";
import FromParagraphToTable from "@/components/03_extractContent/FromParagraphToTable.vue";
import StartFromFormField from "@/components/03_extractContent/StartFromFormField.vue";

// 04_Sections
import AddAndDeleteSections from "@/components/04_sections/AddAndDeleteSections.vue";
import CloneSection from "@/components/04_sections/CloneSection.vue";
import CloneSectionContent from "@/components/04_sections/CloneSectionContent.vue";
import ModifyPageSetupOfSection from "@/components/04_sections/ModifyPageSetupOfSection.vue";
import RemoveSectionContent from "@/components/04_sections/RemoveSectionContent.vue";

// 05_Paragraph
import AddTabStopsToParagraph from "@/components/05_paragraph/AddTabStopsToParagraph.vue";
import AllowLatinTextWrapInMiddleOfAWord from "@/components/05_paragraph/AllowLatinTextWrapInMiddleOfAWord.vue";
import CopyParagraph from "@/components/05_paragraph/CopyParagraph.vue";
import FormACatalogue from "@/components/05_paragraph/FormACatalogue.vue";
import GetParagraphByStyleName from "@/components/05_paragraph/GetParagraphByStyleName.vue";
import GetParagraphRevisionsDetails from "@/components/05_paragraph/GetParagraphRevisionsDetails.vue";
import HideParagraph from "@/components/05_paragraph/HideParagraph.vue";
import InsertRtfStringToDoc from "@/components/05_paragraph/InsertRtfStringToDoc.vue";
import ManagePagination from "@/components/05_paragraph/ManagePagination.vue";
import RemoveAllParagraphs from "@/components/05_paragraph/RemoveAllParagraphs.vue";
import RemoveEmptyLines from "@/components/05_paragraph/RemoveEmptyLines.vue";
import RemoveSpecificParagraph from "@/components/05_paragraph/RemoveSpecificParagraph.vue";
import SetFramePosition from "@/components/05_paragraph/SetFramePosition.vue";
import SetParagraphShading from "@/components/05_paragraph/SetParagraphShading.vue";
import SetSpaceBetweenAsianAndLatinText from "@/components/05_paragraph/SetSpaceBetweenAsianAndLatinText.vue";
import SetSpacing from "@/components/05_paragraph/SetSpacing.vue";

// 06_Text
import ApplyEmphasisMark from "@/components/06_text/ApplyEmphasisMark.vue";
import ChangeCase from "@/components/06_text/ChangeCase.vue";
import CreateBarcode from "@/components/06_text/CreateBarcode.vue";
import GetText from "@/components/06_text/GetText.vue";
import InsertNewText from "@/components/06_text/InsertNewText.vue";
import InsertSymbol from "@/components/06_text/InsertSymbol.vue";
import LoadTextWithEncoding from "@/components/06_text/LoadTextWithEncoding.vue";
import SetSuperscriptAndSubscript from "@/components/06_text/SetSuperscriptAndSubscript.vue";
import SetTextDirection from "@/components/06_text/SetTextDirection.vue";
import SplitText from "@/components/06_text/SplitText.vue";

//07_DocumentOperation
import AlterLanguageDictionary from "@/components/07_documentOperation/AlterLanguageDictionary.vue";
import CheckFileFormat from "@/components/07_documentOperation/CheckFileFormat.vue";
import CompareDocuments from "@/components/07_documentOperation/CompareDocuments.vue";
import CompareDocumentsWithOptions from "@/components/07_documentOperation/CompareDocumentsWithOptions.vue";
import CountWordsNumber from "@/components/07_documentOperation/CountWordsNumber.vue";
import DocumentProperty from "@/components/07_documentOperation/DocumentProperty.vue";
import LoadAndSaveToStream from "@/components/07_documentOperation/LoadAndSaveToStream.vue";
import RecurseAllDocumentObject from "@/components/07_documentOperation/RecurseAllDocumentObject.vue";
import SetWordViewModes from "@/components/07_documentOperation/SetWordViewModes.vue";
import UpdateLastSavedDate from "@/components/07_documentOperation/UpdateLastSavedDate.vue";
import AddSectionFromOtherDoc from "@/components/07_documentOperation/appendDocument/AddSectionFromOtherDoc.vue";
import CloneWordDocument from "@/components/07_documentOperation/appendDocument/CloneWordDocument.vue";
import CopyContentToAnotherDoc from "@/components/07_documentOperation/appendDocument/CopyContentToAnotherDoc.vue";
import KeepSameFormat from "@/components/07_documentOperation/appendDocument/KeepSameFormat.vue";
import LinkHeadersFooters from "@/components/07_documentOperation/appendDocument/LinkHeadersFooters.vue";
import Merge from "@/components/07_documentOperation/appendDocument/Merge.vue";
import MergeDocsOnSamePage from "@/components/07_documentOperation/appendDocument/MergeDocsOnSamePage.vue";
import PreserveTheme from "@/components/07_documentOperation/appendDocument/PreserveTheme.vue";
import SectionBreakContinuous from "@/components/07_documentOperation/appendDocument/SectionBreakContinuous.vue";
import SimpleInsertFile from "@/components/07_documentOperation/appendDocument/SimpleInsertFile.vue";
import SplitDocByPageBreak from "@/components/07_documentOperation/splitDocument/SplitDocByPageBreak.vue";
import SplitDocBySectionBreak from "@/components/07_documentOperation/splitDocument/SplitDocBySectionBreak.vue";
import SplitDocIntoHtmlPages from "@/components/07_documentOperation/splitDocument/SplitDocIntoHtmlPages.vue";
import AcceptOrRejectTrackedChange from "@/components/07_documentOperation/trackChanges/AcceptOrRejectTrackedChange.vue";
import EnableTrackChanges from "@/components/07_documentOperation/trackChanges/EnableTrackChanges.vue";
import GetRevisions from "@/components/07_documentOperation/trackChanges/GetRevisions.vue";
import AddVariables from "@/components/07_documentOperation/variables/AddVariables.vue";
import CountVariables from "@/components/07_documentOperation/variables/CountVariables.vue";
import GetVariables from "@/components/07_documentOperation/variables/GetVariables.vue";
import RemoveVariables from "@/components/07_documentOperation/variables/RemoveVariables.vue";
import RetrieveVariables from "@/components/07_documentOperation/variables/RetrieveVariables.vue";

//08_PageBackground
import SetGradientBackground from "@/components/08_pageBackground/SetGradientBackground.vue";
import SetImageBackground from "@/components/08_pageBackground/SetImageBackground.vue";

//09_PageSetup
import AddGutter from "@/components/09_pageSetup/AddGutter.vue";
import AddLineNumbers from "@/components/09_pageSetup/AddLineNumbers.vue";
import AddPageBorders from "@/components/09_pageSetup/AddPageBorders.vue";
import AddPageNumbersInSections from "@/components/09_pageSetup/AddPageNumbersInSections.vue";
import DifferentPageSetup from "@/components/09_pageSetup/DifferentPageSetup.vue";
import InsertBreak from "@/components/09_pageSetup/InsertBreak.vue";
import InsertPageBreakFirstApproach from "@/components/09_pageSetup/InsertPageBreakFirstApproach.vue";
import InsertPageBreakSecondApproach from "@/components/09_pageSetup/InsertPageBreakSecondApproach.vue";
import InsertSectionBreak from "@/components/09_pageSetup/InsertSectionBreak.vue";
import PageSetup from "@/components/09_pageSetup/PageSetup.vue";
import RemovePageBreaks from "@/components/09_pageSetup/RemovePageBreaks.vue";
import ResetPageNumber from "@/components/09_pageSetup/ResetPageNumber.vue";

//10_conversion
import ConvertDocToByte from "@/components/10_conversion/ConvertDocToByte.vue";
import HtmlToImage from "@/components/10_conversion/HtmlToImage.vue";
import HtmlToPdf from "@/components/10_conversion/HtmlToPdf.vue";
import HtmlToXml from "@/components/10_conversion/HtmlToXml.vue";
import HtmlToXps from "@/components/10_conversion/HtmlToXps.vue";
import ImageToPdf from "@/components/10_conversion/ImageToPdf.vue";
import OdtToWord from "@/components/10_conversion/OdtToWord.vue";
import RtfToHtml from "@/components/10_conversion/RtfToHtml.vue";
import RtfToPDF from "@/components/10_conversion/RtfToPDF.vue";
import ToImage from "@/components/10_conversion/ToImage.vue";
import ToOdt from "@/components/10_conversion/ToOdt.vue";
import ToPCL from "@/components/10_conversion/ToPCL.vue";
import ToPostScript from "@/components/10_conversion/ToPostScript.vue";
import ToRtf from "@/components/10_conversion/ToRtf.vue";
import ToSVG from "@/components/10_conversion/ToSVG.vue";
import ToXML from "@/components/10_conversion/ToXML.vue";
import ToXPS from "@/components/10_conversion/ToXPS.vue";
import TxtToWord from "@/components/10_conversion/TxtToWord.vue";
import WordToPDFA from "@/components/10_conversion/WordToPDFA.vue";
import WordToTxt from "@/components/10_conversion/WordToTxt.vue";
import WordToWordXML from "@/components/10_conversion/WordToWordXML.vue";
import XMLToPDF from "@/components/10_conversion/XMLToPDF.vue";
import XMLToWord from "@/components/10_conversion/XMLToWord.vue";
import HtmlFileToWord from "@/components/10_conversion/FromHtml/HtmlFileToWord.vue";
import HtmlStringToWord from "@/components/10_conversion/FromHtml/HtmlStringToWord.vue";
import AddCoverImage from "@/components/10_conversion/ToEPUB/AddCoverImage.vue";
import ToEpub from "@/components/10_conversion/ToEPUB/ToEpub.vue";
import ToHtml from "@/components/10_conversion/ToHtml/ToHtml.vue";
import ToHtmlExportOption from "@/components/10_conversion/ToHtml/ToHtmlExportOption.vue";
import DisableHyperlinks from "@/components/10_conversion/ToPDF/DisableHyperlinks.vue";
import EmbedAllFontsInPDF from "@/components/10_conversion/ToPDF/EmbedAllFontsInPDF.vue";
import EmbedNoninstalledFonts from "@/components/10_conversion/ToPDF/EmbedNoninstalledFonts.vue";
import KeepHiddenText from "@/components/10_conversion/ToPDF/KeepHiddenText.vue";
import SetImageQuality from "@/components/10_conversion/ToPDF/SetImageQuality.vue";
import SpecifyEmbeddedFont from "@/components/10_conversion/ToPDF/SpecifyEmbeddedFont.vue";
import ToPDF from "@/components/10_conversion/ToPDF/ToPDF.vue";
import ToPDFAndCreateBookmarks from "@/components/10_conversion/ToPDF/ToPDFAndCreateBookmarks.vue";
import ToPdfWithPassword from "@/components/10_conversion/ToPDF/ToPdfWithPassword.vue";

//11_Font
import ChangeFontColor from "@/components/11_font/ChangeFontColor.vue";
import EmbedPrivateFont from "@/components/11_font/EmbedPrivateFont.vue";
import GetListOfUsingFonts from "@/components/11_font/GetListOfUsingFonts.vue";
import SetFont from "@/components/11_font/SetFont.vue";

//12_Styles
import ASCIICharactersBulletStyle from "@/components/12_styles/ASCIICharactersBulletStyle.vue";
import CharacterFormatting from "@/components/12_styles/CharacterFormatting.vue";
import CopyDocumentStyles from "@/components/12_styles/CopyDocumentStyles.vue";
import GetCharacterSpacing from "@/components/12_styles/GetCharacterSpacing.vue";
import GetTextByStyleName from "@/components/12_styles/GetTextByStyleName.vue";
import Lists from "@/components/12_styles/Lists.vue";
import MultiStylesInAParagraph from "@/components/12_styles/MultiStylesInAParagraph.vue";
import ParagraphFormatting from "@/components/12_styles/ParagraphFormatting.vue";
import RestartList from "@/components/12_styles/RestartList.vue";
import RetrieveStyle from "@/components/12_styles/RetrieveStyle.vue";
import Styles from "@/components/12_styles/Styles.vue";

//13_MailMerge
import ChangeLocale from "@/components/13_mailmerge/ChangeLocale.vue";
import ExecuteConditionalField from "@/components/13_mailmerge/ExecuteConditionalField.vue";
import HideEmptyRegions from "@/components/13_mailmerge/HideEmptyRegions.vue";
import IdentifyMergeFieldName from "@/components/13_mailmerge/IdentifyMergeFieldName.vue";
import MailMerge from "@/components/13_mailmerge/MailMerge.vue";
import MailMergeSwitches from "@/components/13_mailmerge/MailMergeSwitches.vue";

//14_Bookmarks
import CopyBookmarkContent from "@/components/14_bookmarks/CopyBookmarkContent.vue";
import CreateBookmark from "@/components/14_bookmarks/CreateBookmark.vue";
import CreateBookmarkForTable from "@/components/14_bookmarks/CreateBookmarkForTable.vue";
import ExtractBookmarkText from "@/components/14_bookmarks/ExtractBookmarkText.vue";
import GetBookmarks from "@/components/14_bookmarks/GetBookmarks.vue";
import InsertDocAtBookmark from "@/components/14_bookmarks/InsertDocAtBookmark.vue";
import InsertImageAtBookmark from "@/components/14_bookmarks/InsertImageAtBookmark.vue";
import RemoveBookmark from "@/components/14_bookmarks/RemoveBookmark.vue";
import RemoveBookmarkContent from "@/components/14_bookmarks/RemoveBookmarkContent.vue";
import ReplaceBookmarkContent from "@/components/14_bookmarks/ReplaceBookmarkContent.vue";
import ReplaceWithTable from "@/components/14_bookmarks/ReplaceWithTable.vue";

//15_comments
import AddCommentForSpecificText from "@/components/15_comments/AddCommentForSpecificText.vue";
import Comment from "@/components/15_comments/Comment.vue";
import ExtractComment from "@/components/15_comments/ExtractComment.vue";
import InsertPictureIntoComment from "@/components/15_comments/InsertPictureIntoComment.vue";
import RemoveAndReplaceComment from "@/components/15_comments/RemoveAndReplaceComment.vue";
import RemoveContentWithComment from "@/components/15_comments/RemoveContentWithComment.vue";
import ReplyToComment from "@/components/15_comments/ReplyToComment.vue";

//16_imageAndShape
import AddBarcodeImage from "@/components/16_imageAndShape/AddBarcodeImage.vue";
import AddHorizontalLine from "@/components/16_imageAndShape/AddHorizontalLine.vue";
import AddImageToEachPage from "@/components/16_imageAndShape/AddImageToEachPage.vue";
import AddShapeGroup from "@/components/16_imageAndShape/AddShapeGroup.vue";
import AddShapes from "@/components/16_imageAndShape/AddShapes.vue";
import AlignShape from "@/components/16_imageAndShape/AlignShape.vue";
import GetAlternativeText from "@/components/16_imageAndShape/GetAlternativeText.vue";
import Image from "@/components/16_imageAndShape/Image.vue";
import InsertImage from "@/components/16_imageAndShape/InsertImage.vue";
import InsertWordArt from "@/components/16_imageAndShape/InsertWordArt.vue";
import RemoveShape from "@/components/16_imageAndShape/RemoveShape.vue";
import ReplaceImageWithText from "@/components/16_imageAndShape/ReplaceImageWithText.vue";
import ResetImageSize from "@/components/16_imageAndShape/ResetImageSize.vue";
import ResetShapeSize from "@/components/16_imageAndShape/ResetShapeSize.vue";
import RotateShape from "@/components/16_imageAndShape/RotateShape.vue";
import SetLineShapeStyle from "@/components/16_imageAndShape/SetLineShapeStyle.vue";
import SetTextWrap from "@/components/16_imageAndShape/SetTextWrap.vue";
import SetTransparencyForTextbox from "@/components/16_imageAndShape/SetTransparencyForTextbox.vue";
import SetTransparentColorForImage from "@/components/16_imageAndShape/SetTransparentColorForImage.vue";
import UpdateImage from "@/components/16_imageAndShape/UpdateImage.vue";

//17_headerAndFooter
import AddHeaderOnlyFirstPage from "@/components/17_headerAndFooter/AddHeaderOnlyFirstPage.vue";
import AdjustHeaderFooterHeight from "@/components/17_headerAndFooter/AdjustHeaderFooterHeight.vue";
import CopyHeaderAndFooter from "@/components/17_headerAndFooter/CopyHeaderAndFooter.vue";
import DifferentFirstPage from "@/components/17_headerAndFooter/DifferentFirstPage.vue";
import HeaderAndFooter from "@/components/17_headerAndFooter/HeaderAndFooter.vue";
import ImageHeaderAndFooter from "@/components/17_headerAndFooter/ImageHeaderAndFooter.vue";
import LockHeader from "@/components/17_headerAndFooter/LockHeader.vue";
import OddAndEvenHeaderFooter from "@/components/17_headerAndFooter/OddAndEvenHeaderFooter.vue";
import PageBorderSurround from "@/components/17_headerAndFooter/PageBorderSurround.vue";
import RemoveFooter from "@/components/17_headerAndFooter/RemoveFooter.vue";
import RemoveHeader from "@/components/17_headerAndFooter/RemoveHeader.vue";

// 18_tables
import AddAlternativeText from "@/components/18_tables/AddAlternativeText.vue";
import AddOrDeleteRow from "@/components/18_tables/AddOrDeleteRow.vue";
import AddOrRemoveColumn from "@/components/18_tables/AddOrRemoveColumn.vue";
import AddPictureToTableCell from "@/components/18_tables/AddPictureToTableCell.vue";
import AllowBreakAcrossPages from "@/components/18_tables/AllowBreakAcrossPages.vue";
import AutoFitToContents from "@/components/18_tables/AutoFitToContents.vue";
import AutoFitToFixed from "@/components/18_tables/AutoFitToFixed.vue";
import AutoFitToWindow from "@/components/18_tables/AutoFitToWindow.vue";
import CellMergeStatus from "@/components/18_tables/CellMergeStatus.vue";
import CloneRow from "@/components/18_tables/CloneRow.vue";
import CloneTable from "@/components/18_tables/CloneTable.vue";
import CombineAndSplitTables from "@/components/18_tables/CombineAndSplitTables.vue";
import CreateNestedTable from "@/components/18_tables/CreateNestedTable.vue";
import CreateTable from "@/components/18_tables/CreateTable.vue";
import CreateTableDirectly from "@/components/18_tables/CreateTableDirectly.vue";
import CreateTableFromHTML from "@/components/18_tables/CreateTableFromHTML.vue";
import CreateVerticalTable from "@/components/18_tables/CreateVerticalTable.vue";
import DifferentBorders from "@/components/18_tables/DifferentBorders.vue";
import FormatMergedCells from "@/components/18_tables/FormatMergedCells.vue";
import GetDiagonalBorder from "@/components/18_tables/GetDiagonalBorder.vue";
import GetRowCellIndex from "@/components/18_tables/GetRowCellIndex.vue";
import GetTablePosition from "@/components/18_tables/GetTablePosition.vue";
import MergeAndSplitTableCell from "@/components/18_tables/MergeAndSplitTableCell.vue";
import ModifyTableFormat from "@/components/18_tables/ModifyTableFormat.vue";
import PreventPageBreakInTable from "@/components/18_tables/PreventPageBreakInTable.vue";
import RemoveTable from "@/components/18_tables/RemoveTable.vue";
import RepeatRowOnEachPage from "@/components/18_tables/RepeatRowOnEachPage.vue";
import ReplaceTextInTable from "@/components/18_tables/ReplaceTextInTable.vue";
import SetColumnWidth from "@/components/18_tables/SetColumnWidth.vue";
import SetOutsidePosition from "@/components/18_tables/SetOutsidePosition.vue";
import SetTableStyleAndBorder from "@/components/18_tables/SetTableStyleAndBorder.vue";
import SetVerticalAlignment from "@/components/18_tables/SetVerticalAlignment.vue";

// 19_hyperlinks
import CreateImageHyperlink from "@/components/19_hyperlinks/CreateImageHyperlink.vue";
import FindHyperlinks from "@/components/19_hyperlinks/FindHyperlinks.vue";
import Hyperlink from "@/components/19_hyperlinks/Hyperlink.vue";
import ModifyHyperlinkText from "@/components/19_hyperlinks/ModifyHyperlinkText.vue";
import RemoveHyperlinks from "@/components/19_hyperlinks/RemoveHyperlinks.vue";
import SetHyperlinkFormat from "@/components/19_hyperlinks/SetHyperlinkFormat.vue";

//20_Security
import Decrypt from "@/components/20_security/Decrypt.vue";
import Encrypt from "@/components/20_security/Encrypt.vue";
import LockSpecifiedSections from "@/components/20_security/LockSpecifiedSections.vue";
import RemoveEditableRange from "@/components/20_security/RemoveEditableRange.vue";
import RemoveReadOnlyRestriction from "@/components/20_security/RemoveReadOnlyRestriction.vue";
import SetEditableRange from "@/components/20_security/SetEditableRange.vue";
import SpecifiedProtectionType from "@/components/20_security/SpecifiedProtectionType.vue";
import WordToPdfEncrypt from "@/components/20_security/WordToPdfEncrypt.vue";

//21_Fields
import AddTCField from "@/components/21_fields/AddTCField.vue";
import ConvertFieldToBodyText from "@/components/21_fields/ConvertFieldToBodyText.vue";
import ConvertIfFieldToText from "@/components/21_fields/ConvertIfFieldToText.vue";
import CreateCrossReference from "@/components/21_fields/CreateCrossReference.vue";
import CreateFormField from "@/components/21_fields/CreateFormField.vue";
import CreateIFField from "@/components/21_fields/CreateIFField.vue";
import CreateNestedField from "@/components/21_fields/CreateNestedField.vue";
import FillFormField from "@/components/21_fields/FillFormField.vue";
import FormFieldsProperties from "@/components/21_fields/FormFieldsProperties.vue";
import GetFieldText from "@/components/21_fields/GetFieldText.vue";
import GetFormFieldByName from "@/components/21_fields/GetFormFieldByName.vue";
import GetFormFieldsCollection from "@/components/21_fields/GetFormFieldsCollection.vue";
import GetMergeFieldName from "@/components/21_fields/GetMergeFieldName.vue";
import InsertAddressBlockField from "@/components/21_fields/InsertAddressBlockField.vue";
import InsertAdvanceField from "@/components/21_fields/InsertAdvanceField.vue";
import InsertMergeField from "@/components/21_fields/InsertMergeField.vue";
import InsertNoneField from "@/components/21_fields/InsertNoneField.vue";
import InsertPageRefField from "@/components/21_fields/InsertPageRefField.vue";
import RemoveCustomPropertyFields from "@/components/21_fields/RemoveCustomPropertyFields.vue";
import RemoveField from "@/components/21_fields/RemoveField.vue";
import ReplaceTextWithMergeField from "@/components/21_fields/ReplaceTextWithMergeField.vue";
import UpdateFields from "@/components/21_fields/UpdateFields.vue";

//22_TableOfContent
import ChangeTOCStyle from "@/components/22_tableOfContent/ChangeTOCStyle.vue";
import ChangeTOCTabStyle from "@/components/22_tableOfContent/ChangeTOCTabStyle.vue";
import CreateTableOfContentByDefault from "@/components/22_tableOfContent/CreateTableOfContentByDefault.vue";
import CustomizeTableOfContent from "@/components/22_tableOfContent/CustomizeTableOfContent.vue";
import RemoveTableOfContent from "@/components/22_tableOfContent/RemoveTableOfContent.vue";

//23_Textbox
import DeleteTableFromTextBox from "@/components/23_textbox/DeleteTableFromTextBox.vue";
import ExtractTextFromTextBoxes from "@/components/23_textbox/ExtractTextFromTextBoxes.vue";
import InsertImageIntoTextBox from "@/components/23_textbox/InsertImageIntoTextBox.vue";
import InsertTableIntoTextBox from "@/components/23_textbox/InsertTableIntoTextBox.vue";
import ReadTableFromTextBox from "@/components/23_textbox/ReadTableFromTextBox.vue";
import RemoveTextBox from "@/components/23_textbox/RemoveTextBox.vue";
import TextBox from "@/components/23_textbox/TextBox.vue";
import TextBoxFormat from "@/components/23_textbox/TextBoxFormat.vue";

//24_Watermark
import ImageWaterMark from "@/components/24_watermark/ImageWaterMark.vue";
import RemoveWatermark from "@/components/24_watermark/RemoveWatermark.vue";
import TextWaterMark from "@/components/24_watermark/TextWaterMark.vue";

//25_OLE
import ExtractOLE from "@/components/25_ole/ExtractOLE.vue";
import InsertOLE from "@/components/25_ole/InsertOLE.vue";
import InsertOLEAsIconViaStream from "@/components/25_ole/InsertOLEAsIconViaStream.vue";

//26_StructuredDocumentTag
import AddCheckBoxContentControl from "@/components/26_structuredDocumentTag/AddCheckBoxContentControl.vue";
import AddContentControls from "@/components/26_structuredDocumentTag/AddContentControls.vue";
import AddRichTextContentControl from "@/components/26_structuredDocumentTag/AddRichTextContentControl.vue";
import ComboBoxItem from "@/components/26_structuredDocumentTag/ComboBoxItem.vue";
import GetContentControlProperty from "@/components/26_structuredDocumentTag/GetContentControlProperty.vue";
import LockContentControlContent from "@/components/26_structuredDocumentTag/LockContentControlContent.vue";
import RemoveContentControls from "@/components/26_structuredDocumentTag/RemoveContentControls.vue";
import UpdateCheckBox from "@/components/26_structuredDocumentTag/UpdateCheckBox.vue";

//27_Footnote
import InsertEndnote from "@/components/27_footnote/InsertEndnote.vue";
import InsertFootnote from "@/components/27_footnote/InsertFootnote.vue";
import RemoveFootnote from "@/components/27_footnote/RemoveFootnote.vue";
import SetPositionAndNumberFormat from "@/components/27_footnote/SetPositionAndNumberFormat.vue";

//28_Macros
import DetectAndRemoveVBAMacros from "@/components/28_macros/DetectAndRemoveVBAMacros.vue";
import Macros from "@/components/28_macros/Macros.vue";

//29_Caption
import AddPictureCaption from "@/components/29_caption/AddPictureCaption.vue";
import AddTableCaption from "@/components/29_caption/AddTableCaption.vue";
import PictureCaptionCrossReference from "@/components/29_caption/PictureCaptionCrossReference.vue";
import SetCaptionWithChapterNumber from "@/components/29_caption/SetCaptionWithChapterNumber.vue";
import TableCaptionCrossReference from "@/components/29_caption/TableCaptionCrossReference.vue";

//30_Page
import FixedLayout from "@/components/30_page/FixedLayout.vue";

const filePageId = ref<string>("0");
const handleChildID = (id: any) => {
  filePageId.value = id;
};
</script>

<style lang="scss" scoped>
* {
  padding: 0%;
  margin: 0%;
}
button {
  padding: 5px 10px;
}
.worker_page {
  width: 100%;
  height: 100vh;
  display: flex;
  padding: 2rem;
  .tree_item {
    width: 20%;
    border-right: 1px solid #eeeeee;
  }
  .file_item {
    display: flex;
    flex: 1;
    justify-content: center;
    align-items: center;
    .file-reader div:first-of-type {
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 20px;
      span {
        font-size: 18px;
        font-weight: 500;
      }
    }
    .non-centered {
      align-items: flex-start;
    }
  }
}

pre {
  text-align: left;
  background-color: #f8f8f8;
  padding: 10px;
  border: 1px solid #ccc;
}
</style>
