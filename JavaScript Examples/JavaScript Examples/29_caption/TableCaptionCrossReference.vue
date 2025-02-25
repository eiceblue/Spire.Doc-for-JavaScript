<template>
  <span>Click the following button to create cross-reference for table caption</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from "vue";

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF","/Library/Fonts/",`${import.meta.env.BASE_URL}static/font/`);
        
        // Create a new document
        const document = wasmModule.Document.Create();

        // Get the first section of the document
        let section = document.AddSection();

        // Add a new table to the section and reset its cells to 2 rows and 3 columns
        let table = section.AddTable({showBorder:true});
        table.ResetCells(2, 3);

        // Add a caption for the table below the item
        let captionParagraph = table.AddCaption("Table", wasmModule.CaptionNumberingFormat.Number, wasmModule.CaptionPosition.BelowItem);

        // Define a bookmark name for the table
        let bookmarkName = "Table_1";

        // Add a paragraph in the section for the bookmark
        let paragraph = section.AddParagraph();
        paragraph.AppendBookmarkStart(bookmarkName); 
        paragraph.AppendBookmarkEnd(bookmarkName);  

        // Create a bookmark navigator to manage the bookmark
        let navigator = wasmModule.BookmarksNavigator.Create(document);
        navigator.MoveToBookmark(bookmarkName); 
        let part = navigator.GetBookmarkContent(); 
        part.BodyItems.Clear(); 
        part.BodyItems.Add(captionParagraph);
        navigator.ReplaceBookmarkContent({ bodyPart: part }); 

        // Create a field for cross-referencing the bookmark
        let field = wasmModule.Field.Create(document);
        field.Type = wasmModule.FieldType.FieldRef; 
        field.Code = "REF Table_1 \\p \\h"; 

        // Add line breaks to create space
        for (let i = 0; i < 3; i++) {
            paragraph.AppendBreak(wasmModule.BreakType.LineBreak); // Add line breaks
        }

        // Add a new paragraph for the caption cross-reference
        paragraph = section.AddParagraph();
        let range = paragraph.AppendText("This is a table caption cross-reference, "); 
        range.CharacterFormat.FontSize = 14; 
        paragraph.ChildObjects.Add(field); 

        // Create a field separator for formatting
        let fieldSeparator = wasmModule.FieldMark.Create(document, wasmModule.FieldMarkType.FieldSeparator);
        paragraph.ChildObjects.Add(fieldSeparator); 

        // Create a text range for the display text of the cross-reference
        let tr = wasmModule.TextRange.Create(document);
        tr.Text = "Table 1"; 
        tr.CharacterFormat.FontSize = 14;
        tr.CharacterFormat.TextColor = wasmModule.Color.get_DeepSkyBlue();
        paragraph.ChildObjects.Add(tr); 

        // Create a field end mark to close the field
        let fieldEnd = wasmModule.FieldMark.Create(document, wasmModule.FieldMarkType.FieldEnd);
        paragraph.ChildObjects.Add(fieldEnd); 

        // Update all fields in the document
        document.IsUpdateFields = true; 
        
        // Define the output file name
        const outputFileName = "TableCaptionCrossReference.docx";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

        // Clean up resources
        document.Dispose();

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
