<template>
  <span>The example demonstrates how to create bookmark for a table in a Word document.</span>
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

        //Create word document.
        let document = wasmModule.Document.Create();

        //Add a new section.
        let section = document.AddSection();

        //Create bookmark for a table
        CreateBookmarkForTableBase(document, section);

        // Define the output file name
        const outputFileName = "CreateBookmarkForTable-result.docx";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});
        
        // Clean up resources
        document.Dispose();

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
      function CreateBookmarkForTableBase(doc, section) {
        //Add a paragraph
        let paragraph = section.AddParagraph();

        //Append text for added paragraph
        let txtRange = paragraph.AppendText("The following example demonstrates how to create bookmark for a table in a Word document.");

        //Set the font in italic
        txtRange.CharacterFormat.Italic = true;

        //Append bookmark start
        paragraph.AppendBookmarkStart("CreateBookmark");

        //Append bookmark end
        paragraph.AppendBookmarkEnd("CreateBookmark");

        //Add table
        let table = section.AddTable({showBorder: true});

        //Set the number of rows and columns
        table.ResetCells(2, 2);

        //Append text for table cells
        let range = table.Rows.get(0).Cells.get(0).AddParagraph().AppendText("sampleA");
        range = table.Rows.get(0).Cells.get(1).AddParagraph().AppendText("sampleB");
        range = table.Rows.get(1).Cells.get(0).AddParagraph().AppendText("120");
        range =table.Rows.get(1).Cells.get(1).AddParagraph().AppendText("260");

        //Get the bookmark by index.
        let bookmark = doc.Bookmarks._get_ItemI(0);

        //Get the name of bookmark.
        let bookmarkName = bookmark.Name;

        //Locate the bookmark by name.
        let navigator = wasmModule.BookmarksNavigator.Create(doc);
        navigator.MoveToBookmark(bookmarkName);

        //Add table to TextBodyPart
        let part = navigator.GetBookmarkContent();
        part.BodyItems.Add(table);

        //Replace bookmark cotent with table
        navigator.ReplaceBookmarkContent({bodyPart : part});
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
