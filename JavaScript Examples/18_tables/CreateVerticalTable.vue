<template>
  <span>Click the following button to create vertical table at one side of the Word document</span>
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

        //Create Word document.
        let document = wasmModule.Document.Create();

        //Add a new section.
        let section = document.AddSection();

        //Add a table with rows and columns and set the text for the table.
        let table = section.AddTable();
        table.ResetCells(1, 1);
        let cell = table.Rows.get(0).Cells.get(0);
        table.Rows.get(0).Height = 150;
        cell.AddParagraph().AppendText("Draft copy in vertical style");

        //Set the TextDirection for the table to RightToLeftRotated.
        cell.CellFormat.TextDirection = wasmModule.TextDirection.RightToLeftRotated;

        //Set the table format.
        table.TableFormat.WrapTextAround = true;
        table.TableFormat.Positioning.VertRelationTo = wasmModule.VerticalRelation.Page;
        table.TableFormat.Positioning.HorizRelationTo = wasmModule.HorizontalRelation.Page;
        table.TableFormat.Positioning.HorizPosition = section.PageSetup.PageSize.Width - table.Width;
        table.TableFormat.Positioning.VertPosition = 200;

        // Define the output file name
        const outputFileName = "CreateVerticalTable_output.docx";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

        // Clean up resources
        document.Close();
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
