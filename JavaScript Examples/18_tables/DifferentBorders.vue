<template>
  <span>Click the following button to set different borders on a table in a Word document</span>
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

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "TableSample.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`); 

        //Open a Word document as template
        let document = wasmModule.Document.Create();
        document.LoadFromFile(inputFileName);

        let table = document.Sections.get_Item(0).Tables.get_Item(0);

        //Set borders of table
        setTableBorders(table);

        //Set borders of cell
        setCellBorders(table.Rows.get(2).Cells.get(0));

        // Define the output file name
        const outputFileName = "DifferentBorders_output.docx";

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

    function setTableBorders(table) {
    table.TableFormat.Borders.BorderType = wasmModule.BorderStyle.Single;
    table.TableFormat.Borders.LineWidth = 3.0;
    table.TableFormat.Borders.Color = wasmModule.Color.get_Red();
}

 function setCellBorders(tableCell) {
     tableCell.CellFormat.Borders.BorderType = wasmModule.BorderStyle.DotDash;
     tableCell.CellFormat.Borders.LineWidth = 1.0;
     tableCell.CellFormat.Borders.Color = wasmModule.Color.get_Green();
 }

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
