<template>
  <span>Click the following button to get row and cell index of table in a Word document</span>
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
        let inputFileName = "ReplaceTextInTable.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);         

        //Load Word from disk
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        //Get the first section
        let section = doc.Sections.get_Item(0);

        //Get the first table in the section
        let table = section.Tables.get_Item(0);

        let content = [];

        //Get table collections
        let collections = section.Tables;

        //Get the table index
        let tableIndex = collections.IndexOf(table);

        //Get the index of the last table row
        let row = table.LastRow;
        let rowIndex = row.GetRowIndex();

        //Get the index of the last table cell
        let cell = row.LastChild;
        let cellIndex = cell.GetCellIndex();

        //Append these information into content
        content.push("Table index is " + tableIndex.toString() + "\n");
        content.push("Row index is " + rowIndex.toString() + "\n");
        content.push("Cell index is " + cellIndex.toString() + "\n");


        // Define the output file name
        const outputFileName = "GetRowCellIndex_output.txt";

        FS.writeFile(outputFileName, content.join('\n'));

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "text/plain"});

        // Clean up resources
        doc.Close();
        doc.Dispose();

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
