<template>
  <span>Click the following button to add or remove columns of table in a Word document.</span>
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
        let inputFileName = "Template_N2.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);        

        //Load the document from disk.
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        //Access the first section
        let section = doc.Sections.get_Item(0);

        //Access the first table
        let table = section.Tables.get_Item(0);

        //Add a blank column
        let columnIndex1 = 0;
        AddColumn(table, columnIndex1);

        //Remove a column
        let columnIndex2 = 2;
        RemoveColumn(table, columnIndex2);

        // Define the output file name
        const outputFileName = "AddOrRemoveColumn_output.docx";

        // Save the document to the specified path
        doc.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

        // Clean up resources
        doc.Close();
        doc.Dispose();

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }

      function AddColumn(table, columnIndex) {
    for (let r = 0; r < table.Rows.Count; r++) {
        let addCell =  wasmModule.TableCell.Create(table.Document);
        table.Rows.get(r).Cells.Insert(columnIndex, addCell);
    }
}
function RemoveColumn(table, columnIndex) {
    for (let r = 0; r < table.Rows.Count; r++) {
        table.Rows.get(r).Cells.RemoveAt(columnIndex);
    }
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
