<template>
  <span>Click the following button to get the merge status of cells in a Word document.</span>
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
        let inputFileName = "CellMergeStatus.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);          

        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        //Get the first section
        let section = doc.Sections.get_Item(0);

        //Get the first table in the section
        let table = section.Tables.get_Item(0);

        let stringBuidler = [];
        for (let i = 0; i < table.Rows.Count; i++) {
            let tableRow = table.Rows.get(i);
            for (let j = 0; j < tableRow.Cells.Count; j++) {
                let tableCell = tableRow.Cells.get(j);
                let verticalMerge = tableCell.CellFormat.VerticalMerge;
                let horizontalMerge = tableCell.GridSpan;
                if (verticalMerge === wasmModule.CellMerge.None && horizontalMerge === 1) {
                    stringBuidler.push("Row " + i + ", cell " + j + ": ");
                    stringBuidler.push("This cell isn't merged.\n");
                } else {
                    stringBuidler.push("Row " + i + ", cell " + j + ": ");
                    stringBuidler.push("This cell is merged.\n");
                }
            }
            stringBuidler.push("\n");
        }

        // Define the output file name
        const outputFileName = "CellMergeStatus_output.txt";

        //Save and launch document
        FS.writeFile(outputFileName, stringBuidler.join('\n'))

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
