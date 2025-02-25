<template>
  <span>Click the following button to read table from text box in a Word document</span>
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
        let inputFileName = "TextBoxTable.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Load a document from the virtual file system
        document.LoadFromFile(inputFileName);

        // Get the first textbox
        let textbox = document.TextBoxes.get_Item(0);

        // Get the first table in the textbox
        let table = textbox.Body.Tables.get_Item(0);

        let str = "";

        // Loop through the paragraphs of the table cells and extract them to a .txt file
        for (let i = 0; i < table.Rows.Count; i++) {
            let row = table.Rows.get_Item(i);
            for (let j = 0; j < row.Cells.Count; j++) {
                let cell = row.Cells.get_Item(j);
                for (let k = 0; k < cell.Paragraphs.Count; k++) {
                    let paragraph = cell.Paragraphs.get_Item(k);
                    str += paragraph.Text + "\t";
                }
            }
            str += "\r\n";
        }

        // Define the output file name
        const outputFileName = "ReadTableFromTextBox.txt";

        // Save the content to the specified path
        wasmModule.FS.writeFile(outputFileName, str);

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "text/plain"});

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
