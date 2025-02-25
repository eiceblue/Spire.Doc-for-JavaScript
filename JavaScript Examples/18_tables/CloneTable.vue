<template>
  <span>Click the following button to insert an existing table by cloning in a Word document.</span>
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
        let inputFileName = "TableTemplate.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`); 

        //Load the document
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        //Get the first section
        let se = doc.Sections.get_Item(0);

        //Get the first table
        let original_Table = se.Tables.get_Item(0);

        //Copy the existing table to copied_Table via Table.clone()
        let copied_Table = original_Table.Clone();
        let st = ["Spire.Presentation for JS", "A professional " +
        "PowerPoint® compatible library that enables developers to create, read, " +
        "write, modify, convert and Print PowerPoint documents on any JS framework."];
        //Get the last row of table
        let lastRow = copied_Table.Rows.get(copied_Table.Rows.Count - 1);
        //Change last row data
        for (let i = 0; i < lastRow.Cells.Count - 1; i++) {
            lastRow.Cells.get(i).Paragraphs.get_Item(0).Text = st[i];
        }
        //Add copied_Table in section
        se.Tables.Add(copied_Table);

        // Define the output file name
        const outputFileName = "CloneTable_output.docx";

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
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
