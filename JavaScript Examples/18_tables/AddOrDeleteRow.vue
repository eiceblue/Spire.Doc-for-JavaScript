<template>
  <span>Click the following button to add or delete row of table in a Word document.</span>
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

        //Create a document
        let document = wasmModule.Document.Create();
        //Load file
        document.LoadFromFile(inputFileName);
        let section = document.Sections.get_Item(0);
        let table = section.Tables.get_Item(0);

        //Delete the seventh row
        table.Rows.RemoveAt(7);

        //Add a row and insert it into specific position
        let row = wasmModule.TableRow.Create(document);
        for (let i = 0; i < table.Rows.get(0).Cells.Count; i++) {
            let tc = row.AddCell();
            let paragraph = tc.AddParagraph();
            paragraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
            paragraph.AppendText("Added");
        }
        table.Rows.Insert(2, row);
        //Add a row at the end of table
        table.AddRow();

        // Define the output file name
        const outputFileName = "AddOrDeleteRow_output.docx";

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
