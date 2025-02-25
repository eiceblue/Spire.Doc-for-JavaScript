<template>
  <span>Click the following button to  create table using html in a Word document</span>
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

        //HTML string
        let HTML = "<table border='2px'>" +
            "<tr>" +
            "<td>Row 1, Cell 1</td>" +
            "<td>Row 1, Cell 2</td>" +
            "</tr>" +
            "<tr>" +
            "<td>Row 2, Cell 2</td>" +
            "<td>Row 2, Cell 2</td>" +
            "</tr>" +
            "</table>";

        //Create a Word document
        let document = wasmModule.Document.Create();

        //Add a section
        let section = document.AddSection();

        //Add a paragraph and append html string
        section.AddParagraph().AppendHTML(HTML);

        // Define the output file name
        const outputFileName = "CreateTableFromHTML_output.docx";

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
