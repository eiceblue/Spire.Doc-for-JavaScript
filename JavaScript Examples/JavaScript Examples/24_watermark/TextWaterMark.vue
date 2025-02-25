<template>
  <span>Click the following button to insert a text watermark into Word document</span>
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
        let inputFileName = "Template.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Load a document from the virtual file system
        document.LoadFromFile(inputFileName);

        // Get the first section
        let section = document.Sections.get_Item(0);

        // Create a new TextWatermark instance
        let txtWatermark = wasmModule.TextWatermark.Create();

        // Set the text for the watermark
        txtWatermark.Text = "E-iceblue";

        // Set the font size for the watermark text
        txtWatermark.FontSize = 95;

        // Set the color of the watermark text to blue
        txtWatermark.Color = wasmModule.Color.get_Blue();

        // Set the layout of the watermark to diagonal
        txtWatermark.Layout = wasmModule.WatermarkLayout.Diagonal;

        // Assign the created text watermark to the document
        section.Document.Watermark = txtWatermark;
        
        // Define the output file name
        const outputFileName = "TextWaterMark.docx";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

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
