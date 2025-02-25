<template>
  <span>Click the following button to set Word document properties.</span>
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
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "Summary_of_Science.doc";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();
        document.LoadFromFile(inputFileName);

        // Set document properties
        document.BuiltinDocumentProperties.Title = "Document Demo Document";
        document.BuiltinDocumentProperties.Subject = "demo";
        document.BuiltinDocumentProperties.Author = "James";
        document.BuiltinDocumentProperties.Company = "e-iceblue";
        document.BuiltinDocumentProperties.Manager = "Jakson";
        document.BuiltinDocumentProperties.Category = "Doc Demos";
        document.BuiltinDocumentProperties.Keywords = "Document, Property, Demo";
        document.BuiltinDocumentProperties.Comments = "This document is just a demo.";

        // Define the output file name
        const outputFileName = "DocumentProperty_out.docx";

        // Save the document to the specified path
        document.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

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
