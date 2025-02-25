<template>
  <span>The example shows how to convert a document to bytes</span>
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
        let inputFileName = "ConvertDocToByte.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const doc = wasmModule.Document.Create();

        // Load a document from the virtual file system
        doc.LoadFromFile(inputFileName);

        // Create a new memory stream.
        let outStream = wasmModule.Stream.Create();
        // Save the document to stream.
        doc.SaveToStream({stream: outStream, fileFormat: wasmModule.FileFormat.Docx});

        // Convert the document to bytes.
        let docBytes = outStream.ToArray();

        // The bytes are now ready to be stored/transmitted.

        // Now reverse the steps to load the bytes back into a document object.
        let inStream = wasmModule.Stream.CreateByBytes(docBytes);

        // Load the stream into a new document object.
        let newDoc = wasmModule.Document.Create();
        newDoc.LoadFromStream({stream: inStream, fileFormat: wasmModule.FileFormat.Auto});

        // Define the output file name
        const outputFileName = "ConvertDocToByte-result.docx";
        
        // Save the document to the specified path
        newDoc.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Clean up resources
        doc.Dispose();
        newDoc.Dispose();

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

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
