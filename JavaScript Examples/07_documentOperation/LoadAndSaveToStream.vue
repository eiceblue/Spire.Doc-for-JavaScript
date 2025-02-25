<template>
  <span>Click the following button to load a document from stream and save a document to stream.</span>
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
        let inputFileName = "Template.docx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        let stream = wasmModule.Stream.CreateByFile(inputFileName);

        // Load the entire document into memory.
        let doc = wasmModule.Document.Create();
        doc.LoadFromStream({ stream: stream, fileFormat: wasmModule.FileFormat.Auto })

        // You can close the stream now, it is no longer needed because the document is in memory.
        stream.Close();

        // Define the output file name
        const outputFileName = "LoadAndSaveToStream_out.rtf";

        // Convert the document to a different format and save to stream.
        let newStream = wasmModule.Stream.CreateByFile(outputFileName);
        doc.SaveToStream({ stream: newStream, fileFormat: wasmModule.FileFormat.Rtf });

        FS.writeFile(outputFileName, newStream.ToArray());

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/rtf" });

        // Clean up resources
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
