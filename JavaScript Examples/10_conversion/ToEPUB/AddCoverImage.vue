<template>
  <span>The example shows how to convert Word to ePub with cover image.
</span>
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
        let inputFileName1 = "AddCoverImage.doc";
        await wasmModule.FetchFileToVFS(inputFileName1,"",`${import.meta.env.BASE_URL}static/data/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName2 = "Cover.png";
        await wasmModule.FetchFileToVFS(inputFileName2,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Load a document from the virtual file system
        document.LoadFromFile(inputFileName1);

        let picture = wasmModule.DocPicture.Create(document);
        picture.LoadImage(inputFileName2);

        // Define the output file name
        const outputFileName = "AddCoverImage-result.epub";

        // Save the document to the specified path
        document.SaveToEpub(outputFileName, picture);

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type:"application/epub+zip"});

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
