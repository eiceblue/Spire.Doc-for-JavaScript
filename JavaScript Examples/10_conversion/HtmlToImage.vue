<template>
  <span>The example shows how to convert Html file to image.</span>
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
        let inputFileName = "HtmlToImage.html";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const doc = wasmModule.Document.Create();

        //Load the file from disk.
        doc.LoadFromFile({fileName:inputFileName, fileFormat : wasmModule.FileFormat.Html, validationType:wasmModule.XHTMLValidationType.None});

        //Save to image. You can convert HTML to BMP, JPEG, PNG, GIF, Tiff，etc.
        let image = doc.SaveImageToStreams({pageIndex : 0, imagetype : wasmModule.ImageType.Bitmap});

        // Define the output file name
        const outputFileName = "HtmlToImage-result.png";
        image.Save(outputFileName);
     
        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray],  {type: 'image/png'});

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
