<template>
  <span>The example demonstrates how to convert Word to PDF and embed noninstalled fonts.</span>
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
        await wasmModule.FetchFileToVFS("PT_Serif-Caption-Web-Regular.ttf","/Library/Fonts/",`${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "ToPDFTemplate.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();
        document.LoadFromFile(inputFileName);

        // Embed the non-installed fonts.
        let parms = wasmModule.ToPdfParameterList.Create();
        let fonts = [wasmModule.PrivateFontPath.Create("PT Serif Caption", "PT_Serif-Caption-Web-Regular.ttf")];
        parms.PrivateFontPaths = fonts;

        // Define the output file name
        const outputFileName = "EmbedNoninstalledFonts-result.pdf";
        //Save doc file to pdf.
        document.SaveToFile({fileName: outputFileName, paramList: parms});
        
        // Clean up resources
        document.Dispose();

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: 'application/pdf'});

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
