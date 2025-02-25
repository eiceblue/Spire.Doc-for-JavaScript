<template>
  <span>The example demonstrates how to convert Word to PDF with password.</span>
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

        let inputFileName = "ToPDFTemplate.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();
        document.LoadFromFile(inputFileName);


        
        //create a parameter
        let toPdf = wasmModule.ToPdfParameterList.Create();

        //set the password
        let password = "E-iceblue";
        toPdf.PdfSecurity.Encrypt("password", password, wasmModule.PdfPermissionsFlags.Default, wasmModule.PdfEncryptionKeySize.Key128Bit);
       
        // Define the output file name
        const outputFileName = "ToPdfWithPassword-result.pdf";

        //save doc file.
        document.SaveToFile({fileName: outputFileName, paramList: toPdf});
        
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
