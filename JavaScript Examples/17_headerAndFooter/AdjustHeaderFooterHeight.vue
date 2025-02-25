<template>
  <span>Click the following button to adjust the height of headers and footers in a Word document.</span>
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
        let inputFileName = "HeaderAndFooter.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Load the document
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        //Get the first section
        let section = doc.Sections.get(0);

        //Adjust the height of headers in the section
        section.PageSetup.HeaderDistance = 100;

        //Adjust the height of footers in the section
        section.PageSetup.FooterDistance = 100;

        let outputFileName = "AdjustHeaderFooterHeight_output.docx"
        //Save and launch document
        doc.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});


        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

        // Clean up resources  
        doc.Close();
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
