<template>
  <span>Click the following button to add one section from other doc.</span>
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
        let inputFileName1 = "Template_Docx_1.docx";
        await wasmModule.FetchFileToVFS(inputFileName1, "", `${import.meta.env.BASE_URL}static/data/`);
        let inputFileName2 = "Sample_two sections.docx";
        await wasmModule.FetchFileToVFS(inputFileName2, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const TarDoc = wasmModule.Document.Create();
        TarDoc.LoadFromFile(inputFileName1);

        //Open a Word document as source document
        let SouDoc = wasmModule.Document.Create();
        SouDoc.LoadFromFile(inputFileName2);

        //Get the second section from source document
        let Ssection = SouDoc.Sections.get(1);

        //Add the section in target document
        TarDoc.Sections.Add(Ssection.Clone());

        // Define the output file name
        const outputFileName = "AddSectionFromOtherDoc_out.docx";

        // Save the document to the specified path
        TarDoc.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

        // Clean up resources
        TarDoc.Dispose();
        SouDoc.Dispose();

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
