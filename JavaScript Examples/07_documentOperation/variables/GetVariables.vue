<template>
  <span>Click the following button to get variables in a Word document .</span>
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
        let inputFileName = "Template_Docx_6.docx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        //Create Word document.
        let document = wasmModule.Document.Create();

        //Load the file.
        document.LoadFromFile(inputFileName);
        let stringBuilder = [];
        stringBuilder.push("This document has following variables:\n");
        for (let i = 0; i < document.Variables.Count; i++) {
          let name = document.Variables.GetNameByIndex(i);
          let value = document.Variables.GetValueByIndex(i);
          stringBuilder.push("Name: " + name + ", " + "Value: " + value + "\n");
        }

        // Define the output file name
        const outputFileName = "GetVariables_out.txt";

        // Save the document to the specified path
        FS.writeFile(outputFileName, stringBuilder.join(""))

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: 'text/plain' });

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
