<template>
  <span>Click the following button to retrieve variables in a Word document.</span>
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

        //Load the file from disk.
        document.LoadFromFile(inputFileName);

        //Retrieve name of the variable by index.
        let s1 = document.Variables.GetNameByIndex(0);

        //Retrieve value of the variable by index.
        let s2 = document.Variables.GetValueByIndex(0);

        //Retrieve the value of the variable by name.
        let s3 = document.Variables.get_Item("A1");

        let content = [];
        content.push("The name of the variable retrieved by index 0 is: " + s1 + "\n");
        content.push("The vaule of the variable retrieved by index 0 is: " + s2 + "\n");
        content.push("The vaule of the variable retrieved by name \"A1\" is: " + s3 + "\n");

        // Define the output file name
        const outputFileName = "RetrieveVariables_out.txt";

        // Save the document to the specified path
        FS.writeFile(outputFileName, content.join(""))

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
