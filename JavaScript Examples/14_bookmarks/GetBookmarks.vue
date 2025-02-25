<template>
  <span>The example shows how to get bookmarks by index and name in a Word document. </span>
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
        let inputFileName = "GetBookmarks.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Load a document from the virtual file system
        document.LoadFromFile(inputFileName);

        //Get the bookmark by index.
        let bookmark1 = document.Bookmarks._get_ItemI(0);

        //Get the bookmark by name.
        let bookmark2 = document.Bookmarks._get_Item("Test2");

        //Create StringBuilder to save
        let content = [];

        //Set string format for displaying
        let result = "The bookmark obtained by index is " + bookmark1.Name + ".\r\nThe bookmark obtained by name is " + bookmark2.Name + ".\n";

        //Add result string to StringBuilder
        content.push(result);

        // Define the output file name
        const outputFileName = "GetBookmarks-result.txt";

        //Write the contents in a TXT file
        wasmModule.FS.writeFile(outputFileName, content.join("\n"));
        document.Close();

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "text/plain"});

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
