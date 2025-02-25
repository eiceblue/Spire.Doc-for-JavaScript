<template>
  <span>The following example shows how to extract comments from Word document</span>
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

        // Load the input file into the virtual file system (VFS)
        const inputFileName = "CommentSample.docx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a new document object  
        let doc = wasmModule.Document.Create();
        // Load file from VFS
        doc.LoadFromFile(inputFileName);

        // Initialize an empty array to store the extracted text from the comments
        let SB = [];

        //Traverse all comments
        for (let i = 0; i < doc.Comments.Count; i++) {
          let comment = doc.Comments.get_Item(i);
          for (let j = 0; j < comment.Body.Paragraphs.Count; j++) {
            let p = comment.Body.Paragraphs.get_Item(j);
            SB.push(p.Text + "\n");
          }
        }

        // Convert the SB array into a Blob object
        const outputFileName = 'GetChartDataPointValues.txt';
        const modifiedFile = new Blob([SB.toString()], { type: "text/plain;charset=utf-8" });

        // Dispose of the document object to free resources
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