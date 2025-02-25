<template>
  <span>Click the following button to copy content from one Word document to another.</span>
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
        let inputFileName2 = "Target.docx";
        await wasmModule.FetchFileToVFS(inputFileName2, "", `${import.meta.env.BASE_URL}static/data/`);

        //Initialize a new object of Document class and load the source document.
        let sourceDoc = wasmModule.Document.Create();
        sourceDoc.LoadFromFile(inputFileName1);

        //Initialize another object to load target document.
        let destinationDoc = wasmModule.Document.Create();
        destinationDoc.LoadFromFile(inputFileName2);

        //Copy content from source file and insert them to the target file.
        for (let i = 0; i < sourceDoc.Sections.Count; i++) {
          let sec = sourceDoc.Sections.get_Item(i);
          for (let j = 0; j < sec.Body.ChildObjects.Count; j++) {
            let obj = sec.Body.ChildObjects.get(j);
            destinationDoc.Sections.get(0).Body.ChildObjects.Add(obj.Clone());
          }
        }

        // Define the output file name
        const outputFileName = "CopyContentToAnotherDoc_out.docx";

        // Save the document to the specified path
        destinationDoc.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

        // Clean up resources
        destinationDoc.Dispose();
        sourceDoc.Dispose();

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
