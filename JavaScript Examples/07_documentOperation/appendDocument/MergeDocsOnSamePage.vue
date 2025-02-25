<template>
  <span>Click the following button to merge documents on same page.</span>
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
        let inputFileName1 = "Insert.docx";
        await wasmModule.FetchFileToVFS(inputFileName1, "", `${import.meta.env.BASE_URL}static/data/`);
        let inputFileName2 = "TableOfContent.docx";
        await wasmModule.FetchFileToVFS(inputFileName2, "", `${import.meta.env.BASE_URL}static/data/`);

        //Create a document
        let document = wasmModule.Document.Create();

        //Load the source document
        document.LoadFromFile(inputFileName1);

        //Clone a destination  document
        let destinationDocument = wasmModule.Document.Create();

        //Load the destination document
        destinationDocument.LoadFromFile(inputFileName2);

        //Traverse sections
        for (let i = 0; i < document.Sections.Count; i++) {
          let section = document.Sections.get_Item(i);
          //Traverse body ChildObjects
          for (let j = 0; j < section.Body.ChildObjects.Count; j++) {
            let obj = section.Body.ChildObjects.get(j);
            //Clone to destination document at the same page
            destinationDocument.Sections.get(0).Body.ChildObjects.Add(obj.Clone());
          }
        }

        // Define the output file name
        const outputFileName = "MergeDocsOnSamePage_out.docx";

        // Save the document to the specified path
        destinationDocument.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

        // Clean up resources
        destinationDocument.Dispose();
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
