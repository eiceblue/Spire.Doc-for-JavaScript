<template>
  <span>The example shows how to copy document style in a Word document.</span>
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
        let inputFileName_1 = "CopyDocumentStyles1.docx";
        await wasmModule.FetchFileToVFS(inputFileName_1,"",`${import.meta.env.BASE_URL}static/data/`);

                // Load the sample file into the virtual file system (VFS)
        let inputFileName_2 = "CopyDocumentStyles2.docx";
        await wasmModule.FetchFileToVFS(inputFileName_2,"",`${import.meta.env.BASE_URL}static/data/`);

        //Load destination document from disk
        let srcDoc = wasmModule.Document.Create();
        srcDoc.LoadFromFile(inputFileName_1);

        //Load destination document from disk
        let destDoc = wasmModule.Document.Create();
        destDoc.LoadFromFile(inputFileName_2);

        //Get the style collections of source document
        let styles = srcDoc.Styles;

        //Add the style to destination document
        for (let i = 0; i < styles.Count; i++) {
            let style = styles.get_Item(i);
            destDoc.Styles.Add(style);
        }

        // Define the output file name
        const outputFileName = "CopyDocumentStyles_result.docx";

        // Save the document to the specified path
        destDoc.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Clean up resources
        destDoc.Dispose();

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

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
