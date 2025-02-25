<template>
  <span>Click the following button to keep same format of source document when merging.</span>
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
        let inputFileName1 = "Template_N2.docx";
        await wasmModule.FetchFileToVFS(inputFileName1, "", `${import.meta.env.BASE_URL}static/data/`);
        let inputFileName2 = "Template_N3.docx";
        await wasmModule.FetchFileToVFS(inputFileName2, "", `${import.meta.env.BASE_URL}static/data/`);

        //Load the source document from disk
        let srcDoc = wasmModule.Document.Create();
        srcDoc.LoadFromFile(inputFileName1);

        //Load the destination document from disk
        let destDoc = wasmModule.Document.Create();
        destDoc.LoadFromFile(inputFileName2);

        //Keep same format of source document
        srcDoc.KeepSameFormat = true;

        //Copy the sections of source document to destination document
        for (let i = 0; i < srcDoc.Sections.Count; i++) {
          let section = srcDoc.Sections.get_Item(i);
          destDoc.Sections.Add(section.Clone());
        }

        // Define the output file name
        const outputFileName = "KeepSameFormat_out.docx";

        // Save the document to the specified path
        destDoc.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

        // Clean up resources
        destDoc.Dispose();
        srcDoc.Dispose();

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
