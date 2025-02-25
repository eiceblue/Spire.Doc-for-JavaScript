<template>
  <span>Click the following button to update the last saved date of the Word document.</span>
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

        function LocalTimeToGreenwishTime(lacalTime) {
          let localTimeZone = wasmModule.TimeZone.get_CurrentTimeZone();
          let timeSpan = localTimeZone.GetUtcOffset(lacalTime);
          let greenwishTime = lacalTime - timeSpan;
          return greenwishTime;
        }

        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "Template.docx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();
        document.LoadFromFile(inputFileName);

        //Update the last saved date
        document.BuiltinDocumentProperties.LastSaveDate = LocalTimeToGreenwishTime(wasmModule.DateTime.get_Now());

        // Define the output file name
        const outputFileName = "UpdateLastSavedDate_out.docx";

        // Save the document to the specified path
        document.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

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
