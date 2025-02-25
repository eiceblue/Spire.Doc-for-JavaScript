<template>
  <span>The example shows how to change the locale when mail merge. </span>
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
        let inputFileName = "MailMerge.doc";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Load word document
        let document = wasmModule.Document.Create();
        document.LoadFromFile(inputFileName);

        // Store the current culture so it can be set back once mail merge is complete.
        const now = new Date();
        const datestr = new Intl.DateTimeFormat("de-DE",
            {year:'numeric',
                month: '2-digit',
                day:'2-digit',
                hour: "2-digit",
                minute: '2-digit',
                second: '2-digit'}).format(now);

        let fieldNames = ["Contact Name", "Fax", "Date"];
        let fieldValues = ["John Smith", "+1 (69) 123456", datestr];
        document.MailMerge.Execute(fieldNames, fieldValues);

        // Define the output file name
        const outputFileName = "ChangeLocale-result.docx";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Clean up resources
        document.Dispose();
        
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
