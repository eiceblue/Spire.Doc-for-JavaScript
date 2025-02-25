<template>
  <span>This sample demonstrates how to hide empty regions when mail merge. </span>
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
        let inputFileName = "HideEmptyRegions.doc";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create word document
        let document = wasmModule.Document.Create();
        document.LoadFromFile(inputFileName);
        let filedNames = ["Contact Name", "Fax", "Date"];
        let filedValues = ["John Smith", "+1 (69) 123456", wasmModule.DateTime.get_Now().Date.ToString()];
        //Set the value to remove paragraphs which contain empty field.
        document.MailMerge.HideEmptyParagraphs = true;
        //Set the value to remove group which contain empty field.
        document.MailMerge.HideEmptyGroup = true;
        document.MailMerge.Execute(filedNames, filedValues);

        // Define the output file name
        const outputFileName = "HideEmptyRegions-result.docx";

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
