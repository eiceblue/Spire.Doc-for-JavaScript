<template>
  <span>Click the following button to replace text by string or regular expression in a table</span>
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
        let inputFileName = "ReplaceTextInTable.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`); 

        //Load Word from disk
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        //Get the first section
        let section = doc.Sections.get_Item(0);

        //Get the first table in the section
        let table = section.Tables.get_Item(0);

        //Define a regular expression to match the {} with its content
        let regex = wasmModule.Regex.Create("{[^\\}]+\\}", wasmModule.RegexOptions.None);

        //Replace the text of table with regex
        table.Replace({pattern: regex, replace: "E-iceblue"});

        //Replace old text with new text in table
        table.Replace({given:"Beijing", replace: "Component", caseSensitive: false, wholeWord: true});

        // Define the output file name
        const outputFileName = "ReplaceTextInTable_output.docx";

        // Save the document to the specified path
        doc.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

        // Clean up resources
        doc.Close();
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
