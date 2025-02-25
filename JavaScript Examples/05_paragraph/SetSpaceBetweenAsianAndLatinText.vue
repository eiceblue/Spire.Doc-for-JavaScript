<template>
  <span>The sample demonstrates how to set the space between Asian text and Latin text in a Word document. </span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName"> Click here to download the generated file </a>
</template>

<script>
import { ref } from 'vue';

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref('');

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = 'SetSpaceBetweenAsianAndLatinText.docx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        let document = wasmModule.Document.Create();

        //Load the file from disk.
        document.LoadFromFile(inputFileName);

        let para = document.Sections.get(0).Paragraphs.get_Item(0);

        //Set whether to automatically adjust space between Asian text and Latin text
        para.Format.AutoSpaceDE = false;
        //Set whether to automatically adjust space between Asian text and numbers
        para.Format.AutoSpaceDN = true;
        // Define the output file name
        const outputFileName = 'SetSpaceBetweenAsianAndLatinText.docx';

        // Save the document to the specified path
        document.SaveToFile({
          fileName: outputFileName,
          fileFormat: wasmModule.FileFormat.Docx2013,
        });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });

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
