<template>
  <span>The following example shows how to replace text with a Word document. </span>
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

        // Load the sample files into the virtual file system (VFS)
        let inputFileName1 = 'Text2.docx';
        await wasmModule.FetchFileToVFS(inputFileName1, '', `${import.meta.env.BASE_URL}static/data/`);

        let inputFileName2 = 'Text1.docx';
        await wasmModule.FetchFileToVFS(inputFileName2, '', `${import.meta.env.BASE_URL}static/data/`);
        //Load a template document
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName1);

        //Load another document to replace text
        let replaceDoc = wasmModule.Document.Create();
        replaceDoc.LoadFromFile(inputFileName2);
        //Replace specified text with the other document
        doc.Replace({
          matchString: 'Document1',
          matchDoc: replaceDoc,
          caseSensitive: false,
          wholeWord: true,
        });

        // Define the output file name
        const outputFileName = 'ReplaceWithDocument.docx';

        // Save the document to the specified path
        doc.SaveToFile({
          fileName: outputFileName,
          fileFormat: wasmModule.FileFormat.Docx2013,
        });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });

        // Clean up resources
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
