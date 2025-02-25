<template>
  <span>The following example demonstrates how to find and highlight in a Word document.</span>
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
        let inputFileName = 'Sample.docx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const doc = wasmModule.Document.Create();

        // Load a document from the virtual file system
        doc.LoadFromFile(inputFileName);

        // Find all occurrences of the string "word" in the document
        let textSelections = doc.FindAllString('word', false, true);

        // Iterate through all found text selections
        for (let i = 0; i < textSelections.length; i++) {
          let selection = textSelections[i];

          // Set the highlight color of the selected text to yellow
          selection.GetAsOneRange().CharacterFormat.HighlightColor = wasmModule.Color.get_Yellow();
        }

        // Define the output file name
        const outputFileName = 'FindAndHighlight.docx';

        // Save the document to the specified path
        doc.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });

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
