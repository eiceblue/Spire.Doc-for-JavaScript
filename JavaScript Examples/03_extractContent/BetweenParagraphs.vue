<template>
  <span>The following example demonstrates how to extract content between paragraphs. </span>
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

        //Create a destination document
        let destinationDoc = wasmModule.Document.Create();

        //Add a section
        let section = destinationDoc.AddSection();

        //Extract content between the first paragraph to the third paragraph
        ExtractBetweenParagraphs(doc, destinationDoc, 1, 3);

        // Define the output file name
        const outputFileName = 'BetweenParagraphs.docx';

        // Save the document to the specified path
        destinationDoc.SaveToFile({
          fileName: outputFileName,
          fileFormat: wasmModule.FileFormat.Docx2013,
        });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });

        // Clean up resources
        destinationDoc.Dispose();

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
    };
    function ExtractBetweenParagraphs(sourceDocument, destinationDocument, startPara, endPara) {
      //Extract the content
      for (let i = startPara - 1; i < endPara; i++) {
        //Clone the ChildObjects of source document
        let doobj = sourceDocument.Sections.get_Item(0).Body.ChildObjects.get_Item(i).Clone();

        //Add to destination document
        destinationDocument.Sections.get_Item(0).Body.ChildObjects.Add(doobj);
      }
    }
    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
