<template>
  <span>The following example demonstrates how to extract content between a paragraph and a table. </span>
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
        let inputFileName = 'IncludingTable.docx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const doc = wasmModule.Document.Create();

        // Load a document from the virtual file system
        doc.LoadFromFile(inputFileName);

        //Create a destination document
        let destinationDoc = wasmModule.Document.Create();

        //Add a section
        let destinationSection = destinationDoc.AddSection();

        //Extract the content from the first paragraph to the first table
        ExtractByTable(doc, destinationDoc, 1, 1);

        // Define the output file name
        const outputFileName = 'FromParagraphToTable.docx';

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
    function ExtractByTable(sourceDocument, destinationDocument, startPara, tableNo) {
      //Get the table from the source document
      let table = sourceDocument.Sections.get(0).Tables.get_Item(tableNo - 1);

      //Get the table index
      let index = sourceDocument.Sections.get_Item(0).Body.ChildObjects.IndexOf(table);
      for (let i = startPara - 1; i <= index; i++) {
        //Clone the ChildObjects of source document
        let doobj = sourceDocument.Sections.get(0).Body.ChildObjects.get(i).Clone();

        //Add to destination document
        destinationDocument.Sections.get(0).Body.ChildObjects.Add(doobj);
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
