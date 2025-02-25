<template>
  <span>The following example shows how to replace text in Word document with table.</span>
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
        let inputFileName = 'Template_Docx_1.docx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const doc = wasmModule.Document.Create();

        // Load a document from the virtual file system
        doc.LoadFromFile(inputFileName);

        //Return TextSection by finding the key text string "Christmas Day, December 25".
        let section = doc.Sections.get_Item(0);
        let selection = doc.FindString('Christmas Day, December 25', true, true);

        //Return TextRange from TextSection, then get OwnerParagraph through TextRange.
        let range = selection.GetAsOneRange();
        let paragraph = range.OwnerParagraph;

        //Return the zero-based index of the specified paragraph.
        let body = paragraph.OwnerTextBody;
        let index = body.ChildObjects.IndexOf(paragraph);

        //Create a new table.
        let table = section.AddTable(true);
        table.ResetCells(3, 3);

        //Remove the paragraph and insert table into the collection at the specified index.
        body.ChildObjects.Remove(paragraph);
        body.ChildObjects.Insert(index, table);

        // Define the output file name
        const outputFileName = 'ReplaceTextWithTable.docx';

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
