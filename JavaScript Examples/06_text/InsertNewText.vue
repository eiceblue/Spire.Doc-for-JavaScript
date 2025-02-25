<template>
  <span>The following example shows how to insert new text after a searched text into Word document. </span>
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

        //Load Document
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        //Find all the text string “New Zealand” from the sample document
        let selections = doc.FindAllString('Word', true, true);
        let index = 0;

        //Defines text range
        let range;

        //Insert new text string (New) after the searched text string
        for (let i = 0; i < selections.length; i++) {
          let selection = selections[i];
          range = selection.GetAsOneRange();
          let newrange = wasmModule.TextRange.Create(doc);
          newrange.Text = '(New text)';
          index = range.OwnerParagraph.ChildObjects.IndexOf(range);
          range.OwnerParagraph.ChildObjects.Insert(index + 1, newrange);
        }

        //Find and highlight the newly added text string New
        let text = doc.FindAllString('New text', true, true);
        for (let i = 0; i < text.length; i++) {
          let seletion = text[i];
          seletion.GetAsOneRange().CharacterFormat.HighlightColor = wasmModule.Color.get_Yellow();
        }

        // Define the output file name
        const outputFileName = 'InsertNewText.docx';

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
