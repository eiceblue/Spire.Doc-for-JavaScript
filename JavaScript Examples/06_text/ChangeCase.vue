<template>
  <span>The following example shows how to change text case to capital letters in a Word document. </span>
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
        let inputFileName = 'Text1.docx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new document and load from file;
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);
        let textRange;
        //Get the first paragraph and set its CharacterFormat to AllCaps
        let para1 = doc.Sections.get(0).Paragraphs.get_Item(1);

        for (let i = 0; i < para1.ChildObjects.Count; i++) {
          let obj = para1.ChildObjects.get(i);
          if (obj instanceof wasmModule.TextRange) {
            textRange = obj;
            textRange.CharacterFormat.AllCaps = true;
          }
        }

        //Get the third paragraph and set its CharacterFormat to IsSmallCaps
        let para2 = doc.Sections.get(0).Paragraphs.get_Item(3);
        for (let j = 0; j < para2.ChildObjects.Count; j++) {
          let obj = para2.ChildObjects.get(j);
          if (obj instanceof wasmModule.TextRange) {
            obj.CharacterFormat.IsSmallCaps = true;
          }
        }

        // Define the output file name
        const outputFileName = 'ChangeCase.docx';

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
