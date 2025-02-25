<template>
  <span>The following example shows how to remove empty lines from a Word document. </span>
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
        let inputFileName = 'Template_Docx_3.docx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create Word document.
        let document = wasmModule.Document.Create();

        //Load the file from disk.
        document.LoadFromFile(inputFileName);

        //Traverse every section on the word document and remove the null and empty paragraphs.
        for (let i = 0; i < document.Sections.Count; i++) {
          let section = document.Sections.get_Item(i);
          for (let j = 0; j < section.Body.ChildObjects.Count; j++) {
            if (section.Body.ChildObjects.get(j).DocumentObjectType === wasmModule.DocumentObjectType.Paragraph) {
              let str = section.Body.ChildObjects.get(j).Text.trim();
              if (str.length === 0) {
                section.Body.ChildObjects.Remove(section.Body.ChildObjects.get(j));
                j--;
              }
            }
          }
        }

        // Define the output file name
        const outputFileName = 'RemoveEmptyLines.docx';

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
