<template>
  <span>The following example shows how to remove section contents in a Word document. </span>
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
        let inputFileName = 'Template_N3.docx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Load the Word document from disk
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        //Loop through all sections
        for (let i = 0; i < doc.Sections.Count; i++) {
          let section = doc.Sections.get(0);
          //Remove header content
          section.HeadersFooters.Header.ChildObjects.Clear();
          //Remove body content
          section.Body.ChildObjects.Clear();
          //Remove footer content
          section.HeadersFooters.Footer.ChildObjects.Clear();
        }

        // Define the output file name
        const outputFileName = 'RemoveSectionContent.docx';

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
