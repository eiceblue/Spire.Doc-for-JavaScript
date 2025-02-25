<template>
  <span>The following example shows how to clone section contents in a Word document. </span>
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
        let inputFileName = 'SectionTemplate.docx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Load the Word document from disk
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        //Get the first section
        let sec1 = doc.Sections.get_Item(0);
        //Get the second section
        let sec2 = doc.Sections.get_Item(1);

        //Loop through the contents of sec1
        for (let i = 0; i < sec1.Body.ChildObjects.Count; i++) {
          //Clone the contents to sec2
          sec2.Body.ChildObjects.Add(sec1.Body.ChildObjects.get(i).Clone());
        }

        // Define the output file name
        const outputFileName = 'CloneSectionContent.docx';

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
