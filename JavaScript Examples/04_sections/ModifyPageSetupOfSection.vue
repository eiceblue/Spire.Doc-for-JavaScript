<template>
  <span>The following example shows how to modify page setup of all sections or one section. </span>
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
        let inputFileName = 'Template_N2.docx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Load Word from disk
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        //Loop through all sections
        for (let i = 0; i < doc.Sections.Count; i++) {
          let section = doc.Sections.get(i);
          //Modify the margins
          section.PageSetup.Margins = wasmModule.MarginsF.Create(100, 80, 100, 80);
          //Modify the page size
          section.PageSetup.PageSize = wasmModule.PageSize.Letter;
        }

        // Or only modify one section
        // For example, modify the page setup of the first section
        let section0 = doc.Sections.get_Item(0);
        section0.PageSetup.Margins = wasmModule.MarginsF.Create(100, 80, 100, 80);
        section0.PageSetup.FooterDistance = 35.4;
        section0.PageSetup.HeaderDistance = 34.4;

        // Define the output file name
        const outputFileName = 'ModifyPageSetupOfSection.docx';

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
