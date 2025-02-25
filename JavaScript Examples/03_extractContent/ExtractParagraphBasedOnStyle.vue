<template>
  <span>The following example shows how to extract paragraphs based on their styles. </span>
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
        let inputFileName = 'ExtractParagraphBasedOnStyle.docx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const doc = wasmModule.Document.Create();

        // Load a document from the virtual file system
        doc.LoadFromFile(inputFileName);

        let styleName1 = 'Heading1';
        let style1Text = '';
        style1Text += 'The following is the content of the paragraph with the style name ' + styleName1 + ': ' + '\n';
        //Extrct paragraph based on style
        for (let i = 0; i < doc.Sections.Count; i++) {
          let section = doc.Sections.get_Item(i);
          //travel the paragraphs
          for (let j = 0; j < section.Paragraphs.Count; j++) {
            let paragraph = section.Paragraphs.get_Item(j);
            if (paragraph.StyleName != null && paragraph.StyleName === styleName1) {
              style1Text += paragraph.Text;
            }
          }
        }

        // Define the output file name
        const outputFileName = 'ExtractParagraphBasedOnStyle.txt';
        // Save result file
        wasmModule.FS.writeFile(outputFileName, style1Text);

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: 'text/plain',
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
