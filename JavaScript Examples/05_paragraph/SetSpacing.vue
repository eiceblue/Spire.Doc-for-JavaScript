<template>
  <span>The following example shows how to set before and after spacing for a paragraph. </span>
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

        //Create Word document.
        let document = wasmModule.Document.Create();

        //Load the file from disk.
        document.LoadFromFile(inputFileName);
        //Add the text strings to the paragraph and set the style.
        let para = wasmModule.Paragraph.Create(document);
        let str = 'This is an inserted paragraph.';
        let textRange1 = para.AppendText(str);
        textRange1.CharacterFormat.TextColor = wasmModule.Color.get_Blue();
        textRange1.CharacterFormat.FontSize = 15;

        //set the spacing before and after.
        para.Format.BeforeAutoSpacing = false;
        para.Format.BeforeSpacing = 10;
        para.Format.AfterAutoSpacing = false;
        para.Format.AfterSpacing = 10;

        //insert the added paragraph to the word document.
        document.Sections.get(0).Paragraphs.Insert(1, para);

        // Define the output file name
        const outputFileName = 'SetSpacing.docx';

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
