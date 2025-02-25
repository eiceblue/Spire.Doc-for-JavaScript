<template>
  <span>The following example shows how to copy paragraphs from one Word document to another. </span>
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
        let inputFileName = 'Template_Docx_5.docx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        let inputFileName2 = 'Logo.jpg';
        await wasmModule.FetchFileToVFS(inputFileName2, '', `${import.meta.env.BASE_URL}static/data/`);
        //Create Word document1.
        let document1 = wasmModule.Document.Create();

        //Load the file from disk.
        document1.LoadFromFile(inputFileName);

        //Create a new document.
        let document2 = wasmModule.Document.Create();

        //Get paragraph 1 and paragraph 2 in document1.
        let s = document1.Sections.get(0);
        let p1 = s.Paragraphs.get_Item(0);
        let p2 = s.Paragraphs.get_Item(1);

        //Copy p1 and p2 to document2.
        let s2 = document2.AddSection();
        let NewPara1 = p1.Clone();
        s2.Paragraphs.Add(NewPara1);

        let NewPara2 = p2.Clone();
        s2.Paragraphs.Add(NewPara2);

        //Add watermark.
        let WM = wasmModule.PictureWatermark.Create();
        // Set the Picture property of WM to an image
        WM.SetPicture(inputFileName2);
        // Set the Watermark property of document2 to WM
        document2.Watermark = WM;
        // Define the output file name
        const outputFileName = 'CopyParagraph.docx';

        // Save the document to the specified path
        document2.SaveToFile({
          fileName: outputFileName,
          fileFormat: wasmModule.FileFormat.Docx2013,
        });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });

        // Clean up resources
        document2.Dispose();

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
