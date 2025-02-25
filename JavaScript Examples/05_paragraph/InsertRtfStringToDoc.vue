<template>
  <span>The following example shows how to insert RTF string in Word document. </span>
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

        //Create Word document.
        let document = wasmModule.Document.Create();

        //Add a new section.
        let section = document.AddSection();

        //Add a paragraph to the section.
        let para = section.AddParagraph();

        //Declare a String variable to store the Rtf string.
        let rtfString = '{\\rtf1\\ansi\\deff0 {\\fonttbl {\\f0 hakuyoxingshu7000;}}\\f0\\fs28 Hello, World}';

        //Append Rtf string to paragraph.
        para.AppendRTF(rtfString);

        // Define the output file name
        const outputFileName = 'InsertRtfStringToDoc.docx';

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
