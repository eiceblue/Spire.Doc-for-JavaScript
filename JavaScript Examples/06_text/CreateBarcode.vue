<template>
  <span>The following example shows how to create barcode in a Word document. </span>
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

        //Create a document
        let doc = wasmModule.Document.Create();

        //Add a paragraph
        let p = doc.AddSection().AddParagraph();

        //Add barcode and set its format
        let txtRang = p.AppendText('H63TWX11072');
        //Set barcode font name, note you need to install the barcode font on your system at first
        txtRang.CharacterFormat.FontName = 'C39HrP60DlTt';
        txtRang.CharacterFormat.FontSize = 80;
        txtRang.CharacterFormat.TextColor = wasmModule.Color.get_SeaGreen();

        // Define the output file name
        const outputFileName = 'CreateBarcode.docx';

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
