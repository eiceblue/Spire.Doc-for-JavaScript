<template>
  <span>Click the following button to check file format of loading file.</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from "vue";

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "Template.docx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();
        document.LoadFromFile(inputFileName);

        //Get file format
        let ff = document.DetectedFormatType;
        let fileFormat = "The file format is ";

        //Check the format info
        switch (ff) {
          case wasmModule.FileFormat.Doc:
            fileFormat += "Microsoft Word 97-2003 document.";
            break;
          case wasmModule.FileFormat.Dot:
            fileFormat += "Microsoft Word 97-2003 template.";
            break;
          case wasmModule.FileFormat.Docx:
            fileFormat += "Office Open XML WordprocessingML Macro-Free Document.";
            break;
          case wasmModule.FileFormat.Docm:
            fileFormat += "Office Open XML WordprocessingML Macro-Enabled Document.";
            break;
          case wasmModule.FileFormat.Dotx:
            fileFormat += "Office Open XML WordprocessingML Macro-Free Template.";
            break;
          case wasmModule.FileFormat.Dotm:
            fileFormat += "Office Open XML WordprocessingML Macro-Enabled Template.";
            break;
          case wasmModule.FileFormat.Rtf:
            fileFormat += "RTF format.";
            break;
          case wasmModule.FileFormat.WordML:
            fileFormat += "Microsoft Word 2003 WordprocessingML format.";
            break;
          case wasmModule.FileFormat.Html:
            fileFormat += "HTML format.";
            break;
          case wasmModule.FileFormat.WordXml:
            fileFormat += "Microsoft Word xml format for word 2007-2013.";
            break;
          case wasmModule.FileFormat.Odt:
            fileFormat += "OpenDocument Text.";
            break;
          case wasmModule.FileFormat.Ott:
            fileFormat += "OpenDocument Text Template.";
            break;
          case wasmModule.FileFormat.DocPre97:
            fileFormat += "Microsoft Word 6 or Word 95 format.";
            break;
          default:
            fileFormat += "Unknown format.";
            break;
        }
        // Define the output file name
        const outputFileName = "CheckFileFormat_out.txt";

        //Save to file.
        FS.writeFile(outputFileName, fileFormat.toString())

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: 'text/plain' });

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
