<template>
  <span>The example shows how to embed private font into Word document.</span>
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
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF","/Library/Fonts/",`${import.meta.env.BASE_URL}static/font/`);
        await wasmModule.FetchFileToVFS("PT Serif Caption.ttf","/Library/Fonts/",`${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "EmbedPrivateFont.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();
        document.LoadFromFile(inputFileName);

        //Get the first section and add a paragraph
        let section = document.Sections.get_Item(0);
        let p = section.AddParagraph();

        //Append text to the paragraph, then set the font name and font size
        let range = p.AppendText("Your Office Development Master");
        range.CharacterFormat.FontName = "PT Serif Caption";
        range.CharacterFormat.FontSize = 20;

        //Allow embedding font in document
        document.EmbedFontsInFile = true;

        //Embed private font from font file into the document
        document.AddPrivateFont(wasmModule.PrivateFontPath.Create("PT Serif Caption",  "PT Serif Caption.ttf"))

        // Define the output file name
        const outputFileName = "EmbedPrivateFont-result.docx";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});
     
        // Clean up resources
        document.Dispose();

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

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
