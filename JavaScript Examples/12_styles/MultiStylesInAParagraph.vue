<template>
  <span>The example shows how to create a paragraph with multiple styles in a Word document.</span>
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

        //Create a Word document
        let doc = wasmModule.Document.Create();

        //Add a section
        let section = doc.AddSection();

        //Add a paragraph
        let para = section.AddParagraph();

        //Add a text range 1 and set its style
        let range = para.AppendText("Spire.Doc for JavaScript ");
        range.CharacterFormat.FontName = "Calibri";
        range.CharacterFormat.FontSize = 16;
        range.CharacterFormat.TextColor = wasmModule.Color.get_Blue();
        range.CharacterFormat.Bold = true;
        range.CharacterFormat.UnderlineStyle = wasmModule.UnderlineStyle.Single;

        //Add a text range 2 and set its style
        range = para.AppendText("is a professional Word JavaScript library");
        range.CharacterFormat.FontName = "Calibri";
        range.CharacterFormat.FontSize = 15;

        // Define the output file name
        const outputFileName = "MultiStylesInAParagraph-result.docx";

        // Save the document to the specified path
        doc.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});
        
        // Clean up resources
        doc.Dispose();

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
