<template>
  <span>Click the following button to set textbox format ( position, line style and internal margin ) in a Word document</span>
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

        // Create a new document
        const document = wasmModule.Document.Create();
        
        // Create a section
        let section = document.AddSection();

        // Add a text box and append sample text
        let TB = section.AddParagraph().AppendTextBox(310, 90);
        let para = TB.Body.AddParagraph();
        let TR = para.AppendText("Using Spire.Doc, developers will find " +
            "a simple and effective method to endow their applications with rich MS Word features. ");
        TR.CharacterFormat.FontName = "Cambria ";
        TR.CharacterFormat.FontSize = 13;

        // Set exact position for the text box
        TB.Format.HorizontalOrigin = wasmModule.HorizontalOrigin.Page;
        TB.Format.HorizontalPosition = 120;
        TB.Format.VerticalOrigin = wasmModule.VerticalOrigin.Page;
        TB.Format.VerticalPosition = 100;

        // Set line style for the text box
        TB.Format.LineStyle = wasmModule.TextBoxLineStyle.Double;
        TB.Format.LineColor = wasmModule.Color.get_CornflowerBlue();
        TB.Format.LineDashing = wasmModule.LineDashing.Solid;
        TB.Format.LineWidth = 5;

        // Set internal margin for the text box
        TB.Format.InternalMargin.Top = 15;
        TB.Format.InternalMargin.Bottom = 10;
        TB.Format.InternalMargin.Left = 12;
        TB.Format.InternalMargin.Right = 10;

        // Define the output file name
        const outputFileName = "TextBoxFormat.docx";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

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
