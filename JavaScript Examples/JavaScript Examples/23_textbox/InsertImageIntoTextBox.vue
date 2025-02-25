<template>
  <span>Click the following button to insert image into text box in a Word document</span>
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

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "Spire.Doc.png";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Create a new section
        let section = document.AddSection();
        
        // Create a new paragraph
        let paragraph = section.AddParagraph();

        // Append a textbox to paragraph
        let tb = paragraph.AppendTextBox(220, 220);

        // Set the position of the textbox
        tb.Format.HorizontalOrigin = wasmModule.HorizontalOrigin.Page;
        tb.Format.HorizontalPosition = 50;
        tb.Format.VerticalOrigin = wasmModule.VerticalOrigin.Page;
        tb.Format.VerticalPosition = 50;

        // Set the fill effect of textbox as picture
        tb.Format.FillEfects.Type = wasmModule.BackgroundType.Picture;

        //Fill the textbox with a picture
        tb.Format.FillEfects.SetPicture(inputFileName);

        // Define the output file name
        const outputFileName = "InsertImageIntoTextBox.docx";

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
