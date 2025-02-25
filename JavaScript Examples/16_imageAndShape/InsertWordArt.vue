<template>
  <span>This sample demonstrates how to insert WordArt in a Word document</span>
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


        // Load the input file into the virtual file system (VFS)
        const inputFileName = "InsertWordArt.docx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create Word document
        let doc = wasmModule.Document.Create();

        // Load Word document
        doc.LoadFromFile(inputFileName);

        // Add a paragraph
        let paragraph = doc.Sections.get(0).AddParagraph();

        // Add a shape
        let shape = paragraph.AppendShape(250, 70, wasmModule.ShapeType.TextWave4);

        // Set the position of the shape
        shape.VerticalPosition = 20;
        shape.HorizontalPosition = 80;

        // Set the text of WordArt
        shape.WordArt.Text = "Thanks for reading.";

        // Set the fill color
        shape.FillColor = wasmModule.Color.get_Red();

        // Set the border color of the text
        shape.StrokeColor = wasmModule.Color.get_Yellow();

        // Save the document
        const outputFileName = "InsertWordArt.docx";
        doc.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx });

        // Read the saved document from the virtual file system and convert it to a byte array
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);

        // Convert the byte array into a Blob object
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

        // Dispose of the document object to free resources
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