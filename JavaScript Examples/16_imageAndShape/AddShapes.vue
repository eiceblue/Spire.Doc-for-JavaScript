<template>
  <span>This sample demonstrates how to add shapes in a Word document</span>
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

        // Create a new document
        let doc = wasmModule.Document.Create();

        // Add a section
        let sec = doc.AddSection();

        // Add a paragraph
        let para = sec.AddParagraph();
        let x = 60, y = 40, lineCount = 0;
        for (let i = 1; i < 20; i++) {
          if (lineCount > 0 && lineCount % 8 == 0) {
            para.AppendBreak(wasmModule.BreakType.PageBreak);
            x = 60;
            y = 40;
            lineCount = 0;
          }
          //Add shape and set its size and position.
          let shape = para.AppendShape(50, 50, wasmModule.ShapeType.fromValue(i));
          shape.HorizontalOrigin = wasmModule.HorizontalOrigin.Page;
          shape.HorizontalPosition = x;
          shape.VerticalOrigin = wasmModule.VerticalOrigin.Page;
          shape.VerticalPosition = y + 50;
          x = x + shape.Width + 50;
          if (i > 0 && i % 5 == 0) {
            y = y + shape.Height + 120;
            lineCount++;
            x = 60;
          }
        }

        // Save the document
        const outputFileName = "AddShapes.docx";
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