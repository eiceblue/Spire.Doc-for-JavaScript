<template>
  <span>The following example shows how to set style of line shape</span>
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

        // Create a document
        let doc = wasmModule.Document.Create();

        // Add a section
        let sec = doc.AddSection();

        // Add a new paragraph
        let para = sec.AddParagraph();

        // Add a line shape
        let shape = para.AppendShape(100, 100, wasmModule.ShapeType.Line);

        // Set style of Line shape
        shape.FillColor = wasmModule.Color.get_Orange();
        shape.StrokeColor = wasmModule.Color.get_Black();
        shape.LineStyle = wasmModule.ShapeLineStyle.Single;
        shape.LineDashing = wasmModule.LineDashing.LongDashDotDotGEL;

        // Save the document
        const outputFileName = "SetLineShapeStyle.docx";
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