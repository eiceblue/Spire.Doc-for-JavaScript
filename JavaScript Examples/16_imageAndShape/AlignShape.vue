<template>
  <span>The following example shows how to align the shape in a Word document</span>
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
        const inputFileName = "Shapes.docx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);
        
        // Create a new document
        let doc = wasmModule.Document.Create();

        // Load Document
        doc.LoadFromFile(inputFileName);

        // Get first section
        let section = doc.Sections.get_Item(0);

        for (let i = 0; i < section.Paragraphs.Count; i++) {
          let para = section.Paragraphs.get_Item(i);
          for (let j = 0; j < para.ChildObjects.Count; j++) {
            let obj = para.ChildObjects.get(j);
            if (obj instanceof wasmModule.ShapeObject) {
              //Set the horizontal alignment as center
              obj.HorizontalAlignment = wasmModule.ShapeHorizontalAlignment.Center;

              ////Set the vertical alignment as top
              //(obj as ShapeObject).VerticalAlignment = ShapeVerticalAlignment.Top;
            }
          }
        }

        //Save and launch document
        const outputFileName = "AlignShape.docx";
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