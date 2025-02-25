<template>
  <span>The following example shows how to remove shape from a Word document</span>
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

        // Load Document
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);
        let section = doc.Sections.get(0);
        //Get all the child objects of paragraph
        for (let i = 0; i < section.Paragraphs.Count; i++) {
          let para = section.Paragraphs.get_Item(i);
          for (let j = 0; j < para.ChildObjects.Count; j++) {
            //If the child objects is shape object
            if (para.ChildObjects.get(j) instanceof wasmModule.ShapeObject) {
              //Remove the shape object
              para.ChildObjects.RemoveAt(j);
              j--;
            }
          }
        }

        //Save docx file
        const outputFileName = "RemoveShape.docx";
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