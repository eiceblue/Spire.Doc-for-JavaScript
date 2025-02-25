<template>
  <span>The following example shows how to replace image with new image in a Word document</span>
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
        const inputFileName = "ImageTemplate.docx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        const imageFileName = "E-iceblue.png";
        await wasmModule.FetchFileToVFS(imageFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Load Document
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        // Get all pictures in the Word document
        let pictures = []
        for (let i = 0; i < doc.Sections.Count; i++) {
          let sec = doc.Sections.get(i);
          for (let j = 0; j < sec.Paragraphs.Count; j++) {
            let para = sec.Paragraphs.get_Item(j);
            for (let k = 0; k < para.ChildObjects.Count; k++) {
              let docObj = para.ChildObjects.get(k);
              if (docObj.DocumentObjectType == wasmModule.DocumentObjectType.Picture) {
                pictures.push(docObj);
              }
            }
          }
        }

        // Replace the first picture with a new image file
        let picture = pictures[0]
        picture.LoadImage({ imgFile: imageFileName });

        // Save the document
        const outputFileName = "UpdateImage.docx";
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