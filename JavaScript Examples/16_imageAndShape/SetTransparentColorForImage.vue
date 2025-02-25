<template>
  <span>The following example shows how to set transeperant color for images in a Word document</span>
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

        // Load Document
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        // Get the first paragraph in the first section
        let paragraph = doc.Sections.get(0).Paragraphs.get_Item(0);

        // Set the blue color of the image(s) in the paragraph to transperant
        for (let i = 0; i < paragraph.ChildObjects.Count; i++) {
          let obj = paragraph.ChildObjects.get(i);
          if (obj instanceof wasmModule.DocPicture) {
            let picture = obj;
            picture.TransparentColor = wasmModule.Color.get_Blue();
          }
        }

        // Save and launch document
        const outputFileName = "SetTransparentColorForImage.docx";
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