<template>
  <span>The following example shows how to reset the size of image in a Word document </span>
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

        // Get the first secion
        let section = doc.Sections.get(0);
        // Get the first paragraph
        let paragraph = section.Paragraphs.get_Item(0);

        // Reset the image size of the first paragraph
        for (let i = 0; i < paragraph.ChildObjects.Count; i++) {
          let docObj = paragraph.ChildObjects.get(i);
          if (docObj instanceof wasmModule.DocPicture) {
            let picture = docObj;
            picture.Width = 50;
            picture.Height = 50;
          }
        }

        //Save the document
        const outputFileName = "ResetImageSize.docx";
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