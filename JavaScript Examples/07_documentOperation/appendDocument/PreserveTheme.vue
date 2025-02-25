<template>
  <span>Click the following button to preserve theme when copying sections from one Word document to another.</span>
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

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "Theme.docx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        //Load the source document
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        //Create a new Word document
        let newWord = wasmModule.Document.Create();

        //Clone default style, theme, compatibility from the source file to the destination document
        doc.CloneDefaultStyleTo(newWord);
        doc.CloneThemesTo(newWord);
        doc.CloneCompatibilityTo(newWord);

        //Add the cloned section to destination document
        newWord.Sections.Add(doc.Sections.get_Item(0).Clone());

        // Define the output file name
        const outputFileName = "PreserveTheme_out.docx";

        // Save the document to the specified path
        newWord.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

        // Clean up resources
        newWord.Dispose();
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
