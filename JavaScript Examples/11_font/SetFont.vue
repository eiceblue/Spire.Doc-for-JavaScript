<template>
  <span>The example shows how to set font in Word document.</span>
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
        let inputFileName = "SetFont.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();
        document.LoadFromFile(inputFileName);

        //Get the first section
        let s = document.Sections.get(0);

        //Get the second paragraph
        let p = s.Paragraphs.get_Item(1);

        //Create a characterFormat object
        let format = wasmModule.CharacterFormat.Create(document);
        //Set font

        format.FontName = "Arial";
        format.FontSize = 16;
        //Loop through the childObjects of paragraph
        for (let i = 0; i < p.ChildObjects.Count; i++) {
            let childObj = p.ChildObjects.get(i);
            if (childObj instanceof wasmModule.TextRange) {
                //Apply character format
                let tr = childObj;
                tr.ApplyCharacterFormat(format);
            }
        }

        // Define the output file name
        const outputFileName = "SetFont-result.docx";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Clean up resources
        document.Dispose();

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

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
