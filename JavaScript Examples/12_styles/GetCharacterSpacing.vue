<template>
  <span>The example demonstrates how to get character spacing in a Word document. </span>
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
        let inputFileName = "GetCharacterSpacing.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Load a document from the virtual file system
        document.LoadFromFile(inputFileName);

        //Get the first section of document
        let section = document.Sections.get(0);

        //Get the first paragraph
        let paragraph = section.Paragraphs.get_Item(0);

        //Define two variables
        let fontName = "";
        let fontSpacing = 0;

        //Traverse the ChildObjects
        for (let i = 0; i < paragraph.ChildObjects.Count; i++) {
            let docObj = paragraph.ChildObjects.get(i);
            //If it is TextRange
            if (docObj instanceof wasmModule.TextRange) {
                let textRange = docObj;
                fontName = textRange.CharacterFormat.FontName;

                //Get the character spacing
                fontSpacing = textRange.CharacterFormat.CharacterSpacing;
            }
        }

        // Define the output file name
        const outputFileName = "GetCharacterSpacing-result.docx";

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
