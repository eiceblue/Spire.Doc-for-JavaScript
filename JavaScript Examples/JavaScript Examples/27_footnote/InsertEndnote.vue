<template>
  <span>Click the following button to insert endnote in Word document</span>
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
        let inputFileName = "InsertEndnote.doc";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Load a document from the virtual file system
        document.LoadFromFile(inputFileName);

        // Get the first section of the document
        let s = document.Sections.get_Item(0);

        // Get the second paragraph in the section
        let p = s.Paragraphs.get_Item(1);

        // Add an endnote to the paragraph
        let endnote = p.AppendFootnote({type: wasmModule.FootnoteType.Endnote});

        // Append a new paragraph to the endnote's text body and add text
        let text = endnote.TextBody.AddParagraph().AppendText("Reference: Wikipedia");

        // Set the text format for the endnote content
        text.CharacterFormat.FontName = "Impact";
        text.CharacterFormat.FontSize = 14;
        text.CharacterFormat.TextColor = wasmModule.Color.get_DarkOrange(); 

        // Set the marker format for the endnote reference
        endnote.MarkerCharacterFormat.FontName = "Calibri"; 
        endnote.MarkerCharacterFormat.FontSize = 25; 
        endnote.MarkerCharacterFormat.TextColor = wasmModule.Color.get_DarkBlue(); 
        
        // Define the output file name
        const outputFileName = "InsertEndnote.docx";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

        // Clean up resources
        document.Dispose();

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
