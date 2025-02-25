<template>
  <span>Click the following button to insert footnote into a Word document</span>
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
        let inputFileName = "SampleB_2.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Load a document from the virtual file system
        document.LoadFromFile(inputFileName);

        // Find the first matched string in the document
        let selection = document.FindString("Spire.Doc", false, true);

        // Get the text range of the found string
        let textRange = selection.GetAsOneRange();

        // Get the paragraph that contains the matched text range
        let paragraph = textRange.OwnerParagraph;

        // Get the index of the text range within the paragraph's child objects
        let index = paragraph.ChildObjects.IndexOf(textRange);

        // Append a footnote to the paragraph
        let footnote = paragraph.AppendFootnote({ type: wasmModule.FootnoteType.Footnote });

        // Insert the footnote into the paragraph just after the matched text range
        paragraph.ChildObjects.Insert(index + 1, footnote);

        // Add a new paragraph to the footnote's text body and append text
        textRange = footnote.TextBody.AddParagraph().AppendText("Welcome to evaluate Spire.Doc");

        // Set the text format for the footnote content
        textRange.CharacterFormat.FontName = "Arial Black"; 
        textRange.CharacterFormat.FontSize = 10; 
        textRange.CharacterFormat.TextColor = wasmModule.Color.get_DarkGray(); 

        // Set the marker format for the footnote reference
        footnote.MarkerCharacterFormat.FontName = "Calibri";
        footnote.MarkerCharacterFormat.FontSize = 12; 
        footnote.MarkerCharacterFormat.Bold = true; 
        footnote.MarkerCharacterFormat.TextColor = wasmModule.Color.get_DarkGreen();
        
        // Define the output file name
        const outputFileName = "InsertFootnote.docx";

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
