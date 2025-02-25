<template>
  <span>The following example demonstrates how to insert comment in a Word document</span>
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
        const inputFileName = "CommentTemplate.docx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a new document object  
        let document = wasmModule.Document.Create();
        // Load file from VFS
        document.LoadFromFile(inputFileName);

        // Call the custom InsertComments function to insert comments into the document,
        InsertComments(document.Sections.get(0));

        //Save the document
        const outputFileName = "Comment.docx";
        document.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx });

        // Read the saved document from the virtual file system and convert it to a byte array
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);

        // Convert the byte array into a Blob object
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

        // Dispose of the document object to free resources
        document.Dispose();

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
    };

    function InsertComments(section) {
      //Insert comment.
      let paragraph = section.Paragraphs.get_Item(1);
      let comment = paragraph.AppendComment("Spire.Doc for .NET");
      comment.Format.Author = "E-iceblue";
      comment.Format.Initial = "CM";
    }
    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>