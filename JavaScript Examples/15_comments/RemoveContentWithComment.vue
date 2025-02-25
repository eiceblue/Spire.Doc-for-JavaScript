<template>
  <span>The example demonstrates how to remove content with comment in a Word document</span>
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
        const inputFileName = "Comments.docx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a new document object  
        let document = wasmModule.Document.Create();

        // Load the document from VFS
        document.LoadFromFile(inputFileName);

        // Get the first comment
        let comment = document.Comments.get_Item(0);

        // Get the paragraph of obtained comment
        let para = comment.OwnerParagraph;

        // Get index of the CommentMarkStart
        let startIndex = para.ChildObjects.IndexOf(comment.CommentMarkStart);

        // Get index of the CommentMarkEnd
        let endIndex = para.ChildObjects.IndexOf(comment.CommentMarkEnd);

        // Create a list
        let list = [];

        // Get TextRanges between the indexes
        for (let i = startIndex; i < endIndex; i++) {
          if (para.ChildObjects.get(i) instanceof wasmModule.TextRange) {
            list.push(para.ChildObjects.get(i));
          }
        }

        // Insert a new TextRange
        let textRange = wasmModule.TextRange.Create(document);

        // Set text is null
        textRange.Text = "";

        // Insert the new textRange
        para.ChildObjects.Insert(endIndex, textRange);

        // Remove previous TextRanges
        for (let i = 0; i < list.length; i++) {
          para.ChildObjects.Remove(list[i]);
        }

        // Save the document
        const outputFileName = "RemoveContentWithComment.docx";
        document.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx });

        // Read the saved document from the virtual file system and convert it to a byte array
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);

        // Convert the byte array into a Blob object
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

        // Dispose of the doc
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