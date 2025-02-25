<template>
  <span>The example demonstrates how to remove bookmark content in a Word document.</span>
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
        let inputFileName = "RemoveBookmarkContent.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        //Load the document from disk.
        document.LoadFromFile(inputFileName);

        //Get the bookmark by name.
        let bookmark = document.Bookmarks._get_Item("Test");

        let para = bookmark.BookmarkStart.Owner;
        let startIndex = para.ChildObjects.IndexOf(bookmark.BookmarkStart);
        para = bookmark.BookmarkEnd.Owner;
        let endIndex = para.ChildObjects.IndexOf(bookmark.BookmarkEnd);

        //Remove the content object, and Start from next of BookmarkStart object, end up with previous of BookmarkEnd object.
        //This method is only to remove the content of the bookmark.
        for (let i = startIndex + 1; i < endIndex; i++) {
            para.ChildObjects.RemoveAt(startIndex + 1);
        }

        // Define the output file name
        const outputFileName = "RemoveBookmarkContent-result.docx";

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
