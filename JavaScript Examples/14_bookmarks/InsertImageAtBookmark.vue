<template>
  <span>The example shows how to insert an image at the location of bookmark in a Word document.</span>
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
        let inputFileName = "InsertImageAtBookmark.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        let imageFileName = "Word.png";
        await wasmModule.FetchFileToVFS(imageFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const doc = wasmModule.Document.Create();

        //Load the document
        doc.LoadFromFile(inputFileName);

        //Create an instance of BookmarksNavigator
        let bn = wasmModule.BookmarksNavigator.Create(doc);

        //Find a bookmark named Test
        bn.MoveToBookmark("Test", true, true);

        //Add a section
        let section0 = doc.AddSection();

        //Add a paragraph for the section
        let paragraph = section0.AddParagraph();

        //Add a picture into the paragraph
        paragraph.AppendPicture({imgFile : imageFileName});

        //Add the paragraph at the position of bookmark
        bn.InsertParagraph(paragraph);

        //Remove the section0
        doc.Sections.Remove(section0);

        // Define the output file name
        const outputFileName = "InsertImageAtBookmark-result.docx";

        // Save the document to the specified path
        doc.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Clean up resources
        doc.Dispose();

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
