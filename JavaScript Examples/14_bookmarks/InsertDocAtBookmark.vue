<template>
  <span>The example shows how to insert a document at the location of bookmark in a Word document. </span>
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
        let inputFileName_1 = "InsertDocAtBookmark1.docx";
        await wasmModule.FetchFileToVFS(inputFileName_1,"",`${import.meta.env.BASE_URL}static/data/`);

                // Load the sample file into the virtual file system (VFS)
        let inputFileName_2 = "InsertDocAtBookmark2.docx";
        await wasmModule.FetchFileToVFS(inputFileName_2,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create the first document
        let document1 = wasmModule.Document.Create();

        //Load the first document from disk.
        document1.LoadFromFile(inputFileName_1);

        //Create the second document
        let document2 = wasmModule.Document.Create();

        //Load the second document from disk.
        document2.LoadFromFile(inputFileName_2);

        //Get the first section of the first document
        let section1 = document1.Sections.get_Item();

        //Locate the bookmark
        let bn = wasmModule.BookmarksNavigator.Create(document1);

        //Find bookmark by name
        bn.MoveToBookmark("Test", true, true);

        //Get bookmarkStart
        let start = bn.CurrentBookmark.BookmarkStart;

        //Get the owner paragraph
        let para = start.OwnerParagraph;

        //Get the para index
        let index = section1.Body.ChildObjects.IndexOf(para);

        //Insert the paragraphs of document2
        for (let i = 0; i < document2.Sections.Count; i++) {
            let section2 = document2.Sections.get_Item(i);
            for (let j = 0; j < section2.Paragraphs.Count; j++) {
                let paragraph = section2.Paragraphs.get_Item(j);
                section1.Body.ChildObjects.Insert(index + 1, paragraph.Clone());
            }
        }

        // Define the output file name
        const outputFileName = "InsertDocAtBookmark-result.docx";

        // Save the document to the specified path
        document1.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Clean up resources
        document1.Dispose();

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
