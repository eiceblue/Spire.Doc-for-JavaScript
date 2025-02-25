<template>
  <span>The example shows how to extract text within a bookmark in a Word document.</span>
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
        let inputFileName = "ExtractBookmarkText.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const doc = wasmModule.Document.Create();

        // Load a document from the virtual file system
        doc.LoadFromFile(inputFileName);

        //Creates a BookmarkNavigator instance to access the bookmark
        let navigator = wasmModule.BookmarksNavigator.Create(doc);
        //Locate a specific bookmark by bookmark name
        navigator.MoveToBookmark("Content");
        let textBodyPart = navigator.GetBookmarkContent();

        //Iterate through the items in the bookmark content to get the text
        let text = "";
        for (let i = 0; i < textBodyPart.BodyItems.Count; i++) {
            let item = textBodyPart.BodyItems.get(i);
            if (item instanceof wasmModule.Paragraph) {
                for (let j = 0; j < item.ChildObjects.Count; j++) {
                    let childObject = item.ChildObjects.get(j);
                    if (childObject instanceof wasmModule.TextRange) {
                        text += childObject.Text;
                  }
                }
            }
        }
          // Define the output file name
          const outputFileName = "ExtractBookmarkText-result.txt";
          //Save to TXT File and launch it
          wasmModule.FS.writeFile(outputFileName, text);
          doc.Close();

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "text/plain"});

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
