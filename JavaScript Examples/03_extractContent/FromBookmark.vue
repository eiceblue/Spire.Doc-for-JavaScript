<template>
  <span>The following example demonstrates how to extract content from a bookmark. </span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName"> Click here to download the generated file </a>
</template>

<script>
import { ref } from 'vue';

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref('');

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = 'Bookmark.docx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const doc = wasmModule.Document.Create();

        // Load a document from the virtual file system
        doc.LoadFromFile(inputFileName);

        //Create a destination document
        let destinationDoc = wasmModule.Document.Create();

        //Add a section for destination document
        let section = destinationDoc.AddSection();

        //Add a paragraph for destination document
        let paragraph = section.AddParagraph();

        //Locate the bookmark in source document
        let navigator = wasmModule.BookmarksNavigator.Create(doc);

        //Find bookmark by name
        navigator.MoveToBookmark({
          bookmarkName: 'Test',
          isStart: true,
          isAfter: true,
        });

        //get text body part
        let textBodyPart = navigator.GetBookmarkContent();

        //Create a TextRange type list
        let list = [];

        //Traverse the items of text body
        for (let i = 0; i < textBodyPart.BodyItems.Count; i++) {
          let item = textBodyPart.BodyItems.get(i);

          for (let j = 0; j < item.ChildObjects.Count; j++) {
            let childObject = item.ChildObjects.get(j);

            //Add it into list
            let range = childObject;
            list.push(range);
          }
        }

        //Add the extract content to destinationDoc document
        for (let m = 0; m < list.length; m++) {
          paragraph.Items.Add(list[m].Clone());
        }

        // Define the output file name
        const outputFileName = 'FromBookmark.docx';

        // Save the document to the specified path
        destinationDoc.SaveToFile({
          fileName: outputFileName,
          fileFormat: wasmModule.FileFormat.Docx2013,
        });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });

        // Clean up resources
        destinationDoc.Dispose();

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
