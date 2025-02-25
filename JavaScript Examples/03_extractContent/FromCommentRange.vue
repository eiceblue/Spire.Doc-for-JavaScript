<template>
  <span>The following example demonstrates how to extract content from comment range. </span>
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
        let inputFileName = 'Comments.docx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create a document
        let sourceDoc = wasmModule.Document.Create();

        //Load the document from disk.
        sourceDoc.LoadFromFile(inputFileName);

        //Create a destination document
        let destinationDoc = wasmModule.Document.Create();

        //Add section for destination document
        let destinationSec = destinationDoc.AddSection();

        //Get the first comment
        let comment = sourceDoc.Comments.get_Item(0);

        //Get the paragraph of obtained comment
        let para = comment.OwnerParagraph;

        //Get index of the CommentMarkStart
        let startIndex = para.ChildObjects.IndexOf(comment.CommentMarkStart);

        //Get index of the CommentMarkEnd
        let endIndex = para.ChildObjects.IndexOf(comment.CommentMarkEnd);

        //Traverse paragraph ChildObjects
        for (let i = startIndex; i <= endIndex; i++) {
          //Clone the ChildObjects of source document
          let doobj = para.ChildObjects.get(i).Clone();

          //Add to destination document
          destinationSec.AddParagraph().ChildObjects.Add(doobj);
        }
        // Define the output file name
        const outputFileName = 'FromCommentRange.docx';

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
