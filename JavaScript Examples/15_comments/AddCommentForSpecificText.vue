<template>
  <span>The following example shows how to add comment for specific text</span>
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

        // Create a new document
        let document = wasmModule.Document.Create();

        // Load file from VFS
        document.LoadFromFile(inputFileName);

        // Call the custom InsertComments function to insert comments into the document,
        InsertComments(document, "development");

        //Save the document
        const outputFileName = "AddCommentForSpecificText.docx";
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
    function InsertComments(doc, keystring) {
      //Find the key string
      let find = doc.FindString(keystring, false, true);

      //Create the commentmarkStart and commentmarkEnd
      let commentmarkStart = wasmModule.CommentMark.Create(doc);
      commentmarkStart.Type = wasmModule.CommentMarkType.CommentStart;
      let commentmarkEnd = wasmModule.CommentMark.Create(doc);
      commentmarkEnd.Type = wasmModule.CommentMarkType.CommentEnd;

      //Add the content for comment
      let comment = wasmModule.Comment.Create(doc);
      comment.Body.AddParagraph().Text = "Test comments";
      comment.Format.Author = "E-iceblue";

      //Get the textRange
      let range = find.GetAsOneRange();

      //Get its paragraph
      let para = range.OwnerParagraph;

      //Get the index of textRange
      let index = para.ChildObjects.IndexOf(range);

      //Add comment
      para.ChildObjects.Add(comment);

      //Insert the commentmarkStart and commentmarkEnd
      para.ChildObjects.Insert(index, commentmarkStart);
      para.ChildObjects.Insert(index + 2, commentmarkEnd);
    }
    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>