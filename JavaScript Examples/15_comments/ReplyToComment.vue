<template>
  <span>The following example demonstrates how to add reply comments in a Word document</span>
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
        const inputFileName = "Comment.docx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Load the image file into the virtual file system (VFS)
        const imageFile = "logo.png";
        await wasmModule.FetchFileToVFS(imageFile, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a new document object  
        let doc = wasmModule.Document.Create();

        // Load the document
        doc.LoadFromFile(inputFileName);

        // Get the first comment
        let comment1 = doc.Comments.get_Item(0);

        // create a new comment and specify the author and content
        let replyComment1 = wasmModule.Comment.Create(doc);
        replyComment1.Format.Author = "E-iceblue";
        replyComment1.Body.AddParagraph().AppendText("Spire.Doc is a professional Word .NET library on operating Word documents.");

        // Add the new comment as a reply to the selected comment.
        comment1.ReplyToComment(replyComment1);

        // Load image
        let docPicture = wasmModule.DocPicture.Create(doc);
        docPicture.LoadImage(imageFile);

        // Insert a picture in the comment
        replyComment1.Body.Paragraphs.get_Item(0).ChildObjects.Add(docPicture);

        // Save the document
        const outputFileName = "ReplyToComment.docx";
        doc.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx });

        // Read the saved document from the virtual file system and convert it to a byte array
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);

        // Convert the byte array into a Blob object
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

        // Dispose of the doc
        doc.Dispose();

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