<template>
  <span>The following example shows how to insert a picture into comment in a Word document</span>
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

        // Load image file into the virtual file system (VFS)
        const imageFile = "E-iceblue.png";
        await wasmModule.FetchFileToVFS(imageFile, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a new document object   
        let doc = wasmModule.Document.Create();
        // Load file from VFS
        doc.LoadFromFile(inputFileName);

        // Get the first paragraph and insert comment
        let paragraph = doc.Sections.get(0).Paragraphs.get_Item(2);
        let comment = paragraph.AppendComment("This is a comment.");
        comment.Format.Author = "E-iceblue";

        // Load a picture
        let docPicture = wasmModule.DocPicture.Create(doc);
        docPicture.LoadImage(imageFile);

        // Insert the picture into the comment body
        comment.Body.AddParagraph().ChildObjects.Add(docPicture);

        // Save the document
        const outputFileName = "InsertPictureIntoComment.docx";
        doc.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx });

        // Read the saved document from the virtual file system and convert it to a byte array
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);

        // Convert the byte array into a Blob object
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

        // Dispose of the document object to free resources
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