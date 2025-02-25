<template>
  <span>The following example demonstrates how to insert image into a Word document</span>
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

        const imageFile = "Spire.Doc.png";
        await wasmModule.FetchFileToVFS(imageFile, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a document
        let document = wasmModule.Document.Create();

        // Add a seciton
        let section = document.AddSection();

        // Insert image
        InsertImage(section,imageFile);

        // Save the document
        const outputFileName = "Image.docx";
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

    function InsertImage(section,imageFile) {
      //Add paragraph
      let paragraph = section.AddParagraph();
      paragraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Left;

      let picture = paragraph.AppendPicture({ imgFile: imageFile });

      picture.Width = 100;
      picture.Height = 100;

      paragraph = section.AddParagraph();
      paragraph.Format.LineSpacing = 20;
      let tr = paragraph.AppendText("Spire.Doc for .NET is a professional Word .NET library specially designed for developers to create, read, write, convert and print Word document files from any .NET( C#, VB.NET, ASP.NET) platform with fast and high quality performance. ");
      tr.CharacterFormat.FontName = "Arial";
      tr.CharacterFormat.FontSize = 14;

      section.AddParagraph();
      paragraph = section.AddParagraph();
      paragraph.Format.LineSpacing = 20;
      tr = paragraph.AppendText("As an independent Word .NET component, Spire.Doc for .NET doesn't need Microsoft Word to be installed on the machine. However, it can incorporate Microsoft Word document creation capabilities into any developers' .NET applications.");
      tr.CharacterFormat.FontName = "Arial";
      tr.CharacterFormat.FontSize = 14;

    }

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>