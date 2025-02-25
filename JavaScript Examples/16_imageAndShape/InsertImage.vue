<template>
  <span>The following example shows how to insert an image at specified location in a Word document</span>
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
        const inputFileName = "BlankTemplate.docx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Load the image file into the virtual file system (VFS)
        const inputImgFileName = "Word.png";
        await wasmModule.FetchFileToVFS(inputImgFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a document
        let doc = wasmModule.Document.Create();
        // Load Document
        doc.LoadFromFile(inputFileName);

        // Get the first section
        let section = doc.Sections.get(0);
        // Add a paragraph
        let paragraph = section.Paragraphs.Count > 0 ? section.Paragraphs.get_Item(0) : section.AddParagraph();
        paragraph.AppendText("The sample demonstrates how to insert an image into a document.");
        paragraph.ApplyStyle({ builtinStyle: wasmModule.BuiltinStyle.Heading2 });
        paragraph = section.AddParagraph();
        paragraph.AppendText("The above is a picture.");

        // Create a picture
        let picture = wasmModule.DocPicture.Create(doc);
        picture.LoadImage({ imgFile: inputImgFileName });
        // Set image's position
        picture.HorizontalPosition = 50.0;
        picture.VerticalPosition = 60.0;

        // Set image's size
        picture.Width = 200;
        picture.Height = 200;

        // Set textWrappingStyle with image;
        picture.TextWrappingStyle = wasmModule.TextWrappingStyle.Through;
        // Insert the picture at the beginning of the second paragraph
        paragraph.ChildObjects.Insert(0, picture);

        // Save the document
        const outputFileName = "InsertImage.docx";
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