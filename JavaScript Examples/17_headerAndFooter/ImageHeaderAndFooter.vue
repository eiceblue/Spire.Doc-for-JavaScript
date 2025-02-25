<template>
  <span>Click the following button to insert images in header and footer in a Word document.</span>
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
        let inputFileName = "Template.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        const inputImgFileName = "E-iceblue.png"
        await wasmModule.FetchFileToVFS(inputImgFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        const inputImgFileName_1 = "logo.png"
        await wasmModule.FetchFileToVFS(inputImgFileName_1,"",`${import.meta.env.BASE_URL}static/data/`);

        //Load the document from disk
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        //Get the header of the first page
        let header = doc.Sections.get(0).HeadersFooters.Header;

        //Add a paragraph for the header
        let paragraph = header.AddParagraph();

        //Set the format of the paragraph
        paragraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;

        //Append a picture in the paragraph
        let headerimage = paragraph.AppendPicture({imgFile: inputImgFileName});
        headerimage.VerticalAlignment = wasmModule.ShapeVerticalAlignment.Bottom;

        //Get the footer of the first section
        let footer = doc.Sections.get_Item(0).HeadersFooters.Footer;

        //Add a paragraph for the footer
        let paragraph2 = footer.AddParagraph();

        //Set the format of the paragraph
        paragraph2.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Left;

        //Append a picture in the paragraph
        let footerimage = paragraph2.AppendPicture({imgFile: inputImgFileName_1});

        //Append text in the paragraph
        let TR = paragraph2.AppendText("Copyright © 2013 e-iceblue. All Rights Reserved.");
        TR.CharacterFormat.FontName = "Arial";
        TR.CharacterFormat.FontSize = 10;
        TR.CharacterFormat.TextColor = wasmModule.Color.get_Black();

        const outputFileName = "ImageHeaderAndFooter.docx";
        // Save the document to the specified path
        doc.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

        // Clean up resources
        doc.Close();
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
