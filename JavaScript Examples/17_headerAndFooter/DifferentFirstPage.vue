<template>
  <span>Click the following button to add different first page header and footer in a Word document</span>
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
        let inputFileName = "MultiplePages.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        let inputImgFileName = "E-iceblue.png";
        await wasmModule.FetchFileToVFS(inputImgFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Load the document
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        //Get the section and set the property true
        let section = doc.Sections.get_Item(0);
        section.PageSetup.DifferentFirstPageHeaderFooter = true;

        //Set the first page header. Here we append a picture in the header

        let paragraph1 = section.HeadersFooters.FirstPageHeader.AddParagraph();
        paragraph1.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;
        let headerimage = paragraph1.AppendPicture({imgFile: inputImgFileName});

        //Set the first page footer
        let paragraph2 = section.HeadersFooters.FirstPageFooter.AddParagraph();
        paragraph2.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
        let FF = paragraph2.AppendText("First Page Footer");
        FF.CharacterFormat.FontSize = 10;

        //Set the other header & footer. If you only need the first page header & footer, don't set this
        let paragraph3 = section.HeadersFooters.Header.AddParagraph();
        paragraph3.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
        let NH = paragraph3.AppendText("Spire.Doc for .NET");
        NH.CharacterFormat.FontSize = 10;

        let paragraph4 = section.HeadersFooters.Footer.AddParagraph();
        paragraph4.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
        let NF = paragraph4.AppendText("E-iceblue");
        NF.CharacterFormat.FontSize = 10;

        // Define the output file name
        const outputFileName = "DifferentFirstPage.docx";

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
