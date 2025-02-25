<template>
  <span>Click the following button to set whether page border surrounds header and footer or not in a Word document.</span>
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

        //Create a new document
        let doc = wasmModule.Document.Create();
        let section = doc.AddSection();

        //Add a sample page border to the document
        section.PageSetup.Borders.BorderType = wasmModule.BorderStyle.Wave;
        section.PageSetup.Borders.Color = wasmModule.Color.get_Green();
        section.PageSetup.Borders.Left.Space = 20;
        section.PageSetup.Borders.Right.Space = 20;

        //Add a header and set its format
        let paragraph1 = section.HeadersFooters.Header.AddParagraph();
        paragraph1.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;
        let headerText = paragraph1.AppendText("Header isn't included in page border");
        headerText.CharacterFormat.FontName = "Calibri";
        headerText.CharacterFormat.FontSize = 20;
        headerText.CharacterFormat.Bold = true;

        //Add a footer and set its format
        let paragraph2 = section.HeadersFooters.Footer.AddParagraph();
        paragraph2.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Left;
        let footerText = paragraph2.AppendText("Footer is included in page border");
        footerText.CharacterFormat.FontName = "Calibri";
        footerText.CharacterFormat.FontSize = 20;
        footerText.CharacterFormat.Bold = true;

        //Set the header not included in the page border while the footer included
        section.PageSetup.PageBorderIncludeHeader = false;
        section.PageSetup.HeaderDistance = 40;
        section.PageSetup.PageBorderIncludeFooter = true;
        section.PageSetup.FooterDistance = 40;


        // Define the output file name
        const outputFileName = "PageBorderSurround_output.docx";

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
