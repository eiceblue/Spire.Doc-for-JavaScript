<template>
  <span>Click the following button to add different headers and footers for odd and even pages in a Word document.</span>
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

        //Load the document
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        //Get the section and
        let section = doc.Sections.get_Item(0);

        //Set the DifferentOddAndEvenPagesHeaderFooter property to ture
        section.PageSetup.DifferentOddAndEvenPagesHeaderFooter = true;

        //Add odd header
        let P3 = section.HeadersFooters.OddHeader.AddParagraph();
        let OH = P3.AppendText("Odd Header");
        P3.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
        OH.CharacterFormat.FontName = "Arial";
        OH.CharacterFormat.FontSize = 10;

        //Add even header
        let P4 = section.HeadersFooters.EvenHeader.AddParagraph();
        let EH = P4.AppendText("Even Header from E-iceblue Using Spire.Doc");
        P4.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
        EH.CharacterFormat.FontName = "Arial";
        EH.CharacterFormat.FontSize = 10;

        //Add odd footer
        let P2 = section.HeadersFooters.OddFooter.AddParagraph();
        let OF = P2.AppendText("Odd Footer");
        P2.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
        OF.CharacterFormat.FontName = "Arial";
        OF.CharacterFormat.FontSize = 10;

        //Add even footer
        let P1 = section.HeadersFooters.EvenFooter.AddParagraph();
        let EF = P1.AppendText("Even Footer from E-iceblue Using Spire.Doc");
        EF.CharacterFormat.FontName = "Arial";
        EF.CharacterFormat.FontSize = 10;
        P1.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;

        const outputFileName = "OddAndEvenHeaderFooter_output.docx";
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
