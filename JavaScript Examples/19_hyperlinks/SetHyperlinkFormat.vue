<template>
  <span>Click the following button to change hyperlink color and remove hyperlink underline</span>
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
        let inputFileName = "BlankTemplate.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);        

        //Load Document
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);
        let section = doc.Sections.get_Item(0);

        //Add a paragraph and append a hyperlink to the paragraph
        let para1 = section.AddParagraph();
        para1.AppendText("Regular Link: ");
        //Format the hyperlink with default color and underline style
        let txtRange1 = para1.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", wasmModule.HyperlinkType.WebLink);
        txtRange1.CharacterFormat.FontName = "Times New Roman";
        txtRange1.CharacterFormat.FontSize = 12;
        let blankPara1 = section.AddParagraph();

        //Add a paragraph and append a hyperlink to the paragraph
        let para2 = section.AddParagraph();
        para2.AppendText("Change Color: ");
        //Format the hyperlink with red color and underline style
        let txtRange2 = para2.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", wasmModule.HyperlinkType.WebLink);
        txtRange2.CharacterFormat.FontName = "Times New Roman";
        txtRange2.CharacterFormat.FontSize = 12;
        txtRange2.CharacterFormat.TextColor = wasmModule.Color.get_Red();
        let blankPara2 = section.AddParagraph();

        //Add a paragraph and append a hyperlink to the paragraph
        let para3 = section.AddParagraph();
        para3.AppendText("Remove Underline: ");
        //Format the hyperlink with red color and no underline style
        let txtRange3 = para3.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", wasmModule.HyperlinkType.WebLink);
        txtRange3.CharacterFormat.FontName = "Times New Roman";
        txtRange3.CharacterFormat.FontSize = 12;
        txtRange3.CharacterFormat.UnderlineStyle = wasmModule.UnderlineStyle.None;

        // Define the output file name
        const outputFileName = "SetHyperlinkFormat_output.docx";

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
