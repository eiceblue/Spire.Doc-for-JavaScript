<template>
  <span>Click the following button to insert hyperlink in a Word document</span>
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
        let inputFileName = "Spire.Doc.png";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);        

        //Open a blank word document as template
        let document = wasmModule.Document.Create();
        let section = document.AddSection();

        //Insert hyperlink
        InsertHyperlink(section,inputFileName);

        // Define the output file name
        const outputFileName = "Hyperlink_output.docx";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

        // Clean up resources
        document.Close();
        document.Dispose();

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }

      function InsertHyperlink(section,inputFileName) {
     let paragraph = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();
     paragraph.AppendText("Spire.Doc for JS\r\n e-iceblue company Ltd. 2002-2010 All rights reserverd");
     paragraph.ApplyStyle({builtinStyle: wasmModule.BuiltinStyle.Heading2});

     paragraph = section.AddParagraph();
     paragraph.AppendText("Home page");
     paragraph.ApplyStyle({builtinStyle: wasmModule.BuiltinStyle.Heading2});
     paragraph = section.AddParagraph();
     paragraph.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", wasmModule.HyperlinkType.WebLink);

     paragraph = section.AddParagraph();
     paragraph.AppendText("Contact US");
     paragraph.ApplyStyle({builtinStyle: wasmModule.BuiltinStyle.Heading2});
     paragraph = section.AddParagraph();
     paragraph.AppendHyperlink("mailto:support@e-iceblue.com", "support@e-iceblue.com", wasmModule.HyperlinkType.EMailLink);

     paragraph = section.AddParagraph();
     paragraph.AppendText("Forum");
     paragraph.ApplyStyle({builtinStyle: wasmModule.BuiltinStyle.Heading2});
     paragraph = section.AddParagraph();
     paragraph.AppendHyperlink("www.e-iceblue.com/forum/", "www.e-iceblue.com/forum/", wasmModule.HyperlinkType.WebLink);

     paragraph = section.AddParagraph();
     paragraph.AppendText("Download Link");
     paragraph.ApplyStyle({builtinStyle: wasmModule.BuiltinStyle.Heading2});
     paragraph = section.AddParagraph();
     paragraph.AppendHyperlink("www.e-iceblue.com/Download/download-word-for-net-now.html", "www.e-iceblue.com/Download/download-word-for-net-now.html", wasmModule.HyperlinkType.WebLink);

     paragraph = section.AddParagraph();
     paragraph.AppendText("Insert Link On Image");
     paragraph.ApplyStyle({builtinStyle: wasmModule.BuiltinStyle.Heading2});
     paragraph = section.AddParagraph();

     let picture = paragraph.AppendPicture({imgFile: inputFileName});

     paragraph.AppendHyperlink("www.e-iceblue.com/Download/download-word-for-net-now.html", picture, wasmModule.HyperlinkType.WebLink);
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
