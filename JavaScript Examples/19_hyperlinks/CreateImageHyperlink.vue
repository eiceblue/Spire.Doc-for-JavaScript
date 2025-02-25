<template>
  <span>Click the following button to create an image hyperlink in a Word document</span>
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
        const inputFileName_1 = "Spire.Doc.png";     
        await wasmModule.FetchFileToVFS(inputFileName_1,"",`${import.meta.env.BASE_URL}static/data/`);   

        //Load Document
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        let section = doc.Sections.get_Item(0);
        //Add a paragraph
        let paragraph = section.AddParagraph();
        //Load an image to a DocPicture object

        let picture = wasmModule.DocPicture.Create(doc);
        //Add an image hyperlink to the paragraph
        picture.LoadImage(inputFileName_1);

        paragraph.AppendHyperlink({
            link: "https://www.e-iceblue.com/Introduce/word-for-net-introduce.html",
            picture: picture,
            type: wasmModule.HyperlinkType.WebLink
        });


        // Define the output file name
        const outputFileName = "CreateImageHyperlink_output.docx";

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
