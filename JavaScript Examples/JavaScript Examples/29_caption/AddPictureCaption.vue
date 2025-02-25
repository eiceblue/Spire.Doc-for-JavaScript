<template>
  <span>Click the following button to add caption for picture</span>
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
        await wasmModule.FetchFileToVFS("arial.ttf","/Library/Fonts/",`${import.meta.env.BASE_URL}static/font/`);
        await wasmModule.FetchFileToVFS("calibri.ttf","/Library/Fonts/",`${import.meta.env.BASE_URL}static/font/`);
        await wasmModule.FetchFileToVFS("cambria_B.ttc","/Library/Fonts/",`${import.meta.env.BASE_URL}static/font/`);
        await wasmModule.FetchFileToVFS("Lucida Sans Unicode.ttf","/Library/Fonts/",`${import.meta.env.BASE_URL}static/font/`);

        await wasmModule.FetchFileToVFS("symbol.ttf","/Library/Fonts/",`${import.meta.env.BASE_URL}static/font/`);
        await wasmModule.FetchFileToVFS("times.ttf","/Library/Fonts/",`${import.meta.env.BASE_URL}static/font/`);
        await wasmModule.FetchFileToVFS("wingding.ttf","/Library/Fonts/",`${import.meta.env.BASE_URL}static/font/`);


        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "Spire.Doc.png";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        let inputFileName1 = "Word.png";
        await wasmModule.FetchFileToVFS(inputFileName1,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Create a new section
        let section = document.AddSection();

        // Add the first paragraph to the section
        let par1 = section.AddParagraph();
        par1.Format.AfterSpacing = 10; 

        // Append the first picture to the paragraph
        let pic1 = par1.AppendPicture({ imgFile: inputFileName });

        // Set the dimensions of the first picture
        pic1.Height = 100;  
        pic1.Width = 120; 

        // Add a caption to the first picture
        let format = wasmModule.CaptionNumberingFormat.Number; 
        pic1.AddCaption("Figure", format, wasmModule.CaptionPosition.BelowItem); 

        // Add the second paragraph to the section
        let par2 = section.AddParagraph();

        // Append the second picture to the second paragraph
        let pic2 = par2.AppendPicture({ imgFile: inputFileName1 });

        // Set the dimensions of the second picture
        pic2.Height = 100; 
        pic2.Width = 120; 

        // Add a caption to the second picture
        pic2.AddCaption("Figure", format, wasmModule.CaptionPosition.BelowItem); 

        // Update fields
        document.IsUpdateFields = true;
        
        // Define the output file name
        const outputFileName = "AddPictureCaption.docx";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

        // Clean up resources
        document.Dispose();

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
